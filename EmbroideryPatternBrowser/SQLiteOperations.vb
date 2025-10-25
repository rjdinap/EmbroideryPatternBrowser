Imports System.Collections.Concurrent
Imports System.Data.SQLite
Imports System.Text
Imports System.Threading

Module SQLiteOperations

    Private _dbPath As String = ""
    Private _insertThread As Thread = Nothing
    Private _queue As New ConcurrentQueue(Of FileRecord)()
    Private _hasItems As New AutoResetEvent(False)
    Private _stopRequested As Boolean = False
    Private _started As Boolean = False

    ' Tunables
    Private _batchSize As Integer = 20000
    Private _idleFlushMs As Integer = 400
    Private _busyTimeoutMs As Integer = 10000
    Private _optimizeOnStop As Boolean = True

    Private _maintenanceRunning As Integer = 0 ' 0 = false, 1 = true

    ' Backpressure
    Private _maxQueueInMemory As Integer = 1_000_000
    Private _queuedCount As Integer = 0

    ' ---- DTO ----
    Public Class FileRecord
        Public FullPath As String
        Public Size As Long
        Public Ext As String
        Public Metadata As String
        Public Sub New(p As String, s As Long, e As String, Optional m As String = "")
            FullPath = p : Size = s : Ext = e : Metadata = m
        End Sub
    End Class

    ' ---------- Open/Create DB & Start Writer ----------
    Public Sub InitializeSQLite(dbFilePath As String,
                                Optional enableWal As Boolean = True,
                                Optional aggressivePragmas As Boolean = True,
                                Optional batchSize As Integer = 20000,
                                Optional optimizeOnStop As Boolean = False)
        Dim sw As New Stopwatch() : sw.Start()
        If String.IsNullOrWhiteSpace(dbFilePath) Then Throw New ArgumentException("dbFilePath is required.")
        If IsSystemishPath(dbFilePath) Then
            Throw New InvalidOperationException("Please choose a database path under a non-system folder (e.g., Documents).")
        End If

        _dbPath = dbFilePath
        _batchSize = Math.Max(1000, batchSize)
        _optimizeOnStop = optimizeOnStop

        Try
            Form1.databaseName = dbFilePath
            Form1.TextBox_Database.Text = Form1.databaseName
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace)
        End Try

        Dim created As Boolean = False
        If Not IO.File.Exists(_dbPath) Then
            Try
                SQLiteConnection.CreateFile(_dbPath)
                created = True
            Catch ex As Exception
                Logger.Error(ex.Message, ex.StackTrace)
                Throw
            End Try
        End If

        Using conn As SQLiteConnection = OpenConnection(dbFilePath, enableWal, aggressivePragmas, created)
            Logger.Debug($"Open: conn.Open in {sw.ElapsedMilliseconds} ms")
            EnsureFts5Available(conn)
            Logger.Debug($"Open: FTS5 probe in {sw.ElapsedMilliseconds} ms total")
            CreateSchema(conn, created)
            Logger.Debug($"Open: schema check in {sw.ElapsedMilliseconds} ms total")
        End Using

        _stopRequested = False
        _insertThread = New Thread(AddressOf InsertWorker) With {
            .IsBackground = True,
            .Name = "SQLite Writer (files + FTS5)"
        }
        _started = True
        _insertThread.Start()

    End Sub

    Public Sub CloseSQLite()
        StopInsertThread()
        _dbPath = ""
    End Sub

    Public Sub StopInsertThread()
        Try
            Form1.databaseName = "No Database Opened."
            Form1.TextBox_Database.Text = Form1.databaseName
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace)
        End Try

        If Not _started Then Return
        _stopRequested = True
        _hasItems.Set()
        If _insertThread IsNot Nothing Then
            _insertThread.Join()
            _insertThread = Nothing
        End If
        _started = False
        Try
            Using conn As New SQLiteConnection(BuildConnStr())
                conn.Open()
                Using ck As New SQLiteCommand("PRAGMA wal_checkpoint(TRUNCATE);", conn)
                    ck.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
        Logger.Info("Database closed. Index writer stopped.")
    End Sub



    ' ---------- Search ----------
    Public Function SearchParsed(userText As String, Optional keyword2 As String = Nothing,
                                 Optional defaultOperator As String = "AND",
                                 Optional autoPrefixWildcard As Boolean = True) As DataTable
        Dim q = BuildFtsQuery(userText, defaultOperator, autoPrefixWildcard)
        If String.IsNullOrWhiteSpace(q) Then Return New DataTable()
        Return SearchSQLite(q, keyword2)
    End Function

    Public Function SearchSQLite(keyword As String, Optional keyword2 As String = Nothing) As DataTable
        Dim dt As New DataTable()
        If String.IsNullOrWhiteSpace(_dbPath) Then Return dt

        Dim sql As String
        If String.IsNullOrWhiteSpace(keyword2) Then
            sql = "SELECT f.fullpath, f.filename, f.ext, f.size, f.metadata " &
                  "FROM files f JOIN files_fts fts ON f.rowid = fts.rowid " &
                  "WHERE files_fts MATCH @kw;"
        Else
            sql = "SELECT f.fullpath, f.filename, f.ext, f.size, f.metadata " &
                  "FROM files f JOIN files_fts fts ON f.rowid = fts.rowid " &
                  "WHERE files_fts MATCH @kw AND files_fts MATCH @kw2;"
        End If

        Try
            Using conn As New SQLiteConnection(BuildConnStr())
                conn.Open()
                ExecPragma(conn, "busy_timeout", _busyTimeoutMs.ToString())

                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.Add("@kw", DbType.String).Value =
                        String.Format("(fullpath:{0} OR metadata:{0})", keyword)
                    If Not String.IsNullOrWhiteSpace(keyword2) Then
                        cmd.Parameters.Add("@kw2", DbType.String).Value =
                            String.Format("ext:{0}", keyword2.TrimStart("."c))
                    End If
                    Using da As New SQLiteDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace)
            MsgBox("SQLite search error: " & ex.Message, MsgBoxStyle.Exclamation, "SQLite Search")
        End Try
        Return dt
    End Function

    Public Function CountRows() As Long
        Try
            Using conn As New SQLiteConnection(BuildConnStr())
                conn.Open()
                ExecPragma(conn, "busy_timeout", _busyTimeoutMs.ToString())
                Using cmd As New SQLiteCommand("SELECT count(*) FROM files;", conn)
                    Dim v = cmd.ExecuteScalar()
                    If v IsNot Nothing AndAlso Not Convert.IsDBNull(v) Then
                        Return Convert.ToInt64(v)
                    End If
                End Using
            End Using
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace)
        End Try
        Return 0
    End Function



    ' ---------- Update single row metadata ----------
    Public Function UpdateMetadataRow(row As DataRow, newMetadata As String) As Boolean
        If row Is Nothing Then Return False
        If String.IsNullOrWhiteSpace(_dbPath) Then Return False

        Dim fullPath As String = TryCast(row("fullpath"), String)
        If String.IsNullOrWhiteSpace(fullPath) Then Return False

        Dim filename As String = TryCast(row("filename"), String)
        Dim ext As String = TryCast(row("ext"), String)
        Dim sizeVal As Long = 0
        Try
            If Not Convert.IsDBNull(row("size")) Then sizeVal = Convert.ToInt64(row("size"))
        Catch
            sizeVal = 0
        End Try

        Try
            Using conn As New SQLiteConnection(BuildConnStr())
                conn.Open()
                ExecPragma(conn, "busy_timeout", _busyTimeoutMs.ToString())

                Dim sql As String =
"INSERT INTO files(fullpath, filename, ext, size, metadata)
 VALUES(@p,@f,@e,@s, COALESCE(NULLIF(@m,''), (SELECT metadata FROM files WHERE fullpath=@p)))
ON CONFLICT(fullpath) DO UPDATE SET
 filename=excluded.filename,
 ext=excluded.ext,
 size=excluded.size,
 metadata=COALESCE(NULLIF(excluded.metadata,''), files.metadata);"

                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.Add("@p", DbType.String).Value = fullPath
                    cmd.Parameters.Add("@f", DbType.String).Value = If(String.IsNullOrEmpty(filename), IO.Path.GetFileName(fullPath), filename)
                    cmd.Parameters.Add("@e", DbType.String).Value = If(ext, "")
                    cmd.Parameters.Add("@s", DbType.Int64).Value = sizeVal
                    cmd.Parameters.Add("@m", DbType.String).Value = If(newMetadata, "")
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            Return True
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace)
            MsgBox("SQLite metadata update error: " & ex.Message, MsgBoxStyle.Exclamation, "SQLite")
            Return False
        End Try
    End Function

    ' ---------- Schema ----------
    Private Sub CreateSchema(conn As SQLiteConnection, created As Boolean)
        Using cmd As New SQLiteCommand(conn)
            cmd.CommandText =
"CREATE TABLE IF NOT EXISTS files(
  fullpath  TEXT PRIMARY KEY,
  filename  TEXT NOT NULL,
  ext       TEXT,
  size      INTEGER NOT NULL,
  metadata  TEXT
);"
            cmd.ExecuteNonQuery()
        End Using

        Using cmd As New SQLiteCommand(conn)
            cmd.CommandText =
"CREATE VIRTUAL TABLE IF NOT EXISTS files_fts USING fts5(
  fullpath, filename, ext, metadata,
  content='files', content_rowid='rowid',
  tokenize='unicode61', prefix='2 3 4 5 6'
);"
            cmd.ExecuteNonQuery()
        End Using

        ' Triggers mirror files -> files_fts
        Using cmd As New SQLiteCommand(conn)
            cmd.CommandText =
"CREATE TRIGGER IF NOT EXISTS files_ai AFTER INSERT ON files BEGIN
  INSERT INTO files_fts(rowid, fullpath, filename, ext, metadata)
  VALUES (new.rowid, new.fullpath, new.filename, new.ext, new.metadata);
END;"
            cmd.ExecuteNonQuery()
        End Using

        Using cmd As New SQLiteCommand(conn)
            cmd.CommandText =
"CREATE TRIGGER IF NOT EXISTS files_ad AFTER DELETE ON files BEGIN
  INSERT INTO files_fts(files_fts, rowid, fullpath) VALUES('delete', old.rowid, old.fullpath);
END;"
            cmd.ExecuteNonQuery()
        End Using

        Using cmd As New SQLiteCommand(conn)
            cmd.CommandText =
"CREATE TRIGGER IF NOT EXISTS files_au AFTER UPDATE ON files BEGIN
  INSERT INTO files_fts(files_fts, rowid, fullpath) VALUES('delete', old.rowid, old.fullpath);
  INSERT INTO files_fts(rowid, fullpath, filename, ext, metadata)
  VALUES (new.rowid, new.fullpath, new.filename, new.ext, new.metadata);
END;"
            cmd.ExecuteNonQuery()
        End Using

        ' Index for range scans
        Using cmd As New SQLiteCommand("CREATE INDEX IF NOT EXISTS idx_files_fullpath ON files(fullpath);", conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub


    ' ---------- Ensure FTS5 ----------
    Private Sub EnsureFts5Available(conn As SQLiteConnection)
        Dim hasFts5 As Boolean = False
        Try
            Using cmd As New SQLiteCommand("PRAGMA compile_options;", conn)
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim opt As String = r.GetString(0)
                        If opt.IndexOf("ENABLE_FTS5", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            hasFts5 = True : Exit While
                        End If
                    End While
                End Using
            End Using
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace)
        End Try

        If Not hasFts5 Then
            Try
                Using cmd As New SQLiteCommand("CREATE VIRTUAL TABLE temp._fts5probe USING fts5(x); DROP TABLE temp._fts5probe;", conn)
                    cmd.ExecuteNonQuery()
                End Using
                hasFts5 = True
            Catch ex As Exception
                Logger.Error(ex.Message, ex.StackTrace)
                Throw New InvalidOperationException(
                    "This SQLite build does not include FTS5 (no such module: fts5). " &
                    "Install System.Data.SQLite with ENABLE_FTS5 and match x86/x64.", ex)
            End Try
        End If
    End Sub

    Private Function BuildConnStr() As String
        Return "Data Source=" & _dbPath & ";Version=3;Pooling=True;BinaryGUID=False;Default IsolationLevel=Serializable;"
    End Function

    Private Function OpenConnection(dbFilePath As String, enableWal As Boolean, aggressivePragmas As Boolean, created As Boolean) As SQLiteConnection
        _dbPath = dbFilePath
        Dim conn As New SQLiteConnection(BuildConnStr())
        conn.Open()
        ExecPragma(conn, "busy_timeout", _busyTimeoutMs.ToString())
        If enableWal Then
            ExecPragma(conn, "journal_mode", "WAL")
            ExecPragma(conn, "wal_autocheckpoint", "25600") ' bumped from 10000
        End If
        ExecPragma(conn, "synchronous", "NORMAL")

        'check for large wal file and clean it
        Dim walPath = _dbPath & "-wal"
        If IO.File.Exists(walPath) AndAlso (New IO.FileInfo(walPath)).Length > (200L * 1024 * 1024) Then
            Using ck As New SQLiteCommand("PRAGMA wal_checkpoint(RESTART); PRAGMA wal_checkpoint(TRUNCATE);", conn)
                ck.ExecuteNonQuery()
            End Using
        Else
            Using ck As New SQLiteCommand("PRAGMA wal_checkpoint(PASSIVE);", conn)
                ck.ExecuteNonQuery()
            End Using
        End If


        If aggressivePragmas Then
            ExecPragma(conn, "temp_store", "MEMORY")
            If created Then ExecPragma(conn, "page_size", "4096")
        End If
        Return conn
    End Function



    'checkpoint the wal file occasionally
    Private Sub MaybeCheckpointBySize(conn As SQLiteConnection, maxBytes As Long)
        Dim walPath = _dbPath & "-wal"
        Dim len As Long = 0
        If IO.File.Exists(walPath) Then len = (New IO.FileInfo(walPath)).Length
        If len >= maxBytes Then
            Using ck As New SQLiteCommand("PRAGMA wal_checkpoint(RESTART);", conn)
                ck.ExecuteNonQuery()
            End Using
            Using ck2 As New SQLiteCommand("PRAGMA wal_checkpoint(TRUNCATE);", conn)
                ck2.ExecuteNonQuery()
            End Using
        Else
            Using ck As New SQLiteCommand("PRAGMA wal_checkpoint(PASSIVE);", conn)
                ck.ExecuteNonQuery()
            End Using
        End If
    End Sub



    Private Sub ExecPragma(conn As SQLiteConnection, name As String, value As String)
        Using cmd As New SQLiteCommand("PRAGMA " & name & "=" & value & ";", conn)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                Logger.Error(" PRAGMA " & name & "=" & value & " failed: " & ex.Message, ex.StackTrace)
            End Try
        End Using
    End Sub

    ' ---------- Writer Thread ----------
    Private Sub InsertWorker()
        Using conn As New SQLiteConnection(BuildConnStr())
            conn.Open()
            ExecPragma(conn, "busy_timeout", _busyTimeoutMs.ToString())
            Using tmp As New SQLiteCommand("PRAGMA journal_mode;", conn)
                Try : tmp.ExecuteScalar() : Catch : End Try
            End Using
            ExecPragma(conn, "synchronous", "NORMAL")
            ExecPragma(conn, "temp_store", "MEMORY")

            Dim sqlUpsert As String =
"INSERT INTO files(fullpath, filename, ext, size, metadata)
 VALUES(@p,@f,@e,@s, COALESCE(NULLIF(@m,''), (SELECT metadata FROM files WHERE fullpath=@p)))
ON CONFLICT(fullpath) DO UPDATE SET
 filename=excluded.filename,
 ext=excluded.ext,
 size=excluded.size,
 metadata=COALESCE(NULLIF(excluded.metadata,''), files.metadata);"

            Using cmdIns As New SQLiteCommand(sqlUpsert, conn)
                Dim pP = cmdIns.Parameters.Add("@p", DbType.String)
                Dim pF = cmdIns.Parameters.Add("@f", DbType.String)
                Dim pE = cmdIns.Parameters.Add("@e", DbType.String)
                Dim pS = cmdIns.Parameters.Add("@s", DbType.Int64)
                Dim pM = cmdIns.Parameters.Add("@m", DbType.String)

                While True
                    Dim batch As New List(Of FileRecord)(_batchSize)
                    Dim item As FileRecord = Nothing

                    If _queue.IsEmpty AndAlso Not _stopRequested Then
                        _hasItems.WaitOne(_idleFlushMs)
                    End If

                    While batch.Count < _batchSize AndAlso _queue.TryDequeue(item)
                        batch.Add(item)
                        Interlocked.Decrement(_queuedCount)
                    End While

                    If batch.Count > 0 Then
                        Using tx = conn.BeginTransaction()
                            cmdIns.Transaction = tx
                            For Each r In batch
                                pP.Value = r.FullPath
                                pF.Value = IO.Path.GetFileName(r.FullPath)
                                pE.Value = If(r.Ext, "")
                                pS.Value = r.Size
                                pM.Value = If(r.Metadata, "")
                                cmdIns.ExecuteNonQuery()
                            Next
                            tx.Commit()
                            cmdIns.Transaction = Nothing
                            MaybeCheckpointBySize(conn, 100L * 1024 * 1024)  ' 100 MB cap
                        End Using
                    ElseIf _stopRequested Then
                        Exit While
                    End If
                End While

                If _optimizeOnStop Then
                    Try
                        Using opt As New SQLiteCommand("INSERT INTO files_fts(files_fts) VALUES('optimize');", conn)
                            opt.CommandTimeout = 0
                            opt.ExecuteNonQuery()
                        End Using
                    Catch
                    End Try
                End If
            End Using
        End Using
    End Sub

    ' ---------- Tokenizer / Query builder (unchanged) ----------
    Public Function BuildFtsQuery(userText As String,
                                  Optional defaultOperator As String = "AND",
                                  Optional autoPrefixWildcard As Boolean = True) As String
        If String.IsNullOrWhiteSpace(userText) Then Return ""

        Dim op = If(String.Equals(defaultOperator, "OR", StringComparison.OrdinalIgnoreCase), "OR", "AND")
        Dim tokens = Tokenize(userText)
        If tokens.Count = 0 Then Return ""

        Dim sb As New StringBuilder()
        Dim needOp As Boolean = False

        For Each t In tokens
            Select Case t.Kind
                Case TokKind.OperatorAnd
                    sb.Append(" AND ") : needOp = False
                Case TokKind.OperatorOr
                    sb.Append(" OR ") : needOp = False
                Case TokKind.OperatorNot
                    If needOp Then sb.Append(" ").Append(op).Append(" ")
                    sb.Append("NOT ") : needOp = False

                Case TokKind.Phrase
                    If needOp Then sb.Append(" ").Append(op).Append(" ")
                    sb.Append("""").Append(EscapeQuotes(t.Text)).Append("""") : needOp = True

                Case TokKind.Field
                    If needOp Then sb.Append(" ").Append(op).Append(" ")
                    Dim fld = t.Field.ToLowerInvariant()
                    Dim val = t.Text
                    Dim hasStar = val.IndexOf("*"c) >= 0
                    Dim isPhrase = val.StartsWith("""") AndAlso val.EndsWith("""")
                    If Not isPhrase AndAlso autoPrefixWildcard AndAlso Not hasStar AndAlso val.Length >= 3 Then val &= "*"
                    If Not val.Contains("*"c) AndAlso NeedsQuoting(val) Then val = """" & EscapeQuotes(val) & """"
                    sb.Append(fld).Append(":").Append(val) : needOp = True

                Case TokKind.Word
                    If needOp Then sb.Append(" ").Append(op).Append(" ")
                    Dim val = t.Text
                    If autoPrefixWildcard AndAlso Not val.Contains("*"c) AndAlso val.Length >= 3 Then val &= "*"
                    If Not val.Contains("*"c) AndAlso NeedsQuoting(val) Then val = """" & EscapeQuotes(val) & """"
                    sb.Append(val) : needOp = True

                Case TokKind.NegatedWord
                    If needOp Then sb.Append(" ").Append(op).Append(" ")
                    Dim val = t.Text
                    If autoPrefixWildcard AndAlso Not val.Contains("*"c) AndAlso val.Length >= 3 Then val &= "*"
                    If Not val.Contains("*"c) AndAlso NeedsQuoting(val) Then val = """" & EscapeQuotes(val) & """"
                    sb.Append("NOT ").Append(val) : needOp = True
            End Select
        Next

        Return sb.ToString()
    End Function

    Private Enum TokKind
        Word
        Phrase
        NegatedWord
        Field
        OperatorAnd
        OperatorOr
        OperatorNot
    End Enum

    Private Class Tok
        Public Kind As TokKind
        Public Text As String
        Public Field As String
    End Class

    Private Function Tokenize(s As String) As List(Of Tok)
        Dim list As New List(Of Tok)()
        Dim i As Integer = 0, n As Integer = s.Length

        While i < n
            While i < n AndAlso Char.IsWhiteSpace(s(i)) : i += 1 : End While
            If i >= n Then Exit While

            If s(i) = """"c Then
                i += 1
                Dim start = i
                While i < n AndAlso s(i) <> """"c : i += 1 : End While
                Dim phrase = s.Substring(start, i - start)
                list.Add(New Tok With {.Kind = TokKind.Phrase, .Text = phrase})
                If i < n Then i += 1
            Else
                Dim start = i
                While i < n AndAlso Not Char.IsWhiteSpace(s(i)) : i += 1 : End While
                Dim raw = s.Substring(start, i - start)

                If raw.Equals("AND", StringComparison.OrdinalIgnoreCase) Then
                    list.Add(New Tok With {.Kind = TokKind.OperatorAnd}) : Continue While
                ElseIf raw.Equals("OR", StringComparison.OrdinalIgnoreCase) Then
                    list.Add(New Tok With {.Kind = TokKind.OperatorOr}) : Continue While
                ElseIf raw.Equals("NOT", StringComparison.OrdinalIgnoreCase) Then
                    list.Add(New Tok With {.Kind = TokKind.OperatorNot}) : Continue While
                End If

                Dim idx = raw.IndexOf(":"c)
                If idx > 0 Then
                    Dim fld = raw.Substring(0, idx)
                    Dim val = raw.Substring(idx + 1)
                    Select Case fld.ToLowerInvariant()
                        Case "ext", "filename", "fullpath", "metadata"
                            list.Add(New Tok With {.Kind = TokKind.Field, .Field = fld, .Text = val})
                            Continue While
                    End Select
                End If

                If raw.StartsWith("-", StringComparison.Ordinal) AndAlso raw.Length > 1 Then
                    list.Add(New Tok With {.Kind = TokKind.NegatedWord, .Text = raw.Substring(1)})
                Else
                    list.Add(New Tok With {.Kind = TokKind.Word, .Text = raw})
                End If
            End If
        End While

        Return list
    End Function

    Private Function NeedsQuoting(token As String) As Boolean
        For Each ch In token
            If Char.IsWhiteSpace(ch) OrElse ch = ":"c OrElse ch = """"c Then Return True
        Next
        Return False
    End Function

    Private Function EscapeQuotes(s As String) As String
        Return s.Replace("""", """""")
    End Function

    ' ---------- Misc helpers ----------
    Private Function IsSystemishPath(p As String) As Boolean
        Try
            Dim full = IO.Path.GetFullPath(p)
            Dim sys = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
            Dim pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
            Dim pfx = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
            Dim pgd = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
            If Not String.IsNullOrEmpty(sys) AndAlso full.StartsWith(sys, StringComparison.OrdinalIgnoreCase) Then Return True
            If Not String.IsNullOrEmpty(pf) AndAlso full.StartsWith(pf, StringComparison.OrdinalIgnoreCase) Then Return True
            If Not String.IsNullOrEmpty(pfx) AndAlso full.StartsWith(pfx, StringComparison.OrdinalIgnoreCase) Then Return True
            If Not String.IsNullOrEmpty(pgd) AndAlso full.StartsWith(pgd, StringComparison.OrdinalIgnoreCase) Then Return True
        Catch
        End Try
        Return False
    End Function


    'maintenance / rebuild index stuff

    Public Sub RebuildIndex(Optional doVacuum As Boolean = False, Optional doOptimize As Boolean = True)
        If String.IsNullOrWhiteSpace(_dbPath) Then Exit Sub

        StatusProgress.ShowPopup(status:="Rebuilding index... this may take a while", indeterminate:=True)

        Dim wasRunning = _started
        If wasRunning Then StopInsertThread()

        Try
            Using conn As New SQLiteConnection(BuildConnStr())
                conn.Open()
                ExecPragma(conn, "busy_timeout", _busyTimeoutMs.ToString())

                Using tx = conn.BeginTransaction()
                    Using rb As New SQLiteCommand("INSERT INTO files_fts(files_fts) VALUES('rebuild');", conn, tx)
                        rb.CommandTimeout = 0
                        rb.ExecuteNonQuery()
                    End Using
                    tx.Commit()
                End Using

                If doOptimize Then
                    Using opt As New SQLiteCommand("INSERT INTO files_fts(files_fts) VALUES('optimize');", conn)
                        opt.CommandTimeout = 0
                        opt.ExecuteNonQuery()
                    End Using
                End If

                ' Light, fast housekeeping
                Using ck As New SQLiteCommand("PRAGMA wal_checkpoint(PASSIVE);", conn)
                    ck.ExecuteNonQuery()
                End Using

                If doVacuum Then
                    Using vac As New SQLiteCommand("VACUUM;", conn)
                        vac.CommandTimeout = 0
                        vac.ExecuteNonQuery()
                    End Using
                End If
            End Using

            Logger.Info("FTS rebuild complete.")
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace)
            MsgBox("RebuildIndex error: " & ex.Message, MsgBoxStyle.Critical, "SQLite")
        Finally
            If wasRunning Then
                Try
                    InitializeSQLite(_dbPath)
                Catch ex As Exception
                    Logger.Error(ex.Message, ex.StackTrace)
                End Try
            End If
            StatusProgress.ClosePopup()
        End Try
    End Sub



    Public Sub OptimizeIndex()
        If String.IsNullOrWhiteSpace(_dbPath) Then
            Logger.Error("Error: No database opened for optimization.")
            Exit Sub
        End If

        ' Re-entry guard
        If Threading.Interlocked.Exchange(_maintenanceRunning, 1) = 1 Then
            Logger.Info("An optimization/maintenance task is already running.")
            Exit Sub
        End If

        ' Show popup on UI (uses thread-safe helpers inside StatusProgress)
        StatusProgress.ShowPopup(status:="Optimizing Index... this may take a while", indeterminate:=True)

        ' Kick work to a background task so the UI stays responsive
        Task.Run(Sub()
                     Try
                         Using conn As New SQLite.SQLiteConnection(BuildConnStr())
                             conn.Open()
                             ExecPragma(conn, "busy_timeout", _busyTimeoutMs.ToString())

                             Using cmd As New SQLite.SQLiteCommand("INSERT INTO files_fts(files_fts) VALUES('optimize');", conn)
                                 cmd.CommandTimeout = 0
                                 cmd.ExecuteNonQuery()
                             End Using
                         End Using
                         Logger.Info("FTS index optimized.")
                     Catch ex As Exception
                         Logger.Error(ex.Message, ex.StackTrace)
                         MsgBox("OptimizeIndex error: " & ex.Message, MsgBoxStyle.Exclamation, "SQLite")
                     Finally
                         ' Close the popup and clear the guard, from any thread
                         Try : StatusProgress.ClosePopup() : Catch : End Try
                         Threading.Interlocked.Exchange(_maintenanceRunning, 0)
                     End Try
                 End Sub)
    End Sub




End Module
