Imports System.Data.SQLite

Public NotInheritable Class SQLMaintenance
    Private Sub New()
    End Sub

    ' Conservative thresholds (used only when we can actually measure)
    Private Const SEGMENT_THRESHOLD As Integer = 150
    Private Const LEVEL_THRESHOLD As Integer = 5

    Public Structure FtsHealth
        Public Segments As Integer?   ' null when unknown
        Public MaxLevel As Integer?   ' null when unknown
        Public HasSize As Boolean
        Public FtsBytes As Long
        Public Overrides Function ToString() As String
            Dim segTxt = If(Segments.HasValue, Segments.Value.ToString(), "?")
            Dim lvlTxt = If(MaxLevel.HasValue, MaxLevel.Value.ToString(), "?")
            If HasSize Then
                Return $"seg={segTxt}, L{lvlTxt}, bytes={FtsBytes:N0}"
            Else
                Return $"seg={segTxt}, L{lvlTxt}"
            End If
        End Function
    End Structure

    Private Shared Function TableExists(conn As SQLiteConnection, tbl As String) As Boolean
        Using cmd As New SQLiteCommand("SELECT 1 FROM sqlite_master WHERE type IN ('table','view') AND name=@n LIMIT 1;", conn)
            cmd.Parameters.AddWithValue("@n", tbl)
            Dim o = cmd.ExecuteScalar()
            Return o IsNot Nothing
        End Using
    End Function




    ''' Try to read FTS fragmentation signals. Supports FTS4 (_segdir) and FTS5 (best-effort).
    ''' Never throws: unknown bits are returned as Null/False.
    Public Shared Function GetFtsHealth(conn As SQLiteConnection, ftsName As String) As FtsHealth
        Dim h As New FtsHealth With {.Segments = Nothing, .MaxLevel = Nothing, .HasSize = False, .FtsBytes = 0}

        ' ---- Size via dbstat (optional but useful)
        Try
            Using cmd As New SQLiteCommand("SELECT SUM(pgsize) FROM dbstat WHERE name=@n OR name LIKE @nlike;", conn)
                cmd.Parameters.AddWithValue("@n", ftsName)
                cmd.Parameters.AddWithValue("@nlike", ftsName & "\_%")
                Dim o = cmd.ExecuteScalar()
                If o IsNot Nothing AndAlso o IsNot DBNull.Value Then
                    h.FtsBytes = CLng(o)
                    h.HasSize = True
                End If
            End Using
        Catch
            ' dbstat not compiled/enabled — ignore
        End Try

        ' ---- Fragmentation / depth
        ' FTS4 has <fts>_segdir with columns (level, idx, start_block, leaves_end_block, end_block)
        Dim segdir As String = ftsName & "_segdir"
        If TableExists(conn, segdir) Then
            Using cmd As New SQLiteCommand($"SELECT COUNT(*), IFNULL(MAX(level),0) FROM {segdir};", conn)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        h.Segments = If(r.IsDBNull(0), CType(Nothing, Integer?), r.GetInt32(0))
                        h.MaxLevel = If(r.IsDBNull(1), CType(Nothing, Integer?), r.GetInt32(1))
                    End If
                End Using
            End Using
            Return h
        End If

        ' FTS5 does NOT expose _segdir.
        ' Heuristic 1 (cheap): count rows in <fts>_data as a proxy for # of segment pages.
        ' It's imperfect but correlates with index fragmentation and stays small cost.
        Dim dataTbl As String = ftsName & "_data"
        If TableExists(conn, dataTbl) Then
            Try
                Using cmd As New SQLiteCommand($"SELECT COUNT(*) FROM {dataTbl};", conn)
                    Dim pages As Long = CLng(cmd.ExecuteScalar())
                    ' Map the rough page count to a conservative "segments" proxy.
                    ' (FTS5 stores many leaf/internal pages; treat every 1,000 pages as ~1 "segment" unit.)
                    Dim segApprox As Integer = CInt(Math.Max(0, Math.Min(Integer.MaxValue, pages \ 1000)))
                    h.Segments = segApprox
                    h.MaxLevel = Nothing ' unknown for FTS5 without decoding rowids
                End Using
            Catch
                ' ignore
            End Try
        End If

        Return h
    End Function

    ''' Conservative: only return True when clearly fragmented.
    ''' If we can't measure reliably, return False (no nagging).
    Public Shared Function ShouldOptimize(conn As SQLiteConnection, ftsName As String) As Boolean
        Dim h = GetFtsHealth(conn, ftsName)

        If h.Segments.HasValue AndAlso h.Segments.Value >= SEGMENT_THRESHOLD Then Return True
        If h.MaxLevel.HasValue AndAlso h.MaxLevel.Value >= LEVEL_THRESHOLD Then Return True

        ' Optional size gate (disabled unless you want size-based nudges)
        ' If h.HasSize AndAlso h.FtsBytes >= (1L << 31) Then Return True ' ~2 GB

        Return False
    End Function

    Public Shared Function CheckDatabaseHealth(dbPath As String, ftsName As String) As Boolean
        If String.IsNullOrWhiteSpace(dbPath) Then Return False
        Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;Pooling=True;BinaryGUID=False;")
            conn.Open()
            Return ShouldOptimize(conn, ftsName)
        End Using
    End Function
End Class
