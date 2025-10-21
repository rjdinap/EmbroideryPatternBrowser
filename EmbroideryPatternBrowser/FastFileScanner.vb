Imports System.Data.SQLite
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading

Public Module FastFileScanner

    ' ---------- Public result object ----------
    Public Structure ScanResult
        Public FilesScanned As Long
        Public DirsScanned As Long
        Public FilesAdded As Long
        Public FilesUpdated As Long
        Public FilesDeleted As Long
        Public ElapsedMs As Long

        Public Overrides Function ToString() As String
            Return String.Format("Dirs: {0}, Files scanned: {1}, Added: {2}, Updated: {3}, Deleted: {4}, Elapsed: {5} ms",
                                 DirsScanned, FilesScanned, FilesAdded, FilesUpdated, FilesDeleted, ElapsedMs)
        End Function
    End Structure

    ' ---------- File attributes ----------
    Private Const FILE_ATTRIBUTE_DIRECTORY As Integer = &H10
    Private Const FILE_ATTRIBUTE_REPARSE_POINT As Integer = &H400
    Private Const FILE_ATTRIBUTE_HIDDEN As Integer = &H2
    Private Const FILE_ATTRIBUTE_SYSTEM As Integer = &H4

    ' ---------- FindFirstFileEx ----------
    Private Enum FINDEX_INFO_LEVELS
        FindExInfoStandard = 0
        FindExInfoBasic = 1
    End Enum

    Private Enum FINDEX_SEARCH_OPS
        FindExSearchNameMatch = 0
        FindExSearchLimitToDirectories = 1
        FindExSearchLimitToDevices = 2
    End Enum

    Private Const FIND_FIRST_EX_LARGE_FETCH As UInteger = &H2UI
    Private ReadOnly INVALID_HANDLE_VALUE As IntPtr = New IntPtr(-1)

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure WIN32_FIND_DATA
        Public dwFileAttributes As Integer
        Public ftCreationTime As ComTypes.FILETIME
        Public ftLastAccessTime As ComTypes.FILETIME
        Public ftLastWriteTime As ComTypes.FILETIME
        Public nFileSizeHigh As Integer
        Public nFileSizeLow As Integer
        Public dwReserved0 As Integer
        Public dwReserved1 As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public cFileName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)>
        Public cAlternate As String
    End Structure

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True, EntryPoint:="FindFirstFileExW")>
    Private Function FindFirstFileEx(lpFileName As String,
                                     fInfoLevelId As FINDEX_INFO_LEVELS,
                                     <Out> ByRef lpFindFileData As WIN32_FIND_DATA,
                                     fSearchOp As FINDEX_SEARCH_OPS,
                                     lpSearchFilter As IntPtr,
                                     dwAdditionalFlags As UInteger) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True, EntryPoint:="FindNextFileW")>
    Private Function FindNextFile(hFindFile As IntPtr,
                                  <Out> ByRef lpFindFileData As WIN32_FIND_DATA) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function FindClose(hFindFile As IntPtr) As Boolean
    End Function


    ' Set once at startup: FastFileScanner.ReportStatus = AddressOf Form1.Status
    Public WriteOnly Property ReportStatus As Action(Of String)
        Set(value As Action(Of String))
            _reportStatus = value
        End Set
    End Property
    Private _reportStatus As Action(Of String)

    Private Sub Status(msg As String)
        Dim cb = _reportStatus
        If cb IsNot Nothing Then
            cb(msg)
        End If
    End Sub


    ' ---------- Public scan entry point ----------
    ''' <summary>
    ''' Scans 1+ roots, reconciles with DB (files table), returns detailed stats.
    ''' Optimizations kept:
    '''  1) Prefetch existing DB rows once per root
    '''  2) Single large transaction per root
    ''' </summary>
    Public Function ScanAll(rootPaths As List(Of String),
                            indexDbPath As String,
                            progressCallback As Action(Of String, Integer, Integer),
                            Optional cancel As CancellationToken = Nothing) As ScanResult

        Dim stats As New ScanResult()
        Dim swTotal As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()
        Dim swDirs As New Diagnostics.Stopwatch()
        Dim swReconcile As New Diagnostics.Stopwatch()
        Dim zp As New ZipProcessing

        If rootPaths Is Nothing OrElse rootPaths.Count = 0 Then
            Throw New ArgumentException("rootPaths must contain at least one path.")
        End If
        If String.IsNullOrWhiteSpace(indexDbPath) Then
            Throw New ArgumentException("indexDbPath is required.")
        End If

        Try
            Form1.Status("Note: for safety and speed, system folders like \Windows and \Program Files are not scanned.")
        Catch
        End Try

        ' PASS 1: collect directories
        swDirs.Start()
        Dim allDirs As New List(Of String)(900000)
        Dim dirCounter As Integer = 0
        Dim totalRoots As Integer = rootPaths.Count

        Dim skipRoots = BuildDefaultSkipRoots()
        Dim skipNames = BuildDefaultSkipNameSet()
        Dim skipContains = BuildDefaultSkipContains()

        'iterative
        For Each root In rootPaths
            If cancel.IsCancellationRequested Then Exit For
            CollectAllDirectoriesIterative(root, allDirs, progressCallback, dirCounter, totalRoots, cancel,
            includeHiddenAndSystem:=False,
            skipRoots:=skipRoots,
            skipNameSet:=skipNames,
            skipContainsFragments:=skipContains)
        Next



        swDirs.Stop()
        stats.DirsScanned = allDirs.Count
        progressCallback?.Invoke("Directory enumeration complete", CInt(stats.DirsScanned), CInt(stats.DirsScanned))

        ' PASS 2: per-root reconcile
        swReconcile.Start()

        Dim exts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
    ".pes", ".hus", ".vp3", ".xxx", ".dst", ".jef", ".sew", ".vip", ".exp", ".pec"
}
        If My.Settings.IncludeZipFilesInScans Then exts.Add(".zip")

        Using conn As New SQLiteConnection("Data Source=" & indexDbPath & ";Version=3;Pooling=True;BinaryGUID=False;")
            conn.Open()
            ExecBusyTimeout(conn, 10000)

            ' Bulk-friendly but safe PRAGMAs (no WAL changes here)
            Using cmd As New SQLiteCommand("PRAGMA synchronous=OFF; PRAGMA temp_store=MEMORY; PRAGMA cache_size=-131072; PRAGMA mmap_size=268435456;", conn)
                Try : cmd.ExecuteNonQuery() : Catch : End Try
            End Using

            ' Group directories by root so we can do one big TX and one prefetch per root
            Dim rootsNorm = rootPaths.Select(Function(r) r.TrimEnd("\"c)).ToList()

            For Each root In rootsNorm
                If cancel.IsCancellationRequested Then Exit For

                ' Collect only the dirs for this root
                Dim rootPrefix As String = root.TrimEnd("\"c) & "\"
                Dim theseDirs = allDirs.Where(Function(d) d.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase) OrElse String.Equals(d.TrimEnd("\"c), root, StringComparison.OrdinalIgnoreCase)).ToList()

                ' 1) Prefetch existing rows once per ROOT
                Dim swSel As New Diagnostics.Stopwatch()
                swSel.Start()
                Dim existing As Dictionary(Of String, (Size As Long, Metadata As String)) = LoadExistingForRoot(conn, rootPrefix)
                swSel.Stop()

                ' 2) One large transaction per ROOT
                Using tx = conn.BeginTransaction()
                    Dim upsertSql As String =
"INSERT INTO files(fullpath, filename, ext, size, metadata)
 VALUES(@p,@f,@e,@s, COALESCE(NULLIF(@m,''), (SELECT metadata FROM files WHERE fullpath=@p)))
ON CONFLICT(fullpath) DO UPDATE SET
 filename=excluded.filename,
 ext=excluded.ext,
 size=excluded.size,
 metadata=COALESCE(NULLIF(excluded.metadata,''), files.metadata);"

                    Using cmdUpsert As New SQLiteCommand(upsertSql, conn, tx),
                          cmdDelete As New SQLiteCommand("DELETE FROM files WHERE fullpath=@p;", conn, tx)

                        Dim pP = cmdUpsert.Parameters.Add("@p", System.Data.DbType.String)
                        Dim pF = cmdUpsert.Parameters.Add("@f", System.Data.DbType.String)
                        Dim pE = cmdUpsert.Parameters.Add("@e", System.Data.DbType.String)
                        Dim pS = cmdUpsert.Parameters.Add("@s", System.Data.DbType.Int64)
                        Dim pM = cmdUpsert.Parameters.Add("@m", System.Data.DbType.String)

                        Dim dP = cmdDelete.Parameters.Add("@p", System.Data.DbType.String)

                        Dim swEnum As New Diagnostics.Stopwatch()
                        Dim swWrite As New Diagnostics.Stopwatch()
                        Dim perRootFilesFound As Integer = 0
                        Dim perRootAdd As Integer = 0
                        Dim perRootUpd As Integer = 0

                        ' Enumerate & process per directory (but write under same TX)
                        For i As Integer = 0 To theseDirs.Count - 1
                            If cancel.IsCancellationRequested Then Exit For
                            Dim dir = theseDirs(i)

                            ' 2a) Enumerate this directory (only files with desired exts)
                            swEnum.Start()
                            Dim found As New List(Of (FullPath As String, Size As Long, Ext As String))()
                            EnumerateFilesFast(dir, exts,
                               Sub(fullPath As String, size As Long, ext As String)
                                   found.Add((fullPath, size, ext))
                               End Sub)
                            swEnum.Stop()

                            If found.Count = 0 Then
                                Continue For
                            End If

                            ' 2b) Reconcile against prefetch map (no SELECTs here)
                            swWrite.Start()
                            For Each f In found
                                Dim storedPath As String = ""
                                Dim fileName As String = ""
                                Dim prev As (Size As Long, Metadata As String) = Nothing
                                Dim hadPrev As Boolean
                                If (f.Ext.ToLower = "zip") Then
                                    'zip files are special case
                                    'Status("Processing zip: " & f.FullPath)
                                    Dim zipFound = zp.ScanZipForFileNames(f.FullPath)
                                    For Each z In zipFound
                                        storedPath = StripLongPathPrefix(z.FullPath)
                                        fileName = Path.GetFileName(storedPath)
                                        prev = Nothing
                                        hadPrev = existing.TryGetValue(storedPath, prev)

                                        If hadPrev Then
                                            existing.Remove(storedPath) ' mark seen
                                            If prev.Size <> z.Size Then
                                                ' UPDATE
                                                pP.Value = storedPath
                                                pF.Value = fileName
                                                pE.Value = z.Ext.TrimStart("."c)
                                                pS.Value = z.Size
                                                pM.Value = ""        ' preserve old metadata
                                                cmdUpsert.ExecuteNonQuery()
                                                perRootUpd += 1
                                            Else
                                                ' unchanged
                                            End If
                                        Else
                                            ' INSERT
                                            pP.Value = storedPath
                                            pF.Value = fileName
                                            pE.Value = z.Ext.TrimStart("."c)
                                            pS.Value = z.Size
                                            pM.Value = ""
                                            cmdUpsert.ExecuteNonQuery()
                                            perRootAdd += 1
                                        End If
                                        perRootFilesFound += 1 ' count files that we find in the zip file
                                    Next
                                    Continue For 'don't process the zip file itself in the outer loop
                                End If

                                storedPath = StripLongPathPrefix(f.FullPath)
                                fileName = Path.GetFileName(storedPath)
                                prev = Nothing
                                hadPrev = existing.TryGetValue(storedPath, prev)

                                If hadPrev Then
                                    existing.Remove(storedPath) ' mark seen
                                    If prev.Size <> f.Size Then
                                        ' UPDATE
                                        pP.Value = storedPath
                                        pF.Value = fileName
                                        pE.Value = f.Ext
                                        pS.Value = f.Size
                                        pM.Value = ""        ' preserve old metadata
                                        cmdUpsert.ExecuteNonQuery()
                                        perRootUpd += 1
                                    Else
                                        ' unchanged
                                    End If
                                Else
                                    ' INSERT
                                    pP.Value = storedPath
                                    pF.Value = fileName
                                    pE.Value = f.Ext
                                    pS.Value = f.Size
                                    pM.Value = ""
                                    cmdUpsert.ExecuteNonQuery()
                                    perRootAdd += 1
                                End If

                                perRootFilesFound += 1
                            Next ' end of the file reconcile loop

                            swWrite.Stop()

                            If (i Mod 200) = 0 Then
                                progressCallback?.Invoke("Reconciling…", i, theseDirs.Count)
                            End If
                        Next

                        ' 2c) DELETE any leftovers for this root
                        Dim delCount As Integer = 0
                        If existing.Count > 0 Then
                            For Each kv In existing
                                dP.Value = kv.Key
                                cmdDelete.ExecuteNonQuery()
                                delCount += 1
                            Next
                        End If

                        ' root-level tallies
                        stats.FilesScanned += perRootFilesFound
                        stats.FilesAdded += perRootAdd
                        stats.FilesUpdated += perRootUpd
                        stats.FilesDeleted += delCount

                        ' Emit per-root summary
                        Form1.Status(String.Format("Perf(root): {0}  sel={1} ms, enum={2} ms, write={3} ms, found={4}, add={5}, upd={6}, del={7}",
                                                   root,
                                                   swSel.ElapsedMilliseconds,
                                                   swEnum.ElapsedMilliseconds,
                                                   swWrite.ElapsedMilliseconds,
                                                   perRootFilesFound, perRootAdd, perRootUpd, delCount))
                    End Using

                    tx.Commit()
                End Using
            Next
        End Using

        swReconcile.Stop()
        swTotal.Stop()
        stats.ElapsedMs = swTotal.ElapsedMilliseconds

        ' Overall timing line
        Form1.Status(String.Format("Timing: dirs={0} ms, reconcile={1} ms, total={2} ms",
                                   swDirs.ElapsedMilliseconds, swReconcile.ElapsedMilliseconds, swTotal.ElapsedMilliseconds))

        Try : StatusProgress.ClosePopup() : Catch : End Try
        Return stats
    End Function

    ' ---------- Prefetch once per ROOT (range scan uses index on files(fullpath)) ----------
    Private Function LoadExistingForRoot(conn As SQLiteConnection,
                                         rootPrefix As String) As Dictionary(Of String, (Size As Long, Metadata As String))
        Dim map As New Dictionary(Of String, (Long, String))(StringComparer.OrdinalIgnoreCase)
        If String.IsNullOrWhiteSpace(rootPrefix) Then Return map

        ' Normalize as canonical prefix "G:\misc1\" (must end with "\" to avoid sibling hits)
        Dim lo As String = StripLongPathPrefix(rootPrefix.TrimEnd("\"c) & "\")
        ' High bound trick: append a sentinel that sorts after all valid path chars
        Dim hi As String = lo & "~~"

        Using c As New SQLiteCommand("
            SELECT fullpath, size, metadata
            FROM files
            WHERE fullpath >= @lo AND fullpath < @hi;", conn)
            c.Parameters.Add("@lo", System.Data.DbType.String).Value = lo
            c.Parameters.Add("@hi", System.Data.DbType.String).Value = hi
            Using r = c.ExecuteReader()
                While r.Read()
                    Dim p As String = If(r.IsDBNull(0), "", r.GetString(0))
                    Dim s As Long = If(r.IsDBNull(1), 0L, CLng(r.GetValue(1)))
                    Dim m As String = If(r.IsDBNull(2), "", r.GetString(2))
                    If Not String.IsNullOrEmpty(p) Then
                        If Not map.ContainsKey(p) Then map.Add(p, (s, m))
                    End If
                End While
            End Using
        End Using

        Return map
    End Function

    ' ---------- Directory enumeration ----------
    Private Sub CollectAllDirectoriesIterative(root As String,
                                           allDirs As List(Of String),
                                           progress As Action(Of String, Integer, Integer),
                                           ByRef counter As Integer,
                                           totalRoots As Integer,
                                           cancel As CancellationToken,
                                           includeHiddenAndSystem As Boolean,
                                           skipRoots As HashSet(Of String),
                                           skipNameSet As HashSet(Of String),
                                           skipContainsFragments As String())

        If String.IsNullOrWhiteSpace(root) Then Exit Sub

        Dim startPath As String = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        If startPath.Length = 0 Then Exit Sub
        If ShouldSkipRoot(startPath, skipRoots) Then Exit Sub

        Dim stack As New Stack(Of String)()
        stack.Push(startPath)

        While stack.Count > 0
            If cancel.IsCancellationRequested Then Exit While
            Dim current As String = stack.Pop()

            If ShouldSkipByNameOrFragment(current, skipNameSet, skipContainsFragments) Then
                Continue While
            End If

            allDirs.Add(current)
            counter += 1
            If (counter Mod 1000) = 0 Then
                progress?.Invoke("Scanning for directories…", counter, totalRoots)
                ' After every N directories, yield a millisecond
                If (counter Mod 2000) = 0 Then Thread.Sleep(2)
            End If

            Dim findData As New WIN32_FIND_DATA()
            ' Ask kernel to LIMIT to directories — reduces per-entry work
            Dim hFind As IntPtr = FindFirstFileEx(MakeSearchPattern(current),
                                              FINDEX_INFO_LEVELS.FindExInfoBasic,
                                              findData,
                                              FINDEX_SEARCH_OPS.FindExSearchLimitToDirectories,
                                              IntPtr.Zero,
                                              FIND_FIRST_EX_LARGE_FETCH)
            If hFind = INVALID_HANDLE_VALUE Then
                Continue While
            End If

            Try
                Do
                    Dim name = findData.cFileName
                    If name = "." OrElse name = ".." Then Continue Do

                    Dim attrs As Integer = findData.dwFileAttributes
                    ' Kernel already limited to directories, but we keep the check
                    Dim isDir As Boolean = (attrs And FILE_ATTRIBUTE_DIRECTORY) <> 0
                    Dim isReparse As Boolean = (attrs And FILE_ATTRIBUTE_REPARSE_POINT) <> 0
                    If Not isDir OrElse isReparse Then Continue Do

                    If (Not includeHiddenAndSystem) AndAlso (((attrs And FILE_ATTRIBUTE_HIDDEN) <> 0) OrElse ((attrs And FILE_ATTRIBUTE_SYSTEM) <> 0)) Then
                        Continue Do
                    End If

                    Dim subdir = CombinePaths(current, name)
                    If Not ShouldSkipByNameOrFragment(subdir, skipNameSet, skipContainsFragments) Then
                        stack.Push(subdir)
                    End If
                Loop While FindNextFile(hFind, findData)
            Finally
                FindClose(hFind)
            End Try
        End While
    End Sub







    ' ---------- File enumeration within one dir ----------
    Private Sub EnumerateFilesFast(dir As String,
                                   exts As HashSet(Of String),
                                   onFile As Action(Of String, Long, String))
        Try
            Dim findData As New WIN32_FIND_DATA()
            Dim searchPattern As String = MakeSearchPattern(dir)

            Dim hFind As IntPtr = FindFirstFileEx(searchPattern,
                                                  FINDEX_INFO_LEVELS.FindExInfoBasic,
                                                  findData,
                                                  FINDEX_SEARCH_OPS.FindExSearchNameMatch,
                                                  IntPtr.Zero,
                                                  FIND_FIRST_EX_LARGE_FETCH)

            If hFind = INVALID_HANDLE_VALUE Then Exit Sub

            Try
                Do
                    Dim name = findData.cFileName
                    If name = "." OrElse name = ".." Then Continue Do

                    Dim attrs As Integer = findData.dwFileAttributes
                    Dim isDir As Boolean = (attrs And FILE_ATTRIBUTE_DIRECTORY) <> 0
                    If Not isDir Then
                        Dim extWithDot = Path.GetExtension(name)
                        If exts.Contains(extWithDot) Then
                            Dim fullPath As String = CombinePaths(dir, name)
                            Dim size As Long = (CLng(findData.nFileSizeHigh) << 32) Or (CLng(CUInt(findData.nFileSizeLow)))
                            onFile?.Invoke(fullPath, size, extWithDot.TrimStart("."c)) ' store ext without leading dot
                        End If
                    End If
                Loop While FindNextFile(hFind, findData)
            Finally
                FindClose(hFind)
            End Try
        Catch
            ' swallow per dir
        End Try
    End Sub

    ' ---------- Skip logic ----------
    Private Function BuildDefaultSkipRoots() As HashSet(Of String)
        Dim hs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each drive In DriveInfo.GetDrives().Where(Function(d) d.DriveType = DriveType.Fixed OrElse d.DriveType = DriveType.Removable)
            Dim root = drive.RootDirectory.FullName.TrimEnd("\"c)
            If String.IsNullOrEmpty(root) Then Continue For
            Dim win = root & "\Windows"
            Dim pf = root & "\Program Files"
            Dim pf86 = root & "\Program Files (x86)"
            Dim pd = root & "\ProgramData"
            hs.Add(win) : hs.Add(pf) : hs.Add(pf86) : hs.Add(pd)
        Next
        Return hs
    End Function

    Private Function BuildDefaultSkipNameSet() As HashSet(Of String)
        Return New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "Windows", "Program Files", "Program Files (x86)", "ProgramData",
            "AppData", "Common Files", "$Recycle.Bin", "System Volume Information"
        }
    End Function

    Private Function BuildDefaultSkipContains() As String()
        Return New String() {"\Windows\", "\Program Files\", "\Program Files (x86)\", "\ProgramData\", "\AppData\"}
    End Function

    Private Function ShouldSkipRoot(p As String, skipRoots As HashSet(Of String)) As Boolean
        If String.IsNullOrEmpty(p) Then Return True
        Dim norm = p.TrimEnd("\"c)
        Return skipRoots.Contains(norm)
    End Function

    Private Function ShouldSkipByNameOrFragment(path As String,
                                            skipNames As HashSet(Of String),
                                            skipContainsFragments As String()) As Boolean
        If String.IsNullOrEmpty(path) Then Return True

        ' Fast fragment checks first
        For Each frag In skipContainsFragments
            If path.IndexOf(frag, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        Next

        ' Avoid New DirectoryInfo(...).Name allocations:
        ' get last segment from string
        Dim p As String = path.TrimEnd("\"c)
        Dim lastSep As Integer = p.LastIndexOf("\"c)
        Dim leaf As String = If(lastSep >= 0, p.Substring(lastSep + 1), p)

        If skipNames.Contains(leaf) Then Return True
        Return False
    End Function


    ' ---------- Path helpers ----------
    Private Function MakeSearchPattern(dir As String) As String
        ' If already long path, skip ToLongPath
        Dim basePath As String
        If dir.StartsWith("\\?\", StringComparison.Ordinal) Then
            basePath = dir
        Else
            basePath = ToLongPath(dir)
        End If

        If basePath.EndsWith("\") Then
            Return basePath & "*"
        Else
            Return basePath & "\*"
        End If
    End Function

    Private Function CombinePaths(baseDir As String, name As String) As String
        If baseDir.EndsWith("\") Then
            Return baseDir & name
        Else
            Return baseDir & "\" & name
        End If
    End Function

    Private Function ToLongPath(p As String) As String
        If String.IsNullOrEmpty(p) Then Return p
        If p.StartsWith("\\?\") Then Return p
        If p.StartsWith("\\") Then
            Return "\\?\UNC\" & p.Substring(2)
        End If
        Return "\\?\" & p
    End Function

    Private Function StripLongPathPrefix(p As String) As String
        If String.IsNullOrEmpty(p) Then Return p
        If p.StartsWith("\\?\UNC\", StringComparison.OrdinalIgnoreCase) Then
            Return "\\" & p.Substring(8)
        End If
        If p.StartsWith("\\?\", StringComparison.OrdinalIgnoreCase) Then
            Return p.Substring(4)
        End If
        Return p
    End Function

    ' ---------- SQLite helpers ----------
    Private Sub ExecBusyTimeout(conn As SQLiteConnection, ms As Integer)
        Using cmd As New SQLiteCommand("PRAGMA busy_timeout=" & Math.Max(0, ms) & ";", conn)
            Try : cmd.ExecuteNonQuery() : Catch : End Try
        End Using
    End Sub







End Module
