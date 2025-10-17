Imports System.IO
Imports System.IO.Compression

Public Class ZipProcessing
    Public Sub New()
    End Sub




    ' Parse filename: returns (pathBeforeQ, extensionOfThatPath, remainderAfterQ)
    Public Function ParseFilename(input As String) As Tuple(Of String, String, String)
        If String.IsNullOrEmpty(input) Then
            Return Tuple.Create("", "", "")
        End If

        Dim q As Integer = input.IndexOf("?"c)
        Dim pathPart As String
        Dim remainder As String

        If q >= 0 Then
            pathPart = input.Substring(0, q)
            remainder = If(q < input.Length - 1, input.Substring(q + 1), "")
        Else
            pathPart = input
            remainder = ""
        End If

        Dim ext As String = ""
        Try
            ext = Path.GetExtension(pathPart)
            If ext Is Nothing Then ext = ""
        Catch
            ext = ""
        End Try

        Return Tuple.Create(pathPart, ext, remainder)
    End Function



    ' === NEW: Read a single entry from a .zip and return it as a MemoryStream ===
    Public Function ReadZip(zipPath As String, innerRelativePath As String) As Stream
        If String.IsNullOrEmpty(zipPath) Then Throw New ArgumentNullException(NameOf(zipPath))
        If innerRelativePath Is Nothing Then innerRelativePath = ""

        ' Normalize "innerRelativePath" to ZIP-style forward slashes and trim leading slashes
        Dim target As String = innerRelativePath.Replace("\"c, "/"c).Trim("/"c)

        Using fs As New FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite Or FileShare.Delete)
            Using za As New ZipArchive(fs, ZipArchiveMode.Read, leaveOpen:=False)
                Dim found As ZipArchiveEntry = Nothing
                For Each entry As ZipArchiveEntry In za.Entries
                    Dim name As String = entry.FullName.Trim("/"c)
                    If String.Equals(name, target, StringComparison.OrdinalIgnoreCase) Then
                        found = entry
                        Exit For
                    End If
                Next
                If found Is Nothing Then
                    Throw New FileNotFoundException($"Entry not found in zip: {zipPath}?{innerRelativePath}")
                End If

                Dim ms As New MemoryStream()
                Using zs As Stream = found.Open()
                    zs.CopyTo(ms)
                End Using
                ms.Position = 0
                Return ms   ' caller disposes
            End Using
        End Using
    End Function




    Public Function ScanZipForFileNames(zipPath As String) As List(Of (FullPath As String, Size As Long, Ext As String))
        ' Open allowing concurrent read (don’t lock the file if AV is touching it, etc.) as 
        Dim zipEntries As New List(Of (FullPath As String, Size As Long, Ext As String))()
        Try
            Using fs As New FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite Or FileShare.Delete)
                Using za As New ZipArchive(fs, ZipArchiveMode.Read, leaveOpen:=False)
                    For Each entry As ZipArchiveEntry In za.Entries
                        ' Skip directories
                        If entry.FullName.EndsWith("/", StringComparison.Ordinal) OrElse entry.FullName.EndsWith("\", StringComparison.Ordinal) Then
                            Continue For
                        End If

                        Dim innerName As String = entry.Name  ' short filename only
                        Dim innerExt As String = NormalizeExt(Path.GetExtension(innerName))
                        If Not IsEmbExt(innerExt) Then Continue For

                        Dim innerPathWin As String = ZipInnerPathToWindows(entry.FullName)
                        Dim compositeFullPath As String = zipPath & "?" & innerPathWin

                        Dim expandedSize As Long = 0
                        Try
                            ' Uncompressed size; may throw on corrupt entries
                            expandedSize = entry.Length
                        Catch
                            expandedSize = 0
                        End Try
                        zipEntries.Add((compositeFullPath, expandedSize, innerExt))
                    Next
                End Using
            End Using

        Catch ex As Exception
            ' Optional: log once per ZIP to avoid log spam on huge archives
            Form1.Status($"ZIP scan failed: {zipPath}: {ex.Message}")
        End Try
        Return zipEntries
    End Function








    'zip file helpers
    '-----------------------------------------------------------
    '-----------------------------------------------------------

    ' --- Allowed embroidery extensions (lowercase, include leading dot) ---
    Private ReadOnly EmbExt As HashSet(Of String) = New HashSet(Of String)(
        New String() {".pes", ".hus", ".vp3", ".xxx", ".dst", ".jef", ".sew", ".vip", ".exp", ".pec"},
        StringComparer.OrdinalIgnoreCase
    )

    Private Function NormalizeExt(ext As String) As String
        If String.IsNullOrEmpty(ext) Then Return ""
        If ext.StartsWith(".") Then
            Return ext.ToLowerInvariant()
        Else
            Return ("."c & ext).ToLowerInvariant()
        End If
    End Function

    Private Function IsEmbExt(ext As String) As Boolean
        Return EmbExt.Contains(NormalizeExt(ext))
    End Function

    Private Function ZipInnerPathToWindows(p As String) As String
        ' Zip stores with forward slashes; convert to backslashes for consistency
        Dim s = p.Replace("/"c, "\"c)
        ' remove any leading backslash (optional style)
        If s.StartsWith("\"c) Then s = s.Substring(1)
        Return s
    End Function



End Class