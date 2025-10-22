Imports System.IO
Imports System.Runtime.InteropServices

Module ShellSelectHelper
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Function SHParseDisplayName(pszName As String, pbc As IntPtr,
                                        ByRef ppidl As IntPtr, sfgaoIn As UInteger,
                                        ByRef psfgaoOut As UInteger) As Integer
    End Function

    <DllImport("shell32.dll", SetLastError:=True)>
    Private Function SHOpenFolderAndSelectItems(pidlFolder As IntPtr,
                                                cidl As UInteger,
                                                <[In]()> apidl() As IntPtr,
                                                dwFlags As UInteger) As Integer
    End Function

    ' Returns a pointer to the last SHITEMID (child) within a PIDL. Pointer is within pidl's memory.
    <DllImport("shell32.dll")>
    Private Function ILFindLastID(pidl As IntPtr) As IntPtr
    End Function

    <DllImport("ole32.dll")>
    Private Sub CoTaskMemFree(pv As IntPtr)
    End Sub

    ''' <summary>
    ''' Opens Explorer at the file's parent folder and selects the file. Robust for long/UNC paths.
    ''' Falls back to /select or just opening the folder.
    ''' </summary>
    Public Sub OpenAndSelect(path As String)
        If String.IsNullOrWhiteSpace(path) Then Return

        Try
            ' If it's a composite path like C:\foo.zip?inner\file.jef,
            ' Explorer can only select the outer .zip.
            Dim effectivePath As String = path
            Dim q As Integer = path.IndexOf("?"c)
            If q > 1 Then
                effectivePath = path.Substring(0, q) ' keep only C:\...\file.zip
            End If

            Dim folder As String = Nothing
            Dim selectFile As String = Nothing

            If File.Exists(effectivePath) Then
                folder = System.IO.Path.GetDirectoryName(effectivePath)
                selectFile = effectivePath
            ElseIf Directory.Exists(effectivePath) Then
                ' If the left side is actually a folder (rare), just open it
                folder = effectivePath
            Else
                ' Fall back to the folder of the left side (even if the zip no longer exists)
                folder = System.IO.Path.GetDirectoryName(effectivePath)
            End If

            ' If we still don’t have a real folder, last-resort open Explorer home
            If String.IsNullOrEmpty(folder) OrElse Not Directory.Exists(folder) Then
                Process.Start("explorer.exe")
                Return
            End If

            ' Prefer the PIDL route for robust selection
            Dim pidlFolder As IntPtr = IntPtr.Zero
            Dim pidlFile As IntPtr = IntPtr.Zero
            Dim attrs As UInteger = 0UI
            Try
                Dim hrF = SHParseDisplayName(folder, IntPtr.Zero, pidlFolder, 0UI, attrs)
                Dim haveFolder As Boolean = (hrF = 0 AndAlso pidlFolder <> IntPtr.Zero)

                Dim haveFile As Boolean = False
                If Not String.IsNullOrEmpty(selectFile) AndAlso File.Exists(selectFile) Then
                    Dim hrI = SHParseDisplayName(selectFile, IntPtr.Zero, pidlFile, 0UI, attrs)
                    haveFile = (hrI = 0 AndAlso pidlFile <> IntPtr.Zero)
                End If

                If haveFolder AndAlso haveFile Then
                    Dim rel As IntPtr = ILFindLastID(pidlFile)
                    Dim children() As IntPtr = {rel}
                    SHOpenFolderAndSelectItems(pidlFolder, 1UI, children, 0UI)
                ElseIf haveFolder Then
                    ' No file to select (missing zip, etc.) – open the folder instead
                    Process.Start("explorer.exe", """" & folder & """")
                Else
                    Process.Start("explorer.exe")
                End If
            Finally
                If pidlFile <> IntPtr.Zero Then CoTaskMemFree(pidlFile)
                If pidlFolder <> IntPtr.Zero Then CoTaskMemFree(pidlFolder)
            End Try

        Catch
            ' Silent best-effort fallback to Home
            Try : Process.Start("explorer.exe") : Catch : End Try
        End Try
    End Sub

End Module
