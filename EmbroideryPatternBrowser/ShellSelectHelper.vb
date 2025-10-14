Imports System
Imports System.Diagnostics
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
            If File.Exists(path) Then
                Dim folder As String = System.IO.Path.GetDirectoryName(path)

                Dim pidlFolder As IntPtr = IntPtr.Zero
                Dim pidlFile As IntPtr = IntPtr.Zero
                Dim attrs As UInteger = 0

                Dim hrF = SHParseDisplayName(folder, IntPtr.Zero, pidlFolder, 0UI, attrs)
                Dim hrI = SHParseDisplayName(path, IntPtr.Zero, pidlFile, 0UI, attrs)

                If hrF = 0 AndAlso pidlFolder <> IntPtr.Zero AndAlso hrI = 0 AndAlso pidlFile <> IntPtr.Zero Then
                    ' Use parent PIDL + child RELATIVE PIDL
                    Dim rel As IntPtr = ILFindLastID(pidlFile)
                    Dim children() As IntPtr = {rel}

                    SHOpenFolderAndSelectItems(pidlFolder, 1UI, children, 0UI)
                Else
                    ' Fallback: /select switch
                    Process.Start("explorer.exe", "/select,""" & path & """")
                End If

                If pidlFile <> IntPtr.Zero Then CoTaskMemFree(pidlFile)
                If pidlFolder <> IntPtr.Zero Then CoTaskMemFree(pidlFolder)

            Else
                ' File missing → open folder if possible
                Dim dir = System.IO.Path.GetDirectoryName(path)
                If Not String.IsNullOrEmpty(dir) AndAlso Directory.Exists(dir) Then
                    Process.Start("explorer.exe", """" & dir & """")
                Else
                    Process.Start("explorer.exe")
                End If
            End If
        Catch
            ' Optional: log or show a message; we silently ignore here.
        End Try
    End Sub
End Module
