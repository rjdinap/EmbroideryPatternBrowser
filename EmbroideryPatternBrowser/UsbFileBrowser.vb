' UsbFileBrowser.vb
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Management

Public Class UsbFileBrowser
    Inherits UserControl

    ' ===== UI =====
    Private ReadOnly header As New Panel()
    Private ReadOnly btnUp As New Button()
    Private ReadOnly txtPath As New TextBox()
    Private ReadOnly lv As New ListView()

    Private ReadOnly ctx As New ContextMenuStrip()
    Private ReadOnly mnuOpenExplorer As New ToolStripMenuItem("Open in Explorer")
    Private ReadOnly mnuNewFolder As New ToolStripMenuItem("New Folder…")
    Private ReadOnly mnuPaste As New ToolStripMenuItem("Paste (copy file)")

    ' Track which item was right-clicked
    Private _ctxItem As ListViewItem = Nothing

    ' ===== State =====
    Private currentFolder As String = Nothing   ' Nothing = ROOT (USB drives)

    ' Columns (no "Modified")
    Private Const COL_NAME As Integer = 0
    Private Const COL_SIZE As Integer = 1
    Private Const COL_TYPE As Integer = 2

    ' Sorting
    Private lastSortCol As Integer = COL_NAME
    Private sortAscending As Boolean = True

    ' ===== Device notification =====
    Private _devNotifyHandle As IntPtr = IntPtr.Zero

    ' ===== Debounce for WM_DEVICECHANGE (avoid COM calls during input-synchronous dispatch) =====
    Private _deviceChangePending As Boolean = False
    Private ReadOnly _deviceChangeTimer As New Timer() With {.Interval = 75}

    Public Sub New()
        Me.DoubleBuffered = True
        Me.Dock = DockStyle.Fill

        ' Bump font +1pt
        Try
            Dim bigger As New Font(Me.Font.FontFamily, Me.Font.SizeInPoints + 1.0F, Me.Font.Style)
            Me.Font = bigger
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try

        ' ===== Header (Up + Path) =====
        header.Dock = DockStyle.Top
        header.Height = 32
        header.Padding = New Padding(6, 4, 6, 4)

        btnUp.Text = "↑ Up"
        btnUp.Width = 68
        btnUp.Dock = DockStyle.Left
        AddHandler btnUp.Click, AddressOf BtnUp_Click

        txtPath.ReadOnly = True
        txtPath.Dock = DockStyle.Fill
        txtPath.BorderStyle = BorderStyle.FixedSingle

        header.Controls.Add(txtPath)
        header.Controls.Add(btnUp)
        Controls.Add(header)

        ' ===== ListView =====
        lv.Dock = DockStyle.Fill
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.HideSelection = False
        lv.Columns.Add("Name", 320, HorizontalAlignment.Left)
        lv.Columns.Add("Size", 110, HorizontalAlignment.Right)
        lv.Columns.Add("Type", 120, HorizontalAlignment.Left)
        AddHandler lv.ItemActivate, AddressOf Lv_ItemActivate   ' double-click / Enter
        AddHandler lv.KeyDown, AddressOf Lv_KeyDown             ' Backspace/Alt+Up
        AddHandler lv.ColumnClick, AddressOf Lv_ColumnClick
        AddHandler lv.MouseUp, AddressOf Lv_MouseUp
        Controls.Add(lv)

        ' ===== Context menu =====
        AddHandler mnuOpenExplorer.Click, AddressOf MnuOpenExplorer_Click
        AddHandler mnuNewFolder.Click, AddressOf MnuNewFolder_Click
        AddHandler mnuPaste.Click, AddressOf MnuPaste_Click
        AddHandler ctx.Opening, AddressOf Ctx_Opening
        ctx.Items.AddRange(New ToolStripItem() {mnuOpenExplorer, mnuNewFolder, mnuPaste})
        lv.ContextMenuStrip = ctx

        ' Debounce handler
        AddHandler _deviceChangeTimer.Tick, AddressOf DeviceChangeTimer_Tick

        ' Register for device notifications & initial populate when handle exists
        AddHandler Me.HandleCreated, Sub(sender As Object, e As EventArgs)
                                         Try
                                             RegisterForVolumeNotifications()
                                         Catch ex As Exception
                                             Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                                         End Try
                                         RefreshView()
                                     End Sub

        ' Ensure unregister when disposed
        AddHandler Me.Disposed, Sub()
                                    Try
                                        _deviceChangeTimer.Stop()
                                        _deviceChangeTimer.Dispose()
                                    Catch ex As Exception
                                        Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                                    End Try
                                    Try
                                        UnregisterVolumeNotifications()
                                    Catch ex As Exception
                                        Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                                    End Try
                                End Sub
    End Sub

#Region "Device Notifications (RegisterDeviceNotification + WM_DEVICECHANGE)"
    ' WM_DEVICECHANGE codes
    Private Const WM_DEVICECHANGE As Integer = &H219
    Private Const DBT_DEVICEARRIVAL As Integer = &H8000
    Private Const DBT_DEVICEREMOVECOMPLETE As Integer = &H8004

    ' RegisterDeviceNotification
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function RegisterDeviceNotification(hRecipient As IntPtr,
                                                       NotificationFilter As IntPtr,
                                                       Flags As Integer) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function UnregisterDeviceNotification(hHandle As IntPtr) As Boolean
    End Function

    Private Const DEVICE_NOTIFY_WINDOW_HANDLE As Integer = &H0

    ' GUID_DEVINTERFACE_VOLUME = {53F5630D-B6BF-11D0-94F2-00A0C91EFB8B}
    Private Shared ReadOnly GUID_DEVINTERFACE_VOLUME As New Guid("53F5630D-B6BF-11D0-94F2-00A0C91EFB8B")

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Private Structure DEV_BROADCAST_DEVICEINTERFACE
        Public dbcc_size As Integer
        Public dbcc_devicetype As Integer
        Public dbcc_reserved As Integer
        Public dbcc_classguid As Guid
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=255)>
        Public dbcc_name As String
    End Structure

    Private Const DBT_DEVTYP_DEVICEINTERFACE As Integer = 5

    Private Sub RegisterForVolumeNotifications()
        If Me.IsHandleCreated AndAlso _devNotifyHandle = IntPtr.Zero Then
            Dim dbi As New DEV_BROADCAST_DEVICEINTERFACE()
            dbi.dbcc_size = Marshal.SizeOf(GetType(DEV_BROADCAST_DEVICEINTERFACE))
            dbi.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
            dbi.dbcc_classguid = GUID_DEVINTERFACE_VOLUME
            dbi.dbcc_name = String.Empty

            Dim buffer As IntPtr = IntPtr.Zero
            Try
                buffer = Marshal.AllocHGlobal(dbi.dbcc_size)
                Marshal.StructureToPtr(dbi, buffer, False)
                _devNotifyHandle = RegisterDeviceNotification(Me.Handle, buffer, DEVICE_NOTIFY_WINDOW_HANDLE)
                If _devNotifyHandle = IntPtr.Zero Then
                    Throw New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "RegisterDeviceNotification failed.")
                End If
            Finally
                If buffer <> IntPtr.Zero Then Marshal.FreeHGlobal(buffer)
            End Try
        End If
    End Sub

    Private Sub UnregisterVolumeNotifications()
        If _devNotifyHandle <> IntPtr.Zero Then
            Try
                UnregisterDeviceNotification(_devNotifyHandle)
            Catch
                ' ignore, but clear anyway
            Finally
                _devNotifyHandle = IntPtr.Zero
            End Try
        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)

        If m.Msg = WM_DEVICECHANGE Then
            Dim code As Integer = m.WParam.ToInt32()
            If code = DBT_DEVICEARRIVAL OrElse code = DBT_DEVICEREMOVECOMPLETE Then
                ' Do not perform IO or WMI here. Debounce to a timer so the input-synchronous call returns.
                If Not _deviceChangePending Then
                    _deviceChangePending = True
                    _deviceChangeTimer.Stop()
                    _deviceChangeTimer.Start()
                End If
            End If
        End If
    End Sub

    Private Sub DeviceChangeTimer_Tick(sender As Object, e As EventArgs)
        _deviceChangeTimer.Stop()
        _deviceChangePending = False

        Try
            If currentFolder Is Nothing Then
                RefreshView()
            Else
                Dim root As String = Nothing
                Try
                    root = Path.GetPathRoot(currentFolder)
                Catch ex As Exception
                    Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                End Try

                If String.IsNullOrEmpty(root) OrElse Not Directory.Exists(root) Then
                    NavigateToRoot()
                Else
                    RefreshView()
                End If
            End If
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub
#End Region

    ' ===================== Public helpers =====================
    Public Sub NavigateToRoot()
        currentFolder = Nothing
        RefreshView()
    End Sub

    Public Sub NavigateToFolder(path As String)
        If String.IsNullOrWhiteSpace(path) OrElse Not Directory.Exists(path) Then Return
        currentFolder = path
        RefreshView()
    End Sub

    ' ===================== UI Actions =====================
    Private Sub BtnUp_Click(sender As Object, e As EventArgs)
        If currentFolder Is Nothing Then Return

        Dim parent As String = Nothing
        Try
            parent = Directory.GetParent(currentFolder)?.FullName
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try

        If String.IsNullOrEmpty(parent) OrElse Not Directory.Exists(parent) Then
            NavigateToRoot()
        Else
            NavigateToFolder(parent)
        End If
    End Sub

    Private Sub Lv_ItemActivate(sender As Object, e As EventArgs)
        If lv.SelectedItems.Count = 0 Then Return
        Dim it = lv.SelectedItems(0)
        Dim kind = TryCast(it.Tag, ItemKind)
        If kind Is Nothing Then Return

        Select Case kind.Kind
            Case ItemKindType.Drive, ItemKindType.Folder
                Try
                    If Directory.Exists(kind.Path) Then NavigateToFolder(kind.Path)
                Catch ex As Exception
                    Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                End Try
            Case ItemKindType.File
                ' No open action per requirements
        End Select
    End Sub

    Private Sub Lv_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Back OrElse (e.Alt AndAlso e.KeyCode = Keys.Up) Then
            BtnUp_Click(sender, e)
            e.Handled = True
        End If
    End Sub

    Private Sub Lv_MouseUp(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            Try
                _ctxItem = lv.HitTest(e.Location).Item   ' may be Nothing
                ctx.Show(lv, e.Location)
            Catch ex As Exception
                Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
            End Try
        End If
    End Sub

    Private Sub Lv_ColumnClick(sender As Object, e As ColumnClickEventArgs)
        If e.Column = lastSortCol Then
            sortAscending = Not sortAscending
        Else
            lastSortCol = e.Column
            sortAscending = True
        End If
        RefreshView()
    End Sub

    ' Context menu visibility rules (no clipboard probing here)
    Private Sub Ctx_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs)
        Dim inRealFolder As Boolean = (currentFolder IsNot Nothing) AndAlso Directory.Exists(currentFolder)
        mnuNewFolder.Enabled = inRealFolder

        ' Don't sniff clipboard on open; just show Paste when in a folder
        mnuPaste.Visible = inRealFolder

        ' Enable "Open in Explorer" only if we have a valid target
        Dim target As String = GetExplorerTargetPath()
        mnuOpenExplorer.Enabled = Not String.IsNullOrEmpty(target)
    End Sub

    ' Choose what to open in Explorer
    Private Function GetExplorerTargetPath() As String
        Try
            If _ctxItem IsNot Nothing Then
                Dim kind = TryCast(_ctxItem.Tag, ItemKind)
                If kind IsNot Nothing Then
                    Select Case kind.Kind
                        Case ItemKindType.Drive, ItemKindType.Folder
                            If Directory.Exists(kind.Path) Then Return kind.Path
                        Case ItemKindType.File
                            If File.Exists(kind.Path) Then Return kind.Path
                    End Select
                End If
            End If

            If currentFolder IsNot Nothing AndAlso Directory.Exists(currentFolder) Then
                Return currentFolder
            End If
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try

        Return Nothing ' At USB root with no selection
    End Function

    Private Sub MnuOpenExplorer_Click(sender As Object, e As EventArgs)
        Dim target As String = GetExplorerTargetPath()
        If String.IsNullOrEmpty(target) Then Return

        Try
            If File.Exists(target) Then
                ' Pre-select file in Explorer
                Process.Start(New ProcessStartInfo With {
                    .FileName = "explorer.exe",
                    .Arguments = "/select,""" & target & """",
                    .UseShellExecute = True
                })
            ElseIf Directory.Exists(target) Then
                Process.Start(New ProcessStartInfo With {
                    .FileName = "explorer.exe",
                    .Arguments = """" & target & """",
                    .UseShellExecute = True
                })
            End If
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
            MessageBox.Show("Could not open in Explorer: " & ex.Message, "USB Browser",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ===================== Populate View =====================
    Private Sub RefreshView()
        Try
            txtPath.Text = If(currentFolder, "(USB Drives)")
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try

        lv.BeginUpdate()
        Try
            lv.Items.Clear()

            If currentFolder Is Nothing Then
                ' ROOT: show only USB drives (filtered)
                Dim drives As List(Of String) = GetUsbDriveLetters()

                Dim items = New List(Of ListViewItem)
                For Each root In drives
                    Dim label As String = root.TrimEnd("\"c)
                    Try
                        Dim di As New DriveInfo(root)
                        If di.IsReady AndAlso Not String.IsNullOrWhiteSpace(di.VolumeLabel) Then
                            label = di.VolumeLabel & " (" & di.Name.TrimEnd("\"c) & ")"
                        End If
                    Catch ex As Exception
                        Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                    End Try

                    Dim li As New ListViewItem(label)
                    li.SubItems.Add("")                ' size
                    li.SubItems.Add("USB Drive")       ' type
                    li.Tag = New ItemKind(ItemKindType.Drive, root)
                    items.Add(li)
                Next

                ApplySort(items, isRoot:=True)
                If items.Count > 0 Then lv.Items.AddRange(items.ToArray())

            Else
                ' FOLDERS first, then FILES
                Dim folderItems As New List(Of ListViewItem)
                Dim fileItems As New List(Of ListViewItem)

                ' Folders
                Dim dirs As IEnumerable(Of String) = Enumerable.Empty(Of String)()
                Try
                    dirs = Directory.EnumerateDirectories(currentFolder)
                Catch ex As Exception
                    Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                End Try
                For Each d In dirs
                    Dim baseName As String = ""
                    Try
                        baseName = Path.GetFileName(d)
                    Catch ex As Exception
                        Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                        Continue For
                    End Try
                    Dim displayName = baseName & " (dir)"
                    Dim li As New ListViewItem(displayName)
                    li.SubItems.Add("")                 ' size
                    li.SubItems.Add("Folder")           ' type
                    li.Tag = New ItemKind(ItemKindType.Folder, d)
                    folderItems.Add(li)
                Next

                ' Files
                Dim files As IEnumerable(Of String) = Enumerable.Empty(Of String)()
                Try
                    files = Directory.EnumerateFiles(currentFolder)
                Catch ex As Exception
                    Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                End Try
                For Each f In files
                    Dim fi As FileInfo = Nothing
                    Try
                        fi = New FileInfo(f)
                    Catch ex As Exception
                        Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                        Continue For
                    End Try

                    Dim li As New ListViewItem(fi.Name)
                    li.SubItems.Add(FormatSize(fi.Length))
                    li.SubItems.Add(fi.Extension.ToUpperInvariant())
                    li.Tag = New ItemKind(ItemKindType.File, fi.FullName)
                    fileItems.Add(li)
                Next

                ApplySort(folderItems, isRoot:=False, isFolder:=True)
                ApplySort(fileItems, isRoot:=False, isFolder:=False)

                If folderItems.Count > 0 Then lv.Items.AddRange(folderItems.ToArray())
                If fileItems.Count > 0 Then lv.Items.AddRange(fileItems.ToArray())

                If lv.Items.Count > 0 Then
                    Try
                        lv.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
                        lv.Columns(COL_NAME).Width = Math.Max(lv.Columns(COL_NAME).Width, 280)
                    Catch ex As Exception
                        Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                    End Try
                End If
            End If
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        Finally
            lv.EndUpdate()
        End Try
    End Sub

    Private Sub ApplySort(ByRef items As List(Of ListViewItem), isRoot As Boolean, Optional isFolder As Boolean = False)
        Dim cmp As Comparison(Of ListViewItem)

        Select Case lastSortCol
            Case COL_NAME
                cmp = Function(a, b) StringComparer.CurrentCultureIgnoreCase.Compare(a.Text, b.Text)
            Case COL_SIZE
                cmp = Function(a, b)
                          Dim sa = If(a.SubItems.Count > COL_SIZE, a.SubItems(COL_SIZE).Text, "0 B")
                          Dim sb = If(b.SubItems.Count > COL_SIZE, b.SubItems(COL_SIZE).Text, "0 B")
                          Return ParseSize(sa).CompareTo(ParseSize(sb))
                      End Function
            Case COL_TYPE
                cmp = Function(a, b) StringComparer.CurrentCultureIgnoreCase.Compare(
                        If(a.SubItems.Count > COL_TYPE, a.SubItems(COL_TYPE).Text, ""),
                        If(b.SubItems.Count > COL_TYPE, b.SubItems(COL_TYPE).Text, ""))
            Case Else
                cmp = Function(a, b) 0
        End Select

        Try
            If cmp IsNot Nothing Then items.Sort(cmp)
            If Not sortAscending Then items.Reverse()
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub

    ' ===================== Context actions =====================
    Private Sub MnuNewFolder_Click(sender As Object, e As EventArgs)
        If currentFolder Is Nothing OrElse Not Directory.Exists(currentFolder) Then
            MessageBox.Show("Select a target folder first.", "USB Browser", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim name As String = Microsoft.VisualBasic.Interaction.InputBox("New folder name:", "Create Folder", "New Folder")
        If String.IsNullOrWhiteSpace(name) Then Return

        Dim dest As String = ""
        Try
            dest = Path.Combine(currentFolder, name.Trim())
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
            Return
        End Try

        Try
            If Not Directory.Exists(dest) Then
                Directory.CreateDirectory(dest)
                RefreshView()
            Else
                MessageBox.Show("Folder already exists.", "USB Browser", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
            MessageBox.Show("Failed to create folder: " & ex.Message, "USB Browser", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MnuPaste_Click(sender As Object, e As EventArgs)
        If currentFolder Is Nothing OrElse Not Directory.Exists(currentFolder) Then
            MessageBox.Show("Select a target folder first.", "USB Browser", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim src As String = TryGetFileFromClipboard()
        If String.IsNullOrEmpty(src) OrElse Not File.Exists(src) Then
            MessageBox.Show("Clipboard doesn't contain a valid file path.", "USB Browser", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim dest As String = ""
        Try
            dest = Path.Combine(currentFolder, Path.GetFileName(src))
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
            Return
        End Try

        Try
            If File.Exists(dest) Then
                If MessageBox.Show("File exists. Overwrite?", "USB Browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    Return
                End If
            End If
            File.Copy(src, dest, overwrite:=True)
            RefreshView()
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
            MessageBox.Show("Copy failed: " & ex.Message, "USB Browser", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ===================== USB enumeration (USB only, light WMI per-drive) =====================
    Private Function GetUsbDriveLetters() As List(Of String)
        Dim letters As New List(Of String)()

        Try
            For Each di In DriveInfo.GetDrives()
                Try
                    If di.IsReady AndAlso di.DriveType = DriveType.Removable Then
                        Dim root As String = di.RootDirectory.FullName ' e.g., "E:\"
                        If IsUsbDrive(root) Then
                            letters.Add(root)
                        End If
                    End If
                Catch ex As Exception
                    Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                End Try
            Next
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try

        letters.Sort()
        Return letters
    End Function

    ''' <summary>
    ''' Light per-drive WMI check: map logical drive -> partition -> diskdrive and verify InterfaceType = "USB"
    ''' or PNPDeviceID contains "USBSTOR". This inspects only the single drive letter passed in.
    ''' </summary>
    Private Function IsUsbDrive(rootWithSep As String) As Boolean
        Try
            Dim driveId As String
            Try
                ' Root like "E:\" -> "E:"
                driveId = rootWithSep.TrimEnd("\"c)
                If Not driveId.EndsWith(":") Then driveId &= ":"
            Catch ex As Exception
                Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                Return False
            End Try

            Using q1 As New ManagementObjectSearcher("SELECT * FROM Win32_LogicalDisk WHERE DeviceID='" & driveId.Replace("\", "\\") & "'")
                For Each ld As ManagementObject In q1.Get()
                    ' ld -> partition
                    Using q2 As New ManagementObjectSearcher("ASSOCIATORS OF {Win32_LogicalDisk.DeviceID='" & driveId.Replace("\", "\\") & "'} WHERE AssocClass = Win32_LogicalDiskToPartition")
                        For Each part As ManagementObject In q2.Get()
                            Dim partId As String = ""
                            Try
                                partId = CStr(part("DeviceID")).Replace("\", "\\")
                            Catch ex As Exception
                                Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                                Continue For
                            End Try

                            ' partition -> diskdrive
                            Using q3 As New ManagementObjectSearcher("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" & partId & "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition")
                                For Each dd As ManagementObject In q3.Get()
                                    Dim iface As String = ""
                                    Dim pnp As String = ""
                                    Try
                                        iface = GetMoString(dd, "InterfaceType")
                                    Catch ex As Exception
                                        Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                                    End Try
                                    Try
                                        pnp = GetMoString(dd, "PNPDeviceID")
                                    Catch ex As Exception
                                        Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                                    End Try

                                    If String.Equals(iface, "USB", StringComparison.OrdinalIgnoreCase) Then Return True
                                    If (Not String.IsNullOrEmpty(pnp)) AndAlso pnp.ToUpperInvariant().Contains("USBSTOR") Then Return True
                                Next
                            End Using
                        Next
                    End Using
                Next
            End Using
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try

        Return False
    End Function

    ' ===================== Clipboard helpers (no text probing) =====================
    ' Reads the clipboard ONLY when Paste is clicked (not during menu opening).
    ' Accepts either a FileDropList or a plain text path (like Explorer "Copy as path").
    Private Function TryGetFileFromClipboard() As String
        Try
            ' 1) Preferred: CF_HDROP (FileDropList)
            If Clipboard.ContainsFileDropList() Then
                Dim files = Clipboard.GetFileDropList()
                If files IsNot Nothing AndAlso files.Count > 0 Then
                    Dim f = files(0)
                    If Not String.IsNullOrWhiteSpace(f) Then
                        Dim p = f.Trim()
                        If File.Exists(p) Then Return p
                    End If
                End If
            End If

            ' 2) Fallback: Unicode text path (Explorer "Copy as path" etc.)
            If Clipboard.ContainsText(TextDataFormat.UnicodeText) Then
                Dim t As String = Clipboard.GetText(TextDataFormat.UnicodeText)
                If Not String.IsNullOrWhiteSpace(t) Then
                    Dim p As String = t.Trim()

                    ' Remove surrounding quotes if present
                    If p.Length >= 2 AndAlso p.StartsWith("""") AndAlso p.EndsWith("""") Then
                        p = p.Substring(1, p.Length - 2)
                    End If

                    ' Normalize whitespace
                    p = p.Trim()

                    ' Ignore trailing backslash for files (but keep UNC \\server\share\file.ext)
                    ' (No special handling needed; File.Exists handles both.)

                    ' Validate it points to a file
                    If File.Exists(p) Then Return p
                End If
            End If

        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try

        Return Nothing
    End Function

    ' ===================== Utilities =====================
    Private Shared Function FormatSize(len As Long) As String
        Dim units = {"B", "KB", "MB", "GB", "TB"}
        Dim size As Double = len
        Dim u As Integer = 0
        While size >= 1024 AndAlso u < units.Length - 1
            size /= 1024
            u += 1
        End While
        Return String.Format("{0:0.##} {1}", size, units(u))
    End Function

    Private Shared Function ParseSize(s As String) As Long
        Try
            Dim parts = s.Split(" "c)
            Dim val As Double = Double.Parse(parts(0))
            Dim unit = parts(1).ToUpperInvariant()
            Dim mult As Long = 1
            Select Case unit
                Case "KB" : mult = 1024L
                Case "MB" : mult = 1024L ^ 2
                Case "GB" : mult = 1024L ^ 3
                Case "TB" : mult = 1024L ^ 4
                Case Else : mult = 1
            End Select
            Return CLng(val * mult)
        Catch
            Return 0
        End Try
    End Function

    Private Class ItemKind
        Public ReadOnly Kind As ItemKindType
        Public ReadOnly Path As String
        Public Sub New(kind As ItemKindType, path As String)
            Me.Kind = kind
            Me.Path = path
        End Sub
    End Class

    Private Enum ItemKindType
        Drive
        Folder
        File
    End Enum

    ' ==== WMI helper ====
    Private Function GetMoString(mo As ManagementObject, propName As String) As String
        Try
            If mo Is Nothing Then Return ""
            Dim prop = mo.Properties(propName)
            If prop Is Nothing Then Return ""
            Dim val = prop.Value
            If val Is Nothing Then Return ""
            Return val.ToString()
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
            Return ""
        End Try
    End Function

End Class
