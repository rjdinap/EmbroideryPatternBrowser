' UsbFileBrowser.vb
Imports System.IO
Imports System.Management
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Linq


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
    ' --- Virtual/zip data exchange format (between our controls) ---
    Private Const DATAFMT_VIRTUAL_LIST As String = "EPB_VIRTUAL_FILE_LIST"


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
    Private ReadOnly _deviceChangeTimer As New System.Windows.Forms.Timer() With {.Interval = 75}

    ' ===== Status UI (Panel_Right_Bottom_Top) =====
    Private _statusLabel As Label = Nothing
    Private _lastStatusText As String = ""
    Private _lastStatusTick As Long = 0
    Private ReadOnly _statusSw As New Stopwatch()
    ' ===== Status UI (progress polling) =====
    Private _progressTimer As New System.Windows.Forms.Timer() With {.Interval = 100} ' 100ms UI tick
    Private _progressTotal As Integer = 0
    Private _progressIndex As Integer = 0
    Private _progressActive As Boolean = False
    ' Serialize copy operations so only one runs at a time
    Private ReadOnly _copySemaphore As New System.Threading.SemaphoreSlim(1, 1)
    ' new: one CTS that cancels the current/queued copies
    Private _copyCts As New CancellationTokenSource()


    Public Sub New()
        Me.DoubleBuffered = True
        Me.Dock = DockStyle.Fill

        'drag drop support
        Me.AllowDrop = True
        lv.AllowDrop = True
        AddHandler Me.DragEnter, AddressOf USB_DragEnter
        AddHandler Me.DragDrop, AddressOf USB_DragDrop
        AddHandler lv.DragEnter, AddressOf USB_DragEnter
        AddHandler lv.DragDrop, AddressOf USB_DragDrop

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

        _statusSw.Start()
        AddHandler _progressTimer.Tick, AddressOf ProgressTimer_Tick
    End Sub


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



    'drag drop version
    Private Async Sub MnuPaste_Click(sender As Object, e As EventArgs)
        If currentFolder Is Nothing OrElse Not Directory.Exists(currentFolder) Then
            MessageBox.Show("Select a target folder first.", "USB Browser", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim files = TryGetFilesFromClipboard()
        If files Is Nothing OrElse files.Count = 0 Then
            MessageBox.Show("Clipboard doesn't contain file paths.", "USB Browser", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Await PasteFilePathsAsync(files)
    End Sub

    'drag drop helper
    Private Function TryGetFilesFromClipboard() As List(Of String)
        Dim out As New List(Of String)
        Try
            ' 1) Our custom format first (if present)
            If Clipboard.ContainsData(DATAFMT_VIRTUAL_LIST) Then
                Dim obj = Clipboard.GetData(DATAFMT_VIRTUAL_LIST)
                If TypeOf obj Is IEnumerable(Of String) Then
                    For Each s In DirectCast(obj, IEnumerable(Of String))
                        If Not String.IsNullOrWhiteSpace(s) Then out.Add(s.Trim())
                    Next
                ElseIf TypeOf obj Is String Then
                    Dim lines = CStr(obj).Split({ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
                    For Each s In lines
                        Dim p = s.Trim()
                        If p.Length > 0 Then out.Add(p)
                    Next
                End If
            End If

            ' 2) Standard FileDrop (real files)
            If Clipboard.ContainsFileDropList() Then
                For Each f As String In Clipboard.GetFileDropList()
                    If Not String.IsNullOrWhiteSpace(f) Then out.Add(f.Trim())
                Next
            End If

            ' 3) Fallback: text (can contain either real paths or composite .zip?inner paths)
            If Clipboard.ContainsText(TextDataFormat.UnicodeText) Then
                Dim t = Clipboard.GetText(TextDataFormat.UnicodeText)
                If Not String.IsNullOrWhiteSpace(t) Then
                    For Each line In t.Split({ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
                        Dim p = line.Trim().Trim(""""c)
                        If p.Length > 0 Then out.Add(p)
                    Next
                End If
            End If

        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try

        ' Keep both real file paths and composite zip paths
        Return out.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function



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

    ' Light per-drive WMI check: map logical drive -> partition -> diskdrive and verify InterfaceType = "USB"
    ' or PNPDeviceID contains "USBSTOR". This inspects only the single drive letter passed in.
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



    'drag drop support

    Private Sub USB_DragEnter(sender As Object, e As DragEventArgs)
        If e.Data Is Nothing OrElse currentFolder Is Nothing OrElse Not Directory.Exists(currentFolder) Then
            e.Effect = DragDropEffects.None : Return
        End If

        If e.Data.GetDataPresent(DATAFMT_VIRTUAL_LIST) OrElse
       e.Data.GetDataPresent(DataFormats.FileDrop) OrElse
       e.Data.GetDataPresent(DataFormats.UnicodeText) OrElse
       e.Data.GetDataPresent(DataFormats.Text) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub



    Private Async Sub USB_DragDrop(sender As Object, e As DragEventArgs)
        If currentFolder Is Nothing OrElse Not Directory.Exists(currentFolder) Then Return
        Dim files = DecodeDragData(e.Data)
        If files Is Nothing OrElse files.Count = 0 Then Return
        Await PasteFilePathsAsync(files)
    End Sub


    Private Function GenerateUniqueDestinationPath(destPath As String) As String
        If Not IO.File.Exists(destPath) Then Return destPath
        Dim dir = IO.Path.GetDirectoryName(destPath)
        Dim baseName = IO.Path.GetFileNameWithoutExtension(destPath)
        Dim ext = IO.Path.GetExtension(destPath)
        Dim i As Integer = 1
        Do
            Dim candidate = IO.Path.Combine(dir, $"{baseName} - Copy ({i}){ext}")
            If Not IO.File.Exists(candidate) Then Return candidate
            i += 1
        Loop
    End Function




    ' Copies the given files to the currently open folder, showing progress in Panel_Right_Bottom_Top
    Private Async Function PasteFilePathsAsync(srcFiles As IEnumerable(Of String)) As Threading.Tasks.Task
        Dim token = _copyCts.Token
        Await _copySemaphore.WaitAsync(token).ConfigureAwait(False)
        Try
            ' Keep both real and composite zip paths; just drop blanks/dupes
            Dim list = srcFiles.
            Where(Function(p) Not String.IsNullOrWhiteSpace(p)).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

            Dim total As Integer = list.Count
            SetStatus($"Copying {total} file(s)…")

            ' ---- start UI-driven progress ----
            _progressTotal = total
            Threading.Interlocked.Exchange(_progressIndex, 0)
            _progressActive = True
            _progressTimer.Stop()
            _progressTimer.Start()
            ' ----------------------------------

            Await Threading.Tasks.Task.Run(
            Sub()
                Dim zp As New ZipProcessing()

                For Each src In list
                    token.ThrowIfCancellationRequested()

                    Try
                        Dim destName As String
                        Dim destPath As String
                        Dim zipPath As String = Nothing
                        Dim innerPath As String = Nothing

                        If IsCompositeZipPath(src, zipPath, innerPath) Then
                            ' Copy out of zip: name comes from inner entry
                            destName = Path.GetFileName(innerPath)
                            destPath = Path.Combine(currentFolder, destName)
                            If File.Exists(destPath) Then destPath = GenerateUniqueDestinationPath(destPath)

                            Using inStream As Stream = zp.ReadZip(zipPath, innerPath)
                                ' CreateNew + pre-generated unique path preserves your existing behavior
                                Using outFs As New FileStream(destPath, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                                    inStream.CopyTo(outFs)
                                End Using
                            End Using

                        Else
                            ' Normal file on disk — keep your no-overwrite policy
                            destName = Path.GetFileName(src)
                            destPath = Path.Combine(currentFolder, destName)
                            If File.Exists(destPath) Then destPath = GenerateUniqueDestinationPath(destPath)
                            File.Copy(src, destPath, overwrite:=False)
                        End If

                    Catch ex As OperationCanceledException
                        Throw
                    Catch ex As Exception
                        ' Log and continue to next file
                        Form1.Status("Copy error: " & ex.Message, ex.StackTrace.ToString)
                    Finally
                        Threading.Interlocked.Increment(_progressIndex)
                    End Try
                Next
            End Sub, token).ConfigureAwait(False)

            ' UI refresh + final status
            If Me.IsHandleCreated AndAlso Not Me.IsDisposed Then
                Me.BeginInvoke(DirectCast(Sub()
                                              RefreshView()
                                              SetStatus("Copy complete.")
                                              ClearStatusSoon()
                                          End Sub, Action))
            End If

        Catch ex As OperationCanceledException
            If Me.IsHandleCreated AndAlso Not Me.IsDisposed Then
                Me.BeginInvoke(DirectCast(Sub()
                                              RefreshView()
                                              SetStatus("Copy cancelled.")
                                              ClearStatusSoon()
                                          End Sub, Action))
            End If
        Finally
            _progressActive = False
            _progressTimer.Stop()
            ClearStatusSoon()
            _copySemaphore.Release()
            If _copyCts.IsCancellationRequested Then _copyCts = New CancellationTokenSource()
        End Try
    End Function



    Private Sub EnsureStatusLabel()
        If _statusLabel IsNot Nothing AndAlso Not _statusLabel.IsDisposed Then Return
        If Form1 Is Nothing OrElse Form1.IsDisposed Then Return

        Dim host = Form1.Panel_Right_Bottom_Top
        If host Is Nothing OrElse host.IsDisposed Then Return

        For Each c As Control In host.Controls
            If TypeOf c Is Label AndAlso String.Equals(c.Name, "lblCopyStatus", StringComparison.OrdinalIgnoreCase) Then
                _statusLabel = DirectCast(c, Label)
                Return
            End If
        Next

        _statusLabel = New Label() With {
        .Name = "lblCopyStatus",
        .AutoSize = True,
        .Dock = DockStyle.Fill,
        .TextAlign = ContentAlignment.MiddleLeft
    }
        host.Controls.Add(_statusLabel)
    End Sub


    Private Sub SetStatus(text As String)
        If Form1 Is Nothing OrElse Form1.IsDisposed OrElse Not Form1.IsHandleCreated Then Return

        ' --- Optional throttle: only update if text changed or >= 80ms elapsed ---
        If Not _statusSw.IsRunning Then _statusSw.Start()
        Dim nowMs = _statusSw.ElapsedMilliseconds
        If text = _lastStatusText AndAlso (nowMs - _lastStatusTick) < 80 Then
            Exit Sub
        End If
        _lastStatusText = text
        _lastStatusTick = nowMs
        ' ------------------------------------------------------------------------

        Dim apply As Action =
        Sub()
            EnsureStatusLabel()
            If _statusLabel Is Nothing OrElse _statusLabel.IsDisposed Then Exit Sub
            _statusLabel.Text = text
            ' Force a paint NOW so intermediate "Copying file i of N" messages are visible
            _statusLabel.Invalidate()
            _statusLabel.Update()    ' processes paint for this control
            ' You can also do: _statusLabel.Refresh()
        End Sub

        If Form1.InvokeRequired Then
            ' we’re on a background thread (e.g., inside Task.Run loop)
            Form1.BeginInvoke(apply)
        Else
            ' we’re already on the UI thread (e.g., the initial “Copying N files…”)
            apply()
        End If
    End Sub

    Private Sub ClearStatus()
        SetStatus(String.Empty)
    End Sub


    Private Async Sub ClearStatusSoon()
        Await Threading.Tasks.Task.Delay(2500)
        ClearStatus()
    End Sub

    Private Sub ProgressTimer_Tick(sender As Object, e As EventArgs)
        If Not _progressActive Then
            _progressTimer.Stop()
            Return
        End If

        EnsureStatusLabel()
        If _statusLabel Is Nothing OrElse _statusLabel.IsDisposed Then Return

        Dim idx = Math.Max(0, Math.Min(_progressIndex, _progressTotal))
        Dim tot = Math.Max(0, _progressTotal)

        If tot <= 0 Then
            _statusLabel.Text = "Copying..."
        Else
            _statusLabel.Text = $"Copying file {idx} of {tot}..."
        End If

        _statusLabel.Invalidate()
        _statusLabel.Update()
    End Sub

    ' Call me from Form1.FormClosing
    Public Function RequestAppClose(e As FormClosingEventArgs) As DialogResult
        ' If a copy is running (semaphore taken), ask the user
        If _copySemaphore.CurrentCount = 0 Then
            Dim r = MessageBox.Show(
                "A file copy is in progress." & Environment.NewLine &
                "Do you want to cancel the copy and exit?",
                "Exit",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question)

            Select Case r
                Case DialogResult.Cancel
                    e.Cancel = True
                    Return DialogResult.Cancel
                Case DialogResult.No
                    ' Don’t exit. Keep app open until copy finishes.
                    e.Cancel = True
                    Return DialogResult.No
                Case DialogResult.Yes
                    ' Cancel the running/queued copies
                    Try : _copyCts.Cancel() : Catch : End Try
                    ' Optionally show a message
                    SetStatus("Cancelling copy…")
                    ' Let closing continue; background loop will observe the token
                    Return DialogResult.Yes
            End Select
        End If
        Return DialogResult.Yes ' normally allow the close - unless files are being copied

    End Function


    ' True if s is like: C:\folder\pack.zip?inner\path\file.pes
    Private Shared Function IsCompositeZipPath(s As String, ByRef zipPath As String, ByRef innerPath As String) As Boolean
        zipPath = Nothing : innerPath = Nothing
        If String.IsNullOrWhiteSpace(s) Then Return False
        Try
            Dim zp As New ZipProcessing()
            Dim t = zp.ParseFilename(s)
            ' t.Item1 = before '?', t.Item2 = ext of outer, t.Item3 = after '?'
            If Not String.IsNullOrEmpty(t.Item1) AndAlso
               String.Equals(t.Item2, ".zip", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.IsNullOrEmpty(t.Item3) Then
                zipPath = t.Item1 : innerPath = t.Item3
                Return True
            End If
        Catch
        End Try
        Return False
    End Function

    Private Shared Function IsCompositeZipPath(s As String) As Boolean
        Dim z As String = Nothing, i As String = Nothing
        Return IsCompositeZipPath(s, z, i)
    End Function

    ' Turn IDataObject into a list of paths. Can include real files and/or composite zip paths.
    Private Function DecodeDragData(data As IDataObject) As List(Of String)
        Dim out As New List(Of String)
        Try
            ' Our custom format (list of strings)
            If data.GetDataPresent(DATAFMT_VIRTUAL_LIST) Then
                Dim obj = data.GetData(DATAFMT_VIRTUAL_LIST)
                If TypeOf obj Is IEnumerable(Of String) Then
                    For Each s In DirectCast(obj, IEnumerable(Of String))
                        If Not String.IsNullOrWhiteSpace(s) Then out.Add(s.Trim())
                    Next
                ElseIf TypeOf obj Is String Then
                    Dim lines = CStr(obj).Split({ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
                    For Each s In lines
                        Dim p = s.Trim()
                        If p.Length > 0 Then out.Add(p)
                    Next
                End If
            End If

            ' Standard FileDrop
            If data.GetDataPresent(DataFormats.FileDrop) Then
                For Each f As String In DirectCast(data.GetData(DataFormats.FileDrop), String())
                    If Not String.IsNullOrWhiteSpace(f) Then out.Add(f.Trim())
                Next
            End If

            ' Fallback: text
            If data.GetDataPresent(DataFormats.UnicodeText) AndAlso out.Count = 0 Then
                Dim t As String = CStr(data.GetData(DataFormats.UnicodeText))
                If Not String.IsNullOrWhiteSpace(t) Then
                    For Each line In t.Split({ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
                        Dim p = line.Trim().Trim(""""c)
                        If p.Length > 0 Then out.Add(p)
                    Next
                End If
            End If
        Catch ex As Exception
            Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try
        Return out.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function





End Class
