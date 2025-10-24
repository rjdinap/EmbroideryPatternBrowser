' requires SQLite
' Install-Package Microsoft.Data.Sqlite
' Install-Package SQLitePCLRaw.bundle_e_sqlite3
' add reference to System.Management, system,io.compression, system.io

' - on the installer, update the Version every time for a new release. it will ask if you want to update the product code - select yes
'  in git, Updating the Single “latest” release Each time
'  Rebuild your installer locally.
'  in git, Go to the Latest release, click Edit, Delete the old asset, Upload the New .exe/.msi, Publish.
'  git: You don't need new tags/versions—just keep overwriting that one release.


Option Strict On
Option Explicit On

Imports System.IO
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Microsoft.WindowsAPICodePack.Dialogs

Public Class Form1


    Public databaseName As String = GetDatabaseToUse()
    Private loadCts As CancellationTokenSource ' (reserved for future use)
    Private usbBrowser As UsbFileBrowser
    ' ---- Logging state ----
    Private Shared ReadOnly _logLock As New Object()
    Private _logPath As String = Nothing
    Private _logInitialized As Boolean = False
    Private Const LOG_MAX_BYTES As Integer = 1024 * 1024 ' 1 MB
    ' Form1.vb (inside the class)
    Private Const HelpPdfName As String = "EmbroideryHelp.pdf"
    Private _isShuttingDown As Boolean = False



    ' ==== Startup ====
    Private Sub EmbroideryPatternBrowser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        StatusProgress.ShowPopup(status:="Setting up...", indeterminate:=True)
        InitializeLogFile()
        Me.WindowState = FormWindowState.Maximized

        ' Hide legacy left panel UI
        Panel_Main_Left.Hide()

        ' USB-only browser control
        usbBrowser = New UsbFileBrowser() With {.Dock = DockStyle.Fill}
        Panel_Right_Bottom_Fill.Controls.Add(usbBrowser)
        usbBrowser.NavigateToRoot()

        Dim ver As String = ""
        Try
            ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
        Catch
            ver = "unknown"
        End Try

        'handle any uncaught exceptions
        AddHandler Application.ThreadException, Sub(sender2, args)
                                                    Try : Status("Error: " & args.Exception.Message, args.Exception.ToString()) : Catch : End Try
                                                End Sub

        AddHandler AppDomain.CurrentDomain.UnhandledException, Sub(sender2, args)
                                                                   Dim ex = TryCast(args.ExceptionObject, Exception)
                                                                   If ex IsNot Nothing Then
                                                                       Try : Status("Error: " & ex.Message, ex.ToString()) : Catch : End Try
                                                                   Else
                                                                       Try : Status("Error: Unhandled exception (no details).") : Catch : End Try
                                                                   End If
                                                               End Sub
        'close the database on application exit
        AddHandler Application.ApplicationExit,
        Sub(_s, _e)
            Try : SQLiteOperations.CloseSQLite() : Catch : End Try
        End Sub

        ' Make the preview always fit inside the box
        With PictureBox_FullImage
            .Dock = DockStyle.Fill
            .AutoSize = False
            .SizeMode = PictureBoxSizeMode.Zoom
            .Margin = New Padding(0)
            .Padding = New Padding(0)
            .BackColor = Color.White
        End With
        PictureBox_FullImage.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox_FullImage.BackColor = Color.White   ' optional, nicer look behind transparent areas

        Status("Welcome to EmbroideryPatternBrowser! : Version " & ver & vbCrLf)

        FastFileScanner.ReportStatus = AddressOf Me.Status
    End Sub


    ' Kick off the DB open only after the form has actually shown & painted.
    Private Async Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' Let the form finish its first paint & layout before we start work
        Await Task.Yield()

        ' Make sure the popup is up and can paint/animate
        Try
            StatusProgress.SetStatus("Opening database…")
            StatusProgress.SetIndeterminate(True)
            Me.UseWaitCursor = True
            Me.Update() ' flush one more paint cycle
        Catch
        End Try

        Dim dbPath As String = databaseName

        Try
            ' Run heavy DB open OFF the UI thread
            Await Task.Run(Sub()
                               ' Do NOT touch UI directly in here.
                               SQLiteOperations.InitializeSQLite(dbPath)

                               ' If you want progress from inside InitializeSQLite,
                               ' expose callbacks there and call BeginInvoke(...) here.
                           End Sub)

            ' Back on UI thread: update UI
            Status("Database ready: " & databaseName & ". Contains: " & SQLiteOperations.CountRows & " patterns.")

        Catch ex As Exception
            Status("Error: " & ex.Message, ex.ToString())
        Finally
            Try : StatusProgress.ClosePopup() : Catch : End Try
            Me.UseWaitCursor = False
        End Try

        'set the database name on the form
        TextBox_Database.Text = databaseName


    End Sub





    ' ==== Help -> About ====
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Try
            AboutForm.ShowDialog()
        Catch ex As Exception
            Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub



    ' ==== Search ====
    Private Sub Button_Search_Click(sender As Object, e As EventArgs) Handles Button_Search.Click
        Try
            ' time the search only
            Dim sw As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim dt = SQLiteOperations.SearchParsed(TextBox_Search.Text, ComboBox_Filter.Text)
            sw.Stop()

            Status(String.Format("Search completed in {0} ms. Keyword: ""{1}"". Rows: {2}",
                            sw.ElapsedMilliseconds, TextBox_Search.Text, If(dt IsNot Nothing, dt.Rows.Count, 0)))

            Dim grid As New VirtualThumbGrid() With {.Dock = DockStyle.Fill}
            Panel_Main_Fill.Controls.Clear()
            Panel_Main_Fill.Controls.Add(grid)

            grid.Bind(dt, AddressOf CreateImageFromPatternWrapper, PictureBox_FullImage, AddressOf SaveMetadataToDb)
            grid.Focus()
        Catch ex As Exception
            Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub



    ' ==== Rotate full-size preview ====
    Private Sub Button_Rotate_Click(sender As Object, e As EventArgs) Handles Button_Rotate.Click
        Try
            If PictureBox_FullImage.Image IsNot Nothing Then
                PictureBox_FullImage.Image = GraphicsFunctions.RotateImage(PictureBox_FullImage.Image, 90.0F)
            End If
        Catch ex As Exception
            Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub


    'check database health
    ' Form1.vb
    Private Async Sub CheckDatabaseHealthToolStripMenuItem_Click(sender As Object, e As EventArgs) _
    Handles CheckDatabaseHealthToolStripMenuItem.Click

        Dim msg As String =
"Database maintenance

This allows you to check the health of your database.

If searches that used to take ~1 second now take ~5 seconds, run this check. The scan may take seconds to minutes depending on DB size/fragmentation and disk speed.

When it finishes, a message will appear at the bottom indicating whether optimization is recommended.

Press OK to run, or Cancel to abort."

        If MessageBox.Show(Me, msg, "Check database health",
                       MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) <> DialogResult.OK Then
            Return
        End If

        StatusProgress.ShowPopup(status:="Checking database health…", indeterminate:=True)

        Dim needsOptimize As Boolean = False
        Try
            ' Run the heavy check off the UI thread
            needsOptimize = Await Task.Run(Function()
                                               Return SQLMaintenance.CheckDatabaseHealth(databaseName, "files_fts")
                                           End Function)
        Catch ex As Exception
            Status("Error checking database health: " & ex.Message, ex.StackTrace)
        Finally
            StatusProgress.ClosePopup()
        End Try

        Status("Database needs optimization: " & needsOptimize.ToString())
        If needsOptimize Then
            Status("Your database needs optimization. Please use Tools → Maintenance → Optimize Search Index to improve performance.",
               textColor:=Color.Red)
        End If
    End Sub



    ' ==== Thumbnail render wrapper for grid ====
    Private Function CreateImageFromPatternWrapper(path As String, w As Integer, h As Integer) As Image
        Try
            Dim pat As EmbPattern = EmbReaderFactory.LoadPattern(path)
            ' Conservative defaults; no heavy effects by default
            Return pat.CreateImageFromPattern(w:=w, h:=h, margin:=20, densityShading:=False, shadingStrength:=1.0)
        Catch ex As Exception
            ' Return a tiny white placeholder and log a concise error
            Status("Error " & ex.Message, ex.StackTrace.ToString)
            Dim bmp As New Bitmap(Math.Max(1, w), Math.Max(1, h))
            Using g = Graphics.FromImage(bmp)
                g.Clear(Color.White)
            End Using
            Return bmp
        End Try
    End Function


    ' ==== Close DB ====
    Private Async Sub closedatabase() Handles CloseDatabaseToolStripMenuItem.Click

        StatusProgress.ShowPopup(status:="Setting up...", indeterminate:=True)

        ' Let the form finish its first paint & layout before we start work
        Await Task.Yield()

        ' Make sure the popup is up and can paint/animate
        Try
            StatusProgress.SetStatus("Closing database…")
            StatusProgress.SetIndeterminate(True)
            Me.UseWaitCursor = True
            Me.Update() ' flush one more paint cycle
        Catch
        End Try

        Try
            ' Run heavy DB open OFF the UI thread
            Await Task.Run(Sub()
                               ' Do NOT touch UI directly in here.
                               SQLiteOperations.CloseSQLite()
                               ' If you want progress from inside InitializeSQLite,
                               ' expose callbacks there and call BeginInvoke(...) here.
                           End Sub)

            ' Back on UI thread: update UI
            Status("Database closed.")
        Catch ex As Exception
            Status("Error closing database " & ex.Message, ex.ToString())
        Finally
            Try : StatusProgress.ClosePopup() : Catch : End Try
            Me.UseWaitCursor = False
        End Try

        databaseName = ""
        TextBox_Database.Text = databaseName

    End Sub



    ' ==== Create DB ====
    Private Sub CreateNewDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateNewDatabaseToolStripMenuItem1.Click

        Try
            Status("Create New database()")
            Dim dirtyName As String = InputBox("Enter a name For your database")
            If String.ReferenceEquals(dirtyName, String.Empty) Then Return ' this is how we check for if the user pressed cancel

            Dim safeBase As String = Regex.Replace(dirtyName, "[^A-Za-z0-9\-_]", "")
            If String.IsNullOrWhiteSpace(safeBase) OrElse safeBase.Length < 3 Then
                MessageBox.Show("The database name you entered Is blank Or has invalid characters. Try again With a different name.")
                Return
            End If

            Dim cleanName As String = safeBase & ".sqlite"
            databaseName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), cleanName)

            If File.Exists(databaseName) Then
                MessageBox.Show("This database name " & databaseName & " already exists. Please choose another name.")
                Return
            End If

            ' Create/open DB
            Status("Creating New database " + databaseName)
            SQLiteOperations.InitializeSQLite(databaseName)


        Catch ex As Exception
            Status("Error creating New database " & ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub




    ' In Form1: ensure we always close the DB when the user clicks the X
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        Dim r = usbBrowser.RequestAppClose(e)  ' replace with your instance name

        If r = DialogResult.Yes Then
            If _isShuttingDown Then Exit Sub
            _isShuttingDown = True
            StatusProgress.ShowPopup(status:="Shutting down, hang On a moment.", indeterminate:=True)
            Try
                ' Fail-safe: close SQLite (stops insert thread, releases handles)
                SQLiteOperations.CloseSQLite()

                ' Close any progress UI first (best-effort)
                Try : StatusProgress.ClosePopup() : Catch : End Try

            Catch ex As Exception
                ' Log a short message on-screen and full details to disk (per your Status signature)
                Try : Status("Error " & ex.Message, ex.ToString()) : Catch : End Try
            End Try
        End If
    End Sub



    'GetDatabaseToUse
    'use the the last database opened. or the default if one doesn't exist
    Private Function GetDatabaseToUse() As String
        Dim defaultDbPath As String = Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "EmbroideryPatternBrowser.sqlite")
        Dim chosen As String = If(My.Settings.LastDatabasePath, "").Trim()
        If Not File.Exists(chosen) Then
            chosen = defaultDbPath
        End If
        Return chosen
    End Function



    Public Function HexDump(bytes As Byte(), Optional width As Integer = 16) As String
        Dim sb As New StringBuilder()
        For i = 0 To bytes.Length - 1 Step width
            Dim slice = bytes.Skip(i).Take(Math.Min(width, bytes.Length - i)).ToArray()
            sb.Append(i.ToString("X6")).Append(":  ")
            ' hex
            For j = 0 To width - 1
                If j < slice.Length Then
                    sb.Append(slice(j).ToString("X2")).Append(" ")
                Else
                    sb.Append("   ")
                End If
            Next
            sb.Append(" | ")
            ' ascii
            For j = 0 To slice.Length - 1
                Dim ch = AscW(ChrW(slice(j)))
                If ch >= 32 AndAlso ch < 127 Then
                    sb.Append(ChrW(ch))
                Else
                    sb.Append(".")
                End If
            Next
            sb.AppendLine()
        Next
        Return sb.ToString()
    End Function




    'Set up log file on startup
    Private Sub InitializeLogFile()
        Try

            Dim docs As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            _logPath = IO.Path.Combine(docs, "EmbroideryPatternBrowser.log")

            ' Ensure file exists
            If Not IO.File.Exists(_logPath) Then
                Using fs = IO.File.Create(_logPath)
                End Using
                _logInitialized = True
                Return
            End If

            ' Truncate to last 1MB (keep most recent) once per startup
            Dim fi As New IO.FileInfo(_logPath)
            If fi.Length > LOG_MAX_BYTES Then
                Dim buf(LOG_MAX_BYTES - 1) As Byte
                Dim read As Integer
                Using fs As New IO.FileStream(_logPath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.ReadWrite)
                    ' Seek to last 1MB and read forward
                    fs.Seek(-LOG_MAX_BYTES, IO.SeekOrigin.End)
                    read = fs.Read(buf, 0, buf.Length)
                End Using
                Using ws As New IO.FileStream(_logPath, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read)
                    ws.Write(buf, 0, read)
                End Using
            End If

            _logInitialized = True
            Status("Initializing log file " + _logPath)
        Catch
            ' Swallow; logging should never crash the UI
            _logInitialized = False
        End Try
    End Sub



    ' ==== Open DB ====
    Private Sub OpenDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenDatabaseToolStripMenuItem.Click
        Try
            Using openFileDialog As New OpenFileDialog() With {
                .Filter = "Embroidery Database Files|*.sqlite",
                .Title = "Select an Embroidery Database File",
                .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                .CheckFileExists = True,
                .Multiselect = False
            }
                If openFileDialog.ShowDialog() = DialogResult.OK Then
                    databaseName = openFileDialog.FileName
                    SQLiteOperations.InitializeSQLite(databaseName)
                    Status("Database ready " & databaseName & ". Contains: " & SQLiteOperations.CountRows & " patterns.")
                    My.Settings.LastDatabasePath = databaseName
                    My.Settings.Save()
                End If
            End Using
        Catch ex As Exception
            Status("Error " & ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub


    'optimize database
    Private Sub OptimizeDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OptimizeDatabaseToolStripMenuItem.Click
        Dim msg As String = "Database maintenance" & vbCrLf & vbCrLf & "This Is Not a normal user Option, And should only be performed if the program has alerted you to perform maintenance." & vbCrLf & vbCrLf
        msg = msg + "This operation will optimize the search index, which may take anywhere from several seconds to several hours depending on database size / fragmentation level / speed of your hard drive." & vbCrLf & vbCrLf
        msg = msg + "Press OK to perform the maintenance, or Cancel to abort."

        Dim res = MessageBox.Show(Me, msg, "Optimize Search Index", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)
        If res <> DialogResult.OK Then Exit Sub

        ' Fire the operation (the method itself shows a progress popup)
        SQLiteOperations.OptimizeIndex()
        Status("Index Optimization complete.")
    End Sub




    ' ==== Grid callback to persist metadata ====
    Private Function SaveMetadataToDb(row As DataRow, newMetadata As String) As Boolean
        Try
            Return SQLiteOperations.UpdateMetadataRow(row, newMetadata)
        Catch ex As Exception
            Status("Error " & ex.Message, ex.StackTrace.ToString)
            Return False
        End Try
    End Function



    'Scan for Images
    Private Sub ScanForImagesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ScanForImagesToolStripMenuItem.Click

        Dim folderPath As String = ""

        ' Pick folder to scan
        Using dlg As New CommonOpenFileDialog()
            dlg.IsFolderPicker = True
            dlg.Title = "Select a folder To scan For embroidery patterns"
            dlg.EnsurePathExists = True

            ' Start location (optional)
            dlg.InitialDirectory = Environment.SpecialFolder.MyComputer.ToString
            dlg.DefaultDirectory = Environment.SpecialFolder.MyComputer.ToString
            ' Pin special folders at the top of the nav pane
            'dlg.AddPlace(KnownFolders.Desktop.Path, FileDialogAddPlaceLocation.Top)
            'dlg.AddPlace(KnownFolders.Documents.Path, FileDialogAddPlaceLocation.Top)

            If dlg.ShowDialog() <> CommonFileDialogResult.Ok Then
                Return
            End If
            folderPath = dlg.FileName

            If String.IsNullOrWhiteSpace(folderPath) OrElse Not Directory.Exists(folderPath) Then
                MessageBox.Show("Invalid folder.")
                Return
            End If

            ' Inform user about skipped system folders (mirrors FastFileScanner policy)
            Try
                Status("Note \Windows And \Program Files are Not scanned.")
            Catch
            End Try

            Dim dirs As New List(Of String) From {folderPath}

            StatusProgress.ShowPopup(status:="Scanning directory information…", indeterminate:=True)
            Dim t As New Threading.Thread(
    Sub()
        Try
            Dim result = FastFileScanner.ScanAll(dirs, databaseName,
                Sub(status As String, current As Integer, total As Integer)
                    ' (keep your popup updates as-is; they already marshal internally)
                    StatusProgress.SetStatus(String.Format("{0} ({1} Of {2})", status, current, total))
                    If total <= 0 Then
                        StatusProgress.SetIndeterminate(True)
                    Else
                        StatusProgress.SetIndeterminate(False)
                        Dim cur = Math.Max(0, Math.Min(current, Math.Max(1, total)))
                        StatusProgress.SetProgress(cur, maximum:=total)
                    End If
                End Sub)

            ' After the scan completes, post a single UI message:
            Me.BeginInvoke(Sub()
                               Status(String.Format("Scan finished {0}", result.ToString()))
                               StatusProgress.ClosePopup()
                           End Sub)

            'we need to check for optimization after a large scan
            If (result.FilesAdded + result.FilesDeleted > 40000) Then
                SQLiteOperations.OptimizeIndex()
            End If


        Catch ex As Exception
            Me.BeginInvoke(Sub()
                               Status("Error during scan " & ex.Message, ex.StackTrace)
                               StatusProgress.ClosePopup()
                           End Sub)
        End Try
    End Sub)
            t.IsBackground = True
            t.Start()
        End Using

        Status("Database contains: " & SQLiteOperations.CountRows & " patterns.")


    End Sub



    ' Tools -> Options
    Private Sub SettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem.Click
        ' capture old size to detect change
        Dim oldW As Integer = My.Settings.ThumbWidth
        Dim oldH As Integer = My.Settings.ThumbHeight

        Using dlg As New OptionsDialog()
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                Dim changed As Boolean = (oldW <> My.Settings.ThumbWidth) OrElse (oldH <> My.Settings.ThumbHeight)

                ' find the current grid instance hosted in Panel_Main_Fill (you create it dynamically)
                Dim grid As VirtualThumbGrid = Nothing
                For Each ctl As Control In Panel_Main_Fill.Controls
                    grid = TryCast(ctl, VirtualThumbGrid)
                    If grid IsNot Nothing Then Exit For
                Next

                If grid IsNot Nothing Then
                    grid.ReloadThumbSettings(regenerate:=changed)
                End If
            End If
        End Using
    End Sub





    'Help
    Private Sub ShowHelpFIleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowHelpFIleToolStripMenuItem.Click
        Try
            Dim pdfPath = Path.Combine(Application.StartupPath, "Help", HelpPdfName)
            If Not File.Exists(pdfPath) Then
                Status("Help PDF Not found " & pdfPath)
                MessageBox.Show("Help file Not found. I'll open the Help folder so you can drop it in.",
                            "Help", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' Open the Help folder (select if exists)
                Dim helpFolder = Path.Combine(Application.StartupPath, "Help")
                If Directory.Exists(helpFolder) Then
                    Process.Start(New ProcessStartInfo("explorer.exe", "/e,""" & helpFolder & """") With {.UseShellExecute = True})
                Else
                    Directory.CreateDirectory(helpFolder)
                    Process.Start(New ProcessStartInfo("explorer.exe", "/e,""" & helpFolder & """") With {.UseShellExecute = True})
                End If
                Return
            End If

            ' Open with default PDF viewer
            Process.Start(New ProcessStartInfo(pdfPath) With {.UseShellExecute = True})

        Catch ex As Exception
            ' Per your logging rules: simple message to screen, full stack to disk
            Status("Error: Couldn't open help PDF: " & ex.Message, ex.ToString())
            MessageBox.Show("Couldn't open help PDF: " & ex.Message,
                        "Help", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



    ' Logs to the on-screen status AND to disk.
    ' Pass stackTrace ONLY for the disk log (kept out of the UI).
    ' Optional textColor lets you color just this message (default = black).
    Public Sub Status(ByVal msg As String,
                  Optional ByVal stackTrace As String = Nothing,
                  Optional ByVal textColor As Color = Nothing)

        ' --- UI log ---
        If Not Me.IsHandleCreated Then Return

        If Me.InvokeRequired Then
            ' hop to UI thread
            Me.BeginInvoke(CType(Sub() SetStatusUI(msg, textColor), MethodInvoker))
        Else
            SetStatusUI(msg, textColor)
        End If

        ' --- Disk log ---
        Try
            If Not _logInitialized OrElse String.IsNullOrEmpty(_logPath) Then
                InitializeLogFile()
            End If
            If String.IsNullOrEmpty(_logPath) Then Exit Sub

            Dim line As String = String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}  {1}", DateTime.Now, msg)

            SyncLock _logLock
                IO.File.AppendAllText(_logPath, line & Environment.NewLine)
                If Not String.IsNullOrEmpty(stackTrace) Then
                    IO.File.AppendAllText(_logPath, "    Stack: " & stackTrace & Environment.NewLine)
                End If
            End SyncLock
        Catch
            ' Never throw from logging
        End Try
    End Sub 'Status



    ' Append 1 line to the status box (optionally colored for this line only).
    Private Sub SetStatusUI(msg As String, Optional textColor As Color = Nothing)
        Try
            ' Move caret to end
            RichTextBox_Status.SelectionStart = Len(RichTextBox_Status.Text)
            RichTextBox_Status.SelectionLength = 0

            ' Choose color (default = black) for this insertion only
            Dim useColor As Color = If(textColor.IsEmpty, Color.Black, textColor)
            RichTextBox_Status.SelectionColor = useColor

            ' Prepend a newline like before, then append text
            RichTextBox_Status.SelectedText = vbCrLf & msg

            ' Reset selection color back to black for future appends
            RichTextBox_Status.SelectionColor = Color.Black

            ' Scroll to bottom
            RichTextBox_Status.SelectionStart = Len(RichTextBox_Status.Text)
            RichTextBox_Status.ScrollToCaret()
        Catch
            ' UI is best-effort; ignore
        End Try
    End Sub



    ' ==== Enter-to-search ====
    Private Sub TextBox_Search_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Search.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button_Search_Click(Nothing, Nothing)
        End If
    End Sub

    'when we enter the text box field, clear anything in there.
    Private Sub TextBox_Search_Enter(ByVal sender As System.Object, ByVal e As EventArgs) Handles TextBox_Search.Enter
        TextBox_Search.Text = ""
    End Sub

    ' 1b) Public, static-like entry point for background threads

    Public Sub StatusFromAnyThread(text As String)
        Dim f As Form1 =
            Application.OpenForms.OfType(Of Form1)().FirstOrDefault()

        If f Is Nothing OrElse f.IsDisposed Then Return
        f.InvokeIfRequired(Sub() f.Status(text))   ' calls your existing instance Status()
    End Sub



End Class



' 1a) Tiny helper to marshal to UI
Module ControlInvokeExtensions
    <Extension()>
    Public Sub InvokeIfRequired(ctrl As Control, action As Action)
        If ctrl Is Nothing OrElse ctrl.IsDisposed Then Return

        If ctrl.IsHandleCreated Then
            If ctrl.InvokeRequired Then
                Try
                    ctrl.BeginInvoke(action)
                Catch
                    ' Control is closing—ignore.
                End Try
            Else
                action()
            End If
        Else
            ' Wait until the handle is created, then run once.
            Dim h As EventHandler = Nothing
            h = Sub(sender As Object, e As EventArgs)
                    RemoveHandler ctrl.HandleCreated, h
                    If ctrl.IsDisposed Then Return
                    Try
                        ctrl.BeginInvoke(action)
                    Catch
                        ' Control is closing—ignore.
                    End Try
                End Sub
            AddHandler ctrl.HandleCreated, h
        End If
    End Sub
End Module



