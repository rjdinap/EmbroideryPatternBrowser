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
    Private Const HelpPdfName As String = "EmbroideryHelp.pdf"
    Private _isShuttingDown As Boolean = False
    Private _dlg As StitchDisplay



    ' ==== Startup ====
    Private Sub EmbroideryPatternBrowser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        StatusProgress.ShowPopup(status:="Setting up...", indeterminate:=True)
        Me.WindowState = FormWindowState.Maximized

    End Sub


    ' Kick off the DB open only after the form has actually shown & painted.
    Private Async Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        ' Let the form finish its first paint & layout before we start work
        Await Task.Yield()

        Logger.Initialize() ' uses Documents\EmbroideryPatternBrowser\logs\EmbroideryPatternBrowser-YYYYMMDD.log
        Logger.AttachUI(Me, RichTextBox_Status) ' your status RichTextBox


        'handle any uncaught exceptions
        AddHandler Application.ThreadException, Sub(sender2, args)
                                                    Try : Logger.Info("Error: " & args.Exception.Message, args.Exception.ToString()) : Catch : End Try
                                                End Sub

        AddHandler AppDomain.CurrentDomain.UnhandledException, Sub(sender2, args)
                                                                   Dim ex = TryCast(args.ExceptionObject, Exception)
                                                                   If ex IsNot Nothing Then
                                                                       Try : Logger.Error(ex.Message, ex.ToString()) : Catch : End Try
                                                                   Else
                                                                       Try : Logger.Error("Unhandled exception (no details).") : Catch : End Try
                                                                   End If
                                                               End Sub
        'close the database on application exit
        AddHandler Application.ApplicationExit,
        Sub(_s, _e)
            Try : SQLiteOperations.CloseSQLite() : Catch : End Try
        End Sub


        ' USB-only browser control
        usbBrowser = New UsbFileBrowser() With {.Dock = DockStyle.Fill}
        Panel_Right_Bottom_Fill.Controls.Add(usbBrowser)

        usbBrowser.NavigateToRoot()

        'no one is going to see this anyway until they can search
        _dlg = New StitchDisplay(
        Panel_TabPage2_Fill_Right,
        Panel_TabPage2_Fill_Bottom,
        ZoomPictureBox_Stitches)
        _dlg.BuildUI()

        'open database
        ' Make sure the popup is up and can paint/animate
        Try
            StatusProgress.SetStatus("Opening database…")
            StatusProgress.SetIndeterminate(True)
            Me.UseWaitCursor = True
            Me.Update() ' flush one more paint cycle
        Catch
        End Try

        Try
            ' Run heavy DB open OFF the UI thread
            Await Task.Run(Sub()
                               SQLiteOperations.InitializeSQLite(databaseName)
                           End Sub)
            Logger.Info("Database ready: " & databaseName)

        Catch ex As Exception
            Logger.Error(ex.Message, ex.ToString())
        Finally
            Try : StatusProgress.ClosePopup() : Catch : End Try
            Me.UseWaitCursor = False
        End Try

        'set the database name on the form
        TextBox_Database.Text = databaseName


        Logger.Noprefix("Welcome to EmbroideryPatternBrowser! : Version " & My.Application.Info.Version.ToString & vbCrLf)

    End Sub





    ' ==== Help -> About ====
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Try
            AboutForm.ShowDialog()
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub



    ' ==== Search ====
    Private Sub Button_Search_Click(sender As Object, e As EventArgs) Handles Button_Search.Click
        Try
            ' time the search only
            Dim sw As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim dt = SQLiteOperations.SearchParsed(TextBox_Search.Text, ComboBox_Filter.Text)
            sw.Stop()

            Logger.Info(String.Format("Search completed in {0} ms. Keyword: ""{1}"". Rows: {2}",
                            sw.ElapsedMilliseconds, TextBox_Search.Text, If(dt IsNot Nothing, dt.Rows.Count, 0)))

            Dim grid As New VirtualThumbGrid() With {.Dock = DockStyle.Fill}
            Panel_Main_Fill.Controls.Clear()
            Panel_Main_Fill.Controls.Add(grid)

            grid.Bind(dt, AddressOf CreateImageFromPatternWrapper, PictureBox_FullImage, AddressOf SaveMetadataToDb, stitchDisplay:=_dlg)
            grid.Focus()
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub



    ' ==== Rotate full-size preview ====
    Private Sub Button_Rotate_Click(sender As Object, e As EventArgs) Handles Button_Rotate.Click
        Try
            If PictureBox_FullImage.Image IsNot Nothing Then
                PictureBox_FullImage.Image = GraphicsFunctions.RotateImage(PictureBox_FullImage.Image, 90.0F)
            End If
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub




    ' ==== Thumbnail render wrapper for grid ====
    Private Function CreateImageFromPatternWrapper(path As String, w As Integer, h As Integer) As Image
        Try
            Dim pat As EmbPattern = EmbReaderFactory.LoadPattern(path)
            ' Conservative defaults; no heavy effects by default
            Return pat.CreateImageFromPattern(w:=w, h:=h, margin:=20, densityShading:=False, shadingStrength:=1.0)
        Catch ex As Exception
            ' Return a tiny white placeholder and log a concise error
            Logger.Error(ex.Message, ex.StackTrace.ToString)
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
            Logger.Info("Database closed.")
        Catch ex As Exception
            Logger.Error("Error closing database " & ex.Message, ex.ToString())
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
            Logger.Info("Create New database()")
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
            Logger.Info("Creating New database " + databaseName)
            SQLiteOperations.InitializeSQLite(databaseName)


        Catch ex As Exception
            Logger.Error("Error creating New database " & ex.Message, ex.StackTrace.ToString)
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
                Try : Logger.Error(ex.Message, ex.ToString()) : Catch : End Try
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
                    Logger.Noprefix("Database ready " & databaseName)
                    My.Settings.LastDatabasePath = databaseName
                    My.Settings.Save()
                End If
            End Using
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub



    'optimize database
    Private Sub OptimizeDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OptimizeDatabaseToolStripMenuItem.Click
        Dim msg As String = "Database maintenance" & vbCrLf & vbCrLf & "This Is Not a normal user Option, And should only be performed If the program has alerted you To perform maintenance." & vbCrLf & vbCrLf
        msg = msg + "This operation will optimize the search index, which may take anywhere from several seconds To several hours depending On database size / fragmentation level / speed Of your hard drive." & vbCrLf & vbCrLf
        msg = msg + "Press OK To perform the maintenance, Or Cancel To abort."

        Dim res = MessageBox.Show(Me, msg, "Optimize Search Index", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)
        If res <> DialogResult.OK Then Exit Sub

        ' Fire the operation (the method itself shows a progress popup)
        SQLiteOperations.OptimizeIndex()
        Logger.Info("Index Optimization complete.")
    End Sub




    ' ==== Grid callback to persist metadata ====
    Private Function SaveMetadataToDb(row As DataRow, newMetadata As String) As Boolean
        Try
            Return SQLiteOperations.UpdateMetadataRow(row, newMetadata)
        Catch ex As Exception
            Logger.Error(ex.Message, ex.StackTrace.ToString)
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
                               Logger.Info(String.Format("Scan finished {0}", result.ToString()))
                               StatusProgress.ClosePopup()
                           End Sub)

            'we need to check to see if we need an optimization
            Dim delta As Long = result.FilesAdded + result.FilesUpdated + result.FilesDeleted
            SQLiteOperations.AddRowsSinceOptimize(delta)
            SQLiteOperations.QuickNeedsOptimize(200_000)


        Catch ex As Exception
            Me.BeginInvoke(Sub()
                               Logger.Error("Error during scan " & ex.Message, ex.StackTrace)
                               StatusProgress.ClosePopup()
                           End Sub)
        End Try
    End Sub)
            t.IsBackground = True
            t.Start()
        End Using

        'Logger.Info("Database contains: " & SQLiteOperations.CountRows & " patterns.")


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
                Logger.Error("Help PDF Not found " & pdfPath)
                MessageBox.Show("Help file Not found.",
                            "Help", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Open with default PDF viewer
            Process.Start(New ProcessStartInfo(pdfPath) With {.UseShellExecute = True})

        Catch ex As Exception
            ' Per your logging rules: simple message to screen, full stack to disk
            Logger.Error("Error: Couldn't open help PDF: " & ex.Message, ex.ToString())
            MessageBox.Show("Couldn't open help PDF: " & ex.Message,
                        "Help", MessageBoxButtons.OK, MessageBoxIcon.Error)
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


End Class



