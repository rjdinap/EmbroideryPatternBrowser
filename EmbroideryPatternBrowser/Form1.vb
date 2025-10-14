' requires SQLite
' Install-Package Microsoft.Data.Sqlite
' Install-Package SQLitePCLRaw.bundle_e_sqlite3
' add reference to System.Management
'add nuget package: microsoft.web.webview2
'on the installer, update the Version every time for a new release. it will ask if you want to update the product code - select yes

Option Strict On
Option Explicit On

Imports System.Data ' DataTable
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Diagnostics

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
    Private Sub SewingPatternOrganizer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitializeLogFile()   ' <-- add this line
        Me.WindowState = FormWindowState.Maximized

        ' Hide legacy left panel UI
        Panel_Main_Left.Hide()

        ' USB-only browser control
        usbBrowser = New UsbFileBrowser() With {.Dock = DockStyle.Fill}
        Panel_Right_Bottom.Controls.Add(usbBrowser)
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
        AddHandler Application.ApplicationExit,
        Sub(_s, _e)
            Try : SQLiteOperations.CloseSQLite() : Catch : End Try
        End Sub

        ' Make the preview always fit inside the box

        With PictureBox_FullImage
            .Dock = DockStyle.Fill
            .AutoSize = False         ' <- critical: never resize to image
            .SizeMode = PictureBoxSizeMode.Zoom
            .Margin = New Padding(0)  ' avoid surprise shrink
            .Padding = New Padding(0)
            .BackColor = Color.White
        End With
        PictureBox_FullImage.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox_FullImage.BackColor = Color.White   ' optional, nicer look behind transparent areas

        Status("Welcome to EmbroideryPatternBrowser! : Version " & ver & vbCrLf)
        SQLiteOperations.InitializeSQLite(databaseName) ' open our default database

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



    ' ==== Thumbnail render wrapper for grid ====
    Private Function CreateImageFromPatternWrapper(path As String, w As Integer, h As Integer) As Image
        Try
            Dim pat As EmbPattern = EmbReaderFactory.LoadPattern(path)
            ' Conservative defaults; no heavy effects by default
            Return pat.CreateImageFromPattern(w:=w, h:=h, margin:=20, densityShading:=False, shadingStrength:=1.0)
        Catch ex As Exception
            ' Return a tiny white placeholder and log a concise error
            Status("Error: " & ex.Message, ex.StackTrace.ToString)
            Dim bmp As New Bitmap(Math.Max(1, w), Math.Max(1, h))
            Using g = Graphics.FromImage(bmp)
                g.Clear(Color.White)
            End Using
            Return bmp
        End Try
    End Function



    ' ==== Create DB ====
    Private Sub CreateNewDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateNewDatabaseToolStripMenuItem.Click

        Try
            Status("Create new database()")
            Dim dirtyName As String = InputBox("Enter a name for your database")
            If String.ReferenceEquals(dirtyName, String.Empty) Then Return ' this is how we check for if the user pressed cancel

            Dim safeBase As String = Regex.Replace(dirtyName, "[^A-Za-z0-9\-_]", "")
            If String.IsNullOrWhiteSpace(safeBase) OrElse safeBase.Length < 3 Then
                MessageBox.Show("The database name you entered is blank or has invalid characters. Try again with a different name.")
                Return
            End If

            Dim cleanName As String = safeBase & ".sqlite"
            databaseName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), cleanName)

            If File.Exists(databaseName) Then
                MessageBox.Show("This database name: " & databaseName & " already exists. Please choose another name.")
                Return
            End If

            ' Create/open DB
            Status("Creating new database: " + databaseName)
            SQLiteOperations.InitializeSQLite(databaseName)


        Catch ex As Exception
            Status("Error creating new database: " & ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub





    ' In Form1: ensure we always close the DB when the user clicks the X
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If _isShuttingDown Then Exit Sub
        _isShuttingDown = True
        Try
            ' Close any progress UI first (best-effort)
            Try : StatusProgress.ClosePopup() : Catch : End Try

            ' Fail-safe: close SQLite (stops insert thread, releases handles)
            SQLiteOperations.CloseSQLite()
        Catch ex As Exception
            ' Log a short message on-screen and full details to disk (per your Status signature)
            Try : Status("Error: " & ex.Message, ex.ToString()) : Catch : End Try
        End Try
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
            Status("Initializing log file: " + _logPath)
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
                    My.Settings.LastDatabasePath = databaseName
                    My.Settings.Save()
                End If
            End Using
        Catch ex As Exception
            Status("Error: " & ex.Message, ex.StackTrace.ToString)
        End Try
    End Sub




    ' ==== Grid callback to persist metadata ====
    Private Function SaveMetadataToDb(row As DataRow, newMetadata As String) As Boolean
        Try
            Return SQLiteOperations.UpdateMetadataRow(row, newMetadata)
        Catch ex As Exception
            Status("Error: " & ex.Message, ex.StackTrace.ToString)
            Return False
        End Try
    End Function



    'Scan for Images
    Private Sub ScanForImagesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ScanForImagesToolStripMenuItem1.Click

        ' Pick folder to scan
        Using fb As New FolderBrowserDialog()
            fb.Description = "Select a folder to scan for embroidery patterns"
            fb.ShowNewFolderButton = False
            fb.RootFolder = Environment.SpecialFolder.MyComputer
            'if the user presses cancel, delete the database that we just created above
            If fb.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            Dim folderPath As String = fb.SelectedPath
            If String.IsNullOrWhiteSpace(folderPath) OrElse Not Directory.Exists(folderPath) Then
                MessageBox.Show("Invalid folder.")
                Return
            End If

            ' Inform user about skipped system folders (mirrors FastFileScanner policy)
            Try
                Status("Note: \Windows and \Program Files are not scanned.")
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
                    StatusProgress.SetStatus(String.Format("{0} ({1} of {2})", status, current, total))
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
                               Status(String.Format("Scan finished: {0}", result.ToString()))
                               StatusProgress.ClosePopup()
                           End Sub)

        Catch ex As Exception
            Me.BeginInvoke(Sub()
                               Status("Error during scan: " & ex.Message, ex.StackTrace)
                               StatusProgress.ClosePopup()
                           End Sub)
        End Try
    End Sub)
            t.IsBackground = True
            t.Start()
        End Using


    End Sub




    Private Sub ShowHelpFIleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowHelpFIleToolStripMenuItem.Click
        Try
            Dim pdfPath = Path.Combine(Application.StartupPath, "Help", HelpPdfName)
            If Not File.Exists(pdfPath) Then
                Status("Help PDF not found: " & pdfPath)
                MessageBox.Show("Help file not found. I’ll open the Help folder so you can drop it in.",
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
    Public Sub Status(ByVal msg As String, Optional ByVal stackTrace As String = Nothing)
        ' --- UI log ---
        Try
            RichTextBox_Status.AppendText(vbCrLf & msg)
            RichTextBox_Status.SelectionStart = Len(RichTextBox_Status.Text)
            RichTextBox_Status.ScrollToCaret()
        Catch
            ' UI is best-effort; ignore
        End Try

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



    ' ==== Enter-to-search ====
    Private Sub TextBox_Search_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Search.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button_Search_Click(Nothing, Nothing)
        End If
    End Sub

    'when we enter the text box field, clear anything in there.
    Private Sub TextBox_Search_KeyDown(ByVal sender As System.Object, ByVal e As EventArgs) Handles TextBox_Search.Enter
        TextBox_Search.Text = ""
    End Sub



End Class






