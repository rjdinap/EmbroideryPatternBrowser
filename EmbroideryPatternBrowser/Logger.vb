Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Drawing

Public Enum LogLevel
    NONE = 0
    INFO = 1
    DEBUG = 2
    ERR = 3
End Enum

Public NotInheritable Class Logger
    Private Shared ReadOnly _sync As New Object()

    Private Shared _filePath As String
    Private Shared _writer As StreamWriter
    Private Shared _ui As RichTextBox
    Private Shared _uiForm As Form

    ' File gets DEBUG+ by default; screen shows INFO & ERROR only (DEBUG -> file only).
    Private Shared _minFileLevel As LogLevel = LogLevel.NONE

    Private Shared ReadOnly _queue As New Collections.Concurrent.ConcurrentQueue(Of String)
    Private Shared ReadOnly _flushSignal As New AutoResetEvent(False)
    Private Shared _bg As Thread
    Private Shared _running As Boolean

    Private Const DEFAULT_MAX_TRIM_BYTES As Long = 1L * 1024 * 1024 ' 1 MB

    ' ---------- Setup ----------
    Public Shared Sub Initialize(Optional logDirectory As String = Nothing,
                                 Optional fileName As String = "EmbroideryPatternBrowser.log",
                                 Optional maxBytesOnStartup As Long = DEFAULT_MAX_TRIM_BYTES)
        SyncLock _sync
            If String.IsNullOrWhiteSpace(logDirectory) Then
                logDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "EmbroideryPatternBrowser", "logs")
            End If
            Directory.CreateDirectory(logDirectory)
            _filePath = Path.Combine(logDirectory, fileName)

            TrimLogFileToMaxBytes(_filePath, maxBytesOnStartup)

            _writer = New StreamWriter(New FileStream(_filePath, FileMode.Append, FileAccess.Write, FileShare.Read)) With {
                .AutoFlush = True
            }

            _running = True
            _bg = New Thread(AddressOf FileFlushLoop) With {.IsBackground = True, .Name = "LoggerFlush"}
            _bg.Start()
        End SyncLock
    End Sub

    Public Shared Sub AttachUI(hostForm As Form, statusBox As RichTextBox)
        SyncLock _sync
            _uiForm = hostForm
            _ui = statusBox
        End SyncLock
    End Sub



    ' ---------- Main API ----------
    ' All sinks include the LEVEL: prefix (e.g., "INFO: message").
    Public Shared Sub Log(level As LogLevel,
                          message As String,
                          Optional stackTrace As String = Nothing,
                          Optional color As Color? = Nothing)

        Dim levelText As String = LevelName(level)
        Dim screenLine As String = ""
        If level = LogLevel.NONE Then
            screenLine = $"{message}"
        Else
            screenLine = $"{levelText}: {message}"
        End If

        ' File line includes timestamp plus the same LEVEL: prefix.
        Dim fileLine As New StringBuilder(256)
        fileLine.AppendFormat("{0:yyyy-MM-dd HH:mm:ss.fff} {1}: {2}", DateTime.Now, levelText, message)
        If Not String.IsNullOrEmpty(stackTrace) Then
            fileLine.AppendLine()
            fileLine.Append("  ").Append(stackTrace.Replace(Environment.NewLine, Environment.NewLine & "  "))
        End If

        'everything goes to disk
        If _writer IsNot Nothing Then
            _queue.Enqueue(fileLine.ToString())
            _flushSignal.Set()
        End If

        ' ---- Screen sink: INFO, NONE and ERROR only, with LEVEL: prefix ----
        If (level = LogLevel.INFO OrElse level = LogLevel.ERR OrElse level = LogLevel.NONE) AndAlso
           _ui IsNot Nothing AndAlso _uiForm IsNot Nothing AndAlso Not _ui.IsDisposed Then

            ' Default ERROR color if none supplied.
            Dim effectiveColor As Color? = color

            Dim apply As MethodInvoker =
                Sub()
                    Try
                        Dim prefix As String = If(_ui.TextLength > 0, Environment.NewLine, "")
                        Dim start = _ui.TextLength
                        _ui.AppendText(prefix & screenLine)

                        If effectiveColor.HasValue Then
                            _ui.Select(start, screenLine.Length)
                            _ui.SelectionColor = effectiveColor.Value
                            _ui.Select(_ui.TextLength, 0)
                            _ui.SelectionColor = _ui.ForeColor
                        End If

                        _ui.SelectionStart = _ui.TextLength
                        _ui.ScrollToCaret()
                    Catch
                        ' ignore transient UI errors
                    End Try
                End Sub

            If _ui.InvokeRequired Then
                _ui.BeginInvoke(apply)
            Else
                apply()
            End If
        End If
    End Sub

    'the same as info, but no prefix of a loglevel on the line printed
    Public Shared Sub Noprefix(message As String,
                           Optional stackTrace As String = Nothing,
                           Optional color As Color? = Nothing)
        Log(LogLevel.NONE, message, stackTrace, color)
    End Sub

    Public Shared Sub Info(message As String,
                           Optional stackTrace As String = Nothing,
                           Optional color As Color? = Nothing)
        Log(LogLevel.INFO, message, stackTrace, color)
    End Sub

    ' DEBUG -> file only (by policy); no color needed.
    Public Shared Sub Debug(message As String, Optional stackTrace As String = Nothing)
        Log(LogLevel.DEBUG, message, stackTrace)
    End Sub

    ' ERROR -> file + screen, with default red color (can override).
    Public Shared Sub [Error](message As String,
                              Optional stackTrace As String = Nothing,
                              Optional color As Color? = Nothing)
        Log(LogLevel.ERR, message, stackTrace, color)
    End Sub

    ' ---------- Teardown ----------
    Public Shared Sub Shutdown()
        SyncLock _sync
            _running = False
            _flushSignal.Set()
        End SyncLock
        Try
            If _bg IsNot Nothing Then _bg.Join(500)
        Catch
        End Try
        Try
            If _writer IsNot Nothing Then _writer.Flush()
            _writer.Dispose()
        Catch
        End Try
    End Sub

    ' ---------- Internals ----------
    Private Shared Sub FileFlushLoop()
        While _running
            _flushSignal.WaitOne(250)
            Dim sb As New StringBuilder(4096)
            Dim pulled As Integer = 0
            Dim line As String = ""
            While pulled < 512 AndAlso _queue.TryDequeue(line)
                If sb.Length > 0 Then sb.AppendLine()
                sb.Append(line)
                pulled += 1
            End While
            If pulled > 0 Then
                Try
                    _writer.WriteLine(sb.ToString())
                Catch
                    ' ignore IO hiccups
                End Try
            End If
        End While

        ' drain remaining
        Dim l As String = Nothing
        While _queue.TryDequeue(l)
            Try : _writer.WriteLine(l) : Catch : End Try
        End While
    End Sub

    ' Trim to last maxBytes, starting cleanly on a line boundary if possible.
    Private Shared Sub TrimLogFileToMaxBytes(path As String, maxBytes As Long)
        If maxBytes <= 0 Then Return
        If Not File.Exists(path) Then Return

        Dim fi = New FileInfo(path)
        If fi.Length <= maxBytes Then Return

        Dim tail As String
        Using fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            fs.Seek(Math.Max(0, fs.Length - maxBytes), SeekOrigin.Begin)
            Dim buf(CInt(maxBytes - 1)) As Byte
            Dim read = fs.Read(buf, 0, buf.Length)
            Dim utf8 As Encoding = New UTF8Encoding(encoderShouldEmitUTF8Identifier:=False, throwOnInvalidBytes:=False)
            tail = utf8.GetString(buf, 0, read)
        End Using

        ' Drop a partial first line if present.
        Dim i As Integer = tail.IndexOfAny(New Char() {ControlChars.Cr, ControlChars.Lf})
        If i >= 0 Then
            Dim j = i
            While j < tail.Length AndAlso (tail(j) = ControlChars.Cr OrElse tail(j) = ControlChars.Lf)
                j += 1
            End While
            If j < tail.Length Then tail = tail.Substring(j)
        End If

        Using w As New StreamWriter(New FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read))
            w.WriteLine($"----- LOG TRIMMED to last {maxBytes} bytes on {DateTime.Now:yyyy-MM-dd HH:mm:ss} -----")
            w.Write(tail)
        End Using
    End Sub

    Private Shared Function LevelName(level As LogLevel) As String
        Select Case level
            Case LogLevel.NONE : Return "NONE"
            Case LogLevel.INFO : Return "INFO"
            Case LogLevel.DEBUG : Return "DEBUG"
            Case LogLevel.ERR : Return "ERROR"
            Case Else : Return level.ToString().ToUpperInvariant()
        End Select
    End Function

    Public Shared ReadOnly Property CurrentLogFile As String
        Get
            Return _filePath
        End Get
    End Property
End Class
