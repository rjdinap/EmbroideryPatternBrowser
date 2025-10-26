Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms

Public Class StitchDisplay
    ' ---------- Dependencies ----------
    Private ReadOnly _rightPanel As Panel
    Private ReadOnly _bottomPanel As Panel
    Private ReadOnly _target As Control   ' ZoomPictureBox or PictureBox (must have .Image)
    ' ---------- UI ----------
    Private _colorListPanel As FlowLayoutPanel
    Private _progressBar As Panel
    Private _track As TrackBar
    Private _lblProgress As Label
    Private _btnPlay As Button
    Private _btnPause As Button
    Private _speed As TrackBar
    Private _tt As New ToolTip()
    ' ---------- State ----------
    Private _timer As New Timer() With {.Interval = 30}
    Private _selectedColors As New HashSet(Of Integer)
    Private _drawUpto As Integer = 0
    Private _totalStitches As Integer = 0
    Private _segments As List(Of (startIdx As Integer, endIdx As Integer, color As Color, threadIdx As Integer))
    Private _bounds As RectangleF = RectangleF.Empty
    Private _margin As Integer = 20
    Private _currentPattern As EmbPattern

    ' ---------- API ----------
    Public Sub New(rightPanel As Panel, bottomPanel As Panel, targetImageControl As Control)
        _rightPanel = rightPanel
        _bottomPanel = bottomPanel
        _target = targetImageControl
    End Sub

    Public Sub BuildUI()
        ' ---- Right: color checklist ----
        _colorListPanel = New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.TopDown,
            .WrapContents = False,
            .AutoScroll = True,
            .Padding = New Padding(4),
            .Margin = New Padding(0)
        }
        _rightPanel.Controls.Clear()
        _rightPanel.Controls.Add(_colorListPanel)
        AddHandler _colorListPanel.Resize, Sub() ResizeColorRows()
        AddHandler _timer.Tick, AddressOf Timer_Tick

        ' ---- Bottom: timeline + controls (stacked vertically) ----
        _bottomPanel.Controls.Clear()
        _bottomPanel.Padding = New Padding(6)
        _bottomPanel.MinimumSize = New Size(0, 92)

        Dim outer As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 3}
        outer.RowStyles.Add(New RowStyle(SizeType.Absolute, 14.0F))
        outer.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0F))
        outer.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))
        outer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

        _progressBar = New Panel() With {.Dock = DockStyle.Fill, .Margin = New Padding(0)}
        AddHandler _progressBar.Paint, AddressOf ProgressBar_Paint
        outer.Controls.Add(_progressBar, 0, 0)

        _track = New TrackBar() With {
            .Minimum = 0, .Maximum = 0,
            .TickStyle = TickStyle.None,
            .AutoSize = False, .Height = 28,
            .Dock = DockStyle.Fill, .Margin = New Padding(0, 2, 0, 0)
        }
        AddHandler _track.Scroll, Sub() ScrubTo(_track.Value)
        outer.Controls.Add(_track, 0, 1)

        Dim controls As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .RowCount = 1, .ColumnCount = 5, .Margin = New Padding(0, 6, 0, 0)}
        controls.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        controls.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        controls.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        controls.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        controls.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

        _btnPlay = New Button() With {.Text = "▶", .Width = 36, .Height = 28, .Margin = New Padding(0, 0, 6, 0)}
        _btnPause = New Button() With {.Text = "⏸", .Width = 36, .Height = 28, .Margin = New Padding(0, 0, 12, 0)}
        AddHandler _btnPlay.Click, Sub() StartPlayback()
        AddHandler _btnPause.Click, Sub() _timer.Stop()

        Dim lblSpeed As New Label() With {.Text = "Speed", .AutoSize = True, .Anchor = AnchorStyles.Left, .Margin = New Padding(0, 6, 6, 0)}
        _speed = New TrackBar() With {.Minimum = 1, .Maximum = 50, .Value = 50, .TickFrequency = 5, .AutoSize = False, .Height = 28, .Dock = DockStyle.Fill, .Margin = New Padding(0)}
        _lblProgress = New Label() With {.Text = "0 / 0", .AutoSize = True, .Anchor = AnchorStyles.Left, .Margin = New Padding(12, 6, 0, 0)}

        controls.Controls.Add(_btnPlay, 0, 0)
        controls.Controls.Add(_btnPause, 1, 0)
        controls.Controls.Add(lblSpeed, 2, 0)
        controls.Controls.Add(_speed, 3, 0)
        controls.Controls.Add(_lblProgress, 4, 0)

        outer.Controls.Add(controls, 0, 2)
        _bottomPanel.Controls.Add(outer)

        AddHandler _rightPanel.Resize, Sub() EnsureRightPanelFitsColorList()
        If _rightPanel.Parent IsNot Nothing Then AddHandler _rightPanel.Parent.Resize, Sub() EnsureRightPanelFitsColorList()
    End Sub

    ' Load by pattern path/name (string). Falls back gracefully on errors.
    Public Sub UpdatePattern(patternPathOrName As String)
        Dim pat As EmbPattern = Nothing
        Try
            If Not String.IsNullOrWhiteSpace(patternPathOrName) Then
                ' Uses the same loader Form1 uses
                pat = EmbReaderFactory.LoadPattern(patternPathOrName)
            End If
        Catch ex As Exception
            Try : Logger.Error("StitchDisplay.UpdatePattern: " & ex.Message, ex.ToString()) : Catch : End Try
            pat = Nothing
        End Try
        UpdatePattern(pat)
    End Sub

    ' Keep an overload if you already have an EmbPattern in hand.
    Public Sub UpdatePattern(pat As EmbPattern)
        _currentPattern = pat
        If _currentPattern Is Nothing Then
            SetTargetImage(New Bitmap(Math.Max(1, _target.Width), Math.Max(1, _target.Height)))
            Return
        End If

        BuildColorChecklist(_currentPattern)
        BuildSegments(_currentPattern)

        _totalStitches = _currentPattern.Stitches.Count
        _track.Maximum = Math.Max(0, _totalStitches)
        _track.Value = 0
        _drawUpto = 0
        _lblProgress.Text = $"0 / {_totalStitches}"
        _bounds = RectangleF.Empty

        RedrawFilteredTo(_target, _currentPattern, _selectedColors, _drawUpto)
        EnsureRightPanelFitsColorList()
    End Sub

    Public Sub StartPlayback()
        If _totalStitches <= 0 Then Return
        If _drawUpto >= _totalStitches Then ScrubTo(0)
        _timer.Start()
    End Sub

    Public Sub StopPlayback()
        _timer.Stop()
    End Sub

    ' ---------- Internals ----------
    Private Sub BuildColorChecklist(pat As EmbPattern)
        _colorListPanel.SuspendLayout()

        ' HARD CLEAR: remove + dispose existing rows to free USER handles
        If _colorListPanel.Controls.Count > 0 Then
            ' ToolTip holds refs to old controls; drop them first
            Try : _tt.RemoveAll() : Catch : End Try

            For i As Integer = _colorListPanel.Controls.Count - 1 To 0 Step -1
                Dim ctl = _colorListPanel.Controls(i)
                _colorListPanel.Controls.RemoveAt(i)
                Try : ctl.Dispose() : Catch : End Try
            Next
        End If

        _selectedColors.Clear()

        ' Build counts
        Dim counts As New Dictionary(Of Integer, Integer)()
        Dim limit As Integer = Math.Min(pat.Stitches.Count, pat.StitchThreadIndices.Count)
        For i = 0 To limit - 1
            Dim t = pat.StitchThreadIndices(i)
            counts(t) = If(counts.ContainsKey(t), counts(t) + 1, 1)
        Next

        ' If file has no thread list, show a single informational row and return
        If pat.ThreadList Is Nothing OrElse pat.ThreadList.Count = 0 Then
            Dim info As New Label() With {.AutoSize = True, .Text = "All stitches (no thread list)", .Padding = New Padding(4)}
            _colorListPanel.Controls.Add(info)
            _colorListPanel.ResumeLayout()
            ResizeColorRows()
            Return
        End If

        'extra guard on large number of theads
        Const MaxRows As Integer = 256   ' adjust as you wish
        Dim shown = Math.Min(pat.ThreadList.Count, MaxRows)
        For t = 0 To shown - 1
            ' ... add rows ...
        Next
        If pat.ThreadList.Count > MaxRows Then
            _colorListPanel.Controls.Add(New Label() With {
        .AutoSize = True,
        .Text = $"… {pat.ThreadList.Count - MaxRows} more colors",
        .Padding = New Padding(4)})
        End If



        ' Build rows
        For t = 0 To pat.ThreadList.Count - 1
            Dim th = pat.ThreadList(t)
            Dim c As Color = th.ColorValue
            Dim n As Integer = If(counts.ContainsKey(t), counts(t), 0)
            _colorListPanel.Controls.Add(MakeColorRow(t, c, n))
            _selectedColors.Add(t)
        Next

        _colorListPanel.ResumeLayout()
        ResizeColorRows()
    End Sub



    Private Function MakeColorRow(threadIdx As Integer, swatchColor As Color, stitchCount As Integer) As Control
        Dim rowH As Integer = 20
        Dim row As New Panel() With {.Height = rowH, .Margin = New Padding(0, 2, 0, 2), .Padding = New Padding(0), .AutoSize = False}

        Dim chk As New CheckBox() With {.Checked = True, .AutoSize = False, .Width = 16, .Height = 16, .Left = 0, .Top = (rowH - 16) \ 2, .Tag = threadIdx}
        AddHandler chk.CheckedChanged,
            Sub()
                If chk.Checked Then _selectedColors.Add(threadIdx) Else _selectedColors.Remove(threadIdx)
                If _currentPattern IsNot Nothing Then RedrawFilteredTo(_target, _currentPattern, _selectedColors, _drawUpto)
                _progressBar.Invalidate()
            End Sub

        Dim swatch As New Panel() With {.BackColor = swatchColor, .Width = 16, .Height = 16, .Left = chk.Right + 2, .Top = (rowH - 16) \ 2, .BorderStyle = BorderStyle.FixedSingle}
        _tt.SetToolTip(swatch, $"Thread {threadIdx}  RGB({swatchColor.R},{swatchColor.G},{swatchColor.B}) #{swatchColor.R:X2}{swatchColor.G:X2}{swatchColor.B:X2}")

        Dim lbl As New Label() With {.AutoSize = True, .Text = $"{stitchCount} sts"}
        lbl.Left = swatch.Right + 4
        lbl.Top = (row.Height - lbl.Height) \ 2

        row.Controls.Add(chk)
        row.Controls.Add(swatch)
        row.Controls.Add(lbl)
        row.Tag = lbl
        Return row
    End Function

    Private Sub BuildSegments(pat As EmbPattern)
        _segments = New List(Of (Integer, Integer, Color, Integer))()
        If pat Is Nothing Then Return

        Dim n As Integer = Math.Min(pat.Stitches.Count, pat.StitchThreadIndices.Count)
        If n <= 0 Then Return

        Dim curIdx As Integer = pat.StitchThreadIndices(0)
        Dim runStart As Integer = 0

        Dim colorForThread As Func(Of Integer, Color) =
        Function(t As Integer) As Color
            If t >= 0 AndAlso t < pat.ThreadList.Count Then
                Return pat.ThreadList(t).ColorValue
            Else
                Return Color.Black   ' fallback when thread table is empty/short
            End If
        End Function

        For i = 1 To n - 1
            Dim idx As Integer = pat.StitchThreadIndices(i)
            If idx <> curIdx Then
                _segments.Add((runStart, i - 1, colorForThread(curIdx), curIdx))
                runStart = i
                curIdx = idx
            End If
        Next

        _segments.Add((runStart, n - 1, colorForThread(curIdx), curIdx))
    End Sub

    Private Sub ProgressBar_Paint(sender As Object, e As PaintEventArgs)
        Dim g = e.Graphics
        g.Clear(Color.Gainsboro)
        If _segments Is Nothing OrElse _segments.Count = 0 OrElse _totalStitches <= 0 Then Return

        Dim r = _progressBar.ClientRectangle
        Dim W = Math.Max(1, r.Width)
        Dim H = r.Height
        Dim N = _totalStitches

        For Each seg In _segments
            If Not _selectedColors.Contains(seg.threadIdx) Then Continue For
            Dim x0 As Integer = CInt(CDbl(seg.startIdx) / N * W)
            Dim x1 As Integer = CInt(CDbl(seg.endIdx + 1) / N * W)
            Using br As New SolidBrush(seg.color)
                g.FillRectangle(br, x0, 0, Math.Max(1, x1 - x0), H)
            End Using
        Next

        Dim xCur As Integer = CInt(CDbl(_drawUpto) / Math.Max(1, N) * W)
        Using pen As New Pen(Color.Black, 2.0F)
            g.DrawLine(pen, xCur, 0, xCur, H)
        End Using
    End Sub

    Private Sub ScrubTo(v As Integer)
        _drawUpto = Math.Max(0, Math.Min(v, _totalStitches))
        _lblProgress.Text = $"{_drawUpto} / {_totalStitches}"
        If _currentPattern Is Nothing Then Return
        RedrawFilteredTo(_target, _currentPattern, _selectedColors, _drawUpto)
        _progressBar.Invalidate()
    End Sub

    Private Sub Timer_Tick(sender As Object, e As EventArgs)
        Dim remaining As Integer = Math.Max(0, _totalStitches - _drawUpto)
        Dim spd As Integer
        If _speed.Value >= _speed.Maximum Then
            spd = Math.Max(remaining \ 8, 1) ' sprint
        Else
            spd = Math.Max(1, _speed.Value)
        End If

        ScrubTo(Math.Min(_track.Maximum, _track.Value + spd))
        _track.Value = _drawUpto
        If _drawUpto >= _track.Maximum Then _timer.Stop()
    End Sub

    Private Sub RedrawFilteredTo(target As Control, pat As EmbPattern, visibleThreads As HashSet(Of Integer), upto As Integer)
        If pat Is Nothing OrElse target Is Nothing Then Return
        Dim W = Math.Max(target.Width, 1), H = Math.Max(target.Height, 1)

        If _bounds.Width = 0 OrElse _bounds.Height = 0 Then
            _bounds = pat.ComputeBounds()
            If _bounds.Width <= 0 OrElse _bounds.Height <= 0 Then
                SetTargetImage(New Bitmap(W, H)) : Return
            End If
        End If

        Dim bmp As New Bitmap(W, H)
        Using g = Graphics.FromImage(bmp)
            g.Clear(Color.White)
            g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

            Dim inner = New RectangleF(_margin, _margin, W - 2 * _margin, H - 2 * _margin)
            Dim s As Single = Math.Min(inner.Width / _bounds.Width, inner.Height / _bounds.Height)
            Dim ox As Single = inner.X - _bounds.Left * s
            Dim oy As Single = inner.Y - _bounds.Top * s

            Dim n As Integer = Math.Min(Math.Max(0, upto), Math.Min(pat.Stitches.Count - 1, pat.StitchThreadIndices.Count - 1))
            Dim threadsAvailable As Integer = pat.ThreadList.Count
            Dim useFilter As Boolean = (threadsAvailable > 0)

            For i = 1 To n
                Dim tPrev = pat.StitchThreadIndices(i - 1)
                Dim tThis = pat.StitchThreadIndices(i)

                ' thread-change boundaries
                If tThis <> tPrev Then Continue For

                ' skip breaks
                Dim breakBefore As Boolean = (i < pat.BreakBefore.Count) AndAlso pat.BreakBefore(i)
                If breakBefore Then Continue For

                ' only apply the color filter if we actually have a thread list
                If useFilter Then
                    If tThis < 0 OrElse tThis >= threadsAvailable Then Continue For
                    If Not visibleThreads.Contains(tThis) Then Continue For
                End If

                Dim p0 = pat.Stitches(i - 1)
                Dim p1 = pat.Stitches(i)
                Dim col As Color = If(useFilter, pat.ThreadList(tThis).ColorValue, Color.Black)

                Using pen As New Pen(col, 1.6F)
                    g.DrawLine(pen, ox + p0.X * s, oy + p0.Y * s, ox + p1.X * s, oy + p1.Y * s)
                End Using
            Next
        End Using

        SetTargetImage(bmp)
    End Sub

    ' Assigns image to PictureBox/ZoomPictureBox (or any control with an Image property); calls FitToWindow() if available.
    Private Sub SetTargetImage(img As Image)
        If _target Is Nothing Then
            Try : img?.Dispose() : Catch : End Try
            Return
        End If

        Dim oldImg As Image = Nothing
        Try
            Dim p = _target.GetType().GetProperty("Image")
            If p IsNot Nothing AndAlso p.CanWrite Then
                Try : oldImg = TryCast(p.GetValue(_target, Nothing), Image) : Catch : End Try
                p.SetValue(_target, img, Nothing)
            End If
        Catch
        End Try

        ' If ZoomPictureBox, fit to window so first render isn't zoomed strangely
        Try
            Dim m = _target.GetType().GetMethod("FitToWindow", Type.EmptyTypes)
            If m IsNot Nothing Then m.Invoke(_target, Nothing)
        Catch
        End Try

        If oldImg IsNot Nothing AndAlso Not Object.ReferenceEquals(oldImg, img) Then
            Try : oldImg.Dispose() : Catch : End Try
        End If
    End Sub

    ' ---------- Sizing helpers ----------
    Private Sub ResizeColorRows()
        If _colorListPanel Is Nothing Then Return
        Dim innerW = _colorListPanel.ClientSize.Width - _colorListPanel.Padding.Horizontal
        For Each row As Control In _colorListPanel.Controls
            row.Width = Math.Max(120, innerW)
            Dim lbl = TryCast(row.Tag, Label)
            If lbl IsNot Nothing Then lbl.Top = (row.Height - lbl.Height) \ 2
        Next
    End Sub

    Private Function GetColorListPreferredWidth() As Integer
        If _colorListPanel Is Nothing OrElse _colorListPanel.Controls.Count = 0 Then Return 140
        Dim maxRight As Integer = 0
        For Each row As Control In _colorListPanel.Controls
            Dim swatch = row.Controls.OfType(Of Panel)().FirstOrDefault(Function(p) p.BorderStyle = BorderStyle.FixedSingle)
            Dim lbl = TryCast(row.Tag, Label)
            If swatch Is Nothing OrElse lbl Is Nothing Then Continue For
            Dim rightEdge As Integer = (swatch.Right + 4) + lbl.PreferredSize.Width
            If rightEdge > maxRight Then maxRight = rightEdge
        Next
        Dim content = maxRight + 6
        Return Math.Max(120, content + _colorListPanel.Padding.Horizontal)
    End Function

    Private Sub EnsureRightPanelFitsColorList()
        If _rightPanel Is Nothing Then Return
        Dim desiredRightWidth As Integer = GetColorListPreferredWidth()
        desiredRightWidth += _rightPanel.Padding.Horizontal + 4

        Dim parent = _rightPanel.Parent
        If TypeOf parent Is SplitContainer Then
            Dim sc = DirectCast(parent, SplitContainer)
            Dim newLeft = Math.Max(100, sc.Width - desiredRightWidth - sc.SplitterWidth)
            sc.SplitterDistance = newLeft
            sc.Panel2MinSize = desiredRightWidth
        ElseIf TypeOf parent Is TableLayoutPanel Then
            Dim tlp = DirectCast(parent, TableLayoutPanel)
            Dim col = tlp.GetColumn(_rightPanel)
            While tlp.ColumnStyles.Count <= col
                tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            End While
            tlp.ColumnStyles(col).SizeType = SizeType.Absolute
            tlp.ColumnStyles(col).Width = desiredRightWidth
        Else
            _rightPanel.Width = desiredRightWidth
            _rightPanel.Dock = DockStyle.Right
        End If

        ResizeColorRows()
    End Sub
End Class
