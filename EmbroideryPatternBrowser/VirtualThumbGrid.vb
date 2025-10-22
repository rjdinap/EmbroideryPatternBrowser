Imports System.Collections.Concurrent
Imports System.Drawing.Drawing2D
Imports System.Threading
Imports System.Linq

' Virtualized thumbnail grid for huge DataTables (100k–1M rows).
' Required columns: fullpath, filename, ext, size, metadata.
' Usage:
'   grid.Bind(dt, AddressOf CreateImageFromPatternWrapper, PictureBox_FullImage, AddressOf SaveMetadataToDb)
' Where:
'   CreateImageFromPatternWrapper(path As String, w As Integer, h As Integer) As Image
'   SaveMetadataToDb(row As DataRow, newMetadata As String) As Boolean
Public Class VirtualThumbGrid
    Inherits ScrollableControl

    ' === Public API / data ===
    Private _table As DataTable
    Private _rowCount As Integer

    ' Optional: full image hookup (left-click)
    Private _fullImageFactory As Func(Of String, Integer, Integer, Image) ' CreateImageFromPattern wrapper
    Private _fullImageTarget As PictureBox

    ' Optional: save metadata delegate (DB persistence)
    Private _saveMetadata As Func(Of DataRow, String, Boolean)

    ' Selection state
    Private _selectedIndex As Integer = -1


    ' --- Drag / drop support fields ---
    Private _dragStart As Point
    Private Const DragThreshold As Integer = 8
    Private _thumbSize As Size = New Size(ThumbW, ThumbH)  ' your thumb size
    Private _cellPadding As Padding = New Padding(Pad) ' space around thumb inside each cell
    Private _cellSize As Size                        ' computed = thumb + padding
    Private _selectedIndices As New HashSet(Of Integer)() ' ==== Selection (multi-select support) ====
    Private _anchorIndex As Integer = -1 '' Selection anchor for Shift-range

    ' --- Virtual/zip data exchange format (between our controls) ---
    Private Const DATAFMT_VIRTUAL_LIST As String = "EPB_VIRTUAL_FILE_LIST"

    ' Example in VirtualThumbGrid (constructor or Bind):
    Dim tw As Integer = Math.Max(32, My.Settings.ThumbWidth)
    Dim th As Integer = Math.Max(32, My.Settings.ThumbHeight)

    ' === Layout / rendering sizes (read from settings) ===
    Private ReadOnly Property ThumbW As Integer
        Get
            Dim w As Integer = My.Settings.ThumbWidth
            If w <= 0 Then w = 200      ' fallback if someone saved 0/negative
            If w > 1000 Then w = 1000
            Return w
        End Get
    End Property

    Private ReadOnly Property ThumbH As Integer
        Get
            Dim h As Integer = My.Settings.ThumbHeight
            If h <= 0 Then h = 200      ' fallback
            If h > 1000 Then h = 1000
            Return h
        End Get
    End Property

    Private Const Pad As Integer = 12

    Private ReadOnly Property TileW As Integer
        Get
            Return ThumbW + Pad
        End Get
    End Property

    Private ReadOnly Property TileH As Integer
        Get
            Return ThumbH + 34   ' room for filename + margins
        End Get
    End Property




    Public ReadOnly Property SelectedIndex As Integer
        Get
            Return _selectedIndex
        End Get
    End Property

    ' Context menu wrapper
    Private _popup As PopupMenuItems


    ' Bind + optional delegates
    Public Overloads Sub Bind(table As DataTable,
                              Optional fullImageFactory As Func(Of String, Integer, Integer, Image) = Nothing,
                              Optional fullImageTarget As PictureBox = Nothing,
                              Optional saveMetadata As Func(Of DataRow, String, Boolean) = Nothing)
        If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))
        _table = table
        _rowCount = _table.Rows.Count

        _fullImageFactory = fullImageFactory
        _fullImageTarget = fullImageTarget
        _saveMetadata = saveMetadata

        ' Build popup with callbacks
        _popup = New PopupMenuItems(
            openDirAction:=Sub() OpenDirectoryForIndex(_selectedIndex),
            editMetadataAction:=Sub() EditMetadataForIndex(_selectedIndex),
             copyPathAction:=Sub() CopySelectedToClipboard()
        )
        'copyPathAction:=Sub() CopyFullPathForIndex(_selectedIndex) 'old call to handle single item

        ResetLayoutMetrics()
        InitWorkers()
        'old version
        'If _rowCount > 0 Then _selectedIndex = 0 Else _selectedIndex = -1
        'drag drop version
        If _rowCount > 0 Then
            _selectedIndex = 0
            _selectedIndices.Clear()
            _selectedIndices.Add(0)
        Else
            _selectedIndex = -1
            _selectedIndices.Clear()
        End If
        Invalidate()

        ' Re-render preview on PictureBox resize so it always fits
        If _fullImageTarget IsNot Nothing Then
            AddHandler _fullImageTarget.Resize, Sub(sender, args)
                                                    If _selectedIndex >= 0 Then
                                                        TryShowFullImageForIndex(_selectedIndex)
                                                    End If
                                                End Sub
        End If
    End Sub


    ' Prefetch rows around viewport
    Private Const PreloadRowsAbove As Integer = 2
    Private Const PreloadRowsBelow As Integer = 4

    ' Background loader degree of parallelism
    Private Const MaxConcurrentLoads As Integer = 6

    ' LRU cache capacity (images)
    Private Const CacheCapacity As Integer = 2000

    ' === State ===
    Private _cols As Integer = 1
    Private _rows As Integer = 0

    ' ToolTip + stable hover (debounced) — owner-drawn to bump font size by +2pt
    Private ReadOnly _tip As New ToolTip() With {
        .AutomaticDelay = 250,
        .AutoPopDelay = 8000,
        .InitialDelay = 250,
        .ReshowDelay = 100,
        .ShowAlways = True,
        .OwnerDraw = True
    }
    Private ReadOnly _hoverTimer As New System.Windows.Forms.Timer() With {.Interval = 350}
    Private _hoverCandidateIndex As Integer = -1
    Private _hoverShownIndex As Integer = -1
    Private _lastMousePt As Point = Point.Empty

    ' Background loading infra
    Private _loadSemaphore As SemaphoreSlim
    Private _cts As CancellationTokenSource
    Private _queue As BlockingCollection(Of Integer)

    ' LRU cache
    Private ReadOnly _cache As New LruImageCache(CacheCapacity)

    ' Placeholders
    Private Shared ReadOnly PlaceholderBack As Brush = SystemBrushes.ControlLight
    Private Shared ReadOnly PlaceholderBorder As Pen = Pens.Gray

    ' Prioritization/queue control
    Private ReadOnly _pending As New System.Collections.Concurrent.ConcurrentDictionary(Of Integer, Byte)()
    Private _activeFirstIndex As Integer = 0
    Private _activeLastIndex As Integer = -1

    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.UserPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw Or
                 ControlStyles.Selectable, True)
        DoubleBuffered = True
        AutoScroll = True
        BackColor = Color.White
        TabStop = True

        _tip.UseAnimation = False
        _tip.UseFading = False
        ' existing _tip init…
        _tip.OwnerDraw = True
        AddHandler _tip.Popup, AddressOf OnToolTipPopup
        AddHandler _tip.Draw, AddressOf OnToolTipDraw
        AddHandler _hoverTimer.Tick, AddressOf OnHoverTimerTick

        ' Owner-draw tooltip to increase font size by +2pt
        AddHandler _tip.Draw,
            Sub(s, e)
                e.Graphics.Clear(Color.FromArgb(255, 255, 255))
                Using f As New Font(Me.Font.FontFamily, Me.Font.Size + 2.0F, FontStyle.Regular, GraphicsUnit.Point)
                    Using textBrush As New SolidBrush(Color.Black)
                        e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
                        e.Graphics.DrawString(e.ToolTipText, f, textBrush, e.Bounds)
                    End Using
                End Using
                e.DrawBorder()
            End Sub
    End Sub


    ' === Owner-drawn ToolTip sizing/drawing ===
    Private ReadOnly _tipFont As New Font("Segoe UI", Me.Font.SizeInPoints + 2.0F, FontStyle.Regular)
    Private Const TipMaxWidth As Integer = 520   ' widen so metadata lines wrap nicely
    Private Const TipPad As Integer = 10

    Private Sub OnToolTipPopup(sender As Object, e As PopupEventArgs)
        ' Measure text with word-wrap to set an explicit size
        Dim text As String = _tip.GetToolTip(Me)
        If String.IsNullOrEmpty(text) Then
            e.ToolTipSize = New Size(0, 0)
            Return
        End If
        Dim fmt As TextFormatFlags = TextFormatFlags.WordBreak Or TextFormatFlags.NoPrefix
        ' Constrain width; TextRenderer will compute needed height
        Dim proposed As New Size(TipMaxWidth, Integer.MaxValue)
        Dim sz As Size = TextRenderer.MeasureText(text, _tipFont, proposed, fmt)
        e.ToolTipSize = New Size(sz.Width + TipPad * 2, sz.Height + TipPad * 2)
    End Sub

    Private Sub OnToolTipDraw(sender As Object, e As DrawToolTipEventArgs)
        ' Background & border
        e.Graphics.Clear(Color.FromArgb(255, 255, 255))
        Using border As New Pen(Color.FromArgb(160, 160, 160))
            e.Graphics.DrawRectangle(border, 0, 0, e.Bounds.Width - 1, e.Bounds.Height - 1)
        End Using

        ' Text
        Dim fmt As TextFormatFlags = TextFormatFlags.WordBreak Or TextFormatFlags.NoPrefix
        Dim textRect As New Rectangle(e.Bounds.X + TipPad, e.Bounds.Y + TipPad,
                                  e.Bounds.Width - TipPad * 2, e.Bounds.Height - TipPad * 2)
        TextRenderer.DrawText(e.Graphics, e.ToolTipText, _tipFont, textRect, Color.Black, fmt)
    End Sub




    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            Try : _cts?.Cancel() : Catch : End Try
            _queue?.Dispose()
            _loadSemaphore?.Dispose()
            _cts?.Dispose()
            _cache.Dispose()
            RemoveHandler _hoverTimer.Tick, AddressOf OnHoverTimerTick
            _tip.RemoveAll()
            _tip.Dispose()
            _hoverTimer.Stop()
            _hoverTimer.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' === Workers ===
    Private Sub InitWorkers()
        _cts?.Cancel()
        _queue?.Dispose()
        _loadSemaphore?.Dispose()

        _cts = New CancellationTokenSource()
        _queue = New BlockingCollection(Of Integer)(boundedCapacity:=4096)
        _loadSemaphore = New SemaphoreSlim(MaxConcurrentLoads)

        For i = 1 To MaxConcurrentLoads
            Task.Run(Function() LoaderLoop(_cts.Token))
        Next

        QueueVisibleAndBuffer()
    End Sub

    Private Async Function LoaderLoop(ct As CancellationToken) As Task
        Try
            For Each idx In _queue.GetConsumingEnumerable(ct)
                If ct.IsCancellationRequested Then Exit For
                If idx < 0 OrElse idx >= _rowCount Then
                    Dim d0 As Byte : _pending.TryRemove(idx, d0)
                    Continue For
                End If

                If Not IsInActiveWindow(idx) AndAlso Not _cache.Contains(idx) Then
                    Dim d1 As Byte : _pending.TryRemove(idx, d1)
                    Continue For
                End If

                Await _loadSemaphore.WaitAsync(ct)
                Try
                    If _cache.Contains(idx) Then
                        Dim d2 As Byte : _pending.TryRemove(idx, d2)
                        Continue For
                    End If

                    Dim row = _table.Rows(idx)
                    'this is where we get the filename that we're going to load
                    'for a normal filename, we can just pull it from fullpath.
                    'we'll need special processing if it's a .zip file
                    Dim fullpath As String = SafeStr(row("fullpath"))

                    If String.IsNullOrEmpty(fullpath) Then
                        Dim d3 As Byte : _pending.TryRemove(idx, d3)
                        Continue For
                    End If

                    Dim img As Image = Nothing
                    ' Render thumbnail off-UI thread
                    Await Task.Run(Sub()
                                       img = EmbThumbnail.GenerateFromFile(fullpath,
                                                                           maxWidth:=ThumbW,
                                                                           maxHeight:=ThumbH,
                                                                           padding:=6,
                                                                           bg:=Color.White,
                                                                           drawBorder:=True,
                                                                           showName:=False)
                                   End Sub, ct)

                    If img IsNot Nothing Then
                        _cache.Add(idx, img, fullpath)
                        BeginInvoke(DirectCast(Sub() InvalidateTile(idx), Action))
                    End If
                Catch ex As OperationCanceledException
                    Exit For
                Catch
                    ' ignore bad rows
                Finally
                    _loadSemaphore.Release()
                    Dim d4 As Byte : _pending.TryRemove(idx, d4)
                End Try
            Next
        Catch ex As OperationCanceledException
        Catch
        End Try
    End Function



    ' === Layout ===
    Private Sub ResetLayoutMetrics()
        Dim colsCalc As Integer = Math.Max(1, (ClientSize.Width - Pad) \ TileW)
        _cols = colsCalc
        _rows = If(_rowCount <= 0, 0, CInt(Math.Ceiling(_rowCount / CDbl(_cols))))
        AutoScrollMinSize = New Size(0, Pad + _rows * TileH + Pad)
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        'drag drop stuff
        _cellSize = New Size(_thumbSize.Width + _cellPadding.Horizontal,
                         _thumbSize.Height + _cellPadding.Vertical)
        _cols = Math.Max(1, Math.Floor(Me.ClientSize.Width / Math.Max(1, _cellSize.Width)))
        _hoverTimer.Stop()
        HideTooltip()
        ResetLayoutMetrics()
        QueueVisibleAndBuffer()
        Invalidate()
    End Sub

    Protected Overrides Sub OnScroll(se As ScrollEventArgs)
        MyBase.OnScroll(se)
        _hoverTimer.Stop()
        HideTooltip()
        QueueVisibleAndBuffer()
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseWheel(e As MouseEventArgs)
        MyBase.OnMouseWheel(e)
        QueueVisibleAndBuffer()
    End Sub

    ' === Painting ===
    Protected Overrides Sub OnPaint(pe As PaintEventArgs)
        MyBase.OnPaint(pe)

        If _table Is Nothing OrElse _rowCount = 0 Then
            Using f = New Font(Font, FontStyle.Italic)
                TextRenderer.DrawText(pe.Graphics, "No items", f, ClientRectangle, Color.Gray,
                                      TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            End Using
            Return
        End If

        pe.Graphics.TranslateTransform(AutoScrollPosition.X, AutoScrollPosition.Y)
        Dim clip = pe.ClipRectangle
        clip.Offset(-AutoScrollPosition.X, -AutoScrollPosition.Y)

        Dim firstVisibleRow As Integer = Math.Max(0, (clip.Top - Pad) \ TileH)
        Dim lastVisibleRow As Integer = Math.Min(Math.Max(0, _rows - 1), (clip.Bottom - Pad) \ TileH)

        For r = firstVisibleRow To lastVisibleRow
            Dim y As Integer = Pad + r * TileH
            For c = 0 To _cols - 1
                Dim idx As Integer = r * _cols + c
                If idx >= _rowCount Then Exit For
                Dim x As Integer = Pad + c * TileW

                Dim rectImg As New Rectangle(x, y, ThumbW, ThumbH)
                Dim img As Image = _cache.GetImage(idx)
                If img IsNot Nothing Then
                    pe.Graphics.DrawImage(img, rectImg)
                Else
                    pe.Graphics.FillRectangle(PlaceholderBack, rectImg)
                    pe.Graphics.DrawRectangle(PlaceholderBorder, rectImg)
                    Using sf0 As New StringFormat()
                        sf0.Alignment = StringAlignment.Center
                        sf0.LineAlignment = StringAlignment.Center
                        pe.Graphics.DrawString("…", Me.Font, Brushes.DimGray, rectImg, sf0)
                    End Using
                End If

                Dim row = _table.Rows(idx)
                Dim name As String = SafeStr(row("filename"))
                Dim textRect As New RectangleF(x, y + ThumbH + 3, ThumbW, TileH - ThumbH - 6)
                Using sf As New StringFormat()
                    sf.Trimming = StringTrimming.EllipsisCharacter
                    sf.Alignment = StringAlignment.Center
                    sf.LineAlignment = StringAlignment.Near
                    pe.Graphics.DrawString(name, Me.Font, Brushes.Black, textRect, sf)
                End Using


                ' selection highlight (clear and obvious for all selected)
                Dim isSelected As Boolean = (_selectedIndices IsNot Nothing AndAlso _selectedIndices.Contains(idx))
                If isSelected Then
                    ' Entire tile rect (image + caption area) with a bit of padding
                    Dim selRect As New Rectangle(x - 3, y - 3, ThumbW + 6, TileH + 6)

                    ' Optional: soft translucent fill so it’s unmistakable
                    Using fillBrush As New SolidBrush(Color.FromArgb(48, &H1E, &H90, &HFF)) ' ~20% DodgerBlue
                        pe.Graphics.FillRectangle(fillBrush, selRect)
                    End Using

                    ' Bold border for all selected; slightly bolder for the caret/anchor
                    Dim isAnchor As Boolean = (idx = _selectedIndex)
                    Dim borderWidth As Single = If(isAnchor AndAlso Me.Focused, 3.0F, 2.0F)
                    Dim borderColor As Color = If(isAnchor AndAlso Me.Focused, Color.DodgerBlue, Color.SteelBlue)
                    Using pen As New Pen(borderColor, borderWidth)
                        pe.Graphics.DrawRectangle(pen, selRect)
                    End Using
                End If
            Next
        Next
    End Sub


    Private Sub InvalidateTile(index As Integer)
        If index < 0 OrElse index >= _rowCount Then Return

        Dim r As Integer = If(_cols = 0, 0, index \ _cols)
        Dim c As Integer = If(_cols = 0, 0, index Mod _cols)

        ' Virtual (content) coordinates where you draw the tile
        Dim vx As Integer = Pad + c * TileW
        Dim vy As Integer = Pad + r * TileH

        ' Convert to client coordinates for invalidation: ADD AutoScrollPosition
        Dim cx As Integer = vx + AutoScrollPosition.X
        Dim cy As Integer = vy + AutoScrollPosition.Y

        Dim rect As New Rectangle(cx - 4, cy - 4, ThumbW + 8, TileH + 8)
        Invalidate(rect)
    End Sub




    ' Mouse down: multi-select aware (Ctrl/Shift) + drag prep
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        Focus()

        Dim idx = IndexFromPoint(e.Location)
        If idx = -1 Then Return

        Dim ctrl As Boolean = (Control.ModifierKeys And Keys.Control) = Keys.Control
        Dim shift As Boolean = (Control.ModifierKeys And Keys.Shift) = Keys.Shift

        If e.Button = MouseButtons.Left Then
            Dim prev = _selectedIndices.ToList() ' snapshot for painting diff

            If shift AndAlso _selectedIndex >= 0 Then
                ' SHIFT range select from the anchor (or current if no anchor yet)
                If _anchorIndex < 0 Then _anchorIndex = _selectedIndex
                Dim a = Math.Min(_anchorIndex, idx)
                Dim b = Math.Max(_anchorIndex, idx)
                _selectedIndices.Clear()
                For i = a To b
                    _selectedIndices.Add(i)
                Next
                SetSelectedIndex(idx, ensureVisible:=False)

            ElseIf ctrl Then
                ' CTRL toggle
                If _selectedIndices.Contains(idx) Then
                    _selectedIndices.Remove(idx)
                Else
                    _selectedIndices.Add(idx)
                End If
                _anchorIndex = idx
                SetSelectedIndex(idx, ensureVisible:=False)

            Else
                ' Plain click => single selection
                _selectedIndices.Clear()
                _selectedIndices.Add(idx)
                _anchorIndex = idx
                SetSelectedIndex(idx, ensureVisible:=False)
            End If

            ' Repaint only changed tiles
            InvalidateSelectionDiff(prev)
            Invalidate()

            TryShowFullImageForIndex(idx)

            ' Drag prep
            _dragStart = e.Location

        ElseIf e.Button = MouseButtons.Right Then
            ' Right-click: keep multi if already contains; otherwise single-select
            Dim prev = _selectedIndices.ToList()
            If Not _selectedIndices.Contains(idx) Then
                _selectedIndices.Clear()
                _selectedIndices.Add(idx)
            End If
            _anchorIndex = idx
            SetSelectedIndex(idx, ensureVisible:=False)
            InvalidateSelectionDiff(prev)
            Invalidate()
            _popup?.Show(Me, e.Location)
        End If
    End Sub



    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        If e.Button = MouseButtons.Right Then
            Dim idx = IndexFromPoint(e.Location)
            If idx <> -1 Then
                SetSelectedIndex(idx, ensureVisible:=False)
                _popup?.Show(Me, e.Location)
            End If
        End If
    End Sub

    ' Fit src into (w x h) preserving aspect; letterbox with white if needed.
    Private Shared Function FitBitmap(ByVal src As Image, ByVal w As Integer, ByVal h As Integer) As Image
        If src Is Nothing OrElse w <= 0 OrElse h <= 0 Then Return src
        Dim scale As Double = Math.Min(w / CDbl(src.Width), h / CDbl(src.Height))
        If scale > 1.0 Then scale = 1.0 ' avoid upscaling; remove if you want larger previews

        Dim outW As Integer = Math.Max(1, CInt(Math.Floor(src.Width * scale)))
        Dim outH As Integer = Math.Max(1, CInt(Math.Floor(src.Height * scale)))

        Dim bmp As New Bitmap(w, h, Drawing.Imaging.PixelFormat.Format24bppRgb)
        Using g = Graphics.FromImage(bmp)
            g.Clear(Color.White)
            g.InterpolationMode = InterpolationMode.HighQualityBicubic
            g.PixelOffsetMode = PixelOffsetMode.HighQuality
            Dim x As Integer = (w - outW) \ 2
            Dim y As Integer = (h - outH) \ 2
            g.DrawImage(src, New Rectangle(x, y, outW, outH))
        End Using
        Return bmp
    End Function

    Private Sub TryShowFullImageForIndex(idx As Integer)
        If _fullImageFactory Is Nothing OrElse _fullImageTarget Is Nothing Then Return
        If idx < 0 OrElse idx >= _rowCount Then Return
        Dim row = _table.Rows(idx)
        Dim path As String = SafeStr(row("fullpath"))
        If String.IsNullOrEmpty(path) Then Return

        ' Snapshot target size now (may be 0 during layout; guard)
        Dim w As Integer = Math.Max(1, _fullImageTarget.ClientSize.Width)
        Dim h As Integer = Math.Max(1, _fullImageTarget.ClientSize.Height)

        Task.Run(Function()
                     Dim img As Image = Nothing
                     Try
                         img = _fullImageFactory(path, w, h) ' renderer already tries to fit
                     Catch
                     End Try
                     Return img
                 End Function).
             ContinueWith(Sub(t)
                              Dim rawImg = t.Result
                              If rawImg Is Nothing Then Return
                              If Not _fullImageTarget.IsHandleCreated Then
                                  Try : rawImg.Dispose() : Catch : End Try
                                  Return
                              End If

                              ' Compute current size on UI thread and hard-fit
                              Dim targetW As Integer = Math.Max(1, _fullImageTarget.ClientSize.Width)
                              Dim targetH As Integer = Math.Max(1, _fullImageTarget.ClientSize.Height)
                              Dim fitted As Image = Nothing
                              Try
                                  fitted = FitBitmap(rawImg, targetW, targetH)
                              Catch
                                  fitted = rawImg
                              End Try

                              _fullImageTarget.BeginInvoke(DirectCast(Sub()
                                                                          Dim old = _fullImageTarget.Image
                                                                          _fullImageTarget.Image = fitted
                                                                          If (Not Object.ReferenceEquals(old, fitted)) AndAlso old IsNot Nothing Then
                                                                              Try : old.Dispose() : Catch : End Try
                                                                          End If
                                                                          If Not Object.ReferenceEquals(fitted, rawImg) AndAlso rawImg IsNot Nothing Then
                                                                              Try : rawImg.Dispose() : Catch : End Try
                                                                          End If
                                                                      End Sub, Action))

                          End Sub, TaskScheduler.FromCurrentSynchronizationContext())
    End Sub

    ' === Popup actions ===
    Private Sub OpenDirectoryForIndex(idx As Integer)
        If idx < 0 OrElse idx >= _rowCount Then Return
        Dim row = _table.Rows(idx)
        Dim path As String = SafeStr(row("fullpath"))
        If String.IsNullOrEmpty(path) Then Return

        Try
            ShellSelectHelper.OpenAndSelect(path)
        Catch
        End Try
    End Sub

    Private Sub EditMetadataForIndex(idx As Integer)
        If idx < 0 OrElse idx >= _rowCount Then Return
        Dim row = _table.Rows(idx)
        Dim current As String = SafeStr(row("metadata"))

        Using dlg As New MetadataEditorForm()
            dlg.MetadataText = current
            If dlg.ShowDialog(FindForm()) = DialogResult.OK Then
                Dim newText As String = dlg.MetadataText

                ' 1) Save to DB if delegate provided
                Dim dbOk As Boolean = True
                If _saveMetadata IsNot Nothing Then
                    Try
                        dbOk = _saveMetadata(row, newText)
                    Catch
                        dbOk = False
                    End Try
                End If

                ' 2) Update the DataTable row (only if DB ok, or if no delegate provided)
                If dbOk OrElse _saveMetadata Is Nothing Then
                    row("metadata") = newText
                    InvalidateTile(idx)
                Else
                    MessageBox.Show(Me, "Failed to save metadata to the database.", "Save Failed",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        End Using
    End Sub

    ' === Keyboard navigation ===
    Protected Overrides Function IsInputKey(keyData As Keys) As Boolean
        Select Case (keyData And Keys.KeyCode)
            Case Keys.Left, Keys.Right, Keys.Up, Keys.Down, Keys.PageUp, Keys.PageDown, Keys.Home, Keys.End
                Return True
        End Select
        Return MyBase.IsInputKey(keyData)
    End Function



    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If _rowCount <= 0 Then Return

        Dim newIndex As Integer = If(_selectedIndex < 0, 0, _selectedIndex)
        Dim changed As Boolean = False

        Dim visibleRows As Integer = Math.Max(1, (ClientSize.Height - Pad) \ TileH)
        Dim pageStep As Integer = visibleRows * Math.Max(1, _cols)

        Select Case e.KeyCode
            Case Keys.Left
                If _cols > 0 Then newIndex = Math.Max(0, newIndex - 1) : changed = True
            Case Keys.Right
                If _cols > 0 Then newIndex = Math.Min(_rowCount - 1, newIndex + 1) : changed = True
            Case Keys.Up
                If _cols > 0 Then newIndex = Math.Max(0, newIndex - _cols) : changed = True
            Case Keys.Down
                If _cols > 0 Then newIndex = Math.Min(_rowCount - 1, newIndex + _cols) : changed = True
            Case Keys.PageUp
                newIndex = Math.Max(0, newIndex - pageStep) : changed = True
            Case Keys.PageDown
                newIndex = Math.Min(_rowCount - 1, newIndex + pageStep) : changed = True
            Case Keys.Home
                newIndex = 0 : changed = True
            Case Keys.End
                newIndex = _rowCount - 1 : changed = True
            Case Else
                ' not a nav key
        End Select

        If Not changed Then Return

        If e.Shift Then
            Dim prev = _selectedIndices.ToList()
            If _anchorIndex < 0 Then _anchorIndex = _selectedIndex
            If _anchorIndex < 0 Then _anchorIndex = 0
            _selectedIndices.Clear()
            Dim a = Math.Min(_anchorIndex, newIndex)
            Dim b = Math.Max(_anchorIndex, newIndex)
            For i = a To b
                _selectedIndices.Add(i)
            Next
            SetSelectedIndex(newIndex, ensureVisible:=True)
            InvalidateSelectionDiff(prev)

        ElseIf e.Control Then
            ' Move caret only
            SetSelectedIndex(newIndex, ensureVisible:=True)

        Else
            ' Single selection
            SetSingleSelection(newIndex, ensureVisible:=True)
        End If

        Invalidate()
        e.Handled = True
    End Sub


    Private Sub SetSelectedIndex(index As Integer, Optional ensureVisible As Boolean = False)
        If index < 0 OrElse index >= _rowCount Then Return
        If index = _selectedIndex Then
            If ensureVisible Then ScrollIntoView(index)
            Return
        End If
        Dim old = _selectedIndex
        _selectedIndex = index
        If old >= 0 Then InvalidateTile(old)
        InvalidateTile(_selectedIndex)
        If ensureVisible Then ScrollIntoView(_selectedIndex)
    End Sub

    Private Sub ScrollIntoView(index As Integer)
        If index < 0 OrElse index >= _rowCount Then Return
        Dim r As Integer = If(_cols = 0, 0, index \ _cols)
        Dim yTop As Integer = Pad + r * TileH
        Dim yBottom As Integer = yTop + TileH

        Dim viewTop As Integer = -AutoScrollPosition.Y
        Dim viewBottom As Integer = viewTop + ClientSize.Height

        Dim targetTop As Integer = viewTop
        If yTop < viewTop Then
            targetTop = yTop
        ElseIf yBottom > viewBottom Then
            targetTop = yBottom - ClientSize.Height
        Else
            Return
        End If

        targetTop = Math.Max(0, Math.Min(targetTop, (Pad + _rows * TileH + Pad) - ClientSize.Height))
        AutoScrollPosition = New Point(0, targetTop)
        QueueVisibleAndBuffer()
        Invalidate()
    End Sub


    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)

        ' --- HOVER / TOOLTIP TRACKING ---
        _lastMousePt = e.Location
        Dim idx = IndexFromPoint(e.Location)

        If idx <> _hoverCandidateIndex Then
            _hoverCandidateIndex = idx
            _hoverTimer.Stop()
            HideTooltip()
            If idx >= 0 AndAlso idx < _rowCount Then
                _hoverTimer.Start()
            End If
        End If

        ' --- DRAG START CHECK ---
        If e.Button = MouseButtons.Left Then
            Dim dx = Math.Abs(e.X - _dragStart.X)
            Dim dy = Math.Abs(e.Y - _dragStart.Y)
            If Math.Max(dx, dy) >= DragThreshold Then
                Dim paths As List(Of String) = GetDragFilePathsAt(e.Location)
                If paths IsNot Nothing AndAlso paths.Count > 0 Then
                    Dim all = paths.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                    Dim real = all.Where(Function(p) IO.File.Exists(p)).ToList()
                    Dim virt = all.Where(Function(p) Not IO.File.Exists(p) AndAlso IsCompositeZipPath(p)).ToList()

                    Dim data As New DataObject()

                    ' Real files => CF_HDROP
                    If real.Count > 0 Then
                        Dim sc As New Specialized.StringCollection()
                        sc.AddRange(real.ToArray())
                        data.SetFileDropList(sc)
                    End If

                    ' All (real + virtual) => custom format for our USB panel
                    If all.Count > 0 Then
                        data.SetData(DATAFMT_VIRTUAL_LIST, False, all)
                        data.SetText(String.Join(Environment.NewLine, all))
                    End If

                    DoDragDrop(data, DragDropEffects.Copy)
                End If
            End If
        End If
    End Sub



    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        _hoverTimer.Stop()
        HideTooltip()
        _hoverCandidateIndex = -1
    End Sub

    Private Sub OnHoverTimerTick(sender As Object, e As EventArgs)
        _hoverTimer.Stop()
        Dim idx = _hoverCandidateIndex
        If idx < 0 OrElse idx >= _rowCount Then
            HideTooltip()
            Return
        End If
        If idx = _hoverShownIndex Then Return

        Dim row = _table.Rows(idx)
        Dim tipText = "fullpath: " & SafeStr(row("fullpath")) & Environment.NewLine &
                      "filename: " & SafeStr(row("filename")) & Environment.NewLine &
                      "size: " & SafeStr(row("size")) & Environment.NewLine &
                      "metadata: " & SafeStr(row("metadata"))

        Dim pt As Point = _lastMousePt + New Size(16, 20)
        _tip.Show(tipText, Me, pt, 7000)
        _hoverShownIndex = idx
    End Sub

    Private Sub HideTooltip()
        If _hoverShownIndex <> -1 Then
            _tip.Hide(Me)
            _hoverShownIndex = -1
        End If
    End Sub

    Private Function IndexFromPoint(p As Point) As Integer
        Dim vx As Integer = p.X - AutoScrollPosition.X
        Dim vy As Integer = p.Y - AutoScrollPosition.Y
        If vy < Pad OrElse vx < Pad Then Return -1
        Dim r As Integer = (vy - Pad) \ TileH
        Dim c As Integer = (vx - Pad) \ TileW
        If r < 0 OrElse c < 0 Then Return -1
        If c >= _cols Then Return -1
        Dim idx As Integer = r * _cols + c
        If idx >= _rowCount Then Return -1

        Dim tileX As Integer = Pad + c * TileW
        Dim tileY As Integer = Pad + r * TileH
        Dim rectTile As New Rectangle(tileX, tileY, ThumbW, TileH)
        Dim vpt As New Point(vx, vy)
        If Not rectTile.Contains(vpt) Then Return -1

        Return idx
    End Function

    ' === Helpers ===
    Private Shared Function SafeStr(o As Object) As String
        If o Is Nothing OrElse o Is DBNull.Value Then Return String.Empty
        Return Convert.ToString(o)
    End Function

    ' --- Prioritized queue management ---
    Private Sub QueueVisibleAndBuffer()
        If _table Is Nothing OrElse _rowCount = 0 OrElse _queue Is Nothing Then Return

        Dim viewTop As Integer = -AutoScrollPosition.Y
        Dim viewBottom As Integer = viewTop + ClientSize.Height

        UpdateActiveWindow(viewTop, viewBottom)
        DrainQueue()

        Dim firstVisibleRow As Integer = Math.Max(0, (viewTop - Pad) \ TileH)
        Dim lastVisibleRow As Integer = Math.Min(Math.Max(0, _rows - 1), (viewBottom - Pad) \ TileH)

        For r = firstVisibleRow To lastVisibleRow
            For c = 0 To _cols - 1
                Dim idx As Integer = r * _cols + c
                If idx >= _rowCount Then Exit For
                TryEnqueue(idx)
            Next
        Next

        For r = firstVisibleRow - 1 To Math.Max(0, firstVisibleRow - PreloadRowsAbove) Step -1
            For c = 0 To _cols - 1
                Dim idx As Integer = r * _cols + c
                If idx >= _rowCount Then Exit For
                TryEnqueue(idx)
            Next
        Next

        For r = lastVisibleRow + 1 To Math.Min(Math.Max(0, _rows - 1), lastVisibleRow + PreloadRowsBelow)
            For c = 0 To _cols - 1
                Dim idx As Integer = r * _cols + c
                If idx >= _rowCount Then Exit For
                TryEnqueue(idx)
            Next
        Next
    End Sub

    Private Sub DrainQueue()
        If _queue Is Nothing Then Return
        Dim tmp As Integer
        While _queue.TryTake(tmp)
            Dim dummy As Byte
            _pending.TryRemove(tmp, dummy)
        End While
    End Sub

    Private Sub UpdateActiveWindow(viewTop As Integer, viewBottom As Integer)
        Dim firstVisibleRow As Integer = Math.Max(0, (viewTop - Pad) \ TileH)
        Dim lastVisibleRow As Integer = Math.Min(Math.Max(0, _rows - 1), (viewBottom - Pad) \ TileH)
        Dim startRow As Integer = Math.Max(0, firstVisibleRow - PreloadRowsAbove)
        Dim endRow As Integer = Math.Min(Math.Max(0, _rows - 1), lastVisibleRow + PreloadRowsBelow)
        Dim firstIdx As Integer = startRow * _cols
        Dim lastIdx As Integer = Math.Min(_rowCount - 1, ((endRow + 1) * _cols) - 1)
        _activeFirstIndex = firstIdx
        _activeLastIndex = lastIdx
    End Sub

    Private Function IsInActiveWindow(index As Integer) As Boolean
        Return index >= _activeFirstIndex AndAlso index <= _activeLastIndex
    End Function

    Private Sub TryEnqueue(index As Integer)
        If index < 0 OrElse index >= _rowCount Then Return
        If _cache.Contains(index) Then Return
        If _pending.TryAdd(index, 0) Then _queue.TryAdd(index)
    End Sub


    'drag drop stuff


    Private Function GetDragFilePathsAt(pt As Point) As List(Of String)
        Dim results As New List(Of String)()

        Dim SubAddPath As Action(Of String) =
        Sub(p As String)
            If String.IsNullOrWhiteSpace(p) Then Return
            Dim s = p.Trim()
            If IO.File.Exists(s) OrElse IsCompositeZipPath(s) Then
                results.Add(s)
            End If
        End Sub

        ' 1) Multi-selection (if any)
        If _selectedIndices IsNot Nothing AndAlso _selectedIndices.Count > 0 Then
            For Each i In _selectedIndices
                SubAddPath(GetPathForIndex(i))
            Next
        End If

        ' 2) Fallback: hit item
        If results.Count = 0 Then
            Dim hitIdx = HitTestThumbIndex(pt)
            If hitIdx <> -1 Then SubAddPath(GetPathForIndex(hitIdx))
        End If

        ' Dedup
        results = results.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        Return results
    End Function



    Private Function HitTestThumbIndex(pt As Point) As Integer
        Return IndexFromPoint(pt)
    End Function

    Private Function GetPathForIndex(idx As Integer) As String
        If idx < 0 OrElse idx >= _rowCount Then Return Nothing

        ' Prefer the cache (fast, already stored)
        Dim p As String = _cache.GetFullPath(idx)
        If Not String.IsNullOrWhiteSpace(p) Then Return p

        ' Fallback to backing table if the item hasn't hit cache yet
        Dim row = _table.Rows(idx)
        p = SafeStr(row("fullpath"))
        If Not String.IsNullOrWhiteSpace(p) Then Return p

        Return Nothing
    End Function

    Private Sub SetSingleSelection(index As Integer, Optional ensureVisible As Boolean = False)
        If index < 0 OrElse index >= _rowCount Then Return
        Dim prev = _selectedIndices.ToList() ' snapshot for diff
        _selectedIndices.Clear()
        _selectedIndices.Add(index)
        _anchorIndex = index
        SetSelectedIndex(index, ensureVisible)
        InvalidateSelectionDiff(prev)
    End Sub

    ' Invalidate only the tiles whose selection state changed
    Private Sub InvalidateSelectionDiff(previous As IEnumerable(Of Integer))
        Dim beforeSet As New HashSet(Of Integer)(previous)
        Dim afterSet As New HashSet(Of Integer)(_selectedIndices)

        ' Symmetric difference = changed tiles
        For Each i In beforeSet.ToArray()
            If afterSet.Contains(i) Then
                beforeSet.Remove(i)  ' unchanged
                afterSet.Remove(i)
            End If
        Next

        ' Remaining in beforeSet = deselected; in afterSet = newly selected
        For Each i In beforeSet
            InvalidateTile(i)
        Next
        For Each i In afterSet
            InvalidateTile(i)
        Next
    End Sub

    Sub CopySelectedToClipboard()
        ' Gather selected paths (fall back to caret if none)
        Dim all As New List(Of String)()

        If _selectedIndices IsNot Nothing AndAlso _selectedIndices.Count > 0 Then
            For Each i In _selectedIndices
                Dim p = GetPathForIndex(i)
                If Not String.IsNullOrWhiteSpace(p) Then all.Add(p.Trim())
            Next
        ElseIf _selectedIndex >= 0 Then
            Dim p = GetPathForIndex(_selectedIndex)
            If Not String.IsNullOrWhiteSpace(p) Then all.Add(p.Trim())
        End If

        all = all.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        If all.Count = 0 Then Exit Sub

        Dim real = all.Where(Function(p) IO.File.Exists(p)).ToList()

        Dim data As New DataObject()
        If real.Count > 0 Then
            Dim sc As New Specialized.StringCollection()
            sc.AddRange(real.ToArray())
            data.SetFileDropList(sc)
        End If

        ' Always include our full list in the custom format (so USB paste works)
        data.SetData(DATAFMT_VIRTUAL_LIST, False, all)
        data.SetText(String.Join(Environment.NewLine, all))

        Try
            Clipboard.SetDataObject(data, True) ' persist even if app closes
        Catch
            ' ignore clipboard races
        End Try
    End Sub




    ' === LRU cache (index -> Image) ===
    Private Class LruImageCache
        Implements IDisposable

        Private ReadOnly _cap As Integer
        Private ReadOnly _dict As New Dictionary(Of Integer, LinkedListNode(Of Entry))()
        Private ReadOnly _list As New LinkedList(Of Entry)()
        Private ReadOnly _lock As New Object()

        Private Class Entry
            Public Property Key As Integer
            Public Property Img As Image
            Public Property FilePath As String
        End Class

        Public Sub New(capacity As Integer)
            _cap = Math.Max(8, capacity)
        End Sub

        Public Function Contains(key As Integer) As Boolean
            SyncLock _lock
                Return _dict.ContainsKey(key)
            End SyncLock
        End Function

        Public Sub Add(key As Integer, img As Image, filepath As String)
            SyncLock _lock
                If _dict.ContainsKey(key) Then
                    Dim node = _dict(key)
                    node.Value.Img?.Dispose()
                    node.Value.Img = img
                    node.Value.FilePath = filepath
                    _list.Remove(node)
                    _list.AddFirst(node)
                Else
                    Dim ent As New Entry With {.Key = key, .Img = img, .FilePath = filepath}
                    Dim node = _list.AddFirst(ent)
                    _dict(key) = node
                    If _dict.Count > _cap Then
                        Dim tail = _list.Last
                        If tail IsNot Nothing Then
                            _list.RemoveLast()
                            _dict.Remove(tail.Value.Key)
                            tail.Value.Img?.Dispose()
                        End If
                    End If
                End If
            End SyncLock
        End Sub

        Public Function GetImage(key As Integer) As Image
            SyncLock _lock
                Dim node As LinkedListNode(Of Entry) = Nothing
                If Not _dict.TryGetValue(key, node) Then Return Nothing
                _list.Remove(node)
                _list.AddFirst(node)
                Return node.Value.Img
            End SyncLock
        End Function

        Public Function GetFullPath(key As Integer) As String
            SyncLock _lock
                Dim node As LinkedListNode(Of Entry) = Nothing
                If Not _dict.TryGetValue(key, node) Then Return Nothing
                _list.Remove(node)
                _list.AddFirst(node)
                Return node.Value.FilePath
            End SyncLock
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            SyncLock _lock
                For Each n In _list
                    n.Img?.Dispose()
                Next
                _list.Clear()
                _dict.Clear()
            End SyncLock
        End Sub
    End Class ' LruImageCache

    Private Shared Function IsCompositeZipPath(s As String) As Boolean
        Try
            Dim zp As New ZipProcessing()
            Dim t = zp.ParseFilename(s)
            Return Not String.IsNullOrEmpty(t.Item1) AndAlso
                   String.Equals(t.Item2, ".zip", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.IsNullOrEmpty(t.Item3)
        Catch
            Return False
        End Try
    End Function

    ' Return fullpaths for currently visible tiles (safe, best-effort).
    Private Function GetVisibleFilePathsSafe() As List(Of String)
        Dim list As New List(Of String)
        Try
            If _rowCount <= 0 OrElse _cols <= 0 Then Return list

            Dim viewTop As Integer = -AutoScrollPosition.Y
            Dim viewBottom As Integer = viewTop + ClientSize.Height

            Dim firstVisibleRow As Integer = Math.Max(0, (viewTop - Pad) \ TileH)
            Dim lastVisibleRow As Integer = Math.Min(Math.Max(0, _rows - 1), (viewBottom - Pad) \ TileH)

            For r = firstVisibleRow To lastVisibleRow
                For c = 0 To _cols - 1
                    Dim idx As Integer = r * _cols + c
                    If idx >= _rowCount Then Exit For
                    Dim p As String = GetPathForIndex(idx)
                    If Not String.IsNullOrWhiteSpace(p) Then list.Add(p)
                Next
            Next
        Catch
            ' ignore; best-effort only
        End Try
        Return list
    End Function


    ' Re-read thumbnail size from My.Settings, refresh layout, and optionally
    ' invalidate and warm the global thumbnail cache for the new size.
    Public Sub ReloadThumbSettings(Optional regenerate As Boolean = False)
        Dim old As Size = _thumbSize
        _thumbSize = New Size(ThumbW, ThumbH)

        ResetLayoutMetrics()
        QueueVisibleAndBuffer()
        Invalidate()

        If regenerate AndAlso (old.Width <> _thumbSize.Width OrElse old.Height <> _thumbSize.Height) Then
            EmbThumbnail.ThumbCacheInvalidateBySize(old.Width, old.Height)

            ' warm only what's visible to keep it snappy
            Dim visible As List(Of String) = GetVisibleFilePathsSafe()
            If visible IsNot Nothing AndAlso visible.Count > 0 Then
                EmbThumbnail.ThumbCacheWarmAsync(visible, _thumbSize.Width, _thumbSize.Height, padding:=6, drawBorder:=True, showName:=False)
            End If
        End If
    End Sub



End Class
