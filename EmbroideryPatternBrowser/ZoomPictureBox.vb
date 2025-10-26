Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class ZoomPictureBox
    Inherits Control

    Private _image As Image
    Private _pan As PointF = PointF.Empty
    Private _panning As Boolean = False
    Private _lastMouse As Point
    Public Event ZoomChanged(ByVal newZoom As Single)
    Private _zoom As Single = 1.0F
    Public Property LogicalSourceSize As SizeF = SizeF.Empty


    Public ReadOnly Property ZoomFactor As Single
        Get
            Return _zoom
        End Get
    End Property

    Public Property Zoom As Double
        Get
            Return _zoom
        End Get
        Set(value As Double)
            If value <= 0 Then value = 0.01F
            If Math.Abs(_zoom - value) > 0.0001F Then
                _zoom = value
                RaiseEvent ZoomChanged(_zoom)
                Invalidate()
            End If
        End Set
    End Property

   Public Property Image As Image
    Get
        Return _image
    End Get
    Set(value As Image)
        _image = value
            Invalidate()
        End Set
End Property


    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.UserPaint Or
                 ControlStyles.ResizeRedraw, True)
        TabStop = True
        BackColor = Color.White
    End Sub

    ' Fit the image fully inside the control (no up-scale)
    Public Sub FitToWindow()
        If _image Is Nothing OrElse ClientSize.Width <= 0 OrElse ClientSize.Height <= 0 Then Return

        ' Use logical size if provided (stitches pane), else fall back to image size.
        Dim srcW As Double = If(Not LogicalSourceSize.IsEmpty, LogicalSourceSize.Width, _image.Width)
        Dim srcH As Double = If(Not LogicalSourceSize.IsEmpty, LogicalSourceSize.Height, _image.Height)

        Dim sx = ClientSize.Width / srcW
        Dim sy = ClientSize.Height / srcH
        _zoom = Math.Min(1.0, Math.Min(sx, sy))

        ' center using the same logical size that determined zoom
        Dim imgW = CSng(srcW * _zoom)
        Dim imgH = CSng(srcH * _zoom)
        _pan = New PointF((ClientSize.Width - imgW) / 2.0F, (ClientSize.Height - imgH) / 2.0F)
        Invalidate()
    End Sub

    'Protected Overrides Sub OnPaint(e As PaintEventArgs)
    '    MyBase.OnPaint(e)
    '    If Me.Image Is Nothing Then Return

    '    Dim g = e.Graphics

    '    e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
    '    e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
    '    e.Graphics.PixelOffsetMode = Drawing2D.PixelOffsetMode.Default
    '    e.Graphics.CompositingQuality = Drawing2D.CompositingQuality.HighQuality

    '    Dim dest = New RectangleF(_pan.X, _pan.Y, CSng(_image.Width * _zoom), CSng(_image.Height * _zoom))
    '    e.Graphics.DrawImage(_image, dest)
    'End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        If Me.Image Is Nothing Then Return

        Dim g = e.Graphics
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
        g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality

        ' Draw to the *logical* size so we can supply a higher-res backing bitmap
        Dim lw As Single = If(LogicalSourceSize.IsEmpty, Me.Image.Width, LogicalSourceSize.Width)
        Dim lh As Single = If(LogicalSourceSize.IsEmpty, Me.Image.Height, LogicalSourceSize.Height)

        Dim dest = New RectangleF(_pan.X, _pan.Y, lw * _zoom, lh * _zoom)
        g.DrawImage(Me.Image, dest)
    End Sub



    Protected Overrides Sub OnMouseWheel(e As MouseEventArgs)
        MyBase.OnMouseWheel(e)
        If _image Is Nothing Then Return
        Focus()

        Dim steps = Math.Sign(e.Delta)
        If steps = 0 Then Return

        Dim factor As Double = If(steps > 0, 1.15, 1.0 / 1.15)
        ZoomAt(e.Location, factor)
    End Sub

    Private Sub ZoomAt(mouseClient As Point, factor As Double)
        Dim oldZoom = _zoom
        Dim newZoom = Math.Max(0.01, Math.Min(100.0, oldZoom * factor))
        If Math.Abs(newZoom - oldZoom) < 0.0001 Then Return

        Dim imgX As Double = (mouseClient.X - _pan.X) / oldZoom
        Dim imgY As Double = (mouseClient.Y - _pan.Y) / oldZoom

        _zoom = newZoom
        RaiseEvent ZoomChanged(_zoom)   ' <-- add this

        _pan.X = CSng(mouseClient.X - imgX * _zoom)
        _pan.Y = CSng(mouseClient.Y - imgY * _zoom)
        Invalidate()
    End Sub



    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If e.Button = MouseButtons.Left Then
            _panning = True
            _lastMouse = e.Location
            Cursor = Cursors.SizeAll
            Focus()
        End If
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        If _panning Then
            Dim dx = e.X - _lastMouse.X
            Dim dy = e.Y - _lastMouse.Y
            _pan.X += dx
            _pan.Y += dy
            _lastMouse = e.Location
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        If e.Button = MouseButtons.Left Then
            _panning = False
            Cursor = Cursors.Default
        End If
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        ' Keep current view anchored at center on resize; or call FitToWindow() if you prefer auto-refit
        Invalidate()
    End Sub
End Class
