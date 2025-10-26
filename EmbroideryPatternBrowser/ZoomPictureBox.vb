Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class ZoomPictureBox
    Inherits Control

    Private _image As Image
    Private _zoom As Double = 1.0      ' current zoom (1.0 = 100%)
    Private _pan As PointF = PointF.Empty
    Private _panning As Boolean = False
    Private _lastMouse As Point

    Public Property Image As Image
        Get
            Return _image
        End Get
        Set(value As Image)
            _image = value
            FitToWindow()
            Invalidate()
        End Set
    End Property

    Public Property Zoom As Double
        Get
            Return _zoom
        End Get
        Set(value As Double)
            _zoom = Math.Max(0.01, Math.Min(100.0, value))
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
        Dim sx = ClientSize.Width / CDbl(_image.Width)
        Dim sy = ClientSize.Height / CDbl(_image.Height)
        _zoom = Math.Min(1.0, Math.Min(sx, sy))
        ' center it
        Dim imgW = _image.Width * _zoom
        Dim imgH = _image.Height * _zoom
        _pan = New PointF((ClientSize.Width - imgW) / 2.0F, (ClientSize.Height - imgH) / 2.0F)
        Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        e.Graphics.Clear(BackColor)

        If _image Is Nothing Then
            Using f = New Font(Font, FontStyle.Italic)
                TextRenderer.DrawText(e.Graphics, "No image", f, ClientRectangle, Color.Gray,
                                      TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            End Using
            Return
        End If

        e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
        e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality

        ' Destination rectangle based on zoom + pan
        Dim dest = New RectangleF(_pan.X, _pan.Y, CSng(_image.Width * _zoom), CSng(_image.Height * _zoom))
        e.Graphics.DrawImage(_image, dest)
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
        ' Keep the point under the cursor stable while zooming
        Dim oldZoom = _zoom
        Dim newZoom = Math.Max(0.01, Math.Min(100.0, oldZoom * factor))
        If Math.Abs(newZoom - oldZoom) < 0.0001 Then Return

        ' Convert mouse point to image-space coordinates before zoom
        Dim imgX As Double = (mouseClient.X - _pan.X) / oldZoom
        Dim imgY As Double = (mouseClient.Y - _pan.Y) / oldZoom

        _zoom = newZoom

        ' Recompute pan so the same image point stays under the cursor
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
