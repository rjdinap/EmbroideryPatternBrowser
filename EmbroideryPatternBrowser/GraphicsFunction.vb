Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging



Public Class GraphicsFunctions
    ''' <summary>
    ''' Rotates an image by angleDegrees (clockwise), returns a new image with a canvas sized to avoid clipping.
    ''' </summary>
    Public Shared Function RotateImage(src As Image, angleDegrees As Single, Optional background As Color = Nothing) As Image
        If src Is Nothing Then Throw New ArgumentNullException(NameOf(src))
        If background = Nothing Then background = Color.Transparent

        Dim rad As Double = angleDegrees * Math.PI / 180.0
        Dim cos As Double = Math.Abs(Math.Cos(rad))
        Dim sin As Double = Math.Abs(Math.Sin(rad))

        Dim newW As Integer = CInt(Math.Ceiling(src.Width * cos + src.Height * sin))
        Dim newH As Integer = CInt(Math.Ceiling(src.Width * sin + src.Height * cos))

        Dim bmp As New Bitmap(newW, newH, PixelFormat.Format32bppArgb)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.InterpolationMode = InterpolationMode.HighQualityBicubic
            g.PixelOffsetMode = PixelOffsetMode.HighQuality
            g.Clear(background)

            ' Move origin to center, rotate, then draw centered
            g.TranslateTransform(newW / 2.0F, newH / 2.0F)
            g.RotateTransform(angleDegrees)
            g.TranslateTransform(-src.Width / 2.0F, -src.Height / 2.0F)
            g.DrawImage(src, New Rectangle(0, 0, src.Width, src.Height), New Rectangle(0, 0, src.Width, src.Height), GraphicsUnit.Pixel)
            g.ResetTransform()
        End Using
        Return bmp
    End Function


    ''' Resize to an explicit W×H (preserving aspect is up to caller).
    Public Function ResizeTo(src As Image,
                             targetWidth As Integer,
                             targetHeight As Integer,
                             Optional highQuality As Boolean = True) As Bitmap
        If src Is Nothing Then Throw New ArgumentNullException(NameOf(src))
        Dim w As Integer = Math.Max(1, targetWidth)
        Dim h As Integer = Math.Max(1, targetHeight)

        Dim bmp As New Bitmap(w, h, PixelFormat.Format32bppPArgb)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.CompositingMode = CompositingMode.SourceOver
            g.CompositingQuality = If(highQuality, CompositingQuality.HighQuality, CompositingQuality.HighSpeed)
            g.InterpolationMode = If(highQuality, InterpolationMode.HighQualityBicubic, InterpolationMode.NearestNeighbor)
            g.SmoothingMode = If(highQuality, SmoothingMode.AntiAlias, SmoothingMode.None)
            g.PixelOffsetMode = If(highQuality, PixelOffsetMode.HighQuality, PixelOffsetMode.None)
            g.DrawImage(src, New Rectangle(0, 0, w, h))
        End Using
        Return bmp
    End Function



    ''' Resize by a scale factor. Example: 1.25 = +25%, 0.5 = 50%.
    Public Function ResizeBy(src As Image,
                             scale As Single,
                             Optional highQuality As Boolean = True) As Bitmap
        If src Is Nothing Then Throw New ArgumentNullException(NameOf(src))
        Dim s As Single = Math.Max(0.01F, scale)
        Dim w As Integer = Math.Max(1, CInt(Math.Round(src.Width * s)))
        Dim h As Integer = Math.Max(1, CInt(Math.Round(src.Height * s)))
        Return ResizeTo(src, w, h, highQuality)
    End Function



    ' Per-control zoom state (so switching images/controls won’t share stale values)
    Private Class ZoomState
        Public Zoom As Double = 1.0
        Public MinZoom As Double = 0.05
        Public MaxZoom As Double = 10.0
    End Class




End Class



