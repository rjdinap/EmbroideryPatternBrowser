Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging


'' Rotate 45 degrees:
'Dim rotated = GraphicsFunctions.RotateImage(img, 45.0F, Color.White)
'PictureBox1.Image = rotated


Public Module GraphicsFunctions
    ''' <summary>
    ''' Rotates an image by angleDegrees (clockwise), returns a new image with a canvas sized to avoid clipping.
    ''' </summary>
    Public Function RotateImage(src As Image, angleDegrees As Single, Optional background As Color = Nothing) As Image
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



End Module
