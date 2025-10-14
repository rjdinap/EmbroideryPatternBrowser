
Public Class AboutForm

    Public t = New System.Windows.Forms.Timer
    Public t2 = New System.Windows.Forms.Timer
    Dim msg As String = ""
    Dim msg2 As String = ""


    Private Sub AboutFormOpen(sender As Object, e As EventArgs) Handles Me.Shown
        AddHandler DirectCast(t, System.Windows.Forms.Timer).Tick, AddressOf ColorChanger
        t.Interval = 100
        t.Start()
        Dim msg As String = "2025 Robert DiNapoli" + vbCrLf + vbCrLf
        msg = msg + "A program to search for embroidery patterns on your drive super fast!." + vbCrLf + vbCrLf
        msg = msg + "Executable path: " + Application.ExecutablePath + vbCrLf + vbCrLf
        msg = msg + "Assmebly version: " + System.Reflection.Assembly.GetExecutingAssembly.GetName().Version.ToString
        RichTextBox_About_About.Text = msg

    End Sub





    Private Sub ColorChanger(sender As Object, e As EventArgs)
        Static Dim red1 As Integer = 0
        Static Dim blue1 As Integer = 0
        Static Dim green1 As Integer = 0
        Static Dim redDir As Integer = CInt(Math.Ceiling(Rnd() * 1))  ' forward , set to 1 for reverse
        Static Dim blueDir As Integer = CInt(Math.Ceiling(Rnd() * 1)) ' forward , set to 1 for reverse
        Static Dim greenDir As Integer = CInt(Math.Ceiling(Rnd() * 1)) ' forward , set to 1 for reverse
        Static Dim redVelocity As Integer = CInt(Math.Ceiling(Rnd() * 5)) + 1
        Static Dim greenVelocity As Integer = CInt(Math.Ceiling(Rnd() * 5)) + 1
        Static Dim blueVelocity As Integer = CInt(Math.Ceiling(Rnd() * 5)) + 1

        'starting values
        If red1 = 0 Then
            red1 = CInt(Math.Ceiling(Rnd() * 200)) + 1
        End If
        If blue1 = 0 Then
            blue1 = CInt(Math.Ceiling(Rnd() * 200)) + 1
        End If
        If green1 = 0 Then
            green1 = CInt(Math.Ceiling(Rnd() * 200)) + 1
        End If


        SetBlendingColor(red1, green1, blue1, 0, 0)

        If redDir = 0 Then
            red1 = red1 + redVelocity
        Else
            red1 = red1 - redVelocity
        End If
        If red1 > 200 Then redDir = 1 ' reverse
        If red1 < 50 Then redDir = 0 ' forward

        If blueDir = 0 Then
            blue1 = blue1 + blueVelocity
        Else
            blue1 = blue1 - blueVelocity
        End If
        If blue1 > 200 Then blueDir = 1 ' reverse
        If blue1 < 50 Then blueDir = 0 ' forward



        If greenDir = 0 Then
            green1 = green1 + greenVelocity
        Else
            green1 = green1 - greenVelocity
        End If
        If green1 > 200 Then greenDir = 1 ' reverse
        If green1 < 50 Then greenDir = 0 ' forward



    End Sub



    Private Sub Button_About_Close_Click(sender As Object, e As EventArgs) Handles Button_About_Close.Click
        DirectCast(t, System.Windows.Forms.Timer).Enabled = False
        DirectCast(t, System.Windows.Forms.Timer).Dispose()
        Me.Close()
    End Sub



    Public Sub SetBlendingColor(red As Integer, green As Integer, blue As Integer, brightness As Integer, alpha As Integer)

        If Me.FormAbout_PictureBox.Image IsNot Nothing Then

            Dim sr, sg, sb, sa, sbr As Single

            'create a bitmap from our original image
            Using origBitmap As New Bitmap(My.Resources.ring)
                Dim cm As New Drawing.Imaging.ColorMatrix
                Dim atr As New Drawing.Imaging.ImageAttributes

                'create a bitmap and graphics object
                Dim newBitmap = New Bitmap(origBitmap.Width, origBitmap.Height)
                Dim colorBoxGraphics As Graphics = Graphics.FromImage(newBitmap)
                Dim r As New Rectangle(0, 0, newBitmap.Width, newBitmap.Height)

                'noramlize the color components to 1
                sr = red / 255
                sg = green / 255
                sb = blue / 255
                sa = alpha / 255
                sbr = brightness / 255

                'create the color matrix
                cm = New Drawing.Imaging.ColorMatrix(New Single()() _
                                   {New Single() {1, 0, 0, 0, 0},
                                    New Single() {0, 1, 0, 0, 0},
                                    New Single() {0, 0, 1, 0, 0},
                                    New Single() {0, 0, 0, 1, 0},
                                    New Single() {sr + sbr, sg + sbr, sb + sbr, sa, 1}})
                atr.SetColorMatrix(cm)
                colorBoxGraphics.DrawImage(origBitmap, r, 0, 0, r.Width, r.Height, Drawing.GraphicsUnit.Pixel, atr)


                Me.FormAbout_PictureBox.Image = newBitmap
            End Using
            GC.Collect()
        End If
    End Sub



End Class