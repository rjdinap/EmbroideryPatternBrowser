Imports System

Public Class SewReader
    Inherits EmbReader

    ' Conservative sanity caps
    Private Const MAX_COLORS As Integer = 4096
    Private Const MAX_STITCH_ITERS As Integer = 50 * 1024 * 1024   ' hard loop limiter
    Private Const STITCH_OFFSET As Integer = &H1D78                 ' format-fixed start

    Public Overrides Sub Read()
        Try
            ' Threads palette per header (with safe fallback)
            Dim threads As EmbThread() = EmbThreadSew.GetThreadSet()
            If threads Is Nothing OrElse threads.Length = 0 Then
                threads = New EmbThread() {New EmbThread(&H0, "SEW Default", "", "Sew", "Sew")}
            End If

            ' Color count (clamped)
            Dim numberOfColors As Integer = ReadInt16LE()
            If numberOfColors < 0 Then numberOfColors = 0
            numberOfColors = Math.Min(numberOfColors, MAX_COLORS)

            ' Palette indices (guarded)
            For i As Integer = 0 To numberOfColors - 1
                Dim idxVal As Integer = ReadInt16LE()
                If threads.Length > 0 Then
                    Dim t As EmbThread = threads(Math.Abs(idxVal) Mod threads.Length)
                    pattern.Add(t)
                End If
            Next

            ' Validate and seek to stitch data
            Dim fileLen As Long = If(reader IsNot Nothing AndAlso reader.BaseStream IsNot Nothing, reader.BaseStream.Length, 0L)
            If STITCH_OFFSET < 0 OrElse STITCH_OFFSET > fileLen Then
                Form1.Status("Error: Invalid SEW stitch offset.")
                pattern.end()
                Return
            End If
            Seek(STITCH_OFFSET)

            ReadSewStitchesSafe(fileLen)

        Catch ex As Exception
            Try : Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString) : Catch : End Try
            Try : pattern.end() : Catch : End Try
        End Try
    End Sub

    Private Sub ReadSewStitchesSafe(fileLen As Long)
        Dim b(1) As Byte
        Dim iters As Integer = 0
        Dim lastPos As Long = reader.BaseStream.Position

        While True
            If iters >= MAX_STITCH_ITERS Then
                Form1.Status("Error: SEW decode iteration limit reached.")
                Exit While
            End If
            iters += 1

            If ReadFully(b) <> b.Length Then Exit While
            ReadPosition += 2

            ' Progress/EOF guards
            If reader.BaseStream.Position = lastPos Then
                Form1.Status("Error: No progress while reading SEW stream.")
                Exit While
            End If
            lastPos = reader.BaseStream.Position
            If lastPos >= fileLen Then Exit While

            ' If first byte is not 0x80, it's a simple stitch (signed deltas)
            If (b(0) And &HFF) <> &H80 Then
                pattern.stitch(Sign8(b(0)), -Sign8(b(1)))
                Continue While
            End If

            ' Control byte follows 0x80
            Dim control As Integer = b(1) And &HFF

            ' Read second pair that contains deltas for the control
            If ReadFully(b) <> b.Length Then Exit While
            ReadPosition += 2
            lastPos = reader.BaseStream.Position

            If (control And &H1) <> 0 Then
                ' Color change
                pattern.color_change()
                Continue While
            End If

            If (control = &H4) OrElse (control = &H2) Then
                ' Move (jump)
                pattern.move(Sign8(b(0)), -Sign8(b(1)))
                Continue While
            End If

            If control = &H10 Then
                ' Explicit stitch
                pattern.stitch(Sign8(b(0)), -Sign8(b(1)))
                Continue While
            End If

            ' Any other control means end
            Exit While
        End While

        pattern.end()
    End Sub

    Private Function Sign8(n As Integer) As Integer
        n = n And &HFF
        If (n And &H80) <> 0 Then n -= &H100
        Return n
    End Function
End Class
