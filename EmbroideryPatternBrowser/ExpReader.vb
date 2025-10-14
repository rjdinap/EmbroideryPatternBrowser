Public Class ExpReader
    Inherits EmbReader

    Public Overrides Sub Read()
        EnsureDefaultPalette()       ' <-- add this line
        pattern.select_thread(0)     ' start on first color (safe even if already 0)
        ReadExpStitches(Me)
    End Sub

    Private Sub EnsureDefaultPalette()
        If pattern Is Nothing Then pattern = New EmbPattern()
        If pattern.ThreadList.Count > 0 Then Exit Sub

        Dim husSet As EmbThread() = EmbThreadHus.GetThreadSet()
        If husSet IsNot Nothing AndAlso husSet.Length > 0 Then
            ' Add a reasonable subset so we don’t bloat memory (EXP can change color many times)
            Dim take As Integer = Math.Min(30, husSet.Length) ' 30 is plenty to cycle through
            For i As Integer = 0 To take - 1
                pattern.Add(husSet(i))
            Next
        Else
            ' Fallback mini palette if HUS set isn’t available
            Dim defaults As Integer() = {
                &HFF0000, &HAA00, &HFF, &HFF7F00, &H800080, &H7F7F,
                &H0, &H808080, &HFFFF00, &HFF00FF, &HFFFF, &H804000
            }
            For i As Integer = 0 To defaults.Length - 1
                pattern.Add(New EmbThread(defaults(i), $"EXP {i + 1}", (i + 1).ToString(), "EXP", "EXP"))
            Next
        End If
    End Sub

    Public Shared Sub ReadExpStitches(r As EmbReader)
        Dim b(1) As Byte

        While True
            If ReadFully(r, b, 0, 2) <> 2 Then Exit While

            If (b(0) And &HFF) <> &H80 Then
                Dim x As Integer = Sign8(b(0))
                Dim y As Integer = -Sign8(b(1))
                r.pattern.stitch(x, y)
                Continue While
            End If

            Dim control As Integer = b(1) And &HFF
            If ReadFully(r, b, 0, 2) <> 2 Then Exit While
            Dim x2 As Integer = Sign8(b(0))
            Dim y2 As Integer = -Sign8(b(1))

            Select Case control
                Case &H80 ' trim
                    r.pattern.trim()

                Case &H2 ' stitch with supplied delta
                    r.pattern.stitch(x2, y2)

                Case &H4 ' move/jump with supplied delta
                    r.pattern.move(x2, y2)

                Case &H1 ' color change, optional move
                    r.pattern.color_change()
                    If x2 <> 0 OrElse y2 <> 0 Then r.pattern.move(x2, y2)

                Case Else
                    Exit While ' unknown control → stop
            End Select
        End While

        r.pattern.end()
    End Sub

    Private Shared Function ReadFully(r As EmbReader, buffer() As Byte, offset As Integer, count As Integer) As Integer
        Dim total As Integer = 0
        While total < count
            Dim got As Integer = r.reader.Read(buffer, offset + total, count - total)
            If got <= 0 Then Exit While
            total += got
            r.ReadPosition += got
        End While
        Return total
    End Function

    Private Shared Function Sign8(n As Integer) As Integer
        n = n And &HFF
        If (n And &H80) <> 0 Then n -= &H100
        Return n
    End Function
End Class
