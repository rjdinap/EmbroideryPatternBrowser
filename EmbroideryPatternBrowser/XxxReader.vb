Imports System

Public Class XxxReader
    Inherits EmbReader

    ' Conservative caps / guards
    Private Const MAX_COLORS As Integer = 4096
    Private Const MAX_STITCH_ITERS As Integer = 50 * 1024 * 1024 ' hard loop limiter

    Public Overrides Sub Read()
        Try
            ' ---- Header ----
            SafeSkip(&H27)
            Dim num_of_colors As Integer = ReadInt16LE()
            If num_of_colors < 0 Then num_of_colors = 0
            If num_of_colors > MAX_COLORS Then num_of_colors = MAX_COLORS

            ' Pre-allocate placeholder threads so color_change() advances indices correctly.
            For i As Integer = 0 To num_of_colors - 1
                pattern.Add(New EmbThread(&H0, "XXX " & (i + 1).ToString(), "", "XXX", "XXX"))
            Next
            If num_of_colors > 0 Then
                pattern.select_thread(0)
            End If

            ' Stitch block starts at 0x100; guard seek
            SafeSeek(&H100)

            ' ---- Decode stitch stream (guarded) ----
            Dim iters As Integer = 0
            Dim lastPos As Long = SafeStreamPos()

            While True
                If iters >= MAX_STITCH_ITERS Then
                    Try : Form1.Status("Error: XXX decode iteration limit reached.") : Catch : End Try
                    Exit While
                End If
                iters += 1

                Dim b1 As Integer = ReadInt8()
                If b1 = Integer.MinValue Then Exit While

                If (b1 = &H7D) OrElse (b1 = &H7E) Then
                    ' 16-bit LE long move
                    Dim x16 As Integer = Signed16(ReadInt16LE())
                    Dim y16 As Integer = Signed16(ReadInt16LE())
                    pattern.move(x16, -y16)
                Else
                    Dim b2 As Integer = ReadInt8()
                    If b2 = Integer.MinValue Then Exit While

                    If b1 <> &H7F Then
                        ' Simple stitch: signed 8-bit deltas
                        pattern.stitch(Sign8(b1), -Sign8(b2))
                    Else
                        ' Extended op 0x7F b2 b3 b4
                        Dim b3 As Integer = ReadInt8()
                        Dim b4 As Integer = ReadInt8()
                        If b3 = Integer.MinValue OrElse b4 = Integer.MinValue Then Exit While

                        Select Case (b2 And &HFF)
                            Case &H1 ' move
                                pattern.move(Sign8(b3), -Sign8(b4))

                            Case &H3 ' trim (then optional move if non-zero)
                                pattern.trim()
                                Dim x As Integer = Sign8(b3)
                                Dim y As Integer = -Sign8(b4)
                                If (x <> 0) OrElse (y <> 0) Then pattern.move(x, y)

                            Case &H8 ' color change
                                pattern.color_change()

                            Case &H7F
                                ' End of design
                                Exit While

                            Case Else
                                ' Unknown extended op; ignore and continue (no risky behavior)
                                ' Intentionally do nothing
                        End Select
                    End If
                End If

                ' forward-progress guard (corruption protection)
                Dim nowPos As Long = SafeStreamPos()
                If nowPos = lastPos Then
                    Try : Form1.Status("Error: No progress while reading XXX stream.") : Catch : End Try
                    Exit While
                End If
                lastPos = nowPos
            End While

            pattern.end()

            ' ---- Color table at end (guarded, optional) ----
            ' Reference: skip(2) then read num_of_colors * 4 (BE) colors.
            ' Only attempt if there are bytes left to avoid IO exceptions.
            If ReaderRemaining() >= 2 Then
                SafeSkip(2)
                For i As Integer = 0 To num_of_colors - 1
                    If ReaderRemaining() < 4 Then Exit For
                    Dim rgb As Integer = ReadInt32BE()
                    If rgb = Integer.MinValue Then Exit For

                    If i >= 0 AndAlso i < pattern.ThreadList.Count Then
                        Dim oldT = pattern.ThreadList(i)
                        Dim desc As String = If(oldT IsNot Nothing AndAlso oldT.Description IsNot Nothing,
                                                oldT.Description, "XXX " & (i + 1).ToString())
                        Dim code As String = If(oldT IsNot Nothing, oldT.ColorCode, "")
                        Dim brand As String = If(oldT IsNot Nothing AndAlso Not String.IsNullOrEmpty(oldT.Brand), oldT.Brand, "XXX")
                        pattern.ThreadList(i) = New EmbThread(rgb, desc, code, brand, "XXX")
                    Else
                        pattern.Add(New EmbThread(rgb, "XXX " & (i + 1).ToString(), "", "XXX", "XXX"))
                    End If
                Next
            End If

        Catch ex As Exception
            ' concise error message for UI; no stack trace (heuristic-friendly)
            Try : Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString) : Catch : End Try
            Try : pattern.end() : Catch : End Try
        End Try
    End Sub

    ' ---- helpers ----
    Private Function Sign8(n As Integer) As Integer
        n = n And &HFF
        If (n And &H80) <> 0 Then n -= &H100
        Return n
    End Function

    Private Function Signed16(v As Integer) As Integer
        If v = Integer.MinValue Then Return 0
        v = v And &HFFFF
        If (v And &H8000) <> 0 Then v -= &H10000
        Return v
    End Function

    Private Sub SafeSkip(count As Integer)
        If count <= 0 Then Return
        Dim avail As Long = ReaderRemaining()
        If avail <= 0 Then Return
        Dim n As Integer = CInt(Math.Min(avail, count))
        Skip(n)
    End Sub

    Private Sub SafeSeek(pos As Integer)
        If reader Is Nothing OrElse reader.BaseStream Is Nothing Then Return
        Dim fileLen As Long = reader.BaseStream.Length
        Dim p As Integer = Math.Max(0, Math.Min(CInt(fileLen), pos))
        Seek(p)
    End Sub

    Private Function ReaderRemaining() As Long
        If reader Is Nothing OrElse reader.BaseStream Is Nothing Then Return 0
        Return Math.Max(0L, reader.BaseStream.Length - reader.BaseStream.Position)
    End Function

    Private Function SafeStreamPos() As Long
        If reader Is Nothing OrElse reader.BaseStream Is Nothing Then Return 0
        Return reader.BaseStream.Position
    End Function
End Class
