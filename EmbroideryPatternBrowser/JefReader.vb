Public Class JefReader
    Inherits EmbReader

    ' Conservative safety caps
    Private Const MAX_COLORS As Integer = 4096
    Private Const MAX_STITCH_BYTES As Integer = 128 * 1024 * 1024 ' 128MB decode guard
    Private Const MAX_LOOP_ITERS As Integer = 50 * 1024 * 1024     ' hard loop limiter

    Public Overrides Sub Read()
        Try
            Dim jefThreads As EmbThread() = EmbThreadJef.GetThreadSet()
            If jefThreads Is Nothing OrElse jefThreads.Length = 0 Then
                jefThreads = New EmbThread() {New EmbThread(&H0, "JEF Default", "", "Jef", "Jef")}
            End If

            ' --- Header ---
            Dim stitchOffset As Integer = ReadInt32LE()
            Skip(20)
            Dim colorCount As Integer = ReadInt32LE()
            Skip(88)

            ' Clamp color count to guard corrupt files
            colorCount = Clamp(colorCount, 0, MAX_COLORS)

            ' --- Thread list (0 index means STOP placeholder) ---
            For i As Integer = 0 To colorCount - 1
                Dim indexRaw As Integer = ReadInt32LE()
                Dim index As Integer = Math.Abs(indexRaw)
                If index = 0 Then
                    ' STOP placeholder; renderer logic handles a Nothing entry
                    pattern.ThreadList.Add(Nothing)
                Else
                    Dim t As EmbThread = jefThreads(index Mod jefThreads.Length)
                    pattern.Add(t)
                End If
            Next

            If pattern.ThreadList.Count > 0 Then
                pattern.select_thread(0)
            End If

            ' --- Validate stitchOffset and seek safely ---
            Dim fileLen As Long = If(reader IsNot Nothing AndAlso reader.BaseStream IsNot Nothing, reader.BaseStream.Length, 0L)
            If stitchOffset < 0 OrElse stitchOffset > fileLen Then
                Form1.StatusFromAnyThread("Error: Invalid JEF stitch offset.")
                pattern.end()
                Return
            End If
            Seek(stitchOffset)

            ReadJefStitchesSafe(fileLen)

        Catch ex As Exception
            Try : Form1.StatusFromAnyThread("Error: " & ex.Message, ex.StackTrace.ToString) : Catch : End Try
            Try : pattern.end() : Catch : End Try
        End Try
    End Sub

    Private Sub ReadJefStitchesSafe(fileLen As Long)
        Dim colorIndex As Integer = 1
        Dim b(1) As Byte

        Dim loopIters As Integer = 0
        Dim bytesRead As Long = 0
        Dim lastPos As Long = reader.BaseStream.Position

        While True
            If loopIters >= MAX_LOOP_ITERS Then
                Form1.StatusFromAnyThread("Error: JEF decode iteration limit reached.")
                Exit While
            End If
            loopIters += 1

            If ReadFully(b) <> b.Length Then Exit While
            ReadPosition += 2
            bytesRead += 2

            ' Progress guard
            If reader.BaseStream.Position = lastPos Then
                Form1.StatusFromAnyThread("Error: No progress while reading JEF stream.")
                Exit While
            End If
            lastPos = reader.BaseStream.Position

            If bytesRead > MAX_STITCH_BYTES Then
                Form1.StatusFromAnyThread("Error: JEF stitch data exceeds maximum size.")
                Exit While
            End If

            ' Regular stitch (2-byte command): signed 8-bit deltas
            If (b(0) And &HFF) <> &H80 Then
                Dim dx As Integer = Sign8(b(0))
                Dim dy As Integer = -Sign8(b(1))
                pattern.stitch(dx, dy)
                Continue While
            End If

            ' Control byte follows 0x80
            Dim ctrl As Integer = b(1) And &HFF

            ' Next 2 bytes are a coordinate pair
            If ReadFully(b) <> b.Length Then Exit While
            ReadPosition += 2
            bytesRead += 2

            Dim dx2 As Integer = Sign8(b(0))
            Dim dy2 As Integer = -Sign8(b(1))

            Select Case ctrl
                Case &H2          ' Trim or jump
                    If dx2 = 0 AndAlso dy2 = 0 Then
                        pattern.trim()
                    Else
                        pattern.move(dx2, dy2)
                    End If

                Case &H1          ' Color change or STOP placeholder
                    If colorIndex < pattern.ThreadList.Count AndAlso pattern.ThreadList(colorIndex) IsNot Nothing Then
                        pattern.color_change()
                        colorIndex += 1
                    Else
                        ' STOP: visual break, keep same color; drop placeholder so indices align
                        pattern.trim()
                        If colorIndex < pattern.ThreadList.Count Then
                            pattern.ThreadList.RemoveAt(colorIndex)
                        End If
                        ' do not advance colorIndex
                    End If

                Case &H10         ' End of design
                    Exit While

                Case Else
                    ' Unknown control; stop safely
                    Exit While
            End Select

            ' Additional forward-progress/EOF guard
            If reader.BaseStream.Position >= fileLen Then Exit While
        End While

        pattern.end()
    End Sub

    Private Function Sign8(n As Integer) As Integer
        n = n And &HFF
        If (n And &H80) <> 0 Then n -= &H100
        Return n
    End Function

    Private Shared Function Clamp(v As Integer, lo As Integer, hi As Integer) As Integer
        If v < lo Then Return lo
        If v > hi Then Return hi
        Return v
    End Function
End Class
