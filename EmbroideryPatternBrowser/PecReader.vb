Imports System
Imports System.Text

Public Class PecReader
    Inherits EmbReader

    Private Const JUMP_CODE As Integer = &H10
    Private Const TRIM_CODE As Integer = &H20
    Private Const FLAG_LONG As Integer = &H80

    ' Conservative sanity caps
    Private Const MAX_COLORS As Integer = 4096
    Private Const MAX_STRIDE As Integer = 4096
    Private Const MAX_ICON_H As Integer = 4096
    Private Const MAX_GRAPHIC_BYTES As Integer = 4 * 1024 * 1024     ' 4MB cap per color graphic
    Private Const MAX_STITCH_ITERS As Integer = 50 * 1024 * 1024      ' hard loop limiter

    Public Overrides Sub Read()
        Try
            ' Some callers (e.g., PES) position the stream at the PEC header ("#PEC0001"),
            ' others position at the PEC payload just after it. Peek to decide.
            Dim mustSkip8 As Boolean = False
            If reader IsNot Nothing AndAlso reader.BaseStream IsNot Nothing AndAlso reader.BaseStream.CanSeek Then
                Dim pos As Long = reader.BaseStream.Position
                Dim hdr(7) As Byte
                Dim got As Integer = reader.Read(hdr, 0, hdr.Length)
                reader.BaseStream.Position = pos
                If got = 8 Then
                    Dim s As String = Encoding.ASCII.GetString(hdr)
                    If s = "#PEC0001" Then mustSkip8 = True
                End If
            End If

            If mustSkip8 Then
                Skip(&H8) ' consume "#PEC0001"
            End If

            ' Decode PEC payload starting at current position
            ReadPecCore()

        Catch ex As Exception
            Try : Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString) : Catch : End Try
            Try : pattern.end() : Catch : End Try
        End Try
    End Sub

    ' ---------------------- Core PEC payload decoder ----------------------
    ' This assumes the stream is already positioned at the PEC payload start
    ' (i.e., right after "#PEC0001" if that marker exists in the file).
    Private Sub ReadPecCore()
        ' --- Label / basic header ---
        Skip(3)
        Dim label As String = SafeTrim(ReadString(16))
        If label.Length > 0 Then pattern.SetMetadata(EmbPattern.PROP_NAME, label)
        Skip(&HF)

        ' Dimensions and icon info (guarded)
        Dim stride As Integer = ReadInt8()
        Dim iconH As Integer = ReadInt8()
        Skip(&HC)

        ' Color section (bounded)
        Dim colorChanges As Integer = ReadInt8()
        If colorChanges = Integer.MinValue Then
            pattern.end() : Return
        End If
        colorChanges = Clamp(colorChanges, 0, MAX_COLORS)

        Dim colorBytes As Byte()
        If colorChanges = 0 Then
            colorBytes = Array.Empty(Of Byte)()
        Else
            colorBytes = New Byte(colorChanges) {} ' +1 as in original logic
            Dim got As Integer = ReadFully(colorBytes)
            If got <> colorBytes.Length Then
                Try : Form1.Status("Error: PEC palette truncated.") : Catch : End Try
                pattern.end() : Return
            End If
        End If
        MapPecColors(colorBytes)

        ' Seek to stitches:
        ' Original: Skip(&H1D0 - colorChanges) then stitchBlockEnd = ReadInt24LE() - 5 + Tell()
        Dim skipCount As Integer = &H1D0 - colorChanges
        If skipCount < 0 Then skipCount = 0
        Skip(skipCount)

        ' The 24-bit length is relative; ensure non-negative and within file
        Dim relLen As Integer = ReadInt24LE()
        If relLen < 5 Then relLen = 5
        Dim curPos As Long = If(reader Is Nothing OrElse reader.BaseStream Is Nothing, 0L, reader.BaseStream.Position)
        Dim fileLen As Long = If(reader Is Nothing OrElse reader.BaseStream Is Nothing, 0L, reader.BaseStream.Length)
        Dim stitchBlockEnd As Long = curPos + (relLen - 5)
        If stitchBlockEnd < 0 Then stitchBlockEnd = 0
        If stitchBlockEnd > fileLen Then stitchBlockEnd = fileLen

        ' 0x31, 0xFF, 0xF0, then 4 shorts => 11 bytes (guarded)
        Skip(&HB)

        ' --- Stitches (safe loop) ---
        ReadPecStitchesSafe()

        ' Seek to computed end (if valid)
        If stitchBlockEnd >= 0 AndAlso reader IsNot Nothing AndAlso reader.BaseStream IsNot Nothing AndAlso stitchBlockEnd <= reader.BaseStream.Length Then
            Seek(CInt(stitchBlockEnd))
        End If

        ' --- Small icon graphics (guard sizes) ---
        Dim wStride As Integer = Clamp(stride, 0, MAX_STRIDE)
        Dim hIcon As Integer = Clamp(iconH, 0, MAX_ICON_H)
        Dim byteSize As Long = CLng(wStride) * CLng(hIcon)
        If byteSize > 0 AndAlso byteSize <= MAX_GRAPHIC_BYTES Then
            ReadPecGraphics(CInt(byteSize), wStride, colorChanges)
        End If
    End Sub

    ' Render the small PEC graphics; avoid storing huge text blobs
    Private Sub ReadPecGraphics(size As Integer, stride As Integer, count As Integer)
        If size <= 0 OrElse count <= 0 Then Return
        Dim graphic(size - 1) As Byte

        For i = 0 To count - 1
            Dim got As Integer = ReadFully(graphic)
            If got <> graphic.Length Then
                ' partial/truncated graphic; stop
                Exit For
            End If

            ' Only embed ASCII-art when reasonably small (<= 64*64)
            If size <= 4096 AndAlso stride > 0 Then
                Try
                    pattern.SetMetadata("pec_graphic_" & i, GetGraphicAsString(graphic, "#"c, " "c, stride))
                Catch
                    ' ignore visual metadata on failure
                End Try
            End If
        Next
    End Sub

    Private Function GetGraphicAsString(graphic As Byte(), one As Char, zero As Char, stride As Integer) As String
        Dim sb As New StringBuilder()
        Dim colStride As Integer = Math.Max(1, stride)
        For i = 0 To graphic.Length - 1
            If colStride > 0 AndAlso (i + 1) Mod colStride = 0 Then sb.AppendLine()
            Dim b As Integer = graphic(i)
            For k = 0 To 7
                sb.Append(If((b And 1) = 0, zero, one))
                b >>= 1
            Next
        Next
        Return sb.ToString()
    End Function

    ' ---------------------- Stitch decoding (hardened) ----------------------
    Private Sub ReadPecStitchesSafe()
        Dim iter As Integer = 0
        Dim lastPos As Long = If(reader Is Nothing OrElse reader.BaseStream Is Nothing, 0L, reader.BaseStream.Position)

        While True
            If iter >= MAX_STITCH_ITERS Then
                Try : Form1.Status("Error: PEC decode iteration limit reached.") : Catch : End Try
                Exit While
            End If
            iter += 1

            Dim val1 = ReadInt8()
            If val1 = Integer.MinValue Then Exit While
            If val1 = &HFF Then Exit While ' end

            Dim val2 = ReadInt8()
            If val2 = Integer.MinValue Then Exit While

            If val1 = &HFE AndAlso val2 = &HB0 Then
                ' Color change triplet (consume the extra byte)
                If ReadInt8() = Integer.MinValue Then Exit While
                pattern.color_change()
                Continue While
            End If

            Dim jump As Boolean = False
            Dim trim As Boolean = False
            Dim x As Integer
            Dim y As Integer

            If (val1 And FLAG_LONG) <> 0 Then
                If (val1 And TRIM_CODE) <> 0 Then trim = True
                If (val1 And JUMP_CODE) <> 0 Then jump = True
                Dim code12 As Integer = ((val1 And &HFF) << 8) Or (val2 And &HFF)
                x = SignExtend12(code12)
                val2 = ReadInt8() ' next byte for Y path
                If val2 = Integer.MinValue Then Exit While
            Else
                x = SignExtend7(val1)
            End If

            If (val2 And FLAG_LONG) <> 0 Then
                If (val2 And TRIM_CODE) <> 0 Then trim = True
                If (val2 And JUMP_CODE) <> 0 Then jump = True
                Dim val3 = ReadInt8()
                If val3 = Integer.MinValue Then Exit While
                Dim code12y As Integer = ((val2 And &HFF) << 8) Or (val3 And &HFF)
                y = SignExtend12(code12y)
            Else
                y = SignExtend7(val2)
            End If

            If jump Then
                pattern.move(x, y)
            ElseIf trim Then
                pattern.trim()
                pattern.move(x, y)
            Else
                pattern.stitch(x, y)
            End If

            ' forward-progress guard
            If reader IsNot Nothing AndAlso reader.BaseStream IsNot Nothing Then
                If reader.BaseStream.Position = lastPos Then
                    Try : Form1.Status("Error: No progress while reading PEC stream.") : Catch : End Try
                    Exit While
                End If
                lastPos = reader.BaseStream.Position
            End If
        End While

        pattern.end()
    End Sub

    ' ---------------------- Color mapping (unchanged behavior, guarded) ----------------------
    Private Sub ProcessPecColors(colorBytes As Byte())
        Dim set64 = EmbThreadPec.GetThreadSet()
        If set64 Is Nothing OrElse set64.Length = 0 OrElse colorBytes Is Nothing Then Exit Sub
        For Each c In colorBytes
            Dim idx = (CInt(c) And &HFF) Mod set64.Length
            pattern.Add(set64(idx))
        Next
    End Sub

    Private Sub ProcessPecTable(colorBytes As Byte())
        Dim set64 = EmbThreadPec.GetThreadSet()
        If set64 Is Nothing OrElse set64.Length = 0 OrElse colorBytes Is Nothing Then Exit Sub

        Dim threadMap(set64.Length - 1) As EmbThread
        Dim queue As New List(Of EmbThread)()
        For Each c In colorBytes
            Dim idx = (CInt(c) And &HFF) Mod set64.Length
            Dim value = threadMap(idx)
            If value Is Nothing Then
                If pattern.ThreadList.Count > 0 Then
                    value = pattern.ThreadList(0)
                    pattern.ThreadList.RemoveAt(0)
                Else
                    value = set64(idx)
                End If
                threadMap(idx) = value
            End If
            queue.Add(value)
        Next
        pattern.ThreadList.Clear()
        pattern.ThreadList.AddRange(queue)
    End Sub

    Private Sub MapPecColors(colorBytes As Byte())
        Dim count = If(colorBytes Is Nothing, 0, colorBytes.Length)
        If pattern.ThreadList.Count = 0 Then
            ProcessPecColors(colorBytes)
            Return
        End If
        If pattern.ThreadList.Count = count Then
            ' 1-to-1; keep existing list (header threads)
            Return
        End If
        ProcessPecTable(colorBytes)
    End Sub

    ' ---------------------- helpers ----------------------
    Private Function SignExtend12(n As Integer) As Integer
        n = n And &HFFF
        If (n And &H800) <> 0 Then n -= &H1000
        Return n
    End Function

    Private Function SignExtend7(n As Integer) As Integer
        n = n And &H7F
        If (n And &H40) <> 0 Then n -= &H80
        Return n
    End Function

    Private Function Clamp(v As Integer, lo As Integer, hi As Integer) As Integer
        If v < lo Then Return lo
        If v > hi Then Return hi
        Return v
    End Function

    Private Function SafeTrim(s As String) As String
        If s Is Nothing Then Return String.Empty
        Return s.Trim()
    End Function
End Class
