
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
    Private _fileName As String = ""
    Public Overrides Sub Read(fn As String)
        _fileName = fn
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
            Try : Logger.Error("PecReader: error in file: " & _fileName & " - " & ex.Message, ex.StackTrace.ToString) : Catch : End Try
            Try : pattern.end() : Catch : End Try
        End Try
    End Sub



    ' ---------------------- Core PEC payload decoder (alignment + 463-color table) ----------------------
    Private Sub ReadPecCore()
        ' --- Label / basic header ---
        Skip(3) ' "LA:"
        Dim label As String = SafeTrim(ReadString(16))
        If label.Length > 0 Then pattern.SetMetadata(EmbPattern.PROP_NAME, label)
        Skip(&HF)

        ' Dimensions and icon info (guarded)
        Dim stride As Integer = ReadInt8()
        Dim iconH As Integer = ReadInt8()
        Skip(&HC)

        ' Color-change count byte (not used to size the table)
        Dim colorChanges As Integer = ReadInt8()
        If colorChanges = Integer.MinValue Then
            pattern.end() : Return
        End If

        ' ---- PEC color table: fixed 463 bytes ----
        ' --- fixed: still read the whole 463-byte table to keep stream aligned ---
        Dim colorListLen As Integer = 463
        Dim colorBytes(colorListLen - 1) As Byte
        Dim gotList As Integer = ReadFully(colorBytes)
        If gotList <> colorBytes.Length Then
            pattern.end() : Return
        End If

        ' Use only the first (colorChanges + 1) entries, like the Java reader
        Dim usedCount As Integer = Math.Max(0, Math.Min(colorListLen, colorChanges + 1))
        If usedCount <= 0 Then
            ' No color changes recorded: leave threadlist as-is
        Else
            Dim used(usedCount - 1) As Byte
            Array.Copy(colorBytes, 0, used, 0, usedCount)
            MapPecColors(used)
        End If


        ' Two zeros, then 4 bytes, then FF F0, then 12 bytes
        Dim z1 As Integer = ReadInt8()
        Dim z2 As Integer = ReadInt8()
        'Try : Form1.StatusFromAnyThread($"[PEC] two zeros after color list = {z1:X2} {z2:X2}") : Catch : End Try
        Skip(4)
        Dim m1 As Integer = ReadInt8()
        Dim m2 As Integer = ReadInt8()
        'ry : Form1.StatusFromAnyThread($"[PEC] pre-stitch marker bytes = {m1:X2} {m2:X2}") : Catch : End Try
        Skip(&HC) ' 12 bytes

        ' --- Stitches ---
        ReadPecStitchesSafe()

        ' --- Small icon graphics (guard sizes) ---
        Dim wStride As Integer = Clamp(stride, 0, MAX_STRIDE)
        Dim hIcon As Integer = Clamp(iconH, 0, MAX_ICON_H)
        Dim byteSize As Long = CLng(wStride) * CLng(hIcon)
        If byteSize > 0 AndAlso byteSize <= MAX_GRAPHIC_BYTES Then
            ReadPecGraphics(CInt(byteSize), wStride, colorChanges)
        End If
    End Sub




    ' Build the EmbPattern.ThreadList from the first N entries of the 463-byte PEC color table.
    ' Each entry is an index into the 64-color PEC palette.
    Private Sub ApplyPecThreads(colorBytes As Byte(), usedCount As Integer)
        Dim set64 = EmbThreadPec.GetThreadSet()
        pattern.ThreadList.Clear()

        If set64 Is Nothing OrElse set64.Length = 0 Then Exit Sub
        If colorBytes Is Nothing Then Exit Sub
        If usedCount <= 0 Then Exit Sub

        Dim n As Integer = Math.Min(usedCount, colorBytes.Length)
        For i As Integer = 0 To n - 1
            Dim idx = (CInt(colorBytes(i)) And &HFF) Mod set64.Length
            pattern.ThreadList.Add(set64(idx))
        Next
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



    ' ============================== PEC stitch decode + verbose logging ==============================
    Private Sub ReadPecStitchesSafe()
        ' - End ONLY on FF 00
        ' - FE B0 <xx> = color change (consume the 3rd byte)
        ' - Decode deltas (7-bit or 12-bit 2’s-complement)
        ' - Classify using extra bits (0x70 mask on long-form bytes):
        '     * x==0 && y==0 → BeginAnchor
        '     * extraX==0x20 && extraY==0x20 → MovementOnly (jump, pen up)
        '     * extraX==0x10 && extraY==0x10 → EndAnchor (pen up)
        '     * also treat the FIRST pair after a color change, and the pair right after a MoveOnly, as EndAnchor
        ' - Emit pattern.move for anchors/moves, pattern.stitch for normal stitches
        ' - No rotate/flip here; renderer will just translate to (0,0) using bounds elsewhere


        Dim iter As Integer = 0
        Dim lastPos As Long = If(reader Is Nothing OrElse reader.BaseStream Is Nothing, 0L, reader.BaseStream.Position)

        ' Logging + state
        Dim logCount As Integer = 0
        Dim lastWasMoveOnly As Boolean = False
        Dim firstAfterColor As Boolean = True    ' C# behavior: first decoded pair after color is treated as EndAnchor

        ' Stats + absolute cursor (for bbox sanity)
        Dim nStitch As Integer = 0, nMove As Integer = 0, nBegin As Integer = 0, nEnd As Integer = 0, nJump As Integer = 0, nTrim As Integer = 0, nColor As Integer = 0
        Dim ax As Integer = 0, ay As Integer = 0
        Dim minX As Integer = 0, minY As Integer = 0, maxX As Integer = 0, maxY As Integer = 0


        While True
            If iter >= MAX_STITCH_ITERS Then
                'Form1.StatusFromAnyThread("ERROR: iteration limit hit")
                Exit While
            End If
            iter += 1

            ' Read two bytes
            Dim b1 = ReadInt8() : If b1 = Integer.MinValue Then Exit While
            Dim b2 = ReadInt8() : If b2 = Integer.MinValue Then Exit While

            ' ---- Proper end condition: FF 00 ----
            If b1 = &HFF AndAlso b2 = &H0 Then
                'Form1.StatusFromAnyThread("END FF 00")
                Exit While
            End If

            ' ---- Color change: FE B0 <byte> ----
            If b1 = &HFE AndAlso b2 = &HB0 Then
                Dim unk = ReadInt8() : If unk = Integer.MinValue Then Exit While
                pattern.color_change() : nColor += 1
                firstAfterColor = True
                lastWasMoveOnly = False
                'Form1.StatusFromAnyThread($"COLOR-CHANGE (byte={unk:X2})")
                Continue While
            End If

            ' ---- Decode X (7-bit or 12-bit) ----
            Dim extraX As Integer = 0
            Dim dx As Integer
            Dim nextForY As Integer

            If (b1 And &H80) <> 0 Then
                extraX = (b1 And &H70)
                Dim code12 As Integer = ((b1 And &HFF) << 8) Or (b2 And &HFF)
                dx = SignExtend12(code12)      ' low 12 bits as signed
                nextForY = ReadInt8() : If nextForY = Integer.MinValue Then Exit While
            Else
                dx = SignExtend7(b1)
                nextForY = b2
            End If

            ' ---- Decode Y (7-bit or 12-bit) ----
            Dim extraY As Integer = 0
            Dim dy As Integer
            If (nextForY And &H80) <> 0 Then
                extraY = (nextForY And &H70)
                Dim yLow = ReadInt8() : If yLow = Integer.MinValue Then Exit While
                Dim code12y As Integer = ((nextForY And &HFF) << 8) Or (yLow And &HFF)
                dy = SignExtend12(code12y)
            Else
                dy = SignExtend7(nextForY)
            End If

            ' ---- Classify per C# rules (anchors/moves) ----
            Dim isBeginAnchor As Boolean = (dx = 0 AndAlso dy = 0)
            Dim isMoveOnly As Boolean = (extraX = &H20 AndAlso extraY = &H20)
            Dim isEndBits As Boolean = (extraX = &H10 AndAlso extraY = &H10)
            Dim doEndAnchor As Boolean = isEndBits OrElse lastWasMoveOnly OrElse firstAfterColor

            ' ---- Emit with pen-up/pen-down semantics ----
            If isBeginAnchor Then
                '  no geometry change needed except ensuring pen-up
                'Form1.StatusFromAnyThread("BEGIN-ANCHOR")
                ' ensure pen-up without changing position
                pattern.move(0, 0)
                nBegin += 1
                lastWasMoveOnly = False
                firstAfterColor = False

            ElseIf isMoveOnly Then
                'Form1.StatusFromAnyThread(String.Format("JUMP/MOVE dx={0,4} dy={1,4}", dx, dy))
                pattern.move(dx, dy)
                nMove += 1 : nJump += 1
                lastWasMoveOnly = True
                firstAfterColor = False

            ElseIf doEndAnchor Then
                'Form1.StatusFromAnyThread(String.Format("END-ANCHOR dx={0,4} dy={1,4}", dx, dy))
                pattern.move(dx, dy)
                nEnd += 1 : nMove += 1
                lastWasMoveOnly = False
                firstAfterColor = False

            Else
                'Form1.StatusFromAnyThread(String.Format("STITCH     dx={0,4} dy={1,4}", dx, dy))
                pattern.stitch(dx, dy)
                nStitch += 1
                lastWasMoveOnly = False
                firstAfterColor = False
            End If

            ' ---- Absolute cursor for bbox sanity ----
            ax += dx : ay += dy
            If ax < minX Then minX = ax
            If ax > maxX Then maxX = ax
            If ay < minY Then minY = ay
            If ay > maxY Then maxY = ay

            ' ---- forward progress guard ----
            If reader IsNot Nothing AndAlso reader.BaseStream IsNot Nothing Then
                If reader.BaseStream.Position = lastPos Then
                    'Form1.StatusFromAnyThread("ERROR: no progress in stitch loop")
                    Exit While
                End If
                lastPos = reader.BaseStream.Position
            End If
        End While

        'Form1.StatusFromAnyThread($"TOTAL: stitch={nStitch} move={nMove} begin={nBegin} end={nEnd} jump={nJump} trim={nTrim} colors={nColor}")
        'Form1.StatusFromAnyThread($"BBox abs => min=({minX},{minY}) max=({maxX},{maxY}) size=({maxX - minX} x {maxY - minY})")

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