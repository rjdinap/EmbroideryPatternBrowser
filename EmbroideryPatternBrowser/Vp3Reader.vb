Imports System.Text

Public Class Vp3Reader
    Inherits EmbReader

    ' ---- behavior toggles ----
    Private Const flipX As Boolean = False
    Private Const flipY As Boolean = False
    Private Const VP3_USE_FIRST_SWATCH As Boolean = True
    Private Const VP3_COLOR_ORDER_BGR As Boolean = False ' your files are RGB

    ' ---- conservative caps / guards (heuristic-friendly) ----
    Private Const MAX_COLORS As Integer = 4096
    Private Const MAX_BLOCK_SIZE As Integer = 16 * 1024 * 1024     ' per color block (16 MB)
    Private Const MAX_STITCH_ITERS As Integer = 50 * 1024 * 1024   ' hard loop limiter
    Private Const MAX_STRING_LEN As Integer = 2048                  ' cap brand/catalog/desc

    Public Overrides Sub Read()
        Try
            ' ---- magic ----
            Dim magic(5) As Byte
            SafeReadFully(magic)

            ' ---- strings + padding with safe skips ----
            SafeSkipVp3String()
            SafeSkipBytes(7)
            SafeSkipVp3String()
            SafeSkipBytes(32)

            ' ---- center (BE int32, 1/100 mm). Java negates center_y. ----
            Dim cxRaw As Integer = SafeReadInt32BE(0)
            Dim cyRaw As Integer = SafeReadInt32BE(0)

            ' sanity clamp (avoid crazy values if previous skips misaligned)
            If cxRaw < -1000000 OrElse cxRaw > 1000000 Then cxRaw = 0
            If cyRaw < -1000000 OrElse cyRaw > 1000000 Then cyRaw = 0

            Dim center_x As Single = CSng(cxRaw / 100.0F)
            Dim center_y As Single = CSng(-cyRaw / 100.0F)

            SafeSkipBytes(27)
            SafeSkipVp3String()
            SafeSkipBytes(24)
            SafeSkipVp3String()

            ' ---- number of colors ----
            Dim count_colors As Integer = SafeReadInt16BE(0)
            If count_colors < 0 OrElse count_colors > MAX_COLORS Then count_colors = 0

            For i As Integer = 0 To count_colors - 1
                If Not Vp3ReadColorBlock(center_x, center_y) Then Exit For
                ' no color_change here; each block picks its own thread
            Next

            pattern.end()

        Catch ex As Exception
            Try : Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString) : Catch : End Try
            Try : pattern.end() : Catch : End Try
        End Try
    End Sub

    ''' <summary>
    ''' Reads one VP3 color block. Returns False if the block cannot be processed safely.
    ''' </summary>
    Private Function Vp3ReadColorBlock(center_x As Single, center_y As Single) As Boolean
        ' Small header (commonly 00 05 00)
        Dim hdr(2) As Byte
        SafeReadFully(hdr)

        ' Distance to next block (relative, big-endian)
        Dim distance_to_next_block As Integer = SafeReadInt32BE(-1)
        If distance_to_next_block < 0 OrElse distance_to_next_block > MAX_BLOCK_SIZE Then
            Form1.Status("Error: VP3 block size invalid.")
            Return False
        End If

        ' Compute absolute end of this block; clamp to file length
        Dim here As Long = reader.BaseStream.Position
        Dim block_end As Long = here + distance_to_next_block
        Dim fileLen As Long = If(reader IsNot Nothing AndAlso reader.BaseStream IsNot Nothing, reader.BaseStream.Length, 0L)
        If block_end < 0 Then block_end = 0
        If block_end > fileLen Then block_end = fileLen

        ' start position (BE int32, 1/100 mm) – Java negates start_y
        Dim start_x As Single = CSng(SafeReadInt32BE(0) / 100.0F)
        Dim start_y As Single = CSng(-SafeReadInt32BE(0) / 100.0F)

        Dim abs_x As Single = start_x + center_x
        Dim abs_y As Single = start_y + center_y
        If flipX Then abs_x = -abs_x
        If flipY Then abs_y = -abs_y

        ' Thread for this block
        Dim thread As EmbThread = Vp3ReadThread()
        If thread Is Nothing Then
            thread = New EmbThread(&H0, "VP3", "", "VP3", "VP3")
        End If
        pattern.Add(thread)
        pattern.select_thread(Math.Max(0, pattern.ThreadList.Count - 1))

        ' Per Java, only move if BOTH are non-zero
        pattern.trim()
        If (abs_x <> 0) AndAlso (abs_y <> 0) Then
            pattern.moveAbs(abs_x, abs_y)
        End If

        ' skip marker chunk (0A F6 00 after 15 pad) – validate minimally but keep tolerant
        SafeSkipBytes(15)
        SafeReadFully(hdr) ' expect 0A F6 00; ignore content

        ' read stitch block until next block boundary
        Dim cur As Long = reader.BaseStream.Position
        Dim stitch_len_long As Long = Math.Max(0, block_end - cur)
        If stitch_len_long <= 0 Then Return True

        Dim stitch_len As Integer = CInt(Math.Min(Integer.MaxValue, stitch_len_long))
        If stitch_len > MAX_BLOCK_SIZE Then
            Form1.Status("Error: VP3 stitch block too large.")
            ' Seek to block end and continue
            Seek(CInt(block_end))
            Return True
        End If

        Dim stitch_bytes As Byte() = SafeReadBytes(stitch_len)
        If stitch_bytes.Length <> stitch_len Then
            ' Truncated; continue safely
            Return True
        End If

        Dim trimmed As Boolean = False
        Dim i As Integer = 0
        Dim iters As Integer = 0

        While i < stitch_len - 1
            If iters >= MAX_STITCH_ITERS Then
                Form1.Status("Error: VP3 decode iteration limit reached.")
                Exit While
            End If
            iters += 1

            Dim xb As Integer = stitch_bytes(i)
            Dim yb As Integer = stitch_bytes(i + 1)
            i += 2

            If (xb And &HFF) <> &H80 Then
                Dim dx As Integer = Sign8(xb)
                Dim dy As Integer = Sign8(yb) ' deltas not negated (we did for center/start)
                If flipX Then dx = -dx
                If flipY Then dy = -dy
                pattern.stitch(dx, dy)
                trimmed = False
                Continue While
            End If

            ' Control paths
            If yb = &H1 Then
                If i + 3 >= stitch_bytes.Length Then Exit While
                Dim lx As Integer = Signed16BE(stitch_bytes(i), stitch_bytes(i + 1)) : i += 2
                Dim ly As Integer = Signed16BE(stitch_bytes(i), stitch_bytes(i + 1)) : i += 2
                Dim dx As Integer = If(flipX, -lx, lx)
                Dim dy As Integer = If(flipY, -ly, ly)
                If (Math.Abs(dx) > 255) OrElse (Math.Abs(dy) > 255) Then
                    If Not trimmed Then pattern.trim()
                    pattern.move(dx, dy)
                    trimmed = True
                Else
                    pattern.stitch(dx, dy)
                    trimmed = False
                End If

            ElseIf yb = &H2 Then
                ' LONG-END (no-op)

            ElseIf yb = &H3 Then
                If Not trimmed Then pattern.trim()
                trimmed = True
            End If
        End While

        ' Ensure reader is at block_end (in case of early exit)
        Dim after As Long = reader.BaseStream.Position
        Dim needSeek As Long = block_end - after
        If needSeek > 0 AndAlso needSeek <= Integer.MaxValue Then
            Seek(CInt(block_end))
        End If

        Return True
    End Function

    ' ----- Color swatch reading (RGB for your files) -----
    Private Function Vp3ReadThread() As EmbThread
        Dim pickedRgb As Integer = 0
        Dim colors As Integer = ReadInt8()
        Dim transition As Integer = ReadInt8() ' unused but read to keep alignment

        If colors <= 0 Then
            pickedRgb = &H0
        Else
            If VP3_USE_FIRST_SWATCH Then
                pickedRgb = ReadVp3ColorRGB()
                SkipVp3ColorMeta()
                For m As Integer = 1 To colors - 1
                    Dim dummy As Integer = ReadVp3ColorRGB()
                    SkipVp3ColorMeta()
                Next
            Else
                For m As Integer = 0 To colors - 1
                    pickedRgb = ReadVp3ColorRGB()
                    SkipVp3ColorMeta()
                Next
            End If
        End If

        Dim thread_type As Integer = ReadInt8() ' unused
        Dim weight As Integer = ReadInt8()      ' unused

        Dim catalog As String = Trunc(ReadVp3StringUtf8(), MAX_STRING_LEN)
        Dim desc As String = Trunc(ReadVp3StringUtf8(), MAX_STRING_LEN)
        Dim brand As String = Trunc(ReadVp3StringUtf8(), MAX_STRING_LEN)
        If String.IsNullOrEmpty(brand) Then brand = "VP3"

        Return New EmbThread(pickedRgb, desc, catalog, brand, "VP3")
    End Function

    Private Function ReadVp3ColorRGB() As Integer
        Dim b0 As Integer = ReadInt8() And &HFF
        Dim b1 As Integer = ReadInt8() And &HFF
        Dim b2 As Integer = ReadInt8() And &HFF
        If VP3_COLOR_ORDER_BGR Then
            Dim r As Integer = b2, g As Integer = b1, b As Integer = b0
            Return (r << 16) Or (g << 8) Or b
        Else
            Dim r As Integer = b0, g As Integer = b1, b As Integer = b2
            Return (r << 16) Or (g << 8) Or b
        End If
    End Function

    Private Sub SkipVp3ColorMeta()
        ' Keep structure alignment; values are not used
        Dim parts As Integer = ReadInt8()
        Dim color_length As Integer = ReadInt16BE()
        ' Intentionally ignored; content (like gradient stops) not needed for thumbnail
    End Sub

    ' ----- string helpers -----
    Private Function ReadVp3StringUtf8() As String
        Dim len As Integer = SafeReadInt16BE(0)
        If len <= 0 Then Return String.Empty
        ' cap read length to avoid large allocations
        Dim toRead As Integer = Math.Min(len, MAX_STRING_LEN)
        Dim bytes = SafeReadBytes(toRead)
        ' if original length was larger, skip the remainder safely
        If len > toRead Then SafeSkipBytes(len - toRead)
        Return Encoding.UTF8.GetString(bytes)
    End Function

    Private Sub SafeSkipVp3String()
        Dim len As Integer = SafeReadInt16BE(0)
        If len > 0 Then SafeSkipBytes(len)
    End Sub

    ' ----- safe IO helpers -----
    Private Sub SafeReadFully(buf As Byte())
        If buf Is Nothing OrElse buf.Length = 0 Then Return
        Dim remaining As Integer = Math.Max(0, CInt(ReaderRemaining()))
        Dim toRead As Integer = Math.Min(remaining, buf.Length)
        If toRead <= 0 Then
            Array.Clear(buf, 0, buf.Length)
            Return
        End If
        reader.Read(buf, 0, toRead)
        ReadPosition += toRead
        If toRead < buf.Length Then
            ' zero-fill the rest
            For i = toRead To buf.Length - 1
                buf(i) = 0
            Next
        End If
    End Sub

    Private Function SafeReadBytes(count As Integer) As Byte()
        Dim avail As Integer = Math.Max(0, CInt(ReaderRemaining()))
        Dim n As Integer = Math.Max(0, Math.Min(count, avail))
        Dim b(n - 1) As Byte
        If n > 0 Then
            reader.Read(b, 0, n)
            ReadPosition += n
        End If
        Return b
    End Function

    Private Sub SafeSkipBytes(count As Integer)
        Dim avail As Integer = Math.Max(0, CInt(ReaderRemaining()))
        Dim n As Integer = Math.Max(0, Math.Min(count, avail))
        If n > 0 Then
            Skip(n)
        End If
    End Sub

    Private Function SafeReadInt16BE(defaultValue As Integer) As Integer
        If ReaderRemaining() < 2 Then Return defaultValue
        Return ReadInt16BE()
    End Function

    Private Function SafeReadInt32BE(defaultValue As Integer) As Integer
        If ReaderRemaining() < 4 Then Return defaultValue
        Return ReadInt32BE()
    End Function

    Private Function ReaderRemaining() As Long
        If reader Is Nothing OrElse reader.BaseStream Is Nothing Then Return 0
        Return Math.Max(0L, reader.BaseStream.Length - reader.BaseStream.Position)
    End Function

    ' ----- math helpers -----
    Private Function Signed16BE(hi As Integer, lo As Integer) As Integer
        Dim v As Integer = ((hi And &HFF) << 8) Or (lo And &HFF)
        If (v And &H8000) <> 0 Then v -= &H10000
        Return v
    End Function

    Private Function Sign8(n As Integer) As Integer
        n = n And &HFF
        If (n And &H80) <> 0 Then n -= &H100
        Return n
    End Function

    Private Function Trunc(s As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(s) Then Return String.Empty
        If maxLen <= 0 Then Return String.Empty
        If s.Length <= maxLen Then Return s
        Return s.Substring(0, maxLen)
    End Function
End Class
