Imports System
Imports System.IO

Public Class HusReader
    Inherits EmbReader

    Public Overrides Sub Read()
        ' --- Header (HUS little-endian) ---
        Dim magic_code As Integer = ReadInt32LE()
        Dim number_of_stitches As Integer = ReadInt32LE()
        Dim number_of_colors As Integer = ReadInt32LE()

        Dim extend_pos_x As Integer = Signed16(ReadInt16LE())
        Dim extend_pos_y As Integer = Signed16(ReadInt16LE())
        Dim extend_neg_x As Integer = Signed16(ReadInt16LE())
        Dim extend_neg_y As Integer = Signed16(ReadInt16LE())

        Dim command_offset As Integer = ReadInt32LE()   ' absolute offsets from file start
        Dim x_offset As Integer = ReadInt32LE()
        Dim y_offset As Integer = ReadInt32LE()

        ' 8-byte padding (present in common HUS specs)
        Dim padding(7) As Byte
        ReadFully(padding)

        Dim unknown_16_bit As Integer = ReadInt16LE() ' present in many samples; ignore

        ' --- palette ---
        Dim husSet As EmbThread() = EmbThreadHus.GetThreadSet()
        If husSet Is Nothing OrElse husSet.Length = 0 Then
            husSet = New EmbThread() {New EmbThread(&H0, "HUS Default", "", "Hus", "Hus")}
        End If

        Dim fileLen As Long = If(reader IsNot Nothing AndAlso reader.BaseStream IsNot Nothing, reader.BaseStream.Length, 0L)
        Dim here As Integer = Tell()

        ' Original (working) monotonic / forward check
        Dim offsetsValid As Boolean =
            (command_offset > here AndAlso command_offset <= fileLen) AndAlso
            (x_offset >= command_offset AndAlso x_offset <= fileLen) AndAlso
            (y_offset >= x_offset AndAlso y_offset <= fileLen)

        ' Read only palette bytes that fit before the commands block
        Dim colorsToRead As Integer = 0
        If offsetsValid Then
            Dim bytesUntilCommands As Long = CLng(command_offset) - CLng(here)
            If bytesUntilCommands < 0 Then bytesUntilCommands = 0
            Dim maxColorsByOffset As Integer = CInt(Math.Min(Integer.MaxValue, bytesUntilCommands \ 2))
            colorsToRead = Math.Max(0, Math.Min(number_of_colors, maxColorsByOffset))
        Else
            ' if header looks odd, do not try to guess; keep empty palette rather than misalign the stream
            colorsToRead = 0
        End If

        For i As Integer = 0 To colorsToRead - 1
            Dim raw As Integer = ReadInt16LE()
            Dim idxU As Integer = CInt(CUShort(raw))
            pattern.Add(husSet(idxU Mod husSet.Length))
        Next

        ' Seek to command stream start; if we can't, we must stop cleanly.
        If Not offsetsValid Then
            Try : Form1.Status("Error: HUS offsets invalid; skipping stitches.") : Catch : End Try
            pattern.end() : Return
        End If
        Seek(command_offset)

        ' --- Compressed blocks (original behavior) ---
        ' Use exactly number_of_stitches for each stream (typical HUS expects this)
        Dim safeCount As Integer = Math.Max(0, number_of_stitches)

        ' COMMANDS
        Dim cmdLen As Integer = Math.Max(0, x_offset - command_offset)
        Dim command_bytes As Byte() = reader.ReadBytes(cmdLen)
        ReadPosition += command_bytes.Length
        Dim command_decompressed As Byte() = EmbCompress.Expand(command_bytes, safeCount)

        ' X deltas
        Seek(x_offset)
        Dim xLen As Integer = Math.Max(0, y_offset - x_offset)
        Dim x_bytes As Byte() = reader.ReadBytes(xLen)
        ReadPosition += x_bytes.Length
        Dim x_decompressed As Byte() = EmbCompress.Expand(x_bytes, safeCount)

        ' Y deltas (to EOF)
        Seek(y_offset)
        Dim y_bytes As Byte()
        If fileLen > 0 Then
            Dim remaining As Long = fileLen - reader.BaseStream.Position
            If remaining < 0 Then remaining = 0
            y_bytes = reader.ReadBytes(CInt(Math.Min(Integer.MaxValue, remaining)))
            ReadPosition += y_bytes.Length
        Else
            y_bytes = reader.ReadBytes(Integer.MaxValue) ' defensive fallback
        End If
        Dim y_decompressed As Byte() = EmbCompress.Expand(y_bytes, safeCount)

        ' --- Stitch decode (byte-per-stitch triplets) ---
        Dim stitch_count As Integer = Math.Min(Math.Min(command_decompressed.Length, x_decompressed.Length), y_decompressed.Length)
        For i As Integer = 0 To stitch_count - 1
            Dim cmd As Integer = command_decompressed(i) And &HFF
            Dim dx As Integer = Sign8(x_decompressed(i) And &HFF)
            Dim dy As Integer = -Sign8(y_decompressed(i) And &HFF)
            Select Case cmd
                Case &H80 : pattern.stitch(dx, dy)                ' stitch
                Case &H81 : pattern.move(dx, dy)                  ' jump/move
                Case &H84 : pattern.color_change(dx, dy)          ' color change
                Case &H88                                        ' trim (+ optional move)
                    If dx <> 0 OrElse dy <> 0 Then pattern.move(dx, dy)
                    pattern.trim()
                Case &H90 : Exit For                              ' end
                Case Else : Exit For                              ' unknown => stop
            End Select
        Next

        pattern.end()
    End Sub

    ' --- helpers ---
    Private Function Sign8(n As Integer) As Integer
        n = n And &HFF
        If (n And &H80) <> 0 Then n -= &H100
        Return n
    End Function

    Private Function Signed16(v As Integer) As Integer
        v = v And &HFFFF
        If (v And &H8000) <> 0 Then v -= &H10000
        Return v
    End Function
End Class
