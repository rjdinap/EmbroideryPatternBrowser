Imports System
Imports System.IO

Public Class VipReader
    Inherits EmbReader

    Public Overrides Sub Read()
        ' ---- VIP header (little-endian) ----
        Dim magic As Integer = ReadInt32LE()            ' expected 0x0190FC5D (not enforced)
        Dim number_of_stitches As Integer = ReadInt32LE()
        Dim number_of_colors As Integer = ReadInt32LE()

        Dim extend_pos_x As Integer = Signed16(ReadInt16LE())
        Dim extend_pos_y As Integer = Signed16(ReadInt16LE())
        Dim extend_neg_x As Integer = Signed16(ReadInt16LE())
        Dim extend_neg_y As Integer = Signed16(ReadInt16LE())

        Dim command_offset As Integer = ReadInt32LE()   ' absolute offsets from file start
        Dim x_offset As Integer = ReadInt32LE()
        Dim y_offset As Integer = ReadInt32LE()

        ' VIP: 8-byte name, then an int16 (unknown), then colorLength (= number_of_colors * 4)
        Dim nameBytes(7) As Byte
        ReadFully(nameBytes)
        Dim unknown_16 As Integer = ReadInt16LE()
        Dim colorLength As Integer = ReadInt32LE()

        ' ---- palette (VIP stores real 24-bit colors, encoded) ----
        Dim here As Integer = Tell()
        Dim fileLen As Long = If(reader IsNot Nothing AndAlso reader.BaseStream IsNot Nothing, reader.BaseStream.Length, 0L)

        ' Make sure offsets look sane and forward-monotonic before reading chunks between them.
        Dim offsetsValid As Boolean =
            (command_offset > here AndAlso command_offset <= fileLen) AndAlso
            (x_offset >= command_offset AndAlso x_offset <= fileLen) AndAlso
            (y_offset >= x_offset AndAlso y_offset <= fileLen)

        ' VIP color table is exactly number_of_colors * 4 bytes, right here after the header fields.
        ' Each color entry is (R,G, B, pad). Bytes are XOR-decoded.
        Dim colors_to_take As Integer = Math.Max(0, number_of_colors)
        Dim encodedColors As Byte() = Array.Empty(Of Byte)()
        If colors_to_take > 0 Then
            ' Read exactly the colorLength bytes if available; otherwise clamp by remaining before command_offset.
            Dim want As Integer = Math.Max(0, Math.Min(colorLength, Integer.MaxValue))
            If offsetsValid Then
                Dim maxAvail As Integer = Math.Max(0, Math.Min(want, CInt(Math.Max(0L, CLng(command_offset) - CLng(Tell())))))
                encodedColors = reader.ReadBytes(maxAvail)
                ReadPosition += encodedColors.Length
            Else
                ' If offsets look bad, still try to read colorLength but clamp to file end.
                Dim remaining As Integer = CInt(Math.Max(0L, fileLen - Tell()))
                Dim toRead As Integer = Math.Min(Math.Max(0, want), remaining)
                encodedColors = reader.ReadBytes(toRead)
                ReadPosition += encodedColors.Length
            End If
        End If

        ' Decode VIP palette: out[i] = (in[i] XOR table[i]) XOR previousInputByte
        If encodedColors.Length >= colors_to_take * 4 Then
            Dim decoded(encodedColors.Length - 1) As Byte
            Dim prevIn As Byte = 0
            For i As Integer = 0 To encodedColors.Length - 1
                Dim inputByte As Byte = encodedColors(i)
                Dim mask As Byte = vipTable(i Mod vipTable.Length)
                decoded(i) = CByte((inputByte Xor mask) Xor prevIn)
                prevIn = inputByte
            Next
            For c As Integer = 0 To colors_to_take - 1
                Dim baseIdx = c * 4
                If baseIdx + 2 < decoded.Length Then
                    Dim r As Integer = decoded(baseIdx + 0)
                    Dim g As Integer = decoded(baseIdx + 1)
                    Dim b As Integer = decoded(baseIdx + 2)
                    pattern.Add(New EmbThread((r << 16) Or (g << 8) Or b, $"VIP {c + 1}", (c + 1).ToString(), "VIP", "VIP"))
                End If
            Next
        End If

        ' ---- seek to command stream start (if invalid offsets, stop cleanly like HusReader) ----
        If Not offsetsValid Then
            Try : Form1.Status("Error: VIP offsets invalid; skipping stitches.") : Catch : End Try
            pattern.end() : Return
        End If
        Seek(command_offset)

        ' ---- compressed blocks, mirroring HusReader style ----
        Dim safeCount As Integer = Math.Max(0, number_of_stitches)

        ' COMMANDS [command_offset .. x_offset)
        Dim cmdLen As Integer = Math.Max(0, x_offset - command_offset)
        Dim command_bytes As Byte() = reader.ReadBytes(cmdLen)
        ReadPosition += command_bytes.Length
        Dim command_decompressed As Byte() = EmbCompress.Expand(command_bytes, safeCount)

        ' X deltas [x_offset .. y_offset)
        Seek(x_offset)
        Dim xLen As Integer = Math.Max(0, y_offset - x_offset)
        Dim x_bytes As Byte() = reader.ReadBytes(xLen)
        ReadPosition += x_bytes.Length
        Dim x_decompressed As Byte() = EmbCompress.Expand(x_bytes, safeCount)

        ' Y deltas [y_offset .. EOF]
        Seek(y_offset)
        Dim y_bytes As Byte()
        If fileLen > 0 Then
            Dim remaining As Long = fileLen - reader.BaseStream.Position
            If remaining < 0 Then remaining = 0
            y_bytes = reader.ReadBytes(CInt(Math.Min(Integer.MaxValue, remaining)))
            ReadPosition += y_bytes.Length
        Else
            y_bytes = reader.ReadBytes(Integer.MaxValue)
        End If
        Dim y_decompressed As Byte() = EmbCompress.Expand(y_bytes, safeCount)

        ' ---- stitch decode (byte-per-stitch triplets), modeled exactly like HusReader ----
        Dim stitch_count As Integer = Math.Min(Math.Min(command_decompressed.Length, x_decompressed.Length), y_decompressed.Length)
        For i As Integer = 0 To stitch_count - 1
            Dim cmd As Integer = command_decompressed(i) And &HFF
            Dim dx As Integer = Sign8(x_decompressed(i) And &HFF)
            Dim dy As Integer = -Sign8(y_decompressed(i) And &HFF)

            Select Case cmd
                Case &H80 : pattern.stitch(dx, dy)                ' stitch / normal
                Case &H81 : pattern.move(dx, dy)                  ' jump / move
                Case &H84 : pattern.color_change(dx, dy)          ' color change
                Case &H88                                        ' trim (+ optional move)
                    If dx <> 0 OrElse dy <> 0 Then pattern.move(dx, dy)
                    pattern.trim()
                Case &H90 : Exit For                              ' end
                Case Else : Exit For                              ' unknown => stop (consistent with HusReader)
            End Select
        Next

        pattern.end()
    End Sub

    ' --- helpers (match HusReader style) ---
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

    ' 512-byte XOR mask table for VIP palette decoding
    Private Shared ReadOnly vipTable As Byte() = {
        &H2E, &H82, &HE4, &H6F, &H38, &HA9, &HDC, &HC6, &H7B, &HB6, &H28, &HAC, &HFD, &HAA, &H8A, &H4E,
        &H76, &H2E, &HF0, &HE4, &H25, &H1B, &H8A, &H68, &H4E, &H92, &HB9, &HB4, &H95, &HF0, &H3E, &HEF,
        &HF7, &H40, &H24, &H18, &H39, &H31, &HBB, &HE1, &H53, &HA8, &H1F, &HB1, &H3A, &H7, &HFB, &HCB,
        &HE6, &H0, &H81, &H50, &HE, &H40, &HE1, &H2C, &H73, &H50, &HD, &H91, &HD6, &HA, &H5D, &HD6,
        &H8B, &HB8, &H62, &HAE, &H47, &H0, &H53, &H5A, &HB7, &H80, &HAA, &H28, &HF7, &H5D, &H70, &H5E,
        &H2C, &HB, &H98, &HE3, &HA0, &H98, &H60, &H47, &H89, &H9B, &H82, &HFB, &H40, &HC9, &HB4, &H0,
        &HE, &H68, &H6A, &H1E, &H9, &H85, &HC0, &H53, &H81, &HD1, &H98, &H89, &HAF, &HE8, &H85, &H4F,
        &HE3, &H69, &H89, &H3, &HA1, &H2E, &H8F, &HCF, &HED, &H91, &H9F, &H58, &H1E, &HD6, &H84, &H3C,
        &H9, &H27, &HBD, &HF4, &HC3, &H90, &HC0, &H51, &H1B, &H2B, &H63, &HBC, &HB9, &H3D, &H40, &H4D,
        &H62, &H6F, &HE0, &H8C, &HF5, &H5D, &H8, &HFD, &H3D, &H50, &H36, &HD7, &HC9, &HC9, &H43, &HE4,
        &H2D, &HCB, &H95, &HB6, &HF4, &HD, &HEA, &HC2, &HFD, &H66, &H3F, &H5E, &HBD, &H69, &H6, &H2A,
        &H3, &H19, &H47, &H2B, &HDF, &H38, &HEA, &H4F, &H80, &H49, &H95, &HB2, &HD6, &HF9, &H9A, &H75,
        &HF4, &HD8, &H9B, &H1D, &HB0, &HA4, &H69, &HDB, &HA9, &H21, &H79, &H6F, &HD8, &HDE, &H33, &HFE,
        &H9F, &H4, &HE5, &H9A, &H6B, &H9B, &H73, &H83, &H62, &H7C, &HB9, &H66, &H76, &HF2, &H5B, &HC9,
        &H5E, &HFC, &H74, &HAA, &H6C, &HF1, &HCD, &H93, &HCE, &HE9, &H80, &H53, &H3, &H3B, &H97, &H4B,
        &H39, &H76, &HC2, &HC1, &H56, &HCB, &H70, &HFD, &H3B, &H3E, &H52, &H57, &H81, &H5D, &H56, &H8D,
        &H51, &H90, &HD4, &H76, &HD7, &HD5, &H16, &H2, &H6D, &HF2, &H4D, &HE1, &HE, &H96, &H4F, &HA1,
        &H3A, &HA0, &H60, &H59, &H64, &H4, &H1A, &HE4, &H67, &HB6, &HED, &H3F, &H74, &H20, &H55, &H1F,
        &HFB, &H23, &H92, &H91, &H53, &HC8, &H65, &HAB, &H9D, &H51, &HD6, &H73, &HDE, &H1, &HB1, &H80,
        &HB7, &HC0, &HD6, &H80, &H1C, &H2E, &H3C, &H83, &H63, &HEE, &HBC, &H33, &H25, &HE2, &HE, &H7A,
        &H67, &HDE, &H3F, &H71, &H14, &H49, &H9C, &H92, &H93, &HD, &H26, &H9A, &HE, &HDA, &HED, &H6F,
        &HA4, &H89, &HC, &H1B, &HF0, &HA1, &HDF, &HE1, &H9E, &H3C, &H4, &H78, &HE4, &HAB, &H6D, &HFF,
        &H9C, &HAF, &HCA, &HC7, &H88, &H17, &H9C, &HE5, &HB7, &H33, &H6D, &HDC, &HED, &H8F, &H6C, &H18,
        &H1D, &H71, &H6, &HB1, &HC5, &HE2, &HCF, &H13, &H77, &H81, &HC5, &HB7, &HA, &H14, &HA, &H6B,
        &H40, &H26, &HA0, &H88, &HD1, &H62, &H6A, &HB3, &H50, &H12, &HB9, &H9B, &HB5, &H83, &H9B, &H37
    }
End Class
