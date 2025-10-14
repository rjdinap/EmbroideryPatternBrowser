Public Class PesReader
    Inherits PecReader   ' If you renamed: Inherits PecReaderSafe

    Private Const PROP_VERSION As String = "version"

    Public Overrides Sub Read()
        ReadPESHeader()
        ' After header handlers call Seek(pecStart), delegate to base PEC reader:
        MyBase.Read()
    End Sub

    Private Sub ReadPESHeader()
        Dim signature As String = ReadString(8)

        Select Case signature
            Case "#PES0100" : pattern.SetMetadata(PROP_VERSION, "10") : ReadPESHeaderV10()
            Case "#PES0090" : pattern.SetMetadata(PROP_VERSION, "9") : ReadPESHeaderV9()
            Case "#PES0080" : pattern.SetMetadata(PROP_VERSION, "8") : ReadPESHeaderV8()
            Case "#PES0070" : pattern.SetMetadata(PROP_VERSION, "7") : ReadPESHeaderV7()
            Case "#PES0060" : pattern.SetMetadata(PROP_VERSION, "6") : ReadPESHeaderV6()
            Case "#PES0056" : pattern.SetMetadata(PROP_VERSION, "5.6") : ReadPESHeaderV5()
            Case "#PES0055" : pattern.SetMetadata(PROP_VERSION, "5.5") : ReadPESHeaderV5()
            Case "#PES0050" : pattern.SetMetadata(PROP_VERSION, "5") : ReadPESHeaderV5()
            Case "#PES0040" : pattern.SetMetadata(PROP_VERSION, "4") : ReadPESHeaderV4()
            Case "#PES0030" : pattern.SetMetadata(PROP_VERSION, "3") : ReadPESHeaderDefault()
            Case "#PES0022" : pattern.SetMetadata(PROP_VERSION, "2.2") : ReadPESHeaderDefault()
            Case "#PES0020" : pattern.SetMetadata(PROP_VERSION, "2") : ReadPESHeaderDefault()
            Case "#PES0001" : pattern.SetMetadata(PROP_VERSION, "1") : ReadPESHeaderDefault()
            Case "#PEC0001"
                ' We're already at a raw PEC stream; nothing else to do here.
                Return
            Case Else
                ReadPESHeaderDefault()
        End Select
    End Sub

    ' --- Common small helpers ---

    Private Sub ReadDescriptions()
        Dim nameLen = ReadInt8()
        Dim name = ReadString(nameLen)
        If name.Length > 0 Then pattern.SetMetadata("design_name", name)

        Dim catLen = ReadInt8()
        pattern.SetMetadata("category", ReadString(catLen))

        Dim authLen = ReadInt8()
        pattern.SetMetadata("author", ReadString(authLen))

        Dim keyLen = ReadInt8()
        pattern.SetMetadata("keywords", ReadString(keyLen))

        Dim comLen = ReadInt8()
        pattern.SetMetadata("comments", ReadString(comLen))
    End Sub

    Private Sub ReadThread()
        Dim colorCodeLen = ReadInt8()
        Dim colorCode = ReadString(colorCodeLen)

        Dim r = ReadInt8()
        Dim g = ReadInt8()
        Dim b = ReadInt8()

        ' Skip 5 bytes of padding/unknowns per format
        Skip(5)

        Dim descLen = ReadInt8()
        Dim desc = ReadString(descLen)

        Dim brandLen = ReadInt8()
        Dim brand = ReadString(brandLen)

        Dim chartLen = ReadInt8()
        Dim chart = ReadString(chartLen)

        Dim rgb = ((r And &HFF) << 16) Or ((g And &HFF) << 8) Or (b And &HFF)
        pattern.Add(New EmbThread(rgb, desc, colorCode, brand, chart))
    End Sub

    ' --- Version-specific headers (all ensure we Seek(pecStart)) ---

    Private Sub ReadPESHeaderDefault()
        Dim pecStart = ReadInt32LE()
        Seek(pecStart)
    End Sub

    Private Sub ReadPESHeaderV4()
        Dim pecStart = ReadInt32LE()
        Skip(4)
        ReadDescriptions()
        Seek(pecStart)
    End Sub

    Private Sub ReadPESHeaderV5()
        Dim pecStart = ReadInt32LE()
        Skip(4)
        ReadDescriptions()

        ' Skip image-from-data block
        Skip(24)
        Dim fromImgLen = ReadInt8()
        Skip(fromImgLen)

        Skip(24)

        ' If any of these sections are non-zero, PES has extra data we're not using.
        Dim numFill = ReadInt16LE() : If numFill <> 0 Then Seek(pecStart) : Return
        Dim numMotif = ReadInt16LE() : If numMotif <> 0 Then Seek(pecStart) : Return
        Dim feather = ReadInt16LE() : If feather <> 0 Then Seek(pecStart) : Return

        Dim numColors = ReadInt16LE()
        For i = 0 To numColors - 1
            ReadThread()
        Next

        Seek(pecStart)
    End Sub

    Private Sub ReadPESHeaderV6()
        Dim pecStart = ReadInt32LE()
        Skip(4)
        ReadDescriptions()

        Skip(36)
        Dim imgLen = ReadInt8()
        Dim imgStr = ReadString(imgLen)
        If imgStr.Length > 0 Then pattern.SetMetadata("image_file", imgStr)

        Skip(24)

        Dim numFill = ReadInt16LE() : If numFill <> 0 Then Seek(pecStart) : Return
        Dim numMotif = ReadInt16LE() : If numMotif <> 0 Then Seek(pecStart) : Return
        Dim feather = ReadInt16LE() : If feather <> 0 Then Seek(pecStart) : Return

        Dim numColors = ReadInt16LE()
        For i = 0 To numColors - 1
            ReadThread()
        Next

        Seek(pecStart)
    End Sub

    Private Sub ReadPESHeaderV7()
        Dim pecStart = ReadInt32LE()
        Skip(4)
        ReadDescriptions()

        Skip(36)
        Dim imgLen = ReadInt8()
        Dim imgStr = ReadString(imgLen)
        If imgStr.Length > 0 Then pattern.SetMetadata("image_file", imgStr)

        Skip(24)

        Dim numFill = ReadInt16LE() : If numFill <> 0 Then Seek(pecStart) : Return
        Dim numMotif = ReadInt16LE() : If numMotif <> 0 Then Seek(pecStart) : Return
        Dim feather = ReadInt16LE() : If feather <> 0 Then Seek(pecStart) : Return

        Dim numColors = ReadInt16LE()
        For i = 0 To numColors - 1
            ReadThread()
        Next

        Seek(pecStart)
    End Sub

    Private Sub ReadPESHeaderV8()
        Dim pecStart = ReadInt32LE()
        Skip(4)
        ReadDescriptions()

        Skip(38)
        Dim imgLen = ReadInt8()
        Dim imgStr = ReadString(imgLen)
        If imgStr.Length > 0 Then pattern.SetMetadata("image_file", imgStr)

        Skip(26)

        Dim numFill = ReadInt16LE() : If numFill <> 0 Then Seek(pecStart) : Return
        Dim numMotif = ReadInt16LE() : If numMotif <> 0 Then Seek(pecStart) : Return
        Dim feather = ReadInt16LE() : If feather <> 0 Then Seek(pecStart) : Return

        Dim numColors = ReadInt16LE()
        For i = 0 To numColors - 1
            ReadThread()
        Next

        Seek(pecStart)
    End Sub

    Private Sub ReadPESHeaderV9()
        Dim pecStart = ReadInt32LE()
        Skip(4)
        ReadDescriptions()

        Skip(14)
        Dim hoopLen = ReadInt8()
        Dim hoopName = ReadString(hoopLen)
        If hoopName.Length > 0 Then pattern.SetMetadata("hoop_name", hoopName)

        Skip(30)
        Dim imgLen = ReadInt8()
        Dim imgStr = ReadString(imgLen)
        If imgStr.Length > 0 Then pattern.SetMetadata("image_file", imgStr)

        Skip(34)

        Dim numFill = ReadInt16LE() : If numFill <> 0 Then Seek(pecStart) : Return
        Dim numMotif = ReadInt16LE() : If numMotif <> 0 Then Seek(pecStart) : Return
        Dim feather = ReadInt16LE() : If feather <> 0 Then Seek(pecStart) : Return

        Dim numColors = ReadInt16LE()
        For i = 0 To numColors - 1
            ReadThread()
        Next

        Seek(pecStart)
    End Sub

    Private Sub ReadPESHeaderV10()
        Dim pecStart = ReadInt32LE()
        Skip(4)
        ReadDescriptions()

        Skip(14)
        Dim hoopLen = ReadInt8()
        Dim hoopName = ReadString(hoopLen)
        If hoopName.Length > 0 Then pattern.SetMetadata("hoop_name", hoopName)

        Skip(38)
        Dim imgLen = ReadInt8()
        Dim imgStr = ReadString(imgLen)
        If imgStr.Length > 0 Then pattern.SetMetadata("image_file", imgStr)

        Skip(34)

        Dim numFill = ReadInt16LE() : If numFill <> 0 Then Seek(pecStart) : Return
        Dim numMotif = ReadInt16LE() : If numMotif <> 0 Then Seek(pecStart) : Return
        Dim feather = ReadInt16LE() : If feather <> 0 Then Seek(pecStart) : Return

        Dim numColors = ReadInt16LE()
        For i = 0 To numColors - 1
            ReadThread()
        Next

        Seek(pecStart)
    End Sub
End Class
