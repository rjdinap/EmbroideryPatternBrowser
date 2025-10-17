Imports System
Imports System.IO
Imports System.Text

Public MustInherit Class EmbReader
    Public pattern As New EmbPattern()
    Public reader As BinaryReader
    Public ReadPosition As Integer
    Protected baseStream As Stream



    Public Overridable Sub Read()
    End Sub


    'old method - doesn't hanlde zip files
    'Public Sub Load(path As String)
    '    Try
    '        Using fs As New FileStream(path, FileMode.Open, FileAccess.Read)
    '            reader = New BinaryReader(fs)
    '            ReadPosition = 0
    '            Read()
    '        End Using
    '    Catch ex As Exception
    '        'fail silently if we can't load a file
    '    End Try
    'End Sub

    ' === NEW: central load that supports disk files AND zip-inner files ===
    Public Overridable Sub Load(inputPath As String)
        ' Clean up any previous stream/reader
        Cleanup()

        Dim zp As New ZipProcessing()
        Dim parsed = zp.ParseFilename(inputPath)
        Dim outerPath As String = parsed.Item1      ' path before '?'
        Dim outerExt As String = parsed.Item2       ' extension of outer path
        Dim innerPath As String = parsed.Item3      ' path after '?', may be ""

        Try
            If IsZipPath(outerExt) AndAlso Not String.IsNullOrEmpty(innerPath) Then
                ' Read from inside the .zip
                baseStream = zp.ReadZip(outerPath, innerPath)   ' returns MemoryStream
            Else
                ' Normal file on disk (either no '?', or it's not a .zip)
                baseStream = New FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read)
            End If

            reader = New BinaryReader(baseStream)
            ' Hand off to the format-specific reader
            Read()

        Finally
            ' If your derived readers need the stream longer, move disposal out as needed.
            ' Default behavior: close as before.
            Cleanup()
        End Try
    End Sub


    Protected Overridable Function IsZipPath(ext As String) As Boolean
        Return String.Equals(ext, ".zip", StringComparison.OrdinalIgnoreCase)
    End Function

    Protected Overridable Sub Cleanup()
        If reader IsNot Nothing Then
            Try : reader.Close() : Catch : End Try
            reader = Nothing
        End If
        If baseStream IsNot Nothing Then
            Try : baseStream.Dispose() : Catch : End Try
            baseStream = Nothing
        End If
    End Sub



    Public Function GetPattern() As EmbPattern
        Return pattern
    End Function

    Protected Function ReadInt8() As Integer
        If reader Is Nothing OrElse reader.BaseStream.Position >= reader.BaseStream.Length Then
            Return Integer.MinValue
        End If
        ReadPosition += 1
        Return CInt(reader.ReadByte())
    End Function

    Protected Function ReadInt16LE() As Integer
        If reader.BaseStream.Position + 2 > reader.BaseStream.Length Then Return Integer.MinValue
        ReadPosition += 2
        Dim b1 = CInt(reader.ReadByte())
        Dim b2 = CInt(reader.ReadByte())
        Return b1 Or (b2 << 8)
    End Function

    Protected Function ReadInt24LE() As Integer
        If reader.BaseStream.Position + 3 > reader.BaseStream.Length Then Return Integer.MinValue
        ReadPosition += 3
        Dim b1 = CInt(reader.ReadByte())
        Dim b2 = CInt(reader.ReadByte())
        Dim b3 = CInt(reader.ReadByte())
        Return b1 Or (b2 << 8) Or (b3 << 16)
    End Function

    Protected Function ReadInt32LE() As Integer
        If reader.BaseStream.Position + 4 > reader.BaseStream.Length Then Return Integer.MinValue
        ReadPosition += 4
        Dim b1 = CInt(reader.ReadByte())
        Dim b2 = CInt(reader.ReadByte())
        Dim b3 = CInt(reader.ReadByte())
        Dim b4 = CInt(reader.ReadByte())
        Return b1 Or (b2 << 8) Or (b3 << 16) Or (b4 << 24)
    End Function

    Protected Function ReadString(length As Integer) As String
        If length <= 0 Then Return String.Empty
        Dim bytes = reader.ReadBytes(length)
        ReadPosition += bytes.Length
        Return Encoding.ASCII.GetString(bytes)
    End Function

    Protected Function ReadFully(buffer As Byte()) As Integer
        Dim read = reader.Read(buffer, 0, buffer.Length)
        ReadPosition += read
        Return read
    End Function

    Protected Sub Skip(count As Integer)
        If count <= 0 Then Return
        reader.BaseStream.Seek(count, SeekOrigin.Current)
        ReadPosition += count
    End Sub

    Protected Sub Seek(pos As Integer)
        reader.BaseStream.Seek(pos, SeekOrigin.Begin)
        ReadPosition = pos
    End Sub

    Protected Function Tell() As Integer
        Return CInt(reader.BaseStream.Position)
    End Function


    Protected Function ReadAllRemainingBytes() As Byte()
        Dim remaining As Long = reader.BaseStream.Length - reader.BaseStream.Position
        If remaining <= 0 Then Return New Byte() {}
        Return reader.ReadBytes(CInt(remaining))
    End Function

    ' --- Big-endian integer readers (needed for VP3) ---

    Protected Function ReadInt16BE() As Integer
        If reader.BaseStream.Position + 2 > reader.BaseStream.Length Then Return Integer.MinValue
        ReadPosition += 2
        Dim b1 = CInt(reader.ReadByte())
        Dim b2 = CInt(reader.ReadByte())
        Return (b1 << 8) Or b2
    End Function

    Protected Function ReadInt24BE() As Integer
        If reader.BaseStream.Position + 3 > reader.BaseStream.Length Then Return Integer.MinValue
        ReadPosition += 3
        Dim b1 = CInt(reader.ReadByte())
        Dim b2 = CInt(reader.ReadByte())
        Dim b3 = CInt(reader.ReadByte())
        Return (b1 << 16) Or (b2 << 8) Or b3
    End Function

    Protected Function ReadInt32BE() As Integer
        If reader.BaseStream.Position + 4 > reader.BaseStream.Length Then Return Integer.MinValue
        ReadPosition += 4
        Dim b1 = CInt(reader.ReadByte())
        Dim b2 = CInt(reader.ReadByte())
        Dim b3 = CInt(reader.ReadByte())
        Dim b4 = CInt(reader.ReadByte())
        Return (b1 << 24) Or (b2 << 16) Or (b3 << 8) Or b4
    End Function


End Class
