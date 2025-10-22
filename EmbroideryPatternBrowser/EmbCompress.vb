' Port of org.embroideryio.embroideryio.EmbCompress (decompress only)
' This version builds a proper *canonical* Huffman code and does bit-accurate lookups.
' Public API is unchanged: Expand(data As Byte(), uncompressedSize As Integer) As Byte()

Public NotInheritable Class EmbCompress
    Private Sub New()
    End Sub

    Public Shared Function Expand(data As Byte(), uncompressedSize As Integer) As Byte()
        Dim c As New EmbCompressInstance()
        Return c.Decompress(data, uncompressedSize)
    End Function

    ' ----------------- Instance -----------------
    Private Class EmbCompressInstance
        Private input As Byte()
        Private bitPos As Integer

        Private blockElements As Integer = -1
        Private charLenHuff As Huffman
        Private charHuff As Huffman
        Private distHuff As Huffman

        Public Function Decompress(src As Byte(), uncompressedSize As Integer) As Byte()
            input = If(src, Array.Empty(Of Byte)())
            bitPos = 0
            blockElements = -1
            charLenHuff = Nothing
            charHuff = Nothing
            distHuff = Nothing

            Dim outBuf As New List(Of Byte)(If(uncompressedSize > 0, uncompressedSize, 4096))
            Dim totalBits As Integer = input.Length * 8

            While (bitPos < totalBits) AndAlso (uncompressedSize = -1 OrElse outBuf.Count < uncompressedSize)
                Dim token As Integer = GetToken()
                If token <= 255 Then
                    outBuf.Add(CByte(token)) ' literal
                ElseIf token = 510 Then
                    Exit While ' end
                Else
                    ' back-reference
                    Dim length As Integer = token - 253 ' min 3
                    Dim back As Integer = GetPosition() + 1
                    Dim start As Integer = outBuf.Count - back
                    If start < 0 Then Exit While

                    ' copy with overlap allowed
                    For i As Integer = 0 To length - 1
                        outBuf.Add(outBuf(start + i))
                    Next
                End If
            End While

            If uncompressedSize > 0 Then
                If outBuf.Count > uncompressedSize Then
                    Return outBuf.GetRange(0, uncompressedSize).ToArray()
                ElseIf outBuf.Count < uncompressedSize Then
                    Dim dst(uncompressedSize - 1) As Byte
                    If outBuf.Count > 0 Then outBuf.CopyTo(dst, 0)
                    Return dst
                End If
            End If

            Return outBuf.ToArray()
        End Function

        ' -------- bit IO --------
        Private Function Peek(bits As Integer) As Integer
            Return GetBits(bitPos, bits)
        End Function
        Private Function Pop(bits As Integer) As Integer
            Dim v = Peek(bits)
            bitPos += bits
            Return v
        End Function

        Private Function GetBits(startBit As Integer, length As Integer) As Integer
            If length <= 0 Then Return 0
            Dim endBit As Integer = startBit + length - 1
            Dim startByte As Integer = startBit \ 8
            Dim endByte As Integer = endBit \ 8

            Dim val As Integer = 0
            For i = startByte To endByte
                val = (val << 8)
                If i >= 0 AndAlso i < input.Length Then
                    val = val Or (input(i) And &HFF)
                End If
            Next

            Dim unusedRight As Integer = (8 - ((endBit + 1) Mod 8)) Mod 8
            Dim mask As Integer = (1 << length) - 1
            Return (val >> unusedRight) And mask
        End Function

        Private Function ReadVarLen() As Integer
            Dim m As Integer = Pop(3)
            If m <> 7 Then Return m
            For q As Integer = 0 To 12
                If Pop(1) = 1 Then
                    m += 1
                Else
                    Exit For
                End If
            Next
            Return m
        End Function

        ' -------- block loaders --------
        Private Sub LoadBlock()
            blockElements = Pop(16)
            charLenHuff = LoadCharacterLengthHuffman()
            charHuff = LoadCharacterHuffman(charLenHuff)
            distHuff = LoadDistanceHuffman()
        End Sub

        Private Function LoadCharacterLengthHuffman() As Huffman
            Dim count As Integer = Pop(5)
            If count = 0 Then
                Dim v As Integer = Pop(5)
                Return Huffman.FromDefault(v)
            End If
            Dim lengths(count - 1) As Integer
            Dim idx As Integer = 0
            While idx < count
                If idx = 3 Then
                    idx += Pop(2) ' skip up to 3
                    If idx >= count Then Exit While
                End If
                lengths(idx) = ReadVarLen()
                idx += 1
            End While
            Return Huffman.FromCodeLengths(lengths, maxLookupBits:=16)
        End Function

        Private Function LoadCharacterHuffman(lengthH As Huffman) As Huffman
            Dim count As Integer = Pop(9)
            If count = 0 Then
                Dim v As Integer = Pop(9)
                Return Huffman.FromDefault(v)
            End If

            Dim lengths(count - 1) As Integer
            Dim idx As Integer = 0
            While idx < count
                Dim res = lengthH.Lookup(Peek(16))
                Dim c As Integer = res.Symbol
                bitPos += res.BitLength
                Select Case c
                    Case 0
                        idx += 1                    ' skip 1
                    Case 1
                        idx += 3 + Pop(4)          ' skip 3 + read(4)
                    Case 2
                        idx += 20 + Pop(9)         ' skip 20 + read(9)
                    Case Else
                        c -= 2
                        If idx < lengths.Length Then
                            lengths(idx) = c : idx += 1
                        Else
                            Exit While
                        End If
                End Select
            End While
            Return Huffman.FromCodeLengths(lengths, maxLookupBits:=16)
        End Function

        Private Function LoadDistanceHuffman() As Huffman
            Dim count As Integer = Pop(5)
            If count = 0 Then
                Dim v As Integer = Pop(5)
                Return Huffman.FromDefault(v)
            End If
            Dim lengths(count - 1) As Integer
            For i = 0 To count - 1
                lengths(i) = ReadVarLen()
            Next
            Return Huffman.FromCodeLengths(lengths, maxLookupBits:=16)
        End Function

        ' -------- token/position --------
        Private Function GetToken() As Integer
            If blockElements <= 0 Then LoadBlock()
            blockElements -= 1
            Dim r = charHuff.Lookup(Peek(16))
            bitPos += r.BitLength
            Return r.Symbol
        End Function

        Private Function GetPosition() As Integer
            Dim r = distHuff.Lookup(Peek(16))
            bitPos += r.BitLength
            If r.Symbol = 0 Then Return 0
            Dim v As Integer = r.Symbol - 1
            v = (1 << v) + Pop(v)
            Return v
        End Function

        ' ----------------- Canonical Huffman -----------------
        Private Class Huffman
            Private ReadOnly _haveTable As Boolean
            Private ReadOnly _defaultValue As Integer
            Private ReadOnly _maxLen As Integer
            Private ReadOnly _tables As Dictionary(Of Integer, Integer)() ' index by bit length (1.._maxLen)
            Private ReadOnly _minLen As Integer

            Private Sub New(defaultVal As Integer)
                _haveTable = False
                _defaultValue = defaultVal
                _maxLen = 0
                _minLen = 0
            End Sub

            Private Sub New(minLen As Integer, maxLen As Integer, tables As Dictionary(Of Integer, Integer)())
                _haveTable = True
                _defaultValue = 0
                _minLen = minLen
                _maxLen = maxLen
                _tables = tables
            End Sub

            Public Shared Function FromDefault(v As Integer) As Huffman
                Return New Huffman(v)
            End Function

            ' Build canonical Huffman from code lengths.
            ' lengths(s) = bit-length for symbol s; 0 => unused symbol.
            Public Shared Function FromCodeLengths(lengths As Integer(), Optional maxLookupBits As Integer = 16) As Huffman
                If lengths Is Nothing OrElse lengths.Length = 0 Then
                    Return New Huffman(0)
                End If

                Dim maxLen As Integer = 0, minLen As Integer = Integer.MaxValue
                Dim counts As New Dictionary(Of Integer, Integer)()
                For i = 0 To lengths.Length - 1
                    Dim L = lengths(i)
                    If L <= 0 Then Continue For
                    If L > maxLen Then maxLen = L
                    If L < minLen Then minLen = L
                    counts(L) = If(counts.ContainsKey(L), counts(L) + 1, 1)
                Next
                If maxLen = 0 OrElse minLen = Integer.MaxValue Then
                    Return New Huffman(0)
                End If
                maxLen = Math.Min(maxLen, maxLookupBits)

                ' Canonical code generation
                Dim nextCode(maxLen) As Integer
                Dim code As Integer = 0
                For L = 1 To maxLen
                    code = (code + If(counts.ContainsKey(L - 1), counts(L - 1), 0)) << 1
                    nextCode(L) = code
                Next

                ' Build per-length tables: code -> symbol
                Dim tables(maxLen) As Dictionary(Of Integer, Integer)
                For L = 0 To maxLen
                    tables(L) = New Dictionary(Of Integer, Integer)()
                Next

                For sym = 0 To lengths.Length - 1
                    Dim L = lengths(sym)
                    If L <= 0 OrElse L > maxLen Then Continue For
                    Dim c As Integer = nextCode(L)
                    nextCode(L) = c + 1
                    If Not tables(L).ContainsKey(c) Then
                        tables(L).Add(c, sym)
                    End If
                Next

                Return New Huffman(minLen, maxLen, tables)
            End Function

            Public Structure LookupResult
                Public Symbol As Integer
                Public BitLength As Integer
            End Structure

            ' Look up using up to 16 preview bits: try shortest length first per canonical rules.
            Public Function Lookup(preview16 As Integer) As LookupResult
                If Not _haveTable Then
                    Return New LookupResult With {.Symbol = _defaultValue, .BitLength = 0}
                End If
                ' Try lengths from min to max; extract top L bits and compare with canonical codes.
                For L = _minLen To _maxLen
                    Dim code As Integer = (preview16 >> (16 - L)) And ((1 << L) - 1)
                    Dim dict = _tables(L)
                    Dim sym As Integer = 0
                    If dict IsNot Nothing AndAlso dict.TryGetValue(code, sym) Then
                        Return New LookupResult With {.Symbol = sym, .BitLength = L}
                    End If
                Next
                ' No match (corrupt stream) – return default
                Return New LookupResult With {.Symbol = _defaultValue, .BitLength = 0}
            End Function
        End Class
    End Class
End Class
