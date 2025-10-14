Imports System
Imports System.IO

Public NotInheritable Class EmbReaderFactory
    Private Sub New()
    End Sub

    Public Shared Function Create(path As String) As EmbReader
        Dim ext As String = System.IO.Path.GetExtension(path).ToLowerInvariant()
        Select Case ext
            Case ".dst" : Return New DstReader()
            Case ".exp" : Return New ExpReader()
            Case ".hus" : Return New HusReader()
            Case ".jef" : Return New JefReader()
            Case ".pes", ".pec" : Return New PesReader() ' PesReader internally can read PEC after header
            Case ".sew" : Return New SewReader()
            Case ".vp3" : Return New Vp3Reader()
            Case ".vip" : Return New VipReader()
            Case ".xxx" : Return New XxxReader()
            Case Else
                Throw New NotSupportedException("Unsupported embroidery format: " & ext)
        End Select
    End Function

    Public Shared Function LoadPattern(path As String) As EmbPattern
        Dim r As EmbReader = Create(path)
        r.Load(path)
        Return r.Pattern
    End Function
End Class
