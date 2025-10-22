Imports System.Globalization

Public Class EmbThreadHus
    Inherits EmbThread

    Public Sub New(colorHex As String, description As String, catalogNumber As String)
        MyBase.New(ParseHexRgb(colorHex), description, catalogNumber, "Hus", "Hus")
    End Sub

    ''' <summary>
    ''' Returns the Hus color set as EmbThread instances (brand+chart = "Hus")
    ''' Catalog number mapped into EmbThread.ColorCode.
    ''' </summary>
    Public Shared Function GetThreadSet() As EmbThread()
        Dim list As New List(Of EmbThread) From {
            New EmbThreadHus("#000000", "Black", "026"),
            New EmbThreadHus("#0000e7", "Blue", "005"),
            New EmbThreadHus("#00c600", "Green", "002"),
            New EmbThreadHus("#ff0000", "Red", "014"),
            New EmbThreadHus("#840084", "Purple", "008"),
            New EmbThreadHus("#ffff00", "Yellow", "020"),
            New EmbThreadHus("#848484", "Grey", "024"),
            New EmbThreadHus("#8484e7", "Light Blue", "006"),
            New EmbThreadHus("#00ff84", "Light Green", "003"),
            New EmbThreadHus("#ff7b31", "Orange", "017"),
            New EmbThreadHus("#ff8ca5", "Pink", "011"),
            New EmbThreadHus("#845200", "Brown", "028"),
            New EmbThreadHus("#ffffff", "White", "022"),
            New EmbThreadHus("#000084", "Dark Blue", "004"),
            New EmbThreadHus("#008400", "Dark Green", "001"),
            New EmbThreadHus("#7b0000", "Dark Red", "013"),
            New EmbThreadHus("#ff6384", "Light Red", "015"),
            New EmbThreadHus("#522952", "Dark Purple", "007"),
            New EmbThreadHus("#ff00ff", "Light Purple", "009"),
            New EmbThreadHus("#ffde00", "Dark Yellow", "019"),
            New EmbThreadHus("#ffff9c", "Light Yellow", "021"),
            New EmbThreadHus("#525252", "Dark Grey", "025"),
            New EmbThreadHus("#d6d6d6", "Light Grey", "023"),
            New EmbThreadHus("#ff5208", "Dark Orange", "016"),
            New EmbThreadHus("#ff9c5a", "Light Orange", "018"),
            New EmbThreadHus("#ff52b5", "Dark Pink", "010"),
            New EmbThreadHus("#ffc6de", "Light Pink", "012"),
            New EmbThreadHus("#523100", "Dark Brown", "027"),
            New EmbThreadHus("#b5a584", "Light Brown", "029")
        }
        Return list.ToArray()
    End Function

    Private Shared Function ParseHexRgb(hex As String) As Integer
        Dim t As String = hex.Trim()
        If t.StartsWith("#") Then t = t.Substring(1)
        If t.StartsWith("0x", StringComparison.OrdinalIgnoreCase) Then t = t.Substring(2)
        If t.Length <> 6 Then Throw New FormatException("Invalid RGB hex: " & hex)
        Dim r As Integer = Integer.Parse(t.Substring(0, 2), NumberStyles.HexNumber)
        Dim g As Integer = Integer.Parse(t.Substring(2, 2), NumberStyles.HexNumber)
        Dim b As Integer = Integer.Parse(t.Substring(4, 2), NumberStyles.HexNumber)
        Return (r << 16) Or (g << 8) Or b
    End Function
End Class
