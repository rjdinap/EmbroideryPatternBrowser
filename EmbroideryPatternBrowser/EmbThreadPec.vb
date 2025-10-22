''' <summary>
''' Brother PEC 64-color palette (hard-coded, no external file IO).
''' Uses the RGB integer values that worked best in your tests.
''' </summary>
Public Class EmbThreadPec
    Inherits EmbThread

    Public Sub New(colorRgb As Integer, description As String)
        MyBase.New(colorRgb, description, "", "Brother", "PEC")
    End Sub

    ''' <summary>
    ''' Returns the 64-color PEC thread set as EmbThread instances.
    ''' </summary>
    Public Shared Function GetThreadSet() As EmbThread()
        ' RGB values are 0xRRGGBB (no alpha).
        Dim builtIn As Integer() = {
            &H0, &HE1F7C, &HA55A3, &H8777, &H4B6BAF, &HED171F,
            &HD15C00, &H913697, &HE49ACB, &H915FAC, &H9DD67D, &HE8A900,
            &HFEBA35, &HFFFF00, &H70BC1F, &HBA9800, &HA8A8A8, &H7B6F00,
            &HFFFFB3, &H4F5556, &H0, &HB3D91, &H770176, &H293133,
            &H2A1301, &HF64A8A, &HB27624, &HFCBBC5, &HFE370F, &HF0F0F0,
            &H6A1C8A, &HA8DDC4, &H2584BB, &HFEB343, &HFFF08D, &HD0A660,
            &HD15400, &H66BA49, &H134A46, &H878787, &HD8CAC6, &H435607,
            &HFEE3C5, &HF993BC, &H3822, &HB2AFD4, &H686AB0, &HEFE3B9,
            &HF73866, &HB54C64, &H132B1A, &HC70156, &HFE9E32, &HA8DEEB,
            &H671A, &H4E2990, &H2F7E20, &HFDD9DE, &HFFD911, &H95BA6,
            &HF0F970, &HE3F35B, &HFFC864, &HFFC896, &HFFC8C8
        }

        Dim list As New List(Of EmbThread)()
        For i As Integer = 0 To builtIn.Length - 1
            list.Add(New EmbThreadPec(builtIn(i), "PEC " & (i + 1).ToString()))
        Next
        Return list.ToArray()
    End Function
End Class
