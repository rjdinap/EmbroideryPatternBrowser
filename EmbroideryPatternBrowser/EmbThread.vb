Imports System.Drawing

Public Class EmbThread
    Public Property ColorValue As Color
    Public Property Description As String
    Public Property ColorCode As String
    Public Property Brand As String
    Public Property ThreadChart As String

    Public Sub New(colorRgb As Integer, description As String, colorCode As String, brand As String, threadChart As String)
        Dim r As Integer = (colorRgb >> 16) And &HFF
        Dim g As Integer = (colorRgb >> 8) And &HFF
        Dim b As Integer = colorRgb And &HFF
        Me.ColorValue = Color.FromArgb(255, r, g, b)
        Me.Description = If(description, "")
        Me.ColorCode = If(colorCode, "")
        Me.Brand = If(brand, "")
        Me.ThreadChart = If(threadChart, "")
    End Sub
End Class
