Imports System.Text

Public Class DstReader
    Inherits EmbReader

    Private Const DSTHEADERSIZE As Integer = 512
    Private Const COMMANDSIZE As Integer = 3

    ' Defensive caps to keep malformed inputs from causing huge work
    Private Const MAX_HEADER_BYTES As Integer = 4096          ' hard ceiling for header reads (should never be hit)
    Private Const MAX_STITCH_COMMANDS As Integer = 5_000_000  ' ~5M tri-bytes ~ 15MB stream; plenty for real DSTs
    Private Const MAX_COLOR_ENTRIES As Integer = 1024         ' guard TC: lines
    Private Const MAX_LABEL_LEN As Integer = 256              ' LA and other text fields
    Private _fileName As String = ""

    Public Overrides Sub Read(fn As String)
        _fileName = fn
        Try
            DstReadHeader()
            DstReadStitches()
        Catch ex As Exception
            Logger.Debug("DstReader: error in file: " & _fileName & " - " & ex.Message, ex.StackTrace.ToString)
            Throw
        End Try
    End Sub

    ' -------- Header --------
    Private Sub DstReadHeader()
        ' Read exactly 512 bytes per DST spec; tolerate shorter files (EOF) gracefully.
        Dim bytesToRead As Integer = Math.Min(DSTHEADERSIZE, MAX_HEADER_BYTES)
        Dim b As Byte() = reader.ReadBytes(bytesToRead)
        ReadPosition += b.Length
        If b.Length = 0 Then Exit Sub ' empty or too short; pattern will just have stitches decoded
        ParseDataStitchHeader(pattern, b)
    End Sub

    Private Shared Sub ParseDataStitchHeader(p As EmbPattern, b As Byte())
        Dim bytestring As String = Encoding.ASCII.GetString(b)
        ' Split on CR/LF; ignore empty lines
        Dim lines = bytestring.Split({ControlChars.Lf, ControlChars.Cr}, StringSplitOptions.RemoveEmptyEntries)

        Dim colorsSeen As Integer = 0
        For Each raw In lines
            If raw Is Nothing OrElse raw.Length < 3 Then Continue For

            ' Defensive: prefix is two chars + space (e.g., "LA:foo" or "LA foo"); normalize
            Dim s As String = raw.Trim()
            If s.Length < 3 Then Continue For

            Dim prefix As String = s.Substring(0, 2)
            Dim value As String = ""
            If s.Length > 3 Then value = s.Substring(3).Trim()

            ' Cap values to avoid huge metadata
            If value IsNot Nothing AndAlso value.Length > MAX_LABEL_LEN Then
                value = value.Substring(0, MAX_LABEL_LEN)
            End If

            Select Case prefix
                Case "LA" ' label
                    If Not String.IsNullOrEmpty(value) Then p.SetMetadata("name", value)
                Case "AU"
                    If Not String.IsNullOrEmpty(value) Then p.SetMetadata("author", value)
                Case "CP"
                    If Not String.IsNullOrEmpty(value) Then p.SetMetadata("copyright", value)
                Case "TC"
                    If colorsSeen >= MAX_COLOR_ENTRIES Then Exit For
                    ' Format commonly: "#RRGGBB,Description,Catalog"
                    Dim parts = value.Split(","c)
                    If parts.Length > 0 Then
                        Dim color As Integer = ParseColorString(parts(0))
                        Dim desc As String = If(parts.Length > 1, SafeSlice(parts(1), MAX_LABEL_LEN), "DST")
                        Dim catalog As String = If(parts.Length > 2, SafeSlice(parts(2), MAX_LABEL_LEN), "")
                        p.Add(New EmbThread(color, desc, catalog, "DST", "DST"))
                        colorsSeen += 1
                    End If
                Case Else
                    ' ignore unknown tokens
            End Select
        Next
    End Sub

    Private Shared Function SafeSlice(s As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(s) Then Return String.Empty
        If s.Length <= maxLen Then Return s
        Return s.Substring(0, maxLen)
    End Function

    Private Shared Function ParseColorString(s As String) As Integer
        If String.IsNullOrWhiteSpace(s) Then Return &H0
        Dim t As String = s.Trim()
        If t.StartsWith("#") Then t = t.Substring(1)
        If t.StartsWith("0x", StringComparison.OrdinalIgnoreCase) Then t = t.Substring(2)
        If t.Length >= 6 Then
            ' guard against non-hex content
            Try
                Dim r As Integer = Convert.ToInt32(t.Substring(0, 2), 16)
                Dim g As Integer = Convert.ToInt32(t.Substring(2, 2), 16)
                Dim b As Integer = Convert.ToInt32(t.Substring(4, 2), 16)
                Return (r << 16) Or (g << 8) Or b
            Catch
                Return &H0
            End Try
        End If
        Dim val As Integer
        If Integer.TryParse(t, val) Then Return val
        Return &H0
    End Function

    ' -------- Stitches --------
    Private Sub DstReadStitches()
        Dim cmd(2) As Byte
        Dim sequinMode As Boolean = False
        Dim commandsRead As Integer = 0

        While True
            Dim got As Integer = reader.Read(cmd, 0, COMMANDSIZE)
            If got <> COMMANDSIZE Then Exit While ' EOF or partial read: stop gracefully
            ReadPosition += got

            commandsRead += 1
            If commandsRead > MAX_STITCH_COMMANDS Then
                Logger.Debug("DstReader: error in filename:" & _fileName & " - Too many stitch commands; file may be corrupt.")
                Exit While
            End If

            Dim b0 As Integer = cmd(0)
            Dim b1 As Integer = cmd(1)
            Dim b2 As Integer = cmd(2)

            Dim dx As Integer = DecodeDx(b0, b1, b2)
            Dim dy As Integer = DecodeDy(b0, b1, b2)

            ' Flag decoding per Tajima DST spec
            If (b2 And &HF3) = &HF3 Then
                ' END
                If dx <> 0 OrElse dy <> 0 Then pattern.move(dx, dy)
                pattern.end()
                Exit While
            ElseIf (b2 And &HC3) = &HC3 Then
                ' COLOR CHANGE
                pattern.color_change(dx, dy)
            ElseIf (b2 And &H43) = &H43 Then
                ' SEQUIN MODE toggle (no direct EmbPattern support; keep as state)
                sequinMode = Not sequinMode
                pattern.trim()
                If dx <> 0 OrElse dy <> 0 Then pattern.move(dx, dy)
            ElseIf (b2 And &H83) = &H83 Then
                ' JUMP or SEQUIN EJECT
                If sequinMode Then
                    pattern.trim()
                    If dx <> 0 OrElse dy <> 0 Then pattern.move(dx, dy)
                Else
                    pattern.move(dx, dy)
                End If
            Else
                ' Normal stitch
                pattern.stitch(dx, dy)
            End If
        End While
    End Sub

    ' Bit helpers
    Private Shared Function GetBit(b As Integer, pos As Integer) As Integer
        Return (b >> pos) And 1
    End Function

    Private Shared Function DecodeDx(b0 As Integer, b1 As Integer, b2 As Integer) As Integer
        Dim x As Integer = 0
        x += GetBit(b2, 2) * (+81)
        x += GetBit(b2, 3) * (-81)
        x += GetBit(b1, 2) * (+27)
        x += GetBit(b1, 3) * (-27)
        x += GetBit(b0, 2) * (+9)
        x += GetBit(b0, 3) * (-9)
        x += GetBit(b1, 0) * (+3)
        x += GetBit(b1, 1) * (-3)
        x += GetBit(b0, 0) * (+1)
        x += GetBit(b0, 1) * (-1)
        Return x
    End Function

    Private Shared Function DecodeDy(b0 As Integer, b1 As Integer, b2 As Integer) As Integer
        Dim y As Integer = 0
        y += GetBit(b2, 5) * (+81)
        y += GetBit(b2, 4) * (-81)
        y += GetBit(b1, 5) * (+27)
        y += GetBit(b1, 4) * (-27)
        y += GetBit(b0, 5) * (+9)
        y += GetBit(b0, 4) * (-9)
        y += GetBit(b1, 7) * (+3)
        y += GetBit(b1, 6) * (-3)
        y += GetBit(b0, 7) * (+1)
        y += GetBit(b0, 6) * (-1)
        Return -y ' Y is inverted in render space
    End Function
End Class
