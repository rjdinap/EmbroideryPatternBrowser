Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Linq


' Full, compact EmbPattern.vb (hardened)
' - Tracks stitches, thread colors, and per-stitch breaks (jump/trim/abs-move/color-change)
' - Renders to Image with optional density shading and debug toggles
' - Adds caps + logging (Form1.Status) to avoid huge allocations and silent failures
' - .NET 4.5+ / 4.8 compatible

Public Class EmbPattern
    Public Const PROP_NAME As String = "name"

    ' ---------- Data ----------
    Private metadata As New Dictionary(Of String, String)()
    Public ReadOnly ThreadList As New List(Of EmbThread)()
    Public ReadOnly Stitches As New List(Of PointF)()
    Public ReadOnly StitchThreadIndices As New List(Of Integer)()
    Public ReadOnly BreakBefore As New List(Of Boolean)() ' True => do not connect line from previous point

    ' ---------- Cursor / State ----------
    Private curX As Single
    Private curY As Single
    Private curThreadIndex As Integer = 0
    Private penUp As Boolean = True ' first stitch should not connect backward

    ' ---------- Guard rails (sizes in pixels) ----------
    Private Const MAX_W As Integer = 6000
    Private Const MAX_H As Integer = 6000
    Private Const MAX_PIXELS As Long = 36_000_000L   ' width * height hard cap (e.g., 6k x 6k)
    Private Const MAX_DENSITY_PIXELS As Long = 36_000_000L

    ' ---------- Metadata ----------
    Public Sub SetMetadata(key As String, value As String)
        metadata(key) = value
    End Sub

    Public Function GetMetadata(key As String) As String
        If metadata.ContainsKey(key) Then Return metadata(key)
        Return Nothing
    End Function

    ' ---------- Threads ----------
    Public Sub Add(t As EmbThread)
        ThreadList.Add(t)
    End Sub

    ' ---------- Embroidery Ops (match Java names) ----------
    Public Sub stitch(dx As Integer, dy As Integer)
        curX += dx : curY += dy
        Stitches.Add(New PointF(curX, curY))
        StitchThreadIndices.Add(curThreadIndex)
        BreakBefore.Add(penUp) ' if a jump/trim/color-change/abs-move happened before, break this segment
        penUp = False
    End Sub

    Public Sub move(dx As Integer, dy As Integer)
        curX += dx : curY += dy
        penUp = True
    End Sub

    Public Sub moveAbs(xAbs As Single, yAbs As Single)
        curX = xAbs : curY = yAbs
        penUp = True
    End Sub

    Public Sub trim()
        penUp = True
    End Sub

    Public Sub color_change()
        If ThreadList.Count > 0 Then
            If curThreadIndex < ThreadList.Count - 1 Then
                curThreadIndex += 1
            Else
                curThreadIndex = ThreadList.Count - 1
            End If
        End If
        penUp = True
    End Sub

    Public Sub color_change(dx As Integer, dy As Integer)
        move(dx, dy)
        color_change()
    End Sub

    Public Sub select_thread(index As Integer)
        If ThreadList.Count = 0 Then
            curThreadIndex = 0
            Return
        End If
        If index < 0 Then index = 0
        If index >= ThreadList.Count Then index = ThreadList.Count - 1
        curThreadIndex = index
        penUp = True   ' start new segment when switching threads
    End Sub

    Public Sub [end]()
        ' finalize if needed
    End Sub

    Public Function CreateImageFromPattern(Optional w As Integer = 0,
                                           Optional h As Integer = 0,
                                           Optional margin As Integer = 20,
                                           Optional densityShading As Boolean = False,
                                           Optional shadingStrength As Double = 1.0) As Image
        Try
            ' --- measure pattern bounds ---
            Dim bounds As RectangleF = ComputeBounds()
            If bounds.Width <= 0 OrElse bounds.Height <= 0 Then
                ' No stitches: draw a small palette preview
                Dim pw As Integer = If(w > 0, Math.Min(w, MAX_W), Math.Min(Math.Max(200, margin * 2 + ThreadList.Count * 50), MAX_W))
                Dim ph As Integer = If(h > 0, Math.Min(h, MAX_H), 140)
                If CLng(pw) * CLng(ph) > MAX_PIXELS Then
                    ScaleToMaxPixels(pw, ph, MAX_PIXELS)
                End If

                Dim bmpEmpty As New Bitmap(pw, ph, PixelFormat.Format24bppRgb)
                Using g = Graphics.FromImage(bmpEmpty)
                    g.SmoothingMode = SmoothingMode.AntiAlias
                    g.Clear(Color.White)
                    Dim x As Integer = margin
                    For Each t In ThreadList
                        Using br As New SolidBrush(t.ColorValue)
                            g.FillRectangle(br, x, margin, 40, Math.Max(1, ph - margin * 2))
                        End Using
                        g.DrawRectangle(Pens.Black, x, margin, 40, Math.Max(1, ph - margin * 2))
                        x += 50
                    Next
                End Using
                Return bmpEmpty
            End If

            ' --- decide scale & canvas size ---
            Dim scale As Double
            If w > 0 OrElse h > 0 Then
                ' Fit to the requested box preserving aspect ratio (inside fit)
                Dim targetW As Integer = If(w > 0, w, Integer.MaxValue)
                Dim targetH As Integer = If(h > 0, h, Integer.MaxValue)
                Dim innerW As Double = Math.Max(1, targetW - margin * 2)
                Dim innerH As Double = Math.Max(1, targetH - margin * 2)
                Dim sx As Double = innerW / bounds.Width
                Dim sy As Double = innerH / bounds.Height
                scale = Math.Max(0.01, Math.Min(sx, sy))
            Else
                ' No requested size: choose a sensible default scale
                Dim defaultScale As Double = 0.2
                scale = defaultScale
            End If

            Dim width As Integer = Math.Max(1, CInt(Math.Ceiling(bounds.Width * scale)) + margin * 2)
            Dim height As Integer = Math.Max(1, CInt(Math.Ceiling(bounds.Height * scale)) + margin * 2)

            ' Cap gigantic canvases (dimension + pixel count)
            width = Math.Min(width, MAX_W)
            height = Math.Min(height, MAX_H)
            If CLng(width) * CLng(height) > MAX_PIXELS Then
                ScaleToMaxPixels(width, height, MAX_PIXELS)
            End If

            ' --- transform origin ---
            Dim tx As Double = margin - (bounds.Left * scale)
            Dim ty As Double = margin - (bounds.Top * scale)

            ' Build color indices for rendering
            Dim useIndices As Integer() = BuildRenderColorIndices()

            ' Optional density map
            Dim heat As Integer(,) = Nothing
            Dim maxHeat As Integer = 0
            If densityShading AndAlso Stitches.Count > 1 Then
                ' guard density grid size, too
                Dim dw As Integer = width, dh As Integer = height
                If CLng(dw) * CLng(dh) > MAX_DENSITY_PIXELS Then
                    ScaleToMaxPixels(dw, dh, MAX_DENSITY_PIXELS)
                End If
                BuildDensityMap(dw, dh, scale, tx, ty, useIndices, heat, maxHeat)
                If maxHeat <= 0 Then densityShading = False
            End If

            Dim bmp As New Bitmap(width, height, PixelFormat.Format24bppRgb)
            Using g As Graphics = Graphics.FromImage(bmp)
                g.SmoothingMode = SmoothingMode.AntiAlias
                g.Clear(Color.White)

                If Stitches.Count < 2 Then
                    ' palette preview
                    Dim x As Integer = margin
                    For Each t In ThreadList
                        Using br As New SolidBrush(t.ColorValue)
                            g.FillRectangle(br, x, margin, 40, Math.Max(1, height - margin * 2))
                        End Using
                        g.DrawRectangle(Pens.Black, x, margin, 40, Math.Max(1, height - margin * 2))
                        x += 50
                    Next
                    Return bmp
                End If

                Dim penW As Single = Math.Max(0.4F / CSng(Math.Max(scale, 0.0001)), 0.1F)

                Dim lastPx As PointF = New PointF(CSng(Stitches(0).X * scale + tx), CSng(Stitches(0).Y * scale + ty))
                Dim lastIdx As Integer = If(useIndices.Length > 0, Math.Max(0, Math.Min(ThreadList.Count - 1, useIndices(0))), 0)

                Using pen As New Pen(GetThreadColor(lastIdx), penW)
                    pen.StartCap = LineCap.Round
                    pen.EndCap = LineCap.Round

                    For i As Integer = 1 To Stitches.Count - 1
                        Dim thisIdx As Integer = If(i < useIndices.Length, Math.Max(0, Math.Min(ThreadList.Count - 1, useIndices(i))), lastIdx)
                        Dim thisBreak As Boolean = If(i < BreakBefore.Count, BreakBefore(i), False)
                        Dim pPx As PointF = New PointF(CSng(Stitches(i).X * scale + tx), CSng(Stitches(i).Y * scale + ty))

                        ' color change => switch pen, don't connect
                        If thisIdx <> lastIdx Then
                            Dim baseCol = GetThreadColor(thisIdx)
                            If pen.Color.ToArgb() <> baseCol.ToArgb() Then pen.Color = baseCol
                            lastPx = pPx
                            lastIdx = thisIdx
                            Continue For
                        End If

                        ' structural break => don't connect
                        If thisBreak Then
                            lastPx = pPx
                            Continue For
                        End If

                        ' density shading (optional)
                        If densityShading AndAlso heat IsNot Nothing Then
                            Dim d1 As Double = SampleHeat(heat, maxHeat, CInt(Math.Round(lastPx.X)), CInt(Math.Round(lastPx.Y)))
                            Dim d2 As Double = SampleHeat(heat, maxHeat, CInt(Math.Round(pPx.X)), CInt(Math.Round(pPx.Y)))
                            Dim d As Double = Math.Min(1.0, Math.Max(0.0, (d1 + d2) * 0.5 * shadingStrength))
                            Dim minA As Integer = 96, maxA As Integer = 255
                            Dim alpha As Integer = ClampInt(CInt(Math.Round(minA + (maxA - minA) * d)), 0, 255)
                            Dim baseCol = GetThreadColor(thisIdx)
                            Dim shaded = Color.FromArgb(alpha, baseCol)
                            If pen.Color.ToArgb() <> shaded.ToArgb() Then pen.Color = shaded
                        Else
                            Dim baseCol = GetThreadColor(thisIdx)
                            If pen.Color.ToArgb() <> baseCol.ToArgb() Then pen.Color = baseCol
                        End If

                        g.DrawLine(pen, lastPx, pPx)
                        lastPx = pPx
                    Next
                End Using
            End Using

            Return bmp

        Catch ex As Exception
            ' Project-wide policy: log succinct error to the UI; avoid stack traces here.
            Try
                Form1.StatusFromAnyThread("Error: " & ex.Message, ex.StackTrace.ToString)
            Catch
                ' ignore secondary failures
            End Try
            Return Nothing
        End Try
    End Function

    ' ---------- Helpers ----------
    Public Function ComputeBounds() As RectangleF
        If Stitches.Count = 0 Then Return RectangleF.Empty
        Dim minX As Single = Single.MaxValue, minY As Single = Single.MaxValue
        Dim maxX As Single = Single.MinValue, maxY As Single = Single.MinValue
        For Each p In Stitches
            If p.X < minX Then minX = p.X
            If p.Y < minY Then minY = p.Y
            If p.X > maxX Then maxX = p.X
            If p.Y > maxY Then maxY = p.Y
        Next
        If maxX < minX OrElse maxY < minY Then Return RectangleF.Empty
        Return New RectangleF(minX, minY, maxX - minX, maxY - minY)
    End Function

    Private Function ToDevice(p As PointF, scale As Double, tx As Double, ty As Double) As PointF
        Return New PointF(CSng(p.X * scale + tx), CSng(p.Y * scale + ty))
    End Function

    Private Function ClampColorIndex(i As Integer) As Integer
        If ThreadList.Count = 0 Then Return 0
        If i < 0 Then Return 0
        If i >= ThreadList.Count Then Return ThreadList.Count - 1
        Return i
    End Function

    Private Function GetThreadColor(index As Integer) As Color
        If index >= 0 AndAlso index < ThreadList.Count Then
            Dim t = ThreadList(index)
            If t IsNot Nothing Then Return t.ColorValue
        End If
        Return Color.Black
    End Function

    ' Use decoded indices when available; otherwise synthesize contiguous blocks by color count.
    'Private Function BuildRenderColorIndices() As Integer()
    '    If Stitches.Count = 0 Then Return New Integer() {}

    '    If StitchThreadIndices.Count = Stitches.Count Then
    '        Dim distinctCount = StitchThreadIndices.Distinct().Count()
    '        If distinctCount > 1 OrElse ThreadList.Count <= 1 Then
    '            Return StitchThreadIndices.ToArray()
    '        End If
    '        ' fall through to synthetic if only one index but multiple threads
    '    End If

    '    Dim result(Stitches.Count - 1) As Integer
    '    If ThreadList.Count <= 1 Then Return result

    '    Dim colors As Integer = ThreadList.Count
    '    Dim baseBlock As Integer = Math.Max(1, Stitches.Count \ colors)
    '    For colorIdx As Integer = 0 To colors - 1
    '        Dim start As Integer = colorIdx * baseBlock
    '        Dim [end] As Integer = If(colorIdx = colors - 1, Stitches.Count, Math.Min(Stitches.Count, start + baseBlock))
    '        For i As Integer = start To [end] - 1
    '            result(i) = colorIdx
    '        Next
    '    Next
    '    Return result
    'End Function

    ' Use decoded indices when available; otherwise synthesize contiguous blocks by color count.
    Private Function BuildRenderColorIndices() As Integer()
        If Stitches.Count = 0 Then Return New Integer() {}

        ' --- Primary path: honor decoded indices exactly when present ---
        ' This matches the thumbnail path and avoids inventing colors for single-color designs.
        If StitchThreadIndices IsNot Nothing AndAlso StitchThreadIndices.Count = Stitches.Count Then
            Return StitchThreadIndices.ToArray()
        End If

        ' --- Fallback: no/partial indices → synthesize a simple block mapping by thread count ---
        Dim result(Stitches.Count - 1) As Integer
        If ThreadList Is Nothing OrElse ThreadList.Count <= 1 Then
            ' 0 or 1 thread → everything at index 0 (already result default)
            'Try : Form1.StatusFromAnyThread("[Render] Using synthetic color indices (<=1 thread or missing indices).") : Catch : End Try
            Return result
        End If

        Dim colors As Integer = ThreadList.Count
        Dim baseBlock As Integer = Math.Max(1, Stitches.Count \ colors)
        For colorIdx As Integer = 0 To colors - 1
            Dim start As Integer = colorIdx * baseBlock
            Dim [end] As Integer = If(colorIdx = colors - 1, Stitches.Count, Math.Min(Stitches.Count, start + baseBlock))
            For i As Integer = start To [end] - 1
                result(i) = colorIdx
            Next
        Next

        'Try : Form1.StatusFromAnyThread($"[Render] Using synthetic color indices (colors={colors}, stitches={Stitches.Count}).") : Catch : End Try
        Return result
    End Function


    ' ---------- Density map (optional shading) ----------
    Private Sub BuildDensityMap(width As Integer, height As Integer,
                                scale As Double, tx As Double, ty As Double,
                                useIndices As Integer(),
                                ByRef heat As Integer(,), ByRef maxHeat As Integer)
        If width <= 0 OrElse height <= 0 Then
            heat = Nothing : maxHeat = 0 : Return
        End If
        ' cap again defensively
        If CLng(width) * CLng(height) > MAX_DENSITY_PIXELS Then
            ScaleToMaxPixels(width, height, MAX_DENSITY_PIXELS)
        End If

        heat = New Integer(width - 1, height - 1) {}
        maxHeat = 0

        Dim last As PointF = ToDevice(Stitches(0), scale, tx, ty)
        Dim lastIdx As Integer = ClampColorIndex(If(useIndices.Length > 0, useIndices(0), 0))

        For i As Integer = 1 To Stitches.Count - 1
            Dim thisIdx As Integer = ClampColorIndex(If(i < useIndices.Length, useIndices(i), lastIdx))
            Dim thisBreak As Boolean = If(i < BreakBefore.Count, BreakBefore(i), False)
            Dim p As PointF = ToDevice(Stitches(i), scale, tx, ty)

            If thisIdx <> lastIdx OrElse thisBreak Then
                IncrementHeat(heat, width, height, CInt(Math.Round(p.X)), CInt(Math.Round(p.Y)), maxHeat)
                last = p : lastIdx = thisIdx
                Continue For
            End If

            Dim dx As Double = p.X - last.X
            Dim dy As Double = p.Y - last.Y
            Dim len As Double = Math.Sqrt(dx * dx + dy * dy)
            Dim steps As Integer = Math.Max(1, CInt(Math.Floor(len)))
            Dim stepx As Double = dx / steps
            Dim stepy As Double = dy / steps
            Dim sx As Double = last.X, sy As Double = last.Y
            For s As Integer = 0 To steps
                IncrementHeat(heat, width, height, CInt(Math.Round(sx)), CInt(Math.Round(sy)), maxHeat)
                sx += stepx : sy += stepy
            Next

            last = p : lastIdx = thisIdx
        Next
    End Sub

    Private Sub IncrementHeat(ByRef heat As Integer(,), width As Integer, height As Integer,
                              x As Integer, y As Integer, ByRef maxHeat As Integer)
        If x < 0 OrElse y < 0 OrElse x >= width OrElse y >= height Then Return
        heat(x, y) += 1
        If heat(x, y) > maxHeat Then maxHeat = heat(x, y)
    End Sub

    Private Function SampleHeat(heat As Integer(,), maxHeat As Integer, x As Integer, y As Integer) As Double
        If heat Is Nothing OrElse maxHeat <= 0 Then Return 0.0
        If x < 0 OrElse y < 0 OrElse x > heat.GetUpperBound(0) OrElse y > heat.GetUpperBound(1) Then Return 0.0
        Return CDbl(heat(x, y)) / CDbl(maxHeat)
    End Function

    ' ---------- Small utils ----------
    Private Shared Sub ScaleToMaxPixels(ByRef w As Integer, ByRef h As Integer, maxPixels As Long)
        If w <= 0 OrElse h <= 0 Then Return
        Dim cur As Long = CLng(w) * CLng(h)
        If cur <= maxPixels Then Return
        Dim scale As Double = Math.Sqrt(maxPixels / CDbl(cur))
        w = Math.Max(1, CInt(Math.Floor(w * scale)))
        h = Math.Max(1, CInt(Math.Floor(h * scale)))
    End Sub

    Private Shared Function ClampInt(v As Integer, lo As Integer, hi As Integer) As Integer
        If v < lo Then Return lo
        If v > hi Then Return hi
        Return v
    End Function
End Class
