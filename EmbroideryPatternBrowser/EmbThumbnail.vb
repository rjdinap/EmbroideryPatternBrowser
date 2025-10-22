Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Linq

' Lightweight thumbnail generator for EmbPattern (hardened)
Public Module EmbThumbnail

    ' === safety caps ===
    Private Const MAX_W As Integer = 3000
    Private Const MAX_H As Integer = 3000
    Private Const MAX_PIXELS As Long = 9_000_000L   ' 3k x 3k

    ' In-memory LRU cache for thumbnails (optional).
    Private ReadOnly _cache As New Dictionary(Of String, CacheEntry)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _lru As New LinkedList(Of String)
    Private ReadOnly _sync As New Object()
    Private _maxCacheItems As Integer = 200 ' tweak for your grid scale

    Private Class CacheEntry
        Public Key As String
        Public Img As Image
        Public Node As LinkedListNode(Of String)
    End Class

    ''' <summary>
    ''' Generate (or fetch cached) thumbnail for a file by path.
    ''' Cache key uses full path + last write time + parameters.
    ''' </summary>
    Public Function GenerateFromFile(filePath As String,
                                     maxWidth As Integer,
                                     maxHeight As Integer,
                                     Optional padding As Integer = 6,
                                     Optional bg As Color = Nothing,
                                     Optional drawBorder As Boolean = False,
                                     Optional showName As Boolean = False) As Image
        Try
            If String.IsNullOrWhiteSpace(filePath) Then Return Nothing
            If bg.IsEmpty Then bg = Color.White

            ' cap request
            Dim wReq As Integer = Clamp(maxWidth, 1, MAX_W)
            Dim hReq As Integer = Clamp(maxHeight, 1, MAX_H)
            If CLng(wReq) * CLng(hReq) > MAX_PIXELS Then ScaleToMaxPixels(wReq, hReq, MAX_PIXELS)

            Dim key As String
            Try
                Dim fi As New FileInfo(filePath)
                key = String.Format("file::{0}|{1:yyyyMMddHHmmss}|{2}x{3}|{4}|{5}|{6}",
                                    fi.FullName, fi.LastWriteTimeUtc, wReq, hReq, padding, drawBorder, showName)
            Catch
                key = String.Format("file::{0}|{1}x{2}|{3}|{4}|{5}", filePath, wReq, hReq, padding, drawBorder, showName)
            End Try

            Dim cached As Image = TryGetFromCache(key)
            If cached IsNot Nothing Then Return cached

            ' Load pattern and render
            Dim pat As EmbPattern = Nothing
            Try
                pat = EmbReaderFactory.LoadPattern(filePath)
            Catch ex As Exception
                Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString)
                Return Nothing
            End Try

            Dim img As Image = Generate(pat, wReq, hReq, padding, bg, drawBorder, showName, pat.GetMetadata(EmbPattern.PROP_NAME))
            AddToCache(key, img)
            Return img

        Catch ex As Exception
            Try : Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString) : Catch : End Try
            Return Nothing
        End Try
    End Function



    ' Generate a thumbnail for an already loaded EmbPattern.
    Public Function Generate(pattern As EmbPattern,
                         maxWidth As Integer,
                         maxHeight As Integer,
                         Optional padding As Integer = 6,
                         Optional bg As Color = Nothing,
                         Optional drawBorder As Boolean = False,
                         Optional showName As Boolean = False,
                         Optional nameText As String = Nothing) As Image
        Try
            If pattern Is Nothing Then Return Nothing
            If bg.IsEmpty Then bg = Color.White

            ' Project setting gate
            Dim useHQ As Boolean = String.Equals(My.Settings.ThumbnailQuality, "High Quality", StringComparison.OrdinalIgnoreCase)

            ' --- cap request ---
            Dim wReq As Integer = Clamp(maxWidth, 1, MAX_W)
            Dim hReq As Integer = Clamp(maxHeight, 1, MAX_H)
            If CLng(wReq) * CLng(hReq) > MAX_PIXELS Then ScaleToMaxPixels(wReq, hReq, MAX_PIXELS)

            ' --- bounds ---
            Dim bounds As RectangleF = pattern.ComputeBounds()
            If bounds.Width <= 0 OrElse bounds.Height <= 0 Then
                ' Fallback: palette swatch box
                Dim w As Integer = Math.Max(64, wReq)
                Dim h As Integer = Math.Max(64, hReq)
                If CLng(w) * CLng(h) > MAX_PIXELS Then ScaleToMaxPixels(w, h, MAX_PIXELS)

                Dim bmp As New Bitmap(w, h, PixelFormat.Format24bppRgb)
                Using g = Graphics.FromImage(bmp)
                    g.SmoothingMode = SmoothingMode.None
                    g.InterpolationMode = InterpolationMode.NearestNeighbor
                    g.CompositingQuality = CompositingQuality.HighSpeed
                    g.Clear(bg)

                    Dim n = Math.Max(1, pattern.ThreadList.Count)
                    Dim sw = Math.Max(1, (w - padding * 2) \ Math.Max(1, Math.Min(n, 8)))
                    Dim x = padding
                    For i = 0 To Math.Min(n, 8) - 1
                        Using br As New SolidBrush(SafeThreadColor(pattern, i))
                            g.FillRectangle(br, x, padding, sw, h - padding * 2)
                        End Using
                        x += sw + 2
                    Next

                    If drawBorder Then g.DrawRectangle(Pens.Gray, 0, 0, w - 1, h - 1)
                End Using
                Return bmp
            End If

            ' --- fit scale (inside padding) ---
            Dim innerW As Double = Math.Max(1, wReq - padding * 2)
            Dim innerH As Double = Math.Max(1, hReq - padding * 2)
            Dim sx As Double = innerW / bounds.Width
            Dim sy As Double = innerH / bounds.Height
            Dim scale As Double = Math.Max(0.01, Math.Min(sx, sy))

            ' Tight final dimensions (not exceeding request)
            Dim wOut As Integer = CInt(Math.Ceiling(bounds.Width * scale)) + padding * 2
            Dim hOut As Integer = CInt(Math.Ceiling(bounds.Height * scale)) + padding * 2
            wOut = Clamp(wOut, 1, wReq)
            hOut = Clamp(hOut, 1, hReq)
            If CLng(wOut) * CLng(hOut) > MAX_PIXELS Then ScaleToMaxPixels(wOut, hOut, MAX_PIXELS)

            ' Base transform at final scale
            Dim tx As Double = padding - (bounds.Left * scale)
            Dim ty As Double = padding - (bounds.Top * scale)

            ' Build color indices (fast)
            Dim colorIdx() As Integer = BuildFastColorIndices(pattern)

            ' ======================== RENDER ========================
            If Not useHQ Then
                ' -------- FAST PATH (original behavior) --------
                Dim bmpOut As New Bitmap(wOut, hOut, PixelFormat.Format24bppRgb)
                Using g As Graphics = Graphics.FromImage(bmpOut)
                    g.SmoothingMode = SmoothingMode.HighSpeed
                    g.InterpolationMode = InterpolationMode.NearestNeighbor
                    g.PixelOffsetMode = PixelOffsetMode.HighSpeed
                    g.CompositingQuality = CompositingQuality.HighSpeed
                    g.Clear(bg)

                    If pattern.Stitches.Count >= 2 Then
                        Dim penW As Single = Math.Max(0.35F / CSng(Math.Max(scale, 0.0001)), 0.1F)
                        Dim last As PointF = ToDevice(pattern.Stitches(0), scale, tx, ty)
                        Dim lastIdx As Integer = ClampIdx(pattern, If(colorIdx.Length > 0, colorIdx(0), 0))

                        Using pen As New Pen(SafeThreadColor(pattern, lastIdx), penW)
                            pen.StartCap = LineCap.Round
                            pen.EndCap = LineCap.Round

                            For i As Integer = 1 To pattern.Stitches.Count - 1
                                Dim thisIdx As Integer = ClampIdx(pattern, If(i < colorIdx.Length, colorIdx(i), lastIdx))
                                Dim hasBreak As Boolean = (i < pattern.BreakBefore.Count) AndAlso pattern.BreakBefore(i)
                                Dim p As PointF = ToDevice(pattern.Stitches(i), scale, tx, ty)

                                If thisIdx <> lastIdx Then
                                    pen.Color = SafeThreadColor(pattern, thisIdx)
                                    last = p : lastIdx = thisIdx
                                    Continue For
                                End If
                                If hasBreak Then
                                    last = p
                                    Continue For
                                End If

                                g.DrawLine(pen, last, p)
                                last = p
                            Next
                        End Using
                    End If

                    ' Label on final
                    If showName Then
                        Dim nm As String = If(String.IsNullOrEmpty(nameText), pattern.GetMetadata(EmbPattern.PROP_NAME), nameText)
                        If Not String.IsNullOrEmpty(nm) Then
                            Using f As New Font("Segoe UI", 7.0F, FontStyle.Regular, GraphicsUnit.Point)
                                Dim rect As New RectangleF(2, hOut - 14, wOut - 4, 12)
                                Using br As New SolidBrush(Color.FromArgb(200, Color.White))
                                    g.FillRectangle(br, rect)
                                End Using
                                Using brT As New SolidBrush(Color.Black)
                                    g.DrawString(nm, f, brT, rect)
                                End Using
                            End Using
                        End If
                    End If

                    If drawBorder Then
                        Using p As New Pen(Color.FromArgb(180, 180, 180))
                            g.DrawRectangle(p, 0, 0, wOut - 1, hOut - 1)
                        End Using
                    End If
                End Using

                Return bmpOut
            Else
                ' -------- HIGH-QUALITY PATH (with tiny-thumb supersampling) --------
                ' Supersample if the smaller side is <= 201 px
                Dim ss As Integer = 1
                If Math.Min(wOut, hOut) <= 201 Then ss = 2

                Dim rw As Integer = wOut * ss
                Dim rh As Integer = hOut * ss
                Dim rScale As Double = scale * ss
                Dim rTx As Double = (tx - padding) * ss + padding
                Dim rTy As Double = (ty - padding) * ss + padding
                If CLng(rw) * CLng(rh) > MAX_PIXELS Then ScaleToMaxPixels(rw, rh, MAX_PIXELS)

                ' Render to scratch
                Dim scratch As New Bitmap(rw, rh, PixelFormat.Format24bppRgb)
                Using g As Graphics = Graphics.FromImage(scratch)
                    If ss > 1 Then
                        g.SmoothingMode = SmoothingMode.AntiAlias
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality
                        g.CompositingQuality = CompositingQuality.HighQuality
                    Else
                        g.SmoothingMode = SmoothingMode.AntiAlias
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality
                        g.CompositingQuality = CompositingQuality.HighQuality
                    End If

                    g.Clear(bg)

                    If pattern.Stitches.Count >= 2 Then
                        ' Slightly higher minimum width at tiny scales so segments don’t vanish
                        Dim penW As Single = Math.Max(0.45F / CSng(Math.Max(rScale, 0.0001)), If(ss > 1, 0.18F, 0.1F))

                        Dim last As PointF = New PointF(
                        CSng(pattern.Stitches(0).X * rScale + rTx),
                        CSng(pattern.Stitches(0).Y * rScale + rTy))
                        Dim lastIdx As Integer = ClampIdx(pattern, If(colorIdx.Length > 0, colorIdx(0), 0))

                        Using pen As New Pen(SafeThreadColor(pattern, lastIdx), penW)
                            pen.StartCap = LineCap.Round
                            pen.EndCap = LineCap.Round

                            For i As Integer = 1 To pattern.Stitches.Count - 1
                                Dim thisIdx As Integer = ClampIdx(pattern, If(i < colorIdx.Length, colorIdx(i), lastIdx))
                                Dim hasBreak As Boolean = (i < pattern.BreakBefore.Count) AndAlso pattern.BreakBefore(i)

                                Dim p As PointF = New PointF(
                                CSng(pattern.Stitches(i).X * rScale + rTx),
                                CSng(pattern.Stitches(i).Y * rScale + rTy))

                                If thisIdx <> lastIdx Then
                                    pen.Color = SafeThreadColor(pattern, thisIdx)
                                    last = p : lastIdx = thisIdx
                                    Continue For
                                End If
                                If hasBreak Then
                                    last = p
                                    Continue For
                                End If

                                g.DrawLine(pen, last, p)
                                last = p
                            Next
                        End Using
                    End If
                End Using

                ' Downscale if supersampled, then draw label/border on final
                Dim bmpOut As Bitmap
                If ss > 1 Then
                    bmpOut = New Bitmap(wOut, hOut, PixelFormat.Format24bppRgb)
                    Using gg As Graphics = Graphics.FromImage(bmpOut)
                        gg.SmoothingMode = SmoothingMode.HighQuality
                        gg.InterpolationMode = InterpolationMode.HighQualityBicubic
                        gg.PixelOffsetMode = PixelOffsetMode.HighQuality
                        gg.CompositingQuality = CompositingQuality.HighQuality
                        gg.Clear(bg)
                        gg.DrawImage(scratch, New Rectangle(0, 0, wOut, hOut))
                    End Using
                    scratch.Dispose()
                Else
                    bmpOut = scratch
                End If

                If showName Then
                    Dim nm As String = If(String.IsNullOrEmpty(nameText), pattern.GetMetadata(EmbPattern.PROP_NAME), nameText)
                    If Not String.IsNullOrEmpty(nm) Then
                        Using g As Graphics = Graphics.FromImage(bmpOut)
                            Using f As New Font("Segoe UI", 7.0F, FontStyle.Regular, GraphicsUnit.Point)
                                Dim rect As New RectangleF(2, bmpOut.Height - 14, bmpOut.Width - 4, 12)
                                Using br As New SolidBrush(Color.FromArgb(200, Color.White))
                                    g.FillRectangle(br, rect)
                                End Using
                                Using brT As New SolidBrush(Color.Black)
                                    g.DrawString(nm, f, brT, rect)
                                End Using
                            End Using
                        End Using
                    End If
                End If

                If drawBorder Then
                    Using p As New Pen(Color.FromArgb(180, 180, 180))
                        Using g As Graphics = Graphics.FromImage(bmpOut)
                            g.DrawRectangle(p, 0, 0, bmpOut.Width - 1, bmpOut.Height - 1)
                        End Using
                    End Using
                End If

                Return bmpOut
            End If
            ' ====================== /RENDER ======================

        Catch ex As Exception
            Try : Form1.Status("Error: " & ex.Message, ex.StackTrace.ToString) : Catch : End Try
            Return Nothing
        End Try
    End Function





    ' -------- cache helpers --------
    Private Function TryGetFromCache(key As String) As Image
        If key Is Nothing Then Return Nothing
        SyncLock _sync
            Dim e As CacheEntry = Nothing
            If _cache.TryGetValue(key, e) AndAlso e IsNot Nothing AndAlso e.Img IsNot Nothing Then
                ' move to front
                If e.Node IsNot Nothing Then
                    _lru.Remove(e.Node)
                    _lru.AddFirst(e.Node)
                    e.Node = _lru.First
                End If
                ' return a safe copy
                Try
                    Return CType(e.Img.Clone(), Image)
                Catch
                    ' If clone fails (disposed/etc.), drop this entry
                    SafeEvict(key)
                End Try
            End If
        End SyncLock
        Return Nothing
    End Function

    Private Sub AddToCache(key As String, img As Image)
        If key Is Nothing OrElse img Is Nothing Then Return
        SyncLock _sync
            If _cache.ContainsKey(key) Then Return
            ' evict if needed
            While _cache.Count >= _maxCacheItems AndAlso _lru.Count > 0
                Dim tailKey = _lru.Last.Value
                SafeEvict(tailKey)
            End While
            Dim node = _lru.AddFirst(key)
            _cache(key) = New CacheEntry With {.Key = key, .Img = CType(img.Clone(), Image), .Node = node}
        End SyncLock
    End Sub

    Private Sub SafeEvict(key As String)
        Try
            If Not _cache.ContainsKey(key) Then Return
            Dim entry = _cache(key)
            If entry IsNot Nothing Then
                Try : entry.Img?.Dispose() : Catch : End Try
                If entry.Node IsNot Nothing Then
                    Try : _lru.Remove(entry.Node) : Catch : End Try
                End If
            End If
            _cache.Remove(key)
        Catch
        End Try
    End Sub

    ' -------- tiny helpers --------
    Private Function ToDevice(p As PointF, scale As Double, tx As Double, ty As Double) As PointF
        Return New PointF(CSng(p.X * scale + tx), CSng(p.Y * scale + ty))
    End Function

    Private Function ClampIdx(pattern As EmbPattern, i As Integer) As Integer
        If pattern.ThreadList Is Nothing OrElse pattern.ThreadList.Count = 0 Then Return 0
        If i < 0 Then Return 0
        If i >= pattern.ThreadList.Count Then Return pattern.ThreadList.Count - 1
        Return i
    End Function

    Private Function SafeThreadColor(pattern As EmbPattern, idx As Integer) As Color
        If idx >= 0 AndAlso idx < pattern.ThreadList.Count Then
            Dim t = pattern.ThreadList(idx)
            If t IsNot Nothing Then Return t.ColorValue
        End If
        Return Color.Black
    End Function

    ' Fast color-index builder: use decoded indices if they match; else fallback to 0.
    Private Function BuildFastColorIndices(pattern As EmbPattern) As Integer()
        Dim n As Integer = pattern.Stitches.Count
        If n = 0 Then Return New Integer() {}
        If pattern.StitchThreadIndices IsNot Nothing AndAlso pattern.StitchThreadIndices.Count = n Then
            Return pattern.StitchThreadIndices.ToArray()
        End If
        Dim a(n - 1) As Integer
        Return a
    End Function

    Private Function Clamp(v As Integer, lo As Integer, hi As Integer) As Integer
        If v < lo Then Return lo
        If v > hi Then Return hi
        Return v
    End Function

    Private Sub ScaleToMaxPixels(ByRef w As Integer, ByRef h As Integer, maxPixels As Long)
        If w <= 0 OrElse h <= 0 Then Return
        Dim cur As Long = CLng(w) * CLng(h)
        If cur <= maxPixels Then Return
        Dim s As Double = Math.Sqrt(maxPixels / CDbl(cur))
        w = Math.Max(1, CInt(Math.Floor(w * s)))
        h = Math.Max(1, CInt(Math.Floor(h * s)))
    End Sub


    ' ===================== public cache utilities =====================

    ''' <summary>Remove all cached thumbnails of a given size (w×h).</summary>
    Public Sub ThumbCacheInvalidateBySize(width As Integer, height As Integer)
        Dim needle As String = "|" & width & "x" & height & "|"
        SyncLock _sync
            Dim keys = _cache.Keys.Where(Function(k) k.IndexOf(needle, StringComparison.Ordinal) >= 0).ToList()
            For Each k In keys
                SafeEvict(k)
            Next
        End SyncLock
    End Sub

    ''' <summary>Clear the entire thumbnail cache.</summary>
    Public Sub ThumbCacheInvalidateAll()
        SyncLock _sync
            Dim keys = _cache.Keys.ToList()
            For Each k In keys
                SafeEvict(k)
            Next
            _lru.Clear()
        End SyncLock
    End Sub

    ''' <summary>Warm (pre-generate) the cache for the provided file paths at size w×h.</summary>
    Public Sub ThumbCacheWarmAsync(paths As IEnumerable(Of String),
                               width As Integer, height As Integer,
                               Optional padding As Integer = 6,
                               Optional drawBorder As Boolean = True,
                               Optional showName As Boolean = False)
        If paths Is Nothing Then Return
        Threading.Tasks.Task.Run(
        Sub()
            For Each p In paths
                Try
                    Dim img As Image = GenerateFromFile(p, width, height, padding, Color.White, drawBorder, showName)
                    img?.Dispose() ' AddToCache inside GenerateFromFile; dispose local
                Catch
                    ' best-effort warm
                End Try
            Next
        End Sub)
    End Sub

    ' =================== end public cache utilities ===================





End Module
