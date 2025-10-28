Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Windows.Forms

Public NotInheritable Class PrintFunctions

    ' -------- Public entry point --------
    Public Shared Sub ShowPrintPane(owner As Form, grid As Control)
        Dim g = TryCast(grid, VirtualThumbGrid)
        Dim dlg As New PrintPane(owner, g)
        dlg.Show(owner) ' modeless; change to ShowDialog(owner) if you prefer modal
    End Sub



    Private Class PrintPane
        Inherits Form

        Private ReadOnly _owner As Form
        Private ReadOnly _grid As VirtualThumbGrid

        ' UI
        Private _preview As PrintPreviewControl
        Private _btnPrint As Button
        Private _btnPrinter As Button
        Private _rThumb As RadioButton
        Private _rFull As RadioButton
        Private _chkThumbDetails As CheckBox
        Private _nudThumbsPerRow As NumericUpDown
        Private _lblNoSelection As Label

        ' Printing state
        Private ReadOnly _doc As PrintDocument = New PrintDocument()

        ' Current selection resolved to real file paths
        Private _selectedFiles As New List(Of String)()

        Private _scroll As Panel
        Private _lblPageInfo As ToolStripLabel

        Private _printMode As String = "thumb"   ' "thumb" or "full"
        Private _thumbIndex As Integer = 0       ' next item to render in thumbnail mode
        Private _fullIndex As Integer = 0        ' next item to render in full-size mode

        Public Sub New(owner As Form, grid As VirtualThumbGrid)
            _owner = owner
            _grid = grid
            Me.Text = "Print & Preview"
            Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.MinimizeBox = False
            Me.MaximizeBox = True
            Me.Size = New Size(980, 700)
            Me.KeyPreview = True ' Enable form to receive key events first

            BuildUi()
            HookEvents()
            RequerySelectionAndRebuild()
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            ' Capture Page Up/Down and arrow keys for page navigation
            If _preview IsNot Nothing AndAlso _preview.Document IsNot Nothing Then
                If keyData = Keys.PageDown OrElse keyData = Keys.Down Then
                    If _preview.StartPage < _preview.Document.PrinterSettings.MaximumPage - 1 Then
                        _preview.StartPage += 1
                        UpdatePageLabel()
                        Return True
                    End If
                ElseIf keyData = Keys.PageUp OrElse keyData = Keys.Up Then
                    If _preview.StartPage > 0 Then
                        _preview.StartPage -= 1
                        UpdatePageLabel()
                        Return True
                    End If
                End If
            End If
            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        Private Sub BuildUi()
            SuspendLayout()
            ' Right: preview host with scrollbars + a tiny zoom strip
            Dim rightHost As New Panel() With {.Dock = DockStyle.Fill}
            Controls.Add(rightHost)

            ' Left: options
            Dim left As New Panel() With {.Dock = DockStyle.Left, .Width = 280, .Padding = New Padding(10)}
            Controls.Add(left)

            Dim grpMode As New GroupBox() With {.Text = "Mode", .Dock = DockStyle.Top, .Height = 130}
            left.Controls.Add(grpMode)

            _rThumb = New RadioButton() With {.Text = "Thumbnail mode", .Checked = True, .Left = 12, .Top = 24, .AutoSize = True}
            _chkThumbDetails = New CheckBox() With {.Text = "Show details under image", .Left = 32, .Top = 50, .AutoSize = True}
            Dim lblPerRow As New Label() With {.Text = "Thumbnails per row:", .Left = 32, .Top = 78, .AutoSize = True}
            _nudThumbsPerRow = New NumericUpDown() With {.Left = 170, .Top = 74, .Width = 60, .Minimum = 1, .Maximum = 6, .Value = 3}
            _rFull = New RadioButton() With {.Text = "Full-size mode", .Left = 12, .Top = 104, .AutoSize = True}

            grpMode.Controls.AddRange({_rThumb, _chkThumbDetails, lblPerRow, _nudThumbsPerRow, _rFull})

            Dim grpPrinter As New GroupBox() With {.Text = "Printer", .Dock = DockStyle.Top, .Height = 70, .Padding = New Padding(8)}
            left.Controls.Add(grpPrinter)

            _btnPrinter = New Button() With {.Text = "Printer…", .Width = 100, .Height = 28, .Left = 8, .Top = 26}
            grpPrinter.Controls.Add(_btnPrinter)

            _btnPrint = New Button() With {.Text = "Print", .Dock = DockStyle.Bottom, .Height = 36}
            left.Controls.Add(_btnPrint)



            Dim ts As New ToolStrip() With {.GripStyle = ToolStripGripStyle.Hidden, .Dock = DockStyle.Top}

            ' Page navigation controls only
            Dim btnPrevPage As New ToolStripButton("◄")
            Dim btnNextPage As New ToolStripButton("►")
            _lblPageInfo = New ToolStripLabel("Page 1 of 1")
            btnPrevPage.ToolTipText = "Previous Page (Page Up)"
            btnNextPage.ToolTipText = "Next Page (Page Down)"

            AddHandler btnPrevPage.Click, Sub()
                                              If _preview.StartPage > 0 Then
                                                  _preview.StartPage -= 1
                                                  UpdatePageLabel()
                                              End If
                                          End Sub
            AddHandler btnNextPage.Click, Sub()
                                              If _preview.StartPage < GetTotalPages() - 1 Then
                                                  _preview.StartPage += 1
                                                  UpdatePageLabel()
                                              End If
                                          End Sub

            ts.Items.AddRange({btnPrevPage, btnNextPage, _lblPageInfo})
            rightHost.Controls.Add(ts)

            _scroll = New Panel() With {.Dock = DockStyle.Fill, .AutoScroll = True}
            rightHost.Controls.Add(_scroll)

            _preview = New PrintPreviewControl() With {
                .Location = New Point(0, 0),   ' not Dock=Fill; we want outer panel scrollbars
                .AutoZoom = False,
                .Zoom = 1.0R
            }
            _scroll.Controls.Add(_preview)

            _lblNoSelection = New Label() With {
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleCenter,
                .ForeColor = Color.DimGray,
                .Font = New Font(Me.Font.FontFamily, 12.0F, FontStyle.Italic),
                .Visible = False,
                .Text = "No items selected in the grid." & Environment.NewLine &
                        "Select 1 or more items to print."
            }
            rightHost.Controls.Add(_lblNoSelection) ' added after scroll so it overlays only when visible

            AddHandler _scroll.Resize, Sub() UpdatePreviewViewport()
            ResumeLayout()
        End Sub


        Private Sub UpdatePreviewViewport()
            If _preview Is Nothing Then Return
            Dim pg = GetPagePixelSize()
            ' Size the control to the (scaled) page so the outer panel shows scrollbars when needed
            _preview.Size = New Size(CInt(pg.Width * _preview.Zoom), CInt(pg.Height * _preview.Zoom))
            _preview.InvalidatePreview()
        End Sub



        Private Function GetPagePixelSize() As Size
            Try
                Dim ps = _doc.DefaultPageSettings
                Dim dpiX As Single, dpiY As Single
                Using gr = Me.CreateGraphics()
                    dpiX = gr.DpiX : dpiY = gr.DpiY
                End Using
                Dim wPx = CInt((ps.Bounds.Width / 100.0F) * dpiX)
                Dim hPx = CInt((ps.Bounds.Height / 100.0F) * dpiY)
                Return New Size(Math.Max(1, wPx + 20), Math.Max(1, hPx + 20)) ' tiny extra so border shows
            Catch
                Return New Size(850, 1100)
            End Try
        End Function



        Private Sub HookEvents()
            AddHandler _rThumb.CheckedChanged, Sub()
                                                   RebuildPreview()
                                                   _preview.Focus()
                                               End Sub
            AddHandler _rFull.CheckedChanged, Sub()
                                                  RebuildPreview()
                                                  _preview.Focus()
                                              End Sub
            AddHandler _chkThumbDetails.CheckedChanged, Sub() RebuildPreview()
            AddHandler _nudThumbsPerRow.ValueChanged, Sub() RebuildPreview()

            AddHandler _btnPrinter.Click, AddressOf OnPrinter
            AddHandler _btnPrint.Click, AddressOf OnPrint

            AddHandler _doc.BeginPrint, AddressOf OnBeginPrint
            AddHandler _doc.PrintPage, AddressOf OnPrintPage
        End Sub




        Private Sub UpdatePageLabel()
            If _lblPageInfo IsNot Nothing AndAlso _preview IsNot Nothing Then
                Dim totalPages As Integer = GetTotalPages()
                _lblPageInfo.Text = $"Page {_preview.StartPage + 1} of {totalPages}"
            End If
        End Sub



        Private Function GetTotalPages() As Integer
            If _selectedFiles.Count = 0 Then Return 1

            If _printMode = "thumb" Then
                Dim perRow As Integer = CInt(_nudThumbsPerRow.Value)
                Dim thumbsPerPage As Integer = perRow * 4
                Return Math.Max(1, CInt(Math.Ceiling(_selectedFiles.Count / CDbl(thumbsPerPage))))
            Else
                Return _selectedFiles.Count
            End If
        End Function


        Private Sub OnPrinter(sender As Object, e As EventArgs)
            Using dlg As New PrintDialog()
                dlg.UseEXDialog = True
                dlg.Document = _doc
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    _doc.PrinterSettings = dlg.PrinterSettings
                    RebuildPreview()
                    UpdatePreviewViewport()
                End If
            End Using
        End Sub

        Private Sub OnPrint(sender As Object, e As EventArgs)
            If _selectedFiles.Count = 0 Then
                MessageBox.Show(Me, "Select one or more images in the main grid first.", "Nothing to print",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            Try
                _doc.Print()
            Catch ex As Exception
                Try : Logger.Error("Print error: " & ex.Message, ex.ToString()) : Catch : End Try
                MessageBox.Show(Me, "Printer reported an error: " & ex.Message, "Print", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
        End Sub



        ' Build selection → file list (from VirtualThumbGrid)
        Private Sub RequerySelectionAndRebuild()
            _selectedFiles = GetSelectedFullPathsFromGrid(_grid)
            _lblNoSelection.Visible = (_selectedFiles.Count = 0)
            If _lblNoSelection.Visible Then
                _lblNoSelection.BringToFront()
            Else
                _lblNoSelection.SendToBack()
            End If
            RebuildPreview()
        End Sub




        Private Sub RebuildPreview()
            If _selectedFiles.Count = 0 Then
                _preview.Document = Nothing
                _preview.Invalidate()
                Return
            End If

            _printMode = If(_rThumb.Checked, "thumb", "full")
            _thumbIndex = 0
            _fullIndex = 0

            _preview.Document = _doc
            _preview.InvalidatePreview()
            UpdatePreviewViewport()
            UpdatePageLabel()
            _preview.BackColor = Color.White
        End Sub




        Private Sub OnBeginPrint(sender As Object, e As PrintEventArgs)
            _thumbIndex = 0
            _fullIndex = 0
        End Sub



        Private Sub OnPrintPage(sender As Object, e As PrintPageEventArgs)
            ' Work in pixels for layout/drawing
            e.Graphics.PageUnit = GraphicsUnit.Pixel

            ' Convert printable margin box (1/100") → pixels
            Dim dx As Single = e.Graphics.DpiX, dy As Single = e.Graphics.DpiY
            Dim pxMargin As New Rectangle(
        CInt(e.MarginBounds.Left / 100.0F * dx),
        CInt(e.MarginBounds.Top / 100.0F * dy),
        CInt(e.MarginBounds.Width / 100.0F * dx),
        CInt(e.MarginBounds.Height / 100.0F * dy)
    )

            ' Draw everything relative to (0,0) inside the margin
            e.Graphics.TranslateTransform(pxMargin.Left, pxMargin.Top)
            Dim inner As New Rectangle(0, 0, pxMargin.Width, pxMargin.Height)

            If _printMode = "thumb" Then
                Dim drawn As Integer = RenderThumbnailPage(e.Graphics, inner)
                _thumbIndex += drawn
                e.HasMorePages = (_thumbIndex < _selectedFiles.Count)
            Else
                If _fullIndex < _selectedFiles.Count Then
                    DrawFullSizeLayout(e.Graphics, inner, _selectedFiles(_fullIndex))
                    _fullIndex += 1
                    e.HasMorePages = (_fullIndex < _selectedFiles.Count)
                Else
                    e.HasMorePages = False
                End If
            End If
        End Sub



        ' ---------------- THUMBNAILS MODE ----------------
        Private Sub BuildPages_Thumbnails()
            ' Intentionally empty – thumbnails now render per page via RenderThumbnailPage()
        End Sub

        Private Function RenderThumbnailPage(g As Graphics, bounds As Rectangle) As Integer
            Dim perRow As Integer = CInt(_nudThumbsPerRow.Value)

            Dim ts As Size = GetThumbSize()
            ' Scale thumbnail size from screen DPI (96) to printer DPI
            Dim dpiScale As Single = g.DpiX / 96.0F

            ' Calculate base thumbnail size
            Dim baseCellPad As Integer = CInt(12 * dpiScale)
            Dim baseThumbW As Integer = CInt(ts.Width * dpiScale)
            Dim baseThumbH As Integer = CInt(ts.Height * dpiScale)

            ' Calculate how wide each cell would be with the desired number per row
            Dim availableWidth As Integer = bounds.Width
            Dim maxCellW As Integer = availableWidth \ perRow
            Dim maxThumbW As Integer = maxCellW - (baseCellPad * 2)

            ' Use the smaller of the calculated size or the desired size
            Dim thumbW As Integer = Math.Min(baseThumbW, Math.Max(64, maxThumbW))
            Dim thumbH As Integer = CInt(thumbW * (ts.Height / CDbl(ts.Width))) ' maintain aspect ratio

            Dim cellPad As Integer = baseCellPad
            Dim cellW As Integer = thumbW + cellPad * 2
            ' AFTER (more room for 12pt title + 10.5pt meta, 2–3 lines)
            Dim detailsExtra As Integer
            If _chkThumbDetails.Checked Then
                Using titleFont As New Font(Me.Font.FontFamily, 12.0F, FontStyle.Bold, GraphicsUnit.Point),
          metaFont As New Font(Me.Font.FontFamily, 10.5F, FontStyle.Regular, GraphicsUnit.Point)
                    Dim sf As New StringFormat(StringFormatFlags.LineLimit)
                    ' Allow for 3 lines of title (full path can be long) + 2 lines of meta
                    Dim titleH = CInt(Math.Ceiling(g.MeasureString("Wy", titleFont, 600, sf).Height)) * 3
                    Dim metaH = CInt(Math.Ceiling(g.MeasureString("Wy", metaFont, 600, sf).Height)) * 2
                    detailsExtra = titleH + metaH + 18
                End Using
            Else
                detailsExtra = 18
            End If
            Dim cellH As Integer = thumbH + detailsExtra + cellPad * 2

            ' Always use the requested number of columns
            Dim cols As Integer = perRow
            Dim rows As Integer = Math.Max(1, bounds.Height \ cellH)
            Dim slots As Integer = cols * rows

            Dim drawn As Integer = 0
            Dim x As Integer = bounds.Left
            Dim y As Integer = bounds.Top

            Dim k As Integer = _thumbIndex
            For r = 0 To rows - 1
                x = bounds.Left
                For c = 0 To cols - 1
                    If k >= _selectedFiles.Count Then Exit For
                    Dim fp As String = _selectedFiles(k)
                    Dim cell As New Rectangle(x, y, cellW, cellH)
                    DrawThumbnailCell(g, cell, fp, thumbW, thumbH, _chkThumbDetails.Checked)
                    x += cellW
                    k += 1
                    drawn += 1
                Next
                y += cellH
                If k >= _selectedFiles.Count Then Exit For
            Next

            Return drawn ' number of items rendered on this page
        End Function



        Private Sub DrawThumbnailCell(g As Graphics, cell As Rectangle, filePath As String, thumbW As Integer, thumbH As Integer, showDetails As Boolean)
            ' where the image will go, centered within the cell
            Dim imgRect As New Rectangle(cell.Left + (cell.Width - thumbW) \ 2, cell.Top + 8, thumbW, thumbH)

            ' draw the thumbnail
            Try
                Dim img As Image = EmbThumbnail.GenerateFromFile(
                    ImageDriveManager.Resolve(filePath),
                    maxWidth:=thumbW,
                    maxHeight:=thumbH,
                    padding:=6,
                    bg:=Color.White,
                    drawBorder:=True,
                    showName:=False)
                ' Scale the generated image to fill the full imgRect (thumbW x thumbH)
                DrawFit(g, img, imgRect)
                img.Dispose()
            Catch
                ' ignore; we still draw caption area under where the image would be
            End Try

            ' ----- CAPTION UNDER IMAGE (anchored to image box) -----
            Dim captionTop As Integer = imgRect.Bottom + 6
            Dim captionRect As New Rectangle(imgRect.Left, captionTop, imgRect.Width, Math.Max(0, cell.Bottom - captionTop - 6))
            Dim centerFlags As TextFormatFlags = TextFormatFlags.NoPrefix Or TextFormatFlags.WordBreak Or TextFormatFlags.HorizontalCenter

            ' For show details mode, use FULL PATH; otherwise just filename
            Dim name As String = If(showDetails, filePath, SafeFileName(filePath))

            If showDetails Then
                Using titleFont As New Font(Me.Font.FontFamily, 10.0F, FontStyle.Regular, GraphicsUnit.Point),
              metaFont As New Font(Me.Font.FontFamily, 10.5F, FontStyle.Regular, GraphicsUnit.Point)

                    ' draw filename (bold, centered)
                    Dim sf As New StringFormat(StringFormatFlags.LineLimit)
                    sf.Trimming = StringTrimming.EllipsisWord
                    sf.Alignment = StringAlignment.Center
                    sf.LineAlignment = StringAlignment.Near
                    Dim titleSize As SizeF = g.MeasureString(name, titleFont, captionRect.Width, sf)
                    Dim titleHeight As Integer = Math.Min(captionRect.Height, Math.Max(CInt(Math.Ceiling(titleSize.Height)), titleFont.Height))
                    Dim titleRect As New Rectangle(captionRect.Left, captionRect.Top, captionRect.Width, Math.Min(titleHeight, captionRect.Height))
                    DrawStringWrap(g, name, titleFont, titleRect, center:=True)

                    ' draw meta just below
                    Dim metaText As String = ""
                    Try : metaText = GetPatternFacts(filePath) : Catch : End Try

                    Dim metaRect As New Rectangle(captionRect.Left, titleRect.Bottom + 2, captionRect.Width,
                                          Math.Max(0, captionRect.Bottom - (titleRect.Bottom + 2)))
                    If metaRect.Height > 0 AndAlso metaText.Length > 0 Then
                        DrawStringWrap(g, metaText, metaFont, metaRect, center:=True)
                    End If
                End Using
            Else
                ' draw nothing
            End If

            ' (optional) debug cell boundary to verify layout:
            ' g.DrawRectangle(Pens.Gainsboro, cell)
        End Sub




        Private Shared Function SafeFileName(p As String) As String
            If String.IsNullOrEmpty(p) Then Return ""
            Try
                Return Path.GetFileName(p)
            Catch
                ' If it's a virtual path (zip://…|…), just show the tail after the last separator we recognize
                Dim s As String = p.Replace("\", "/")
                Dim barIx = s.LastIndexOf("|"c)
                If barIx >= 0 AndAlso barIx < s.Length - 1 Then Return s.Substring(barIx + 1)
                Dim slashIx = s.LastIndexOf("/"c)
                If slashIx >= 0 AndAlso slashIx < s.Length - 1 Then Return s.Substring(slashIx + 1)
                Return p ' last resort: show raw
            End Try
        End Function



        ' ---------------- FULL-SIZE MODE ----------------

        Private Sub DrawFullSizeLayout(g As Graphics, bounds As Rectangle, filePath As String)
            ' Minimal padding - printer already has margins
            Dim pad As Integer = 4
            Dim headerRect As New Rectangle(bounds.Left + pad, bounds.Top + pad, bounds.Width - pad * 2, 80)
            Dim contentRect As New Rectangle(bounds.Left + pad, headerRect.Bottom + 8, bounds.Width - pad * 2, bounds.Height - headerRect.Bottom - pad - 8)

            ' nice quality
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality

            ' resolve once
            Dim resolved As String = filePath
            Try : resolved = ImageDriveManager.Resolve(filePath) : Catch : End Try

            ' ----- Header (full path only) -----
            Using headerFont As New Font(Me.Font.FontFamily, 10.0F, FontStyle.Bold)
                ' Use word wrap to show full path
                Using sf As New StringFormat()
                    sf.Trimming = StringTrimming.None
                    sf.FormatFlags = StringFormatFlags.LineLimit
                    sf.Alignment = StringAlignment.Near
                    sf.LineAlignment = StringAlignment.Near
                    g.DrawString(resolved, headerFont, Brushes.Black, RectangleF.op_Implicit(headerRect), sf)
                End Using
            End Using

            ' load pattern once to determine if we have thread list
            Dim pat As EmbPattern = Nothing
            Try : pat = EmbReaderFactory.LoadPattern(resolved) : Catch : pat = Nothing : End Try

            ' ----- Columns: adjust layout based on whether we have a thread list -----
            Dim colGap As Integer = 24
            Dim leftW As Integer
            Dim leftRect As Rectangle
            Dim rightRect As Rectangle
            Dim hasThreadList As Boolean = (pat IsNot Nothing AndAlso pat.ThreadList IsNot Nothing AndAlso pat.ThreadList.Count > 0)

            If hasThreadList Then
                ' 70% for image, 30% for thread list (more room for image)
                leftW = CInt(contentRect.Width * 0.7)
                leftRect = New Rectangle(contentRect.Left, contentRect.Top, leftW, contentRect.Height)
                ' Fixed: rightRect should extend to the right edge of contentRect
                rightRect = New Rectangle(leftRect.Right + colGap, contentRect.Top,
                                         contentRect.Right - (leftRect.Right + colGap), contentRect.Height)
            Else
                ' No thread list - use full width for image
                leftRect = contentRect
                rightRect = Rectangle.Empty
            End If

            ' ----- Left: image (pattern or big thumbnail) + CAPTION UNDER IMAGE -----
            Dim drewImg As Boolean = False
            Dim img As Image = Nothing

            Try
                If pat IsNot Nothing Then
                    img = pat.CreateImageFromPattern(w:=leftRect.Width, h:=leftRect.Height, margin:=20, densityShading:=False)
                    drewImg = True
                End If
            Catch
                img = Nothing
                drewImg = False
            End Try

            If Not drewImg Then
                Try
                    img = EmbThumbnail.GenerateFromFile(resolved,
                                                maxWidth:=leftRect.Width,
                                                maxHeight:=leftRect.Height,
                                                padding:=12,
                                                bg:=Color.White,
                                                drawBorder:=True,
                                                showName:=False)
                Catch
                    img = Nothing
                End Try
            End If

            Dim usedImgRect As Rectangle = Rectangle.Empty
            If img IsNot Nothing Then
                usedImgRect = FitRect(img, leftRect)
                g.DrawImage(img, usedImgRect)
                img.Dispose()
            End If

            ' caption block directly under the image box (or at top if image missing)
            Dim capTop As Integer = If(usedImgRect.IsEmpty, leftRect.Top, Math.Min(leftRect.Bottom - 1, usedImgRect.Bottom + 8))
            Dim captionRect As New Rectangle(leftRect.Left, capTop, leftRect.Width, Math.Max(0, leftRect.Bottom - capTop))

            Using titleFont As New Font(Me.Font.FontFamily, 11.5F, FontStyle.Bold),
          metaFont As New Font(Me.Font.FontFamily, 10.5F, FontStyle.Regular)
                ' Line 1: Filename
                'Dim nameLine As String = SafeFileName(resolved) ' switched to filepath
                Dim sf As New StringFormat(StringFormatFlags.LineLimit)
                Dim nameSize As SizeF = g.MeasureString(filePath, titleFont, captionRect.Width, sf)
                Dim nameH As Integer = CInt(Math.Ceiling(nameSize.Height))
                Dim nameRect As New Rectangle(captionRect.Left, captionRect.Top, captionRect.Width, nameH)
                DrawStringWrap(g, filePath, titleFont, nameRect, center:=False)

                ' Get size and stitch count for lines 2 and 3
                Dim sizeText As String = ""
                Dim stitchText As String = ""
                Try
                    Dim fi As New FileInfo(resolved)
                    sizeText = "Size: " & FormatBytes(fi.Length)
                Catch
                End Try

                ' Use the pat variable from outer scope (don't redeclare)
                Try
                    If pat IsNot Nothing Then
                        stitchText = "Stitches: " & pat.Stitches.Count.ToString("N0")
                    End If
                Catch
                End Try

                ' Line 2: Size
                Dim lineH As Integer = CInt(Math.Ceiling(g.MeasureString("Wy", metaFont, captionRect.Width, sf).Height))
                Dim currentY As Integer = nameRect.Bottom + 4
                If Not String.IsNullOrEmpty(sizeText) AndAlso currentY + lineH <= captionRect.Bottom Then
                    Dim sizeRect As New Rectangle(captionRect.Left, currentY, captionRect.Width, lineH)
                    DrawStringWrap(g, sizeText, metaFont, sizeRect, center:=False)
                    currentY += lineH + 2
                End If

                ' Line 3: Stitches
                If Not String.IsNullOrEmpty(stitchText) AndAlso currentY + lineH <= captionRect.Bottom Then
                    Dim stitchRect As New Rectangle(captionRect.Left, currentY, captionRect.Width, lineH)
                    DrawStringWrap(g, stitchText, metaFont, stitchRect, center:=False)
                End If
            End Using

            ' ----- Right: thread list (if pattern available) -----
            If hasThreadList AndAlso Not rightRect.IsEmpty Then
                ' Debug: Draw a border around the thread list area to verify size
                g.DrawRectangle(Pens.LightGray, rightRect)

                ' Ensure minimum width for thread list
                If rightRect.Width < 100 Then
                    ' Area too narrow, skip thread list
                    Using errFont As New Font(Me.Font.FontFamily, 9.0F, FontStyle.Italic)
                        DrawStringWrap(g, $"(area too small: {rightRect.Width}px wide)", errFont, rightRect, center:=False)
                    End Using
                Else
                    Try
                        DrawThreadList(g, rightRect, pat)
                    Catch ex As Exception
                        Using errFont As New Font(Me.Font.FontFamily, 9.0F, FontStyle.Italic)
                            DrawStringWrap(g, "(Error: " & ex.Message & ")", errFont, rightRect, center:=False)
                        End Using
                    End Try
                End If
            End If
        End Sub




        ' Return the destination rectangle that fits the image inside dest while preserving aspect.
        Private Shared Function FitRect(img As Image, dest As Rectangle) As Rectangle
            If img Is Nothing OrElse dest.Width <= 0 OrElse dest.Height <= 0 Then Return Rectangle.Empty
            Dim sx As Double = dest.Width / CDbl(img.Width)
            Dim sy As Double = dest.Height / CDbl(img.Height)
            Dim k As Double = Math.Min(1.0, Math.Min(sx, sy))
            Dim w As Integer = Math.Max(1, CInt(Math.Floor(img.Width * k)))
            Dim h As Integer = Math.Max(1, CInt(Math.Floor(img.Height * k)))
            ' Top-left align instead of center
            Dim x As Integer = dest.Left
            Dim y As Integer = dest.Top
            Return New Rectangle(x, y, w, h)
        End Function



        Private Shared Sub DrawFit(g As Graphics, img As Image, dest As Rectangle)
            If img Is Nothing OrElse dest.Width <= 0 OrElse dest.Height <= 0 Then Return
            Dim scale As Double = Math.Min(dest.Width / CDbl(img.Width), dest.Height / CDbl(img.Height))
            scale = Math.Min(1.0, scale)
            Dim w As Integer = Math.Max(1, CInt(Math.Floor(img.Width * scale)))
            Dim h As Integer = Math.Max(1, CInt(Math.Floor(img.Height * scale)))
            Dim x As Integer = dest.Left + (dest.Width - w) \ 2
            Dim y As Integer = dest.Top + (dest.Height - h) \ 2
            g.DrawImage(img, New Rectangle(x, y, w, h))
        End Sub

        Private Sub DrawThreadList(g As Graphics, area As Rectangle, pat As EmbPattern)
            If pat Is Nothing Then
                DrawStringWrap(g, "(no thread list)", Me.Font, area, center:=False)
                Return
            End If

            ' Build per-thread counts
            Dim n As Integer = Math.Min(pat.Stitches.Count, pat.StitchThreadIndices.Count)
            Dim counts As New Dictionary(Of Integer, Integer)
            For i = 0 To n - 1
                Dim t = pat.StitchThreadIndices(i)
                counts(t) = If(counts.ContainsKey(t), counts(t) + 1, 1)
            Next

            ' Scale based on printer DPI but keep text reasonable
            Dim dpiScale As Single = g.DpiX / 96.0F
            Dim y As Integer = area.Top
            Dim rowH As Integer = CInt(28 * dpiScale)
            Dim swatch As Integer = CInt(22 * dpiScale)
            Dim gap As Integer = CInt(4 * dpiScale)
            Dim textLeft As Integer = area.Left + swatch + CInt(12 * dpiScale)

            Using penBorder As New Pen(Color.Gray),
                  fnt As New Font(Me.Font.FontFamily, 11.0F, FontStyle.Regular, GraphicsUnit.Point)
                For t = 0 To Math.Max(0, pat.ThreadList.Count) - 1
                    If y + rowH > area.Bottom Then Exit For
                    Dim c As Color = If(t < pat.ThreadList.Count, pat.ThreadList(t).ColorValue, Color.Black)
                    Using br As New SolidBrush(c)
                        g.FillRectangle(br, area.Left, y + gap, swatch, swatch)
                    End Using
                    g.DrawRectangle(penBorder, area.Left, y + gap, swatch, swatch)

                    Dim v As Integer = 0
                    counts.TryGetValue(t, v)
                    Dim txt As String = $"{v:N0} sts"
                    Dim rowRect As New Rectangle(textLeft, y, area.Right - textLeft, rowH)
                    DrawStringWrap(g, txt, fnt, rowRect, center:=False)
                    y += rowH
                Next
            End Using
        End Sub

        Private Function GetPatternFacts(filePath As String, Optional includeFilename As Boolean = False, Optional includeFullPath As Boolean = False) As String
            Dim resolved As String = ""
            Try
                resolved = ImageDriveManager.Resolve(filePath)
            Catch
                resolved = filePath ' fall back to original if Resolve throws
            End Try

            Dim displayName As String = If(includeFullPath, resolved, SafeFileName(resolved))
            Dim sizeText As String = ""
            Try
                Dim fi As New FileInfo(resolved)
                sizeText = FormatBytes(fi.Length)
            Catch
            End Try

            Dim pat As EmbPattern = Nothing
            Dim stitchCount As Integer = 0
            Dim inches As String = ""
            Try
                pat = EmbReaderFactory.LoadPattern(resolved)
                stitchCount = pat.Stitches.Count
                inches = TryGetSizeInches(pat)
            Catch
            End Try

            Dim facts As String =
                If(includeFilename, displayName & "    ", "") &
                $"Size: {sizeText}    Stitches: {stitchCount:N0}" &
                If(String.IsNullOrEmpty(inches), "", $"    Pattern: {inches}")
            Return facts
        End Function

        Private Function TryGetSizeInches(pat As EmbPattern) As String
            ' Best-effort: use stitch bounds as “units”; if metadata contains width_in/height_in, prefer those.
            Try
                Dim wIn As String = pat.GetMetadata("width_in")
                Dim hIn As String = pat.GetMetadata("height_in")
                If Not String.IsNullOrWhiteSpace(wIn) AndAlso Not String.IsNullOrWhiteSpace(hIn) Then
                    Return $"{wIn} × {hIn} in"
                End If
            Catch
            End Try
            ' Fallback: derive from stitch coordinate bounds (unitless → show as inches only if plausible)
            Try
                Dim b As RectangleF = pat.ComputeBounds()
                If b.Width > 0 AndAlso b.Height > 0 Then
                    ' If your project has a real units conversion, replace this.
                    ' For now, just present as inches if bounds look like inch-scale (<= 20")
                    If b.Width <= 20 AndAlso b.Height <= 20 Then
                        Return $"{b.Width:0.##} × {b.Height:0.##} in"
                    End If
                End If
            Catch
            End Try
            Return ""
        End Function

        Private Shared Function FormatBytes(len As Long) As String
            Dim units = {"B", "KB", "MB", "GB"}
            Dim f As Double = len
            Dim u As Integer = 0
            While f >= 1024 AndAlso u < units.Length - 1
                f /= 1024 : u += 1
            End While
            Return f.ToString(If(u = 0, "0", "0.0")) & " " & units(u)
        End Function




        ' Return the thumb size the grid is using, or a sane default.
        Private Function GetThumbSize() As Size
            If _grid IsNot Nothing Then
                Dim pW = _grid.GetType().GetProperty("ThumbWidth")
                Dim pH = _grid.GetType().GetProperty("ThumbHeight")
                If pW IsNot Nothing AndAlso pH IsNot Nothing Then
                    Dim w = CInt(pW.GetValue(_grid, Nothing))
                    Dim h = CInt(pH.GetValue(_grid, Nothing))
                    If w >= 64 AndAlso h >= 64 Then Return New Size(w, h)
                End If
            End If
            ' fall back to settings then a good default
            Dim wSet As Integer = 0, hSet As Integer = 0
            Try : wSet = My.Settings.ThumbWidth : hSet = My.Settings.ThumbHeight : Catch : End Try
            If wSet < 64 OrElse hSet < 64 Then Return New Size(200, 200)
            Return New Size(wSet, hSet)
        End Function



        ' Draw wrapped text (center or left) on a printer/preview Graphics
        Private Shared Sub DrawStringWrap(g As Graphics, text As String, font As Font, rect As Rectangle, center As Boolean)
            Using br As New SolidBrush(Color.Black)
                Using sf As New StringFormat(StringFormatFlags.LineLimit)
                    sf.Trimming = StringTrimming.EllipsisWord
                    sf.Alignment = If(center, StringAlignment.Center, StringAlignment.Near)
                    sf.LineAlignment = StringAlignment.Near
                    g.DrawString(text, font, br, RectangleF.op_Implicit(rect), sf)
                End Using
            End Using
        End Sub




        Private Shared Function GetSelectedFullPathsFromGrid(grid As VirtualThumbGrid) As List(Of String)
            Dim result As New List(Of String)
            If grid Is Nothing Then Return result

            Try
                ' Preferred: multi-select via reflection (_selectedIndices)
                Dim fSel As FieldInfo = GetType(VirtualThumbGrid).GetField("_selectedIndices", BindingFlags.NonPublic Or BindingFlags.Instance)
                Dim fTable As FieldInfo = GetType(VirtualThumbGrid).GetField("_table", BindingFlags.NonPublic Or BindingFlags.Instance)
                Dim selSet = TryCast(fSel?.GetValue(grid), System.Collections.Generic.HashSet(Of Integer))
                Dim dt = TryCast(fTable?.GetValue(grid), DataTable)

                If selSet IsNot Nothing AndAlso dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim ids = selSet.OrderBy(Function(i) i).ToArray()
                    For Each i In ids
                        If i >= 0 AndAlso i < dt.Rows.Count Then
                            Dim stored As String = SafeStr(dt.Rows(i)("fullpath"))
                            If Not String.IsNullOrWhiteSpace(stored) Then
                                result.Add(stored)
                            End If
                        End If
                    Next
                End If

                ' Fallback to single selection if needed
                If result.Count = 0 Then
                    Dim selIdxProp = GetType(VirtualThumbGrid).GetProperty("SelectedIndex", BindingFlags.Public Or BindingFlags.Instance)
                    Dim ix As Integer = CInt(selIdxProp.GetValue(grid, Nothing))
                    If ix >= 0 AndAlso dt IsNot Nothing AndAlso ix < dt.Rows.Count Then
                        Dim stored As String = SafeStr(dt.Rows(ix)("fullpath"))
                        If Not String.IsNullOrWhiteSpace(stored) Then result.Add(stored)
                    End If
                End If
            Catch
                ' ignore; show the “no selection” nudge
            End Try

            Return result
        End Function

        Private Shared Function SafeStr(o As Object) As String
            If o Is Nothing OrElse Convert.IsDBNull(o) Then Return ""
            Return CStr(o)
        End Function
    End Class



End Class
