Public Class OptionsDialog
    Inherits Form

    Private chkZip As CheckBox
    Private grpThumb As GroupBox
    Private rSmall As RadioButton
    Private rMedium As RadioButton
    Private rLarge As RadioButton
    Private rCustom As RadioButton
    Private lblWH As Label
    Private nudW As NumericUpDown
    Private nudH As NumericUpDown
    Private btnOK As Button
    Private btnCancel As Button
    ' === Thumbnail quality UI (created at runtime so no designer edits needed) ===
    Private lblThumbQuality As Label
    Private cboThumbQuality As ComboBox

    Public Sub New()
        Me.Text = "Options"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Font = New Drawing.Font("Segoe UI", 9.0F)

        ' We'll compute the height dynamically after laying out controls
        Me.ClientSize = New Drawing.Size(520, 240)

        ' --- Zip checkbox ---
        chkZip = New CheckBox() With {
        .Text = "Include zip files in scans?",
        .Left = 16, .Top = 16, .Width = 300,
        .Name = "chkIncludeZips"
    }
        Controls.Add(chkZip)

        ' --- Thumbnails group ---
        grpThumb = New GroupBox() With {
        .Text = "Thumbnail Image Size",
        .Left = 12,
        .Top = chkZip.Bottom + 12,
        .Width = 496,
        .Height = 120
    }
        Controls.Add(grpThumb)

        rSmall = New RadioButton() With {.Text = "Small (100 × 100)", .Left = 16, .Top = 24, .Width = 160}
        rMedium = New RadioButton() With {.Text = "Medium (200 × 200)", .Left = 16, .Top = 48, .Width = 170}
        rLarge = New RadioButton() With {.Text = "Large (300 × 300)", .Left = 16, .Top = 72, .Width = 160}
        rCustom = New RadioButton() With {.Text = "Custom:", .Left = 260, .Top = 24, .Width = 80}

        lblWH = New Label() With {.Text = "W × H", .Left = 260, .Top = 50, .Width = 48}
        nudW = New NumericUpDown() With {.Left = 312, .Top = 46, .Width = 70, .Minimum = 32, .Maximum = 1000, .Increment = 10}
        nudH = New NumericUpDown() With {.Left = 392, .Top = 46, .Width = 70, .Minimum = 32, .Maximum = 1000, .Increment = 10}

        AddHandler rSmall.CheckedChanged, AddressOf SizeModeChanged
        AddHandler rMedium.CheckedChanged, AddressOf SizeModeChanged
        AddHandler rLarge.CheckedChanged, AddressOf SizeModeChanged
        AddHandler rCustom.CheckedChanged, AddressOf SizeModeChanged

        grpThumb.Controls.AddRange(New Control() {rSmall, rMedium, rLarge, rCustom, lblWH, nudW, nudH})

        ' --- NEW: Thumbnail quality label + dropdown (UNDER the group) ---
        Dim lblThumbQuality As New Label() With {
        .AutoSize = True,
        .Text = "Thumbnail Image Quality",
        .Left = grpThumb.Left + 4,
        .Top = grpThumb.Bottom + 12,
        .Name = "lblThumbQuality"
    }
        Controls.Add(lblThumbQuality)

        cboThumbQuality = New ComboBox() With {
        .Name = "cboThumbnailQuality",
        .DropDownStyle = ComboBoxStyle.DropDownList,
        .Left = lblThumbQuality.Right + 12,
        .Top = lblThumbQuality.Top - 2,
        .Width = 180
    }
        cboThumbQuality.Items.AddRange(New Object() {"Fast", "High Quality"})
        cboThumbQuality.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Controls.Add(cboThumbQuality)

        ' --- Buttons (placed under the quality row) ---
        Dim buttonsTop As Integer = cboThumbQuality.Bottom + 16
        btnOK = New Button() With {.Text = "OK", .DialogResult = DialogResult.OK, .Left = 316, .Top = buttonsTop, .Width = 90, .Name = "btnOK"}
        btnCancel = New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Left = 418, .Top = buttonsTop, .Width = 90}

        Controls.Add(btnOK)
        Controls.Add(btnCancel)
        Me.AcceptButton = btnOK
        Me.CancelButton = btnCancel

        ' Adjust overall height so buttons are comfortably inside the client area
        Me.ClientSize = New Drawing.Size(Me.ClientSize.Width, btnCancel.Bottom + 16)

        ' Tab order tweaks
        chkZip.TabIndex = 0
        grpThumb.TabIndex = 1
        lblThumbQuality.TabIndex = 2
        cboThumbQuality.TabIndex = 3
        btnOK.TabIndex = 4
        btnCancel.TabIndex = 5

        ' Load & bind
        LoadFromSettings()
        UpdateCustomEnable()
        AddHandler btnOK.Click, AddressOf OnOk
    End Sub



    Private Sub LoadFromSettings()
        ' Checkbox
        chkZip.Checked = My.Settings.IncludeZipFilesInScans

        ' Size mode with safe fallback
        Dim mode As String = If(String.IsNullOrWhiteSpace(My.Settings.ThumbSizeMode), "Medium", My.Settings.ThumbSizeMode).Trim()
        Select Case mode.ToLowerInvariant()
            Case "small" : rSmall.Checked = True
            Case "large" : rLarge.Checked = True
            Case "custom" : rCustom.Checked = True
            Case Else : rMedium.Checked = True
        End Select

        ' Custom W/H
        Dim w As Integer = My.Settings.ThumbWidth
        Dim h As Integer = My.Settings.ThumbHeight
        If w <= nudW.Minimum OrElse w > nudW.Maximum Then w = 200
        If h <= nudH.Minimum OrElse h > nudH.Maximum Then h = 200
        nudW.Value = w
        nudH.Value = h

        'Thumbnail quality
        ' Normalize and select a valid value
        Dim q As String = My.Settings.ThumbnailQuality
        If cboThumbQuality Is Nothing Then
            ' Defensive: nothing to do if the control isn't ready yet
            Exit Sub
        End If
        If Not cboThumbQuality.Items.Contains(q) Then
            q = "Fast"
        End If
        cboThumbQuality.SelectedItem = q
    End Sub


    Private Sub OnOk(sender As Object, e As EventArgs)
        ' Save checkbox
        My.Settings.IncludeZipFilesInScans = chkZip.Checked

        ' Save size mode + resolve preset sizes
        Dim mode As String
        If rSmall.Checked Then
            mode = "Small"
            My.Settings.ThumbWidth = 100 : My.Settings.ThumbHeight = 100
        ElseIf rMedium.Checked Then
            mode = "Medium"
            My.Settings.ThumbWidth = 200 : My.Settings.ThumbHeight = 200
        ElseIf rLarge.Checked Then
            mode = "Large"
            My.Settings.ThumbWidth = 300 : My.Settings.ThumbHeight = 300
        Else
            mode = "Custom"
            My.Settings.ThumbWidth = CInt(nudW.Value)
            My.Settings.ThumbHeight = CInt(nudH.Value)
        End If
        My.Settings.ThumbSizeMode = mode
        'thumbnail quality
        Dim selected As String = If(TryCast(cboThumbQuality.SelectedItem, String), Nothing)
        If String.IsNullOrWhiteSpace(selected) Then selected = "Fast"
        My.Settings.ThumbnailQuality = selected

        ' Persist
        My.Settings.Save()
        ' Let caller close the dialog (DialogResult already OK)
    End Sub

    Private Sub SizeModeChanged(sender As Object, e As EventArgs)
        UpdateCustomEnable()
    End Sub

    Private Sub UpdateCustomEnable()
        Dim enableCustom = rCustom.Checked
        nudW.Enabled = enableCustom
        nudH.Enabled = enableCustom
        lblWH.Enabled = enableCustom
    End Sub
End Class
