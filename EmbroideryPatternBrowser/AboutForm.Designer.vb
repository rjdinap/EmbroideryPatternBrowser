<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AboutForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.RichTextBox_About_About = New System.Windows.Forms.RichTextBox()
        Me.Button_About_Close = New System.Windows.Forms.Button()
        Me.Panel_About_Top = New System.Windows.Forms.Panel()
        Me.FormAbout_PictureBox = New System.Windows.Forms.PictureBox()
        Me.Panel_About_Bottom = New System.Windows.Forms.Panel()
        Me.Panel_About_Fill = New System.Windows.Forms.Panel()
        Me.Panel_About_Top.SuspendLayout()
        CType(Me.FormAbout_PictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_About_Bottom.SuspendLayout()
        Me.Panel_About_Fill.SuspendLayout()
        Me.SuspendLayout()
        '
        'RichTextBox_About_About
        '
        Me.RichTextBox_About_About.BackColor = System.Drawing.SystemColors.Control
        Me.RichTextBox_About_About.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox_About_About.CausesValidation = False
        Me.RichTextBox_About_About.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBox_About_About.Location = New System.Drawing.Point(20, 0)
        Me.RichTextBox_About_About.Margin = New System.Windows.Forms.Padding(20)
        Me.RichTextBox_About_About.Name = "RichTextBox_About_About"
        Me.RichTextBox_About_About.ReadOnly = True
        Me.RichTextBox_About_About.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
        Me.RichTextBox_About_About.Size = New System.Drawing.Size(722, 197)
        Me.RichTextBox_About_About.TabIndex = 4
        Me.RichTextBox_About_About.Text = ""
        '
        'Button_About_Close
        '
        Me.Button_About_Close.Location = New System.Drawing.Point(716, 9)
        Me.Button_About_Close.Name = "Button_About_Close"
        Me.Button_About_Close.Size = New System.Drawing.Size(34, 23)
        Me.Button_About_Close.TabIndex = 6
        Me.Button_About_Close.Text = "Ok"
        Me.Button_About_Close.UseVisualStyleBackColor = True
        '
        'Panel_About_Top
        '
        Me.Panel_About_Top.Controls.Add(Me.FormAbout_PictureBox)
        Me.Panel_About_Top.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_About_Top.Location = New System.Drawing.Point(0, 0)
        Me.Panel_About_Top.Name = "Panel_About_Top"
        Me.Panel_About_Top.Size = New System.Drawing.Size(762, 75)
        Me.Panel_About_Top.TabIndex = 7
        '
        'FormAbout_PictureBox
        '
        Me.FormAbout_PictureBox.Image = Global.EmbroideryPatternBrowser.My.Resources.Resources.ring
        Me.FormAbout_PictureBox.Location = New System.Drawing.Point(11, 11)
        Me.FormAbout_PictureBox.Name = "FormAbout_PictureBox"
        Me.FormAbout_PictureBox.Size = New System.Drawing.Size(52, 46)
        Me.FormAbout_PictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.FormAbout_PictureBox.TabIndex = 5
        Me.FormAbout_PictureBox.TabStop = False
        '
        'Panel_About_Bottom
        '
        Me.Panel_About_Bottom.Controls.Add(Me.Button_About_Close)
        Me.Panel_About_Bottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel_About_Bottom.Location = New System.Drawing.Point(0, 272)
        Me.Panel_About_Bottom.Name = "Panel_About_Bottom"
        Me.Panel_About_Bottom.Size = New System.Drawing.Size(762, 44)
        Me.Panel_About_Bottom.TabIndex = 8
        '
        'Panel_About_Fill
        '
        Me.Panel_About_Fill.Controls.Add(Me.RichTextBox_About_About)
        Me.Panel_About_Fill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel_About_Fill.Location = New System.Drawing.Point(0, 75)
        Me.Panel_About_Fill.Name = "Panel_About_Fill"
        Me.Panel_About_Fill.Size = New System.Drawing.Size(762, 197)
        Me.Panel_About_Fill.TabIndex = 9
        '
        'AboutForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(762, 316)
        Me.Controls.Add(Me.Panel_About_Fill)
        Me.Controls.Add(Me.Panel_About_Bottom)
        Me.Controls.Add(Me.Panel_About_Top)
        Me.Name = "AboutForm"
        Me.Text = "Embroidery Pattern Organizer"
        Me.Panel_About_Top.ResumeLayout(False)
        CType(Me.FormAbout_PictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_About_Bottom.ResumeLayout(False)
        Me.Panel_About_Fill.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents RichTextBox_About_About As RichTextBox
    Friend WithEvents FormAbout_PictureBox As PictureBox
    Friend WithEvents Button_About_Close As Button
    Friend WithEvents Panel_About_Top As Panel
    Friend WithEvents Panel_About_Bottom As Panel
    Friend WithEvents Panel_About_Fill As Panel
End Class
