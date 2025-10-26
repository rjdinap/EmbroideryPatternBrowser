<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Panel_Main_Top = New System.Windows.Forms.Panel()
        Me.ComboBox_Filter = New System.Windows.Forms.ComboBox()
        Me.Label_Filter = New System.Windows.Forms.Label()
        Me.Button_Search = New System.Windows.Forms.Button()
        Me.TextBox_Search = New System.Windows.Forms.TextBox()
        Me.Label_Search = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenDatabaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CloseDatabaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CreateNewDatabaseToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ScanForImagesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SettingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.MaintenanceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptimizeDatabaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowHelpFIleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel_Main_Bottom = New System.Windows.Forms.Panel()
        Me.Panel_Main_Bottom_Fill = New System.Windows.Forms.Panel()
        Me.RichTextBox_Status = New System.Windows.Forms.RichTextBox()
        Me.Panel_Main_Botton_Top = New System.Windows.Forms.Panel()
        Me.Label_Database = New System.Windows.Forms.Label()
        Me.TextBox_Database = New System.Windows.Forms.TextBox()
        Me.Panel_Main_Fill = New System.Windows.Forms.Panel()
        Me.Panel_Right = New System.Windows.Forms.Panel()
        Me.Panel_Right_Fill = New System.Windows.Forms.Panel()
        Me.TabControl_Panel_Right_Fill = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Panel_TabPage1_Fill = New System.Windows.Forms.Panel()
        Me.PictureBox_FullImage = New EmbroideryPatternBrowser.ZoomPictureBox()
        Me.Panel_TabPage1_Top = New System.Windows.Forms.Panel()
        Me.Button_Rotate = New System.Windows.Forms.Button()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.ZoomPictureBox_Stitches = New EmbroideryPatternBrowser.ZoomPictureBox()
        Me.Panel_TabPage2_Fill_Bottom = New System.Windows.Forms.Panel()
        Me.Panel_TabPage2_Fill_Right = New System.Windows.Forms.Panel()
        Me.Panel_Right_Bottom = New System.Windows.Forms.Panel()
        Me.Panel_Right_Bottom_Fill = New System.Windows.Forms.Panel()
        Me.Panel_Right_Bottom_Top = New System.Windows.Forms.Panel()
        Me.Panel_TabPage2_Fill_Fill = New System.Windows.Forms.Panel()
        Me.Panel_Main_Top.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.Panel_Main_Bottom.SuspendLayout()
        Me.Panel_Main_Bottom_Fill.SuspendLayout()
        Me.Panel_Main_Botton_Top.SuspendLayout()
        Me.Panel_Right.SuspendLayout()
        Me.Panel_Right_Fill.SuspendLayout()
        Me.TabControl_Panel_Right_Fill.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.Panel_TabPage1_Fill.SuspendLayout()
        Me.Panel_TabPage1_Top.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.Panel_Right_Bottom.SuspendLayout()
        Me.Panel_TabPage2_Fill_Fill.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel_Main_Top
        '
        Me.Panel_Main_Top.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Panel_Main_Top.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Main_Top.Controls.Add(Me.ComboBox_Filter)
        Me.Panel_Main_Top.Controls.Add(Me.Label_Filter)
        Me.Panel_Main_Top.Controls.Add(Me.Button_Search)
        Me.Panel_Main_Top.Controls.Add(Me.TextBox_Search)
        Me.Panel_Main_Top.Controls.Add(Me.Label_Search)
        Me.Panel_Main_Top.Controls.Add(Me.MenuStrip1)
        Me.Panel_Main_Top.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Main_Top.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Main_Top.Name = "Panel_Main_Top"
        Me.Panel_Main_Top.Size = New System.Drawing.Size(1356, 59)
        Me.Panel_Main_Top.TabIndex = 0
        '
        'ComboBox_Filter
        '
        Me.ComboBox_Filter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_Filter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox_Filter.FormattingEnabled = True
        Me.ComboBox_Filter.Items.AddRange(New Object() {"", "dst", "exp", "hus", "jef", "pec", "pes", "sew", "vip", "vp3", "xxx"})
        Me.ComboBox_Filter.Location = New System.Drawing.Point(506, 31)
        Me.ComboBox_Filter.Name = "ComboBox_Filter"
        Me.ComboBox_Filter.Size = New System.Drawing.Size(121, 24)
        Me.ComboBox_Filter.TabIndex = 6
        '
        'Label_Filter
        '
        Me.Label_Filter.AutoSize = True
        Me.Label_Filter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Filter.Location = New System.Drawing.Point(463, 33)
        Me.Label_Filter.Name = "Label_Filter"
        Me.Label_Filter.Size = New System.Drawing.Size(39, 16)
        Me.Label_Filter.TabIndex = 5
        Me.Label_Filter.Text = "Filter:"
        '
        'Button_Search
        '
        Me.Button_Search.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Search.Location = New System.Drawing.Point(355, 32)
        Me.Button_Search.Name = "Button_Search"
        Me.Button_Search.Size = New System.Drawing.Size(65, 23)
        Me.Button_Search.TabIndex = 4
        Me.Button_Search.Text = "Search"
        Me.Button_Search.UseVisualStyleBackColor = True
        '
        'TextBox_Search
        '
        Me.TextBox_Search.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Search.Location = New System.Drawing.Point(66, 32)
        Me.TextBox_Search.Name = "TextBox_Search"
        Me.TextBox_Search.Size = New System.Drawing.Size(283, 22)
        Me.TextBox_Search.TabIndex = 3
        '
        'Label_Search
        '
        Me.Label_Search.AutoSize = True
        Me.Label_Search.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Search.Location = New System.Drawing.Point(13, 33)
        Me.Label_Search.Name = "Label_Search"
        Me.Label_Search.Size = New System.Drawing.Size(50, 16)
        Me.Label_Search.TabIndex = 2
        Me.Label_Search.Text = "&Search"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1354, 27)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenDatabaseToolStripMenuItem, Me.CloseDatabaseToolStripMenuItem, Me.CreateNewDatabaseToolStripMenuItem1})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(41, 23)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'OpenDatabaseToolStripMenuItem
        '
        Me.OpenDatabaseToolStripMenuItem.Name = "OpenDatabaseToolStripMenuItem"
        Me.OpenDatabaseToolStripMenuItem.Size = New System.Drawing.Size(210, 24)
        Me.OpenDatabaseToolStripMenuItem.Text = "Open Database"
        '
        'CloseDatabaseToolStripMenuItem
        '
        Me.CloseDatabaseToolStripMenuItem.Name = "CloseDatabaseToolStripMenuItem"
        Me.CloseDatabaseToolStripMenuItem.Size = New System.Drawing.Size(210, 24)
        Me.CloseDatabaseToolStripMenuItem.Text = "Close Database"
        '
        'CreateNewDatabaseToolStripMenuItem1
        '
        Me.CreateNewDatabaseToolStripMenuItem1.Name = "CreateNewDatabaseToolStripMenuItem1"
        Me.CreateNewDatabaseToolStripMenuItem1.Size = New System.Drawing.Size(210, 24)
        Me.CreateNewDatabaseToolStripMenuItem1.Text = "Create New Database"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ScanForImagesToolStripMenuItem, Me.SettingsToolStripMenuItem, Me.ToolStripMenuItem1, Me.MaintenanceToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(52, 23)
        Me.ToolsToolStripMenuItem.Text = "Tools"
        '
        'ScanForImagesToolStripMenuItem
        '
        Me.ScanForImagesToolStripMenuItem.Name = "ScanForImagesToolStripMenuItem"
        Me.ScanForImagesToolStripMenuItem.Size = New System.Drawing.Size(175, 24)
        Me.ScanForImagesToolStripMenuItem.Text = "Scan for Images"
        '
        'SettingsToolStripMenuItem
        '
        Me.SettingsToolStripMenuItem.Name = "SettingsToolStripMenuItem"
        Me.SettingsToolStripMenuItem.Size = New System.Drawing.Size(175, 24)
        Me.SettingsToolStripMenuItem.Text = "Options"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(172, 6)
        '
        'MaintenanceToolStripMenuItem
        '
        Me.MaintenanceToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OptimizeDatabaseToolStripMenuItem})
        Me.MaintenanceToolStripMenuItem.Name = "MaintenanceToolStripMenuItem"
        Me.MaintenanceToolStripMenuItem.Size = New System.Drawing.Size(175, 24)
        Me.MaintenanceToolStripMenuItem.Text = "Maintenance"
        '
        'OptimizeDatabaseToolStripMenuItem
        '
        Me.OptimizeDatabaseToolStripMenuItem.Name = "OptimizeDatabaseToolStripMenuItem"
        Me.OptimizeDatabaseToolStripMenuItem.Size = New System.Drawing.Size(214, 24)
        Me.OptimizeDatabaseToolStripMenuItem.Text = "Optimize Search Index"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem, Me.ShowHelpFIleToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(49, 23)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(168, 24)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'ShowHelpFIleToolStripMenuItem
        '
        Me.ShowHelpFIleToolStripMenuItem.Name = "ShowHelpFIleToolStripMenuItem"
        Me.ShowHelpFIleToolStripMenuItem.Size = New System.Drawing.Size(168, 24)
        Me.ShowHelpFIleToolStripMenuItem.Text = "Show Help FIle"
        '
        'Panel_Main_Bottom
        '
        Me.Panel_Main_Bottom.BackColor = System.Drawing.SystemColors.Control
        Me.Panel_Main_Bottom.Controls.Add(Me.Panel_Main_Bottom_Fill)
        Me.Panel_Main_Bottom.Controls.Add(Me.Panel_Main_Botton_Top)
        Me.Panel_Main_Bottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel_Main_Bottom.Location = New System.Drawing.Point(0, 477)
        Me.Panel_Main_Bottom.Name = "Panel_Main_Bottom"
        Me.Panel_Main_Bottom.Size = New System.Drawing.Size(924, 130)
        Me.Panel_Main_Bottom.TabIndex = 2
        '
        'Panel_Main_Bottom_Fill
        '
        Me.Panel_Main_Bottom_Fill.Controls.Add(Me.RichTextBox_Status)
        Me.Panel_Main_Bottom_Fill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel_Main_Bottom_Fill.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Main_Bottom_Fill.Name = "Panel_Main_Bottom_Fill"
        Me.Panel_Main_Bottom_Fill.Size = New System.Drawing.Size(924, 100)
        Me.Panel_Main_Bottom_Fill.TabIndex = 5
        '
        'RichTextBox_Status
        '
        Me.RichTextBox_Status.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox_Status.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBox_Status.Location = New System.Drawing.Point(0, 0)
        Me.RichTextBox_Status.Name = "RichTextBox_Status"
        Me.RichTextBox_Status.Size = New System.Drawing.Size(924, 100)
        Me.RichTextBox_Status.TabIndex = 0
        Me.RichTextBox_Status.Text = ""
        '
        'Panel_Main_Botton_Top
        '
        Me.Panel_Main_Botton_Top.Controls.Add(Me.Label_Database)
        Me.Panel_Main_Botton_Top.Controls.Add(Me.TextBox_Database)
        Me.Panel_Main_Botton_Top.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel_Main_Botton_Top.Location = New System.Drawing.Point(0, 100)
        Me.Panel_Main_Botton_Top.Name = "Panel_Main_Botton_Top"
        Me.Panel_Main_Botton_Top.Size = New System.Drawing.Size(924, 30)
        Me.Panel_Main_Botton_Top.TabIndex = 4
        '
        'Label_Database
        '
        Me.Label_Database.AutoSize = True
        Me.Label_Database.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Database.Location = New System.Drawing.Point(3, 9)
        Me.Label_Database.Name = "Label_Database"
        Me.Label_Database.Size = New System.Drawing.Size(92, 20)
        Me.Label_Database.TabIndex = 2
        Me.Label_Database.Text = "Database:"
        '
        'TextBox_Database
        '
        Me.TextBox_Database.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox_Database.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox_Database.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Database.Location = New System.Drawing.Point(101, 10)
        Me.TextBox_Database.Name = "TextBox_Database"
        Me.TextBox_Database.ReadOnly = True
        Me.TextBox_Database.Size = New System.Drawing.Size(515, 19)
        Me.TextBox_Database.TabIndex = 3
        Me.TextBox_Database.Text = "No Database Opened."
        '
        'Panel_Main_Fill
        '
        Me.Panel_Main_Fill.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Main_Fill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel_Main_Fill.Location = New System.Drawing.Point(0, 59)
        Me.Panel_Main_Fill.Name = "Panel_Main_Fill"
        Me.Panel_Main_Fill.Size = New System.Drawing.Size(924, 418)
        Me.Panel_Main_Fill.TabIndex = 4
        '
        'Panel_Right
        '
        Me.Panel_Right.Controls.Add(Me.Panel_Right_Fill)
        Me.Panel_Right.Controls.Add(Me.Panel_Right_Bottom)
        Me.Panel_Right.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel_Right.Location = New System.Drawing.Point(924, 59)
        Me.Panel_Right.Name = "Panel_Right"
        Me.Panel_Right.Size = New System.Drawing.Size(432, 548)
        Me.Panel_Right.TabIndex = 1
        '
        'Panel_Right_Fill
        '
        Me.Panel_Right_Fill.AutoScroll = True
        Me.Panel_Right_Fill.Controls.Add(Me.TabControl_Panel_Right_Fill)
        Me.Panel_Right_Fill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel_Right_Fill.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Right_Fill.Name = "Panel_Right_Fill"
        Me.Panel_Right_Fill.Size = New System.Drawing.Size(432, 319)
        Me.Panel_Right_Fill.TabIndex = 3
        '
        'TabControl_Panel_Right_Fill
        '
        Me.TabControl_Panel_Right_Fill.Controls.Add(Me.TabPage1)
        Me.TabControl_Panel_Right_Fill.Controls.Add(Me.TabPage2)
        Me.TabControl_Panel_Right_Fill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl_Panel_Right_Fill.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl_Panel_Right_Fill.Location = New System.Drawing.Point(0, 0)
        Me.TabControl_Panel_Right_Fill.Name = "TabControl_Panel_Right_Fill"
        Me.TabControl_Panel_Right_Fill.SelectedIndex = 0
        Me.TabControl_Panel_Right_Fill.Size = New System.Drawing.Size(432, 319)
        Me.TabControl_Panel_Right_Fill.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Panel_TabPage1_Fill)
        Me.TabPage1.Controls.Add(Me.Panel_TabPage1_Top)
        Me.TabPage1.Location = New System.Drawing.Point(4, 29)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(424, 286)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Image"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Panel_TabPage1_Fill
        '
        Me.Panel_TabPage1_Fill.Controls.Add(Me.PictureBox_FullImage)
        Me.Panel_TabPage1_Fill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel_TabPage1_Fill.Location = New System.Drawing.Point(3, 39)
        Me.Panel_TabPage1_Fill.Name = "Panel_TabPage1_Fill"
        Me.Panel_TabPage1_Fill.Size = New System.Drawing.Size(418, 244)
        Me.Panel_TabPage1_Fill.TabIndex = 1
        '
        'PictureBox_FullImage
        '
        Me.PictureBox_FullImage.BackColor = System.Drawing.Color.White
        Me.PictureBox_FullImage.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox_FullImage.Image = Nothing
        Me.PictureBox_FullImage.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox_FullImage.Name = "PictureBox_FullImage"
        Me.PictureBox_FullImage.Size = New System.Drawing.Size(418, 244)
        Me.PictureBox_FullImage.TabIndex = 0
        Me.PictureBox_FullImage.Zoom = 1.0R
        '
        'Panel_TabPage1_Top
        '
        Me.Panel_TabPage1_Top.BackColor = System.Drawing.SystemColors.Control
        Me.Panel_TabPage1_Top.Controls.Add(Me.Button_Rotate)
        Me.Panel_TabPage1_Top.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_TabPage1_Top.Location = New System.Drawing.Point(3, 3)
        Me.Panel_TabPage1_Top.Name = "Panel_TabPage1_Top"
        Me.Panel_TabPage1_Top.Size = New System.Drawing.Size(418, 36)
        Me.Panel_TabPage1_Top.TabIndex = 0
        '
        'Button_Rotate
        '
        Me.Button_Rotate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Rotate.Location = New System.Drawing.Point(3, 6)
        Me.Button_Rotate.Name = "Button_Rotate"
        Me.Button_Rotate.Size = New System.Drawing.Size(66, 23)
        Me.Button_Rotate.TabIndex = 0
        Me.Button_Rotate.Text = "Rotate"
        Me.Button_Rotate.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Panel_TabPage2_Fill_Fill)
        Me.TabPage2.Controls.Add(Me.Panel_TabPage2_Fill_Right)
        Me.TabPage2.Controls.Add(Me.Panel_TabPage2_Fill_Bottom)
        Me.TabPage2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage2.Location = New System.Drawing.Point(4, 29)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(424, 286)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Stitches"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'ZoomPictureBox_Stitches
        '
        Me.ZoomPictureBox_Stitches.BackColor = System.Drawing.Color.White
        Me.ZoomPictureBox_Stitches.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ZoomPictureBox_Stitches.Image = Nothing
        Me.ZoomPictureBox_Stitches.Location = New System.Drawing.Point(0, 0)
        Me.ZoomPictureBox_Stitches.Name = "ZoomPictureBox_Stitches"
        Me.ZoomPictureBox_Stitches.Size = New System.Drawing.Size(332, 234)
        Me.ZoomPictureBox_Stitches.TabIndex = 0
        Me.ZoomPictureBox_Stitches.Text = "ZoomPictureBox1"
        Me.ZoomPictureBox_Stitches.Zoom = 1.0R
        '
        'Panel_TabPage2_Fill_Bottom
        '
        Me.Panel_TabPage2_Fill_Bottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel_TabPage2_Fill_Bottom.Location = New System.Drawing.Point(3, 237)
        Me.Panel_TabPage2_Fill_Bottom.Name = "Panel_TabPage2_Fill_Bottom"
        Me.Panel_TabPage2_Fill_Bottom.Size = New System.Drawing.Size(418, 46)
        Me.Panel_TabPage2_Fill_Bottom.TabIndex = 1
        '
        'Panel_TabPage2_Fill_Right
        '
        Me.Panel_TabPage2_Fill_Right.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel_TabPage2_Fill_Right.Location = New System.Drawing.Point(335, 3)
        Me.Panel_TabPage2_Fill_Right.Name = "Panel_TabPage2_Fill_Right"
        Me.Panel_TabPage2_Fill_Right.Size = New System.Drawing.Size(86, 234)
        Me.Panel_TabPage2_Fill_Right.TabIndex = 0
        '
        'Panel_Right_Bottom
        '
        Me.Panel_Right_Bottom.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Panel_Right_Bottom.Controls.Add(Me.Panel_Right_Bottom_Fill)
        Me.Panel_Right_Bottom.Controls.Add(Me.Panel_Right_Bottom_Top)
        Me.Panel_Right_Bottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel_Right_Bottom.Location = New System.Drawing.Point(0, 319)
        Me.Panel_Right_Bottom.Name = "Panel_Right_Bottom"
        Me.Panel_Right_Bottom.Size = New System.Drawing.Size(432, 229)
        Me.Panel_Right_Bottom.TabIndex = 2
        '
        'Panel_Right_Bottom_Fill
        '
        Me.Panel_Right_Bottom_Fill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel_Right_Bottom_Fill.Location = New System.Drawing.Point(0, 34)
        Me.Panel_Right_Bottom_Fill.Name = "Panel_Right_Bottom_Fill"
        Me.Panel_Right_Bottom_Fill.Size = New System.Drawing.Size(432, 195)
        Me.Panel_Right_Bottom_Fill.TabIndex = 1
        '
        'Panel_Right_Bottom_Top
        '
        Me.Panel_Right_Bottom_Top.BackColor = System.Drawing.SystemColors.Control
        Me.Panel_Right_Bottom_Top.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Right_Bottom_Top.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Right_Bottom_Top.Name = "Panel_Right_Bottom_Top"
        Me.Panel_Right_Bottom_Top.Size = New System.Drawing.Size(432, 34)
        Me.Panel_Right_Bottom_Top.TabIndex = 0
        '
        'Panel_TabPage2_Fill_Fill
        '
        Me.Panel_TabPage2_Fill_Fill.Controls.Add(Me.ZoomPictureBox_Stitches)
        Me.Panel_TabPage2_Fill_Fill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel_TabPage2_Fill_Fill.Location = New System.Drawing.Point(3, 3)
        Me.Panel_TabPage2_Fill_Fill.Name = "Panel_TabPage2_Fill_Fill"
        Me.Panel_TabPage2_Fill_Fill.Size = New System.Drawing.Size(332, 234)
        Me.Panel_TabPage2_Fill_Fill.TabIndex = 2
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1356, 607)
        Me.Controls.Add(Me.Panel_Main_Fill)
        Me.Controls.Add(Me.Panel_Main_Bottom)
        Me.Controls.Add(Me.Panel_Right)
        Me.Controls.Add(Me.Panel_Main_Top)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.Text = "Embroidery Pattern Browser"
        Me.Panel_Main_Top.ResumeLayout(False)
        Me.Panel_Main_Top.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.Panel_Main_Bottom.ResumeLayout(False)
        Me.Panel_Main_Bottom_Fill.ResumeLayout(False)
        Me.Panel_Main_Botton_Top.ResumeLayout(False)
        Me.Panel_Main_Botton_Top.PerformLayout()
        Me.Panel_Right.ResumeLayout(False)
        Me.Panel_Right_Fill.ResumeLayout(False)
        Me.TabControl_Panel_Right_Fill.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.Panel_TabPage1_Fill.ResumeLayout(False)
        Me.Panel_TabPage1_Top.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.Panel_Right_Bottom.ResumeLayout(False)
        Me.Panel_TabPage2_Fill_Fill.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel_Main_Top As System.Windows.Forms.Panel
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Panel_Main_Bottom As System.Windows.Forms.Panel
    Friend WithEvents Panel_Main_Fill As System.Windows.Forms.Panel
    Friend WithEvents ToolsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Button_Search As Button
    Friend WithEvents TextBox_Search As TextBox
    Friend WithEvents Label_Search As Label
    Friend WithEvents Panel_Right As Panel
    Friend WithEvents Button_Rotate As Button
    Friend WithEvents ComboBox_Filter As ComboBox
    Friend WithEvents Label_Filter As Label
    Friend WithEvents TextBox_Database As TextBox
    Friend WithEvents Label_Database As Label
    Friend WithEvents HelpToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Panel_Main_Bottom_Fill As Panel
    Friend WithEvents RichTextBox_Status As RichTextBox
    Friend WithEvents Panel_Main_Botton_Top As Panel
    Friend WithEvents Panel_Right_Bottom As Panel
    Friend WithEvents ShowHelpFIleToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Panel_Right_Fill As Panel
    Friend WithEvents OpenDatabaseToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Panel_Right_Bottom_Fill As Panel
    Friend WithEvents Panel_Right_Bottom_Top As Panel
    Friend WithEvents CloseDatabaseToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CreateNewDatabaseToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ScanForImagesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SettingsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripSeparator
    Friend WithEvents MaintenanceToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OptimizeDatabaseToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PictureBox_FullImage As ZoomPictureBox
    Friend WithEvents TabControl_Panel_Right_Fill As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents Panel_TabPage1_Fill As Panel
    Friend WithEvents Panel_TabPage1_Top As Panel
    Friend WithEvents Panel_TabPage2_Fill_Right As Panel
    Friend WithEvents ZoomPictureBox_Stitches As ZoomPictureBox
    Friend WithEvents Panel_TabPage2_Fill_Bottom As Panel
    Friend WithEvents Panel_TabPage2_Fill_Fill As Panel
End Class
