<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Main
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.FolderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Pnl_ComparisonFilters = New System.Windows.Forms.Panel()
        Me.FontBold_Checkbox = New System.Windows.Forms.CheckBox()
        Me.FontType_Checkbox = New System.Windows.Forms.CheckBox()
        Me.FontSize_Checkbox = New System.Windows.Forms.CheckBox()
        Me.Left_Checkbox = New System.Windows.Forms.CheckBox()
        Me.Top_Checkbox = New System.Windows.Forms.CheckBox()
        Me.Width_Checkbox = New System.Windows.Forms.CheckBox()
        Me.Height_Checkbox = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.UncheckAll_Button = New System.Windows.Forms.Button()
        Me.CheckAll_Button = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditFiltersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.Pnl_ComparisonFilters.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.EditToolStripMenuItem, Me.HelpToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(800, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem1, Me.FolderToolStripMenuItem})
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.OpenToolStripMenuItem.Text = "Open"
        '
        'FileToolStripMenuItem1
        '
        Me.FileToolStripMenuItem1.Name = "FileToolStripMenuItem1"
        Me.FileToolStripMenuItem1.Size = New System.Drawing.Size(180, 22)
        Me.FileToolStripMenuItem1.Text = "File"
        '
        'FolderToolStripMenuItem
        '
        Me.FolderToolStripMenuItem.Name = "FolderToolStripMenuItem"
        Me.FolderToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.FolderToolStripMenuItem.Text = "Folder"
        Me.FolderToolStripMenuItem.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(713, 362)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 25)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Compare"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(418, 366)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Quick Load"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'Pnl_ComparisonFilters
        '
        Me.Pnl_ComparisonFilters.AutoScroll = True
        Me.Pnl_ComparisonFilters.BackColor = System.Drawing.SystemColors.Control
        Me.Pnl_ComparisonFilters.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Pnl_ComparisonFilters.Controls.Add(Me.FontBold_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.FontType_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.FontSize_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.Left_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.Top_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.Width_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.Height_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.Label1)
        Me.Pnl_ComparisonFilters.Location = New System.Drawing.Point(554, 53)
        Me.Pnl_ComparisonFilters.Name = "Pnl_ComparisonFilters"
        Me.Pnl_ComparisonFilters.Size = New System.Drawing.Size(234, 303)
        Me.Pnl_ComparisonFilters.TabIndex = 6
        '
        'FontBold_Checkbox
        '
        Me.FontBold_Checkbox.AutoSize = True
        Me.FontBold_Checkbox.Checked = True
        Me.FontBold_Checkbox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.FontBold_Checkbox.Location = New System.Drawing.Point(26, 180)
        Me.FontBold_Checkbox.Name = "FontBold_Checkbox"
        Me.FontBold_Checkbox.Size = New System.Drawing.Size(46, 17)
        Me.FontBold_Checkbox.TabIndex = 7
        Me.FontBold_Checkbox.Text = "bold"
        Me.FontBold_Checkbox.UseVisualStyleBackColor = True
        '
        'FontType_Checkbox
        '
        Me.FontType_Checkbox.AutoSize = True
        Me.FontType_Checkbox.Checked = True
        Me.FontType_Checkbox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.FontType_Checkbox.Location = New System.Drawing.Point(26, 157)
        Me.FontType_Checkbox.Name = "FontType_Checkbox"
        Me.FontType_Checkbox.Size = New System.Drawing.Size(73, 17)
        Me.FontType_Checkbox.TabIndex = 6
        Me.FontType_Checkbox.Text = "fontFamily"
        Me.FontType_Checkbox.UseVisualStyleBackColor = True
        '
        'FontSize_Checkbox
        '
        Me.FontSize_Checkbox.AutoSize = True
        Me.FontSize_Checkbox.Checked = True
        Me.FontSize_Checkbox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.FontSize_Checkbox.Location = New System.Drawing.Point(26, 134)
        Me.FontSize_Checkbox.Name = "FontSize_Checkbox"
        Me.FontSize_Checkbox.Size = New System.Drawing.Size(64, 17)
        Me.FontSize_Checkbox.TabIndex = 5
        Me.FontSize_Checkbox.Text = "fontSize"
        Me.FontSize_Checkbox.UseVisualStyleBackColor = True
        '
        'Left_Checkbox
        '
        Me.Left_Checkbox.AutoSize = True
        Me.Left_Checkbox.Checked = True
        Me.Left_Checkbox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Left_Checkbox.Location = New System.Drawing.Point(26, 111)
        Me.Left_Checkbox.Name = "Left_Checkbox"
        Me.Left_Checkbox.Size = New System.Drawing.Size(40, 17)
        Me.Left_Checkbox.TabIndex = 4
        Me.Left_Checkbox.Text = "left"
        Me.Left_Checkbox.UseVisualStyleBackColor = True
        '
        'Top_Checkbox
        '
        Me.Top_Checkbox.AutoSize = True
        Me.Top_Checkbox.Checked = True
        Me.Top_Checkbox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Top_Checkbox.Location = New System.Drawing.Point(26, 88)
        Me.Top_Checkbox.Name = "Top_Checkbox"
        Me.Top_Checkbox.Size = New System.Drawing.Size(41, 17)
        Me.Top_Checkbox.TabIndex = 3
        Me.Top_Checkbox.Text = "top"
        Me.Top_Checkbox.UseVisualStyleBackColor = True
        '
        'Width_Checkbox
        '
        Me.Width_Checkbox.AutoSize = True
        Me.Width_Checkbox.Checked = True
        Me.Width_Checkbox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Width_Checkbox.Location = New System.Drawing.Point(26, 65)
        Me.Width_Checkbox.Name = "Width_Checkbox"
        Me.Width_Checkbox.Size = New System.Drawing.Size(51, 17)
        Me.Width_Checkbox.TabIndex = 2
        Me.Width_Checkbox.Text = "width"
        Me.Width_Checkbox.UseVisualStyleBackColor = True
        '
        'Height_Checkbox
        '
        Me.Height_Checkbox.AutoSize = True
        Me.Height_Checkbox.Checked = True
        Me.Height_Checkbox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Height_Checkbox.Location = New System.Drawing.Point(26, 42)
        Me.Height_Checkbox.Name = "Height_Checkbox"
        Me.Height_Checkbox.Size = New System.Drawing.Size(55, 17)
        Me.Height_Checkbox.TabIndex = 1
        Me.Height_Checkbox.Text = "height"
        Me.Height_Checkbox.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(23, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 15)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Comparison Filters"
        '
        'UncheckAll_Button
        '
        Me.UncheckAll_Button.Location = New System.Drawing.Point(635, 362)
        Me.UncheckAll_Button.Name = "UncheckAll_Button"
        Me.UncheckAll_Button.Size = New System.Drawing.Size(75, 25)
        Me.UncheckAll_Button.TabIndex = 8
        Me.UncheckAll_Button.Text = "Uncheck All"
        Me.UncheckAll_Button.UseVisualStyleBackColor = True
        '
        'CheckAll_Button
        '
        Me.CheckAll_Button.Location = New System.Drawing.Point(554, 362)
        Me.CheckAll_Button.Name = "CheckAll_Button"
        Me.CheckAll_Button.Size = New System.Drawing.Size(75, 25)
        Me.CheckAll_Button.TabIndex = 7
        Me.CheckAll_Button.Text = "Check All"
        Me.CheckAll_Button.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(327, 53)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(183, 303)
        Me.ListBox1.TabIndex = 7
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(337, 364)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 25)
        Me.Button3.TabIndex = 9
        Me.Button3.Text = "Edit Filters"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(251, 364)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 10
        Me.Button4.Text = "Test"
        Me.Button4.UseVisualStyleBackColor = True
        Me.Button4.Visible = False
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(155, 364)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(75, 23)
        Me.Button5.TabIndex = 11
        Me.Button5.Text = "Test"
        Me.Button5.UseVisualStyleBackColor = True
        Me.Button5.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(551, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 15)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Filter Selection"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(324, 35)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 15)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Active Filters"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EditFiltersToolStripMenuItem})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.EditToolStripMenuItem.Text = "Edit"
        '
        'EditFiltersToolStripMenuItem
        '
        Me.EditFiltersToolStripMenuItem.Name = "EditFiltersToolStripMenuItem"
        Me.EditFiltersToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.EditFiltersToolStripMenuItem.Text = "Edit Filters"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(55, 106)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(210, 250)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 14
        Me.PictureBox1.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(97, 35)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(110, 31)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "ME Diff"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(51, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(203, 20)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "HMI Page Comparison Tool"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(55, 388)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(455, 29)
        Me.ProgressBar1.TabIndex = 17
        Me.ProgressBar1.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(52, 369)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 15)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "Active Filters"
        Me.Label6.Visible = False
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(800, 429)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.UncheckAll_Button)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.CheckAll_Button)
        Me.Controls.Add(Me.Pnl_ComparisonFilters)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Main"
        Me.Text = "ME Diff - Comparison Tool"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.Pnl_ComparisonFilters.ResumeLayout(False)
        Me.Pnl_ComparisonFilters.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FileToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents FolderToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Pnl_ComparisonFilters As Panel
    Friend WithEvents UncheckAll_Button As Button
    Friend WithEvents CheckAll_Button As Button
    Friend WithEvents FontBold_Checkbox As CheckBox
    Friend WithEvents FontType_Checkbox As CheckBox
    Friend WithEvents FontSize_Checkbox As CheckBox
    Friend WithEvents Left_Checkbox As CheckBox
    Friend WithEvents Top_Checkbox As CheckBox
    Friend WithEvents Width_Checkbox As CheckBox
    Friend WithEvents Height_Checkbox As CheckBox
    Friend WithEvents Label1 As Label
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents EditToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents EditFiltersToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents HelpToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents Label6 As Label
End Class
