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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.FolderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Pnl_ComparisonFilters = New System.Windows.Forms.Panel()
        Me.UncheckAll_Button = New System.Windows.Forms.Button()
        Me.CheckAll_Button = New System.Windows.Forms.Button()
        Me.FontBold_Checkbox = New System.Windows.Forms.CheckBox()
        Me.FontType_Checkbox = New System.Windows.Forms.CheckBox()
        Me.FontSize_Checkbox = New System.Windows.Forms.CheckBox()
        Me.Left_Checkbox = New System.Windows.Forms.CheckBox()
        Me.Top_Checkbox = New System.Windows.Forms.CheckBox()
        Me.Width_Checkbox = New System.Windows.Forms.CheckBox()
        Me.Height_Checkbox = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.MenuStrip1.SuspendLayout()
        Me.Pnl_ComparisonFilters.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
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
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(103, 22)
        Me.OpenToolStripMenuItem.Text = "Open"
        '
        'FileToolStripMenuItem1
        '
        Me.FileToolStripMenuItem1.Name = "FileToolStripMenuItem1"
        Me.FileToolStripMenuItem1.Size = New System.Drawing.Size(107, 22)
        Me.FileToolStripMenuItem1.Text = "File"
        '
        'FolderToolStripMenuItem
        '
        Me.FolderToolStripMenuItem.Name = "FolderToolStripMenuItem"
        Me.FolderToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.FolderToolStripMenuItem.Text = "Folder"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(713, 415)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Compare"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(632, 415)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Quick Load"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Pnl_ComparisonFilters
        '
        Me.Pnl_ComparisonFilters.BackColor = System.Drawing.SystemColors.Control
        Me.Pnl_ComparisonFilters.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Pnl_ComparisonFilters.Controls.Add(Me.UncheckAll_Button)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.CheckAll_Button)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.FontBold_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.FontType_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.FontSize_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.Left_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.Top_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.Width_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.Height_Checkbox)
        Me.Pnl_ComparisonFilters.Controls.Add(Me.Label1)
        Me.Pnl_ComparisonFilters.Location = New System.Drawing.Point(570, 53)
        Me.Pnl_ComparisonFilters.Name = "Pnl_ComparisonFilters"
        Me.Pnl_ComparisonFilters.Size = New System.Drawing.Size(207, 341)
        Me.Pnl_ComparisonFilters.TabIndex = 6
        '
        'UncheckAll_Button
        '
        Me.UncheckAll_Button.Location = New System.Drawing.Point(84, 309)
        Me.UncheckAll_Button.Name = "UncheckAll_Button"
        Me.UncheckAll_Button.Size = New System.Drawing.Size(75, 25)
        Me.UncheckAll_Button.TabIndex = 8
        Me.UncheckAll_Button.Text = "Uncheck All"
        Me.UncheckAll_Button.UseVisualStyleBackColor = True
        '
        'CheckAll_Button
        '
        Me.CheckAll_Button.Location = New System.Drawing.Point(3, 309)
        Me.CheckAll_Button.Name = "CheckAll_Button"
        Me.CheckAll_Button.Size = New System.Drawing.Size(75, 25)
        Me.CheckAll_Button.TabIndex = 7
        Me.CheckAll_Button.Text = "Check All"
        Me.CheckAll_Button.UseVisualStyleBackColor = True
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
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(191, 53)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(183, 342)
        Me.ListBox1.TabIndex = 7
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Pnl_ComparisonFilters)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Main"
        Me.Text = "Main"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.Pnl_ComparisonFilters.ResumeLayout(False)
        Me.Pnl_ComparisonFilters.PerformLayout()
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
End Class
