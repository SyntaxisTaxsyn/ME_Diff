<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EditFilterList
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
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.Save_Button = New System.Windows.Forms.Button()
        Me.Close_Button = New System.Windows.Forms.Button()
        Me.Add_Button = New System.Windows.Forms.Button()
        Me.Help_Button = New System.Windows.Forms.Button()
        Me.Update_Button = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(12, 12)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(446, 424)
        Me.CheckedListBox1.TabIndex = 0
        '
        'Save_Button
        '
        Me.Save_Button.Location = New System.Drawing.Point(12, 442)
        Me.Save_Button.Name = "Save_Button"
        Me.Save_Button.Size = New System.Drawing.Size(80, 32)
        Me.Save_Button.TabIndex = 1
        Me.Save_Button.Text = "Save"
        Me.Save_Button.UseVisualStyleBackColor = True
        '
        'Close_Button
        '
        Me.Close_Button.Location = New System.Drawing.Point(378, 442)
        Me.Close_Button.Name = "Close_Button"
        Me.Close_Button.Size = New System.Drawing.Size(80, 32)
        Me.Close_Button.TabIndex = 2
        Me.Close_Button.Text = "Close"
        Me.Close_Button.UseVisualStyleBackColor = True
        '
        'Add_Button
        '
        Me.Add_Button.Location = New System.Drawing.Point(98, 442)
        Me.Add_Button.Name = "Add_Button"
        Me.Add_Button.Size = New System.Drawing.Size(88, 32)
        Me.Add_Button.TabIndex = 3
        Me.Add_Button.Text = "Add New Item"
        Me.Add_Button.UseVisualStyleBackColor = True
        '
        'Help_Button
        '
        Me.Help_Button.Location = New System.Drawing.Point(286, 443)
        Me.Help_Button.Name = "Help_Button"
        Me.Help_Button.Size = New System.Drawing.Size(86, 32)
        Me.Help_Button.TabIndex = 4
        Me.Help_Button.Text = "Help"
        Me.Help_Button.UseVisualStyleBackColor = True
        '
        'Update_Button
        '
        Me.Update_Button.Location = New System.Drawing.Point(192, 443)
        Me.Update_Button.Name = "Update_Button"
        Me.Update_Button.Size = New System.Drawing.Size(88, 32)
        Me.Update_Button.TabIndex = 5
        Me.Update_Button.Text = "Update"
        Me.Update_Button.UseVisualStyleBackColor = True
        '
        'EditFilterList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(470, 487)
        Me.ControlBox = False
        Me.Controls.Add(Me.Update_Button)
        Me.Controls.Add(Me.Help_Button)
        Me.Controls.Add(Me.Add_Button)
        Me.Controls.Add(Me.Close_Button)
        Me.Controls.Add(Me.Save_Button)
        Me.Controls.Add(Me.CheckedListBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "EditFilterList"
        Me.Text = "Edit Filter List"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents CheckedListBox1 As CheckedListBox
    Friend WithEvents Save_Button As Button
    Friend WithEvents Close_Button As Button
    Friend WithEvents Add_Button As Button
    Friend WithEvents Help_Button As Button
    Friend WithEvents Update_Button As Button
End Class
