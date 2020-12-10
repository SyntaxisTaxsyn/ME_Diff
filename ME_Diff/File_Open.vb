Public Class File_Open
    Private Sub File_Open_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        Main.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ' Display File Picker for HMI XML export file
        Me.OpenFileDialog1.Filter = "XML FILES (.XML)|*.XML" ' set file filter
        Me.OpenFileDialog1.Title = "Select the HMI XML Display Export"
        Me.OpenFileDialog1.ShowDialog(Me) 'display file picker
        STR_LeftPathFile = OpenFileDialog1.FileName.ToString 'assign file path to variable

        If STR_LeftPathFile Like "*OpenFileDialog*" Then
            MsgBox("No file picked")
            Application.Exit()
        End If

        Me.TextBox1.Text = STR_LeftPathFile

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ' Display File Picker for HMI XML export file
        Me.OpenFileDialog1.Filter = "XML FILES (.XML)|*.XML" ' set file filter
        Me.OpenFileDialog1.Title = "Select the HMI XML Display Export"
        Me.OpenFileDialog1.ShowDialog(Me) 'display file picker
        STR_RightPathFile = OpenFileDialog1.FileName.ToString 'assign file path to variable

        If STR_RightPathFile Like "*OpenFileDialog*" Then
            MsgBox("No file picked")
            Application.Exit()
        End If

        Me.TextBox2.Text = STR_RightPathFile

    End Sub
End Class