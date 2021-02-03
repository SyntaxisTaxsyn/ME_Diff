Public Class EditFilterList

    Private ItemAdded As Boolean = False
    Private Sub EditFilterList_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Call ClearFilterList()
        Call PopulateFilterList()
    End Sub

    Private Sub ClearFilterList()
        CheckedListBox1.Items.Clear()
    End Sub

    Private Sub PopulateFilterList()
        For Each item In FilterConfigurationList
            CheckedListBox1.Items.Add(item, True)
        Next
    End Sub


    Private Sub Update_Button_Click(sender As Object, e As EventArgs) Handles Update_Button.Click
        FilterConfigurationList.Clear()
        For Each item In CheckedListBox1.CheckedItems
            FilterConfigurationList.Add(item)
        Next
        Call WriteFilterListIntoTextFile()
        Call ClearFilterList()
        Call PopulateFilterList()
        Call Main.UpdateFilterPanelControls()
        ItemAdded = False ' Any added entries will still be checked so cancel this flag
    End Sub

    Private Sub Save_Button_Click(sender As Object, e As EventArgs) Handles Save_Button.Click
        FilterConfigurationList.Clear()
        For Each item In CheckedListBox1.Items
            FilterConfigurationList.Add(item)
        Next
        Call WriteFilterListIntoTextFile()
        Call ClearFilterList()
        Call PopulateFilterList()
        Call Main.UpdateFilterPanelControls()
        ItemAdded = False
    End Sub

    Private Sub Add_Button_Click(sender As Object, e As EventArgs) Handles Add_Button.Click
        Dim str As String = InputBox("Enter new name", "New name", "")
        If Not str = "" Then
            CheckedListBox1.Items.Add(str, True)
            ItemAdded = True
        End If
    End Sub

    Private Sub Help_Button_Click(sender As Object, e As EventArgs) Handles Help_Button.Click
        EditFilterListHelp.Show()
    End Sub

    Private Sub Close_Button_Click(sender As Object, e As EventArgs) Handles Close_Button.Click
        If ItemAdded Then
            Dim Result As MsgBoxResult
            Result = MsgBox("Unsaved items have been added that will not be saved, are you sure you want to continue closing?", vbOKCancel, "Items not saved!")
            Select Case Result
                Case MsgBoxResult.Cancel
                    Exit Sub
                Case MsgBoxResult.Ok
                    ' conntinue to next check
            End Select
        End If
        If Not CheckedListBox1.CheckedItems.Count = CheckedListBox1.Items.Count Then
            Dim Result As MsgBoxResult
            Result = MsgBox("Unchecked items exist that have not been saved, are you sure you want to continue closing?", vbOKCancel, "Items not saved!")
            Select Case Result
                Case MsgBoxResult.Cancel
                    ' do nothing the user chose to stay on this form
                Case MsgBoxResult.Ok
                    ' continue closing
                    Me.Close()
            End Select
        Else
            ' no unsaved changes detected
            Me.Close()
        End If
    End Sub
End Class