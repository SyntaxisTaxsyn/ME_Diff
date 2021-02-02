Imports System.CodeDom.Compiler
Imports System.IO
Imports System.Xml
Imports System.Xml.Schema
Imports ME_Diff.gfx
Imports System.Reflection
Imports System.Runtime.Serialization.Formatters.Binary


Public Class Main
    Private Sub FileToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles FileToolStripMenuItem1.Click
        File_Open.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Fcompare As New ME_Diff.FileCompare
        Dim str As String = Fcompare.CompareFiles(FileCompare.UnitTestType.DontUse)
        If Not str = "" Then
            ' Only display errors/differences returned by the comparison operation
            MsgBox(str)
        End If
    End Sub


    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load
        List_FileCompareContentMatch = New List(Of String)
        List_FileCompareContentNoMatch = New List(Of String)
        List_FolderCompare = New List(Of String)
        FilterList = New List(Of String)
        Call RemoveAllCheckBoxSubscriptions(Pnl_ComparisonFilters)
        Call AddAllCheckBoxSubscriptions(Pnl_ComparisonFilters)
    End Sub


    ''' <summary>
    ''' Temporary code, delete on cleanup
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(sender As Object, e As EventArgs)

        Dim Filestr As New gfx

        Filestr = DeserializeObjectFromXMLFile("C:/Temp/Drawing.xml")

        Dim stream As FileStream = File.Create("C:/Temp/page2test.gfx")
        Dim formatter As New BinaryFormatter()
        formatter.Serialize(stream, Filestr)
        stream.Close()

        ' Restore from file
        'Dim stream As FileStream
        'Dim formatter As New BinaryFormatter()
        stream = File.OpenRead("C:/Temp/Page2test.gfx")
        Dim Page As gfx = formatter.Deserialize(stream)
        stream.Close()
        MsgBox("")

    End Sub

    ''' <summary>
    ''' Temp code delete on cleanup
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        ' Temporary code used to quickly set the names of the test files im using during application developmment
        STR_LeftPathFile = "C:\temp\ME_Diff\Page2.xml"
        STR_RightPathFile = "C:\temp\ME_Diff\Page21.xml"
    End Sub

    Private Sub CheckAll_Button_Click(sender As Object, e As EventArgs) Handles CheckAll_Button.Click
        Call RemoveAllCheckBoxSubscriptions(Pnl_ComparisonFilters) ' This will suppress the event handlers firing for each checkbox state change in the loop
        Call CheckAll(Pnl_ComparisonFilters)
        Call RemoveAllFilters(Pnl_ComparisonFilters)
        Call AddAllFilters(Pnl_ComparisonFilters)
        Call AddAllCheckBoxSubscriptions(Pnl_ComparisonFilters) ' Reinstate event handlers for checkboxes after the processing is complete
        Call UpdateListBoxFilterViewer()
    End Sub

    Private Sub UncheckAll_Button_Click(sender As Object, e As EventArgs) Handles UncheckAll_Button.Click
        Call RemoveAllCheckBoxSubscriptions(Pnl_ComparisonFilters) ' This will suppress the event handlers firing for each checkbox state change in the loop
        Call UncheckAll(Me.Pnl_ComparisonFilters)
        Call RemoveAllFilters(Pnl_ComparisonFilters)
        Call AddAllFilters(Pnl_ComparisonFilters) ' for consistency sake, we know that each box is unchecked but we check this anyway
        Call AddAllCheckBoxSubscriptions(Pnl_ComparisonFilters) ' Reinstate event handlers for checkboxes after the processing is complete
        Call UpdateListBoxFilterViewer()
    End Sub

    Public Sub CheckBoxStateChangeHandler(sender As Object, e As EventArgs)
        Dim chkbx As CheckBox = TryCast(sender, CheckBox)
        UpdateCheckBoxFilterList(Pnl_ComparisonFilters)
        UpdateListBoxFilterViewer()
    End Sub

    Public Sub UpdateListBoxFilterViewer()
        ListBox1.Items.Clear()
        For Each itm In FilterList
            ListBox1.Items.Add(itm)
        Next
    End Sub
End Class