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

        ' Todo, refactor this code so that the majority of its function is embedded in subs/functions outside of this userform for cleanliness

        ' Intended workflow here is
        ' select hmi xml output file
        ' create an instance of it in memory using the class file generated from the xsd
        ' parse the files with multiple algorithms looking for differences
        ' The following list intends to highlight the kinds of changes we are looking for
        ' - Check page contents match object/group count only
        ' - If content match then check for object data parity
        ' - If content dont match then check the following
        ' - Look for same amount of objects but the order is changed - report the objects moved in output report
        ' - When same amounts of objects cannot be matched list left/right that cannnot be found elsewhere

        ' create a report entry function which will write an entry in a list to be output at the end of the program
        ' create a mutli-file handler that uses the single file protocol to analyse a whole directory
        ' when looking at directory compare, list differences in files
        ' when comparing objects add optional filters for content we know will not be of use in comparison i.e height,width,left,top, font size etc
        ' this is to allow for like comparison between different sizes of HMI
        ' 

        ' Check the file names are populated
        If STR_LeftPathFile = "" Then
            MsgBox("Unable to compare, file name missing")
            Exit Sub
        End If
        If STR_RightPathFile = "" Then
            MsgBox("Unable to compare, file name missing")
            Exit Sub
        End If

        ' get left file name for current working file
        Dim Fname As String
        Dim wrk()
        wrk = Split(STR_LeftPathFile, "\")
        wrk = Split(wrk(wrk.Length - 1), ".")
        Fname = wrk(0)

        ' deserialize the files into program objects
        Dim LeftObj As New gfx
        Dim RightObj As New gfx

        LeftObj = DeserializeObjectFromXMLFile(STR_LeftPathFile)
        RightObj = DeserializeObjectFromXMLFile(STR_RightPathFile)

        If LeftObj.Items.Count = RightObj.Items.Count Then
            ' Item count matches so check the names/types match
            Dim matchflag As Boolean = True
            For a = 0 To LeftObj.Items.Count - 1
                If Not LeftObj.Items(a).GetType = RightObj.Items(a).GetType Then
                    matchflag = False
                End If
                If Not LeftObj.ItemsElementName(a) = RightObj.ItemsElementName(a) Then
                    matchflag = False
                End If
            Next
            If matchflag Then
                ' basic type/name checking has passed so proceed to the next step of the algorithm and compare the elements/properties
                ' if group type of object, then compare the group properties first, then pass the group into the comparegroup function
                ' If any other type of object then pass to compareitemsbytype_obj
                For a = 0 To LeftObj.Items.Count - 1
                    If LeftObj.ItemsElementName(a) = ME_Diff.ItemsChoiceType.group Then
                        ' compare elements of group first
                        Call CompareItemsByType_Obj(LeftObj.Items(a), RightObj.Items(a), Fname, "Root")
                        Dim msg As String
                        Dim gobj As ME_Diff.groupType = LeftObj.Items(a)
                        msg = CompareLikeForLike(LeftObj.Items(a), RightObj.Items(a), Fname, "Root/" & gobj.name)
                        If Not msg = "" Then
                            Dim exgroup As ME_Diff.groupType = LeftObj.Items(a)
                            Call AddListContentMatchGroupException(exgroup.name, Fname, "Root", "Group content match exception encountered - Contents are not comparable, manual check is required - " & msg)
                        End If
                    Else
                        Call CompareItemsByType_Obj(LeftObj.Items(a), RightObj.Items(a), Fname, "Root")
                    End If
                Next
            Else
                ' other cases when top level content doesnt match are handled here
            End If
        End If

        'For a = 0 To LeftObj.ItemsElementName.Count - 1
        '    If LeftObj.ItemsElementName(a) = RightObj.ItemsElementName(a) Then
        '        ' Item count matches but check the names/types match up first
        '        ' Compare like for like elements
        '        Dim msg As String
        '        If LeftObj.ItemsElementName(a) = ME_Diff.ItemsChoiceType.group Then
        '            msg = CompareLikeForLike(LeftObj.Items(a), RightObj.Items(a), Fname, "Root")
        '            If Not msg = "" Then
        '                Dim exgroup As ME_Diff.groupType = LeftObj.Items(a)
        '                Call AddListContentMatchGroupException(exgroup.name, Fname, "Root", "Group content match exception encountered - Contents are not comparable, manual check is required")
        '            End If
        '            MsgBox("boo")
        '        Else
        '            ' not a group so process this separately
        '            MsgBox("not handled yet")
        '        End If

        '    Else
        '        ' object item counts dont match up so process the file differently
        '    End If
        'Next

        'For Each itm As Object In LeftObj.Items
        '    MsgBox(itm.name)
        'Next
        'For Each itm As display

        'Next
        Call outputreport(STR_LeftPathFile)
        MsgBox("done")
    End Sub


    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load
        List_FileCompareContentMatch = New List(Of String)
        List_FileCompareContentNoMatch = New List(Of String)
        List_FolderCompare = New List(Of String)
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
        Call CheckAll(Me.Pnl_ComparisonFilters)
    End Sub

    Private Sub UncheckAll_Button_Click(sender As Object, e As EventArgs) Handles UncheckAll_Button.Click
        Call UncheckAll(Me.Pnl_ComparisonFilters)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim chcar As String
        Dim incar As String
        Dim Check As String = "String"
        Dim Input As String = "StrIng"
        For a = 0 To Len(Check) - 1
            chcar = Check.Substring(a, 1)
            incar = Input.Substring(a, 1)
            If Not chcar = incar Then
                MsgBox("")
            End If
        Next
    End Sub
End Class