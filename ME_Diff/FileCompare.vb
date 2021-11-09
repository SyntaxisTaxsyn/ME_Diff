Imports System.IO
Imports System.Reflection
Public Class FileCompare

    Public Sub New()
        List_FileCompareContentMatch = New List(Of String)
        List_FileCompareContentNoMatch = New List(Of String)
        List_FolderCompare = New List(Of String)
    End Sub

    Public Sub SetFilePaths(ByRef Lfile As String, ByRef Rfile As String)
        STR_LeftPathFile = Lfile
        STR_RightPathFile = Rfile
    End Sub

    Public Function CompareFiles(ByVal TestMode As UnitTestType)

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
            'MsgBox("Unable to compare, file name missing")
            CompareFiles = "Error - Unable to compare, left path file name missing"
            Exit Function
        End If
        If STR_RightPathFile = "" Then
            'MsgBox("Unable to compare, file name missing")
            CompareFiles = "Error - Unable to compare, right path file name missing"
            Exit Function
        End If

        If Not File.Exists(STR_LeftPathFile) Then
            CompareFiles = "Error - Unable to compare, Left file does not exist"
            Exit Function
        End If

        If Not File.Exists(STR_RightPathFile) Then
            CompareFiles = "Error - Unable to compare, Right file does not exist"
            Exit Function
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

        Call CompareDisplayParameters(LeftObj, RightObj, Fname)

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
                ' this would get triggered if the top level item count matches, but the types/names dont
            End If
        Else

            ' Find items in groups that do match and run regular comparison on them
            For Each Litm In LeftObj.Items
                For Each Ritm In RightObj.Items
                    If Litm.name = Ritm.name Then
                        If Litm.GetType = Ritm.GetType Then

                            If Litm.GetType.ToString = "ME_Diff.groupType" Then
                                ' compare elements of group first
                                Call CompareItemsByType_Obj(Litm, Ritm, Fname, "Root")
                                Dim msg As String
                                Dim gobj As ME_Diff.groupType = Litm
                                msg = CompareLikeForLike(Litm, Ritm, Fname, "Root/" & gobj.name)
                                If Not msg = "" Then
                                    Dim exgroup As ME_Diff.groupType = Litm
                                    Call AddListContentMatchGroupException(exgroup.name, Fname, "Root", "Group content match exception encountered - Contents are not comparable, manual check is required - " & msg)
                                End If
                            Else
                                Call CompareItemsByType_Obj(Litm, Ritm, Fname, "Root")
                            End If

                        End If
                    End If
                Next
            Next

            ' Find groups that dont match on both sides
            If LeftObj.Items.Count > RightObj.Items.Count Then
                ' Left has more, find the left items that dont exist
                Dim NoMatch As Boolean = False
                For Each Litm In LeftObj.Items
                    For Each Ritm In RightObj.Items
                        If Ritm.name = Litm.name Then
                            If Ritm.GetType = Litm.GetType Then
                                NoMatch = True
                                Exit For
                            End If
                        End If
                    Next
                    If NoMatch = False Then
                        Call AddListContentMatchGroupException(Litm.name.ToString, Fname, "Root", "Object Type : " & GetTypeFromString(Litm.GetType.ToString) & " - Parameter : Group item mismatch - Left Value = Present, Right Value = None")
                    End If
                    NoMatch = False
                Next
            End If

            If RightObj.Items.Count > LeftObj.Items.Count Then
                ' Right has more, find the right items that dont exist
                Dim NoMatch As Boolean = False
                For Each Ritm In RightObj.Items
                    For Each Litm In LeftObj.Items
                        If Ritm.name = Litm.name Then
                            If Ritm.GetType = Litm.GetType Then
                                NoMatch = True
                                Exit For
                            End If
                        End If
                    Next
                    If NoMatch = False Then
                        Call AddListContentMatchGroupException(Ritm.name.ToString, Fname, "Root", "Object Type : " & GetTypeFromString(Ritm.GetType.ToString) & " - Parameter : Group item mismatch - Left Value = None, Right Value = Present")
                    End If
                    NoMatch = False
                Next
            End If

            'Throw New Exception("Object count no match exception, not implemented yet")


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

        ' Select output type
        Select Case TestMode
            Case UnitTestType.FixedFile
                CompareFiles = OutputToTest_File()
                Exit Function
            Case UnitTestType.DontUse
                Select Case ReportOutputType
                    Case ReportType.Text
                        Call Outputreport(STR_LeftPathFile)
                    Case ReportType.HTML
                        Call OutputHTMLReport(STR_LeftPathFile)
                End Select

                CompareFiles = ""
                Exit Function
            Case Else
                CompareFiles = ""
                Exit Function
        End Select

    End Function

    Public Sub OutputHTMLReport(ByRef path As String)

        Dim ReportObject As New ReportOutput

        ' Get left file name
        Dim LFname As String
        Dim wrk()
        wrk = Split(STR_LeftPathFile, "\")
        wrk = Split(wrk(wrk.Length - 1), ".")
        LFname = wrk(0)

        ' Get right file name
        Dim RFname As String
        wrk = Split(STR_RightPathFile, "\")
        wrk = Split(wrk(wrk.Length - 1), ".")
        RFname = wrk(0)

        ' add the page to report object
        CreateOutputReportPage(ReportObject, List_FileCompareContentMatch, LFname, RFname)

        ' Output the report
        ReportObject.GenerateHTMLReport()

    End Sub

    Public Sub CompareDisplayParameters(ByRef lobj As gfx, ByRef robj As gfx, ByRef Fname As String)
        Dim linfo() As PropertyInfo = lobj.displaySettings.GetType().GetProperties()
        Dim rinfo() As PropertyInfo = robj.displaySettings.GetType().GetProperties()
        For a = 0 To linfo.Count - 1
            If linfo(a).CanRead Then
                Dim str As String = linfo(a).PropertyType().ToString
                If Not str = "ME_Diff.objectBaseType[]" Then ' ignore base class types as these are groups and dont convert to a single value
                    Try
                        If Not linfo(a).GetValue(lobj.displaySettings, Nothing) = rinfo(a).GetValue(robj.displaySettings, Nothing) Then
                            Call AddListContentMatch("", Fname, "Display", GetMEObjectType(lobj), linfo(a).Name, linfo(a).GetValue(lobj.displaySettings, Nothing).ToString, rinfo(a).GetValue(robj.displaySettings, Nothing).ToString)

                        End If
                    Catch ex As Exception
                        ' Im not going to do anything with these exceptions, i just want the code to find any elements that can be compared with the = operator
                    End Try

                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Unit test type to apply to the function call. Should be used to allow for redirection of the function output to allow for
    ''' internal comparison of test results
    ''' </summary>
    Public Enum UnitTestType
        ''' <summary>
        ''' Uses a fixed file in the test project folder with known results
        ''' </summary>
        FixedFile
        ''' <summary>
        ''' Dont use the unit test output redirect feature (sends output directly to the text file and displays as before)
        ''' </summary>
        DontUse
    End Enum

    Public Enum ReportType
        Text
        HTML
    End Enum
End Class
