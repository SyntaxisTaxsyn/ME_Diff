﻿Imports System.Reflection
Imports System.IO
Imports System.Web

Public Module Subroutines

    Public Sub ReadFilterTextFileIntoList()
        Dim wrkstr As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) & "\Filters.txt"
        Using reader As StreamReader = New StreamReader(wrkstr)
            Do
                Dim Temp As String
                Temp = reader.ReadLine
                If Not Temp = "" Then ' Filter empty lines in the text file
                    FilterConfigurationList.Add(Temp)
                End If
            Loop Until reader.EndOfStream
        End Using
    End Sub

    Public Sub WriteFilterListIntoTextFile()
        Dim wrkstr As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) & "\Filters.txt"
        Using writer As StreamWriter = New StreamWriter(wrkstr, False)
            For Each item In FilterConfigurationList
                writer.WriteLine(item)
            Next
        End Using
    End Sub

    ''' <summary>
    ''' Sets the checked value to true for all checkbox controls contained within the panel control passed in by reference
    ''' </summary>
    ''' <param name="Pnl">Panel userform control object to search for checkboxs</param>
    Public Sub CheckAll(ByRef Pnl As Panel)
        For Each ctrl As Control In Pnl.Controls
            If TypeOf ctrl Is CheckBox Then
                Dim chkbx As CheckBox = TryCast(ctrl, CheckBox)
                If chkbx IsNot Nothing Then chkbx.Checked = True
            End If
        Next
    End Sub

    ''' <summary>
    ''' Sets the checked value to false for all checkbox controls contained within the panel control passed in by reference
    ''' </summary>
    ''' <param name="Pnl">Panel userform control object to search for checkboxs</param>
    Public Sub UncheckAll(ByRef Pnl As Panel)
        For Each ctrl As Control In Pnl.Controls
            If TypeOf ctrl Is CheckBox Then
                Dim chkbx As CheckBox = TryCast(ctrl, CheckBox)
                If chkbx IsNot Nothing Then chkbx.Checked = False
            End If
        Next
    End Sub

    ''' <summary>
    ''' do something
    ''' </summary>
    Public Sub UpdateFilters(ByRef Pnl As Panel)

        ' Todo use this routine to update the filter switches used by the comparison algorithm
        ' Add a call to this routine in each check/uncheck all routine
        ' Also create handlers for checkbox checked state changed calls this and updates also

        ' Loop through existing checkboxes
        ' Check if their name exists as an entry in the filter list 
        ' Add it if not
        For Each ctrl As Control In Pnl.Controls
            If TypeOf ctrl Is CheckBox Then
                Dim chkbx As CheckBox = TryCast(ctrl, CheckBox)
                If chkbx IsNot Nothing Then
                    If chkbx.Checked Then
                        Dim ListCheck As Boolean = False
                        For Each itm As String In FilterList.Where(Function(x) x = chkbx.Text)
                            ListCheck = True
                            Exit For
                        Next
                        If Not ListCheck Then
                            FilterList.Add(chkbx.Text)
                        End If
                    End If
                End If
            End If
        Next

    End Sub
    Public Sub AddAllFilters(ByRef Pnl As Panel)
        For Each ctrl As Control In Pnl.Controls
            If TypeOf ctrl Is CheckBox Then
                Dim chkbx As CheckBox = TryCast(ctrl, CheckBox)
                If chkbx IsNot Nothing Then
                    If chkbx.Checked Then
                        FilterList.Add(chkbx.Text)
                    End If
                End If
            End If
        Next
    End Sub
    Public Sub RemoveAllFilters(ByRef Pnl As Panel)
        FilterList.Clear()
    End Sub

    Public Sub RemoveFilter(ByRef Chkbox As CheckBox)
        FilterList.RemoveAt(Chkbox.Text)
    End Sub

    Public Sub RemoveAllCheckBoxSubscriptions(ByRef Pnl As Panel)
        For Each ctrl As Control In Pnl.Controls
            If TypeOf ctrl Is CheckBox Then
                Dim chkbx As CheckBox = TryCast(ctrl, CheckBox)
                Try
                    RemoveHandler chkbx.CheckedChanged, AddressOf Main.CheckBoxStateChangeHandler
                Catch ex As Exception
                    ' If this fails its only because the subscription didnt exist, which im not interested in
                End Try
            End If
        Next
    End Sub

    Public Sub AddAllCheckBoxSubscriptions(ByRef Pnl As Panel)
        For Each ctrl As Control In Pnl.Controls
            If TypeOf ctrl Is CheckBox Then
                Dim chkbx As CheckBox = TryCast(ctrl, CheckBox)
                Try
                    AddHandler chkbx.CheckedChanged, AddressOf Main.CheckBoxStateChangeHandler
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Throw New Exception("Failed to add checkbox subscription event", ex)
                End Try
            End If
        Next
    End Sub

    Public Sub UpdateCheckBoxFilterList(ByRef Pnl As Panel)
        Call RemoveAllFilters(Pnl)
        Call AddAllFilters(Pnl)
    End Sub


    ''' <summary>
    ''' Subroutine that controls adding content match exceptions to the report output list
    ''' </summary>
    ''' <param name="Gnest">Group nesting level. Is a string path containing the preceding groups that make up the path to this object</param>
    ''' <param name="Fname">File name</param>
    ''' <param name="Oname">Object name</param>
    ''' <param name="Otype">Object type</param>
    ''' <param name="ParType">Parameter type</param>
    ''' <param name="lval">Left value</param>
    ''' <param name="rval">Right value</param>
    Public Sub AddListContentMatch(ByVal Gnest As String, ByVal Fname As String, ByVal Oname As String, ByVal Otype As String, ByVal ParType As String, ByVal lval As String, ByVal rval As String)
        Dim msg As String = ""
        msg = msg & "File Name : " & Fname & " - " ' name of the file for finding the related xml
        msg = msg & "Nest Level : " & Gnest & " - " ' string value containing nesting level path i.e Root, Root/Group1, Root/Group1/Group2 etc
        msg = msg & "Object Name : " & Oname & " - " ' User defined object name
        msg = msg & "Object Type : " & Otype & " - " ' ME project defined object type
        msg = msg & "Parameter : " & ParType & " - " ' Name of parameter where difference was detected
        msg = msg & "Left Value = " & lval & ", " ' Left file parameter value
        msg = msg & "Right Value = " & rval ' Right file parameter value
        List_FileCompareContentMatch.Add(msg)
    End Sub


    ''' <summary>
    ''' Loop through the properties of 2 gfx objects and compare them for differences
    ''' </summary>
    ''' <param name="lobj">left object</param>
    ''' <param name="robj">right object</param>
    ''' <param name="fname">Name of the xml file. Is used for producing the output report</param>
    ''' <param name="Gnest">Group Nesting String. Contains a string value that indicates the group nesting path of the object</param>
    Public Sub CompareItemsByType_Obj(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' todo - add filter controls for ignoring parameters
        ' todo - add select case handler for complex types that need to be compared, do an HMI page with every type possible and work on code that detects the changes in each of those special object types 

        'If lobj.GetType.ToString = "ME_Diff.groupType" Then
        '    If robj.GetType.ToString = "ME_Diff.groupType" Then
        '        Dim lGpObj As ME_Diff.groupType = lobj
        '        Dim rGpObj As ME_Diff.groupType = robj
        '        For a = 0 To lGpObj.Items.Count - 1
        '            If lGpObj.name = rGpObj.name Then
        '                MsgBox(lGpObj.name)
        '            End If
        '        Next
        '    End If
        'End If


        Dim linfo() As PropertyInfo = lobj.GetType().GetProperties()
        Dim rinfo() As PropertyInfo = robj.GetType().GetProperties()
        For a = 0 To linfo.Count - 1
            If linfo(a).CanRead Then
                Dim str As String = linfo(a).PropertyType().ToString
                Select Case str
                    Case "ME_Diff.animationsAllMEType"
                        ' call special case handler for dealing with animation property comparison
                        Call CompareAnimationsProperties(lobj.animations, robj.animations, Gnest, fname, lobj.name.ToString)
                    Case "ME_Diff.multistateButtonBaseStateType"
                        MsgBox("") ' Was there a good reason this was left without an attached action
                    Case "ME_Diff.connectionType"
                        MsgBox("") ' Same for here, is it because this type of object cannot be encountered at this level?
                    Case "ME_Diff.animationsVisibilityType"
                        Call CompareAnimationsProperties_AnimateVisibilityAlone(lobj.animations, robj.animations, Gnest, fname, lobj.name.ToString)
                    Case "ME_Diff.animationsAllButRotMEType"
                        CompareAnimationsPropertiesAllButRotType(lobj.animations, robj.animations, Gnest, fname, lobj.name.ToString)
                End Select
                If Not str = "ME_Diff.objectBaseType[]" Then ' ignore base class types as these are groups and dont convert to a single value
                    Try
                        If Not linfo(a).GetValue(lobj, Nothing) = rinfo(a).GetValue(robj, Nothing) Then
                            Call AddListContentMatch(Gnest, fname, lobj.name, GetMEObjectType(lobj), linfo(a).Name, linfo(a).GetValue(lobj, Nothing).ToString, rinfo(a).GetValue(robj, Nothing).ToString)
                            'MsgBox("DifferenceFound" & linfo(a).Name & " - " & linfo(a).GetValue(lobj, Nothing) & " <> " & rinfo(a).Name & " - " & rinfo(a).GetValue(robj, Nothing))
                        End If
                    Catch ex As Exception
                        ' Im not going to do anything with these exceptions, i just want the code to find any elements that can be compared with the = operator
                    End Try

                End If
            End If
        Next

        Call CompareItemsByType_SpecificCases(lobj, robj, fname, Gnest)
        Call CheckConnectionsByType_SpecificCases(lobj, robj, fname, Gnest)

    End Sub

    ''' <summary>
    ''' Case handler for specific comparison of certian object types who contain arrays of other types that cant be picked up using the generic code
    ''' </summary>
    ''' <param name="lobj">Left object</param>
    ''' <param name="robj">Right object</param>
    ''' <param name="fname">File name</param>
    ''' <param name="Gnest">Group nest level</param>
    Public Sub CompareItemsByType_SpecificCases(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Select the special case handling behaviour by object type
        Dim LobjTyp As String = lobj.GetType.ToString
        Select Case LobjTyp
            Case "ME_Diff.momentaryButtonType"
                Call CompareMomentaryButtonMultistateProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.maintainedButtonType"
                Call CompareMaintainedButtonMultistateProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.latchedButtonType"
                Call CompareLatchedButtonMultistateProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.multistateButtonType"
                Call CompareMultistateProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.interlockedButtonType"
                Call CompareInterlockedProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.rampButtonType"
                Call CompareRampProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.numericInputEnableType"
                Call CompareNumericInputEnableProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.stringInputEnableType"
                Call CompareStringInputEnableProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.gotoButtonType"
                Call CompareGotoProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.pveSimpleButtonType"
                Call ComparePVESimpleProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.closeButtonType"
                Call CompareCloseButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.displayListSelectorType"
                Call CompareDisplayListSelectorMultistateProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.multistateIndicatorType"
                Call CompareMultistateIndicatorProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.symbolIndicatorType"
                Call CompareSymbolIndicatorProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.listIndicatorType"
                Call CompareListIndicatorProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.barGraphType"
                Call CompareBarGraphProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.gaugeType"
                Call CompareGaugeProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.linkedButtonType"
                Call CompareLinkedButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.activeXObjectType"
                Call CompareActiveXProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.recipePlusButtonType"
                Call CompareRecipePlusButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.keyButtonType"
                Call CompareKeyButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.timedKeyButtonType"
                Call CompareTimedKeyButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.GenericButtonType"
                Call CompareGenericButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.loginButtonType"
                Call CompareLoginButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.logoutButtonType"
                Call CompareLogoutButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.passwordButtonType"
                Call ComparePasswordButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.controlListSelectorType"
                Call CompareControlListSelectorProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.pilotedControlListSelectorType"
                Call ComparePilotedControlListSelectorProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.languageSwitchButtonType"
                Call CompareLanguageSwitchButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.macroButtonType"
                Call CompareMacroProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.gotoConfigureButtonType"
                Call CompareGotoConfigureModeProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.HMIServiceLinkedButtonType"
                Call CompareHMILinkedServiceButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.HMIServiceButtonType"
                Call CompareHMIServiceButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.ackAllAlarmButtonType"
                Call CompareAckAllAlarmButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.clearAlarmHistButtonType"
                Call CompareClearAlarmHistoryButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.printAlarmHistoryType"
                Call ComparePrintAlarmHistoryButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.printAlarmStatusType"
                Call ComparePrintAlarmStatusButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.simpleButtonType"
                Call CompareSimpleButtonProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.clearAuditTrailButtonType"
                Call CompareClearAuditTrailButtonProperties(lobj, robj, fname, Gnest)
        End Select

    End Sub

    ''' <summary>
    ''' Case handler for specific comparison of certian object types who contain arrays of other types that cant be picked up using the generic code
    ''' </summary>
    ''' <param name="lobj">Left object</param>
    ''' <param name="robj">Right object</param>
    ''' <param name="fname">File name</param>
    ''' <param name="Gnest">Group nest level</param>
    Public Sub CheckConnectionsByType_SpecificCases(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Select the special case handling behaviour by object type
        Dim LobjTyp As String = lobj.GetType.ToString
        Select Case LobjTyp
            'Case "ME_Diff.multistateButtonType"
            '    Call CompareMultistateProperties(lobj, robj, fname, Gnest)
            Case "ME_Diff.momentaryButtonType"
                Call CheckMomentaryPBConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.maintainedButtonType"
                Call CheckMaintainedPBConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.latchedButtonType"
                Call CheckLatchedPBConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.multistateButtonType"
                Call CheckMultistatePBConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.interlockedButtonType"
                Call CheckInterlockedPBConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.rampButtonType"
                Call CheckRampPBConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.numericDisplayType"
                Call CheckNumericDisplayConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.numericInputEnableType"
                Call CheckNumericInputEnableConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.numericInputCursorPointType"
                Call CheckNumericInputCursorConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.stringDisplayType"
                Call CheckStringDisplayConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.stringInputEnableType"
                Call CheckStringInputEnableConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.gotoButtonType"
                Call CheckGotoConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.closeButtonType"
                Call CheckCloseButtonConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.multistateIndicatorType"
                Call CheckMultistateIndicatorConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.symbolIndicatorType"
                Call CheckSymbolIndicatorConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.listIndicatorType"
                Call CheckListIndicatorConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.barGraphType"
                Call CheckBarGraphConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.gaugeType"
                Call CheckGaugeConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.activeXObjectType"
                Call CheckActiveXConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.controlListSelectorType"
                Call CheckControlListSelectorConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.pilotedControlListSelectorType"
                Call CheckPilotedControlListSelectorConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.localMessageDisplayType"
                Call CheckLocalMessageDisplayConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.macroButtonType"
                Call CheckMacroConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.alarmListType"
                Call CheckAlarmListConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.alarmBannerType"
                Call CheckAlarmBannerConnections(lobj, robj, fname, Gnest)
            Case "ME_Diff.alarmStatusListType"
                Call CheckAlarmStatusListConnections(lobj, robj, fname, Gnest)
        End Select

    End Sub

    Public Sub CheckMomentaryPBConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.momentaryButtonType = lobj
        Dim RMSB As ME_Diff.momentaryButtonType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Momentary Button - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Momentary Button - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Momentary Button - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If




    End Sub

    Public Sub CheckMultistatePBConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.multistateButtonType = lobj
        Dim RMSB As ME_Diff.multistateButtonType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Multistate Button - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Multistate Button - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Multistate Button - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If




    End Sub

    Public Sub CheckLatchedPBConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.latchedButtonType = lobj
        Dim RMSB As ME_Diff.latchedButtonType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Latched Button - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Latched Button - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Latched Button - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If




    End Sub

    Public Sub CheckInterlockedPBConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.interlockedButtonType = lobj
        Dim RMSB As ME_Diff.interlockedButtonType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Interlocked Button - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Interlocked Button - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Interlocked Button - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If




    End Sub


    Public Sub CheckRampPBConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.rampButtonType = lobj
        Dim RMSB As ME_Diff.rampButtonType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Ramp Button - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Ramp Button - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Ramp Button - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If




    End Sub

    Public Sub CheckNumericDisplayConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.numericDisplayType = lobj
        Dim RMSB As ME_Diff.numericDisplayType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Numeric Display - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Numeric Display - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Numeric Display - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If




    End Sub

    Public Sub CheckNumericInputEnableConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.numericInputEnableType = lobj
        Dim RMSB As ME_Diff.numericInputEnableType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Numeric Input Enable - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Numeric Input Enable - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Numeric Input Enable - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If




    End Sub

    Public Sub CheckNumericInputCursorConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.numericInputCursorPointType = lobj
        Dim RMSB As ME_Diff.numericInputCursorPointType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Numeric Input Cursor - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Numeric Input Cursor - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Numeric Input Cursor - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If




    End Sub

    Public Sub CheckStringDisplayConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.stringDisplayType = lobj
        Dim RMSB As ME_Diff.stringDisplayType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "String Display - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "String Display - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "String Display - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckStringInputEnableConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.stringInputEnableType = lobj
        Dim RMSB As ME_Diff.stringInputEnableType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "String Input Enable - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "String Input Enable - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "String Input Enable - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckGotoConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.gotoButtonType = lobj
        Dim RMSB As ME_Diff.gotoButtonType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Goto Button - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Goto Button - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Goto Button - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub


    Public Sub CheckCloseButtonConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.closeButtonType = lobj
        Dim RMSB As ME_Diff.closeButtonType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Close Button - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Close Button - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Close Button - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckMultistateIndicatorConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.multistateIndicatorType = lobj
        Dim RMSB As ME_Diff.multistateIndicatorType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Multistate Indicator - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Multistate Indicator - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Multistate Indicator - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckSymbolIndicatorConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.symbolIndicatorType = lobj
        Dim RMSB As ME_Diff.symbolIndicatorType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Symbol Indicator - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Symbol Indicator - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Symbol Indicator - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckListIndicatorConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.listIndicatorType = lobj
        Dim RMSB As ME_Diff.listIndicatorType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "List Indicator - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "List Indicator - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "List Indicator - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckBarGraphConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.barGraphType = lobj
        Dim RMSB As ME_Diff.barGraphType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Bar Graph - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Bar Graph - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Bar Graph - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckGaugeConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.gaugeType = lobj
        Dim RMSB As ME_Diff.gaugeType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Gauge - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Gauge - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Gauge - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckControlListSelectorConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.controlListSelectorType = lobj
        Dim RMSB As ME_Diff.controlListSelectorType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Control List Selector - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Control List Selector - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Control List Selector - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckPilotedControlListSelectorConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.pilotedControlListSelectorType = lobj
        Dim RMSB As ME_Diff.pilotedControlListSelectorType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Piloted Control List Selector - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Piloted Control List Selector - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Piloted Control List Selector - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckLocalMessageDisplayConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.localMessageDisplayType = lobj
        Dim RMSB As ME_Diff.localMessageDisplayType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Local Message Display - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Local Message Display - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Local Message Display - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckMacroConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.macroButtonType = lobj
        Dim RMSB As ME_Diff.macroButtonType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Macro - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Macro - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Macro - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckAlarmListConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.alarmListType = lobj
        Dim RMSB As ME_Diff.alarmListType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Alarm List - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Alarm List - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Alarm List - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckAlarmBannerConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.alarmBannerType = lobj
        Dim RMSB As ME_Diff.alarmBannerType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Alarm Banner - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Alarm Banner - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Alarm Banner - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckAlarmStatusListConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.alarmStatusListType = lobj
        Dim RMSB As ME_Diff.alarmStatusListType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Alarm Status List - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Alarm Status List - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Alarm Status List - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckActiveXConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.activeXObjectType = lobj
        Dim RMSB As ME_Diff.activeXObjectType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "ActiveX - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            Else
                                ' left defined, right nothing
                                If LConn.optionalExpression IsNot Nothing Then
                                    Desc = "optionalExpression"
                                    Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Connection " & a & " - " & Desc,
                                                                 "defined",
                                                                 "nothing")
                                End If
                            End If
                        Else
                            ' left nothing, right defined
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             "nothing",
                                                             "defined")
                            End If
                        End If
                    Next
                End If
            Else
                ' left defined, right not
                If LMSB.connections IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "ActiveX - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If
            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' right defined left not
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "ActiveX - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If

    End Sub

    Public Sub CheckMaintainedPBConnections(ByRef lobj As Object,
                                           ByRef robj As Object,
                                           ByRef fname As String,
                                           ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.maintainedButtonType = lobj
        Dim RMSB As ME_Diff.maintainedButtonType = robj

        If LMSB.connections IsNot Nothing Then
            If RMSB.connections IsNot Nothing Then
                ' both sides have connections defined
                If Not LMSB.connections.Count = RMSB.connections.Count Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Maintained Button - Connection Count Mismatch",
                                             LMSB.connections.Count,
                                             RMSB.connections.Count)
                Else
                    For a = 0 To LMSB.connections.Count - 1

                        ' Declare reuseable loop variables
                        Dim MEObjType As String = GetMEObjectType(lobj)
                        Dim LConn As ME_Diff.connectionType = LMSB.connections(a)
                        Dim RConn As ME_Diff.connectionType = RMSB.connections(a)
                        Dim Left As String = ""
                        Dim Right As String = ""
                        Dim Desc As String = ""

                        ' expression
                        Desc = "expression"
                        Left = LConn.expression.ToString
                        Right = RConn.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' name
                        Desc = "name"
                        Left = LConn.name.ToString
                        Right = RConn.name.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Connection " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' optionalExpression
                        If LConn.optionalExpression IsNot Nothing Then
                            If RConn.optionalExpression IsNot Nothing Then
                                Desc = "optionalExpression"
                                Left = LConn.optionalExpression.ToString
                                Right = RConn.optionalExpression.ToString
                                If Not Left = Right Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Connection " & a & " - " & Desc,
                                                             Left,
                                                             Right)
                                End If
                            End If
                        End If


                    Next
                End If
            Else
                If LMSB.connections IsNot Nothing Then
                    ' right has nothing, left does
                    Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Maintained Button - Connection Count Mismatch",
                                     "defined",
                                     "nothing")
                End If

            End If
        Else
            If RMSB.connections IsNot Nothing Then
                ' Left has no connections defined, right does
                Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Maintained Button - Connection Count Mismatch",
                                     "nothing",
                                     "defined")
            End If
        End If



    End Sub

    Public Sub CompareMomentaryButtonMultistateProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.momentaryButtonType = lobj
        Dim RMSB As ME_Diff.momentaryButtonType = robj

        If Not LMSB.states.Count = RMSB.states.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "MultistateButton - State Count Mismatch",
                                     LMSB.states.Count,
                                     RMSB.states.Count)
        Else

            For a = 0 To LMSB.states.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.pushButtonBaseStateValueType = LMSB.states(a)
                Dim Rstate As ME_Diff.pushButtonBaseStateValueType = RMSB.states(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""

                ' backColor
                Desc = "backColor"
                Left = LState.backColor.ToString
                Right = Rstate.backColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blink
                Desc = "blink"
                Left = LState.blink.ToString
                Right = Rstate.blink.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blinkSpecified
                Desc = "blinkSpecified"
                Left = LState.blinkSpecified.ToString
                Right = Rstate.blinkSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' borderColor
                Desc = "borderColor"
                Left = LState.borderColor.ToString
                Right = Rstate.borderColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' endColor
                Desc = "endColor"
                Left = LState.endColor.ToString
                Right = Rstate.endColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirection
                Desc = "gradientDirection"
                Left = LState.gradientDirection.ToString
                Right = Rstate.gradientDirection.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirectionSpecified
                Desc = "gradientDirectionSpecified"
                Left = LState.gradientDirectionSpecified.ToString
                Right = Rstate.gradientDirectionSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyle
                Desc = "gradientShadingStyle"
                Left = LState.gradientShadingStyle.ToString
                Right = Rstate.gradientShadingStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyleSpecified
                Desc = "gradientShadingStyleSpecified"
                Left = LState.gradientShadingStyleSpecified.ToString
                Right = Rstate.gradientShadingStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStop
                Desc = "gradientStop"
                Left = LState.gradientStop.ToString
                Right = Rstate.gradientStop.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStopSpecified
                Desc = "gradientStopSpecified"
                Left = LState.gradientStopSpecified.ToString
                Right = Rstate.gradientStopSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternColor
                Desc = "patternColor"
                Left = LState.patternColor.ToString
                Right = Rstate.patternColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyle
                Desc = "patternStyle"
                Left = LState.patternStyle.ToString
                Right = Rstate.patternStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyleSpecified
                Desc = "patternStyleSpecified"
                Left = LState.patternStyleSpecified.ToString
                Right = Rstate.patternStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' stateId
                Desc = "stateId"
                Left = LState.stateId.ToString
                Right = Rstate.stateId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' value
                Desc = "value"
                If LState.value IsNot Nothing Then
                    If Rstate.value IsNot Nothing Then
                        Left = LState.value.ToString
                        Right = Rstate.value.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If
                    Else
                        If LState.value IsNot Nothing Then
                            ' Left is set to value but right is not
                            Left = LState.value.ToString
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     Left,
                                                     "Nothing")
                        End If
                    End If

                Else
                    If Rstate.value IsNot Nothing Then
                        Right = Rstate.value.ToString
                        Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     "Nothing",
                                                     Right)
                    End If
                End If



                ' Handle special datatypes within the multistate indicator
                Call Compare_BasicCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                Call Compare_BasicImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub

    Public Sub CompareDisplayListSelectorMultistateProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.displayListSelectorType = lobj
        Dim RMSB As ME_Diff.displayListSelectorType = robj

        If Not LMSB.states.Count = RMSB.states.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Display List Selector - State Count Mismatch",
                                     LMSB.states.Count,
                                     RMSB.states.Count)
        Else

            For a = 0 To LMSB.states.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.displayListSelectorStateType = LMSB.states(a)
                Dim Rstate As ME_Diff.displayListSelectorStateType = RMSB.states(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""

                ' caption
                Desc = "caption"
                Left = LState.caption.ToString
                Right = Rstate.caption.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' display
                Desc = "display"
                Left = LState.display.ToString
                Right = Rstate.display.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' displayLeftPosition
                Desc = "displayLeftPosition"
                Left = LState.displayLeftPosition.ToString
                Right = Rstate.displayLeftPosition.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' displayLeftPositionSpecified
                Desc = "displayLeftPositionSpecified"
                Left = LState.displayLeftPositionSpecified.ToString
                Right = Rstate.displayLeftPositionSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' displayPosition
                Desc = "displayPosition"
                Left = LState.displayPosition.ToString
                Right = Rstate.displayPosition.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' displayPositionSpecified
                Desc = "displayPositionSpecified"
                Left = LState.displayPositionSpecified.ToString
                Right = Rstate.displayPositionSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' displayTopPosition
                Desc = "displayTopPosition"
                Left = LState.displayTopPosition.ToString
                Right = Rstate.displayTopPosition.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' displayTopPositionSpecified
                Desc = "displayTopPositionSpecified"
                Left = LState.displayTopPositionSpecified.ToString
                Right = Rstate.displayTopPositionSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' parameterFile
                Desc = "parameterFile"
                Left = LState.parameterFile.ToString
                Right = Rstate.parameterFile.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' parameterList
                Desc = "parameterList"
                Left = LState.parameterList.ToString
                Right = Rstate.parameterList.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' parameterType
                Desc = "parameterType"
                Left = LState.parameterType.ToString
                Right = Rstate.parameterType.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' parameterTypeSpecified
                Desc = "parameterTypeSpecified"
                Left = LState.parameterTypeSpecified.ToString
                Right = Rstate.parameterTypeSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' stateId
                Desc = "stateId"
                Left = LState.stateId.ToString
                Right = Rstate.stateId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If





                ' Handle special datatypes within the multistate indicator
                Call Compare_DisplayListSelectorCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                'Call Compare_BasicImageType(lobj, robj, fname, Gnest, LState, Rstate, a) ' this object type doesnt have image as a subtype

            Next

        End If

    End Sub

    Public Sub CompareMultistateIndicatorProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.multistateIndicatorType = lobj
        Dim RMSB As ME_Diff.multistateIndicatorType = robj

        If Not LMSB.states.Count = RMSB.states.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Multistate Indicator - State Count Mismatch",
                                     LMSB.states.Count,
                                     RMSB.states.Count)
        Else

            For a = 0 To LMSB.states.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.basicStateType = LMSB.states(a)
                Dim Rstate As ME_Diff.basicStateType = RMSB.states(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""

                ' backColor
                Desc = "backColor"
                Left = LState.backColor.ToString
                Right = Rstate.backColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blink
                Desc = "blink"
                Left = LState.blink.ToString
                Right = Rstate.blink.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blinkSpecified
                Desc = "blinkSpecified"
                Left = LState.blinkSpecified.ToString
                Right = Rstate.blinkSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' borderColor
                Desc = "borderColor"
                Left = LState.borderColor.ToString
                Right = Rstate.borderColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                '' caption
                'Desc = "caption"
                'Left = LState.caption.ToString
                'Right = Rstate.caption.ToString
                'If Not Left = Right Then
                '    Call AddListContentMatch(Gnest,
                '                             fname,
                '                             lobj.name,
                '                             GetMEObjectType(lobj),
                '                             "State " & a & " - " & Desc,
                '                             Left,
                '                             Right)
                'End If

                ' endColor
                Desc = "endColor"
                Left = LState.endColor.ToString
                Right = Rstate.endColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirection
                Desc = "gradientDirection"
                Left = LState.gradientDirection.ToString
                Right = Rstate.gradientDirection.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirectionSpecified
                Desc = "gradientDirectionSpecified"
                Left = LState.gradientDirectionSpecified.ToString
                Right = Rstate.gradientDirectionSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyle
                Desc = "gradientShadingStyle"
                Left = LState.gradientShadingStyle.ToString
                Right = Rstate.gradientShadingStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyleSpecified
                Desc = "gradientShadingStyleSpecified"
                Left = LState.gradientShadingStyleSpecified.ToString
                Right = Rstate.gradientShadingStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStop
                Desc = "gradientStop"
                Left = LState.gradientStop.ToString
                Right = Rstate.gradientStop.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStopSpecified
                Desc = "gradientStopSpecified"
                Left = LState.gradientStopSpecified.ToString
                Right = Rstate.gradientStopSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternColor
                Desc = "patternColor"
                Left = LState.patternColor.ToString
                Right = Rstate.patternColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyle
                Desc = "patternStyle"
                Left = LState.patternStyle.ToString
                Right = Rstate.patternStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyleSpecified
                Desc = "patternStyleSpecified"
                Left = LState.patternStyleSpecified.ToString
                Right = Rstate.patternStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' stateId
                Desc = "stateId"
                Left = LState.stateId.ToString
                Right = Rstate.stateId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                If LState.value IsNot Nothing Then
                    If Rstate.value IsNot Nothing Then
                        ' value
                        Desc = "value"
                        Left = LState.value.ToString
                        Right = Rstate.value.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If
                    End If
                End If


                ' Handle special datatypes within the multistate indicator
                Call Compare_BasicCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                Call Compare_BasicImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub

    Public Sub CompareSymbolIndicatorProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.symbolIndicatorType = lobj
        Dim RMSB As ME_Diff.symbolIndicatorType = robj

        If Not LMSB.states.Count = RMSB.states.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Multistate Indicator - State Count Mismatch",
                                     LMSB.states.Count,
                                     RMSB.states.Count)
        Else

            For a = 0 To LMSB.states.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.symbolIndicatorStateType = LMSB.states(a)
                Dim Rstate As ME_Diff.symbolIndicatorStateType = RMSB.states(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""

                ' borderColor
                Desc = "borderColor"
                Left = LState.borderColor.ToString
                Right = Rstate.borderColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' stateId
                Desc = "stateId"
                Left = LState.stateId.ToString
                Right = Rstate.stateId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' value
                If LState.value IsNot Nothing Then ' handle exception case where the error state has no value attribute
                    If Rstate.value IsNot Nothing Then
                        Desc = "value"
                        Left = LState.value.ToString
                        Right = Rstate.value.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If
                    End If
                End If


                ' Handle special datatypes within the multistate indicator
                'Call Compare_BasicCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                Call Compare_SymbolIndicatorImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub

    Public Sub CompareListIndicatorProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.listIndicatorType = lobj
        Dim RMSB As ME_Diff.listIndicatorType = robj

        If Not LMSB.states.Count = RMSB.states.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Multistate Indicator - State Count Mismatch",
                                     LMSB.states.Count,
                                     RMSB.states.Count)
        Else

            For a = 0 To LMSB.states.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.listIndicatorStateType = LMSB.states(a)
                Dim Rstate As ME_Diff.listIndicatorStateType = RMSB.states(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""

                ' stateId
                Desc = "stateId"
                Left = LState.stateId.ToString
                Right = Rstate.stateId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' value
                If LState.value IsNot Nothing Then ' handle exception case where the error state has no value attribute
                    If Rstate.value IsNot Nothing Then
                        Desc = "value"
                        Left = LState.value.ToString
                        Right = Rstate.value.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If
                    End If
                End If


                ' Handle special datatypes within the multistate indicator
                Call Compare_ListIndicatorCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                'Call Compare_SymbolIndicatorImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub
    Public Sub CompareMaintainedButtonMultistateProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.maintainedButtonType = lobj
        Dim RMSB As ME_Diff.maintainedButtonType = robj

        If Not LMSB.states.Count = RMSB.states.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "MultistateButton - State Count Mismatch",
                                     LMSB.states.Count,
                                     RMSB.states.Count)
        Else

            For a = 0 To LMSB.states.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.pushButtonBaseStateValueType = LMSB.states(a)
                Dim Rstate As ME_Diff.pushButtonBaseStateValueType = RMSB.states(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""

                ' backColor
                Desc = "backColor"
                Left = LState.backColor.ToString
                Right = Rstate.backColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blink
                Desc = "blink"
                Left = LState.blink.ToString
                Right = Rstate.blink.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blinkSpecified
                Desc = "blinkSpecified"
                Left = LState.blinkSpecified.ToString
                Right = Rstate.blinkSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' borderColor
                Desc = "borderColor"
                Left = LState.borderColor.ToString
                Right = Rstate.borderColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' endColor
                Desc = "endColor"
                Left = LState.endColor.ToString
                Right = Rstate.endColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirection
                Desc = "gradientDirection"
                Left = LState.gradientDirection.ToString
                Right = Rstate.gradientDirection.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirectionSpecified
                Desc = "gradientDirectionSpecified"
                Left = LState.gradientDirectionSpecified.ToString
                Right = Rstate.gradientDirectionSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyle
                Desc = "gradientShadingStyle"
                Left = LState.gradientShadingStyle.ToString
                Right = Rstate.gradientShadingStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyleSpecified
                Desc = "gradientShadingStyleSpecified"
                Left = LState.gradientShadingStyleSpecified.ToString
                Right = Rstate.gradientShadingStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStop
                Desc = "gradientStop"
                Left = LState.gradientStop.ToString
                Right = Rstate.gradientStop.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStopSpecified
                Desc = "gradientStopSpecified"
                Left = LState.gradientStopSpecified.ToString
                Right = Rstate.gradientStopSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternColor
                Desc = "patternColor"
                Left = LState.patternColor.ToString
                Right = Rstate.patternColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyle
                Desc = "patternStyle"
                Left = LState.patternStyle.ToString
                Right = Rstate.patternStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyleSpecified
                Desc = "patternStyleSpecified"
                Left = LState.patternStyleSpecified.ToString
                Right = Rstate.patternStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' stateId
                Desc = "stateId"
                Left = LState.stateId.ToString
                Right = Rstate.stateId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' value
                Desc = "value"
                If LState.value IsNot Nothing Then
                    If Rstate.value IsNot Nothing Then
                        Left = LState.value.ToString
                        Right = Rstate.value.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If
                    Else
                        If LState.value IsNot Nothing Then
                            ' Left is set to value but right is not
                            Left = LState.value.ToString
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     Left,
                                                     "Nothing")
                        End If
                    End If

                Else
                    If Rstate.value IsNot Nothing Then
                        Right = Rstate.value.ToString
                        Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     "Nothing",
                                                     Right)
                    End If
                End If



                ' Handle special datatypes within the multistate indicator
                Call Compare_BasicCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                Call Compare_BasicImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub

    Public Sub CompareBarGraphProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.barGraphType = lobj
        Dim RMSB As ME_Diff.barGraphType = robj

        If Not LMSB.threshold.Count = RMSB.threshold.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Bar graph - Threshold Count Mismatch",
                                     LMSB.threshold.Count,
                                     RMSB.threshold.Count)
        Else

            For a = 0 To LMSB.threshold.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.thresholdType = LMSB.threshold(a)
                Dim Rstate As ME_Diff.thresholdType = RMSB.threshold(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""

                ' blink
                Desc = "blink"
                Left = LState.blink.ToString
                Right = Rstate.blink.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blinkSpecified
                Desc = "blinkSpecified"
                Left = LState.blinkSpecified.ToString
                Right = Rstate.blinkSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' fillColor
                Desc = "fillColor"
                Left = LState.fillColor.ToString
                Right = Rstate.fillColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' thresholdIndex
                Desc = "thresholdIndex"
                Left = LState.thresholdIndex.ToString
                Right = Rstate.thresholdIndex.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' thresholdIndexSpecified
                Desc = "thresholdIndexSpecified"
                Left = LState.thresholdIndexSpecified.ToString
                Right = Rstate.thresholdIndexSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' thresholdValue
                Desc = "thresholdValue"
                Left = LState.thresholdValue.ToString
                Right = Rstate.thresholdValue.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' thresholdValueSpecified
                Desc = "thresholdValueSpecified"
                Left = LState.thresholdValueSpecified.ToString
                Right = Rstate.thresholdValueSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' Handle special datatypes within the multistate indicator
                'Call Compare_ListIndicatorCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                'Call Compare_SymbolIndicatorImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub

    Public Sub CompareGaugeProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.gaugeType = lobj
        Dim RMSB As ME_Diff.gaugeType = robj

        If Not LMSB.threshold.Count = RMSB.threshold.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Gauge - Threshold Count Mismatch",
                                     LMSB.threshold.Count,
                                     RMSB.threshold.Count)
        Else

            For a = 0 To LMSB.threshold.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.thresholdType = LMSB.threshold(a)
                Dim Rstate As ME_Diff.thresholdType = RMSB.threshold(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""

                ' blink
                Desc = "blink"
                Left = LState.blink.ToString
                Right = Rstate.blink.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blinkSpecified
                Desc = "blinkSpecified"
                Left = LState.blinkSpecified.ToString
                Right = Rstate.blinkSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' fillColor
                Desc = "fillColor"
                Left = LState.fillColor.ToString
                Right = Rstate.fillColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' thresholdIndex
                Desc = "thresholdIndex"
                Left = LState.thresholdIndex.ToString
                Right = Rstate.thresholdIndex.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' thresholdIndexSpecified
                Desc = "thresholdIndexSpecified"
                Left = LState.thresholdIndexSpecified.ToString
                Right = Rstate.thresholdIndexSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' thresholdValue
                Desc = "thresholdValue"
                Left = LState.thresholdValue.ToString
                Right = Rstate.thresholdValue.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' thresholdValueSpecified
                Desc = "thresholdValueSpecified"
                Left = LState.thresholdValueSpecified.ToString
                Right = Rstate.thresholdValueSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "Threshold " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' Handle special datatypes within the multistate indicator
                'Call Compare_ListIndicatorCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                'Call Compare_SymbolIndicatorImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub

    Public Sub CompareLatchedButtonMultistateProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.latchedButtonType = lobj
        Dim RMSB As ME_Diff.latchedButtonType = robj

        If Not LMSB.states.Count = RMSB.states.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "LatchedButton - State Count Mismatch",
                                     LMSB.states.Count,
                                     RMSB.states.Count)
        Else

            For a = 0 To LMSB.states.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.pushButtonBaseStateValueType = LMSB.states(a)
                Dim Rstate As ME_Diff.pushButtonBaseStateValueType = RMSB.states(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""

                ' backColor
                Desc = "backColor"
                Left = LState.backColor.ToString
                Right = Rstate.backColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blink
                Desc = "blink"
                Left = LState.blink.ToString
                Right = Rstate.blink.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blinkSpecified
                Desc = "blinkSpecified"
                Left = LState.blinkSpecified.ToString
                Right = Rstate.blinkSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' borderColor
                Desc = "borderColor"
                Left = LState.borderColor.ToString
                Right = Rstate.borderColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' endColor
                Desc = "endColor"
                Left = LState.endColor.ToString
                Right = Rstate.endColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirection
                Desc = "gradientDirection"
                Left = LState.gradientDirection.ToString
                Right = Rstate.gradientDirection.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirectionSpecified
                Desc = "gradientDirectionSpecified"
                Left = LState.gradientDirectionSpecified.ToString
                Right = Rstate.gradientDirectionSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyle
                Desc = "gradientShadingStyle"
                Left = LState.gradientShadingStyle.ToString
                Right = Rstate.gradientShadingStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyleSpecified
                Desc = "gradientShadingStyleSpecified"
                Left = LState.gradientShadingStyleSpecified.ToString
                Right = Rstate.gradientShadingStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStop
                Desc = "gradientStop"
                Left = LState.gradientStop.ToString
                Right = Rstate.gradientStop.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStopSpecified
                Desc = "gradientStopSpecified"
                Left = LState.gradientStopSpecified.ToString
                Right = Rstate.gradientStopSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternColor
                Desc = "patternColor"
                Left = LState.patternColor.ToString
                Right = Rstate.patternColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyle
                Desc = "patternStyle"
                Left = LState.patternStyle.ToString
                Right = Rstate.patternStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyleSpecified
                Desc = "patternStyleSpecified"
                Left = LState.patternStyleSpecified.ToString
                Right = Rstate.patternStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' stateId
                Desc = "stateId"
                Left = LState.stateId.ToString
                Right = Rstate.stateId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' value
                Desc = "value"
                If LState.value IsNot Nothing Then
                    If Rstate.value IsNot Nothing Then
                        Left = LState.value.ToString
                        Right = Rstate.value.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If
                    Else
                        If LState.value IsNot Nothing Then
                            ' Left is set to value but right is not
                            Left = LState.value.ToString
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     Left,
                                                     "Nothing")
                        End If
                    End If

                Else
                    If Rstate.value IsNot Nothing Then
                        Right = Rstate.value.ToString
                        Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     "Nothing",
                                                     Right)
                    End If
                End If



                ' Handle special datatypes within the multistate indicator
                Call Compare_BasicCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                Call Compare_BasicImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub

    Public Sub CompareCloseButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.closeButtonType = lobj
        Dim RMSB As ME_Diff.closeButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub ComparePVESimpleProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.pveSimpleButtonType = lobj
        Dim RMSB As ME_Diff.pveSimpleButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareLinkedButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.linkedButtonType = lobj
        Dim RMSB As ME_Diff.linkedButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareRecipePlusButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.recipePlusButtonType = lobj
        Dim RMSB As ME_Diff.recipePlusButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareKeyButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.keyButtonType = lobj
        Dim RMSB As ME_Diff.keyButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareTimedKeyButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.timedKeyButtonType = lobj
        Dim RMSB As ME_Diff.timedKeyButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareGenericButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.GenericButtonType = lobj
        Dim RMSB As ME_Diff.GenericButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareLoginButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.loginButtonType = lobj
        Dim RMSB As ME_Diff.loginButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareLogoutButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.logoutButtonType = lobj
        Dim RMSB As ME_Diff.logoutButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub ComparePasswordButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.passwordButtonType = lobj
        Dim RMSB As ME_Diff.passwordButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareLanguageSwitchButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.languageSwitchButtonType = lobj
        Dim RMSB As ME_Diff.languageSwitchButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareMacroProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.macroButtonType = lobj
        Dim RMSB As ME_Diff.macroButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareGotoConfigureModeProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.gotoConfigureButtonType = lobj
        Dim RMSB As ME_Diff.gotoConfigureButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareHMILinkedServiceButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.HMIServiceLinkedButtonType = lobj
        Dim RMSB As ME_Diff.HMIServiceLinkedButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareHMIServiceButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.HMIServiceButtonType = lobj
        Dim RMSB As ME_Diff.HMIServiceButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareAckAllAlarmButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.ackAllAlarmButtonType = lobj
        Dim RMSB As ME_Diff.ackAllAlarmButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareClearAlarmHistoryButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.clearAlarmHistButtonType = lobj
        Dim RMSB As ME_Diff.clearAlarmHistButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub ComparePrintAlarmHistoryButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.printAlarmHistoryType = lobj
        Dim RMSB As ME_Diff.printAlarmHistoryType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub ComparePrintAlarmStatusButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.printAlarmStatusType = lobj
        Dim RMSB As ME_Diff.printAlarmStatusType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareSimpleButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.simpleButtonType = lobj
        Dim RMSB As ME_Diff.simpleButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareClearAuditTrailButtonProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.clearAuditTrailButtonType = lobj
        Dim RMSB As ME_Diff.clearAuditTrailButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub


    Public Sub CompareControlListSelectorProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.controlListSelectorType = lobj
        Dim RMSB As ME_Diff.controlListSelectorType = robj

        If Not LMSB.states.Count = RMSB.states.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Control List Selector - State Count Mismatch",
                                     LMSB.states.Count,
                                     RMSB.states.Count)
        Else

            For a = 0 To LMSB.states.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.controlListSelectorStateType = LMSB.states(a)
                Dim Rstate As ME_Diff.controlListSelectorStateType = RMSB.states(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""

                ' stateId
                Desc = "stateId"
                Left = LState.stateId.ToString
                Right = Rstate.stateId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' value
                Desc = "value"
                Left = LState.value.ToString
                Right = Rstate.value.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' Handle special datatypes within the multistate indicator
                Call Compare_ControlListCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                'Call Compare_BasicImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub

    Public Sub ComparePilotedControlListSelectorProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.pilotedControlListSelectorType = lobj
        Dim RMSB As ME_Diff.pilotedControlListSelectorType = robj

        If Not LMSB.states.Count = RMSB.states.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "Piloted Control List Selector - State Count Mismatch",
                                     LMSB.states.Count,
                                     RMSB.states.Count)
        Else

            For a = 0 To LMSB.states.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.pilotedControlListSelectorStateType = LMSB.states(a)
                Dim Rstate As ME_Diff.pilotedControlListSelectorStateType = RMSB.states(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""

                ' access
                Desc = "access"
                Left = LState.access.ToString
                Right = Rstate.access.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' accessSpecified
                Desc = "accessSpecified"
                Left = LState.accessSpecified.ToString
                Right = Rstate.accessSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' stateId
                Desc = "stateId"
                Left = LState.stateId.ToString
                Right = Rstate.stateId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' value
                Desc = "value"
                Left = LState.value.ToString
                Right = Rstate.value.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If


                ' Handle special datatypes within the multistate indicator
                Call Compare_ControlListCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                'Call Compare_BasicImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub

    Public Sub CompareGotoProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.gotoButtonType = lobj
        Dim RMSB As ME_Diff.gotoButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareInterlockedProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.interlockedButtonType = lobj
        Dim RMSB As ME_Diff.interlockedButtonType = robj

        If Not LMSB.states.Count = RMSB.states.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "InterlockedButton - State Count Mismatch",
                                     LMSB.states.Count,
                                     RMSB.states.Count)
        Else
            For a = 0 To LMSB.states.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.interlockedButtonBaseStateType = LMSB.states(a)
                Dim Rstate As ME_Diff.interlockedButtonBaseStateType = RMSB.states(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""


                ' backColor
                Desc = "backColor"
                Left = LState.backColor.ToString
                Right = Rstate.backColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blink
                Desc = "blink"
                Left = LState.blink.ToString
                Right = Rstate.blink.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blinkSpecified
                Desc = "blinkSpecified"
                Left = LState.blinkSpecified.ToString
                Right = Rstate.blinkSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' borderColor
                Desc = "borderColor"
                Left = LState.borderColor.ToString
                Right = Rstate.borderColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' endColor
                Desc = "endColor"
                Left = LState.endColor.ToString
                Right = Rstate.endColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirection
                Desc = "gradientDirection"
                Left = LState.gradientDirection.ToString
                Right = Rstate.gradientDirection.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirectionSpecified
                Desc = "gradientDirectionSpecified"
                Left = LState.gradientDirectionSpecified.ToString
                Right = Rstate.gradientDirectionSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyle
                Desc = "gradientShadingStyle"
                Left = LState.gradientShadingStyle.ToString
                Right = Rstate.gradientShadingStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyleSpecified
                Desc = "gradientShadingStyleSpecified"
                Left = LState.gradientShadingStyleSpecified.ToString
                Right = Rstate.gradientShadingStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStop
                Desc = "gradientStop"
                Left = LState.gradientStop.ToString
                Right = Rstate.gradientStop.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStopSpecified
                Desc = "gradientStopSpecified"
                Left = LState.gradientStopSpecified.ToString
                Right = Rstate.gradientStopSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternColor
                Desc = "patternColor"
                Left = LState.patternColor.ToString
                Right = Rstate.patternColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyle
                Desc = "patternStyle"
                Left = LState.patternStyle.ToString
                Right = Rstate.patternStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyleSpecified
                Desc = "patternStyleSpecified"
                Left = LState.patternStyleSpecified.ToString
                Right = Rstate.patternStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' stateId
                Desc = "stateId"
                Left = LState.stateId.ToString
                Right = Rstate.stateId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' Handle special datatypes within the multistate indicator
                Call Compare_BasicCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                Call Compare_BasicImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub

    Public Sub CompareActiveXProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LDat As ME_Diff.activeXObjectType = lobj
        Dim RDat As ME_Diff.activeXObjectType = robj
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        'Active X Oject Animations.animatevisibility

        If LDat.animations IsNot Nothing Then
            If RDat.animations IsNot Nothing Then
                If LDat.animations.animateVisibility IsNot Nothing Then
                    If RDat.animations.animateVisibility IsNot Nothing Then
                        ' expression
                        Desc = "expression"
                        Left = LDat.animations.animateVisibility.expression.ToString
                        Right = RDat.animations.animateVisibility.expression.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' expressionTrueState
                        Desc = "expressionTrueState"
                        Left = LDat.animations.animateVisibility.expressionTrueState.ToString
                        Right = RDat.animations.animateVisibility.expressionTrueState.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                        End If

                        ' expressionTrueStateSpecified
                        Desc = "expressionTrueStateSpecified"
                        Left = LDat.animations.animateVisibility.expressionTrueStateSpecified.ToString
                        Right = RDat.animations.animateVisibility.expressionTrueStateSpecified.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                        End If
                    Else
                        ' right animations visibility not defined
                        If LDat.animations.animateVisibility IsNot Nothing Then
                            Call AddListContentMatch(Gnest,
                                                                 fname,
                                                                 lobj.name,
                                                                 GetMEObjectType(lobj),
                                                                 "Animations-AnimateVisibility Object Mismatch",
                                                                 "Defined",
                                                                 "Defined")
                        End If
                    End If
                Else
                    ' left animations visibility not defined
                    If RDat.animations.animateVisibility IsNot Nothing Then
                        Call AddListContentMatch(Gnest,
                                                             fname,
                                                             lobj.name,
                                                             GetMEObjectType(lobj),
                                                             "Animations-AnimateVisibility Object Mismatch",
                                                             "Not Defined",
                                                             "Defined")
                    End If
                End If
            Else
                ' left defined right not
                If Not LDat.animations IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                                         fname,
                                                         lobj.name,
                                                         GetMEObjectType(lobj),
                                                         "Animations Object Mismatch",
                                                         "Defined",
                                                         "Not Defined")
                End If
            End If
        Else
            ' check if rdat object isnot nothing and create message if true
            If RDat.animations IsNot Nothing Then
                Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Animations Object Mismatch",
                                                     "Not Defined",
                                                     "Defined")
            End If
        End If

        ' Active X Object Data
        If LDat.data IsNot Nothing Then
            If RDat.data IsNot Nothing Then

                ' data
                Desc = "data"
                Left = GetStringFromByteArray(LDat.data.data)
                Right = GetStringFromByteArray(RDat.data.data)
                'Left = LDat.data.data.ToString
                'Right = RDat.data.data.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                End If

                ' fileName
                Desc = "fileName"
                Left = LDat.data.fileName.ToString
                Right = RDat.data.fileName.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                End If

                ' fileVersion
                Desc = "fileVersion"
                Left = LDat.data.fileVersion.ToString
                Right = RDat.data.fileVersion.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                End If

                ' progId
                Desc = "progId"
                Left = LDat.data.progId.ToString
                Right = RDat.data.progId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                End If

                ' typeLibrary
                Desc = "typeLibrary"
                Left = LDat.data.typeLibrary.ToString
                Right = RDat.data.typeLibrary.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                End If

                ' typeLibVersion
                Desc = "typeLibVersion"
                Left = LDat.data.typeLibVersion.ToString
                Right = RDat.data.typeLibVersion.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                End If

            Else
                ' left defined right not
                If LDat.data IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                                         fname,
                                                         lobj.name,
                                                         GetMEObjectType(lobj),
                                                         "Data Object Mismatch",
                                                         "Defined",
                                                         "Not Defined")
                End If
            End If
        Else
            ' check if rdat object isnot nothing and create message if true
            If RDat.data IsNot Nothing Then
                Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Data Object Mismatch",
                                                     "Not Defined",
                                                     "Defined")
            End If
        End If

        ' Active X Object Parameters
        If LDat.parameters IsNot Nothing Then
            If RDat.parameters IsNot Nothing Then

                For a = 0 To LDat.parameters.Count - 1

                    ' description
                    Desc = "description"
                    Left = LDat.parameters(a).description.ToString
                    Right = RDat.parameters(a).description.ToString
                    If Not Left = Right Then
                        Call AddListContentMatch(Gnest,
                                                         fname,
                                                         lobj.name,
                                                         GetMEObjectType(lobj),
                                                         "Parameter - " & a & " - " & Desc,
                                                         Left,
                                                         Right)
                    End If

                    ' name
                    Desc = "name"
                    Left = LDat.parameters(a).name.ToString
                    Right = RDat.parameters(a).name.ToString
                    If Not Left = Right Then
                        Call AddListContentMatch(Gnest,
                                                         fname,
                                                         lobj.name,
                                                         GetMEObjectType(lobj),
                                                         "Parameter - " & a & " - " & Desc,
                                                         Left,
                                                         Right)
                    End If

                    ' value
                    Desc = "value"
                    Left = LDat.parameters(a).value.ToString
                    Right = RDat.parameters(a).value.ToString
                    If Not Left = Right Then
                        Call AddListContentMatch(Gnest,
                                                         fname,
                                                         lobj.name,
                                                         GetMEObjectType(lobj),
                                                         "Parameter - " & a & " - " & Desc,
                                                         Left,
                                                         Right)
                    End If

                Next



                ' fileName
                Desc = "fileName"
                Left = LDat.data.fileName.ToString
                Right = RDat.data.fileName.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                End If

                ' fileVersion
                Desc = "fileVersion"
                Left = LDat.data.fileVersion.ToString
                Right = RDat.data.fileVersion.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                End If

                ' progId
                Desc = "progId"
                Left = LDat.data.progId.ToString
                Right = RDat.data.progId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                End If

                ' typeLibrary
                Desc = "typeLibrary"
                Left = LDat.data.typeLibrary.ToString
                Right = RDat.data.typeLibrary.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                End If

                ' typeLibVersion
                Desc = "typeLibVersion"
                Left = LDat.data.typeLibVersion.ToString
                Right = RDat.data.typeLibVersion.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     Desc,
                                                     Left,
                                                     Right)
                End If

            Else
                ' left defined right not
                If LDat.parameters IsNot Nothing Then
                    Call AddListContentMatch(Gnest,
                                                         fname,
                                                         lobj.name,
                                                         GetMEObjectType(lobj),
                                                         "Data Parameters Mismatch",
                                                         "Defined",
                                                         "Not Defined")
                End If
            End If
        Else
            ' check if rdat object isnot nothing and create message if true
            If RDat.parameters IsNot Nothing Then
                Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "Data Parameters Mismatch",
                                                     "Not Defined",
                                                     "Defined")
            End If
        End If

        ' Handle special datatypes within the multistate indicator
        'Call Compare_BasicCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
        'Call Compare_BasicImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

    End Sub

    Public Function GetStringFromByteArray(ByRef barray As Byte()) As String
        Dim tstr As String = ""
        Dim tarr() As String
        ReDim tarr(barray.Count - 1)
        For a = 0 To barray.Count - 1
            tarr(a) = barray(a).ToString
        Next
        tstr = Join(tarr, "")
        Return tstr
    End Function

    Public Sub CompareNumericInputEnableProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.numericInputEnableType = lobj
        Dim RMSB As ME_Diff.numericInputEnableType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareStringInputEnableProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.stringInputEnableType = lobj
        Dim RMSB As ME_Diff.stringInputEnableType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub CompareRampProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.rampButtonType = lobj
        Dim RMSB As ME_Diff.rampButtonType = robj

        ' this object type has no states to compare hence this section is empty at present

        ' Handle special datatypes within the multistate indicator
        Call Compare_BasicCaptionType(lobj, robj, fname, Gnest)
        Call Compare_BasicImageType(lobj, robj, fname, Gnest)

    End Sub

    Public Sub Compare_BasicCaptionType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.interlockedButtonBaseStateType,
                                        ByRef RState As ME_Diff.interlockedButtonBaseStateType,
                                        ByRef StateIndex As Integer)

        Dim LCap As ME_Diff.basicCaptionType = LState.caption
        Dim RCap As ME_Diff.basicCaptionType = RState.caption
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LCap.alignment.ToString
        Right = RCap.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LCap.alignmentSpecified.ToString
        Right = RCap.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LCap.backColor.ToString
        Right = RCap.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LCap.backStyle.ToString
        Right = RCap.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LCap.backStyleSpecified.ToString
        Right = RCap.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LCap.blink.ToString
        Right = RCap.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LCap.blinkSpecified.ToString
        Right = RCap.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' bold
        Desc = "bold"
        Left = LCap.bold.ToString
        Right = RCap.bold.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' boldSpecified
        Desc = "boldSpecified"
        Left = LCap.boldSpecified.ToString
        Right = RCap.boldSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' Caption
        Desc = "caption"
        Left = LCap.caption.ToString
        Right = RCap.caption.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LCap.color.ToString
        Right = RCap.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontFamily
        Desc = "fontFamily"
        Left = LCap.fontFamily.ToString
        Right = RCap.fontFamily.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontSize
        Desc = "fontSize"
        Left = LCap.fontSize.ToString
        Right = RCap.fontSize.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontSizeSpecified
        Desc = "fontSizeSpecified"
        Left = LCap.fontSizeSpecified.ToString
        Right = RCap.fontSizeSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' italic
        Desc = "italic"
        Left = LCap.italic.ToString
        Right = RCap.italic.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' italicSpecified
        Desc = "italicSpecified"
        Left = LCap.italicSpecified.ToString
        Right = RCap.italicSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' strikethrough
        Desc = "strikethrough"
        Left = LCap.strikethrough.ToString
        Right = RCap.strikethrough.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' strikethroughSpecified
        Desc = "strikethroughSpecified"
        Left = LCap.strikethroughSpecified.ToString
        Right = RCap.strikethroughSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' underline
        Desc = "underline"
        Left = LCap.underline.ToString
        Right = RCap.underline.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' underlineSpecified
        Desc = "underlineSpecified"
        Left = LCap.underlineSpecified.ToString
        Right = RCap.underlineSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' wordWrap
        Desc = "wordWrap"
        Left = LCap.wordWrap.ToString
        Right = RCap.wordWrap.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' wordWrapSpecified
        Desc = "wordWrapSpecified"
        Left = LCap.wordWrapSpecified.ToString
        Right = RCap.wordWrapSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub Compare_ControlListCaptionType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.pilotedControlListSelectorStateType,
                                        ByRef RState As ME_Diff.pilotedControlListSelectorStateType,
                                        ByRef StateIndex As Integer)

        Dim LCap As ME_Diff.controlListSelectorCaptionType = LState.caption
        Dim RCap As ME_Diff.controlListSelectorCaptionType = RState.caption
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LCap.alignment.ToString
        Right = RCap.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LCap.alignmentSpecified.ToString
        Right = RCap.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LCap.backColor.ToString
        Right = RCap.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LCap.backStyle.ToString
        Right = RCap.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LCap.backStyleSpecified.ToString
        Right = RCap.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LCap.blink.ToString
        Right = RCap.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LCap.blinkSpecified.ToString
        Right = RCap.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If


        ' Caption
        Desc = "caption"
        Left = LCap.caption.ToString
        Right = RCap.caption.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LCap.color.ToString
        Right = RCap.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub Compare_ControlListCaptionType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.controlListSelectorStateType,
                                        ByRef RState As ME_Diff.controlListSelectorStateType,
                                        ByRef StateIndex As Integer)

        Dim LCap As ME_Diff.controlListSelectorCaptionType = LState.caption
        Dim RCap As ME_Diff.controlListSelectorCaptionType = RState.caption
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LCap.alignment.ToString
        Right = RCap.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LCap.alignmentSpecified.ToString
        Right = RCap.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LCap.backColor.ToString
        Right = RCap.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LCap.backStyle.ToString
        Right = RCap.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LCap.backStyleSpecified.ToString
        Right = RCap.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LCap.blink.ToString
        Right = RCap.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LCap.blinkSpecified.ToString
        Right = RCap.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If


        ' Caption
        Desc = "caption"
        Left = LCap.caption.ToString
        Right = RCap.caption.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LCap.color.ToString
        Right = RCap.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub Compare_BasicCaptionType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.basicStateType,
                                        ByRef RState As ME_Diff.basicStateType,
                                        ByRef StateIndex As Integer)

        Dim LCap As ME_Diff.basicCaptionType = LState.caption
        Dim RCap As ME_Diff.basicCaptionType = RState.caption
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LCap.alignment.ToString
        Right = RCap.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LCap.alignmentSpecified.ToString
        Right = RCap.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LCap.backColor.ToString
        Right = RCap.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LCap.backStyle.ToString
        Right = RCap.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LCap.backStyleSpecified.ToString
        Right = RCap.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LCap.blink.ToString
        Right = RCap.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LCap.blinkSpecified.ToString
        Right = RCap.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' bold
        Desc = "bold"
        Left = LCap.bold.ToString
        Right = RCap.bold.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' boldSpecified
        Desc = "boldSpecified"
        Left = LCap.boldSpecified.ToString
        Right = RCap.boldSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' Caption
        Desc = "caption"
        Left = LCap.caption.ToString
        Right = RCap.caption.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LCap.color.ToString
        Right = RCap.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontFamily
        Desc = "fontFamily"
        Left = LCap.fontFamily.ToString
        Right = RCap.fontFamily.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontSize
        Desc = "fontSize"
        Left = LCap.fontSize.ToString
        Right = RCap.fontSize.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontSizeSpecified
        Desc = "fontSizeSpecified"
        Left = LCap.fontSizeSpecified.ToString
        Right = RCap.fontSizeSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' italic
        Desc = "italic"
        Left = LCap.italic.ToString
        Right = RCap.italic.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' italicSpecified
        Desc = "italicSpecified"
        Left = LCap.italicSpecified.ToString
        Right = RCap.italicSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' strikethrough
        Desc = "strikethrough"
        Left = LCap.strikethrough.ToString
        Right = RCap.strikethrough.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' strikethroughSpecified
        Desc = "strikethroughSpecified"
        Left = LCap.strikethroughSpecified.ToString
        Right = RCap.strikethroughSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' underline
        Desc = "underline"
        Left = LCap.underline.ToString
        Right = RCap.underline.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' underlineSpecified
        Desc = "underlineSpecified"
        Left = LCap.underlineSpecified.ToString
        Right = RCap.underlineSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' wordWrap
        Desc = "wordWrap"
        Left = LCap.wordWrap.ToString
        Right = RCap.wordWrap.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' wordWrapSpecified
        Desc = "wordWrapSpecified"
        Left = LCap.wordWrapSpecified.ToString
        Right = RCap.wordWrapSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub Compare_ListIndicatorCaptionType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.listIndicatorStateType,
                                        ByRef RState As ME_Diff.listIndicatorStateType,
                                        ByRef StateIndex As Integer)

        Dim LCap As ME_Diff.listIndicatorCaptionType = LState.caption
        Dim RCap As ME_Diff.listIndicatorCaptionType = RState.caption
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LCap.alignment.ToString
        Right = RCap.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LCap.alignmentSpecified.ToString
        Right = RCap.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LCap.backColor.ToString
        Right = RCap.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LCap.backStyle.ToString
        Right = RCap.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LCap.backStyleSpecified.ToString
        Right = RCap.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LCap.blink.ToString
        Right = RCap.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LCap.blinkSpecified.ToString
        Right = RCap.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' Caption
        Desc = "caption"
        Left = LCap.caption.ToString
        Right = RCap.caption.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LCap.color.ToString
        Right = RCap.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub Compare_DisplayListSelectorCaptionType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.displayListSelectorStateType,
                                        ByRef RState As ME_Diff.displayListSelectorStateType,
                                        ByRef StateIndex As Integer)

        Dim LCap As ME_Diff.displayListSelectorCaptionType = LState.caption
        Dim RCap As ME_Diff.displayListSelectorCaptionType = RState.caption
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LCap.alignment.ToString
        Right = RCap.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LCap.alignmentSpecified.ToString
        Right = RCap.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LCap.backColor.ToString
        Right = RCap.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LCap.backStyle.ToString
        Right = RCap.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LCap.backStyleSpecified.ToString
        Right = RCap.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LCap.blink.ToString
        Right = RCap.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LCap.blinkSpecified.ToString
        Right = RCap.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' Caption
        Desc = "caption"
        Left = LCap.caption.ToString
        Right = RCap.caption.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LCap.color.ToString
        Right = RCap.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' usesDisplayName
        Desc = "usesDisplayName"
        Left = LCap.usesDisplayName.ToString
        Right = RCap.usesDisplayName.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' usesDisplayNameSpecified
        Desc = "usesDisplayNameSpecified"
        Left = LCap.usesDisplayNameSpecified.ToString
        Right = RCap.usesDisplayNameSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If


    End Sub

    Public Sub Compare_BasicCaptionType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String)

        Dim LCap As ME_Diff.basicCaptionType = lobj.caption
        Dim RCap As ME_Diff.basicCaptionType = robj.caption
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LCap.alignment.ToString
        Right = RCap.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LCap.alignmentSpecified.ToString
        Right = RCap.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LCap.backColor.ToString
        Right = RCap.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LCap.backStyle.ToString
        Right = RCap.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LCap.backStyleSpecified.ToString
        Right = RCap.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LCap.blink.ToString
        Right = RCap.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LCap.blinkSpecified.ToString
        Right = RCap.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' bold
        Desc = "bold"
        Left = LCap.bold.ToString
        Right = RCap.bold.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' boldSpecified
        Desc = "boldSpecified"
        Left = LCap.boldSpecified.ToString
        Right = RCap.boldSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' Caption
        Desc = "caption"
        Left = LCap.caption.ToString
        Right = RCap.caption.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LCap.color.ToString
        Right = RCap.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' fontFamily
        Desc = "fontFamily"
        Left = LCap.fontFamily.ToString
        Right = RCap.fontFamily.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' fontSize
        Desc = "fontSize"
        Left = LCap.fontSize.ToString
        Right = RCap.fontSize.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' fontSizeSpecified
        Desc = "fontSizeSpecified"
        Left = LCap.fontSizeSpecified.ToString
        Right = RCap.fontSizeSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' italic
        Desc = "italic"
        Left = LCap.italic.ToString
        Right = RCap.italic.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' italicSpecified
        Desc = "italicSpecified"
        Left = LCap.italicSpecified.ToString
        Right = RCap.italicSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' strikethrough
        Desc = "strikethrough"
        Left = LCap.strikethrough.ToString
        Right = RCap.strikethrough.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' strikethroughSpecified
        Desc = "strikethroughSpecified"
        Left = LCap.strikethroughSpecified.ToString
        Right = RCap.strikethroughSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' underline
        Desc = "underline"
        Left = LCap.underline.ToString
        Right = RCap.underline.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' underlineSpecified
        Desc = "underlineSpecified"
        Left = LCap.underlineSpecified.ToString
        Right = RCap.underlineSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' wordWrap
        Desc = "wordWrap"
        Left = LCap.wordWrap.ToString
        Right = RCap.wordWrap.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' wordWrapSpecified
        Desc = "wordWrapSpecified"
        Left = LCap.wordWrapSpecified.ToString
        Right = RCap.wordWrapSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub CompareMultistateProperties(ByRef lobj As Object, ByRef robj As Object, ByRef fname As String, ByRef Gnest As String)

        ' Declare temporary variables for breaking out object properties to compare against
        Dim LMSB As ME_Diff.multistateButtonType = lobj
        Dim RMSB As ME_Diff.multistateButtonType = robj

        If Not LMSB.states.Count = RMSB.states.Count Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "MultistateButton - State Count Mismatch",
                                     LMSB.states.Count,
                                     RMSB.states.Count)
        Else
            For a = 0 To LMSB.states.Count - 1

                ' Declare reuseable loop variables
                Dim MEObjType As String = GetMEObjectType(lobj)
                Dim LState As ME_Diff.multistateButtonBaseStateType = LMSB.states(a)
                Dim Rstate As ME_Diff.multistateButtonBaseStateType = RMSB.states(a)
                Dim Left As String = ""
                Dim Right As String = ""
                Dim Desc As String = ""


                ' backColor
                Desc = "backColor"
                Left = LState.backColor.ToString
                Right = Rstate.backColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blink
                Desc = "blink"
                Left = LState.blink.ToString
                Right = Rstate.blink.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' blinkSpecified
                Desc = "blinkSpecified"
                Left = LState.blinkSpecified.ToString
                Right = Rstate.blinkSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' borderColor
                Desc = "borderColor"
                Left = LState.borderColor.ToString
                Right = Rstate.borderColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' endColor
                Desc = "endColor"
                Left = LState.endColor.ToString
                Right = Rstate.endColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirection
                Desc = "gradientDirection"
                Left = LState.gradientDirection.ToString
                Right = Rstate.gradientDirection.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientDirectionSpecified
                Desc = "gradientDirectionSpecified"
                Left = LState.gradientDirectionSpecified.ToString
                Right = Rstate.gradientDirectionSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyle
                Desc = "gradientShadingStyle"
                Left = LState.gradientShadingStyle.ToString
                Right = Rstate.gradientShadingStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientShadingStyleSpecified
                Desc = "gradientShadingStyleSpecified"
                Left = LState.gradientShadingStyleSpecified.ToString
                Right = Rstate.gradientShadingStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStop
                Desc = "gradientStop"
                Left = LState.gradientStop.ToString
                Right = Rstate.gradientStop.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' gradientStopSpecified
                Desc = "gradientStopSpecified"
                Left = LState.gradientStopSpecified.ToString
                Right = Rstate.gradientStopSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternColor
                Desc = "patternColor"
                Left = LState.patternColor.ToString
                Right = Rstate.patternColor.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyle
                Desc = "patternStyle"
                Left = LState.patternStyle.ToString
                Right = Rstate.patternStyle.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' patternStyleSpecified
                Desc = "patternStyleSpecified"
                Left = LState.patternStyleSpecified.ToString
                Right = Rstate.patternStyleSpecified.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' stateId
                Desc = "stateId"
                Left = LState.stateId.ToString
                Right = Rstate.stateId.ToString
                If Not Left = Right Then
                    Call AddListContentMatch(Gnest,
                                             fname,
                                             lobj.name,
                                             GetMEObjectType(lobj),
                                             "State " & a & " - " & Desc,
                                             Left,
                                             Right)
                End If

                ' value
                Desc = "value"
                If LState.value IsNot Nothing Then
                    If Rstate.value IsNot Nothing Then
                        Left = LState.value.ToString
                        Right = Rstate.value.ToString
                        If Not Left = Right Then
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     Left,
                                                     Right)
                        End If
                    Else
                        If LState.value IsNot Nothing Then
                            ' Left is set to value but right is not
                            Left = LState.value.ToString
                            Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     Left,
                                                     "Nothing")
                        End If
                    End If

                Else
                    If Rstate.value IsNot Nothing Then
                        Right = Rstate.value.ToString
                        Call AddListContentMatch(Gnest,
                                                     fname,
                                                     lobj.name,
                                                     GetMEObjectType(lobj),
                                                     "State " & a & " - " & Desc,
                                                     "Nothing",
                                                     Right)
                    End If
                End If



                ' Handle special datatypes within the multistate indicator
                Call Compare_BasicCaptionType(lobj, robj, fname, Gnest, LState, Rstate, a)
                Call Compare_BasicImageType(lobj, robj, fname, Gnest, LState, Rstate, a)

            Next

        End If

    End Sub

    Public Sub Compare_BasicCaptionType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.multistateButtonBaseStateType,
                                        ByRef RState As ME_Diff.multistateButtonBaseStateType,
                                        ByRef StateIndex As Integer)

        Dim LCap As ME_Diff.basicCaptionType = LState.caption
        Dim RCap As ME_Diff.basicCaptionType = RState.caption
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LCap.alignment.ToString
        Right = RCap.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LCap.alignmentSpecified.ToString
        Right = RCap.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LCap.backColor.ToString
        Right = RCap.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LCap.backStyle.ToString
        Right = RCap.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LCap.backStyleSpecified.ToString
        Right = RCap.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LCap.blink.ToString
        Right = RCap.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LCap.blinkSpecified.ToString
        Right = RCap.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' bold
        Desc = "bold"
        Left = LCap.bold.ToString
        Right = RCap.bold.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' boldSpecified
        Desc = "boldSpecified"
        Left = LCap.boldSpecified.ToString
        Right = RCap.boldSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' Caption
        Desc = "caption"
        Left = LCap.caption.ToString
        Right = RCap.caption.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LCap.color.ToString
        Right = RCap.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontFamily
        Desc = "fontFamily"
        Left = LCap.fontFamily.ToString
        Right = RCap.fontFamily.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontSize
        Desc = "fontSize"
        Left = LCap.fontSize.ToString
        Right = RCap.fontSize.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontSizeSpecified
        Desc = "fontSizeSpecified"
        Left = LCap.fontSizeSpecified.ToString
        Right = RCap.fontSizeSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' italic
        Desc = "italic"
        Left = LCap.italic.ToString
        Right = RCap.italic.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' italicSpecified
        Desc = "italicSpecified"
        Left = LCap.italicSpecified.ToString
        Right = RCap.italicSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' strikethrough
        Desc = "strikethrough"
        Left = LCap.strikethrough.ToString
        Right = RCap.strikethrough.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' strikethroughSpecified
        Desc = "strikethroughSpecified"
        Left = LCap.strikethroughSpecified.ToString
        Right = RCap.strikethroughSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' underline
        Desc = "underline"
        Left = LCap.underline.ToString
        Right = RCap.underline.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' underlineSpecified
        Desc = "underlineSpecified"
        Left = LCap.underlineSpecified.ToString
        Right = RCap.underlineSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' wordWrap
        Desc = "wordWrap"
        Left = LCap.wordWrap.ToString
        Right = RCap.wordWrap.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' wordWrapSpecified
        Desc = "wordWrapSpecified"
        Left = LCap.wordWrapSpecified.ToString
        Right = RCap.wordWrapSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub Compare_BasicCaptionType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.pushButtonBaseStateValueType,
                                        ByRef RState As ME_Diff.pushButtonBaseStateValueType,
                                        ByRef StateIndex As Integer)

        Dim LCap As ME_Diff.basicCaptionType = LState.caption
        Dim RCap As ME_Diff.basicCaptionType = RState.caption
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LCap.alignment.ToString
        Right = RCap.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LCap.alignmentSpecified.ToString
        Right = RCap.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LCap.backColor.ToString
        Right = RCap.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LCap.backStyle.ToString
        Right = RCap.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LCap.backStyleSpecified.ToString
        Right = RCap.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LCap.blink.ToString
        Right = RCap.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LCap.blinkSpecified.ToString
        Right = RCap.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' bold
        Desc = "bold"
        Left = LCap.bold.ToString
        Right = RCap.bold.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' boldSpecified
        Desc = "boldSpecified"
        Left = LCap.boldSpecified.ToString
        Right = RCap.boldSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' Caption
        Desc = "caption"
        Left = LCap.caption.ToString
        Right = RCap.caption.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LCap.color.ToString
        Right = RCap.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontFamily
        Desc = "fontFamily"
        Left = LCap.fontFamily.ToString
        Right = RCap.fontFamily.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontSize
        Desc = "fontSize"
        Left = LCap.fontSize.ToString
        Right = RCap.fontSize.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' fontSizeSpecified
        Desc = "fontSizeSpecified"
        Left = LCap.fontSizeSpecified.ToString
        Right = RCap.fontSizeSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' italic
        Desc = "italic"
        Left = LCap.italic.ToString
        Right = RCap.italic.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' italicSpecified
        Desc = "italicSpecified"
        Left = LCap.italicSpecified.ToString
        Right = RCap.italicSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' strikethrough
        Desc = "strikethrough"
        Left = LCap.strikethrough.ToString
        Right = RCap.strikethrough.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' strikethroughSpecified
        Desc = "strikethroughSpecified"
        Left = LCap.strikethroughSpecified.ToString
        Right = RCap.strikethroughSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' underline
        Desc = "underline"
        Left = LCap.underline.ToString
        Right = RCap.underline.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' underlineSpecified
        Desc = "underlineSpecified"
        Left = LCap.underlineSpecified.ToString
        Right = RCap.underlineSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' wordWrap
        Desc = "wordWrap"
        Left = LCap.wordWrap.ToString
        Right = RCap.wordWrap.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' wordWrapSpecified
        Desc = "wordWrapSpecified"
        Left = LCap.wordWrapSpecified.ToString
        Right = RCap.wordWrapSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub Compare_BasicImageType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String)

        Dim LImg As ME_Diff.basicImageType = lobj.imageSettings
        Dim RImg As ME_Diff.basicImageType = robj.imageSettings
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LImg.alignment.ToString
        Right = RImg.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LImg.alignmentSpecified.ToString
        Right = RImg.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LImg.backColor.ToString
        Right = RImg.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LImg.backStyle.ToString
        Right = RImg.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LImg.backStyleSpecified.ToString
        Right = RImg.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LImg.blink.ToString
        Right = RImg.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LImg.blinkSpecified.ToString
        Right = RImg.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LImg.color.ToString
        Right = RImg.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' imageName
        Desc = "imageName"
        Left = LImg.imageName.ToString
        Right = RImg.imageName.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' scaled
        Desc = "scaled"
        Left = LImg.scaled.ToString
        Right = RImg.scaled.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

        ' scaledSpecified
        Desc = "scaledSpecified"
        Left = LImg.scaledSpecified.ToString
        Right = RImg.scaledSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub Compare_BasicImageType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.interlockedButtonBaseStateType,
                                        ByRef RState As ME_Diff.interlockedButtonBaseStateType,
                                        ByRef StateIndex As Integer)

        Dim LImg As ME_Diff.basicImageType = LState.imageSettings
        Dim RImg As ME_Diff.basicImageType = RState.imageSettings
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LImg.alignment.ToString
        Right = RImg.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LImg.alignmentSpecified.ToString
        Right = RImg.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LImg.backColor.ToString
        Right = RImg.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LImg.backStyle.ToString
        Right = RImg.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LImg.backStyleSpecified.ToString
        Right = RImg.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LImg.blink.ToString
        Right = RImg.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LImg.blinkSpecified.ToString
        Right = RImg.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LImg.color.ToString
        Right = RImg.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' imageName
        Desc = "imageName"
        Left = LImg.imageName.ToString
        Right = RImg.imageName.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' scaled
        Desc = "scaled"
        Left = LImg.scaled.ToString
        Right = RImg.scaled.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' scaledSpecified
        Desc = "scaledSpecified"
        Left = LImg.scaledSpecified.ToString
        Right = RImg.scaledSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub Compare_BasicImageType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.basicStateType,
                                        ByRef RState As ME_Diff.basicStateType,
                                        ByRef StateIndex As Integer)

        Dim LImg As ME_Diff.basicImageType = LState.imageSettings
        Dim RImg As ME_Diff.basicImageType = RState.imageSettings
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LImg.alignment.ToString
        Right = RImg.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LImg.alignmentSpecified.ToString
        Right = RImg.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LImg.backColor.ToString
        Right = RImg.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LImg.backStyle.ToString
        Right = RImg.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LImg.backStyleSpecified.ToString
        Right = RImg.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LImg.blink.ToString
        Right = RImg.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LImg.blinkSpecified.ToString
        Right = RImg.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LImg.color.ToString
        Right = RImg.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' imageName
        Desc = "imageName"
        Left = LImg.imageName.ToString
        Right = RImg.imageName.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' scaled
        Desc = "scaled"
        Left = LImg.scaled.ToString
        Right = RImg.scaled.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' scaledSpecified
        Desc = "scaledSpecified"
        Left = LImg.scaledSpecified.ToString
        Right = RImg.scaledSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub Compare_SymbolIndicatorImageType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.symbolIndicatorStateType,
                                        ByRef RState As ME_Diff.symbolIndicatorStateType,
                                        ByRef StateIndex As Integer)

        Dim LImg As ME_Diff.symbolIndicatorImageType = LState.imageSettings
        Dim RImg As ME_Diff.symbolIndicatorImageType = RState.imageSettings
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' backColor
        Desc = "backColor"
        Left = LImg.backColor.ToString
        Right = RImg.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LImg.backStyle.ToString
        Right = RImg.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LImg.backStyleSpecified.ToString
        Right = RImg.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LImg.blink.ToString
        Right = RImg.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LImg.blinkSpecified.ToString
        Right = RImg.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LImg.color.ToString
        Right = RImg.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub
    Public Sub Compare_BasicImageType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.multistateButtonBaseStateType,
                                        ByRef RState As ME_Diff.multistateButtonBaseStateType,
                                        ByRef StateIndex As Integer)

        Dim LImg As ME_Diff.basicImageType = LState.imageSettings
        Dim RImg As ME_Diff.basicImageType = RState.imageSettings
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LImg.alignment.ToString
        Right = RImg.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LImg.alignmentSpecified.ToString
        Right = RImg.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LImg.backColor.ToString
        Right = RImg.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LImg.backStyle.ToString
        Right = RImg.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LImg.backStyleSpecified.ToString
        Right = RImg.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LImg.blink.ToString
        Right = RImg.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LImg.blinkSpecified.ToString
        Right = RImg.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LImg.color.ToString
        Right = RImg.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' imageName
        Desc = "imageName"
        Left = LImg.imageName.ToString
        Right = RImg.imageName.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' scaled
        Desc = "scaled"
        Left = LImg.scaled.ToString
        Right = RImg.scaled.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' scaledSpecified
        Desc = "scaledSpecified"
        Left = LImg.scaledSpecified.ToString
        Right = RImg.scaledSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    Public Sub Compare_BasicImageType(ByRef lobj As Object,
                                        ByRef robj As Object,
                                        ByRef fname As String,
                                        ByRef Gnest As String,
                                        ByRef LState As ME_Diff.pushButtonBaseStateValueType,
                                        ByRef RState As ME_Diff.pushButtonBaseStateValueType,
                                        ByRef StateIndex As Integer)

        Dim LImg As ME_Diff.basicImageType = LState.imageSettings
        Dim RImg As ME_Diff.basicImageType = RState.imageSettings
        Dim Desc As String = ""
        Dim Left As String = ""
        Dim Right As String = ""

        ' alignment
        Desc = "alignment"
        Left = LImg.alignment.ToString
        Right = RImg.alignment.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' alignmentSpecified
        Desc = "alignmentSpecified"
        Left = LImg.alignmentSpecified.ToString
        Right = RImg.alignmentSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backColor
        Desc = "backColor"
        Left = LImg.backColor.ToString
        Right = RImg.backColor.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyle
        Desc = "backStyle"
        Left = LImg.backStyle.ToString
        Right = RImg.backStyle.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' backStyleSpecified
        Desc = "backStyleSpecified"
        Left = LImg.backStyleSpecified.ToString
        Right = RImg.backStyleSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blink
        Desc = "blink"
        Left = LImg.blink.ToString
        Right = RImg.blink.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' blinkSpecified
        Desc = "blinkSpecified"
        Left = LImg.blinkSpecified.ToString
        Right = RImg.blinkSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' color
        Desc = "color"
        Left = LImg.color.ToString
        Right = RImg.color.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' imageName
        Desc = "imageName"
        Left = LImg.imageName.ToString
        Right = RImg.imageName.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' scaled
        Desc = "scaled"
        Left = LImg.scaled.ToString
        Right = RImg.scaled.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

        ' scaledSpecified
        Desc = "scaledSpecified"
        Left = LImg.scaledSpecified.ToString
        Right = RImg.scaledSpecified.ToString
        If Not Left = Right Then
            Call AddListContentMatch(Gnest,
                                     fname,
                                     lobj.name,
                                     GetMEObjectType(lobj),
                                     "State " & StateIndex & " - " & Desc,
                                     Left,
                                     Right)
        End If

    End Sub

    ''' <summary>
    ''' Subroutine that controls adding group content match exceptions to the output report
    ''' </summary>
    ''' <param name="Gname">Group Name. Will be Root if object exists on top level</param>
    ''' <param name="fname">File name</param>
    ''' <param name="Gnest">Group nesting level. Is a string path containing the preceding groups that make up the path to this object</param>
    ''' <param name="ExMessage">Exception message. This is a user generated string that details the nature of the exception that caused an entry to be generated for this list</param>
    Public Sub AddListContentMatchGroupException(ByVal Gname As String, ByVal fname As String, ByVal Gnest As String, ByVal ExMessage As String)
        Dim msg As String = ""
        msg = msg & "File Name : " & fname & " - "
        msg = msg & "Nest Level : " & Gnest & " - "
        msg = msg & "Object Name : " & Gname & " - "
        msg = msg & ExMessage
        List_FileCompareContentMatch.Add(msg)
    End Sub

    ''' <summary>
    ''' Writes the contents of the report lists out to a textfile and displays it
    ''' </summary>
    ''' <param name="pathtofile">String path to the file to output the text to. Is also what is opened in notepad.exe automatically at the end of the routine</param>
    Public Sub Outputreport(ByRef pathtofile As String)
        Using output As StreamWriter = New StreamWriter(Path.GetDirectoryName(pathtofile) & "\Output.txt", False)
            For a = 0 To List_FileCompareContentMatch.Count - 1
                output.WriteLine(List_FileCompareContentMatch.Item(a))
            Next
            System.Diagnostics.Process.Start("notepad.exe", Path.GetDirectoryName(pathtofile) & "\Output.txt")
        End Using
    End Sub

    ''' <summary>
    ''' Returns the contents of the output file as a string to pass back to the calling method instead of outputting to a text file
    ''' </summary>
    Public Function OutputToTest_File()
        'Dim strout As String = ""
        'OutputToTest_File = ""
        'For a = 0 To List_FileCompareContentMatch.Count - 1
        '    strout = strout & List_FileCompareContentMatch.Item(a) & vbCrLf
        'Next
        'OutputToTest_File = strout
        OutputToTest_File = List_FileCompareContentMatch ' return the complete list
    End Function

    ''' <summary>
    ''' Creates a new output report page instance within an existing reportoutput object and applies the currently configured filter parameters whilst doing so
    ''' </summary>
    ''' <param name="ReportObject">The master report object file you wish to add a report page to</param>
    ''' <param name="ReportList">The in memory report list to use to generate the page from</param>
    ''' <param name="LeftFile">Left file name</param>
    ''' <param name="RightFile">Right file name</param>
    Public Sub CreateOutputReportPage(ByRef ReportObject As ReportOutput,
                                      ByRef ReportList As List(Of String),
                                      ByRef LeftFile As String,
                                      ByRef RightFile As String)
        Dim NewReportPage As ReportOutputPage = New ReportOutputPage(LeftFile, RightFile)
        For Each item In ReportList
            If Not NeedsFiltered(item) Then
                NewReportPage.AddListItem(item)
            End If
        Next
        ReportObject.AddPage(NewReportPage)
    End Sub
    ''' <summary>
    ''' Checks the supplied string to see if the parameter contained within matches a filter parameter
    ''' </summary>
    ''' <param name="item">string containing the parameter to check</param>
    ''' <returns>True if the parameter in the string matches a filter, False if not </returns>
    Public Function NeedsFiltered(ByRef item As String) As Boolean
        Dim TempMember As ReportMember = New ReportMember(item)
        Dim MatchFound As Boolean = False
        For Each item In FilterList.Where(Function(x) x = TempMember.Parameter)
            MatchFound = True
            Exit For
        Next
        Return MatchFound
    End Function

End Module

