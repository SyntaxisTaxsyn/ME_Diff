﻿Imports System.Reflection
Imports System.IO
Module Subroutines

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
    Private Sub UpdateFilters()

        ' Todo use this routine to update the filter switches used by the comparison algorithm
        ' Add a call to this routine in each check/uncheck all routine
        ' Also create handlers for checkbox checked state changed calls this and updates also
        Call UpdateFilters()

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

        Dim linfo() As PropertyInfo = lobj.GetType().GetProperties()
        Dim rinfo() As PropertyInfo = lobj.GetType().GetProperties()
        For a = 0 To linfo.Count - 1
            If linfo(a).CanRead Then
                Dim str As String = linfo(a).PropertyType().ToString
                Select Case str
                    Case "ME_Diff.animationsAllMEType"
                        ' call special case handler for dealing with animation property comparison
                        Call CompareAnimationsProperties(lobj.animations, robj.animations, Gnest, fname, lobj.name.ToString)
                    Case "ME_Diff.multistateButtonBaseStateType"
                        MsgBox("")

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
            Case "ME_Diff.multistateButtonType"
                Call CompareMultistateProperties(lobj, robj, fname, Gnest)
        End Select

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


                ' Backcolor
                Desc = "Backcolor"
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
        Desc = "caption"
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

        ' Todo - fill in the rest of the properties for this data type

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
        msg = msg & "Group Name : " & Gname & " - "
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

End Module