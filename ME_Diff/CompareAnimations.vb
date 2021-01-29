Module CompareAnimations
    Public Sub CompareAnimationsProperties(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        Call CompareAnimationsProperties_AnimateColor(Lobj, Robj, Gnest, fname, oname)
        Call CompareAnimationsProperties_AnimateFill(Lobj, Robj, Gnest, fname, oname)
        Call CompareAnimationsProperties_AnimateHeight(Lobj, Robj, Gnest, fname, oname)
        Call CompareAnimationsProperties_AnimateHorizontalPosition(Lobj, Robj, Gnest, fname, oname)
        Call CompareAnimationsProperties_Animatehorizontalslider(Lobj, Robj, Gnest, fname, oname)
        Call CompareAnimationsProperties_AnimateHyperlink(Lobj, Robj, Gnest, fname, oname)
        Call CompareAnimationsProperties_AnimateRotation(Lobj, Robj, Gnest, fname, oname)
        Call CompareAnimationsProperties_AnimateVerticalPosition(Lobj, Robj, Gnest, fname, oname)
        Call CompareAnimationsProperties_AnimateVerticalSlider(Lobj, Robj, Gnest, fname, oname)
        Call CompareAnimationsProperties_AnimateVisibility(Lobj, Robj, Gnest, fname, oname)
        Call CompareAnimationsProperties_AnimateWidth(Lobj, Robj, Gnest, fname, oname)

    End Sub
    Public Sub CompareAnimationsProperties_AnimateColor(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        '----------------------------------
        ' Expression
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateColor IsNot Nothing Then
                    If Robj.animateColor IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateColor.expression = Robj.animateColor.expression Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Color-expression", Lobj.animateColor.expression.ToString, Robj.animateColor.expression.ToString)
                        End If
                    Else
                        ' This is a difference where the right side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Color", "Configured", "Nothing")
                    End If
                Else
                    If Robj.animateColor IsNot Nothing Then
                        ' This is a difference where the left side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Color", "Nothing", "Configured")
                    End If
                End If
            Else
                ' right is empty and left is not
                If Lobj.animateColor IsNot Nothing Then
                    ' report the right as missing the animation
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "Color", "Configured", "Nothing")
                End If
            End If
        Else
            If Robj IsNot Nothing Then
                If Robj.animateColor IsNot Nothing Then
                    ' left is empty and right is not
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "Color", "Nothing", "Configured")
                End If
            End If
        End If

        '----------------------------------
        ' blink rate
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateColor IsNot Nothing Then
                    If Robj.animateColor IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateColor.blinkRate = Robj.animateColor.blinkRate Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Color-blinkRate", Lobj.animateColor.blinkRate.ToString, Robj.animateColor.blinkRate.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' blink rate specified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateColor IsNot Nothing Then
                    If Robj.animateColor IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateColor.blinkRateSpecified = Robj.animateColor.blinkRateSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Color-blinkRateSpecified", Lobj.animateColor.blinkRateSpecified.ToString, Robj.animateColor.blinkRateSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If


        '----------------------------------
        ' Animate Color
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateColor IsNot Nothing Then
                    If Robj.animateColor IsNot Nothing Then
                        ' Both have properties set so compare them
                        Try
                            For a = 0 To Lobj.animateColor.color.Length - 1

                                Dim LColor As animateColorEntryType = Lobj.animateColor.color(a)
                                Dim Rcolor As animateColorEntryType = Robj.animateColor.color(a)

                                ' Back behavior
                                If Not LColor.backBehavior = Rcolor.backBehavior Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - backBehavior",
                                                             LColor.backBehavior.ToString,
                                                             Rcolor.backBehavior.ToString)
                                End If

                                ' Back behavior specified
                                If Not LColor.backBehaviorSpecified = Rcolor.backBehaviorSpecified Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - backBehaviorSpecified",
                                                             LColor.backBehaviorSpecified.ToString,
                                                             Rcolor.backBehaviorSpecified.ToString)
                                End If

                                ' Back color 1
                                If Not LColor.backColor1 = Rcolor.backColor1 Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - backColor1",
                                                             LColor.backColor1.ToString,
                                                             Rcolor.backColor1.ToString)
                                End If

                                ' Back color 2
                                If Not LColor.backColor2 = Rcolor.backColor2 Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - backColor2",
                                                             LColor.backColor2.ToString,
                                                             Rcolor.backColor2.ToString)
                                End If

                                ' End color
                                If Not LColor.endColor = Rcolor.endColor Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - endColor",
                                                             LColor.endColor.ToString,
                                                             Rcolor.endColor.ToString)
                                End If

                                ' Fill color mode
                                If Not LColor.fillColorMode = Rcolor.fillColorMode Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - fillColorMode",
                                                             LColor.fillColorMode.ToString,
                                                             Rcolor.fillColorMode.ToString)
                                End If

                                ' Fill color mode specified
                                If Not LColor.fillColorModeSpecified = Rcolor.fillColorModeSpecified Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - fillColorModeSpecified",
                                                             LColor.fillColorModeSpecified.ToString,
                                                             Rcolor.fillColorModeSpecified.ToString)
                                End If

                                ' Fill end color
                                If Not LColor.fillEndColor = Rcolor.fillEndColor Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - fillEndColor",
                                                             LColor.fillEndColor.ToString,
                                                             Rcolor.fillEndColor.ToString)
                                End If

                                ' Fill gradient direction
                                If Not LColor.fillGradientDirection = Rcolor.fillGradientDirection Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - fillGradientDirection",
                                                             LColor.fillGradientDirection.ToString,
                                                             Rcolor.fillGradientDirection.ToString)
                                End If

                                ' Fill gradient direction specified
                                If Not LColor.fillGradientDirectionSpecified = Rcolor.fillGradientDirectionSpecified Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - fillGradientDirectionspecified",
                                                             LColor.fillGradientDirectionSpecified.ToString,
                                                             Rcolor.fillGradientDirectionSpecified.ToString)
                                End If

                                ' Fill gradient shading style
                                If Not LColor.fillGradientShadingStyle = Rcolor.fillGradientShadingStyle Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - fillGradientShadingStyle",
                                                             LColor.fillGradientShadingStyle.ToString,
                                                             Rcolor.fillGradientShadingStyle.ToString)
                                End If

                                ' Fill gradient shading style specified
                                If Not LColor.fillGradientShadingStyleSpecified = Rcolor.fillGradientShadingStyleSpecified Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - fillGradientShadingStyleSpecified",
                                                             LColor.fillGradientShadingStyleSpecified.ToString,
                                                             Rcolor.fillGradientShadingStyleSpecified.ToString)
                                End If

                                ' Fill gradient stop
                                If Not LColor.fillGradientStop = Rcolor.fillGradientStop Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - fillGradientStop",
                                                             LColor.fillGradientStop.ToString,
                                                             Rcolor.fillGradientStop.ToString)
                                End If

                                ' Fill gradient stop specified
                                If Not LColor.fillGradientStopSpecified = Rcolor.fillGradientStopSpecified Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - fillGradientStopSpecified",
                                                             LColor.fillGradientStopSpecified.ToString,
                                                             Rcolor.fillGradientStopSpecified.ToString)
                                End If

                                ' Fore behavior
                                If Not LColor.foreBehavior = Rcolor.foreBehavior Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - foreBehavior",
                                                             LColor.foreBehavior.ToString,
                                                             Rcolor.foreBehavior.ToString)
                                End If

                                ' Fore behavior specified
                                If Not LColor.foreBehaviorSpecified = Rcolor.foreBehaviorSpecified Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - foreBehaviorSpecified",
                                                             LColor.foreBehaviorSpecified.ToString,
                                                             Rcolor.foreBehaviorSpecified.ToString)
                                End If

                                ' Fore color 1
                                If Not LColor.foreColor1 = Rcolor.foreColor1 Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - foreColor1",
                                                             LColor.foreColor1.ToString,
                                                             Rcolor.foreColor1.ToString)
                                End If

                                ' Fore color 2
                                If Not LColor.foreColor2 = Rcolor.foreColor2 Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - foreColor2",
                                                             LColor.foreColor2.ToString,
                                                             Rcolor.foreColor2.ToString)
                                End If

                                ' Gradient direction
                                If Not LColor.gradientDirection = Rcolor.gradientDirection Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - gradientDirection",
                                                             LColor.gradientDirection.ToString,
                                                             Rcolor.gradientDirection.ToString)
                                End If

                                ' Gradient direction specified
                                If Not LColor.gradientDirectionSpecified = Rcolor.gradientDirectionSpecified Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - gradientDirectionSpecified",
                                                             LColor.gradientDirectionSpecified.ToString,
                                                             Rcolor.gradientDirectionSpecified.ToString)
                                End If

                                ' Gradient shading style
                                If Not LColor.gradientShadingStyle = Rcolor.gradientShadingStyle Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - gradientShadingStyle",
                                                             LColor.gradientShadingStyle.ToString,
                                                             Rcolor.gradientShadingStyle.ToString)
                                End If

                                ' Gradient shading style specified
                                If Not LColor.gradientShadingStyleSpecified = Rcolor.gradientShadingStyleSpecified Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - gradientShadingStyleSpecified",
                                                             LColor.gradientShadingStyleSpecified.ToString,
                                                             Rcolor.gradientShadingStyleSpecified.ToString)
                                End If

                                ' Gradient stop
                                If Not LColor.gradientStop = Rcolor.gradientStop Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - gradientStop",
                                                             LColor.gradientStop.ToString,
                                                             Rcolor.gradientStop.ToString)
                                End If

                                ' Gradient stop specified
                                If Not LColor.gradientStopSpecified = Rcolor.gradientStopSpecified Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - gradientStopSpecified",
                                                             LColor.gradientStopSpecified.ToString,
                                                             Rcolor.gradientStopSpecified.ToString)
                                End If

                                ' Value
                                If Not LColor.value = Rcolor.value Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - value",
                                                             LColor.value.ToString,
                                                             Rcolor.value.ToString)
                                End If

                                ' Value specified
                                If Not LColor.valueSpecified = Rcolor.valueSpecified Then
                                    Call AddListContentMatch(Gnest,
                                                             fname,
                                                             oname,
                                                             "Animation",
                                                             "Animate Color" & " Item:" & a & " - valueSpecified",
                                                             LColor.valueSpecified.ToString,
                                                             Rcolor.valueSpecified.ToString)
                                End If

                            Next


                        Catch ex As Exception

                        End Try

                    End If
                End If
            End If
        End If

    End Sub
    Public Sub CompareAnimationsProperties_AnimateFill(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        '----------------------------------
        ' Expression
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateFill IsNot Nothing Then
                    If Robj.animateFill IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateFill.expression = Robj.animateFill.expression Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-expression", Lobj.animateFill.expression.ToString, Robj.animateFill.expression.ToString)
                        End If
                    Else
                        ' check left property is set
                        If Lobj.animateFill IsNot Nothing Then
                            ' This is a difference where the left side has no property set
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill", "Configured", "Nothing")
                        End If
                    End If
                Else
                    ' check right property is set
                    If Robj.animateFill IsNot Nothing Then
                        ' This is a difference where the left side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill", "Nothing", "Configured")
                    End If
                End If
            Else
                ' Case for left object isnot nothing
                If Lobj IsNot Nothing Then
                    If Lobj.animateFill IsNot Nothing Then
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill", "Configured", "Nothing")
                    End If
                End If
            End If
        Else
            ' Case for right object isnot nothing
            If Robj IsNot Nothing Then
                If Robj.animateFill IsNot Nothing Then
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill", "Nothing", "Configured")
                End If
            End If
        End If

        '----------------------------------
        ' fill direction
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateFill IsNot Nothing Then
                    If Robj.animateFill IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateFill.fillDirection = Robj.animateFill.fillDirection Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-fillDirection", Lobj.animateFill.fillDirection.ToString, Robj.animateFill.fillDirection.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' fill direction specified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateFill IsNot Nothing Then
                    If Robj.animateFill IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateFill.fillDirectionSpecified = Robj.animateFill.fillDirectionSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-fillDirectionSpecified", Lobj.animateFill.fillDirectionSpecified.ToString, Robj.animateFill.fillDirectionSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' fill max
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateFill IsNot Nothing Then
                    If Robj.animateFill IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateFill.fillMax = Robj.animateFill.fillMax Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-fillMax", Lobj.animateFill.fillMax.ToString, Robj.animateFill.fillMax.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' fill max specified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateFill IsNot Nothing Then
                    If Robj.animateFill IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateFill.fillMaxSpecified = Robj.animateFill.fillMaxSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-fillMaxSpecified", Lobj.animateFill.fillMaxSpecified.ToString, Robj.animateFill.fillMaxSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' fill min
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateFill IsNot Nothing Then
                    If Robj.animateFill IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateFill.fillMin = Robj.animateFill.fillMin Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-fillMin", Lobj.animateFill.fillMin.ToString, Robj.animateFill.fillMin.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' fill min specified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateFill IsNot Nothing Then
                    If Robj.animateFill IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateFill.fillMinSpecified = Robj.animateFill.fillMinSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-fillMinSpecified", Lobj.animateFill.fillMinSpecified.ToString, Robj.animateFill.fillMinSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' fill inside only
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateFill IsNot Nothing Then
                    If Robj.animateFill IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateFill.insideOnly = Robj.animateFill.insideOnly Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-insideOnly", Lobj.animateFill.insideOnly.ToString, Robj.animateFill.insideOnly.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' fill inside only specified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateFill IsNot Nothing Then
                    If Robj.animateFill IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateFill.insideOnlySpecified = Robj.animateFill.insideOnlySpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-insideOnlySpecified", Lobj.animateFill.insideOnlySpecified.ToString, Robj.animateFill.insideOnlySpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateFill IsNot Nothing Then
                    If Robj.animateFill IsNot Nothing Then
                        ' Both have properties set so compare them
                        Try
                            Dim Lstr As String = Lobj.animateFill.Item.GetType.ToString
                            Dim Rstr As String = Robj.animateFill.Item.GetType.ToString

                            Select Case Lstr
                                Case "ME_Diff.defaultExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As defaultExpressionRangeType = Lobj.animateFill.Item
                                        Dim RFill As defaultExpressionRangeType = Robj.animateFill.Item

                                        ' Nothing to compare in here default item type has no comparable properties

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.constantExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As constantExpressionRangeType = Lobj.animateFill.Item
                                        Dim RFill As constantExpressionRangeType = Robj.animateFill.Item

                                        ' Item Max
                                        If Not LFill.max = RFill.max Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Animate Fill - Item max",
                                                                     LFill.max.ToString,
                                                                     RFill.max.ToString)
                                        End If

                                        ' Item Max specified
                                        If Not LFill.maxSpecified = RFill.maxSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Animate Fill - Item Max maxSpecified",
                                                                     LFill.maxSpecified.ToString,
                                                                     RFill.maxSpecified.ToString)
                                        End If

                                        ' Item Min
                                        If Not LFill.min = RFill.min Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Animate Fill - Item min",
                                                                     LFill.min.ToString,
                                                                     RFill.min.ToString)
                                        End If

                                        ' Item Min specified
                                        If Not LFill.minSpecified = RFill.minSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Animate Fill - Item minSpecified",
                                                                     LFill.minSpecified.ToString,
                                                                     RFill.minSpecified.ToString)
                                        End If
                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.readFromTagExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As readFromTagExpressionRangeType = Lobj.animateFill.Item
                                        Dim RFill As readFromTagExpressionRangeType = Robj.animateFill.Item

                                        ' Item max tag
                                        If Not LFill.maxTag = RFill.maxTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Animate Fill - Item maxTag",
                                                                     LFill.maxTag.ToString,
                                                                     RFill.maxTag.ToString)
                                        End If

                                        ' Item min tag
                                        If Not LFill.minTag = RFill.minTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Animate Fill - Item minTag",
                                                                     LFill.minTag.ToString,
                                                                     RFill.minTag.ToString)
                                        End If

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Fill-Item Type Mismatch", Lstr, Rstr)
                                    End If
                            End Select
                        Catch
                        End Try
                    End If
                End If
            End If
        End If

    End Sub
    Public Sub CompareAnimationsProperties_AnimateHeight(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        '----------------------------------
        ' Anchor
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHeight IsNot Nothing Then
                    If Robj.animateHeight IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHeight.anchor = Robj.animateHeight.anchor Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height-anchor", Lobj.animateHeight.anchor.ToString, Robj.animateHeight.anchor.ToString)
                        End If
                    Else
                        ' check left property is set
                        If Lobj.animateHeight IsNot Nothing Then
                            ' This is a difference where the left side has no property set
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height", "Configured", "Nothing")
                        End If
                    End If
                Else
                    ' check right property is set
                    If Robj.animateHeight IsNot Nothing Then
                        ' This is a difference where the left side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height", "Nothing", "Configured")
                    End If
                End If
            Else
                If Lobj IsNot Nothing Then
                    If Lobj.animateHeight IsNot Nothing Then
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height", "Configured", "Nothing")
                    End If
                End If
            End If

        Else
            ' Case for right object isnot nothing
            If Robj IsNot Nothing Then
                If Robj.animateHeight IsNot Nothing Then
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height", "Nothing", "Configured")
                End If
            End If
        End If

        '----------------------------------
        ' Anchor Specified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHeight IsNot Nothing Then
                    If Robj.animateHeight IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHeight.anchorSpecified = Robj.animateHeight.anchorSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height-anchorSpecified", Lobj.animateHeight.anchorSpecified.ToString, Robj.animateHeight.anchorSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' Expression
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHeight IsNot Nothing Then
                    If Robj.animateHeight IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHeight.expression = Robj.animateHeight.expression Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height-expression", Lobj.animateHeight.expression.ToString, Robj.animateHeight.expression.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' verticalChangeMax
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHeight IsNot Nothing Then
                    If Robj.animateHeight IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHeight.verticalChangeMax = Robj.animateHeight.verticalChangeMax Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height-verticalChangeMax", Lobj.animateHeight.verticalChangeMax.ToString, Robj.animateHeight.verticalChangeMax.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' verticalChangeMaxSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHeight IsNot Nothing Then
                    If Robj.animateHeight IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHeight.verticalChangeMaxSpecified = Robj.animateHeight.verticalChangeMaxSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height-verticalChangeMaxSpecified", Lobj.animateHeight.verticalChangeMaxSpecified.ToString, Robj.animateHeight.verticalChangeMaxSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' verticalChangeMin
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHeight IsNot Nothing Then
                    If Robj.animateHeight IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHeight.verticalChangeMin = Robj.animateHeight.verticalChangeMin Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height-verticalChangeMin", Lobj.animateHeight.verticalChangeMin.ToString, Robj.animateHeight.verticalChangeMin.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' verticalChangeMinSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHeight IsNot Nothing Then
                    If Robj.animateHeight IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHeight.verticalChangeMinSpecified = Robj.animateHeight.verticalChangeMinSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height-verticalChangeMinSpecified", Lobj.animateHeight.verticalChangeMinSpecified.ToString, Robj.animateHeight.verticalChangeMinSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHeight IsNot Nothing Then
                    If Robj.animateHeight IsNot Nothing Then
                        ' Both have properties set so compare them
                        Try
                            Dim Lstr As String = Lobj.animateHeight.Item.GetType.ToString
                            Dim Rstr As String = Robj.animateHeight.Item.GetType.ToString

                            Select Case Lstr
                                Case "ME_Diff.defaultExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As defaultExpressionRangeType = Lobj.animateHeight.Item
                                        Dim RFill As defaultExpressionRangeType = Robj.animateHeight.Item

                                        ' Nothing to compare in here default item type has no comparable properties

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.constantExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As constantExpressionRangeType = Lobj.animateHeight.Item
                                        Dim RFill As constantExpressionRangeType = Robj.animateHeight.Item

                                        ' Item Max
                                        If Not LFill.max = RFill.max Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Height - Item max",
                                                                     LFill.max.ToString,
                                                                     RFill.max.ToString)
                                        End If

                                        ' Item Max specified
                                        If Not LFill.maxSpecified = RFill.maxSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Height - Item maxSpecified",
                                                                     LFill.maxSpecified.ToString,
                                                                     RFill.maxSpecified.ToString)
                                        End If

                                        ' Item Min
                                        If Not LFill.min = RFill.min Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Height - Item min",
                                                                     LFill.min.ToString,
                                                                     RFill.min.ToString)
                                        End If

                                        ' Item Min specified
                                        If Not LFill.minSpecified = RFill.minSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Height - Item minSpecified",
                                                                     LFill.minSpecified.ToString,
                                                                     RFill.minSpecified.ToString)
                                        End If
                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.readFromTagExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As readFromTagExpressionRangeType = Lobj.animateHeight.Item
                                        Dim RFill As readFromTagExpressionRangeType = Robj.animateHeight.Item

                                        ' Item max tag
                                        If Not LFill.maxTag = RFill.maxTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Height - Item maxTag",
                                                                     LFill.maxTag.ToString,
                                                                     RFill.maxTag.ToString)
                                        End If

                                        ' Item min tag
                                        If Not LFill.minTag = RFill.minTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Height - Item minTag",
                                                                     LFill.minTag.ToString,
                                                                     RFill.minTag.ToString)
                                        End If

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Height-Item Type Mismatch", Lstr, Rstr)
                                    End If
                            End Select
                        Catch
                        End Try
                    End If
                End If
            End If
        End If

    End Sub
    Public Sub CompareAnimationsProperties_AnimateHorizontalPosition(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        '----------------------------------
        ' Expression
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalPosition IsNot Nothing Then
                    If Robj.animateHorizontalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHorizontalPosition.expression = Robj.animateHorizontalPosition.expression Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition-expression", Lobj.animateHorizontalPosition.expression.ToString, Robj.animateHorizontalPosition.expression.ToString)
                        End If
                    Else
                        ' check left property is set
                        If Lobj.animateHorizontalPosition IsNot Nothing Then
                            ' This is a difference where the left side has no property set
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition", "Configured", "Nothing")
                        End If
                    End If
                Else
                    ' check right property is set
                    If Robj.animateHorizontalPosition IsNot Nothing Then
                        ' This is a difference where the left side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition", "Nothing", "Configured")
                    End If
                End If
            Else
                If Lobj IsNot Nothing Then
                    If Lobj.animateHorizontalPosition IsNot Nothing Then
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition", "Configured", "Nothing")
                    End If
                End If
            End If

        Else
            ' Case for right object isnot nothing
            If Robj IsNot Nothing Then
                If Robj.animateHorizontalPosition IsNot Nothing Then
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition", "Nothing", "Configured")
                End If
            End If
        End If

        '----------------------------------
        ' horizontalOffsetMax
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalPosition IsNot Nothing Then
                    If Robj.animateHorizontalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHorizontalPosition.horizontalOffsetMax = Robj.animateHorizontalPosition.horizontalOffsetMax Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition-horizontalOffsetMax", Lobj.animateHorizontalPosition.horizontalOffsetMax.ToString, Robj.animateHorizontalPosition.horizontalOffsetMax.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' horizontalOffsetMaxSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalPosition IsNot Nothing Then
                    If Robj.animateHorizontalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHorizontalPosition.horizontalOffsetMaxSpecified = Robj.animateHorizontalPosition.horizontalOffsetMaxSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition-horizontalOffsetMaxSpecified", Lobj.animateHorizontalPosition.horizontalOffsetMaxSpecified.ToString, Robj.animateHorizontalPosition.horizontalOffsetMaxSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' horizontalOffsetMin
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalPosition IsNot Nothing Then
                    If Robj.animateHorizontalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHorizontalPosition.horizontalOffsetMin = Robj.animateHorizontalPosition.horizontalOffsetMin Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition-horizontalOffsetMin", Lobj.animateHorizontalPosition.horizontalOffsetMin.ToString, Robj.animateHorizontalPosition.horizontalOffsetMin.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' horizontalOffsetMinSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalPosition IsNot Nothing Then
                    If Robj.animateHorizontalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHorizontalPosition.horizontalOffsetMinSpecified = Robj.animateHorizontalPosition.horizontalOffsetMinSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition-horizontalOffsetMinSpecified", Lobj.animateHorizontalPosition.horizontalOffsetMinSpecified.ToString, Robj.animateHorizontalPosition.horizontalOffsetMinSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalPosition IsNot Nothing Then
                    If Robj.animateHorizontalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        Try
                            Dim Lstr As String = Lobj.animateHorizontalPosition.Item.GetType.ToString
                            Dim Rstr As String = Robj.animateHorizontalPosition.Item.GetType.ToString

                            Select Case Lstr
                                Case "ME_Diff.defaultExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As defaultExpressionRangeType = Lobj.animateHorizontalPosition.Item
                                        Dim RFill As defaultExpressionRangeType = Robj.animateHorizontalPosition.Item

                                        ' Nothing to compare in here default item type has no comparable properties

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.constantExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As constantExpressionRangeType = Lobj.animateHorizontalPosition.Item
                                        Dim RFill As constantExpressionRangeType = Robj.animateHorizontalPosition.Item

                                        ' Item Max
                                        If Not LFill.max = RFill.max Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalPosition - Item max",
                                                                     LFill.max.ToString,
                                                                     RFill.max.ToString)
                                        End If

                                        ' Item Max specified
                                        If Not LFill.maxSpecified = RFill.maxSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalPosition - Item maxSpecified",
                                                                     LFill.maxSpecified.ToString,
                                                                     RFill.maxSpecified.ToString)
                                        End If

                                        ' Item Min
                                        If Not LFill.min = RFill.min Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalPosition - Item min",
                                                                     LFill.min.ToString,
                                                                     RFill.min.ToString)
                                        End If

                                        ' Item Min specified
                                        If Not LFill.minSpecified = RFill.minSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalPosition - Item minSpecified",
                                                                     LFill.minSpecified.ToString,
                                                                     RFill.minSpecified.ToString)
                                        End If
                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.readFromTagExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As readFromTagExpressionRangeType = Lobj.animateHorizontalPosition.Item
                                        Dim RFill As readFromTagExpressionRangeType = Robj.animateHorizontalPosition.Item

                                        ' Item max tag
                                        If Not LFill.maxTag = RFill.maxTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalPosition - Item maxTag",
                                                                     LFill.maxTag.ToString,
                                                                     RFill.maxTag.ToString)
                                        End If

                                        ' Item min tag
                                        If Not LFill.minTag = RFill.minTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalPosition - Item minTag",
                                                                     LFill.minTag.ToString,
                                                                     RFill.minTag.ToString)
                                        End If

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalPosition-Item Type Mismatch", Lstr, Rstr)
                                    End If
                            End Select
                        Catch
                        End Try
                    End If
                End If
            End If
        End If

    End Sub
    Public Sub CompareAnimationsProperties_Animatehorizontalslider(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        '----------------------------------
        ' Expression
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalSlider IsNot Nothing Then
                    If Robj.animateHorizontalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHorizontalSlider.expression = Robj.animateHorizontalSlider.expression Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSlider-expression", Lobj.animateHorizontalSlider.expression.ToString, Robj.animateHorizontalSlider.expression.ToString)
                        End If
                    Else
                        ' check left property is set
                        If Lobj.animateHorizontalSlider IsNot Nothing Then
                            ' This is a difference where the left side has no property set
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSlider", "Configured", "Nothing")
                        End If
                    End If
                Else
                    ' check right property is set
                    If Robj.animateHorizontalSlider IsNot Nothing Then
                        ' This is a difference where the left side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSlider", "Nothing", "Configured")
                    End If
                End If
            Else
                If Lobj IsNot Nothing Then
                    If Lobj.animateHorizontalSlider IsNot Nothing Then
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSlider", "Configured", "Nothing")
                    End If
                End If
            End If

        Else
            ' Case for right object isnot nothing
            If Robj IsNot Nothing Then
                If Robj.animateHorizontalSlider IsNot Nothing Then
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSlider", "Nothing", "Configured")
                End If
            End If
        End If

        '----------------------------------
        ' horizontalOffsetMax
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalSlider IsNot Nothing Then
                    If Robj.animateHorizontalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHorizontalSlider.horizontalOffsetMax = Robj.animateHorizontalSlider.horizontalOffsetMax Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSlider-horizontalOffsetMax", Lobj.animateHorizontalSlider.horizontalOffsetMax.ToString, Robj.animateHorizontalSlider.horizontalOffsetMax.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' horizontalOffsetMaxSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalSlider IsNot Nothing Then
                    If Robj.animateHorizontalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHorizontalSlider.horizontalOffsetMaxSpecified = Robj.animateHorizontalSlider.horizontalOffsetMaxSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSlider-horizontalOffsetMaxSpecified", Lobj.animateHorizontalSlider.horizontalOffsetMaxSpecified.ToString, Robj.animateHorizontalSlider.horizontalOffsetMaxSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' horizontalOffsetMin
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalSlider IsNot Nothing Then
                    If Robj.animateHorizontalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHorizontalSlider.horizontalOffsetMin = Robj.animateHorizontalSlider.horizontalOffsetMin Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSlider-horizontalOffsetMin", Lobj.animateHorizontalSlider.horizontalOffsetMin.ToString, Robj.animateHorizontalSlider.horizontalOffsetMin.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' horizontalOffsetMinSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalSlider IsNot Nothing Then
                    If Robj.animateHorizontalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHorizontalSlider.horizontalOffsetMinSpecified = Robj.animateHorizontalSlider.horizontalOffsetMinSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSlider-horizontalOffsetMinSpecified", Lobj.animateHorizontalSlider.horizontalOffsetMinSpecified.ToString, Robj.animateHorizontalSlider.horizontalOffsetMinSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHorizontalSlider IsNot Nothing Then
                    If Robj.animateHorizontalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        Try
                            Dim Lstr As String = Lobj.animateHorizontalSlider.Item.GetType.ToString
                            Dim Rstr As String = Robj.animateHorizontalSlider.Item.GetType.ToString

                            Select Case Lstr
                                Case "ME_Diff.defaultExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As defaultExpressionRangeType = Lobj.animateHorizontalSlider.Item
                                        Dim RFill As defaultExpressionRangeType = Robj.animateHorizontalSlider.Item

                                        ' Nothing to compare in here default item type has no comparable properties

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSlider-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.constantExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As constantExpressionRangeType = Lobj.animateHorizontalSlider.Item
                                        Dim RFill As constantExpressionRangeType = Robj.animateHorizontalSlider.Item

                                        ' Item Max
                                        If Not LFill.max = RFill.max Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalSlider - Item Max",
                                                                     LFill.max.ToString,
                                                                     RFill.max.ToString)
                                        End If

                                        ' Item Max specified
                                        If Not LFill.maxSpecified = RFill.maxSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalSlider - Item Max Specified",
                                                                     LFill.maxSpecified.ToString,
                                                                     RFill.maxSpecified.ToString)
                                        End If

                                        ' Item Min
                                        If Not LFill.min = RFill.min Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalSlider - Item Min",
                                                                     LFill.min.ToString,
                                                                     RFill.min.ToString)
                                        End If

                                        ' Item Min specified
                                        If Not LFill.minSpecified = RFill.minSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalSlider - Item Min Specified",
                                                                     LFill.minSpecified.ToString,
                                                                     RFill.minSpecified.ToString)
                                        End If
                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSllider-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.readFromTagExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As readFromTagExpressionRangeType = Lobj.animateHorizontalSlider.Item
                                        Dim RFill As readFromTagExpressionRangeType = Robj.animateHorizontalSlider.Item

                                        ' Item max tag
                                        If Not LFill.maxTag = RFill.maxTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalSlider - Item Max tag",
                                                                     LFill.maxTag.ToString,
                                                                     RFill.maxTag.ToString)
                                        End If

                                        ' Item min tag
                                        If Not LFill.minTag = RFill.minTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "HorizontalSlider - Item Min tag",
                                                                     LFill.minTag.ToString,
                                                                     RFill.minTag.ToString)
                                        End If

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "HorizontalSlider-Item Type Mismatch", Lstr, Rstr)
                                    End If
                            End Select
                        Catch
                        End Try
                    End If
                End If
            End If
        End If

    End Sub
    Public Sub CompareAnimationsProperties_AnimateRotation(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        '----------------------------------
        ' Expression
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateRotation IsNot Nothing Then
                    If Robj.animateRotation IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateRotation.expression = Robj.animateRotation.expression Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation-expression", Lobj.animateRotation.expression.ToString, Robj.animateRotation.expression.ToString)
                        End If
                    Else
                        ' check left property is set
                        If Lobj.animateRotation IsNot Nothing Then
                            ' This is a difference where the left side has no property set
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation", "Configured", "Nothing")
                        End If
                    End If
                Else
                    ' check right property is set
                    If Robj.animateRotation IsNot Nothing Then
                        ' This is a difference where the left side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation", "Nothing", "Configured")
                    End If
                End If
            Else
                If Lobj IsNot Nothing Then
                    If Lobj.animateRotation IsNot Nothing Then
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation", "Configured", "Nothing")
                    End If
                End If
            End If

        Else
            ' Case for right object isnot nothing
            If Robj IsNot Nothing Then
                If Robj.animateRotation IsNot Nothing Then
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation", "Nothing", "Configured")
                End If
            End If
        End If

        '----------------------------------
        ' centerOfRotation
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateRotation IsNot Nothing Then
                    If Robj.animateRotation IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateRotation.centerOfRotation = Robj.animateRotation.centerOfRotation Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation-centerOfRotation", Lobj.animateRotation.centerOfRotation.ToString, Robj.animateRotation.centerOfRotation.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' rotationMax
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateRotation IsNot Nothing Then
                    If Robj.animateRotation IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateRotation.rotationMax = Robj.animateRotation.rotationMax Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation-rotationMax", Lobj.animateRotation.rotationMax.ToString, Robj.animateRotation.rotationMax.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' rotationMaxSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateRotation IsNot Nothing Then
                    If Robj.animateRotation IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateRotation.rotationMaxSpecified = Robj.animateRotation.rotationMaxSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation-rotationMaxSpecified", Lobj.animateRotation.rotationMaxSpecified.ToString, Robj.animateRotation.rotationMaxSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' rotationMin
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateRotation IsNot Nothing Then
                    If Robj.animateRotation IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateRotation.rotationMin = Robj.animateRotation.rotationMin Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation-rotationMin", Lobj.animateRotation.rotationMin.ToString, Robj.animateRotation.rotationMin.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' rotationMinSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateRotation IsNot Nothing Then
                    If Robj.animateRotation IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateRotation.rotationMinSpecified = Robj.animateRotation.rotationMinSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation-rotationMinSpecified", Lobj.animateRotation.rotationMinSpecified.ToString, Robj.animateRotation.rotationMinSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateRotation IsNot Nothing Then
                    If Robj.animateRotation IsNot Nothing Then
                        ' Both have properties set so compare them
                        Try
                            Dim Lstr As String = Lobj.animateRotation.Item.GetType.ToString
                            Dim Rstr As String = Robj.animateRotation.Item.GetType.ToString

                            Select Case Lstr
                                Case "ME_Diff.defaultExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As defaultExpressionRangeType = Lobj.animateRotation.Item
                                        Dim RFill As defaultExpressionRangeType = Robj.animateRotation.Item

                                        ' Nothing to compare in here default item type has no comparable properties

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.constantExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As constantExpressionRangeType = Lobj.animateRotation.Item
                                        Dim RFill As constantExpressionRangeType = Robj.animateRotation.Item

                                        ' Item Max
                                        If Not LFill.max = RFill.max Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Rotation - Item max",
                                                                     LFill.max.ToString,
                                                                     RFill.max.ToString)
                                        End If

                                        ' Item Max specified
                                        If Not LFill.maxSpecified = RFill.maxSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Rotation - Item maxSpecified",
                                                                     LFill.maxSpecified.ToString,
                                                                     RFill.maxSpecified.ToString)
                                        End If

                                        ' Item Min
                                        If Not LFill.min = RFill.min Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Rotation - Item min",
                                                                     LFill.min.ToString,
                                                                     RFill.min.ToString)
                                        End If

                                        ' Item Min specified
                                        If Not LFill.minSpecified = RFill.minSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Rotation - Item minSpecified",
                                                                     LFill.minSpecified.ToString,
                                                                     RFill.minSpecified.ToString)
                                        End If
                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.readFromTagExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As readFromTagExpressionRangeType = Lobj.animateRotation.Item
                                        Dim RFill As readFromTagExpressionRangeType = Robj.animateRotation.Item

                                        ' Item max tag
                                        If Not LFill.maxTag = RFill.maxTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Rotation - Item maxTag",
                                                                     LFill.maxTag.ToString,
                                                                     RFill.maxTag.ToString)
                                        End If

                                        ' Item min tag
                                        If Not LFill.minTag = RFill.minTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Rotation - Item minTag",
                                                                     LFill.minTag.ToString,
                                                                     RFill.minTag.ToString)
                                        End If

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Rotation-Item Type Mismatch", Lstr, Rstr)
                                    End If
                            End Select
                        Catch
                        End Try
                    End If
                End If
            End If
        End If

    End Sub

    Public Sub CompareAnimationsProperties_AnimateVerticalPosition(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        '----------------------------------
        ' Expression
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalPosition IsNot Nothing Then
                    If Robj.animateVerticalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVerticalPosition.expression = Robj.animateVerticalPosition.expression Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition-expression", Lobj.animateVerticalPosition.expression.ToString, Robj.animateVerticalPosition.expression.ToString)
                        End If
                    Else
                        ' check left property is set
                        If Lobj.animateVerticalPosition IsNot Nothing Then
                            ' This is a difference where the left side has no property set
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition", "Configured", "Nothing")
                        End If
                    End If
                Else
                    ' check right property is set
                    If Robj.animateVerticalPosition IsNot Nothing Then
                        ' This is a difference where the left side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition", "Nothing", "Configured")
                    End If
                End If
            Else
                If Lobj IsNot Nothing Then
                    If Lobj.animateVerticalPosition IsNot Nothing Then
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition", "Configured", "Nothing")
                    End If
                End If
            End If

        Else
            ' Case for right object isnot nothing
            If Robj IsNot Nothing Then
                If Robj.animateVerticalPosition IsNot Nothing Then
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition", "Nothing", "Configured")
                End If
            End If
        End If

        '----------------------------------
        ' VerticalOffsetMax
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalPosition IsNot Nothing Then
                    If Robj.animateVerticalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVerticalPosition.verticalOffsetMax = Robj.animateVerticalPosition.verticalOffsetMax Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition-verticalOffsetMax", Lobj.animateVerticalPosition.verticalOffsetMax.ToString, Robj.animateVerticalPosition.verticalOffsetMax.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' VerticalOffsetMaxSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalPosition IsNot Nothing Then
                    If Robj.animateVerticalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVerticalPosition.verticalOffsetMaxSpecified = Robj.animateVerticalPosition.verticalOffsetMaxSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition-verticalOffsetMaxSpecified", Lobj.animateVerticalPosition.verticalOffsetMaxSpecified.ToString, Robj.animateVerticalPosition.verticalOffsetMaxSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' VerticalOffsetMin
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalPosition IsNot Nothing Then
                    If Robj.animateVerticalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVerticalPosition.verticalOffsetMin = Robj.animateVerticalPosition.verticalOffsetMin Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition-verticalOffsetMin", Lobj.animateVerticalPosition.verticalOffsetMin.ToString, Robj.animateVerticalPosition.verticalOffsetMin.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' VerticalOffsetMinSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalPosition IsNot Nothing Then
                    If Robj.animateVerticalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVerticalPosition.verticalOffsetMinSpecified = Robj.animateVerticalPosition.verticalOffsetMinSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition-verticalOffsetMinSpecified", Lobj.animateVerticalPosition.verticalOffsetMinSpecified.ToString, Robj.animateVerticalPosition.verticalOffsetMinSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalPosition IsNot Nothing Then
                    If Robj.animateVerticalPosition IsNot Nothing Then
                        ' Both have properties set so compare them
                        Try
                            Dim Lstr As String = Lobj.animateVerticalPosition.Item.GetType.ToString
                            Dim Rstr As String = Robj.animateVerticalPosition.Item.GetType.ToString

                            Select Case Lstr
                                Case "ME_Diff.defaultExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As defaultExpressionRangeType = Lobj.animateVerticalPosition.Item
                                        Dim RFill As defaultExpressionRangeType = Robj.animateVerticalPosition.Item

                                        ' Nothing to compare in here default item type has no comparable properties

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.constantExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As constantExpressionRangeType = Lobj.animateVerticalPosition.Item
                                        Dim RFill As constantExpressionRangeType = Robj.animateVerticalPosition.Item

                                        ' Item Max
                                        If Not LFill.max = RFill.max Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalPosition - Item max",
                                                                     LFill.max.ToString,
                                                                     RFill.max.ToString)
                                        End If

                                        ' Item Max specified
                                        If Not LFill.maxSpecified = RFill.maxSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalPosition - Item maxSpecified",
                                                                     LFill.maxSpecified.ToString,
                                                                     RFill.maxSpecified.ToString)
                                        End If

                                        ' Item Min
                                        If Not LFill.min = RFill.min Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalPosition - Item min",
                                                                     LFill.min.ToString,
                                                                     RFill.min.ToString)
                                        End If

                                        ' Item Min specified
                                        If Not LFill.minSpecified = RFill.minSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalPosition - Item minSpecified",
                                                                     LFill.minSpecified.ToString,
                                                                     RFill.minSpecified.ToString)
                                        End If
                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.readFromTagExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As readFromTagExpressionRangeType = Lobj.animateVerticalPosition.Item
                                        Dim RFill As readFromTagExpressionRangeType = Robj.animateVerticalPosition.Item

                                        ' Item max tag
                                        If Not LFill.maxTag = RFill.maxTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalPosition - Item maxTag",
                                                                     LFill.maxTag.ToString,
                                                                     RFill.maxTag.ToString)
                                        End If

                                        ' Item min tag
                                        If Not LFill.minTag = RFill.minTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalPosition - Item minTag",
                                                                     LFill.minTag.ToString,
                                                                     RFill.minTag.ToString)
                                        End If

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalPosition-Item Type Mismatch", Lstr, Rstr)
                                    End If
                            End Select
                        Catch
                        End Try
                    End If
                End If
            End If
        End If


    End Sub

    Public Sub CompareAnimationsProperties_AnimateVerticalSlider(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        '----------------------------------
        ' Expression
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalSlider IsNot Nothing Then
                    If Robj.animateVerticalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVerticalSlider.expression = Robj.animateVerticalSlider.expression Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSlider-expression", Lobj.animateVerticalSlider.expression.ToString, Robj.animateVerticalSlider.expression.ToString)
                        End If
                    Else
                        ' check left property is set
                        If Lobj.animateVerticalSlider IsNot Nothing Then
                            ' This is a difference where the left side has no property set
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSlider", "Configured", "Nothing")
                        End If
                    End If
                Else
                    ' check right property is set
                    If Robj.animateVerticalSlider IsNot Nothing Then
                        ' This is a difference where the left side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSlider", "Nothing", "Configured")
                    End If
                End If
            Else
                If Lobj IsNot Nothing Then
                    If Lobj.animateVerticalSlider IsNot Nothing Then
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSlider", "Configured", "Nothing")
                    End If
                End If
            End If

        Else
            ' Case for right object isnot nothing
            If Robj IsNot Nothing Then
                If Robj.animateVerticalSlider IsNot Nothing Then
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSlider", "Nothing", "Configured")
                End If
            End If
        End If

        '----------------------------------
        ' VerticalOffsetMax
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalSlider IsNot Nothing Then
                    If Robj.animateVerticalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVerticalSlider.verticalOffsetMax = Robj.animateVerticalSlider.verticalOffsetMax Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSlider-VerticalOffsetMax", Lobj.animateVerticalSlider.verticalOffsetMax.ToString, Robj.animateVerticalSlider.verticalOffsetMax.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' VerticalOffsetMaxSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalSlider IsNot Nothing Then
                    If Robj.animateVerticalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVerticalSlider.verticalOffsetMaxSpecified = Robj.animateVerticalSlider.verticalOffsetMaxSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSlider-VerticalOffsetMaxSpecified", Lobj.animateVerticalSlider.verticalOffsetMaxSpecified.ToString, Robj.animateVerticalSlider.verticalOffsetMaxSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' VerticalOffsetMin
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalSlider IsNot Nothing Then
                    If Robj.animateVerticalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVerticalSlider.verticalOffsetMin = Robj.animateVerticalSlider.verticalOffsetMin Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSlider-VerticalOffsetMin", Lobj.animateVerticalSlider.verticalOffsetMin.ToString, Robj.animateVerticalSlider.verticalOffsetMin.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' VerticalOffsetMinSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalSlider IsNot Nothing Then
                    If Robj.animateVerticalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVerticalSlider.verticalOffsetMinSpecified = Robj.animateVerticalSlider.verticalOffsetMinSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSlider-VerticalOffsetMinSpecified", Lobj.animateVerticalSlider.verticalOffsetMinSpecified.ToString, Robj.animateVerticalSlider.verticalOffsetMinSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVerticalSlider IsNot Nothing Then
                    If Robj.animateVerticalSlider IsNot Nothing Then
                        ' Both have properties set so compare them
                        Try
                            Dim Lstr As String = Lobj.animateVerticalSlider.Item.GetType.ToString
                            Dim Rstr As String = Robj.animateVerticalSlider.Item.GetType.ToString

                            Select Case Lstr
                                Case "ME_Diff.defaultExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As defaultExpressionRangeType = Lobj.animateVerticalSlider.Item
                                        Dim RFill As defaultExpressionRangeType = Robj.animateVerticalSlider.Item

                                        ' Nothing to compare in here default item type has no comparable properties

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSlider-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.constantExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As constantExpressionRangeType = Lobj.animateVerticalSlider.Item
                                        Dim RFill As constantExpressionRangeType = Robj.animateVerticalSlider.Item

                                        ' Item Max
                                        If Not LFill.max = RFill.max Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalSlider - Item Max",
                                                                     LFill.max.ToString,
                                                                     RFill.max.ToString)
                                        End If

                                        ' Item Max specified
                                        If Not LFill.maxSpecified = RFill.maxSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalSlider - Item Max Specified",
                                                                     LFill.maxSpecified.ToString,
                                                                     RFill.maxSpecified.ToString)
                                        End If

                                        ' Item Min
                                        If Not LFill.min = RFill.min Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalSlider - Item Min",
                                                                     LFill.min.ToString,
                                                                     RFill.min.ToString)
                                        End If

                                        ' Item Min specified
                                        If Not LFill.minSpecified = RFill.minSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalSlider - Item Min Specified",
                                                                     LFill.minSpecified.ToString,
                                                                     RFill.minSpecified.ToString)
                                        End If
                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSllider-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.readFromTagExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As readFromTagExpressionRangeType = Lobj.animateVerticalSlider.Item
                                        Dim RFill As readFromTagExpressionRangeType = Robj.animateVerticalSlider.Item

                                        ' Item max tag
                                        If Not LFill.maxTag = RFill.maxTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalSlider - Item Max tag",
                                                                     LFill.maxTag.ToString,
                                                                     RFill.maxTag.ToString)
                                        End If

                                        ' Item min tag
                                        If Not LFill.minTag = RFill.minTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "VerticalSlider - Item Min tag",
                                                                     LFill.minTag.ToString,
                                                                     RFill.minTag.ToString)
                                        End If

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "VerticalSlider-Item Type Mismatch", Lstr, Rstr)
                                    End If
                            End Select
                        Catch
                        End Try
                    End If
                End If
            End If
        End If

    End Sub

    Public Sub CompareAnimationsProperties_AnimateVisibility(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        '----------------------------------
        ' Expression
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVisibility IsNot Nothing Then
                    If Robj.animateVisibility IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVisibility.expression = Robj.animateVisibility.expression Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Visibility-expression", Lobj.animateVisibility.expression.ToString, Robj.animateVisibility.expression.ToString)
                        End If
                    Else
                        ' This is a difference where the right side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Visibility", "Configured", "Nothing")
                    End If
                Else
                    If Robj.animateColor IsNot Nothing Then
                        ' This is a difference where the left side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Visibility", "Nothing", "Configured")
                    End If
                End If
            Else
                ' right is empty and left is not
                If Lobj.animateVisibility IsNot Nothing Then
                    ' report the right as missing the animation
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "Visibility", "Configured", "Nothing")
                End If
            End If
        Else
            If Robj IsNot Nothing Then
                If Robj.animateVisibility IsNot Nothing Then
                    ' left is empty and right is not
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "Visibility", "Nothing", "Configured")
                End If
            End If
        End If

        '----------------------------------
        ' expressionTrueState
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVisibility IsNot Nothing Then
                    If Robj.animateVisibility IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVisibility.expressionTrueState = Robj.animateVisibility.expressionTrueState Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Visibility-expressionTrueState", Lobj.animateVisibility.expressionTrueState.ToString, Robj.animateVisibility.expressionTrueState.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' expressionTrueStateSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateVisibility IsNot Nothing Then
                    If Robj.animateVisibility IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateVisibility.expressionTrueStateSpecified = Robj.animateVisibility.expressionTrueStateSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Visibility-expressionTrueStateSpecified", Lobj.animateVisibility.expressionTrueStateSpecified.ToString, Robj.animateVisibility.expressionTrueStateSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Public Sub CompareAnimationsProperties_AnimateWidth(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        '----------------------------------
        ' Anchor
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateWidth IsNot Nothing Then
                    If Robj.animateWidth IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateWidth.anchor = Robj.animateWidth.anchor Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width-anchor", Lobj.animateWidth.anchor.ToString, Robj.animateWidth.anchor.ToString)
                        End If
                    Else
                        ' check left property is set
                        If Lobj.animateWidth IsNot Nothing Then
                            ' This is a difference where the left side has no property set
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width", "Configured", "Nothing")
                        End If
                    End If
                Else
                    ' check right property is set
                    If Robj.animateWidth IsNot Nothing Then
                        ' This is a difference where the left side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width", "Nothing", "Configured")
                    End If
                End If
            Else
                If Lobj IsNot Nothing Then
                    If Lobj.animateWidth IsNot Nothing Then
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width", "Configured", "Nothing")
                    End If
                End If
            End If

        Else
            ' Case for right object isnot nothing
            If Robj IsNot Nothing Then
                If Robj.animateWidth IsNot Nothing Then
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width", "Nothing", "Configured")
                End If
            End If
        End If

        '----------------------------------
        ' Anchor Specified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateWidth IsNot Nothing Then
                    If Robj.animateWidth IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateWidth.anchorSpecified = Robj.animateWidth.anchorSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width-anchorSpecified", Lobj.animateWidth.anchorSpecified.ToString, Robj.animateWidth.anchorSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' Expression
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateWidth IsNot Nothing Then
                    If Robj.animateWidth IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateWidth.expression = Robj.animateWidth.expression Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width-expression", Lobj.animateWidth.expression.ToString, Robj.animateWidth.expression.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' HorizontalChangeMax
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateWidth IsNot Nothing Then
                    If Robj.animateWidth IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateWidth.horizontalChangeMax = Robj.animateWidth.horizontalChangeMax Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width-horizontalChangeMax", Lobj.animateWidth.horizontalChangeMax.ToString, Robj.animateWidth.horizontalChangeMax.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' HorizontalChangeMaxSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateWidth IsNot Nothing Then
                    If Robj.animateWidth IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateWidth.horizontalChangeMaxSpecified = Robj.animateWidth.horizontalChangeMaxSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width-horizontalChangeMaxSpecified", Lobj.animateWidth.horizontalChangeMaxSpecified.ToString, Robj.animateWidth.horizontalChangeMaxSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' HorizontalChangeMin
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateWidth IsNot Nothing Then
                    If Robj.animateWidth IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateWidth.horizontalChangeMin = Robj.animateWidth.horizontalChangeMin Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width-horizontalChangeMin", Lobj.animateWidth.horizontalChangeMin.ToString, Robj.animateWidth.horizontalChangeMin.ToString)
                        End If
                    End If
                End If
            End If
        End If

        '----------------------------------
        ' HorizontalChangeMinSpecified
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateWidth IsNot Nothing Then
                    If Robj.animateWidth IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateWidth.horizontalChangeMinSpecified = Robj.animateWidth.horizontalChangeMinSpecified Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width-horizontalChangeMinSpecified", Lobj.animateWidth.horizontalChangeMinSpecified.ToString, Robj.animateWidth.horizontalChangeMinSpecified.ToString)
                        End If
                    End If
                End If
            End If
        End If

        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateWidth IsNot Nothing Then
                    If Robj.animateWidth IsNot Nothing Then
                        ' Both have properties set so compare them
                        Try
                            Dim Lstr As String = Lobj.animateWidth.Item.GetType.ToString
                            Dim Rstr As String = Robj.animateWidth.Item.GetType.ToString

                            Select Case Lstr
                                Case "ME_Diff.defaultExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As defaultExpressionRangeType = Lobj.animateWidth.Item
                                        Dim RFill As defaultExpressionRangeType = Robj.animateWidth.Item

                                        ' Nothing to compare in here default item type has no comparable properties

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.constantExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As constantExpressionRangeType = Lobj.animateWidth.Item
                                        Dim RFill As constantExpressionRangeType = Robj.animateWidth.Item

                                        ' Item Max
                                        If Not LFill.max = RFill.max Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Width - Item max",
                                                                     LFill.max.ToString,
                                                                     RFill.max.ToString)
                                        End If

                                        ' Item Max specified
                                        If Not LFill.maxSpecified = RFill.maxSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Width - Item maxSpecified",
                                                                     LFill.maxSpecified.ToString,
                                                                     RFill.maxSpecified.ToString)
                                        End If

                                        ' Item Min
                                        If Not LFill.min = RFill.min Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Width - Item min",
                                                                     LFill.min.ToString,
                                                                     RFill.min.ToString)
                                        End If

                                        ' Item Min specified
                                        If Not LFill.minSpecified = RFill.minSpecified Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Width - Item minSpecified",
                                                                     LFill.minSpecified.ToString,
                                                                     RFill.minSpecified.ToString)
                                        End If
                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width-Item Type Mismatch", Lstr, Rstr)
                                    End If
                                Case "ME_Diff.readFromTagExpressionRangeType"
                                    If Lstr = Rstr Then
                                        Dim LFill As readFromTagExpressionRangeType = Lobj.animateWidth.Item
                                        Dim RFill As readFromTagExpressionRangeType = Robj.animateWidth.Item

                                        ' Item max tag
                                        If Not LFill.maxTag = RFill.maxTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Width - Item maxTag",
                                                                     LFill.maxTag.ToString,
                                                                     RFill.maxTag.ToString)
                                        End If

                                        ' Item min tag
                                        If Not LFill.minTag = RFill.minTag Then
                                            Call AddListContentMatch(Gnest,
                                                                     fname,
                                                                     oname,
                                                                     "Animation",
                                                                     "Width - Item minTag",
                                                                     LFill.minTag.ToString,
                                                                     RFill.minTag.ToString)
                                        End If

                                    Else
                                        ' Left and right types dont match
                                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Width-Item Type Mismatch", Lstr, Rstr)
                                    End If
                            End Select
                        Catch
                        End Try
                    End If
                End If
            End If
        End If

    End Sub

    Public Sub CompareAnimationsProperties_AnimateHyperlink(ByRef Lobj As animationsAllMEType, ByRef Robj As animationsAllMEType, ByRef Gnest As String, ByRef fname As String, ByRef oname As String)

        '----------------------------------
        ' Expression
        '----------------------------------
        If Lobj IsNot Nothing Then
            If Robj IsNot Nothing Then
                ' both have objects
                If Lobj.animateHyperlink IsNot Nothing Then
                    If Robj.animateHyperlink IsNot Nothing Then
                        ' Both have properties set so compare them
                        If Not Lobj.animateHyperlink.expression = Robj.animateHyperlink.expression Then
                            Call AddListContentMatch(Gnest, fname, oname, "Animation", "Hyperlink-Expression", Lobj.animateHyperlink.expression.ToString, Robj.animateHyperlink.expression.ToString)
                        End If
                    Else
                        ' This is a difference where the right side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Hyperlink", "Configured", "Nothing")
                    End If
                Else
                    If Robj.animateHyperlink IsNot Nothing Then
                        ' This is a difference where the left side has no property set
                        Call AddListContentMatch(Gnest, fname, oname, "Animation", "Hyperlink", "Nothing", "Configured")
                    End If
                End If
            Else
                ' right is empty and left is not
                If Lobj.animateHyperlink IsNot Nothing Then
                    ' report the right as missing the animation
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "Hyperlink", "Configured", "Nothing")
                End If
            End If
        Else
            If Robj IsNot Nothing Then
                If Robj.animateHyperlink IsNot Nothing Then
                    ' left is empty and right is not
                    Call AddListContentMatch(Gnest, fname, oname, "Animation", "Hyperlink", "Nothing", "Configured")
                End If
            End If
        End If

    End Sub

End Module
