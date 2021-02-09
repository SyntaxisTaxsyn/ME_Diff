﻿Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class ReportObjectTests

    <TestMethod()> Public Sub CreateNewItem()
        Dim Testitem As ME_Diff.ReportMember = New ME_Diff.ReportMember("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")
        Assert.IsTrue(Testitem.FileName = "00000Left")
        Assert.IsTrue(Testitem.NestLevel = "")
        Assert.IsTrue(Testitem.ObjectName = "Display")
        Assert.IsTrue(Testitem.ObjectType = "gfx")
        Assert.IsTrue(Testitem.Parameter = "displayType")
        Assert.IsTrue(Testitem.LeftValue = "replace")
        Assert.IsTrue(Testitem.RightValue = "onTop")

        Testitem = New ME_Diff.ReportMember("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        Assert.IsTrue(Testitem.FileName = "00000Left")
        Assert.IsTrue(Testitem.NestLevel = "")
        Assert.IsTrue(Testitem.ObjectName = "Display")
        Assert.IsTrue(Testitem.ObjectType = "gfx")
        Assert.IsTrue(Testitem.Parameter = "titleBarSpecified")
        Assert.IsTrue(Testitem.LeftValue = "False")
        Assert.IsTrue(Testitem.RightValue = "True")

        Testitem = New ME_Diff.ReportMember("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        Assert.IsTrue(Testitem.FileName = "00000Left")
        Assert.IsTrue(Testitem.NestLevel = "Root/Group2/Group1")
        Assert.IsTrue(Testitem.ObjectName = "Text2")
        Assert.IsTrue(Testitem.ObjectType = "textDrawingObjectType")
        Assert.IsTrue(Testitem.Parameter = "width")
        Assert.IsTrue(Testitem.LeftValue = "80")
        Assert.IsTrue(Testitem.RightValue = "88")

        Testitem = New ME_Diff.ReportMember("File Name : 00000Left - Nest Level : Root/Group2/Group3 - Object Name : Text7 - Object Type : textDrawingObjectType - Parameter : backStyle - Left Value = transparent, Right Value = solid")
        Assert.IsTrue(Testitem.FileName = "00000Left")
        Assert.IsTrue(Testitem.NestLevel = "Root/Group2/Group3")
        Assert.IsTrue(Testitem.ObjectName = "Text7")
        Assert.IsTrue(Testitem.ObjectType = "textDrawingObjectType")
        Assert.IsTrue(Testitem.Parameter = "backStyle")
        Assert.IsTrue(Testitem.LeftValue = "transparent")
        Assert.IsTrue(Testitem.RightValue = "solid")

        Testitem = New ME_Diff.ReportMember("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        Assert.IsTrue(Testitem.FileName = "00000Left")
        Assert.IsTrue(Testitem.NestLevel = "Root")
        Assert.IsTrue(Testitem.ObjectName = "EditComboControl5")
        Assert.IsTrue(Testitem.ObjectType = "activeXObjectType")
        Assert.IsTrue(Testitem.Parameter = "data")
        Assert.IsTrue(Testitem.LeftValue = "00102171600200110040002555004002552553484648100")
        Assert.IsTrue(Testitem.RightValue = "0010217160020011004000255500400103484648100")
    End Sub

    <TestMethod()> Public Sub CreateNewPage()

        Dim newpage As ME_Diff.ReportOutputPage = New ME_Diff.ReportOutputPage("Left", "Right")
        Assert.IsTrue(newpage.LeftFile = "Left")
        Assert.IsTrue(newpage.RightFile = "Right")

        newpage.AddListItem("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")
        Assert.IsTrue(newpage.OutputList.Item(0).FileName = "00000Left")
        Assert.IsTrue(newpage.OutputList.Item(0).NestLevel = "")
        Assert.IsTrue(newpage.OutputList.Item(0).ObjectName = "Display")
        Assert.IsTrue(newpage.OutputList.Item(0).ObjectType = "gfx")
        Assert.IsTrue(newpage.OutputList.Item(0).Parameter = "displayType")
        Assert.IsTrue(newpage.OutputList.Item(0).LeftValue = "replace")
        Assert.IsTrue(newpage.OutputList.Item(0).RightValue = "onTop")

        newpage.AddListItem("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        Assert.IsTrue(newpage.OutputList.Item(1).FileName = "00000Left")
        Assert.IsTrue(newpage.OutputList.Item(1).NestLevel = "")
        Assert.IsTrue(newpage.OutputList.Item(1).ObjectName = "Display")
        Assert.IsTrue(newpage.OutputList.Item(1).ObjectType = "gfx")
        Assert.IsTrue(newpage.OutputList.Item(1).Parameter = "titleBarSpecified")
        Assert.IsTrue(newpage.OutputList.Item(1).LeftValue = "False")
        Assert.IsTrue(newpage.OutputList.Item(1).RightValue = "True")

        newpage.AddListItem("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        Assert.IsTrue(newpage.OutputList.Item(2).FileName = "00000Left")
        Assert.IsTrue(newpage.OutputList.Item(2).NestLevel = "Root/Group2/Group1")
        Assert.IsTrue(newpage.OutputList.Item(2).ObjectName = "Text2")
        Assert.IsTrue(newpage.OutputList.Item(2).ObjectType = "textDrawingObjectType")
        Assert.IsTrue(newpage.OutputList.Item(2).Parameter = "width")
        Assert.IsTrue(newpage.OutputList.Item(2).LeftValue = "80")
        Assert.IsTrue(newpage.OutputList.Item(2).RightValue = "88")

        newpage.AddListItem("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        Assert.IsTrue(newpage.OutputList.Item(3).FileName = "00000Left")
        Assert.IsTrue(newpage.OutputList.Item(3).NestLevel = "Root")
        Assert.IsTrue(newpage.OutputList.Item(3).ObjectName = "EditComboControl5")
        Assert.IsTrue(newpage.OutputList.Item(3).ObjectType = "activeXObjectType")
        Assert.IsTrue(newpage.OutputList.Item(3).Parameter = "data")
        Assert.IsTrue(newpage.OutputList.Item(3).LeftValue = "00102171600200110040002555004002552553484648100")
        Assert.IsTrue(newpage.OutputList.Item(3).RightValue = "0010217160020011004000255500400103484648100")
    End Sub

    <TestMethod()> Public Sub CreateNewReport()

        Dim Newreport As New ME_Diff.ReportOutput()
        Newreport.AddPage("Left", "Right")
        Assert.IsTrue(Newreport.PageList.Item(0).LeftFile = "Left")
        Assert.IsTrue(Newreport.PageList.Item(0).RightFile = "Right")

        Newreport.PageList.Item(0).AddListItem("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")

        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(0).FileName = "00000Left")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(0).NestLevel = "")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(0).ObjectName = "Display")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(0).ObjectType = "gfx")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(0).Parameter = "displayType")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(0).LeftValue = "replace")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(0).RightValue = "onTop")

        Newreport.PageList.Item(0).AddListItem("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(1).FileName = "00000Left")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(1).NestLevel = "")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(1).ObjectName = "Display")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(1).ObjectType = "gfx")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(1).Parameter = "titleBarSpecified")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(1).LeftValue = "False")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(1).RightValue = "True")

        Newreport.PageList.Item(0).AddListItem("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(2).FileName = "00000Left")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(2).NestLevel = "Root/Group2/Group1")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(2).ObjectName = "Text2")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(2).ObjectType = "textDrawingObjectType")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(2).Parameter = "width")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(2).LeftValue = "80")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(2).RightValue = "88")

        Newreport.PageList.Item(0).AddListItem("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(3).FileName = "00000Left")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(3).NestLevel = "Root")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(3).ObjectName = "EditComboControl5")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(3).ObjectType = "activeXObjectType")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(3).Parameter = "data")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(3).LeftValue = "00102171600200110040002555004002552553484648100")
        Assert.IsTrue(Newreport.PageList.Item(0).OutputList.Item(3).RightValue = "0010217160020011004000255500400103484648100")


        Newreport.AddPage("Left1", "Right1")
        Assert.IsTrue(Newreport.PageList.Item(1).LeftFile = "Left1")
        Assert.IsTrue(Newreport.PageList.Item(1).RightFile = "Right1")

        Newreport.PageList.Item(1).AddListItem("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")

        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(0).FileName = "00000Left")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(0).NestLevel = "")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(0).ObjectName = "Display")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(0).ObjectType = "gfx")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(0).Parameter = "displayType")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(0).LeftValue = "replace")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(0).RightValue = "onTop")

        Newreport.PageList.Item(1).AddListItem("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(1).FileName = "00000Left")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(1).NestLevel = "")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(1).ObjectName = "Display")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(1).ObjectType = "gfx")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(1).Parameter = "titleBarSpecified")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(1).LeftValue = "False")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(1).RightValue = "True")

        Newreport.PageList.Item(1).AddListItem("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(2).FileName = "00000Left")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(2).NestLevel = "Root/Group2/Group1")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(2).ObjectName = "Text2")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(2).ObjectType = "textDrawingObjectType")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(2).Parameter = "width")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(2).LeftValue = "80")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(2).RightValue = "88")

        Newreport.PageList.Item(1).AddListItem("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(3).FileName = "00000Left")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(3).NestLevel = "Root")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(3).ObjectName = "EditComboControl5")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(3).ObjectType = "activeXObjectType")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(3).Parameter = "data")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(3).LeftValue = "00102171600200110040002555004002552553484648100")
        Assert.IsTrue(Newreport.PageList.Item(1).OutputList.Item(3).RightValue = "0010217160020011004000255500400103484648100")

    End Sub

    <TestMethod()> Public Sub CheckReportFilter()
        Dim result As Boolean = False
        ME_Diff.FilterList = New List(Of String)
        ME_Diff.FilterList.Add("backStyle")
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")
        result = ME_Diff.NeedsFiltered("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        Assert.IsTrue(result = True)
        result = ME_Diff.NeedsFiltered("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        Assert.IsTrue(result = True)
        result = ME_Diff.NeedsFiltered("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        Assert.IsTrue(result = False)
        result = ME_Diff.NeedsFiltered("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")
        Assert.IsTrue(result = False)
        ME_Diff.FilterList.Add("displayType")
        result = ME_Diff.NeedsFiltered("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")
        Assert.IsTrue(result = True)
    End Sub

    <TestMethod()> Public Sub CheckReportCreate()
        Dim ResultList As List(Of String) = New List(Of String)
        Dim ReportObject As ME_Diff.ReportOutput = New ME_Diff.ReportOutput
        ME_Diff.FilterList = New List(Of String)
        ME_Diff.FilterList.Add("backStyle")
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")
        ResultList.Add("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group3 - Object Name : Text7 - Object Type : textDrawingObjectType - Parameter : backStyle - Left Value = transparent, Right Value = solid")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfile", "RightFile")

        Assert.IsTrue(ReportObject.PageList.Item(0).LeftFile = "Leftfile")
        Assert.IsTrue(ReportObject.PageList.Item(0).RightFile = "RightFile")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(0).FileName = "00000Left")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(0).NestLevel = "")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(0).ObjectName = "Display")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(0).ObjectType = "gfx")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(0).Parameter = "titleBarSpecified")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(0).LeftValue = "False")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(0).RightValue = "True")

        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(1).FileName = "00000Left")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(1).NestLevel = "")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(1).ObjectName = "Display")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(1).ObjectType = "gfx")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(1).Parameter = "displayType")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(1).LeftValue = "replace")
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Item(1).RightValue = "onTop")

        Assert.IsTrue(ReportObject.PageList.Count = 1)
        Assert.IsTrue(ReportObject.PageList.Item(0).OutputList.Count = 2)

        ME_Diff.FilterList.Clear()
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfiles", "RightFiles")

        Assert.IsTrue(ReportObject.PageList.Item(1).LeftFile = "Leftfiles")
        Assert.IsTrue(ReportObject.PageList.Item(1).RightFile = "RightFiles")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(0).FileName = "00000Left")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(0).NestLevel = "")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(0).ObjectName = "Display")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(0).ObjectType = "gfx")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(0).Parameter = "titleBarSpecified")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(0).LeftValue = "False")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(0).RightValue = "True")

        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(1).FileName = "00000Left")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(1).NestLevel = "")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(1).ObjectName = "Display")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(1).ObjectType = "gfx")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(1).Parameter = "displayType")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(1).LeftValue = "replace")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(1).RightValue = "onTop")

        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(2).FileName = "00000Left")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(2).NestLevel = "Root/Group2/Group3")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(2).ObjectName = "Text7")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(2).ObjectType = "textDrawingObjectType")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(2).Parameter = "backStyle")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(2).LeftValue = "transparent")
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Item(2).RightValue = "solid")

        Assert.IsTrue(ReportObject.PageList.Count = 2)
        Assert.IsTrue(ReportObject.PageList.Item(1).OutputList.Count = 3)


    End Sub
    <TestMethod()> Public Sub CheckDefaultBrowserStart()
        Dim result As Boolean = False
        result = ME_Diff.Run(ME_Diff.DefaultWebBrowser, "C:\temp\default.html")
        Assert.IsTrue(result = True)
        result = ME_Diff.Run("d.exe", "C:\temp\default.html")
        Assert.IsTrue(result = False)
    End Sub

    <TestMethod()> Public Sub CheckReportSummaryOutput_1PageHasDiffs()
        Dim results(50) As Boolean
        For a = 0 To results.Length - 1
            results(a) = False
        Next
        ' Create the mock report data
        Dim ResultList As List(Of String) = New List(Of String)
        Dim ReportObject As ME_Diff.ReportOutput = New ME_Diff.ReportOutput
        ME_Diff.FilterList = New List(Of String)
        ME_Diff.FilterList.Add("backStyle")
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")
        ResultList.Add("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group3 - Object Name : Text7 - Object Type : textDrawingObjectType - Parameter : backStyle - Left Value = transparent, Right Value = solid")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfile", "RightFile")
        ' Generate the report output
        ReportObject.GenerateHTMLReport()
        ' collect the report output into variables
        Dim TotalDifferenceString As String = ReportObject.ReportSummary.Item(8)
        Dim TotalPagesString As String = ReportObject.ReportSummary.Item(9)
        Dim PagesNoDifferenceString As String = ReportObject.ReportSummary.Item(10)
        Dim PagesWithDifferenceString As String = ReportObject.ReportSummary.Item(11)

        ' Check the content
        If InStr(TotalDifferenceString, "- 2") > 0 Then
            results(0) = True
        End If
        If InStr(TotalPagesString, "- 1") > 0 Then
            results(1) = True
        End If
        If InStr(PagesNoDifferenceString, "- 0") > 0 Then
            results(2) = True
        End If
        If InStr(PagesWithDifferenceString, "- 1") > 0 Then
            results(3) = True
        End If

        Assert.IsTrue(results(0))
        Assert.IsTrue(results(1))
        Assert.IsTrue(results(2))
        Assert.IsTrue(results(3))

    End Sub

    <TestMethod()> Public Sub CheckReportSummaryOutput_2PagesHaveDiffs()
        Dim results(50) As Boolean
        For a = 0 To results.Length - 1
            results(a) = False
        Next
        ' Create the mock report data
        Dim ResultList As List(Of String) = New List(Of String)
        Dim ReportObject As ME_Diff.ReportOutput = New ME_Diff.ReportOutput
        ME_Diff.FilterList = New List(Of String)
        ME_Diff.FilterList.Add("backStyle")
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")
        ResultList.Add("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group3 - Object Name : Text7 - Object Type : textDrawingObjectType - Parameter : backStyle - Left Value = transparent, Right Value = solid")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfile", "RightFile")
        ME_Diff.FilterList.Clear()
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfiles", "RightFiles")

        ' Generate the report output
        ReportObject.GenerateHTMLReport()
        ' collect the report output into variables
        Dim TotalDifferenceString As String = ReportObject.ReportSummary.Item(8)
        Dim TotalPagesString As String = ReportObject.ReportSummary.Item(9)
        Dim PagesNoDifferenceString As String = ReportObject.ReportSummary.Item(10)
        Dim PagesWithDifferenceString As String = ReportObject.ReportSummary.Item(11)

        ' Check the content
        If InStr(TotalDifferenceString, "- 5") > 0 Then
            results(0) = True
        End If
        If InStr(TotalPagesString, "- 2") > 0 Then
            results(1) = True
        End If
        If InStr(PagesNoDifferenceString, "- 0") > 0 Then
            results(2) = True
        End If
        If InStr(PagesWithDifferenceString, "- 2") > 0 Then
            results(3) = True
        End If

        Assert.IsTrue(results(0))
        Assert.IsTrue(results(1))
        Assert.IsTrue(results(2))
        Assert.IsTrue(results(3))

    End Sub

    <TestMethod()> Public Sub CheckReportSummaryOutput_3Pages1NoDiffs()

        Dim results(50) As Boolean
        For a = 0 To results.Length - 1
            results(a) = False
        Next

        ' Create the mock report data
        Dim ResultList As List(Of String) = New List(Of String)
        Dim ReportObject As ME_Diff.ReportOutput = New ME_Diff.ReportOutput
        ME_Diff.FilterList = New List(Of String)
        ME_Diff.FilterList.Add("backStyle")
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")
        ResultList.Add("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group3 - Object Name : Text7 - Object Type : textDrawingObjectType - Parameter : backStyle - Left Value = transparent, Right Value = solid")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfile", "RightFile")
        ME_Diff.FilterList.Clear()
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfiles", "RightFiles")

        ' add all items to the filter list to create a blank page entry
        ME_Diff.FilterList.Clear()
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")
        ME_Diff.FilterList.Add("titleBarSpecified")
        ME_Diff.FilterList.Add("displayType")
        ME_Diff.FilterList.Add("backStyle")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "LeftNoDiff", "RightNoDiff")

        ' Generate the report output
        ReportObject.GenerateHTMLReport()
        ' collect the report output into variables
        Dim TotalDifferenceString As String = ReportObject.ReportSummary.Item(8)
        Dim TotalPagesString As String = ReportObject.ReportSummary.Item(9)
        Dim PagesNoDifferenceString As String = ReportObject.ReportSummary.Item(10)
        Dim PagesWithDifferenceString As String = ReportObject.ReportSummary.Item(11)

        ' Check the content
        If InStr(TotalDifferenceString, "- 5") > 0 Then
            results(0) = True
        End If
        If InStr(TotalPagesString, "- 3") > 0 Then
            results(1) = True
        End If
        If InStr(PagesNoDifferenceString, "- 1") > 0 Then
            results(2) = True
        End If
        If InStr(PagesWithDifferenceString, "- 2") > 0 Then
            results(3) = True
        End If

        Assert.IsTrue(results(0))
        Assert.IsTrue(results(1))
        Assert.IsTrue(results(2))
        Assert.IsTrue(results(3))

    End Sub

    <TestMethod()> Public Sub CheckSummary1Output_5Pages3NoDiffs()

        Dim results(50) As Boolean
        For a = 0 To results.Length - 1
            results(a) = False
        Next

        ' Create the mock report data
        Dim ResultList As List(Of String) = New List(Of String)
        Dim ReportObject As ME_Diff.ReportOutput = New ME_Diff.ReportOutput
        ME_Diff.FilterList = New List(Of String)
        ME_Diff.FilterList.Add("backStyle")
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")
        ResultList.Add("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group3 - Object Name : Text7 - Object Type : textDrawingObjectType - Parameter : backStyle - Left Value = transparent, Right Value = solid")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfile", "RightFile")
        ME_Diff.FilterList.Clear()
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfiles", "RightFiles")

        ' add all items to the filter list to create a blank page entry
        ME_Diff.FilterList.Clear()
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")
        ME_Diff.FilterList.Add("titleBarSpecified")
        ME_Diff.FilterList.Add("displayType")
        ME_Diff.FilterList.Add("backStyle")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "LeftNoDiff", "RightNoDiff")
        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "LeftNoDiff1", "RightNoDiff2")
        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "LeftNoDiff2", "RightNoDiff3")

        ' Generate the report output
        ReportObject.GenerateHTMLReport()
        ' collect the report output into variables
        Dim ListItem1 As String = ReportObject.Summary1.Item(5)
        Dim ListItem2 As String = ReportObject.Summary1.Item(6)
        Dim ListItem3 As String = ReportObject.Summary1.Item(7)


        ' Check the content
        If InStr(ListItem1, "LeftNoDiff") > 0 Then
            results(0) = True
        End If
        ' Check the content
        If InStr(ListItem2, "LeftNoDiff1") > 0 Then
            results(1) = True
        End If ' Check the content
        If InStr(ListItem3, "LeftNoDiff2") > 0 Then
            results(2) = True
        End If

        Assert.IsTrue(results(0))
        Assert.IsTrue(results(1))
        Assert.IsTrue(results(2))

    End Sub

    <TestMethod()> Public Sub CheckSummary1Output_3Pages1NoDiffs()

        Dim results(50) As Boolean
        For a = 0 To results.Length - 1
            results(a) = False
        Next

        ' Create the mock report data
        Dim ResultList As List(Of String) = New List(Of String)
        Dim ReportObject As ME_Diff.ReportOutput = New ME_Diff.ReportOutput
        ME_Diff.FilterList = New List(Of String)
        ME_Diff.FilterList.Add("backStyle")
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")
        ResultList.Add("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group3 - Object Name : Text7 - Object Type : textDrawingObjectType - Parameter : backStyle - Left Value = transparent, Right Value = solid")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfile", "RightFile")
        ME_Diff.FilterList.Clear()
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfiles", "RightFiles")

        ' add all items to the filter list to create a blank page entry
        ME_Diff.FilterList.Clear()
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")
        ME_Diff.FilterList.Add("titleBarSpecified")
        ME_Diff.FilterList.Add("displayType")
        ME_Diff.FilterList.Add("backStyle")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "LeftNoDiff", "RightNoDiff")

        ' Generate the report output
        ReportObject.GenerateHTMLReport()
        ' collect the report output into variables
        Dim ListItem1 As String = ReportObject.Summary1.Item(5)



        ' Check the content
        If InStr(ListItem1, "LeftNoDiff") > 0 Then
            results(0) = True
        End If


        Assert.IsTrue(results(0))


    End Sub

    <TestMethod()> Public Sub CheckSummary1Output_2PagesHaveDiffs()

        Dim results(50) As Boolean
        For a = 0 To results.Length - 1
            results(a) = False
        Next

        ' Create the mock report data
        Dim ResultList As List(Of String) = New List(Of String)
        Dim ReportObject As ME_Diff.ReportOutput = New ME_Diff.ReportOutput
        ME_Diff.FilterList = New List(Of String)
        ME_Diff.FilterList.Add("backStyle")
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")
        ResultList.Add("File Name : 00000Left - Nest Level : Root - Object Name : EditComboControl5 - Object Type : activeXObjectType - Parameter : data - Left Value = 00102171600200110040002555004002552553484648100, Right Value = 0010217160020011004000255500400103484648100")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group1 - Object Name : Text2 - Object Type : textDrawingObjectType - Parameter : width - Left Value = 80, Right Value = 88")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : titleBarSpecified - Left Value = False, Right Value = True")
        ResultList.Add("File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop")
        ResultList.Add("File Name : 00000Left - Nest Level : Root/Group2/Group3 - Object Name : Text7 - Object Type : textDrawingObjectType - Parameter : backStyle - Left Value = transparent, Right Value = solid")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfile", "RightFile")
        ME_Diff.FilterList.Clear()
        ME_Diff.FilterList.Add("data")
        ME_Diff.FilterList.Add("width")

        Call ME_Diff.CreateOutputReportPage(ReportObject, ResultList, "Leftfiles", "RightFiles")


        ' Generate the report output
        ReportObject.GenerateHTMLReport()
        ' collect the report output into variables
        Dim ListItem1 As String = ReportObject.Summary1.Item(5)

        ' Check the content
        If InStr(ListItem1, "</ul>") > 0 Then
            results(0) = True
        End If


        Assert.IsTrue(results(0))


    End Sub


End Class