﻿Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ME_Diff
Imports System.IO
Imports System.Windows
Imports System.Reflection

<TestClass()> Public Class _00DrawingTests

    <TestMethod()> Public Sub _0000TextTest()
        Dim fcomp As New FileCompare
        Dim tdfstr As String = GetPathToFileCSV("00Drawing", "00Text", "0000", "Text")
        Dim Lfile As String = GetPathToFileXML("00Drawing", "00Text", "0000", "Text", "L", "text")
        Dim Rfile As String = GetPathToFileXML("00Drawing", "00Text", "0000", "Text", "R", "text")
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0001ImageTest()
        Dim fcomp As New FileCompare
        Dim tdfstr As String = GetPathToFileCSV("00Drawing", "01Image", "0001", "Image")
        Dim Lfile As String = GetPathToFileXML("00Drawing", "01Image", "0001", "Image", "L", "main")
        Dim Rfile As String = GetPathToFileXML("00Drawing", "01Image", "0001", "Image", "R", "main")
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0002PanelTest()
        Dim fcomp As New FileCompare
        Dim tdfstr As String = GetPathToFileCSV("00Drawing", "02Panel", "0002", "Panel")
        Dim Lfile As String = GetPathToFileXML("00Drawing", "02Panel", "0002", "Panel", "L", "main")
        Dim Rfile As String = GetPathToFileXML("00Drawing", "02Panel", "0002", "Panel", "R", "main")
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0003ArcTest()
        Dim fcomp As New FileCompare
        Dim tdfstr As String = GetPathToFileCSV("00Drawing", "03Arc", "0003", "Arc")
        Dim Lfile As String = GetPathToFileXML("00Drawing", "03Arc", "0003", "Arc", "L", "main")
        Dim Rfile As String = GetPathToFileXML("00Drawing", "03Arc", "0003", "Arc", "R", "main")
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0004EllipseTest()
        Dim fcomp As New FileCompare
        Dim tdfstr As String = GetPathToFileCSV("00Drawing", "04Ellipse", "0004", "Ellipse")
        Dim Lfile As String = GetPathToFileXML("00Drawing", "04Ellipse", "0004", "Ellipse", "L", "main")
        Dim Rfile As String = GetPathToFileXML("00Drawing", "04Ellipse", "0004", "Ellipse", "R", "main")
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0005FreeHandTest()
        Dim fcomp As New FileCompare
        Dim tdfstr As String = GetPathToFileCSV("00Drawing", "05Freehand", "0005", "Freehand")
        Dim Lfile As String = GetPathToFileXML("00Drawing", "05Freehand", "0005", "Freehand", "L", "main")
        Dim Rfile As String = GetPathToFileXML("00Drawing", "05Freehand", "0005", "Freehand", "R", "main")
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0006LineTest()
        Dim fcomp As New FileCompare
        Dim P As String = "00Drawing"
        Dim C As String = "06Line"
        Dim I As String = "0006"
        Dim N As String = "Line"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N)
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0007PolygonTest()
        Dim fcomp As New FileCompare
        Dim P As String = "00Drawing"
        Dim C As String = "07Polygon"
        Dim I As String = "0007"
        Dim N As String = "Polygon"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N)
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0008PolylineTest()
        Dim fcomp As New FileCompare
        Dim P As String = "00Drawing"
        Dim C As String = "08Polyline"
        Dim I As String = "0008"
        Dim N As String = "Polyline"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N)
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0009RectangleTest()
        Dim fcomp As New FileCompare
        Dim P As String = "00Drawing"
        Dim C As String = "09Rectangle"
        Dim I As String = "0009"
        Dim N As String = "Rectangle"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N)
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0010RoundedRectangleTest()
        Dim fcomp As New FileCompare
        Dim P As String = "00Drawing"
        Dim C As String = "10RoundedRectangle"
        Dim I As String = "0010"
        Dim N As String = "RoundedRectangle"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N)
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0011WedgeTest()
        Dim fcomp As New FileCompare
        Dim P As String = "00Drawing"
        Dim C As String = "11Wedge"
        Dim I As String = "0011"
        Dim N As String = "Wedge"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N)
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

End Class

<TestClass()> Public Class _00PushbuttonTests

    <TestMethod()> Public Sub _0100MomentaryPBTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "01Pushbutton"
        Dim C As String = "00Momentary"
        Dim I As String = "0100"
        Dim N As String = "MomentaryPushbutton"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' state0 and caption tests
        T = "state1"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' image and connection tests
        T = "image"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' state 2 tests
        T = "state2"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' connections tests
        T = "connections"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next


    End Sub

    <TestMethod()> Public Sub _0101MaintainedPBTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "01Pushbutton"
        Dim C As String = "01Maintained"
        Dim I As String = "0101"
        Dim N As String = "MaintainedPushbutton"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' state1 and caption tests
        T = "state1"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' image tests
        T = "image"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' state 2 tests
        T = "state2"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' connections tests
        T = "connections"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

    <TestMethod()> Public Sub _0102LatchedPBTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "01Pushbutton"
        Dim C As String = "02Latched"
        Dim I As String = "0102"
        Dim N As String = "LatchedPushbutton"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(50) As TestResult

        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' state1 and caption tests
        T = "state1"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' image tests
        T = "image"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' state 2 tests
        T = "state2"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

        ' connections tests
        T = "connections"
        tdfstr = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Lfile = GetPathToFileXML(P, C, I, N, "L", T)
        Rfile = GetPathToFileXML(P, C, I, N, "R", T)
        fcomp = Nothing
        fcomp = New FileCompare
        ThisTestCase = New TestCases(tdfstr)
        strlst = Nothing
        strlst = New List(Of String)
        For a = 0 To Result.Length - 1
            Result(a) = TestResult.Fail
        Next

        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        For a = 0 To ThisTestCase.TestList.Count - 1
            Result(a) = CheckTestCriteria(ThisTestCase.TestList.Item(a), strlst)
            Assert.IsTrue(Result(a), "Failed at index " & a &
                          ",Property - " & ThisTestCase.TestList.Item(a).sProperty &
                          ",Left - " & ThisTestCase.TestList.Item(a).sLeftval &
                          ",Right - " & ThisTestCase.TestList.Item(a).sRightval)
        Next

    End Sub

End Class