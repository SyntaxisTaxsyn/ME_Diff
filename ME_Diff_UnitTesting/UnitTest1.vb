Imports System.Text
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

<TestClass()> Public Class _01PushbuttonTests

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

    <TestMethod()> Public Sub _0103MultistatePBTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "01Pushbutton"
        Dim C As String = "03Multistate"
        Dim I As String = "0103"
        Dim N As String = "MultistatePushbutton"
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

        ' state 1 and caption tests
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

        ' state 3 tests
        T = "state3"
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

    <TestMethod()> Public Sub _0104InterlockedPBTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "01Pushbutton"
        Dim C As String = "04Interlocked"
        Dim I As String = "0104"
        Dim N As String = "InterlockedPushbutton"
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

        ' state 0 and caption tests
        T = "state0"
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

        ' state 1 tests
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

    End Sub

    <TestMethod()> Public Sub _0105RampPBTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "01Pushbutton"
        Dim C As String = "05Ramp"
        Dim I As String = "0105"
        Dim N As String = "RampPushbutton"
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

        ' image and caption tests
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

<TestClass()> Public Class _02NumericandStringTests

    <TestMethod()> Public Sub _0200NumericDisplayTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "02Numeric and string"
        Dim C As String = "00NumericDisplay"
        Dim I As String = "0200"
        Dim N As String = "NumericDisplay"
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

        ' connection tests
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

    <TestMethod()> Public Sub _0201NumericInputEnableTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "02Numeric and string"
        Dim C As String = "01NumericInputEnable"
        Dim I As String = "0201"
        Dim N As String = "NumericInputEnable"
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

        ' image and caption tests
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

        ' connection tests
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

    <TestMethod()> Public Sub _0202NumericInputCursorTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "02Numeric and string"
        Dim C As String = "02NumericInputCursorPoint"
        Dim I As String = "0202"
        Dim N As String = "NumericInputCursor"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(100) As TestResult

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

        ' connection tests
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

    <TestMethod()> Public Sub _0203StringDisplayTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "02Numeric and string"
        Dim C As String = "03StringDisplay"
        Dim I As String = "0203"
        Dim N As String = "StringDisplay"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(100) As TestResult

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

    <TestMethod()> Public Sub _0204StringInputEnableTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "02Numeric and string"
        Dim C As String = "04StringInputEnable"
        Dim I As String = "0204"
        Dim N As String = "StringInputEnable"
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

        ' image and caption tests
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

        ' connection tests
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

<TestClass()> Public Class _03DisplayNavigationTests

    <TestMethod()> Public Sub _0300GotoTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "03Display Navigation"
        Dim C As String = "00Goto"
        Dim I As String = "0300"
        Dim N As String = "Goto"
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

        ' caption tests
        T = "caption"
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

        ' connection tests
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

    <TestMethod()> Public Sub _0301ReturnToTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "03Display Navigation"
        Dim C As String = "01ReturnTo"
        Dim I As String = "0301"
        Dim N As String = "ReturnTo"
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

        ' caption tests
        T = "caption"
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

    <TestMethod()> Public Sub _0302CloseButtonTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "03Display Navigation"
        Dim C As String = "02Close"
        Dim I As String = "0302"
        Dim N As String = "CloseButton"
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

        ' caption tests
        T = "caption"
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

        ' connection tests
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

    <TestMethod()> Public Sub _0303DisplayListSelectorTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "03Display Navigation"
        Dim C As String = "03DisplayListSelector"
        Dim I As String = "0303"
        Dim N As String = "DisplayListSelector"
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


        ' caption tests
        T = "caption"
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

        ' state tests
        T = "state"
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