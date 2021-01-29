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

<TestClass()> Public Class _04IndicatorTests

    <TestMethod()> Public Sub _0400MultistateIndicatorTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "04Indicator"
        Dim C As String = "00Multistate"
        Dim I As String = "0400"
        Dim N As String = "MultistateIndicator"
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

    <TestMethod()> Public Sub _0401SymbolIndicatorTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "04Indicator"
        Dim C As String = "01Symbol"
        Dim I As String = "0401"
        Dim N As String = "SymbolIndicator"
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

    <TestMethod()> Public Sub _0402ListIndicatorTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "04Indicator"
        Dim C As String = "02List"
        Dim I As String = "0402"
        Dim N As String = "ListIndicator"
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

<TestClass()> Public Class _05GaugeandGraphTests

    <TestMethod()> Public Sub _0500BarGraphTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "05Gauge and graph"
        Dim C As String = "00BarGraph"
        Dim I As String = "0500"
        Dim N As String = "BarGraph"
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


        ' threshold tests
        T = "threshold"
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

    <TestMethod()> Public Sub _0501GaugeTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "05Gauge and graph"
        Dim C As String = "01Gauge"
        Dim I As String = "0501"
        Dim N As String = "Gauge"
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


        ' threshold tests
        T = "threshold"
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

    <TestMethod()> Public Sub _0502ScaleTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "05Gauge and graph"
        Dim C As String = "02Scale"
        Dim I As String = "0502"
        Dim N As String = "Scale"
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

    End Sub

End Class


<TestClass()> Public Class _06TrendingTests

    <TestMethod()> Public Sub _0600TrendPausePenTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "06Trending"
        Dim C As String = "00TrendPausePen"
        Dim I As String = "0600"
        Dim N As String = "TrendPausePen"
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

    <TestMethod()> Public Sub _0601TrendNextPenTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "06Trending"
        Dim C As String = "01TrendNextPen"
        Dim I As String = "0601"
        Dim N As String = "TrendNextPen"
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

    <TestMethod()> Public Sub _0602TrendTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "06Trending"
        Dim C As String = "02Trend"
        Dim I As String = "0602"
        Dim N As String = "Trend"
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


        ' Connections tests
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

        ' data tests
        T = "data"
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

<TestClass()> Public Class _07RecipePlusTests

    <TestMethod()> Public Sub _0700RecipePlusButtonTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "07Recipe plus"
        Dim C As String = "00RecipePlusButton"
        Dim I As String = "0700"
        Dim N As String = "RecipePlusButton"
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

    <TestMethod()> Public Sub _0701RecipePlusSelectorTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "07Recipe plus"
        Dim C As String = "01RecipePlusSelector"
        Dim I As String = "0701"
        Dim N As String = "RecipePlusSelector"
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

    End Sub

    <TestMethod()> Public Sub _0702RecipePlusTableTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "07Recipe plus"
        Dim C As String = "02RecipePlusTable"
        Dim I As String = "0702"
        Dim N As String = "RecipePlusTable"
        Dim T As String = "main"
        Dim tdfstr As String = GetPathToFileCSV(P, C, I, N & T) ' Added "& T" to N param as this test has multiple subtypes
        Dim Lfile As String = GetPathToFileXML(P, C, I, N, "L", T)
        Dim Rfile As String = GetPathToFileXML(P, C, I, N, "R", T)
        Dim ThisTestCase As New TestCases(tdfstr)
        Dim strlst As List(Of String)
        Dim Result(60) As TestResult

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

<TestClass()> Public Class _08KeyTests

    <TestMethod()> Public Sub _0800BackSpaceKeyTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "08Key"
        Dim C As String = "00BackSpaceKey"
        Dim I As String = "0800"
        Dim N As String = "BackSpaceKey"
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

    <TestMethod()> Public Sub _0801EndKeyTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "08Key"
        Dim C As String = "01EndKey"
        Dim I As String = "0801"
        Dim N As String = "EndKey"
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

    <TestMethod()> Public Sub _0802EnterKeyTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "08Key"
        Dim C As String = "02EnterKey"
        Dim I As String = "0802"
        Dim N As String = "EnterKey"
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

    <TestMethod()> Public Sub _0803HomeKeyTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "08Key"
        Dim C As String = "03HomeKey"
        Dim I As String = "0803"
        Dim N As String = "HomeKey"
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

    <TestMethod()> Public Sub _0804MoveDownKeyTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "08Key"
        Dim C As String = "04MoveDownKey"
        Dim I As String = "0804"
        Dim N As String = "MoveDownKey"
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

    <TestMethod()> Public Sub _0805MoveLeftKeyTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "08Key"
        Dim C As String = "05MoveLeftKey"
        Dim I As String = "0805"
        Dim N As String = "MoveLeftKey"
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

    <TestMethod()> Public Sub _0806MoveRightKeyTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "08Key"
        Dim C As String = "06MoveRightKey"
        Dim I As String = "0806"
        Dim N As String = "MoveRightKey"
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

    <TestMethod()> Public Sub _0807MoveUpKeyTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "08Key"
        Dim C As String = "07MoveUpKey"
        Dim I As String = "0807"
        Dim N As String = "MoveUpKey"
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

    <TestMethod()> Public Sub _0808PageDownKeyTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "08Key"
        Dim C As String = "08PageDownKey"
        Dim I As String = "0808"
        Dim N As String = "PageDownKey"
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

    <TestMethod()> Public Sub _0809PageUpKeyTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "08Key"
        Dim C As String = "09PageUpKey"
        Dim I As String = "0809"
        Dim N As String = "PageUpKey"
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

End Class

<TestClass()> Public Class _09UserManagementTests

    <TestMethod()> Public Sub _0900AddUserGroupTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "09User management"
        Dim C As String = "00AddUserGroup"
        Dim I As String = "0900"
        Dim N As String = "AddUserGroup"
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

    <TestMethod()> Public Sub _0901DeleteUserGroupTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "09User management"
        Dim C As String = "01DeleteUserGroup"
        Dim I As String = "0901"
        Dim N As String = "DeleteUserGroup"
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

    <TestMethod()> Public Sub _0902ModifyGroupMembershipTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "09User management"
        Dim C As String = "02ModifyGroupMembership"
        Dim I As String = "0902"
        Dim N As String = "ModifyGroupMembership"
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

    <TestMethod()> Public Sub _0903UnlockUserTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "09User management"
        Dim C As String = "03UnlockUser"
        Dim I As String = "0903"
        Dim N As String = "UnlockUser"
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

    <TestMethod()> Public Sub _0904EnableUserTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "09User management"
        Dim C As String = "04EnableUser"
        Dim I As String = "0904"
        Dim N As String = "EnableUser"
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

    <TestMethod()> Public Sub _0905DisableUserTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "09User management"
        Dim C As String = "05DisableUser"
        Dim I As String = "0905"
        Dim N As String = "DisableUser"
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

    <TestMethod()> Public Sub _0906LoginTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "09User management"
        Dim C As String = "06Login"
        Dim I As String = "0906"
        Dim N As String = "Login"
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

    <TestMethod()> Public Sub _0907LogoutTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "09User management"
        Dim C As String = "07Logout"
        Dim I As String = "0907"
        Dim N As String = "Logout"
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

    <TestMethod()> Public Sub _0908PasswordTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "09User management"
        Dim C As String = "08Password"
        Dim I As String = "0908"
        Dim N As String = "Password"
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

    <TestMethod()> Public Sub _0909ChangeUserPropertiesTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "09User management"
        Dim C As String = "09ChangeUserProperties"
        Dim I As String = "0909"
        Dim N As String = "ChangeUserProperties"
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

End Class


<TestClass()> Public Class _10AdvancedTests

    <TestMethod()> Public Sub _1000ControlListSelectorTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "00ControlListSelector"
        Dim I As String = "1000"
        Dim N As String = "ControlListSelector"
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

    <TestMethod()> Public Sub _1001PilotedControlListSelectorTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "01PilotedControlListSelector"
        Dim I As String = "1001"
        Dim N As String = "PilotedControlListSelector"
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

    <TestMethod()> Public Sub _1002DisplayPrintTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "02DisplayPrint"
        Dim I As String = "1002"
        Dim N As String = "DisplayPrint"
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

    End Sub

    <TestMethod()> Public Sub _1003LanguageSwitchButtonTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "03LanguageSwitchButton"
        Dim I As String = "1003"
        Dim N As String = "LanguageSwitchButton"
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

    End Sub

    <TestMethod()> Public Sub _1004LocalMessageDisplayTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "04LocalMessageDisplay"
        Dim I As String = "1004"
        Dim N As String = "LocalMessageDisplay"
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

    <TestMethod()> Public Sub _1005MacroTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "05Macro"
        Dim I As String = "1005"
        Dim N As String = "Macro"
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

    End Sub

    <TestMethod()> Public Sub _1006ShutdownTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "06Shutdown"
        Dim I As String = "1006"
        Dim N As String = "Shutdown"
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

    End Sub

    <TestMethod()> Public Sub _1007GotoConfigureModeTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "07GotoConfigureMode"
        Dim I As String = "1007"
        Dim N As String = "GotoConfigureMode"
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

    End Sub

    <TestMethod()> Public Sub _1008TimeAndDateTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "08TimeAndDate"
        Dim I As String = "1008"
        Dim N As String = "TimeAndDate"
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
    End Sub

    <TestMethod()> Public Sub _100900AcknowledgeTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0900Acknowledge"
        Dim I As String = "0900"
        Dim N As String = "Acknowledge"
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


    End Sub

    <TestMethod()> Public Sub _100901AcknowledgeAllTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0901AcknowledgeAll"
        Dim I As String = "0901"
        Dim N As String = "AcknowledgeAll"
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


    End Sub

    <TestMethod()> Public Sub _100902AlarmStatusModeTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0902AlarmStatusMode"
        Dim I As String = "0902"
        Dim N As String = "AlarmStatusMode"
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


    End Sub


    <TestMethod()> Public Sub _100903ClearAlarmBannerTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0903ClearAlarmBanner"
        Dim I As String = "0903"
        Dim N As String = "ClearAlarmBanner"
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


    End Sub

    <TestMethod()> Public Sub _100904ClearAlarmHistoryTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0904ClearAlarmHistory"
        Dim I As String = "0904"
        Dim N As String = "ClearAlarmHistory"
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


    End Sub

    <TestMethod()> Public Sub _100905PrintAlarmHistoryTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0905PrintAlarmHistory"
        Dim I As String = "0905"
        Dim N As String = "PrintAlarmHistory"
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


    End Sub

    <TestMethod()> Public Sub _100906PrintAlarmStatusTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0906PrintAlarmStatus"
        Dim I As String = "0906"
        Dim N As String = "PrintAlarmStatus"
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


    End Sub

    <TestMethod()> Public Sub _100907ResetAlarmStatusTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0907ResetAlarmStatus"
        Dim I As String = "0907"
        Dim N As String = "ResetAlarmStatus"
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


    End Sub

    <TestMethod()> Public Sub _100908SilenceTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0908Silence"
        Dim I As String = "0908"
        Dim N As String = "Silence"
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


    End Sub

    <TestMethod()> Public Sub _100909SortAlarmsTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0909SortAlarms"
        Dim I As String = "0909"
        Dim N As String = "SortAlarms"
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


    End Sub

    <TestMethod()> Public Sub _100910AlarmListTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0910AlarmList"
        Dim I As String = "0910"
        Dim N As String = "AlarmList"
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

    <TestMethod()> Public Sub _100911AlarmBannerTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0911AlarmBanner"
        Dim I As String = "0911"
        Dim N As String = "AlarmBanner"
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

    <TestMethod()> Public Sub _100912AlarmStatusListTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "09Alarm/0912AlarmStatusList"
        Dim I As String = "0912"
        Dim N As String = "AlarmStatusList"
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

    <TestMethod()> Public Sub _101000ClearTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "10Diagnostics/1000Clear"
        Dim I As String = "1000"
        Dim N As String = "Clear"
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

    <TestMethod()> Public Sub _101001ClearAllTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "10Diagnostics/1001ClearAll"
        Dim I As String = "1001"
        Dim N As String = "ClearAll"
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

    <TestMethod()> Public Sub _101002DiagnosticsListTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "10Diagnostics/1002DiagnosticsList"
        Dim I As String = "1002"
        Dim N As String = "DiagnosticsList"
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

    <TestMethod()> Public Sub _101100ClearAuditTrailTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "11Audit/1100ClearAuditTrail"
        Dim I As String = "1100"
        Dim N As String = "ClearAuditTrail"
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

    <TestMethod()> Public Sub _101101AuditTrailListTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "11Audit/1101AuditTrailList"
        Dim I As String = "1101"
        Dim N As String = "AuditTrailList"
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

    <TestMethod()> Public Sub _101102AuditTrailDetailTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "11Audit/1102AuditTrailDetail"
        Dim I As String = "1102"
        Dim N As String = "AuditTrailDetail"
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

    <TestMethod()> Public Sub _101200InfoAcknowledgeTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "12Information/1200Acknowledge"
        Dim I As String = "1200"
        Dim N As String = "Acknowledge"
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

    <TestMethod()> Public Sub _101201InfoMessageDisplayTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "10Advanced"
        Dim C As String = "12Information/1201MessageDisplay"
        Dim I As String = "1201"
        Dim N As String = "MessageDisplay"
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
End Class

<TestClass()> Public Class _11ActiveXTests

    <TestMethod()> Public Sub _1100MEProgramLauncherTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "11Active X Control"
        Dim C As String = "00MEProgramLauncher"
        Dim I As String = "1100"
        Dim N As String = "MEProgramLauncher"
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

        ' data tests
        T = "data"
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

    <TestMethod()> Public Sub _1101AdobePDFViewerTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "11Active X Control"
        Dim C As String = "01AdobePDFViewer"
        Dim I As String = "1101"
        Dim N As String = "AdobePDFViewer"
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

        ' data tests
        T = "data"
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

    <TestMethod()> Public Sub _1102CogDisplayTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "11Active X Control"
        Dim C As String = "02CogDisplay"
        Dim I As String = "1102"
        Dim N As String = "CogDisplay"
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

        ' data tests
        T = "data"
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

    <TestMethod()> Public Sub _1103FTViewDiagnosticsTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "11Active X Control"
        Dim C As String = "03FTViewDiagnosticsViewer"
        Dim I As String = "1103"
        Dim N As String = "FTViewDiagnostics"
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

        ' data tests
        T = "data"
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

    <TestMethod()> Public Sub _1104DBGridControlTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "11Active X Control"
        Dim C As String = "04DBGridControl"
        Dim I As String = "1104"
        Dim N As String = "DBGridControl"
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

End Class

<TestClass()> Public Class _12SymbolFactoryTests

    <TestMethod()> Public Sub _1200MEPolygon1Test()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "12Symbol factory"
        Dim C As String = "00Polygon1"
        Dim I As String = "1200"
        Dim N As String = "Polygon1"
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

    <TestMethod()> Public Sub _1201MEPolygon2Test()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "12Symbol factory"
        Dim C As String = "01Polygon2"
        Dim I As String = "1201"
        Dim N As String = "Polygon2"
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

    <TestMethod()> Public Sub _1202MEPolygon3Test()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "12Symbol factory"
        Dim C As String = "02Polygon3"
        Dim I As String = "1202"
        Dim N As String = "Polygon3"
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

    <TestMethod()> Public Sub _1203MEPolygon4Test()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "12Symbol factory"
        Dim C As String = "03Polygon4"
        Dim I As String = "1203"
        Dim N As String = "Polygon4"
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

    <TestMethod()> Public Sub _1204MEPolygon5Test()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "12Symbol factory"
        Dim C As String = "04Polygon5"
        Dim I As String = "1204"
        Dim N As String = "Polygon5"
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

    <TestMethod()> Public Sub _1205MEPolygon6Test()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "12Symbol factory"
        Dim C As String = "05Polygon6"
        Dim I As String = "1205"
        Dim N As String = "Polygon6"
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

    <TestMethod()> Public Sub _1206MEPolygon7Test()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "12Symbol factory"
        Dim C As String = "06Polygon7"
        Dim I As String = "1206"
        Dim N As String = "Polygon7"
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

    <TestMethod()> Public Sub _1207MEPolygon8Test()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "12Symbol factory"
        Dim C As String = "07Polygon8"
        Dim I As String = "1207"
        Dim N As String = "Polygon8"
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

End Class

<TestClass()> Public Class _13AnimationTests

    <TestMethod()> Public Sub _1300MEAnimationVisibilityTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "13Animations"
        Dim C As String = "00Visibility"
        Dim I As String = "1300"
        Dim N As String = "AnimationVisibility"
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

        ' animate tests
        T = "animate"
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

    <TestMethod()> Public Sub _1301MEAnimateColorTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "13Animations"
        Dim C As String = "01Color"
        Dim I As String = "1301"
        Dim N As String = "AnimateColor"
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

        ' animate tests
        T = "animate"
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

        ' animatecolor tests
        T = "animatecolor"
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

    <TestMethod()> Public Sub _1302MEAnimateFillTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "13Animations"
        Dim C As String = "02Fill"
        Dim I As String = "1302"
        Dim N As String = "AnimateFill"
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

        ' animate tests
        T = "animate"
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

        ' animateexpressionrange tests
        T = "animateexpressionrange"
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

        ' animateconstant tests
        T = "animateconstant"
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

    <TestMethod()> Public Sub _1303MEAnimateHorizontalPositionTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "13Animations"
        Dim C As String = "03HorizontalPosition"
        Dim I As String = "1303"
        Dim N As String = "AnimateHorizontalPosition"
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

        ' animate tests
        T = "animate"
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

        ' animateexpressionrange tests
        T = "animatereadfromtag"
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

        ' animateconstant tests
        T = "animateconstant"
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

    <TestMethod()> Public Sub _1304MEAnimateVerticalPositionTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "13Animations"
        Dim C As String = "04VerticalPosition"
        Dim I As String = "1304"
        Dim N As String = "AnimateVerticalPosition"
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

        ' animate tests
        T = "animate"
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

        ' animateexpressionrange tests
        T = "animatereadfromtag"
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

        ' animateconstant tests
        T = "animateconstant"
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

    <TestMethod()> Public Sub _1305MEAnimateWidthTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "13Animations"
        Dim C As String = "05Width"
        Dim I As String = "1305"
        Dim N As String = "AnimateWidth"
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

        ' animate tests
        T = "animate"
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

        ' animateexpressionrange tests
        T = "animatereadfromtag"
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

        ' animateconstant tests
        T = "animateconstant"
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

    <TestMethod()> Public Sub _1306MEAnimateHeightTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "13Animations"
        Dim C As String = "06Height"
        Dim I As String = "1306"
        Dim N As String = "AnimateHeight"
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

        ' animate tests
        T = "animate"
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

        ' animateexpressionrange tests
        T = "animatereadfromtag"
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

        ' animateconstant tests
        T = "animateconstant"
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

    <TestMethod()> Public Sub _1307MEAnimateRotationTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "13Animations"
        Dim C As String = "07Rotation"
        Dim I As String = "1307"
        Dim N As String = "AnimateRotation"
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

        ' animate tests
        T = "animate"
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

        ' animateexpressionrange tests
        T = "animatereadfromtag"
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

        ' animateconstant tests
        T = "animateconstant"
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

    <TestMethod()> Public Sub _1308MEAnimateHyperLinkTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "13Animations"
        Dim C As String = "08Hyperlink"
        Dim I As String = "1308"
        Dim N As String = "AnimateHyperlink"
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

        ' animate tests
        T = "animate"
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

    <TestMethod()> Public Sub _1309MEAnimateHorizontalSliderTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "13Animations"
        Dim C As String = "09HorizontalSlider"
        Dim I As String = "1309"
        Dim N As String = "AnimateHorizontalSlider"
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

        ' animate tests
        T = "animate"
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

        ' animateexpressionrange tests
        T = "animatereadfromtag"
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

        ' animateconstant tests
        T = "animateconstant"
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

    <TestMethod()> Public Sub _1310MEAnimateVerticalSliderTest()

        ' main tests
        Dim fcomp As New FileCompare
        Dim P As String = "13Animations"
        Dim C As String = "10VerticalSlider"
        Dim I As String = "1310"
        Dim N As String = "AnimateVerticalSlider"
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

        ' animate tests
        T = "animate"
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

        ' animateexpressionrange tests
        T = "animatereadfromtag"
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

        ' animateconstant tests
        T = "animateconstant"
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