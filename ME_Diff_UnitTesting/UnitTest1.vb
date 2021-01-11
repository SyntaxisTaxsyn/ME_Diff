Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ME_Diff
Imports System.IO
Imports System.Windows
Imports System.Reflection

<TestClass()> Public Class DrawingTests

    <TestMethod()> Public Sub TextTest()
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

    <TestMethod()> Public Sub ImageTest()
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

    <TestMethod()> Public Sub PanelTest()
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

    <TestMethod()> Public Sub ArcTest()
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

    <TestMethod()> Public Sub EllipseTest()
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

End Class