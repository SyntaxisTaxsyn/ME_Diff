﻿Imports System.Text
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
        Dim Result As TestResult = TestResult.Fail
        fcomp.SetFilePaths(Lfile, Rfile)
        strlst = fcomp.CompareFiles(FileCompare.UnitTestType.FixedFile)
        Result = CheckTestCriteria(ThisTestCase.TestList.Item(0), strlst)
        Assert.IsTrue(Result)


    End Sub

End Class