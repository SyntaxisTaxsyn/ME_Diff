Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ME_Diff
Imports System.IO
Imports System.Windows
Imports System.Reflection

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub TestMethod1()
        Dim fcomp As New FileCompare
        Dim lstr As String = Path.Combine((Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)), "/Test Files/TextFile1.txt")
        Dim rstr As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) & "\Test Files\TextFile1.txt"
        'Dim rstr As String
        'fcomp.SetFilePaths()
        Dim sstr As String = ""
        Using ffile As StreamReader = New StreamReader(rstr)
            sstr = sstr & ffile.ReadLine
        End Using
        MsgBox(sstr)

    End Sub

End Class