Imports System.IO

''' <summary>
''' A class containing the definition for 1 test case instance
''' </summary>
Public Class TestCase

    Public iTestNumber As Integer
    Public sProperty As String
    Public sLeftval As String
    Public sRightval As String

    ''' <summary>
    ''' Class initialiser, assign the input values to the class members
    ''' </summary>
    ''' <param name="i">Integer test number</param>
    ''' <param name="s">String property</param>
    ''' <param name="l">String left value</param>
    ''' <param name="r">String right value</param>
    Public Sub New(ByVal i As Integer,
                   ByVal s As String,
                   ByVal l As String,
                   ByVal r As String)

        iTestNumber = i
        sProperty = s
        sLeftval = l
        sRightval = r

    End Sub

End Class

''' <summary>
''' A helper class to extract csv files containing the test criteria to compare against the test files when performing the unit tests
''' </summary>
Public Class TestCases

    ''' <summary>
    ''' A list type containing all the test results for the current unit test
    ''' </summary>
    Public TestList As List(Of TestCase)

    ''' <summary>
    ''' Class initialiser, takes the name of the file to read the test instances from and converts those into a list of type Testcase
    ''' </summary>
    ''' <param name="TDF">Test definition file - A csv file containing the test definitions</param>
    Public Sub New(ByVal TDF As String)
        Dim strline As String
        Dim tar()
        Dim linecount As Integer = 0
        TestList = New List(Of TestCase)
        Using reader As StreamReader = New StreamReader(TDF)
            ' Read the first line and check it for headers
            strline = reader.ReadLine
            tar = Split(strline, ",")
            If Not tar(0) = "Test number" Then
                Throw New Exception("Invalid header found in TDF csv @ line " & linecount & ", File - " & TDF)
                Exit Sub
            End If
            ' Read the rest of the file
            Do
                linecount = linecount + 1
                Dim newlistentry As TestCase
                strline = reader.ReadLine
                tar = Split(strline, ",")
                If Not tar.Length = 4 Then
                    Throw New Exception("Invalid line member count in TDF csv @ line " & linecount & ", File - " & TDF)
                End If
                newlistentry = New TestCase(Val(tar(0)), tar(1), tar(2), tar(3))
                TestList.Add(newlistentry)
            Loop Until reader.EndOfStream

        End Using

    End Sub
End Class
