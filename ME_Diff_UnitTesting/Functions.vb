Imports System.IO
Imports System.Reflection
Module Functions
    ''' <summary>
    ''' Calculate the path to the test instance file
    ''' </summary>
    ''' <param name="P">Parent folder</param>
    ''' <param name="C">Child folder</param>
    ''' <param name="I">Test instance integer as string</param>
    ''' <param name="N">Test name</param>
    ''' <param name="S">Test salutation (right R or left L)</param>
    ''' <param name="T">Test type</param>
    ''' <returns>A string containing the full path to the desired test instance file</returns>
    Public Function GetPathToFileXML(ByVal P As String,
                                     ByVal C As String,
                                     ByVal I As String,
                                     ByVal N As String,
                                     ByVal S As String,
                                     ByVal T As String)

        ' Build the path string
        Dim wrkstr As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) & "\Test Pages\"
        wrkstr = wrkstr & P & "\"
        wrkstr = wrkstr & C & "\"
        wrkstr = wrkstr & I & N & S & "_" & T & ".xml"

        GetPathToFileXML = wrkstr

    End Function

    ''' <summary>
    ''' Calculate the path to the test instance csv definition file
    ''' </summary>
    ''' <param name="P">Parent folder</param>
    ''' <param name="C">Child folder</param>
    ''' <param name="I">Test instance integer as string</param>
    ''' <param name="N">Test name</param>
    ''' <returns>A string containing the full path to the desired test instance definition file</returns>
    Public Function GetPathToFileCSV(ByVal P As String,
                                     ByVal C As String,
                                     ByVal I As String,
                                     ByVal N As String)

        ' Build the path string
        Dim wrkstr As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) & "\Test Pages\"
        wrkstr = wrkstr & P & "\"
        wrkstr = wrkstr & C & "\"
        wrkstr = wrkstr & I & N & ".csv"

        GetPathToFileCSV = wrkstr

    End Function

    ''' <summary>
    ''' Check the string list for passing of all test criteria before outputing the result to the calling method
    ''' </summary>
    ''' <param name="TCase">Test case object</param>
    ''' <param name="TList">Test list object</param>
    ''' <returns>Test result as an object of type TestResult</returns>
    Public Function CheckTestCriteria(ByRef TCase As TestCase,
                                      ByRef TList As List(Of String)) As TestResult

        CheckTestCriteria = TestResult.Fail
        Dim bResult As Boolean = False

        For Each itm In TList
            If itm.Contains("Test" & TCase.iTestNumber) Then
                If itm.Contains("Parameter : " & TCase.sProperty) Then
                    If itm.Contains("Left Value = " & TCase.sLeftval) Then
                        If itm.Contains("Right Value = " & TCase.sRightval) Then
                            bResult = True
                        End If
                    End If
                End If
            End If
            If bResult = True Then
                CheckTestCriteria = TestResult.Pass
                Exit Function
            End If
        Next

        If bResult = False Then
            CheckTestCriteria = TestResult.Fail
        End If

    End Function

    Public Enum TestResult
        Fail
        Pass
    End Enum
End Module
