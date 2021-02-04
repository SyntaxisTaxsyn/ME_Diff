Public Class ReportOutput
    Public PageList As List(Of ReportOutputPage)
    Public Sub New()
        PageList = New List(Of ReportOutputPage)
    End Sub
    Public Sub AddPage(ByRef LeftFileStr As String, ByRef RightFileStr As String)
        Dim NewReportPage As New ReportOutputPage(LeftFileStr, RightFileStr)
        PageList.Add(NewReportPage)
    End Sub
End Class

Public Class ReportOutputPage
    Public LeftFile As String
    Public RightFile As String
    Public OutputList As List(Of ReportMember)
    Public Sub New(ByRef Lfile As String, ByRef RFile As String)
        OutputList = New List(Of ReportMember)
        LeftFile = Lfile
        RightFile = RFile
    End Sub
    Public Sub AddListItem(ByRef reportstr As String)
        Dim ListMember As New ReportMember(reportstr)
        OutputList.Add(ListMember)
    End Sub
End Class

Public Class ReportMember
    Public FileName As String
    Public NestLevel As String
    Public ObjectName As String
    Public ObjectType As String
    Public Parameter As String
    Public LeftValue As String
    Public RightValue As String
    ''' <summary>
    '''  Initialise a new member of this class by converting the outputlinn string into the constituent class parameters
    '''  this effectiveely prepares the output data for reformatting/redisplaying in a number of ways
    ''' </summary>
    ''' <param name="outputline"></param>
    Public Sub New(ByRef outputline As String)
        '"File Name : 00000Left - Nest Level :  - Object Name : Display - Object Type : gfx - Parameter : displayType - Left Value = replace, Right Value = onTop"
        Dim tarr() As String
        Dim SearchParams As String() = {
            OutputMembers.FileName,
            OutputMembers.NestLevel,
            OutputMembers.ObjectName,
            OutputMembers.ObjectType,
            OutputMembers.Parameter,
            OutputMembers.LeftValue,
            OutputMembers.RightValue}

        tarr = Split(outputline, " - ") ' whitespace added to the - split character to deliberately consume the additonal whitespacing in the input message
        For Each item In SearchParams
            For Each member In tarr
                If member.Contains(item) Then
                    If item Like "*" & " Value" & "*" Then
                        ' handle the special case here as left/right values are within this array element also
                        Dim intr()
                        intr = Split(member, ", ")
                        LeftValue = Replace(intr(0), OutputMembers.LeftValue & " = ", "")
                        RightValue = Replace(intr(1), OutputMembers.RightValue & " = ", "")
                    Else
                        Dim Proceed As Boolean = True
                        If item = OutputMembers.LeftValue Then Proceed = False
                        If item = OutputMembers.RightValue Then Proceed = False
                        If Proceed Then
                            Dim intr()
                            intr = Split(member, " : ")
                            Select Case item
                                Case OutputMembers.FileName
                                    FileName = intr(1)
                                Case OutputMembers.NestLevel
                                    NestLevel = intr(1)
                                Case OutputMembers.ObjectName
                                    ObjectName = intr(1)
                                Case OutputMembers.ObjectType
                                    ObjectType = intr(1)
                                Case OutputMembers.Parameter
                                    Parameter = intr(1)
                            End Select
                        End If
                    End If
                End If
            Next
        Next

    End Sub
End Class

Public Class OutputMembers
    Public Const FileName As String = "File Name"
    Public Const NestLevel As String = "Nest Level"
    Public Const ObjectName As String = "Object Name"
    Public Const ObjectType As String = "Object Type"
    Public Const Parameter As String = "Parameter"
    Public Const LeftValue As String = "Left Value"
    Public Const RightValue As String = "Right Value"
End Class
