Imports System.IO
Imports System.Reflection
Public Class ReportOutput
    Public PageList As List(Of ReportOutputPage)
    Public Header As List(Of String)
    Public ReportSummary As List(Of String)
    Public Summary1 As List(Of String)
    Public Summary2 As List(Of String)
    Public PageReport As List(Of String)
    Public PageReports As List(Of List(Of String))
    Public Footer As List(Of String)
    Public HTMLReport As List(Of String)
    Public Sub New()
        PageList = New List(Of ReportOutputPage)
    End Sub
    Public Sub AddPage(ByRef LeftFileStr As String, ByRef RightFileStr As String)
        Dim NewReportPage As New ReportOutputPage(LeftFileStr, RightFileStr)
        PageList.Add(NewReportPage)
    End Sub
    Public Sub AddPage(ByRef PageObject As ReportOutputPage)
        PageList.Add(PageObject)
    End Sub

    ''' <summary>
    ''' Once the page list has been populated with all desired entries this method can be called to generate an
    ''' HTML report using the stored template files from the project
    ''' </summary>
    Public Sub GenerateHTMLReport()
        ' Get paths to files
        Dim HeaderPath As String
        Dim ReportSummaryPath As String
        Dim Summary1Path As String
        Dim Summary2Path As String
        Dim PageReportPath As String
        Dim FooterPath As String

        HeaderPath = GetWorkingDirectoryPath("ReportHeader.html")
        ReportSummaryPath = GetWorkingDirectoryPath("ReportSummary.html")
        Summary1Path = GetWorkingDirectoryPath("PagesSummary1.html")
        Summary2Path = GetWorkingDirectoryPath("PagesSummary2.html")
        PageReportPath = GetWorkingDirectoryPath("PageReport.html")
        FooterPath = GetWorkingDirectoryPath("ReportFooter.html")

        ' Get template file content into working memory
        Header = ReadFileToList(HeaderPath)
        ReportSummary = ReadFileToList(ReportSummaryPath)
        Summary1 = ReadFileToList(Summary1Path)
        Summary2 = ReadFileToList(Summary2Path)
        PageReport = ReadFileToList(PageReportPath)
        PageReports = New List(Of List(Of String))
        Footer = ReadFileToList(FooterPath)

        ' Get report content
        ReportSummary = GetReportSummary(ReportSummary)
        Summary1 = GetSummary1(Summary1)
        Summary2 = GetSummary2(Summary2)
        PageReports = GetPageReports(PageReports, PageReport)

        ' Build report and display
        Call BuildHTMLReport()
        Call WriteHTMLReportOutToFile()
        Call DisplayHTMLReport


    End Sub

    Public Sub DisplayHTMLReport()

        Dim ReportPath As String = Path.GetDirectoryName(STR_LeftPathFile) & "\Report.html"
        Dim RunResult As Boolean
        RunResult = Run(DefaultWebBrowser, ReportPath)

    End Sub
    Public Sub WriteHTMLReportOutToFile()
        Using output As StreamWriter = New StreamWriter(Path.GetDirectoryName(STR_LeftPathFile) & "\Report.html", False)
            For Each line In HTMLReport
                output.WriteLine(line)
            Next
        End Using
    End Sub

    Public Sub BuildHTMLReport()
        HTMLReport = New List(Of String)

        For Each line In Header
            HTMLReport.Add(line)
        Next

        For Each line In ReportSummary
            HTMLReport.Add(line)
        Next

        For Each line In Summary1
            HTMLReport.Add(line)
        Next

        For Each line In Summary2
            HTMLReport.Add(line)
        Next

        For Each page In PageReports
            For Each item In page
                HTMLReport.Add(item)
            Next
        Next

        For Each line In Footer
            HTMLReport.Add(line)
        Next

    End Sub
    ''' <summary>
    ''' Takes the list of lists of strings pagereports object and adds 1 object containing a list of strings populated with the report details
    ''' For each page found in the pagelist that had differences on it
    ''' </summary>
    ''' <param name="PageReports">Object containing a list of page lists</param>
    ''' <param name="PageReport">Template page list for generating individual page lists</param>
    ''' <returns>An object containing a list of page list reports</returns>
    Public Function GetPageReports(ByRef PageReports As List(Of List(Of String)),
                                   ByRef PageReport As List(Of String)) As List(Of List(Of String))

        Dim PageReportList As List(Of List(Of String)) = New List(Of List(Of String))

        For Each Page In PageList
            If Page.OutputList.Count > 0 Then
                Dim NewPageReport As List(Of String)
                NewPageReport = GetPageReport(Page, PageReport)
                PageReportList.Add(NewPageReport)
            End If
        Next

        Return PageReportList

    End Function

    ''' <summary>
    ''' Takes a page report from the reportoutput object and a page report template and produces a list of strings containing
    ''' the correctly formatted page report with all the data contained in the correct places
    ''' </summary>
    ''' <param name="Page">Page object containing the page details to add to the report template</param>
    ''' <param name="PageReport">Template page list for generating individual page lists</param>
    ''' <returns>Formatted list of strings containing the report data for the specific page</returns>
    Public Function GetPageReport(ByRef Page As ReportOutputPage, ByRef PageReport As List(Of String)) As List(Of String)

        ' Order of events in here 

        ' Create a new list object for the table content
        ' Add the table content from the page report
        ' Combine the 2 lists into an output report and return it

        ' Get a copy of the report template into a new list object
        Dim ReportCopy As List(Of String) = New List(Of String)
        For Each item In PageReport
            ReportCopy.Add(item)
        Next

        ' Declare variables to use for page content substitution
        Dim PageLinkID As String
        Dim PageName As String
        Dim LeftFileName As String
        Dim RightFileName As String
        Dim DifferenceQuantity As Integer

        ' Assign values to page content replacements
        PageLinkID = Replace(Page.LeftFile, " ", "_")
        PageName = Page.LeftFile
        LeftFileName = Page.LeftFile
        RightFileName = Page.RightFile
        DifferenceQuantity = Page.OutputList.Count

        ' Parse report list copy and replace values with variable content
        ReportCopy = SearchListForSubstituteString(ReportCopy, "{PageLinkID}", PageLinkID)
        ReportCopy = SearchListForSubstituteString(ReportCopy, "{PageName}", PageName)
        ReportCopy = SearchListForSubstituteString(ReportCopy, "{LeftFileName}", LeftFileName)
        ReportCopy = SearchListForSubstituteString(ReportCopy, "{RightFileName}", RightFileName)
        ReportCopy = SearchListForSubstituteString(ReportCopy, "{DifferenceQuantity}", DifferenceQuantity)

        Dim ReportTableList As List(Of String) = New List(Of String)
        ReportTableList = GetTableReportList(Page.OutputList)

        Dim NewList As New List(Of String)
        For Each item In ReportCopy
            If Not InStr(item, "{TableContent}") > 0 Then
                NewList.Add(item)
            Else
                For Each line In ReportTableList
                    NewList.Add(line)
                Next
            End If
        Next

        Return NewList

    End Function

    Public Function GetTableReportList(ByRef ReportOutputList As List(Of ReportMember)) As List(Of String)

        Dim NewList As New List(Of String)
        Dim TableReportList As New List(Of List(Of String))
        Dim OriginalContent As String = ""

        ' Get original content to maintain line spacing from template
        For Each item In PageReport
            If InStr(item, "{TableContent}") Then
                OriginalContent = item
            End If
        Next

        For Each item In ReportOutputList
            Dim NewTableRow As New List(Of String)
            NewTableRow = GetTableReportRowList(item, OriginalContent)
            TableReportList.Add(NewTableRow)
        Next

        For Each tRow In TableReportList
            For Each item In tRow
                NewList.Add(item)
            Next
        Next

        Return NewList

    End Function

    Public Function GetTableReportRowList(ByRef Item As ReportMember, ByRef OC As String) As List(Of String)

        Dim NewList As New List(Of String)
        Dim OFT As String ' Original Format
        OFT = Replace(OC, "{TableContent}", "") ' Gets the line indent from the report template
        NewList.Add(OFT & "<tr>")
        NewList.Add(OFT & "    <td>" & Item.NestLevel & "</td>")
        NewList.Add(OFT & "    <td>" & Item.ObjectName & "</td>")
        NewList.Add(OFT & "    <td>" & Item.ObjectType & "</td>")
        NewList.Add(OFT & "    <td>" & Item.Parameter & "</td>")
        NewList.Add(OFT & "    <td>" & Item.LeftValue & "</td>")
        NewList.Add(OFT & "    <td>" & Item.RightValue & "</td>")
        NewList.Add(OFT & "</tr>")
        Return NewList

    End Function

    ''' <summary>
    ''' Process the summary 2 page. Takes the template report page as an argument
    ''' </summary>
    ''' <param name="Slist">Template report page</param>
    ''' <returns>A new list containing the original template content, with a list of pages included whose report objects indicate differences were found</returns>
    Public Function GetSummary2(ByRef Slist As List(Of String)) As List(Of String)

        Dim PageList As List(Of String)
        Dim NewList As List(Of String) = New List(Of String)
        PageList = GetPagesWithDifferenceList()
        PageList = FormatListasHTMLAnchorListElement(PageList)

        For Each item In Slist
            If Not InStr(item, "{ListContent}") > 0 Then
                NewList.Add(item)
            Else
                Dim OriginalFormat As String = item
                For Each PageItem In PageList
                    Dim NewItem As String = OriginalFormat.Replace("{ListContent}", PageItem)
                    NewList.Add(NewItem)
                Next
            End If
        Next

        Return NewList

    End Function

    ''' <summary>
    ''' Wrap a supplied string with html list element and anchor tags
    ''' </summary>
    ''' <param name="Slist">String to format</param>
    ''' <returns>String formatted with html list elements, anchor tag and wrapped in bold tags</returns>
    Public Function FormatListasHTMLAnchorListElement(ByRef Slist As List(Of String)) As List(Of String)
        '<li><a href="#Page3"><b>Page 3</b></a></li>
        For a = 0 To Slist.Count - 1
            Slist.Item(a) = "<li><a href=""#" & Replace(Slist.Item(a), " ", "_") &
                            """><b>" & Slist.Item(a) & "</b></a></li>"
        Next
        Return Slist
    End Function

    ''' <summary>
    ''' Process the summary 1 page. Takes the template report page as an argument
    ''' </summary>
    ''' <param name="Slist">Template report page</param>
    ''' <returns>A new list containing the original template content, with a list of pages included whose report objects indicate no differences were found</returns>
    Public Function GetSummary1(ByRef Slist As List(Of String)) As List(Of String)

        Dim PageList As List(Of String)
        Dim NewList As List(Of String) = New List(Of String)
        PageList = GetPagesNoDifferenceList()
        PageList = FormatListasHTMLListElement(PageList)

        For Each item In Slist
            If Not InStr(item, "{ListContent}") > 0 Then
                NewList.Add(item)
            Else
                Dim OriginalFormat As String = item
                For Each PageItem In PageList
                    Dim NewItem As String = OriginalFormat.Replace("{ListContent}", PageItem)
                    NewList.Add(NewItem)
                Next
            End If
        Next

        Return NewList

    End Function

    ''' <summary>
    ''' Wrap a supplied string with html list element tags
    ''' </summary>
    ''' <param name="Slist">String to format</param>
    ''' <returns>String formatted with html list elements and wrapped in bold tags</returns>
    Public Function FormatListasHTMLListElement(ByRef Slist As List(Of String)) As List(Of String)
        For a = 0 To Slist.Count - 1
            Slist.Item(a) = "<li><b>" & Slist.Item(a) & "</b></li>"
        Next
        Return Slist
    End Function

    ''' <summary>
    ''' Processes the supplied string list containing the reportsummary.html template and substitutes the tag values for their correct variable values
    ''' The variable values to use are calculated in this function also
    ''' </summary>
    ''' <param name="Slist">String list containing the tag values to substitute</param>
    ''' <returns>The supplied list with the tag values substituted with the calculated values</returns>
    Public Function GetReportSummary(ByRef Slist As List(Of String)) As List(Of String)

        Dim TotalDifferences As Integer = 0
        Dim TotalPages As Integer = 0
        Dim PagesWithNoDifference As Integer = 0
        Dim PagesWithDifference As Integer = 0

        For Each page In PageList
            TotalPages += 1
            If page.OutputList.Count = 0 Then
                PagesWithNoDifference += 1
            Else
                PagesWithDifference += 1
                TotalDifferences += page.OutputList.Count
            End If
        Next

        Slist = SearchListForSubstituteString(Slist, "{TotalDifferences}", TotalDifferences)
        Slist = SearchListForSubstituteString(Slist, "{TotalPages}", TotalPages)
        Slist = SearchListForSubstituteString(Slist, "{PagesNoDifference}", PagesWithNoDifference)
        Slist = SearchListForSubstituteString(Slist, "{PagesWithDifference}", PagesWithDifference)

        Return Slist

    End Function

    ''' <summary>
    ''' Searches a given list of strings for the presence of a substitute tag and replaces it with the substitution value supplied.
    ''' Each value is expected to be found once only hence the function returns immediately after the first substitution is found
    ''' </summary>
    ''' <param name="Slist">List of strings to check for the tag</param>
    ''' <param name="SubstitutionTag">Tag value of substitution parameter to find. Tag strings must be enclosed in {} for proper function e.g "{TagName}"</param>
    ''' <param name="SubstitutionValue">Value to replace in the string with if the tag is found</param>
    ''' <returns>The supplied list with the value substituted if found</returns>
    Public Function SearchListForSubstituteString(ByRef Slist As List(Of String),
                                                  ByRef SubstitutionTag As String,
                                                  ByRef SubstitutionValue As String) As List(Of String)
        For a = 0 To Slist.Count - 1
            Slist.Item(a) = ReturnStringWithSubstitutionValue(Slist.Item(a),
                                                              SubstitutionTag,
                                                              SubstitutionValue)
        Next

        Return Slist

    End Function

    ''' <summary>
    ''' Searches a given string for the presence of a substitution tag and replaces it if found
    ''' </summary>
    ''' <param name="SearchString">String to search</param>
    ''' <param name="SubstitutionTag">Tag value of substitution parameter to find. Tag strings must be enclosed in {} for proper function  e.g "{TagName}"</param>
    ''' <param name="SubstitutionValue">Value to replace in the string with if the tag is found</param>
    ''' <returns>The string with the tag replaced with the supplied value if found</returns>
    Public Function ReturnStringWithSubstitutionValue(ByRef SearchString As String,
                                                      ByRef SubstitutionTag As String,
                                                      ByRef SubstitutionValue As String)
        Dim Wrkstr As String = SearchString
        Wrkstr = Replace(Wrkstr, SubstitutionTag, SubstitutionValue)
        Return Wrkstr

    End Function

    Public Function GetPagesWithDifferenceList() As List(Of String)
        Dim Nlist As New List(Of String)
        For Each Page In PageList
            If Page.OutputList.Count > 0 Then
                Nlist.Add(Page.LeftFile)
            End If
        Next
        Return Nlist
    End Function

    Public Function GetPagesNoDifferenceList() As List(Of String)
        Dim Nlist As New List(Of String)
        For Each Page In PageList
            If Page.OutputList.Count = 0 Then
                Nlist.Add(Page.LeftFile)
            End If
        Next
        Return Nlist
    End Function

    ''' <summary>
    ''' Reads a complete file into a list of string variable
    ''' </summary>
    ''' <param name="Path">Fully qualified path to the file to read</param>
    ''' <returns>A new list of string containing the file contents</returns>
    Public Function ReadFileToList(ByRef Path As String) As List(Of String)
        Dim Nlist As List(Of String) = New List(Of String)
        Using R As StreamReader = New StreamReader(Path)
            Do
                Nlist.Add(R.ReadLine)
            Loop Until R.EndOfStream
        End Using
        Return Nlist
    End Function

    ''' <summary>
    ''' Gets the path to current executing assembly to provide a fully qualified path to the supplied file name
    ''' </summary>
    ''' <param name="Filename">Name of file to add to the working directory path</param>
    ''' <returns>A string containing the fully qualified path to the supplied file name using the current working directory as context</returns>
    Public Function GetWorkingDirectoryPath(ByRef Filename As String) As String
        Dim wrkstr As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
        wrkstr &= "\" & Filename
        Return wrkstr
    End Function

    ''' <summary>
    ''' Look in user system registry and determine the default browser
    ''' </summary>
    ''' <returns>a string containing the name of the default browser</returns>
    Public Function DefaultWebBrowser() As String
        Dim path As String = "\http\shell\open\command"
        Using Reg As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(path)
            If Reg IsNot Nothing Then
                Dim webBrowserPath As String = Reg.GetValue(String.Empty).ToString
                If Not webBrowserPath = "" Then
                    If webBrowserPath IsNot Nothing Then
                        If webBrowserPath.First() = """" Then
                            Dim tarr()
                            tarr = Split(webBrowserPath, """")
                            Return tarr(1)
                        Else
                            Dim tarr()
                            tarr = Split(webBrowserPath, " ")
                            Return tarr(0)
                        End If
                    End If
                End If
            End If
            Return ""
        End Using
    End Function

    ''' <summary>
    ''' run a program with supplied arguments
    ''' </summary>
    ''' <param name="FileName">Filename and path of the executable to run (must have .exe on the end to work)</param>
    ''' <param name="Args">Any arguments supplied to the program at runtime</param>
    ''' <returns>Boolean True if process start was called successfully, False if a problem was encountered</returns>
    Public Function Run(ByRef FileName As String, ByRef Args As String) As Boolean

        Try
            Dim proc As New ProcessStartInfo
            proc.FileName = FileName
            proc.WindowStyle = ProcessWindowStyle.Normal
            If Not Args = "" Then
                proc.Arguments = Args
            End If
            System.Diagnostics.Process.Start(proc)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

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
        Dim isMultiState As Boolean = False
        Dim SearchParams As String() = {
            OutputMembers.FileName,
            OutputMembers.NestLevel,
            OutputMembers.ObjectName,
            OutputMembers.ObjectType,
            OutputMembers.Parameter,
            OutputMembers.LeftValue,
            OutputMembers.RightValue}

        If outputline.Contains("State") Then
            isMultiState = True
        End If

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
                                    If isMultiState Then
                                        Parameter = intr(1) & " - " & tarr(5)
                                    Else
                                        Parameter = intr(1)
                                    End If
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
