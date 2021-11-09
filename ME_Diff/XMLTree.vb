Public Class XMLTree
    Public Members As List(Of XMLMember)
    Public Sub New()
        Members = New List(Of XMLMember)
    End Sub

    Public Sub ParseList(ByRef xmllist As List(Of String))
        Dim CNL As Integer = 0 ' current nest level
        Dim PNL As Integer = 0 ' previous nest level
        Dim CMX As XMLMember ' Current member context
        ' 1st pass get all content
        For Each itm In xmllist
            CNL = GetNestLevel(itm)
            CMX = New XMLMember(CNL, itm)
            Members.Add(CMX)
        Next

    End Sub

    Public Function CheckisCompleted(ByRef newlist As List(Of XMLMember), ByRef NLvl As Integer) As Boolean
        Dim retval As Boolean = True
        For Each itm In newlist
            If itm.NestLevel > NLvl Then
                retval = False ' some sub level components remain on the top level list
            End If
            If itm.Children.Count > 0 Then
                If CheckisCompleted(itm.Children, itm.NestLevel) = False Then
                    retval = False
                End If
            End If
        Next
        Return retval
    End Function

    Public Function GetNestLevel(ByRef xmlstring As String) As Integer
        Dim WSC As Integer = 0 ' White space count
        For i = 0 To xmlstring.Length - 1
            If xmlstring(i) = " " Then WSC += 1
        Next
        Dim retval = (WSC / 3) - 1
        If retval < 0 Then retval = 0
        Return retval
    End Function
End Class

Public Class XMLMember
    Public NestLevel As Integer
    Public MemberID As String
    Public Properties As List(Of XMLProperty)
    Public Children As List(Of XMLMember)
    Private data As String
    Public Sub New(ByRef NLVL As Integer, ByRef XMLDAT As String)
        Properties = New List(Of XMLProperty)
        Children = New List(Of XMLMember)
        NestLevel = NLVL
        data = RemoveWhitespace(XMLDAT)
        MemberID = GetMemberName(XMLDAT)
        Properties = GetProperties(XMLDAT)
    End Sub
    Public Function GetProperties(ByRef str As String) As List(Of XMLProperty)
        Dim tar() As String
        Dim wrklist As New List(Of XMLProperty)
        tar = str.Split(" ")
        For a = 1 To tar.Length - 1 ' index starts at 1 because the 0 element is the member ID which we have already processed 
            Dim newXMLProp As New XMLProperty
            newXMLProp = GetProperty(tar(a))
            wrklist.Add(newXMLProp)
        Next
        Return wrklist
    End Function
    Public Function GetProperty(ByRef str As String) As XMLProperty
        Dim tar() As String
        Dim retval As New XMLProperty
        tar = str.Split("=")
        retval.Parameter = tar(0)
        retval.Value = tar(1).Replace("""", "")
        retval.Value = retval.Value.Replace(">", "")
        Return retval
    End Function
    Public Function GetMemberName(ByRef str As String) As String
        Dim tar() As String
        Dim wrkstr As String
        tar = Split(str, " ")
        wrkstr = tar(0)
        wrkstr = wrkstr.Replace("<", "")
        wrkstr = wrkstr.Replace(">", "")
        wrkstr = wrkstr.Replace("/", "")
        Return wrkstr
    End Function
    Public Function RemoveWhitespace(ByRef str As String) As String
        str = str.Trim()
        Return str
    End Function
    Public Sub AddChildren(ByRef NewMember As XMLMember)
        Children.Add(NewMember)
    End Sub
    Public Sub AddMember(ByRef NLVL As Integer, ByRef XMLDAT As String)

    End Sub
End Class

Public Class XMLProperty
    Public Parameter As String
    Public Value As String
End Class
