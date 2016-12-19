Module modGlobal
    '''<summary>
    ''' The Get Command will looks for Command Line Arguments, this on will return as string
    ''' the switch will be something like /mystring="this is fun"
    ''' if it is just /mystring then it will return what is set in the sDefault string.
    ''' </summary>
    Public Function GetCommand(ByVal strLookFor As String, ByVal sDefault As String, Optional ByRef DidExist As Boolean = False, Optional ByRef Switch As String = "/") As String
        Dim sAns As String = sDefault
        DidExist = False
        Dim cmdLine() As String = System.Environment.GetCommandLineArgs
        Dim i As Integer = 0
        Dim intCount As Integer = cmdLine.Length
        Dim strValue As String = ""
        If intCount > 1 Then
            For i = 1 To intCount - 1
                strValue = cmdLine(i)
                strValue = Replace(strValue, Switch, "")
                Dim strSplit() As String = Split(strValue, "=")
                Dim intLBound As Integer = LBound(strSplit)
                Dim intUBound As Integer = UBound(strSplit)
                If LCase(strSplit(intLBound)) = LCase(strLookFor) Then
                    If intUBound <> 0 Then
                        sAns = strSplit(intUBound)
                    Else
                        sAns = sDefault
                    End If
                    DidExist = True
                    Exit For
                End If
            Next
        End If
        Return LCase(sAns)
    End Function
    '''<summary>
    ''' The Get Command will looks for Command Line Arguments, this on will return as long
    ''' the switch will be something like /mylongvalue=92
    ''' if it is just /mylongvalue it will return the lDefault value
    ''' </summary>
    Public Function GetCommand(ByVal strLookFor As String, ByVal lDefault As Long, Optional ByRef DidExist As Boolean = False, Optional ByRef Switch As String = "/") As Long
        Dim lAns As Long = 0
        DidExist = False
        Dim cmdLine() As String = System.Environment.GetCommandLineArgs
        Dim i As Integer = 0
        Dim intCount As Integer = cmdLine.Length
        Dim strValue As String = ""
        If intCount > 1 Then
            For i = 1 To intCount - 1
                strValue = cmdLine(i)
                strValue = Replace(strValue, Switch, "")
                Dim strSplit() As String = Split(strValue, "=")
                Dim intLBound As Integer = LBound(strSplit)
                Dim intUBound As Integer = UBound(strSplit)
                If LCase(strSplit(intLBound)) = LCase(strLookFor) Then
                    If intUBound <> 0 Then
                        lAns = strSplit(intUBound)
                    Else
                        lAns = lDefault
                    End If
                    DidExist = True
                    Exit For
                End If
            Next
        End If
        Return lAns
    End Function
    '''<summary>
    ''' The Get Command will looks for Command Line Arguments, this on will return as boolean.
    ''' if the command is /swtich it will return as true since it did exist
    ''' you can also use /switch=false
    ''' </summary>
    Public Function GetCommand(ByVal strLookFor As String, ByVal bDefault As Boolean, Optional ByRef DidExist As Boolean = False, Optional ByRef Switch As String = "/") As Boolean
        Dim bAns As Boolean = bDefault
        DidExist = False
        Dim cmdLine() As String = System.Environment.GetCommandLineArgs
        Dim i As Integer = 0
        Dim intCount As Integer = cmdLine.Length
        Dim strValue As String = ""
        If intCount > 1 Then
            For i = 1 To intCount - 1
                strValue = cmdLine(i)
                strValue = Replace(strValue, Switch, "")
                Dim strSplit() As String = Split(strValue, "=")
                Dim intLBound As Integer = LBound(strSplit)
                Dim intUBound As Integer = UBound(strSplit)
                If LCase(strSplit(intLBound)) = LCase(strLookFor) Then
                    If intUBound <> 0 Then
                        bAns = strSplit(intUBound)
                    Else
                        bAns = True
                    End If
                    DidExist = True
                    Exit For
                End If
            Next
        End If
        Return bAns
    End Function
End Module
