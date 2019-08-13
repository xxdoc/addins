Attribute VB_Name = "modErrors"
Option Explicit                                    ' ©Rd

Private mFile As String
Private mInit As Boolean

Public Sub InitErr(Optional sCompName As String, Optional ByVal fClearMsgLog As Boolean)
   On Error GoTo Fail
    Dim i As Integer
    If Not mInit Then
        If LenB(sCompName) = 0& Then sCompName = App.EXEName
        If Right$(App.Path, 1&) = "\" Then
            mFile = App.Path & sCompName
        Else
            mFile = App.Path & "\" & sCompName
        End If
        If fClearMsgLog Then
            i = FreeFile()
            Open mFile & "_Msg.log" For Output As #i
            Close #i
        End If
        mInit = True
    End If
Fail:
End Sub

Public Sub LogError(sProcName As String, Optional sExtraInfo As String)
    Dim Num As Long, Src As String, Desc As String
    With Err
      Num = .Number: Src = .Source: Desc = .Description
    End With
    If Erl Then Desc = Desc & vbCrLf & "Error on line " & Erl
   On Error GoTo Fail
    If LenB(sExtraInfo) Then Desc = Desc & vbCrLf & sExtraInfo
    If mInit Then Else InitErr
    Dim i As Integer: i = FreeFile()
    Open mFile & "_Error.log" For Append As #i
        Print #i, Src; " error ";
        Print #i, Format$(Now, "h:nn:ss am/pm mmmm d, yyyy")
        Print #i, sProcName; " error!"
        Print #i, "Error #"; Num; " - "; Desc
        Print #i, " * * * * * * * * * * * * * * * * * * *"
Fail:
    Close #i
    Beep
End Sub

Public Sub LogMsg(Msg As String)
   On Error GoTo Fail
    If mInit Then Else InitErr
    Dim i As Integer: i = FreeFile()
    Open mFile & "_Msg.log" For Append As #i
        Print #i, Format$(Now, "h:nn:ss am/pm mmmm d, yyyy")
        Print #i, Msg
        Print #i, " * * * * * * * * * * * * * * * * * * *"
    Close #i
Fail:
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  :›)
