Attribute VB_Name = "mKeyCapt"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapseMilliseconds As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private hWndVBinst As Long
Private lTimerId As Long

Sub KeyCaptureStart(ByVal VB_hWnd As Long)
    If lTimerId = 0& Then
        hWndVBinst = VB_hWnd
        lTimerId = SetTimer(0&, 0&, 50&, AddressOf TimerProc)
    End If
End Sub

Sub KeyCaptureEnd()
    KillTimer 0&, lTimerId
    lTimerId = 0&
End Sub

Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)

    Dim fAlt As Boolean
    Dim fWin As Boolean
    Dim fF2 As Boolean
    Dim fF3 As Boolean
    Const vbKeyWinL = &H5B&

    If GetForegroundWindow <> hWndVBinst Then Exit Sub
    KillTimer 0&, idEvent

    fAlt = (GetAsyncKeyState(vbKeyMenu) And &H8000) = &H8000
    If fAlt Then

        fWin = (GetAsyncKeyState(vbKeyWinL) And &H8000) = &H8000
        fF2 = (GetAsyncKeyState(vbKeyF2) And &H8001) = &H8001
        fF3 = (GetAsyncKeyState(vbKeyF3) And &H8001) = &H8001

        If fWin And fF2 Then
            If nCallCnt Then DisplayCallee

        ElseIf fWin And fF3 Then
            If nCallCnt Then JumpToPrevReference

        ElseIf fF3 Then
            If nCallers Then JumpToNextReference

        ElseIf fF2 Then
            Call RefreshMemberReferences
            If nCallers Then oPopupMenu.ShowPopup

        End If
    End If

    If lTimerId <> 0& Then
        lTimerId = SetTimer(0&, 0&, 50&, AddressOf TimerProc)
    End If
End Sub
