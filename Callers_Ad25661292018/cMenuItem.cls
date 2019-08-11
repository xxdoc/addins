VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit ' Menu Item Class - adapted from Darryl Hasieber :)

Private WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
Private oCmdBarCtl As Office.CommandBarControl

Public Sub Add(oPopup As Office.CommandBar, sCaption As String, ByVal ArrayIndex As Long, Optional Bitmap As Object)
11  On Error GoTo ErrHandler              ' Create sub-item in the menu
12    Set oCmdBarCtl = oPopup.Controls.Add(Type:=msoControlButton, Parameter:=ArrayIndex)
13    oCmdBarCtl.Caption = sCaption       ' Assign the specified caption
14    Set MenuHandler = oVBE.Events.CommandBarEvents(oCmdBarCtl)
15    If Not Bitmap Is Nothing Then
16        Dim oTemp As Object, sClipText As String
17      On Error GoTo ErrWith
18        With Clipboard
19            sClipText = .GetText
20            Set oTemp = .GetData
21            .SetData Bitmap, vbCFBitmap ' Copy the icon to the clipboard
22            oCmdBarCtl.PasteFace        ' Set the icon for the button
23            .Clear
24            If Not oTemp Is Nothing Then .SetData oTemp
25            .SetText sClipText
ErrWith:
        End With
    End If
ErrHandler:
  If Err Then LogError "cMenuItem.Add", sCaption
End Sub

Public Sub Remove()
    If Not oCmdBarCtl Is Nothing Then
       oCmdBarCtl.Delete
       Set oCmdBarCtl = Nothing
       Set MenuHandler = Nothing
    End If
End Sub

Private Sub Class_Terminate()
    Call Remove
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    With CommandBarControl
        CodePaneMenuItem_Click .Parameter
    End With
End Sub

