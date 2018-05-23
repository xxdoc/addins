VERSION 5.00
Begin VB.Form frmSelelectedControls 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4152
   ClientLeft      =   5760
   ClientTop       =   6300
   ClientWidth     =   6012
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4152
   ScaleWidth      =   6012
   Begin VB.TextBox txtSelectedControls 
      Height          =   3504
      Left            =   144
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   432
      Width           =   5664
   End
   Begin VB.Label Label1 
      Caption         =   "Selected controls:"
      Height          =   264
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1452
   End
End
Attribute VB_Name = "frmSelelectedControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mCtrlHandler As VBIDE.SelectedVBControlsEvents
Attribute mCtrlHandler.VB_VarHelpID = -1
Private WithEvents mCompHandler As VBIDE.VBComponentsEvents
Attribute mCompHandler.VB_VarHelpID = -1
Private mSelectedVBComponent As VBIDE.VBComponent
Private WithEvents mProjHandler As VBIDE.VBProjectsEvents
Attribute mProjHandler.VB_VarHelpID = -1

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub AlwaysOnTop(frm As Form, ByVal bOnTop As Boolean)
    ' Toggles "AlwaysOnTop" property of a window
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const Flags = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If bOnTop Then
        SetWindowPos frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags
    Else
        SetWindowPos frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags
    End If
End Sub

Private Sub Form_Load()
    AlwaysOnTop Me, True
    UpdateHandlers
End Sub

Private Sub mCompHandler_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
    Set mCtrlHandler = VBInstance.Events.SelectedVBControlsEvents(VBInstance.ActiveVBProject, VBInstance.SelectedVBComponent.Designer)
    Set mSelectedVBComponent = VBComponent
    UpdateList
End Sub

Private Sub mCompHandler_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
    UpdateList
End Sub

Private Sub mCompHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    Set mCtrlHandler = VBInstance.Events.SelectedVBControlsEvents(VBInstance.ActiveVBProject, VBInstance.SelectedVBComponent.Designer)
    Set mSelectedVBComponent = VBComponent
    UpdateList
End Sub

Private Sub mCtrlHandler_ItemAdded(ByVal VBControl As VBIDE.VBControl)
    UpdateList VBControl.ControlObject.Name
End Sub

Private Sub mCtrlHandler_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
    UpdateList , VBControl.ControlObject.Name
End Sub

Private Sub UpdateList(Optional CtrlAdd, Optional CtrlRemove)
    Dim iCtrl As VBControl
    Dim iInclude As Boolean
    
    
    Me.Caption = VBInstance.ActiveVBProject.Name & "." & mSelectedVBComponent.Name
    txtSelectedControls.Text = ""
    For Each iCtrl In mSelectedVBComponent.Designer.VBControls
        If iCtrl.InSelection Then
            iInclude = True
            If Not IsMissing(CtrlRemove) Then
                If iCtrl.ControlObject.Name = CtrlRemove Then
                    iInclude = False
                End If
            End If
            If Not IsMissing(CtrlAdd) Then
                If iCtrl.ControlObject.Name = CtrlAdd Then
                    CtrlAdd = ""
                End If
            End If
            If iInclude Then
                txtSelectedControls.Text = txtSelectedControls.Text & iCtrl.ControlObject.Name & " (" & TypeName(iCtrl.ControlObject) & ")" & vbCrLf
            End If
        End If
    Next
    If Not IsMissing(CtrlAdd) Then
        If CtrlAdd <> "" Then
            txtSelectedControls.Text = txtSelectedControls.Text & CtrlAdd & " (" & TypeName(mSelectedVBComponent.Designer.VBControls(CtrlAdd).ControlObject) & ")" & vbCrLf
        End If
    End If
End Sub

Private Sub mProjHandler_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    SetCurrentHandlers
End Sub

Private Sub mProjHandler_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    SetCurrentHandlers
End Sub

Private Sub mProjHandler_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    SetCurrentHandlers
End Sub

Private Sub mProjHandler_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    SetCurrentHandlers
End Sub


Private Sub SetCurrentHandlers()
    
    If Not VBInstance.ActiveVBProject Is Nothing Then
        Set mCompHandler = Nothing
        Set mCompHandler = VBInstance.Events.VBComponentsEvents(VBInstance.ActiveVBProject)
    End If
    
    If Not VBInstance.SelectedVBComponent Is Nothing Then
        If VBInstance.SelectedVBComponent.HasOpenDesigner Then
            Set mCtrlHandler = Nothing
            Set mCtrlHandler = VBInstance.Events.SelectedVBControlsEvents(VBInstance.ActiveVBProject, VBInstance.SelectedVBComponent.Designer)
            Set mSelectedVBComponent = Nothing
            Set mSelectedVBComponent = VBInstance.SelectedVBComponent
            UpdateList
        Else
            Me.Caption = "No designer selected"
            txtSelectedControls.Text = ""
        End If
    Else
        Me.Caption = "No designer selected"
        txtSelectedControls.Text = ""
    End If
End Sub

Public Sub UpdateHandlers()
    Set mProjHandler = Nothing
    Set mProjHandler = VBInstance.Events.VBProjectsEvents
    SetCurrentHandlers
End Sub
