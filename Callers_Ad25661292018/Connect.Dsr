VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9870
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   17025
   _ExtentX        =   30030
   _ExtentY        =   17410
   _Version        =   393216
   Description     =   "Callers AddIn"
   DisplayName     =   "Callers AddIn"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private WithEvents ProjectsEvents As VBProjectsEvents
'Private WithEvents ComponentEvents As VBComponentsEvents

Private iGroupIdx As Long
Private iHotKeyId As Long

Private oCodePaneMenuCallers As Office.CommandBarControl
Private oCodePaneMenuCallee As Office.CommandBarControl
Private oImmediateWindMenu As Office.CommandBarControl
Private oCodePaneMenuMembers As Office.CommandBarControl

Private oCodePaneMethods As Office.CommandBarControl
Private oCodePaneProperties As Office.CommandBarControl
Private oCodePaneVariables As Office.CommandBarControl
Private oCodePaneEvents As Office.CommandBarControl
Private oCodePaneConstants As Office.CommandBarControl
Private oCodePaneTypesEnums As Office.CommandBarControl

Private WithEvents CallersHandler As CommandBarEvents
Attribute CallersHandler.VB_VarHelpID = -1
Private WithEvents CalleeHandler As CommandBarEvents
Attribute CalleeHandler.VB_VarHelpID = -1
Private WithEvents ImmediateHandler As CommandBarEvents
Attribute ImmediateHandler.VB_VarHelpID = -1

Private WithEvents MethodsHandler As CommandBarEvents
Attribute MethodsHandler.VB_VarHelpID = -1
Private WithEvents PropertiesHandler As CommandBarEvents
Attribute PropertiesHandler.VB_VarHelpID = -1
Private WithEvents VariablesHandler As CommandBarEvents
Attribute VariablesHandler.VB_VarHelpID = -1
Private WithEvents EventsHandler As CommandBarEvents
Attribute EventsHandler.VB_VarHelpID = -1
Private WithEvents ConstantsHandler As CommandBarEvents
Attribute ConstantsHandler.VB_VarHelpID = -1
Private WithEvents TypesEnumsHandler As CommandBarEvents
Attribute TypesEnumsHandler.VB_VarHelpID = -1

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
 On Error GoTo ErrHandler
    'Set ProjectsEvents = oVBE.Events.VBProjectsEvents()
   DoEvents
   If AddItemToMenu("L&ist Code Members...", "Code Window", msoControlPopup, oCodePaneMenuMembers, , 4, True) Then
      If AddItemToSubMenu(oCodePaneMenuMembers, "Constants", oCodePaneConstants) Then
           Set ConstantsHandler = oVBE.Events.CommandBarEvents(oCodePaneConstants)
      End If
      If AddItemToSubMenu(oCodePaneMenuMembers, "Events", oCodePaneEvents) Then
           Set EventsHandler = oVBE.Events.CommandBarEvents(oCodePaneEvents)
      End If
      If AddItemToSubMenu(oCodePaneMenuMembers, "Methods", oCodePaneMethods) Then
           Set MethodsHandler = oVBE.Events.CommandBarEvents(oCodePaneMethods)
      End If
      If AddItemToSubMenu(oCodePaneMenuMembers, "Properties", oCodePaneProperties) Then
           Set PropertiesHandler = oVBE.Events.CommandBarEvents(oCodePaneProperties)
      End If
      If AddItemToSubMenu(oCodePaneMenuMembers, "Types-Enums", oCodePaneTypesEnums) Then
           Set TypesEnumsHandler = oVBE.Events.CommandBarEvents(oCodePaneTypesEnums)
      End If
      If AddItemToSubMenu(oCodePaneMenuMembers, "Variables", oCodePaneVariables) Then
           Set VariablesHandler = oVBE.Events.CommandBarEvents(oCodePaneVariables)
      End If
   End If
   If AddItemToMenu("Return to Call&ee", "Code Window", msoControlButton, oCodePaneMenuCallee, LoadResPicture("Left", vbResBitmap), 4) Then
      Set CalleeHandler = oVBE.Events.CommandBarEvents(oCodePaneMenuCallee)
   End If
   If AddItemToMenu("Display Ca&llers...", "Code Window", msoControlButton, oCodePaneMenuCallers, LoadResPicture("Right", vbResBitmap), 4, True) Then
      Set CallersHandler = oVBE.Events.CommandBarEvents(oCodePaneMenuCallers)
   End If
   If AddItemToMenu("C&lear", "Immediate Window", msoControlButton, oImmediateWindMenu, , oVBE.CommandBars("Immediate Window").Controls.Count - 1&) Then
      Set ImmediateHandler = oVBE.Events.CommandBarEvents(oImmediateWindMenu)
   End If
   Call RedimCallers(100)
   KeyCaptureStart oVBE.MainWindow.hWnd
ErrHandler:
 If Err Then LogError "Connect.AddinInstance_OnStartupComplete"
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
 On Error GoTo ErrHandler
   If ConnectMode = ext_cm_Startup Then
      Call InitErr("CallersAddin")
      Set oVBE = Application
   End If
ErrHandler:
 If Err Then LogError "Connect.AddinInstance_OnConnection"
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    Call KeyCaptureEnd
    Call ResetContextMenu
    Call EraseCallerArrays
    If Not oCodePaneMethods Is Nothing Then
       Call oCodePaneMethods.Delete
       Set oCodePaneMethods = Nothing
       Set MethodsHandler = Nothing
    End If
    If Not oCodePaneProperties Is Nothing Then
       Call oCodePaneProperties.Delete
       Set oCodePaneProperties = Nothing
       Set PropertiesHandler = Nothing
    End If
    If Not oCodePaneVariables Is Nothing Then
       Call oCodePaneVariables.Delete
       Set oCodePaneVariables = Nothing
       Set VariablesHandler = Nothing
    End If
    If Not oCodePaneEvents Is Nothing Then
       Call oCodePaneEvents.Delete
       Set oCodePaneEvents = Nothing
       Set EventsHandler = Nothing
    End If
    If Not oCodePaneConstants Is Nothing Then
       Call oCodePaneConstants.Delete
       Set oCodePaneConstants = Nothing
       Set ConstantsHandler = Nothing
    End If
    If Not oCodePaneTypesEnums Is Nothing Then
       Call oCodePaneTypesEnums.Delete
       Set oCodePaneTypesEnums = Nothing
       Set TypesEnumsHandler = Nothing
    End If
    If Not oCodePaneMenuMembers Is Nothing Then
       Call oCodePaneMenuMembers.Delete
       Set oCodePaneMenuMembers = Nothing
    End If
    If Not oCodePaneMenuCallers Is Nothing Then
       Call oCodePaneMenuCallers.Delete
       Set oCodePaneMenuCallers = Nothing
       Set CallersHandler = Nothing
    End If
    If Not oCodePaneMenuCallee Is Nothing Then
       Call oCodePaneMenuCallee.Delete
       Set oCodePaneMenuCallee = Nothing
       Set CalleeHandler = Nothing
    End If
    If Not oImmediateWindMenu Is Nothing Then
       Call oImmediateWindMenu.Delete
       Set oImmediateWindMenu = Nothing
       Set ImmediateHandler = Nothing
    End If
    'Set ComponentEvents = Nothing
    'Set ProjectsEvents = Nothing
    Set oVBE = Nothing
End Sub

Private Sub AddinInstance_OnAddInsUpdate(custom() As Variant)
'
End Sub

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
'
End Sub

Private Function AddItemToMenu(sCaption As String, sMenuName As String, cbcType As Office.MsoControlType, cbMenuCBar As Office.CommandBarControl, Optional oBitmap As Object, Optional ByVal lBefore As Long = 4, Optional ByVal bGroup As Boolean) As Boolean
   Dim cbMenu As Office.CommandBar
   Dim oTemp As Object, sClipText As String
  On Error GoTo AddItemError
   Set cbMenu = oVBE.CommandBars(sMenuName)
   If cbMenu Is Nothing Then Exit Function

   If lBefore = 4& Then
      If iGroupIdx = 0& Then
         iGroupIdx = 4&
         Do Until cbMenu.Controls(iGroupIdx).BeginGroup Or iGroupIdx = cbMenu.Controls.Count - 1&
            iGroupIdx = iGroupIdx + 1&
         Loop
      End If
      lBefore = iGroupIdx
      ' Create new item in the menu
      Set cbMenuCBar = cbMenu.Controls.Add(Type:=cbcType, Before:=lBefore)
      cbMenuCBar.BeginGroup = bGroup
   Else
      ' Create new item in the menu
      Set cbMenuCBar = cbMenu.Controls.Add(Type:=cbcType)
   End If

   cbMenuCBar.Caption = sCaption  ' Assign the specified caption
    AddItemToMenu = True          ' Return success to caller

     If Not oBitmap Is Nothing Then
       On Error GoTo ErrWith
         With Clipboard
            sClipText = .GetText
            Set oTemp = .GetData
            .SetData oBitmap, vbCFBitmap ' Copy the icon to the clipboard
            cbMenuCBar.PasteFace         ' Set the icon for the button
            .Clear
            If Not oTemp Is Nothing Then
               .SetData oTemp
               Set oTemp = Nothing
            End If
            .SetText sClipText
ErrWith:
        End With
    End If
AddItemError:
  If Err = 521 Then ' Can't open clipboard
  ElseIf Err Then LogError "Connect.AddItemToMenu", sMenuName
  End If
End Function

Private Function AddItemToSubMenu(cbMenu As Office.CommandBarControl, sCaption As String, cbMenuCBar As Office.CommandBarControl) As Boolean
    Dim oTemp As Object, sClipText As String
  On Error GoTo AddSubItemError
    If cbMenu Is Nothing Then Exit Function
                                   ' Create new item in the menu
    Set cbMenuCBar = cbMenu.Controls.Add(Type:=msoControlButton)
    cbMenuCBar.Caption = sCaption  ' Assign the caption"
    AddItemToSubMenu = True        ' Return success to caller

    On Error GoTo ErrWith
    With Clipboard
       sClipText = .GetText
       Set oTemp = .GetData ' Copy the icon to the clipboard
       .SetData LoadResPicture(sCaption, vbResBitmap), vbCFBitmap
       cbMenuCBar.PasteFace ' Set the icon for the button
       .Clear
       If Not oTemp Is Nothing Then
          .SetData oTemp
          Set oTemp = Nothing
       End If
       .SetText sClipText
ErrWith:
    End With
AddSubItemError:
  If Err = 521 Then ' Can't open clipboard
  ElseIf Err Then LogError "Connect.AddItemToSubMenu", cbMenu.Caption
  End If
End Function

Private Sub CallersHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    Call RefreshMemberReferences
    If nCallers Then oPopupMenu.ShowPopup
End Sub

Private Sub CalleeHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    If nCallCnt Then DisplayCallee
End Sub

Private Sub MethodsHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    Call RefreshComponentMembers(vbext_mt_Method, "Methods")
    If nCallers Then oPopupMenu.ShowPopup
End Sub

Private Sub PropertiesHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    Call RefreshComponentMembers(vbext_mt_Property, "Properties")
    If nCallers Then oPopupMenu.ShowPopup
End Sub

Private Sub TypesEnumsHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    Call RefreshComponentMembers(6&, "Types-Enums")
    If nCallers Then oPopupMenu.ShowPopup
End Sub

Private Sub VariablesHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    Call RefreshComponentMembers(vbext_mt_Variable, "Variables")
    If nCallers Then oPopupMenu.ShowPopup
End Sub

Private Sub ConstantsHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    Call RefreshComponentMembers(vbext_mt_Const, "Constants")
    If nCallers Then oPopupMenu.ShowPopup
End Sub

Private Sub EventsHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    Call RefreshComponentMembers(vbext_mt_Event, "Events")
    If nCallers Then oPopupMenu.ShowPopup
End Sub

Private Sub ImmediateHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
 On Error GoTo ErrHandler
    oVBE.Windows("Immediate").SetFocus
    SendKeys "^{Home}", True
    SendKeys "^+{End}", True
    SendKeys "{Del}", True
ErrHandler:
End Sub

'Private Sub ProjectsEvents_ItemActivated(ByVal VBProject As VBIDE.VBProject)
'    Set ComponentEvents = oVBE.Events.VBComponentsEvents(VBProject)
'End Sub
'
'Private Sub ProjectsEvents_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
'    If VBProject Is oVBE.ActiveVBProject Then
'        Set ComponentEvents = oVBE.Events.VBComponentsEvents(VBProject)
'    End If
'End Sub
'
'Private Sub ComponentEvents_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
'    RefreshMemberReferences
'End Sub
'
'Private Sub ComponentEvents_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
'    RefreshMemberReferences
'End Sub
'
'Private Sub ComponentEvents_ItemReloaded(ByVal VBComponent As VBIDE.VBComponent)
'    RefreshMemberReferences
'End Sub
'
'Private Sub ComponentEvents_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
'    RefreshMemberReferences
'End Sub
'
'Private Sub ComponentEvents_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
'    RefreshMemberReferences
'End Sub
