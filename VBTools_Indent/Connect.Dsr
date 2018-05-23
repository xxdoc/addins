VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11040
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   17550
   _ExtentX        =   30956
   _ExtentY        =   19473
   _Version        =   393216
   Description     =   "VBTools"
   DisplayName     =   "VBTools"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
' Made By Michael Ciurescu (CVMichael)
'

Public FormDisplayed As Boolean
Public VBInstance As VBIDE.VBE
Private mcbMenuCommandBar As CommandBar

Private cmdBar1 As CommandBarButton
Private cmdBar2 As CommandBarButton

Public WithEvents PrjHandler  As VBProjectsEvents    ' projects event handler
Attribute PrjHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler1 As CommandBarEvents   ' Auto Indent
Attribute MenuHandler1.VB_VarHelpID = -1
Public WithEvents MenuHandler2 As CommandBarEvents
Attribute MenuHandler2.VB_VarHelpID = -1

Public WithEvents HandlerOptions As CommandBarEvents
Attribute HandlerOptions.VB_VarHelpID = -1

Private objSortWindow As Window

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    Set VBInstance = Application
    Set mcbMenuCommandBar = VBInstance.CommandBars.Add("VBTools - Made By Michael Ciurescu (CVMichael) - Version: " & App.Major & "." & App.Minor, msoBarTop, False, True)
    
'    http://p2p.wrox.com/topic.asp?TOPIC_ID=18319
    With mcbMenuCommandBar
        .Enabled = True
        .Visible = True
        
        Set cmdBar1 = .Controls.Add(msoControlButton, 694)
        cmdBar1.ToolTipText = "Auto Indent Code"
        
        Set cmdBar2 = .Controls.Add(msoControlButton, 3157)
        cmdBar2.ToolTipText = "Procedure List"
    End With
    
    Set Me.MenuHandler1 = VBInstance.Events.CommandBarEvents(cmdBar1)
    Set Me.MenuHandler2 = VBInstance.Events.CommandBarEvents(cmdBar2)
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    If Not (mcbMenuCommandBar Is Nothing) Then mcbMenuCommandBar.Delete
    Set objSortWindow = Nothing
End Sub

Private Sub CmpHandler_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
    Debug.Print "CmpHandler_ItemAdded: " & VBComponent.Name
End Sub

Private Sub MenuHandler1_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If Not (VBInstance.ActiveCodePane Is Nothing) Then
        IndentInitialize
        
        IndentCodeModule VBInstance.ActiveCodePane.CodeModule
    Else
        MsgBox "It would help if you actually open a code module...", vbInformation
    End If
End Sub

Private Sub MenuHandler2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim docSP As Object
    Dim udSP As udSortProcedures
    Const guidMYTOOL As String = "{CDA313D0-AFE0-01A3-B621-591AF24708E1}"
    
    If Not (objSortWindow Is Nothing) Then
        objSortWindow.Visible = True
    Else
        Set objSortWindow = VBInstance.Windows.CreateToolWindow(Me.VBInstance.Addins("VBTools.Connect"), "VBTools.udSortProcedures", "Procedure List", guidMYTOOL, docSP)
        
        Set udSP = docSP
        Set udSP.MyWindow = objSortWindow
        Set udSP.VBInstance = Me.VBInstance
        
        objSortWindow.Visible = True
    End If
End Sub
