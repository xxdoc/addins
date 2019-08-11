VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11715
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   21540
   _ExtentX        =   37994
   _ExtentY        =   20664
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "Document Map Add-In"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
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

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Const CommandBarTitle = "Document Map Window"

Const guidTasks$ = "AB3075C1-B54F-11d3-941A-00A0CC547B23"
Dim gwinWindow As VBIDE.Window
Dim mUserDoc As UserDoc

Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    gwinWindow.Visible = False
   
End Sub

Sub Show()
    On Error Resume Next
    
    If (gwinWindow Is Nothing) Then
        Set gwinWindow = VBInstance.Windows.CreateToolWindow(Me.VBInstance.Addins("DocumentMapAddIn.Connect"), "DocumentMapAddIn.UserDoc", "Document Map", "{CDA313D0-AFE0-01A3-B621-591AF24708E1}", mUserDoc)
        If Not (gwinWindow Is Nothing) Then
            Set mUserDoc.VBInstance = VBInstance
            Set mUserDoc.Connect = Me
            mUserDoc.UserDocumentLoad
            gwinWindow.Visible = True
        End If
    Else
        gwinWindow.Visible = True
    End If
    
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'Save the vb instance
    Set VBInstance = Application
    
    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show

    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar(CommandBarTitle)
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
    On Error Resume Next
    
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    MsgBox Err.Description
End Sub

'------------------------------------------------------
'This method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload gwinWindow
    Set gwinWindow = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
Dim cbMenu As Object

On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    Set cbMenuCommandBar = VBInstance.CommandBars.FindControl(msoControlButton, , CommandBarTitle)
    If (cbMenuCommandBar Is Nothing) Then
        'add it to the command bar
        Set cbMenuCommandBar = cbMenu.Controls.Add(1)
        'set the caption
        cbMenuCommandBar.Caption = sCaption
        cbMenuCommandBar.Tag = sCaption
    End If
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:
End Function

