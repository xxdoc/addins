VERSION 5.00
Begin VB.Form frmIPC 
   Caption         =   "Form1"
   ClientHeight    =   555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2115
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleWidth      =   2115
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIPCServer 
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   1860
   End
End
Attribute VB_Name = "frmIPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtIPCServer_Change()
    If Len(txtIPCServer) = 0 Then Exit Sub
    'MsgBox txtIPCServer
    IPCCommand txtIPCServer
    txtIPCServer = Empty
End Sub
