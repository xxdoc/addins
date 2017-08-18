VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mi complemento"
   ClientHeight    =   3204
   ClientLeft      =   2172
   ClientTop       =   1932
   ClientWidth     =   6036
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
   ScaleHeight     =   3204
   ScaleWidth      =   6036
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelectedControls 
      Caption         =   "Show selected controls"
      Height          =   372
      Left            =   360
      TabIndex        =   1
      Top             =   252
      Width           =   2424
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Close"
      Height          =   375
      Left            =   4356
      TabIndex        =   0
      Top             =   216
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub cmdSelectedControls_Click()
    frmSelelectedControls.Show
End Sub

