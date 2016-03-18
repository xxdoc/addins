VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Add In"
   ClientHeight    =   3195
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmAddIn.frx":0000
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
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



Private Sub OKButton_Click()
    Me.Hide
End Sub

