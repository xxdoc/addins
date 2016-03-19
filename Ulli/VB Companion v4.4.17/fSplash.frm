VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4620
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog cdlDB 
      Left            =   1140
      Top             =   -15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Ulli's VB Companion - Need Path to API Database"
      FileName        =   "Win32Api.mdb"
      Filter          =   "Microsoft Access Databases (*.mdb)|*.MDB"
      FilterIndex     =   1
   End
   Begin VB.Image img 
      BorderStyle     =   1  'Fest Einfach
      Height          =   765
      Left            =   195
      Picture         =   "fSplash.frx":000C
      Top             =   188
      Width           =   825
   End
   Begin VB.Label lblAbout 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading VB Companion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1230
      TabIndex        =   0
      Top             =   450
      Width           =   2445
   End
End
Attribute VB_Name = "fSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This form has no code

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 16:14)  Decl: 6  Code: 0  Total: 6 Lines
':) CommentOnly: 1 (16,7%)  Commented: 0 (0%)  Empty: 1 (16,7%)  Max Logic Depth: 0
