VERSION 5.00
Begin VB.Form fIcon 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3870
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "fIcon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fIcon.frx":08CA
   ScaleHeight     =   40
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   258
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.PictureBox picMenuCopy 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2565
      Picture         =   "fIcon.frx":1BCC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   180
      Width           =   240
   End
   Begin VB.PictureBox picAa 
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   3075
      Picture         =   "fIcon.frx":1F0E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   75
      Width           =   480
   End
   Begin VB.PictureBox picMenuCompare 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2220
      Picture         =   "fIcon.frx":2BDC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   195
      Width           =   240
   End
   Begin VB.PictureBox picMenuResetRed 
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   1590
      Picture         =   "fIcon.frx":2F1E
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   165
      Width           =   240
   End
   Begin VB.PictureBox picMenuOpenAll 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1905
      Picture         =   "fIcon.frx":3260
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   195
      Width           =   240
   End
   Begin VB.PictureBox picMenuResetGreen 
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   1290
      Picture         =   "fIcon.frx":35A2
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   165
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'Kein
      Height          =   600
      Left            =   0
      Picture         =   "fIcon.frx":38E4
      ScaleHeight     =   600
      ScaleWidth      =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "fIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This form has no code

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 16:14)  Decl: 6  Code: 0  Total: 6 Lines
':) CommentOnly: 1 (16,7%)  Commented: 0 (0%)  Empty: 1 (16,7%)  Max Logic Depth: 0
