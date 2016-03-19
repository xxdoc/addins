VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fSelect 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCount 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   225
      Left            =   2280
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   2
      Top             =   270
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid fgdMembers 
      Height          =   1935
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   9
      Cols            =   4
      FixedCols       =   0
      BackColor       =   15794175
      ForeColor       =   4210752
      BackColorFixed  =   13697023
      ForeColorFixed  =   128
      ForeColorSel    =   0
      BackColorBkg    =   15794175
      GridColorFixed  =   12632256
      ScrollTrack     =   -1  'True
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      MousePointer    =   99
      FormatString    =   " Name                                     |Declared in                           |Type         |Scope   "
      MouseIcon       =   "fSelect.frx":0000
   End
   Begin VB.PictureBox picExit 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   30
      Picture         =   "fSelect.frx":08DA
      ScaleHeight     =   240
      ScaleWidth      =   1695
      TabIndex        =   1
      Top             =   255
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VBCompanion.Fader Fader2 
      Left            =   2580
      Top             =   825
      _ExtentX        =   1164
      _ExtentY        =   450
      FadeInSpeed     =   8
      FadeOutSpeed    =   8
   End
End
Attribute VB_Name = "fSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VB member selector popup form

Option Explicit

Private CompoName       As String

Friend Property Let CurrentComponentName(Name As String)

  'copy all applicable member names to flexgrid and position window

    CompoName = Name
    With fgdMembers
        'prepare
        .Rows = 2  'min 2 rows
        .Row = 0 'Headline
        .Col = 0 'col 0 is default throughout
        .CellFontBold = True
        .Row = 1 'this is the "Exit"line
        .CellBackColor = &HD0D0D0 'light gray
        .CellAlignment = flexAlignCenterCenter
        .RowHeight(.Row) = picExit.Height * Screen.TwipsPerPixelY
        Set .CellPicture = picExit
        'fill it
        For Each Item In CodeMembers 'Item is a Variant of Type Array
            If Item(idxMemberScope) <> IIf(Item(idxCompoName) = Name, 0, vbext_Private) Then
                .Rows = .Rows + 1 'needs more rows
                .Row = .Row + 1 'next row
                .CellFontBold = True 'member name bold
                .TextMatrix(.Row, 0) = Item(idxMembername)
                .TextMatrix(.Row, 1) = Item(idxCompoName)
                If Item(idxMemberScope) = NoScope Then
                    .TextMatrix(.Row, 2) = Choose(Item(idxMemberType), "Module", "Class", "Form", vbNullString, "Form", "MDIForm", "PropPage", "Control", "Document", vbNullString, "Designer")
                    .CellForeColor = &HA00000
                  Else 'NOT Item(IDXMEMBERSCOPE)...
                    .TextMatrix(.Row, 2) = Choose(Item(idxMemberType), "Method", "Property", "Variable", "Event", "Constant")
                    .CellForeColor = Choose(Item(idxMemberType), &H8000&, &H6060&, &H700070, &H705000, &H80&)
                End If
                .TextMatrix(.Row, 3) = Choose(Item(idxMemberScope), "Private", "Public", "Friend", vbNullString)
            End If
        Next Item
        'select top line
        .Row = 1
        'sort it
        .Col = 0
        .ColSel = 0
        .Sort = flexSortStringNoCaseAscending
        'adjust to fit
        If .Rows <= 8 Then
            .Height = .Rows * .RowHeight(1) / Screen.TwipsPerPixelY + 1
            .Width = .Width - 16
            Width = (.Width + 4) * Screen.TwipsPerPixelX
        End If
        Height = (.Height + 4) * Screen.TwipsPerPixelY
    End With 'fgdMembers

End Property

Private Sub fgdMembers_Click()

    With fgdMembers
        If .MouseRow = 0 Then
            .Col = .MouseCol
            .ColSel = .MouseCol
            .Sort = flexSortStringNoCaseAscending
        End If
    End With 'FGDMEMBERS

End Sub

Private Sub fgdMembers_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
      Case vbKeyEscape, vbKeyPause
        Tag = vbNullString
        GoBack
      Case vbKeyReturn, vbKeySpace 'same as a Click or rather MouseUp
        fgdMembers_MouseUp 0, 0, 0, 0
    End Select

End Sub

Private Sub fgdMembers_KeyPress(KeyAscii As Integer)

  Dim RowNum        As Long

    Select Case KeyAscii
      Case vbKeyEscape, vbKeySpace, vbKeyReturn
        'do nothing
      Case Else
        With fgdMembers
            .Col = 0
            .ColSel = 0
            .Sort = flexSortStringNoCaseAscending
            For RowNum = 2 To .Rows - 1
                If LCase$(Left$(.TextMatrix(RowNum, 0), 1)) = LCase$(Chr$(KeyAscii)) Then
                    .Row = RowNum
                    .TopRow = RowNum
                    Exit For 'loop varying rownum
                End If
            Next RowNum
            If RowNum = .Rows Then
                Beeper 4000, 20
            End If
        End With 'fgdMembers
    End Select

End Sub

Private Sub fgdMembers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With fgdMembers
        Select Case .MouseRow
          Case 0
            .MousePointer = flexCustom
          Case Else
            .MousePointer = flexDefault
        End Select
    End With 'FGDMEMBERS

End Sub

Private Sub fgdMembers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With fgdMembers
        If .MouseRow <> 0 Then
            If .MouseRow = 1 Or Button = vbRightButton Then
                Tag = vbNullString
              Else 'NOT .MOUSEROW...
                Tag = .TextMatrix(.Row, 0)
            End If
            GoBack
        End If
    End With 'fgdMembers

End Sub

Private Sub fgdMembers_Scroll()

    picCount.Visible = (fgdMembers.TopRow = 1)

End Sub

Private Sub Form_Activate()

    Fader2.FadeIn FadeMedium

End Sub

Private Sub GoBack()

  'exit fSelect

    CaretRgn Arrow Or Inval 'invalidate additional arrow region before hiding Me
    Fader2.FadeOut FadeMedium
    Hide

End Sub

Private Sub picCount_Click()

    Tag = vbNullString
    GoBack

End Sub

Private Sub picCount_Paint()

  'display member count

    With picCount
        .Cls
        .CurrentY = 1
        .ForeColor = &H9090&
        picCount.Print "(";
        .ForeColor = &H90&
        picCount.Print CompoName;
        .ForeColor = &H9090&
        picCount.Print " can see";
        .ForeColor = &H900000
        With fgdMembers
            picCount.Print .Rows - 2;
            picCount.ForeColor = &H9090&
            picCount.Print "member" & IIf(.Rows = 3, ")", "s)")
        End With 'FGDMEMBERS
    End With 'PICCOUNT

End Sub

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 16:14)  Decl: 5  Code: 185  Total: 190 Lines
':) CommentOnly: 10 (5,3%)  Commented: 20 (10,5%)  Empty: 39 (20,5%)  Max Logic Depth: 5
