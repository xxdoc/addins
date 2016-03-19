VERSION 5.00
Begin VB.Form fMultiline 
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3390
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H00C0FFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txMultiline 
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   2730
      Left            =   53
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   330
      Width           =   7155
   End
   Begin VB.Label lbLines 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Line(s) -- max 25"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   390
      TabIndex        =   9
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Label lbLineCnt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   150
      TabIndex        =   8
      Top             =   3120
      Width           =   180
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
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
      Height          =   195
      Index           =   4
      Left            =   6495
      TabIndex        =   7
      Top             =   3105
      Width           =   405
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Paste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   3
      Left            =   5655
      TabIndex        =   6
      ToolTipText     =   "Paste copied code into textbox and un-convert"
      Top             =   3105
      Width           =   495
   End
   Begin VB.Label lbTit 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(press Shift+Return when done)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A07070&
      Height          =   195
      Index           =   2
      Left            =   4260
      TabIndex        =   5
      Top             =   45
      Width           =   2745
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   4725
      TabIndex        =   4
      ToolTipText     =   "Cancel"
      Top             =   3105
      Width           =   600
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   1
      Left            =   3900
      TabIndex        =   3
      ToolTipText     =   "Convert and insert into code at the caret position"
      Top             =   3105
      Width           =   480
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00707070&
      Height          =   195
      Index           =   0
      Left            =   2880
      TabIndex        =   2
      ToolTipText     =   "Convert into VB compliant form and send to a message box"
      Top             =   3105
      Width           =   690
   End
   Begin VBCompanion.Fader Fader3 
      Left            =   1860
      Top             =   660
      _ExtentX        =   1164
      _ExtentY        =   450
      FadeInSpeed     =   8
      FadeOutSpeed    =   8
   End
   Begin VB.Label lbTit 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Multi-Line Literal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "fMultiline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Textbox to enter Multiline VB text literals
'based on an idea by Evan Toder

Option Explicit

Private LastIndex       As Long
Private Const ContMark  As String = " _" & vbCrLf
Private sInset          As String
Private sContinue       As String
Private Pasted          As Boolean
Private Const HelpText  As String = "Enter your text as you would like it to appear when the application runs." & vbCrLf & _
                                    vbCrLf & _
                                    """Preview"" lets you see what your text looks like in a message box." & vbCrLf & _
                                    vbCrLf & _
                                    """Apply"" will format your text into a VB-compliant form and then paste it into your code" & vbCrLf & _
                                    "at the caret position (currently indicated by the black arrow)." & vbCrLf & _
                                    vbCrLf & _
                                    """Cancel"" abandons the operation and leaves your code unaltered." & vbCrLf & _
                                    vbCrLf & _
                                    "And finally ""Paste"": " & vbCrLf & _
                                    vbCrLf & _
                                    "1 Mark any literal or expression in your code and copy it to the clipboard. You can't do" & vbCrLf & _
                                    "   that right now, but you will once you have closed this box." & vbCrLf & _
                                    vbCrLf & _
                                    "2 Re-open this box and click Paste. Your copied text is stripped of all VB-relevant stuff" & vbCrLf & _
                                    "   and is inserted into this box." & vbCrLf & _
                                    vbCrLf & _
                                    "   If the pasted text is an expression it will be evaluated and the result shown, unless" & vbCrLf & _
                                    "   there is an error in the expression or elsewhere in your code." & vbCrLf & _
                                    vbCrLf & _
                                    "   Note: Pasted text is limited to approximately 800 characters after stripping."

Private HelpToolTip     As cToolTip

Private Function Convert(TextToConvert As String) As String

  'format text into VB compliant form

    If Len(TextToConvert) Then
        Convert = Replace$(Replace$(Replace$("""" & Replace$(Replace$(Replace$(TextToConvert, """", Chr$(160)), vbCrLf, sContinue), """"" & ", vbNullString) & """", " & _" & vbCrLf & sInset & """""", vbNullString), """"" & ", vbNullString), Chr$(160), """""")
    End If

End Function

Private Sub Form_Activate()

    RedrawArrow
    Fader3.FadeIn FadeFast

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetCursor

End Sub

Private Sub GoBack()

  'exit fMultiline

    CaretRgn Arrow Or Inval 'invalidate additional arrow region before hiding Me
    Fader3.FadeOut FadeFast
    Hide

End Sub

Friend Property Let Indent(ToCol As Long)

    sInset = Space$(ToCol - 1)
    sContinue = """ & vbCrlf &" & ContMark & sInset & """" 'closing quote - runtime line feed - space.underline.IDE line feed - inset - opening quote

End Property

Private Sub lb_Click(Index As Integer)

    Select Case Index
      Case 0
        Preview
      Case 1
        Tag = Convert(txMultiline) 'to vb compliant form
        GoBack 'exit
      Case 2
        Tag = vbNullString
        GoBack
      Case 3
        Paste
      Case 4
        Set HelpToolTip = Nothing
        Set HelpToolTip = New cToolTip
        HelpToolTip.Create lb(4), HelpText, TTBalloonIfActive, False, TTIconInfo, "Help", vbBlack, &HFFFFF0, 50, 30000
    End Select

End Sub

Private Sub lb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    lb(LastIndex).FontUnderline = False
    lb(Index).FontUnderline = True
    LastIndex = Index
    SetCursor LoadCursor(0, CursorRightHand)

End Sub

Private Sub Paste()

  Dim tmpText   As String

    With Clipboard
        If .GetFormat(vbCFText) Then
            tmpText = .GetText
            'clipboard text may have line continuations in it - so we strip them
            'here because ExecuteVBCode expects a single un-continued line, and then
            'let VB untangle what we have
            ExecuteVBCode "Clipboard.SetText " & Strip(tmpText)
            RedrawArrow
        End If
        Pasted = True
        txMultiline = .GetText
        .SetText tmpText
    End With 'CLIPBOARD

End Sub

Private Sub Preview()

    lb(0).FontUnderline = False
    'Convert may have line continuations in it - so we strip them
    'here because ExecuteVBCode expects a single un-continued line
    ExecuteVBCode "MsgBox " & Strip(Convert(txMultiline)) & ",vbInformation , ""Preview"""
    RedrawArrow

End Sub

Private Sub ResetCursor()

    If lb(LastIndex).FontUnderline Then
        lb(LastIndex).FontUnderline = False
        SetCursor LoadCursor(0, CursorArrow)
        Set HelpToolTip = Nothing
    End If

End Sub

Private Function Strip(TextToStrip As String) As String

  'strips trailing empty lines and contmarks from text and then leading spaces from each individual line

  Dim Lines()   As String
  Dim i         As Long
  Dim sTmp      As String

    sTmp = TextToStrip
    Do
        sTmp = RTrim$(sTmp)
        If Right$(sTmp, 2) = vbCrLf Then
            sTmp = Left$(sTmp, Len(sTmp) - 2)
          Else 'NOT RIGHT$(STMP,...
            Exit Do 'loop 
        End If
    Loop
    If Right$(sTmp, 3) = "& _" Then
        sTmp = Left$(sTmp, Len(sTmp) - 3)
    End If
    Lines = Split(sTmp, ContMark)
    For i = 0 To UBound(Lines)
        Lines(i) = LTrim$(Lines(i))
    Next i
    Strip = Join$(Lines)

End Function

Private Sub txMultiline_Change()

  'limit text to 25 lines because thats the maximum VB will accept

  Dim lc As Long

    lc = SendMessage(txMultiline.hWnd, EM_GETLINECOUNT, 0&, ByVal 0&)
    lbLineCnt = lc
    If lc > 25 Then
        If Pasted Then
            Pasted = False
            txMultiline = "[Can't paste - too many Lines]"
          Else 'PASTED = FALSE/0
            SendMessage txMultiline.hWnd, EM_UNDO, 0&, ByVal 0&
            Beeper 2000, 20
            txMultiline.SelStart = Len(txMultiline)
        End If
        lbLineCnt = SendMessage(txMultiline.hWnd, EM_GETLINECOUNT, 0&, ByVal 0&)
      Else 'NOT LC...
        SendMessage txMultiline.hWnd, EM_EMPTYUNDOBUFFER, 0&, ByVal 0&
    End If

End Sub

Private Sub txMultiline_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
      Case vbKeyReturn
        If Shift And vbShiftMask Then 'shift+return key simultaneously
            KeyCode = 0 'kill keycode
            lb_Click 1
        End If
      Case vbKeyEscape
        lb_Click 2 'exit without
      Case vbKeyPause
        Preview
    End Select

End Sub

Private Sub txMultiline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetCursor

End Sub

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 16:14)  Decl: 33  Code: 187  Total: 220 Lines
':) CommentOnly: 11 (5%)  Commented: 11 (5%)  Empty: 52 (23,6%)  Max Logic Depth: 3
