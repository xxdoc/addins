VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fSetOpts 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "x"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4350
   ControlBox      =   0   'False
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fSetOpts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fr 
      Height          =   360
      Left            =   3585
      TabIndex        =   37
      Top             =   90
      Width           =   600
      Begin VB.Label lbAbout 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " About "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   165
         Left            =   45
         TabIndex        =   38
         Top             =   135
         Width           =   450
      End
   End
   Begin VB.Frame frMail 
      BackColor       =   &H00E0E0E0&
      Height          =   720
      Left            =   3015
      TabIndex        =   32
      Top             =   555
      Width           =   1170
      Begin VB.TextBox txAuthor 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   330
         Left            =   45
         Locked          =   -1  'True
         MouseIcon       =   "fSetOpts.frx":000C
         MousePointer    =   99  'Custom
         MultiLine       =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "fSetOpts.frx":08D6
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame frRaster 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Raster"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1065
      Left            =   180
      TabIndex        =   30
      Top             =   2940
      Width           =   2670
      Begin VB.CheckBox ckROP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Over&write"
         ForeColor       =   &H00800080&
         Height          =   330
         Left            =   1785
         TabIndex        =   14
         ToolTipText     =   "MaskPen"
         Top             =   645
         Width           =   705
      End
      Begin VB.Frame frRasterOpts 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   105
         TabIndex        =   39
         Top             =   240
         Width           =   1065
         Begin VB.OptionButton opRaster 
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   " None "
            Top             =   15
            Width           =   180
         End
         Begin VB.OptionButton opRaster 
            Height          =   195
            Index           =   1
            Left            =   285
            TabIndex        =   8
            ToolTipText     =   " Vertical "
            Top             =   15
            Width           =   180
         End
         Begin VB.OptionButton opRaster 
            Height          =   195
            Index           =   2
            Left            =   585
            TabIndex        =   9
            ToolTipText     =   " Horizontal "
            Top             =   15
            Width           =   180
         End
         Begin VB.OptionButton opRaster 
            Height          =   195
            Index           =   3
            Left            =   870
            TabIndex        =   10
            ToolTipText     =   " Both "
            Top             =   15
            Width           =   180
         End
      End
      Begin VB.OptionButton opDotted 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Dotted"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   705
         Width           =   795
      End
      Begin VB.OptionButton opSolid 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Soli&d"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   975
         TabIndex        =   13
         Top             =   705
         Width           =   645
      End
      Begin VB.CheckBox btColor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Color"
         ForeColor       =   &H00800080&
         Height          =   315
         Left            =   1365
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   660
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "N    V    H    B"
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   36
         Top             =   435
         Width           =   990
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   2115
         Top             =   165
         Width           =   465
      End
   End
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   2895
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Color           =   15790320
   End
   Begin VB.Frame frButtons 
      BackColor       =   &H00E0E0E0&
      Height          =   3975
      Left            =   3015
      TabIndex        =   33
      Top             =   1350
      Width           =   1170
      Begin VB.CommandButton btCAO 
         Caption         =   "&Refresh Reset"
         Height          =   465
         Index           =   4
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Reset settings and close box"
         Top             =   2553
         Width           =   915
      End
      Begin VB.CommandButton btCAO 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         CausesValidation=   0   'False
         Height          =   465
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   225
         Width           =   915
      End
      Begin VB.CommandButton btCAO 
         Caption         =   "&Apply"
         Height          =   465
         Index           =   1
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Apply settings"
         Top             =   1777
         Width           =   915
      End
      Begin VB.CommandButton btCAO 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   465
         Index           =   2
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Apply settings and close box"
         Top             =   3330
         Width           =   915
      End
      Begin VB.CommandButton btCAO 
         Caption         =   "&Save"
         Height          =   465
         Index           =   3
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Apply and save settings"
         Top             =   1001
         Width           =   915
      End
   End
   Begin VB.Frame frAutoComplete 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto Complete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00006000&
      Height          =   1275
      Left            =   180
      TabIndex        =   31
      Top             =   4050
      Width           =   2670
      Begin VB.CheckBox ckUnique 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Unique"
         ForeColor       =   &H00006000&
         Height          =   315
         Left            =   105
         TabIndex        =   17
         Top             =   570
         Width           =   825
      End
      Begin VB.TextBox txtTriggerLength 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   2310
         MaxLength       =   1
         TabIndex        =   19
         Top             =   570
         Width           =   195
      End
      Begin VB.CheckBox ckNoisy 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Noisy"
         ForeColor       =   &H00006000&
         Height          =   195
         Left            =   105
         TabIndex        =   20
         Top             =   960
         Width           =   705
      End
      Begin VB.OptionButton opAcOn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Acti&ve"
         ForeColor       =   &H00006000&
         Height          =   210
         Left            =   90
         TabIndex        =   15
         Top             =   300
         Width           =   765
      End
      Begin VB.OptionButton opAcOff 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Inacti&ve"
         ForeColor       =   &H00006000&
         Height          =   195
         Left            =   1605
         TabIndex        =   16
         Top             =   300
         Width           =   885
      End
      Begin VB.Label lb 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "&Trigger Length"
         ForeColor       =   &H00006000&
         Height          =   195
         Index           =   1
         Left            =   1155
         TabIndex        =   18
         Top             =   615
         Width           =   1080
      End
   End
   Begin VB.Frame frScroll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lines to Scroll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1230
      Left            =   180
      TabIndex        =   27
      Top             =   570
      Width           =   2670
      Begin VB.OptionButton opPage 
         BackColor       =   &H00E0E0E0&
         Caption         =   "W&hole Page"
         CausesValidation=   0   'False
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   930
         Width           =   1200
      End
      Begin VB.OptionButton opHalfPage 
         BackColor       =   &H00E0E0E0&
         Caption         =   "[set at runtime]"
         CausesValidation=   0   'False
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   615
         Width           =   2535
      End
      Begin VB.OptionButton opAbsValue 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter Number of &Lines"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   90
         TabIndex        =   0
         Top             =   300
         Width           =   1935
      End
      Begin VB.TextBox txLines 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2205
         MaxLength       =   2
         TabIndex        =   1
         ToolTipText     =   "1 thru 99"
         Top             =   270
         Width           =   285
      End
   End
   Begin VB.Frame frSmooth 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Smooth Scrolling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1020
      Left            =   180
      TabIndex        =   28
      Top             =   1860
      Width           =   2670
      Begin VB.HScrollBar scrSpeed 
         Height          =   240
         LargeChange     =   50
         Left            =   705
         Max             =   20
         Min             =   200
         SmallChange     =   10
         TabIndex        =   6
         Top             =   630
         Value           =   20
         Width           =   1800
      End
      Begin VB.OptionButton opScOff 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Off"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1965
         TabIndex        =   5
         Top             =   300
         Width           =   510
      End
      Begin VB.OptionButton opScOn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&On"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   90
         TabIndex        =   4
         Top             =   300
         Width           =   525
      End
      Begin VB.Label lb 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Sp&eed"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   645
         Width           =   465
      End
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   30
      Picture         =   "fSetOpts.frx":08F3
      Top             =   30
      Width           =   480
   End
   Begin VBCompanion.Fader Fader1 
      Left            =   2205
      Top             =   135
      _ExtentX        =   1164
      _ExtentY        =   450
      FadeInSpeed     =   8
      FadeOutSpeed    =   8
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Set Options:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606000&
      Height          =   330
      Index           =   0
      Left            =   555
      TabIndex        =   26
      Top             =   90
      Width           =   1455
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Set Options:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0FFFF&
      Height          =   330
      Index           =   3
      Left            =   570
      TabIndex        =   34
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "fSetOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Options Form

Option Explicit

Private Tooltip                 As New cToolTip
Private PrevFocus               As Long

Private Enum ButtonIndexes
    idxCancel = 0
    idxApply = 1
    idxOK = 2
    idxSave = 3
    idxRefresh = 4
End Enum
#If False Then ':) Line inserted by Formatter
Private idxCancel, idxApply, idxOK, idxSave, idxRefresh ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Private Sub btCAO_Click(Index As Integer)

  'this works by combining several clicks recursively thru itself

    Select Case Index
      Case idxCancel
        Fader1.FadeOut FadeVeryFast
        Hide
      Case idxApply
        Select Case True
          Case opAbsValue
            NumLinesToScroll = txLines
          Case opPage
            NumLinesToScroll = sFullPage
          Case opHalfPage
            NumLinesToScroll = sHalfPage
        End Select
        Smooth = opScOn
        ScrollDelay = scrSpeed / SSScale
        AutoComplete = opAcOn
        Noisy = (ckNoisy = vbChecked)
        LineStyle = IIf(opSolid, PS_SOLID, PS_DOT)
        GridColor = shpColor.FillColor
        RasDrwMode = IIf(ckROP = vbChecked, vbCopyPen, vbMaskPen)
        Pen DestroyItCreateIt
        EraseRaster True
        If Raster Then
            DrawRaster
        End If
        TriggerLength = Val(txtTriggerLength)
        NonUnique = (ckUnique = vbUnchecked) 'saving the opposite
        LoadKeywords IIf(NonUnique, TriggerLength, 2)
      Case idxOK
        btCAO_Click idxApply
        btCAO_Click idxCancel
      Case idxSave
        btCAO_Click idxApply
        With App 'save settings
            SaveSetting .Title, sOptions, sLinesToScroll, NumLinesToScroll
            SaveSetting .Title, sOptions, sMode, IIf(Smooth, sSmooth, sInstant)
            SaveSetting .Title, sOptions, sSpeed, Format$(scrSpeed)
            SaveSetting .Title, sOptions, sAutoComplete, IIf(AutoComplete, sOn, sOff)
            SaveSetting .Title, sOptions, sNoisy, IIf(Noisy, sOn, sOff)
            SaveSetting .Title, sOptions, sROP, IIf(RasDrwMode = vbMaskPen, sMaskPen, sCopyPen)
            SaveSetting .Title, sOptions, sRaster, Choose(Raster + 1, sNone, sVertical, sHorizontal, sBoth)
            SaveSetting .Title, sOptions, sRasterColor, Hex$(GridColor)
            SaveSetting .Title, sOptions, sRasterStyle, IIf(LineStyle = PS_DOT, sDotted, sSolid)
            SaveSetting .Title, sOptions, sTriggerLength, txtTriggerLength
            SaveSetting .Title, sOptions, sUnique, IIf(NonUnique, sOff, sOn)
        End With 'APP
      Case idxRefresh
        GetRegistrySettings
        btCAO_Click idxCancel
    End Select

End Sub

Private Sub btColor_Click()

    If btColor = vbChecked Then
        On Error Resume Next
            With cdlColor
                .Flags = cdlCCFullOpen Or cdlCCRGBInit
                .Color = shpColor.FillColor
                .ShowColor
                If Err = 0 Then
                    shpColor.FillColor = .Color
                End If
            End With 'CDLCOLOR
        On Error GoTo 0
        btColor = vbUnchecked
    End If

End Sub

Private Sub ckROP_Click()

    If ckROP = vbChecked Then
        ckROP.ToolTipText = "CopyPen"
      Else 'NOT CKROP...
        ckROP.ToolTipText = "MaskPen"
    End If

End Sub

Private Sub ckUnique_Click()

    opAcOff_Click 'en-/disable some controls

End Sub

Private Sub Form_Activate()

    Tooltip.Create txAuthor, "Click here to send mail to author", TTBalloonIfActive, False, TTIconInfo, "Standard mail program", , , 250, 10000
    Fader1.FadeIn FadeVeryFast

End Sub

Private Sub Form_Load()

  Const Margin          As Long = 5 'pixels - prevent Me from being placed directly at the screen borders
  Dim MarginX           As Long
  Dim MarginY           As Long

    GetCursorPos CursorPos 'where's the mouse cursor
    With CursorPos
        .X = .X * Screen.TwipsPerPixelX - Width / 2  'adjust to twips and also reflect my dimensions
        .Y = .Y * Screen.TwipsPerPixelY - Height / 2
        MarginX = Margin * Screen.TwipsPerPixelX
        MarginY = Margin * Screen.TwipsPerPixelY
        Select Case True 'limit x to be within screen
          Case .X < MarginX
            .X = MarginX
          Case .X + Width > Screen.Width - MarginX
            .X = Screen.Width - Width - MarginX
        End Select
        Select Case True 'limit y to be within screen
          Case .Y < MarginY
            .Y = MarginY
          Case .Y + Height > Screen.Height - MarginY
            .Y = Screen.Height - Height - MarginY
        End Select
        Move .X, .Y 'move Me to that position
    End With 'CURSORPOS

    'preset initial captions and values
    With App
        Caption = .Title & " V" & .Major & "." & .Minor & "." & .Revision
    End With 'APP
    opHalfPage.Caption = opHpCapt
    Select Case NumLinesToScroll
      Case sFullPage
        opPage = True
      Case sHalfPage
        opHalfPage = True
      Case Else
        opAbsValue = True
        txLines = Abs(Val(NumLinesToScroll)) Mod 30
    End Select
    opScOn = Smooth
    opScOff = (Smooth = False)
    lb(2).Enabled = Smooth
    scrSpeed.Enabled = Smooth
    On Error Resume Next
        scrSpeed = ScrollDelay * SSScale
    On Error GoTo 0
    opAcOn = AutoComplete
    opAcOff = (AutoComplete = False)
    txtTriggerLength = TriggerLength
    txtTriggerLength_LostFocus 'presets the validation rules
    ckUnique = IIf(NonUnique, vbUnchecked, vbChecked)
    ckNoisy = IIf(Noisy, vbChecked, vbUnchecked)
    ckROP = IIf(RasDrwMode = vbCopyPen, vbChecked, vbUnchecked)
    shpColor.FillColor = GridColor
    opSolid = (LineStyle = PS_SOLID)
    opDotted = (LineStyle = PS_DOT)
    opRaster(Raster And 3) = True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        ReleaseCapture  'release the Mouse 'NOT RIGHT$(STMP,...
        SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0& 'non-client area button down (in caption)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PrevFocus = 0
    opAcOff.CausesValidation = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Tooltip = Nothing

End Sub

Private Sub frAutoComplete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub frButtons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub frMail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub frRaster_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub frScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub frSmooth_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub lb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub lbAbout_Click()

    With App
        ShellAbout hWnd, "About " & AppDetails & "#Operating System:", .Title & " V" & .Major & "." & .Minor & "." & .Revision & vbCrLf & .LegalCopyright, fIcon.Icon.Handle
    End With 'APP

End Sub

Private Sub opAbsValue_Click()

  'user wishes to input number of lines to scroll

    With txLines
        .Enabled = opAbsValue
        .TabStop = opAbsValue
        If opAbsValue Then
            .SelStart = 0
            .SelLength = 2
            On Error Resume Next 'this may be called during form load when we cannot set focus
                .SetFocus
            On Error GoTo 0
        End If
    End With 'TXLINES

End Sub

Private Sub opAcOff_Click()

    txtTriggerLength.Enabled = opAcOn And (ckUnique = vbUnchecked)
    ckNoisy.Enabled = opAcOn
    lb(1).Enabled = opAcOn And (ckUnique = vbUnchecked)
    ckUnique.Enabled = opAcOn

End Sub

Private Sub opAcOff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If GetFocus = txtTriggerLength.hWnd Then
        opAcOff.CausesValidation = False
    End If

End Sub

Private Sub opAcOn_Click()

    opAcOff_Click
    If opAcOn Then
        If Len(txtTriggerLength) = 0 Then
            On Error Resume Next
                txtTriggerLength.SetFocus
            On Error GoTo 0
        End If
    End If

End Sub

Private Sub opHalfPage_Click()

    opAbsValue_Click
    txLines = vbNullString

End Sub

Private Sub opPage_Click()

    opAbsValue_Click
    txLines = vbNullString

End Sub

Private Sub opRaster_Click(Index As Integer)

    Raster = Index
    Select Case Index
      Case 0
        shpColor.FillStyle = vbFSTransparent
      Case 1
        shpColor.FillStyle = vbVerticalLine
      Case 2
        shpColor.FillStyle = vbHorizontalLine
      Case 3
        shpColor.FillStyle = vbCross
    End Select
    If Index <> 0 Then
        'measure font
        With fIcon
            .FontName = IDEFontName
            .FontSize = IDEFontSize
            If .TextWidth("I") <> .TextWidth("W") Then
                MsgBoxEx "You cannot use the raster option with variable pitch fonts like " & IDEFontName & ".", vbExclamation, Sound:=-440.02
                opRaster(0) = True
            End If
        End With 'FICON
    End If
    btColor.Enabled = Raster
    opDotted.Enabled = Raster
    opSolid.Enabled = Raster
    ckROP.Enabled = Raster

End Sub

Private Sub opScOff_Click()

    scrSpeed.Enabled = opScOn
    lb(2).Enabled = opScOn

End Sub

Private Sub opScOn_Click()

    opScOff_Click

End Sub

Private Sub scrSpeed_Change()

    lb(2).ToolTipText = "Approx " & Int(5000 / scrSpeed) & " Lines per Second"

End Sub

Private Sub scrSpeed_Scroll()

    scrSpeed_Change

End Sub

Private Sub txAuthor_Click()

    With App
        txAuthor.SelStart = 0
        PutFocus PrevFocus
        SendMeMail hWnd, .ProductName & " V" & .Major & "." & .Minor & "." & .Revision
    End With 'APP

End Sub

Private Sub txAuthor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If PrevFocus = 0 Then
        PrevFocus = GetFocus
    End If

End Sub

Private Sub txLines_KeyPress(KeyAscii As Integer)

    If InStr("0123456789" & Chr$(vbKeyBack), Chr$(KeyAscii)) = 0 Then 'neither numeric nor backspace
        KeyAscii = 0
        Beep
    End If

End Sub

Private Sub txLines_Validate(Cancel As Boolean)

    Cancel = Not IsNumeric(txLines)
    If Cancel Then
        Beep
    End If

End Sub

Private Sub txtTriggerLength_Change()

    If InStr("23456", txtTriggerLength) = 0 Then 'is not in a legal range
        txtTriggerLength = vbNullString
        Beep
    End If

End Sub

Private Sub txtTriggerLength_GotFocus()

  'these two do not normally cause validation so as to let the user out of txLines unvalidated and
  'checkmark one of these. but when TriggerLength has the focus they also cause validation so that
  'the user cannot leave txtTriggerlength if if fails validation

    opHalfPage.CausesValidation = True
    opPage.CausesValidation = True

End Sub

Private Sub txtTriggerLength_LostFocus()

  'lost focus - so it has been validated and we can reset

    opHalfPage.CausesValidation = False
    opPage.CausesValidation = False

End Sub

Private Sub txtTriggerLength_Validate(Cancel As Boolean)

    If Len(txtTriggerLength) = 0 Then
        Beep
        Cancel = True
    End If

End Sub

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 16:14)  Decl: 17  Code: 427  Total: 444 Lines
':) CommentOnly: 9 (2%)  Commented: 27 (6,1%)  Empty: 117 (26,4%)  Max Logic Depth: 4
