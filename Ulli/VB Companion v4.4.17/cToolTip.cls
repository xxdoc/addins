VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Custom Tooltip Class
'
'This class was inspired by code by Eidos (found at PSC some time ago) and others.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jun29 2003   UMG
'
'Added tooltip for hWnd-less controls.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jan02 2003   UMG
'
'Three new options have been added - display Tooltip always / only if parent form is active / never
'see TTStyle.
'
'Added missing Style private property variable.
'Rearranged code a little.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private TThWnd                          As Long         'Tooltip window handle
Attribute TThWnd.VB_VarDescription = "Tooltip window handle."

Private Const ToolTipWindowClassName    As String = "Tooltips_Class32"
Attribute ToolTipWindowClassName.VB_VarDescription = "Window Style."
Private Const TTS_STANDARD              As Long = 0
Private Const TTS_BALLOON               As Long = &H40
Private Const TTS_ALWAYSTIP             As Long = 1     'display even if parent window is inactive
Private Const TTS_NOPREFIX              As Long = 2     'does not remove "&" from text

Public Enum TTStyle
    TTStandardIfActive = TTS_STANDARD                   'suppress if parent form is not active
    TTBalloonIfActive = TTS_BALLOON                     'suppress if parent form is not active
    TTStandardAlways = TTS_STANDARD Or TTS_ALWAYSTIP    'display even if parent form is not active
    TTBalloonAlways = TTS_BALLOON Or TTS_ALWAYSTIP      'display even if parent form is not active
    TTNone = -1                                         'kill tooltip (this is simply treated as illegal)
End Enum
#If False Then
Private TTStandardIfActive, TTBalloonIfActive, TTStandardAlways, TTBalloonAlways, TTNone 'to preserve the case
#End If
Public Enum TTIcon
    TTIconNone = 0
    TTIconInfo = 1         'i in white balloon
    TTIconWarning = 2      '! in yellow triangle
    TTIconError = 3        'x in red circle
    'all have a light gray shadow so be careful when selecting the ToolTip BackColor
End Enum
#If False Then 'preserve capitalization
Private TTIconNone, TTIconInfo, TTIconWarning, TTIconError
#End If

'my properties
Private myStyle                     As TTStyle
Private myIcon                      As TTIcon
Private myForeColor                 As Long
Private myBackColor                 As Long
Private myTitle                     As String   'has the current title
Attribute myTitle.VB_VarDescription = "Private Property Variable."
Private myHoverTime                 As Long     'time im millisecs (-1 = use default)
Private myPopupTime                 As Long     'time im millisecs (-1 = use default)
Private myInitialText               As Variant  'has the initial text
Private myInitialTitle              As Variant  'has the initial title

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Enum ApiConstants
    SWP_NOSIZE = 1
    SWP_NOMOVE = 2
    SWP_NOACTIVATE = 16
    SWP_FLAGS = SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    TOPMOST = -1
    WM_USER = &H400
    TTM_SETDELAYTIME = WM_USER + 3
    TTM_ADDTOOL = WM_USER + 4
    TTM_SETTIPBKCOLOR = WM_USER + 19
    TTM_SETTIPTEXTCOLOR = WM_USER + 20
    TTM_SETTITLE = WM_USER + 32
    TTF_CENTERTIP = 2 'center tool on parent
    TTF_SUBCLASS = &H10 'use implicit subclassing
    CW_USEDEFAULT = &H80000000
    TTDT_AUTOPOP = 2
    TTDT_INITIAL = 3
    LoBits = &H7FFFFFFF
End Enum
#If False Then ':) Line inserted by Formatter
Private SWP_NOSIZE, SWP_NOMOVE, SWP_NOACTIVATE, SWP_FLAGS, TOPMOST, WM_USER, TTM_SETDELAYTIME, TTM_ADDTOOL, TTM_SETTIPBKCOLOR, TTM_SETTIPTEXTCOLOR, _
        TTM_SETTITLE, TTF_CENTERTIP, TTF_SUBCLASS, CW_USEDEFAULT, TTDT_AUTOPOP, TTDT_INITIAL, LoBits ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECTANGLE) As Long
Private Type RECTANGLE
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

Private Type ToolInfo
    ttSize          As Long
    myFlags         As Long
    ttParhWnd       As Long
    ttId            As Long
    ParentRect      As RECTANGLE
    hInstance       As Long
    myText          As String
    lParam          As Long
End Type

Private ToolInfo    As ToolInfo
Attribute ToolInfo.VB_VarDescription = "Tool information structure."

Public Property Get BackCol() As Long
Attribute BackCol.VB_Description = "Returns the current tooltip backcolor."

  'this returns the current tooltip backcolor

    BackCol = myBackColor

End Property

Public Property Get Centered() As Boolean
Attribute Centered.VB_Description = "Returns the current tooltip alignment."

  'this returns the current tooltip alignment

    Centered = CBool(ToolInfo.myFlags And TTF_CENTERTIP)

End Property

Private Sub Class_Initialize()

    InitCommonControls 'doesn't matter that this is called for every class instance
    myStyle = TTNone

End Sub

Private Sub Class_Terminate()

  'kill tooltip window if one exists

    If TThWnd Then
        DestroyWindow TThWnd
        TThWnd = 0
    End If
    myStyle = TTNone

End Sub

Public Function Create(Parent As Control, _
                       Text As String, _
                       Optional ByVal Style As TTStyle = TTBalloonAlways, _
                       Optional ByVal Centered As Boolean = False, _
                       Optional ByVal Icon As TTIcon = TTIconNone, _
                       Optional Title As String = "", _
                       Optional ByVal ForeColor As Long = vbButtonText, _
                       Optional ByVal BackColor As Long = vbInfoBackground, _
                       Optional ByVal HoverTime As Long = -1, _
                       Optional ByVal PopupTime As Long = -1) As Long

  'Create the tooltip window for parent control that has an hWnd
  'This can now also create custom tooltips for hWnd-less controls,
  'just supply the containing form.hWnd for windowless controls.

    Class_Terminate 'kill tooltip window if one exists
    With ToolInfo
        On Error Resume Next
            .ttParhWnd = Parent.hWnd 'the control's hWnd
            If Err Then 'has no hWnd
                Err.Clear
                .ttParhWnd = Parent.Parent.hWnd 'the control's parent's hWnd (normaly the form)
            End If
            If (Err = 0) And _
                (Style = TTBalloonAlways Or Style = TTStandardAlways Or Style = TTBalloonIfActive Or Style = TTStandardIfActive) And _
                (Icon = TTIconError Or Icon = TTIconInfo Or Icon = TTIconNone Or Icon = TTIconWarning) Then
                'the tooltip parent control has an hWnd and the params are acceptable
                .ttSize = Len(ToolInfo)
                .myFlags = TTF_SUBCLASS Or IIf(Centered, TTF_CENTERTIP, 0&)
                GetClientRect .ttParhWnd, .ParentRect
                .hInstance = App.hInstance
                myTitle = Title
                If myInitialTitle = Empty Then
                    myInitialTitle = myTitle
                End If
                .myText = Replace$(Text, "|", vbCrLf) 'the vertical bar is used as line break character
                If Len(myTitle) = 0 Then
                    .myText = Replace$(.myText, vbCrLf, Spce)
                End If
                If myInitialText = Empty Then
                    myInitialText = .myText
                End If
                If ForeColor < 0 Then
                    ForeColor = GetSysColor(ForeColor And LoBits)
                End If
                If BackColor < 0 Then
                    BackColor = GetSysColor(BackColor And LoBits)
                End If
                If ForeColor = BackColor Then
                    ForeColor = vbButtonText
                    BackColor = vbInfoBackground
                End If
                myForeColor = ForeColor
                myBackColor = BackColor
                myStyle = Style
                myIcon = Icon
                myHoverTime = HoverTime
                myPopupTime = PopupTime
                'create tooltip window and set it's properties
                TThWnd = CreateWindowEx(0&, ToolTipWindowClassName, vbNullString, TTS_NOPREFIX Or Style, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, .ttParhWnd, 0&, .hInstance, 0&)
                SetWindowPos TThWnd, TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS
                SendMessage TThWnd, TTM_ADDTOOL, 0&, ToolInfo
                SendMessage TThWnd, TTM_SETTITLE, Icon, ByVal myTitle
                SendMessage TThWnd, TTM_SETTIPTEXTCOLOR, myForeColor, ByVal 0&
                SendMessage TThWnd, TTM_SETTIPBKCOLOR, myBackColor, ByVal 0&
                SendMessage TThWnd, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal myHoverTime
                SendMessage TThWnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal myPopupTime
                Create = TThWnd
            End If
        On Error GoTo 0
    End With 'TOOLINFO

End Function

Public Property Get ForeCol() As Long
Attribute ForeCol.VB_Description = "Returns the current tooltip forecolor."

  'this returns the current tooltip forecolor

    ForeCol = myForeColor

End Property

Public Property Get HoverTime() As Long

  'this returns the current mouse hover time time in millicecs (-1 for default)

    HoverTime = myHoverTime

End Property

Public Property Get Icon() As TTIcon
Attribute Icon.VB_Description = "Returns the current tooltip icon."

  'this returns the current tooltip icon

    Icon = myIcon

End Property

Public Property Get InitialText() As String
Attribute InitialText.VB_Description = "Returns the inital tooltip text."

  'this returns the inital tooltip text, ie the one that was supplied on creation

    InitialText = myInitialText

End Property

Public Property Get InitialTitle() As String
Attribute InitialTitle.VB_Description = "Returns the inital tooltip title."

  'this returns the inital tooltip title, ie the one that was supplied on creation

    InitialTitle = myInitialTitle

End Property

Public Property Get PopupTime() As Long

  'this returns the current max PopupTime time in millisecs (-1 for default)

    PopupTime = myPopupTime

End Property

Public Property Get Style() As TTStyle
Attribute Style.VB_Description = "Returns the current tooltip style."

  'this returns the current tooltip style

    Style = myStyle

End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns the current tooltip text."

  'this returns the current tooltip text

    Text = ToolInfo.myText

End Property

Public Property Get Title() As String
Attribute Title.VB_Description = "Returns the current tooltip title."

  'this returns the current tooltip Title

    Title = myTitle

End Property

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 16:14)  Decl: 111  Code: 184  Total: 295 Lines
':) CommentOnly: 38 (12,9%)  Commented: 30 (10,2%)  Empty: 64 (21,7%)  Max Logic Depth: 4
