Attribute VB_Name = "mSubclass"
'Subclassing the IDE (and plenty of other goodies)

Option Explicit

'==============================================================================================
'
'changeable at compile time
'
' - - - - - - - - - - - - - - - - - modify both values to correspond- - - - - - - - - - - - - -

Private Const ScrollFraction        As Single = 1 / 2 'fraction of page to scroll
Public Const opHpCapt               As String = "Half a &Page"

' - - - - - - - - - - - - - - - - - Raster drawing margins- - - - - - - - - - - - - - - - - - -

'these margins were arrived at experimentally
Private Enum RasterMargins 'in pixels
    RasterTop = 30
    RasterStartPos = 13 'without the indicator bar which is 21 pixels wide
    IndicatorBarWidth = 21
End Enum
#If False Then ':) Line inserted by Formatter
Private RasterTop, RasterStartPos, IndicatorBarWidth ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

'==============================================================================================

'this depends on tabwidth and charwidth
Private RasterPitch                 As Long
'width of indicator bar if on
Private RasterLeftMargin            As Long
'...and this comes from SystemMetrics
Private RasterRightMargin           As Long

'the VB IDE
Public VBInstance                   As VBIDE.VBE  'this has a reference to the instantiated VB IDE
Public ActiveCompo                  As VBComponent
Public ResetMenuButton              As Office.CommandBarButton
Private IsGreen                     As Boolean
Public MoveableFormIsShowing        As Boolean
Private InhibitGrid                 As Boolean
Public OpenAllMenuButton            As Office.CommandBarButton
Public CompareMenuButton            As Office.CommandBarButton
Public CopyMenuButton               As Office.CommandBarButton
Public CurrentCodeName              As String
Public Const NAC                    As String = "No active component"

'subclassing
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const IDX_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const IDX_WINDOWPROC        As Long = -4
Private Const WM_SETFOCUS           As Long = 7
Private Const WM_KILLFOCUS          As Long = 8
Private Const WM_PAINT              As Long = 15
Public Const WM_CLOSE               As Long = 16
Private Const WM_MOUSEACTIVATE      As Long = &H21
Private Const WM_KEYDOWN            As Long = &H100
Private Const WM_KEYUP              As Long = &H101
Private Const WM_CHAR               As Long = &H102
Private Const WM_HSCROLL            As Long = &H114
Private Const WM_VSCROLL            As Long = &H115
Private Const WM_LBUTTONUP          As Long = &H202
Private Const WM_MBUTTONDOWN        As Long = &H207
Private Const WM_MOUSEWHEEL         As Long = &H20A
Private Const WM_MDIACTIVATE        As Long = &H222

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN       As Long = &HA1
Public Const HTCAPTION              As Long = 2

'yes, we will be sending messages
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'timing
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public HiResTimerPresent            As Long
Public CPUFreq                      As Currency
Private CPUTicksStart               As Currency
Private CPUTicksNow                 As Currency

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE            As Long = 1
Private Const SWP_NOMOVE            As Long = 2
Private Const SWP_NOZORDER          As Long = 4
Private Const SWP_FRAMECHANGED      As Long = &H20
Private Const SWP_AFTERFRAMECHANGE  As Long = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED

'properties
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Const PropName              As String = "HookedByUlli"

'focus, keyboard and positions
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As tPOINT) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As tPOINT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As tRECT) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'graphic
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As tPOINT, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function InvertRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function InvalidateRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As tRECT, ByVal bErase As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Private Declare Function SetPenPosition Lib "gdi32" Alias "MoveToEx" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function DrawLine Lib "gdi32" Alias "LineTo" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As Long, lpScrollInfo As tSCROLLINFO) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const WINDING               As Long = 2
Public Const PS_SOLID               As Long = 0
Public Const PS_DOT                 As Long = 2
Private Const SM_CXVSCROLL          As Long = 2

'scrollbar info
Private Const SB_CTL                As Long = 2
Private Const SB_ENDSCROLL          As Long = 8
Private Const SIF_POS               As Long = 4
Private hWndScrollbar               As Long 'handle of the scrollbar window
Private Type tSCROLLINFO
    cbSize                          As Long
    fMask                           As Long
    nMin                            As Long
    nMax                            As Long
    nPage                           As Long
    nPos                            As Long
    nTrackPos                       As Long
End Type
Private SCROLLINFO                  As tSCROLLINFO

'vb code execution
Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal PtrToStringToExec As Long, ByVal Any1 As Long, ByVal Any2 As Long, ByVal CheckOnly As Long) As Long

'About box
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'cursor
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Const CursorArrow            As Long = 32512
Public Const CursorRightHand        As Long = 32649

'Send Mail
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL         As Long = 1
Private Const SE_NO_ERROR           As Long = 33 'Values below 33 are error returns

'Registry
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Sub RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long)
Private Const HKEY_CURRENT_USER     As Long = &H80000001
Private Const KEY_QUERY_VALUE       As Long = 1
Private Const REG_OPTION_RESERVED   As Long = 0
Private Const ERROR_NONE            As Long = 0
Private RegHandle                   As Long
Private DataType                    As Long
Private DataLength                  As Long

'noise
Public Declare Function Beeper Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

'idle
Private Declare Function BeginIdleDetection Lib "Msidle.dll" Alias "#3" (ByVal pfnCallback As Long, ByVal dwIdleMin As Long, ByVal dwReserved As Long) As Long
Private Declare Function EndIdleDetection Lib "Msidle.dll" Alias "#4" (ByVal dwReserved As Long) As Long
Private Const USER_IDLE_BEGIN       As Long = 1
Private Const USER_IDLE_END         As Long = 2
Private Idletime                    As Long

'types and structures
Public Type tPOINT
    X       As Long
    Y       As Long
End Type
Public CursorPos                    As tPOINT
Private CaretPos                    As tPOINT
Public XPointerVertices(1 To 4)     As tPOINT
Public ArrowVertices(1 To 7)        As tPOINT
Private RegionVertices(1 To 7)      As tPOINT

Public Type tRECT 'defined by top left and bottom right corner points
    TopLeft     As tPOINT
    BottomRight As tPOINT
End Type
Public WindowRect                   As tRECT

Public Enum ArrayIndexes 'and other goodies
    idxCompoName = 0
    idxMembername = 1
    idxMemberScope = 2
    idxMemberType = 3
    NoScope = 4
End Enum
#If False Then ':) Line inserted by Formatter
Private idxCompoName, idxMembername, idxMemberScope, idxMemberType, NoScope ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Public Enum FlashAction
    Inval = 1
    Flash = 2
    Arrow = 4
    XPointer = 8
End Enum
#If False Then ':) Line inserted by Formatter
Private Inval, Flash, Arrow, XPointer ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Private Enum MouseButtons
    PrimaryButton = 1
    SecondaryButton = 2
End Enum
#If False Then ':) Line inserted by Formatter
Private PrimaryButton, SecondaryButton ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Public Enum PenAction
    DestroyIt = 1
    CreateIt = 2
    DestroyItCreateIt = DestroyIt Or CreateIt
End Enum
#If False Then ':) Line inserted by Formatter
Private DestroyIt, CreateIt, DestroyItCreateIt ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

'database
Public ApiDBFileName                As String
Public ApiDatabase                  As Database
Public MasterSet                    As Recordset
Public SlaveSet                     As Recordset

'hooks, hWnds and icon
Public hWndCodePane                 As Long
Private hDCCodePane                 As Long
Public hWndCompare                  As Long
Private hGridPen                    As Long
Public GridColor                    As Long
Public RasDrwMode                   As Long
Private hPrevPen                    As Long
Public hWndMDIClient                As Long
Private hDCMDIClient                As Long
Private CheckDirtyTimerId           As Long
Public hDCIcon                      As Long
Public wIcon                        As Long
Public hIcon                        As Long

'SQL
Private Const sSelectFrom           As String = "SELECT * FROM "
Public Const sDeclares              As String = "Declares"
Private Const sConstants            As String = "Constants"
Private Const sTypes                As String = "Types"
Private Const sTypeItems            As String = "TypeItems"

'DB-fieldnames
Private Const sFullName             As String = "FullName"
Private Const sName                 As String = "Name"
Private Const sTypeItem             As String = "TypeItem"

'assortment of chars to be replaced by space
Private Const CharsToSpace          As String = "()[],""<>=+&-*/\^';:"

'registry custom colors

'registry VB tabwidth, indicator bar and font settings/colors
Private Const VBASettings           As String = "Software\Microsoft\VBA\Microsoft Visual Basic"
Private Const sTabWidth             As String = "TabWidth"
Private Const sFontheight           As String = "FontHeight"
Private Const sFontface             As String = "FontFace"
Private Const sIndicator            As String = "IndicatorBar"
Private Const sCFC                  As String = "CodeForeColors"
Private Const DefaultFontName       As String = "Fixedsys"
Private Const DefaultFontSize       As Long = 9
Public IDEFontName                  As String
Public IDEFontSize                  As Long
Private Const DefaultColors         As String = "0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0" '17 colors
Public IDEColors                    As String
Private CharWidth                   As Long
Public LineHeight                   As Long
Public RegionCenterY                As Long

'our own settings...
Public Const sOptions               As String = "Options"
Public Const sLinesToScroll         As String = "LinesToScroll"
Public sHalfPage                    As String
Public Const sFullPage              As String = "Full Page"
Public Const sMode                  As String = "Mode"
Public Const sSmooth                As String = "Smooth"
Public Const sInstant               As String = "Instant"
Public Const sSpeed                 As String = "Speed"
Public Const sAutoComplete          As String = "AutoComplete"
Public Const sNoisy                 As String = "Noisy"
Public Const sApiLocation           As String = "APILocation"
Public Const sAskUser               As String = "[askuser]"
Public Const sUnknown               As String = "[unknown]"
Public Const sTriggerLength         As String = "TriggerLength"
Public Const sUnique                As String = "UniqueMatch"
Public Const sRaster                As String = "Raster"
Public Const sROP                   As String = "RasterDrawMode"
Public Const sNone                  As String = "None"
Public Const sVertical              As String = "Vertical"
Public Const sHorizontal            As String = "Horizontal"
Public Const sBoth                  As String = "Horizontal and Vertical"
Public Const sRasterStyle           As String = "RasterStyle"
Public Const sSolid                 As String = "Solid"
Public Const sDotted                As String = "Dotted"
Public Const sRasterColor           As String = "RasterColor"
Public Const sCopyPen               As String = "Copy Pen"
Public Const sMaskPen               As String = "Mask Pen"
Public Const sOn                    As String = "On"
Public Const sOff                   As String = "Off"
Public Const Spce                   As String = " "

Public Const UndefinedCaret         As String = "You cannot open the $ while the caret position is undefined." & vbCrLf & vbCrLf & "Set proper insertion point first..."

'what we got (or didn't get) from the Registry or from our own options
Private VBTabWidth                  As Long
Public Smooth                       As Long
Public ScrollDelay                  As Single
Public AutoComplete                 As Boolean
Public Noisy                        As Boolean
Public NonUnique                    As Boolean
Public Raster                       As Long
Public FixedFontPitch               As Boolean
Public LineStyle                    As Long
Public TriggerLength                As Long 'the minimum length of word fragment to trigger autocomplete

'wheel scrolling
Public NumLinesToScroll             As String
Private IsScrolling                 As Boolean
Private ScrollTo                    As Long 'line number
Private LastTop                     As Long 'last known topline number
Private PrevTop                     As Long 'and the one before that
Public Const SSScale                As Single = 10000

'misc
Public CodeMembers                  As Collection
Public Item                         As Variant
Public NumTexts                     As Variant
Private UserTypedCode               As Boolean
Public DontAskAgain                 As Boolean
Public InDesignMode                 As String

Public Function AppDetails() As String

    With App
        AppDetails = .ProductName & " V" & .Major & "." & .Minor & "." & .Revision
    End With 'APP

End Function

Public Sub CaretRgn(ByVal Action As FlashAction, Optional ByVal Cnt As Long = 0)

  'flashes or invalidates a region around the caret
  'a positive count gets the caret position first

  Dim hRgn          As Long

    If (Action And Flash) And Cnt > 0 Then
        GetCaretPos CaretPos 'where's the caret
    End If
    If Action And XPointer Then
        For hRgn = 1 To UBound(XPointerVertices) 'hRgn is mis-used as counter here
            RegionVertices(hRgn) = OffsetVertex(XPointerVertices(hRgn), CaretPos)
        Next hRgn
      Else 'Arrow 'NOT ACTION...
        For hRgn = 1 To UBound(ArrowVertices)
            RegionVertices(hRgn) = OffsetVertex(ArrowVertices(hRgn), CaretPos)
        Next hRgn
    End If 'hRgn now has the number of vertices plus one
    hRgn = CreatePolygonRgn(RegionVertices(1), hRgn - 1, WINDING) 'and now hRgn has the handle to the region
    If Action And Inval Then
        InvalidateRgn hWndCodePane, hRgn, False 'so that the next paint will remove the arrow
      Else 'NOT ACTION...
        HideCaret hWndCodePane 'hide caret
        Do Until Cnt = 0
            If Cnt > 0 Then 'its the normal flashing action; otherwise we are just redrawing the region
                Wait 0.08
            End If
            InvertRgn hDCCodePane, hRgn 'flash the region
            Cnt = Cnt - Sgn(Cnt)
        Loop
        ShowCaret hWndCodePane 'show caret again
    End If
    DeleteObject hRgn 'and finally delete the region

End Sub

Private Function CodePaneProc(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  'Window Procedure for the active IDE Codepane

  Dim i             As Long
  Dim j             As Long
  Dim SelStartLine  As Long
  Dim WordStartsAt  As Long
  Dim LenWord       As Long
  Dim LenRepl       As Long
  Dim MchCnt        As Long
  Dim FromWhichBox  As Long 'governs the behavior regarding selections after the deed
  Dim OriginalLine  As String
  Dim TypeBody      As String
  Dim CurrentWord   As String
  Dim TstStr        As String
  Dim TmpStr        As String
  Dim RplStr        As String
  Dim TableName     As String

    With VBInstance
        CodePaneProc = CallWindowProc(GetProp(hWnd, PropName), hWnd, nMsg, wParam, lParam)
        TmpStr = .MainWindow.Caption
        If InStr(TmpStr, InDesignMode) Then
            With .ActiveCodePane
                Select Case nMsg
                  Case WM_PAINT, WM_KEYUP, WM_VSCROLL
                    DrawRaster
                  Case WM_HSCROLL
                    hWndScrollbar = lParam 'save scrollbar handle
                    If wParam And SB_ENDSCROLL Then 'user released the mouse, probably just clicked
                        EraseRaster 'so erase to redraw the raster
                    End If
                    DrawRaster
                  Case WM_LBUTTONUP    'primary (left) button up
                    LastTop = .TopLine 'save current positioning on left mouse up
                    DrawRaster
                  Case WM_MBUTTONDOWN 'middle button or wheel down
                    If .CountOfVisibleLines Then
                        .GetSelection SelStartLine, i, WordStartsAt, j 'this looks funny but the effect is that it
                        .SetSelection SelStartLine, i, WordStartsAt, j 'repositions correctly in the x-direction
                        If SelStartLine >= LastTop + .CountOfVisibleLines Then 'codewindow has gotten smaller
                            LastTop = SelStartLine - .CountOfVisibleLines + 1
                        End If
                        .TopLine = LastTop
                        DrawRaster
                        If SelStartLine = WordStartsAt And i = j Then 'no selection
                            RegionCenterY = LineHeight \ 2 + 1
                            CaretRgn XPointer Or Flash, 12 'even - flash and remove
                        End If
                      Else '.COUNTOFVISIBLELINES = FALSE/0
                        MsgBoxEx "Why would you try and see the current line" & vbCrLf & _
                                 "when you don't see any line in this window?", PosX:=-2, OffsetY:=-38, Icon:=fIcon.Icon, OCapt:=OK, NCapt:="&Ooops..."
                    End If
                  Case WM_KEYDOWN
                    PrevTop = LastTop
                    LastTop = .TopLine 'save current positioning on keydown
                    Select Case wParam
                      Case vbKeyReturn, vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown 'the user has just left a line
                        DoEvents
                        RefreshMembers UserTypedCode 'so refresh our members collection (if that line was typed into)
                        UserTypedCode = False
                      Case vbKeyPause 'the user wishes to open the member list box (fSelect) or the multiline box (fMultiline)
                        LastTop = PrevTop
                        .TopLine = LastTop
                        .GetSelection SelStartLine, WordStartsAt, i, j
                        .SetSelection SelStartLine, WordStartsAt, i, j
                        If SelStartLine = i And WordStartsAt = j Then 'no selection is okay
                            RefreshMembers UserTypedCode 'so refresh our members collection if necessary
                            UserTypedCode = False
                            RegionCenterY = LineHeight \ 2 + 1
                            If GetAsyncKeyState(vbKeyShift) < 0 Then 'shift is also pressed
                                MoveToArrow fMultiline
                                fMultiline.Indent = 1
                                On Error Resume Next
                                    fMultiline.Indent = j - (Mid$(.CodeModule.Lines(i, 1), WordStartsAt - 1, 1) <> Spce)
                                On Error GoTo 0
                                With fMultiline
                                    .Show vbModal
                                    TmpStr = .Tag
                                End With 'FMULTILINE
                                Unload fMultiline
                                FromWhichBox = 2 + Len(TmpStr)
                              Else 'NOT GETASYNCKEYSTATE(VBKEYSHIFT)...
                                MoveToArrow fSelect
                                With fSelect
                                    .CurrentComponentName = CurrentCodeName
                                    .Show vbModal
                                    TmpStr = Spce & .Tag & Spce 'enclose by two spaces front and back
                                End With 'FSELECT
                                Unload fSelect
                                FromWhichBox = 1
                            End If
                            With .CodeModule
                                OriginalLine = .Lines(SelStartLine, 1)
                                CurrentWord = SpaceReplace(OriginalLine) 'replace some special chars by space
                                If Len(TmpStr) > 2 Then 'more than just two spaces
                                    If WordStartsAt > 1 Then
                                        If Mid$(CurrentWord, WordStartsAt - 1, 1) = Spce Then 'space left of insertion point
                                            TmpStr = LTrim$(TmpStr)
                                        End If
                                      Else 'insertion at left boundary 'NOT WORDSTARTSAT...
                                        TmpStr = LTrim$(TmpStr)
                                    End If
                                    If Mid$(CurrentWord, WordStartsAt, 1) = Spce Then 'space right of insertion point
                                        TmpStr = RTrim$(TmpStr)
                                    End If
                                    .ReplaceLine SelStartLine, Left$(OriginalLine, WordStartsAt - 1) & TmpStr & Mid$(OriginalLine, j)
                                  Else 'NOT LEN(TMPSTR)...
                                    TmpStr = vbNullString
                                End If
                            End With '.CODEMODULE
                            If FromWhichBox = 1 Then 'from fSelect
                                .SetSelection SelStartLine, WordStartsAt, SelStartLine, WordStartsAt + Len(TmpStr)
                              ElseIf FromWhichBox <> 2 Then 'from fMultiline (not cancelled) 'NOT FROMWHICHBOX...
                                .GetSelection i, i, j, i
                                .SetSelection SelStartLine, WordStartsAt, j - 1, 1023
                            End If
                          Else 'caret posn undefined 'NOT SELSTARTLINE...
                            MsgBoxEx Replace$(UndefinedCaret, "$", "the Member Selection or Multiline Box"), PosX:=-2, OffsetY:=-45, Icon:=fIcon.Icon, OCapt:=OK, NCapt:="&Close"
                        End If
                        .Window.SetFocus 'reset focus
                    End Select
                  Case WM_CHAR 'the user has typed a char
                    Select Case wParam
                      Case vbKeyBack, vbKeySpace, vbKeyReturn, vbKeyTab
                        'do nothing - would be kinda indecent to autocomplete if the user is backspace-deleting or inserting a space or an empty line
                        UserTypedCode = False
                      Case Else
                        UserTypedCode = True
                        If AutoComplete Then
                            .GetSelection SelStartLine, i, j, j 'where's the caret
                            OriginalLine = .CodeModule.Lines(SelStartLine, 1) 'get the line the user is currently typing into
                            If Len(OriginalLine) And i > 1 Then 'there is something in that line - have a closer look
                                CurrentWord = SpaceReplace(OriginalLine) 'replace some special chars by space
                                WordStartsAt = InStrRev(CurrentWord, Spce, i - 1) + 1 'start posn of the word the user is typing
                                If IsInCode(OriginalLine, WordStartsAt) Then 'this is part of the coding (not comment nor literal)
                                    CurrentWord = LCase$(Mid$(CurrentWord, WordStartsAt)) 'isolate the word-fragment the user is typing right now
                                    i = InStr(CurrentWord, Spce)
                                    If i Then
                                        CurrentWord = Left$(CurrentWord, i - 1)
                                    End If
                                    LenWord = Len(CurrentWord)
                                    If LenWord >= TriggerLength Or NonUnique = False Then 'don't harrass the user with too short words unless he wants unique matches anyway
                                        If SelStartLine > .CodeModule.CountOfDeclarationLines Then 'user is typing in the code part
                                            If NotAnApiLine(OriginalLine) Then '..and its not an api so search the Members Collection
                                                MchCnt = 0
                                                For Each Item In CodeMembers 'Item is a Variant of Type Array
                                                    If Item(idxCompoName) = CurrentCodeName Then 'check for applicable scope
                                                        i = 0
                                                      Else 'NOT ITEM(IDXCOMPONAME)...
                                                        i = vbext_Private 'i has forbidden scope
                                                    End If
                                                    If Item(idxMemberScope) <> i Then 'this member has an appropriate scope
                                                        TstStr = CStr(Item(idxMembername)) 'get name of member
                                                        LenRepl = Len(TstStr)
                                                        If LenWord < LenRepl Then 'user has not yet finished typing the word so offer autocomplete
                                                            If CurrentWord = LCase$(Left$(TstStr, LenWord)) Then 'this is a possible autocomplete
                                                                Inc MchCnt
                                                                TmpStr = TstStr
                                                                If NonUnique Then
                                                                    Exit For 'loop varying item
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Next Item
                                                If MchCnt = 1 Then
                                                    RplStr = TmpStr
                                                    LenRepl = Len(TmpStr)
                                                End If
                                            End If
                                          Else 'its in the declarations part 'NOT SELSTARTLINE...
                                            If Not ApiDatabase Is Nothing Then 'we have an api database...
                                                If InStr(1, OriginalLine, sApiDeclare, vbTextCompare) Then
                                                    TableName = sDeclares
                                                  ElseIf InStr(1, OriginalLine, sApiConst, vbTextCompare) Then 'NOT INSTR(1,...
                                                    TableName = sConstants
                                                  ElseIf InStr(1, OriginalLine, sApiType, vbTextCompare) Then 'NOT INSTR(1,...
                                                    TableName = sTypes
                                                End If
                                                If Len(TableName) Then '...and the user is attempting an api - we have a table name
                                                    'Open recordset from composed SQL string
                                                    Set MasterSet = ApiDatabase.OpenRecordset(sSelectFrom & TableName & " WHERE Name LIKE """ & IIf(Right$(CurrentWord, 1) = "?", Left$(CurrentWord, Len(CurrentWord) - 1), CurrentWord & "*") & """")
                                                    With MasterSet
                                                        If Not .EOF Then 'there is something in that recordset
                                                            .MoveLast 'populate recordset to get the number of rows
                                                            If .RecordCount = 1 Then 'found one single row
                                                                Select Case TableName
                                                                  Case sDeclares
                                                                    WordStartsAt = InStr(1, OriginalLine, sApiDeclare, vbTextCompare) 'where will the replacement go
                                                                    RplStr = .Fields(sFullName) '...and the replacement itself
                                                                  Case sConstants
                                                                    WordStartsAt = InStr(1, OriginalLine, sApiConst, vbTextCompare) 'as above
                                                                    RplStr = .Fields(sFullName)
                                                                    If InStr(RplStr, " As ") = 0 Then
                                                                        RplStr = Replace$(RplStr, "=", "As Long =", , 1) 'add "As Long" if there is no "As"
                                                                    End If
                                                                  Case sTypes
                                                                    WordStartsAt = InStr(1, OriginalLine, sApiType, vbTextCompare)
                                                                    RplStr = "Type " & .Fields(sName)
                                                                    'open recordset having Type-members by a composed SQL string
                                                                    Set SlaveSet = ApiDatabase.OpenRecordset(sSelectFrom & sTypeItems & " WHERE TypeID = " & .Fields("ID") & " ORDER BY Id")
                                                                    With SlaveSet
                                                                        Do Until .EOF 'cycle thru Type-members and append them via TypeBody
                                                                            TypeBody = TypeBody & vbCrLf & Space$(VBTabWidth) & LTrim$(.Fields(sTypeItem))
                                                                            .MoveNext
                                                                        Loop
                                                                        .Close
                                                                    End With 'SLAVESET
                                                                    TypeBody = TypeBody & vbCrLf & "End Type"
                                                                End Select
                                                                OriginalLine = Left$(OriginalLine, WordStartsAt - 1) 'trunc original line
                                                            End If
                                                        End If
                                                        .Close
                                                    End With 'MASTERSET
                                                End If
                                            End If
                                        End If
                                        If Len(RplStr) = 0 Then 'no replacement found yet, so search the KeyWords Collection
                                            If NotAnApiLine(OriginalLine) Then
                                                MchCnt = 0
                                                For Each Item In KeyWords
                                                    TstStr = CStr(Item) 'here Item is a Variant of Type String
                                                    LenRepl = Len(TstStr)
                                                    If LenWord < LenRepl Then 'user has not yet finished typing the word so offer autocomplete
                                                        If CurrentWord = LCase$(Left$(TstStr, LenWord)) Then
                                                            Inc MchCnt
                                                            TmpStr = TstStr
                                                            If NonUnique Then
                                                                Exit For 'loop varying item
                                                            End If
                                                        End If
                                                    End If
                                                Next Item
                                                If MchCnt = 1 Then
                                                    RplStr = TmpStr
                                                    LenRepl = Len(TmpStr)
                                                End If
                                            End If
                                        End If
                                        If Len(RplStr) Then 'we have something to offer
                                            OriginalLine = Replace$(Replace$(Replace$(Left$(OriginalLine, WordStartsAt - 1) & RplStr & Mid$(OriginalLine, WordStartsAt + LenWord), "><", "<>"), "=<", "<="), "=>", ">=") 'a little cosmetic
                                            .CodeModule.ReplaceLine SelStartLine, OriginalLine & TypeBody 'put it in the code
                                            If Len(TableName) Then 'this came from API DB
                                                i = 0
                                                LenWord = Len(RplStr)
                                                LenRepl = LenWord
                                              Else 'this came from member- or keyword-list 'LEN(TABLENAME) = FALSE/0
                                                TmpStr = .CodeModule.Lines(SelStartLine, 1) 'get line back to see whether VB has expanded or compacted the line and adjust selection accordingly
                                                i = 1
                                                j = 1
                                                Do 'determine new selection
                                                    Select Case Spce
                                                      Case Mid$(OriginalLine, i, 1)
                                                        Inc i
                                                      Case Mid$(TmpStr, j, 1)
                                                        Inc j
                                                      Case Else
                                                        Inc i
                                                        Inc j
                                                    End Select
                                                Loop Until i = WordStartsAt + LenRepl Or j = WordStartsAt + LenRepl
                                                i = j - i
                                            End If
                                            .SetSelection SelStartLine, WordStartsAt + LenRepl + i, SelStartLine, WordStartsAt + LenWord + i 'put selection in line
                                            If Noisy Then
                                                Beeper 4000, 20 'make a short tick noise
                                            End If 'noisy
                                        End If 'something there to offer
                                    End If 'long enough to autocomplete
                                End If 'is in code
                            End If 'something in current line
                        End If 'Autocomplete
                    End Select 'wParam
                End Select 'nMsg
            End With '.ACTIVECODEPANE
        End If 'in design mode
    End With 'VBINSTANCE

End Function

Public Function Dec(ByRef What As Long, Optional ByVal By As Long = 1) As Long

    What = What - By
    Dec = What

End Function

Public Sub DirtyTimer(ByVal Action As PenAction)

    If Action And DestroyIt Then
        If CheckDirtyTimerId Then
            KillTimer 0, CheckDirtyTimerId
            CheckDirtyTimerId = 0
        End If
    End If
    If Action And CreateIt Then
        If CheckDirtyTimerId = 0 Then
            CheckDirtyTimerId = SetTimer(0, 0, 200, AddressOf PollDirty)
        End If
    End If

End Sub

Public Sub DrawRaster()

  Dim HorizPosn     As Long
  Dim VertPosn      As Long
  Dim HorizOffset   As Long
  Dim RasterBottom  As Long
  Dim PreviousROP   As Long

    If Not InhibitGrid Then
        If Raster And FixedFontPitch Then
            RasterBottom = RasterTop + VBInstance.ActiveCodePane.CountOfVisibleLines * LineHeight
            If RasterBottom > RasterTop Then 'there is at least one line
                GetScrollInfo hWndScrollbar, SB_CTL, SCROLLINFO 'so we get it directly from the scrollbar once we know it's hWnd
                HorizOffset = SCROLLINFO.nPos Mod VBTabWidth 'just the remainder
                If HorizOffset > 0 Then
                    HorizOffset = (VBTabWidth - HorizOffset) * CharWidth 'pixels to the left
                End If
                GetWindowRect hWndCodePane, WindowRect 'so get canvas size to draw on
                HideCaret hWndCodePane 'if the caret sits on a raster line it would get drawn over so we better hide it
                PreviousROP = SetROP2(hDCCodePane, RasDrwMode) 'set to seleceted mode

                If Raster And 1 Then 'draw vertical raster
                    For HorizPosn = RasterLeftMargin + RasterStartPos + HorizOffset To WindowRect.BottomRight.X - WindowRect.TopLeft.X - RasterRightMargin Step RasterPitch
                        SetPenPosition hDCCodePane, HorizPosn, RasterTop, ByVal 0&  'move pen to top
                        DrawLine hDCCodePane, HorizPosn, RasterBottom 'draw line to bottom
                    Next HorizPosn
                End If

                If Raster And 2 Then 'draw horizontal raster
                    HorizPosn = WindowRect.BottomRight.X - WindowRect.TopLeft.X - RasterRightMargin
                    HorizOffset = RasterLeftMargin + RasterStartPos
                    For VertPosn = RasterTop To RasterBottom Step LineHeight
                        SetPenPosition hDCCodePane, HorizOffset, VertPosn, ByVal 0& 'move pen to left
                        DrawLine hDCCodePane, HorizPosn, VertPosn 'draw line to right
                    Next VertPosn
                End If

                'tidy up
                SetROP2 hDCCodePane, PreviousROP 'restore rop code
                ShowCaret hWndCodePane 'restore caret
            End If
        End If
        If MoveableFormIsShowing Then 'drawing the raster was caused by a change of Form Size or Posn
            InhibitGrid = True 'prevent avalanche - drawing the black arrow causes drawing the raster
            RedrawArrow 'so we redraw the arrow in case it was unhidden by the moveable form
            InhibitGrid = False
        End If
    End If

End Sub

Public Sub EraseRaster(Optional ByVal Always As Boolean = False)

    If Raster Or Always Then
        GetClientRect hWndCodePane, WindowRect
        WindowRect.TopLeft.Y = RasterTop
        InvalidateRect hWndCodePane, WindowRect, False 'so that the grid is erased on next redraw
    End If

End Sub

Public Sub ExecuteVBCode(CodeLine As String, Optional DebugOnly As Boolean = False)

    If EbExecuteLine(StrPtr(CodeLine), 0&, 0&, CLng(Abs(DebugOnly))) Then
        Beeper 660, 10
        Beeper 440, 10
    End If

End Sub

Public Function GetNumText(Number As Long) As String

    If Number < LBound(NumTexts) Or Number > UBound(NumTexts) Then
        GetNumText = CStr(Number)
      Else 'NOT NUMBER...
        GetNumText = NumTexts(Number)
    End If

End Function

Public Sub GetRegistrySettings() 'and some other initializing
Attribute GetRegistrySettings.VB_Description = "Gets all relevant registry settings."

    If RegOpenKeyEx(HKEY_CURRENT_USER, VBASettings, REG_OPTION_RESERVED, KEY_QUERY_VALUE, RegHandle) = ERROR_NONE Then
        'get VB registry settings
        DataLength = Len(VBTabWidth)
        If RegQueryValueEx(RegHandle, sTabWidth, REG_OPTION_RESERVED, DataType, VBTabWidth, DataLength) <> ERROR_NONE Then
            VBTabWidth = 4 'default
        End If
        'the IDE code pane is cheating so we get fontname and -size from the registry
        DataLength = Len(IDEFontSize)
        If RegQueryValueEx(RegHandle, sFontheight, REG_OPTION_RESERVED, DataType, IDEFontSize, DataLength) <> ERROR_NONE Then
            IDEFontSize = DefaultFontSize
        End If
        DataLength = Len(RasterLeftMargin)
        If RegQueryValueEx(RegHandle, sIndicator, REG_OPTION_RESERVED, DataType, RasterLeftMargin, DataLength) <> ERROR_NONE Then
            RasterLeftMargin = 1 'default: is present
        End If
        IDEFontName = String$(128, 0)
        DataLength = Len(IDEFontName)
        If RegQueryValueEx(RegHandle, sFontface, REG_OPTION_RESERVED, DataType, ByVal IDEFontName, DataLength) <> ERROR_NONE Then
            IDEFontName = DefaultFontName
          Else 'NOT REGQUERYVALUEEX(REGHANDLE,...
            IDEFontName = Left$(IDEFontName, DataLength + (Asc(Mid$(IDEFontName, DataLength, 1)) = 0))
        End If
        IDEColors = String$(128, 0)
        DataLength = Len(IDEColors)
        If RegQueryValueEx(RegHandle, sCFC, REG_OPTION_RESERVED, DataType, ByVal IDEColors, DataLength) <> ERROR_NONE Then
            IDEColors = DefaultColors
          Else 'NOT REGQUERYVALUEEX(REGHANDLE,...
            IDEColors = Left$(IDEColors, DataLength + (Asc(Mid$(IDEColors, DataLength, 1)) = 0))
        End If
        RegCloseKey RegHandle
      Else 'all default to std values 'NOT REGOPENKEYEX(HKEY_CURRENT_USER,...
        VBTabWidth = 4
        IDEFontSize = DefaultFontSize
        IDEFontName = DefaultFontName
        IDEColors = DefaultColors
        RasterLeftMargin = 1
    End If

    'measure font
    fIcon.FontName = IDEFontName
    fIcon.FontSize = IDEFontSize
    CharWidth = fIcon.TextWidth("I")
    LineHeight = fIcon.TextHeight("I") 'hah! gotcha!!
    RasterPitch = VBTabWidth * CharWidth

    'indicator bar on the left (yes, it can be switched off)
    If RasterLeftMargin Then
        RasterLeftMargin = IndicatorBarWidth 'width in pixels
    End If

    'the width of the scrollbar on the right
    RasterRightMargin = GetSystemMetrics(SM_CXVSCROLL) + 10

    'get our own settings
    With App
        NumLinesToScroll = GetSetting(.Title, sOptions, sLinesToScroll, sHalfPage)
        Smooth = (GetSetting(.Title, sOptions, sMode, sSmooth) = sSmooth)
        ScrollDelay = Val(GetSetting(.Title, sOptions, sSpeed, "75")) / SSScale
        AutoComplete = (GetSetting(.Title, sOptions, sAutoComplete, sOn) = sOn)
        Noisy = (GetSetting(.Title, sOptions, sNoisy, sOn) = sOn)
        Select Case GetSetting(.Title, sOptions, sRaster, "Both")
          Case sNone
            Raster = 0
          Case sVertical
            Raster = 1
          Case sHorizontal
            Raster = 2
          Case Else
            Raster = 3
        End Select
        FixedFontPitch = (fIcon.TextWidth("W") = CharWidth)
        GridColor = Val("&H" & GetSetting(.Title, sOptions, sRasterColor, Hex$(RGB(236, 236, 240)))) 'a lite shade of reddish gray
        RasDrwMode = IIf(GetSetting(.Title, sOptions, sROP, sMaskPen) = sMaskPen, vbMaskPen, vbCopyPen)
        LineStyle = IIf((GetSetting(.Title, sOptions, sRasterStyle, sSolid)) = sSolid, PS_SOLID, PS_DOT)
        TriggerLength = GetSetting(.Title, sOptions, sTriggerLength, "3")
        NonUnique = (GetSetting(.Title, sOptions, sUnique, sOn) = sOff)
        ApiDBFileName = GetSetting(.Title, sOptions, sApiLocation, sAskUser)
    End With 'APP

    '...and finally initialize scrollinfo struct
    With SCROLLINFO
        .cbSize = Len(SCROLLINFO)
        .fMask = SIF_POS 'Scroll Info Flag <- Get Position
    End With 'SCROLLINFO

End Sub

Private Sub HookCodePane()

  'subclass the code pane

    On Error Resume Next
        With VBInstance
            If .ActiveWindow Is .ActiveCodePane.Window Then
                hWndCodePane = FindWindowEx(hWndMDIClient, 0&, "VbaWindow", .ActiveWindow.Caption)
                If hWndCodePane Then 'we have an open code pane
                    Set ActiveCompo = .ActiveCodePane.CodeModule.Parent
                    CurrentCodeName = ActiveCompo.Name
                    If GetProp(hWndCodePane, PropName) = 0 Then 'not yet hooked
                        UpdateTooltips
                        SetProp hWndCodePane, PropName, SetWindowLong(hWndCodePane, IDX_WINDOWPROC, AddressOf CodePaneProc)
                        hDCCodePane = GetDC(hWndCodePane) 'get codepane device context
                        BitBlt hDCMDIClient, 0&, 0&, wIcon, hIcon, hDCIcon, 0&, 0&, vbSrcCopy 'show icon
                        Pen CreateIt
                        DirtyTimer DestroyItCreateIt
                    End If
                End If
                RefreshMembers True
                DrawRaster
            End If
        End With 'VBINSTANCE
    On Error GoTo 0

End Sub

Public Sub HookMDIClient()
Attribute HookMDIClient.VB_Description = "Hooks (subclasses) VB' s MDI client window."

  'subclass the MDI

    On Error Resume Next
        hWndMDIClient = FindWindowEx(VBInstance.MainWindow.hWnd, 0&, "MDIClient", vbNullString)
        If hWndMDIClient Then 'found the MDI client window
            If GetProp(hWndMDIClient, PropName) = 0 Then 'not yet hooked
                SetProp hWndMDIClient, PropName, SetWindowLong(hWndMDIClient, IDX_WINDOWPROC, AddressOf MDIClientProc)
                hDCMDIClient = GetDC(hWndMDIClient)
                LastTop = 1
            End If
            HookCodePane
        End If
    On Error GoTo 0

End Sub

Public Sub IdleBeginDetection(Optional ByVal IdleMinutes As Long = 5)

    BeginIdleDetection AddressOf IdleCallBack, IdleMinutes, 0&
    Idletime = IdleMinutes

End Sub

Public Sub IdleCallBack(ByVal dwState As Long)

  Dim Proj      As VBProject
  Dim Compo     As VBComponent
  Dim PName     As String
  Dim MsgText   As String
  Dim Sound1    As Long
  Dim Sound2    As Long
  Const IdD     As String = "Idle Detection"

    If dwState = USER_IDLE_BEGIN Then
        Sound1 = vbQuestion
        Sound2 = vbInformation
        For Each Proj In VBInstance.VBProjects
            PName = Proj.Name & "."
            For Each Compo In Proj.VBComponents
                With Compo
                    If .IsDirty Then
                        MsgText = "You may want to save " & PName & .Name & " now that you've been idle for more than " & LCase$(GetNumText(Idletime)) & " minute" & IIf(Idletime = 1, ".", "s.")
                        If Len(.FileNames(1)) Then
                            If MsgBoxEx(MsgText, vbQuestion Or vbOKCancel, IdD, -1, OCapt:="800|801", NCapt:="&Save|&Later", Sound:=Sound1) = vbOK Then
                                If .SaveAs(.FileNames(1)) Then
                                    MsgBoxEx "Saved " & PName & .Name & " in " & .FileNames(1), vbInformation, "Success", TimeOut:=1500
                                  Else '.SAVEAS(.FILENAMES(1)) = FALSE/0
                                    MsgBoxEx "Could not save " & PName & .Name, vbCritical, "Failed", TimeOut:=-1
                                End If
                            End If
                          Else 'LEN(.FILENAMES(1)) = FALSE/0
                            MsgBoxEx MsgText, , IdD, -1, OCapt:="800", NCapt:="&Acknowleged", Sound:=Sound2
                        End If
                        Sound1 = 0
                        Sound2 = 0
                    End If
                End With 'COMPO
        Next Compo, Proj
    End If

End Sub

Public Sub IdleStopDetection()

    EndIdleDetection 0&

End Sub

Public Function Inc(ByRef What As Long, Optional ByVal By As Long = 1) As Long

    What = What + By
    Inc = What

End Function

Private Function IsInCode(Line As String, Posn As Long) As Boolean

  'this checks that the current word (at Posn) is not in a literal or in a comment
  'and is not preceded by the Dim-, the Const- or the As-keyword

  Dim Ptr       As Long
  Dim Cnt       As Long
  Dim Tmp       As String

    IsInCode = True
    Ptr = Posn
    Do Until Ptr = 0
        Ptr = InStrRev(Line, """", Ptr)
        If Ptr Then
            Inc Cnt 'count the number of quotes preceeding the current word
            Dec Ptr 'move pointer left before quote
        End If
    Loop
    If Cnt And 1 Then 'odd - we're inside a literal
        IsInCode = False
      Else 'NOT CNT...
        'see if an Apostophe or Rem, or the As/Dim/Const keywords preceed the current word
        Ptr = InStrRev(Line, "'", Posn)
        If Ptr = 0 Then
            If InStrRev(Line, " AS ", Posn) = 0 Then
                Ptr = InStrRev(Line, " as ", Posn, vbTextCompare) 'bug fix - was looking into Tmp
                If Ptr = 0 Then
                    Tmp = Spce & Line
                    Ptr = InStrRev(Tmp, " dim ", Posn, vbTextCompare)
                    If Ptr = 0 Then
                        Ptr = InStrRev(Tmp, " const ", Posn, vbTextCompare)
                        If Ptr = 0 Then
                            Ptr = InStrRev(Tmp, " rem ", Posn, vbTextCompare)
                        End If
                    End If
                End If
            End If
        End If
        If Ptr Then
            IsInCode = False
        End If
    End If

End Function

Public Function MakeMoveable(hWnd As Long, NewValue As Boolean) As Boolean

  'not yet implemented; I don't know where this bloody style bit is (if! it is a style bit)

  Dim i0 As Long 'dummy code

    i0 = 0 'dummy code

End Function

Public Function MakePoint(ByVal X As Long, ByVal Y As Long) As tPOINT

  'little helper - simply moves x and y coords into a point structure

    With MakePoint
        .X = X
        .Y = Y
    End With 'MAKEPOINT

End Function

Private Function MakeRect(ByVal l As Long, ByVal t As Long, ByVal r As Long, ByVal b As Long) As tRECT

  'little helper - simply moves topleft and bottomright points into a rect structure

    With MakeRect
        .TopLeft = MakePoint(l, t)
        .BottomRight = MakePoint(r, b)
    End With 'MAKERECT

End Function

Public Function MakeSizeable(hWnd As Long, NewValue As Boolean) As Boolean

  Dim OldStyle As Long

    OldStyle = GetWindowLong(hWnd, IDX_STYLE)
    If NewValue Then
        SetWindowLong hWnd, IDX_STYLE, OldStyle Or WS_THICKFRAME
      Else 'NEWVALUE = FALSE/0
        SetWindowLong hWnd, IDX_STYLE, OldStyle And Not WS_THICKFRAME
    End If
    SetWindowPos hWnd, 0&, 0&, 0&, 0&, 0&, SWP_AFTERFRAMECHANGE
    MakeSizeable = Not CBool(OldStyle And WS_THICKFRAME) 'return previous value

End Function

Private Function MDIClientProc(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  'Window procedure for VB's MDIClient window

  Dim TopLn     As Long
  Dim NumLines  As Long
  Dim Delta     As Long

    MDIClientProc = CallWindowProc(GetProp(hWnd, PropName), hWnd, nMsg, wParam, lParam) 'call the original winproc to do what has to be done
    With VBInstance
        'bug fix: we have to find out 1st if there's a code pane at all
        If Not .ActiveCodePane Is Nothing Then 'we have a codepane to scroll in
            Select Case nMsg 'and now split on message type
              Case WM_KILLFOCUS 'this codepane just lost the focus (remember - the original procedure has already been performed)
                UnhookCodePane
              Case WM_MDIACTIVATE, WM_MOUSEACTIVATE, WM_SETFOCUS 'another codepane has been (re)activated by the user
                HookCodePane
              Case WM_PAINT
                BitBlt hDCMDIClient, 0&, 0&, wIcon, hIcon, hDCIcon, 0&, 0&, vbSrcCopy 'refresh icon
              Case WM_MOUSEWHEEL 'hah!!! there it is - the user fingers the mouse wheel
                If wParam And SecondaryButton Then 'secondary (right) mouse button is down while scrolling
                    fSetOpts.Show vbModal 'show options dialog
                    Unload fSetOpts
                    .ActiveCodePane.Window.SetFocus
                  Else 'NOT WPARAM...
                    With .ActiveCodePane
                        If .CodeModule.CountOfLines Then 'codepane is not empty
                            'translate mousewheel and pressed key (Shift or Cntl)
                            If IsScrolling Then
                                TopLn = ScrollTo 'accumulate mouse scrolls (start from where it WILL be)
                              Else 'ISSCROLLING = FALSE/0
                                TopLn = .TopLine 'start from where it is right now
                            End If
                            Select Case NumLinesToScroll
                              Case sHalfPage 'half page (in fact fraction of page, see ScrollFraction)
                                NumLines = Int(.CountOfVisibleLines * ScrollFraction)
                              Case sFullPage
                                NumLines = .CountOfVisibleLines - 1 'so that the bottom line is at the top after scrolling and vice versa
                              Case Else
                                NumLines = Abs(Val(NumLinesToScroll))
                                If NumLines >= .CountOfVisibleLines Then 'not more than a page
                                    NumLines = .CountOfVisibleLines - 1
                                End If
                            End Select
                            ScrollTo = TopLn - Sgn(wParam) * NumLines / ((wParam And &HFFFF&) \ 4 + 1) 'compute new top line
                            If ScrollTo = TopLn Then 'zero lines to scroll (probably a very small window)
                                ScrollTo = TopLn - Sgn(wParam) 'make it one line
                            End If
                            With .CodeModule
                                Select Case ScrollTo 'correct it if it is out of range
                                  Case Is < 1
                                    ScrollTo = 1
                                  Case Is > .CountOfLines
                                    ScrollTo = .CountOfLines
                                End Select
                            End With '.CODEMODULE
                            If Smooth Then
                                If Not IsScrolling Then
                                    IsScrolling = True
                                    ShowCursor False
                                    Do
                                        Delta = ScrollTo - TopLn
                                        If Abs(Delta) < 2 Then
                                            Exit Do 'loop 
                                        End If
                                        TopLn = TopLn + Sgn(Delta) * (Abs(Delta) \ NumLines + 1) 'speeding up on distance
                                        .TopLine = TopLn
                                        DoEvents
                                        Wait ScrollDelay
                                    Loop
                                    .TopLine = ScrollTo 'make sure it is where it ought to be
                                    ShowCursor True
                                    IsScrolling = False
                                End If
                              Else 'SMOOTH = FALSE/0
                                .TopLine = ScrollTo
                            End If 'smooth
                            DrawRaster
                        End If 'codepane is not empty
                    End With '.ACTIVECODEPANE
                End If 'secondary button
            End Select 'split on msg type
        End If 'we have a code pane
    End With 'VBINSTANCE

End Function

Public Sub MoveToArrow(Frm As Form, Optional Top As Long, Optional Lft As Long, Optional Bot As Long, Optional Rgt As Long)

  '¤  maximize and readjust the active codepane
  '¤  move the From next to caret and centered vertically if possible
  '¤  create the arrow pointing to the caret and flash leaving it on

  Const Margin  As Long = 5 'pixels - prevent Me  from being placed directly at the screen borders
  Dim MarginX   As Long
  Dim MarginY   As Long
  ' Dim Top       As Long
  ' Dim Lft       As Long
  ' Dim Bot       As Long
  ' Dim Rgt       As Long

    With VBInstance.ActiveCodePane
        .Window.WindowState = vbext_ws_Maximize
        .CodeModule.CodePane.GetSelection Top, Lft, Bot, Rgt
        .CodeModule.CodePane.SetSelection Top, Lft, Bot, Rgt
        DoEvents
        GetWindowRect hWndCodePane, WindowRect 'where's the codepane window
        With CursorPos
            If GetCaretPos(CursorPos) = False Then
                .X = (WindowRect.BottomRight.X - Frm.ScaleWidth) / 2 'x-center
            End If
            .X = (.X + WindowRect.TopLeft.X + ArrowVertices(1).X + 6) * Screen.TwipsPerPixelX 'adjust to twips
            .Y = (.Y + WindowRect.TopLeft.Y + 30) * Screen.TwipsPerPixelY - Frm.Height / 2
            MarginX = Margin * Screen.TwipsPerPixelX
            MarginY = Margin * Screen.TwipsPerPixelY
            Select Case True 'limit x to be within screen
              Case .X < MarginX
                .X = MarginX
              Case .X + Frm.Width > Screen.Width - MarginX
                .X = Screen.Width - Frm.Width - MarginX
            End Select
            Select Case True 'limit y to be within screen
              Case .Y < MarginY
                .Y = MarginY
              Case .Y + Frm.Height > Screen.Height - MarginY
                .Y = Screen.Height - Frm.Height - MarginY
            End Select
            Frm.Move .X, .Y 'move Me to that position
        End With 'CURSORPOS
    End With 'VBINSTANCE.ACTIVECODEPANE
    CaretRgn Arrow Or Flash, 11 'odd - flash and leave

End Sub

Private Function NotAnApiLine(Line As String) As Boolean

  'returns true when the line is not an api declaration line nor an api const nor an api type

    NotAnApiLine = ((InStr(1, Line, sApiConst, vbTextCompare) Or InStr(1, Line, sApiDeclare, vbTextCompare) Or InStr(1, Line, sApiType, vbTextCompare)) = 0)

End Function

Public Function OffsetVertex(Vertex As tPOINT, By As tPOINT) As tPOINT

  'offset a point

    With OffsetVertex
        .X = Vertex.X + By.X
        .Y = Vertex.Y + By.Y + RegionCenterY
    End With 'OFFSETVERTEX

End Function

Public Sub Pen(ByVal Action As PenAction)

    If Action And DestroyIt Then
        If hGridPen Then
            SelectObject hDCCodePane, hPrevPen 're-select previous object
            DeleteObject hGridPen 'delete the pen
            hGridPen = 0
        End If
    End If
    If Action And CreateIt Then
        If hGridPen = 0 Then
            hGridPen = CreatePen(LineStyle, 1, GridColor) 'create a new pen
            hPrevPen = SelectObject(hDCCodePane, hGridPen) 'select pen into code pane device context
        End If
    End If

End Sub

Private Sub PollDirty()

    On Error Resume Next 'there may not be an active component or the user may have removed the button
        If ResetMenuButton.Visible Then
            If ActiveCompo.IsDirty Then
                If Not IsGreen Then
                    SetMenuIcon ResetMenuButton, fIcon.picMenuResetGreen
                    IsGreen = True
                End If
              Else 'ACTIVECOMPO.ISDIRTY = FALSE/0
                If IsGreen Then
                    SetMenuIcon ResetMenuButton, fIcon.picMenuResetRed
                    IsGreen = False
                End If
            End If
        End If
    On Error GoTo 0

End Sub

Public Sub RedrawArrow()

    CaretRgn Arrow Or Inval
    DoEvents
    CaretRgn Arrow Or Flash, -1 'without finding the caret (we know where that was) and without flash delay

End Sub

Private Sub RefreshMembers(NeedsRefresh As Boolean)

  'Copy all members to our own collection

  'This is necessary because VB apparently verifies and expands/compacts the current active line
  'thru access to it's Members collection. Originally this search was done in direct response to
  'the WM_CHAR message in the codepane proc, which resulted in the line being expanded/compacted
  'while typing. So now we build our own collection after VB has processed the current line anyway
  '(see CodePaneProc WM_KEYDOWN). Unfortunately this is quite a bit of an overhead, but it's mini-
  'mized because we only do it after the user really typed some code.

  Dim Compo     As VBComponent
  Dim Memb      As Member

    If NeedsRefresh Then
        Set CodeMembers = New Collection
        For Each Compo In VBInstance.ActiveVBProject.VBComponents
            With Compo
                If .Type <> vbext_ct_RelatedDocument And .Type <> vbext_ct_ResFile Then
                    CodeMembers.Add Array(.Name, .Name, NoScope, .Type, "CMF")
                    For Each Memb In .CodeModule.Members
                        With Memb
                            CodeMembers.Add Array(Compo.Name, .Name, .Scope, .Type) 'arrayindex=0, 1, 2, 3
                        End With 'MEMB
                    Next Memb
                End If
            End With 'COMPO
        Next Compo
    End If

End Sub

Public Sub SendMeMail(FromhWnd As Long, Subject As String)
Attribute SendMeMail.VB_Description = "What it says: opens mail prog."

  Dim UserName  As String
  Dim Lng       As Long

    Lng = 128
    UserName = String$(Lng, 0)
    GetUserName UserName, Lng
    UserName = Left$(UserName, Lng + (Asc(Mid$(UserName, Lng, 1)) = 0))
    If ShellExecute(FromhWnd, vbNullString, "mailto:UMGEDV@Yahoo.com?subject=" & Subject & " &body=Hi Ulli,<br><br>[your message]<br><br>Best regards from " & UserName, vbNullString, App.Path, SW_SHOWNORMAL) < SE_NO_ERROR Then
        Beep
        MsgBoxEx "Cannot send Mail from this System.", Title:="Mail disabled/not installed", PosX:=-2, Icon:=fIcon.Icon, OCapt:=OK, NCapt:="&Close"
    End If

End Sub

Public Sub SetMenuIcon(MenuButton As Office.CommandBarButton, Pic As PictureBox)

  Dim TmpStr As String

    With Clipboard
        TmpStr = .GetText
        .SetData Pic.Image
        MenuButton.PasteFace
        .Clear
        .SetText TmpStr
    End With 'CLIPBOARD

End Sub

Private Function SpaceReplace(Line As String) As String

  'replace some special chars in a line of code by space

  Dim i             As Long

    SpaceReplace = Line
    For i = 1 To Len(CharsToSpace)
        SpaceReplace = Replace$(SpaceReplace, Mid$(CharsToSpace, i, 1), Spce)
    Next i

End Function

Public Sub UnhookCodePane()

    If hWndCodePane Then
        Pen DestroyIt
        EraseRaster
        BitBlt hDCMDIClient, 0&, 0&, wIcon, hIcon, hDCIcon, wIcon, 0&, vbSrcCopy 'erase icon
        ReleaseDC hWndCodePane, hDCCodePane 'release device context
        SetWindowLong hWndCodePane, IDX_WINDOWPROC, GetProp(hWndCodePane, PropName) 'reset window proc ptr
        RemoveProp hWndCodePane, PropName 'remove the property
        hWndCodePane = 0
        hWndScrollbar = 0
        On Error Resume Next 'the user may have removed the button
            ResetMenuButton.Enabled = False
            ResetMenuButton.ToolTipText = NAC
            CompareMenuButton.Enabled = False
            CompareMenuButton.ToolTipText = NAC
            CopyMenuButton.Enabled = False
            CopyMenuButton.ToolTipText = NAC
        On Error GoTo 0
    End If

End Sub

Public Sub UnhookMDIClient()
Attribute UnhookMDIClient.VB_Description = "Restore the original window procedure address and remove our property."

    If hWndMDIClient Then
        DirtyTimer DestroyIt
        BitBlt hDCMDIClient, 0&, 0&, wIcon, hIcon, hDCIcon, wIcon, 0&, vbSrcCopy 'erase icon
        ReleaseDC hWndMDIClient, hDCMDIClient 'release device context
        UnhookCodePane
        SetWindowLong hWndMDIClient, IDX_WINDOWPROC, GetProp(hWndMDIClient, PropName)
        RemoveProp hWndMDIClient, PropName 'remove the property
        hWndMDIClient = 0
    End If

End Sub

Private Sub UpdateTooltips()

    On Error Resume Next 'the user may have removed the button
        With ActiveCompo
            ResetMenuButton.Enabled = True
            ResetMenuButton.ToolTipText = "Reset " & .Name & " to previous state"
            CompareMenuButton.Enabled = True
            CompareMenuButton.ToolTipText = "Compare " & .Name & " with previous state"
            CopyMenuButton.Enabled = True
            CopyMenuButton.ToolTipText = "Copy code into " & .Name & " at caret position"
        End With 'ACTIVECOMPO
    On Error GoTo 0

End Sub

Private Sub Wait(Secs As Single)

    If HiResTimerPresent Then
        'high resolution timing functions
        QueryPerformanceCounter CPUTicksStart
        Do
            QueryPerformanceCounter CPUTicksNow
        Loop Until (CPUTicksNow - CPUTicksStart) / CPUFreq > Secs
      Else 'HIRESTIMERPRESENT = FALSE/0
        Sleep Secs * 1000
    End If

End Sub

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 16:14)  Decl: 364  Code: 1060  Total: 1424 Lines
':) CommentOnly: 82 (5,8%)  Commented: 252 (17,7%)  Empty: 188 (13,2%)  Max Logic Depth: 19
