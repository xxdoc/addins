Attribute VB_Name = "Modul"
' ctrl-B        : find brace opposite and find block command beginning and ending
' ctrl-shift-B  : select ctrl-B area
'
' Braces    : () [] {}
'
' Blocks    : IF-END IF,FOR-NEXT,WITH-END WITH,DO WHILE-LOOP,WHILE-WEND,SELECT-END SELECT
'
' Author: HASAN AYDIN
' 07/01/2008
' hassanaydin@gmail.com
'
' Thanks:   Google      : What can we do if not exist (return to BBS)
'           Hotkey      : How to Define ctrl-b
'           VBCompanion : DrawLine on codewindow

Option Explicit

Private Const GWL_WNDPROC = -4
Public Const WM_HOTKEY = &H312
Private m_hkCount As Long
Public oldProc As Long
Public fMain As frmAddIn


Public Enum ModConst
    MOD_ALT = &H1
    MOD_CONTROL = &H2
    MOD_SHIFT = &H4
End Enum
    
Public startLine As Long, startCol As Long
Public endLine As Long, endCol As Long
Public st As String
Public i As Integer
Public j As Integer
Public satir As Integer
Public sutun As Integer
Public ch As String
Public ch2 As String
Public parantez As String
Public adet As Integer
Public countoflines
Public kolon As Long
Public wParam As Long
     
'vbcompanion
Public hDCCodePane   As Long
Public HorizPosn     As Long
Public VertPosn      As Long
Public HorizOffset   As Long
Public RasterBottom  As Long
Public PreviousROP   As Long
Public hWndCodePane                 As Long
Public LineHeight                   As Long
Public hGridPen                    As Long
Public LineStyle                    As Long
Public GridColor                    As Long
Public hPrevPen                    As Long
Public IDEFontName                  As String
Public IDEFontSize                  As Long
Public DataLength                  As Long
Public CharWidth                   As Long
Public RegHandle                   As Long
Public Const sFontface             As String = "FontFace"
Public Const REG_OPTION_RESERVED   As Long = 0
Public DataType                    As Long
Public Const ERROR_NONE            As Long = 0
Public Const DefaultFontName       As String = "Fixedsys"
Public Const sFontheight           As String = "FontHeight"
Public Const DefaultFontSize       As Long = 9
Public Enum PenAction
    DestroyIt = 1
    CreateIt = 2
    DestroyItCreateIt = DestroyIt Or CreateIt
End Enum
Public Enum RasterMargins 'in pixels
    RasterTop = 30
    RasterStartPos = 13 'without the indicator bar which is 21 pixels wide
    IndicatorBarWidth = 21
End Enum
Public Type tPOINT
    X       As Long
    Y       As Long
End Type
Public Type tRECT 'defined by top left and bottom right corner points
    TopLeft     As tPOINT
    BottomRight As tPOINT
End Type
Public WindowRect                   As tRECT

Public ArrowVertices(1 To 4)        As tPOINT
Public RegionVertices(1 To 4)       As tPOINT
Public RegionCenterY                As Long
Public CaretPos                     As tPOINT
Public Const WINDING                As Long = 2
Public hWndMDIClient                As Long

Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DrawLine Lib "gdi32" Alias "LineTo" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPenPosition Lib "gdi32" Alias "MoveToEx" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As tRECT) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As tPOINT) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As tPOINT, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function InvalidateRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function InvertRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long


Sub Main()
    oldProc = SetWindowLongA(frmAddIn.hWnd, GWL_WNDPROC, AddressOf WndProc)
    HotKeyActivate frmAddIn.hWnd, MOD_CONTROL, Asc("B")                 'CTRL+B
    HotKeyActivate frmAddIn.hWnd, MOD_CONTROL + MOD_SHIFT, Asc("B")     'CHIFT+CTRL+B
    'HotKeyActivate frmAddIn.hWnd, 0, vbKeyRight
End Sub

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal gparam As Long, ByVal lParam As Long) As Long
    
    'wParam - the number of the hotkey, its identification.
    'lParam - HiWord is the Modifiere e.g. Shift, Ctrl, Alt
    'lParam . LoWord is the KeyCode, it is the same Code found in the Objectbrowser (F2)
    'under KeyCode
    'but I think you need only the number (identifier) of the hotkey, given in wParam.
    wParam = gparam
    
'    Debug.Print wParam, lParam
    WndProc = 0
    If uMsg = WM_HOTKEY Then
        'The Hotkey message
        'If wParam = 1 Then
        PARANTEZ_KONTROL
        'End If
    Else
        'All other messages to the old Windowprocedure
        WndProc = CallWindowProcA(oldProc, hWnd, uMsg, wParam, lParam)
    End If
End Function

Function HotKeyActivate(ByVal hWnd As Long, Modifier As ModConst, Optional KeyCode As Integer) As Long
    
    m_hkCount = m_hkCount + 1
    
    ' 0 for no success, otherwise success
    HotKeyActivate = RegisterHotKey(hWnd, m_hkCount, Modifier, KeyCode)
    
End Function

Function HotKeyDeactivate(ByVal hWnd As Long)
    Dim i As Integer
    For i = 1 To m_hkCount
        UnregisterHotKey hWnd, i
    Next i
    m_hkCount = 0
End Function

Sub PARANTEZ_KONTROL() 'wparam  1:ctrl-b   2:shift+ctrl+b
    Dim curline As Long
    

    hWndMDIClient = FindWindowEx(frmAddIn.VBInstance.MainWindow.hWnd, 0&, "MDIClient", vbNullString)
    hWndCodePane = FindWindowEx(hWndMDIClient, 0&, "VbaWindow", frmAddIn.VBInstance.ActiveWindow.Caption)
    
    hDCCodePane = GetDC(hWndCodePane) 'get codepane device context

    'MsgBox startCol & ":" & CharWidth & ":" & curline & ":" & LineHeight

    'SetPenPosition hDCCodePane, startCol * CharWidth, curline * LineHeight, ByVal 0&          'move pen to top
    'DrawLine hDCCodePane, startCol * CharWidth, curline * LineHeight - LineHeight     'draw line to bottom
    
     
     
    On Error Resume Next
    
    frmAddIn.VBInstance.ActiveCodePane.Show
    frmAddIn.VBInstance.ActiveCodePane.CodeModule.CodePane.GetSelection startLine, startCol, endLine, endCol
    
    curline = startLine - frmAddIn.VBInstance.ActiveCodePane.CodeModule.CodePane.TopLine
     
     
     countoflines = frmAddIn.VBInstance.ActiveCodePane.CodeModule.countoflines
     
     parantez = "(){}[]"
   

    
    st = frmAddIn.VBInstance.ActiveCodePane.CodeModule.Lines(startLine, 1)
    
                            
    If Err Then MsgBox "Hata var." & vbCrLf & "Error.": Exit Sub
       
    j = startCol
    kolon = j
    ch = Mid(st, j, 1)  'kursorun saðýndaki iþaret       right of the cursor's sign
    adet = 0
    
    i = InStr(parantez, ch)
    If i > 0 Then
        ch2 = Mid(parantez, i, 1)
        'parantez açma ise kapamayý bul.  Finding closer brace
        If i Mod 2 = 1 Then
            For satir = startLine To countoflines
                st = frmAddIn.VBInstance.ActiveCodePane.CodeModule.Lines(satir, 1)
                
devam:
                If Mid(st, j, 1) = ch2 Then adet = adet + 1
                If Mid(st, j, 1) = Mid(parantez, i + 1, 1) Then adet = adet - 1
                If adet = 0 Then
                    If wParam = 1 Then
                        frmAddIn.VBInstance.ActiveCodePane.SetSelection satir, j, satir, j
                    End If
                    If wParam = 2 Then
                        frmAddIn.VBInstance.ActiveCodePane.SetSelection startLine, startCol, satir, j + 1
                    End If
                    ok_goster
                    Exit Sub
                End If
                j = j + 1
                If j <= Len(st) Then GoTo devam
                j = 1
            Next
            MsgBox "Parantez kapatýlmamýþ." & vbCrLf & "Brace isn't closed"
            Exit Sub
        End If
        'parantez kapama ise açmayý bul. Finding opener brace
        If i Mod 2 = 0 Then
            For satir = startLine To 1 Step -1
                st = frmAddIn.VBInstance.ActiveCodePane.CodeModule.Lines(satir, 1)
                If j = 0 Then j = Len(st)
devam2:
                If Mid(st, j, 1) = ch2 Then adet = adet + 1
                If Mid(st, j, 1) = Mid(parantez, i - 1, 1) Then adet = adet - 1
                If adet = 0 Then
                    If wParam = 1 Then
                        frmAddIn.VBInstance.ActiveCodePane.SetSelection satir, j, satir, j
                    End If
                    If wParam = 2 Then
                        frmAddIn.VBInstance.ActiveCodePane.SetSelection startLine, startCol + 1, satir, j
                    End If
                    ok_goster
                    Exit Sub
                End If
                j = j - 1
                If j > 1 Then GoTo devam2
            Next
            MsgBox "Parantez açýlmamýþ." & "Brace isn't opened."
            Exit Sub
        End If
    End If
    
    KONTROL2 "If", "End If", ""
    KONTROL2 "For", "Next", "Exit For", "Resume Next"
    KONTROL2 "With", "End With", ""
    KONTROL2 "Do While", "Loop", ""
    KONTROL2 "Do Until", "Loop", ""
    KONTROL2 "While", "Wend", ""
    KONTROL2 "Select Case", "End Select", ""
    KONTROL2 "Sub", "End Sub", "Exit Sub"
    KONTROL2 "Public Sub", "End Sub", "Exit Sub"
    KONTROL2 "Private Sub", "End Sub", "Exit Sub"
    KONTROL2 "Function", "End Function", "Exit Function"
    KONTROL2 "Public Function", "End Function", "Exit Function"
    KONTROL2 "Private Function", "End Function", "Exit Function"
    KONTROL2 "Enum", "End Enum"
    KONTROL2 "Public Enum", "End Enum"
    KONTROL2 "Private Enum", "End Enum"
    
'cikis: If wParam = 3 Then
'          SendKeys "{right}"
'          ok_goster
'       End If
End Sub

Sub KONTROL2(baslama As String, bitis As String, ParamArray gozardi())
    'baslama:start of command
    'bitis  :end of command
    'gozardi: exclude text   (For  ->  Exit For)
    
    Dim st2 As String
    Dim i2 As Integer
    
    Dim gozardi2() As Variant
    
    adet = 0
    gozardi2 = gozardi
    
    'yukarýdan aþaðýya
    'up to down
    ch = Mid(st & " ", kolon, Len(baslama) + 1)
    If ch = baslama & " " Then
        For satir = startLine To countoflines
            st2 = frmAddIn.VBInstance.ActiveCodePane.CodeModule.Lines(satir, 1)
            st2 = Remove_StrAndComment(st2, gozardi2())
            i2 = InStr(" " & st2 & " ", " " & bitis & " ")
            If InStr(" " & st2 & " ", " " & baslama & " ") > 0 And i2 = 0 Then adet = adet + 1
            If i2 > 0 Then adet = adet - 1
            
            If baslama = "If" Then
                If InStr(st2, "Then ") > 0 Then
                
                    'Then'den sonra kod varsa If kapanmýþ demektir.
                    'if there is code after Then 'If' is closed
                    If Trim(Mid(st2, InStr(st2, "Then ") + 4)) > "" Then adet = adet - 1
                    
                End If
            End If
            If adet = 0 Then
                If wParam = 1 Then
                    frmAddIn.VBInstance.ActiveCodePane.SetSelection satir, i2, satir, i2
                   Else
                    frmAddIn.VBInstance.ActiveCodePane.SetSelection startLine, 1, satir, 255
                End If
                Exit Sub
            End If
        Next
        frmAddIn.VBInstance.ActiveCodePane.CodeModule.VBE.de
        'MsgBox baslama & "-" & bitis & " bloðu kapatýlmamýþ." & vbCrLf & "Block isn't closed."
        Exit Sub
    End If
    
    'aþaðýdan yukarýya
    'down to up
    ch = Mid(st & " ", kolon, Len(bitis) + 1)
    If ch = bitis & " " Then
        For satir = startLine To 1 Step -1
            st2 = frmAddIn.VBInstance.ActiveCodePane.CodeModule.Lines(satir, 1)
            st2 = Remove_StrAndComment(st2, gozardi2())
            i2 = InStr(" " & st2 & " ", " " & baslama & " ")
            If InStr(" " & st2 & " ", " " & bitis & " ") > 0 Then adet = adet + 1
            If i2 > 0 And InStr(" " & st2 & " ", " " & bitis & " ") = 0 Then adet = adet - 1
            
            If baslama = "If" Then
                If InStr(st2, "Then ") > 0 Then
                    'Then'den sonra kod varsa If kapanmýþ demektir.
                    'if there is code after 'Then' 'If' is closed
                    If Trim(Mid(st2, InStr(st2, "Then ") + 4)) > "" Then adet = adet + 1
                End If
            End If
            If adet = 0 Then
                If wParam = 1 Then
                    frmAddIn.VBInstance.ActiveCodePane.SetSelection satir, i2, satir, i2
                   Else
                    frmAddIn.VBInstance.ActiveCodePane.SetSelection startLine, 255, satir, 1
                End If
                Exit Sub
            End If
        Next
        Debug.Print baslama & "-" & bitis & " bloðu Açýlmamýþ." & vbCrLf & "Block isn't opened."
        Exit Sub
    End If
    
End Sub

Function Remove_StrAndComment(X As String, gozardi() As Variant) As String
    Dim z As Integer
    Dim z2 As Integer
    Remove_StrAndComment = X
    
    'if there isn't string
    If InStr(X, Chr(34)) = 0 Then GoTo devam
    
    'replace double quate
    While InStr(X, Chr(34) & Chr(34)) > 0
        X = Replace(X, Chr(34) & Chr(34), "")
    Wend
    
    'remove inner string
tekrar:
    z = InStr(X, Chr(34))
    z2 = InStr(z + 1, X, Chr(34))
    If z * z2 = 0 Then GoTo devam
    X = Left(X, z - 1) & Mid(X, z2 + 1)
    GoTo tekrar
    
devam:
    'remove comment
    z = InStr(X, "'")
    If z Then X = Left(X, z - 1)
    
    'gözardi edilecek kelimeleri sil
    'remove exclude text
    For z = 0 To UBound(gozardi())
        X = Replace(X, gozardi(z), "")
    Next
    Remove_StrAndComment = X
    
End Function





Public Sub TestArea()
   'Test code
    For i = 1 To 100
        For j = 1 To 50
             i = " ""  """"return Next " ' FOR Next
             i = "(2*(45+45)" & "/(45)/45)"
             i = (2 * (45 + 45 / (12 / (58.45) * 23 + 23)))
             If j = 1 Then
                i = 12
                Resume Next
               Else
                i = 12
                Exit For
             End If
             Do While i = 2
                While i = 0
                    Do Until i = 2
                        'Do While'da hata var
                    Loop
                Wend
             Loop
        Next
    Next

End Sub



Public Sub Pen(ByVal Action As PenAction)

    If Action And CreateIt Then
        If hGridPen = 0 Then
            hGridPen = CreatePen(LineStyle, 1, GridColor) 'create a new pen
            hPrevPen = SelectObject(hDCCodePane, hGridPen) 'select pen into code pane device context
        End If
    End If

End Sub

Public Function MakePoint(ByVal X As Long, ByVal Y As Long) As tPOINT

  'little helper - simply moves x and y coords into a point structure

    With MakePoint
        .X = X
        .Y = Y
    End With 'MAKEPOINT

End Function

Public Function OffsetVertex(Vertex As tPOINT, By As tPOINT) As tPOINT

  'offset a point

    With OffsetVertex
        .X = Vertex.X + By.X
        .Y = Vertex.Y + By.Y + RegionCenterY
    End With 'OFFSETVERTEX

End Function


Public Sub ok_goster()
    Dim hRgn          As Long
    GetCaretPos CaretPos
    CaretPos.X = CaretPos.X '+ Int(frmAddIn.TextWidth("T") / 2)
    CaretPos.Y = CaretPos.Y + 2 '+ frmAddIn.TextHeight("T") / 2
    For hRgn = 1 To UBound(ArrowVertices)
         RegionVertices(hRgn) = OffsetVertex(ArrowVertices(hRgn), CaretPos)
    Next hRgn
    hRgn = CreatePolygonRgn(RegionVertices(1), hRgn - 1, WINDING) 'and now hRgn has the handle to the region
    InvertRgn hDCCodePane, hRgn
    DeleteObject hRgn 'and finally delete the region

End Sub
