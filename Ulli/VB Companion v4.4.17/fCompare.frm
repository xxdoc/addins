VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form fCompare 
   BackColor       =   &H00E0E0E0&
   Caption         =   "File Compare"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   Icon            =   "fCompare.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximiert
   Begin VB.CommandButton btSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Save compare results"
      Top             =   30
      Width           =   810
   End
   Begin VB.CommandButton btColorize 
      Caption         =   "&Colorize"
      Enabled         =   0   'False
      Height          =   315
      Left            =   30
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Color syntax"
      Top             =   30
      Width           =   810
   End
   Begin RichTextLib.RichTextBox rtBox 
      Height          =   2760
      Left            =   990
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4868
      _Version        =   393217
      BackColor       =   16580607
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      MousePointer    =   1
      DisableNoScroll =   -1  'True
      RightMargin     =   9999
      TextRTF         =   $"fCompare.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrTooltip 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   4500
      Top             =   3570
   End
   Begin VB.CommandButton btFindNextDiff 
      Caption         =   "&Find next "
      Height          =   315
      Left            =   855
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Find next difference"
      Top             =   30
      Width           =   810
   End
   Begin VB.Label lbMessage 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Left            =   1230
      TabIndex        =   3
      Top             =   3615
      Width           =   3165
   End
   Begin VB.Label lbEqual 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2595
      TabIndex        =   2
      Top             =   45
      Width           =   4740
   End
   Begin VB.Label lbBar 
      BorderStyle     =   1  'Fest Einfach
      Height          =   6705
      Left            =   15
      TabIndex        =   4
      Top             =   720
      Width           =   285
   End
End
Attribute VB_Name = "fCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Comparator
'compares two textfiles line by line
'synchronization is by content rather than by line  numbers

Option Explicit

'###########################################################################
#Const Wingdings = True 'marker font switch
'           (set to False if you don't have font Wingdings 3 installed
'            and it will use Arial instead; Wingdings 3 looks better though)
'###########################################################################

DefLng A-Z  'we're 32 bit!

Private Const EM_SCROLL     As Long = &HB5
Private Const EM_GETFIRSTVISIBLELINE    As Long = &HCE

Private Const SB_LINEUP     As Long = 0
Private Const SB_LINEDOWN   As Long = 1

Private Digest              As cMD5
Private hFileOld            As Long
Private WhereInOld          As Long
Private hFileNew            As Long
Private WhereInNew          As Long
Private i                   As Long
Private LenFiles            As Long
Private LenBoxText          As Long
Private DoneSoFar           As Long
Private DiffPointer         As Long 'used to scan through the text to find differences
Private SkipToNext          As Long
Private PrevStep            As Long
Private NumDiffs            As Long
Private NumNonDiffs         As Long
Private ProgressStep        As Single
Private CorrIfAtt           As Boolean 'controls the position for the first divider line
Private DrawSep             As Boolean

Private LastDeclSig         As String
Private Textline            As String
Private Sig                 As String
Private SigsOld             As String
Private SigsNew             As String
Private BoxText             As String
Private LenSig              As Long
Private DisregardCase       As Boolean

Private Const StepChar      As String = "" 'used in progressbar
Private Const LenSep        As Long = 1023

Private Const InsColor      As Long = &HC000&     'RGB(0, 196, 0)       used for insertions
Private Const DelColor      As Long = vbRed       'RGB(0, 0, 255)       used for deletions

Private SyntaxColor1        As Long 'used for reserved words
Private SyntaxColor2        As Long 'used for comments
Private SyntaxColor3        As Long 'not used yet; reserved for literals and such
Private DefForeColor        As Long 'used for the remainder

Private RTFSyntaxColor1     As String 'RTF-syntaxed color
Private RTFSyntaxColor2     As String 'RTF-syntaxed color
Private RTFSyntaxColor3     As String 'RTF-syntaxed color

#If Wingdings Then
Private Const MarkerFont    As String = "Wingdings 3"
Private Const MarkerBold    As Boolean = False
Private Const MarkerSize    As Long = 10
'both chars MUST be below Asc = (128)
Private Const ArrowOut      As String = "t"
Private Const ArrowIn       As String = "u"
#Else
Rem Mark Off Silent
Private Const MarkerFont    As String = "Arial"
Private Const MarkerBold    As Boolean = True
Private Const MarkerSize    As Long = 14
'both chars MUST be below Asc = (128)
Private Const ArrowOut      As String = "<"
Private Const ArrowIn       As String = ">"
Rem Mark On
#End If

Private Const Delimiters    As String = " %&@!#$():,"

Private Sub btColorize_Click()

    btColorize.Enabled = False 'just once is enough
    Screen.MousePointer = vbHourglass
    With rtBox
        .Visible = False 'prevent flicker thru updates, meanwhile show label underneath
        .TextRTF = Colorize(.TextRTF) 'colorize the rtf-text
        .Visible = True
        .SetFocus
    End With 'RTBOX
    Screen.MousePointer = vbDefault
    InformDiff
    Flash
    btSave.Enabled = (NumDiffs > 0 And NumNonDiffs > 0)

End Sub

Private Sub btFindNextDiff_Click()

  'search rich text box for next difference (ie next colored line after current (block of) colored lines)

  Dim Color1    As Long 'initially SyntaxColor1
  Dim Color2    As Long 'initially SyntaxColor2
  Dim Color3    As Long 'not used yet
  Dim Color4    As Long 'initially DefForeColor so that it skips colored lines if DiffPointer points to one
  Dim CurrColor As Long

    Color1 = ColorFromRTF(RTFSyntaxColor1)
    Color2 = ColorFromRTF(RTFSyntaxColor2)
    Color3 = ColorFromRTF(RTFSyntaxColor3)
    Color4 = DefForeColor

    Screen.MousePointer = vbHourglass
    Enabled = False
    lbMessage = vbCrLf & "Searching, please wait..."
    With rtBox
        .Visible = False 'prevent flicker thru updates, meanwhile show label underneath
        DoEvents
        Do
            If DiffPointer >= LenBoxText Then
                DiffPointer = 1
                Beeper 4000, 10  'wrap around beep
            End If
            Do 'skip 'normal' colored lines
                DiffPointer = InStr(DiffPointer, BoxText, vbLf) + 1 'find next line
                .SelStart = DiffPointer - 1
                .SelLength = 1
                CurrColor = .SelColor 'get the color
            Loop Until CurrColor = Color1 Or CurrColor = Color2 Or CurrColor = Color4 Or DiffPointer >= LenBoxText
            Color1 = InsColor 'now that it's found a syn-colored line search for del/ins-colored line again
            Color2 = DelColor
            Color4 = Color2 'oops...
        Loop Until CurrColor = InsColor Or CurrColor = DelColor
        Enabled = True
        If SendMessage(.hWnd, EM_GETFIRSTVISIBLELINE, 0&, ByVal 0&) > 3 Then
            SendMessage .hWnd, EM_SCROLL, SB_LINEDOWN, ByVal 0&
            SendMessage .hWnd, EM_SCROLL, SB_LINEDOWN, ByVal 0&
            SendMessage .hWnd, EM_SCROLL, SB_LINEDOWN, ByVal 0&
        End If
        .Visible = True
        .SetFocus
    End With 'RTBOX
    Screen.MousePointer = vbDefault
    Flash

End Sub

Private Sub btSave_Click()

  Dim RTFFile   As String

    RTFFile = VBInstance.ActiveVBProject.Name & Spce & VBInstance.SelectedVBComponent.Name & Spce & "CompareResult.RTF"
    rtBox.SaveFile RTFFile
    MsgBoxEx "Compare results saved in " & CurDir & "\" & RTFFile, vbOKOnly Or vbInformation, "Saved"

End Sub

Private Function ColorFromRTF(RTFColor As String) As Long

  'will also accept capitals and hex or mixed strings like this: \RED&Hc0\green122\Blue&hHFF

  Dim sTmp As String

    sTmp = LCase$(Replace$(RTFColor, "\", ""))
    ColorFromRTF = RGB(Val(Mid$(RTFColor, InStr(sTmp, "red") + 3)), Val(Mid$(sTmp, InStr(sTmp, "green") + 5)), Val(Mid$(sTmp, InStr(sTmp, "blue") + 4)))

End Function

Private Function Colorize(ORTF As String) As String

  Dim lTmp                  As Long 'used for various things

  Dim NRTF                  As String 'the colorized rtf
  Dim CurrPtr               As Long 'pointer to scan the rtf text
  Dim OCurrPtr              As Long 'keeps track of the old rtf position
  Dim NCurrPtr              As Long 'keeps track of the new rtf position
  Dim LenOChunk             As Long 'the length of a sending chunk
  Dim LenNChunk             As Long 'the length of a receiving chunk
  Dim Word                  As String 'a word from it
  Dim Char                  As String 'and a character from it
  Dim SetSyntaxColor1       As String 'rtf command
  Dim SetSyntaxColor2       As String 'rtf command
  Dim SetSyntaxColor3       As String 'rtf command
  Dim ResetSyntaxColor      As String 'rtf command
  Dim LenSkpOn1             As Long
  Dim LenSkpOn2             As Long
  Dim LenSkpOn3             As Long
  Dim LenSkpOff             As Long
  Dim NumSeps               As Long
  Dim CurrentSyntaxColorIs  As Long 'keeps track of the currently selected color

    lbMessage = vbCrLf & "Coloring, please wait..."
    lbMessage.Refresh
    NRTF = Space$(Len(ORTF) * 1.2) 'good guess
    CurrPtr = InStr(ORTF, "{\colortbl")
    Do 'count the colors already present
        CurrPtr = InStr(CurrPtr + 1, ORTF, ";")
        Inc lTmp
    Loop Until Mid$(ORTF, CurrPtr + 1, 1) = "}"

    RTFSyntaxColor1 = ColorToRTF(SyntaxColor1)
    RTFSyntaxColor2 = ColorToRTF(SyntaxColor2)
    RTFSyntaxColor3 = ColorToRTF(SyntaxColor3)
    'make the four foreground color indexes into the color table
    ResetSyntaxColor = "\cf1 "
    SetSyntaxColor1 = "\cf" & CStr(lTmp) & Spce
    SetSyntaxColor2 = "\cf" & CStr(lTmp + 1) & Spce
    SetSyntaxColor3 = "\cf" & CStr(lTmp + 2) & Spce

    'and their lengths for skipping over them
    LenSkpOn1 = Len(SetSyntaxColor1)
    LenSkpOn2 = Len(SetSyntaxColor2)
    LenSkpOn3 = Len(SetSyntaxColor3)
    LenSkpOff = Len(ResetSyntaxColor)

    OCurrPtr = 1
    NCurrPtr = 1

    'now add our two syntax colors to the \colortable
    LenOChunk = CurrPtr - OCurrPtr
    Mid$(NRTF, NCurrPtr, LenOChunk + Len(RTFSyntaxColor1) + Len(RTFSyntaxColor2) + 2) = Mid$(ORTF, OCurrPtr, LenOChunk) & ";" & RTFSyntaxColor1 & ";" & RTFSyntaxColor2
    OCurrPtr = OCurrPtr + LenOChunk
    NCurrPtr = NCurrPtr + LenOChunk + Len(RTFSyntaxColor1) + Len(RTFSyntaxColor2) + 2

    'okay - lets go and fight rich text format
    CurrPtr = InStr(ORTF, "\par ") 'set CurrPtr to first paragraph
    Do Until CurrPtr >= Len(ORTF)
        Word = vbNullString
        lTmp = CurrPtr 'remember where we are
        Do While CurrPtr < Len(ORTF)
            Inc CurrPtr
            Char = Mid$(ORTF, CurrPtr, 1)
            If InStr(Delimiters, Char) Then 'delimiter found
                Exit Do 'word is ready 'loop 
              Else 'NOT INSTR(DELIMITERS,...
                Select Case Char
                  Case "\" 'an RTF thingy - skip to next space
                    CurrPtr = InStr(CurrPtr + 1, ORTF, Spce)
                    If CurrPtr = 0 Then 'oops - there ain't no more spaces - strange! but then you never know...
                        CurrPtr = Len(ORTF) + 1
                    End If
                    Exit Do 'loop 
                  Case """" 'quote - skip literal
                    Word = Char
                    CurrPtr = InStr(CurrPtr + 1, ORTF, """") 'no need to check for not found; they come in pairs
                    Exit Do 'loop 
                  Case "'" 'apostrophe - that's a legal reserved word although there may not be a delimiter after it
                    Word = Char
                    Exit Do 'loop 
                  Case "[" 'one of the more unknown secrets of VB
                    Word = Char
                    CurrPtr = InStr(CurrPtr + 1, ORTF, "]") 'no need to check for not found; they come in pairs
                    Exit Do 'loop 
                  Case ArrowOut, ArrowIn 'our marking chars for In and Out
                    If Mid$(ORTF, CurrPtr - 6, 2) = "\f" Then 'special font also - skip to next paragraph
                        CurrPtr = InStr(CurrPtr + 1, ORTF, "\par ") - 1
                        If CurrPtr < 0 Then 'pooh! - this was the last paragraph; we're thru
                            CurrPtr = Len(ORTF) + 1
                        End If
                        CurrentSyntaxColorIs = 0 'after this we assume default color
                      Else 'NOT MID$(ORTF,...
                        Word = Word & Char
                    End If
                  Case vbCr, vbLf 'have no meaning in RTF although they may be present
                    'do nothing
                  Case Else 'ahhh! a character - append it to the word
                    Word = Word & Char
                End Select
            End If
        Loop
        If Len(Word) Then 'we have a word
            NumSeps = 0
            If InStr(ResWds, "." & Word & ".") Then 'it's a keyword
                If Word = "'" Then 'a comment
                    CurrentSyntaxColorIs = 0
                  ElseIf Word = "Rem" Then 'another way of commenting 'NOT WORD...
                    CurrentSyntaxColorIs = 0 'assume default color
                End If
                If CurrentSyntaxColorIs <> 1 Then
                    LenOChunk = lTmp - OCurrPtr + 1
                    LenNChunk = LenOChunk + LenSkpOn1
                    Mid$(NRTF, NCurrPtr, LenNChunk) = Mid$(ORTF, OCurrPtr, LenOChunk) & SetSyntaxColor1
                    OCurrPtr = OCurrPtr + LenOChunk
                    NCurrPtr = NCurrPtr + LenNChunk
                    If Word = "Rem" Or Word = "'" Then
                        LenOChunk = CurrPtr - OCurrPtr + 1
                        LenNChunk = LenOChunk + LenSkpOn2
                        Mid$(NRTF, NCurrPtr, LenNChunk) = Mid$(ORTF, OCurrPtr, LenOChunk) & SetSyntaxColor2
                        Inc OCurrPtr, LenOChunk
                        Inc NCurrPtr, LenNChunk
                        CurrPtr = InStr(CurrPtr + 1, ORTF, "\par ") - 1
                        CurrentSyntaxColorIs = 2
                      Else 'NOT WORD...
                        CurrentSyntaxColorIs = 1
                    End If
                End If
              Else 'NOT INSTR(RESWDS,...
                If CurrentSyntaxColorIs <> 0 Then
                    LenOChunk = lTmp - OCurrPtr + 1
                    LenNChunk = LenOChunk + LenSkpOff
                    Mid$(NRTF, NCurrPtr, LenNChunk) = Mid$(ORTF, OCurrPtr, LenOChunk) & ResetSyntaxColor
                    OCurrPtr = OCurrPtr + LenOChunk
                    NCurrPtr = NCurrPtr + LenNChunk
                    If Len(NRTF) < NCurrPtr * 0.9 Then 'give it some more, it tends to run low
                        NRTF = NRTF & Space$(Len(NRTF) / 10)
                    End If
                    CurrentSyntaxColorIs = 0
                End If
            End If
          Else 'LEN(WORD) = FALSE/0
            Inc NumSeps
            If NumSeps = LenSep Then
                CurrentSyntaxColorIs = 2
            End If
        End If
        UpdateProgress CurrPtr / Len(ORTF)
        'grin - at one time during development this was running forward first and then
        'backwards because RTF grew faster the CurrPtr could march thru it...
        'a progress bar going backwards - a horrifying sight indeed!
    Loop
    'and finally xfer the remainder
    Mid$(NRTF, NCurrPtr, Len(ORTF) - OCurrPtr + 1) = Mid$(ORTF, OCurrPtr)

    'pooh! - at least we got thru without injuries
    Colorize = RTrim$(NRTF) 'so return what we made of it and hope for the best
    Enabled = True

End Function

Private Function ColorToRTF(ByVal Color As Long) As String

    ColorToRTF = "\red" & CStr(Color And 255) & "\green" & CStr(Color \ 256 And 255) & "\blue" & CStr(Color \ 256 \ 256 And 255)

End Function

Public Sub Compare(OldFileName As String, NewFileName As String)

    DisregardCase = (MsgBoxEx("During the comparison process upper/lower case characters are considered...", vbOKCancel Or vbQuestion Or vbDefaultButton2, "Comparator", 8000, -2, -2, Icon:=fIcon.picAa, OCapt:=OK & "|" & Abbrechen, NCapt:="&Identical|&Different") = vbOK)
    lbMessage = vbCrLf & "Comparing, please wait..."
    Form_Resize
    Show
    Enabled = False
    DiffPointer = 0
    NumDiffs = 0
    NumNonDiffs = 0
    CorrIfAtt = False
    DoneSoFar = 0
    PrevStep = 0
    Screen.MousePointer = vbHourglass
    Sig = Digest.Signature("A") 'test-call to find out how long signature is
    LenSig = Len(Sig)
    SkipToNext = LenSig + 1

    With rtBox
        .Visible = False 'prevent flicker thru updates, meanwhile show label underneath
        DoEvents
        .Text = vbNullString
        hFileOld = FreeFile
        Open OldFileName For Input _
             Access Read Shared As hFileOld
        SigsOld = vbNullString
        hFileNew = FreeFile
        Open NewFileName For Input _
             Access Read Shared As hFileNew
        SigsNew = vbNullString
        LenFiles = 2 * (LOF(hFileOld) + LOF(hFileNew)) 'double - we have two passes

        With Digest

            'create signatures for old file
            Do Until EOF(hFileOld)
                Line Input #hFileOld, Textline
                If Len(Trim$(Textline)) Then
                    If DisregardCase Then
                        Sig = .Signature(LCase$(StripMultipleSpacesFrom(Textline)))
                      Else 'DISREGARDCASE = FALSE/0
                        Sig = .Signature(StripMultipleSpacesFrom(Textline))
                    End If
                    SigsOld = SigsOld & Sig
                End If
                DoneSoFar = DoneSoFar + Len(Textline) + 2  'for crlf
                UpdateProgress DoneSoFar / LenFiles
            Loop

            'create signatures for new file
            Do Until EOF(hFileNew)
                Line Input #hFileNew, Textline
                If Len(Trim$(Textline)) Then
                    If DisregardCase Then
                        Sig = .Signature(LCase$(StripMultipleSpacesFrom(Textline)))
                      Else 'DISREGARDCASE = FALSE/0
                        Sig = .Signature(StripMultipleSpacesFrom(Textline))
                    End If
                    SigsNew = SigsNew & Sig
                    If CorrIfAtt Then
                        If InStr(Textline, "Attribute") = 1 Then  'last declaration has one ore more attributes
                            LastDeclSig = Sig 'so modify the signature to that line
                        End If
                    End If
                    CorrIfAtt = (Sig = LastDeclSig) 'CorrIfAtt comes on when processing last declaration line
                End If
                DoneSoFar = DoneSoFar + Len(Textline) + 2 'for crlf
                UpdateProgress DoneSoFar / LenFiles
            Loop

        End With 'DIGEST
        Close hFileOld, hFileNew

        'compare files
        hFileOld = FreeFile
        Open OldFileName For Input _
             Access Read Shared As hFileOld
        hFileNew = FreeFile
        Open NewFileName For Input _
             Access Read Shared As hFileNew

        Do
            If Len(SigsOld) Then
                WhereInOld = 0
                Do
                    WhereInOld = InStr(WhereInOld + 1, SigsOld, Left$(SigsNew, LenSig))
                Loop Until (WhereInOld Mod LenSig) < 2 'make sure we have hit a signature and not someting in between that matches by coincidence
              Else 'LEN(SIGSOLD) = FALSE/0
                WhereInOld = -1 'at end
            End If
            If Len(SigsNew) Then
                WhereInNew = 0
                Do
                    WhereInNew = InStr(WhereInNew + 1, SigsNew, Left$(SigsOld, LenSig))
                Loop Until (WhereInNew Mod LenSig) < 2 'make sure we have hit a signature and not someting in between that matches by coincidence
              Else 'LEN(SIGSNEW) = FALSE/0
                WhereInNew = -1 'at end
            End If
            If WhereInOld = -1 And WhereInNew = -1 Then
                Exit Do 'loop 
            End If
            Inc NumDiffs
            Select Case True 'now let's see...
              Case WhereInOld = 1 And WhereInNew = 1
                'we're in sync
                'output new line and skip old line
                OutNewLine DefForeColor
                GetLine (hFileOld)
                SigsOld = Mid$(SigsOld, SkipToNext)
                SigsNew = Mid$(SigsNew, SkipToNext)
                Dec NumDiffs
                Inc NumNonDiffs
              Case WhereInOld = 0 And WhereInNew = 0
                '1st line of Old is not in New and viceversa
                'output lines as deleted and new
                OutDelLine GetLine(hFileOld) & vbCrLf, DelColor
                OutNewLine InsColor
                SigsOld = Mid$(SigsOld, SkipToNext)
                SigsNew = Mid$(SigsNew, SkipToNext)
              Case WhereInOld < 1
                '1st line of New is not in Old
                'there are is a new line in New - output it
                OutNewLine InsColor
                SigsNew = Mid$(SigsNew, SkipToNext)
              Case WhereInNew < 1
                '1st line of Old is not in New
                'line has been deleted
                OutDelLine GetLine(hFileOld) & vbCrLf, DelColor
                SigsOld = Mid$(SigsOld, SkipToNext)
              Case Else
                'line is in both files at different positions
                If WhereInNew < WhereInOld Then
                    OutNewLine InsColor
                    SigsNew = Mid$(SigsNew, SkipToNext)
                  Else 'NOT WHEREINNEW...
                    OutDelLine GetLine(hFileOld) & vbCrLf, DelColor
                    SigsOld = Mid$(SigsOld, SkipToNext)
                End If
            End Select
            UpdateProgress DoneSoFar / LenFiles
        Loop
        Close hFileOld, hFileNew
        InformDiff
        btFindNextDiff.Enabled = (NumDiffs > 0 And NumNonDiffs > 0)
        btColorize.Enabled = True
        BoxText = .Text
        LenBoxText = Len(BoxText)
        Enabled = True
        .Visible = True
        .SetFocus
        Flash
    End With 'RTBOX
    Screen.MousePointer = vbDefault

End Sub

Private Sub Flash()

    If NumDiffs Then
        For i = 1 To 8
            DoEvents
            rtBox.SelLength = i
            Sleep 80
            rtBox.SelLength = 0
            Sleep 80
        Next i
    End If

End Sub

Private Sub Form_Initialize()

    hWndCompare = hWnd
    SetParent hWndCompare, hWndMDIClient 'make me a child of the IDE

End Sub

Private Sub Form_Load()

  Dim ForeColors() As String
  Dim Xlat As Variant

    Set Digest = New cMD5
    rtBox.Font.Name = IDEFontName
    rtBox.Font.Size = IDEFontSize
    btFindNextDiff.Enabled = False
    btSave.Enabled = False

    'get VB's forecolors
    ForeColors = Split(Trim$(IDEColors), Spce)

    'VB stores them in a queer way so we have to translate
    'the default color(0) also becomes black, btw
    Xlat = Array(0, 15, 7, 8, 0, 12, 4, 14, 6, 10, 2, 11, 3, 9, 1, 13, 5)

    'the four colors we're interested in
    SyntaxColor1 = QBColor(Xlat(Val(ForeColors(6)))) 'reserved words
    SyntaxColor2 = QBColor(Xlat(Val(ForeColors(5)))) 'comments
    SyntaxColor3 = QBColor(Xlat(Val(ForeColors(0)))) 'literals and such
    DefForeColor = QBColor(Xlat(Val(ForeColors(7)))) 'remainder

End Sub

Private Sub Form_Resize()

  Const TopMargin     As Long = 26
  Const BottomMargin  As Long = TopMargin + 2

    If WindowState <> vbMinimized Then
        MakeSizeable hWnd, WindowState <> vbMaximized
        MakeMoveable hWnd, WindowState <> vbMaximized
        If Height <= rtBox.Top * Screen.TwipsPerPixelY Then
            ReleaseCapture 'prevent the user from making the window too small
          Else 'NOT HEIGHT...
            On Error Resume Next
                lbBar.Move 2, TopMargin, 24, ScaleHeight - BottomMargin
                With rtBox
                    .Move lbBar.Width - 2, TopMargin, ScaleWidth - 25, ScaleHeight - BottomMargin
                    'keep the lines where they are
                    If (SendMessage(.hWnd, EM_SCROLL, SB_LINEDOWN, ByVal 0&) And &HFFFF&) = 1 Then 'did in deed scroll one line
                        SendMessage .hWnd, EM_SCROLL, SB_LINEUP, ByVal 0& 'so scroll back that line
                    End If
                End With 'RTBOX
                With lbEqual
                    .Width = ScaleWidth - lbEqual.Left - 2
                    .ForeColor = vbRed
                    Set Font = .Font 'to measure the font
                    'include spacing distance - stepchar & stepchar includes one sd and
                    'the single stepchar does not
                    ProgressStep = (TextWidth(StepChar & StepChar) - TextWidth(StepChar)) / (.Width - 2)
                End With 'LBEQUAL
                lbMessage.Move lbBar.Width - 2, TopMargin, ScaleWidth - 25, ScaleHeight - BottomMargin
            On Error GoTo 0
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Digest = Nothing

End Sub

Private Function GetLine(FromFile As Long) As String

  Dim TmpString As String

    Do
        GetLine = ""
        If Not EOF(FromFile) Then
            Line Input #FromFile, GetLine
            DoneSoFar = DoneSoFar + Len(GetLine) + 2 'for crlf
        End If
        TmpString = Trim$(GetLine)
    Loop Until Len(TmpString) Or EOF(FromFile)
    DrawSep = ((InStr(TmpString, "End Sub") Or InStr(TmpString, "End Function") Or InStr(GetLine, "End Property")) = 1)

End Function

Private Sub InformDiff()

    With lbEqual
        If NumDiffs Then
            .Caption = Spce & NumDiffs & " difference" & IIf(NumDiffs > 1, "s", "") & " found."
          Else 'NUMDIFFS = FALSE/0
            .Caption = " No differences found."
        End If
    End With 'LBEQUAL
    If DiffPointer > 0 Then
        rtBox.SelStart = DiffPointer - 1
      Else 'NOT DIFFPOINTER...
        rtBox.SelStart = 0
    End If

End Sub

Public Property Let LastDeclLine(Line As String)

  'the last declaration line is handed over to enable us
  'to draw the first separator after that line

    LastDeclSig = Digest.Signature(StripMultipleSpacesFrom(Line))

End Property

Private Sub OutDelLine(Textline As String, Color As Long)

    PrepColorAndIcon Color, 1 'arrow out
    With rtBox
        'skip leading blanks
        For i = 1 To Len(Textline)
            If Mid$(Textline, i, 1) = Spce Then
                .SelText = Spce
              Else 'NOT MID$(TEXTLINE,...
                Exit For 'loop varying i
            End If
        Next i
        'and now strike thru
        .SelStrikeThru = True
        .SelText = Mid$(Textline, i)
    End With 'RTBOX

End Sub

Private Sub OutNewLine(Color As Long)

    PrepColorAndIcon Color, 2 'arrow in
    With rtBox
        .SelText = GetLine(hFileNew) & vbCrLf
        If Left$(SigsNew, LenSig) = LastDeclSig Then 'this was last declaration line or last attribute of it
            DrawSep = True 'draw a separator line
        End If
        If DrawSep Then
            .SelColor = &HC0C0C0
            .SelStrikeThru = True
            .SelIndent = 0
            .SelText = Space$(LenSep) & Chr$(160) & vbCrLf
        End If
    End With 'RTBOX

End Sub

Private Sub PrepColorAndIcon(Color As Long, WhichMarker As Long)

  Dim Tmp   As Long

    With rtBox
        .SelStrikeThru = False
        If Color <> DefForeColor Then
            If DiffPointer = 0 Then
                DiffPointer = .SelStart + 1
            End If
            Tmp = .SelFontSize
            .SelIndent = 0
            .SelColor = Color
            'mark this line -it's either new or deleted
            .SelFontName = MarkerFont
            .SelBold = MarkerBold
            .SelFontSize = MarkerSize
            .SelText = Mid$(ArrowOut & ArrowIn, WhichMarker, 1)
            '...and reset
            .SelFontName = IDEFontName
            .SelBold = False
            .SelFontSize = Tmp
          Else 'NOT COLOR...
            .SelIndent = 11
        End If
        .SelColor = Color
    End With 'RTBOX

End Sub

Private Sub rtBox_GotFocus()

    HideCaret rtBox.hWnd

End Sub

Private Sub rtBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    HideCaret rtBox.hWnd

End Sub

Private Function StripMultipleSpacesFrom(Text As String) As String

  Dim p     As Long
  Dim q     As Long

    StripMultipleSpacesFrom = Trim$(Text)

    p = Len(StripMultipleSpacesFrom)
    For q = 1 To p
        Select Case Mid$(StripMultipleSpacesFrom, q, 1)
          Case """"
            Do
                Inc q
                If Mid$(StripMultipleSpacesFrom, q, 1) = Spce Then
                    Mid$(StripMultipleSpacesFrom, q, 1) = Chr$(0)
                End If
            Loop Until Mid$(StripMultipleSpacesFrom, q, 1) = """"
          Case "["
            Do
                Inc q
                If Mid$(StripMultipleSpacesFrom, q, 1) = Spce Then
                    Mid$(StripMultipleSpacesFrom, q, 1) = Chr$(0)
                End If
            Loop Until Mid$(StripMultipleSpacesFrom, q, 1) = "]"
        End Select
    Next q
    Do
        q = p
        StripMultipleSpacesFrom = Replace$(StripMultipleSpacesFrom, "    ", Spce)
        StripMultipleSpacesFrom = Replace$(StripMultipleSpacesFrom, "   ", Spce)
        StripMultipleSpacesFrom = Replace$(StripMultipleSpacesFrom, "  ", Spce)
        p = Len(StripMultipleSpacesFrom)
    Loop While p <> q

End Function

Private Sub UpdateProgress(Percent As Single)

  Dim CurrStep  As Long

    CurrStep = Percent / ProgressStep
    If CurrStep <> PrevStep Then
        lbEqual = String$(CurrStep, StepChar)
        PrevStep = CurrStep
        lbEqual.Refresh
    End If

End Sub

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 16:14)  Decl: 81  Code: 672  Total: 753 Lines
':) CommentOnly: 48 (6,4%)  Commented: 120 (15,9%)  Empty: 111 (14,7%)  Max Logic Depth: 7
