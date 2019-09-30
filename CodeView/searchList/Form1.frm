VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFile 
      Height          =   330
      Left            =   585
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Text            =   "D:\_code\vbdec2\notes\_name_diffs.txt"
      Top             =   45
      Width           =   3120
   End
   Begin VB.CheckBox chkScript 
      Caption         =   "List"
      Height          =   240
      Left            =   45
      TabIndex        =   13
      Top             =   90
      Width           =   780
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   ".."
      Height          =   285
      Left            =   3780
      TabIndex        =   12
      Top             =   45
      Width           =   285
   End
   Begin VB.Frame Frame1 
      Caption         =   "Use JS To Process Listitem entries on click before sending to search UI"
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   495
      Visible         =   0   'False
      Width           =   9420
      Begin VB.CheckBox chkScriptEnabled 
         Caption         =   "Enable"
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   315
         Width           =   1140
      End
      Begin MSScriptControlCtl.ScriptControl sc 
         Left            =   5940
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         Language        =   "jscript"
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   1395
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Text            =   "Form1.frx":0000
         Top             =   225
         Width           =   7890
      End
      Begin VB.CommandButton Command2 
         Caption         =   "def"
         Height          =   330
         Left            =   720
         TabIndex        =   9
         Top             =   3015
         Width           =   510
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   1395
         TabIndex        =   8
         Top             =   2475
         Width           =   7890
      End
      Begin VB.TextBox buf1 
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Text            =   "buf1"
         Top             =   1350
         Width           =   1185
      End
      Begin VB.TextBox buf2 
         Height          =   285
         Left            =   135
         TabIndex        =   6
         Text            =   "buf2"
         Top             =   1665
         Width           =   1185
      End
      Begin VB.CommandButton Command3 
         Caption         =   "?"
         Height          =   330
         Left            =   90
         TabIndex        =   5
         Top             =   3015
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Click to copy buf"
         Height          =   285
         Left            =   90
         TabIndex        =   11
         Top             =   1080
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   285
      Left            =   4140
      TabIndex        =   3
      Top             =   45
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   285
      Left            =   4770
      TabIndex        =   2
      Top             =   45
      Width           =   510
   End
   Begin Project1.ucFilterList lv 
      Height          =   4335
      Left            =   45
      TabIndex        =   0
      Top             =   4230
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   7646
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Dim defScr As String
Dim dlg As New CCmnDlg

Private Sub chkScript_Click()
    Frame1.Visible = chkScript.value
    Form_Resize
End Sub

Private Sub cmdBrowse_Click()
    tmp = dlg.OpenDialog()
    If Len(tmp) > 0 Then txtFile = tmp
End Sub

Private Sub Command3_Click()
    Const h = "p.alert(x), p.lst(x), p.trim(x), p.buf1=, p.buf2="
    MsgBox h, vbInformation
End Sub

Sub alert(x)
    MsgBox x
End Sub

Sub lst(x)
    List1.AddItem x
End Sub

Function trim(x)
    x = Replace(x, vbTab, " ")
    trim = VBA.trim(x)
End Function

Function SendIPCCommand(msg As String)
    
    Const WM_PASTE = &H302
    Dim tmp As String
    Dim hAddin As Long
    
    hAddin = GetSetting("codeview", "ipc", "txtIPC", 0)
   
    If IsWindow(hAddin) = 0 Then
        List1.AddItem "SendIPCCommand: CodeView.Search all IPC window not found?"
        Exit Function
    End If
    
    List1.AddItem "SendIPCCommand hwnd: " & hAddin
    
    Clipboard.Clear
    Clipboard.SetText msg
    PostMessage hAddin, WM_PASTE, 0, 0

End Function

Private Sub cmdLoad_Click()
    lv.Clear
    tmp = fso.ReadFile(txtFile)
    tmp = Split(tmp, vbCrLf)
    For Each x In tmp
        lv.ListItems.Add , , x
    Next
End Sub

Function process(ByVal x)
    On Error Resume Next
    x = Replace(x, "'", Empty)
    sc.Reset
    sc.AddObject "p", Me, True
    sc.AddCode Text1.Text
    process = sc.Eval("ItemClick_Process('" & x & "')")
    If Err.Number <> 0 Then
        List1.AddItem sc.Error.Line & " " & sc.Error.Description
    End If
End Function

Private Sub Command1_Click()
     fso.WriteFile txtFile, lv.GetAllElements
End Sub

Private Sub Command2_Click()
    Text1 = defScr
End Sub



Private Sub Form_Load()
    lv.SetColumnHeaders "Text"
    lv.AllowDelete = True
    lv.MultiSelect = True
    lv.SetFont "Courier", 12
    defScr = Text1
    p = App.Path & "\lastScript.txt"
    If fso.FileExists(p) Then Text1 = fso.ReadFile(p)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Frame1.Visible Then
        lv.Top = Frame1.Top + Frame1.Height + 200
    Else
        lv.Top = txtFile.Top + txtFile.Height + 100
    End If
    lv.Height = Me.Height - lv.Top - 350
    lv.Width = Me.Width - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    p = App.Path & "\lastScript.txt"
    fso.WriteFile p, Text1
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    List1.Clear
    copyAfter = Empty
    If chkScriptEnabled.value Then
        x = process(Item.Text)
    Else
        x = Item.Text
    End If
    SendIPCCommand CStr(x)
    If Err.Number <> 0 Then List1.AddItem Err.Description
End Sub

Private Sub Text2_Click()
    Clipboard.Clear
    Clipboard.SetText Text2
End Sub

Private Sub Text3_Click()
    Clipboard.Clear
    Clipboard.SetText Text3
End Sub

Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    txtFile = Data.Files(1)
End Sub
