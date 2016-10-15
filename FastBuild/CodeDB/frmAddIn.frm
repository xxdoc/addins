VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAddIn 
   Caption         =   "Code Database Addin -dzzie@yahoo.com"
   ClientHeight    =   6735
   ClientLeft      =   2190
   ClientTop       =   2235
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   13095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "Add File"
      Height          =   285
      Left            =   9945
      TabIndex        =   17
      Top             =   6300
      Width           =   915
   End
   Begin VB.Frame fraAdd 
      Caption         =   " Add New Code "
      Height          =   4875
      Left            =   3195
      TabIndex        =   7
      Top             =   540
      Visible         =   0   'False
      Width           =   8700
      Begin VB.CommandButton cmdIPCTest 
         Caption         =   "IPC Test"
         Height          =   330
         Left            =   270
         TabIndex        =   15
         Top             =   1845
         Width           =   960
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         TabIndex        =   11
         Top             =   270
         Width           =   6570
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   1575
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   780
         Width           =   6660
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Extract Prototype"
         Height          =   255
         Index           =   0
         Left            =   5805
         TabIndex        =   9
         Top             =   4500
         Width           =   1620
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   255
         Index           =   2
         Left            =   7515
         TabIndex        =   8
         Top             =   4500
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Double click body textbox to paste and parse prototype"
         Height          =   285
         Left            =   1575
         TabIndex        =   18
         Top             =   4500
         Width           =   3930
      End
      Begin VB.Label lblClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   8325
         TabIndex        =   14
         Top             =   180
         Width           =   330
      End
      Begin VB.Label Label2 
         Caption         =   "Code Body"
         Height          =   375
         Left            =   225
         TabIndex        =   13
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Prototype"
         Height          =   285
         Left            =   270
         TabIndex        =   12
         Top             =   315
         Width           =   1005
      End
   End
   Begin VB.OptionButton optFile 
      Caption         =   "Files"
      Height          =   285
      Left            =   8415
      TabIndex        =   6
      Top             =   6300
      Width           =   960
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Function"
      Height          =   285
      Left            =   7290
      TabIndex        =   5
      Top             =   6300
      Value           =   -1  'True
      Width           =   1050
   End
   Begin VB.ComboBox cboLang 
      Height          =   315
      Left            =   5310
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6300
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   255
      Index           =   5
      Left            =   12090
      TabIndex        =   1
      Top             =   6300
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   255
      Index           =   4
      Left            =   11250
      TabIndex        =   0
      Top             =   6300
      Width           =   855
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   6045
      Left            =   4500
      TabIndex        =   2
      Top             =   135
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   10663
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   50000
      TextRTF         =   $"frmAddIn.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Codedb.ucFilterList lv 
      Height          =   6495
      Left            =   90
      TabIndex        =   16
      Top             =   135
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   11456
   End
   Begin VB.Label Lang 
      Caption         =   "Lang"
      Height          =   240
      Left            =   4770
      TabIndex        =   3
      Top             =   6300
      Width           =   420
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuAddCode 
         Caption         =   "Add Code"
      End
      Begin VB.Menu mnuSaveFile 
         Caption         =   "Add New File to DB (drop on lv)"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStrings 
         Caption         =   "String Tools"
      End
      Begin VB.Menu mnuAdoConstr 
         Caption         =   "Connection Str Builder"
      End
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Const LB_FINDSTRING = &H18F

Dim rs As ADODB.Recordset
Dim dlg As New clsCmnDlg2
Dim cn As New ADODB.Connection
Dim loadedFile As String
Dim loadedFileText As String
Dim projDir As String
Dim hAddin As Long

Enum cmdTypes
    ct_addFile
End Enum

Private Sub cboLang_Click()
    
    Dim txt As String, lang_id As Long
    Dim li As ListItem
    
    txt = cboLang.Text
    lang_id = Mid(txt, InStrRev(txt, "@") + 1, Len(txt))

    Set rs = cn.Execute("Select * from CodeDB where lang_id=" & lang_id & " and isFile=" & IIf(optFile.value, 1, 0))
    lv.Clear
    
    If rs.BOF And rs.EOF Then Exit Sub
    
    rs.MoveFirst
    
    While Not rs.EOF
        Set li = lv.AddItem(rs.Fields("NAME").value)
        li.Tag = rs.Fields("ID").value
        rs.MoveNext
    Wend
    
    rs.Close
    
End Sub

Private Sub mnuSaveFile_Click()
    Dim p As String
    p = dlg.OpenDialog(AllFiles, , , Me.hwnd)
    If Len(p) = 0 Then Exit Sub
    AddFile p
End Sub


Private Sub lv_ItemDeleted(Item As MSComctlLib.ListItem, cancel As Boolean)
    cn.Execute "Delete from CodeDB where ID=" & Item.Tag
    cboLang_Click
End Sub

Private Sub cmdIPCTest_Click()

    Const WM_PASTE = &H302
    
    If IsWindow(hAddin) = 0 Then
        MsgBox "Parent form closed? Hwnd not found? we could rescan but not here.."
        Exit Sub
    End If
    
    Clipboard.Clear
    Clipboard.SetText InputBox("Send:", , "test")
    PostMessage hAddin, WM_PASTE, 0, 0
    
End Sub

Function SendIPCCommand(cmdType As cmdTypes, msg As String)
    
    Const WM_PASTE = &H302
    Dim tmp As String
    
    If hAddin = 0 Then Exit Function
    
    If IsWindow(hAddin) = 0 Then
        MsgBox "Parent form closed? Hwnd not found? we could rescan but not here.."
        Exit Function
    End If
    
    Select Case cmdType
        Case ct_addFile: tmp = "add:"
    End Select
    
    If Len(tmp) = 0 Then
        MsgBox "dev error unknown cmd type"
        Exit Function
    End If
    
    tmp = tmp & msg

    Clipboard.Clear
    Clipboard.SetText tmp
    PostMessage hAddin, WM_PASTE, 0, 0

End Function

Private Sub cmdAddFile_Click()
    On Error GoTo hell
    Dim p As String
    If Len(loadedFileText) = 0 Then Exit Sub
    p = dlg.SaveDialog(AllFiles, projDir, "Save As", , Me.hwnd, loadedFile)
    If Len(p) = 0 Then Exit Sub
    WriteFile p, loadedFileText
    SendIPCCommand ct_addFile, p
    Exit Sub
hell:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    
    Dim cmd As String, a As Long
    
    cmdAddFile.Enabled = False
    cmdIPCTest.Visible = IsIde()
    lv.SetColumnHeaders "name*"
    lv.MultiSelect = False
    lv.AllowDelete = True
    lv.font = "tahoma"
    
    fraAdd.Move lv.Left, lv.top, Me.Width - 400, Text2.Height
    
    cmd = Replace(Command, """", Empty)
    a = InStr(cmd, "hwnd=")
    
    If a > 0 Then
        hAddin = CLng(Mid(cmd, a + 5))
        cmd = Trim(Mid(cmd, 1, a))
    End If
        
    On Error GoTo oops
    
    If FileExists(cmd) Then
        projDir = GetParentFolder(cmd)
    End If

    cn.ConnectionString = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.path & "\db1.mdb;"
    cn.Open
    
    Set rs = cn.Execute("Select * from langs")

    cboLang.Clear
    While Not rs.EOF
        cboLang.AddItem rs.Fields("lang").value & String(80, " ") & "@" & rs.Fields("autoid").value
        rs.MoveNext
    Wend
    
    cboLang.ListIndex = 0
    
    Exit Sub
oops: MsgBox Err.Description
End Sub

Private Sub Form_Unload(cancel As Integer)
    On Error Resume Next
    cn.Close
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo oops
    Select Case Index
        Case 0: Call extract
        Case 2: Call AddNewCode: Text3 = Empty: Text4 = Empty
        Case 4: Clipboard.Clear: Clipboard.SetText Text2.Text
        Case 5: Text2.Text = Empty
    End Select
    Exit Sub
oops: MsgBox Err.Description
End Sub

Sub AddFile(pth As String)
    
    If Not FileExists(pth) Then Exit Sub
    
    Dim txt As String, lang_id As Long
    Dim n As String
    
    n = FileNameFromPath(pth)
    txt = cboLang.Text
    lang_id = Mid(txt, InStrRev(txt, "@") + 1, Len(txt))
 
    rs.Open "SELECT * FROM CODEDB", cn, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs.Fields("NAME").value = n
    rs.Fields("lang_id").value = lang_id
    rs.Fields("isFile").value = 1
    SaveFileToDB pth, rs, "CODE"
    rs.Update
    rs.Close
    
End Sub

Public Function SaveFileToDB(ByVal FileName As String, rs As Object, FieldName As String) As Boolean
'**************************************************************
'PURPOSE: SAVES DATA FROM BINARY FILE (e.g., .EXE, WORD DOCUMENT
'CONTROL TO RECORDSET RS IN FIELD NAME FIELDNAME

'FIELD TYPE MUST BE BINARY (OLE OBJECT IN ACCESS)

'REQUIRES: REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS 2.0 or ABOVE

'SAMPLE USAGE
'Dim sConn As String
'Dim oConn As New ADODB.Connection
'Dim oRs As New ADODB.Recordset
'
'
'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
'
'oConn.Open sConn
'oRs.Open "SELECT * FROM MYTABLE", oConn, adOpenKeyset, _
   adLockOptimistic
'oRs.AddNew

'SaveFileToDB "C:\MyDocuments\MyDoc.Doc", oRs, "MyFieldName"
'oRs.Update
'oRs.Close
'**************************************************************

Dim iFileNum As Integer
Dim lFileLength As Long

Dim abBytes() As Byte
Dim iCtr As Integer

On Error GoTo ErrorHandler
If Dir(FileName) = "" Then Exit Function
If Not TypeOf rs Is ADODB.Recordset Then Exit Function

'read file contents to byte array
iFileNum = FreeFile
Open FileName For Binary Access Read As #iFileNum
lFileLength = LOF(iFileNum)
If (lFileLength - 1) Mod 2 <> 0 Then lFileLength = lFileLength + 1
ReDim abBytes(lFileLength)
Get #iFileNum, , abBytes()

'put byte array contents into db field
rs.Fields(FieldName).AppendChunk abBytes()
Close #iFileNum

SaveFileToDB = True
ErrorHandler:
End Function

Public Function LoadFileFromDB(FileName As String, rs As Object, FieldName As String, Optional ByRef emsg As String) As Boolean
'************************************************
'PURPOSE: LOADS BINARY DATA IN RECORDSET RS,
'FIELD FieldName TO a File Named by the FileName parameter

'REQUIRES: REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS 2.0 or ABOVE

'SAMPLE USAGE
'Dim sConn As String
'Dim oConn As New ADODB.Connection
'Dim oRs As New ADODB.Recordset
'
'
'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
'
'oConn.Open sConn
'oRs.Open "SELECT * FROM MyTable", oConn, adOpenKeyset,
' adLockOptimistic
'LoadFileFromDB "C:\MyDocuments\MyDoc.Doc",  oRs, "MyFieldName"
'oRs.Close
'************************************************
Dim iFileNum As Integer
Dim lFileLength As Long
Dim abBytes() As Byte
Dim iCtr As Integer

    On Error GoTo ErrorHandler
    If Not TypeOf rs Is ADODB.Recordset Then Exit Function
    
    iFileNum = FreeFile
    Open FileName For Binary As #iFileNum
    lFileLength = LenB(rs(FieldName))
    abBytes = rs(FieldName).GetChunk(lFileLength)
    Put #iFileNum, , abBytes()
    Close #iFileNum
    
    LoadFileFromDB = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    emsg = Err.Description
    Close #iFileNum
End Function

Private Sub AddNewCode()

    If Text3 = Empty Or Text4 = Empty Then
        MsgBox "Need code or name duh"
        Exit Sub
    End If
    
    Dim txt As String, lang_id As Long
    
    txt = cboLang.Text
    lang_id = Mid(txt, InStrRev(txt, "@") + 1, Len(txt))

    rs.Open "SELECT * FROM CODEDB", cn, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs.Fields("NAME").value = Text3
    rs.Fields("CODE").value = Text4
    rs.Fields("lang_id").value = lang_id
    rs.Fields("isFile").value = 0
    rs.Update
    rs.Close
    cboLang_Click
   
End Sub

Private Sub extract()
    On Error Resume Next
    Dim tmp As String, fs
    If Text4 = Empty Then MsgBox "Ughh need function to extract name from!": Exit Sub
    tmp = firstLine(Text4)
    fs = InStrRev(tmp, " ", InStr(tmp, "("))
    tmp = Mid(tmp, fs + 1, Len(tmp))
    If Len(tmp) > 254 Then tmp = Mid(tmp, 1, 254)
    Text3 = tmp
End Sub

Private Sub CopyCode()
 
    Dim tmp As String, emsg As String
    Dim txt As String, cid
    Dim rs As New Recordset
    'On Error Resume Next
    
    loadedFile = Empty
    loadedFileText = Empty
    cmdAddFile.Enabled = False
    Close 'all open file handles..
    
    If lv.selItem Is Nothing Then Exit Sub
    
    cid = lv.selItem.Tag
    
    If optFile.value Then
        rs.Open "SELECT * FROM CodeDB where id=" & cid, cn, adOpenKeyset, adLockOptimistic
        tmp = GetFreeFileName()
        If FileExists(tmp) Then Kill tmp
        If LoadFileFromDB(tmp, rs, "CODE", emsg) Then
            If FileExists(tmp) Then
                loadedFile = lv.selItem.Text
                loadedFileText = stripAnyFromEnd(ReadFile(tmp), vbCr, vbLf, Chr(0))
                Text2 = loadedFileText
                Kill tmp
                Text2.selStart = 1
                cmdAddFile.Enabled = True
            End If
        Else
            MsgBox "Error loading file from db" & emsg
        End If
        rs.Close
    Else
        Set rs = cn.Execute("SELECT * FROM CodeDB where ID=" & cid)
        Text2.Text = Text2.Text & vbCrLf & vbCrLf & rs("CODE")
    End If

End Sub

Function firstLine(it)
    Dim t
    t = Split(it, vbCrLf)
    firstLine = t(0)
End Function

Private Sub lblClose_Click()
    fraAdd.Visible = False
End Sub

Private Sub lv_BeforeDelete(cancel As Boolean)
    If MsgBox("Are you sure you want to delete " & lv.selCount & " items?", vbYesNo) = vbNo Then
        cancel = True
    End If
End Sub

Private Sub lv_Click()
    CopyCode
End Sub

Private Sub lv_OleDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim f
    For Each f In Data.Files
        AddFile CStr(f)
    Next
    optFile_Click
End Sub

Private Sub lv_UserHitReturnInFilter()
    CopyCode
End Sub

Private Sub mnuAddCode_Click()
    fraAdd.Visible = True
End Sub

Private Sub mnuAdoConstr_Click()
    frmAdo.Show
End Sub

Private Sub mnuStrings_Click()
    frmLazy.Show
End Sub

Private Sub optFile_Click()
    cboLang_Click
    cmdAddFile.Enabled = True
End Sub

Private Sub Option1_Click()
    cboLang_Click
    cmdAddFile.Enabled = False
End Sub

Private Sub Text2_Change()
    modSyntaxHighlighting.SyntaxHighlight Text2
    Text2.selStart = Len(Text2)
End Sub

Private Sub Text4_Change()
    Text4.selStart = 0
    Text4.selLength = 0
End Sub

Private Sub Text4_DblClick()
    Dim c
    c = Clipboard.GetText
    If c <> Empty Then Text4 = c: Command1_Click 0
End Sub

 

Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Sub WriteFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Function ReadFile(FileName)
    Dim f As Long, temp
  f = FreeFile
  temp = ""
   Open FileName For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(FileName), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Function FolderExists(path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Function GetFreeFileName(Optional ByVal folder As String, Optional extension = ".txt") As String
    
    On Error GoTo handler 'can have overflow err once in awhile :(
    Dim i As Integer
    Dim tmp As String

    If Len(folder) = 0 Then folder = Environ("temp")
    If Not FolderExists(folder) Then Exit Function
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Left(extension, 1) <> "." Then extension = "." & extension
    
again:
    Do
      tmp = folder & RandomNum() & extension
    Loop Until Not FileExists(tmp)
    
    GetFreeFileName = tmp
    
Exit Function
handler:

    If i < 10 Then
        i = i + 1
        GoTo again
    End If
    
End Function

Function RandomNum() As Long
    Dim tmp As Long
    Dim tries As Long
    
    On Error Resume Next

    Do While 1
        Err.Clear
        Randomize
        tmp = Round(Timer * Now * Rnd(), 0)
        RandomNum = tmp
        If Err.Number = 0 Then Exit Function
        If tries < 100 Then
            tries = tries + 1
        Else
            Exit Do
        End If
    Loop
    
    RandomNum = GetTickCount
    
End Function

Function FileNameFromPath(fullpath) As String
    Dim tmp
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

