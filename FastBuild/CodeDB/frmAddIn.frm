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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   285
      Left            =   10485
      TabIndex        =   23
      Top             =   6300
      Width           =   690
   End
   Begin VB.CommandButton cmdSaveFile 
      Caption         =   "Save File"
      Height          =   240
      Left            =   9450
      TabIndex        =   22
      Top             =   6300
      Width           =   915
   End
   Begin VB.Frame fraAdd 
      Caption         =   " Add New Code "
      Height          =   4875
      Left            =   3195
      TabIndex        =   10
      Top             =   540
      Visible         =   0   'False
      Width           =   8700
      Begin VB.CommandButton cmdBrowse 
         Caption         =   ".."
         Height          =   240
         Left            =   5355
         TabIndex        =   21
         Top             =   4500
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add File"
         Height          =   240
         Index           =   1
         Left            =   5625
         TabIndex        =   20
         Top             =   4500
         Width           =   915
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   855
         OLEDropMode     =   1  'Manual
         TabIndex        =   19
         Top             =   4500
         Width           =   4470
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   780
         Width           =   6660
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Extract"
         Height          =   255
         Index           =   0
         Left            =   6615
         TabIndex        =   12
         Top             =   4500
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   255
         Index           =   2
         Left            =   7515
         TabIndex        =   11
         Top             =   4500
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Add File"
         Height          =   285
         Left            =   90
         TabIndex        =   18
         Top             =   4500
         Width           =   1005
      End
      Begin VB.Label lblClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   8325
         TabIndex        =   17
         Top             =   180
         Width           =   330
      End
      Begin VB.Label Label2 
         Caption         =   "Code Body"
         Height          =   375
         Left            =   225
         TabIndex        =   16
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Prototype"
         Height          =   285
         Left            =   270
         TabIndex        =   15
         Top             =   315
         Width           =   1005
      End
   End
   Begin VB.OptionButton optFile 
      Caption         =   "Files"
      Height          =   285
      Left            =   8415
      TabIndex        =   9
      Top             =   6300
      Width           =   960
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Function"
      Height          =   285
      Left            =   7290
      TabIndex        =   8
      Top             =   6300
      Value           =   -1  'True
      Width           =   1050
   End
   Begin VB.ComboBox cboLang 
      Height          =   315
      Left            =   5310
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6300
      Width           =   1860
   End
   Begin VB.ListBox lstFilter 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   1080
      TabIndex        =   4
      Top             =   1305
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   255
      Index           =   5
      Left            =   12090
      TabIndex        =   3
      Top             =   6300
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   255
      Index           =   4
      Left            =   11250
      TabIndex        =   2
      Top             =   6300
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Left            =   90
      TabIndex        =   1
      Top             =   6255
      Width           =   4365
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6060
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4365
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   6045
      Left            =   4500
      TabIndex        =   5
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
   Begin VB.Label Lang 
      Caption         =   "Lang"
      Height          =   240
      Left            =   4770
      TabIndex        =   6
      Top             =   6300
      Width           =   420
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuStrings 
         Caption         =   "Strings"
      End
      Begin VB.Menu mnuAddCode 
         Caption         =   "Add Code"
      End
      Begin VB.Menu mnuAdoConstr 
         Caption         =   "Ado ConStr"
      End
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Const LB_FINDSTRING = &H18F

#Const IS_ADDIN = False

#If IS_ADDIN Then
    Public VBInstance As VBIDE.VBE
    Public Connect As Connect
#End If

Dim ws As Workspace
Dim db As Database
Dim rs As Recordset
Dim dlg As New clsCmnDlg2
Dim cn As New ADODB.Connection
Dim loadedFile As String
Dim loadedFileText As String
Dim projDir As String

'todo: switch over to ado completely..
'todo: load projDir from command passed in from addin..
'todo: ipc server in vb ide addin to allow for remote process to add files to work space?


Private Sub cboLang_Click()
    
    Dim txt As String, lang_id As Long
    
    txt = cboLang.Text
    
    lang_id = Mid(txt, InStrRev(txt, "@") + 1, Len(txt))
      
    Dim rs As ADODB.Recordset
    If cn.State <> 1 Then cn.Open
    
    Set rs = cn.Execute("Select * from CodeDB where lang_id=" & lang_id & " and isFile=" & IIf(optFile.Value, 1, 0))
    List1.Clear
    
    If rs.BOF And rs.EOF Then Exit Sub
    
    rs.MoveFirst
    
    While Not rs.EOF
        List1.AddItem rs.Fields("NAME").Value & String(80, " ") & "@" & rs.Fields("ID").Value
        rs.MoveNext
    Wend
    
    rs.Close
    cn.Close
    
End Sub

Private Sub cmdBrowse_Click()
    Dim p As String
    p = dlg.OpenDialog(AllFiles, , , Me.hwnd)
    If Len(p) > 0 Then txtFile = p
End Sub

Private Sub cmdDelete_Click()
    
    If MsgBox("Are you sure you want to delete this entry?", vbYesNo) = vbNo Then Exit Sub
    
    If lstFilter.Visible Then
        txt = lstFilter.List(lstFilter.ListIndex)
    Else
        txt = List1.List(List1.ListIndex)
    End If
    
    If Len(txt) = 0 Then
        'they did not select an entry, but there is only one filtered result so thats it..
        'user action: the entered a filter, saw one result and hit return
        If lstFilter.Visible And lstFilter.ListCount = 1 Then
            txt = lstFilter.List(0)
        End If
    End If
    
    If Len(txt) = 0 Then Exit Sub
    cid = Mid(txt, InStrRev(txt, "@") + 1, Len(txt))
    
    If cn.State <> 1 Then cn.Open
    cn.Execute "Delete from CodeDB where ID=" & cid
    cn.Close
    cboLang_Click
    
End Sub

Private Sub cmdSaveFile_Click()
    On Error GoTo hell
    Dim p As String
    p = dlg.SaveDialog(AllFiles, projDir, "Save As", , Me.hwnd, loadedFile)
    If Len(p) = 0 Then Exit Sub
    WriteFile p, loadedFileText
    Exit Sub
hell:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    On Error GoTo oops
    
    With List1
        lstFilter.Move .Left, .top, .Width, .Height
        fraAdd.Move .Left, .top, Me.Width - 400, Text2.Height
    End With
    
    cn.ConnectionString = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.path & "\db1.mdb;"

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.path & "\db1.mdb")
    
    Set rs = db.OpenRecordset("langs", dbOpenDynaset)
    rs.MoveFirst
    cboLang.Clear
    While Not rs.EOF
        cboLang.AddItem rs.Fields("lang").Value & String(80, " ") & "@" & rs.Fields("autoid").Value
        rs.MoveNext
    Wend
    
    cboLang.ListIndex = 0
    
    Exit Sub
oops: MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs.Close: db.Close: ws.Close
    Set rs = Nothing: Set db = Nothing: Set ws = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo oops
    Select Case Index
        Case 0: Call extract
        Case 1: Call AddFile
        Case 2: Call AddNewCode: Text3 = Empty: Text4 = Empty
                
        'Case 2: Call comment
        'Case 3: Call comment(False)
        Case 4: Clipboard.Clear: Clipboard.SetText Text2.Text
        Case 5: Text2.Text = Empty
    End Select
    Exit Sub
oops: MsgBox Err.Description
End Sub

Sub AddFile()
    
    If Not FileExists(txtFile) Then Exit Sub
    
    Dim txt As String, lang_id As Long
    Dim n As String
    
    n = FileNameFromPath(txtFile)
    txt = cboLang.Text
    lang_id = Mid(txt, InStrRev(txt, "@") + 1, Len(txt))
    
    Dim rs As New ADODB.Recordset
     
    If cn.State <> 1 Then cn.Open
    rs.Open "SELECT * FROM CODEDB", cn, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs.Fields("NAME").Value = n
    rs.Fields("lang_id").Value = lang_id
    rs.Fields("isFile").Value = 1
    SaveFileToDB txtFile, rs, "CODE"
    rs.Update
    rs.Close
    cn.Close
    
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

Public Function LoadFileFromDB(FileName As String, rs As Object, FieldName As String) As Boolean
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

ErrorHandler:
End Function

Private Sub AddNewCode()

    If Text3 = Empty Or Text4 = Empty Then
        MsgBox "Need code or name duh"
        Exit Sub
    End If
    
    Dim txt As String, lang_id As Long
    
    txt = cboLang.Text
    lang_id = Mid(txt, InStrRev(txt, "@") + 1, Len(txt))
    
    q = Chr(34) 'quote
    dq = Chr(34) & Chr(34)
    v = """" & Replace(Text3, q, dq) & """,""" & Replace(Text4, q, dq) & """"
    sSQL = "INSERT INTO CodeDB (NAME,CODE,lang_id) VALUES(" & v & ", " & lang_id & ");"
    db.Execute sSQL
    
    cboLang_Click
   
End Sub

Private Sub extract()
    On Error Resume Next
    If Text4 = Empty Then MsgBox "Ughh need function to extract name from!": Exit Sub
    tmp = firstLine(Text4)
    fs = InStrRev(tmp, " ", InStr(tmp, "("))
    tmp = Mid(tmp, fs + 1, Len(tmp))
    If Len(tmp) > 254 Then tmp = Mid(tmp, 1, 254)
    Text3 = tmp
End Sub

'Private Sub comment(Optional out As Boolean = True)
'    'basic outline of sub from Palidan on pscode
'    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
'    VBInstance.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
'    If StartLine = EndLine And StartColumn = EndColumn Then Exit Sub
'    For i = StartLine To EndLine
'        If i = EndLine And EndColumn = 1 Then Exit For
'        l = VBInstance.ActiveCodePane.CodeModule.Lines(i, 1)
'        If out Then
'            VBInstance.ActiveCodePane.CodeModule.ReplaceLine i, "'" + l
'        Else
'            VBInstance.ActiveCodePane.CodeModule.ReplaceLine i, Mid(l, 2)
'        End If
'    Next
'    Connect.Hide
'End Sub

Private Sub CopyCode()
 
    Dim tmp As String
    Dim adors As New ADODB.Recordset
    
    loadedFile = Empty
    loadedFileText = Empty
    cmdSaveFile.Enabled = False
    
    If lstFilter.Visible Then
        txt = lstFilter.List(lstFilter.ListIndex)
    Else
        txt = List1.List(List1.ListIndex)
    End If
    
    If Len(txt) = 0 Then
        'they did not select an entry, but there is only one filtered result so thats it..
        'user action: the entered a filter, saw one result and hit return
        If lstFilter.Visible And lstFilter.ListCount = 1 Then
            txt = lstFilter.List(0)
        End If
    End If
    
    If Len(txt) = 0 Then Exit Sub
    
    cid = Mid(txt, InStrRev(txt, "@") + 1, Len(txt))
    
    If optFile.Value Then
        If cn.State <> 1 Then cn.Open
        adors.Open "SELECT * FROM CodeDB where ID=" & cid, cn, adOpenKeyset, adLockOptimistic
        tmp = Environ("temp") & "\tmp.code"
        If FileExists(tmp) Then Kill tmp
        LoadFileFromDB tmp, adors, "CODE"
        If FileExists(tmp) Then
            loadedFile = Trim(Mid(txt, 1, InStr(txt, "@") - 1))
            loadedFileText = ReadFile(tmp)
            Text2 = loadedFileText
            Kill tmp
            Text2.selStart = 1
            cmdSaveFile.Enabled = True
        End If
        adors.Close
    Else
        If cn.State <> 1 Then cn.Open
        adors.Open "SELECT * FROM CodeDB where ID=" & cid, cn, adOpenKeyset, adLockOptimistic
        Text2.Text = Text2.Text & vbCrLf & vbCrLf & adors("CODE")
    End If
    
    cn.Close
    
    
    
End Sub

Function firstLine(it)
    t = Split(it, vbCrLf)
    firstLine = t(0)
End Function

Private Sub lblClose_Click()
    fraAdd.Visible = False
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
End Sub

Private Sub Option1_Click()
    cboLang_Click
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
    c = Clipboard.GetText
    If c <> Empty Then Text4 = c: Command1_Click 0
End Sub

Private Sub Text1_Change()

    'List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRING, -1, ByVal CStr(Text1.Text))
    
    If Len(Text1) = 0 Then
        lstFilter.Visible = False
    Else
        lstFilter.Visible = True
        Dim i As Long
        lstFilter.Clear
        For i = 0 To List1.ListCount - 1
            If InStr(1, List1.List(i), Text1, vbTextCompare) > 0 Then
                lstFilter.AddItem List1.List(i)
            End If
        Next
    End If
    
        
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call CopyCode
End Sub

Private Sub List1_DblClick()
    Call CopyCode
End Sub

Private Sub lstFilter_Click()
    Call CopyCode
End Sub

Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    txtFile.Text = Data.Files(1)
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
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Function ReadFile(FileName)
  f = FreeFile
  temp = ""
   Open FileName For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(FileName), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function



Function FileNameFromPath(fullpath) As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

