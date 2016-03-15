VERSION 5.00
Begin VB.Form frmAdo 
   Caption         =   "Ado Connection String Generator"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOut 
      Height          =   1635
      Left            =   90
      TabIndex        =   6
      Top             =   630
      Width           =   8610
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   285
      Left            =   6795
      TabIndex        =   5
      Top             =   135
      Width           =   1410
   End
   Begin VB.CheckBox chkAppPath 
      Caption         =   "app.path"
      Height          =   240
      Left            =   5265
      TabIndex        =   4
      Top             =   135
      Width           =   1185
   End
   Begin VB.TextBox txtDB 
      Height          =   330
      Left            =   3195
      TabIndex        =   3
      Top             =   90
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "DB"
      Height          =   240
      Left            =   2745
      TabIndex        =   2
      Top             =   135
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Type"
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   510
   End
End
Attribute VB_Name = "frmAdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'all but jet are odbc jet is ole
Enum dbServers
    Access
    JetAccess2k
    MsSql2k
    mysql
    DSN
    FileDsn
    dBase
    csvtext
End Enum

Private Sub cmdGenerate_Click()
    
    Dim i As Long
    
    tmp = Array("Access", "JetAccess2k", "MsSql2k", "mysql", "DSN", "FileDsn", "dBase", "csvtext")
    For i = 0 To UBound(tmp)
        If Combo1.Text = tmp(i) Then Exit For
    Next
    
    If i > UBound(tmp) Then
        MsgBox "Select db type", vbInformation
        Exit Sub
    End If
    
    db = txtDB
    If LCase(Right(db, 4)) <> ".mdb" Then db = db & ".mdb"
    
    If chkAppPath.Value = 1 Then
        db = """ & app.path & ""\" & db
    End If
    
    txtOut.Text = BuildConnectionString(i, CStr(db))
    txtOut.selLength = Len(txtOut.Text)
    
End Sub

Private Sub Form_Load()

    tmp = Array("Access", "JetAccess2k", "MsSql2k", "mysql", "DSN", "FileDsn", "dBase", "csvtext")
    For Each t In tmp
        Combo1.AddItem t
    Next
    Combo1.ListIndex = 0
       
End Sub


Function BuildConnectionString(dbServer As dbServers, dbName As String, Optional server As String, _
                          Optional Port = 3306, Optional User As String, Optional pass As String)
    
    Dim dbPath As String
    Dim baseString As String
    Dim blnInlineAuth As Boolean
    
    Select Case dbServer
        Case Access
            baseString = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=____;"
        Case FileDsn
            baseString = "FILEDSN=____;"
        Case DSN
            baseString = "DSN=____;"
        Case dBase
            baseString = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=____;"
        Case mysql
            baseString = "Driver={mySQL};Server=" & server & ";Port=" & Port & ";Stmt=;Option=16834;Database=____;"
        Case MsSql2k
            baseString = "Driver={SQL Server};Server=" & server & ";Database=____;"
        Case JetAccess2k
            baseString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=____;" & _
                         "User Id=" & User & ";" & _
                         "Password=" & pass & ";"
                         blnInlineAuth = True
        Case csvtext
                baseString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                             "Data Source=____;" & _
                             "Extended Properties=""text;HDR=YES;FMT=Delimited"""
    End Select
                         
        
    If Not blnInlineAuth Then
        If User <> Empty Then baseString = baseString & "Uid:" & User & ";"
        If pass <> Empty Then baseString = baseString & "Pwd:" & User & ";"
    End If
       
    '%AP% is like enviromental variable for app.path i am lazy :P
    dbPath = Replace(dbName, "%AP%", App.path)
    
    BuildConnectionString = "constr = """ & Replace(baseString, "____", dbPath) & """"
    
End Function



