VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fast Build Addin             http://sandsprite.com"
   ClientHeight    =   4950
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDisplayAsHex 
      Caption         =   "ints as hex in dbg tooltips"
      Height          =   315
      Left            =   150
      TabIndex        =   19
      Top             =   1350
      Width           =   2190
   End
   Begin VB.CheckBox chkShowPostBuildOutput 
      Caption         =   "Show Build Output"
      Height          =   195
      Left            =   2340
      TabIndex        =   18
      Top             =   1395
      Width           =   1725
   End
   Begin VB.CommandButton cmdSaveExec 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8865
      TabIndex        =   17
      Top             =   585
      Width           =   1095
   End
   Begin VB.TextBox txtExecute 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2385
      OLEDropMode     =   1  'Manual
      TabIndex        =   16
      Top             =   540
      Width           =   6360
   End
   Begin VB.CheckBox chkClearImmediate 
      Caption         =   "Clear Immediate Window On Start"
      Height          =   375
      Left            =   4275
      TabIndex        =   14
      Top             =   1305
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Last CMD Output"
      Height          =   285
      Left            =   7155
      TabIndex        =   13
      Top             =   1350
      Width           =   1590
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   285
      Left            =   8865
      TabIndex        =   12
      Top             =   1350
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8865
      TabIndex        =   11
      Top             =   990
      Width           =   1095
   End
   Begin VB.TextBox txtPostBuild 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2385
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Top             =   945
      Width           =   6360
   End
   Begin VB.TextBox txtAbout 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1710
      Width           =   9870
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   135
      TabIndex        =   4
      Top             =   7425
      Width           =   9825
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   90
         Width           =   1050
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   7470
         TabIndex        =   6
         Top             =   90
         Width           =   1050
      End
      Begin VB.CommandButton cmdClearLog 
         Caption         =   "Clear Log"
         Height          =   375
         Left            =   8640
         TabIndex        =   5
         Top             =   90
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8865
      TabIndex        =   3
      Top             =   180
      Width           =   1095
   End
   Begin VB.TextBox txtBuildPath 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2385
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   180
      Width           =   6360
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   135
      TabIndex        =   0
      Top             =   5130
      Width           =   9825
   End
   Begin VB.Label Label1 
      Caption         =   "Execute Btn CmdLine:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   45
      TabIndex        =   15
      Top             =   585
      Width           =   2490
   End
   Begin VB.Label Label1 
      Caption         =   "Post Build Command"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   45
      TabIndex        =   9
      Top             =   990
      Width           =   2220
   End
   Begin VB.Label Label1 
      Caption         =   "Default Build Path:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   225
      Width           =   2400
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Name:   VB6 FastBuild Addin
' Author: David Zimmer
' Site:   http://sandsprite.com
'
Dim loaded As Boolean


Private Sub chkDisplayAsHex_Click()
        
    If Not loaded Then Exit Sub
    
    If chkDisplayAsHex.value = 0 Then
        SaveSetting "FastBuild", "Settings", "DisplayAsHex", 0
    Else
        SaveSetting "FastBuild", "Settings", "DisplayAsHex", 1
        If Not LoadHexToolTipsDll() Then
            MsgBox "Could not find hextooTip.dll?"
            chkDisplayAsHex.value = 0
        End If
    End If
    
End Sub


'Private Sub chkClearImmediate_Click()
'
'    ClearImmediateOnStart = chkClearImmediate.value
'    SaveSetting "fastbuild", "settings", "ClearImmediateOnStart", chkClearImmediate.value
'
'    If loaded And ClearImmediateOnStart = 1 Then
'        MsgBox "Change takes effect next time to start IDE"
'    End If
'
'End Sub

Private Sub chkShowPostBuildOutput_Click()
    ShowPostBuildOutput = chkShowPostBuildOutput.value
    SaveSetting "fastbuild", "settings", "ShowPostBuildOutput", chkShowPostBuildOutput.value
End Sub

Private Sub cmdSaveExec_Click()
    On Error Resume Next
    VBInstance.ActiveVBProject.WriteProperty "fastBuild", "ExecBtnCmdLine", Trim(txtExecute)
End Sub

Private Sub cmdTest_Click()
    Dim postbuild As String
    Dim ret As String
    Dim homeDir As String
    
    SetHomeDir
    postbuild = ExpandVars(txtPostBuild, txtBuildPath)
    ret = GetCommandOutput("cmd /c " & postbuild, True, True)
    
    txtAbout = "Expanded command script: " & vbCrLf & postbuild & vbCrLf & vbCrLf & _
               "Command output: " & vbCrLf & ret
    
End Sub

Private Sub cmdUpdate_Click()
     On Error Resume Next
     
     'If Len(txtBuildPath) = 0 Then
     '   MsgBox "You must enter a valid path including file name and extension", vbExclamation
     '   Exit Sub
     'End If
     
     txtBuildPath = Replace(txtBuildPath, "/", "\")
     
     VBInstance.ActiveVBProject.WriteProperty "fastBuild", "fullPath", txtBuildPath
     
     If Err.Number = 0 Then
        txtAbout = "This build path has been set as default, you will not be prompted " & vbCrLf & _
                    "again and must change it here or in the projects vbp file as long " & vbCrLf & _
                    "as this plugin is active. " & vbCrLf & vbCrLf & _
                    "Note you can use relative paths such as .\file.exe or .\..\file.exe"
     Else
        MsgBox "Save failed: " & Err.Description, vbCritical
     End If
        
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    VBInstance.ActiveVBProject.WriteProperty "fastBuild", "PostBuild", txtPostBuild
    If Err.Number <> 0 Then
        MsgBox "Error saving postbuild command: " & Err.Description
    End If
End Sub

Private Sub Command2_Click()
    txtAbout = Module2.LastCommandOutput
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim x
    
    If isBuildPathSet() Then
        txtBuildPath = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
    End If
        
    txtExecute = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "ExecBtnCmdLine")
    
    chkClearImmediate.value = ClearImmediateOnStart
    chkShowPostBuildOutput.value = ShowPostBuildOutput
    
    txtPostBuild = GetPostBuildCommand
    
    chkDisplayAsHex.value = IIf(GetSetting("FastBuild", "Settings", "DisplayAsHex", 0) = "1", 1, 0)
    
    loaded = True
    
     txtAbout = "Build Path once set will be used from then on out automatically as default" & vbCrLf & _
                "and you will no longer be prompted every single time you want to generate the " & vbCrLf & _
                "executable." & vbCrLf & _
                "" & vbCrLf & _
                "The post build command allows you to specify a command (or batch file) that" & vbCrLf & _
                "you want run after your executable is built. Click the last cmd output " & vbCrLf & _
                "button to see the results, or test to watch it live. This command" & vbCrLf & _
                " supports several envirnoment variables which it can expand: " & vbCrLf & _
                "" & vbCrLf & _
                "%1     = exe FullPath" & vbCrLf & _
                "%app   = exe Home dir " & vbCrLf & _
                "%fname = exe file name" & vbCrLf
                
End Sub



Private Sub txtBuildPath_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    Dim pth As String
    
    pth = data.Files(1)
    
    If FileExists(pth) Then
        txtBuildPath = pth
        Exit Sub
    End If
    
    If FolderExists(pth) Then
        txtBuildPath = pth
    End If
    
End Sub

Private Sub txtPostBuild_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    On Error Resume Next
    Dim pth As String
    
    pth = data.Files(1)
    
    If FileExists(pth) Then
        txtPostBuild = pth
        Exit Sub
    End If
    
End Sub



'Private Sub cmdAbout_Click()
'    txtAbout.Visible = Not txtAbout.Visible
'End Sub
'
'Private Sub cmdClearLog_Click()
'    Set hookLog = New Collection
'    List1.Clear
'End Sub
'
'Public Sub cmdRefresh_Click()
'    List1.Clear
'    For Each X In Module1.hookLog
'        List1.AddItem X
'    Next
'End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'
'    Frame1.Top = Me.Height - Frame1.Height - 600
'    List1.Height = Frame1.Top - List1.Top - 200
'    List1.Width = Me.Width - List1.Left - 200
'    Frame1.Left = Me.Width - Frame1.Width - 200
'
'    With List1
'        txtAbout.Move .Left, .Top, .Width, .Height
'    End With
'
'End Sub
