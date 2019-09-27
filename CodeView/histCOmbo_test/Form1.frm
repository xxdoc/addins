VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin Project1.HistoryCombo hist 
      Height          =   420
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   4110
      _extentx        =   7250
      _extenty        =   741
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    hist.LoadHistory App.path & "\_VBDecompiler.vbp.search.txt"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    hist.SaveHistory
End Sub

Private Sub hist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        hist.RecordIfNew
        Debug.Print "save if new"
    End If
End Sub
