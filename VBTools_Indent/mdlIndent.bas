Attribute VB_Name = "mdlIndent"
Option Explicit

'
' Made By Michael Ciurescu (CVMichael)
'

Private BlockStart() As String
Private BlockEnd() As String
Private BlockMiddle() As String


Public Sub IndentInitialize()
    BlockStart = Split("If * Then;For *;Do *;Do;Select Case *;While *;With *;Private Function *;Public Function *;Friend Function *;Function *;Private Sub *;Public Sub *;Friend Sub *;Sub *;Private Property *;Public Property *;Friend Property *;Property *;Private Enum *;Public Enum *;Friend Enum *;Enum *;Private Type *;Public Type *;Friend Type *;Type *", ";")
    BlockEnd = Split("End If;Next*;Loop;Loop *;End Select;Wend;End With;End Function;End Function;End Function;End Function;End Sub;End Sub;End Sub;End Sub;End Property;End Property;End Property;End Property;End Enum;End Enum;End Enum;End Enum;End Type;End Type;End Type;End Type", ";")
    
    BlockMiddle = Split("ElseIf * Then;Else;Case *", ";")
End Sub

Public Sub IndentCodeModule(CModule As VBIDE.CodeModule)
    Dim AllCode As String
    Dim AllLines() As String, LineNumbers() As String
    Dim K As Long, Q As Long
    Dim VBLine As String
    Dim StartCursorPos As Long, TopLine As Long
    
    With CModule
        TopLine = .CodePane.TopLine
        .CodePane.GetSelection StartCursorPos, 0, 0, 0
        
        AllCode = .Lines(1, .CountOfLines)
        
        AllCode = Replace$(AllCode, "_" & vbNewLine, "")
        AllLines = Split(AllCode, vbNewLine)
        ReDim LineNumbers(UBound(AllLines))
        
        For K = 0 To UBound(AllLines)
            Q = InStr(1, AllLines(K), " ")
            If Q = 0 Then Q = InStr(1, AllLines(K), vbTab)
            
            If Q > 0 Then
                If CStr(Val(Left(AllLines(K), Q - 1))) = Left(AllLines(K), Q - 1) Then
                    LineNumbers(K) = Left(AllLines(K), Q - 1)
                    AllLines(K) = Mid$(AllLines(K), Q)
                End If
            End If
            
            AllLines(K) = Trim(AllLines(K))
            
            If Left(AllLines(K), 1) = "'" Then
                AllLines(K) = "' " & Trim$(Mid$(AllLines(K), 2))
            End If
        Next K
        
        IndentBlock AllLines
        
        For K = 0 To UBound(AllLines)
            VBLine = RemoveLineComments(AllLines(K))
            
            For Q = 0 To UBound(BlockMiddle)
                If Replace$(VBLine, vbTab, "") Like BlockMiddle(Q) And Left(VBLine, 1) = vbTab Then
                    AllLines(K) = Mid$(AllLines(K), 2)
                End If
            Next Q
        Next K
        
        For K = 0 To UBound(AllLines)
            If Len(LineNumbers(K)) > 0 Then
                AllLines(K) = LineNumbers(K) & vbTab & AllLines(K)
            End If
        Next K
        
        AllCode = Replace$(Join$(AllLines, vbNewLine), Chr(1), "")
        
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, AllCode
        
        CModule.CodePane.SetSelection StartCursorPos, 1, StartCursorPos, 1
        .CodePane.TopLine = TopLine
    End With
End Sub

Public Sub IndentBlock(VBLines() As String)
    Dim K As Long, Q As Long
    Dim EndPos As Long
    Dim VBLine As String
    Dim StartPos As Long
    Dim FoundStartEnd As Boolean
    
    Do
        StartPos = 0
        EndPos = UBound(VBLines)
        
        Do
            FoundStartEnd = False
            
            For K = EndPos To StartPos Step -1
                VBLine = RemoveLineComments(VBLines(K))
                
                If Len(VBLine) > 0 Then
                    For Q = 0 To UBound(BlockStart)
                        If VBLine Like BlockStart(Q) Then
                            StartPos = K + 1
                            FoundStartEnd = True
                            Exit For
                        End If
                    Next Q
                    
                    If Q <= UBound(BlockStart) Then Exit For
                End If
            Next K
            
            If FoundStartEnd Then
                For K = StartPos To EndPos
                    VBLine = RemoveLineComments(VBLines(K))
                    
                    If Len(VBLine) > 0 Then
                        If VBLine Like BlockEnd(Q) Then
                            EndPos = K - 1
                            FoundStartEnd = True
                            Exit For
                        End If
                    End If
                Next K
            End If
        Loop While FoundStartEnd
        
        If Not (Not FoundStartEnd And StartPos = 0 And EndPos = UBound(VBLines)) Then
'            Debug.Print StartPos, EndPos
            IndentLineBlock VBLines, StartPos, EndPos
        End If
    Loop Until Not FoundStartEnd And StartPos = 0 And EndPos = UBound(VBLines)
End Sub

Public Function RemoveLineComments(ByVal VBLine As String) As String
    Dim K As Long, Q As Long
    
    If InStr(1, VBLine, "'") > 0 Then
        K = 1
        Do
            If Mid$(VBLine, K, 1) = """" Then
                For Q = K + 1 To Len(VBLine)
                    If Mid$(VBLine, Q, 1) = """" Then
                        VBLine = Left(VBLine, K) & Mid$(VBLine, Q)
                        K = K + 1
                        Exit For
                    End If
                Next Q
            End If
            
            K = K + 1
        Loop Until K >= Len(VBLine)
        
        For K = 1 To Len(VBLine)
            If Mid$(VBLine, K, 1) = "'" Then
                VBLine = Left$(VBLine, K - 1)
                Exit For
            End If
        Next K
    End If
    
    RemoveLineComments = Trim$(VBLine)
End Function

Public Sub IndentLineBlock(VBLines() As String, ByVal StartLine As Long, ByVal EndLine As Long, Optional NoIndentChar As Byte = 1)
    Dim K As Long
    
    If NoIndentChar > 0 Then
        If StartLine > 0 Then VBLines(StartLine - 1) = Chr(NoIndentChar) & VBLines(StartLine - 1)
        If EndLine < UBound(VBLines) Then VBLines(EndLine + 1) = Chr(NoIndentChar) & VBLines(EndLine + 1)
    End If
    
    For K = StartLine To EndLine
        If Left(VBLines(K), 1) = "'" Then
            VBLines(K) = "' " & vbTab & Mid$(VBLines(K), 2)
        Else
            VBLines(K) = vbTab & VBLines(K)
        End If
    Next K
End Sub

