Attribute VB_Name = "modVBE"
Option Explicit

Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
Private Declare Sub CopyMemByR Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)

' Application reference - do not store in the Connect designer
Public oVBE As VBIDE.VBE
Public oPopupMenu As Office.CommandBar

Private Const MAX_POS_ITEMS = 200& ' Bug fix 20 Aug 2016 - Nasty menu errors around 250 items!

Private Type tCallerHits
    sCaller As String          ' Component/member reference
    sMember As String          ' Member/proc name
    iDecLine As Long           ' Start of proc with reference
    iOffset As Long            ' Line offset of reference
    iCurCol As Long            ' Column of reference
    vbeKind As vbext_ProcKind  ' Proc kind of proc with reference
End Type

Private taCallers() As tCallerHits
Private taCallees() As tCallerHits ' Return to previous callees 3.6
Private cMenuItem() As cMenuItem
Private sTermStrs() As String

Private fExempt As Long
Private mCallee As Long ' Return to previous callees
Private mMenuCnt As Long
Private mCurIndex As Long
Private mBlockedName As String

Public nCallCnt As Long
Public nCallers As Long

Public Sub RedimCallers(ByVal iNewUB As Long)
    ReDim Preserve sTermStrs(0 To iNewUB) As String
    ReDim Preserve taCallers(0 To iNewUB) As tCallerHits
    ReDim Preserve taCallees(1 To iNewUB) As tCallerHits
    ReDim Preserve cMenuItem(1 To iNewUB) As cMenuItem
End Sub

Public Sub EraseCallerArrays()
    Erase sTermStrs()
    Erase taCallers()
    Erase taCallees()
    Erase cMenuItem()
End Sub

Public Sub ResetContextMenu()
   Do While mMenuCnt
      cMenuItem(mMenuCnt).Remove
      Set cMenuItem(mMenuCnt) = Nothing
      mMenuCnt = mMenuCnt - 1&
   Loop
   If Not oPopupMenu Is Nothing Then
      Call oPopupMenu.Delete
      Set oPopupMenu = Nothing
   End If
End Sub

Public Sub DisplayCallee()
    CodePaneMenuItem_Click 0&
End Sub

Public Sub CodePaneMenuItem_Click(ByVal Idx As Long)
    Dim sCompName As String
    Dim eKind As vbext_ProcKind
    Dim lStartLine As Long
    Dim lStartCol As Long
    Dim lTopLine As Long
    Dim i As Long, j As Long
  On Error GoTo ErrHandler

    If (nCallers = 0& Or mCurIndex = 0&) And Idx = 0& Then ' Previous callees
        mCallee = (mCallee + 1&) Mod nCallCnt
        i = nCallCnt - mCallee
        With taCallers(0)
            .sCaller = taCallees(i).sCaller
            .sMember = taCallees(i).sMember
            .iDecLine = taCallees(i).iDecLine
            .iOffset = taCallees(i).iOffset
            .iCurCol = taCallees(i).iCurCol
            .vbeKind = taCallees(i).vbeKind
        End With
    Else
        mCallee = 0&
    End If
    mCurIndex = Idx

    With taCallers(Idx)
        i = InStr(1, .sCaller, ".") ' Get component reference
        If i = 0& Then
            i = InStr(1, taCallers(0).sCaller, ".")
            sCompName = Left$(taCallers(0).sCaller, i - 1&)
        Else
            sCompName = Left$(.sCaller, i - 1&)
        End If
        lStartLine = .iOffset ' Reference offset in procedure
        lStartCol = .iCurCol
        eKind = .vbeKind
    End With

  On Error GoTo ErrWith
    With oVBE.ActiveVBProject.VBComponents(sCompName).CodeModule
       If eKind <> -1& Then ' Catch changes in procedure offset
           lStartLine = lStartLine + .ProcBodyLine(taCallers(Idx).sMember, eKind)
       Else
           lStartLine = lStartLine + taCallers(Idx).iDecLine
       End If
       i = .CodePane.TopLine
       j = .CodePane.CountOfVisibleLines
       If Not (lStartLine > i And lStartLine < i + j) Then
           lTopLine = lStartLine - CLng(j / 2.5)
           If lTopLine < 1 Then lTopLine = 1
          .CodePane.TopLine = lTopLine
       End If
       Call .CodePane.SetSelection(lStartLine, lStartCol, lStartLine, lStartCol)
       If Not oVBE.ActiveCodePane Is .CodePane Then
          .Parent.Activate ' Activate the component
       End If
      .CodePane.Show
ErrWith:
    End With
ErrHandler:
 If Err = 9 Then
    ' Occurs when stored callers are deleted or the module removed
 ElseIf Err Then
    LogError "modVBE.CodePaneMenuItem_Click", taCallers(Idx).sCaller
 End If
End Sub

Private Sub AddNewRef(sProcRef As String, sMember As String, ByVal lDecLine As Long, ByVal lCurLine As Long, ByVal lCurCol As Long, ByVal eKind As vbext_ProcKind)
    nCallers = nCallers + 1&
    If nCallers > UBound(taCallers) Then RedimCallers nCallers + 100&
    With taCallers(nCallers)
        .sCaller = sProcRef
        .sMember = sMember
        .iDecLine = lDecLine
        .iOffset = lCurLine - lDecLine
        .iCurCol = lCurCol
        .vbeKind = eKind
    End With
End Sub

Private Sub RecordCallee(sProcRef As String, sMember As String, ByVal lDecLine As Long, ByVal lCurLine As Long, ByVal lCurCol As Long, ByVal eKind As vbext_ProcKind)
    With taCallers(0)
        .sCaller = sProcRef
        .sMember = sMember
        .iDecLine = lDecLine
        .iOffset = lCurLine - lDecLine
        .iCurCol = lCurCol
        .vbeKind = eKind
    End With
    nCallCnt = nCallCnt + 1&
    If nCallCnt > UBound(taCallees) Then
        ReDim Preserve taCallees(1 To nCallCnt + 100) As tCallerHits
    End If
    With taCallees(nCallCnt) ' Add to previous callees
        .sCaller = sProcRef
        .sMember = sMember
        .iDecLine = lDecLine
        .iOffset = lCurLine - lDecLine
        .iCurCol = lCurCol
        .vbeKind = eKind
    End With
End Sub

Public Sub JumpToPrevReference()
    If mCurIndex = 1& Then mCurIndex = nCallers + 1& ' Cycle references
    CodePaneMenuItem_Click mCurIndex - 1&
End Sub

Public Sub JumpToNextReference()
    If mCurIndex = nCallers Then mCurIndex = 0& ' Cycle references
    CodePaneMenuItem_Click mCurIndex + 1&
End Sub

Public Sub RefreshComponentMembers(ByVal eMembType As vbext_MemberType, sMemberType As String)
   Dim oCodeMod As CodeModule
   Dim oMember As Member
   Dim vbeKind As vbext_ProcKind
   Dim lDeclarationL As Long
   Dim lCurrentLine As Long
   Dim lCurrentCol As Long
   Dim lCntDecLines As Long
   Dim lEndLine As Long
   Dim sTempLine As String
   Dim sProcKind As String
   Dim sMemberName As String
   Dim sMemberLine As String
   Dim sFoundName As String
   Dim sTermName As String

   Dim i As Long, j As Long
   Dim idxA() As Long, k As Long

  On Error GoTo ErrHandler
   ' Exit if we're not in the code pane
   If oVBE.ActiveCodePane Is Nothing Then Exit Sub

  ' Retrieve the current line the cursor is on
   oVBE.ActiveCodePane.GetSelection lCurrentLine, lCurrentCol, j, k

   Set oCodeMod = oVBE.ActiveCodePane.CodeModule

   ' Callee just returns to current position
   RecordCallee oCodeMod.Parent.Name & ".", vbNullString, lCurrentLine, lCurrentLine, lCurrentCol, -1&
   nCallers = 0
   Call ResetContextMenu

  On Error GoTo EndWith
   With oCodeMod

      lCntDecLines = .CountOfDeclarationLines + 1&
      vbeKind = -1&

      For i = 1 To .Members.Count
         sMemberName = .Members(i).Name
         If eMembType = .Members(i).Type Then
            lCurrentLine = 1&
            lCurrentCol = 1&
            lEndLine = lCntDecLines
            Select Case eMembType
               Case Is = vbext_mt_Const
                   If .Find("Const " & sMemberName & " ", lCurrentLine, lCurrentCol, lEndLine, -1&, , True) Then
                       sFoundName = "Const  " & sMemberName
                       sTermName = sMemberName
                       lCurrentCol = lCurrentCol + 6
                       lDeclarationL = lCurrentLine
                   End If
               Case Is = vbext_mt_Event
                   If .Find("Event " & sMemberName & "(", lCurrentLine, lCurrentCol, lEndLine, -1&, , True) Then
                       sFoundName = "Event " & sMemberName
                       sTermName = "E" & sMemberName
                       lCurrentCol = lCurrentCol + 6
                       lDeclarationL = lCurrentLine
                   End If
               Case Is = vbext_mt_Variable
                   If .Find(" " & sMemberName & " As ", lCurrentLine, lCurrentCol, lEndLine, -1&, , True) Then
                       sFoundName = " " & sMemberName
                       sTermName = sMemberName
                   ElseIf .Find(" " & sMemberName, lCurrentLine, lCurrentCol, lEndLine, -1&, , True) Then
                        sTempLine = .Lines(lCurrentLine, 1&)
                        If InStr(sTempLine, "(") Then
                           sFoundName = " " & sMemberName & "()"
                        Else
                           sFoundName = " " & sMemberName
                        End If
                        sTermName = sMemberName
                   End If
                   lCurrentCol = lCurrentCol + 1
                   lDeclarationL = lCurrentLine
               Case Is = vbext_mt_Method
                   lCurrentLine = .ProcBodyLine(sMemberName, vbext_pk_Proc)
                   lDeclarationL = lCurrentLine
                   If lCurrentLine = 1& Then
                        .Find sMemberName & " Lib", lCurrentLine, lCurrentCol, -1&, -1&, , True
                        sFoundName = sMemberName & "  (Declare Lib)"
                        sTermName = "A" & sMemberName
                   Else
                        sMemberLine = Trim$(.Lines(lCurrentLine, 1&))
                        j = InStr(1&, sMemberLine, sMemberName)
                        sFoundName = sMemberName & "  (" & LTrim$(Left$(sMemberLine, j - 2&)) & ")"
                        sTermName = "B" & sMemberName
                        lCurrentCol = j
                   End If
                   vbeKind = vbext_pk_Proc
               Case Is = vbext_mt_Property
                   lCurrentLine = lCntDecLines
                   Do While .Find(" " & sMemberName & "(", lCurrentLine, lCurrentCol, -1&, -1&, , True)
                       If sMemberName = .ProcOfLine(lCurrentLine, vbeKind) Then ' Gets vbeKind
                           lCurrentLine = .ProcBodyLine(sMemberName, vbeKind)
                           Select Case vbeKind
                              Case Is = vbext_pk_Get:  sProcKind = "(Get)  "
                              Case Is = vbext_pk_Let:  sProcKind = "(Let)   "
                              Case Is = vbext_pk_Set:  sProcKind = "(Set)  "
                           End Select ' Unique name for this prop
                           AddNewRef sProcKind & sMemberName, sMemberName, lCurrentLine, lCurrentLine, lCurrentCol + 1, vbeKind
                           sTermStrs(nCallers) = sMemberName & sProcKind
                           lCurrentLine = .ProcStartLine(sMemberName, vbeKind) + .ProcCountLines(sMemberName, vbeKind)
                       Else ' Shouldn't happen?
                           lCurrentLine = lCurrentLine + 1&
                       End If
                   Loop
            End Select
            If LenB(sFoundName) <> 0& Then
               AddNewRef sFoundName, sMemberName, lDeclarationL, lCurrentLine, lCurrentCol, vbeKind
               sTermStrs(nCallers) = sTermName
               sFoundName = vbNullString
            End If
         End If
      Next i
      If eMembType = vbext_mt_Event Then ' Check for Implements objects
         lCurrentLine = 1&
         lCurrentCol = 1&
         lEndLine = lCntDecLines
         Do While .Find("Implements ", lCurrentLine, lCurrentCol, lEndLine, -1&, , True)
            sTempLine = .Lines(lCurrentLine, 1&)
            If Left$(Trim$(sTempLine), 11&) = "Implements " Then
                j = lCurrentCol + 11&
                k = InStr(j, sTempLine, " ")
                If k = 0& Then
                   sMemberName = Mid$(sTempLine, j)
                Else
                   sMemberName = Mid$(sTempLine, j, k - j)
                End If
                AddNewRef "Implements " & sMemberName, sMemberName, lCurrentLine, lCurrentLine, j, -1&
                sTermStrs(nCallers) = "A" & sMemberName
            End If
            lCurrentLine = lCurrentLine + 1&
            lCurrentCol = 1&
            lEndLine = lCntDecLines
         Loop
         lCurrentLine = 1&
         lCurrentCol = 1&
         lEndLine = lCntDecLines
         Do While .Find(" WithEvents ", lCurrentLine, lCurrentCol, lEndLine, -1&, , True)
            sTempLine = .Lines(lCurrentLine, 1&)
            If IsCode(sTempLine, lCurrentCol, lCurrentCol) Then
                j = lCurrentCol + 12&
                k = InStr(j, sTempLine, " ")
                If k = 0& Then ' Can't happen?
                   sMemberName = Mid$(sTempLine, j)
                Else
                   sMemberName = Mid$(sTempLine, j, k - j)
                End If
                AddNewRef "WithEvents " & sMemberName, sMemberName, lCurrentLine, lCurrentLine, j, -1&
                sTermStrs(nCallers) = sMemberName
                lCurrentCol = k
            Else
                lCurrentLine = lCurrentLine + 1&
                lCurrentCol = 1&
            End If
            lEndLine = lCntDecLines
        Loop
     ElseIf eMembType = 6 Then ' Check for Enums and Types
        lCurrentLine = 1&
        lCurrentCol = 1&
        lEndLine = lCntDecLines
        Do While .Find("Enum ", lCurrentLine, lCurrentCol, lEndLine, -1&, , True)
            sTempLine = .Lines(lCurrentLine, 1&)
            If InStr(" " & sTempLine, " Enum ") Then
                If IsCode(sTempLine, lCurrentCol, lCurrentCol) Then
                    j = lCurrentCol + 5&
                    k = InStr(j, sTempLine, " ")
                    If k = 0& Then
                       sMemberName = Mid$(sTempLine, j)
                    Else
                       sMemberName = Mid$(sTempLine, j, k - j)
                    End If
                    AddNewRef "Enum  " & sMemberName, sMemberName, lCurrentLine, lCurrentLine, j, -1&
                    sTermStrs(nCallers) = "E" & sMemberName
                    If Not .Find("End Enum", lCurrentLine, lCurrentCol, lEndLine, -1&, , True) Then
                        lCurrentLine = taCallers(nCallers).iDecLine
                       'lCurrentCol = taCallers(nCallers).iCurCol
                        nCallers = nCallers - 1&
                    End If
                End If
            End If
            lCurrentLine = lCurrentLine + 1&
            lCurrentCol = 1&
            lEndLine = lCntDecLines
        Loop
        lCurrentLine = 1&
        lCurrentCol = 1&
        lEndLine = lCntDecLines
        Do While .Find("Type ", lCurrentLine, lCurrentCol, lEndLine, -1&, , True)
            sTempLine = .Lines(lCurrentLine, 1&)
            If InStr(" " & sTempLine, " Type ") Then
                If IsCode(sTempLine, lCurrentCol, lCurrentCol) Then
                    j = lCurrentCol + 5&
                    k = InStr(j, sTempLine, " ")
                    If k = 0& Then
                       sMemberName = Mid$(sTempLine, j)
                    Else
                       sMemberName = Mid$(sTempLine, j, k - j)
                    End If
                    AddNewRef "Type   " & sMemberName, sMemberName, lCurrentLine, lCurrentLine, j, -1&
                    sTermStrs(nCallers) = "T" & sMemberName
                    If Not .Find("End Type", lCurrentLine, lCurrentCol, lEndLine, -1&, , True) Then
                        lCurrentLine = taCallers(nCallers).iDecLine
                       'lCurrentCol = taCallers(nCallers).iCurCol
                        nCallers = nCallers - 1&
                    End If
                End If
            End If
            lCurrentLine = lCurrentLine + 1&
            lCurrentCol = 1&
            lEndLine = lCntDecLines
        Loop
     End If
EndWith:
    End With
  On Error GoTo ErrHandler
    Set oCodeMod = Nothing

    If nCallers Then
       ReDim idxA(nCallers) As Long
       For j = 1 To nCallers
           idxA(j) = j ' Initialize the index array
       Next
       strShellIndexed sTermStrs, idxA

      ' (Re-)create the popup menu using the Position argument
       Set oPopupMenu = oVBE.CommandBars.Add(Name:="Members", Position:=msoBarPopup)

       k = 0&
       sFoundName = vbNullString
       For j = 1 To nCallers
           i = idxA(j) ' Sorted lookup

           If taCallers(i).sCaller <> sFoundName Then ' Add to menu only the first hit in a proc

                sFoundName = taCallers(i).sCaller
                k = k + 1&
                Set cMenuItem(k) = New cMenuItem ' Create new item in popup-menu
                cMenuItem(k).Add oPopupMenu, sFoundName, i, LoadResPicture(sMemberType, vbResBitmap)

                If k = MAX_POS_ITEMS Then Exit For
           End If
       Next
       mMenuCnt = k
       mCurIndex = 0
    End If

ErrHandler:
   Set oCodeMod = Nothing
End Sub

Public Sub RefreshMemberReferences()
          ' Adapted from Add Error Handling addin by Kamilche
          Dim eProcKind As vbext_ProcKind
          Dim oComp As VBComponent
          Dim oCodePane As CodePane
          Dim oThisMod As CodeModule
          Dim oNextMod As CodeModule
          Dim oMember As Member

          Dim i As Long, j As Long, k As Long
          Dim lDeclarationL As Long
          Dim lCurrentLine As Long
          Dim lCurrentCol As Long
          Dim vbeScope As vbext_Scope
          Dim sMembName As String
          Dim sCompName As String
          Dim sProvName As String
          Dim sProcRef As String
          Dim sCodeLine As String

130      On Error GoTo ErrHandler
140        Set oCodePane = oVBE.ActiveCodePane
          ' Exit if we're not in the code pane
150        If oCodePane Is Nothing Then Exit Sub
160        Set oThisMod = oCodePane.CodeModule

          ' Retrieve the current line the cursor is on
170        oCodePane.GetSelection lCurrentLine, lCurrentCol, j, k

          ' Retrieve the procedure name and the ProcKind
180        sMembName = oThisMod.ProcOfLine(lCurrentLine, eProcKind)

           fExempt = 0&
           If LenB(sMembName) Then
185            vbeScope = oThisMod.Members(sMembName).Scope
               lDeclarationL = oThisMod.ProcBodyLine(sMembName, eProcKind)
           Else 'LenB(sMembName) = 0
               eProcKind = -1& ' Not a procedure, find member declared name
               lDeclarationL = lCurrentLine
190            sMembName = GetDeclareName(oThisMod, lDeclarationL, lCurrentCol, vbeScope, eProcKind)
           End If

200        If LenB(sMembName) Then

210           mBlockedName = vbNullString
              sCompName = oThisMod.Parent.Name
220           sProcRef = sCompName & "." & sMembName

230           RecordCallee sProcRef, sMembName, lDeclarationL, lCurrentLine, lCurrentCol, eProcKind
240           nCallers = 0&
260           Call ResetContextMenu
              '(bug fix 27 July 2016) ImplementedEvent:
HandleEvents: '(bug fix 19 Aug 2016) WithEvents:

             ' Search current code module for references to current member
270           FindCallers oThisMod, sMembName, vbNullString, lCurrentLine

280           If vbeScope <> vbext_Private Then ' vbext_Friend | vbext_Public

                 ' Search the other components for references to current member
290               For Each oComp In oVBE.ActiveVBProject.VBComponents
300                  If Not oComp Is Nothing Then
310                   If Not oComp.Name = vbNullString Then
320                    If Not oComp Is oThisMod.Parent Then
                         On Error Resume Next
330                       Set oNextMod = Nothing  ' Bug fix Dec 18, 2010
                           Select Case oComp.Type
                              Case vbext_ct_RelatedDocument, vbext_ct_ResFile
                               ' Related docs, res files throw exception
                              Case Else
340                            Set oNextMod = oComp.CodeModule
                              'Case vbext_ct_DocObject
                              'Case vbext_ct_ClassModule
                              'Case vbext_ct_MSForm
                              'Case vbext_ct_PropPage
                              'Case vbext_ct_StdModule
                              'Case vbext_ct_UserControl
                              'Case vbext_ct_VBForm
                              'Case vbext_ct_VBMDIForm
                              'Case vbext_ct_ActiveXDesigner
                           End Select
                          On Error GoTo ErrHandler

350                        If Not oNextMod Is Nothing Then
360                            FindCallers oNextMod, sMembName, sCompName
370                        End If
380                     End If
                      End If
                    End If
390               Next oComp
400           End If

410           If nCallers Then
                 ' (Re-)create the popup menu using the Position argument
420               Set oPopupMenu = oVBE.CommandBars.Add(Name:="Callers", Position:=msoBarPopup)
430
                  k = 0&
                  sProcRef = vbNullString
                  For i = 1 To nCallers
440
                      If taCallers(i).sCaller <> sProcRef Then ' Add to menu only the first hit in a proc
450
                           sProcRef = taCallers(i).sCaller
                           k = k + 1&
                           Set cMenuItem(k) = New cMenuItem ' Create new item in popup-menu
                           cMenuItem(k).Add oPopupMenu, sProcRef, i, LoadResPicture("Right", vbResBitmap)
460
                           If k = MAX_POS_ITEMS Then Exit For
                      End If
                  Next
                  mMenuCnt = k
                  mCurIndex = 0&

470           Else ' Check if the receiver of an Implemented event is seeking to know who raises it
                  i = InStr(1, sMembName, "_") ' Fix Implemented events not identifying component
                  If i <> 0 Then               ' that Implements/Raises the events - 26 July, 2016
                      mBlockedName = sCompName ' Block receiver of the event from being identified
                      sProvName = Left$(sMembName, i - 1&) ' Extract component? name from event name
                      sMembName = Mid$(sMembName, i + 1&)  ' Get member name out from event name
                      Set oMember = Nothing
                      Set oComp = Nothing
                      On Error Resume Next ' If we're wrong will error here (Form events, WithEvents var)
                      Set oComp = oVBE.ActiveVBProject.VBComponents(sProvName)
                      If Not oComp Is Nothing Then         ' Make sure we're not chasing our tail
                          Set oThisMod = oComp.CodeModule
                          Set oMember = oThisMod.Members(sMembName)
                          If Not oMember Is Nothing Then   ' Is it a code member
                              lCurrentLine = oMember.CodeLocation
                              vbeScope = oMember.Scope
                              sCompName = sProvName
                              sProcRef = sCompName & "." & sMembName
                              On Error GoTo ErrHandler
                              GoTo HandleEvents   ' Now find callers of the Implements class
                          End If
                      Else ' WithEvents?  Bug fix 19 Aug 2016
                          Set oMember = oThisMod.Members(sProvName) ' WithEvents var name?
                          If Not oMember Is Nothing Then            ' Is it a code member
                              lCurrentLine = 1&: lCurrentCol = 1&   ' Can't trust CodeLocation prop
                              k = oThisMod.CountOfDeclarationLines  ' Get declaration line
                              Do While oThisMod.Find("WithEvents " & sProvName, lCurrentLine, lCurrentCol, k, -1&, , True)
                                  k = lCurrentCol + Len("WithEvents " & sProvName)
                                  sCodeLine = oThisMod.Lines(lCurrentLine, 1)
                                  If IsCode(sCodeLine, lCurrentCol, k) Then ' Only if it's code
                                      vbeScope = oMember.Scope
                                      i = InStr(k, sCodeLine, " As ") + 4&
                                      j = InStr(i, sCodeLine, " ")
                                      If j = 0& Then j = Len(sCodeLine) + 1&
                                      sCompName = Mid$(sCodeLine, i, j - i)
                                      sProcRef = sCompName & "." & sMembName ' Event source Comp and Member
                                      Set oComp = oVBE.ActiveVBProject.VBComponents(sCompName)
                                      If Not oComp Is Nothing Then  ' Make sure we're not chasing our tail
                                          Set oThisMod = oComp.CodeModule
                                          lCurrentLine = oThisMod.Members(sMembName).CodeLocation
                                          On Error GoTo ErrHandler
                                          GoTo HandleEvents   ' Now find callers of the source Event
                                      End If
                                  End If
                                  lCurrentCol = k
                                  k = oThisMod.CountOfDeclarationLines
                              Loop
                          End If
                      End If
                      On Error GoTo 0
                  End If
              End If

480       Else 'LenB(sMembName) = 0
500           Call ResetContextMenu
              nCallers = 0
510       End If

ErrHandler:
      If Err Then LogError "modVBE.RefreshMemberReferences", sProcRef
End Sub

Private Function GetInstanceName(ByVal oCodeMod As CodeModule, sClassName As String, ByVal lStartLine As Long, ByVal lEndLine As Long, sInsts() As String) As Long
          Dim sCodeLine As String
          Dim lInstCnt As Long
          Dim lLineEnd As Long
          Dim lStartCol As Long
          Dim lEndColumn As Long
          Dim i As Long, j As Long
          Dim fTryAgain As Boolean
610     On Error GoTo EndWith
620       With oCodeMod
630        lStartCol = 1
           lEndColumn = -1
           lLineEnd = lEndLine
640        Do
            ' Search the procedure for class instantiation (also frm As New form, ctl As New ctrl, etc)
650          Do While .Find(sClassName, lStartLine, lStartCol, lEndLine, lEndColumn, True, True)
660              sCodeLine = .Lines(lStartLine, 1)
670              If IsCode(sCodeLine, lStartCol, lEndColumn) Then
680                  If lStartCol > 4 Then
690                      If Mid$(sCodeLine, lStartCol - 4, 4) = " As " Then
700                         j = 4   ' cClass  As 'sClassName'
                           'Do While IsDelim(Mid$(sCodeLine, lStartCol - j - 1, 1))
710                         Do While IsDelimI(MidI(sCodeLine, lStartCol - j - 1))
720                             j = j + 1
730                         Loop
740                         i = j + 1
750                         Do While Not IsDelimI(MidI(sCodeLine, lStartCol - i - 1))
760                             i = i + 1
770                         Loop
780                         ReDim Preserve sInsts(lInstCnt) As String
                            sInsts(lInstCnt) = Mid$(sCodeLine, lStartCol - i, i - j)
                            lInstCnt = lInstCnt + 1&
                            GoTo LoopNext
800                     End If
810                 End If
820                 If lStartCol > 8 Then
830                     If Mid$(sCodeLine, lStartCol - 8, 8) = " As New " Then
840                         j = 8   ' cClass  As New 'sClassName'
850                         Do While IsDelimI(MidI(sCodeLine, lStartCol - j - 1))
860                             j = j + 1
870                         Loop
880                         i = j + 1
890                         Do While Not IsDelimI(MidI(sCodeLine, lStartCol - i - 1))
900                             i = i + 1
910                         Loop
920                         ReDim Preserve sInsts(lInstCnt) As String
                            sInsts(lInstCnt) = Mid$(sCodeLine, lStartCol - i, i - j)
                            lInstCnt = lInstCnt + 1&
                            GoTo LoopNext
940                     End If
950                 End If
960                 If lStartCol > 7 Then
970                     If Mid$(sCodeLine, lStartCol - 7, 7) = " = New " Then
980                         i = InStr(sCodeLine, "Set ") + 4  'Set cClass = New 'sClassName'
990                         If i > 4 And i < lStartCol - 7 Then ' Bug fix Dec 12, 2010
1000                            ReDim Preserve sInsts(lInstCnt) As String
1010                            sInsts(lInstCnt) = Mid$(sCodeLine, i, lStartCol - 7 - i)
                                lInstCnt = lInstCnt + 1&
                                GoTo LoopNext
1020                         End If
1030                     End If
1040                 End If
1050                 If lStartCol > 3 Then
1060                     If Mid$(sCodeLine, lStartCol - 3, 3) = " = " Then
1070                         i = InStr(sCodeLine, "Set ") + 4  'Set cClass = 'sClassName'
1080                         If i > 4 And i < lStartCol - 3 Then ' Bug fix Dec 12, 2010
1090                            ReDim Preserve sInsts(lInstCnt) As String
1100                            sInsts(lInstCnt) = Mid$(sCodeLine, i, lStartCol - 3 - i)
                                lInstCnt = lInstCnt + 1&
1110                         End If
1120                     End If
1130                 End If
1140             End If
LoopNext:        lStartCol = lEndColumn + 1
1160             lEndColumn = -1
                 lEndLine = lLineEnd
1170         Loop

1180         fTryAgain = (lLineEnd > .CountOfDeclarationLines + 1)

            ' Search the component for class instantiation (also frm As New form, ctl As New ctrl, etc)
             If fTryAgain Then
               lStartLine = 1
               lLineEnd = .CountOfDeclarationLines + 1
               lEndLine = lLineEnd
               lStartCol = 1
               lEndColumn = -1
             End If

1190       Loop While fTryAgain
EndWith:
        End With
        GetInstanceName = lInstCnt
    If Err Then LogError "modVBE.GetInstanceName", sClassName
End Function

Private Sub FindCallers(ByVal oCodeMod As CodeModule, sMembName As String, sCompName As String, Optional ByVal lProcLine As Long)
          ' Adapted from Project References addin by ':) Ulli
          Dim eProcKind As vbext_ProcKind
          Dim oMember As Member
          Dim sProcName As String
          Dim sCodeLine As String
          Dim sCaller As String
          Dim lInstCnt As Long
          Dim lStartCol As Long
          Dim lEndColumn As Long
          Dim lLineStart As Long
          Dim lLineCount As Long
          Dim lLineLen As Long
          Dim lCodeLine As Long
          Dim lEndLine As Long
          Dim lContinue As Long
          Dim fIsCode As Long
          Dim fMustQualify As Long
          Dim i As Long, j As Long, k As Long
          ReDim sInstName(0) As String

1200     On Error GoTo EndWith
1210       With oCodeMod

            ' First search in the declarations section
1220         lCodeLine = 1
1230         lEndLine = .CountOfDeclarationLines + 1

1240         lStartCol = 1
1250         lEndColumn = -1

1260         Do While .Find(sMembName, lCodeLine, lStartCol, lEndLine, lEndColumn, True, True)

1270            sCodeLine = .Lines(lCodeLine, 1)
                ' The VBE's CodeModule Find function accepts an underscore as a delimiter
1280            If IsWholeWord(sCodeLine, lStartCol, lEndColumn) Then

1290               fIsCode = IsCode(sCodeLine, lStartCol, lEndColumn)
1300               If fIsCode Then

1310                  If LenB(sCompName) Then
1320                     If oVBE.ActiveVBProject.VBComponents(sCompName).Type <> vbext_ct_StdModule Then
1330                       If Not fExempt Then fMustQualify = -1
                         End If
                      End If

1340                  If IsValid(oCodeMod, lCodeLine, lStartCol, sCompName, lCodeLine, fMustQualify, sMembName, lProcLine) Then

1350                     lContinue = lCodeLine
1360                     Do While lContinue > 1& ' .ProcOfLine kinda thing
                           'If Right$(.Lines(lContinue - 1, 1), 1) = "_" Then
1370                        If RightI(.Lines(lContinue - 1, 1), 1) = 95 Then
1380                           lContinue = lContinue - 1& ' Line continuation
1390                           sCodeLine = .Lines(lContinue, 1)
1400                           lLineLen = Len(sCodeLine)                       ' Check for beginning a comment
1410                           fIsCode = IsCode(sCodeLine, lLineLen, lLineLen) ' before the end of the line
                               If Not fIsCode Then Exit Do
                            Else
                               Exit Do
                            End If
                         Loop
1420                     If fIsCode Then
1430                        If lContinue <> lProcLine Then ' lProcLine is zero if not current comp

1440                           sProcName = GetDeclareName(oCodeMod, lContinue, 1)
1450                           If LenB(sProcName) And (sProcName <> sMembName) Then

1480                              AddNewRef .Parent.Name & "." & sProcName, sProcName, lContinue, lCodeLine, lStartCol, -1&

1550                           End If
                            End If
                         End If
1560                  End If
                   End If
                End If
1570            lStartCol = lEndColumn + 1
1580            lEndColumn = -1
1590            lEndLine = .CountOfDeclarationLines + 1
             Loop

            ' Locate the first line of procedure code
1600         lCodeLine = .CountOfDeclarationLines + 1
1610         lEndLine = -1

1620         lStartCol = 1
1630         lEndColumn = -1

1640         Do While .Find(sMembName, lCodeLine, lStartCol, lEndLine, lEndColumn, True, True)
1650            If lCodeLine <> lProcLine Then
1660               sCodeLine = .Lines(lCodeLine, 1)

1670               If IsCode(sCodeLine, lStartCol, lEndColumn) Then

                      ' Grab member name (and procedure kind) of procedure
1680                   sProcName = .ProcOfLine(lCodeLine, eProcKind)

1690                    If sProcName <> sMembName Then

1700                       lLineStart = .ProcBodyLine(sProcName, eProcKind)
1710                       If lLineStart <> lCodeLine Then
1720                          ' The VBE's CodeModule Find function accepts an underscore as a delimiter
                              If Not IsWholeWord(sCodeLine, lStartCol, lEndColumn) Then GoTo DoNext
1730                       Else ' Member name exists within declaration line of procedure
1740                          i = InStr(sProcName, sMembName)
                              If i = 1 Then ' Procedure name starts with member name
                                  If InStr(sProcName, sMembName & "_") = 1 Then
                                      ' "sMembName_" is a valid hit
                                      AddNewRef .Parent.Name & "." & sProcName, sProcName, lLineStart, lCodeLine, lStartCol, eProcKind
                                  End If
                                  GoTo DoNext ' Naming conflict with procedure name
                              ' If procedure name doesn't start with member name
1750                          ElseIf i > 1 And mBlockedName <> .Parent.Name Then
                                  i = InStr(1, sProcName, "_")
                                  If i <> 0 Then
                                     ' Possibly "sCompName_MembName" ' Implements
                                     If sProcName = sCompName & "_" & sMembName Then ' Bug fix 2.26
                                        'sInstName = sCompName
                                        'fExempt = -1&
                                        ' "sCompName_MembName" is a valid hit
                                        AddNewRef .Parent.Name & "." & sProcName, sProcName, lLineStart, lCodeLine, lStartCol, eProcKind
                                     End If
                                     ' WithEvents Memb As sCompName?
                                     sCaller = Left$(sProcName, i - 1&)
                                     Set oMember = Nothing
                                     On Error Resume Next
                                     Set oMember = .Members(sCaller)
                                     On Error GoTo EndWith
                                     If Not oMember Is Nothing Then  ' Just got to be WithEvents var
                                        i = 1&: j = 1&               ' But is it of type sCompName?
                                        k = .CountOfDeclarationLines ' Can't trust CodeLocation prop
                                        Do While .Find("WithEvents " & sCaller, i, j, k, -1&, , True)
                                            k = j + Len("WithEvents " & sCaller)
                                            sCodeLine = .Lines(i, 1&)       ' Get declaration line
                                            If IsCode(sCodeLine, j, k) Then ' Only if its code
                                               If InStr(1, sCodeLine, "As " & sCompName) Then
                                                  ' "Memb_MembName" is a valid hit
                                                  AddNewRef .Parent.Name & "." & sProcName, sProcName, lLineStart, lCodeLine, lStartCol, eProcKind
                                               End If
                                            End If
                                            j = k ' Skip comment header?
                                            k = .CountOfDeclarationLines
                                        Loop
                                     End If
                                  End If
                              Else
                                  If InStr(sCodeLine, " As " & sMembName) > 0 Then ' Bug fix 2.3 June 16, 2016
                                      ' "As sMembName" is a valid hit
                                      AddNewRef .Parent.Name & "." & sProcName, sProcName, lLineStart, lCodeLine, lStartCol, eProcKind
                                  End If
                                  GoTo DoNext ' Naming conflict with param name
                              End If
                           End If

                           fMustQualify = 0
1760                       If LenB(sCompName) Then
1770                          lLineCount = .ProcCountLines(sProcName, eProcKind)
1780                          lInstCnt = GetInstanceName(oCodeMod, sCompName, lLineStart, lLineStart + lLineCount, sInstName)

                              If lInstCnt = 0 Then
                                 ' If we have failed to find any class instantiation we
                                 ' may still succeed if this is a member of an MDI form
                                 ' calling a procedure in an MDI child
1783                              If oCodeMod.Parent.Type = vbext_ct_VBMDIForm Then
1786                                  If oVBE.ActiveVBProject.VBComponents(sCompName).Type = vbext_ct_VBForm Then
                                          If lStartCol - 11 > 0 Then ' Len("ActiveForm.")
                                              If Mid$(sCodeLine, lStartCol - 11, 11) = "ActiveForm." Then
                                                  sInstName(0) = "ActiveForm" ' Bug fix March 20, 2014
                                                  lInstCnt = 1
                                              End If
                                          End If
                                      End If
                                  End If

                                  If lInstCnt = 0 Then
                                    ' Or perhaps a user control
1788                                 If oVBE.ActiveVBProject.VBComponents(sCompName).Type = vbext_ct_UserControl Then
                                       ' oVBE.ActiveVBProject.VBComponents(sCompName).Name
                                        If oVBE.ActiveVBProject.Type <> vbext_pt_ActiveXControl Then
                                          If lStartCol - 14 > 0 Then ' Len("ActiveControl.")
                                             If Mid$(sCodeLine, lStartCol - 14, 14) = "ActiveControl." Then
                                                 sInstName(0) = "ActiveControl" ' Bug fix March 21, 2014
                                                 lInstCnt = 1
                                             End If
                                          End If
                                        End If
                                     End If
                                  End If

                                 ' Default to Class name
                                  If lInstCnt = 0 Then
                                     sInstName(0) = sCompName
                                     lInstCnt = 1
                                  End If
                              End If

1790                          If oVBE.ActiveVBProject.VBComponents(sCompName).Type <> vbext_ct_StdModule Then
1800                             If Not fExempt Then fMustQualify = -1
                              End If
                           Else 'LenB(sCompName) = 0
                              sInstName(0) = .Parent.Name
                              lInstCnt = 1
                           End If

1810                       i = 0&
                           Do
                              If IsValid(oCodeMod, lCodeLine, lStartCol, sInstName(i), lLineStart, fMustQualify, sMembName, lProcLine) Then

1840                             AddNewRef .Parent.Name & "." & sProcName, sProcName, lLineStart, lCodeLine, lStartCol, eProcKind

1910                          End If
                              i = i + 1&
                           Loop Until i = lInstCnt
1920                    End If
1930                 End If
                  End If
DoNext:
1940              lStartCol = lEndColumn + 1
1950              lEndColumn = -1
1960              lEndLine = -1
1970         Loop
EndWith:
1980       End With
       If Err Then LogError "modVBE.FindCallers", sCaller
End Sub

Private Function GetDeclareName(ByVal oCodeMod As CodeModule, ByRef lCodeLine As Long, ByVal lStartCol As Long, Optional vbeScope As vbext_Scope, Optional eKind As vbext_ProcKind) As String
         Dim i As Long, j As Long, k As Long
         Dim sCodeLine As String
         Dim sTempLine As String
         Dim sMembName As String
         Dim sBuffer As String
         
         Const sDECS As String = " Public Private Declare Const Dim Global "

2000     On Error GoTo ErrHandler

2010     sCodeLine = oCodeMod.Lines(lCodeLine, 1)
2020     If LenB(sCodeLine) = 0 Then Exit Function

'         Do While lCodeLine > 1
'            sCodeLine = oCodeMod.Lines(lCodeLine, 1)
'            If LenB(sCodeLine) Then Exit Do
'            lCodeLine = lCodeLine - 1
'         Loop
'         If lCodeLine = 0 Then Exit Function

2030     If IsCode(sCodeLine, lStartCol, lStartCol) Then

           ' Try to match the member name with the selection
2040        i = lStartCol
2050        j = lStartCol
2060        k = Len(sCodeLine)

2070        Do While i > 1  ' Step back to a delimiter
2080           If IsDelimI(MidI(sCodeLine, i - 1)) Then Exit Do
2090           i = i - 1
2100        Loop
2110        Do Until j > k ' Step forward to a delimiter
2120           If IsDelimI(MidI(sCodeLine, j)) Then Exit Do
2130           j = j + 1
2140        Loop

2150        sMembName = Mid$(sCodeLine, i, j - i)
           'sMembName = Trim$(Mid$(sCodeLine, i, j - i))
2160     End If

2170     sCodeLine = LTrim$(sCodeLine)
2180    'If AscW(sCodeLine) = 35 Then Exit Function '# Line
2190     If AscW(sCodeLine) = 39 Then Exit Function 'Comment Line

2200     If LenB(sMembName) Then
2210       On Error Resume Next ' Is it a member name
2220        sBuffer = oCodeMod.Members(sMembName).Name

2230        If LenB(sBuffer) Then ' If so, is it a code member
2240          vbeScope = oCodeMod.Members(sBuffer).Scope

2250          If vbeScope <> 0 Then ' If so, we have it
2260            GetDeclareName = sBuffer
2270            Exit Function
2280          End If
2290        End If
2300       On Error GoTo ErrHandler
2310     End If

2320     If lCodeLine <= oCodeMod.CountOfDeclarationLines Then

          ' Check for line continuation
           Do While lCodeLine > 1 ' .ProcOfLine kinda thing
             'If Right$(oCodeMod.Lines(lCodeLine - 1, 1), 1) = "_" Then
              If RightI(oCodeMod.Lines(lCodeLine - 1, 1), 1) = 95 Then
                 lCodeLine = lCodeLine - 1 ' Line continuation
              Else
                 Exit Do
              End If
           Loop

2330       sCodeLine = LTrim$(oCodeMod.Lines(lCodeLine, 1))
           sTempLine = " " & sCodeLine

          ' Enums and Types are not included in the
          ' members collection so try them first
2340       If InStr(sTempLine, " Enum ") Then
2350          j = InStr(sTempLine, " Enum ") + 6
2360          k = InStr(j, sTempLine, " ")
2370          If k = 0 Then
2380             GetDeclareName = Mid$(sTempLine, j)
2390          Else
2400             GetDeclareName = Mid$(sTempLine, j, k - j)
2410          End If
2420          j = InStr(sCodeLine, " ")
2430          Select Case Left$(sCodeLine, j - 1)
                Case "Public"
2440              vbeScope = vbext_Public
2460              fExempt = -1
               'Case "Private"
                 'vbeScope = vbext_Private
2480            Case Else
2490              vbeScope = vbext_Private
2500          End Select
2510          Exit Function

2520       ElseIf InStr(sTempLine, " Type ") Then
2530          j = InStr(sTempLine, " Type ") + 6
2540          k = InStr(j, sTempLine, " ")
2550          If k = 0 Then
2560             GetDeclareName = Mid$(sTempLine, j)
2570          Else
2580             GetDeclareName = Mid$(sTempLine, j, k - j)
2590          End If
2600          j = InStr(sCodeLine, " ")
2610          Select Case Left$(sCodeLine, j - 1)
                Case "Private"
2630              vbeScope = vbext_Private
               'Case "Public"
                 'vbeScope = vbext_Public
2650            Case Else
2660              vbeScope = vbext_Public
2670          End Select
2680          Exit Function

          ' An Implements object is also overlooked
2690       ElseIf Left$(sTempLine, 12) = " Implements " Then
2700          j = 13
2710          k = InStr(j, sTempLine, " ")
2720          If k = 0 Then
2725             GetDeclareName = Mid$(sTempLine, j)
2730          Else
2735             GetDeclareName = Mid$(sTempLine, j, k - j)
2740          End If
2745          vbeScope = vbext_Private
2750          Exit Function

          ' Also try a raised Event
2755       ElseIf Left$(sTempLine, 7) = " Event " Then
2760          j = 8
2765          k = InStr(j, sTempLine, "(")
2770          If Not (k = 0) Then
2775             GetDeclareName = Mid$(sTempLine, j, k - j)
2780             vbeScope = vbext_Private
2785             Exit Function
2790          End If

2795       End If

2800       j = InStr(sCodeLine, " ") ' Member of a Type or Enum?
2810       If j = 0 Then j = Len(sCodeLine) + 1
2820       sMembName = Left$(sCodeLine, j - 1)

2830      'sDECS = " Public Private Declare Const Dim Global "
2840       If InStr(sDECS, " " & sMembName & " ") = 0 Then

2850          i = -1
2855          Do While (lCodeLine > 1) And i ' .ProcOfLine kinda thing
2860             lCodeLine = lCodeLine - 1
2865             sTempLine = " " & LTrim$(oCodeMod.Lines(lCodeLine, 1))

2870             j = InStr(2, sTempLine, " ")
2875             If j = 0 Then j = Len(sTempLine) + 1
2880             sMembName = LTrim$(Left$(sTempLine, j - 1))

2885             If InStr(sDECS, " " & sMembName & " ") <> 0 Then i = 0

2890             If InStr(sTempLine, " Type ") Then
2900                 j = InStr(sTempLine, " Type ") + 6
2910                 k = InStr(j, sTempLine, " ")
2920                 If k = 0 Then
2930                    GetDeclareName = Mid$(sTempLine, j)
2940                 Else
2950                    GetDeclareName = Mid$(sTempLine, j, k - j)
2960                 End If
2970                 sCodeLine = LTrim$(sTempLine)
2980                 j = InStr(sCodeLine, " ")
2990                 Select Case Left$(sCodeLine, j - 1)
                        Case "Public"
3010                       vbeScope = vbext_Public
                       'Case "Private"
                         'vbeScope = vbext_Private
3030                    Case Else
3040                      vbeScope = vbext_Private
3050                End Select
3060                Exit Function

3070             ElseIf InStr(sTempLine, " Enum ") Then
3080                j = InStr(sTempLine, " Enum ") + 6
3090                k = InStr(j, sTempLine, " ")
3100                If k = 0 Then
3110                   GetDeclareName = Mid$(sTempLine, j)
3120                Else
3130                   GetDeclareName = Mid$(sTempLine, j, k - j)
3140                End If
3150                sCodeLine = LTrim$(sTempLine)
3160                j = InStr(sCodeLine, " ")
3170                Select Case Left$(sCodeLine, j - 1)
                       Case "Public"
3180                      vbeScope = vbext_Public
3200                      fExempt = -1
                      'Case "Private"
                         'vbeScope = vbext_Private
3220                   Case Else
3230                      vbeScope = vbext_Private
3240                End Select
3250                Exit Function
3260             End If

3270         Loop
3280      End If 'Not a declaration keyword?

3290    End If 'If lCodeLine <= oCodeMod.CountOfDeclarationLines

       ' User did not click on top of a member so loop through
       ' the members to try to find one within this code line
3300    For i = 1 To oCodeMod.Members.Count

3310      j = InStr(sCodeLine, " " & oCodeMod.Members(i).Name)
3320      If j > 0 Then

3330         sMembName = oCodeMod.Members(i).Name
3340         k = j + 1 + Len(sMembName)
3350         Select Case oCodeMod.Members(i).Type

                Case vbext_mt_Variable
3360               If MidI(sCodeLine, k) = 32 Or MidI(sCodeLine, k) = 40 Then ' " " Or "("
3370                  If InStr(sCodeLine, " ") = j Then
3380                     GetDeclareName = sMembName
3390                     vbeScope = oCodeMod.Members(i).Scope
3400                     Exit For
3410                  ElseIf InStr(sCodeLine, " WithEvents ") = j - 11 Then
3420                     GetDeclareName = sMembName
3430                     vbeScope = oCodeMod.Members(i).Scope
3440                     Exit For
3450                  End If
3460               ElseIf k > Len(sCodeLine) Then
3470                  GetDeclareName = sMembName
3480                  vbeScope = oCodeMod.Members(i).Scope
3490                  Exit For
3500               End If

3510            Case vbext_mt_Const
3520               If InStr(sCodeLine, "Const") = j - 5 Then
3530                  If MidI(sCodeLine, k) = 32 Then ' " "
3540                     GetDeclareName = sMembName
3550                     vbeScope = oCodeMod.Members(i).Scope
3560                     Exit For
3570                  End If
3580               End If

3590            Case vbext_mt_Event  ' Raised Events
3600               If InStr(sCodeLine, "Event") = j - 5 Then
3610                  GetDeclareName = sMembName
3620                  vbeScope = oCodeMod.Members(i).Scope
3630                  Exit For
3640               End If

3650            Case vbext_mt_Method      ' Includes ALL procedures except properties,
3660               If Mid$(sCodeLine, k, 5) = " Lib " Then ' including API Declares
3670                  GetDeclareName = sMembName
3680                  vbeScope = oCodeMod.Members(i).Scope
                      eKind = vbext_pk_Proc
3690                  Exit For
3700               End If ' Non-API Declares handled by ProcOfLine in RefreshMemberReferences

               'Case vbext_mt_Property ' Handled by ProcOfLine in RefreshMemberReferences
3710         End Select
           
3720      End If
3730    Next

ErrHandler:
   If Err Then LogError "modVBE.GetDeclareName", sMembName
End Function

Private Function IsWholeWord(sLine As String, ByVal lStartCol As Long, ByVal lEndColumn As Long) As Boolean
  On Error GoTo ErrHandler
    Dim fDelim As Long ' The VBE's CodeModule Find function accepts an underscore as a delimiter
    If (lStartCol > 1) Then
        fDelim = MidI(sLine, lStartCol - 1) <> 95 '<> "_"
    Else
        fDelim = -1
    End If
    If fDelim Then
       If lEndColumn <= Len(sLine) Then
          fDelim = MidI(sLine, lEndColumn) <> 95 '<> "_"
       Else
          fDelim = -1
       End If
    End If
    IsWholeWord = fDelim
ErrHandler:
  If Err Then LogError "modVBE.IsWholeWord", sLine
End Function

Private Function IsCode(sLine As String, ByVal lStartCol As Long, ByVal lEndColumn As Long) As Boolean
    ' Adapted from Project References addin by ':) Ulli
    Const Comment As String = "'"
    Const Quote As String = """"
    Dim i As Long, j As Long, k As Long
   On Error GoTo ErrHandler
   ' See if word is in a comment
    If Left$(LTrim$(sLine), 4) = "Rem " Then Exit Function
    If Left$(LTrim$(sLine), 1) = Comment Then Exit Function

    i = InStr(1, sLine, Comment)
    k = InStr(1, sLine, Quote)

    Do Until i = 0 Or k = 0

       j = InStr(k + 1, sLine, Quote)
       If j = 0 Then Exit Function ' Error out!

       If k < i And j > i Then ' If comment is in a string literal
            i = InStr(i + 1, sLine, Comment)
       End If
       k = InStr(j + 1, sLine, Quote)
    Loop

    If i = 0 Then i = Len(sLine)
    If lStartCol > i Then Exit Function

    k = InStr(1, sLine, Quote)
    Do Until k = 0 Or k > i

       j = InStr(k + 1, sLine, Quote)
       If j = 0 Then Exit Function ' Error out!

       If k < lStartCol And j >= lEndColumn Then ' Bug fix Dec 19, 2010
             Exit Function ' The word is in a string literal
       End If
       k = InStr(j + 1, sLine, Quote)
    Loop
    IsCode = True

ErrHandler:
 If Err Then LogError "modVBE.IsCode", sLine
End Function

Private Function IsValid(ByVal oCodeMod As CodeModule, ByVal lCodeLine As Long, ByVal lStartCol As Long, sCompName As String, ByVal lLineStart As Long, ByVal fMustQualify As Long, sMembName As String, ByVal lProcLine As Long) As Boolean
    Dim sCodeLine As String
    Dim fWithBlock As Long
    Dim i As Long, j As Long
    
   On Error GoTo ErrHandler

     If lStartCol > 1 Then
         sCodeLine = oCodeMod.Lines(lCodeLine, 1)
         If MidI(sCodeLine, lStartCol - 1) = 46 Then ' "."
             If lStartCol > 2 Then
                'If Not IsDelim(Mid$(sCodeLine, lStartCol - 2, 1)) Then
                 If Not IsDelimI(MidI(sCodeLine, lStartCol - 2)) Then
                     If lStartCol > 3 Then
                        If Mid$(sCodeLine, lStartCol - 3, 2) = "Me" Then
                           If lStartCol > 4 Then ' Bug fix Dec 11, 2010
                              If IsDelimI(MidI(sCodeLine, lStartCol - 4)) Then
                                 IsValid = Not fMustQualify
                              End If
                           Else ' If Me is current class, fMustQualify is false
                              IsValid = Not fMustQualify
                           End If
                        End If
                     End If ' Qualified, but is it our component?
                     i = Len(sCompName)
                     If lStartCol > i + 1 Then ' "comp.proc"
                        IsValid = (Mid$(sCodeLine, lStartCol - i - 1, i) = sCompName)
                     End If
                 ElseIf MidI(sCodeLine, lStartCol - 2) = 41 Then ' ")"  Bug fix March 20, 2014
                    ' We may have an array of this class object
                     i = InStr(1, sCompName, "(")
                     If i Then
                        j = lStartCol - 3
                        Do While j
                            If MidI(sCodeLine, j) = 40 Then Exit Do ' "("
                            j = j - 1
                        Loop
                        If Not j < i Then
                            i = i - 1
                            IsValid = (Mid$(sCodeLine, j - i, i) = Left$(sCompName, i))
                        End If
                     End If
                 Else ' We have a With block " .proc" | "(.proc" etc
                     fWithBlock = -1
                 End If
             Else ' We have a With block ".proc"
                 fWithBlock = -1
             End If

         Else ' " memb" | "(memb" etc
            IsValid = Not fMustQualify
            If IsValid And (lProcLine = 0) Then ' lProcLine is zero if not current comp
              ' Check for duplicate name with narrower scope
               For i = 1 To oCodeMod.Members.Count
                  If oCodeMod.Members(i).Name = sMembName Then
                     IsValid = False
                     Exit For
                  End If
               Next i
            End If
         End If

         If fWithBlock Then
            Do While (lCodeLine > lLineStart) ' .ProcOfLine kinda thing
               lCodeLine = lCodeLine - 1
               sCodeLine = " " & oCodeMod.Lines(lCodeLine, 1)
               i = InStr(sCodeLine, " With ")
               If i Then
                   i = i + 6
                   j = InStr(i, sCodeLine, " ")
                   If j = 0 Then j = Len(sCodeLine) + 1

                   IsValid = (Mid$(sCodeLine, i, j - i) = sCompName)
                   Exit Do
               End If
            Loop
         End If
     Else ' "proc"
         IsValid = Not fMustQualify
         If IsValid And (lProcLine = 0) Then
           ' Check for duplicate name with narrower scope
            For i = 1 To oCodeMod.Members.Count
               If oCodeMod.Members(i).Name = sMembName Then
                  IsValid = False
                  Exit For
               End If
            Next i
         End If
     End If

ErrHandler:
 If Err Then LogError "modVBE.IsValid", sCodeLine
End Function

' ¤¤ IsDelim ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  This function checks if the character passed is a common word
'  delimiter, and then returns True or False accordingly.
'
'  By default, any non-alphabetic character (except for an
'  underscore) is considered a word delimiter, including numbers.
'
'  By default, an underscore is treated as part of a whole word,
'  and so is not considered a word delimiter.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Public Function IsDelimI(ByVal iAscW As Long) As Boolean ' ©Rd
    Select Case iAscW
        ' Uppercase, Underscore, Lowercase chars not delimiters
        Case 65 To 90, 95, 97 To 122: IsDelimI = False

        'Case 39, 146: IsDelimI = False  ' Apostrophes not delimiters
        Case 48 To 57: IsDelimI = False ' Numeric chars not delimiters

        Case Else: IsDelimI = True ' Any other character is delimiter
    End Select
End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Sub strShellIndexed(sA() As String, idxA() As Long)
   ' Thanks heaps to LukeH for this great algorithm :)
   Dim i As Long, j As Long, k As Long
   Dim lpStr As Long, lp As Long, n As Long
   Dim Idx As Long, s As String
   Const n1 As Long = 1, n4 As Long = 4
   lp = VarPtr(sA(n1)) - n4 '(- lb * 4)                 ' Cache the pointer to the array
   lpStr = VarPtr(s)                                    ' Cache the pointer to the string variable
   k = nCallers       ' -----=====================----- ' Get the distance from lowerbound to upperbound
   Do: j = j + j + j + n1: Loop Until j > k             ' Find the initial value for distance
   Do: j = j \ 3&                                       ' Reduce distance by two thirds
      For i = n1 + j To nCallers                        ' Loop through each position in our current section
         CopyMemByV lpStr, lp + idxA(i) * n4, n4        ' Put the current value in the string buffer (using its pointer)
         Idx = idxA(i)                                  ' Put the current index in the index buffer
         n = i - j                                      ' Set the pointer to the value below
         If StrComp(sA(idxA(n)), s, n1) = n1 Then       ' Compare the current value with the immediately previous value
            k = i                                       ' If the wrong order then set our temp pointer to the current index
            Do: idxA(k) = idxA(n)                       ' Copy the lower index to the current index
               k = n: n = n - j '-vb2themax-            ' Adjust the pointers to compare down a level
               If n < n1 Then Exit Do                   ' Make sure we're in-bounds or exit the loop
            Loop While StrComp(sA(idxA(n)), s, n1) = n1 ' Keep going as long as current value needs to move down
            idxA(k) = Idx                               ' Put the buffered index back in the correct position
      End If: Next                                      ' Increment the inner for loop
   Loop Until j = n1  ' -----=====================----- ' Drop out when we're done
   CopyMemByR ByVal lpStr, 0&, n4                       ' De-reference our pointer to variable s
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Twice as fast as AscW and Mid$ when compiled.
'        iChr = AscW(Mid$(sStr, lPos, 1))
'        iChr = MidI(sStr, lPos)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get MidI(sStr As String, ByVal lpos As Long) As Integer
    CopyMemByR MidI, ByVal StrPtr(sStr) + lpos + lpos - 2&, 2&
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Mid$(sStr, lPos, 1) = Chr$(iChr)
'        MidI(sStr, lPos) = iChr
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let MidI(sStr As String, ByVal lpos As Long, ByVal iChrW As Integer)
    CopyMemByR ByVal StrPtr(sStr) + lpos + lpos - 2&, iChrW, 2&
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       iChr = AscW(Right$(sStr, lPos, 1))
'       iChr = RightI(sStr, lPos)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get RightI(sStr As String, ByVal lRightPos As Long) As Integer
    CopyMemByR RightI, ByVal StrPtr(sStr) + LenB(sStr) - lRightPos - lRightPos, 2&
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Right$(sStr, lPos, 1) = Chr$(iChr)
'      RightI(sStr, lPos) = iChr
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let RightI(sStr As String, ByVal lRightPos As Long, ByVal iChrW As Integer)
    CopyMemByR ByVal StrPtr(sStr) + LenB(sStr) - lRightPos - lRightPos, iChrW, 2&
End Property

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

