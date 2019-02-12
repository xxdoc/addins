Imports System.Runtime.CompilerServices
Imports Microsoft.Win32
Imports System.Drawing
'---------------------------------------------------------------------
'
'Various extension methods used in BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------

Module modExtensions

    Public Enum LineModsEnum
        Added
        Removed
        Changed
    End Enum


    Private _AllBookmarks As Dictionary(Of VBIDE.VBProject, Bookmarks) = New Dictionary(Of VBIDE.VBProject, Bookmarks)
    <Extension()> _
    Public Function Bookmarks(VBProject As VBIDE.VBProject) As Bookmarks
        If _AllBookmarks.ContainsKey(VBProject) Then
            Return _AllBookmarks(VBProject)
        Else
            Dim bm = New Bookmarks(VBProject)
            _AllBookmarks.Add(VBProject, bm)
            Return bm
        End If
    End Function


    Private _AllBreakPoints As Dictionary(Of VBIDE.VBProject, Breakpoints) = New Dictionary(Of VBIDE.VBProject, Breakpoints)
    <Extension()> _
    Public Function BreakPoints(VBProject As VBIDE.VBProject) As Breakpoints
        If _AllBreakPoints.ContainsKey(VBProject) Then
            Return _AllBreakPoints(VBProject)
        Else
            Dim bp = New Breakpoints(VBProject)
            _AllBreakPoints.Add(VBProject, bp)
            Return bp
        End If
    End Function


    ''' <summary>
    ''' When the project is unloaded, release all the marks we're tracking
    ''' </summary>
    ''' <param name="VBProject"></param>
    ''' <remarks></remarks>
    <Extension()> _
    Public Sub Release(VBProject As VBIDE.VBProject)
        If _AllBreakPoints.ContainsKey(VBProject) Then
            _AllBreakPoints.Remove(VBProject)
        End If
        If _AllBookmarks.ContainsKey(VBProject) Then
            _AllBookmarks.Remove(VBProject)
        End If
    End Sub


    <Extension()> _
    Public Sub ClearBookmarks(VBProjects As VBIDE.VBProjects)
        For Each p As VBIDE.VBProject In VBProjects
            p.Bookmarks.Clear()
        Next
    End Sub


    <Extension()> _
    Public Sub ClearBreakPoints(VBProjects As VBIDE.VBProjects)
        For Each p As VBIDE.VBProject In VBProjects
            p.BreakPoints.Clear()
        Next
    End Sub



    <Extension()> _
    Public Sub SaveBookmarks(VBProjects As VBIDE.VBProjects)
        For Each p As VBIDE.VBProject In VBProjects
            p.Bookmarks.Save()
        Next
    End Sub


    <Extension()> _
    Public Sub SaveBreakPoints(VBProjects As VBIDE.VBProjects)
        For Each p As VBIDE.VBProject In VBProjects
            p.BreakPoints.Save()
        Next
    End Sub


    <Extension()> _
    Public Sub LineChange(CodeModule As VBIDE.CodeModule, ByVal LineMods As LineModsEnum, ByVal CurLineNumber As Integer, ByVal NumLinesToChange As Integer)
        Dim Bookmarks = CodeModule.Parent.Collection.Parent.Bookmarks
        Dim BreakPoints = CodeModule.Parent.Collection.Parent.BreakPoints
        Dim ModName = CodeModule.Parent.Name
        Static BPsToCheck As List(Of Breakpoint) = New List(Of Breakpoint)

        If Bookmarks.Count = 0 And BreakPoints.Count = 0 Then Exit Sub

        '====================================================
        ' BOOKMARK HANDLING
        '====================================================
        Dim BMRemoved As HashSet(Of Bookmark) = New HashSet(Of Bookmark)
        For Each bm In Bookmarks
            If StrComp(bm.ModuleName, ModName, CompareMethod.Text) = 0 Then
                If LineMods = LineModsEnum.Added Then
                    If bm.LineNumber > CurLineNumber - NumLinesToChange Then
                        Logger.Log("Bookmark was " & bm.ToString)
                        bm.LineNumber = bm.LineNumber + NumLinesToChange
                        Logger.Log("Bookmark is now " & bm.ToString)
                    End If
                ElseIf LineMods = LineModsEnum.Removed Then
                    If bm.LineNumber >= CurLineNumber + NumLinesToChange Then
                        Logger.Log("Bookmark was " & bm.ToString)
                        bm.LineNumber = bm.LineNumber - NumLinesToChange
                        Logger.Log("Bookmark is now " & bm.ToString)
                    ElseIf bm.LineNumber >= CurLineNumber Then
                        BMRemoved.Add(bm)
                        Logger.Log("Bookmark Line was removed " & bm.ToString)
                    End If
                Else
                    '---- the lines were changed, but in the case of bookmarks
                    '     VB will keep the bookmarks even if you comment out a line
                End If
            End If
        Next
        Bookmarks.RemoveAll(Function(bm) BMRemoved.Contains(bm))
        '====================================================
        '====================================================


        '====================================================
        ' BREAKPOINT HANDLING
        '====================================================
        Dim BPRemoved As HashSet(Of Breakpoint) = New HashSet(Of Breakpoint)
        For Each bp In BreakPoints
            If StrComp(bp.ModuleName, ModName, CompareMethod.Text) = 0 Then
                If LineMods = LineModsEnum.Added Then
                    Logger.Log(String.Format("Lines Added, bp.linenum={0}  CurLine={1}  LinestoChange={2}", bp.LineNumber, CurLineNumber, NumLinesToChange))
                    If bp.LineNumber > CurLineNumber - NumLinesToChange Then
                        Logger.Log("Breakpoint was " & bp.ToString)
                        bp.LineNumber = bp.LineNumber + NumLinesToChange
                        Logger.Log("Breakpoint is now " & bp.ToString)
                    Else
                        Logger.Log("Breakpoint Skipped because line is before changed lines")
                    End If
                ElseIf LineMods = LineModsEnum.Removed Then
                    Logger.Log(String.Format("Lines Removed, bp.linenum={0}  CurLine={1}  LinestoChange={2}", bp.LineNumber, CurLineNumber, NumLinesToChange))
                    If bp.LineNumber >= CurLineNumber + NumLinesToChange Then
                        Logger.Log("Breakpoint was " & bp.ToString)
                        bp.LineNumber = bp.LineNumber - NumLinesToChange
                        Logger.Log("Breakpoint is now " & bp.ToString)
                    ElseIf bp.LineNumber >= CurLineNumber Then
                        BPRemoved.Add(bp)
                        Logger.Log("Breakpoint Line was removed " & bp.ToString)
                    Else
                        Logger.Log("Breakpoint Skipped because line is after or same as changed lines")
                    End If
                Else
                    '---- the lines where changed
                    '     in the case of breakpoints, VB MIGHT remove any breakpoints that 
                    '     were set on the line, the prob is it won't remove them till you move off
                    '     the line (you can remove the comment char, and the breakpoint will remain)
                    '
                    '     So, i add this BP to a list, and check it later
                    '     once the user is no longer editing the line where this breakpoint is defined
                    '
                    '     Yeah, this approach is chock full of possibilities for holes, but what can you do?
                    If bp.LineNumber >= CurLineNumber And bp.LineNumber < (CurLineNumber + NumLinesToChange) Then
                        If CodeModule.IsBreakpointCommentedNow(bp) Then
                            If Not BPsToCheck.Contains(bp) Then BPsToCheck.Add(bp)
                        End If
                    End If
                End If
            End If
        Next

        '---- if there are any Breakpoints in the check list
        '     see if they're still commented now, as long as we're no longer 
        '     pointing at their range
        '
        '     Not great because this represents quite a bit of work to do on each timer
        '     tick, but in testing, it doesn't seem to register in Procmon, and it's relatively
        '     straightforward, so it'll do for now.
        Dim RemoveFromBPsToCheck As List(Of Breakpoint) = Nothing
        For Each bp In BPsToCheck
            If StrComp(bp.ModuleName, ModName, CompareMethod.Text) = 0 Then
                If bp.LineNumber >= CurLineNumber And bp.LineNumber < (CurLineNumber + NumLinesToChange) Then
                    '---- we're still looking at the same range, so skip this one
                Else
                    '---- no longer looking at the line that had the breakpoint
                    '     so if it's still commented, it can't be a breakpoint anymore
                    If CodeModule.IsBreakpointCommentedNow(bp) Then
                        Logger.Log("Breakpoint was removed because it's been commented out: " & bp.ToString)
                        BPRemoved.Add(bp)
                    End If
                    If RemoveFromBPsToCheck Is Nothing Then RemoveFromBPsToCheck = New List(Of Breakpoint)
                    RemoveFromBPsToCheck.Add(bp)
                End If
            End If
        Next
        '---- once we've checked a breakpoint above, remove it from the check list
        If RemoveFromBPsToCheck IsNot Nothing Then
            BPsToCheck.RemoveAll(Function(bp) RemoveFromBPsToCheck.Contains(bp))
        End If

        '---- and finally, remove any breakpoints that have been flagged as cleared
        BreakPoints.RemoveAll(Function(bp) BPRemoved.Contains(bp))
    End Sub


    <Extension()> _
    Public Function IsBreakpointCommentedNow(CodeModule As VBIDE.CodeModule, bp As Breakpoint) As Boolean
        Dim l = CodeModule.Lines(bp.LineNumber, 1)
        If Left(Trim(l), 1) = "'" Or Left(Trim(l), 4).ToLower = "rem " Then
            Return True
        End If
        Return False
    End Function


    <Extension()> _
    Public Function Find(VBComponents As VBIDE.VBComponents, Name As String) As VBIDE.VBComponent
        Return (From comp As VBIDE.VBComponent In VBComponents Where StrComp(comp.Name, Name, CompareMethod.Text) = 0 Select comp).FirstOrDefault
    End Function


    <Extension()> _
    Public Function IsCodeWindowVisible(VBComponent As VBIDE.VBComponent) As Boolean
        'Logger.Log("Search visible windows...")
        'For Each win As VBIDE.Window In VBComponent.VBE.Windows
        '    Logger.Log(" Window caption=" & win.Caption)
        'Next
        '---- CodeSmart will change the code window caption (so it fits better in the tab)
        '     so accomodate both versions
        Dim w = (From win As VBIDE.Window In VBComponent.VBE.Windows Where win.Caption Like "* - " & VBComponent.Name & " (Code)" Or win.Caption = VBComponent.Name & " (Code)").FirstOrDefault
        If w IsNot Nothing Then Return True
        Return False
    End Function


    <Extension()> _
    Public Function MainWindow(VBInstance As VBIDE.VBE) As VBIDE.Window
        Return (From win As VBIDE.Window In VBInstance.Windows Where win.Type = VBIDE.vbext_WindowType.vbext_wt_MainWindow).FirstOrDefault
    End Function


    ''' <summary>
    ''' retrieve the font used for code edit windows in VB6
    ''' </summary>
    ''' <param name="VBInstance"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function EditorFont(VBInstance As VBIDE.VBE) As Font
        '---- note, I'm caching the font to reduce overhead for this call
        Dim Face As String
        Static f As Font
        Static LastFace As String
        Static LastHeight As Integer
        Dim h As Integer = 10
        Using regKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\VBA\Microsoft Visual Basic", False)
            Face = regKey.GetValue("FontFace", "Courier New")
            h = regKey.GetValue("FontHeight", 10)
        End Using
        If Face <> LastFace Or h <> LastHeight Then
            LastFace = Face : LastHeight = h
            f = New Font(Face, h)
        End If
        Return f
    End Function
End Module
