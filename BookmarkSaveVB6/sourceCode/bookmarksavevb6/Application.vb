Option Explicit On

Imports VBIDE
Imports Microsoft.Office.Core
Imports System.Windows.Forms.Application
Imports System.Drawing

'---------------------------------------------------------------------
'
'Main Application entry point for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------


''' <summary>
''' Main Application Class for the BookmarkSave Addin
''' </summary>
''' <remarks></remarks>
Public Class Application
    Private Const VBA_WIN As String = "VbaWindow"

    Private _VBInstance As VBIDE.VBE

    Private WithEvents _KeyHook As KeyboardHook = New KeyboardHook

    '---- When restoring bookmarks or breakpoints, this is set to prevent 
    '     the interception code from also kicking in
    Private _bSettingMarks As Boolean

    '---- this is used to monitor edits to the active code window, as the user
    '     edits code, we have to adjust bookmark and breakpoint locations
    Private WithEvents _EditTimer As System.Windows.Forms.Timer = New System.Windows.Forms.Timer

    '---- subclass the MDIChild window of VB to be alerted of when the user clicks
    '     into the margin area of a code window (to set/clear a breakpoint)
    Private WithEvents _MDIChildSubClasser As Subclasser = New Subclasser

    '---- used to track whether or not the last breakpoint set was successful
    '     in some cases, VB can act like it's set a breakpoint, then pop a msgbox saying that
    '     it can't actually set that breakpoint after all
    Private WithEvents _InvalidBreakTimer As System.Windows.Forms.Timer = New System.Windows.Forms.Timer
    Private _LastBP As Breakpoint

    Private WithEvents _MenuHandler As CommandBarEvents          'command bar event handler
    Private _MenuCommandBar As CommandBarControl

    Private mcbMenuBreakPointToggle As CommandBarControl
    Private mcbMenuBookmarkToggle As CommandBarControl
    Private mcbMenuBookmarkNext As CommandBarControl
    Private mcbMenuBookmarkPrev As CommandBarControl

    Private Enum VBCommandIDs
        EnmBreakPointToggle = 51
        EnmBreakPointRemoveALL = 579
        EnmBookmarkToggle = 2525
        EnmBookmarkremoveAll = 2528
        EnmBookmarkNext = 2526
        EnmBookmarkPrev = 2527
    End Enum


    'Sink events for all the commands we need to pay attention to
    'MENU TITLE 'MENU BAR, SUB MENU DEBUG
    Public WithEvents _MenuMainDebug_BreakPointToggle As CommandBarEvents
    Public WithEvents _MenuMainDebug_BreakPointRemoveAll As CommandBarEvents

    '"MENU TITLE 'Debug'
    Public WithEvents _MenuDebug_BreakPointToggle As CommandBarEvents
    Public WithEvents _MenuDebug_BreakPointRemoveAll As CommandBarEvents

    'Menu Title 'Edit'
    Public WithEvents _Edit_BreakPointToggle As CommandBarEvents
    Public WithEvents _Edit_BookmarkToggle As CommandBarEvents
    Public WithEvents _Edit_BookmarkRemoveAll As CommandBarEvents

    'Menu Title 'Toggle'
    Public WithEvents _Toggle_BreakPointToggle As CommandBarEvents
    Public WithEvents _Toggle_BookmarkToggle As CommandBarEvents

    'MenuBookmarks
    Public WithEvents _MenuBookmarks_BookmarkToggle As CommandBarEvents
    Public WithEvents _MenuBookmarks_BookmarkRemoveAll As CommandBarEvents

    'For Project Events
    Public WithEvents _VBProjectsEvents As VBProjectsEvents
    Public WithEvents _VBComponentsEvents As VBComponentsEvents


    '---- these properties just define the hotkeys we'll use for 
    '     setting/naving bookmarks
    Public Property BookMarkToggleHotkey As System.Windows.Forms.Keys = System.Windows.Forms.Keys.K Or System.Windows.Forms.Keys.Control
    Public Property BookMarkNextHotkey As System.Windows.Forms.Keys = System.Windows.Forms.Keys.Right Or System.Windows.Forms.Keys.Alt
    Public Property BookMarkPrevHotkey As System.Windows.Forms.Keys = System.Windows.Forms.Keys.Left Or System.Windows.Forms.Keys.Alt

    Public Property BreakPointNextHotkey As System.Windows.Forms.Keys = System.Windows.Forms.Keys.Down Or System.Windows.Forms.Keys.Alt
    Public Property BreakPointPrevHotkey As System.Windows.Forms.Keys = System.Windows.Forms.Keys.Up Or System.Windows.Forms.Keys.Alt


    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        _EditTimer.Interval = 10
        _EditTimer.Start()
    End Sub


    Public Sub OnConnection(VBInst As VBIDE.VBE, ByVal ConnectMode As VBIDE.vbext_ConnectMode, custom() As Object)
        '---- make the VBE globally available
        _VBInstance = VBInst

        '---- Setup the exception handling stuff
        ExceptionExtensions.Initialize()

        'Do your initializing stuff here
        Try
            _MenuCommandBar = AddToAddInCommandBar("Bookmark Save")

            'sink the event
            _MenuHandler = _VBInstance.Events.CommandBarEvents(_MenuCommandBar)

            'sink all necessary buttons on menus
            Call InitMenuSink("Menu Bar", 30165, "MenubarDebug")
            Call InitMenuSink("Bookmarks")
            Call InitMenuSink("Debug")
            Call InitMenuSink("Edit")
            Call InitMenuSink("Toggle")

        Catch ex As Exception
            ex.Show("Failed to startup addin properly.")
        End Try
    End Sub


    Public Sub OnDisconnect(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Object)
        'Menu Title 'Edit'
        _Edit_BookmarkRemoveAll = Nothing
        _Edit_BookmarkToggle = Nothing
        _Edit_BreakPointToggle = Nothing

        '---- stop the code checking timer
        _EditTimer.Stop()
        _EditTimer = Nothing

        _InvalidBreakTimer.Stop()
        _InvalidBreakTimer = Nothing

        _KeyHook = Nothing

        If _MDIChildSubClasser IsNot Nothing Then _MDIChildSubClasser.Stop()
        _MDIChildSubClasser = Nothing

        'MenuBookmarks
        _MenuBookmarks_BookmarkRemoveAll = Nothing
        _MenuBookmarks_BookmarkToggle = Nothing

        _MenuDebug_BreakPointRemoveAll = Nothing
        _MenuDebug_BreakPointToggle = Nothing

        _MenuHandler = Nothing

        _MenuMainDebug_BreakPointRemoveAll = Nothing
        _MenuMainDebug_BreakPointToggle = Nothing

        'Menu Title 'Toggle'
        _Toggle_BreakPointToggle = Nothing
        _Toggle_BookmarkToggle = Nothing

        _VBComponentsEvents = Nothing
        _VBProjectsEvents = Nothing

        'delete the command bar entry
        _MenuCommandBar.Delete()
        _MenuCommandBar = Nothing

        Logger.Stop()

        _VBInstance = Nothing
    End Sub


    Public Sub OnStartupComplete(custom() As Object)
        '---- after the projects have been loaded sink the VB project events
        _VBProjectsEvents = _VBInstance.Events.VBProjectsEvents
        _VBComponentsEvents = _VBInstance.Events.VBComponentsEvents(_VBInstance.ActiveVBProject)

        '---- auto load bookmarks and breaks
        'If GetSetting(My.Application.Info.Title, "Settings", "StartupBKM", vbChecked) = vbChecked Then
        Call RestoreAllBookmarks()
        'End If

        'If GetSetting(My.Application.Info.Title, "Settings", "StartupBKP", vbChecked) = vbChecked Then
        Call RestoreAllBreakpoints()
        'End If

        '---- we have to subclass the MDIClient window of the Main App window
        '     subclassing the MDIClient appears to mess up CodeSmart, so we'll watch 
        '     the main window instead
        Dim MDIClient = FindWindowLike(_VBInstance.MainWindow.HWnd, "*", "MDIClient")
        If MDIClient.Count = 1 Then
            'Dim hwnd = _VBInstance.MainWindow.HWnd
            Dim hwnd = MDIClient(0)
            _MDIChildSubClasser.PostProcSubclass(hwnd, {Win32.Messages.WM_PARENTNOTIFY, Win32.Messages.WM_SETCURSOR})
        End If
    End Sub


    Public Sub OnAddInsUpdate(custom() As Object)
        '---- nothing neccessary here
    End Sub


    Private Function AddToAddInCommandBar(ByVal Caption As String) As CommandBarControl
        Try
            'see if we can find the Add-Ins menu
            Dim cbMenu = _VBInstance.CommandBars("Add-Ins")
            If cbMenu Is Nothing Then Return Nothing

            'add it to the command bar
            Dim cbMenuCommandBar As CommandBarButton = cbMenu.Controls.Add(1)
            'set the caption
            cbMenuCommandBar.Caption = Caption
            '---- set the icon, but this looks ugly, so I'm not bothering right now
            'My.Computer.Clipboard.SetImage(My.Resources.BookmarkSave_Icon_Image)
            'cbMenuCommandBar.PasteFace()

            Return cbMenuCommandBar
        Catch
            Return Nothing
        End Try
    End Function


    ''' <summary>
    ''' Handles setting up event sinks for all the bookmark/breakpoint related menu items
    ''' </summary>
    ''' <param name="MenuToSubClass"></param>
    ''' <param name="SubMenuTopID"></param>
    ''' <param name="SubIDSubName"></param>
    ''' <remarks></remarks>
    Private Sub InitMenuSink(ByVal MenuToSubClass As String, Optional ByVal SubMenuTopID As Integer = 0, Optional ByVal SubIDSubName As String = "")
        Try
            Dim MenuBar As Object
            Dim MenuControl As CommandBarControl
            Dim SubMenuControl As CommandBarControl
            Dim SubMenuBar As Object

            MenuBar = _VBInstance.CommandBars(MenuToSubClass)

            For Each MenuControl In MenuBar.Controls
                If SubIDSubName = "" Then
                    Call SetEventToSink(MenuControl, MenuToSubClass)
                    MenuControl.Enabled = True
                Else
                    If MenuControl.Id = SubMenuTopID Then
                        For Each SubMenuBar In MenuControl.Controls
                            'Logger.Log(SubMenuBar.Caption & "|| " & SubMenuBar.Id)
                            Call SetEventToSink(SubMenuBar, SubIDSubName)
                            SubMenuBar.Enabled = True
                        Next
                    End If
                    'Logger.Log(MenuControl.Caption & "|| " & MenuControl.Id)
                End If
            Next

            'kill all objects
            MenuBar = Nothing
            MenuControl = Nothing
            SubMenuControl = Nothing

        Catch ex As Exception
            ex.Show("There was an error trying to initialize one or more of the Menu Bar buttons!")
        End Try
    End Sub


    ''' <summary>
    ''' This routine seems a tad overblown, but I'm leaving it as is for now
    ''' Essentially, it just finds and caches referenced to specific VB Commands
    ''' so that I can tell when they've been "executed".
    ''' </summary>
    ''' <param name="CommandBar"></param>
    ''' <param name="MenuToSubClass"></param>
    ''' <remarks></remarks>
    Private Sub SetEventToSink(ByVal CommandBar As CommandBarControl, ByVal MenuToSubClass As String)
        Select Case CommandBar.Id
            Case VBCommandIDs.EnmBreakPointToggle
                Select Case UCase(MenuToSubClass)
                    Case "DEBUG"
                        'used to set breakpoints, only need one reference to this
                        If mcbMenuBreakPointToggle Is Nothing Then
                            mcbMenuBreakPointToggle = CommandBar
                            mcbMenuBreakPointToggle.Enabled = True
                        End If
                        _MenuDebug_BreakPointToggle = _VBInstance.Events.CommandBarEvents(CommandBar)
                    Case "EDIT"
                        'used to set breakpoints, only need one reference to this
                        If mcbMenuBreakPointToggle Is Nothing Then
                            mcbMenuBreakPointToggle = CommandBar
                            mcbMenuBreakPointToggle.Enabled = True
                        End If
                        _Edit_BreakPointToggle = _VBInstance.Events.CommandBarEvents(CommandBar)
                    Case "TOGGLE"
                        'used to set breakpoints, only need one reference to this
                        If mcbMenuBreakPointToggle Is Nothing Then
                            mcbMenuBreakPointToggle = CommandBar
                            mcbMenuBreakPointToggle.Enabled = True
                        End If
                        _Toggle_BreakPointToggle = _VBInstance.Events.CommandBarEvents(CommandBar)
                    Case "MENUBARDEBUG"
                        'MENU TITLE 'MENU BAR, SUB MENU DEBUG
                        _MenuMainDebug_BreakPointToggle = _VBInstance.Events.CommandBarEvents(CommandBar)
                        'used to set breakpoints, only need one reference to this
                        If mcbMenuBreakPointToggle Is Nothing Then
                            mcbMenuBreakPointToggle = CommandBar
                            mcbMenuBreakPointToggle.Enabled = True
                        End If
                End Select

            Case VBCommandIDs.EnmBreakPointRemoveALL
                'Can't get any
                Select Case UCase(MenuToSubClass)
                    Case "MENUBARDEBUG"
                        _MenuMainDebug_BreakPointRemoveAll = _VBInstance.Events.CommandBarEvents(CommandBar)
                End Select

            Case VBCommandIDs.EnmBookmarkToggle
                Select Case UCase(MenuToSubClass)
                    Case "EDIT"
                        'used to set bookmarks, only need one reference to this
                        If mcbMenuBookmarkToggle Is Nothing Then
                            mcbMenuBookmarkToggle = CommandBar
                            mcbMenuBookmarkToggle.Enabled = True
                        End If
                        _Edit_BookmarkToggle = _VBInstance.Events.CommandBarEvents(CommandBar)

                    Case "TOGGLE"
                        'used to set bookmarks, only need one reference to this
                        If mcbMenuBookmarkToggle Is Nothing Then
                            mcbMenuBookmarkToggle = CommandBar
                            mcbMenuBookmarkToggle.Enabled = True
                        End If
                        _Toggle_BookmarkToggle = _VBInstance.Events.CommandBarEvents(CommandBar)
                    Case "BOOKMARKS"
                        'used to set bookmarks, only need one reference to this
                        If mcbMenuBookmarkToggle Is Nothing Then
                            mcbMenuBookmarkToggle = CommandBar
                            mcbMenuBookmarkToggle.Enabled = True
                        End If
                        _MenuBookmarks_BookmarkToggle = _VBInstance.Events.CommandBarEvents(CommandBar)
                End Select

            Case VBCommandIDs.EnmBookmarkNext
                Select Case UCase(MenuToSubClass)
                    Case "EDIT"
                        'used to set bookmarks, only need one reference to this
                        If mcbMenuBookmarkNext Is Nothing Then
                            mcbMenuBookmarkNext = CommandBar
                            mcbMenuBookmarkNext.Enabled = True
                        End If

                    Case "TOGGLE"
                        'used to set bookmarks, only need one reference to this
                        If mcbMenuBookmarkNext Is Nothing Then
                            mcbMenuBookmarkNext = CommandBar
                            mcbMenuBookmarkNext.Enabled = True
                        End If

                    Case "BOOKMARKS"
                        'used to set bookmarks, only need one reference to this
                        If mcbMenuBookmarkNext Is Nothing Then
                            mcbMenuBookmarkNext = CommandBar
                            mcbMenuBookmarkNext.Enabled = True
                        End If
                End Select

            Case VBCommandIDs.EnmBookmarkPrev
                Select Case UCase(MenuToSubClass)
                    Case "EDIT"
                        'used to set bookmarks, only need one reference to this
                        If mcbMenuBookmarkPrev Is Nothing Then
                            mcbMenuBookmarkPrev = CommandBar
                            mcbMenuBookmarkPrev.Enabled = True
                        End If

                    Case "TOGGLE"
                        'used to set bookmarks, only need one reference to this
                        If mcbMenuBookmarkPrev Is Nothing Then
                            mcbMenuBookmarkPrev = CommandBar
                            mcbMenuBookmarkPrev.Enabled = True
                        End If

                    Case "BOOKMARKS"
                        'used to set bookmarks, only need one reference to this
                        If mcbMenuBookmarkPrev Is Nothing Then
                            mcbMenuBookmarkPrev = CommandBar
                            mcbMenuBookmarkPrev.Enabled = True
                        End If
                End Select

            Case VBCommandIDs.EnmBookmarkremoveAll
                Select Case UCase(MenuToSubClass)
                    Case "EDIT"
                        _Edit_BookmarkRemoveAll = _VBInstance.Events.CommandBarEvents(CommandBar)
                    Case "BOOKMARKS"
                        _MenuBookmarks_BookmarkRemoveAll = _VBInstance.Events.CommandBarEvents(CommandBar)
                End Select
        End Select
    End Sub


    Private Sub _Edit_BookmarkRemoveAll_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _Edit_BookmarkRemoveAll.Click
        _VBInstance.VBProjects.ClearBookmarks()
    End Sub


    Private Sub _Edit_BookmarkToggle_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _Edit_BookmarkToggle.Click
        Call BookmarkToggle(CommandBarControl, handled, CancelDefault)
    End Sub


    Private Sub _Edit_BreakPointToggle_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _Edit_BreakPointToggle.Click
        If CommandBarControl.Enabled = True Then
            Call BreakPointToggle()
        End If
    End Sub


    Private Sub _MenuBookmarks_BookmarkRemoveAll_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _MenuBookmarks_BookmarkRemoveAll.click
        _VBInstance.VBProjects.ClearBookmarks()
    End Sub


    Private Sub _MenuBookmarks_BookmarkToggle_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _MenuBookmarks_BookmarkToggle.Click
        Call BookmarkToggle(CommandBarControl, handled, CancelDefault)
    End Sub


    Private Sub _MenuDebug_BreakPointRemoveAll_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _MenuDebug_BreakPointRemoveAll.Click
        _VBInstance.VBProjects.ClearBreakPoints()
    End Sub


    Private Sub _MenuDebug_BreakPointToggle_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _MenuDebug_BreakPointToggle.Click
        If CommandBarControl.enabled Then
            Call BreakPointToggle()
        End If
    End Sub


    Private Sub _MenuMainDebug_BreakPointRemoveAll_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _MenuMainDebug_BreakPointRemoveAll.Click
        _VBInstance.VBProjects.ClearBreakPoints()
    End Sub


    Private Sub _MenuMainDebug_BreakPointToggle_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _MenuMainDebug_BreakPointToggle.click
        If CommandBarControl.enabled Then
            Call BreakPointToggle()
        End If
    End Sub


    Private Sub _Toggle_BookmarkToggle_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _Toggle_BookmarkToggle.Click
        Call BookmarkToggle(CommandBarControl, handled, CancelDefault)
    End Sub


    Private Sub _Toggle_BreakPointToggle_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _Toggle_BreakPointToggle.Click
        If CommandBarControl.enabled Then
            Call BreakPointToggle()
        End If
    End Sub


    ''' <summary>
    ''' Actually handle toggling a breakpoint
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub BreakPointToggle()
        '---- if we're setting the mark, ignore this click
        If _bSettingMarks = True Then Exit Sub

        Dim StartPos As Integer

        With _VBInstance
            '---- Can't setfocus because the VB may be in run mode
            '.ActiveCodePane.Window.SetFocus()

            'we just need to current cursor position
            .ActiveCodePane.GetSelection(StartPos, 0, 0, 0)

            Dim BreakPoints = .ActiveVBProject.BreakPoints

            Dim bp = BreakPoints.Find(.ActiveCodePane.CodeModule.Parent.Name, StartPos)
            If bp IsNot Nothing Then
                BreakPoints.Remove(bp)
                Logger.Log("Clearing BreakPoint " & bp.ToString)
            Else
                bp = BreakPoints.Add(.ActiveCodePane.CodeModule.Parent.Name, StartPos)
                Logger.Log("Setting BreakPoint " & bp.ToString)
                _InvalidBreakTimer.Stop()
                _InvalidBreakTimer.Tag = 0
                _LastBP = bp
                _InvalidBreakTimer.Interval = 50
                _InvalidBreakTimer.Start()
            End If
        End With

        '---- can't do this cause VB may be in run mode
        '_VBInstance.ActiveCodePane.Window.SetFocus()
    End Sub


    Private Sub BookmarkToggle(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean)
        '---- if we're setting the mark, ignore this click
        If _bSettingMarks = True Then Exit Sub

        Try
            Dim StartPos As Integer

            With _VBInstance
                'we just need to current cursor position
                .ActiveCodePane.Window.SetFocus()
                .ActiveCodePane.GetSelection(StartPos, 0, 0, 0)

                Dim Bookmarks = .ActiveVBProject.Bookmarks

                Dim bm = Bookmarks.Find(.ActiveCodePane.CodeModule.Parent.Name, StartPos)
                If bm IsNot Nothing Then
                    Bookmarks.Remove(bm)
                    Logger.Log("Clearing Bookmark " & bm.ToString)
                Else
                    'init if nothing
                    If CommandBarControl.Enabled = True Then
                        bm = Bookmarks.Add(.ActiveCodePane.CodeModule.Parent.Name, StartPos)
                        Logger.Log("Setting Bookmark " & bm.ToString)
                    End If
                End If
            End With

            _VBInstance.ActiveCodePane.Window.SetFocus()
        Catch ex As Exception
            ex.Show()
        End Try
    End Sub


    ''' <summary>
    ''' This routine is called when we detect a breakpoint set or cleared via the left hand margin
    ''' </summary>
    ''' <param name="LineOffset"></param>
    ''' <param name="bSet"></param>
    ''' <remarks></remarks>
    Private Sub BreakpointSetValue(LineOffset As Integer, bSet As Boolean)
        Try
            Dim StartPos As Integer

            With _VBInstance
                'we just need to current cursor position
                '---- this can cause issues when in run mode, plus there should
                '     be no need for it, since this func can only be called after the user
                '     has actually clicked in the window.
                '.ActiveCodePane.Window.SetFocus()
                StartPos = .ActiveCodePane.TopLine + LineOffset - 1

                Dim Breakpoints = .ActiveVBProject.BreakPoints

                Dim bp = Breakpoints.Find(.ActiveCodePane.CodeModule.Parent.Name, StartPos)
                If bp IsNot Nothing Then
                    If bSet = False Then
                        Breakpoints.Remove(bp)
                        Logger.Log("Clearing Breakpoint " & bp.ToString)
                    Else
                        '---- Already set
                        Logger.Log("Breakpoint already set at line " & bp.ToString)
                    End If
                Else
                    'init if nothing
                    If bSet = True Then
                        bp = Breakpoints.Add(.ActiveCodePane.CodeModule.Parent.Name, StartPos)
                        Logger.Log("Setting Breakpoint " & bp.ToString)
                    Else
                        Logger.Log("BreakPoint already clear at line " & StartPos.ToString)
                    End If
                End If
            End With

            '---- this can cause issues when in run mode, plus there should
            '     be no need for it, since this func can only be called after the user
            '     has actually clicked in the window.
            '_VBInstance.ActiveCodePane.Window.SetFocus()
        Catch ex As Exception
            ex.Show()
        End Try
    End Sub


    Public Sub RestoreAllBookmarks()
        Try
            For Each Project As VBIDE.VBProject In _VBInstance.VBProjects
                Project.Bookmarks.Load()
                RestoreProjectBookMarks(Project)
                DoEvents()
            Next

        Catch ex As Exception
            ex.Show()
        End Try
    End Sub


    Private Sub RestoreProjectBookMarks(ByVal Project As VBIDE.VBProject)
        Try
            Logger.Log("Restoring bookmarks for " & Project.Name)
            _bSettingMarks = True
            Dim Bookmarks = Project.Bookmarks

            For Each bm In Bookmarks
                Dim Comp = Project.VBComponents.Find(bm.ModuleName)
                If Comp IsNot Nothing Then
                    '---- use this to check whether the components window is visible or not
                    Dim bVis = Comp.IsCodeWindowVisible
                    '---- note, just accessing the CodePane object will show the window
                    With Comp.CodeModule.CodePane
                        .Window.Visible = True
                        DoEvents()
                        .Window.SetFocus()
                        DoEvents()
                        .SetSelection(bm.LineNumber, 1, bm.LineNumber, 1)
                        'DoEvents()
                        '.Window.SetFocus()
                        'DoEvents()
                        If mcbMenuBookmarkToggle.Enabled = True Then
                            mcbMenuBookmarkToggle.Execute()
                        End If
                        DoEvents()
                        '.Window.SetFocus()
                        'DoEvents()
                        If bVis = False Then .Window.Visible = False
                        DoEvents()
                    End With
                End If
            Next

        Catch ex As Exception
            ex.Show("Unable to restore bookmarks")

        Finally
            _bSettingMarks = False
        End Try
    End Sub


    Public Sub RestoreAllBreakpoints()
        Try
            For Each Project As VBIDE.VBProject In _VBInstance.VBProjects
                Project.BreakPoints.Load()
                RestoreProjectBreakpoints(Project)
                DoEvents()
            Next

        Catch ex As Exception
            ex.Show()
        End Try
    End Sub


    Private Sub RestoreProjectBreakpoints(ByVal Project As VBIDE.VBProject)
        Try
            Logger.Log("Restoring breakpoints for " & Project.Name)
            _bSettingMarks = True

            Dim BreakPoints = Project.BreakPoints

            For Each bp In BreakPoints
                Dim Comp = Project.VBComponents.Find(bp.ModuleName)
                If Comp IsNot Nothing Then
                    '---- use this to check if a code window is visible
                    '     because actually accessing the CODEPANE object will
                    '     make it visible
                    Dim bVis = Comp.IsCodeWindowVisible
                    With Comp.CodeModule.CodePane
                        .Window.Visible = True
                        DoEvents()
                        .Window.SetFocus()
                        DoEvents()
                        .SetSelection(bp.LineNumber, 1, bp.LineNumber, 1)
                        DoEvents()
                        .Window.SetFocus()
                        DoEvents()
                        If mcbMenuBreakPointToggle.Enabled = True Then
                            mcbMenuBreakPointToggle.Execute()
                        End If
                        DoEvents()
                        '.Window.SetFocus()
                        'DoEvents()
                        If bVis = False Then .Window.Visible = False
                        DoEvents()
                    End With
                End If
            Next

        Catch ex As Exception
            ex.Show()

        Finally
            _bSettingMarks = False
        End Try
    End Sub


    Private Sub BreakPointNext(Optional DirUp As Boolean = False)
        Try
            '---- get ALL breakpoints collected into a single, sortable collection
            Dim AllBreakPoints As List(Of Breakpoint) = New List(Of Breakpoint)
            For Each proj As VBProject In _VBInstance.VBProjects
                AllBreakPoints.AddRange(proj.BreakPoints)
            Next

            '---- order all breakpoints (backwards if searching prev)
            If DirUp Then
                AllBreakPoints = (From bp In AllBreakPoints Order By bp.Parent.Parent.Name Descending, bp.ModuleName Descending, bp.LineNumber Descending).ToList
            Else
                AllBreakPoints = (From bp In AllBreakPoints Order By bp.Parent.Parent.Name, bp.ModuleName, bp.LineNumber).ToList
            End If

            Dim CurLine As Integer
            _VBInstance.ActiveCodePane.GetSelection(CurLine, 0, 0, 0)
            Dim CurProject = _VBInstance.ActiveVBProject.Name
            Dim CurModule = _VBInstance.ActiveCodePane.CodeModule.Parent.Name

            Dim NextBreakPoint As Breakpoint = Nothing
            If DirUp Then
                For Each bp In AllBreakPoints
                    If bp.Parent.Parent.Name < CurProject Then NextBreakPoint = bp : Exit For
                    If bp.Parent.Parent.Name = CurProject Then
                        If bp.ModuleName < CurModule Then NextBreakPoint = bp : Exit For
                        If bp.ModuleName = CurModule Then
                            If bp.LineNumber < CurLine Then NextBreakPoint = bp : Exit For
                        End If
                    End If
                Next
            Else
                For Each bp In AllBreakPoints
                    If bp.Parent.Parent.Name > CurProject Then NextBreakPoint = bp : Exit For
                    If bp.Parent.Parent.Name = CurProject Then
                        If bp.ModuleName > CurModule Then NextBreakPoint = bp : Exit For
                        If bp.ModuleName = CurModule Then
                            If bp.LineNumber > CurLine Then NextBreakPoint = bp : Exit For
                        End If
                    End If
                Next
            End If
            If NextBreakPoint Is Nothing Then
                '---- just wrap to first breakpoint
                NextBreakPoint = AllBreakPoints.FirstOrDefault
            End If
            If NextBreakPoint IsNot Nothing Then
                Dim Comp = NextBreakPoint.Parent.Parent.VBComponents.Find(NextBreakPoint.ModuleName)
                If Comp IsNot Nothing Then
                    '---- use this to check if a code window is visible
                    '     because actually accessing the CODEPANE object will
                    '     make it visible
                    Dim bVis = Comp.IsCodeWindowVisible
                    With Comp.CodeModule.CodePane
                        .Window.Visible = True
                        DoEvents()
                        .Window.SetFocus()
                        DoEvents()
                        .SetSelection(NextBreakPoint.LineNumber, 1, NextBreakPoint.LineNumber, 1)
                        DoEvents()
                        .Window.SetFocus()
                        DoEvents()
                    End With
                End If
            End If

        Catch ex As Exception
            ex.Show()

        End Try
    End Sub


    ''' <summary>
    ''' If any collections have become dirty, save them
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CheckForSaves()
        For Each Project As VBProject In _VBInstance.VBProjects
            If Project.Bookmarks.IsDirty Then Project.Bookmarks.Save() : Logger.Log("SAVED Bookmarks")
            If Project.BreakPoints.IsDirty Then Project.BreakPoints.Save() : Logger.Log("SAVED Breakpoints")
        Next
    End Sub


    ''' <summary>
    ''' Returns whether or not VB6 is in runmode
    ''' You could also use the VBBuildEvents sink
    ''' This is not currently used, but kept for reference
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsInRunMode() As Boolean
        Return (_VBInstance.CommandBars("File").Controls(1).Enabled = False)
    End Function


    ''' <summary>
    ''' Trap several specific keys, plus the macro keys we want to implement (for bookmarks)
    ''' </summary>
    ''' <param name="Key"></param>
    ''' <remarks></remarks>
    Private Sub _KeyHook_KeyUp(Key As System.Windows.Forms.Keys) Handles _KeyHook.KeyUp
        Try
            Select Case Key
                '---- in this case, i'm just watching for an f9 so I can handle what I need to do
                '     internally
                Case (System.Windows.Forms.Keys.F9 Or System.Windows.Forms.Keys.Shift)
                    'Break Point Clear
                    _VBInstance.VBProjects.ClearBreakPoints()

                Case System.Windows.Forms.Keys.F9
                    '---- ie f9 key alone, no modifiers
                    '     The menu is indeed disabled while in run mode
                    '     BUT the F9 key still toggles the breakpoint
                    '     so we can't ignore the keystroke
                    'If mcbMenuBreakPointToggle.Enabled = True Then
                    BreakPointToggle()
                    'End If

                    '---- 
                    '---- 
                    '---- These are driven by properties so that ostensibly they can be configured
                    '---- 
                    '---- 
                Case Me.BookMarkToggleHotkey
                    '---- in this case, the hook is to actually toggle a bookmark
                    '     No need to test whether the menu is enabled or not
                    '     if it's disabled, VB won't let it happen anyway
                    mcbMenuBookmarkToggle.Execute()
                Case Me.BookMarkPrevHotkey
                    mcbMenuBookmarkPrev.Execute()
                Case Me.BookMarkNextHotkey
                    mcbMenuBookmarkNext.Execute()
                Case Me.BreakPointPrevHotkey
                    BreakPointNext(True)
                Case Me.BreakPointNextHotkey
                    BreakPointNext()
            End Select
        Catch
        End Try
    End Sub


    Private Sub _VBProjectsEvents_ItemActivated(VBProject As VBIDE.VBProject) Handles _VBProjectsEvents.ItemActivated
        _VBComponentsEvents = _VBInstance.Events.VBComponentsEvents(VBProject)
    End Sub


    ''' <summary>
    ''' When a project is added, load it's bookmarks and breakpoints
    ''' </summary>
    ''' <param name="VBProject"></param>
    ''' <remarks></remarks>
    Private Sub _VBProjectsEvents_ItemAdded(VBProject As VBIDE.VBProject) Handles _VBProjectsEvents.ItemAdded
        Try
            Logger.Log("Restoring marks on Project added")
            VBProject.Bookmarks.Load()
            RestoreProjectBookMarks(VBProject)

            VBProject.BreakPoints.Load()
            RestoreProjectBreakpoints(VBProject)

        Catch ex As Exception
            ex.Show("Unable to restore breakpoints and bookmarks.")
        End Try
    End Sub


    ''' <summary>
    ''' Interestingly, this event fires when the VBG is closed, AND when projects are removed from the group
    ''' </summary>
    ''' <param name="VBProject"></param>
    ''' <remarks></remarks>
    Private Sub _VBProjectsEvents_ItemRemoved(VBProject As VBIDE.VBProject) Handles _VBProjectsEvents.ItemRemoved

        '---- i'm always saving the marks, no options for this right now
        'If GetSetting(My.Application.Info.Title, "Settings", "AutoSaveBKP", vbChecked) = vbChecked Then
        VBProject.Bookmarks.Save()

        'If GetSetting(My.Application.Info.Title, "Settings", "AutoSaveBKP", vbChecked) = vbChecked Then
        VBProject.BreakPoints.Save()

        '---- and release all the marks
        VBProject.Release()
    End Sub


    Private Sub _EditTimer_Tick(sender As Object, e As System.EventArgs) Handles _EditTimer.Tick
        '---- as this timer ticks, we have to look at the active code window (if any)
        '     and account for any coding change that might have repositioned 
        '     bookmarks and breakpoints
        Try
            If _VBInstance.ActiveCodePane Is Nothing Then Exit Sub

            CheckForCodeChanges()

            '---- check every x minutes if anything's gotten dirty and save if it has
            Static LastSave As Date = Now()
            If Now.Subtract(LastSave).Seconds > 30 Then
                CheckForSaves()
            End If

        Catch ex As Exception
            ex.Show()
        End Try
    End Sub


    ''' <summary>
    ''' This routine is called constantly to monitor the active code window for any edits
    ''' Any edits that change then number of lines in the window will cause bookmarks and
    ''' breakpoints to shift, which I have to account for here
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CheckForCodeChanges()
        '---- This routine assumes that a codepane is active!
        Static PreviousCountofLines As Integer
        Static PrevStartLine As Integer
        Static PrevStartCol As Integer
        Static PrevEndLine As Integer
        Static PrevEndCol As Integer
        Static PrevAtBOL As Boolean
        Static PrevAtEOL As Boolean
        Static PrevBlankLine As Boolean
        Static PreviousCodePaneModule As String

        Try
            Dim StartLine As Integer
            Dim StartCol As Integer
            Dim EndLine As Integer
            Dim EndCol As Integer
            Dim CurCountOfLines As Integer
            Dim NumberOfLineDiff As Integer
            Dim CurCodePaneModule = _VBInstance.ActiveCodePane.CodeModule.Parent.Name

            'get count of lines
            CurCountOfLines = _VBInstance.ActiveCodePane.CodeModule.CountOfLines

            'get current line pos
            _VBInstance.ActiveCodePane.GetSelection(StartLine, StartCol, EndLine, EndCol)
            Dim CurLine = _VBInstance.ActiveCodePane.CodeModule.Lines(StartLine, 1)
            Dim ChangedLineCount = EndLine - StartLine + 1
            Dim AtBOL = StartCol <= (Len(CurLine) - Len(CurLine.TrimStart) + 1)
            Dim AtEOL = StartCol >= (Len(CurLine) - Len(CurLine.TrimEnd) + 1)
            Dim IsBlankLine = Len(CurLine.Trim) = 0

            '---- if we've switched codepanes, reset some vars
            If PreviousCodePaneModule <> CurCodePaneModule Then
                PreviousCodePaneModule = CurCodePaneModule
                PreviousCountofLines = CurCountOfLines
                PrevStartLine = StartLine
                PrevStartCol = StartCol
                PrevEndLine = EndLine
                PrevEndCol = EndCol
                PrevAtBOL = AtBOL
                PrevAtEOL = AtEOL
                PrevBlankLine = IsBlankLine
            End If

            '---- there has to be some lines
            If PreviousCountofLines <> 0 Then
                'check if lines have been added
                If PreviousCountofLines < CurCountOfLines Then
                    Logger.Log(String.Format("Line count dropped. Previous Line count: {0}", PreviousCountofLines))
                    '---- adjust if cursor was and still is at start of line, because VB would
                    '     have moved any bps/bms down with the line
                    If PrevAtBOL And AtBOL Then StartLine -= 1
                    NumberOfLineDiff = CurCountOfLines - PreviousCountofLines
                    Logger.Log(String.Format("Adjustment A: NumLinesDiff={0} PrevAtBol={1} AtBol={2}  CurCountOfLines={3} StartLine={4}", NumberOfLineDiff, PrevAtBOL, AtBOL, CurCountOfLines, StartLine))
                    _VBInstance.ActiveCodePane.CodeModule.LineChange(LineModsEnum.Added, StartLine, NumberOfLineDiff)
                End If

                'check if lines have been removed
                If PreviousCountofLines > CurCountOfLines Then
                    Logger.Log(String.Format("Line count increased. Previous Line count: {0}", PreviousCountofLines))
                    '---- if the line we were on was blank, VB will pull the break point from the nextline
                    '     on to the new line
                    Dim Adj = 0
                    If PrevBlankLine And StartLine > PrevStartLine Then Adj += -1

                    NumberOfLineDiff = PreviousCountofLines - CurCountOfLines

                    '---- if we were at the end of prev line, and we're still on that line
                    '     then this line was joined with the next line, so the NEXT line's breakpoint etc
                    '     has been lost, so act like the NEXT line was deleted
                    If PrevAtEOL And NumberOfLineDiff = 1 And StartLine = PrevStartLine Then Adj += 1

                    If AtEOL And NumberOfLineDiff = 1 And StartLine < PrevStartLine Then Adj += 1
                    '---- apply any adjustments
                    StartLine += Adj

                    Logger.Log(String.Format("Adjustment B: NumLinesDiff={0} PrevAtEOL={1} AtEOL={2}  Adj={3} StartLine={4}", NumberOfLineDiff, PrevAtEOL, AtEOL, Adj, StartLine))
                    _VBInstance.ActiveCodePane.CodeModule.LineChange(LineModsEnum.Removed, StartLine, NumberOfLineDiff)
                End If

                If PreviousCountofLines = CurCountOfLines Then
                    '---- lines have just been changed, the prob is, if a line with a breakpoint was commented
                    '     we've just lost the breakpoint
                    _VBInstance.ActiveCodePane.CodeModule.LineChange(LineModsEnum.Changed, StartLine, ChangedLineCount)
                End If
            End If
            PreviousCountofLines = CurCountOfLines
            PrevStartLine = StartLine
            PrevStartCol = StartCol
            PrevEndLine = EndLine
            PrevEndCol = EndCol
            PrevAtBOL = AtBOL
            PrevAtEOL = AtEOL
            PrevBlankLine = IsBlankLine
        Catch ex As Exception
            Logger.Log(ex.ToString)
        End Try
    End Sub


    ''' <summary>
    ''' This is the watchdog timer that stops us watching for a "bad breakpoint" msg eventually
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub _InvalidBreakTimer_Tick(sender As Object, e As System.EventArgs) Handles _InvalidBreakTimer.Tick
        '---- when this ticks
        '     enum all windows to find if the warning window is up
        Try
            _InvalidBreakTimer.Tag = Val(_InvalidBreakTimer.Tag) + 1
            If Val(_InvalidBreakTimer.Tag) < 100 Then
                Dim hwnds() = FindWindowLike(0, "Breakpoint not allowed on this line", "Static")
                If hwnds.Count > 0 Then
                    If _VBInstance.ActiveVBProject.BreakPoints.Contains(_LastBP) Then
                        '---- looks like it was (and never got set) so remove it
                        Logger.Log("Removing " & _LastBP.ToString)
                        _VBInstance.ActiveVBProject.BreakPoints.Remove(_LastBP)
                    End If
                Else
                    '---- keep watching
                    Exit Sub
                End If
            End If
            _InvalidBreakTimer.Stop()
            _LastBP = Nothing

        Catch ex As Exception
            ex.Show("Problem while watching for invalid breakpoint message")
        End Try
    End Sub


    ''' <summary>
    ''' This subclasser watches the MDIChild window for WM_MOUSEACTIVATE messages
    ''' so I can know what code window is currently active when it's activated
    ''' and subclass it for mouse events
    ''' </summary>
    ''' <param name="Sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub _MDIChildSubClasser_PostWndProc(Sender As Object, e As Subclasser.WndProcEventArgs) Handles _MDIChildSubClasser.PostWndProc
        Static ActiveVBCodeHwnd As IntPtr
        Static LastPt As User32.POINT
        Static DownHandled As Integer = -1
        Try
            Select Case e.Msg
                Case Win32.Messages.WM_PARENTNOTIFY
                    '---- we get this message everytime the lbutton is clicked in a VBCode child window
                    Dim msg As DWordAsXY
                    msg.DWord = e.wParam.ToInt32
                    If msg.X = Win32.Messages.WM_LBUTTONDOWN Then
                        '---- this is the only one we care about
                        '     split the lparam into x and y
                        msg.DWord = e.lParam.ToInt32
                        '---- convert the point to screen coords
                        LastPt.X = msg.X
                        LastPt.Y = msg.Y
                        User32.ClientToScreen(e.hWnd, LastPt)
                        'Logger.Log("ParentNotify={0} {1}", LastPt.X, LastPt.Y)
                        DownHandled = 0
                    End If

                Case Win32.Messages.WM_SETCURSOR
                    Dim hwnd = e.wParam
                    If hwnd.ToInt32 > 0 AndAlso hwnd <> ActiveVBCodeHwnd Then
                        If WinAPI.GetClassName(hwnd) = "VbaWindow" Then
                            ActiveVBCodeHwnd = hwnd
                        Else
                            ActiveVBCodeHwnd = 0
                        End If
                    End If

                    If DownHandled = 0 Then
                        DownHandled = 1
                    ElseIf DownHandled = 1 Then
                        '---- key off the second SETCURSOR message
                        If hwnd.ToInt32 <> 0 AndAlso hwnd = ActiveVBCodeHwnd Then
                            '---- this is the window we care about
                            '     convert the screen coords to this window's client coords
                            User32.ScreenToClient(hwnd, LastPt)
                            'Logger.Log("Subclassed msg={0} {1} {2}", Hex(e.Msg.ToInt32), LastPt.X, LastPt.Y)
                            HandleCodeWindowClick(hwnd, LastPt)
                        End If
                        DownHandled = -1
                    End If
            End Select
            'Logger.Log("Subclassed msg={0} {1} {2}", Hex(e.Msg.ToInt32), e.wParam, e.lParam)

        Catch ex As Exception
            ex.Show()
        End Try
    End Sub


    Private Sub HandleCodeWindowClick(hwnd As IntPtr, Pt As User32.POINT)
        Try
            '---- if they click outside the margin, just ignore it
            If Pt.X > 22 Then
                Logger.Log("Click Not in margin")
                Exit Sub
            End If

            '---- Well, it turns out that .net's text size calcs are almost the same 
            '     as win32's, but not quite, so I have to use the real API
            '     calls to get +exactly+ the same size text.
            Static LineHeight As Integer
            Static LastFont As Font
            Dim hdc = WinAPI.GetDC(hwnd)
            Dim sz As apiSIZE
            Dim f = _VBInstance.EditorFont
            If LineHeight = 0 OrElse LastFont Is Nothing OrElse Not f.Equals(LastFont) Then
                Dim hFont = _VBInstance.EditorFont.ToHfont
                Dim hFontOld = WinAPI.SelectObject(hdc, hFont)
                Dim r = WinAPI.GetTextExtentPoint32(hdc, "Wg", 2, sz)
                LineHeight = sz.cy
                WinAPI.SelectObject(hdc, hFontOld)
                WinAPI.ReleaseDC(hwnd, hdc)
            End If

            '---- we got a mousedown, so capture the screen rect for the line we're on
            Dim LineNum As Integer = Fix((Pt.Y - 30) \ LineHeight) + 1
            Dim img = ScreenShot.Capture.ClientWindowRect(hwnd, New Rectangle(0, ((LineNum - 1) * LineHeight) + 30, 22 + 100, LineHeight + 100))

            '---- draw a square over where I think the dot should be, for testing
            'Using g = Graphics.FromImage(img)
            '    g.DrawRectangle(New Pen(Color.DarkSalmon), 5, LineHeight \ 2 - 6, 12, 12)
            'End Using

            'show form for testing
            'Static frm As frmScreenShot
            'If frm Is Nothing Then frm = New frmScreenShot
            'frm.Show()
            'frm.BackgroundImage = img

            '---- Test for dot, note that the bookmark mark doesn't overlap the edges of the
            '     breakpoint dot, so I can still test for them.
            '     test all 4 sides of the breakpoint dot for black, if they're all there, 
            '     we have to have a breakpoint!
            Dim bmp As Bitmap = img
            Dim clr = If(bmp.GetPixel(6, LineHeight \ 2).ToArgb = Color.Black.ToArgb, 1, 0)
            clr += If(bmp.GetPixel(16, LineHeight \ 2).ToArgb = Color.Black.ToArgb, 1, 0)
            clr += If(bmp.GetPixel(11, LineHeight \ 2 - 5).ToArgb = Color.Black.ToArgb, 1, 0)
            clr += If(bmp.GetPixel(11, LineHeight \ 2 + 5).ToArgb = Color.Black.ToArgb, 1, 0)
            Dim IsDot = clr = 4

            Logger.Log(String.Format("Down at Line {0} {1} {2} lh={3} IsDot={4}", LineNum, Pt.X, Pt.Y, LineHeight, IsDot))
            BreakpointSetValue(LineNum, IsDot)

        Catch ex As Exception
            ex.Show()
        End Try
    End Sub


    ''' <summary>
    ''' When a component is removed, remove all breakpoints and bookmarks associated with it
    ''' </summary>
    ''' <param name="VBComponent"></param>
    ''' <remarks></remarks>
    Private Sub _VBComponentsEvents_ItemRemoved(VBComponent As VBIDE.VBComponent) Handles _VBComponentsEvents.ItemRemoved
        Dim BMRemove As List(Of Bookmark) = New List(Of Bookmark)
        For Each bm As Bookmark In VBComponent.Collection.Parent.Bookmarks
            If StrComp(bm.ModuleName, VBComponent.Name, CompareMethod.Text) = 0 Then
                BMRemove.Add(bm)
            End If
        Next
        For Each bm As Bookmark In BMRemove
            VBComponent.Collection.Parent.Bookmarks.Remove(bm)
        Next


        Dim BPRemove As List(Of Breakpoint) = New List(Of Breakpoint)
        For Each bp As Breakpoint In VBComponent.Collection.Parent.BreakPoints
            If StrComp(bp.ModuleName, VBComponent.Name, CompareMethod.Text) = 0 Then
                BPRemove.Add(bp)
            End If
        Next
        For Each bp As Breakpoint In BPRemove
            VBComponent.Collection.Parent.BreakPoints.Remove(bp)
        Next
    End Sub


    ''' <summary>
    ''' when a component is renamed, we have to change all bookmark/breakpoint names as well
    ''' </summary>
    ''' <param name="VBComponent"></param>
    ''' <param name="OldName"></param>
    ''' <remarks></remarks>
    Private Sub _VBComponentsEvents_ItemRenamed(VBComponent As VBIDE.VBComponent, OldName As String) Handles _VBComponentsEvents.ItemRenamed
        For Each bm As Bookmark In VBComponent.Collection.Parent.Bookmarks
            If StrComp(bm.ModuleName, OldName, CompareMethod.Text) = 0 Then
                bm.ModuleName = VBComponent.Name
            End If
        Next

        For Each bp As Breakpoint In VBComponent.Collection.Parent.BreakPoints
            If StrComp(bp.ModuleName, OldName, CompareMethod.Text) = 0 Then
                bp.ModuleName = VBComponent.Name
            End If
        Next
    End Sub


    Private Sub _MenuHandler_Click(CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles _MenuHandler.Click
        Dim frm = New ConfigurationForm
        frm.ShowDialog()
    End Sub
End Class



<StructLayout(LayoutKind.Explicit)> _
<ComVisible(False)> _
Public Structure DWordAsXY
    <FieldOffset(0)> Public DWord As Int32
    <FieldOffset(0)> Public X As Int16
    <FieldOffset(2)> Public Y As Int16
End Structure
