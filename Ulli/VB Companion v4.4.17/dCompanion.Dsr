VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} dCompanion 
   ClientHeight    =   2925
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   3270
   _ExtentX        =   5768
   _ExtentY        =   5159
   _Version        =   393216
   Description     =   "Adds mousewheel support, autocomplete and a few other goodies to the VB IDE"
   DisplayName     =   "Ulli's VB Companion"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "dCompanion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'© 2002/2007    UMGEDV GmbH  (umgedv@yahoo.com)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Author         UMG (Ulli K. Muehlenweg)
'
'
'Title          Ulli's VB Companion
'
'               Adds Mouse Wheel Support and Auto Complete plus a few other Goodies to
'               the VB IDE.
'
'               Simply compile the .DLL into your VB folder and restart VB.
'
'Notes          Using this Add-In together with other Add-Ins which also subclass the IDE
'               may crash VB because VB apparently fires the OnBeginShutdown events not
'               in reverse order of OnStartupComplete. Therefore unhooking the IDE will
'               not work correctly and will eventually leave VB with a dead hook to a
'               non-existent window procedure.
'======
'How To     >>  Hold down secondary (right) mouse button and rotate wheel to open
'======         Options Dialog Box.
'
'           >>  Hold down Shift or Cntl or both while scrolling to decrease scroll
'               distance temporarily.
'
'           >>  Press Pause key to open member list.
'               The member names shown are those only which are within scope of the current
'               code module.
'               Click on any member name to insert it into your code.
'               Click on column header to sort by that column.
'               Click on Exit or press Escape to abandon.
'
'   NEW     >>  Press Shift+Pause keys to open multiline literal box. The literal box
'               will help you to design long multiline text literals and convert them to
'               the proper VB syntax, including newlines, quotes and line continuation
'               marks.
'               Enter the text as you would like to see it during program execution.
'               Press Pause key or click on Preview to see an actual example of the
'               VB-interpretation of the converted "syntaxed" text in a message box.
'               Press Shift+Return to insert the converted text into your code.
'               Press Escape to abandon.
'               You can also copy a literal from the codepane and paste that into
'               the box, the VB syntax will be un-converted during the paste process.
'
'   NEW     >>  Click on Reset (main menu bar) when it's green to reset all changes
'               (code AND visual elements) which were made since you last saved the
'               current component.
'
'   NEW     >>  Click on Compare (main menu bar) to see all alterations of the current
'               component (code AND visual elements) which were made since you last
'               saved the it.
'
'   NEW     >>  Click on OpenAll (main menu bar) to open all available codepanes.
'
'   NEW     >>  Click on Copy (main menu bar) to open the Copy Falility. In the CF you
'               can open and display any VB file, mark text and add it incrementally to
'               an internal clipboard. Once done you can then paste the contents of the
'               clipboard into the current module at the caret position.
'
'           >>  Press any dead key (Cursor left/right/up/down, Page up/down, Pos1, End)
'               to get you out of the selected text and confirm the autocomplete.
'
'           >>  You can now type a questionmark after an API name which fails to trigger
'               autocompletion. If it still fails it isn't there (or faulty).
'
'           >>  Click middle button or mousewheel to return to caret position.
'
'           >>  Click on the horiz scrollbar thumb just once if the raster fails to adjust
'               to the horizontal position or doesn't show at all.
'
'           >>  In the Options Popup, when you turn on the Raster option a color dialog
'               will let you select the raster color.
'
'           >>  Click 'Refresh/Reset' if you have altered anything in the IDE that would
'               affect this AddIn (Font, FontSize, Indicatorbar or TabWidth) or if you
'               wish to go back to the last saved registry settings.
'
'**********************************************************************************
'Development History
'**********************************************************************************
'
'01Feb2007 Version 4.4.17   UMG
'
'Added Copy Facility and command bar button.
'The CF was generally overhauled. Formatter now uses almost the same CF.
'
'Added option to select ROP for drawing the grid. MaskPen did only work for white bckgnds.
'Tnx to Kibe for reporting this quirk.
'
'Streamlined some code, in particular with regard to the flashing arrow <===.
'Modified flashing arrow to honor font height.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'01Mar2006 Version 4.3.7    UMG
'
'Changed fCompare to make it insensitive to multiple spaces (except in literals).
'Fixed bug in cMD5 - Round 4 ACs(53) was numbered incorrectly.
'Fixed bug in fMultiline - had difficulties with quoted literals.
'Rewrote 25 lines limitation in fMultiline.
'Dis- and enabled Compare- and Reset-Menu-Buttons when required.
'Hide Caret while flashing caret region. Did not look too good with an overwrite caret.
'Added Idle Detection.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'22Nov2004 Version 4.2.8     UMG
'
'Fixed bug in compare window (fCompare) code coloring routine.
'Moved Other VB words into a loooong string.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'03Sep2004 Version 4.2.6    UMG
'
'Bugfix
'Added check in mSubclass.UnhookCodePane to see whether the current ActiveComponent
'is in fact 'something'.
'
'Added Autocomplete trigger Mode Unique: this mode autocompletes the words only
'when there is a unique match between the typed word fragment and the names
'available. An example will clarify the difference to the trigger length mode:
'
'   avalilable names McGyver
'                    McGovern
'                    Nobody
'
'   User types                  Response
'                   Length Mode(3)      Unique Mode
'
'   N                                   Nobody
'   No                                  Nobody
'   Nob             Nobody              Nobody
'
'   McG             McGyver
'   McGo            McGovern            McGovern
'
'That is - the Unique mode will respond when there is a single match only no matter how long,
'whereas the Trigger Length mode will respond on the first match (of possibly several)
'which has the reqired trigger length.
'
'This is in a way equivalent to a variable auto-adjusted trigger length.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'27Aug2004 Version 4.1.12    UMG
'
'Mousewheel now has a new name - you can call me Companion.
'
'Added fMultiline - this box allows you to enter multiline literals as you would
'like them to appear in a message box for example (WYSIWYG). The text is then expanded
'and formatted into a VB compliant syntax and inserted into your code at the caret position.
'Tnx to Evan Toder for letting me steal his idea; the code is completely different from
'his though: expansion/formatting is accomplished in one (admittedly long) line of code
'which only the compiler understands (and me ;-} of course).
'
'Added menu button Reset: clicking this button will undo ALL changes in the active module
'(code AND visual components) and reset it to the last saved state; the dirty flag is
'also reset. The dirty flag indicates that changes were made during the current session
'which have not yet been saved.
'
'Added menu button OpenAll: clicking this button will open all codepanes of each loaded
'project.
'
'Added Comparator. Click on menu button Compare to compare the current state of a module with
'it's last saved state.
'
'Fixed annoying bug in Function IsInCode regarding the 'As' keyword preceeding the currently
'typed word or word fragment; will no more interfere with VBs variable type popups
'unless you type 'AS' in capitals to permit autocompleting the variable type.
'
'Added resource string lookup in mMsgBoxEx.
'
'Added horizontal grid or raster lines. Selecting this option may make mouse-scrolling
'a little sluggish if your PC is one of the slower kind.
'
'Added resource string lookup to find the main window caption fragment which indicates
'that we are in design mode.
'
'Clarified some code passages and fixed a few minor qirks. Code cosmetics, in particular
'variable names and duplicated names or definitions / declarations...
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'03Jul2004 Version 3.2.3    UMG
'
'Added mMsgBoxEx - extended message box based on Ray Mercer's code but heavily modified
'and debugged.
'Fixed quirk in fSetOpts validation of txtTriggerLength
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'27Jan2004 Version 3.1.7    UMG
'
'Added check for correct api database.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'24Jan2004 Version 3.1.6    UMG
'
'GPF bug fix. Occured when there was no open code pane but an open form.
'Modified Icon show and hide.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'20Jan2004 Version 3.1.5    UMG
'
'Changed Options popup to secondary (right) mousebutton + scrollwheel.
'Added Fader.
'Changed Flashing Region again (is now a polygonic region).
'Fixed fSelect sizing/positioning and added border.
'Added scope count in fgdMembers.
'Added code window topline correction if code window size was changed.
'Added VB tab width recognition with ApiType body.
'No scan thru fgdMembers on exit keys.
'Added a few comments.
'
'New:
'----
'This AddIn can now optionally display a vertical raster or grid at the tab stop positions,
'however this function is by no means perfect but it will serve the purpose it is meant for,
'namely as a guide for indentation.
'
'To save time and resources re-drawing the raster is deferred until the end of a smooth mouse
'sroll, so it seems a bit 'sluggish'
'
'Known quirks:
'1  The VB tabwidth, font and indicatorbar may be altered while this AddIn is active, but this
'   will not be noticed until next startup or click on Refresh/Reset in fSetOpts.
'
'2  The Code Window Horizontal Scrollbar is a very stubborn thing indeed >:-(
'
'   Apparently the Scrollbar does not belong to the Code Window. So instead of using SB_HORZ
'   we have to wait for the first Scroll Msg and then we know an hWndScrollbar and can ask the
'   scrollbar for it's value via SB_CTL. The effect is that switching from one code pane to
'   another may result in the raster being off the correct position until the horiz scroll bar
'   is used again - click on thumb to redraw correct raster.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'15Jan2004 Version 2.6.3    UMG
'
'Added DeleteObject for Region (may have been a memory leak)
'Modified repositioning
'Made Flashing Region a Rounded tRECT
'Some code cosmetics
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'18Dec2003 Version 2.5.7    UMG
'
'Made Return to Cursor (mouse middle button) and Selection Box (pause key)
'independent of AutoComplete option.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'28Dec2002 Version 2.4.3    UMG
'
'Added custom tooltip class.
'Added BitBlt Icon into VB's MDI window.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'21Dec2002 Version 2.3.6    UMG
'
'Replaced WM_SETFOCUS by WM_MOUSEACTIVATE = &H21
'Altered LastTop-Algorithm.
'Changed App.Title, so the entries in the Registry are no longer found, you will
'have to re-enter and save again. Also delete the previous Registry keys.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'11Dec2002 Version 2.3.3    UMG
'
'Changed Wheel Scrolling Algorithm so as to accumulate wheel events while scolling.
'Scroll speed increases with distance.
'Added High Resolution Timer functions.
'Added Scroll Speed Option - hover mouse over "Speed" label in Selection Box.
'Added RightButtonClick for exiting the Selection Box.
'Added Middle- or Wheel-Click to return to caret position. The circle drawn has
'    the correct size for FixedSys font, other fonts may need some adjustment
'    (find CreateEllipticRgn (near line 202 in mSubclass) to modify)
'
'Fixed Bug with RegSave Trigger Length.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'19Nov2002 Version 2.2.11   UMG
'
'Fixed or rather circumvented 'Paused-bug'; I don't know what's wrong there but
'AutoComplete is now suspended while we're in Paused mode and that apparently cures
'it. If anybody has a clue I'd be pleased to know...
'
'Changed selection list to flexgrid
'you can now sort the selection list by column and use the keyboard to position it.
'
'#################################################################################'
'==========
'IMPORTANT:
'==========
'This was developed using a German version of VB, so you have to translate one item:
'Search for [Entwerfen] and translate that to what your Main Window Caption says in
'square brackets when you're in Design Mode'(probably [design] in english) keeping
'the case (UPPER/lower).
'
'(this is no longer true after V3.3.12)
'
'See line 6 of mSubclass.
'##################################################################################
'
'Some code optimization.
'
'You can now type a questionmark after an API name which fails to trigger auto-
'completion (see ReadMe.txt)
'
'    for example
'    Private ApiDeclare SendMessage?
'
'    ...to distingish it from SendMessageCallback and SendMessageTimeout.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'13Oct2002 Version 2.2.8    UMG
'
'Added API database
'
'Use keywords:    ApiDeclare instead of  Declare Sub/Function
'                 ApiConst   instead of  Const
'                 ApiType    instead of  Type
'
' ...and then type a space and the name of the API member you wish to autocomplete.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'08Oct2002 Version 2.1.15   UMG
'
'Added fSelect - press Pause button to open member list
'    known quirk: the selection may be off a few bytes after the list is closed.
'                 (fixed - I hope)
'
'Added check for MDI / SDI.
'
'Some code optimization.
'
'Fixed quirk with space or newline input before a part keyword or a name.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'02Oct2002 Version 2.1.12   UMG
'
'Added IDE Codepane Auto Complete Function.
'
'Added fSetOpts options for Auto Complete.
'
'Modified subclassing.
'
'Fixed bug regarding scrolling an empty code pane.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'08Sep2002 Version 1.2.4    UMG
'
'Get scroll options from Registry, or from our own settings; compile time option
'no longer exists.
'
'Hold down left mouse button and rotate wheel to open the Scroll Settings Dialog Box.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'06Sep2002 Version 1.1.5    UMG
'
'"Exit" bug fixed - wasn't an Exit bug really: this happened when the user tried to
'scroll AND all codepanes were closed AND at least one codepane had been open before.
'       ===                           ===
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'05Sep2002 Version 1.1.4    UMG
'
'Scrolling method changed - no sending keystrokes anymore.
'
'You can now slow down scrolling by factors 2, 3 or 4 by holding down the Shift key,
'the Cntl key or both respectively, while scrolling the mouse wheel.
'
'You also have the choice between two alternative scrolling modes at compile time
'by altering a single #Const.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'05Sep2002 Version 1.0.1    UMG
'
'Now has a "fraction of page to scroll" - constant, currently set to 1/2, modify that
'as you like. If you feel like storing/getting this value from/in Settings: the only
'limit is your imagination.
'
'A little code cosmetic and plenty of comments.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'04Sep2002 Version 1.0.0    UMG
'
'Prototype
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Get temp file name for saving the current source
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private WithEvents ResetMenuButtonEvent     As CommandBarEvents
Attribute ResetMenuButtonEvent.VB_VarHelpID = -1
Private WithEvents OpenAllMenuButtonEvent   As CommandBarEvents
Attribute OpenAllMenuButtonEvent.VB_VarHelpID = -1
Private WithEvents CompareMenuButtonEvent   As CommandBarEvents
Attribute CompareMenuButtonEvent.VB_VarHelpID = -1
Private WithEvents CopyMenuButtonEvent      As CommandBarEvents
Attribute CopyMenuButtonEvent.VB_VarHelpID = -1

Private CommandBarMenu  As CommandBar
Private Proj            As VBProject
Private Compo           As VBComponent
Private OrigFilenames() As String
Private TempFilenames() As String
Private i               As Long
Private CaptText        As String
Private Const NoComp    As String = "Cannot see any active component. You must open a code pane first."
Private Const Backslash As String = "\"

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)

  'tidy up: close database, remove buttons, deactivate button events and unload forms

    If VBInstance.DisplayModel = vbext_dm_MDI Then

        'the user is unloading/disconnecting the addin...

        'so destroy the timer...
        DirtyTimer DestroyIt

        '...unhook any hooked window
        UnhookMDIClient

        On Error Resume Next

            '...close the database
            Set SlaveSet = Nothing
            Set MasterSet = Nothing
            ApiDatabase.Close
            Set ApiDatabase = Nothing

            '...delete my menu buttons
            ResetMenuButton.Delete
            Set ResetMenuButtonEvent = Nothing
            Set ResetMenuButton = Nothing

            OpenAllMenuButton.Delete
            Set OpenAllMenuButtonEvent = Nothing
            Set OpenAllMenuButton = Nothing

            CompareMenuButton.Delete
            Set CompareMenuButtonEvent = Nothing
            Set CompareMenuButton = Nothing

            CopyMenuButton.Delete
            Set CopyMenuButtonEvent = Nothing
            Set CopyMenuButton = Nothing

        On Error GoTo 0

        '...unlog the icons
        Unload fIcon

        '...and finally kill the compare window if it is open
        If hWndCompare Then
            SendMessage hWndCompare, WM_CLOSE, 0&, ByVal 0&
            hWndCompare = 0
        End If

    End If

End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

  Dim e         As Long

    Set VBInstance = Application 'save object variable pointing to the VB application instance
    With VBInstance
        If .DisplayModel = vbext_dm_MDI Then
            fSplash.Show
            DoEvents
            On Error Resume Next
                Set CommandBarMenu = .CommandBars(1)
            On Error GoTo 0
            If CommandBarMenu Is Nothing Then
                MsgBox "Companion Add-In was loaded but could not be connected to the main VB Menu.", vbCritical
              Else 'NOT COMMANDBARMENU...
                With CommandBarMenu

                    Set OpenAllMenuButton = .Controls.Add(msoControlButton)
                    With OpenAllMenuButton
                        .BeginGroup = True
                        .Caption = "&OpenAll"
                        .State = msoButtonUp
                        .Style = msoButtonIconAndCaption
                        .ToolTipText = "Open all code panes"
                        SetMenuIcon OpenAllMenuButton, fIcon.picMenuOpenAll
                    End With 'OPENALLMENUBUTTON

                    Set CompareMenuButton = .Controls.Add(msoControlButton)
                    With CompareMenuButton
                        .BeginGroup = True
                        .Caption = "&Compare"
                        .State = msoButtonUp
                        .Style = msoButtonIconAndCaption
                        .ToolTipText = NAC
                        SetMenuIcon CompareMenuButton, fIcon.picMenuCompare
                    End With 'COMPAREMENUBUTTON

                    Set CopyMenuButton = .Controls.Add(msoControlButton)
                    With CopyMenuButton
                        .BeginGroup = True
                        .Caption = "Cop&y"
                        .State = msoButtonUp
                        .Style = msoButtonIconAndCaption
                        .ToolTipText = NAC
                        SetMenuIcon CopyMenuButton, fIcon.picMenuCopy
                    End With 'RESETMENUBUTTON 'COPYMENUBUTTON

                    Set ResetMenuButton = .Controls.Add(msoControlButton)
                    With ResetMenuButton
                        .BeginGroup = True
                        .Caption = "R&eset"
                        .State = msoButtonUp
                        .Style = msoButtonIconAndCaption
                        .ToolTipText = NAC
                        SetMenuIcon ResetMenuButton, fIcon.picMenuResetRed
                    End With 'RESETMENUBUTTON

                End With 'COMMANDBARMENU

                With .Events
                    Set OpenAllMenuButtonEvent = .CommandBarEvents(OpenAllMenuButton) 'hook events for this menu button
                    Set CompareMenuButtonEvent = .CommandBarEvents(CompareMenuButton) 'hook events for this menu button
                    Set CopyMenuButtonEvent = .CommandBarEvents(CopyMenuButton) 'hook events for this menu button
                    Set ResetMenuButtonEvent = .CommandBarEvents(ResetMenuButton) 'hook events for this menu button
                End With '.EVENTS

            End If
            DoEvents '.COUNTOFVISIBLELINES = FALSE/0
            Sleep 555
            NumTexts = Array("No", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve")
            'get VB TabWidth and Registry settings
            GetRegistrySettings
            'check API database
            If ApiDBFileName <> sUnknown Then 'either we have a name or we ask the user
                Do 'try to open database
                    On Error Resume Next
                        Set ApiDatabase = DBEngine.Workspaces(0).OpenDatabase(ApiDBFileName, , True, vbNullString)
                        e = Err
                        If e Then 'couldnt open Database so we ask the user
                            Err.Clear
                            With fSplash.cdlDB 'prepare common dialog
                                .Flags = cdlOFNLongNames Or cdlOFNFileMustExist Or cdlOFNReadOnly
                                .ShowOpen
                                If Err Then 'user clicked cancel in common dialog popup
                                    ApiDBFileName = sUnknown 'so that we will not ask again
                                  Else 'got a file name 'ERR = FALSE/0
                                    ApiDBFileName = .FileName
                                End If
                                SaveSetting App.Title, sOptions, sApiLocation, ApiDBFileName
                            End With 'FSPLASH.CDLDB
                          Else 'database was opened successfully 'E = FALSE/0
                            If ApiDatabase.TableDefs(1).Name <> sDeclares Then 'see if db is the correct one
                                ApiDatabase.Close 'wrong db so close it
                                ApiDBFileName = sAskUser 'reset wrong name
                                DoEvents
                                Beeper 440, 20
                                Sleep 333
                                e = 1 'try again
                            End If
                        End If
                    On Error GoTo 0
                Loop Until e = 0 Or ApiDBFileName = sUnknown 'no error or no file name - exit
            End If
            LoadKeywords IIf(NonUnique, TriggerLength, 2)
            sHalfPage = Replace$(opHpCapt, "&", vbNullString)
            Load fIcon
            With fIcon.picIcon
                hDCIcon = .hDC
                wIcon = .Width / 2
                hIcon = .Height
            End With 'FICON.PICICON
            'get text constant from VB resource dll
            InDesignMode = "[" & GetResourceString("vb6ide.dll", 13137) & "]"
            'create x-pointer vertices
            XPointerVertices(1) = MakePoint(-11, -11)   'top left           |\  /|
            XPointerVertices(2) = MakePoint(11, 11)     'bottom right       | \/ |
            XPointerVertices(3) = MakePoint(11, -11)    'top right          | /\ |
            XPointerVertices(4) = MakePoint(-11, 11)    'bottom left        |/  \|
            'create arrow vertices
            ArrowVertices(1) = MakePoint(38, -4)        'shaft top right
            ArrowVertices(2) = MakePoint(10, -4)        'shaft top left       /|
            ArrowVertices(3) = MakePoint(10, -10)       'wedge top           / +------+
            ArrowVertices(4) = MakePoint(0, 0)          'wedge point        <         |
            ArrowVertices(5) = MakePoint(10, 10)        'wedge bottom        \ +------+
            ArrowVertices(6) = MakePoint(10, 5)         'shaft bottom left    \|
            ArrowVertices(7) = MakePoint(38, 5)         'shaft bottom right
            'check high resolution timer present
            HiResTimerPresent = 0
            On Error Resume Next
                HiResTimerPresent = QueryPerformanceFrequency(CPUFreq)
            On Error GoTo 0
            Unload fSplash
        End If
    End With 'VBINSTANCE

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    AddinInstance_OnBeginShutdown custom() 'disconnect is similar to shutdown
    IdleStopDetection

End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)

    If VBInstance.DisplayModel = vbext_dm_MDI Then
        HookMDIClient 'hook the MDI client window which in turn hooks the active code pane window, if present
        IdleBeginDetection
    End If

End Sub

Private Sub CompareMenuButtonEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

  Dim fc    As Long

    With VBInstance
        If .CodePanes.Count Then
            CaptText = App.ProductName & " - " & .ActiveVBProject.Name & " [" & .SelectedVBComponent.Name & "]"
            With .SelectedVBComponent
                If Len(.FileNames(1)) = 0 Then
                    MsgBoxEx "Component " & .Name & " is new; there is no previous state to compare it with.", vbInformation, CaptText, PosX:=-2, PosY:=-2
                  Else 'NOT LEN(.FILENAMES(1))...
                    If .IsDirty Then
                        If hWndCompare Then
                            SendMessage hWndCompare, WM_CLOSE, 0&, ByVal 0&
                            hWndCompare = 0
                        End If
                        fCompare.Caption = CaptText
                        fc = .FileCount
                        ReDim OrigFilenames(1 To fc), TempFilenames(1 To fc)
                        For i = 1 To fc
                            OrigFilenames(i) = .FileNames(i)
                            TempFilenames(i) = String$(255, 0)
                            GetTempFileName Left$(OrigFilenames(i), InStrRev(OrigFilenames(i), Backslash)), "UMG", 0, TempFilenames(i)
                            TempFilenames(i) = Left$(TempFilenames(i), InStr(TempFilenames(i), Chr$(0)) - 1)
                            FileCopy OrigFilenames(i), TempFilenames(i)
                        Next i
                        .SaveAs OrigFilenames(1) 'may alter the .frx file (VB bug?)
                        fCompare.LastDeclLine = .CodeModule.Lines(.CodeModule.CountOfDeclarationLines, 1)
                        fCompare.Compare TempFilenames(1), OrigFilenames(1)
                        Set fCompare = Nothing
                        For i = 1 To fc
                            FileCopy TempFilenames(i), OrigFilenames(i)
                            Kill TempFilenames(i)
                        Next i
                        .IsDirty = True 'since the user my have inserted spaces or empty lines we cannot reset IsDirty once it is set, even though the comparator may think that there are no changes
                      Else '.ISDIRTY = FALSE/0
                        MsgBoxEx "No changes were found, " & .Name & " is in its initial state.", vbInformation, "No need to compare " & .Name, PosX:=-2, PosY:=-2, OCapt:=OK, NCapt:="Okay"
                    End If
                End If
            End With '.SELECTEDVBCOMPONENT
          Else '.CODEPANES.COUNT = FALSE/0
            MsgBox NoComp, vbCritical
        End If
    End With 'VBINSTANCE

End Sub

Private Sub CopyMenuButtonEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

  Dim TopLn     As Long
  Dim WS        As Long
  Dim Top       As Long
  Dim Lft       As Long
  Dim Rgt       As Long
  Dim Bot       As Long

    With VBInstance
        If .CodePanes.Count Then
            With .ActiveCodePane
                WS = .Window.WindowState
                TopLn = .TopLine
                .CodeModule.CodePane.GetSelection Top, Lft, Bot, Rgt
                If Lft = 1 Then
                    RegionCenterY = 0
                  Else 'NOT LFT...
                    RegionCenterY = LineHeight
                End If
                If Top = Bot And Lft = Rgt Then
                    MoveToArrow fCopy, Top, Lft, Bot, Rgt
                    With fCopy
                        .TextToPaste = vbNullString
                        .sFontName = IDEFontName
                        .lFontSize = IDEFontSize
                        .Show vbModal
                    End With 'FCOPY
                    CaretRgn Arrow Or Inval 'invalidate additional arrow region before hiding Me
                    If Len(fCopy.TextToPaste) Then
                        .CodeModule.InsertLines Bot - (Lft <> 1), fCopy.TextToPaste
                      Else 'LEN(FCOPY.TEXTTOPASTE) = FALSE/0
                        .Window.WindowState = WS
                        .TopLine = TopLn
                    End If
                    Unload fCopy
                  Else 'NOT TOP...
                    MsgBoxEx Replace$(UndefinedCaret, "$", "Copy Facility"), vbInformation, PosX:=-2, OffsetY:=-40, OCapt:=OK, NCapt:="&Close"
                End If
                .Window.SetFocus 'reset focus
            End With '.ACTIVECODEPANE
          Else '.CODEPANES.COUNT = FALSE/0
            MsgBox NoComp, vbCritical
        End If
    End With 'VBINSTANCE

End Sub

Private Sub OpenAllMenuButtonEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

  Dim ActivePane    As CodePane
  Dim WS            As Long
  Dim NumModules    As Long

    With VBInstance
        For Each Proj In .VBProjects 'count components is this (these) project(s)
            For Each Compo In Proj.VBComponents
                NumModules = NumModules - (Compo.Type <> vbext_ct_RelatedDocument And Compo.Type <> vbext_ct_ResFile)
        Next Compo, Proj
        Select Case True
          Case .VBProjects.Count = 0
            MsgBox "Cannot see any project. You must open a project first.", vbCritical
          Case NumModules = 0
            MsgBoxEx "Project " & .ActiveVBProject.Name & " has no codepanes.", vbInformation, PosX:=-2, PosY:=-2, OCapt:=OK, NCapt:="Okay"
          Case .CodePanes.Count = NumModules
            MsgBoxEx GetNumText(NumModules) & " codepane" & IIf(NumModules = 1, " is", "s are") & " already open.", vbInformation, PosX:=-2, PosY:=-2, OCapt:=OK, NCapt:="Okay"
          Case Else
            Set ActivePane = .ActiveCodePane
            If Not ActivePane Is Nothing Then
                WS = ActivePane.Window.WindowState
            End If
            For Each Proj In .VBProjects
                With Proj
                    For Each Compo In .VBComponents
                        With Compo
                            If .Type <> vbext_ct_ResFile And .Type <> vbext_ct_RelatedDocument Then
                                With .CodeModule.CodePane
                                    If Not .Window.Visible Then
                                        .Show
                                        DoEvents
                                    End If
                                End With '.CODEMODULE.CODEPANE
                            End If
                        End With 'COMPO
                    Next Compo
                End With 'PROJ
            Next Proj
            SendKeys "%W", True 'Tile windows
            If Not ActivePane Is Nothing Then
                ActivePane.Show
                ActivePane.Window.WindowState = WS
            End If
        End Select
    End With 'VBINSTANCE

End Sub

Private Sub ResetMenuButtonEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

  Dim Reply   As Long

    With VBInstance
        If .CodePanes.Count = 0 Then
            MsgBox NoComp, vbCritical
          Else 'NOT .CODEPANES.COUNT...
            With ActiveCompo
                If .IsDirty Then
                    If DontAskAgain Then
                        Reply = vbYes
                      Else 'DONTASKAGAIN = FALSE/0
                        Reply = MsgBoxEx("You will lose ALL changes you have made." & vbCrLf & vbCrLf & "Are you sure you want to reset " & .Name & " to it's previous state?", vbQuestion Or vbYesNoCancel Or vbDefaultButton3, Title:="Reset " & .Name, PosX:=-2, PosY:=-2, TimeOut:=15000, OCapt:=Ja & "|" & Nein & "|" & Abbrechen, NCapt:="&Go ahead|Yes, &always|&No")
                    End If
                    On Error Resume Next
                        Select Case Reply
                          Case vbYes 'in fact this is the Go ahead button
                            .Reload
                            Reply = Err
                          Case vbNo 'this is the Yes allways button
                            .Reload
                            Reply = Err
                            DontAskAgain = True 'LEN(TABLENAME) = FALSE/0
                          Case Else 'no button
                            Reply = 0
                        End Select
                        If Reply Then
                            MsgBoxEx "Cannot find any previous state for " & .Name & ".", vbInformation, "Cannot reset " & .Name, PosX:=-2, PosY:=-2, OCapt:=OK, NCapt:="Okay"
                          Else 'REPLY = FALSE/0
                            If hWndCompare Then
                                SendMessage hWndCompare, WM_CLOSE, 0&, ByVal 0&
                                hWndCompare = 0
                            End If
                        End If
                    On Error GoTo 0
                  Else '.ISDIRTY = FALSE/0
                    MsgBoxEx "No changes were found; " & .Name & " is in its initial state.", vbInformation, "Cannot reset " & .Name, PosX:=-2, PosY:=-2, OCapt:=OK, NCapt:="Okay"
                End If
            End With 'ACTIVECOMPO
        End If
    End With 'VBINSTANCE

End Sub

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 16:14)  Decl: 410  Code: 399  Total: 809 Lines
':) CommentOnly: 406 (50,2%)  Commented: 74 (9,1%)  Empty: 53 (6,6%)  Max Logic Depth: 10
