Imports System.Collections
Imports System.Windows.Forms


''' <summary>
''' A simple control that allows the user to select pretty much any valid hotkey combination
''' Originally from serenity@exscape.org, 2006-08-03
''' </summary>
Public Class HotkeyControl
    Inherits TextBox

    ' These variables store the current hotkey and modifier(s)
    Private _hotkey As Keys = Keys.None
    Private _modifiers As Keys = Keys.None
    ' ArrayLists used to enforce the use of proper modifiers.
    ' Shift+A isn't a valid hotkey, for instance, as it would screw up when the user is typing.
    Private needNonShiftModifier As ArrayList = Nothing
    Private needNonAltGrModifier As ArrayList = Nothing
    Private dummy As New ContextMenu()


    ''' <summary>
    ''' Used to make sure that there is no right-click menu available
    ''' </summary>
    Public Overrides Property ContextMenu() As ContextMenu
        Get
            Return dummy
        End Get
        Set(value As ContextMenu)
            MyBase.ContextMenu = dummy
        End Set
    End Property


    ''' <summary>
    ''' Forces the control to be non-multiline
    ''' </summary>
    Public Overrides Property Multiline() As Boolean
        Get
            Return MyBase.Multiline
        End Get
        Set(value As Boolean)
            ' Ignore what the user wants; force Multiline to false
            MyBase.Multiline = False
        End Set
    End Property


    ''' <summary>
    ''' Creates a new HotkeyControl
    ''' </summary>
    Public Sub New()
        Me.ContextMenu = dummy
        ' Disable right-clicking
        Me.Text = "None"
        ' Handle events that occurs when keys are pressed
        AddHandler Me.KeyPress, New KeyPressEventHandler(AddressOf HotkeyControl_KeyPress)
        AddHandler Me.KeyUp, New KeyEventHandler(AddressOf HotkeyControl_KeyUp)
        AddHandler Me.KeyDown, New KeyEventHandler(AddressOf HotkeyControl_KeyDown)
        ' Fill the ArrayLists that contain all invalid hotkey combinations
        needNonShiftModifier = New ArrayList()
        needNonAltGrModifier = New ArrayList()

        PopulateModifierLists()
    End Sub


    ''' <summary>
    ''' Populates the ArrayLists specifying disallowed hotkeys
    ''' such as Shift+A, Ctrl+Alt+4 (would produce a dollar sign) etc
    ''' </summary>
    Private Sub PopulateModifierLists()
        ' Shift + 0 - 9, A - Z
        For k As Keys = Keys.D0 To Keys.Z
            needNonShiftModifier.Add(CInt(k))
        Next
        ' Shift + Numpad keys
        For k As Keys = Keys.NumPad0 To Keys.NumPad9
            needNonShiftModifier.Add(CInt(k))
        Next
        ' Shift + Misc (,;<./ etc)
        For k As Keys = Keys.Oem1 To Keys.OemBackslash
            needNonShiftModifier.Add(CInt(k))
        Next
        ' Shift + Space, PgUp, PgDn, End, Home
        For k As Keys = Keys.Space To Keys.Home
            needNonShiftModifier.Add(CInt(k))
        Next
        ' Misc keys that we can't loop through
        needNonShiftModifier.Add(CInt(Keys.Insert))
        needNonShiftModifier.Add(CInt(Keys.Help))
        needNonShiftModifier.Add(CInt(Keys.Multiply))
        needNonShiftModifier.Add(CInt(Keys.Add))
        needNonShiftModifier.Add(CInt(Keys.Subtract))
        needNonShiftModifier.Add(CInt(Keys.Divide))
        needNonShiftModifier.Add(CInt(Keys.[Decimal]))
        needNonShiftModifier.Add(CInt(Keys.[Return]))
        needNonShiftModifier.Add(CInt(Keys.Escape))
        needNonShiftModifier.Add(CInt(Keys.NumLock))
        needNonShiftModifier.Add(CInt(Keys.Scroll))
        needNonShiftModifier.Add(CInt(Keys.Pause))
        ' Ctrl+Alt + 0 - 9
        For k As Keys = Keys.D0 To Keys.D9
            needNonAltGrModifier.Add(CInt(k))
        Next
    End Sub


    ''' <summary>
    ''' Resets this hotkey control to None
    ''' </summary>
    Public Shadows Sub Clear()
        Me.Hotkey = Keys.None
        Me.HotkeyModifiers = Keys.None
    End Sub


    ''' <summary>
    ''' Fires when a key is pushed down. Here, we'll want to update the text in the box
    ''' to notify the user what combination is currently pressed.
    ''' </summary>
    Private Sub HotkeyControl_KeyDown(sender As Object, e As KeyEventArgs)
        ' Clear the current hotkey
        ' ONLY if it's already been set as the back or delete
        ' ie, you can set the hotkey to back by pressing it once
        ' but if you press it again, it's cleared.
        If (e.KeyCode = Keys.Back AndAlso Me.Hotkey = e.KeyCode) OrElse (e.KeyCode = Keys.Delete AndAlso Me.Hotkey = e.KeyCode) Then
            ResetHotkey()
            Return
        Else
            Me._modifiers = e.Modifiers
            Me._hotkey = e.KeyCode

            Redraw()
        End If
    End Sub


    ''' <summary>
    ''' Fires when all keys are released. If the current hotkey isn't valid, reset it.
    ''' Otherwise, do nothing and keep the text and hotkey as it was.
    ''' </summary>
    Private Sub HotkeyControl_KeyUp(sender As Object, e As KeyEventArgs)
        If Me._hotkey = Keys.None AndAlso Control.ModifierKeys = Keys.None Then
            ResetHotkey()
            Return
        End If
    End Sub


    ''' <summary>
    ''' Prevents the letter/whatever entered to show up in the TextBox
    ''' Without this, a "A" key press would appear as "aControl, Alt + A"
    ''' </summary>
    Private Sub HotkeyControl_KeyPress(sender As Object, e As KeyPressEventArgs)
        e.Handled = True
    End Sub


    ''' <summary>
    ''' Handles some misc keys, such as Ctrl+Delete and Shift+Insert
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        If keyData = Keys.Delete OrElse keyData = (Keys.Control Or Keys.Delete) Then
            ResetHotkey()
            Return True
        End If
        If keyData = (Keys.Shift Or Keys.Insert) Then
            ' Paste
            Return True
        End If
        ' Don't allow
        ' Allow the rest
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function


    ''' <summary>
    ''' Clears the current hotkey and resets the TextBox
    ''' </summary>
    Public Sub ResetHotkey()
        Me._hotkey = Keys.None
        Me._modifiers = Keys.None
        Redraw()
    End Sub


    ''' <summary>
    ''' Used to get/set the hotkey (e.g. Keys.A)
    ''' </summary>
    Public Property Hotkey() As Keys
        Get
            Return Me._hotkey
        End Get
        Set(value As Keys)
            Me._hotkey = value
            Redraw(True)
        End Set
    End Property


    ''' <summary>
    ''' Used to get/set the modifier keys (e.g. Keys.Alt | Keys.Control)
    ''' </summary>
    Public Property HotkeyModifiers() As Keys
        Get
            Return Me._modifiers
        End Get
        Set(value As Keys)
            Me._modifiers = value
            Redraw(True)
        End Set
    End Property


    ''' <summary>
    ''' Helper function
    ''' </summary>
    Private Sub Redraw()
        Redraw(False)
    End Sub


    ''' <summary>
    ''' Redraws the TextBox when necessary.
    ''' </summary>
    ''' <param name="bCalledProgramatically">Specifies whether this function was called by the Hotkey/HotkeyModifiers properties or by the user.</param>
    Private Sub Redraw(bCalledProgramatically As Boolean)
        ' No hotkey set
        If Me._hotkey = Keys.None Then
            Me.Text = "None"
            Return
        End If
        ' LWin/RWin doesn't work as hotkeys (neither do they work as modifier keys in .NET 2.0)
        If Me._hotkey = Keys.LWin OrElse Me._hotkey = Keys.RWin Then
            Me.Text = "None"
            Return
        End If
        ' Only validate input if it comes from the user
        If bCalledProgramatically = False Then
            ' No modifier or shift only, AND a hotkey that needs another modifier
            If (Me._modifiers = Keys.Shift OrElse Me._modifiers = Keys.None) AndAlso Me.needNonShiftModifier.Contains(CInt(Me._hotkey)) Then
                If Me._modifiers = Keys.None Then
                    ' Set Ctrl+Alt as the modifier unless Ctrl+Alt+<key> won't work...
                    If needNonAltGrModifier.Contains(CInt(Me._hotkey)) = False Then
                        Me._modifiers = Keys.Alt Or Keys.Control
                    Else
                        ' ... in that case, use Shift+Alt instead.
                        Me._modifiers = Keys.Alt Or Keys.Shift
                    End If
                Else

                    ' User pressed Shift and an invalid key (e.g. a letter or a number),
                    ' that needs another set of modifier keys
                    Me._hotkey = Keys.None
                    Me.Text = Me._modifiers.ToString() & " + Invalid key"

                    Return
                End If
            End If
            ' Check all Ctrl+Alt keys

            If (Me._modifiers = (Keys.Alt Or Keys.Control)) AndAlso Me.needNonAltGrModifier.Contains(CInt(Me._hotkey)) Then

                ' Ctrl+Alt+4 etc won't work; reset hotkey and tell the user
                Me._hotkey = Keys.None
                Me.Text = Me._modifiers.ToString() & " + Invalid key"

                Return
            End If
        End If
        If Me._modifiers = Keys.None Then
            If Me._hotkey = Keys.None Then
                Me.Text = "None"
                Return
            Else
                ' We get here if we've got a hotkey that is valid without a modifier,
                ' like F1-F12, Media-keys etc.
                Me.Text = Me._hotkey.ToString()

                Return
            End If
        End If
        ' I have no idea why this is needed, but it is. Without this code, pressing only Ctrl
        ' will show up as "Control + ControlKey", etc.
        ' Alt 
        If Me._hotkey = Keys.Menu OrElse Me._hotkey = Keys.ShiftKey OrElse Me._hotkey = Keys.ControlKey Then
            Me._hotkey = Keys.None
        End If
        Me.Text = Me._modifiers.ToString() & " + " & Me._hotkey.ToString()
    End Sub
End Class
