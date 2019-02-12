
''' <summary>
''' Simple winapi subclassing Class
''' Allows for preproc and postproc subclassing, which is handy!
''' </summary>
''' <remarks></remarks>
Public Class Subclasser
    Implements IDisposable

    ' Prototype delegate to handle callbacks from subclassed control
    Private Delegate Function SubclassCallback(ByVal hWnd As IntPtr, ByVal Msg As IntPtr, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

    ' Subclasses a control with a callback
    <DllImport("User32", EntryPoint:="SetWindowLongA", CharSet:=CharSet.Unicode)> _
    Private Shared Function SetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As SubclassCallback) As IntPtr
    End Function

    ' Unsubclasses a control
    <DllImport("User32", EntryPoint:="SetWindowLongA", CharSet:=CharSet.Unicode)> _
    Private Shared Function SetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As IntPtr) As IntPtr
    End Function

    ' Forwards a message to a subclassed control
    <DllImport("User32", EntryPoint:="CallWindowProcA", CharSet:=CharSet.Unicode)> _
    Private Shared Function CallWindowProc(ByVal lpPrevWndFunc As IntPtr, ByVal hWnd As IntPtr, ByVal Msg As IntPtr, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function


    Private GWL_WNDPROC As IntPtr = New IntPtr((-4))
    Private prevHandle As IntPtr
    Private myCallback As New SubclassCallback(AddressOf Me.SubclassCallbackHandler)
    Private _hwnd As Integer
    Private _prevHandle As Integer

    '---- use a hashset for fastest lookup speed
    Private _msgs As HashSet(Of Win32.Messages) = New HashSet(Of Win32.Messages)


    Public Event WndProc(Sender As Object, e As WndProcEventArgs)
    Public Event PostWndProc(Sender As Object, e As WndProcEventArgs)


    Public Sub PostProcSubclass(ByVal hwnd As Integer)
        PostProc = True
        Me.Subclass(hwnd)
    End Sub


    Public Sub PostProcSubclass(ByVal hwnd As Integer, msgs() As Win32.Messages)
        PostProc = True
        Me.Subclass(hwnd, msgs)
    End Sub


    Public Sub Subclass(ByVal hwnd As Integer)
        ' Subclass control by setting the new window proc but also remember was it was originally
        ClearSubClassing()
        _hwnd = hwnd
        _prevHandle = SetWindowLong(hwnd, GWL_WNDPROC, Me.myCallback)
    End Sub


    Public Sub Subclass(ByVal hwnd As Integer, msgs() As Win32.Messages)
        _msgs.Clear()
        For Each m In msgs
            _msgs.Add(m)
        Next
        Me.Subclass(hwnd)
    End Sub


    Public Property PostProc As Boolean
        Get
            Return _PostProc
        End Get
        Set(value As Boolean)
            _PostProc = value
        End Set
    End Property
    Private _PostProc As Boolean


    Public ReadOnly Property hwnd As IntPtr
        Get
            Return _hwnd
        End Get
    End Property


    ''' <summary>
    ''' Stop subclassing but don't clear out messages
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearSubClassing()
        ' Unsubclass control (use overloaded version)
        If _hwnd <> 0 Then
            SetWindowLong(_hwnd, GWL_WNDPROC, _prevHandle)
        End If
        _prevHandle = 0
        _hwnd = 0
    End Sub


    Public Sub [Stop]()
        ClearSubClassing()
        _msgs.Clear()
    End Sub


    ' Handles all the messages sent to original handle
    Private Function SubclassCallbackHandler(ByVal hWnd As IntPtr, ByVal Msg As IntPtr, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        ' Process whatever messages you want here
        Dim e = New WndProcEventArgs
        e.hWnd = hWnd
        e.lParam = lParam
        e.wParam = wParam
        e.Msg = Msg

        '---- in the case of a normal subclass, raise the event BEFORE
        '     calling on down the chain
        Dim bIntercept = _msgs.Count = 0 OrElse _msgs.Contains(Msg)
        If bIntercept Then
            RaiseEvent WndProc(Me, e)
        End If
        
        '---- call down the windproc chain
        Dim r = CallWindowProc(_prevHandle, e.hWnd, e.Msg, e.wParam, e.lParam)

        If bIntercept Then
            RaiseEvent PostWndProc(Me, e)
        End If
        Return r
    End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            Me.Stop()
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region


    ''' <summary>
    ''' Nested class that describes the arguments of a subclassing event
    ''' (either pre or post proc)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class WndProcEventArgs
        Inherits EventArgs

        Public hWnd As IntPtr
        Public Msg As IntPtr
        Public wParam As IntPtr
        Public lParam As IntPtr
        Public [Return] As IntPtr
    End Class
End Class


