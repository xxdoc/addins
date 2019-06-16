'---------------------------------------------------------------------
'
'Main IDTExtensibility entry point for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------

''' <summary>
''' This is the barebones COM visible CONNECT object that has to be registered with
''' VB6 to act as an addin
''' All it does is forward the interface calls to the Application object
''' 
''' NOTE: the GUIDs used HAVE to be unique!
''' </summary>
''' <remarks></remarks>
<ComClass("96861C1E-73A0-46E2-9993-AE66D2BC6A91", "1AEA0235-959D-4424-8231-8EBB9B9C85FE")> _
<ProgId("BookmarkSave.Connect")> _
Public Class Connect
    Implements VBIDE.IDTExtensibility

    Private _App As Application

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements VBIDE.IDTExtensibility.OnAddInsUpdate
        _App.OnAddInsUpdate(custom)
    End Sub

    Public Sub OnConnection(VBInst As Object, ConnectMode As VBIDE.vbext_ConnectMode, AddInInst As VBIDE.AddIn, ByRef custom As System.Array) Implements VBIDE.IDTExtensibility.OnConnection
        Try
            If _App Is Nothing Then
                _App = New Application
            End If
            _App.OnConnection(VBInst, ConnectMode, custom)
        Catch ex As Exception
            ex.Show("Unable to start addin.")
        End Try
    End Sub

    Public Sub OnDisconnection(RemoveMode As VBIDE.vbext_DisconnectMode, ByRef custom As System.Array) Implements VBIDE.IDTExtensibility.OnDisconnection
        _App.OnDisconnect(RemoveMode, custom)
    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements VBIDE.IDTExtensibility.OnStartupComplete
        _App.OnStartupComplete(custom)
    End Sub
End Class