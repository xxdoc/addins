Option Explicit On
'---------------------------------------------------------------------
'
'CodeLocation base class for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------

''' <summary>
''' Represents a generic location in code (both breakpoints and bookmarks are locations)
''' </summary>
''' <remarks></remarks>
<DataContract()> _
Public Class CodeLocation

    <DataMember()> _
    Public Property ModuleName() As String
        Get
            Return _ModuleName
        End Get
        Set(ByVal value As String)
            _ModuleName = value
            Me.SetDirty()
        End Set
    End Property
    Private _ModuleName As String


    <DataMember()> _
    Public Property LineNumber() As Integer
        Get
            Return _LineNumber
        End Get
        Set(ByVal value As Integer)
            _LineNumber = value
            Me.SetDirty()
        End Set
    End Property
    Private _LineNumber As Integer


    ''' <summary>
    ''' can contain the contents of the line, mainly used for checking if the 
    ''' line has been commented out, thus it couldn't have a breakpoint set to it
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LineContent() As String


    Public Overrides Function ToString() As String
        Return Me.ModuleName & ":" & Me.LineNumber
    End Function


    Public Overridable Sub SetDirty()
    End Sub
End Class
