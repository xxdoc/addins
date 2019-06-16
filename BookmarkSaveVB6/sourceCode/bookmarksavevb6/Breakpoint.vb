'---------------------------------------------------------------------
'
'Breakpoint class for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------

''' <summary>
''' Represents a saved Breakpoint
''' </summary>
''' <remarks></remarks>
<DataContract()> _
Public Class Breakpoint
    Inherits CodeLocation


    Public Property Parent As Breakpoints
        Get
            Return _Parent
        End Get
        Set(value As Breakpoints)
            _Parent = value
        End Set
    End Property
    Private _Parent As Breakpoints


    Public Overrides Sub SetDirty()
        If Me.Parent IsNot Nothing Then Me.Parent.IsDirty = True
    End Sub
End Class
