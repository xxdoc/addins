'---------------------------------------------------------------------
'
'Bookmark class BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------

''' <summary>
''' Represents a saved persistent bookmark
''' </summary>
''' <remarks></remarks>
<DataContract()> _
Public Class Bookmark
    Inherits CodeLocation


    Public Property Parent As Bookmarks
        Get
            Return _Parent
        End Get
        Set(value As Bookmarks)
            _Parent = value
        End Set
    End Property
    Private _Parent As Bookmarks


    Public Overrides Sub SetDirty()
        If Me.Parent IsNot Nothing Then Me.Parent.IsDirty = True
    End Sub
End Class
