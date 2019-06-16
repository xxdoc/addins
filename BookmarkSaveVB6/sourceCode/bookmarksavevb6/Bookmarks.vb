Option Explicit On
Imports System.IO
'---------------------------------------------------------------------
'
'Bookmarks class for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------


''' <summary>
''' Represents a collection of persisten bookmarks
''' </summary>
''' <remarks></remarks>
<ComVisible(False)>
Public Class Bookmarks
    Inherits List(Of Bookmark)

    Public Property Parent As VBIDE.VBProject

    ''' <summary>
    ''' Must have a default constructor to make serializer happy
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Parent = Nothing
    End Sub


    Public Sub New(VBProject As VBIDE.VBProject)
        Parent = VBProject
    End Sub


    Public Overloads Function Add(ModuleName As String, LineNumber As Long) As Bookmark
        'create a new object
        Dim bm = New Bookmark

        'set the properties passed into the method
        bm.ModuleName = ModuleName
        bm.LineNumber = LineNumber
        bm.Parent = Me
        MyBase.Add(bm)

        'return the object created
        Return bm
    End Function


    Public Overloads Function Find(ModuleName As String, LineNumber As Integer) As Bookmark
        Return (From b In Me Where b.ModuleName = ModuleName And b.LineNumber = LineNumber Select b).FirstOrDefault
    End Function


    ''' <summary>
    ''' Save with default project filename
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Save()
        Me.Save(Me.DefaultFile)
    End Sub


    Public Sub Save(ByVal FileName As String)
        Try
            If FileName IsNot Nothing Then
                Logger.Log("Saving Bookmarks to " & FileName)
                Dim Buf = Serialize(Of Bookmarks)(Me)
                My.Computer.FileSystem.WriteAllText(FileName, Buf, False)
            End If

        Catch ex As Exception
            ex.Show("Unable to Save Bookmarks.")
        End Try
        Me.IsDirty = False
    End Sub


    ''' <summary>
    ''' Load default persisted bookmarks
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Load()
        Me.Load(Me.DefaultFile)
    End Sub


    Public Sub Load(ByVal FileName As String)
        Dim Entries = New Bookmarks(Nothing)

        Try
            If My.Computer.FileSystem.FileExists(FileName) Then
                Dim File = My.Computer.FileSystem.ReadAllText(FileName)
                Entries = Deserialize(Of Bookmarks)(File)
                Me.Clear()
                For Each entry In Entries
                    entry.Parent = Me
                    Me.Add(entry)
                Next
            Else
                Me.Clear()
            End If
        Catch ex As Exception
            ex.Show("Unable to Load Bookmarks")
        End Try
    End Sub


    Public ReadOnly Property DefaultFile() As String
        Get
            If Parent Is Nothing OrElse String.IsNullOrEmpty(Parent.FileName) Then Return Nothing
            Dim pth = Path.GetDirectoryName(Parent.FileName)
            Return Path.Combine(pth, Path.GetFileNameWithoutExtension(Parent.FileName) & ".bm")
        End Get
    End Property


    Public Property IsDirty() As Boolean
        Get
            Return bDirty
        End Get
        Set(value As Boolean)
            bDirty = value
        End Set
    End Property
    Private bDirty As Boolean


    Public Shadows Sub Clear()
        Logger.Log("Bookmarks Cleared")
        MyBase.Clear()
        Me.IsDirty = False
    End Sub
End Class
