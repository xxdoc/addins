Option Explicit On
Imports System.IO
'---------------------------------------------------------------------
'
'Breakpoints class for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------

''' <summary>
''' Represents a collection of persistent breakpoints
''' </summary>
''' <remarks></remarks>
<ComVisible(False)>
Public Class Breakpoints
    Inherits List(Of Breakpoint)

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


    Public Overloads Function Add(ModuleName As String, LineNumber As Long) As Breakpoint
        'create a new object
        Dim bp = New Breakpoint

        'set the properties passed into the method
        bp.ModuleName = ModuleName
        bp.LineNumber = LineNumber
        bp.Parent = Me
        Logger.Log("Adding BP at " & ModuleName & ":" & LineNumber.ToString)
        MyBase.Add(bp)
        Me.IsDirty = True

        'return the object created
        Return bp
    End Function


    Public Overloads Function Find(ModuleName As String, LineNumber As Integer) As Breakpoint
        Return (From b In Me Where b.ModuleName = ModuleName And b.LineNumber = LineNumber Select b).FirstOrDefault
    End Function


    ''' <summary>
    ''' Save to default project file
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Save()
        Me.Save(Me.DefaultFile)
    End Sub


    Public Sub Save(ByVal FileName As String)
        Try
            If FileName IsNot Nothing Then
                Dim Buf = Serialize(Of Breakpoints)(Me)
                My.Computer.FileSystem.WriteAllText(FileName, Buf, False)
            End If
        Catch ex As Exception
            ex.Show("Unable to Save Breakpoints")
        End Try
        Me.IsDirty = False
    End Sub


    ''' <summary>
    ''' Load from Default project file
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Load()
        Me.Load(Me.DefaultFile)
    End Sub


    Public Sub Load(FileName As String)
        Dim Entries = New Breakpoints(Nothing)

        Try
            If My.Computer.FileSystem.FileExists(FileName) Then
                Dim File = My.Computer.FileSystem.ReadAllText(FileName)
                Entries = Deserialize(Of Breakpoints)(File)
                Me.Clear()
                For Each entry In Entries
                    entry.Parent = Me
                    Me.Add(entry)
                Next
                Me.IsDirty = False
            Else
                Me.Clear()
            End If
        Catch ex As Exception
            ex.Show("Unable to load Breakpoints")
        End Try
    End Sub


    Public ReadOnly Property DefaultFile() As String
        Get
            If Parent Is Nothing OrElse String.IsNullOrEmpty(Parent.FileName) Then Return Nothing
            Dim pth = Path.GetDirectoryName(Parent.FileName)
            Return Path.Combine(pth, Path.GetFileNameWithoutExtension(Parent.FileName) & ".bp")
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
        MyBase.Clear()
        Me.IsDirty = False
    End Sub
End Class
