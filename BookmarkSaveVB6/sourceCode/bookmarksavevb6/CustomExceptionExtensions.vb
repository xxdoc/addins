Imports System.Runtime.CompilerServices
Imports System.IO
'---------------------------------------------------------------------
'
'Exception Handling Extensions for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------


''' <summary>
''' Provides extended methods for the general Exception object
''' allowing easier logging and display of exception information
''' </summary>
''' <remarks></remarks>
Public Module CustomExceptionExtensions

    ''' <summary>
    ''' Logs an exception 
    ''' </summary>
    ''' <param name="ex"></param>
    ''' <remarks></remarks>
    <Extension()> _
    Public Sub Log(ByVal ex As Exception, Optional ByVal Message As String = "", Optional ByVal Details As String = "")
        Try
            Using Writer = My.Computer.FileSystem.OpenTextFileWriter(Path.Combine(My.Application.Info.DirectoryPath, "BookmarkSave.log"), True)
                Writer.WriteLine(String.Format("{0:dd/MM/yyyy HH:mm:ss}  -  {1} Msg: {2}   Details: {3}", Now(), ex.ToString, Message, Details))
                Writer.Flush()
            End Using
        Catch
            '---- just ignore any exceptions here
        End Try
    End Sub
End Module
