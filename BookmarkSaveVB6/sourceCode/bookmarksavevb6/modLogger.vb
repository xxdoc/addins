'---------------------------------------------------------------------
'
'Simple Debug logging methods for BookmarkSave project
'Why I didn't just use OutputDebugString? No idea anymore!
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------

Public Class Logger

    Private Shared _frmLogger As frmLogger

    ''' <summary>
    ''' Write message to Debug window (and to log window if it's open)
    ''' </summary>
    ''' <param name="Message"></param>
    ''' <remarks></remarks>
    Public Shared Sub Log(Message As String)
        Static LogVal As String = ""
        Debug.Print(Message)

        If Len(LogVal) > 20000 Then
            Dim p = InStr(5000, LogVal, vbCrLf)
            LogVal = Mid(LogVal, p + 2)
        End If
        LogVal = LogVal & vbCrLf & Message

        If _frmLogger IsNot Nothing Then
            _frmLogger.SetLog(LogVal)
        End If
    End Sub


    Public Shared Sub Start()
        If _frmLogger Is Nothing Then _frmLogger = New frmLogger
        _frmLogger.Show()
        '---- force to show the log contents
        Log("")
    End Sub


    Public Shared Sub [Stop]()
        If _frmLogger IsNot Nothing Then _frmLogger.Visible = False
        _frmLogger = Nothing
    End Sub
End Class
