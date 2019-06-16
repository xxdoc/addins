'---------------------------------------------------------------------
'
'Simple Debug logging form for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------


Public Class frmLogger

    Public Sub SetLog(FullLog As String)
        tbxLog.Text = FullLog
        tbxLog.SelectionStart = Len(FullLog) - 1
        tbxLog.ScrollToCaret()
    End Sub


    Private Sub frmLogger_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Logger.Stop()
    End Sub
End Class