'---------------------------------------------------------------------
'
'Options control for the configuration dialog for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------


Public Class OptionsControl

    Private Sub btnShowLogger_Click(sender As System.Object, e As System.EventArgs) Handles btnShowLogger.Click
        Logger.Start()
    End Sub
End Class
