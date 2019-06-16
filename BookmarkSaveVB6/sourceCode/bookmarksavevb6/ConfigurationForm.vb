Imports System.Windows.Forms
'---------------------------------------------------------------------
'
'Simple config form for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------


''' <summary>
''' The configuration form (there's no really much to it other than the about box)
''' </summary>
''' <remarks></remarks>
Public Class ConfigurationForm

    Private Sub ConfigurationForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '---- every node needs a control associated with it
        treeOptions.Nodes("ndGeneral").Tag = New OptionsControl
        treeOptions.Nodes("ndGeneral").Nodes("ndGeneralOptions").Tag = treeOptions.Nodes("ndGeneral").Tag
        treeOptions.Nodes("ndAbout").Tag = New AboutBoxControl

        treeOptions.ExpandAll()
    End Sub


    Private Sub treeOptions_AfterSelect(sender As Object, e As System.Windows.Forms.TreeViewEventArgs) Handles treeOptions.AfterSelect
        If e.Node Is Nothing Then Exit Sub

        splitMain.Panel2.Controls.Clear()
        Dim ctrl As UserControl = e.Node.Tag
        splitMain.Panel2.Controls.Add(ctrl)
        ctrl.Dock = DockStyle.Fill
    End Sub


    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub
End Class