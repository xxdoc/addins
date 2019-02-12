<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OptionsControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnShowLogger = New System.Windows.Forms.Button()
        Me.lblBookmarkHotKey = New System.Windows.Forms.Label()
        Me.cboBookmarkToggle = New BookmarkSave.HotkeyControl()
        Me.SuspendLayout()
        '
        'btnShowLogger
        '
        Me.btnShowLogger.Location = New System.Drawing.Point(119, 38)
        Me.btnShowLogger.Name = "btnShowLogger"
        Me.btnShowLogger.Size = New System.Drawing.Size(87, 26)
        Me.btnShowLogger.TabIndex = 1
        Me.btnShowLogger.Text = "Show Log"
        Me.btnShowLogger.UseVisualStyleBackColor = True
        '
        'lblBookmarkHotKey
        '
        Me.lblBookmarkHotKey.AutoSize = True
        Me.lblBookmarkHotKey.Location = New System.Drawing.Point(3, 15)
        Me.lblBookmarkHotKey.Name = "lblBookmarkHotKey"
        Me.lblBookmarkHotKey.Size = New System.Drawing.Size(94, 13)
        Me.lblBookmarkHotKey.TabIndex = 3
        Me.lblBookmarkHotKey.Text = "Toggle Bookmark:"
        Me.lblBookmarkHotKey.Visible = False
        '
        'cboBookmarkToggle
        '
        Me.cboBookmarkToggle.Hotkey = System.Windows.Forms.Keys.None
        Me.cboBookmarkToggle.HotkeyModifiers = System.Windows.Forms.Keys.None
        Me.cboBookmarkToggle.Location = New System.Drawing.Point(119, 12)
        Me.cboBookmarkToggle.Name = "cboBookmarkToggle"
        Me.cboBookmarkToggle.Size = New System.Drawing.Size(130, 20)
        Me.cboBookmarkToggle.TabIndex = 2
        Me.cboBookmarkToggle.Text = "None"
        Me.cboBookmarkToggle.Visible = False
        '
        'OptionsControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.lblBookmarkHotKey)
        Me.Controls.Add(Me.cboBookmarkToggle)
        Me.Controls.Add(Me.btnShowLogger)
        Me.Name = "OptionsControl"
        Me.Size = New System.Drawing.Size(377, 339)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnShowLogger As System.Windows.Forms.Button
    Friend WithEvents cboBookmarkToggle As BookmarkSave.HotkeyControl
    Friend WithEvents lblBookmarkHotKey As System.Windows.Forms.Label

End Class
