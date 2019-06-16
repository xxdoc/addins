<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLogger
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.tbxLog = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'tbxLog
        '
        Me.tbxLog.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbxLog.Location = New System.Drawing.Point(0, 0)
        Me.tbxLog.Multiline = True
        Me.tbxLog.Name = "tbxLog"
        Me.tbxLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.tbxLog.Size = New System.Drawing.Size(456, 338)
        Me.tbxLog.TabIndex = 0
        '
        'frmLogger
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(456, 338)
        Me.Controls.Add(Me.tbxLog)
        Me.Name = "frmLogger"
        Me.Text = "frmLogger"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tbxLog As System.Windows.Forms.TextBox
End Class
