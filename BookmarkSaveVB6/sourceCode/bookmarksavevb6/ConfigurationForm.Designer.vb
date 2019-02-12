<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ConfigurationForm
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
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Options")
        Dim TreeNode2 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("General", New System.Windows.Forms.TreeNode() {TreeNode1})
        Dim TreeNode3 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("About")
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ConfigurationForm))
        Me.splitMain = New System.Windows.Forms.SplitContainer()
        Me.treeOptions = New System.Windows.Forms.TreeView()
        Me.btnOK = New System.Windows.Forms.Button()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitMain.Panel1.SuspendLayout()
        Me.splitMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'splitMain
        '
        Me.splitMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.splitMain.Location = New System.Drawing.Point(12, 12)
        Me.splitMain.Name = "splitMain"
        '
        'splitMain.Panel1
        '
        Me.splitMain.Panel1.Controls.Add(Me.treeOptions)
        '
        'splitMain.Panel2
        '
        Me.splitMain.Panel2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.splitMain.Size = New System.Drawing.Size(481, 309)
        Me.splitMain.SplitterDistance = 108
        Me.splitMain.TabIndex = 0
        '
        'treeOptions
        '
        Me.treeOptions.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.treeOptions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treeOptions.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.treeOptions.FullRowSelect = True
        Me.treeOptions.Location = New System.Drawing.Point(0, 0)
        Me.treeOptions.Name = "treeOptions"
        TreeNode1.Name = "ndGeneralOptions"
        TreeNode1.Text = "Options"
        TreeNode2.Name = "ndGeneral"
        TreeNode2.Text = "General"
        TreeNode3.Name = "ndAbout"
        TreeNode3.Text = "About"
        Me.treeOptions.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode2, TreeNode3})
        Me.treeOptions.ShowLines = False
        Me.treeOptions.ShowPlusMinus = False
        Me.treeOptions.Size = New System.Drawing.Size(108, 309)
        Me.treeOptions.TabIndex = 0
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(405, 327)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(88, 26)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'ConfigurationForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(505, 358)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.splitMain)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ConfigurationForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ConfigurationForm"
        Me.splitMain.Panel1.ResumeLayout(False)
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitMain.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents treeOptions As System.Windows.Forms.TreeView
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
