<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        SplitContainer1 = New SplitContainer()
        tvFolders = New TreeView()
        TableLayoutPanel1 = New TableLayoutPanel()
        lvFiles = New ListView()
        Panel1 = New Panel()
        txtPath = New TextBox()
        btnForward = New Button()
        btnBack = New Button()
        btnGo = New Button()
        CType(SplitContainer1, ComponentModel.ISupportInitialize).BeginInit()
        SplitContainer1.Panel1.SuspendLayout()
        SplitContainer1.Panel2.SuspendLayout()
        SplitContainer1.SuspendLayout()
        TableLayoutPanel1.SuspendLayout()
        Panel1.SuspendLayout()
        SuspendLayout()
        ' 
        ' SplitContainer1
        ' 
        SplitContainer1.Dock = DockStyle.Fill
        SplitContainer1.Location = New Point(0, 0)
        SplitContainer1.Name = "SplitContainer1"
        ' 
        ' SplitContainer1.Panel1
        ' 
        SplitContainer1.Panel1.Controls.Add(tvFolders)
        ' 
        ' SplitContainer1.Panel2
        ' 
        SplitContainer1.Panel2.Controls.Add(TableLayoutPanel1)
        SplitContainer1.Size = New Size(800, 712)
        SplitContainer1.SplitterDistance = 266
        SplitContainer1.TabIndex = 0
        ' 
        ' tvFolders
        ' 
        tvFolders.Dock = DockStyle.Fill
        tvFolders.Location = New Point(0, 0)
        tvFolders.Name = "tvFolders"
        tvFolders.Size = New Size(266, 712)
        tvFolders.TabIndex = 0
        ' 
        ' TableLayoutPanel1
        ' 
        TableLayoutPanel1.BackgroundImageLayout = ImageLayout.None
        TableLayoutPanel1.ColumnCount = 1
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100F))
        TableLayoutPanel1.Controls.Add(lvFiles, 0, 1)
        TableLayoutPanel1.Controls.Add(Panel1, 0, 0)
        TableLayoutPanel1.Dock = DockStyle.Fill
        TableLayoutPanel1.Location = New Point(0, 0)
        TableLayoutPanel1.Name = "TableLayoutPanel1"
        TableLayoutPanel1.RowCount = 2
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 40F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Percent, 100F))
        TableLayoutPanel1.Size = New Size(530, 712)
        TableLayoutPanel1.TabIndex = 0
        ' 
        ' lvFiles
        ' 
        lvFiles.Dock = DockStyle.Fill
        lvFiles.Location = New Point(3, 43)
        lvFiles.Name = "lvFiles"
        lvFiles.Size = New Size(524, 666)
        lvFiles.TabIndex = 0
        lvFiles.UseCompatibleStateImageBehavior = False
        ' 
        ' Panel1
        ' 
        Panel1.Controls.Add(txtPath)
        Panel1.Controls.Add(btnForward)
        Panel1.Controls.Add(btnBack)
        Panel1.Controls.Add(btnGo)
        Panel1.Dock = DockStyle.Fill
        Panel1.Location = New Point(3, 3)
        Panel1.Name = "Panel1"
        Panel1.Size = New Size(524, 34)
        Panel1.TabIndex = 1
        ' 
        ' txtPath
        ' 
        txtPath.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        txtPath.Font = New Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        txtPath.Location = New Point(72, 3)
        txtPath.Name = "txtPath"
        txtPath.Size = New Size(383, 29)
        txtPath.TabIndex = 3
        ' 
        ' btnForward
        ' 
        btnForward.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnForward.Location = New Point(36, 3)
        btnForward.Name = "btnForward"
        btnForward.Size = New Size(30, 30)
        btnForward.TabIndex = 2
        btnForward.Text = "⏵"
        btnForward.UseVisualStyleBackColor = True
        ' 
        ' btnBack
        ' 
        btnBack.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnBack.Location = New Point(3, 3)
        btnBack.Name = "btnBack"
        btnBack.Size = New Size(30, 30)
        btnBack.TabIndex = 1
        btnBack.Text = "⏴"
        btnBack.UseVisualStyleBackColor = True
        ' 
        ' btnGo
        ' 
        btnGo.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnGo.Font = New Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        btnGo.Location = New Point(461, 3)
        btnGo.Name = "btnGo"
        btnGo.Size = New Size(60, 30)
        btnGo.TabIndex = 0
        btnGo.Text = "Go"
        btnGo.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 712)
        Controls.Add(SplitContainer1)
        Name = "Form1"
        Text = "Form1"
        SplitContainer1.Panel1.ResumeLayout(False)
        SplitContainer1.Panel2.ResumeLayout(False)
        CType(SplitContainer1, ComponentModel.ISupportInitialize).EndInit()
        SplitContainer1.ResumeLayout(False)
        TableLayoutPanel1.ResumeLayout(False)
        Panel1.ResumeLayout(False)
        Panel1.PerformLayout()
        ResumeLayout(False)
    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents tvFolders As TreeView
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents lvFiles As ListView
    Friend WithEvents Panel1 As Panel
    Friend WithEvents txtPath As TextBox
    Friend WithEvents btnForward As Button
    Friend WithEvents btnBack As Button
    Friend WithEvents btnGo As Button

End Class
