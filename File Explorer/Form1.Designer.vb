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
        btnRefresh = New Button()
        bntHome = New Button()
        btnNewFolder = New Button()
        btnRename = New Button()
        btnDelete = New Button()
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
        SplitContainer1.Size = New Size(1264, 630)
        SplitContainer1.SplitterDistance = 238
        SplitContainer1.TabIndex = 0
        ' 
        ' tvFolders
        ' 
        tvFolders.Dock = DockStyle.Fill
        tvFolders.Font = New Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        tvFolders.Location = New Point(0, 0)
        tvFolders.Name = "tvFolders"
        tvFolders.ShowLines = False
        tvFolders.Size = New Size(238, 630)
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
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 34F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Percent, 100F))
        TableLayoutPanel1.Size = New Size(1022, 630)
        TableLayoutPanel1.TabIndex = 0
        ' 
        ' lvFiles
        ' 
        lvFiles.Dock = DockStyle.Fill
        lvFiles.Font = New Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        lvFiles.LabelEdit = True
        lvFiles.Location = New Point(3, 37)
        lvFiles.Name = "lvFiles"
        lvFiles.Size = New Size(1016, 590)
        lvFiles.TabIndex = 0
        lvFiles.UseCompatibleStateImageBehavior = False
        ' 
        ' Panel1
        ' 
        Panel1.Controls.Add(btnRefresh)
        Panel1.Controls.Add(bntHome)
        Panel1.Controls.Add(btnNewFolder)
        Panel1.Controls.Add(btnRename)
        Panel1.Controls.Add(btnDelete)
        Panel1.Controls.Add(txtPath)
        Panel1.Controls.Add(btnForward)
        Panel1.Controls.Add(btnBack)
        Panel1.Controls.Add(btnGo)
        Panel1.Dock = DockStyle.Fill
        Panel1.Location = New Point(3, 3)
        Panel1.Name = "Panel1"
        Panel1.Size = New Size(1016, 28)
        Panel1.TabIndex = 1
        ' 
        ' btnRefresh
        ' 
        btnRefresh.Font = New Font("Segoe UI Symbol", 8.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        btnRefresh.Location = New Point(72, 0)
        btnRefresh.Name = "btnRefresh"
        btnRefresh.Size = New Size(30, 28)
        btnRefresh.TabIndex = 8
        btnRefresh.Text = ""
        btnRefresh.UseVisualStyleBackColor = True
        ' 
        ' bntHome
        ' 
        bntHome.Anchor = AnchorStyles.Left
        bntHome.Font = New Font("Segoe UI Symbol", 8.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        bntHome.Location = New Point(108, 0)
        bntHome.Name = "bntHome"
        bntHome.Size = New Size(30, 28)
        bntHome.TabIndex = 7
        bntHome.Text = ""
        bntHome.UseVisualStyleBackColor = True
        ' 
        ' btnNewFolder
        ' 
        btnNewFolder.Anchor = AnchorStyles.Right
        btnNewFolder.Font = New Font("Segoe UI Symbol", 8.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        btnNewFolder.ForeColor = Color.Orange
        btnNewFolder.Location = New Point(914, 0)
        btnNewFolder.Name = "btnNewFolder"
        btnNewFolder.Size = New Size(30, 28)
        btnNewFolder.TabIndex = 6
        btnNewFolder.Text = ""
        btnNewFolder.UseVisualStyleBackColor = True
        ' 
        ' btnRename
        ' 
        btnRename.Anchor = AnchorStyles.Right
        btnRename.Font = New Font("Segoe UI Symbol", 8.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        btnRename.ForeColor = SystemColors.ControlText
        btnRename.Location = New Point(950, 0)
        btnRename.Name = "btnRename"
        btnRename.Size = New Size(30, 28)
        btnRename.TabIndex = 5
        btnRename.Text = ""
        btnRename.UseVisualStyleBackColor = True
        ' 
        ' btnDelete
        ' 
        btnDelete.Anchor = AnchorStyles.Right
        btnDelete.Font = New Font("Segoe UI Symbol", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        btnDelete.ForeColor = Color.Black
        btnDelete.Location = New Point(986, 0)
        btnDelete.Name = "btnDelete"
        btnDelete.Size = New Size(30, 28)
        btnDelete.TabIndex = 4
        btnDelete.Text = ""
        btnDelete.UseVisualStyleBackColor = True
        ' 
        ' txtPath
        ' 
        txtPath.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        txtPath.Font = New Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        txtPath.Location = New Point(144, 1)
        txtPath.Name = "txtPath"
        txtPath.Size = New Size(728, 27)
        txtPath.TabIndex = 3
        ' 
        ' btnForward
        ' 
        btnForward.Anchor = AnchorStyles.Left
        btnForward.Font = New Font("Segoe UI Symbol", 8.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        btnForward.Location = New Point(36, 0)
        btnForward.Name = "btnForward"
        btnForward.Size = New Size(30, 28)
        btnForward.TabIndex = 2
        btnForward.Text = ""
        btnForward.UseVisualStyleBackColor = True
        ' 
        ' btnBack
        ' 
        btnBack.Anchor = AnchorStyles.Left
        btnBack.Font = New Font("Segoe UI Symbol", 8.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        btnBack.Location = New Point(0, 0)
        btnBack.Name = "btnBack"
        btnBack.Size = New Size(30, 28)
        btnBack.TabIndex = 1
        btnBack.Text = ""
        btnBack.UseVisualStyleBackColor = True
        ' 
        ' btnGo
        ' 
        btnGo.Anchor = AnchorStyles.Right
        btnGo.Font = New Font("Segoe UI Symbol", 8.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        btnGo.Location = New Point(878, 0)
        btnGo.Name = "btnGo"
        btnGo.Size = New Size(30, 28)
        btnGo.TabIndex = 0
        btnGo.Text = "Go"
        btnGo.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1264, 630)
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
    Friend WithEvents btnDelete As Button
    Friend WithEvents btnRename As Button
    Friend WithEvents btnNewFolder As Button
    Friend WithEvents bntHome As Button
    Friend WithEvents btnRefresh As Button

End Class
