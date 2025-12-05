' File Explorer - A simple file explorer.

' MIT License
' Copyright(c) 2025 Joseph W. Lumbley

' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Imports System.IO



Public Class Form1

    ' Simple navigation history
    Private ReadOnly _history As New List(Of String)
    Private _historyIndex As Integer = -1

    ' Context menu for files
    Private cmsFiles As New ContextMenuStrip()

    Private lblStatus As New ToolStripStatusLabel()

    Private imgList As New ImageList()

    Private statusTimer As New Timer() With {.Interval = 5000}

    Private Sub ShowStatus(message As String)
        lblStatus.Text = message
        statusTimer.Stop()
        AddHandler statusTimer.Tick, AddressOf ClearStatus
        statusTimer.Start()
    End Sub

    Private Sub ClearStatus(sender As Object, e As EventArgs)
        lblStatus.Text = ""
        statusTimer.Stop()
    End Sub

    Private Sub InitImageList()
        imgList.ImageSize = New Size(16, 16)
        imgList.ColorDepth = ColorDepth.Depth32Bit

        ' Load icons (replace with actual file paths or resources)
        imgList.Images.Add("Folder", My.Resources.Resource1.Folder_16X16)
        imgList.Images.Add("Drive", My.Resources.Resource1.Drive_16X16)
        imgList.Images.Add("Documents", My.Resources.Resource1.Documents_16X16)

        imgList.Images.Add("Downloads", My.Resources.Resource1.Downloads_16X16)
        imgList.Images.Add("Desktop", My.Resources.Resource1.Desktop_16X16)
        imgList.Images.Add("EasyAccess", My.Resources.Resource1.Easy_Access_16X16)

        imgList.Images.Add("Music", My.Resources.Resource1.Music_16X16)
        imgList.Images.Add("Pictures", My.Resources.Resource1.Pictures_16X16)
        imgList.Images.Add("Videos", My.Resources.Resource1.Videos_16X16)

        imgList.Images.Add("Executable", My.Resources.Resource1.Executable_16X16)

        tvFolders.ImageList = imgList
        lvFiles.SmallImageList = imgList
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "File Explorer - Code with Joe"

        ShowStatus("Loading...")

        InitImageList()

        InitListView()
        InitTreeRoots()
        NavigateTo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))

        InitStatusBar()

        ' Context menu setup
        cmsFiles.Items.Add("Open", Nothing, AddressOf Open_Click)
        cmsFiles.Items.Add("New Folder", Nothing, AddressOf NewFolder_Click)
        cmsFiles.Items.Add("Rename", Nothing, AddressOf RenameFile_Click)
        cmsFiles.Items.Add("Copy Name", Nothing, AddressOf CopyFileName_Click)
        cmsFiles.Items.Add("Copy Path", Nothing, AddressOf CopyFilePath_Click)
        cmsFiles.Items.Add("Delete", Nothing, AddressOf Delete_Click)
        lvFiles.ContextMenuStrip = cmsFiles

        ShowStatus("Ready")

    End Sub

    ' -------- UI init --------
    Private Sub InitListView()
        lvFiles.View = View.Details
        lvFiles.FullRowSelect = True
        lvFiles.MultiSelect = True
        lvFiles.Columns.Clear()
        lvFiles.Columns.Add("Name", 475)
        lvFiles.Columns.Add("Type", 120)
        lvFiles.Columns.Add("Size", 100)
        lvFiles.Columns.Add("Modified", 160)
    End Sub

    Private Sub InitTreeRoots()

        tvFolders.Nodes.Clear()
        tvFolders.ShowRootLines = True

        ' --- Easy Access node ---
        Dim easyAccessNode As New TreeNode("Easy Access") With {
        .ImageKey = "EasyAccess",
        .SelectedImageKey = "EasyAccess"
        }

        ' Define special folders
        Dim specialFolders As (String, Environment.SpecialFolder)() = {
        ("Documents", Environment.SpecialFolder.MyDocuments),
        ("Music", Environment.SpecialFolder.MyMusic),
        ("Pictures", Environment.SpecialFolder.MyPictures),
        ("Videos", Environment.SpecialFolder.MyVideos),
        ("Downloads", Environment.SpecialFolder.UserProfile), ' Downloads handled separately
        ("Desktop", Environment.SpecialFolder.Desktop)
    }

        For Each sf In specialFolders
            Dim specialFolderPath As String = Environment.GetFolderPath(sf.Item2)

            ' Handle Downloads manually (UserProfile\Downloads)
            If sf.Item1 = "Downloads" Then
                specialFolderPath = Path.Combine(specialFolderPath, "Downloads")
            End If

            If Directory.Exists(specialFolderPath) Then
                Dim node = New TreeNode(sf.Item1) With {
                .Tag = specialFolderPath,
                .ImageKey = sf.Item1,
                .SelectedImageKey = sf.Item1
            }
                node.Nodes.Add("Loading...") ' placeholder for lazy-load
                easyAccessNode.Nodes.Add(node)
            End If
        Next

        ' Add Easy Access to tree
        tvFolders.Nodes.Add(easyAccessNode)

        ' Expand the Easy Access node
        easyAccessNode.Expand()

        ' --- Drives as separate roots ---
        For Each di In DriveInfo.GetDrives()
            Dim rootNode = New TreeNode(di.Name) With {
                .Tag = di.RootDirectory.FullName,
                .ImageKey = "Drive",
                .SelectedImageKey = "Drive"
                }
            rootNode.Nodes.Add("Loading...")
            tvFolders.Nodes.Add(rootNode)
        Next

    End Sub

    Private Sub InitStatusBar()
        Dim statusStrip As New StatusStrip()
        statusStrip.Items.Add(lblStatus)
        Me.Controls.Add(statusStrip)
    End Sub

    'Private Sub ShowStatus(message As String)
    '    lblStatus.Text = message
    'End Sub

    ' -------- Navigation --------
    Private Sub NavigateTo(path As String, Optional recordHistory As Boolean = True)
        If String.IsNullOrWhiteSpace(path) Then Exit Sub
        If Not Directory.Exists(path) Then
            MessageBox.Show("Folder not found: " & path, "Navigation", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ShowStatus("Folder not found: " & path)
            Exit Sub
        End If

        ShowStatus("Navigated To: " & path)

        txtPath.Text = path
        PopulateFiles(path)
        EnsureTreeSelection(path)

        If recordHistory Then
            ' Trim forward history if we branch
            If _historyIndex >= 0 AndAlso _historyIndex < _history.Count - 1 Then
                _history.RemoveRange(_historyIndex + 1, _history.Count - (_historyIndex + 1))
            End If
            _history.Add(path)
            _historyIndex = _history.Count - 1
            UpdateNavButtons()
        End If
    End Sub

    Private Sub UpdateNavButtons()
        btnBack.Enabled = _historyIndex > 0
        btnForward.Enabled = _historyIndex >= 0 AndAlso _historyIndex < _history.Count - 1
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        GoToFolderOrOpenFile(txtPath.Text)
    End Sub

    Private Sub GoToFolderOrOpenFile(Path As String)
        ' Navigate to folder or open file.

        ' If folder exists go there
        If Directory.Exists(Path) Then
            NavigateTo(Path)
            ' If file exists open it
        ElseIf File.Exists(Path) Then
            ' Open file with default application.
            Try
                Process.Start(New ProcessStartInfo(Path) With {.UseShellExecute = True})
                ShowStatus("Opened " & Path)
            Catch ex As Exception
                MessageBox.Show("Cannot open: " & ex.Message, "Open", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ShowStatus("Cannot open: " & ex.Message)
            End Try
        End If

    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        If _historyIndex <= 0 Then Exit Sub
        _historyIndex -= 1
        NavigateTo(_history(_historyIndex), recordHistory:=False)
        UpdateNavButtons()
    End Sub

    Private Sub btnForward_Click(sender As Object, e As EventArgs) Handles btnForward.Click
        If _historyIndex >= _history.Count - 1 Then Exit Sub
        _historyIndex += 1
        NavigateTo(_history(_historyIndex), recordHistory:=False)
        UpdateNavButtons()
    End Sub

    Private Sub tvFolders_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles tvFolders.BeforeExpand
        Dim node = e.Node
        If node.Nodes.Count = 1 AndAlso node.Nodes(0).Text = "Loading..." Then
            node.Nodes.Clear()
            Try
                For Each dirPath In Directory.GetDirectories(CStr(node.Tag))
                    Dim child = New TreeNode(Path.GetFileName(dirPath)) With {.Tag = dirPath,
                        .ImageKey = "Folder",
                .SelectedImageKey = "Folder"
                        }
                    ' Add placeholder only if it has subdirectories
                    If HasSubdirectories(dirPath) Then child.Nodes.Add("Loading...")
                    node.Nodes.Add(child)
                Next
            Catch ex As UnauthorizedAccessException
                node.Nodes.Add(New TreeNode("[Access denied]"))
            Catch ex As IOException
                node.Nodes.Add(New TreeNode("[Unavailable]"))
            End Try
        End If
    End Sub

    Private Function HasSubdirectories(path As String) As Boolean
        Try
            Return Directory.EnumerateDirectories(path).Any()
        Catch
            Return False
        End Try
    End Function

    Private Sub tvFolders_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvFolders.AfterSelect
        Dim node = e.Node
        If node Is Nothing Then Exit Sub
        NavigateTo(CStr(node.Tag))
    End Sub

    ' -------- Files population --------
    Private Sub PopulateFiles(path As String)
        lvFiles.BeginUpdate()
        lvFiles.Items.Clear()

        ' Folders first
        Try
            For Each mDir In Directory.GetDirectories(path)
                Dim di = New DirectoryInfo(mDir)
                Dim item = New ListViewItem(di.Name)
                item.SubItems.Add("Folder")
                item.SubItems.Add("") ' size blank for folders
                item.SubItems.Add(di.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
                item.Tag = di.FullName
                item.ImageKey = "Folder"
                lvFiles.Items.Add(item)
            Next
        Catch ex As UnauthorizedAccessException
            lvFiles.Items.Add(New ListViewItem("[Access denied]") With {.ForeColor = Color.DarkRed})
        End Try

        ' Files
        Try
            For Each file In Directory.GetFiles(path)
                Dim fi = New FileInfo(file)
                Dim item = New ListViewItem(fi.Name)
                item.SubItems.Add(fi.Extension.ToLowerInvariant())
                item.SubItems.Add(FormatSize(fi.Length))
                item.SubItems.Add(fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
                item.Tag = fi.FullName

                ' Assign image based on file type
                Select Case fi.Extension.ToLowerInvariant()
                    ' Music Files
                    Case ".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma",
                         ".m4a", ".alac", ".aiff", ".dsd"
                        item.ImageKey = "Music"

                    ' Image Files
                    Case ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".svg", ".webp", ".heic",
                         ".raw", ".cr2", ".nef", ".orf", ".sr2"
                        item.ImageKey = "Pictures"

                    ' Document Files
                    Case ".doc", ".docx", ".pdf", ".txt", ".xls", ".xlsx", ".ppt", ".pptx",
                         ".odt", ".ods", ".odp", ".rtf", ".html", ".htm", ".md"
                        item.ImageKey = "Documents"

                    ' Video Files
                    Case ".mp4", ".avi", ".mov", ".wmv", ".mkv", ".flv", ".webm", ".mpeg", ".mpg",
                         ".3gp", ".vob", ".ogv", ".ts"
                        item.ImageKey = "Videos"

                    ' Downloaded Files (generic)
                    Case ".zip", ".rar", ".iso", ".7z", ".tar", ".gz", ".dmg",
                         ".epub", ".mobi", ".apk", ".crx"
                        item.ImageKey = "Downloads"

                    ' Executable Files
                    Case ".exe", ".bat", ".cmd", ".msi", ".com", ".scr", ".pif",
                         ".jar", ".vbs", ".ps1", ".wsf", ".dll", ".json", ".pdb", ".sln"
                        item.ImageKey = "Executable"

                        ' Other file types can be categorized as needed
                    Case Else
                        item.ImageKey = "Documents" ' Default icon for unrecognized file types
                End Select

                lvFiles.Items.Add(item)
            Next
        Catch ex As UnauthorizedAccessException
            lvFiles.Items.Add(New ListViewItem("[Access denied]") With {.ForeColor = Color.DarkRed})
        End Try

        lvFiles.EndUpdate()
    End Sub

    Private Function FormatSize(bytes As Long) As String
        Dim units = New String() {"B", "KB", "MB", "GB", "TB"}
        Dim size = CDbl(bytes)
        Dim unitIdx = 0
        While size >= 1024 AndAlso unitIdx < units.Length - 1
            size /= 1024
            unitIdx += 1
        End While
        Return $"{size:0.##} {units(unitIdx)}"
    End Function

    ' -------- Sync tree selection --------
    Private Sub EnsureTreeSelection(path As String)
        ' Simple sync: find node by tag
        Dim match = FindNodeByTag(tvFolders.Nodes, path)
        If match IsNot Nothing Then
            tvFolders.SelectedNode = match
            match.EnsureVisible()
        End If
    End Sub

    Private Function FindNodeByTag(nodes As TreeNodeCollection, path As String) As TreeNode
        For Each n As TreeNode In nodes
            If String.Equals(CStr(n.Tag), path, StringComparison.OrdinalIgnoreCase) Then
                Return n
            End If
            Dim child = FindNodeByTag(n.Nodes, path)
            If child IsNot Nothing Then Return child
        Next
        Return Nothing
    End Function

    ' -------- Open on double-click --------
    Private Sub lvFiles_ItemActivate(sender As Object, e As EventArgs) Handles lvFiles.ItemActivate
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        Dim sel = lvFiles.SelectedItems(0)
        Dim fullPath = CStr(sel.Tag)

        GoToFolderOrOpenFile(fullPath)

    End Sub

    Private Sub lvFiles_AfterLabelEdit(sender As Object, e As LabelEditEventArgs) Handles lvFiles.AfterLabelEdit
        If e.Label Is Nothing Then Return ' user cancelled

        Dim item = lvFiles.Items(e.Item)
        Dim oldPath = CStr(item.Tag)
        Dim newName = e.Label
        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        If oldPath = newPath Then Return ' no change

        Try
            If Directory.Exists(oldPath) Then
                Directory.Move(oldPath, newPath)
                ShowStatus("Renamed Folder to: " & newName)
            ElseIf File.Exists(oldPath) Then
                File.Move(oldPath, newPath)
                ShowStatus("Renamed File to: " & newName)
            End If
            item.Tag = newPath
        Catch ex As Exception
            MessageBox.Show("Rename failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus("Rename failed: " & ex.Message)
            e.CancelEdit = True
        End Try
    End Sub

    Private Sub CopyFileName_Click(sender As Object, e As EventArgs)
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        Clipboard.SetText(lvFiles.SelectedItems(0).Text)
        ShowStatus("Copied File Name " & lvFiles.SelectedItems(0).Text)

    End Sub

    Private Sub CopyFilePath_Click(sender As Object, e As EventArgs)
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        Clipboard.SetText(CStr(lvFiles.SelectedItems(0).Tag))
        ShowStatus("Copied File Path " & lvFiles.SelectedItems(0).Tag)

    End Sub

    Private Sub RenameFile_Click(sender As Object, e As EventArgs)
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        lvFiles.SelectedItems(0).BeginEdit() ' triggers inline rename
    End Sub

    Private Sub Open_Click(sender As Object, e As EventArgs)
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        Dim fullPath = CStr(lvFiles.SelectedItems(0).Tag)
        GoToFolderOrOpenFile(fullPath)
    End Sub

    Private Sub Delete_Click(sender As Object, e As EventArgs)
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        Dim item = lvFiles.SelectedItems(0)
        Dim fullPath = CStr(item.Tag)
        Dim result = MessageBox.Show("Are you sure you want to delete '" & item.Text & "'?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If result <> DialogResult.Yes Then Exit Sub
        Try
            If Directory.Exists(fullPath) Then
                Directory.Delete(fullPath, recursive:=True)
                ShowStatus("Deleted folder " & item.Text)
            ElseIf File.Exists(fullPath) Then
                File.Delete(fullPath)
                ShowStatus("Deleted file " & item.Text)
            End If
            lvFiles.Items.Remove(item)
        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus("Delete failed: " & ex.Message)
        End Try
    End Sub


    'Private Sub DeleteSelectedItem()
    '    If lvFiles.SelectedItems.Count = 0 Then Exit Sub
    '    Dim sel = lvFiles.SelectedItems(0)
    '    Dim fullPath = CStr(sel.Tag)

    '    If MessageBox.Show("Delete '" & sel.Text & "'?", "Confirm Delete",
    '                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
    '        Try
    '            If File.Exists(fullPath) Then
    '                File.Delete(fullPath)
    '                ShowStatus("Deleted file " & sel.Text)
    '            ElseIf Directory.Exists(fullPath) Then
    '                Directory.Delete(fullPath, recursive:=True)
    '                ShowStatus("Deleted folder " & sel.Text)
    '            End If
    '            PopulateFiles(txtPath.Text)
    '        Catch ex As Exception
    '            MessageBox.Show("Delete failed: " & ex.Message, "Error",
    '                        MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Try
    '    End If
    'End Sub


    Private Sub NewFolder_Click(sender As Object, e As EventArgs)
        Dim currentPath = txtPath.Text
        Dim newFolderName = "New Folder"
        Dim newFolderPath = Path.Combine(currentPath, newFolderName)
        Dim count = 1
        While Directory.Exists(newFolderPath)
            newFolderName = $"New Folder ({count})"
            newFolderPath = Path.Combine(currentPath, newFolderName)
            count += 1
        End While
        Try
            Directory.CreateDirectory(newFolderPath)
            Dim di = New DirectoryInfo(newFolderPath)
            Dim item = New ListViewItem(di.Name)
            item.SubItems.Add("Folder")
            item.SubItems.Add("") ' size blank for folders
            item.SubItems.Add(di.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
            item.Tag = di.FullName
            item.ImageKey = "Folder"

            lvFiles.Items.Add(item)
            item.BeginEdit() ' allow user to rename immediately
            ShowStatus("Created folder: " & di.Name)
        Catch ex As Exception
            MessageBox.Show("Failed to create folder: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus("Failed to create folder: " & ex.Message)
        End Try
    End Sub

    Private Sub txtPath_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPath.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            NavigateTo(txtPath.Text)
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click

        Delete_Click(sender, e)

    End Sub

    Private Sub btnRename_Click(sender As Object, e As EventArgs) Handles btnRename.Click

        RenameFile_Click(sender, e)

    End Sub

    Private Sub btnNewFolder_Click(sender As Object, e As EventArgs) Handles btnNewFolder.Click

        NewFolder_Click(sender, e)

    End Sub

    Private Sub bntHome_Click(sender As Object, e As EventArgs) Handles bntHome.Click

        NavigateTo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))

    End Sub

End Class