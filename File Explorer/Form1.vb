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

' Github repo: https://github.com/JoeLumbley/File-Explorer


Imports System.IO
Imports System.Text.RegularExpressions

Public Class Form1

    ' Simple navigation history
    Private ReadOnly _history As New List(Of String)
    Private _historyIndex As Integer = -1

    ' Clipboard for copy/paste
    Private _clipboardPath As String = String.Empty
    Private _clipboardIsCut As Boolean = False

    ' Context menu for files
    Private cmsFiles As New ContextMenuStrip()

    Private lblStatus As New ToolStripStatusLabel()

    Private imgList As New ImageList()

    Private imgArrows As New ImageList()

    Private statusTimer As New Timer() With {.Interval = 10000}

    Private currentFolder As String = String.Empty

    Private ShowHiddenFiles As Boolean = False

    Private ColumnTypes As New Dictionary(Of Integer, ListViewItemComparer.ColumnDataType) From {
        {0, ListViewItemComparer.ColumnDataType.Text},       ' Name
        {1, ListViewItemComparer.ColumnDataType.Text},       ' Type
        {2, ListViewItemComparer.ColumnDataType.Number},     ' Size
        {3, ListViewItemComparer.ColumnDataType.DateValue}   ' Modified
    }

    Private Shared ReadOnly SizeUnits As (Unit As String, Factor As Long)() = {
        ("B", 1L), ' Bytes
        ("KB", 1024L), ' Kilobytes
        ("MB", 1024L ^ 2), ' Megabytes
        ("GB", 1024L ^ 3), ' Gigabytes
        ("TB", 1024L ^ 4), ' Terabytes
        ("PB", 1024L ^ 5) ' Petabytes
    }

    Private _lastColumn As Integer = -1
    Private _lastOrder As SortOrder = SortOrder.Ascending

    ' Icons for status messages
    ' These require a font that supports these glyphs, such as Segoe MDL2 Assets or similar.
    Private IconError As String = ""
    Private IconDialog As String = ""
    Private IconSuccess As String = ""
    Private IconOpen As String = ""
    Private IconCopy As String = ""
    Private IconPaste As String = ""
    Private IconProtect As String = ""
    Private IconNavigate As String = ""
    Private IconSmile As String = ""
    Private IconWarning As String = "⚠"
    Private IconDelete As String = ""
    Private IconNewFolder As String = ""
    Private IconCut As String = "✂"
    Private IconSearch As String = ""

    Dim SearchResults As New List(Of String)
    Private SearchIndex As Integer = -1

    Private Sub Form_Load(sender As Object, e As EventArgs) _
        Handles MyBase.Load

        InitApp()

    End Sub

    Private Sub txtPath_KeyDown(sender As Object, e As KeyEventArgs) _
        Handles txtPath.KeyDown

        Path_KeyDown(e)

    End Sub

    Private Sub tvFolders_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) _
        Handles tvFolders.NodeMouseClick

        Dim info = tvFolders.HitTest(e.Location)

        ' Only toggle node when the arrow is clicked
        If info.Location = TreeViewHitTestLocations.StateImage Then

            If e.Node.IsExpanded Then
                e.Node.Collapse()
            Else
                e.Node.Expand()
            End If

        End If

    End Sub

    Private Sub tvFolders_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) _
        Handles tvFolders.BeforeExpand

        ExpandNode_LazyLoad(e.Node)

    End Sub

    Private Sub tvFolders_AfterSelect(sender As Object, e As TreeViewEventArgs) _
        Handles tvFolders.AfterSelect

        NavigateToSelectedFolderTreeView_AfterSelect(sender, e)

    End Sub

    Private Sub tvFolders_BeforeCollapse(sender As Object, e As TreeViewCancelEventArgs) _
        Handles tvFolders.BeforeCollapse

        e.Node.StateImageIndex = 0   ' ▶ collapsed

    End Sub


    Private Sub lvFiles_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) _
        Handles lvFiles.ItemSelectionChanged

        UpdateEditButtons()

        UpdateEditContextMenu()

    End Sub

    Private Sub lvFiles_ItemActivate(sender As Object, e As EventArgs) _
        Handles lvFiles.ItemActivate
        ' The ItemActivate event is raised when the user double-clicks an item or
        ' presses the Enter key when an item is selected.

        ' Validate selection
        If Not PathExists(CStr(lvFiles.SelectedItems(0).Tag)) Then
            ShowStatus(IconError & " The selected item isn't a file or folder and can't be opened.")
            Exit Sub
        End If

        GoToFolderOrOpenFile_EnterKeyDownOrDoubleClick()

    End Sub

    Private Sub lvFiles_BeforeLabelEdit(sender As Object, e As LabelEditEventArgs) _
        Handles lvFiles.BeforeLabelEdit

        Dim item As ListViewItem = lvFiles.Items(e.Item)
        Dim fullPath As String = CStr(item.Tag)

        ' Rule 1: Path must exist
        If Not PathExists(fullPath) Then
            e.CancelEdit = True
            ShowStatus(IconError & " The selected item doesn't exist and cannot be renamed.")
            Exit Sub
        End If

        ' Rule 2: Protected items cannot be renamed
        If IsProtectedPathOrFolder(fullPath) Then
            e.CancelEdit = True
            ShowStatus(IconProtect & " This item is protected and cannot be renamed.")
            Exit Sub
        End If

        ' Rule 3: User must have rename permission
        Dim parentDir As String
        If Directory.Exists(fullPath) Then
            ' Item is a folder → check write access ON the folder
            parentDir = fullPath
        Else
            ' Item is a file → check write access on its parent directory
            parentDir = Path.GetDirectoryName(fullPath)
        End If

        If Not HasWriteAccessToDirectory(parentDir) Then
            e.CancelEdit = True
            ShowStatus(IconError & " This location does not allow renaming.")
            Exit Sub
        End If

    End Sub

    Private Sub lvFiles_AfterLabelEdit(sender As Object, e As LabelEditEventArgs) _
        Handles lvFiles.AfterLabelEdit

        RenameFileOrFolder_AfterLabelEdit(e)

    End Sub

    Private Sub lvFiles_ColumnClick(sender As Object, e As ColumnClickEventArgs) _
        Handles lvFiles.ColumnClick

        If e.Column = _lastColumn Then
            _lastOrder = If(_lastOrder = SortOrder.Ascending,
                        SortOrder.Descending,
                        SortOrder.Ascending)
        Else
            _lastColumn = e.Column
            _lastOrder = SortOrder.Ascending
        End If

        UpdateColumnHeaders(e.Column, _lastOrder)

        lvFiles.ListViewItemSorter =
        New ListViewItemComparer(_lastColumn, _lastOrder, ColumnTypes)

        lvFiles.Sort()

    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) _
        Handles btnBack.Click

        NavigateBackward_Click()

    End Sub

    Private Sub btnForward_Click(sender As Object, e As EventArgs) _
        Handles btnForward.Click

        NavigateForward_Click()

    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) _
        Handles btnRefresh.Click

        InitTreeRoots()

        NavigateTo(currentFolder, recordHistory:=False)

    End Sub

    Private Sub bntHome_Click(sender As Object, e As EventArgs) _
        Handles bntHome.Click

        NavigateTo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), True)

        UpdateNavButtons()

    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) _
        Handles btnGo.Click

        ExecuteCommand(txtPath.Text.Trim())

    End Sub

    Private Sub btnNewFolder_Click(sender As Object, e As EventArgs) _
        Handles btnNewFolder.Click

        NewFolder_Click(sender, e)

    End Sub

    Private Sub btnNewTextFile_Click(sender As Object, e As EventArgs) _
        Handles btnNewTextFile.Click

        NewTextFile_Click(sender, e)

    End Sub

    Private Sub btnCut_Click(sender As Object, e As EventArgs) _
        Handles btnCut.Click

        CutSelected_Click(sender, e)

    End Sub

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) _
        Handles btnCopy.Click

        CopySelected_Click(sender, e)

    End Sub

    Private Sub btnPaste_Click(sender As Object, e As EventArgs) _
        Handles btnPaste.Click

        PasteSelected_Click(sender, e)

    End Sub

    Private Sub btnRename_Click(sender As Object, e As EventArgs) _
        Handles btnRename.Click

        RenameFile_Click(sender, e)

    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) _
        Handles btnDelete.Click

        Delete_Click(sender, e)

    End Sub

    Private Sub Path_KeyDown(e As KeyEventArgs)
        ' Path command input box key handling

        ' Check for Enter key
        If e.KeyCode = Keys.Enter Then

            e.SuppressKeyPress = True

            Dim command As String = txtPath.Text.Trim()

            ExecuteCommand(command)

        End If

    End Sub

    Private Sub ExpandNode_LazyLoad(node As TreeNode)

        ' Only lazy-load if placeholder exists
        If node.Nodes.Count = 1 AndAlso node.Nodes(0).Text = "Loading..." Then

            node.Nodes.Clear()

            Dim basePath As String = CStr(node.Tag)

            Try
                For Each dirPath In Directory.GetDirectories(basePath)
                    Dim di As New DirectoryInfo(dirPath)

                    ' Skip hidden/system folders
                    If (di.Attributes And (FileAttributes.Hidden Or FileAttributes.System)) <> 0 Then
                        Continue For
                    End If

                    Dim child As New TreeNode(di.Name) With {
                        .Tag = dirPath,
                        .ImageKey = "Folder",
                        .SelectedImageKey = "Folder"
                    }

                    If HasSubdirectories(dirPath) Then
                        child.Nodes.Add("Loading...")
                        child.StateImageIndex = 0   ' ▶ collapsed
                    Else
                        child.StateImageIndex = 2   ' no arrow
                    End If

                    node.Nodes.Add(child)
                Next

            Catch ex As UnauthorizedAccessException
                node.Nodes.Add(New TreeNode("[Access denied]") With {.ForeColor = Color.Gray})
                Debug.WriteLine($"[Access denied]: {ex.Message}")
            Catch ex As IOException
                node.Nodes.Add(New TreeNode("[Unavailable]") With {.ForeColor = Color.Gray})
                Debug.WriteLine($"[Unavailable]: {ex.Message}")
            Catch ex As Exception
                Debug.WriteLine($"ExpandNode_LazyLoad Error: {ex.Message}")
            End Try
        End If

        ' Set expanded icon
        node.StateImageIndex = 1   ' ▼ expanded

    End Sub

    Private Sub NavigateToSelectedFolderTreeView_AfterSelect(sender As Object, e As TreeViewEventArgs)

        ' Get the selected node
        Dim node As TreeNode = e.Node
        If node Is Nothing Then Exit Sub

        ' Ensure the Tag is a string path
        Dim path2Nav As String = TryCast(node.Tag, String)
        If String.IsNullOrEmpty(path2Nav) Then Exit Sub

        ' Check if the node represents a drive root
        Try
            Dim driveInfo As New DriveInfo(IO.Path.GetPathRoot(path2Nav))

            ' Prune the tree if the drive is not ready
            If driveInfo.IsReady = False Then
                tvFolders.Nodes.Remove(node)
                ShowStatus(IconWarning & " Drive is not ready and has been removed.")
                Return
            End If

        Catch ex As Exception
            ' Handle any exceptions when accessing DriveInfo
            ShowStatus(IconError & " NavTree: Error accessing drive: " & ex.Message)
            Debug.WriteLine("NavTree AfterSelect: Error accessing drive: " & ex.Message)
            Return
        End Try

        ' If the drive is ready, navigate to the folder
        NavigateTo(path2Nav)

    End Sub


    Private Sub PopulateFiles(path As String)

        lvFiles.BeginUpdate()

        lvFiles.Items.Clear()

        ' Folders first
        Try

            For Each mDir In Directory.GetDirectories(path)

                Dim di = New DirectoryInfo(mDir)

                ' Skip hidden/system folders unless checkbox is checked
                If Not ShowHiddenFiles AndAlso
                   (di.Attributes And (FileAttributes.Hidden Or FileAttributes.System)) <> 0 Then
                    Continue For
                End If

                Dim item = New ListViewItem(di.Name)
                item.SubItems.Add("Folder")
                item.SubItems.Add("") ' size blank for folders
                item.SubItems.Add(di.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
                item.Tag = di.FullName
                item.ImageKey = "Folder"

                lvFiles.Items.Add(item)

            Next

        Catch ex As UnauthorizedAccessException
            ShowStatus(IconError & " Access Denied")
            Debug.WriteLine($"PopulateFiles [Folder Access Denied]: {ex.Message}")
        Catch ex As Exception
            ShowStatus(IconError & $" Error: {ex.Message}")
            Debug.WriteLine($"PopulateFiles Folder [Error]: {ex.Message}")
        End Try


        ' Files
        Try
            For Each file In Directory.GetFiles(path)
                Dim fi = New FileInfo(file)

                ' Skip hidden/system files unless checkbox is checked
                If Not ShowHiddenFiles AndAlso
               (fi.Attributes And (FileAttributes.Hidden Or FileAttributes.System)) <> 0 Then
                    Continue For
                End If

                Dim item = New ListViewItem(fi.Name)
                item.SubItems.Add(fi.Extension.ToLowerInvariant())
                item.SubItems.Add(FormatSize(fi.Length))
                item.SubItems.Add(fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
                item.Tag = fi.FullName

                ' Assign image based on file type
                Select Case fi.Extension.ToLowerInvariant()
                    Case ".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma",
                     ".m4a", ".alac", ".aiff", ".dsd"
                        item.ImageKey = "Music"
                    Case ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".svg", ".webp", ".heic",
                     ".raw", ".cr2", ".nef", ".orf", ".sr2"
                        item.ImageKey = "Pictures"
                    Case ".doc", ".docx", ".pdf", ".txt", ".xls", ".xlsx", ".ppt", ".pptx",
                     ".odt", ".ods", ".odp", ".rtf", ".html", ".htm", ".md"
                        item.ImageKey = "Documents"
                    Case ".mp4", ".avi", ".mov", ".wmv", ".mkv", ".flv", ".webm", ".mpeg", ".mpg",
                     ".3gp", ".vob", ".ogv", ".ts"
                        item.ImageKey = "Videos"
                    Case ".zip", ".rar", ".iso", ".7z", ".tar", ".gz", ".dmg",
                     ".epub", ".mobi", ".apk", ".crx"
                        item.ImageKey = "Downloads"
                    Case ".exe", ".bat", ".cmd", ".msi", ".com", ".scr", ".pif",
                     ".jar", ".vbs", ".ps1", ".wsf", ".dll", ".json", ".pdb", ".sln"
                        item.ImageKey = "Executable"
                    Case Else
                        item.ImageKey = "Documents"
                End Select

                lvFiles.Items.Add(item)

            Next

        Catch ex As UnauthorizedAccessException
            ShowStatus(IconError & " Access Denied")
            Debug.WriteLine($"PopulateFiles [File Access Denied]: {ex.Message}")
        Catch ex As Exception
            ShowStatus(IconError & $" Error: {ex.Message}")
            Debug.WriteLine($"PopulateFiles File [Error]: {ex.Message}")
        End Try

        lvFiles.EndUpdate()

    End Sub

    Private Sub GoToFolderOrOpenFile_EnterKeyDownOrDoubleClick()
        ' This event is triggered when the user double-clicks a file or folder in lvFiles or
        ' presses the Enter key when a file or folder is selected.

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        Dim sel As ListViewItem = lvFiles.SelectedItems(0)

        Dim fullPath As String = CStr(sel.Tag)

        GoToFolderOrOpenFile(fullPath)

    End Sub

    Private Sub RenameFileOrFolder_AfterLabelEdit(ByRef e As LabelEditEventArgs)
        ' -------- Rename file or folder after label edit in lvFiles --------

        If e.Label Is Nothing Then Return ' user cancelled

        Dim item = lvFiles.Items(e.Item)
        Dim oldPath = CStr(item.Tag)
        Dim newName = e.Label
        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        If oldPath = newPath Then Return ' no change

        Try
            ' Validate new name
            If Directory.Exists(oldPath) Then

                Directory.Move(oldPath, newPath)

                ShowStatus(IconSuccess & " Renamed Folder to: " & newName)

            ElseIf File.Exists(oldPath) Then

                File.Move(oldPath, newPath)

                ShowStatus(IconSuccess & " Renamed File to: " & newName)

            End If

            item.Tag = newPath

        Catch ex As Exception

            e.CancelEdit = True

            ShowStatus(IconError & " Rename Failed. " & ex.Message)

            Debug.WriteLine("RenameFileOrFolder_AfterLabelEdit Error: " & ex.Message)

        End Try

    End Sub


    Private Sub NavigateBackward_Click()
        ' Navigate backward in the history list

        ' If we're already at the first entry, there's nowhere to go
        If _historyIndex <= 0 Then Exit Sub

        ' Move one step back in history
        _historyIndex -= 1

        ' Navigate to the previous location without recording a new history entry
        NavigateTo(_history(_historyIndex), recordHistory:=False)

        ' Refresh the enabled/disabled state of Back/Forward buttons
        UpdateNavButtons()

    End Sub

    Private Sub NavigateForward_Click()
        ' Navigate forward in the history list

        ' If we're at the most recent entry, we can't go forward
        If _historyIndex >= _history.Count - 1 Then Exit Sub

        ' Move one step forward in history
        _historyIndex += 1

        ' Navigate to the next location without recording a new history entry
        NavigateTo(_history(_historyIndex), recordHistory:=False)

        ' Refresh the enabled/disabled state of Back/Forward buttons
        UpdateNavButtons()

    End Sub


    Private Sub Open_Click(sender As Object, e As EventArgs)
        ' Open selected file or folder - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        Dim fullPath = CStr(lvFiles.SelectedItems(0).Tag)
        GoToFolderOrOpenFile(fullPath)

    End Sub


    Private Sub NewFolder_Click(sender As Object, e As EventArgs)
        ' Create new folder in current directory - Mouse right-click context menu for lvFiles

        Dim currentPath As String = currentFolder
        Dim newFolderName As String = "New Folder"
        Dim newFolderPath As String = Path.Combine(currentPath, newFolderName)

        ' Ensure unique folder name
        Dim count As Integer = 1
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

            ShowStatus(IconNewFolder & " Created folder: " & di.Name)

        Catch ex As Exception

            ShowStatus(IconError & " New folder creation failed. " & ex.Message)

            Debug.WriteLine("NewFolder_Click Error: " & ex.Message)

        End Try

    End Sub

    Private Sub NewTextFile_Click(sender As Object, e As EventArgs)

        Dim destDir As String = currentFolder

        ' Validate destination folder
        If String.IsNullOrWhiteSpace(destDir) OrElse Not Directory.Exists(destDir) Then
            ShowStatus(IconWarning & " Invalid folder. Cannot create file.")
            Return
        End If

        ' Base filename
        Dim baseName As String = "New Text File"
        Dim newFilePath As String = Path.Combine(destDir, baseName & ".txt")

        ' Ensure unique name
        Dim counter As Integer = 1
        While File.Exists(newFilePath)
            newFilePath = Path.Combine(destDir, $"{baseName} ({counter}).txt")
            counter += 1
        End While

        Try
            ' Create the file with initial content
            File.WriteAllText(newFilePath, $"Created on {DateTime.Now:G}")

            ShowStatus(IconSuccess & " Text file created: " & newFilePath)

            ' Refresh the folder view so the user sees the new file
            NavigateTo(destDir)

            ' Open the newly created file
            GoToFolderOrOpenFile(newFilePath)

        Catch ex As Exception

            ShowStatus(IconError & " New text file creation failed. " & ex.Message)

            Debug.WriteLine("NewTextFile_Click Error: " & ex.Message)

        End Try

    End Sub


    Private Sub CutSelected_Click(sender As Object, e As EventArgs)

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        ' Store in internal clipboard
        _clipboardPath = CStr(lvFiles.SelectedItems(0).Tag)

        _clipboardIsCut = True

        ' Fade the item to indicate "cut"
        Dim sel = lvFiles.SelectedItems(0)
        sel.ForeColor = Color.Gray
        sel.Font = New Font(sel.Font, FontStyle.Italic)

        ShowStatus(IconCut & " Cut to clipboard: " & _clipboardPath)

    End Sub

    Private Sub CopySelected_Click(sender As Object, e As EventArgs)

        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        Dim path As String = CStr(lvFiles.SelectedItems(0).Tag)

        ' Validate that the selected item actually exists
        If Not PathExists(path) Then
            ShowStatus(IconWarning & " Copy failed: Selected item does not exist.")
            Exit Sub
        End If

        ' Copy to system clipboard
        'Clipboard.SetText(path)

        ' Update internal clipboard
        _clipboardPath = path
        _clipboardIsCut = False

        ShowStatus(IconCopy & " Copied to clipboard: " & _clipboardPath)

        UpdateEditButtons()
        UpdateEditContextMenu()


    End Sub

    Private Sub PasteSelected_Click(sender As Object, e As EventArgs)

        ' Is a file or folder selected?
        If String.IsNullOrEmpty(_clipboardPath) Then Exit Sub

        Dim destDir = currentFolder
        Dim destPath = Path.Combine(destDir, Path.GetFileName(_clipboardPath))

        If destPath = Nothing Then

            ShowStatus(IconError & "Paste failed: No Path")

            Exit Sub

        End If

        Try
            If File.Exists(_clipboardPath) Then
                If _clipboardIsCut Then
                    File.Move(_clipboardPath, destPath)
                Else
                    File.Copy(_clipboardPath, destPath, overwrite:=False)
                End If
            ElseIf Directory.Exists(_clipboardPath) Then
                If _clipboardIsCut Then
                    Directory.Move(_clipboardPath, destPath)
                Else
                    CopyDirectory(_clipboardPath, destPath)
                End If

            Else

                ShowStatus("Paste failed: No Path")

                Exit Sub

            End If

            ' Clear cut state
            _clipboardPath = Nothing
            _clipboardIsCut = False

            ' Refresh current folder view
            PopulateFiles(destDir)

            ResetCutVisuals()

            ShowStatus(IconPaste & " Pasted into " & txtPath.Text)

        Catch ex As Exception
            MessageBox.Show("Paste failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus(IconError & " Paste failed: " & ex.Message)
            Debug.WriteLine("PasteSelected_Click Error: " & ex.Message)
        End Try

    End Sub


    Private Sub RenameFile_Click(sender As Object, e As EventArgs)
        ' Rename selected file or folder - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        lvFiles.SelectedItems(0).BeginEdit() ' triggers inline rename

    End Sub

    Private Sub Delete_Click(sender As Object, e As EventArgs)
        ' Delete selected file or folder - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        Dim item = lvFiles.SelectedItems(0)
        Dim fullPath = CStr(item.Tag)

        ' Reject relative paths outright
        If Not Path.IsPathRooted(fullPath) Then

            ShowStatus(IconWarning & " Delete failed: Path must be absolute. Example: C:\folder")

            Exit Sub

        End If

        ' Check if the path is in the protected list
        If IsProtectedPathOrFolder(fullPath) Then
            ' The path is protected; prevent deletion
            ShowStatus(IconProtect & " Deletion prevented for protected path: " & fullPath)
            Dim msg As String = "Deletion prevented for protected path: " & Environment.NewLine & fullPath
            MsgBox(msg, MsgBoxStyle.Critical, "Deletion Prevented")
            Exit Sub
        End If

        ' Confirm deletion
        Dim result = MessageBox.Show("Are you sure you want to delete '" & item.Text & "'?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If result <> DialogResult.Yes Then Exit Sub

        Try
            ' Check if it's a directory
            If Directory.Exists(fullPath) Then
                Directory.Delete(fullPath, recursive:=True)
                lvFiles.Items.Remove(item)
                ShowStatus(IconDelete & " Deleted folder: " & item.Text)
                ' Check if it's a file
            ElseIf File.Exists(fullPath) Then
                File.Delete(fullPath)
                lvFiles.Items.Remove(item)
                ShowStatus(IconDelete & " Deleted file: " & item.Text)
            Else
                ShowStatus(IconWarning & " Path not found.")
            End If
        Catch ex As Exception

            'MessageBox.Show("Delete failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus(IconError & " Delete failed: " & ex.Message)

            Debug.WriteLine("Delete_Click Error: " & ex.Message)

        End Try

    End Sub


    Private Sub CopyFileName_Click(sender As Object, e As EventArgs)
        ' Copy selected file name to clipboard - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        Clipboard.SetText(lvFiles.SelectedItems(0).Text)

        ShowStatus(IconCopy & " Copied File Name " & lvFiles.SelectedItems(0).Text)

    End Sub

    Private Sub CopyFilePath_Click(sender As Object, e As EventArgs)
        ' Copy selected file path to clipboard - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        ' Copy the full path stored in the Tag property to clipboard
        Clipboard.SetText(CStr(lvFiles.SelectedItems(0).Tag))

        ShowStatus(IconCopy & " Copied File Path " & lvFiles.SelectedItems(0).Tag)

    End Sub


    Private Sub ExecuteCommand(command As String)

        ' Use regex to split by spaces but keep quoted substrings together
        Dim parts As String() = Regex.Matches(command, "[\""].+?[\""]|[^ ]+").
        Cast(Of Match)().
        Select(Function(m) m.Value.Trim(""""c)).
        ToArray()

        If parts.Length = 0 Then
            ShowStatus(IconDialog & " No command entered.")
            Return
        End If

        Dim cmd As String = parts(0).ToLower()

        Select Case cmd

            Case "cd"

                If parts.Length > 1 Then
                    Dim newPath As String = String.Join(" ", parts.Skip(1)).Trim()
                    NavigateTo(newPath)
                Else
                    ShowStatus(IconDialog & " Usage: cd [directory] - cd C:\ ")
                End If

            Case "copy"

                If parts.Length > 2 Then

                    Dim source As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()
                    Dim destination As String = parts(parts.Length - 1).Trim()

                    ' Check if source file or directory exists
                    If Not (File.Exists(source) OrElse Directory.Exists(source)) Then
                        ShowStatus(IconWarning & " Copy Failed - Source: ''" & source & "'' does not exist.")
                        Return
                    End If

                    ' Check if destination directory exists
                    If Not Directory.Exists(destination) Then
                        ShowStatus(IconWarning & " Copy Failed - Destination folder: ''" & destination & "'' does not exist.")
                        Return
                    End If

                    ' If source is a file, copy it
                    If File.Exists(source) Then
                        CopyFile(source, destination)
                    Else
                        ' If source is a directory, copy it
                        CopyDirectory(source, Path.Combine(destination, Path.GetFileName(source)))
                    End If

                Else
                    ShowStatus(IconDialog & " Usage: copy [source] [destination] - e.g., copy C:\folder1\file.doc C:\folder2")
                End If

            Case "move"

                If parts.Length > 2 Then

                    Dim source As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()
                    Dim destination As String = parts(parts.Length - 1).Trim()

                    MoveFileOrDirectory(source, destination)

                Else
                    ShowStatus(IconDialog & " Usage: move [source] [destination] - move C:\folder1\directoryToMove C:\folder2\directoryToMove")
                End If

            Case "delete"

                If parts.Length > 1 Then

                    Dim pathToDelete As String = String.Join(" ", parts.Skip(1)).Trim()

                    DeleteFileOrDirectory(pathToDelete)

                Else
                    ShowStatus(IconDialog & " Usage: delete [file_or_directory]")
                End If

            Case "mkdir", "make" ' You can use "mkdir" or "make" as the command
                If parts.Length > 1 Then
                    Dim directoryPath As String = String.Join(" ", parts.Skip(1)).Trim()

                    ' Validate the directory path
                    If String.IsNullOrWhiteSpace(directoryPath) Then
                        ShowStatus(IconDialog & " Usage: mkdir [directory_path] - e.g., mkdir C:\newfolder")
                        Return
                    End If

                    CreateDirectory(directoryPath)

                Else
                    ShowStatus(IconDialog & " Usage: mkdir [directory_path] - e.g., mkdir C:\newfolder")
                End If

            Case "rename"
                ' Rename file or directory

                If parts.Length > 2 Then

                    Dim sourcePath As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()

                    Dim newName As String = parts(parts.Length - 1).Trim()

                    RenameFileOrDirectory(sourcePath, newName)

                Else
                    ShowStatus(IconDialog & " Usage: rename [source_path] [new_name] - e.g., rename C:\folder\oldname.txt newname.txt")
                End If

            Case "text", "txt"

                If parts.Length > 1 Then

                    Dim rawPath = String.Join(" ", parts.Skip(1))

                    Dim filePath = NormalizeTextFilePath(rawPath)

                    If filePath Is Nothing Then

                        ShowStatus(IconDialog & " Usage: text [file_path]  e.g., text C:\example.txt")

                        Return

                    End If

                    CreateTextFile(filePath)

                    Return

                End If

                ' No file path provided
                Dim destDir = currentFolder

                ' Validate destination folder
                If String.IsNullOrWhiteSpace(destDir) OrElse Not Directory.Exists(destDir) Then

                    ShowStatus(IconWarning & " Invalid folder. Cannot create file.")

                    Return

                End If

                ' Ensure unique file name
                Dim newFilePath = GetUniqueFilePath(destDir, "New Text File", ".txt")

                Try

                    ' Create the file with initial content
                    File.WriteAllText(newFilePath, $"Created on {DateTime.Now:G}")

                    ShowStatus(IconSuccess & " Text file created: " & newFilePath)

                    ' Refresh the folder view so the user sees the new file
                    NavigateTo(destDir)

                    ' Open the newly created file
                    GoToFolderOrOpenFile(newFilePath)

                Catch ex As Exception
                    ShowStatus(IconError & " Failed to create text file: " & ex.Message)
                    Debug.WriteLine("Text Command Error: " & ex.Message)
                End Try

            Case "help"

                Dim helpText As String = "Available Commands:" & Environment.NewLine & Environment.NewLine &
                                         "cd [directory]" & Environment.NewLine &
                                         "Change directory to the specified path" & Environment.NewLine &
                                         "cd C:\" & Environment.NewLine &
                                         "cd C:\folder" & Environment.NewLine & Environment.NewLine &
                                         "mkdir [directory_path]" & Environment.NewLine &
                                         "Creates a new folder" & Environment.NewLine &
                                         "mkdir C:\newfolder" & Environment.NewLine &
                                         "make C:\newfolder" & Environment.NewLine & Environment.NewLine &
                                         "copy [source] [destination]" & Environment.NewLine &
                                         "Copy file or folder to destination folder" & Environment.NewLine &
                                         "copy C:\folderA\file.doc C:\folderB" & Environment.NewLine &
                                         "copy C:\folderA C:\folderB" & Environment.NewLine & Environment.NewLine &
                                         "move [source] [destination]" & Environment.NewLine &
                                         "Move file or folder to destination" & Environment.NewLine &
                                         "move C:\folderA\file.doc C:\folderB\file.doc" & Environment.NewLine &
                                         "move C:\folderA\file.doc C:\folderB\rename.doc" & Environment.NewLine &
                                         "move C:\folderA\folder C:\folderB\folder" & Environment.NewLine &
                                         "move C:\folderA\folder C:\folderB\rename" & Environment.NewLine & Environment.NewLine &
                                         "delete [file or directory]" & Environment.NewLine &
                                         "delete C:\file" & Environment.NewLine &
                                         "delete C:\folder" & Environment.NewLine & Environment.NewLine &
                                         "text [file]" & Environment.NewLine &
                                         "Creates a new text file at the specified path." & Environment.NewLine &
                                         "text C:\folder\example.txt" & Environment.NewLine & Environment.NewLine &
                                         "help - Show this help message"

                MessageBox.Show(helpText, "Help", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Case "find", "search"

                If parts.Length > 1 Then

                    Dim searchTerm As String = String.Join(" ", parts.Skip(1)).Trim()

                    ShowStatus(IconSearch & " Searching for: " & searchTerm)

                    SearchInCurrentFolder(searchTerm)

                Else
                    ShowStatus(IconDialog & " Usage: find [search_term] - e.g., find document")
                End If

                SearchIndex = 0

            Case "findnext", "searchnext"

                If SearchResults.Count = 0 Then
                    ShowStatus(IconDialog & " No previous search results. Use 'find [search_term]' to start a search.")
                    Return
                End If

                ' Advance index
                SearchIndex += 1

                ' Wrap around if needed
                If SearchIndex >= SearchResults.Count Then
                    SearchIndex = 0
                End If

                ' Clear item selection for lvFiles
                lvFiles.SelectedItems.Clear()

                ' Select the item
                Dim nextPath As String = SearchResults(SearchIndex)
                SelectListViewItemByPath(nextPath)
                lvFiles.Focus()

                ShowStatus(IconSmile & " Showing result " & (SearchIndex + 1) & " of " & SearchResults.Count)


            Case "exit", "quit"
                ' Exit the application

                Me.Close()

            Case Else

                ' Is the input a folder?
                If Directory.Exists(command) Then

                    ' Go to that folder.
                    NavigateTo(command)

                    ' Is the input a file?
                ElseIf File.Exists(command) Then

                    ' Open the file or go to its folder
                    GoToFolderOrOpenFile(command)

                Else

                    ' The input isn't a folder or a file,
                    ' at this point, the interpreter treats it as an unknown command.
                    ShowStatus(IconDialog & " Unknown command: " & cmd)

                End If

        End Select

    End Sub


    Private Sub NavigateTo(path As String, Optional recordHistory As Boolean = True)
        ' Navigate to the specified folder path.

        '  Updates the current folder, path textbox, and file list.
        If String.IsNullOrWhiteSpace(path) Then Exit Sub

        ' Validate that the folder exists
        If Not Directory.Exists(path) Then
            MessageBox.Show("Folder not found: " & path, "Navigation", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ShowStatus(IconWarning & " Folder not found: " & path)
            Exit Sub
        End If

        ShowStatus(IconNavigate & " Navigated To: " & path)

        currentFolder = path
        txtPath.Text = path
        PopulateFiles(path)

        If recordHistory Then
            ' Trim forward history if we branch
            If _historyIndex >= 0 AndAlso _historyIndex < _history.Count - 1 Then
                _history.RemoveRange(_historyIndex + 1, _history.Count - (_historyIndex + 1))
            End If
            _history.Add(path)
            _historyIndex = _history.Count - 1
            UpdateNavButtons()
        End If

        UpdateFileButtons()
        UpdateEditButtons()
        UpdateEditContextMenu()

    End Sub

    Private Sub GoToFolderOrOpenFile(Path As String)
        ' Navigate to folder or open file.

        ' If folder exists, go there
        If Directory.Exists(Path) Then

            NavigateTo(Path, True)

            ' If file exists, open it
        ElseIf File.Exists(Path) Then

            Try

                ' Open file with default application.
                Dim processStartInfo As New ProcessStartInfo(Path) With {.UseShellExecute = True}
                Dim process As Process = Process.Start(processStartInfo)

                ShowStatus(IconOpen & " Opened " & Path)

            Catch ex As Exception
                ShowStatus(IconError & " Cannot open: " & ex.Message)

                Debug.WriteLine("GoToFolderOrOpenFile: Error opening file: " & ex.Message)
            End Try

        Else
            ShowStatus(IconWarning & " Path does not exist: " & Path)
        End If

    End Sub


    Private Sub CreateDirectory(directoryPath As String)

        Try
            ' Create the directory
            Dim dirInfo As DirectoryInfo = Directory.CreateDirectory(directoryPath)

            ShowStatus(IconNewFolder & " Directory created: " & dirInfo.FullName)

            NavigateTo(directoryPath, True)

        Catch ex As Exception
            ShowStatus(IconError & " Failed to create directory: " & ex.Message)
            Debug.WriteLine("CreateDirectory Error: " & ex.Message)
        End Try

    End Sub

    Private Sub CreateTextFile(filePath As String)

        Try
            Dim destDir As String = Path.GetDirectoryName(filePath)

            ' If the file does not exist, create it
            If Not File.Exists(filePath) Then

                ' Create the file with initial content
                File.WriteAllText(filePath, $"Created on {DateTime.Now:G}")

                ShowStatus(IconSuccess & " Text file created: " & filePath)

            Else
                ShowStatus(IconDialog & " File already exists: " & filePath)
            End If

            ' Refresh the folder view so the user sees the new or existing file
            NavigateTo(destDir)

            ' Open the newly created or existing file
            GoToFolderOrOpenFile(filePath)

        Catch ex As Exception
            ShowStatus(IconError & " Failed to create text file: " & ex.Message)
            Debug.WriteLine("CreateTextFile Error: " & ex.Message)

        End Try

    End Sub


    Private Sub CopyFile(source As String, destination As String)
        Try
            ' Check if the destination file already exists
            Dim fileName As String = Path.GetFileName(source)
            Dim destDirFileName = Path.Combine(destination, fileName)
            If File.Exists(destDirFileName) Then
                Dim msg As String = "The file '" & fileName & "' already exists in the destination folder." & Environment.NewLine &
                                "Do you want to overwrite it?"
                Dim result = MessageBox.Show(msg, "File Exists", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                If result = DialogResult.No Then
                    ShowStatus(IconWarning & " Copy operation canceled.")
                    Return
                End If
            End If

            File.Copy(source, destDirFileName, overwrite:=True)

            ShowStatus(IconCopy & " Copied file: " & fileName & " to: " & destination)

            NavigateTo(destination)

        Catch ex As Exception
            ShowStatus(IconError & " Copy Failed: " & ex.Message)
            Debug.WriteLine("CopyFile Error: " & ex.Message)
        End Try
    End Sub

    Private Sub MoveFileOrDirectory(source As String, destination As String)
        Try
            ' Validate parameters
            If String.IsNullOrWhiteSpace(source) OrElse String.IsNullOrWhiteSpace(destination) Then
                ShowStatus(IconWarning & " Source or destination path is invalid.")
                Return
            End If

            ' Check if the source is a file
            If File.Exists(source) Then
                If Not File.Exists(destination) Then
                    File.Move(source, destination)
                    'ShowStatus("Moved file to: " & destination)
                    MsgBox("Moved file to: " & destination)
                    Dim destDir As String = Path.GetDirectoryName(destination)
                    NavigateTo(destDir)
                Else
                    ShowStatus(IconWarning & " Destination file already exists.")
                End If

                ' Check if the source is a directory
            ElseIf Directory.Exists(source) Then
                If Not Directory.Exists(destination) Then
                    Directory.Move(source, destination)
                    'ShowStatus("Moved directory to: " & destination)
                    MsgBox("Moved directory to: " & destination)
                    NavigateTo(destination)
                Else
                    ShowStatus(IconWarning & " Destination directory already exists.")
                End If

                ' If neither a file nor a directory exists
            Else
                ShowStatus(IconWarning & " Source not found.")
            End If
        Catch ex As Exception
            ShowStatus(IconError & " Move failed: " & ex.Message)
            Debug.WriteLine("MoveFileOrDirectory Error: " & ex.Message)
        End Try
    End Sub

    Private Sub DeleteFileOrDirectory(path2Delete As String)

        ' Reject relative paths outright
        If Not Path.IsPathRooted(path2Delete) Then

            ShowStatus(IconWarning & " Delete failed: Path must be absolute. Example: C:\folder")

            Exit Sub

        End If

        ' Check if the path is in the protected list
        If IsProtectedPathOrFolder(path2Delete) Then
            ' The path is protected; prevent deletion

            ' Notify the user of the prevention so the user knows why it didn't delete.
            ShowStatus(IconProtect & " Deletion prevented for protected path: " & path2Delete)
            Dim msg As String = "Deletion prevented for protected path: " & Environment.NewLine & path2Delete
            MsgBox(msg, MsgBoxStyle.Critical, "Deletion Prevented")

            NavigateTo(path2Delete)

            Exit Sub

        End If

        Try

            ' Check if it's a file
            If File.Exists(path2Delete) Then

                ' Goto the directory of the file to be deleted.
                ' So the user can see what is about to be deleted.
                Dim destDir As String = IO.Path.GetDirectoryName(path2Delete)
                NavigateTo(destDir)

                ' Make the user confirm deletion.
                Dim confirmMsg As String = "Are you sure you want to delete the file '" &
                                            path2Delete & "'?"
                Dim result = MessageBox.Show(confirmMsg,
                                             "Delete",
                                             MessageBoxButtons.YesNo,
                                             MessageBoxIcon.Warning)
                If result <> DialogResult.Yes Then Exit Sub

                File.Delete(path2Delete)

                ShowStatus(IconDelete & "Deleted file: " & path2Delete)

                ' Go to the directory of the file that was deleted.
                ' So the user can see that it has been deleted.
                NavigateTo(destDir)

                ' Check if it's a directory
            ElseIf Directory.Exists(path2Delete) Then

                ' Go to the directory to be deleted.
                ' So the user can see what is about to be deleted.
                NavigateTo(path2Delete)

                ' Make the user confirm deletion.
                Dim confirmMsg As String = IconDialog & " Are you sure you want to delete the folder '" &
                                           path2Delete & "' and all its contents?"
                ShowStatus(confirmMsg)
                Dim result = MessageBox.Show(confirmMsg,
                                             "Delete",
                                             MessageBoxButtons.YesNo,
                                             MessageBoxIcon.Warning)
                If result <> DialogResult.Yes Then Exit Sub

                Directory.Delete(path2Delete, recursive:=True)

                ' Go to the parent directory of the deleted directory.
                ' So the user can see that it has been deleted.
                Dim parentDir As String = IO.Path.GetDirectoryName(path2Delete)
                NavigateTo(parentDir, True)

                ShowStatus(IconDelete & " Deleted folder: " & path2Delete)

            Else
                ShowStatus(IconWarning & " Delete failed: Path not found.")
            End If

        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message,
                            "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            ShowStatus(IconError & " Delete failed: " & ex.Message)
            Debug.WriteLine("DeleteFileOrDirectory Error: " & ex.Message)
        End Try

    End Sub

    Private Sub CopyDirectory(sourceDir As String, destDir As String)

        Dim dirInfo As New DirectoryInfo(sourceDir)

        If Not dirInfo.Exists Then Throw New DirectoryNotFoundException("Source not found: " & sourceDir)

        Try

            ShowStatus(IconDialog & "Copying directory")

            ' Create destination directory
            Directory.CreateDirectory(destDir)
            ' Copy files
            For Each file In dirInfo.GetFiles()
                Dim targetFilePath = Path.Combine(destDir, file.Name)
                file.CopyTo(targetFilePath, overwrite:=True)
            Next
            ' Copy subdirectories recursively
            For Each subDir In dirInfo.GetDirectories()
                Dim newDest = Path.Combine(destDir, subDir.Name)
                CopyDirectory(subDir.FullName, newDest)
            Next

            'ShowStatus("Copied into " & destDir)

            'MessageBox.Show("Copied into " & destDir, "Copied Folder", MessageBoxButtons.OK, MessageBoxIcon.Information)

            NavigateTo(destDir)

        Catch ex As Exception
            ShowStatus(IconError & "Copy failed: " & ex.Message)
            Debug.WriteLine("CopyDirectory Error: " & ex.Message)
        End Try

    End Sub

    Private Sub RenameFileOrDirectory(sourcePath As String, newName As String)

        Dim newPath As String = Path.Combine(Path.GetDirectoryName(sourcePath), newName)

        ' Rule 1: Path must be absolute (start with C:\ or similar).
        ' Reject relative paths outright
        If Not Path.IsPathRooted(sourcePath) Then

            ShowStatus(IconDialog & " Rename failed: Path must be absolute. Example: C:\folder")

            Exit Sub

        End If

        ' Rule 2: Protected paths are never renamed.
        ' Check if the path is in the protected list
        If IsProtectedPathOrFolder(sourcePath) Then
            ' The path is protected; prevent rename

            ' Notify the user of the prevention so the user knows why it didn't rename.
            'Dim msg As String = "Rename prevented for protected path: " & Environment.NewLine & sourcePath
            'MsgBox(msg, MsgBoxStyle.Exclamation, "Can't Rename")

            ' Show user the directory so they can see it wasn't renamed.
            NavigateTo(sourcePath)

            ' Notify the user of the prevention so the user knows why it didn't rename.
            ShowStatus(IconProtect & "  Rename prevented for protected path or folder: " & sourcePath)

            Exit Sub

        End If

        Try

            ' Rule 3: If it’s a folder, rename the folder and show the new folder.
            ' If source is a directory
            If Directory.Exists(sourcePath) Then

                ' Rename directory
                Directory.Move(sourcePath, newPath)

                ' Navigate to the renamed directory
                NavigateTo(newPath)

                ShowStatus(IconSuccess & " Renamed Folder to: " & newName)

                ' Rule 4: If it’s a file, rename the file and show its folder.
                ' If source is a file
            ElseIf File.Exists(sourcePath) Then

                ' Rename file
                File.Move(sourcePath, newPath)

                ' Navigate to the directory of the renamed file
                NavigateTo(Path.GetDirectoryName(sourcePath))

                ShowStatus(IconSuccess & " Renamed File to: " & newName)

                ' Rule 5: If nothing exists at that path, explain the quoting rule for spaces.
            Else
                ' Path does not exist
                ShowStatus(IconError & " Renamed failed: No path. Paths with spaces must be enclosed in quotes. Example: rename ""[source_path]"" ""[new_name]"" e.g., rename ""C:\folder\old name.txt"" ""new name.txt""")

            End If

        Catch ex As Exception
            ' Rule 6: If anything goes wrong,show a status message.
            ShowStatus(IconError & " Rename failed: " & ex.Message)
            Debug.WriteLine("RenameFileOrDirectory Error: " & ex.Message)
        End Try

    End Sub

    Private Sub SearchInCurrentFolder(searchTerm As String)

        ' Clear item selection for lvFiles
        lvFiles.SelectedItems.Clear()

        Try
            SearchResults = New List(Of String)

            ' Search files
            For Each filePath In Directory.GetFiles(currentFolder, "*", SearchOption.AllDirectories)
                If Path.GetFileName(filePath).IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    SearchResults.Add(filePath)
                End If
            Next

            ' Search directories
            For Each dirPath In Directory.GetDirectories(currentFolder, "*", SearchOption.AllDirectories)
                If Path.GetFileName(dirPath).IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    SearchResults.Add(dirPath)
                End If
            Next

            ' Highlight first match in ListView
            If SearchResults.Count > 0 Then
                SelectListViewItemByPath(SearchResults(0))
                lvFiles.Focus()

                ShowStatus(IconSmile & " Found " & SearchResults.Count & " item(s) matching: " & searchTerm)
            Else
                ShowStatus(IconDialog & " No items found matching: " & searchTerm)
            End If

        Catch ex As Exception
            ShowStatus(IconError & " Search failed: " & ex.Message)
            Debug.WriteLine("SearchInCurrentFolder Error: " & ex.Message)
        End Try

    End Sub


    Private Sub UpdateNavButtons()
        btnBack.Enabled = _historyIndex > 0
        btnForward.Enabled = _historyIndex >= 0 AndAlso _historyIndex < _history.Count - 1
    End Sub

    Private Sub UpdateDestructiveButton()

        ' --- Rule 0: Something must be selected ---
        Dim itemSelected As Boolean = (lvFiles.SelectedItems.Count > 0)
        If Not itemSelected Then
            btnCut.Enabled = False
            btnRename.Enabled = False
            btnDelete.Enabled = False
            Exit Sub
        End If

        Dim fullPath As String = CStr(lvFiles.SelectedItems(0).Tag)

        ' --- Rule 1: Path must exist ---
        If Not PathExists(fullPath) Then
            btnCut.Enabled = False
            btnRename.Enabled = False
            btnDelete.Enabled = False
            Exit Sub
        End If

        ' --- Rule 2: Protected items cannot be renamed ---
        If IsProtectedPathOrFolder(fullPath) Then
            btnCut.Enabled = False
            btnRename.Enabled = False
            btnDelete.Enabled = False
            Exit Sub
        End If

        ' --- Rule 3: User must have rename permission ---
        Dim parentDir As String =
        If(Directory.Exists(fullPath),
           fullPath,                          ' Folder → check write access ON the folder
           Path.GetDirectoryName(fullPath))   ' File → check write access on parent folder

        If Not HasWriteAccessToDirectory(parentDir) Then
            btnRename.Enabled = False
            Exit Sub
        End If

        ' --- All rules passed ---
        btnCut.Enabled = True
        btnRename.Enabled = True
        btnDelete.Enabled = True

    End Sub

    Private Sub UpdateRenameButton()

        ' --- Rule 0: Something must be selected ---
        Dim itemSelected As Boolean = (lvFiles.SelectedItems.Count > 0)
        If Not itemSelected Then
            btnRename.Enabled = False
            Exit Sub
        End If

        Dim fullPath As String = CStr(lvFiles.SelectedItems(0).Tag)

        ' --- Rule 1: Path must exist ---
        If Not PathExists(fullPath) Then
            btnRename.Enabled = False
            Exit Sub
        End If

        ' --- Rule 2: Protected items cannot be renamed ---
        If IsProtectedPathOrFolder(fullPath) Then
            btnRename.Enabled = False
            Exit Sub
        End If

        ' --- Rule 3: User must have rename permission ---
        Dim parentDir As String =
        If(Directory.Exists(fullPath),
           fullPath,                          ' Folder → check write access ON the folder
           Path.GetDirectoryName(fullPath))   ' File → check write access on parent folder

        If Not HasWriteAccessToDirectory(parentDir) Then
            btnRename.Enabled = False
            Exit Sub
        End If

        ' --- All rules passed ---
        btnRename.Enabled = True

    End Sub

    Private Sub UpdateDeleteButton()

        ' --- Rule 0: Something must be selected ---
        Dim itemSelected As Boolean = (lvFiles.SelectedItems.Count > 0)
        If Not itemSelected Then
            btnDelete.Enabled = False
            Exit Sub
        End If

        Dim fullPath As String = CStr(lvFiles.SelectedItems(0).Tag)

        ' --- Rule 1: Path must exist ---
        If Not PathExists(fullPath) Then
            btnDelete.Enabled = False
            Exit Sub
        End If

        ' --- Rule 2: Protected items cannot be renamed ---
        If IsProtectedPathOrFolder(fullPath) Then
            btnDelete.Enabled = False
            Exit Sub
        End If

        ' --- Rule 3: User must have rename permission ---
        Dim parentDir As String =
        If(Directory.Exists(fullPath),
           fullPath,                          ' Folder → check write access ON the folder
           Path.GetDirectoryName(fullPath))   ' File → check write access on parent folder

        If Not HasWriteAccessToDirectory(parentDir) Then
            btnDelete.Enabled = False
            Exit Sub
        End If

        ' --- All rules passed ---
        btnDelete.Enabled = True

    End Sub

    Private Sub UpdateCutButton()

        ' --- Rule 0: Something must be selected ---
        Dim itemSelected As Boolean = (lvFiles.SelectedItems.Count > 0)
        If Not itemSelected Then
            btnCut.Enabled = False
            Exit Sub
        End If

        Dim fullPath As String = CStr(lvFiles.SelectedItems(0).Tag)

        ' --- Rule 1: Path must exist ---
        If Not PathExists(fullPath) Then
            btnCut.Enabled = False
            Exit Sub
        End If

        ' --- Rule 2: Protected items cannot be cut ---
        If IsProtectedPathOrFolder(fullPath) Then
            btnCut.Enabled = False
            Exit Sub
        End If

        ' --- Rule 3: User must have cut permission ---
        Dim parentDir As String =
        If(Directory.Exists(fullPath),
           fullPath,                          ' Folder → check write access ON the folder
           Path.GetDirectoryName(fullPath))   ' File → check write access on parent folder

        If Not HasWriteAccessToDirectory(parentDir) Then
            btnCut.Enabled = False
            Exit Sub
        End If

        ' --- All rules passed ---
        btnCut.Enabled = True

    End Sub

    Private Sub UpdateCopyButton()

        ' --- Rule 0: Something must be selected ---
        Dim itemSelected As Boolean = (lvFiles.SelectedItems.Count > 0)
        If Not itemSelected Then
            btnCopy.Enabled = False
            Exit Sub
        End If

        Dim fullPath As String = CStr(lvFiles.SelectedItems(0).Tag)

        ' --- Rule 1: Path must exist ---
        If Not PathExists(fullPath) Then
            btnCopy.Enabled = False
            Exit Sub
        End If

        ' --- All rules passed ---
        btnCopy.Enabled = True

    End Sub

    Private Sub UpdateEditButtons()

        'UpdateCutButton()

        UpdateCopyButton()

        'UpdateRenameButton()

        'UpdateDeleteButton()

        UpdateDestructiveButton()

        ' Update Paste button
        If _clipboardIsCut Then
            btnPaste.Enabled = True
        Else
            btnPaste.Enabled = Not String.IsNullOrEmpty(_clipboardPath)
        End If

    End Sub

    Private Sub UpdateFileButtons()

        If HasWriteAccessToDirectory(currentFolder) Then
            ' Enable buttons if current folder is valid and writable
            btnNewTextFile.Enabled = True
        Else
            ' Disable buttons if current folder is invalid
            btnNewTextFile.Enabled = False
        End If

        If HasDirectoryCreationAccessToDirectory(currentFolder) Then
            btnNewFolder.Enabled = True
        Else
            btnNewFolder.Enabled = False
        End If

    End Sub


    Private Sub UpdateEditContextMenu()

        If cmsFiles.Items.Count = 0 Then Return

        If _clipboardIsCut Then
            cmsFiles.Items("Paste").Enabled = True
        Else
            cmsFiles.Items("Paste").Enabled = Not String.IsNullOrEmpty(_clipboardPath)
        End If





        If lvFiles.SelectedItems.Count = 0 Then

            ' Context menu items
            cmsFiles.Items("Cut").Enabled = False
            cmsFiles.Items("Copy").Enabled = False
            cmsFiles.Items("Rename").Enabled = False
            cmsFiles.Items("Delete").Enabled = False
            cmsFiles.Items("Open").Enabled = False
            'cmsFiles.Items("CopyName").Enabled = False
            cmsFiles.Items("CopyPath").Enabled = False

            Exit Sub

        End If

        Dim path As String = CStr(lvFiles.SelectedItems(0).Tag)
        Dim exists As Boolean = PathExists(path)

        ' Context menu items
        cmsFiles.Items("Cut").Enabled = exists
        cmsFiles.Items("Copy").Enabled = exists
        cmsFiles.Items("Rename").Enabled = exists
        cmsFiles.Items("Delete").Enabled = exists
        cmsFiles.Items("Open").Enabled = exists
        'cmsFiles.Items("CopyName").Enabled = exists
        cmsFiles.Items("CopyPath").Enabled = exists

    End Sub

    Private Sub UpdateColumnHeaders(sortedColumn As Integer, order As SortOrder)
        For i As Integer = 0 To lvFiles.Columns.Count - 1
            Dim baseText = lvFiles.Columns(i).Text.Replace(" ▲", "").Replace(" ▼", "")
            If i = sortedColumn Then
                lvFiles.Columns(i).Text = baseText & If(order = SortOrder.Ascending, " ▲", " ▼")
            Else
                lvFiles.Columns(i).Text = baseText
            End If
        Next
    End Sub

    Private Sub ResetCutVisuals()
        For Each item As ListViewItem In lvFiles.Items
            item.ForeColor = SystemColors.WindowText
            item.Font = New Font(lvFiles.Font, FontStyle.Regular)
        Next
    End Sub

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


    Private Function PathExists(path As String) As Boolean
        Return File.Exists(path) OrElse Directory.Exists(path)
    End Function

    Private Function HasWriteAccessToDirectory(dirPath As String) As Boolean
        ' Check if we can create, rename, and delete a temporary file in the directory to test write access 

        Dim testFile As String =
        Path.Combine(dirPath, ".__access_test_" & Guid.NewGuid().ToString("N") & ".tmp")

        Dim testFile2 As String = testFile & "_renamed"

        Try
            ' Create
            Using fs As FileStream = File.Create(testFile, 1, FileOptions.None)
            End Using

            ' Rename
            File.Move(testFile, testFile2)

            ' Cleanup
            File.Delete(testFile2)

            Return True

        Catch ex As UnauthorizedAccessException
            Return False
        Catch ex As IOException
            Return False
        End Try

    End Function


    'Private Function HasDirectoryCreationAccessToDirectory(dirPath As String) As Boolean
    '    ' Test whether we can create, rename, and delete a directory inside dirPath.

    '    Dim testDir As String =
    '    Path.Combine(dirPath, ".__dir_access_test_" & Guid.NewGuid().ToString("N"))

    '    Dim testDirRenamed As String = testDir & "_renamed"

    '    Try
    '        ' Create directory
    '        Directory.CreateDirectory(testDir)

    '        ' Rename directory
    '        Directory.Move(testDir, testDirRenamed)

    '        ' Delete directory
    '        Directory.Delete(testDirRenamed)

    '        Return True

    '    Catch ex As UnauthorizedAccessException
    '        Return False
    '    Catch ex As IOException
    '        Return False
    '    End Try

    'End Function



    Private Function HasDirectoryCreationAccessToDirectory(dirPath As String) As Boolean
        ' Tests whether we can create, rename, and delete a directory inside dirPath.

        Dim testDir As String =
        Path.Combine(dirPath, ".__dir_access_test_" & Guid.NewGuid().ToString("N"))

        Dim testDirRenamed As String = testDir & "_renamed"

        Try
            ' Create directory
            Directory.CreateDirectory(testDir)

            ' Rename directory
            Directory.Move(testDir, testDirRenamed)

            ' Delete directory
            Directory.Delete(testDirRenamed, recursive:=False)

            Return True

        Catch ex As UnauthorizedAccessException
            Return False

        Catch ex As IOException
            ' Could be permissions, could be other IO issues.
            ' For access testing, treat as failure.
            Return False

        Finally
            ' Cleanup in case of partial success
            Try
                If Directory.Exists(testDir) Then
                    Directory.Delete(testDir, recursive:=False)
                End If

                If Directory.Exists(testDirRenamed) Then
                    Directory.Delete(testDirRenamed, recursive:=False)
                End If
            Catch
                ' Swallow cleanup errors intentionally.
            End Try
        End Try
    End Function



    Private Function NormalizeTextFilePath(raw As String) As String
        If raw Is Nothing Then Return Nothing

        Dim trimmed = raw.Trim()
        If trimmed.Length = 0 Then Return Nothing

        ' Reject folders — this function is for files only
        If Directory.Exists(trimmed) Then Return Nothing

        ' Auto-append .txt if missing
        If Path.GetExtension(trimmed) = "" Then
            trimmed &= ".txt"
        End If

        Return trimmed
    End Function

    Private Function GetUniqueFilePath(baseDir As String, baseName As String, ext As String) As String
        Dim candidate = Path.Combine(baseDir, baseName & ext)
        Dim counter = 1

        While File.Exists(candidate)
            candidate = Path.Combine(baseDir, $"{baseName} ({counter}){ext}")
            counter += 1
        End While

        Return candidate
    End Function

    Private Sub SelectListViewItemByPath(fullPath As String)
        For Each item As ListViewItem In lvFiles.Items
            If String.Equals(item.Tag.ToString(), fullPath, StringComparison.OrdinalIgnoreCase) Then
                item.Selected = True
                item.Focused = True
                item.EnsureVisible()
                Exit For
            End If
        Next
    End Sub

    Private Function HasSubdirectories(path As String) As Boolean

        Try

            Return Directory.EnumerateDirectories(path).Any()

        Catch ex As Exception

            Debug.WriteLine($"HasSubdirectories: Access denied or error for path: {path}")

            Return False

        End Try

    End Function

    Private Function IsProtectedPathOrFolder(path2Check As String) As Boolean

        If String.IsNullOrWhiteSpace(path2Check) Then Return False

        ' Reject relative paths outright
        If Not Path.IsPathRooted(path2Check) Then Return False

        ' Normalize the input path: full path, trim trailing slashes
        Dim normalizedInput As String = Path.GetFullPath(path2Check).TrimEnd("\"c)

        ' Define protected paths (normalized) - Exact match and subpaths
        Dim protectedPaths As String() = {
            "C:\Windows",
            "C:\Program Files",
            "C:\Program Files (x86)",
            "C:\ProgramData"
        }

        ' Normalize protected paths too
        For Each protectedPath In protectedPaths
            Dim normalizedProtected As String = Path.GetFullPath(protectedPath).TrimEnd("\"c)

            ' Exact match
            If normalizedInput.Equals(normalizedProtected, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            ' Subdirectory match (ensure proper boundary with "\")
            If normalizedInput.StartsWith(normalizedProtected & "\", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

        Next

        ' Define protected folders (normalized) for the current user - Exact match only
        Dim protectedFolders As String() = {
            "C:\Users",
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Documents"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Desktop"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Pictures"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Music"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Videos"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "AppData\Local"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "AppData\Roaming")
        }

        ' Normalize protected folders too
        For Each protectedFolder In protectedFolders
            Dim normalizedProtected As String = Path.GetFullPath(protectedFolder).TrimEnd("\"c)

            ' Exact match
            If normalizedInput.Equals(normalizedProtected, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

        Next

        Return False

    End Function

    Private Function FormatSize(bytes As Long) As String
        Dim absBytes = Math.Abs(bytes)

        For i = SizeUnits.Length - 1 To 0 Step -1
            If absBytes >= SizeUnits(i).Factor Then
                Dim value = absBytes / SizeUnits(i).Factor
                Dim formatted = $"{value:0.##} {SizeUnits(i).Unit}"
                Return If(bytes < 0, "-" & formatted, formatted)
            End If
        Next

        Return $"{bytes} B"
    End Function

    Private Function ParseSize(input As String) As Long
        If String.IsNullOrWhiteSpace(input) Then Return 0

        ' Normalize
        Dim text = input.Trim().Replace(",", "").ToUpperInvariant()

        ' Extract sign
        Dim isNegative As Boolean = text.StartsWith("-")
        If isNegative Then text = text.Substring(1).Trim()

        ' Split into number + unit
        Dim parts = text.Split(" "c, StringSplitOptions.RemoveEmptyEntries)

        Dim numberPart As String = parts(0)
        Dim unitPart As String = If(parts.Length > 1, parts(1), "B")

        ' Parse numeric portion
        Dim value As Double
        If Not Double.TryParse(numberPart, Globalization.NumberStyles.Float,
                               Globalization.CultureInfo.InvariantCulture, value) Then
            Return 0
        End If

        ' Unit multipliers (binary units)
        Dim multipliers As New Dictionary(Of String, Long)(StringComparer.OrdinalIgnoreCase) From {
            {"B", 1L},
            {"KB", 1024L},
            {"MB", 1024L ^ 2},
            {"GB", 1024L ^ 3},
            {"TB", 1024L ^ 4},
            {"PB", 1024L ^ 5}
        }

        ' Resolve multiplier
        Dim factor As Long = 1
        If multipliers.ContainsKey(unitPart) Then
            factor = multipliers(unitPart)
        End If

        ' Compute final byte count
        Dim bytes As Double = value * factor
        Dim result As Long = CLng(Math.Round(bytes))

        Return If(isNegative, -result, result)
    End Function


    Private Sub InitApp()

        Me.Text = "File Explorer - Code with Joe"

        InitStatusBar()

        ShowStatus(IconDialog & " Loading...")

        InitContextMenu()

        InitImageList()

        InitArrows()

        InitTreeRoots()

        InitListView()

        NavigateTo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))

        'InitContextMenu()

        RunTests()

        ShowStatus(IconSuccess & " Ready")

    End Sub

    Private Sub InitListView()
        lvFiles.View = View.Details
        lvFiles.FullRowSelect = True
        lvFiles.MultiSelect = True
        lvFiles.Columns.Clear()
        lvFiles.Columns.Add("Name", 550)
        lvFiles.Columns.Add("Type", 120)
        lvFiles.Columns.Add("Size", 150)
        lvFiles.Columns.Add("Modified", 200)

    End Sub

    Private Sub InitTreeRoots()
        '  Initialize TreeView Roots

        tvFolders.Nodes.Clear()
        tvFolders.ShowRootLines = True

        ' Use arrow icons instead of plus/minus
        tvFolders.ShowPlusMinus = False
        tvFolders.StateImageList = imgArrows   ' Index 0 = ▶, Index 1 = ▼, Index 2 = no arrow

        ' --- Easy Access node ---
        Dim easyAccessNode As New TreeNode("Easy Access") With {
            .ImageKey = "EasyAccess",
            .SelectedImageKey = "EasyAccess",
            .StateImageIndex = 0   ' ▶ collapsed
        }

        ' Define special folders
        Dim specialFolders As (String, Environment.SpecialFolder)() = {
            ("Documents", Environment.SpecialFolder.MyDocuments),
            ("Music", Environment.SpecialFolder.MyMusic),
            ("Pictures", Environment.SpecialFolder.MyPictures),
            ("Videos", Environment.SpecialFolder.MyVideos),
            ("Downloads", Environment.SpecialFolder.UserProfile), ' handled manually
            ("Desktop", Environment.SpecialFolder.Desktop)
        }

        For Each sf In specialFolders
            Dim specialFolderPath As String = Environment.GetFolderPath(sf.Item2)

            ' Handle Downloads manually
            If sf.Item1 = "Downloads" Then
                specialFolderPath = Path.Combine(specialFolderPath, "Downloads")
            End If

            If Directory.Exists(specialFolderPath) Then

                Dim node As New TreeNode(sf.Item1) With {
                    .Tag = specialFolderPath,
                    .ImageKey = sf.Item1,
                    .SelectedImageKey = sf.Item1
                }

                If HasSubdirectories(specialFolderPath) Then
                    node.Nodes.Add("Loading...")
                    node.StateImageIndex = 0   ' ▶ collapsed
                Else
                    node.StateImageIndex = 2  ' no arrow
                End If

                easyAccessNode.Nodes.Add(node)
            End If
        Next

        ' Add Easy Access to tree
        tvFolders.Nodes.Add(easyAccessNode)

        ' Expand Easy Access and update arrow
        easyAccessNode.Expand()
        easyAccessNode.StateImageIndex = 1   ' ▼ expanded

        ' --- Drives ---
        For Each di In DriveInfo.GetDrives()
            If di.IsReady Then
                Try
                    Dim rootNode As New TreeNode(di.Name & " - " & di.VolumeLabel) With {
                        .Tag = di.RootDirectory.FullName
                    }

                    If di.DriveType = DriveType.CDRom Then
                        rootNode.ImageKey = "Optical"
                        rootNode.SelectedImageKey = "Optical"
                    Else
                        rootNode.ImageKey = "Drive"
                        rootNode.SelectedImageKey = "Drive"
                    End If

                    If HasSubdirectories(di.RootDirectory.FullName) Then
                        rootNode.Nodes.Add("Loading...")
                        rootNode.StateImageIndex = 0   ' ▶ collapsed
                    Else
                        rootNode.StateImageIndex = 2  ' no arrow
                    End If

                    tvFolders.Nodes.Add(rootNode)

                Catch ex As IOException
                    Debug.WriteLine($"Error accessing drive {di.Name}: {ex.Message}")
                Catch ex As UnauthorizedAccessException
                    Debug.WriteLine($"Access denied to drive {di.Name}: {ex.Message}")

                Catch ex As Exception
                    Debug.WriteLine($"Unexpected error with drive {di.Name}: {ex.Message}")

                End Try
            Else
                Debug.WriteLine($"Drive {di.Name} is not ready.")
            End If
        Next

    End Sub

    Private Sub InitStatusBar()

        Dim statusStrip As New StatusStrip()

        ' Set font to Segoe UI Symbol, 12pt
        statusStrip.Font = New Font("Segoe UI Symbol", 9.0F, FontStyle.Regular)

        ' If lblStatus should also use this font, set it explicitly
        lblStatus.Font = statusStrip.Font

        statusStrip.Items.Add(lblStatus)

        Me.Controls.Add(statusStrip)

    End Sub

    Private Sub InitImageList()

        imgList.ImageSize = New Size(16, 16)
        imgList.ColorDepth = ColorDepth.Depth32Bit

        ' Load icons 
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
        imgList.Images.Add("Optical", My.Resources.Resource1.Optical_16X16)
        imgList.Images.Add("AccessDenied", My.Resources.Resource1.Access_Denied_16X16)
        imgList.Images.Add("Error", My.Resources.Resource1.Error_16X16)


        ' Assign ImageList to controls
        tvFolders.ImageList = imgList
        lvFiles.SmallImageList = imgList

    End Sub

    Private Sub InitArrows()
        imgArrows.ImageSize = New Size(16, 16)
        imgArrows.ColorDepth = ColorDepth.Depth32Bit
        imgArrows.Images.Add("Collapsed", My.Resources.Resource1.Arrow_Right_16X16)
        imgArrows.Images.Add("Expanded", My.Resources.Resource1.Arrow_Down_16X16)
        imgArrows.Images.Add("NoArrow", My.Resources.Resource1.No_Arrow_16X16)


    End Sub

    Private Sub InitContextMenu()

        cmsFiles.Items.Add("Open", Nothing, AddressOf Open_Click).Name = "Open"

        cmsFiles.Items.Add("New Folder", Nothing, AddressOf NewFolder_Click).Name = "NewFolder"
        cmsFiles.Items.Add("New Text File", Nothing, AddressOf NewTextFile_Click).Name = "NewTextFile"

        cmsFiles.Items.Add("Cut", Nothing, AddressOf CutSelected_Click).Name = "Cut"
        cmsFiles.Items.Add("Copy", Nothing, AddressOf CopySelected_Click).Name = "Copy"
        cmsFiles.Items.Add("Paste", Nothing, AddressOf PasteSelected_Click).Name = "Paste"

        cmsFiles.Items.Add("Rename", Nothing, AddressOf RenameFile_Click).Name = "Rename"
        cmsFiles.Items.Add("Delete", Nothing, AddressOf Delete_Click).Name = "Delete"

        'cmsFiles.Items.Add("Copy Name", Nothing, AddressOf CopyFileName_Click).Name = "CopyName"
        cmsFiles.Items.Add("Copy Path", Nothing, AddressOf CopyFilePath_Click).Name = "CopyPath"

        lvFiles.ContextMenuStrip = cmsFiles

    End Sub


    Private Sub RunTests()

        Debug.WriteLine("Running tests...")

        TestIsProtectedPath()

        TestParseSize()

        TestFormatSize()

        Test_NormalizeTextFilePath()

        Test_GetUniqueFilePath()

        Debug.WriteLine("All tests executed.")

    End Sub

    Private Sub TestIsProtectedPath()

        TestIsProtectedPath_ExactMode()

        TestIsProtectedPath_SubdirMode()

    End Sub

    Private Sub TestIsProtectedPath_ExactMode()

        Debug.WriteLine("→ Testing IsProtectedPathOrFolder (Exact Mode)")

        Dim userProfile As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)

        ' === Positive Tests (Exact Matches, expected True) ===
        AssertTrue(IsProtectedPathOrFolder(userProfile), "User profile should be protected")
        AssertTrue(IsProtectedPathOrFolder(Path.Combine(userProfile, "Documents")), "Documents should be protected")
        AssertTrue(IsProtectedPathOrFolder(Path.Combine(userProfile, "Desktop")), "Desktop should be protected")
        AssertTrue(IsProtectedPathOrFolder(Path.Combine(userProfile, "Downloads")), "Downloads should be protected")

        ' === Negative Tests (Subdirectories should fail in exact-only mode) ===
        AssertFalse(IsProtectedPathOrFolder(Path.Combine(userProfile, "Documents\MyProject")), "Subfolder under Documents should NOT be protected in exact mode")
        AssertFalse(IsProtectedPathOrFolder(Path.Combine(userProfile, "Desktop\TestFolder")), "Subfolder under Desktop should NOT be protected in exact mode")

        ' === Edge Cases ===
        AssertFalse(IsProtectedPathOrFolder(""), "Empty string should not be protected")
        AssertFalse(IsProtectedPathOrFolder(Nothing), "Nothing should not be protected")

        Debug.WriteLine("✓ Exact-mode tests passed")

    End Sub

    Private Sub TestIsProtectedPath_SubdirMode()

        Debug.WriteLine("→ Testing IsProtectedPathOrFolder (Subdirectory Mode)")

        ' === Positive Tests (Exact Matches, expected True) ===
        AssertTrue(IsProtectedPathOrFolder("C:\Windows"), "Windows root should be protected")
        AssertTrue(IsProtectedPathOrFolder("C:\Program Files"), "Program Files root should be protected")

        ' === Positive Tests (Subdirectories now succeed) ===
        AssertTrue(IsProtectedPathOrFolder("C:\Windows\System32"), "System32 should be protected as a subdirectory")
        AssertTrue(IsProtectedPathOrFolder("C:\Program Files\MyApp"), "Subfolder under Program Files should be protected")
        AssertTrue(IsProtectedPathOrFolder("C:\Windows\SysWOW64\WindowsPowerShell\v1.0"),
               "Deep Windows subdirectory should be protected")

        ' === Negative Tests (Unrelated paths still fail) ===
        AssertFalse(IsProtectedPathOrFolder("C:\Temp"), "Temp should NOT be protected")
        AssertFalse(IsProtectedPathOrFolder("D:\Games"), "Unrelated drive should NOT be protected")

        ' === Edge Cases ===
        AssertTrue(IsProtectedPathOrFolder("c:\windows"), "Case-insensitive match should be protected")
        AssertTrue(IsProtectedPathOrFolder("C:\Windows\"), "Trailing slash should still match protected path")

        Debug.WriteLine("✓ Subdirectory-inclusive tests passed")

    End Sub

    Private Sub TestFormatSize()

        Debug.WriteLine("→ Testing FormatSize")

        AssertTrue(FormatSize(0) = "0 B", "0 bytes should format as '0 B'")
        AssertTrue(FormatSize(500) = "500 B", "500 bytes should format as '500 B'")
        AssertTrue(FormatSize(1024) = "1 KB", "1024 bytes should format as '1 KB'")
        AssertTrue(FormatSize(1536) = "1.5 KB", "1536 bytes should format as '1.5 KB'")
        AssertTrue(FormatSize(1048576) = "1 MB", "1,048,576 bytes should format as '1 MB'")
        AssertTrue(FormatSize(1073741824) = "1 GB", "1,073,741,824 bytes should format as '1 GB'")
        AssertTrue(FormatSize(1099511627776) = "1 TB", "1,099,511,627,776 bytes should format as '1 TB'")

        Debug.WriteLine("✓ FormatSize tests passed")

    End Sub

    Private Sub TestParseSize()

        Debug.WriteLine("→ Testing ParseSize")

        Dim cmp As New ListViewItemComparer(0, SortOrder.Ascending)

        AssertTrue(cmp.Test_ParseSize("0 B") = 0, "0 B should parse to 0 bytes")
        AssertTrue(cmp.Test_ParseSize("500 B") = 500, "500 B should parse to 500 bytes")
        AssertTrue(cmp.Test_ParseSize("1 KB") = 1024, "1 KB should parse to 1024 bytes")
        AssertTrue(cmp.Test_ParseSize("1.5 KB") = 1536, "1.5 KB should parse to 1536 bytes")
        AssertTrue(cmp.Test_ParseSize("1 MB") = 1048576, "1 MB should parse to 1,048,576 bytes")
        AssertTrue(cmp.Test_ParseSize("1 GB") = 1073741824, "1 GB should parse to 1,073,741,824 bytes")
        AssertTrue(cmp.Test_ParseSize("1 TB") = 1099511627776, "1 TB should parse to 1,099,511,627,776 bytes")

        Debug.WriteLine("✓ ParseSize tests passed")

    End Sub

    Private Sub Test_NormalizeTextFilePath()

        Debug.WriteLine("→ Testing NormalizeTextFilePath")

        ' === Null / Empty ===
        AssertTrue(NormalizeTextFilePath(Nothing) Is Nothing, "Nothing should return Nothing")
        AssertTrue(NormalizeTextFilePath("") Is Nothing, "Empty string should return Nothing")
        AssertTrue(NormalizeTextFilePath("   ") Is Nothing, "Whitespace should return Nothing")

        ' === Reject folders ===
        Dim tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString())
        Directory.CreateDirectory(tempDir)
        AssertTrue(NormalizeTextFilePath(tempDir) Is Nothing, "Existing directory should return Nothing")
        Directory.Delete(tempDir)

        ' === Auto-append .txt ===
        AssertTrue(NormalizeTextFilePath("C:\Test\Notes") = "C:\Test\Notes.txt",
               "Missing extension should auto-append .txt")

        ' === Preserve existing extension ===
        AssertTrue(NormalizeTextFilePath("C:\Test\Notes.md") = "C:\Test\Notes.md",
               "Existing extension should be preserved")

        ' === Trim whitespace ===
        AssertTrue(NormalizeTextFilePath("   C:\File   ") = "C:\File.txt",
               "Whitespace should be trimmed before processing")

        Debug.WriteLine("✓ NormalizeTextFilePath tests passed")

    End Sub

    Private Sub Test_GetUniqueFilePath()

        Debug.WriteLine("→ Testing GetUniqueFilePath")

        ' Create a temporary directory for isolated testing
        Dim tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString())
        Directory.CreateDirectory(tempDir)

        ' === No conflicts ===
        Dim p1 = GetUniqueFilePath(tempDir, "New Text File", ".txt")
        AssertTrue(p1 = Path.Combine(tempDir, "New Text File.txt"),
               "Should return base name when no conflict exists")

        ' === One conflict ===
        File.WriteAllText(Path.Combine(tempDir, "New Text File.txt"), "")
        Dim p2 = GetUniqueFilePath(tempDir, "New Text File", ".txt")
        AssertTrue(p2 = Path.Combine(tempDir, "New Text File (1).txt"),
               "Should append (1) when base file exists")

        ' === Multiple conflicts ===
        File.WriteAllText(Path.Combine(tempDir, "New Text File (1).txt"), "")
        File.WriteAllText(Path.Combine(tempDir, "New Text File (2).txt"), "")
        Dim p3 = GetUniqueFilePath(tempDir, "New Text File", ".txt")
        AssertTrue(p3 = Path.Combine(tempDir, "New Text File (3).txt"),
               "Should increment until a free name is found")

        ' Cleanup
        Directory.Delete(tempDir, recursive:=True)

        Debug.WriteLine("✓ GetUniqueFilePath tests passed")

    End Sub

    Private Sub TestListViewParser()

        lvFiles.Items.Add(New ListViewItem({"alpha.txt", "Text", "12 KB", "1/2/2024"}))
        lvFiles.Items.Add(New ListViewItem({"beta.txt", "Text", "3,200", "12/25/2023"}))
        lvFiles.Items.Add(New ListViewItem({"gamma.txt", "Text", "1.5 MB", "5/10/2024"}))
        lvFiles.Items.Add(New ListViewItem({"delta.txt", "Text", "1.2 GB", "3/1/2023"}))

    End Sub


    Private Sub AssertTrue(condition As Boolean, message As String)
        Debug.Assert(condition, message)
    End Sub

    Private Sub AssertFalse(condition As Boolean, message As String)
        Debug.Assert(Not condition, message)
    End Sub


End Class

Public Class ListViewItemComparer
    Implements IComparer

    Private ReadOnly _column As Integer
    Private ReadOnly _order As SortOrder
    Private ReadOnly _columnTypes As Dictionary(Of Integer, ColumnDataType)

    Public Enum ColumnDataType
        Text
        Number
        DateValue
    End Enum

    Public Sub New(column As Integer, order As SortOrder,
                   Optional columnTypes As Dictionary(Of Integer, ColumnDataType) = Nothing)

        _column = column
        _order = order
        _columnTypes = If(columnTypes, New Dictionary(Of Integer, ColumnDataType))
    End Sub

    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Dim itemX As ListViewItem = CType(x, ListViewItem)
        Dim itemY As ListViewItem = CType(y, ListViewItem)

        Dim valX As String = itemX.SubItems(_column).Text
        Dim valY As String = itemY.SubItems(_column).Text

        Dim result As Integer

        Select Case GetColumnType(_column)
            Case ColumnDataType.Number
                result = CompareNumbers(valX, valY)

            Case ColumnDataType.DateValue
                result = CompareDates(valX, valY)

            Case Else
                result = String.Compare(valX, valY, StringComparison.OrdinalIgnoreCase)
        End Select

        If _order = SortOrder.Descending Then result = -result
        Return result
    End Function

    Private Function GetColumnType(col As Integer) As ColumnDataType
        If _columnTypes.ContainsKey(col) Then
            Return _columnTypes(col)
        End If
        Return ColumnDataType.Text
    End Function

    ' -----------------------------
    '   NUMBER SORTING (FILE SIZE)
    ' -----------------------------
    Private Function CompareNumbers(a As String, b As String) As Integer
        Dim n1 As Long = ParseSize(a)
        Dim n2 As Long = ParseSize(b)
        Return n1.CompareTo(n2)
    End Function

    ' -----------------------------
    '   DATE SORTING
    ' -----------------------------
    Private Function CompareDates(a As String, b As String) As Integer
        Dim d1, d2 As DateTime
        DateTime.TryParse(a, d1)
        DateTime.TryParse(b, d2)
        Return d1.CompareTo(d2)
    End Function

    ' -----------------------------
    '   PARSE SIZE (SELF-CONTAINED)
    ' -----------------------------
    Private Function ParseSize(input As String) As Long
        If String.IsNullOrWhiteSpace(input) Then Return 0

        ' Normalize
        Dim text = input.Trim().Replace(",", "").ToUpperInvariant()

        ' Extract sign
        Dim isNegative As Boolean = text.StartsWith("-")
        If isNegative Then text = text.Substring(1).Trim()

        ' Split into number + unit
        Dim parts = text.Split(" "c, StringSplitOptions.RemoveEmptyEntries)

        Dim numberPart As String = parts(0)
        Dim unitPart As String = If(parts.Length > 1, parts(1), "B")

        ' Parse numeric portion
        Dim value As Double
        If Not Double.TryParse(numberPart, Globalization.NumberStyles.Float,
                               Globalization.CultureInfo.InvariantCulture, value) Then
            Return 0
        End If

        ' Unit multipliers (binary units)
        Dim multipliers As New Dictionary(Of String, Long)(StringComparer.OrdinalIgnoreCase) From {
            {"B", 1L},
            {"KB", 1024L},
            {"MB", 1024L ^ 2},
            {"GB", 1024L ^ 3},
            {"TB", 1024L ^ 4},
            {"PB", 1024L ^ 5}
        }

        Dim factor As Long = 1
        If multipliers.ContainsKey(unitPart) Then
            factor = multipliers(unitPart)
        End If

        Dim bytes As Double = value * factor
        Dim result As Long = CLng(Math.Round(bytes))

        Return If(isNegative, -result, result)
    End Function

    ' TEMP: expose parser for testing
    Public Function Test_ParseSize(input As String) As Long
        Return ParseSize(input)
    End Function

End Class