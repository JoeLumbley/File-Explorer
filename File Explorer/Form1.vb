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

    Private showHiddenFiles As Boolean = False

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

    Private ErrorChar As Char = "⛔" ' ⛔
    Private DialogChar As String = "" ' 💬
    Private SuccessChar As Char = "" ' ✔
    Private OpenChar As Char = "" ' 📂
    Private CopyChar As Char = ""     ' 📋
    Private PasteChar As Char = ""    ' 📥
    Private ProtectChar As Char = "" ' lock
    Private NavChar As Char = "" ' →
    Private SmileChar As Char = "" ' 😊
    Private WarningChar As Char = "⚠" ' ⚠
    Private DeleteChar As Char = "" ' Trash can
    Private NewFolderChar As Char = "" ' 📁
    Private CutChar As Char = "✂" ' ✂

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        InitApp()

    End Sub

    Private Sub txtPath_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPath.KeyDown

        Path_KeyDown(e)

    End Sub

    Private Sub tvFolders_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvFolders.AfterSelect

        NavigateToSelectedFolderTreeView_AfterSelect(sender, e)

    End Sub

    Private Sub lvFiles_ItemActivate(sender As Object, e As EventArgs) Handles lvFiles.ItemActivate
        ' The ItemActivate event is raised when the user double-clicks an item or
        ' presses the Enter key when an item is selected.

        GoToFolderOrOpenFile_EnterKeyDownOrDoubleClick()

    End Sub

    'Private Sub tvFolders_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles tvFolders.BeforeExpand
    '    ' Load subdirectories on demand

    '    Dim node As TreeNode = e.Node
    '    If node.Nodes.Count = 1 AndAlso node.Nodes(0).Text = "Loading..." Then
    '        node.Nodes.Clear()
    '        Try
    '            For Each dirPath In Directory.GetDirectories(CStr(node.Tag))
    '                Dim di As New DirectoryInfo(dirPath)

    '                ' Skip hidden/system folders unless you want them visible
    '                If (di.Attributes And (FileAttributes.Hidden Or FileAttributes.System)) <> 0 Then
    '                    Continue For
    '                End If

    '                Dim child As New TreeNode(di.Name) With {
    '                .Tag = dirPath,
    '                .ImageKey = "Folder",
    '                .SelectedImageKey = "Folder"
    '                }

    '                ' Add placeholder only if it has subdirectories
    '                If HasSubdirectories(dirPath) Then child.Nodes.Add("Loading...")
    '                node.Nodes.Add(child)
    '            Next
    '        Catch ex As UnauthorizedAccessException
    '            node.Nodes.Add(New TreeNode("[Access denied]") With {.ForeColor = Color.Gray})
    '        Catch ex As IOException
    '            node.Nodes.Add(New TreeNode("[Unavailable]") With {.ForeColor = Color.Gray})
    '        End Try
    '    End If

    'End Sub


    'Private Sub tvFolders_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) _
    'Handles tvFolders.BeforeExpand

    '    e.Node.StateImageIndex = 1   ' ▼ expanded

    '    ' Your existing lazy-load logic...
    '    Dim node As TreeNode = e.Node
    '    If node.Nodes.Count = 1 AndAlso node.Nodes(0).Text = "Loading..." Then
    '        node.Nodes.Clear()
    '        Try
    '            For Each dirPath In Directory.GetDirectories(CStr(node.Tag))
    '                Dim di As New DirectoryInfo(dirPath)

    '                If (di.Attributes And (FileAttributes.Hidden Or FileAttributes.System)) <> 0 Then
    '                    Continue For
    '                End If

    '                Dim child As New TreeNode(di.Name) With {
    '                .Tag = dirPath,
    '                .ImageKey = "Folder",
    '                .SelectedImageKey = "Folder",
    '                .StateImageIndex = 0
    '            }

    '                If HasSubdirectories(dirPath) Then child.Nodes.Add("Loading...")
    '                node.Nodes.Add(child)
    '            Next

    '        Catch ex As UnauthorizedAccessException
    '            node.Nodes.Add(New TreeNode("[Access denied]") With {.ForeColor = Color.Gray})
    '        Catch ex As IOException
    '            node.Nodes.Add(New TreeNode("[Unavailable]") With {.ForeColor = Color.Gray})
    '        End Try
    '    End If
    'End Sub


    'Private Sub tvFolders_BeforeCollapse(sender As Object, e As TreeViewCancelEventArgs) _
    'Handles tvFolders.BeforeCollapse

    '    e.Node.StateImageIndex = 0   ' ▶ collapsed

    'End Sub






    '' ============================
    ''  Expand Node (Lazy Load)
    '' ============================
    'Private Sub tvFolders_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) _
    'Handles tvFolders.BeforeExpand

    '    e.Node.StateImageIndex = 1   ' ▼ expanded

    '    Dim node As TreeNode = e.Node

    '    ' Lazy-load subdirectories
    '    If node.Nodes.Count = 1 AndAlso node.Nodes(0).Text = "Loading..." Then
    '        node.Nodes.Clear()

    '        Try
    '            For Each dirPath In Directory.GetDirectories(CStr(node.Tag))
    '                Dim di As New DirectoryInfo(dirPath)

    '                ' Skip hidden/system folders
    '                If (di.Attributes And (FileAttributes.Hidden Or FileAttributes.System)) <> 0 Then
    '                    Continue For
    '                End If

    '                Dim child As New TreeNode(di.Name) With {
    '                .Tag = dirPath,
    '                .ImageKey = "Folder",
    '                .SelectedImageKey = "Folder",
    '                .StateImageIndex = 0   ' ▶ collapsed
    '            }

    '                If HasSubdirectories(dirPath) Then
    '                    child.Nodes.Add("Loading...")
    '                End If

    '                node.Nodes.Add(child)
    '            Next

    '        Catch ex As UnauthorizedAccessException
    '            node.Nodes.Add(New TreeNode("[Access denied]") With {.ForeColor = Color.Gray})
    '        Catch ex As IOException
    '            node.Nodes.Add(New TreeNode("[Unavailable]") With {.ForeColor = Color.Gray})
    '        End Try
    '    End If

    'End Sub




    ' ============================
    '  Expand Node (Lazy Load)
    ' ============================
    Private Sub tvFolders_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) _
    Handles tvFolders.BeforeExpand

        e.Node.StateImageIndex = 1   ' ▼ expanded

        Dim node As TreeNode = e.Node

        ' Lazy-load subdirectories
        If node.Nodes.Count = 1 AndAlso node.Nodes(0).Text = "Loading..." Then
            node.Nodes.Clear()

            Try
                For Each dirPath In Directory.GetDirectories(CStr(node.Tag))
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
            Catch ex As IOException
                node.Nodes.Add(New TreeNode("[Unavailable]") With {.ForeColor = Color.Gray})
            End Try
        End If

    End Sub


    ' ============================
    '  Collapse Node
    ' ============================
    Private Sub tvFolders_BeforeCollapse(sender As Object, e As TreeViewCancelEventArgs) _
    Handles tvFolders.BeforeCollapse

        e.Node.StateImageIndex = 0   ' ▶ collapsed

    End Sub




















    Private Sub lvFiles_BeforeLabelEdit(sender As Object, e As LabelEditEventArgs) Handles lvFiles.BeforeLabelEdit

        ' Prevent renaming of protected paths
        Dim item As ListViewItem = lvFiles.Items(e.Item)
        Dim fullPath As String = CStr(item.Tag)

        If IsProtectedPathOrFolder(fullPath) Then
            e.CancelEdit = True
            ShowStatus(ProtectChar & " Renaming protected items is not allowed.")
        End If

    End Sub

    Private Sub lvFiles_AfterLabelEdit(sender As Object, e As LabelEditEventArgs) Handles lvFiles.AfterLabelEdit

        RenameFileOrFolder_AfterLabelEdit(e)

    End Sub

    Private Sub lvFiles_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lvFiles.ColumnClick

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

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click

        InitTreeRoots()

        NavigateTo(currentFolder, recordHistory:=False)

    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click

        ExecuteCommand(txtPath.Text.Trim())

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

                ShowStatus(SuccessChar & " Renamed Folder to: " & newName)

            ElseIf File.Exists(oldPath) Then

                File.Move(oldPath, newPath)

                ShowStatus(SuccessChar & " Renamed File to: " & newName)

            End If

            item.Tag = newPath

        Catch ex As Exception

            MessageBox.Show("Rename failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus(ErrorChar & " Rename failed: " & ex.Message)
            e.CancelEdit = True

        End Try

    End Sub

    Private Sub PasteSelected_Click(sender As Object, e As EventArgs)

        ' Is a file or folder selected?
        If String.IsNullOrEmpty(_clipboardPath) Then Exit Sub

        Dim destDir = txtPath.Text
        Dim destPath = Path.Combine(destDir, Path.GetFileName(_clipboardPath))

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
            End If

            ' Clear cut state
            _clipboardPath = Nothing
            _clipboardIsCut = False

            ' Refresh current folder view
            PopulateFiles(destDir)

            ResetCutVisuals()

            ShowStatus(PasteChar & " Pasted into " & txtPath.Text)

        Catch ex As Exception
            MessageBox.Show("Paste failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus(ErrorChar & " Paste failed: " & ex.Message)

        End Try
    End Sub

    Private Sub CopySelected_Click(sender As Object, e As EventArgs)

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        ' Copy the full path stored in the Tag property to clipboard
        Clipboard.SetText(CStr(lvFiles.SelectedItems(0).Tag))

        ' Store in internal clipboard
        _clipboardPath = CStr(lvFiles.SelectedItems(0).Tag)

        _clipboardIsCut = False

        ShowStatus(CopyChar & " Copied to clipboard: " & _clipboardPath)

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

        ShowStatus(CutChar & " Cut to clipboard: " & _clipboardPath)

    End Sub

    'Private Sub NewTextFile_Click(sender As Object, e As EventArgs)

    '    ' Go to the folder of the file to be created.
    '    ' So the user can see where is about to be created.
    '    Dim destDir As String = currentFolder
    '    Dim newFilePath As String = Path.Combine(destDir, "New Text File.txt")

    '    ' Ensure unique name if file already exists
    '    Dim counter As Integer = 1
    '    While File.Exists(newFilePath)
    '        newFilePath = Path.Combine(destDir, $"New Text File ({counter}).txt")
    '        counter += 1
    '    End While

    '    Try

    '        ' Check if the file exists
    '        If Not System.IO.File.Exists(newFilePath) Then

    '            ' Create a new file if it doesn't exist
    '            Using writer As New System.IO.StreamWriter(newFilePath)

    '                writer.WriteLine("Created on " & DateTime.Now.ToString())

    '            End Using

    '            ShowStatus(SuccessChar & " Text file created: " & newFilePath)

    '            ' Go to the folder of the file that was created.
    '            ' So the user can see that it has been created.
    '            NavigateTo(destDir)

    '            ' Open the newly created text file.
    '            GoToFolderOrOpenFile(newFilePath)

    '        End If

    '    Catch ex As Exception
    '        ShowStatus(ErrorChar & " Failed to create text file: " & ex.Message)
    '    End Try

    'End Sub






    Private Sub NewTextFile_Click(sender As Object, e As EventArgs)

        Dim destDir As String = currentFolder

        ' Validate destination folder
        If String.IsNullOrWhiteSpace(destDir) OrElse Not Directory.Exists(destDir) Then
            ShowStatus(WarningChar & " Invalid folder. Cannot create file.")
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

            ShowStatus(SuccessChar & " Text file created: " & newFilePath)

            ' Refresh the folder view so the user sees the new file
            NavigateTo(destDir)

            ' Open the newly created file
            GoToFolderOrOpenFile(newFilePath)

        Catch ex As Exception
            ShowStatus(ErrorChar & " Failed to create text file: " & ex.Message)
        End Try

    End Sub







    Private Sub CopyFileName_Click(sender As Object, e As EventArgs)
        ' Copy selected file name to clipboard - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        Clipboard.SetText(lvFiles.SelectedItems(0).Text)

        ShowStatus(CopyChar & " Copied File Name " & lvFiles.SelectedItems(0).Text)

    End Sub

    Private Sub CopyFilePath_Click(sender As Object, e As EventArgs)
        ' Copy selected file path to clipboard - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        ' Copy the full path stored in the Tag property to clipboard
        Clipboard.SetText(CStr(lvFiles.SelectedItems(0).Tag))

        ShowStatus(CopyChar & " Copied File Path " & lvFiles.SelectedItems(0).Tag)

    End Sub

    Private Sub RenameFile_Click(sender As Object, e As EventArgs)
        ' Rename selected file or folder - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        lvFiles.SelectedItems(0).BeginEdit() ' triggers inline rename

    End Sub

    Private Sub Open_Click(sender As Object, e As EventArgs)
        ' Open selected file or folder - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        Dim fullPath = CStr(lvFiles.SelectedItems(0).Tag)
        GoToFolderOrOpenFile(fullPath)
    End Sub

    Private Sub Delete_Click(sender As Object, e As EventArgs)
        ' Delete selected file or folder - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        Dim item = lvFiles.SelectedItems(0)
        Dim fullPath = CStr(item.Tag)

        ' Reject relative paths outright
        If Not Path.IsPathRooted(fullPath) Then

            ShowStatus(WarningChar & " Delete failed: Path must be absolute. Example: C:\folder")

            Exit Sub

        End If

        ' Check if the path is in the protected list
        If IsProtectedPathOrFolder(fullPath) Then
            ' The path is protected; prevent deletion
            ShowStatus(ProtectChar & " Deletion prevented for protected path: " & fullPath)
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
                ShowStatus(DeleteChar & " Deleted folder: " & item.Text)
                ' Check if it's a file
            ElseIf File.Exists(fullPath) Then
                File.Delete(fullPath)
                lvFiles.Items.Remove(item)
                ShowStatus(DeleteChar & " Deleted file: " & item.Text)
            Else
                ShowStatus(WarningChar & " Path not found.")
            End If
        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus(ErrorChar & " Delete failed: " & ex.Message)
        End Try

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

            ShowStatus(NewFolderChar & " Created folder: " & di.Name)

        Catch ex As Exception
            MessageBox.Show("Failed to create folder: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus(ErrorChar & " Failed to create folder: " & ex.Message)
        End Try

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

    Private Sub ExecuteCommand(command As String)

        ' Use regex to split by spaces but keep quoted substrings together
        Dim parts As String() = Regex.Matches(command, "[\""].+?[\""]|[^ ]+").
        Cast(Of Match)().
        Select(Function(m) m.Value.Trim(""""c)).
        ToArray()

        If parts.Length = 0 Then
            ShowStatus(DialogChar & " No command entered.")
            Return
        End If

        Dim cmd As String = parts(0).ToLower()

        Select Case cmd

            Case "cd"

                If parts.Length > 1 Then
                    Dim newPath As String = String.Join(" ", parts.Skip(1)).Trim()
                    NavigateTo(newPath)
                Else
                    ShowStatus(DialogChar & " Usage: cd [directory] - cd C:\ ")
                End If

            Case "copy"

                If parts.Length > 2 Then

                    Dim source As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()
                    Dim destination As String = parts(parts.Length - 1).Trim()

                    ' Check if source file or directory exists
                    If Not (File.Exists(source) OrElse Directory.Exists(source)) Then
                        ShowStatus(WarningChar & " Copy Failed - Source: ''" & source & "'' does not exist.")
                        Return
                    End If

                    ' Check if destination directory exists
                    If Not Directory.Exists(destination) Then
                        ShowStatus(WarningChar & " Copy Failed - Destination folder: ''" & destination & "'' does not exist.")
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
                    ShowStatus(DialogChar & " Usage: copy [source] [destination] - e.g., copy C:\folder1\file.doc C:\folder2")
                End If


            Case "move"

                If parts.Length > 2 Then

                    Dim source As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()
                    Dim destination As String = parts(parts.Length - 1).Trim()

                    MoveFileOrDirectory(source, destination)

                Else
                    ShowStatus(DialogChar & " Usage: move [source] [destination] - move C:\folder1\directoryToMove C:\folder2\directoryToMove")
                End If

            Case "delete"

                If parts.Length > 1 Then

                    Dim pathToDelete As String = String.Join(" ", parts.Skip(1)).Trim()

                    DeleteFileOrDirectory(pathToDelete)

                Else
                    ShowStatus(DialogChar & " Usage: delete [file_or_directory]")
                End If

            Case "mkdir", "make" ' You can use "mkdir" or "make" as the command
                If parts.Length > 1 Then
                    Dim directoryPath As String = String.Join(" ", parts.Skip(1)).Trim()

                    ' Validate the directory path
                    If String.IsNullOrWhiteSpace(directoryPath) Then
                        ShowStatus(DialogChar & " Usage: mkdir [directory_path] - e.g., mkdir C:\newfolder")
                        Return
                    End If

                    CreateDirectory(directoryPath)

                Else
                    ShowStatus(DialogChar & " Usage: mkdir [directory_path] - e.g., mkdir C:\newfolder")
                End If


            Case "rename"
                ' Rename the file or directory

                If parts.Length > 2 Then
                    Dim sourcePath As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()
                    Dim newName As String = parts(parts.Length - 1).Trim()
                    Dim newPath As String = Path.Combine(Path.GetDirectoryName(sourcePath), newName)

                    ' Reject relative paths outright
                    If Not Path.IsPathRooted(sourcePath) Then

                        ShowStatus(DialogChar & " Rename failed: Path must be absolute. Example: C:\folder")

                        Exit Sub

                    End If

                    ' Check if the path is in the protected list
                    If IsProtectedPathOrFolder(sourcePath) Then
                        ' The path is protected; prevent rename

                        ' Notify the user of the prevention so the user knows why it didn't rename.
                        Dim msg As String = "Rename prevented for protected path: " & Environment.NewLine & sourcePath
                        MsgBox(msg, MsgBoxStyle.Critical, "Rename Prevented")

                        NavigateTo(sourcePath)

                        Exit Sub

                    End If

                    Try
                        ' Validate new name
                        If Directory.Exists(sourcePath) Then
                            Directory.Move(sourcePath, newPath)
                            ShowStatus(SuccessChar & " Renamed Folder to: " & newName)
                        ElseIf File.Exists(sourcePath) Then
                            File.Move(sourcePath, newPath)
                            ShowStatus(SuccessChar & " Renamed File to: " & newName)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Rename failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ShowStatus(ErrorChar & " Rename failed: " & ex.Message)
                    End Try
                Else
                    ShowStatus(DialogChar & " Usage: rename [source_path] [new_name] - e.g., rename C:\folder\oldname.txt newname.txt")
                End If

            Case "text", "txt"

                If parts.Length > 1 Then

                    Dim filePath As String = String.Join(" ", parts.Skip(1)).Trim()

                    ' Validate the file path
                    If String.IsNullOrWhiteSpace(filePath) Then
                        ShowStatus(DialogChar & " Usage: text [file_path] - e.g., text C:\example.txt")
                        Return
                    End If

                    ' Create or open the text file
                    CreateTextFile(filePath)

                Else
                    ShowStatus(DialogChar & " Usage: text [file_path] - e.g., text C:\example.txt")
                End If


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

            Case Else
                If Directory.Exists(command) Then
                    NavigateTo(command)
                ElseIf File.Exists(command) Then
                    GoToFolderOrOpenFile(command)
                Else
                    ShowStatus(DialogChar & " Unknown command: " & cmd)
                End If
        End Select
    End Sub

    Sub CreateTextFile(filePath As String)

        Try

            ' Check if the file exists
            If Not System.IO.File.Exists(filePath) Then

                ' Create a new file if it doesn't exist
                Using writer As New System.IO.StreamWriter(filePath)

                    writer.WriteLine("Created on " & DateTime.Now.ToString())

                    ShowStatus(SuccessChar & " Text file created: " & filePath)

                    Dim destDir As String = Path.GetDirectoryName(filePath)

                    NavigateTo(destDir)

                    GoToFolderOrOpenFile(filePath)

                End Using

            End If

        Catch ex As Exception
            ShowStatus(ErrorChar & " Failed to create text file: " & ex.Message)
        End Try

    End Sub

    Private Sub CreateDirectory(directoryPath As String)

        Try
            ' Create the directory
            Dim dirInfo As DirectoryInfo = Directory.CreateDirectory(directoryPath)

            ShowStatus(NewFolderChar & " Directory created: " & dirInfo.FullName)

            NavigateTo(directoryPath, True)

        Catch ex As Exception
            ShowStatus(ErrorChar & " Failed to create directory: " & ex.Message)
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
                    ShowStatus(WarningChar & " Copy operation canceled.")
                    Return
                End If
            End If

            File.Copy(source, destDirFileName, overwrite:=True)

            ShowStatus(CopyChar & " Copied file: " & fileName & " to: " & destination)

            NavigateTo(destination)

        Catch ex As Exception
            ShowStatus(ErrorChar & " Copy Failed: " & ex.Message)
        End Try
    End Sub

    Private Sub MoveFileOrDirectory(source As String, destination As String)
        Try
            ' Validate parameters
            If String.IsNullOrWhiteSpace(source) OrElse String.IsNullOrWhiteSpace(destination) Then
                ShowStatus(WarningChar & " Source or destination path is invalid.")
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
                    ShowStatus(WarningChar & " Destination file already exists.")
                End If

                ' Check if the source is a directory
            ElseIf Directory.Exists(source) Then
                If Not Directory.Exists(destination) Then
                    Directory.Move(source, destination)
                    'ShowStatus("Moved directory to: " & destination)
                    MsgBox("Moved directory to: " & destination)
                    NavigateTo(destination)
                Else
                    ShowStatus(WarningChar & " Destination directory already exists.")
                End If

                ' If neither a file nor a directory exists
            Else
                ShowStatus(WarningChar & " Source not found.")
            End If
        Catch ex As Exception
            ShowStatus(ErrorChar & " Move failed: " & ex.Message)
        End Try
    End Sub

    Private Sub DeleteFileOrDirectory(path2Delete As String)

        ' Reject relative paths outright
        If Not Path.IsPathRooted(path2Delete) Then

            ShowStatus(WarningChar & " Delete failed: Path must be absolute. Example: C:\folder")

            Exit Sub

        End If

        ' Check if the path is in the protected list
        If IsProtectedPathOrFolder(path2Delete) Then
            ' The path is protected; prevent deletion

            ' Notify the user of the prevention so the user knows why it didn't delete.
            ShowStatus(ProtectChar & " Deletion prevented for protected path: " & path2Delete)
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

                ShowStatus(DeleteChar & "Deleted file: " & path2Delete)

                ' Go to the directory of the file that was deleted.
                ' So the user can see that it has been deleted.
                NavigateTo(destDir)

                ' Check if it's a directory
            ElseIf Directory.Exists(path2Delete) Then

                ' Go to the directory to be deleted.
                ' So the user can see what is about to be deleted.
                NavigateTo(path2Delete)

                ' Make the user confirm deletion.
                Dim confirmMsg As String = DialogChar & " Are you sure you want to delete the folder '" &
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

                ShowStatus(DeleteChar & " Deleted folder: " & path2Delete)

            Else
                ShowStatus(WarningChar & " Delete failed: Path not found.")
            End If

        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message,
                            "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            ShowStatus(ErrorChar & " Delete failed: " & ex.Message)
        End Try

    End Sub


    Private Sub CopyDirectory(sourceDir As String, destDir As String)

        Dim dirInfo As New DirectoryInfo(sourceDir)

        If Not dirInfo.Exists Then Throw New DirectoryNotFoundException("Source not found: " & sourceDir)

        Try

            ShowStatus(DialogChar & "Copying directory")

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
            ShowStatus(ErrorChar & "Copy failed: " & ex.Message)
        End Try

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

    Private Sub InitApp()

        Me.Text = "File Explorer - Code with Joe"

        ShowStatus(DialogChar & " Loading...")

        InitImageList()

        InitListView()

        InitArrows()

        InitTreeRoots()

        NavigateTo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))

        InitStatusBar()

        InitContextMenu()

        RunTests()

        ShowStatus(SuccessChar & " Ready")

    End Sub




    ' -------- UI init --------
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

    'Private Sub InitTreeRoots()

    '    tvFolders.Nodes.Clear()
    '    tvFolders.ShowRootLines = True

    '    tvFolders.ShowPlusMinus = False
    '    tvFolders.StateImageList = imgArrows

    '    ' --- Easy Access node ---
    '    Dim easyAccessNode As New TreeNode("Easy Access") With {
    '    .ImageKey = "EasyAccess",
    '    .SelectedImageKey = "EasyAccess"
    '    }

    '    ' Define special folders
    '    Dim specialFolders As (String, Environment.SpecialFolder)() = {
    '    ("Documents", Environment.SpecialFolder.MyDocuments),
    '    ("Music", Environment.SpecialFolder.MyMusic),
    '    ("Pictures", Environment.SpecialFolder.MyPictures),
    '    ("Videos", Environment.SpecialFolder.MyVideos),
    '    ("Downloads", Environment.SpecialFolder.UserProfile), ' Downloads handled separately
    '    ("Desktop", Environment.SpecialFolder.Desktop)
    '}

    '    For Each sf In specialFolders
    '        Dim specialFolderPath As String = Environment.GetFolderPath(sf.Item2)

    '        ' Handle Downloads manually (UserProfile\Downloads)
    '        If sf.Item1 = "Downloads" Then
    '            specialFolderPath = Path.Combine(specialFolderPath, "Downloads")
    '        End If

    '        If Directory.Exists(specialFolderPath) Then
    '            Dim node = New TreeNode(sf.Item1) With {
    '            .Tag = specialFolderPath,
    '            .ImageKey = sf.Item1,
    '            .SelectedImageKey = sf.Item1
    '        }
    '            node.Nodes.Add("Loading...") ' placeholder for lazy-load
    '            easyAccessNode.Nodes.Add(node)
    '        End If
    '    Next

    '    ' Add Easy Access to tree
    '    tvFolders.Nodes.Add(easyAccessNode)

    '    ' Expand the Easy Access node
    '    easyAccessNode.Expand()

    '    For Each di In DriveInfo.GetDrives()
    '        If di.IsReady Then
    '            Try
    '                Dim rootNode As New TreeNode(di.Name & " - " & di.VolumeLabel) With {
    '            .Tag = di.RootDirectory.FullName
    '        }

    '                ' Check if the drive is an optical drive
    '                If di.DriveType = DriveType.CDRom Then
    '                    rootNode.ImageKey = "Optical"
    '                    rootNode.SelectedImageKey = "Optical"
    '                Else
    '                    rootNode.ImageKey = "Drive"
    '                    rootNode.SelectedImageKey = "Drive"
    '                End If

    '                rootNode.Nodes.Add("Loading...")
    '                tvFolders.Nodes.Add(rootNode)
    '            Catch ex As IOException
    '                ' Handle the exception (e.g., log it or show a message)
    '                Debug.WriteLine($"Error accessing drive {di.Name}: {ex.Message}")
    '            End Try
    '        Else
    '            ' Optionally handle the case where the drive is not ready
    '            Debug.WriteLine($"Drive {di.Name} is not ready.")
    '        End If
    '    Next

    'End Sub


    'Private Sub InitTreeRoots()

    '    tvFolders.Nodes.Clear()
    '    tvFolders.ShowRootLines = True

    '    tvFolders.ShowPlusMinus = False
    '    tvFolders.StateImageList = imgArrows

    '    ' --- Easy Access node ---
    '    Dim easyAccessNode As New TreeNode("Easy Access") With {
    '    .ImageKey = "EasyAccess",
    '    .SelectedImageKey = "EasyAccess",
    '    .StateImageIndex = 0
    '}

    '    ' Define special folders
    '    Dim specialFolders As (String, Environment.SpecialFolder)() = {
    '    ("Documents", Environment.SpecialFolder.MyDocuments),
    '    ("Music", Environment.SpecialFolder.MyMusic),
    '    ("Pictures", Environment.SpecialFolder.MyPictures),
    '    ("Videos", Environment.SpecialFolder.MyVideos),
    '    ("Downloads", Environment.SpecialFolder.UserProfile),
    '    ("Desktop", Environment.SpecialFolder.Desktop)
    '}

    '    For Each sf In specialFolders
    '        Dim specialFolderPath As String = Environment.GetFolderPath(sf.Item2)

    '        If sf.Item1 = "Downloads" Then
    '            specialFolderPath = Path.Combine(specialFolderPath, "Downloads")
    '        End If

    '        If Directory.Exists(specialFolderPath) Then
    '            Dim node = New TreeNode(sf.Item1) With {
    '            .Tag = specialFolderPath,
    '            .ImageKey = sf.Item1,
    '            .SelectedImageKey = sf.Item1,
    '            .StateImageIndex = 0
    '        }

    '            node.Nodes.Add("Loading...")
    '            easyAccessNode.Nodes.Add(node)
    '        End If
    '    Next

    '    tvFolders.Nodes.Add(easyAccessNode)
    '    easyAccessNode.Expand()

    '    For Each di In DriveInfo.GetDrives()
    '        If di.IsReady Then
    '            Try
    '                Dim rootNode As New TreeNode(di.Name & " - " & di.VolumeLabel) With {
    '                .Tag = di.RootDirectory.FullName,
    '                .StateImageIndex = 0
    '            }

    '                If di.DriveType = DriveType.CDRom Then
    '                    rootNode.ImageKey = "Optical"
    '                    rootNode.SelectedImageKey = "Optical"
    '                Else
    '                    rootNode.ImageKey = "Drive"
    '                    rootNode.SelectedImageKey = "Drive"
    '                End If

    '                rootNode.Nodes.Add("Loading...")
    '                tvFolders.Nodes.Add(rootNode)

    '            Catch ex As IOException
    '                Debug.WriteLine($"Error accessing drive {di.Name}: {ex.Message}")
    '            End Try
    '        Else
    '            Debug.WriteLine($"Drive {di.Name} is not ready.")
    '        End If
    '    Next

    'End Sub


    '' ============================
    ''  Initialize TreeView Roots
    '' ============================
    'Private Sub InitTreeRoots()

    '    tvFolders.Nodes.Clear()
    '    tvFolders.ShowRootLines = True

    '    ' Use arrow icons instead of plus/minus
    '    tvFolders.ShowPlusMinus = False
    '    tvFolders.StateImageList = imgArrows   ' Index 0 = ▶, Index 1 = ▼

    '    ' --- Easy Access node ---
    '    Dim easyAccessNode As New TreeNode("Easy Access") With {
    '    .ImageKey = "EasyAccess",
    '    .SelectedImageKey = "EasyAccess",
    '    .StateImageIndex = 0   ' ▶ collapsed
    '}

    '    ' Define special folders
    '    Dim specialFolders As (String, Environment.SpecialFolder)() = {
    '    ("Documents", Environment.SpecialFolder.MyDocuments),
    '    ("Music", Environment.SpecialFolder.MyMusic),
    '    ("Pictures", Environment.SpecialFolder.MyPictures),
    '    ("Videos", Environment.SpecialFolder.MyVideos),
    '    ("Downloads", Environment.SpecialFolder.UserProfile), ' handled manually
    '    ("Desktop", Environment.SpecialFolder.Desktop)
    '}

    '    For Each sf In specialFolders
    '        Dim specialFolderPath As String = Environment.GetFolderPath(sf.Item2)

    '        ' Handle Downloads manually
    '        If sf.Item1 = "Downloads" Then
    '            specialFolderPath = Path.Combine(specialFolderPath, "Downloads")
    '        End If

    '        If Directory.Exists(specialFolderPath) Then
    '            Dim node As New TreeNode(sf.Item1) With {
    '            .Tag = specialFolderPath,
    '            .ImageKey = sf.Item1,
    '            .SelectedImageKey = sf.Item1,
    '            .StateImageIndex = 0   ' ▶ collapsed
    '        }

    '            node.Nodes.Add("Loading...") ' placeholder for lazy-load
    '            easyAccessNode.Nodes.Add(node)
    '        End If
    '    Next

    '    ' Add Easy Access to tree
    '    tvFolders.Nodes.Add(easyAccessNode)

    '    ' Expand Easy Access and update arrow
    '    easyAccessNode.Expand()
    '    easyAccessNode.StateImageIndex = 1   ' ▼ expanded

    '    ' --- Drives ---
    '    For Each di In DriveInfo.GetDrives()
    '        If di.IsReady Then
    '            Try
    '                Dim rootNode As New TreeNode(di.Name & " - " & di.VolumeLabel) With {
    '                .Tag = di.RootDirectory.FullName,
    '                .StateImageIndex = 0   ' ▶ collapsed
    '            }

    '                If di.DriveType = DriveType.CDRom Then
    '                    rootNode.ImageKey = "Optical"
    '                    rootNode.SelectedImageKey = "Optical"
    '                Else
    '                    rootNode.ImageKey = "Drive"
    '                    rootNode.SelectedImageKey = "Drive"
    '                End If

    '                rootNode.Nodes.Add("Loading...")
    '                tvFolders.Nodes.Add(rootNode)

    '            Catch ex As IOException
    '                Debug.WriteLine($"Error accessing drive {di.Name}: {ex.Message}")
    '            End Try
    '        Else
    '            Debug.WriteLine($"Drive {di.Name} is not ready.")
    '        End If
    '    Next

    'End Sub



    ' ============================
    '  Initialize TreeView Roots
    ' ============================
    Private Sub InitTreeRoots()

        tvFolders.Nodes.Clear()
        tvFolders.ShowRootLines = True

        ' Use arrow icons instead of plus/minus
        tvFolders.ShowPlusMinus = False
        tvFolders.StateImageList = imgArrows   ' Index 0 = ▶, Index 1 = ▼

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
                End Try
            Else
                Debug.WriteLine($"Drive {di.Name} is not ready.")
            End If
        Next

    End Sub



    'Private Sub InitStatusBar()
    '    Dim statusStrip As New StatusStrip()
    '    ' Change font to Segoe UI Symbol Regular 12 pt


    '    statusStrip.Items.Add(lblStatus)
    '    Me.Controls.Add(statusStrip)
    'End Sub


    Private Sub InitStatusBar()

        Dim statusStrip As New StatusStrip()

        ' Set font to Segoe UI Symbol, 12pt
        statusStrip.Font = New Font("Segoe UI Symbol", 10.0F, FontStyle.Regular)

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
        ' Context menu setup

        cmsFiles.Items.Add("Cut", Nothing, AddressOf CutSelected_Click)
        cmsFiles.Items.Add("Copy", Nothing, AddressOf CopySelected_Click)
        cmsFiles.Items.Add("Paste", Nothing, AddressOf PasteSelected_Click)

        cmsFiles.Items.Add("Open", Nothing, AddressOf Open_Click)


        cmsFiles.Items.Add("Copy Name", Nothing, AddressOf CopyFileName_Click)
        cmsFiles.Items.Add("Copy Path", Nothing, AddressOf CopyFilePath_Click)

        cmsFiles.Items.Add("New Folder", Nothing, AddressOf NewFolder_Click)
        cmsFiles.Items.Add("New Text File", Nothing, AddressOf NewTextFile_Click) ' ✅ new item

        cmsFiles.Items.Add("Rename", Nothing, AddressOf RenameFile_Click)
        cmsFiles.Items.Add("Delete", Nothing, AddressOf Delete_Click)

        lvFiles.ContextMenuStrip = cmsFiles

    End Sub

    ' -------- Navigation --------
    Private Sub NavigateTo(path As String, Optional recordHistory As Boolean = True)
        If String.IsNullOrWhiteSpace(path) Then Exit Sub
        If Not Directory.Exists(path) Then
            MessageBox.Show("Folder not found: " & path, "Navigation", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ShowStatus(WarningChar & " Folder not found: " & path)
            Exit Sub
        End If

        ShowStatus(NavChar & " Navigated To: " & path)

        currentFolder = path
        txtPath.Text = path
        PopulateFiles(path)
        'EnsureTreeSelection(path)

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
                ShowStatus(WarningChar & " Drive is not ready and has been removed.")
                Return
            End If

        Catch ex As Exception
            ' Handle any exceptions when accessing DriveInfo
            ShowStatus(ErrorChar & " NavTree: Error accessing drive: " & ex.Message)
            Return
        End Try

        ' If the drive is ready, navigate to the folder
        NavigateTo(path2Nav)

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

                ShowStatus(OpenChar & " Opened " & Path)

            Catch ex As Exception
                ShowStatus(ErrorChar & " Cannot open: " & ex.Message)
            End Try

        Else
            ShowStatus(WarningChar & " Path does not exist: " & Path)
        End If

    End Sub

    Private Function HasSubdirectories(path As String) As Boolean
        Try
            Return Directory.EnumerateDirectories(path).Any()
        Catch
            Return False
        End Try
    End Function

    Private Sub PopulateFiles(path As String)
        lvFiles.BeginUpdate()
        lvFiles.Items.Clear()

        ' Folders first
        Try
            For Each mDir In Directory.GetDirectories(path)
                Dim di = New DirectoryInfo(mDir)

                ' Skip hidden/system folders unless checkbox is checked
                If Not showHiddenFiles AndAlso
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
            lvFiles.Items.Add(New ListViewItem("[Access denied]") With {.ForeColor = Color.Gray})
        End Try

        ' Files
        Try
            For Each file In Directory.GetFiles(path)
                Dim fi = New FileInfo(file)

                ' Skip hidden/system files unless checkbox is checked
                If Not showHiddenFiles AndAlso
               (fi.Attributes And (FileAttributes.Hidden Or FileAttributes.System)) <> 0 Then
                    Continue For
                End If

                Dim item = New ListViewItem(fi.Name)
                item.SubItems.Add(fi.Extension.ToLowerInvariant())
                item.SubItems.Add(FormatSize(fi.Length))
                item.SubItems.Add(fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
                item.Tag = fi.FullName

                ' Assign image based on file type (same as before)
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
            lvFiles.Items.Add(New ListViewItem("[Access denied]") With {.ForeColor = Color.Gray})
        End Try

        lvFiles.EndUpdate()
    End Sub

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

    Private Sub RunTests()

        Debug.WriteLine("Running tests...")

        TestIsProtectedPath_ExactMode()

        TestIsProtectedPath_SubdirMode()

        TestParseSize()

        TestFormatSize()

        'TestListViewParser()

        Debug.WriteLine("All tests executed.")

    End Sub

    Private Sub TestIsProtectedPath_ExactMode()

        Dim userProfile As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)

        ' === Positive Tests (Exact Matches, expected True) ===
        Debug.Assert(IsProtectedPathOrFolder(userProfile) = True)

        Debug.Assert(IsProtectedPathOrFolder(Path.Combine(userProfile, "Documents")) = True)
        Debug.Assert(IsProtectedPathOrFolder(Path.Combine(userProfile, "Desktop")) = True)
        Debug.Assert(IsProtectedPathOrFolder(Path.Combine(userProfile, "Downloads")) = True)

        ' === Negative Tests (Subdirectories should fail in exact-only mode) ===
        Debug.Assert(IsProtectedPathOrFolder(Path.Combine(userProfile, "Documents\MyProject")) = False)
        Debug.Assert(IsProtectedPathOrFolder(Path.Combine(userProfile, "Desktop\TestFolder")) = False)

        ' === Edge Cases ===
        Debug.Assert(IsProtectedPathOrFolder("") = False)
        Debug.Assert(IsProtectedPathOrFolder(Nothing) = False)

        Debug.WriteLine("Exact-mode tests executed.")

    End Sub

    Private Sub TestIsProtectedPath_SubdirMode()

        ' === Positive Tests (Exact Matches, expected True) ===
        Debug.Assert(IsProtectedPathOrFolder("C:\Windows") = True)
        Debug.Assert(IsProtectedPathOrFolder("C:\Program Files") = True)

        ' === Positive Tests (Subdirectories now succeed) ===
        Debug.Assert(IsProtectedPathOrFolder("C:\Windows\System32") = True)
        Debug.Assert(IsProtectedPathOrFolder("C:\Program Files\MyApp") = True)
        Debug.Assert(IsProtectedPathOrFolder("C:\Windows\SysWOW64\WindowsPowerShell\v1.0") = True)

        ' === Negative Tests (Unrelated paths still fail) ===
        Debug.Assert(IsProtectedPathOrFolder("C:\Temp") = False)
        Debug.Assert(IsProtectedPathOrFolder("D:\Games") = False)

        ' === Edge Cases ===
        Debug.Assert(IsProtectedPathOrFolder("c:\windows") = True)   ' Case-insensitive
        Debug.Assert(IsProtectedPathOrFolder("C:\Windows\") = True)  ' Trailing slash

        Debug.WriteLine("Subdirectory-inclusive tests executed.")

    End Sub

    Private Sub TestFormatSize()

        ' Test FormatSize
        Debug.Assert(FormatSize(0) = "0 B")
        Debug.Assert(FormatSize(500) = "500 B")
        Debug.Assert(FormatSize(1024) = "1 KB")
        Debug.Assert(FormatSize(1536) = "1.5 KB")
        Debug.Assert(FormatSize(1048576) = "1 MB")
        Debug.Assert(FormatSize(1073741824) = "1 GB")
        Debug.Assert(FormatSize(1099511627776) = "1 TB")

        Debug.WriteLine("FormatSize tests executed.")

    End Sub

    Private Sub TestParseSize()

        ' Test ParseSize (using exposed test wrapper)
        Dim cmp As New ListViewItemComparer(0, SortOrder.Ascending)

        Debug.Assert(cmp.Test_ParseSize("0 B") = 0)
        Debug.Assert(cmp.Test_ParseSize("500 B") = 500)
        Debug.Assert(cmp.Test_ParseSize("1 KB") = 1024)
        Debug.Assert(cmp.Test_ParseSize("1.5 KB") = 1536)
        Debug.Assert(cmp.Test_ParseSize("1 MB") = 1048576)
        Debug.Assert(cmp.Test_ParseSize("1 GB") = 1073741824)
        Debug.Assert(cmp.Test_ParseSize("1 TB") = 1099511627776)

        Debug.WriteLine("ParseSize tests executed.")

    End Sub

    Private Sub TestListViewParser()

        lvFiles.Items.Add(New ListViewItem({"alpha.txt", "Text", "12 KB", "1/2/2024"}))
        lvFiles.Items.Add(New ListViewItem({"beta.txt", "Text", "3,200", "12/25/2023"}))
        lvFiles.Items.Add(New ListViewItem({"gamma.txt", "Text", "1.5 MB", "5/10/2024"}))
        lvFiles.Items.Add(New ListViewItem({"delta.txt", "Text", "1.2 GB", "3/1/2023"}))

    End Sub

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