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

    Private statusTimer As New Timer() With {.Interval = 10000}

    Private currentFolder As String = String.Empty

    Private showHiddenFiles As Boolean = False

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        InitApp()

    End Sub

    Private Sub txtPath_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPath.KeyDown

        Path_KeyDown(e)

    End Sub

    Private Sub tvFolders_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvFolders.AfterSelect

        Dim node As TreeNode = e.Node

        If node Is Nothing Then Exit Sub

        NavigateTo(CStr(node.Tag))

    End Sub

    Private Sub lvFiles_ItemActivate(sender As Object, e As EventArgs) Handles lvFiles.ItemActivate

        'GoToFolderOrOpenFile_DoubleClick(sender, e)
        ' -------- Open on double-click --------

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        Dim sel As ListViewItem = lvFiles.SelectedItems(0)

        Dim fullPath As String = CStr(sel.Tag)

        GoToFolderOrOpenFile(fullPath)

    End Sub

    Private Sub tvFolders_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles tvFolders.BeforeExpand
        ' Load subdirectories on demand

        Dim node As TreeNode = e.Node
        If node.Nodes.Count = 1 AndAlso node.Nodes(0).Text = "Loading..." Then
            node.Nodes.Clear()
            Try
                For Each dirPath In Directory.GetDirectories(CStr(node.Tag))
                    Dim di As New DirectoryInfo(dirPath)

                    ' Skip hidden/system folders unless you want them visible
                    If (di.Attributes And (FileAttributes.Hidden Or FileAttributes.System)) <> 0 Then
                        Continue For
                    End If

                    Dim child As New TreeNode(di.Name) With {
                    .Tag = dirPath,
                    .ImageKey = "Folder",
                    .SelectedImageKey = "Folder"
                }

                    ' Add placeholder only if it has subdirectories
                    If HasSubdirectories(dirPath) Then child.Nodes.Add("Loading...")
                    node.Nodes.Add(child)
                Next
            Catch ex As UnauthorizedAccessException
                node.Nodes.Add(New TreeNode("[Access denied]") With {.ForeColor = Color.Gray})
            Catch ex As IOException
                node.Nodes.Add(New TreeNode("[Unavailable]") With {.ForeColor = Color.Gray})
            End Try
        End If

    End Sub


    Private Sub lvFiles_AfterLabelEdit(sender As Object, e As LabelEditEventArgs) Handles lvFiles.AfterLabelEdit
        ' -------- Rename file or folder --------

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

    Private Sub lvFiles_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lvFiles.ColumnClick
        ' -------- Sort by column --------

        ' Toggle between ascending and descending
        Dim sortOrder As SortOrder = If(lvFiles.Sorting = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)

        ' Sort the ListView
        lvFiles.Sorting = sortOrder
        lvFiles.ListViewItemSorter = New ListViewItemComparer(e.Column, sortOrder)
        lvFiles.Sort()

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

            ShowStatus("Pasted into " & txtPath.Text)

        Catch ex As Exception
            MessageBox.Show("Paste failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus("Paste failed: " & ex.Message)

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

        ShowStatus("Copied to clipboard: " & _clipboardPath)

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

        ShowStatus("Cut to clipboard: " & _clipboardPath)

    End Sub

    Private Sub NewTextFile_Click(sender As Object, e As EventArgs)

        ' Go to the folder of the file to be created.
        ' So the user can see where is about to be created.
        Dim destDir As String = currentFolder
        Dim newFilePath As String = Path.Combine(destDir, "New Text File.txt")

        ' Ensure unique name if file already exists
        Dim counter As Integer = 1
        While File.Exists(newFilePath)
            newFilePath = Path.Combine(destDir, $"New Text File ({counter}).txt")
            counter += 1
        End While

        Try

            ' Check if the file exists
            If Not System.IO.File.Exists(newFilePath) Then

                ' Create a new file if it doesn't exist
                Using writer As New System.IO.StreamWriter(newFilePath)

                    writer.WriteLine("Created on " & DateTime.Now.ToString())

                End Using

                ShowStatus("Text file created: " & newFilePath)

                ' Go to the folder of the file that was created.
                ' So the user can see that it has been created.
                NavigateTo(destDir)

                ' Open the newly created text file.
                GoToFolderOrOpenFile(newFilePath)

            End If

        Catch ex As Exception
            ShowStatus("Failed to create text file: " & ex.Message)
        End Try

    End Sub

    Private Sub CopyFileName_Click(sender As Object, e As EventArgs)
        ' Copy selected file name to clipboard - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        Clipboard.SetText(lvFiles.SelectedItems(0).Text)

        ShowStatus("Copied File Name " & lvFiles.SelectedItems(0).Text)

    End Sub

    Private Sub CopyFilePath_Click(sender As Object, e As EventArgs)
        ' Copy selected file path to clipboard - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        ' Copy the full path stored in the Tag property to clipboard
        Clipboard.SetText(CStr(lvFiles.SelectedItems(0).Tag))

        ShowStatus("Copied File Path " & lvFiles.SelectedItems(0).Tag)

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

            ShowStatus("Delete failed: Path must be absolute. Example: C:\folder")

            Exit Sub

        End If

        ' Check if the path is in the protected list
        If IsProtectedPath(fullPath) Then
            ' The path is protected; prevent deletion
            ShowStatus("Deletion prevented for protected path: " & fullPath)
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
                ShowStatus("Deleted folder: " & item.Text)
                ' Check if it's a file
            ElseIf File.Exists(fullPath) Then
                File.Delete(fullPath)
                lvFiles.Items.Remove(item)
                ShowStatus("Deleted file: " & item.Text)
            Else
                ShowStatus("Path not found.")
            End If
        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus("Delete failed: " & ex.Message)
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

            ShowStatus("Created folder: " & di.Name)

        Catch ex As Exception
            MessageBox.Show("Failed to create folder: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ShowStatus("Failed to create folder: " & ex.Message)
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
            ShowStatus("No command entered.")
            Return
        End If

        Dim cmd As String = parts(0).ToLower()

        Select Case cmd

            Case "cd"

                If parts.Length > 1 Then
                    Dim newPath As String = String.Join(" ", parts.Skip(1)).Trim()
                    NavigateTo(newPath)
                Else
                    ShowStatus("Usage: cd [directory] - cd C:\ ")
                End If

            Case "copy"

                If parts.Length > 2 Then

                    Dim source As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()
                    Dim destination As String = parts(parts.Length - 1).Trim()

                    ' Check if source file or directory exists
                    If Not (File.Exists(source) OrElse Directory.Exists(source)) Then
                        ShowStatus("Copy Failed - Source: ''" & source & "'' does not exist.")
                        Return
                    End If

                    ' Check if destination directory exists
                    If Not Directory.Exists(destination) Then
                        ShowStatus("Copy Failed - Destination folder: ''" & destination & "'' does not exist.")
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
                    ShowStatus("Usage: copy [source] [destination] - e.g., copy C:\folder1\file.doc C:\folder2")
                End If


            Case "move"

                If parts.Length > 2 Then

                    Dim source As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()
                    Dim destination As String = parts(parts.Length - 1).Trim()

                    MoveFileOrDirectory(source, destination)

                Else
                    ShowStatus("Usage: move [source] [destination] - move C:\folder1\directoryToMove C:\folder2\directoryToMove")
                End If

            Case "delete"

                If parts.Length > 1 Then

                    Dim pathToDelete As String = String.Join(" ", parts.Skip(1)).Trim()

                    DeleteFileOrDirectory(pathToDelete)

                Else
                    ShowStatus("Usage: delete [file_or_directory]")
                End If

            Case "mkdir", "make" ' You can use "mkdir" or "make" as the command
                If parts.Length > 1 Then
                    Dim directoryPath As String = String.Join(" ", parts.Skip(1)).Trim()

                    ' Validate the directory path
                    If String.IsNullOrWhiteSpace(directoryPath) Then
                        ShowStatus("Usage: mkdir [directory_path] - e.g., mkdir C:\newfolder")
                        Return
                    End If

                    CreateDirectory(directoryPath)

                Else
                    ShowStatus("Usage: mkdir [directory_path] - e.g., mkdir C:\newfolder")
                End If

            Case "text", "txt"

                If parts.Length > 1 Then

                    Dim filePath As String = String.Join(" ", parts.Skip(1)).Trim()

                    ' Validate the file path
                    If String.IsNullOrWhiteSpace(filePath) Then
                        ShowStatus("Usage: text [file_path] - e.g., text C:\example.txt")
                        Return
                    End If

                    ' Create or open the text file
                    CreateTextFile(filePath)

                Else
                    ShowStatus("Usage: text [file_path] - e.g., text C:\example.txt")
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
                    ShowStatus("Unknown command: " & cmd)
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

                    ShowStatus("Text file created: " & filePath)

                    Dim destDir As String = Path.GetDirectoryName(filePath)

                    NavigateTo(destDir)

                    GoToFolderOrOpenFile(filePath)

                End Using

            End If

        Catch ex As Exception
            ShowStatus("Failed to create text file: " & ex.Message)
        End Try

    End Sub

    Private Sub CreateDirectory(directoryPath As String)

        Try
            ' Create the directory
            Dim dirInfo As DirectoryInfo = Directory.CreateDirectory(directoryPath)

            ShowStatus("Directory created: " & dirInfo.FullName)

            NavigateTo(directoryPath, True)

        Catch ex As Exception
            ShowStatus("Failed to create directory: " & ex.Message)
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
                    ShowStatus("Copy operation canceled.")
                    Return
                End If
            End If

            File.Copy(source, destDirFileName, overwrite:=True)

            ShowStatus("Copied file: " & fileName & " to: " & destination)

            NavigateTo(destination)

        Catch ex As Exception
            ShowStatus("Copy Failed: " & ex.Message)
        End Try
    End Sub

    Private Sub MoveFileOrDirectory(source As String, destination As String)
        Try
            ' Validate parameters
            If String.IsNullOrWhiteSpace(source) OrElse String.IsNullOrWhiteSpace(destination) Then
                ShowStatus("Source or destination path is invalid.")
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
                    ShowStatus("Destination file already exists.")
                End If

                ' Check if the source is a directory
            ElseIf Directory.Exists(source) Then
                If Not Directory.Exists(destination) Then
                    Directory.Move(source, destination)
                    'ShowStatus("Moved directory to: " & destination)
                    MsgBox("Moved directory to: " & destination)
                    NavigateTo(destination)
                Else
                    ShowStatus("Destination directory already exists.")
                End If

                ' If neither a file nor a directory exists
            Else
                ShowStatus("Source not found.")
            End If
        Catch ex As Exception
            ShowStatus("Move failed: " & ex.Message)
        End Try
    End Sub

    Private Sub DeleteFileOrDirectory(path2Delete As String)

        ' Reject relative paths outright
        If Not Path.IsPathRooted(path2Delete) Then

            ShowStatus("Delete failed: Path must be absolute. Example: C:\folder")

            Exit Sub

        End If

        ' Check if the path is in the protected list
        If IsProtectedPath(path2Delete) Then
            ' The path is protected; prevent deletion

            ' Notify the user of the prevention so the user knows why it didn't delete.
            ShowStatus("Deletion prevented for protected path: " & path2Delete)
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

                ShowStatus("Deleted file: " & path2Delete)

                ' Go to the directory of the file that was deleted.
                ' So the user can see that it has been deleted.
                NavigateTo(destDir)

                ' Check if it's a directory
            ElseIf Directory.Exists(path2Delete) Then

                ' Go to the directory to be deleted.
                ' So the user can see what is about to be deleted.
                NavigateTo(path2Delete)

                ' Make the user confirm deletion.
                Dim confirmMsg As String = "Are you sure you want to delete the folder '" &
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

                ShowStatus("Deleted folder: " & path2Delete)

            Else
                ShowStatus("Delete failed: Path not found.")
            End If

        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message,
                            "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            ShowStatus("Delete failed: " & ex.Message)
        End Try

    End Sub


    Private Sub CopyDirectory(sourceDir As String, destDir As String)

        Dim dirInfo As New DirectoryInfo(sourceDir)

        If Not dirInfo.Exists Then Throw New DirectoryNotFoundException("Source not found: " & sourceDir)

        Try

            ShowStatus("Copying directory")

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
            ShowStatus("Copy failed: " & ex.Message)
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

        ShowStatus("Loading...")

        InitImageList()

        InitListView()

        InitTreeRoots()

        NavigateTo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))

        InitStatusBar()

        InitContextMenu()

        RunTests()

        ShowStatus("Ready")

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
            ShowStatus("Folder not found: " & path)
            Exit Sub
        End If

        ShowStatus("Navigated To: " & path)

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

    Private Function IsProtectedPath(path2Check As String) As Boolean

        If String.IsNullOrWhiteSpace(path2Check) Then Return False

        ' Reject relative paths outright
        If Not Path.IsPathRooted(path2Check) Then Return False

        ' Normalize the input path: full path, trim trailing slashes
        Dim normalizedInput As String = Path.GetFullPath(path2Check).TrimEnd("\"c)

        ' Define protected paths (normalized)
        Dim protectedPaths As String() = {
        "C:\Windows",
        "C:\Program Files",
        "C:\Program Files (x86)",
        "C:\ProgramData",
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Documents"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Desktop"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Pictures"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Music"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Videos"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "AppData\Local"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "AppData\Roaming")
    }

        ' Normalize protected paths too
        For Each protectedPath In protectedPaths
            Dim normalizedProtected As String = Path.GetFullPath(protectedPath).TrimEnd("\"c)

            ' Exact match OR subdirectory match
            If normalizedInput.Equals(normalizedProtected, StringComparison.OrdinalIgnoreCase) _
           OrElse normalizedInput.StartsWith(normalizedProtected & "\", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Return False

    End Function

    Private Sub RunTests()

        Debug.WriteLine("Running tests...")

        TestIsProtectedPath()

        Debug.WriteLine("All tests executed.")

    End Sub

    Private Sub TestIsProtectedPath()

        ' === Positive Tests (expected True) ===
        Debug.Assert(IsProtectedPath("C:\Windows") = True, "Should detect Windows folder")
        Debug.Assert(IsProtectedPath("C:\Program Files") = True, "Should detect Program Files")
        Debug.Assert(IsProtectedPath("C:\Program Files (x86)") = True, "Should detect Program Files (x86)")
        Debug.Assert(IsProtectedPath("C:\ProgramData") = True, "Should detect ProgramData")

        Dim userProfile As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        Debug.Assert(IsProtectedPath(Path.Combine(userProfile, "Documents")) = True, "Should detect Documents")
        Debug.Assert(IsProtectedPath(Path.Combine(userProfile, "Desktop")) = True, "Should detect Desktop")
        Debug.Assert(IsProtectedPath(Path.Combine(userProfile, "Downloads")) = True, "Should detect Downloads")
        Debug.Assert(IsProtectedPath(Path.Combine(userProfile, "Pictures")) = True, "Should detect Pictures")
        Debug.Assert(IsProtectedPath(Path.Combine(userProfile, "Music")) = True, "Should detect Music")
        Debug.Assert(IsProtectedPath(Path.Combine(userProfile, "Videos")) = True, "Should detect Videos")
        Debug.Assert(IsProtectedPath(Path.Combine(userProfile, "AppData\Local")) = True, "Should detect Local AppData")
        Debug.Assert(IsProtectedPath(Path.Combine(userProfile, "AppData\Roaming")) = True, "Should detect Roaming AppData")

        ' === New Positive Tests for Subdirectories (expected True) ===
        Debug.Assert(IsProtectedPath("C:\Windows\System32") = True, "Subfolder of Windows should be protected")
        Debug.Assert(IsProtectedPath("C:\Program Files\MyApp") = True, "Subfolder of Program Files should be protected")
        Debug.Assert(IsProtectedPath(Path.Combine(userProfile, "Documents\MyProject")) = True, "Subfolder of Documents should be protected")
        Debug.Assert(IsProtectedPath(Path.Combine(userProfile, "Desktop\TestFolder")) = True, "Subfolder of Desktop should be protected")

        ' === Negative Tests (expected False) ===
        Debug.Assert(IsProtectedPath("C:\Temp") = False, "Temp should not be protected")
        Debug.Assert(IsProtectedPath("D:\Games") = False, "Games folder should not be protected")
        Debug.Assert(IsProtectedPath("C:\Users\Public") = False, "Public folder should not be protected")

        ' === Edge Case Tests ===
        Debug.Assert(IsProtectedPath("c:\windows") = True, "Case-insensitive path should be detected as protected")
        Debug.Assert(IsProtectedPath("C:\Windows\") = True, "Trailing slash should still be detected as protected")
        Debug.Assert(IsProtectedPath(".\Windows") = False, "Relative path should not match protected directories")
        Debug.Assert(IsProtectedPath("") = False, "Empty path should not be protected")
        Debug.Assert(IsProtectedPath(Nothing) = False, "Nothing should not be protected")

        Debug.WriteLine("All IsProtectedPath tests executed.")

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

End Class

' Custom comparer class
Public Class ListViewItemComparer
    Implements IComparer

    Private col As Integer
    Private sortOrder As SortOrder

    Public Sub New(column As Integer, order As SortOrder)
        col = column
        sortOrder = order
    End Sub

    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Dim returnVal As Integer = String.Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)

        If sortOrder = SortOrder.Descending Then
            returnVal *= -1
        End If

        Return returnVal
    End Function

End Class
