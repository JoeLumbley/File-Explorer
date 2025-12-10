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

                'If parts.Length > 2 Then
                '    Dim source As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()
                '    Dim destination As String = parts(parts.Length - 1).Trim()

                '    ' Check if source file exists
                '    If Not File.Exists(source) Then

                '        ' Show error message
                '        ShowStatus("Copy Failed - Source file: ''" & source & "'' does not exist.")

                '        ' Challenging to recover from this error, so just
                '        Return

                '    End If

                '    ' Check if destination directory exists
                '    If Not Directory.Exists(destination) Then

                '        ' Show error message
                '        ShowStatus("Copy Failed - Destination folder: ''" & destination & "'' does not exist.")

                '        ' Challenging to recover from this error, so just
                '        Return

                '    End If

                '    CopyFile(source, destination)

                'Else
                '    ShowStatus("Usage: copy [source] [destination] - copy C:\folder1\file.doc C:\folder2 ")
                'End If

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

            Case "help"

                Dim helpText As String = "Available Commands:" & Environment.NewLine & Environment.NewLine &
                                         "cd [directory] - Change directory" & Environment.NewLine & Environment.NewLine &
                                         "copy [source] [destination]" & Environment.NewLine & Environment.NewLine &
                                         "Copy file to destination folder" & Environment.NewLine & Environment.NewLine &
                                         "copy C:\folderA\file.doc C:\folderB" & Environment.NewLine &
                                         "copy C:\folderA\file.doc C:\folderB\rename.doc" & Environment.NewLine &
                                         "move [source] [destination]" & Environment.NewLine &
                                         "Move file or folder to destination" & Environment.NewLine &
                                         "move C:\folderA\file.doc C:\folderB\file.doc" & Environment.NewLine &
                                         "move C:\folderA\file.doc C:\folderB\rename.doc" & Environment.NewLine &
                                         "move C:\folderA\folder C:\folderB\folder" & Environment.NewLine &
                                         "move C:\folderA\folder C:\folderB\rename" & Environment.NewLine & Environment.NewLine &
                                         "delete [file_or_directory] - Delete file or directory" & Environment.NewLine & Environment.NewLine &
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

    Private Sub txtPath_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPath.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim command As String = txtPath.Text.Trim()
            ExecuteCommand(command)
        End If
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
            'ShowStatus("Copied file: " & fileName & " to: " & destination)
            MsgBox("Copied file: " & fileName & " to: " & destination, MsgBoxStyle.Information, "File Copied")
            Dim destDir As String = Path.GetDirectoryName(destination)
            NavigateTo(destDir)

        Catch ex As Exception
            ShowStatus("Copy Failed: " & ex.Message)
        End Try
    End Sub




    'Private Sub CopyFile(sourcePath As String, destPath As String)

    '    Try

    '        ' Check if the destination file already exists
    '        Dim fileName As String = Path.GetFileName(sourcePath)
    '        Dim destDirFileName = Path.Combine(destPath, fileName)
    '        If File.Exists(destDirFileName) Then
    '            Dim msg As String = "The file '" & fileName & "' already exists in the destination folder." & Environment.NewLine & Environment.NewLine &
    '                                "Do you want to overwrite it?"
    '            Dim result = MessageBox.Show(msg, "File Exists", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
    '            If result = DialogResult.No Then
    '                ShowStatus("Copy operation canceled.")
    '                Return
    '            End If
    '        End If

    '        File.Copy(sourcePath, destDirFileName, overwrite:=True)

    '        ShowStatus("Copied file: " & fileName & " to: " & destPath)

    '        NavigateTo(destPath)

    '    Catch ex As Exception

    '        ShowStatus("Copy Failed: " & ex.Message)

    '    End Try

    'End Sub


    'Private Sub MoveFile(sourcePath As String, destPath As String)

    '    Try
    '        Dim fileName As String = Path.GetFileName(sourcePath)
    '        Dim destDirFileName As String

    '        If Directory.Exists(destPath) Then
    '            ' Destination is an existing directory
    '            destDirFileName = Path.Combine(destPath, fileName)

    '        ElseIf File.Exists(destPath) Then
    '            ' Destination is an existing file
    '            destDirFileName = destPath

    '        ElseIf destPath.EndsWith(Path.DirectorySeparatorChar) OrElse destPath.EndsWith(Path.AltDirectorySeparatorChar) Then
    '            ' Destination does not exist but ends with "\" → treat as directory
    '            Directory.CreateDirectory(destPath)
    '            destDirFileName = Path.Combine(destPath, fileName)

    '        Else
    '            ' Destination does not exist and does not end with "\" → treat as file
    '            destDirFileName = destPath
    '        End If

    '        ' Ensure the destination directory exists
    '        Dim destDir As String = Path.GetDirectoryName(destDirFileName)
    '        If Not Directory.Exists(destDir) Then
    '            Directory.CreateDirectory(destDir)
    '        End If

    '        ' Overwrite prompt if file exists
    '        If File.Exists(destDirFileName) Then
    '            Dim msg As String = "The file '" & Path.GetFileName(destDirFileName) & "' already exists." &
    '                            Environment.NewLine & Environment.NewLine &
    '                            "Do you want to overwrite it?"
    '            Dim result = MessageBox.Show(msg, "File Exists", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
    '            If result = DialogResult.No Then
    '                ShowStatus("Move operation canceled.")
    '                Return
    '            End If
    '        End If

    '        ' Perform the move
    '        If File.Exists(sourcePath) Then
    '            File.Move(sourcePath, destDirFileName)
    '            ShowStatus("Moved file to: " & destDirFileName)
    '        ElseIf Directory.Exists(sourcePath) Then
    '            Directory.Move(sourcePath, destPath)
    '            ShowStatus("Moved directory to: " & destPath)
    '        Else
    '            ShowStatus("Source not found.")
    '        End If

    '        NavigateTo(destPath)

    '    Catch ex As Exception
    '        ShowStatus("Move failed: " & ex.Message)
    '    End Try

    'End Sub

    'Private Sub MoveFile(sourcePath As String, destPath As String)

    '    Try

    '        ' Check if the destination file already exists
    '        Dim fileName As String = Path.GetFileName(sourcePath)
    '        Dim destDirFileName = Path.Combine(destPath, fileName)
    '        If File.Exists(destPath) Then
    '            Dim msg As String = "The file '" & fileName & "' already exists in the destination folder." & Environment.NewLine & Environment.NewLine &
    '                                "Do you want to overwrite it?"
    '            Dim result = MessageBox.Show(msg, "File Exists", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
    '            If result = DialogResult.No Then
    '                ShowStatus("Move operation canceled.")
    '                Return
    '            End If
    '        End If

    '        If File.Exists(sourcePath) Then

    '            File.Move(sourcePath, destDirFileName)

    '            ShowStatus("Moved file to: " & destPath)

    '        ElseIf Directory.Exists(sourcePath) Then

    '            Directory.Move(sourcePath, destPath)

    '            ShowStatus("Moved directory to: " & destPath)

    '        Else
    '            ShowStatus("Source not found.")
    '        End If

    '    Catch ex As Exception

    '        ShowStatus("Move failed: " & ex.Message)

    '    End Try

    'End Sub


    'Private Sub MoveFile(source As String, destination As String)
    '    Try
    '        If File.Exists(source) Then
    '            File.Move(source, destination)
    '            ShowStatus("Moved file to: " & destination)
    '            NavigateTo(destination)

    '        ElseIf Directory.Exists(source) Then
    '            Directory.Move(source, destination)
    '            ShowStatus("Moved directory to: " & destination)
    '            NavigateTo(destination)
    '        Else
    '            ShowStatus("Source not found.")
    '        End If
    '    Catch ex As Exception
    '        ShowStatus("Move failed: " & ex.Message)
    '    End Try
    'End Sub

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


    Private Sub DeleteFileOrDirectory(path As String)
        Try
            If File.Exists(path) Then
                File.Delete(path)
                ShowStatus("Deleted file: " & path)
            ElseIf Directory.Exists(path) Then
                Directory.Delete(path, recursive:=True)
                ShowStatus("Deleted directory: " & path)
            Else
                ShowStatus("Path not found.")
            End If
        Catch ex As Exception
            ShowStatus("Delete failed: " & ex.Message)
        End Try
    End Sub

    Private Sub lvFiles_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lvFiles.ColumnClick

        ' Toggle between ascending and descending
        Dim sortOrder As SortOrder = If(lvFiles.Sorting = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)

        ' Sort the ListView
        lvFiles.Sorting = sortOrder
        lvFiles.ListViewItemSorter = New ListViewItemComparer(e.Column, sortOrder)
        lvFiles.Sort()

    End Sub

    'Private Sub CopyDirectory(sourceDir As String, destDir As String)
    '    Dim dirInfo As New DirectoryInfo(sourceDir)
    '    If Not dirInfo.Exists Then Throw New DirectoryNotFoundException("Source not found: " & sourceDir)

    '    Directory.CreateDirectory(destDir)

    '    ' Copy files
    '    For Each file In dirInfo.GetFiles()
    '        Dim targetFilePath = Path.Combine(destDir, file.Name)
    '        file.CopyTo(targetFilePath, overwrite:=True)
    '    Next

    '    ' Copy subdirectories recursively
    '    For Each subDir In dirInfo.GetDirectories()
    '        Dim newDest = Path.Combine(destDir, subDir.Name)
    '        CopyDirectory(subDir.FullName, newDest)
    '    Next
    'End Sub


    Private Sub CopyDirectory(sourceDir As String, destDir As String)

        Dim dirInfo As New DirectoryInfo(sourceDir)

        If Not dirInfo.Exists Then Throw New DirectoryNotFoundException("Source not found: " & sourceDir)

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

    End Sub

    Private Sub CopySelectedFile_Click(sender As Object, e As EventArgs)
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        Dim fullPath = CStr(lvFiles.SelectedItems(0).Tag)

        ' Example: copy to Desktop
        Dim destPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), Path.GetFileName(fullPath))

        If File.Exists(fullPath) Then
            CopyFile(fullPath, destPath)
        ElseIf Directory.Exists(fullPath) Then
            CopyDirectory(fullPath, destPath)
        End If
    End Sub

    Private Sub PasteSelected_Click(sender As Object, e As EventArgs)
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

    Private Sub ResetCutVisuals()
        For Each item As ListViewItem In lvFiles.Items
            item.ForeColor = SystemColors.WindowText
            item.Font = New Font(lvFiles.Font, FontStyle.Regular)
        Next
    End Sub

    Private Sub CopySelected_Click(sender As Object, e As EventArgs)
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        _clipboardPath = CStr(lvFiles.SelectedItems(0).Tag)
        ShowStatus("Copied to clipboard: " & _clipboardPath)
    End Sub

    Private Sub CutSelected_Click(sender As Object, e As EventArgs)
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        _clipboardPath = CStr(lvFiles.SelectedItems(0).Tag)
        _clipboardIsCut = True

        ' Fade the item to indicate "cut"
        Dim sel = lvFiles.SelectedItems(0)
        sel.ForeColor = Color.Gray
        sel.Font = New Font(sel.Font, FontStyle.Italic)
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
        cmsFiles.Items.Add("Cut", Nothing, AddressOf CutSelected_Click)
        cmsFiles.Items.Add("Copy", Nothing, AddressOf CopySelected_Click)
        cmsFiles.Items.Add("Paste", Nothing, AddressOf PasteSelected_Click)

        cmsFiles.Items.Add("Open", Nothing, AddressOf Open_Click)


        cmsFiles.Items.Add("Copy Name", Nothing, AddressOf CopyFileName_Click)
        cmsFiles.Items.Add("Copy Path", Nothing, AddressOf CopyFilePath_Click)

        cmsFiles.Items.Add("New Folder", Nothing, AddressOf NewFolder_Click)
        cmsFiles.Items.Add("Rename", Nothing, AddressOf RenameFile_Click)
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
        lvFiles.Columns.Add("Name", 550)
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
        'GoToFolderOrOpenFile(txtPath.Text)
        ExecuteCommand(txtPath.Text.Trim())

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
