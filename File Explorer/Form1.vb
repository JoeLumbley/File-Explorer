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
' Documentation: You can find the documentation at the GitHub repository.


Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading

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

    Private statusTimer As New System.Windows.Forms.Timer() With {.Interval = 10000}



    Private currentFolder As String = String.Empty

    Private ShowHiddenFiles As Boolean = False

    Private ColumnTypes As New Dictionary(Of Integer, ListViewItemComparer.ColumnDataType) From {
        {0, ListViewItemComparer.ColumnDataType.Text},       ' Name
        {1, ListViewItemComparer.ColumnDataType.Text},       ' Type
        {2, ListViewItemComparer.ColumnDataType.Number},     ' Size
        {3, ListViewItemComparer.ColumnDataType.DateValue}   ' Modified
    }

    Friend Shared ReadOnly SizeUnits As (Unit As String, Factor As Long)() = {
        ("B", 1L), ' Bytes
        ("KB", 1024L), ' Kilobytes
        ("MB", 1024L ^ 2), ' Megabytes
        ("GB", 1024L ^ 3), ' Gigabytes
        ("TB", 1024L ^ 4), ' Terabytes
        ("PB", 1024L ^ 5), ' Petabytes
        ("EB", 1024L ^ 6)  ' Exabytes
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
    Private IconMoving As String = ""
    Private IconQuestion As String = ""

    Dim SearchResults As New List(Of String)
    Private SearchIndex As Integer = -1

    Private ReadOnly tips As New ToolTip()

    Private fileTypeMap As New Dictionary(Of String, String) From {
        {".aac", "Audio"}, {".aiff", "Audio"}, {".alac", "Audio"},
        {".dsd", "Audio"}, {".flac", "Audio"}, {".m4a", "Audio"},
        {".mp3", "Audio"}, {".ogg", "Audio"}, {".wav", "Audio"},
        {".wma", "Audio"},
        {".bmp", "Image"}, {".cr2", "Image"}, {".gif", "Image"},
        {".heic", "Image"}, {".jpeg", "Image"}, {".jpg", "Image"},
        {".nef", "Image"}, {".orf", "Image"}, {".png", "Image"},
        {".raw", "Image"}, {".sr2", "Image"}, {".svg", "Image"},
        {".tiff", "Image"}, {".webp", "Image"},
        {".doc", "Document"}, {".docx", "Document"}, {".htm", "Document"},
        {".html", "Document"}, {".md", "Document"}, {".odp", "Document"},
        {".ods", "Document"}, {".odt", "Document"}, {".pdf", "Document"},
        {".ppt", "Document"}, {".pptx", "Document"}, {".rtf", "Document"},
        {".txt", "Document"}, {".xls", "Document"}, {".xlsx", "Document"},
        {".3gp", "Video"}, {".avi", "Video"}, {".flv", "Video"},
        {".mkv", "Video"}, {".mov", "Video"}, {".mp4", "Video"},
        {".mpeg", "Video"}, {".mpg", "Video"}, {".ogv", "Video"},
        {".vob", "Video"}, {".webm", "Video"},
        {".7z", "Archive"}, {".apk", "Archive"}, {".crx", "Archive"},
        {".dmg", "Archive"}, {".epub", "Archive"}, {".gz", "Archive"},
        {".iso", "Archive"}, {".mobi", "Archive"}, {".rar", "Archive"},
        {".tar", "Archive"}, {".zip", "Archive"},
        {".bat", "Executable"}, {".cmd", "Executable"}, {".com", "Executable"},
        {".dll", "Executable"}, {".exe", "Executable"}, {".jar", "Executable"},
        {".msi", "Executable"}, {".pdb", "Executable"}, {".pif", "Executable"},
        {".ps1", "Executable"}, {".scr", "Executable"}, {".vbs", "Executable"},
        {".wsf", "Executable"},
        {".asm", "Code"}, {".bash", "Code"}, {".c", "Code"},
        {".cc", "Code"}, {".cpp", "Code"}, {".cs", "Code"},
        {".csproj", "C# Project"}, {".vbproj", "Visual Basic Project"},
        {".fsproj", "F# Project"}, {".sln", "Visual Studio Solution"},
        {".cxx", "Code"}, {".go", "Code"}, {".h", "Code"},
        {".hh", "Code"}, {".hpp", "Code"}, {".java", "Code"},
        {".js", "Code"}, {".json", "Code"}, {".jsonc", "Code"},
        {".jsx", "Code"}, {".kts", "Code"}, {".kt", "Code"},
        {".lua", "Code"}, {".m", "Code"}, {".php", "Code"},
        {".psql", "Code"}, {".py", "Code"}, {".r", "Code"},
        {".rb", "Code"}, {".resx", "Resource"}, {".rs", "Code"},
        {".s", "Code"}, {".sh", "Code"}, {".sql", "Code"},
        {".swift", "Code"}, {".toml", "Code"}, {".ts", "Code"},
        {".tsx", "Code"}, {".vb", "Visual Basic Source"},
        {".user", "Project Options"}, {".vbx", "Code"},
        {".xaml", "Code"}, {".xml", "Code"}, {".yaml", "Code"},
        {".yml", "Code"},
        {".cfg", "Config"}, {".ini", "Config"},
        {".lnk", "Shortcut"}
    }

    Dim StatusPad As String = New String(" "c, 2)

    Private copyCts As CancellationTokenSource

    Private _isRenaming As Boolean = False

    Private ReadOnly EasyAccessFile As String =
    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "easyaccess.txt")

    Private Sub EnsureEasyAccessFile()
        Dim dir = Path.GetDirectoryName(EasyAccessFile)
        If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
        If Not File.Exists(EasyAccessFile) Then File.WriteAllText(EasyAccessFile, "")
    End Sub

    Private Function LoadEasyAccessEntries() As List(Of (Name As String, Path As String))
        EnsureEasyAccessFile()

        Dim list As New List(Of (String, String))

        For Each line In File.ReadAllLines(EasyAccessFile)
            If String.IsNullOrWhiteSpace(line) Then Continue For

            Dim parts = line.Split({","c}, 2)
            If parts.Length = 2 Then
                Dim name = parts(0).Trim()
                Dim path = parts(1).Trim()

                If Directory.Exists(path) Then
                    list.Add((name, path))
                End If
            End If
        Next

        Return list
    End Function

    'Public Sub AddToEasyAccess(name As String, path As String)
    '    EnsureEasyAccessFile()

    '    Dim entry = $"{name},{path}"
    '    Dim existing = File.ReadAllLines(EasyAccessFile)

    '    If Not existing.Contains(entry) Then
    '        File.AppendAllLines(EasyAccessFile, {entry})
    '    End If

    '    UpdateTreeRoots()
    'End Sub


    Public Sub AddToEasyAccess(name As String, path As String)
        EnsureEasyAccessFile()

        Dim entry = $"{name},{path}"
        Dim existing = File.ReadAllLines(EasyAccessFile)

        If Not existing.Contains(entry) Then
            File.AppendAllLines(EasyAccessFile, {entry})
        End If

        UpdateTreeRoots()
    End Sub

    Public Sub RemoveFromEasyAccess(path As String)
        EnsureEasyAccessFile()

        Dim lines = File.ReadAllLines(EasyAccessFile).ToList()
        Dim updated = lines.Where(Function(l) Not l.EndsWith("," & path)).ToList()

        File.WriteAllLines(EasyAccessFile, updated)

        UpdateTreeRoots()
    End Sub




    Private Sub Form_Load(sender As Object, e As EventArgs) _
        Handles MyBase.Load

        InitApp()

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' Ensure the address bar receives focus after the form is shown and after any
        ' sync/async initialization finishes. Using BeginInvoke guarantees this runs
        ' once the message loop is ready (avoids focus being stolen by other init).

        Me.BeginInvoke(New Action(Sub()
                                      If txtAddressBar IsNot Nothing AndAlso txtAddressBar.CanFocus Then
                                          txtAddressBar.Focus()
                                          PlaceCaretAtEndOfAddressBar()
                                      End If
                                  End Sub)
        )

    End Sub


    'Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean

    '    ' Handle custom key commands
    '    If HandleEnterKey(keyData) Then Return True
    '    If HandleAddressBarShortcuts(keyData) Then Return True

    '    If HandleTabNavigation(keyData) Then Return True
    '    If HandleShiftTabNavigation(keyData) Then Return True

    '    If HandleNavigationShortcuts(keyData) Then Return True

    '    If HandleSearchShortcuts(keyData) Then Return True
    '    If HandleFileFolderOperations(Nothing, keyData) Then Return True

    '    Return MyBase.ProcessCmdKey(msg, keyData)
    'End Function



    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean

        ' Highest‑priority, context‑sensitive actions
        If HandleEnterKey(keyData) Then Return True
        If HandleAddressBarShortcuts(keyData) Then Return True

        ' Focus navigation (Shift+Tab should be checked before Tab)
        If HandleShiftTabNavigation(keyData) Then Return True
        If HandleTabNavigation(keyData) Then Return True

        ' Explorer‑style navigation shortcuts
        If HandleNavigationShortcuts(keyData) Then Return True

        ' Search and file operations (lowest priority)
        If HandleSearchShortcuts(keyData) Then Return True
        If HandleFileFolderOperations(Nothing, keyData) Then Return True

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function


    Private Sub tvFolders_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) _
        Handles tvFolders.NodeMouseClick

        Dim info = tvFolders.HitTest(e.Location)

        ' Only toggle node when the arrow is clicked
        If info.Location = TreeViewHitTestLocations.StateImage Then

            tvFolders.BeginUpdate()

            If e.Node.IsExpanded Then
                e.Node.Collapse()
            Else
                e.Node.Expand()
            End If

            tvFolders.EndUpdate()

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

        ' Set collapsed icon
        e.Node.StateImageIndex = 0   ' ▶ collapsed

    End Sub


    Private Sub lvFiles_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) _
        Handles lvFiles.ItemSelectionChanged

        UpdateEditButtonsAndMenus()

        UpdateFileButtonsAndMenus()

    End Sub

    Private Sub lvFiles_ItemActivate(sender As Object, e As EventArgs) _
        Handles lvFiles.ItemActivate
        ' The ItemActivate event is raised when the user double-clicks an item or
        ' presses the Enter key when an item is selected.

        If lvFiles.SelectedItems().Count = 0 Then Exit Sub


        ' Validate selection
        If Not PathExists(CStr(lvFiles.SelectedItems(0).Tag)) Then
            ShowStatus(StatusPad & IconError & " The selected item isn't a file or folder and can't be opened.")
            Exit Sub
        End If

        GoToFolderOrOpenFile_EnterKeyDownOrDoubleClick()

    End Sub

    Private Sub lvFiles_BeforeLabelEdit(sender As Object, e As LabelEditEventArgs) _
        Handles lvFiles.BeforeLabelEdit

        _isRenaming = True


        Dim item As ListViewItem = lvFiles.Items(e.Item)
        Dim fullPath As String = CStr(item.Tag)

        ' Rule 1: Path must exist
        If Not PathExists(fullPath) Then
            e.CancelEdit = True
            ShowStatus(StatusPad & IconError & " The selected item doesn't exist and cannot be renamed.")
            Exit Sub
        End If

        ' Rule 2: Protected items cannot be renamed
        If IsProtectedPathOrFolder(fullPath) Then
            e.CancelEdit = True
            ShowStatus(StatusPad & IconProtect & " This item is protected and cannot be renamed.")
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

        If Not HasWriteAccess(parentDir) Then
            e.CancelEdit = True
            ShowStatus(StatusPad & IconError & " This location does not allow renaming.")
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

        NavigateTo(currentFolder, recordHistory:=False)

        UpdateTreeRoots()

    End Sub

    Private Sub bntHome_Click(sender As Object, e As EventArgs) _
        Handles bntHome.Click

        ' Go to user's home directory
        NavigateTo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), True)

        UpdateNavButtons()

    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) _
        Handles btnGo.Click

        ExecuteCommand(txtAddressBar.Text.Trim)

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

    'Private Function HandleAddressBarShortcuts(keyData As Keys) As Boolean

    '    ' ===========================
    '    '   FOCUS ADDRESS BAR (Ctrl+L, Alt+D, F4)
    '    ' ===========================
    '    If keyData = (Keys.Control Or Keys.L) _
    '    OrElse keyData = (Keys.Alt Or Keys.D) _
    '    OrElse keyData = Keys.F4 Then
    '        txtAddressBar.Focus()
    '        txtAddressBar.SelectAll()
    '        Return True
    '    End If

    '    ' ===========================
    '    '   ESCAPE (Address Bar reset)
    '    ' ===========================


    '    If keyData = Keys.Escape AndAlso Not _isRenaming Then
    '        txtAddressBar.Text = currentFolder

    '        'GoToFolderOrOpenFile(currentFolder)
    '        'PopulateFiles(currentFolder)

    '        ' Reset index for new search
    '        SearchIndex = 0

    '        SearchResults = New List(Of String)

    '        NavigateTo(currentFolder, recordHistory:=False)

    '        PlaceCaretAtEndOfAddressBar()
    '        Return True
    '    End If


    '    Return False
    'End Function



    Private Function HandleAddressBarShortcuts(keyData As Keys) As Boolean

        ' ===========================
        '   FOCUS ADDRESS BAR (Ctrl+L, Alt+D, F4)
        ' ===========================
        If keyData = (Keys.Control Or Keys.L) _
            OrElse keyData = (Keys.Alt Or Keys.D) _
            OrElse keyData = Keys.F4 Then

            txtAddressBar.Focus()
            txtAddressBar.SelectAll()
            Return True
        End If

        ' ===========================
        '   ESCAPE (Address Bar reset)
        ' ===========================
        If keyData = Keys.Escape AndAlso Not _isRenaming Then

            txtAddressBar.Text = currentFolder

            ' Reset search state
            SearchIndex = 0
            SearchResults = New List(Of String)

            ' Navigate without recording history
            NavigateTo(currentFolder, recordHistory:=False)

            PlaceCaretAtEndOfAddressBar()
            Return True
        End If

        Return False
    End Function


    'Private Function HandleEnterKey(keyData As Keys) As Boolean

    '    If keyData <> Keys.Enter Then
    '        Return False
    '    End If

    '    ' ===========================
    '    '   ENTER (Address Bar execute)
    '    ' ===========================
    '    If txtAddressBar.Focused Then
    '        ExecuteCommand(txtAddressBar.Text.Trim())
    '        Return True
    '    End If

    '    ' ===========================
    '    '   ENTER (TreeView toggle)
    '    ' ===========================
    '    If tvFolders.Focused Then
    '        ToggleExpandCollapse()
    '        Return True
    '    End If

    '    ' ===========================
    '    '   ENTER (File List open)
    '    ' ===========================
    '    If lvFiles.Focused Then
    '        OpenSelectedItem()
    '        Return True
    '    End If

    '    Return False
    'End Function


    Private Function HandleEnterKey(keyData As Keys) As Boolean

        If keyData <> Keys.Enter Then
            Return False
        End If

        ' Block global Enter behavior during rename mode
        If _isRenaming Then
            Return False
        End If

        ' ===========================
        '   ENTER (Address Bar execute)
        ' ===========================
        If txtAddressBar.Focused Then
            ExecuteCommand(txtAddressBar.Text.Trim())
            Return True
        End If

        ' ===========================
        '   ENTER (TreeView toggle)
        ' ===========================
        If tvFolders.Focused Then
            ToggleExpandCollapse()
            Return True
        End If

        ' ===========================
        '   ENTER (File List open)
        ' ===========================
        If lvFiles.Focused Then
            OpenSelectedItem()
            Return True
        End If

        Return False
    End Function

    Private Sub OpenSelectedItem()
        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        Dim fullPath = CStr(lvFiles.SelectedItems(0).Tag)
        GoToFolderOrOpenFile(fullPath)
    End Sub

    'Private Function HandleTabNavigation(keyData As Keys) As Boolean
    '    If keyData = Keys.Tab Then

    '        ' ==========================
    '        ' Address Bar → File List
    '        ' ==========================
    '        If txtAddressBar.Focused Then
    '            lvFiles.Focus()

    '            If lvFiles.Items.Count > 0 Then
    '                If lvFiles.SelectedItems.Count > 0 Then
    '                    Dim sel = lvFiles.SelectedItems(0)
    '                    sel.Focused = True
    '                    sel.EnsureVisible()
    '                Else
    '                    lvFiles.Items(0).Selected = True
    '                    lvFiles.Items(0).Focused = True
    '                    lvFiles.Items(0).EnsureVisible()
    '                End If
    '            End If

    '            Return True
    '        End If

    '        ' ==========================
    '        ' File List → TreeView
    '        ' ==========================
    '        If lvFiles.Focused Then
    '            tvFolders.Focus()

    '            If tvFolders.SelectedNode Is Nothing AndAlso tvFolders.Nodes.Count > 0 Then
    '                tvFolders.SelectedNode = tvFolders.Nodes(0)
    '            End If

    '            tvFolders.SelectedNode?.EnsureVisible()
    '            Return True
    '        End If

    '        ' ==========================
    '        ' TreeView → Address Bar
    '        ' ==========================
    '        If tvFolders.Focused Then
    '            txtAddressBar.Focus()
    '            PlaceCaretAtEndOfAddressBar()
    '            Return True
    '        End If

    '        ' ==========================
    '        ' Fallback
    '        ' ==========================
    '        txtAddressBar.Focus()
    '        PlaceCaretAtEndOfAddressBar()
    '        Return True
    '    End If

    '    Return False
    'End Function

    Private Function HandleTabNavigation(keyData As Keys) As Boolean

        If keyData <> Keys.Tab Then
            Return False
        End If

        ' Prevent Tab navigation while renaming
        If _isRenaming Then
            Return False
        End If

        ' ==========================
        ' Address Bar → File List
        ' ==========================
        If txtAddressBar.Focused Then
            lvFiles.Focus()

            If lvFiles.Items.Count > 0 Then
                If lvFiles.SelectedItems.Count > 0 Then
                    Dim sel = lvFiles.SelectedItems(0)
                    sel.Focused = True
                    sel.EnsureVisible()
                Else
                    lvFiles.Items(0).Selected = True
                    lvFiles.Items(0).Focused = True
                    lvFiles.Items(0).EnsureVisible()
                End If
            End If

            Return True
        End If

        ' ==========================
        ' File List → TreeView
        ' ==========================
        If lvFiles.Focused Then
            tvFolders.Focus()

            If tvFolders.SelectedNode Is Nothing AndAlso tvFolders.Nodes.Count > 0 Then
                tvFolders.SelectedNode = tvFolders.Nodes(0)
            End If

            tvFolders.SelectedNode?.EnsureVisible()
            Return True
        End If

        ' ==========================
        ' TreeView → Address Bar
        ' ==========================
        If tvFolders.Focused Then
            txtAddressBar.Focus()
            PlaceCaretAtEndOfAddressBar()
            Return True
        End If

        ' ==========================
        ' Fallback
        ' ==========================
        txtAddressBar.Focus()
        PlaceCaretAtEndOfAddressBar()
        Return True
    End Function

    'Private Function HandleShiftTabNavigation(keyData As Keys) As Boolean

    '    ' Detect SHIFT + TAB correctly inside ProcessCmdKey
    '    If keyData = (Keys.Shift Or Keys.Tab) Then

    '        ' ===========================
    '        '   TreeView → File List
    '        ' ===========================
    '        If tvFolders.Focused Then
    '            lvFiles.Focus()

    '            If lvFiles.Items.Count > 0 Then
    '                If lvFiles.SelectedItems.Count > 0 Then
    '                    Dim sel = lvFiles.SelectedItems(0)
    '                    sel.Focused = True
    '                    sel.EnsureVisible()
    '                Else
    '                    lvFiles.Items(0).Selected = True
    '                    lvFiles.Items(0).Focused = True
    '                    lvFiles.Items(0).EnsureVisible()
    '                End If
    '            End If

    '            Return True
    '        End If

    '        ' ===========================
    '        '   File List → Address Bar
    '        ' ===========================
    '        If lvFiles.Focused Then
    '            txtAddressBar.Focus()
    '            PlaceCaretAtEndOfAddressBar()
    '            Return True
    '        End If

    '        ' ===========================
    '        '   Address Bar → TreeView
    '        ' ===========================
    '        If txtAddressBar.Focused Then
    '            tvFolders.Focus()

    '            If tvFolders.SelectedNode Is Nothing AndAlso tvFolders.Nodes.Count > 0 Then
    '                tvFolders.SelectedNode = tvFolders.Nodes(0)
    '            End If

    '            tvFolders.SelectedNode?.EnsureVisible()
    '            Return True
    '        End If

    '        ' ===========================
    '        '   Fallback
    '        ' ===========================
    '        txtAddressBar.Focus()
    '        PlaceCaretAtEndOfAddressBar()
    '        Return True
    '    End If

    '    Return False
    'End Function

    Private Function HandleShiftTabNavigation(keyData As Keys) As Boolean

        ' Detect SHIFT + TAB correctly inside ProcessCmdKey
        If keyData <> (Keys.Shift Or Keys.Tab) Then
            Return False
        End If

        ' Prevent Shift+Tab navigation while renaming
        If _isRenaming Then
            Return False
        End If

        ' ===========================
        '   TreeView → File List
        ' ===========================
        If tvFolders.Focused Then
            lvFiles.Focus()

            If lvFiles.Items.Count > 0 Then
                If lvFiles.SelectedItems.Count > 0 Then
                    Dim sel = lvFiles.SelectedItems(0)
                    sel.Focused = True
                    sel.EnsureVisible()
                Else
                    lvFiles.Items(0).Selected = True
                    lvFiles.Items(0).Focused = True
                    lvFiles.Items(0).EnsureVisible()
                End If
            End If

            Return True
        End If

        ' ===========================
        '   File List → Address Bar
        ' ===========================
        If lvFiles.Focused Then
            txtAddressBar.Focus()
            PlaceCaretAtEndOfAddressBar()
            Return True
        End If

        ' ===========================
        '   Address Bar → TreeView
        ' ===========================
        If txtAddressBar.Focused Then
            tvFolders.Focus()

            If tvFolders.SelectedNode Is Nothing AndAlso tvFolders.Nodes.Count > 0 Then
                tvFolders.SelectedNode = tvFolders.Nodes(0)
            End If

            tvFolders.SelectedNode?.EnsureVisible()
            Return True
        End If

        ' ===========================
        '   Fallback
        ' ===========================
        txtAddressBar.Focus()
        PlaceCaretAtEndOfAddressBar()
        Return True
    End Function

    Private Function GlobalShortcutsAllowed() As Boolean
        Return Not txtAddressBar.Focused AndAlso Not _isRenaming
    End Function

    'Private Function HandleNavigationShortcuts(keyData As Keys) As Boolean

    '    '' ===========================
    '    '' ALT + HOME (Goto User Folder)
    '    '' ===========================
    '    'If keyData = (Keys.Alt Or Keys.Home) AndAlso
    '    '    Not txtAddressBar.Focused AndAlso
    '    '    Not _isRenaming Then
    '    '    GoToFolderOrOpenFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))
    '    '    Return True
    '    'End If

    '    ' ===========================
    '    ' ALT + HOME (Goto User Folder)
    '    ' ===========================
    '    If keyData = (Keys.Alt Or Keys.Home) AndAlso GlobalShortcutsAllowed() Then
    '        GoToFolderOrOpenFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))
    '        Return True
    '    End If

    '    ' ===========================
    '    ' ALT + LEFT (Back)
    '    ' ===========================
    '    If keyData = (Keys.Alt Or Keys.Left) Then
    '        NavigateBackward_Click()
    '        PlaceCaretAtEndOfAddressBar()
    '        Return True
    '    End If

    '    ' ===========================
    '    ' ALT + RIGHT (Forward)
    '    ' ===========================
    '    If keyData = (Keys.Alt Or Keys.Right) Then
    '        NavigateForward_Click()
    '        PlaceCaretAtEndOfAddressBar()
    '        Return True
    '    End If

    '    ' ===========================
    '    ' ALT + UP (Parent folder)
    '    ' ===========================
    '    If keyData = (Keys.Alt Or Keys.Up) Then
    '        NavigateToParent()
    '        PlaceCaretAtEndOfAddressBar()
    '        Return True
    '    End If

    '    ' ===========================
    '    ' F5 (Refresh)
    '    ' ===========================
    '    If keyData = Keys.F5 Then
    '        RefreshCurrentFolder()
    '        txtAddressBar.Focus()

    '        PlaceCaretAtEndOfAddressBar()

    '        Return True
    '    End If

    '    ' ===========================
    '    ' F11 (Full screen)
    '    ' ===========================
    '    If keyData = Keys.F11 Then
    '        ToggleFullScreen()
    '        Return True
    '    End If

    '    Return False
    'End Function


    Private Function HandleNavigationShortcuts(keyData As Keys) As Boolean

        ' ===========================
        ' ALT + HOME (Goto User Folder)
        ' ===========================
        If keyData = (Keys.Alt Or Keys.Home) AndAlso Not _isRenaming Then
            GoToFolderOrOpenFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))
            PlaceCaretAtEndOfAddressBar()
            Return True
        End If

        ' ===========================
        ' ALT + LEFT (Back)
        ' ===========================
        If keyData = (Keys.Alt Or Keys.Left) AndAlso Not _isRenaming Then
            NavigateBackward_Click()
            PlaceCaretAtEndOfAddressBar()
            Return True
        End If

        ' ===========================
        ' ALT + RIGHT (Forward)
        ' ===========================
        If keyData = (Keys.Alt Or Keys.Right) AndAlso Not _isRenaming Then
            NavigateForward_Click()
            PlaceCaretAtEndOfAddressBar()
            Return True
        End If

        ' ===========================
        ' ALT + UP (Parent folder)
        ' ===========================
        If keyData = (Keys.Alt Or Keys.Up) AndAlso Not _isRenaming Then
            NavigateToParent()
            PlaceCaretAtEndOfAddressBar()
            Return True
        End If


        ' ===========================
        ' F5 (Refresh)
        ' ===========================
        If keyData = Keys.F5 AndAlso Not _isRenaming Then
            RefreshCurrentFolder()
            txtAddressBar.Focus()
            PlaceCaretAtEndOfAddressBar()
            Return True
        End If

        ' ===========================
        ' F11 (Full screen)
        ' ===========================
        If keyData = Keys.F11 AndAlso Not _isRenaming Then
            ToggleFullScreen()
            Return True
        End If

        Return False
    End Function

    'Private Function HandleSearchShortcuts(keyData As Keys) As Boolean

    '    ' ===========================
    '    '   CTRL + F (Find)
    '    ' ===========================
    '    If keyData = (Keys.Control Or Keys.F) Then
    '        InitiateSearch()
    '        Return True
    '    End If

    '    ' ===========================
    '    '   F3 (Find Next)
    '    ' ===========================
    '    If keyData = Keys.F3 Then
    '        HandleFindNextCommand()
    '        Return True
    '    End If

    '    Return False
    'End Function

    Private Function HandleSearchShortcuts(keyData As Keys) As Boolean

        ' ===========================
        '   CTRL + F (Find)
        ' ===========================
        If keyData = (Keys.Control Or Keys.F) AndAlso GlobalShortcutsAllowed() Then
            InitiateSearch()
            Return True
        End If

        ' ===========================
        '   F3 (Find Next)
        ' ===========================
        If keyData = Keys.F3 AndAlso GlobalShortcutsAllowed() Then
            HandleFindNextCommand()
            Return True
        End If

        Return False
    End Function

    'Private Function HandleFileFolderOperations(sender As Object, keyData As Keys) As Boolean
    '    Try
    '        ' ===========================
    '        '   CTRL + O  (Open)
    '        ' ===========================
    '        If keyData = (Keys.Control Or Keys.O) Then
    '            OpenSelectedOrStartCommand()
    '            Return True
    '        End If

    '        ' ===========================
    '        '   CTRL + SHIFT + E  (Expand one level)
    '        ' ===========================
    '        If keyData = (Keys.Control Or Keys.Shift Or Keys.E) Then
    '            ExpandOneLevel()
    '            Return True
    '        End If

    '        ' ===========================
    '        '   CTRL + SHIFT + C  (Collapse one level)
    '        ' ===========================
    '        If keyData = (Keys.Control Or Keys.Shift Or Keys.C) Then
    '            CollapseOneLevel()
    '            Return True
    '        End If

    '        ' ===========================
    '        '   CTRL + SHIFT + N  (New folder)
    '        ' ===========================
    '        If keyData = (Keys.Control Or Keys.Shift Or Keys.N) Then
    '            NewFolder_Click(sender, EventArgs.Empty)
    '            Return True
    '        End If

    '        ' ===========================
    '        '   CTRL + SHIFT + T  (New text file)
    '        ' ===========================
    '        If keyData = (Keys.Control Or Keys.Shift Or Keys.T) Then
    '            NewTextFile_Click(sender, EventArgs.Empty)
    '            Return True
    '        End If












    '        '' ===========================
    '        ''   F2 (Rename)
    '        '' ===========================
    '        'If keyData = Keys.F2 Then
    '        '    RenameFile_Click(sender, EventArgs.Empty)
    '        '    Return True
    '        'End If


    '        ' ===========================
    '        '   F2 (Rename)
    '        ' ===========================
    '        If keyData = Keys.F2 AndAlso
    '        Not txtAddressBar.Focused AndAlso
    '        Not _isRenaming Then
    '            RenameFile_Click(sender, EventArgs.Empty)
    '            Return True
    '        End If








    '        '' ===========================
    '        ''   CTRL + C (Copy)
    '        '' ===========================
    '        'If keyData = (Keys.Control Or Keys.C) AndAlso Not txtAddressBar.Focused Then
    '        '    CopySelected_Click(sender, EventArgs.Empty)
    '        '    Return True
    '        'End If

    '        ' ===========================
    '        '   CTRL + C (Copy)
    '        ' ===========================
    '        If keyData = (Keys.Control Or Keys.C) AndAlso
    '        Not txtAddressBar.Focused AndAlso
    '        Not _isRenaming Then
    '            CopySelected_Click(sender, EventArgs.Empty)
    '            Return True
    '        End If


    '        ' ===========================
    '        '   CTRL + V (Paste)
    '        ' ===========================
    '        'If keyData = (Keys.Control Or Keys.V) AndAlso Not txtAddressBar.Focused Then
    '        '    PasteSelected_Click(sender, EventArgs.Empty)
    '        '    Return True
    '        'End If

    '        If keyData = (Keys.Control Or Keys.V) AndAlso
    '        Not txtAddressBar.Focused AndAlso
    '        Not _isRenaming Then
    '            PasteSelected_Click(sender, EventArgs.Empty)
    '            Return True
    '        End If






    '        '' ===========================
    '        ''   CTRL + X (Cut)
    '        '' ===========================
    '        'If keyData = (Keys.Control Or Keys.X) AndAlso Not txtAddressBar.Focused Then
    '        '    CutSelected_Click(sender, EventArgs.Empty)
    '        '    Return True
    '        'End If








    '        ' ===========================
    '        '   CTRL + X (Cut)
    '        ' ===========================
    '        If keyData = (Keys.Control Or Keys.X) AndAlso
    '        Not txtAddressBar.Focused AndAlso
    '        Not _isRenaming Then
    '            CutSelected_Click(sender, EventArgs.Empty)
    '            Return True
    '        End If











    '        '' ===========================
    '        ''   CTRL + A (Select all)
    '        '' ===========================
    '        'If keyData = (Keys.Control Or Keys.A) Then
    '        '    SelectAllItems()
    '        '    lvFiles.Focus()
    '        '    Return True
    '        'End If


    '        ' ===========================
    '        '   CTRL + A (Select all)
    '        ' ===========================
    '        If keyData = (Keys.Control Or Keys.A) AndAlso
    '        Not txtAddressBar.Focused AndAlso
    '        Not _isRenaming Then
    '            SelectAllItems()
    '            lvFiles.Focus()
    '            Return True
    '        End If







    '        ' ' ===========================
    '        ' '   CTRL + D  OR  DELETE  (Delete)
    '        ' ' ===========================
    '        ' If keyData = (Keys.Control Or Keys.D) _
    '        'OrElse keyData = Keys.Delete Then

    '        '     Delete_Click(sender, EventArgs.Empty)
    '        '     Return True
    '        ' End If

    '        ' ===========================
    '        '   CTRL + D  OR  DELETE  (Delete)
    '        ' ===========================
    '        If (keyData = (Keys.Control Or Keys.D) OrElse keyData = Keys.Delete) AndAlso
    '           Not txtAddressBar.Focused AndAlso
    '           Not _isRenaming Then
    '            Delete_Click(sender, EventArgs.Empty)
    '            Return True
    '        End If









    '    Catch ex As Exception
    '        MessageBox.Show("An error occurred: " & ex.Message,
    '                    "Error",
    '                    MessageBoxButtons.OK,
    '                    MessageBoxIcon.Error)
    '    End Try

    '    Return False
    'End Function

    Private Function HandleFileFolderOperations(sender As Object, keyData As Keys) As Boolean
        Try
            ' ===========================
            '   CTRL + O  (Open)
            ' ===========================
            If keyData = (Keys.Control Or Keys.O) AndAlso GlobalShortcutsAllowed() Then
                OpenSelectedOrStartCommand()
                Return True
            End If

            ' ===========================
            '   CTRL + SHIFT + E  (Expand one level)
            ' ===========================
            If keyData = (Keys.Control Or Keys.Shift Or Keys.E) AndAlso GlobalShortcutsAllowed() Then
                ExpandOneLevel()
                Return True
            End If

            ' ===========================
            '   CTRL + SHIFT + C  (Collapse one level)
            ' ===========================
            If keyData = (Keys.Control Or Keys.Shift Or Keys.C) AndAlso GlobalShortcutsAllowed() Then
                CollapseOneLevel()
                Return True
            End If

            ' ===========================
            '   CTRL + SHIFT + N  (New folder)
            ' ===========================
            If keyData = (Keys.Control Or Keys.Shift Or Keys.N) AndAlso GlobalShortcutsAllowed() Then
                NewFolder_Click(sender, EventArgs.Empty)
                Return True
            End If

            ' ===========================
            '   CTRL + SHIFT + T  (New text file)
            ' ===========================
            If keyData = (Keys.Control Or Keys.Shift Or Keys.T) AndAlso GlobalShortcutsAllowed() Then
                NewTextFile_Click(sender, EventArgs.Empty)
                Return True
            End If

            ' ===========================
            '   F2 (Rename)
            ' ===========================
            If keyData = Keys.F2 AndAlso GlobalShortcutsAllowed() Then
                RenameFile_Click(sender, EventArgs.Empty)
                Return True
            End If

            ' ===========================
            '   CTRL + C (Copy)
            ' ===========================
            If keyData = (Keys.Control Or Keys.C) AndAlso GlobalShortcutsAllowed() Then
                CopySelected_Click(sender, EventArgs.Empty)
                Return True
            End If

            ' ===========================
            '   CTRL + V (Paste)
            ' ===========================
            If keyData = (Keys.Control Or Keys.V) AndAlso GlobalShortcutsAllowed() Then
                PasteSelected_Click(sender, EventArgs.Empty)
                Return True
            End If

            ' ===========================
            '   CTRL + X (Cut)
            ' ===========================
            If keyData = (Keys.Control Or Keys.X) AndAlso GlobalShortcutsAllowed() Then
                CutSelected_Click(sender, EventArgs.Empty)
                Return True
            End If

            ' ===========================
            '   CTRL + A (Select all)
            ' ===========================
            If keyData = (Keys.Control Or Keys.A) AndAlso GlobalShortcutsAllowed() Then
                SelectAllItems()
                lvFiles.Focus()
                Return True
            End If

            ' ===========================
            '   CTRL + D  OR  DELETE  (Delete)
            ' ===========================
            If (keyData = (Keys.Control Or Keys.D) OrElse keyData = Keys.Delete) AndAlso GlobalShortcutsAllowed() Then
                Delete_Click(sender, EventArgs.Empty)
                Return True
            End If

        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message,
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
        End Try

        Return False
    End Function


    Private Async Sub ExecuteCommand(command As String)

        ' Use regex to split by spaces but keep quoted substrings together
        Dim parts As String() = Regex.Matches(command, "[\""].+?[\""]|[^ ]+").
        Cast(Of Match)().
        Select(Function(m) m.Value.Trim(""""c)).
        ToArray()

        If parts.Length = 0 Then
            ShowStatus(StatusPad & IconDialog & "  No command entered.")
            Return
        End If

        Dim cmd As String = parts(0).ToLower()

        Select Case cmd

            Case "cd"

                If parts.Length > 1 Then
                    Dim newPath As String = String.Join(" ", parts.Skip(1)).Trim()
                    NavigateTo(newPath)
                Else
                    ShowStatus(StatusPad & IconDialog & " Usage: cd [directory] - cd C:\ ")
                End If

            Case "copy", "cp"
                If parts.Length > 2 Then
                    Dim source As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()
                    Dim destination As String = parts(parts.Length - 1).Trim()

                    ' Check if source file or directory exists
                    If Not (File.Exists(source) OrElse Directory.Exists(source)) Then

                        ShowStatus(StatusPad & IconError &
                                   " Copy failed:  Source """ &
                                   source &
                                   """ does not exist.  Paths with spaces must be enclosed in quotes.  Example: copy ""C:\folder A"" ""C:\folder B""")
                        Return
                    End If

                    ' Check if destination directory exists
                    If Not Directory.Exists(destination) Then

                        ShowStatus(StatusPad & IconError &
                                   " Copy failed.  Destination: """ &
                                   destination &
                                   """ does not exist.  Paths with spaces must be enclosed in quotes.  Example: copy ""C:\folder A"" ""C:\folder B""")

                        Return
                    End If

                    copyCts = New CancellationTokenSource()

                    Await CopyFileOrDirectory(source, destination, copyCts.Token)

                Else
                    ShowStatus(StatusPad & IconDialog & " Usage: copy [source] [destination] - e.g., copy C:\folder1\file.doc C:\folder2")
                End If


            Case "move", "mv"

                If parts.Length > 2 Then

                    Dim source As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()
                    Dim destination As String = parts(parts.Length - 1).Trim()

                    MoveFileOrDirectory(source, destination)

                Else
                    ShowStatus(StatusPad & IconDialog & " Usage: move [source] [destination] - move C:\folder1\directoryToMove C:\folder2\directoryToMove")
                End If

            Case "delete", "rm"

                If parts.Length > 1 Then

                    Dim pathToDelete As String = String.Join(" ", parts.Skip(1)).Trim()

                    DeleteFileOrDirectory(pathToDelete)

                Else
                    ShowStatus(StatusPad & IconDialog & " Usage: delete [file_or_directory]")
                End If

            Case "mkdir", "make", "md" ' You can use "mkdir" or "make" as the command
                If parts.Length > 1 Then
                    Dim directoryPath As String = String.Join(" ", parts.Skip(1)).Trim()

                    ' Validate the directory path
                    If String.IsNullOrWhiteSpace(directoryPath) Then
                        ShowStatus(StatusPad & IconDialog & " Usage: mkdir [directory_path] - e.g., mkdir C:\newfolder")
                        Return
                    End If

                    CreateDirectory(directoryPath)

                Else
                    ShowStatus(StatusPad & IconDialog & " Usage: mkdir [directory_path] - e.g., mkdir C:\newfolder")
                End If

            Case "rename"
                ' Rename file or directory

                If parts.Length > 2 Then

                    Dim sourcePath As String = String.Join(" ", parts.Skip(1).Take(parts.Length - 2)).Trim()

                    Dim newName As String = parts(parts.Length - 1).Trim()

                    RenameFileOrDirectory(sourcePath, newName)

                Else
                    ShowStatus(StatusPad & IconDialog & " Usage: rename [source_path] [new_name] - e.g., rename C:\folder\oldname.txt newname.txt")
                End If

            Case "text", "txt"

                If parts.Length > 1 Then

                    Dim rawPath = String.Join(" ", parts.Skip(1))

                    Dim filePath = NormalizeTextFilePath(rawPath)

                    If filePath Is Nothing Then

                        ShowStatus(StatusPad & IconDialog & " Usage: text [file_path]  e.g., text C:\example.txt")

                        Return

                    End If

                    CreateTextFile(filePath)

                    Return

                End If

                ' No file path provided
                Dim destDir = currentFolder

                ' Validate destination folder
                If String.IsNullOrWhiteSpace(destDir) OrElse Not Directory.Exists(destDir) Then

                    ShowStatus(StatusPad & IconWarning & " Invalid folder. Cannot create file.")

                    Return

                End If

                ' Ensure unique file name
                Dim newFilePath = GetUniqueFilePath(destDir, "New Text File", ".txt")

                Try

                    ' Create the file with initial content
                    File.WriteAllText(newFilePath, $"Created on {DateTime.Now:G}")

                    ShowStatus(StatusPad & IconSuccess & " Text file created: " & newFilePath)

                    ' Refresh the folder view so the user sees the new file
                    NavigateTo(destDir)

                    ' Open the newly created file
                    GoToFolderOrOpenFile(newFilePath)

                Catch ex As Exception
                    ShowStatus(StatusPad & IconError & " Failed to create text file: " & ex.Message)
                    Debug.WriteLine("Text Command Error: " & ex.Message)
                End Try

            Case "help", "man"
                ShowHelpFile()
                Return

            Case "open"

                ' --- 1. If the user typed: open "C:\path\to\something"
                If parts.Length > 1 Then

                    Dim targetPath As String =
                        String.Join(" ", parts.Skip(1)).
                        Trim().
                        Trim(""""c)

                    HandleOpenPath(targetPath)
                    Return

                End If

                ' --- 2. If no path was typed, use the selected item in the ListView
                If lvFiles.SelectedItems.Count = 0 Then
                    ShowStatus(StatusPad & IconDialog &
                               "  Usage: open [file_or_folder]  — or select an item first.")
                    Return
                End If

                Dim selected As ListViewItem = lvFiles.SelectedItems(0)
                Dim fullPath As String = selected.Tag.ToString()

                HandleOpenPath(fullPath)
                Return

            'Case "find", "search"

            '    If parts.Length > 1 Then

            '        Dim searchTerm As String = String.Join(" ", parts.Skip(1)).Trim()

            '        If String.IsNullOrWhiteSpace(searchTerm) Then
            '            ShowStatus(StatusPad & IconDialog & " Usage: find [search_term] - e.g., find document")
            '            Return
            '        End If

            '        ShowStatus(StatusPad & IconSearch & " Searching for: " & searchTerm)

            '        'SearchInCurrentFolder(searchTerm)
            '        OnlySearchForFilesInCurrentFolder(searchTerm)

            '        ' Reset index for new search
            '        SearchIndex = 0

            '        ' Auto-select first result if available
            '        If SearchResults.Count > 0 Then
            '            lvFiles.SelectedItems.Clear()
            '            SelectListViewItemByPath(SearchResults(0))

            '            Dim nextPath As String = SearchResults(SearchIndex)
            '            Dim fileName As String = Path.GetFileNameWithoutExtension(nextPath)

            '            txtAddressBar.Focus()

            '            ShowStatus(
            '                StatusPad & IconSearch &
            '                "    Result " &
            '                (SearchIndex + 1) &
            '                " of " &
            '                SearchResults.Count &
            '                $"     ""{fileName}""       Next result  -  F3      Open  -  Ctrl + O      Reset  -  Esc"
            '            )


            '        Else
            '            ShowStatus(StatusPad & IconDialog & "  No results found for: " & searchTerm)
            '        End If

            '    Else
            '        ShowStatus(StatusPad & IconDialog & "  Usage: find [search_term] - e.g., find document")
            '    End If

            '    Return


            Case "find", "search"

                If parts.Length > 1 Then

                    Dim searchTerm As String = String.Join(" ", parts.Skip(1)).Trim()

                    If String.IsNullOrWhiteSpace(searchTerm) Then
                        ShowStatus(
                            StatusPad & IconDialog &
                            "  Usage: find [search_term]   e.g., find document"
                        )
                        Return
                    End If

                    ' Announce search
                    ShowStatus(StatusPad & IconSearch & "  Searching for: " & searchTerm)

                    ' Perform search
                    OnlySearchForFilesInCurrentFolder(searchTerm)


                    ' Reset index for new search
                    SearchIndex = 0

                    ' If results exist, auto-select the first one
                    If SearchResults.Count > 0 Then

                        lvFiles.SelectedItems.Clear()
                        SelectListViewItemByPath(SearchResults(0))

                        Dim nextPath As String = SearchResults(SearchIndex)
                        Dim fileName As String = Path.GetFileNameWithoutExtension(nextPath)

                        'txtAddressBar.Focus()
                        lvFiles.Focus()

                        HighlightSearchMatches()
                        'HighlightCurrentResult()


                        ' Unified search HUD
                        ShowStatus(
                            StatusPad & IconSearch &
                            $"  Result {SearchIndex + 1} of {SearchResults.Count}    " &
                            $""“{fileName}”"    Next  F3    Open  Ctrl+O    Reset  Esc"
                        )
                    Else
                        ShowStatus(
                            StatusPad & IconDialog &
                            "  No results found for: " & searchTerm
                        )
                    End If

                Else
                    ShowStatus(
                        StatusPad & IconDialog &
                        "  Usage: find [search_term]   e.g., find document"
                    )
                End If

                Return

            Case "findnext", "searchnext", "next"

                HandleFindNextCommand()


                'If SearchResults.Count = 0 Then
                '    ShowStatus(StatusPad & IconDialog & "  No previous search results. Use 'find [search_term]' to start a search.")
                '    Return
                'End If

                '' Advance index
                'SearchIndex += 1

                '' Wrap around
                'If SearchIndex >= SearchResults.Count Then
                '    SearchIndex = 0
                'End If

                '' Select the next result
                'lvFiles.SelectedItems.Clear()
                'Dim nextPath As String = SearchResults(SearchIndex)
                'SelectListViewItemByPath(nextPath)

                'ShowStatus(
                '    StatusPad & IconSearch &
                '    " Showing result " &
                '    (SearchIndex + 1) &
                '    " of " &
                '    SearchResults.Count &
                '    " To show the next result, enter: findnext")

                Return

            Case "exit", "quit", "close", "bye", "shutdown", "logoff", "signout", "poweroff", "halt", "end", "terminate", "stop", "leave", "farewell", "adios", "ciao", "sayonara", "goodbye", "later"

                ' Confirm exit
                If MessageBox.Show("Are you sure you want to exit?",
                                   "Confirm Exit",
                                   MessageBoxButtons.YesNo,
                                   MessageBoxIcon.Question) = DialogResult.Yes Then
                    Application.Exit()
                Else
                    ShowStatus("Exit cancelled.")
                End If

                Return

            Case Else

                ' Is the input a folder?
                If Directory.Exists(command) Then
                    ' Go to that folder.
                    NavigateTo(command)
                    Return
                    ' Is the input a file?
                ElseIf File.Exists(command) Then
                    OpenFileWithDefaultApp(command)
                    RestoreAddressBar()
                    Return
                Else
                    ' The input isn't a folder or a file,
                    ' at this point, the interpreter treats it as an unknown command.
                    ShowStatus(StatusPad & IconQuestion &
                               $"  Unknown command:     ""{cmd}""       Esc to reset.       Type ""help"" for a list of commands.")
                    Return
                End If

        End Select

    End Sub

    Private Sub ShowHelpFile()
        Dim helpFilePath As String =
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cli_help.txt")

        EnsureHelpFileExists(helpFilePath)
        OpenFileWithDefaultApp(helpFilePath)
        RestoreAddressBar()
    End Sub

    Private Sub HandleOpenPath(path As String)

        If File.Exists(path) Then
            OpenFileWithDefaultApp(path)
            RestoreAddressBar()
            Return
        End If

        If Directory.Exists(path) Then
            NavigateTo(path)
            Return
        End If

        ShowStatus(StatusPad & IconError & "  Path not found: " & path)

    End Sub

    Private Sub RestoreAddressBar()
        txtAddressBar.Text = currentFolder
        txtAddressBar.Focus()
        PlaceCaretAtEndOfAddressBar()
    End Sub

    Private Sub OpenFileWithDefaultApp(filePath As String)
        ' Open file with default application.

        Try

            Dim psi As New ProcessStartInfo(filePath) With {
                .UseShellExecute = True
            }

            Process.Start(psi)

            Dim fileName As String = Path.GetFileNameWithoutExtension(filePath)
            ShowStatus(StatusPad & IconOpen & $"  Opened:   ""{fileName}""")

        Catch ex As Exception
            ShowStatus(StatusPad & IconError & " Cannot open: " & ex.Message)
            Debug.WriteLine("OpenFileWithDefaultApp: Error opening file: " & ex.Message)
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

        OpenSelectedItem()

    End Sub

    Private Sub NewFolder_Click(sender As Object, e As EventArgs)
        ' Create new folder in current directory - Mouse right-click context menu for lvFiles or button click 

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

            ShowStatus(StatusPad & IconNewFolder & " Created folder: " & di.Name)

        Catch ex As Exception

            ShowStatus(StatusPad & IconError & " New folder creation failed. " & ex.Message)

            Debug.WriteLine("NewFolder_Click Error: " & ex.Message)

        End Try

    End Sub

    Private Sub NewTextFile_Click(sender As Object, e As EventArgs)

        Dim destDir As String = currentFolder

        ' Validate destination folder
        If String.IsNullOrWhiteSpace(destDir) OrElse Not Directory.Exists(destDir) Then
            ShowStatus(StatusPad & IconWarning & " Invalid folder. Cannot create file.")
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

            ShowStatus(StatusPad & IconSuccess & " Text file created: " & newFilePath)

            ' Refresh the folder view so the user sees the new file
            NavigateTo(destDir)

            ' Open the newly created file
            GoToFolderOrOpenFile(newFilePath)

        Catch ex As Exception

            ShowStatus(StatusPad & IconError & " New text file creation failed. " & ex.Message)

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

        ShowStatus(StatusPad & IconCut & " Cut to clipboard: " & _clipboardPath)

    End Sub

    Private Sub CopySelected_Click(sender As Object, e As EventArgs)

        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        Dim path As String = CStr(lvFiles.SelectedItems(0).Tag)

        ' Validate that the selected item actually exists
        If Not PathExists(path) Then
            ShowStatus(StatusPad & IconWarning & " Copy failed: Selected item does not exist.")
            Exit Sub
        End If

        ' Update internal clipboard
        _clipboardPath = path
        _clipboardIsCut = False

        ShowStatus(StatusPad & IconCopy & " Copied to clipboard: " & _clipboardPath)

        UpdateEditButtonsAndMenus()

        UpdateFileButtonsAndMenus()


    End Sub

    Private Async Sub PasteSelected_Click(sender As Object, e As EventArgs)
        ' ------------------------------------------------------------
        ' Paste handler with free-space checks for file + directory
        ' ------------------------------------------------------------

        If String.IsNullOrWhiteSpace(_clipboardPath) Then
            ShowStatus(StatusPad & IconError & " Paste failed: No item in clipboard.")
            Exit Sub
        End If

        Dim sourcePath As String = _clipboardPath
        Dim destDir As String = currentFolder

        If String.IsNullOrWhiteSpace(destDir) OrElse Not Path.IsPathRooted(destDir) Then
            ShowStatus(StatusPad & IconError & " Paste failed: Invalid destination folder.")
            Exit Sub
        End If

        Dim name As String = Path.GetFileName(sourcePath)
        Dim destPath As String = Path.Combine(destDir, name)

        Dim isFile As Boolean = File.Exists(sourcePath)
        Dim isDir As Boolean = Directory.Exists(sourcePath)

        If Not isFile AndAlso Not isDir Then
            ShowStatus(StatusPad & IconError & " Paste failed: Source not found.")
            Exit Sub
        End If

        ' Auto‑rename only when copying (not cutting)
        If Not _clipboardIsCut Then
            destPath = ResolveDestinationPathWithAutoRename(destPath, isDir)
        End If

        copyCts = New CancellationTokenSource()
        Dim ct = copyCts.Token

        Try
            ct.ThrowIfCancellationRequested()

            ' ------------------------------------------------------------
            ' FILE PASTE
            ' ------------------------------------------------------------
            If isFile Then
                Dim src As New FileInfo(sourcePath)

                ' Free-space check
                If Not _clipboardIsCut AndAlso Not HasEnoughSpace(src, destPath) Then
                    ShowStatus(StatusPad & IconError & " Not enough space to paste file: " & src.Name)
                    Exit Sub
                End If

                Dim result As CopyResult = Await CopyFile(sourcePath, destPath, ct)

                If result.Success Then
                    ShowStatus(StatusPad & IconPaste & " File pasted into " & txtAddressBar.Text)
                End If

                ' ------------------------------------------------------------
                ' DIRECTORY PASTE
                ' ------------------------------------------------------------
            ElseIf isDir Then

                ' Free-space check (sum of all files)
                If Not _clipboardIsCut Then
                    Dim totalSize As Long =
                    Directory.EnumerateFiles(sourcePath, "*", SearchOption.AllDirectories).
                    Sum(Function(f) New FileInfo(f).Length)

                    Dim root = Path.GetPathRoot(destPath)
                    Dim drive As New DriveInfo(root)

                    If totalSize > drive.AvailableFreeSpace Then
                        ShowStatus(StatusPad & IconError & " Not enough space to paste folder.")
                        Exit Sub
                    End If
                End If

                Dim result As CopyResult = Await CopyDirectory(sourcePath, destPath, ct)

                ShowStatus(StatusPad & IconPaste &
                       $" Pasted folder: {result.FilesCopied} files, {result.FilesSkipped} skipped.")
            End If

            ' Reset clipboard state
            _clipboardPath = Nothing
            _clipboardIsCut = False

            Await PopulateFiles(destDir)

            ' ------------------------------------------------------------
            ' Select the newly pasted item
            ' ------------------------------------------------------------
            Dim pastedName As String = Path.GetFileName(destPath)

            For Each item As ListViewItem In lvFiles.Items
                If String.Equals(item.Text, pastedName, StringComparison.OrdinalIgnoreCase) Then
                    item.Selected = True
                    item.Focused = True
                    item.EnsureVisible()
                    Exit For
                End If
            Next

            lvFiles.Focus()

            ResetCutVisuals()
            UpdateEditButtonsAndMenus()

        Catch ex As OperationCanceledException
            ShowStatus(StatusPad & IconWarning & " Paste canceled.")

        Catch ex As Exception
            ShowStatus(StatusPad & IconError & " Paste failed: " & ex.Message)
        End Try

    End Sub


    Private Sub RenameFile_Click(sender As Object, e As EventArgs)
        ' Rename selected file or folder - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub
        lvFiles.SelectedItems(0).BeginEdit() ' triggers inline rename

    End Sub

    Private Async Sub Delete_Click(sender As Object, e As EventArgs)

        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        Dim selected = lvFiles.SelectedItems.Cast(Of ListViewItem)().ToList()
        Dim count As Integer = selected.Count

        ' Confirm deletion
        Dim msg As String =
        If(count = 1,
           $"Are you sure you want to delete '{selected(0).Text}'?",
           $"Are you sure you want to delete these {count} items?")

        Dim confirm = MessageBox.Show(msg, "Delete",
                                  MessageBoxButtons.YesNo,
                                  MessageBoxIcon.Warning)

        If confirm <> DialogResult.Yes Then Exit Sub

        ShowStatus(StatusPad & IconDelete & $"  Deleting {count} items.")

        ' Remember index of first selected item for post-refresh selection
        Dim firstIndex As Integer = selected(0).Index

        For Each item In selected
            Dim fullPath As String = CStr(item.Tag)

            ' Reject relative paths
            If Not Path.IsPathRooted(fullPath) Then
                ShowStatus(StatusPad & IconWarning &
                       " Delete failed: Path must be absolute. Example: C:\folder")
                Continue For
            End If

            ' Protected path check
            If IsProtectedPathOrFolder(fullPath) Then
                ShowStatus(StatusPad & IconProtect &
                       " Deletion prevented for protected path: " & fullPath)
                Continue For
            End If

            Try

                If Directory.Exists(fullPath) Then
                    Directory.Delete(fullPath, recursive:=True)
                    ShowStatus(StatusPad & IconDelete &
                           " Deleted folder: " & item.Text)

                ElseIf File.Exists(fullPath) Then
                    File.Delete(fullPath)
                    ShowStatus(StatusPad & IconDelete &
                           " Deleted file: " & item.Text)

                Else
                    ShowStatus(StatusPad & IconWarning &
                           " Path not found: " & fullPath)
                End If

            Catch ex As Exception
                ShowStatus(StatusPad & IconError &
                       " Delete failed: " & ex.Message)
                Debug.WriteLine("Delete_Click Error: " & ex.Message)
            End Try
        Next

        ' Refresh the folder
        Await PopulateFiles(currentFolder)

        ' ------------------------------------------------------------
        ' Select the next logical item (Explorer-style)
        ' ------------------------------------------------------------
        If lvFiles.Items.Count > 0 Then
            Dim newIndex As Integer = Math.Min(firstIndex, lvFiles.Items.Count - 1)
            lvFiles.Items(newIndex).Selected = True
            lvFiles.Items(newIndex).Focused = True
            lvFiles.Items(newIndex).EnsureVisible()
        End If

        lvFiles.Focus()
        UpdateEditButtonsAndMenus()

    End Sub


    Private Sub CopyFileName_Click(sender As Object, e As EventArgs)
        ' Copy selected file name to clipboard - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        Clipboard.SetText(lvFiles.SelectedItems(0).Text)

        ShowStatus(StatusPad & IconCopy & " Copied File Name " & lvFiles.SelectedItems(0).Text)

    End Sub

    Private Sub CopyFilePath_Click(sender As Object, e As EventArgs)
        ' Copy selected file path to clipboard - Mouse right-click context menu for lvFiles

        ' Is a file or folder selected?
        If lvFiles.SelectedItems.Count = 0 Then Exit Sub

        ' Copy the full path stored in the Tag property to clipboard
        Clipboard.SetText(CStr(lvFiles.SelectedItems(0).Tag))

        ShowStatus(StatusPad & IconCopy & " Copied File Path " & lvFiles.SelectedItems(0).Tag)

    End Sub


    Private Sub RefreshCurrentFolder()
        ' Refresh the current folder view
        NavigateTo(currentFolder, recordHistory:=False)
        UpdateTreeRoots()
    End Sub

    Private Sub NavigateToParent()
        Dim parent = Directory.GetParent(currentFolder)
        If parent IsNot Nothing Then
            NavigateTo(parent.FullName, recordHistory:=True)
        End If
    End Sub

    Private Sub SelectAllItems()
        For Each item As ListViewItem In lvFiles.Items
            item.Selected = True
        Next
    End Sub

    Private Sub ExpandOneLevel()

        Dim node As TreeNode = tvFolders.SelectedNode
        If node Is Nothing Then Exit Sub

        ' If the node is collapsed, expand it (one level)
        If Not node.IsExpanded Then
            ShowStatus(StatusPad & IconDialog & "  Expanding folder...")

            tvFolders.BeginUpdate()
            node.Expand()
            tvFolders.EndUpdate()

            ShowStatus(StatusPad & IconSuccess & "  Expanded: " & node.FullPath)
            Return
        End If

        ' If the node is expanded, move to its first child (like Explorer)
        If node.Nodes.Count > 0 Then
            tvFolders.SelectedNode = node.Nodes(0)
            tvFolders.SelectedNode.EnsureVisible()

            ShowStatus(StatusPad & IconSuccess & "  Navigated into: " & tvFolders.SelectedNode.FullPath)
            Return
        End If

        ' If expanded but no children, do nothing
        ShowStatus(StatusPad & IconDialog & "  No subfolders to expand.")

    End Sub

    Private Sub CollapseOneLevel()

        Dim node As TreeNode = tvFolders.SelectedNode
        If node Is Nothing Then Exit Sub

        ' If the node is expanded, collapse it (one level)
        If node.IsExpanded Then
            ShowStatus(StatusPad & IconDialog & "  Collapsing folder...")

            tvFolders.BeginUpdate()
            node.Collapse()
            tvFolders.EndUpdate()

            ShowStatus(StatusPad & IconSuccess & "  Collapsed: " & node.FullPath)
            Return
        End If

        ' If the node is collapsed, move selection to its parent
        If node.Parent IsNot Nothing Then
            tvFolders.SelectedNode = node.Parent
            tvFolders.SelectedNode.EnsureVisible()

            ShowStatus(StatusPad & IconSuccess & "  Moved to parent: " & node.Parent.FullPath)
            Return
        End If

        ' If it's a root node and already collapsed, do nothing
        ShowStatus(StatusPad & IconDialog & "  Already at the top level.")

    End Sub

    Private Sub ToggleExpandCollapse()

        Dim node As TreeNode = tvFolders.SelectedNode
        If node Is Nothing Then Exit Sub

        If node.IsExpanded Then
            ' Collapse one level
            tvFolders.BeginUpdate()
            node.Collapse()
            tvFolders.EndUpdate()

        Else
            ' Expand one level
            tvFolders.BeginUpdate()
            node.Expand()
            tvFolders.EndUpdate()

        End If

    End Sub

    Private Sub ToggleFullScreen()
        If Me.FormBorderStyle = FormBorderStyle.None Then
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.WindowState = FormWindowState.Normal
        Else
            Me.FormBorderStyle = FormBorderStyle.None
            Me.WindowState = FormWindowState.Maximized
        End If
    End Sub


    Private Sub ExpandNode_LazyLoad(node As TreeNode)

        ' Only lazy-load if placeholder exists
        If node.Nodes.Count = 1 AndAlso node.Nodes(0).Text = "Loading..." Then

            tvFolders.BeginUpdate()

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

        tvFolders.EndUpdate()

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

                tvFolders.BeginUpdate()
                tvFolders.Nodes.Remove(node)
                tvFolders.EndUpdate()
                ShowStatus(StatusPad & IconWarning & " Drive is not ready and has been removed.")
                Return
            End If

        Catch ex As Exception
            ' Handle any exceptions when accessing DriveInfo
            ShowStatus(StatusPad & IconError & " NavTree: Error accessing drive: " & ex.Message)
            Debug.WriteLine("NavTree AfterSelect: Error accessing drive: " & ex.Message)
            Return
        End Try

        ' If the drive is ready, navigate to the folder
        NavigateTo(path2Nav)

    End Sub

    Private Async Function PopulateFiles(path As String) As Task
        lvFiles.BeginUpdate()
        lvFiles.Items.Clear()

        Try

            ' -------------------------------
            ' Load directories asynchronously
            ' -------------------------------
            Dim directories = Await Task.Run(Function()

                                                 Dim dirList As New List(Of String)()
                                                 Try

                                                     dirList.AddRange(
                                                     Directory.GetDirectories(path).
                                                     Where(Function(d) _
                                                     ShowHiddenFiles OrElse
                                                     Not (New DirectoryInfo(d).Attributes And (FileAttributes.Hidden Or FileAttributes.System) <> 0)))

                                                 Catch ex As UnauthorizedAccessException
                                                     ShowStatus(StatusPad & IconWarning & " Access denied to some directories in: " & path)
                                                     Debug.WriteLine($"PopulateFiles - {path} - Directories - UnauthorizedAccessException - {ex.Message}") ' Log exception details
                                                 Catch ex As Exception
                                                     ShowStatus(StatusPad & IconError & " An error occurred: " & ex.Message)
                                                     Debug.WriteLine($"PopulateFiles - {path} - Directories - Error - {ex.Message}")
                                                 End Try

                                                 Return dirList ' Ensure that an empty list is returned if an exception occurs
                                             End Function)



            Dim itemsToAdd As New List(Of ListViewItem)


            For Each mDir In directories
                Dim di As New DirectoryInfo(mDir)
                Dim item As New ListViewItem(di.Name)

                item.SubItems.Add("Folder")
                item.SubItems.Add("") ' No size for folders
                item.SubItems.Add(di.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
                item.Tag = di.FullName
                item.ImageKey = "Folder"

                itemsToAdd.Add(item)
            Next

            ' ---------------------------
            ' Load files asynchronously
            ' ---------------------------
            Dim files = Await Task.Run(Function()

                                           Dim fileList As New List(Of String)
                                           Try

                                               fileList.AddRange(
                                               Directory.GetFiles(path).
                                               Where(Function(f) _
                                               ShowHiddenFiles OrElse
                                               (New FileInfo(f).Attributes And
                                               (FileAttributes.Hidden Or FileAttributes.System)) = 0))

                                           Catch ex As UnauthorizedAccessException
                                               ShowStatus(StatusPad & IconError & " Access denied to some files in: " & path)
                                               Debug.WriteLine($"PopulateFiles - {path} - Files - UnauthorizedAccessException - {ex.Message}")
                                           Catch ex As Exception
                                               ShowStatus(StatusPad & IconError & " An error occurred: " & ex.Message)
                                               Debug.WriteLine($"PopulateFiles - {path} - Files - Error - {ex.Message}")

                                           End Try

                                           Return fileList
                                       End Function)

            For Each file In files
                Dim fi As New FileInfo(file)
                Dim item As New ListViewItem(fi.Name)

                Dim ext = fi.Extension.ToLowerInvariant()
                Dim category = fileTypeMap.GetValueOrDefault(ext, "Document")

                ' Category column
                item.SubItems.Add(category)

                ' Size + date
                item.SubItems.Add(FormatSize(fi.Length))
                item.SubItems.Add(fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
                item.Tag = fi.FullName

                ' Image assignment
                Select Case category
                    Case "Audio" : item.ImageKey = "Music"
                    Case "Image" : item.ImageKey = "Pictures"
                    Case "Document" : item.ImageKey = "Documents"
                    Case "Video" : item.ImageKey = "Videos"
                    Case "Archive" : item.ImageKey = "Downloads"
                    Case "Executable" : item.ImageKey = "Executable"
                    Case "Shortcut" : item.ImageKey = "Shortcut"
                    Case Else : item.ImageKey = "Documents"
                End Select

                itemsToAdd.Add(item)
            Next

            lvFiles.Items.AddRange(itemsToAdd.ToArray())

            ShowStatus(StatusPad & lvFiles.Items.Count & "  items")

        Catch ex As Exception
            ShowStatus(StatusPad & IconError & $" Error: {ex.Message}")
            Debug.WriteLine($"General Error: {ex.Message}")

        Finally
            lvFiles.EndUpdate()
        End Try

    End Function

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

        _isRenaming = False


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

                ShowStatus(StatusPad & IconSuccess & " Renamed Folder to: " & newName)

            ElseIf File.Exists(oldPath) Then

                File.Move(oldPath, newPath)

                ShowStatus(StatusPad & IconSuccess & " Renamed File to: " & newName)

            End If

            item.Tag = newPath

        Catch ex As Exception

            e.CancelEdit = True

            ShowStatus(StatusPad & IconError & " Rename Failed. " & ex.Message)

            Debug.WriteLine("RenameFileOrFolder_AfterLabelEdit Error: " & ex.Message)

        End Try

    End Sub

    Private Async Function CopyFile(
        source As String,
        destination As String,
        ct As CancellationToken
    ) As Task(Of CopyResult)
        ' ------------------------------------------------------------
        ' Copy a single file with auto‑rename + free-space check
        ' ------------------------------------------------------------

        Dim result As New CopyResult()
        Dim src As New FileInfo(source)

        ' Proactive free-space check
        If Not HasEnoughSpace(src, destination) Then
            result.FilesSkipped = 1
            result.Errors.Add("Not enough space for: " & source)
            ShowStatus(StatusPad & IconError & " Not enough space, skipping: " & src.Name)
            Return result
        End If

        ' Auto‑rename if needed
        Dim finalDest As String = destination
        If File.Exists(finalDest) Then
            finalDest = ResolveDestinationPathWithAutoRename(finalDest, isDirectory:=False)
            result.FilesRenamed = 1
            ShowStatus(StatusPad & IconDialog &
                   " Auto‑renamed → " & Path.GetFileName(finalDest))
            Debug.WriteLine("Auto‑renamed to: " & finalDest)
        End If

        Try
            ct.ThrowIfCancellationRequested()

            Await Task.Run(Sub()
                               ct.ThrowIfCancellationRequested()
                               File.Copy(source, finalDest, overwrite:=False)
                           End Sub, ct)

            result.FilesCopied = 1
            ShowStatus(StatusPad & IconCopy & " Copied file: " & finalDest)
            Debug.WriteLine("Copied file: " & finalDest)

        Catch ex As IOException When ex.HResult = &H80070070
            result.FilesSkipped = 1
            result.Errors.Add("Disk full while copying: " & source)
            ShowStatus(StatusPad & IconError & " Copy failed: Not enough disk space.")
            MessageBox.Show("There is not enough space on the destination drive.",
                        "Disk Full",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

        Catch ex As IOException
            result.FilesSkipped = 1
            result.Errors.Add("I/O error: " & source & " - " & ex.Message)
            ShowStatus(StatusPad & IconError & " I/O error copying: " & Path.GetFileName(source))

        Catch ex As UnauthorizedAccessException
            result.FilesSkipped = 1
            result.Errors.Add("Unauthorized: " & source)
            ShowStatus(StatusPad & IconError & " Unauthorized: " & Path.GetFileName(source))

        Catch ex As Exception
            result.FilesSkipped = 1
            result.Errors.Add("Copy failed: " & source & " - " & ex.Message)
            ShowStatus(StatusPad & IconError & " Copy failed: " & Path.GetFileName(source))
        End Try

        Return result
    End Function

    Public Async Function CopyDirectory(
        sourceDir As String,
        destDir As String,
        ct As CancellationToken
    ) As Task(Of CopyResult)

        ' ------------------------------------------------------------
        ' Recursive directory copy with free-space checks + taxonomy
        ' ------------------------------------------------------------

        Dim result As New CopyResult()
        Dim dirInfo As New DirectoryInfo(sourceDir)

        If Not dirInfo.Exists Then
            result.Errors.Add("Source directory not found: " & sourceDir)
            ShowStatus(StatusPad & IconError & " Source directory not found: " & sourceDir)
            Return result
        End If

        Try
            ct.ThrowIfCancellationRequested()

            ' Create destination directory
            Try
                Directory.CreateDirectory(destDir)
                result.DirectoriesCreated += 1
            Catch ex As Exception
                result.Errors.Add("Failed to create directory: " & destDir & " - " & ex.Message)
                ShowStatus(StatusPad & IconError & " Failed to create directory: " & destDir)
                Return result
            End Try

            ' ------------------------------------------------------------
            ' Copy files
            ' ------------------------------------------------------------
            For Each srcFile In dirInfo.GetFiles()
                ct.ThrowIfCancellationRequested()

                Dim targetFile = Path.Combine(destDir, srcFile.Name)

                ' Locked file skip
                If IsFileLocked(srcFile) Then
                    result.FilesSkipped += 1
                    ShowStatus(StatusPad & IconWarning & " Skipped locked file: " & srcFile.FullName)
                    Continue For
                End If

                ' Free-space check
                If Not HasEnoughSpace(srcFile, targetFile) Then
                    result.FilesSkipped += 1
                    result.Errors.Add("Not enough space for: " & srcFile.FullName)
                    ShowStatus(StatusPad & IconError & " Not enough space, skipping: " & srcFile.Name)
                    Continue For
                End If

                Try
                    Await Task.Run(Sub()
                                       ct.ThrowIfCancellationRequested()
                                       srcFile.CopyTo(targetFile, overwrite:=False)
                                   End Sub, ct)

                    result.FilesCopied += 1
                    ShowStatus(StatusPad & IconCopy & " Copied file: " & targetFile)

                Catch ex As IOException When ex.HResult = &H80070070
                    result.FilesSkipped += 1
                    result.Errors.Add("Disk full while copying: " & srcFile.FullName)
                    ShowStatus(StatusPad & IconError & " Disk full, skipping: " & srcFile.Name)

                Catch ex As IOException When ex.HResult = &H80070050
                    result.FilesSkipped += 1
                    result.Errors.Add("File already exists: " & srcFile.FullName)
                    ShowStatus(StatusPad & IconWarning & " Skipped existing file: " & srcFile.Name)

                Catch ex As IOException
                    result.FilesSkipped += 1
                    result.Errors.Add("I/O error: " & srcFile.FullName & " - " & ex.Message)
                    ShowStatus(StatusPad & IconError & " I/O error copying: " & srcFile.Name)

                Catch ex As UnauthorizedAccessException
                    result.FilesSkipped += 1
                    result.Errors.Add("Unauthorized: " & srcFile.FullName)
                    ShowStatus(StatusPad & IconError & " Unauthorized: " & srcFile.Name)

                Catch ex As Exception
                    result.FilesSkipped += 1
                    result.Errors.Add("Copy failed: " & srcFile.FullName & " - " & ex.Message)
                    ShowStatus(StatusPad & IconError & " Copy failed: " & srcFile.Name)
                End Try
            Next

            ' ------------------------------------------------------------
            ' Copy subdirectories in parallel
            ' ------------------------------------------------------------
            Dim subTasks As New List(Of Task(Of CopyResult))()

            For Each subDir In dirInfo.GetDirectories()
                Dim newDest = Path.Combine(destDir, subDir.Name)
                subTasks.Add(CopyDirectory(subDir.FullName, newDest, ct))
            Next

            Dim subResults = Await Task.WhenAll(subTasks)

            For Each r In subResults
                result.FilesCopied += r.FilesCopied
                result.FilesSkipped += r.FilesSkipped
                result.DirectoriesCreated += r.DirectoriesCreated
                result.Errors.AddRange(r.Errors)
            Next

        Catch ex As OperationCanceledException
            result.Errors.Add("Canceled: " & ex.Message)
            ShowStatus(StatusPad & IconWarning & " Directory copy canceled.")
            Debug.WriteLine("CopyDirectory canceled: " & ex.Message)

        Catch ex As Exception
            result.Errors.Add("Error: " & ex.Message)
            ShowStatus(StatusPad & IconError & " Copy failed: " & ex.Message)
            Debug.WriteLine("CopyDirectory error: " & ex.Message)
        End Try

        Return result
    End Function

    Private Async Function CopyFileOrDirectory(
        source As String,
        destination As String,
        ct As CancellationToken
    ) As Task

        ' ------------------------------------------------------------
        ' Wrapper for CLI: copy file or directory using unified engine
        ' ------------------------------------------------------------

        Try
            ct.ThrowIfCancellationRequested()

            ' -------------------------
            ' FILE COPY
            ' -------------------------
            If File.Exists(source) Then

                Dim destFile As String = Path.Combine(destination, Path.GetFileName(source))

                ' Free-space check
                Dim srcInfo As New FileInfo(source)
                If Not HasEnoughSpace(srcInfo, destFile) Then
                    ShowStatus(StatusPad & IconError &
                           " Not enough space to copy file: " & srcInfo.Name)
                    Return
                End If

                Dim result As CopyResult = Await CopyFile(source, destFile, ct)

                If result.Success Then
                    NavigateTo(destination)

                    txtAddressBar.Focus()
                    PlaceCaretAtEndOfAddressBar()

                End If

                Return
            End If

            ' -------------------------
            ' DIRECTORY COPY
            ' -------------------------
            If Directory.Exists(source) Then

                Dim targetDir As String = Path.Combine(destination, Path.GetFileName(source))

                ' NEW: Auto‑rename directory if needed
                If Directory.Exists(targetDir) Then
                    Dim newDir = ResolveDestinationPathWithAutoRename(targetDir, isDirectory:=True)
                    ShowStatus(StatusPad & IconDialog &
                           " Auto‑renamed folder → " & Path.GetFileName(newDir))
                    Debug.WriteLine("Auto‑renamed directory to: " & newDir)
                    targetDir = newDir
                End If

                ' Free-space check (sum of all files)
                Dim totalSize As Long =
                Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories).
                Sum(Function(f) New FileInfo(f).Length)

                Dim root = Path.GetPathRoot(targetDir)
                Dim drive As New DriveInfo(root)

                If totalSize > drive.AvailableFreeSpace Then
                    ShowStatus(StatusPad & IconError &
                           " Not enough space to copy folder.")
                    Return
                End If

                Dim result As CopyResult = Await CopyDirectory(source, targetDir, ct)

                If result.Success Then
                    Dim parent = Directory.GetParent(targetDir)
                    If parent IsNot Nothing Then
                        NavigateTo(parent.FullName)

                        txtAddressBar.Focus()
                        PlaceCaretAtEndOfAddressBar()

                    End If
                End If


                Return
            End If

            ' -------------------------
            ' SOURCE NOT FOUND
            ' -------------------------
            ShowStatus(StatusPad & IconError &
                   " Copy failed: Source does not exist or is not a valid file or directory.")

        Catch ex As OperationCanceledException
            ShowStatus(StatusPad & IconWarning & " Copy operation canceled.")

        Catch ex As Exception
            ShowStatus(StatusPad & IconError & " Copy failed: " & ex.Message)
        End Try

    End Function


    Private Async Sub NavigateTo(path As String, Optional recordHistory As Boolean = True)
        ' Navigate to the specified folder path.
        ' Updates the current folder, path textbox, and file list.

        ' Validate input
        If String.IsNullOrWhiteSpace(path) Then Exit Sub

        ' Validate that the folder exists
        If Not Directory.Exists(path) Then
            MessageBox.Show("Folder not found: " & path, "Navigation", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ShowStatus(StatusPad & IconWarning & " Folder not found: " & path)
            Exit Sub
        End If

        If txtAddressBar.InvokeRequired Then
            txtAddressBar.Invoke(Sub() NavigateTo(path, recordHistory))
            Return
        End If

        ShowStatus(StatusPad & IconNavigate & " Navigated To: " & path)

        currentFolder = path
        txtAddressBar.Text = path
        Await PopulateFiles(path) ' Await the async method

        If recordHistory Then
            ' Trim forward history if we branch
            If _historyIndex >= 0 AndAlso _historyIndex < _history.Count - 1 Then
                _history.RemoveRange(_historyIndex + 1, _history.Count - (_historyIndex + 1))
            End If
            _history.Add(path)
            _historyIndex = _history.Count - 1
            UpdateNavButtons()
        End If

        UpdateFileButtonsAndMenus()

        UpdateEditButtonsAndMenus()
    End Sub

    Private Sub GoToFolderOrOpenFile(FileOrFolder As String)
        ' Navigate to folder or open file.

        ' If folder exists, go there
        If Directory.Exists(FileOrFolder) Then

            NavigateTo(FileOrFolder, True)

            ' If file exists, open it
        ElseIf File.Exists(FileOrFolder) Then

            OpenFileWithDefaultApp(FileOrFolder)

        Else
            ShowStatus(StatusPad & IconWarning & " Path does not exist: " & FileOrFolder)
        End If

    End Sub

    Private Sub OpenSelectedOrStartCommand()

        ' If a file or folder is selected, open it
        If lvFiles.SelectedItems.Count > 0 Then

            Dim selected As ListViewItem = lvFiles.SelectedItems(0)
            Dim fullPath As String = selected.Tag.ToString()

            GoToFolderOrOpenFile(fullPath)
            Return

        End If

        ' Nothing selected → start an open command
        txtAddressBar.Focus()
        txtAddressBar.Text = "open "
        PlaceCaretAtEndOfAddressBar()
    End Sub

    Private Sub CreateDirectory(directoryPath As String)

        Try
            ' Create the directory
            Dim dirInfo As DirectoryInfo = Directory.CreateDirectory(directoryPath)

            ShowStatus(StatusPad & IconNewFolder & " Directory created: " & dirInfo.FullName)

            NavigateTo(directoryPath, True)

        Catch ex As Exception
            ShowStatus(StatusPad & IconError & " Failed to create directory: " & ex.Message)
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

                ShowStatus(StatusPad & IconSuccess & " Text file created: " & filePath)

            Else
                ShowStatus(StatusPad & IconDialog & " File already exists: " & filePath)
            End If

            ' Refresh the folder view so the user sees the new or existing file
            NavigateTo(destDir)

            ' Open the newly created or existing file
            GoToFolderOrOpenFile(filePath)

        Catch ex As Exception
            ShowStatus(StatusPad & IconError & " Failed to create text file: " & ex.Message)
            Debug.WriteLine("CreateTextFile Error: " & ex.Message)

        End Try

    End Sub

    Private Sub MoveFileOrDirectory(source As String, destination As String)
        Try
            ' Validate parameters
            If String.IsNullOrWhiteSpace(source) OrElse String.IsNullOrWhiteSpace(destination) Then
                ShowStatus(StatusPad & IconWarning & " Source or destination path is invalid.")
                Return
            End If

            ' if source and destination are the same, do nothing
            If String.Equals(source.TrimEnd("\"c), destination.TrimEnd("\"c), StringComparison.OrdinalIgnoreCase) Then
                ShowStatus(StatusPad & IconWarning & " Source and destination paths are the same. Move operation canceled.")
                Return
            End If

            ' Is source on the protected paths list?
            If IsProtectedPathOrFolder(source) Then
                ShowStatus(StatusPad & IconProtect & " Move operation prevented for protected path: " & source)
                Return
            End If

            ' Is destination on the protected paths list?
            If IsProtectedPathOrFolder(destination) Then
                ShowStatus(StatusPad & IconProtect & " Move operation prevented for protected path: " & destination)
                Return
            End If

            ' Prevent moving a directory into itself or its subdirectory
            If Directory.Exists(source) AndAlso
               (String.Equals(source.TrimEnd("\"c), destination.TrimEnd("\"c), StringComparison.OrdinalIgnoreCase) OrElse
                destination.StartsWith(source.TrimEnd("\"c) & "\", StringComparison.OrdinalIgnoreCase)) Then
                ShowStatus(StatusPad & IconWarning & " Cannot move a directory into itself or its subdirectory.")
                Return
            End If

            ' Check if the source is a file
            If File.Exists(source) Then

                ' Check if the destination file already exists
                If Not File.Exists(destination) Then

                    ' Navigate to the directory of the source file
                    NavigateTo(Path.GetDirectoryName(source))

                    ShowStatus(StatusPad & IconMoving & "  Moving file to: " & destination)

                    ' Ensure destination directory exists
                    Directory.CreateDirectory(Path.GetDirectoryName(destination))

                    File.Move(source, destination)

                    ' Navigate to the destination folder (corrected)
                    NavigateTo(destination)

                    ShowStatus(StatusPad & IconMoving & "  Moved file to: " & destination)

                Else
                    ShowStatus(StatusPad & IconWarning & " Destination file already exists.")
                End If

            ElseIf Directory.Exists(source) Then

                ' Check if the destination directory already exists
                If Not Directory.Exists(destination) Then

                    ' Navigate to the directory being moved so the user can see it
                    NavigateTo(source)

                    ShowStatus(StatusPad & IconMoving & "  Moving directory to: " & destination)

                    ' Ensure destination parent exists
                    Directory.CreateDirectory(Path.GetDirectoryName(destination))

                    ' Perform the move
                    Directory.Move(source, destination)

                    ' Navigate to the new location FIRST
                    NavigateTo(destination)

                    ' Now refresh the tree roots
                    UpdateTreeRoots()

                    ShowStatus(StatusPad & IconMoving & "  Moved directory to: " & destination)

                Else
                    ShowStatus(StatusPad & IconWarning & " Destination directory already exists.")
                End If

            Else
                ShowStatus(StatusPad & IconWarning & "  Move failed: Source path not found. Paths with spaces must be enclosed in quotes. Example: move ""C:\folder A"" ""C:\folder B""")
            End If


        Catch ex As Exception
            ShowStatus(StatusPad & IconError & " Move failed: " & ex.Message)
            Debug.WriteLine("MoveFileOrDirectory Error: " & ex.Message)
        End Try
    End Sub

    Private Async Sub DeleteFileOrDirectory(path2Delete As String)
        ' Reject relative paths outright
        If Not Path.IsPathRooted(path2Delete) Then
            ShowStatus(StatusPad & IconWarning & "  Delete failed: Path must be absolute. Example: C:\folder")
            Exit Sub
        End If

        ' Check if the path is in the protected list
        If IsProtectedPathOrFolder(path2Delete) Then
            ' Navigate to the protected path so the user can see it was not deleted
            NavigateTo(path2Delete)
            ' Inform the user that deletion was prevented
            ShowStatus(StatusPad & IconProtect & "  Deletion prevented for protected path: " & path2Delete)
            Exit Sub
        End If

        Try
            ' Check if it's a file
            If File.Exists(path2Delete) Then
                ' Go to the directory of the file to be deleted so the user can see what is about to be deleted.
                Dim destDir As String = IO.Path.GetDirectoryName(path2Delete)
                NavigateTo(destDir)

                ' Make the user confirm file deletion.
                Dim fileName As String = Path.GetFileName(path2Delete)
                Dim confirmMsg As String = "Are you sure you want to delete the file:" & Environment.NewLine &
                                        "''" & fileName & "''?"
                Dim result = MessageBox.Show(confirmMsg,
                                         "Confirm File Deletion",
                                         MessageBoxButtons.YesNo,
                                         MessageBoxIcon.Question)
                If result <> DialogResult.Yes Then Exit Sub

                File.Delete(path2Delete)

                ' Refresh the view to show the file is deleted
                Await PopulateFiles(destDir) ' Await the async method

                ShowStatus(StatusPad & IconDelete & "  Deleted file: " & fileName)

                ' Check if it's a directory
            ElseIf Directory.Exists(path2Delete) Then
                ' Navigate into the folder to be deleted
                NavigateTo(path2Delete)

                ' Get the parent directory so we can navigate there after deletion
                Dim parentDir As String = IO.Path.GetDirectoryName(path2Delete)

                ' Ask the user to confirm deletion
                Dim folderName As String = Path.GetFileName(path2Delete)
                Dim confirmMsg As String =
                "Are you sure you want to delete the following folder:" & Environment.NewLine &
                "''" & folderName & "'' and all of its contents?"
                Dim result = MessageBox.Show(confirmMsg,
                                         "Confirm Folder Deletion",
                                         MessageBoxButtons.YesNo,
                                         MessageBoxIcon.Question)
                If result <> DialogResult.Yes Then Exit Sub

                Directory.Delete(path2Delete, recursive:=True)

                ' Navigate to the parent so the user sees the result
                NavigateTo(parentDir, True)

                ' Refresh the view to show the folder is deleted
                Await PopulateFiles(parentDir) ' Await the async method

                ShowStatus(StatusPad & IconDelete & "  Deleted folder: " & folderName)

            Else
                ShowStatus(StatusPad & IconWarning & "  Delete failed: Path not found.")
            End If

        Catch ex As Exception
            ShowStatus(StatusPad & IconError & " Delete failed: " & ex.Message)
            Debug.WriteLine("DeleteFileOrDirectory Error: " & ex.Message)
        End Try
    End Sub

    Private Sub RenameFileOrDirectory(sourcePath As String, newName As String)

        Dim newPath As String = Path.Combine(Path.GetDirectoryName(sourcePath), newName)

        ' Rule 1: Path must be absolute (start with C:\ or similar).
        ' Reject relative paths outright
        If Not Path.IsPathRooted(sourcePath) Then
            ShowStatus(StatusPad & IconDialog & " Rename failed: Path must be absolute. Example: C:\folder")
            Exit Sub
        End If

        ' Rule 2: Protected paths are never renamed.
        ' Check if the path is in the protected list
        If IsProtectedPathOrFolder(sourcePath) Then
            ' The path is protected; prevent rename

            ' Show user the directory so they can see it wasn't renamed.
            NavigateTo(sourcePath)

            ' Notify the user of the prevention so the user knows why it didn't rename.
            ShowStatus(StatusPad & IconProtect & "  Rename prevented for protected path or folder: " & sourcePath)

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

                ShowStatus(StatusPad & IconSuccess & " Renamed Folder to: " & newName)

                ' Rule 4: If it’s a file, rename the file and show its folder.
                ' If source is a file
            ElseIf File.Exists(sourcePath) Then

                ' Rename file
                File.Move(sourcePath, newPath)

                ' Navigate to the directory of the renamed file
                NavigateTo(Path.GetDirectoryName(sourcePath))

                ShowStatus(StatusPad & IconSuccess & " Renamed File to: " & newName)

                ' Rule 5: If nothing exists at that path, explain the quoting rule for spaces.
            Else
                ' Path does not exist
                ShowStatus(StatusPad & IconError & " Renamed failed: No path. Paths with spaces must be enclosed in quotes. Example: rename ""[source_path]"" ""[new_name]"" e.g., rename ""C:\folder\old name.txt"" ""new name.txt""")

            End If

        Catch ex As Exception
            ' Rule 6: If anything goes wrong,show a status message.
            ShowStatus(StatusPad & IconError & " Rename failed: " & ex.Message)
            Debug.WriteLine("RenameFileOrDirectory Error: " & ex.Message)
        End Try

    End Sub


    Private Sub OnlySearchForFilesInCurrentFolder(searchTerm As String)

        ' Clear item selection for lvFiles
        lvFiles.SelectedItems.Clear()

        Try
            SearchResults = New List(Of String)

            ' Search files only in the current folder
            For Each filePath In Directory.GetFiles(currentFolder, "*.*")
                If Path.GetFileName(filePath).IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    SearchResults.Add(filePath)
                End If
            Next

            ' Highlight first match in ListView
            If SearchResults.Count > 0 Then
                SelectListViewItemByPath(SearchResults(0))
                lvFiles.Focus()

                ShowStatus(StatusPad & IconSmile & " Found " & SearchResults.Count & " item(s) matching: " & searchTerm)
            Else
                ShowStatus(StatusPad & IconDialog & " No items found matching: " & searchTerm)
            End If

        Catch ex As Exception
            ShowStatus(StatusPad & IconError & " Search failed: " & ex.Message)
            Debug.WriteLine("OnlySearchForFilesInCurrentFolder Error: " & ex.Message)
        End Try

    End Sub

    Private Sub UpdateNavButtons()
        btnBack.Enabled = _historyIndex > 0
        btnForward.Enabled = _historyIndex >= 0 AndAlso _historyIndex < _history.Count - 1
    End Sub

    Private Sub UpdateDestructiveButtonsAndMenus()

        ' --- Rule 0: Something must be selected ---
        Dim hasSelection As Boolean = (lvFiles.SelectedItems.Count > 0)
        If Not hasSelection Then
            SetDestructiveButtons(False, False, False)
            SetContextDestructiveItems(False)
            Exit Sub
        End If

        Dim fullPath As String = CStr(lvFiles.SelectedItems(0).Tag)

        ' --- Rule 1: Path must exist ---
        If Not PathExists(fullPath) Then
            SetDestructiveButtons(False, False, False)
            SetContextDestructiveItems(False)
            Exit Sub
        End If

        ' --- Rule 2: Protected items cannot be renamed or deleted ---
        If IsProtectedPathOrFolder(fullPath) Then
            SetDestructiveButtons(False, False, False)
            SetContextDestructiveItems(False)
            Exit Sub
        End If

        ' --- Rule 3: User must have rename permission ---
        Dim parentDir As String =
        If(Directory.Exists(fullPath),
           fullPath,                          ' Folder → check write access ON the folder
           Path.GetDirectoryName(fullPath))   ' File → check write access on parent folder

        Dim canRename As Boolean = HasWriteAccess(parentDir)

        ' --- Apply toolbar rules ---
        btnCut.Enabled = canRename
        btnRename.Enabled = canRename
        btnDelete.Enabled = True   ' Delete allowed if path exists + not protected

        ' --- Apply context menu rules ---
        SetContextDestructiveItems(True)

    End Sub

    Private Sub SetDestructiveButtons(cut As Boolean, rename As Boolean, delete As Boolean)
        btnCut.Enabled = cut
        btnRename.Enabled = rename
        btnDelete.Enabled = delete
    End Sub

    Private Sub SetContextDestructiveItems(enabled As Boolean)
        cmsFiles.Items("Cut").Enabled = enabled
        cmsFiles.Items("Rename").Enabled = enabled
        cmsFiles.Items("Delete").Enabled = enabled
    End Sub

    Private Sub UpdateCopyButtonAndMenus()

        Dim hasSelection As Boolean = (lvFiles.SelectedItems.Count > 0)
        Dim fullPath As String = If(hasSelection, CStr(lvFiles.SelectedItems(0).Tag), Nothing)
        Dim exists As Boolean = hasSelection AndAlso PathExists(fullPath)

        Dim canCopy As Boolean = exists

        btnCopy.Enabled = canCopy

        If cmsFiles.Items.Count > 0 Then
            cmsFiles.Items("Copy").Enabled = canCopy
        End If

    End Sub

    Private Sub UpdateEditButtonsAndMenus()

        UpdateCopyButtonAndMenus()

        UpdateDestructiveButtonsAndMenus()

        ' Update Paste button
        If _clipboardIsCut Then
            btnPaste.Enabled = True
        Else
            btnPaste.Enabled = Not String.IsNullOrEmpty(_clipboardPath)
        End If

        If cmsFiles.Items.Count = 0 Then Return

        If _clipboardIsCut Then
            cmsFiles.Items("Paste").Enabled = True
        Else
            cmsFiles.Items("Paste").Enabled = Not String.IsNullOrEmpty(_clipboardPath)
        End If

    End Sub

    Private Sub UpdateFileButtonsAndMenus()

        ' --- Folder-level permissions ---
        Dim canWrite As Boolean = HasWriteAccess(currentFolder)
        Dim canCreateFolders As Boolean = HasDirectoryCreationAccess(currentFolder)

        btnNewTextFile.Enabled = canWrite
        btnNewFolder.Enabled = canCreateFolders

        If cmsFiles.Items.Count = 0 Then Exit Sub

        cmsFiles.Items("NewFolder").Enabled = canCreateFolders
        cmsFiles.Items("NewTextFile").Enabled = canWrite


        ' --- Item-level actions ---
        Dim hasSelection As Boolean = (lvFiles.SelectedItems.Count > 0)

        If Not hasSelection Then
            cmsFiles.Items("Open").Enabled = False
            cmsFiles.Items("CopyPath").Enabled = False
            Exit Sub
        End If

        Dim selectedPath As String = CStr(lvFiles.SelectedItems(0).Tag)
        Dim exists As Boolean = PathExists(selectedPath)

        cmsFiles.Items("Open").Enabled = exists
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
        ' Check if lblStatus is not null
        If lblStatus IsNot Nothing Then
            ' Use the parent control's Invoke method
            If lblStatus.GetCurrentParent.InvokeRequired Then
                lblStatus.GetCurrentParent.Invoke(New Action(Of String)(AddressOf ShowStatus), message)
            Else
                ' Update the ToolStripStatusLabel text
                lblStatus.Text = message
                statusTimer.Stop()
                AddHandler statusTimer.Tick, AddressOf ClearStatus
                statusTimer.Start()
            End If
        Else
            ' Handle the case where lblStatus is not initialized
            Debug.WriteLine("lblStatus control is not initialized.")
        End If
    End Sub

    Private Sub ClearStatus(sender As Object, e As EventArgs)
        lblStatus.Text = ""
        statusTimer.Stop()
    End Sub


    Private Sub PlaceCaretAtEndOfAddressBar()
        ' Place caret at end of text
        txtAddressBar.SelectionStart = txtAddressBar.Text.Length
    End Sub

    'Private Sub HandleFindNextCommand()

    '    If SearchResults.Count = 0 Then
    '        ShowStatus(StatusPad & IconDialog & "  No previous search results. Press: Ctrl + F or enter: 'find [search_term]' to start a search.")
    '        Return
    '    End If

    '    ' Advance index
    '    SearchIndex += 1

    '    ' Wrap around
    '    If SearchIndex >= SearchResults.Count Then
    '        SearchIndex = 0
    '    End If

    '    ' Select the next result
    '    lvFiles.SelectedItems.Clear()
    '    Dim nextPath As String = SearchResults(SearchIndex)

    '    SelectListViewItemByPath(nextPath)
    '    Dim fileName As String = Path.GetFileNameWithoutExtension(nextPath)

    '    ShowStatus(
    '        StatusPad & IconSearch &
    '        "    Result " &
    '        (SearchIndex + 1) &
    '        " of " &
    '        SearchResults.Count &
    '        $"     ""{fileName}""       Next result  -  F3      Open  -  Ctrl + O      Reset  -  Esc"
    '    )

    'End Sub


    Private Sub HandleFindNextCommand()

        ' No active search
        If SearchResults.Count = 0 Then
            ShowStatus(
            StatusPad & IconDialog &
            "  No previous search results. Press Ctrl+F or enter: find [search_term] to start a search."
        )
            Return
        End If

        ' Advance index with wraparound
        SearchIndex += 1
        If SearchIndex >= SearchResults.Count Then
            SearchIndex = 0
        End If

        ' Select the next result
        lvFiles.SelectedItems.Clear()
        Dim nextPath As String = SearchResults(SearchIndex)
        SelectListViewItemByPath(nextPath)

        'HighlightSearchMatches()
        'HighlightCurrentResult()


        Dim fileName As String = Path.GetFileNameWithoutExtension(nextPath)

        ' Status HUD
        ShowStatus(
        StatusPad & IconSearch &
        $"  Result {SearchIndex + 1} of {SearchResults.Count}    " &
        $""“{fileName}”"    Next  F3    Open  Ctrl+O    Reset  Esc"
    )

    End Sub


    'Private Sub HighlightSearchMatches()
    '    ' Soft pastel highlight that feels calm and learner-friendly
    '    Dim highlightColor As Color = Color.FromArgb(235, 245, 255)  ' very light blue

    '    ' Reset all items first
    '    For Each item As ListViewItem In lvFiles.Items
    '        lvFiles.BeginUpdate()
    '        item.BackColor = Color.White
    '        lvFiles.EndUpdate()
    '    Next

    '    ' Apply highlight to matched items
    '    For Each path As String In SearchResults
    '        Dim item As ListViewItem = FindListViewItemByPath(path)
    '        If item IsNot Nothing Then
    '            lvFiles.BeginUpdate()
    '            item.BackColor = highlightColor
    '            lvFiles.EndUpdate()
    '        End If
    '    Next
    'End Sub

    Private Sub HighlightSearchMatches()
        ' Soft pastel highlight that feels calm and learner-friendly
        'Dim highlightColor As Color = Color.FromArgb(235, 245, 255)  ' very light blue
        Dim highlightColor As Color = Color.FromArgb(215, 240, 251)  ' very light blue


        lvFiles.BeginUpdate()

        ' Reset all items first
        For Each item As ListViewItem In lvFiles.Items
            item.BackColor = Color.White
        Next

        ' Apply highlight to matched items
        For Each path As String In SearchResults
            Dim item As ListViewItem = FindListViewItemByPath(path)
            If item IsNot Nothing Then
                item.BackColor = highlightColor
            End If
        Next

        lvFiles.EndUpdate()
    End Sub

    Private Function FindListViewItemByPath(fullPath As String) As ListViewItem
        For Each item As ListViewItem In lvFiles.Items
            If String.Equals(item.Tag?.ToString(), fullPath, StringComparison.OrdinalIgnoreCase) Then
                Return item
            End If
        Next
        Return Nothing
    End Function


    Private Sub InitiateSearch()
        txtAddressBar.Focus()
        txtAddressBar.Text = "find "
        PlaceCaretAtEndOfAddressBar()
    End Sub

    Private Sub ConsumeKey(e As KeyEventArgs)
        e.Handled = True
        e.SuppressKeyPress = True
    End Sub

    Private Function ResolveDestinationPathWithAutoRename(
        initialDestPath As String,
        isDirectory As Boolean
    ) As String
        ' Generates a non‑conflicting destination path using Windows‑style
        ' rename‑on‑copy semantics:
        '   File.txt → File - Copy.txt → File - Copy (2).txt → ...
        ' For directories, the extension is omitted.

        If String.IsNullOrWhiteSpace(initialDestPath) Then
            Return initialDestPath
        End If

        Dim dir As String = Path.GetDirectoryName(initialDestPath)
        Dim baseName As String = Path.GetFileNameWithoutExtension(initialDestPath)
        Dim ext As String = If(isDirectory, String.Empty, Path.GetExtension(initialDestPath))

        ' If the directory doesn't exist, we can't check collisions
        If String.IsNullOrWhiteSpace(dir) OrElse Not Directory.Exists(dir) Then
            Return initialDestPath
        End If

        Dim candidate As String = initialDestPath
        Dim index As Integer = 1

        While (Not isDirectory AndAlso File.Exists(candidate)) OrElse
          (isDirectory AndAlso Directory.Exists(candidate))

            If index = 1 Then
                candidate = Path.Combine(dir, $"{baseName} - Copy{ext}")
            Else
                candidate = Path.Combine(dir, $"{baseName} - Copy ({index}){ext}")
            End If

            index += 1
        End While

        Return candidate
    End Function

    Private Function HasEnoughSpace(sourceFile As FileInfo, destinationPath As String) As Boolean
        Try
            Dim root = Path.GetPathRoot(destinationPath)
            Dim drive As New DriveInfo(root)
            Return sourceFile.Length <= drive.AvailableFreeSpace
        Catch
            ' If free space cannot be determined, allow the copy attempt
            Return True
        End Try
    End Function

    Private Function IsFileLocked(file As FileInfo) As Boolean
        Try
            Using stream As FileStream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None)
                ' If we can open it with no sharing, it's not locked
            End Using
            Return False
        Catch ex As IOException
            ' Sharing violation or lock
            Return True
        Catch ex As UnauthorizedAccessException
            ' File may be a directory or protected
            Return True
        End Try
    End Function

    Private Sub EnsureHelpFileExists(helpFilePath As String)

        Try
            If File.Exists(helpFilePath) Then
                Exit Sub
            End If

            ShowStatus(StatusPad & IconDialog & "  Creating help file.")

            Dim helpText As String = BuildHelpText()

            Dim dir As String = Path.GetDirectoryName(helpFilePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            File.WriteAllText(helpFilePath, helpText)

            ShowStatus(StatusPad & IconSuccess & "  Help file created.")

        Catch ex As Exception
            ShowStatus(StatusPad & IconError & "  Unable to create help file: " & ex.Message)
            Debug.WriteLine("Unable to create help file: " & ex.Message)
        End Try

    End Sub

    Private Function BuildHelpText() As String
        Dim sb As New StringBuilder()

        ' ---------------------------------------------------------
        ' HEADER
        ' ---------------------------------------------------------
        sb.AppendLine("========================================")
        sb.AppendLine("            File Explorer CLI")
        sb.AppendLine("========================================")
        sb.AppendLine()
        sb.AppendLine("Available Commands:")
        sb.AppendLine()

        ' ---------------------------------------------------------
        ' NAVIGATION
        ' ---------------------------------------------------------
        sb.AppendLine("cd")
        sb.AppendLine("cd [directory]")
        sb.AppendLine("  Change directory to the specified path.")
        sb.AppendLine("  Examples:")
        sb.AppendLine("    cd C:\")
        sb.AppendLine("    cd ""C:\My Folder""")
        sb.AppendLine()

        ' ---------------------------------------------------------
        ' DIRECTORY CREATION
        ' ---------------------------------------------------------
        sb.AppendLine("mkdir, make")
        sb.AppendLine("mkdir [directory_path]")
        sb.AppendLine("  Create a new folder.")
        sb.AppendLine("  Examples:")
        sb.AppendLine("    mkdir C:\newfolder")
        sb.AppendLine("    make ""C:\My New Folder""")
        sb.AppendLine()

        ' ---------------------------------------------------------
        ' COPY
        ' ---------------------------------------------------------
        sb.AppendLine("copy")
        sb.AppendLine("copy [source] [destination]")
        sb.AppendLine("  Copy a file or folder to a destination folder.")
        sb.AppendLine("  Examples:")
        sb.AppendLine("    copy C:\folderA\file.doc C:\folderB")
        sb.AppendLine("    copy ""C:\folder A"" ""C:\folder B""")
        sb.AppendLine()

        ' ---------------------------------------------------------
        ' MOVE
        ' ---------------------------------------------------------
        sb.AppendLine("move")
        sb.AppendLine("move [source] [destination]")
        sb.AppendLine("  Move a file or folder to a new location.")
        sb.AppendLine("  Examples:")
        sb.AppendLine("    move C:\folderA\file.doc C:\folderB\file.doc")
        sb.AppendLine("    move ""C:\folder A\file.doc"" ""C:\folder B\renamed.doc""")
        sb.AppendLine("    move C:\folderA\folder C:\folderB\folder")
        sb.AppendLine("    move ""C:\folder A"" ""C:\folder B\New Name""")
        sb.AppendLine()

        ' ---------------------------------------------------------
        ' DELETE
        ' ---------------------------------------------------------
        sb.AppendLine("delete")
        sb.AppendLine("delete [file_or_directory]")
        sb.AppendLine("  Delete a file or folder.")
        sb.AppendLine("  Examples:")
        sb.AppendLine("    delete C:\file.txt")
        sb.AppendLine("    delete ""C:\My Folder""")
        sb.AppendLine()

        ' ---------------------------------------------------------
        ' RENAME
        ' ---------------------------------------------------------
        sb.AppendLine("rename")
        sb.AppendLine("rename [source_path] [new_name]")
        sb.AppendLine("  Rename a file or directory.")
        sb.AppendLine("  Paths with spaces must be enclosed in quotes.")
        sb.AppendLine("  Examples:")
        sb.AppendLine("    rename ""C:\folder\oldname.txt"" ""newname.txt""")
        sb.AppendLine("    rename ""C:\folder\old name.txt"" ""new name.txt""")
        sb.AppendLine()

        ' ---------------------------------------------------------
        ' OPEN
        ' ---------------------------------------------------------
        sb.AppendLine("open")
        sb.AppendLine("open [file_or_directory]")
        sb.AppendLine("  Open a file with the default application, or navigate into a folder.")
        sb.AppendLine("  Examples:")
        sb.AppendLine("    open C:\folder\file.txt")
        sb.AppendLine("    open ""C:\My Folder""")
        sb.AppendLine("    open")
        sb.AppendLine("      (opens the selected file or folder)")
        sb.AppendLine()

        ' ---------------------------------------------------------
        ' TEXT FILE CREATION
        ' ---------------------------------------------------------
        sb.AppendLine("text, txt")
        sb.AppendLine("text [file_path]")
        sb.AppendLine("  Create a new text file at the specified path.")
        sb.AppendLine("  Example:")
        sb.AppendLine("    text ""C:\folder\example.txt""")
        sb.AppendLine()

        ' ---------------------------------------------------------
        ' SEARCH
        ' ---------------------------------------------------------
        sb.AppendLine("find, search")
        sb.AppendLine("find [search_term]")
        sb.AppendLine("  Search for files and folders in the current directory.")
        sb.AppendLine("  Example:")
        sb.AppendLine("    find document")
        sb.AppendLine()

        sb.AppendLine("findnext, searchnext")
        sb.AppendLine("findnext")
        sb.AppendLine("  Show the next search result from the previous search.")
        sb.AppendLine()

        ' ---------------------------------------------------------
        ' EXIT
        ' ---------------------------------------------------------
        sb.AppendLine("exit, quit")
        sb.AppendLine("  Exit the application.")
        sb.AppendLine()

        ' ---------------------------------------------------------
        ' HELP
        ' ---------------------------------------------------------
        sb.AppendLine("help")
        sb.AppendLine("  Opens this help file.")
        sb.AppendLine()

        Return sb.ToString()
    End Function

    Private Function PathExists(path As String) As Boolean
        Return File.Exists(path) OrElse Directory.Exists(path)
    End Function

    Private Function HasWriteAccess(dirPath As String) As Boolean
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
            Debug.WriteLine("HasWriteAccessToDirectory - " & dirPath & " - False - UnauthorizedAccessException - " & ex.Message)

            Return False
        Catch ex As IOException
            Debug.WriteLine("HasWriteAccessToDirectory - " & dirPath & " - False - IOException - " & ex.Message)

            Return False
        End Try

    End Function

    Private Function HasDirectoryCreationAccess(dirPath As String) As Boolean
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
            Debug.WriteLine("HasDirectoryCreationAccessToDirectory - " & dirPath & " - False - UnauthorizedAccessException:  " & ex.Message)

            Return False

        Catch ex As IOException
            ' Could be permissions, could be other IO issues.
            ' For access testing, treat as failure.
            Debug.WriteLine("HasDirectoryCreationAccessToDirectory - " & dirPath & " - False - IOException:  " & ex.Message)

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
            Catch ex As Exception
                ' Swallow cleanup errors intentionally.
                Debug.WriteLine(dirPath & " Swallow cleanup errors intentionally " & ex.Message)

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

        Catch ex As UnauthorizedAccessException
            Debug.WriteLine($"HasSubdirectories - {path} - False - UnauthorizedAccessException - {ex.Message}")
            Return False

        Catch ex As Exception
            Debug.WriteLine($"HasSubdirectories - {path} - False - Error - {ex.Message}")
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

        If bytes = 0 Then Return "0 B"

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



    Private Sub InitApp()

        Me.Text = "File Explorer - Code with Joe"

        Me.KeyPreview = True

        Me.CenterToScreen()

        ConfigureTooltips()

        InitStatusBar()

        ShowStatus(StatusPad & IconDialog & " Loading...")

        InitContextMenu()

        InitImageList()

        InitArrows()

        InitListView()

        InitTreeView()

        '  Start in User Profile folder
        NavigateTo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))

        UpdateTreeRoots()


        RunTests()

        ShowStatus(StatusPad & IconSuccess & "  Ready")

    End Sub

    Private Sub ConfigureTooltips()

        ' General tooltip settings
        tips.AutoPopDelay = 6000
        tips.InitialDelay = 300
        tips.ReshowDelay = 100
        tips.ShowAlways = True

        ' ============================
        ' Navigation Buttons
        ' ============================
        tips.SetToolTip(btnBack, "Go back to the previous folder  (Alt + ← or Backspace)")
        tips.SetToolTip(btnForward, "Go forward to the next folder  (Alt + →)")
        tips.SetToolTip(btnRefresh, "Refresh the current folder  (F5)")
        tips.SetToolTip(bntHome, "Go to your Home directory (Alt + Home)")
        tips.SetToolTip(btnGo, "Go to path or run command (Enter)")

        ' ============================
        ' File / Folder Creation
        ' ============================
        tips.SetToolTip(btnNewFolder, "Create a new folder  (Ctrl + Shift + N)")
        tips.SetToolTip(btnNewTextFile, "Create a new text file (Ctrl + Shift + T)")

        ' ============================
        ' Clipboard Operations
        ' ============================
        tips.SetToolTip(btnCut, "Cut selected items  (Ctrl + X)")
        tips.SetToolTip(btnCopy, "Copy selected items  (Ctrl + C)")
        tips.SetToolTip(btnPaste, "Paste items  (Ctrl + V)")

        ' ============================
        ' File Operations
        ' ============================
        tips.SetToolTip(btnRename, "Rename the selected item  (F2)")
        tips.SetToolTip(btnDelete, "Delete the selected item  (Delete or Ctrl + D)")

        'tips.SetToolTip(txtAddressBar, "Address Bar: Type a path or command here.  (Ctrl + L, Alt + D, or F4 to focus)")
        tips.SetToolTip(txtAddressBar,
                        "Type a path or command. Enter runs it, Esc resets. Ctrl+L, Alt+D, or F4 to focus.")

    End Sub

    Private Sub InitListView()
        lvFiles.View = View.Details
        lvFiles.FullRowSelect = True
        lvFiles.MultiSelect = True
        lvFiles.Columns.Clear()
        lvFiles.Columns.Add("Name", 500)
        lvFiles.Columns.Add("Type", 175)
        lvFiles.Columns.Add("Size", 100)
        lvFiles.Columns.Add("Modified", 150)

    End Sub

    Private Sub InitTreeView()

        ' Don't show lines connecting nodes
        tvFolders.ShowRootLines = False

        ' Use arrow icons instead of plus/minus
        tvFolders.ShowPlusMinus = False
        tvFolders.StateImageList = imgArrows   ' Index 0 = ▶, Index 1 = ▼, Index 2 = no arrow

    End Sub




    'Private Sub EnsureEasyAccessSettings()
    '    If My.Settings.EasyAccessPaths Is Nothing Then
    '        My.Settings.EasyAccessPaths = New Specialized.StringCollection()
    '        My.Settings.Save()
    '    End If
    'End Sub




    Private Sub UpdateTreeRoots()
        ' Update the TreeView with current drives and special folders at the top level.

        tvFolders.BeginUpdate()

        ' Clear existing nodes and add new ones for:
        tvFolders.Nodes.Clear()

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

        ' --- User Easy Access entries ---
        Dim userEntries = LoadEasyAccessEntries()

        For Each entry In userEntries
            Dim node As New TreeNode(entry.Name) With {
        .Tag = entry.Path,
        .ImageKey = "Folder",
        .SelectedImageKey = "Folder"
    }

            If HasSubdirectories(entry.Path) Then
                node.Nodes.Add("Loading...")
                node.StateImageIndex = 0
            Else
                node.StateImageIndex = 2
            End If

            easyAccessNode.Nodes.Add(node)
        Next


        ' Add Easy Access to tree
        tvFolders.Nodes.Add(easyAccessNode)

        ' Expand Easy Access and update arrow
        easyAccessNode.Expand()
        easyAccessNode.StateImageIndex = 1   ' ▼ expanded

        '' --- Drives ---
        'For Each di In DriveInfo.GetDrives()
        '    If di.IsReady Then
        '        Try
        '            Dim rootNode As New TreeNode(di.Name & " - " & di.VolumeLabel) With {
        '                .Tag = di.RootDirectory.FullName
        '            }

        '            If di.DriveType = DriveType.CDRom Then
        '                rootNode.ImageKey = "Optical"
        '                rootNode.SelectedImageKey = "Optical"
        '            Else
        '                rootNode.ImageKey = "Drive"
        '                rootNode.SelectedImageKey = "Drive"
        '            End If

        '            If HasSubdirectories(di.RootDirectory.FullName) Then
        '                rootNode.Nodes.Add("Loading...")
        '                rootNode.StateImageIndex = 0   ' ▶ collapsed
        '            Else
        '                rootNode.StateImageIndex = 2  ' no arrow
        '            End If

        '            tvFolders.Nodes.Add(rootNode)

        '        Catch ex As IOException
        '            Debug.WriteLine($"Error accessing drive {di.Name}: {ex.Message}")
        '        Catch ex As UnauthorizedAccessException
        '            Debug.WriteLine($"Access denied to drive {di.Name}: {ex.Message}")

        '        Catch ex As Exception
        '            Debug.WriteLine($"Unexpected error with drive {di.Name}: {ex.Message}")

        '        End Try
        '    Else
        '        Debug.WriteLine($"Drive {di.Name} is not ready.")
        '    End If
        'Next


        ' --- Drives ---
        For Each di In DriveInfo.GetDrives()
            If di.IsReady Then
                Try
                    Dim freeSpace As String = FormatBytes(di.AvailableFreeSpace)
                    Dim totalSpace As String = FormatBytes(di.TotalSize)

                    Dim displayText As String =
                $"{di.Name} - {di.VolumeLabel} ({freeSpace} free of {totalSpace})"

                    Dim rootNode As New TreeNode(displayText) With {
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

        tvFolders.EndUpdate()

    End Sub



    Private Function FormatBytes(bytes As Long) As String
        Dim sizes() As String = {"B", "KB", "MB", "GB", "TB"}
        Dim order As Integer = 0
        Dim value As Double = bytes

        While value >= 1024 AndAlso order < sizes.Length - 1
            order += 1
            value /= 1024
        End While

        Return $"{value:0.##} {sizes(order)}"
    End Function

    Private Sub InitStatusBar()

        Dim statusStrip As New StatusStrip()

        ' Set font to Segoe UI Symbol, 9pt
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
        imgList.Images.Add("Shortcut", My.Resources.Resource1.Shortcut_16X16)


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


    'Private Sub InitContextMenu()

    '    cmsFiles.Items.Add(New ToolStripMenuItem("Open", Nothing, AddressOf Open_Click) With {
    '    .Name = "Open",
    '    .ShortcutKeyDisplayString = "Ctrl+O"
    '})

    '    cmsFiles.Items.Add(New ToolStripMenuItem("New Folder", Nothing, AddressOf NewFolder_Click) With {
    '    .Name = "NewFolder",
    '    .ShortcutKeyDisplayString = "Ctrl+Shift+N"
    '})

    '    cmsFiles.Items.Add(New ToolStripMenuItem("New Text File", Nothing, AddressOf NewTextFile_Click) With {
    '    .Name = "NewTextFile",
    '    .ShortcutKeyDisplayString = "Ctrl+Shift+T"
    '})

    '    cmsFiles.Items.Add(New ToolStripMenuItem("Cut", Nothing, AddressOf CutSelected_Click) With {
    '    .Name = "Cut",
    '    .ShortcutKeyDisplayString = "Ctrl+X"
    '})

    '    cmsFiles.Items.Add(New ToolStripMenuItem("Copy", Nothing, AddressOf CopySelected_Click) With {
    '    .Name = "Copy",
    '    .ShortcutKeyDisplayString = "Ctrl+C"
    '})

    '    cmsFiles.Items.Add(New ToolStripMenuItem("Paste", Nothing, AddressOf PasteSelected_Click) With {
    '    .Name = "Paste",
    '    .ShortcutKeyDisplayString = "Ctrl+V"
    '})

    '    cmsFiles.Items.Add(New ToolStripMenuItem("Rename", Nothing, AddressOf RenameFile_Click) With {
    '    .Name = "Rename",
    '    .ShortcutKeyDisplayString = "F2"
    '})

    '    cmsFiles.Items.Add(New ToolStripMenuItem("Delete", Nothing, AddressOf Delete_Click) With {
    '    .Name = "Delete",
    '    .ShortcutKeyDisplayString = "Delete"
    '})

    '    cmsFiles.Items.Add(New ToolStripMenuItem("Copy Path", Nothing, AddressOf CopyFilePath_Click) With {
    '    .Name = "CopyPath",
    '    .ShortcutKeyDisplayString = "Ctrl+P"
    '})

    '    cmsFiles.Items.Add(New ToolStripMenuItem("Full-Screen", Nothing, AddressOf ToggleFullScreen) With {
    '    .Name = "FullScreen",
    '    .ShortcutKeyDisplayString = "F11"
    '})

    '    lvFiles.ContextMenuStrip = cmsFiles
    'End Sub

    Private Sub InitContextMenu()

        cmsFiles.Items.Add(New ToolStripMenuItem("Open", Nothing, AddressOf Open_Click) With {
            .Name = "Open",
            .ShortcutKeyDisplayString = "Ctrl+O"
        })

        cmsFiles.Items.Add(New ToolStripMenuItem("Home Folder", Nothing,
            Sub() GoToFolderOrOpenFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))
        ) With {
            .Name = "HomeFolder",
            .ShortcutKeyDisplayString = "Alt+Home"
        })

        cmsFiles.Items.Add(New ToolStripMenuItem("New Folder", Nothing, AddressOf NewFolder_Click) With {
            .Name = "NewFolder",
            .ShortcutKeyDisplayString = "Ctrl+Shift+N"
        })

        cmsFiles.Items.Add(New ToolStripMenuItem("New Text File", Nothing, AddressOf NewTextFile_Click) With {
            .Name = "NewTextFile",
            .ShortcutKeyDisplayString = "Ctrl+Shift+T"
        })

        cmsFiles.Items.Add(New ToolStripMenuItem("Cut", Nothing, AddressOf CutSelected_Click) With {
            .Name = "Cut",
            .ShortcutKeyDisplayString = "Ctrl+X"
        })

        cmsFiles.Items.Add(New ToolStripMenuItem("Copy", Nothing, AddressOf CopySelected_Click) With {
            .Name = "Copy",
            .ShortcutKeyDisplayString = "Ctrl+C"
        })

        cmsFiles.Items.Add(New ToolStripMenuItem("Paste", Nothing, AddressOf PasteSelected_Click) With {
            .Name = "Paste",
            .ShortcutKeyDisplayString = "Ctrl+V"
        })

        cmsFiles.Items.Add(New ToolStripMenuItem("Rename", Nothing, AddressOf RenameFile_Click) With {
            .Name = "Rename",
            .ShortcutKeyDisplayString = "F2"
        })

        cmsFiles.Items.Add(New ToolStripMenuItem("Delete", Nothing, AddressOf Delete_Click) With {
            .Name = "Delete",
            .ShortcutKeyDisplayString = "Delete"
        })

        cmsFiles.Items.Add(New ToolStripMenuItem("Copy Path", Nothing, AddressOf CopyFilePath_Click) With {
            .Name = "CopyPath",
            .ShortcutKeyDisplayString = "Ctrl+P"
        })

        cmsFiles.Items.Add(New ToolStripMenuItem("Full-Screen", Nothing, AddressOf ToggleFullScreen) With {
            .Name = "FullScreen",
            .ShortcutKeyDisplayString = "F11"
        })

        lvFiles.ContextMenuStrip = cmsFiles

    End Sub

    Private Sub RunTests()

        Debug.WriteLine("Running tests...")

        TestIsProtectedPath()

        TestParseSize()

        TestFormatSize()

        Test_NormalizeTextFilePath()

        Test_GetUniqueFilePath()

        Test_RoundTrip()

        Debug.WriteLine("All tests executed.")

    End Sub

    Private Sub Test_RoundTrip()
        Debug.WriteLine("→ Testing FormatSize ↔ ParseSize round‑trip")

        Dim cmp As New ListViewItemComparer(0, SortOrder.Ascending)

        For Each u In SizeUnits
            Dim unit = u.Unit
            Dim factor = u.Factor

            ' Pick a representative value for this unit
            Dim original As Long = factor * 3

            Dim formatted As String = FormatSize(original)
            Dim parsed As Long = cmp.Test_ParseSize(formatted)

            AssertTrue(parsed = original, $"Round‑trip failed for {unit}: '{formatted}' parsed as {parsed}")
        Next

        Debug.WriteLine("✓ Round‑trip tests passed")
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
        AssertTrue(FormatSize(1152921504606846976L) = "1 EB", "1 EB should format as '1 EB'")
        AssertTrue(FormatSize(1152921504606846976L * 2) = "2 EB", "2 EB should format a '2 EB'")

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

        AssertTrue(cmp.Test_ParseSize("1 EB") = 1152921504606846976L, "1 EB should parse to 1,152,921,504,606,846,976 bytes")

        AssertTrue(cmp.Test_ParseSize("1.5 EB") = CLng(1.5 * 1152921504606846976L), "1.5 EB should parse to 1,729,382,256,910,270,464 bytes")

        AssertTrue(cmp.Test_ParseSize("2 EB") = 1152921504606846976L * 2, "2 EB should parse to 2,305,843,009,213,693,952 bytes")

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

    'Private Sub btnPin_Click(sender As Object, e As EventArgs) Handles btnPin.Click

    '    AddToEasyAccess(currentFolder)




    'End Sub


    Private Sub btnPin_Click(sender As Object, e As EventArgs) Handles btnPin.Click
        Dim name As String = GetFolderDisplayName(currentFolder)
        AddToEasyAccess(name, currentFolder)
    End Sub

    Private Function GetFolderDisplayName(folderPath As String) As String
        Dim name = Path.GetFileName(folderPath.TrimEnd("\"c))
        If String.IsNullOrWhiteSpace(name) Then
            name = folderPath
        End If
        Return name
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

        ' Look up the factor from the shared SizeUnits array
        Dim factor As Long = 1
        For Each u In Form1.SizeUnits
            If u.Unit.Equals(unitPart, StringComparison.OrdinalIgnoreCase) Then
                factor = u.Factor
                Exit For
            End If
        Next

        ' Compute final byte count
        Dim bytes As Double = value * factor
        Dim result As Long = CLng(Math.Round(bytes))

        Return If(isNegative, -result, result)
    End Function

    ' TEMP: expose parser for testing
    Public Function Test_ParseSize(input As String) As Long
        Return ParseSize(input)
    End Function

End Class



' ------------------------------------------------------------
' Unified result object for file and directory copy operations
' ------------------------------------------------------------
Public Class CopyResult
    Public Property FilesCopied As Integer = 0
    Public Property FilesSkipped As Integer = 0

    Public Property FilesRenamed As Integer = 0

    Public Property DirectoriesCreated As Integer
    Public Property Errors As New List(Of String)

    Public ReadOnly Property Success As Boolean
        Get
            Return Errors.Count = 0
        End Get
    End Property
End Class

' This app was developed with the help of Copilot through many human + AI pairing sessions.
' The goal: Explorer‑grade behavior with learner‑friendly clarity.

' Maximum Effort


