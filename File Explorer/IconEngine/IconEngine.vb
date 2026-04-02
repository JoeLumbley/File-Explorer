Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Threading
Imports File_Explorer.Explorer.Interop.Shell


Public Class IconEngine


    Private Const ThisPCGUID As String = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    Private Const ThisPCString As String = "This PC"
    Private Const ThisPCKey As String = "ThisPC"
    Private Const ThisPCPath As String = "shell:MyComputerFolder"
    Private Const FallbackThisPCKey As String = "FallbackThisPC"

    Private Const RecycleBinGUID As String = "shell:::{645FF040-5081-101B-9F08-00AA002F954E}"
    Private Const RecycleBinString As String = "Recycle Bin"
    Private Const RecycleBinKey As String = "RecycleBin"
    Private Const RecycleBinPath As String = "shell:RecycleBinFolder"
    Private Const FallbackRecycleBinKey As String = "Recycle"

    Private Const OpticalKey As String = "Optical"
    Private Const DriveKey As String = "Drive"
    Private Const FolderKey As String = "Folder"
    Private Const FallbackFolderKey As String = "FallbackFolder"

    Private Const EasyAccessString As String = "Easy Access"
    Private Const EasyAccessKey As String = "EasyAccess"


    Private Const FileKey As String = "File"
    Private Const FallbackFileKey As String = "FallbackFile"

    Private Const ExecutableKey As String = "FallbackExecutable"


    Private IconCache As New ImageList()
    Private UIForm As Form
    Private IconSize As Integer

    Public Sub New(form As Form)
        ' Initialize the ImageList with appropriate settings for caching icons.

        UIForm = form

        IconSize = GetScaledIconSize(UIForm)

        IconCache.ColorDepth = ColorDepth.Depth32Bit

        IconCache.ImageSize = New Size(IconSize, IconSize)

    End Sub

    Public Sub SetSpecialFolderNodeIcon(
    specialFolderNode As TreeNode,
    iconCache As ImageList,
    iconSize As Integer,
    specialFolderPath As String,
    FallbackKey As String)

        ' Assigns the correct icon for a special folder (e.g., Desktop, Documents, Downloads).
        ' Icons are cached by their folder path to avoid repeated ShellInterop calls and
        ' to keep TreeView population fast and responsive.

        ' If the icon for this special folder is not already cached, retrieve and store it.
        If Not iconCache.Images.ContainsKey(specialFolderPath) Then

            ' Attempt to retrieve the Explorer-accurate icon for this folder path.
            ' This supports DPI scaling and returns the correct Windows shell icon.
            Dim specialFolderIcon = ShellInterop.GetIconForPath(specialFolderPath, iconSize)
            'Dim specialFolderIcon As Icon = Nothing ' Uncomment to test fallback behavior.

            ' If retrieval succeeded, cache the icon and assign it to the node.
            If specialFolderIcon IsNot Nothing Then

                ' Store the icon in the ImageList using the folder path as the unique key.
                iconCache.Images.Add(specialFolderPath, specialFolderIcon.ToBitmap())

                ' Assign the cached icon to the TreeNode.
                specialFolderNode.ImageKey = specialFolderPath
                specialFolderNode.SelectedImageKey = specialFolderPath

            Else
                ' Retrieval failed—use the provided fallback icon to maintain UI consistency.
                specialFolderNode.ImageKey = FallbackKey
                specialFolderNode.SelectedImageKey = FallbackKey
            End If

        Else
            ' The icon is already cached—simply assign it to the node.
            specialFolderNode.ImageKey = specialFolderPath
            specialFolderNode.SelectedImageKey = specialFolderPath
        End If

    End Sub

    Public Sub SetEasyAccessUserEntryNodeIcon(userEntryNode As TreeNode, imageList As ImageList, iconSize As Integer)

        If Not imageList.Images.ContainsKey(FolderKey) Then
            Dim userEntryIcon = IconLibrary.GenericFolder(iconSize)
            'Dim userEasyAccessIcon = Nothing ' Uncomment to test fallback behavior.

            If userEntryIcon IsNot Nothing Then
                imageList.Images.Add(FolderKey, userEntryIcon.ToBitmap())
                userEntryNode.ImageKey = FolderKey
                userEntryNode.SelectedImageKey = FolderKey
            Else
                userEntryNode.ImageKey = FallbackFolderKey
                userEntryNode.SelectedImageKey = FallbackFolderKey
            End If
        Else
            userEntryNode.ImageKey = FolderKey
            userEntryNode.SelectedImageKey = FolderKey
        End If

    End Sub

    Public Sub SetThisPCNodeIcon(thisPCNode As TreeNode, iconCache As ImageList, iconSize As Integer)

        If Not iconCache.Images.ContainsKey(ThisPCKey) Then

            Dim thisPCIcon = ShellInterop.GetIconForVirtualFolder(ThisPCGUID, iconSize)
            'Dim thisPCIcon As Icon = Nothing ' Uncomment to test fallback behavior.

            If thisPCIcon IsNot Nothing Then
                iconCache.Images.Add(ThisPCKey, thisPCIcon.ToBitmap())
                thisPCNode.ImageKey = ThisPCKey
                thisPCNode.SelectedImageKey = ThisPCKey
            Else
                thisPCNode.ImageKey = FallbackThisPCKey
                thisPCNode.SelectedImageKey = FallbackThisPCKey
            End If
        Else
            thisPCNode.ImageKey = ThisPCKey
            thisPCNode.SelectedImageKey = ThisPCKey
        End If

    End Sub


    Public Sub SetDriveNodeIcon(driveNode As TreeNode, di As DriveInfo, imageList As ImageList, iconSize As Integer)

        If Not imageList.Images.ContainsKey(di.RootDirectory.FullName) Then

            Dim driveIcon = ShellInterop.GetIconForPath(di.RootDirectory.FullName, iconSize)
            'Dim driveIcon As Icon = Nothing ' Uncomment to test fallback behavior.

            If driveIcon IsNot Nothing Then
                imageList.Images.Add(di.RootDirectory.FullName, driveIcon.ToBitmap())
                driveNode.ImageKey = di.RootDirectory.FullName
                driveNode.SelectedImageKey = di.RootDirectory.FullName
            Else
                If di.DriveType = DriveType.CDRom Then
                    driveNode.ImageKey = OpticalKey
                    driveNode.SelectedImageKey = OpticalKey
                Else
                    driveNode.ImageKey = DriveKey
                    driveNode.SelectedImageKey = DriveKey
                End If
            End If
        Else
            driveNode.ImageKey = di.RootDirectory.FullName
            driveNode.SelectedImageKey = di.RootDirectory.FullName
        End If

    End Sub

    Public Sub SetRecycleNodeIcon(recycleBinNode As TreeNode, iconCache As ImageList, iconSize As Integer)
        ' Assigns the correct Recycle Bin icon to the provided TreeNode.
        ' Icons are cached in the ImageList to avoid redundant extraction and improve performance.

        ' If the ImageList does not already contain the Recycle Bin icon, we need to retrieve and cache it.
        If Not iconCache.Images.ContainsKey(RecycleBinKey) Then

            ' Attempt to retrieve the Recycle Bin icon from ShellInterop.
            ' This handles virtual folders and returns an Explorer-accurate icon when available.
            Dim recycleBinIcon = ShellInterop.GetIconForVirtualFolder(RecycleBinGUID, iconSize)
            'Dim recycleBinIcon = Nothing ' Uncomment to test fallback behavior.

            ' If ShellInterop successfully returned an icon, cache it and assign it to the node.
            If recycleBinIcon IsNot Nothing Then

                ' Store the icon in the ImageList under the canonical key.
                iconCache.Images.Add(RecycleBinKey, recycleBinIcon.ToBitmap())

                ' Point the TreeNode to the cached icon so future loads do not require ShellInterop again.
                recycleBinNode.ImageKey = RecycleBinKey
                recycleBinNode.SelectedImageKey = RecycleBinKey

            Else
                ' ShellInterop failed to retrieve the icon (rare but possible).
                ' Use the predefined fallback icon to keep the UI consistent and predictable.
                recycleBinNode.ImageKey = FallbackRecycleBinKey
                recycleBinNode.SelectedImageKey = FallbackRecycleBinKey
            End If

        Else
            ' The icon is already cached—simply assign it to the node.
            ' This avoids unnecessary interop calls and keeps UI updates fast.
            recycleBinNode.ImageKey = RecycleBinKey
            recycleBinNode.SelectedImageKey = RecycleBinKey
        End If

    End Sub

    Public Sub SetListViewItemIconForFile(item As ListViewItem, iconCache As ImageList, fi As FileInfo)

        Dim ext = fi.Extension.ToLowerInvariant()
        Dim IconSize As Integer = GetScaledIconSize(UIForm)


        If ext = "" Then
            If Not iconCache.Images.ContainsKey(FileKey) Then
                Dim fileIcon = IconLibrary.GenericFile(IconSize)
                If fileIcon IsNot Nothing Then
                    iconCache.Images.Add(FileKey, fileIcon.ToBitmap())
                    ' use generic file icon for no-extension files
                    item.ImageKey = FileKey
                Else
                    item.ImageKey = FallbackFileKey
                End If
            Else
                item.ImageKey = FileKey
            End If
        ElseIf Not ext = ".exe" Then
            If Not iconCache.Images.ContainsKey(ext) Then
                Dim extensionIcon = ShellInterop.GetIconForPath(fi.FullName, IconSize)
                If extensionIcon IsNot Nothing Then
                    iconCache.Images.Add(ext, extensionIcon.ToBitmap())
                    ' use extension as key to group same-type files
                    item.ImageKey = ext
                Else
                    item.ImageKey = FileKey
                End If
            Else
                item.ImageKey = ext
            End If
        Else
            If Not iconCache.Images.ContainsKey(fi.FullName) Then
                Dim executableIcon = ShellInterop.GetIconForPath(fi.FullName, IconSize)
                If executableIcon IsNot Nothing Then
                    iconCache.Images.Add(fi.FullName, executableIcon.ToBitmap())
                    ' use full path as key to preserve unique icons for different executables
                    item.ImageKey = fi.FullName
                Else
                    item.ImageKey = ExecutableKey
                End If
            Else
                item.ImageKey = fi.FullName
            End If
        End If


    End Sub

    Public Sub SetListViewItemIconForDirectory(item As ListViewItem, iconCache As ImageList, di As DirectoryInfo)
        ' Set icon, with caching in the image list to avoid duplicates and improve performance
        Dim folderIcon As Icon
        Dim IconSize As Integer = GetScaledIconSize(UIForm)

        If Not iconCache.Images.ContainsKey(di.FullName) Then

            If Not iconCache.Images.ContainsKey(FolderKey) Then

                folderIcon = IconLibrary.GenericFolder(IconSize)

                If folderIcon IsNot Nothing Then
                    iconCache.Images.Add(FolderKey, folderIcon.ToBitmap())
                    item.ImageKey = FolderKey
                Else
                    item.ImageKey = FallbackFolderKey
                End If

            Else
                item.ImageKey = FolderKey
            End If

        Else
            item.ImageKey = di.FullName
        End If

    End Sub

    Public Sub UpdateIconSize()
        ' Updates the icon size based on the current DPI scaling of the form.
        ' This should be called when the form's DPI changes to ensure icons are rendered at the correct size.
        IconSize = GetScaledIconSize(UIForm)
        IconCache.ImageSize = New Size(IconSize, IconSize)
        ClearCache() ' Clear cache to ensure icons are reloaded at the new size.
    End Sub

    Private Sub ClearCache()
        ' Clears the icon cache. This can be used to free memory or to refresh icons if needed.
        IconCache.Images.Clear()
    End Sub

    Public Function GetIconForPath(folderPath As String) As Bitmap
        ' Retrieves an icon for a given folder path, using the cache if available.
        ' This is a convenience method that combines caching logic with ShellInterop retrieval.
        If Not IconCache.Images.ContainsKey(folderPath) Then
            Dim icon = ShellInterop.GetIconForPath(folderPath, IconSize)
            If icon IsNot Nothing Then
                IconCache.Images.Add(folderPath, icon.ToBitmap())
            End If
        End If
        Return If(IconCache.Images.ContainsKey(folderPath), IconCache.Images(folderPath), Nothing)
    End Function

    Public Function GetIconForVirtualFolder(folderGUID As String) As Bitmap
        ' Retrieves an icon for a given virtual folder GUID, using the cache if available.
        ' This is a convenience method that combines caching logic with ShellInterop retrieval for virtual folders.
        If Not IconCache.Images.ContainsKey(folderGUID) Then
            Dim icon = ShellInterop.GetIconForVirtualFolder(folderGUID, IconSize)
            If icon IsNot Nothing Then
                IconCache.Images.Add(folderGUID, icon.ToBitmap())
            End If
        End If
        Return If(IconCache.Images.ContainsKey(folderGUID), IconCache.Images(folderGUID), Nothing)
    End Function

    Public Sub AddFallbackIcon(key As String, icon As Bitmap)
        ' Allows adding a fallback icon to the cache under a specific key.
        ' This can be used to ensure that there is always an icon available for certain keys, even if ShellInterop retrieval fails.
        If Not IconCache.Images.ContainsKey(key) Then
            IconCache.Images.Add(key, icon)
        End If
    End Sub

    Public Function GetIconCache() As ImageList
        ' Provides access to the shared ImageList cache for icons.
        ' This allows other parts of the application to retrieve cached icons by key.
        Return IconCache
    End Function

    Public Function GetScaledIconSize(form As Form) As Integer
        Dim scale As Double = form.DeviceDpi / 96.0
        Return CInt(16 * scale)
    End Function


End Class