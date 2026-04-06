'Imports System.Diagnostics
'Imports System.Drawing
'Imports System.IO
'Imports System.Threading
'Imports File_Explorer.Explorer.Interop.Shell


'Public Class IconEngine


'    Private Const ThisPCGUID As String = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
'    Private Const ThisPCString As String = "This PC"
'    Private Const ThisPCKey As String = "ThisPC"
'    Private Const ThisPCPath As String = "shell:MyComputerFolder"
'    Private Const FallbackThisPCKey As String = "FallbackThisPC"

'    Private Const RecycleBinGUID As String = "shell:::{645FF040-5081-101B-9F08-00AA002F954E}"
'    Private Const RecycleBinString As String = "Recycle Bin"
'    Private Const RecycleBinKey As String = "RecycleBin"
'    Private Const RecycleBinPath As String = "shell:RecycleBinFolder"
'    Private Const FallbackRecycleBinKey As String = "Recycle"

'    Private Const OpticalKey As String = "Optical"
'    Private Const DriveKey As String = "Drive"
'    Private Const FolderKey As String = "Folder"
'    Private Const FallbackFolderKey As String = "FallbackFolder"

'    Private Const EasyAccessString As String = "Easy Access"
'    Private Const EasyAccessKey As String = "EasyAccess"


'    Private Const FileKey As String = "File"
'    Private Const FallbackFileKey As String = "FallbackFile"

'    Private Const ExecutableKey As String = "FallbackExecutable"


'    Private IconCache As New ImageList()
'    Private UIForm As Form
'    Private iconSizePix As Integer

'    Public Sub New(form As Form)
'        ' Initialize the ImageList with appropriate settings for caching icons.

'        UIForm = form

'        iconSizePix = GetScaledIconSize(UIForm)

'        IconCache.ColorDepth = ColorDepth.Depth32Bit

'        IconCache.ImageSize = New Size(iconSizePix, iconSizePix)

'    End Sub

'    Public Sub SetSpecialFolderNodeIcon(
'    specialFolderNode As TreeNode,
'    specialFolderPath As String,
'    FallbackKey As String)

'        ' Assigns the correct icon for a special folder (e.g., Desktop, Documents, Downloads).
'        ' Icons are cached by their folder path to avoid repeated ShellInterop calls and
'        ' to keep TreeView population fast and responsive.

'        ' If the icon for this special folder is not already cached, retrieve and store it.
'        If Not IconCache.Images.ContainsKey(specialFolderPath) Then

'            ' Attempt to retrieve the Explorer-accurate icon for this folder path.
'            ' This supports DPI scaling and returns the correct Windows shell icon.
'            Dim specialFolderIcon = ShellInterop.GetIconForPath(specialFolderPath, iconSizePix)
'            'Dim specialFolderIcon As Icon = Nothing ' Uncomment to test fallback behavior.

'            ' If retrieval succeeded, cache the icon and assign it to the node.
'            If specialFolderIcon IsNot Nothing Then

'                ' Store the icon in the ImageList using the folder path as the unique key.
'                IconCache.Images.Add(specialFolderPath, specialFolderIcon.ToBitmap())

'                ' Assign the cached icon to the TreeNode.
'                specialFolderNode.ImageKey = specialFolderPath
'                specialFolderNode.SelectedImageKey = specialFolderPath

'            Else
'                ' Retrieval failed—use the provided fallback icon to maintain UI consistency.
'                specialFolderNode.ImageKey = FallbackKey
'                specialFolderNode.SelectedImageKey = FallbackKey
'            End If

'        Else
'            ' The icon is already cached—simply assign it to the node.
'            specialFolderNode.ImageKey = specialFolderPath
'            specialFolderNode.SelectedImageKey = specialFolderPath
'        End If

'    End Sub

'    Public Sub SetEasyAccessUserEntryNodeIcon(userEntryNode As TreeNode, folderPath As String)

'        Dim key As String = folderPath

'        ' If we already have the icon cached, use it immediately
'        If IconCache.Images.ContainsKey(key) Then
'            userEntryNode.ImageKey = key
'            Return
'        End If

'        ' Try to get the real Explorer folder icon
'        Dim folderIcon = ShellInterop.GetIconForPath(folderPath, iconSizePix)

'        If folderIcon IsNot Nothing Then
'            IconCache.Images.Add(key, folderIcon.ToBitmap())
'            userEntryNode.ImageKey = key
'            Return
'        End If

'        ' Fallback: generic folder icon, but store it under THIS folder's key
'        Dim genericIcon As Icon = Nothing

'        ' Ensure the shared generic icon exists
'        If Not IconCache.Images.ContainsKey(FolderKey) Then
'            genericIcon = IconLibrary.GenericFolder(iconSizePix)
'            If genericIcon IsNot Nothing Then
'                IconCache.Images.Add(FolderKey, genericIcon.ToBitmap())
'            End If
'        End If

'        ' Now assign the generic icon to THIS folder's key
'        If Not IconCache.Images.ContainsKey(key) Then
'            If genericIcon Is Nothing Then
'                genericIcon = IconLibrary.GenericFolder(iconSizePix)
'            End If

'            If genericIcon IsNot Nothing Then
'                IconCache.Images.Add(key, genericIcon.ToBitmap())
'            End If
'        End If

'        userEntryNode.ImageKey = If(IconCache.Images.ContainsKey(key), key, FallbackFolderKey)








'        'If Not IconCache.Images.ContainsKey(FolderKey) Then

'        '    Dim userEntryIcon = IconLibrary.GenericFolder(iconSizePix)
'        '    'Dim userEasyAccessIcon = Nothing ' Uncomment to test fallback behavior.

'        '    If userEntryIcon IsNot Nothing Then
'        '        IconCache.Images.Add(FolderKey, userEntryIcon.ToBitmap())
'        '        userEntryNode.ImageKey = FolderKey
'        '        userEntryNode.SelectedImageKey = FolderKey
'        '    Else
'        '        userEntryNode.ImageKey = FallbackFolderKey
'        '        userEntryNode.SelectedImageKey = FallbackFolderKey
'        '    End If
'        'Else
'        '    userEntryNode.ImageKey = FolderKey
'        '    userEntryNode.SelectedImageKey = FolderKey
'        'End If

'    End Sub

'    Public Sub SetChildNodeIcon(child As TreeNode, folderPath As String)


'        Dim key As String = folderPath

'        ' If we already have the icon cached, use it immediately
'        If IconCache.Images.ContainsKey(key) Then
'            child.ImageKey = key
'            Return
'        End If

'        ' Try to get the real Explorer folder icon
'        Dim folderIcon = ShellInterop.GetIconForPath(folderPath, iconSizePix)

'        If folderIcon IsNot Nothing Then
'            IconCache.Images.Add(key, folderIcon.ToBitmap())
'            child.ImageKey = key
'            Return
'        End If

'        ' Fallback: generic folder icon, but store it under THIS folder's key
'        Dim genericIcon As Icon = Nothing

'        ' Ensure the shared generic icon exists
'        If Not IconCache.Images.ContainsKey(FolderKey) Then
'            genericIcon = IconLibrary.GenericFolder(iconSizePix)
'            If genericIcon IsNot Nothing Then
'                IconCache.Images.Add(FolderKey, genericIcon.ToBitmap())
'            End If
'        End If

'        ' Now assign the generic icon to THIS folder's key
'        If Not IconCache.Images.ContainsKey(key) Then
'            If genericIcon Is Nothing Then
'                genericIcon = IconLibrary.GenericFolder(iconSizePix)
'            End If

'            If genericIcon IsNot Nothing Then
'                IconCache.Images.Add(key, genericIcon.ToBitmap())
'            End If
'        End If

'        child.ImageKey = If(IconCache.Images.ContainsKey(key), key, FallbackFolderKey)


'        'If Not IconCache.Images.ContainsKey(FolderKey) Then

'        '    Dim folderIcon = IconLibrary.GenericFolder(iconSizePix)
'        '    'Dim folderIcon = Nothing ' force fallback testing

'        '    If folderIcon IsNot Nothing Then
'        '        IconCache.Images.Add(FolderKey, folderIcon.ToBitmap())
'        '        child.ImageKey = FolderKey
'        '        child.SelectedImageKey = FolderKey
'        '    Else
'        '        child.ImageKey = FallbackFolderKey
'        '        child.SelectedImageKey = FallbackFolderKey
'        '    End If
'        'Else
'        '    child.ImageKey = FolderKey
'        '    child.SelectedImageKey = FolderKey
'        'End If

'    End Sub


'    Public Sub SetThisPCNodeIcon(thisPCNode As TreeNode)

'        If Not IconCache.Images.ContainsKey(ThisPCKey) Then

'            Dim thisPCIcon = ShellInterop.GetIconForVirtualFolder(ThisPCGUID, iconSizePix)
'            'Dim thisPCIcon As Icon = Nothing ' Uncomment to test fallback behavior.

'            If thisPCIcon IsNot Nothing Then
'                IconCache.Images.Add(ThisPCKey, thisPCIcon.ToBitmap())
'                thisPCNode.ImageKey = ThisPCKey
'                thisPCNode.SelectedImageKey = ThisPCKey
'            Else
'                thisPCNode.ImageKey = FallbackThisPCKey
'                thisPCNode.SelectedImageKey = FallbackThisPCKey
'            End If
'        Else
'            thisPCNode.ImageKey = ThisPCKey
'            thisPCNode.SelectedImageKey = ThisPCKey
'        End If

'    End Sub


'    Public Sub SetDriveNodeIcon(driveNode As TreeNode, di As DriveInfo)

'        If Not IconCache.Images.ContainsKey(di.RootDirectory.FullName) Then

'            Dim driveIcon = ShellInterop.GetIconForPath(di.RootDirectory.FullName, iconSizePix)
'            'Dim driveIcon As Icon = Nothing ' Uncomment to test fallback behavior.

'            If driveIcon IsNot Nothing Then
'                IconCache.Images.Add(di.RootDirectory.FullName, driveIcon.ToBitmap())
'                driveNode.ImageKey = di.RootDirectory.FullName
'                driveNode.SelectedImageKey = di.RootDirectory.FullName
'            Else
'                If di.DriveType = DriveType.CDRom Then
'                    driveNode.ImageKey = OpticalKey
'                    driveNode.SelectedImageKey = OpticalKey
'                Else
'                    driveNode.ImageKey = DriveKey
'                    driveNode.SelectedImageKey = DriveKey
'                End If
'            End If
'        Else
'            driveNode.ImageKey = di.RootDirectory.FullName
'            driveNode.SelectedImageKey = di.RootDirectory.FullName
'        End If

'    End Sub

'    Public Sub SetRecycleNodeIcon(recycleBinNode As TreeNode)
'        ' Assigns the correct Recycle Bin icon to the provided TreeNode.
'        ' Icons are cached in the ImageList to avoid redundant extraction and improve performance.

'        ' If the ImageList does not already contain the Recycle Bin icon, we need to retrieve and cache it.
'        If Not IconCache.Images.ContainsKey(RecycleBinKey) Then

'            ' Attempt to retrieve the Recycle Bin icon from ShellInterop.
'            ' This handles virtual folders and returns an Explorer-accurate icon when available.
'            Dim recycleBinIcon = ShellInterop.GetIconForVirtualFolder(RecycleBinGUID, iconSizePix)
'            'Dim recycleBinIcon = Nothing ' Uncomment to test fallback behavior.

'            ' If ShellInterop successfully returned an icon, cache it and assign it to the node.
'            If recycleBinIcon IsNot Nothing Then

'                ' Store the icon in the ImageList under the canonical key.
'                IconCache.Images.Add(RecycleBinKey, recycleBinIcon.ToBitmap())

'                ' Point the TreeNode to the cached icon so future loads do not require ShellInterop again.
'                recycleBinNode.ImageKey = RecycleBinKey
'                recycleBinNode.SelectedImageKey = RecycleBinKey

'            Else
'                ' ShellInterop failed to retrieve the icon (rare but possible).
'                ' Use the predefined fallback icon to keep the UI consistent and predictable.
'                recycleBinNode.ImageKey = FallbackRecycleBinKey
'                recycleBinNode.SelectedImageKey = FallbackRecycleBinKey
'            End If

'        Else
'            ' The icon is already cached—simply assign it to the node.
'            ' This avoids unnecessary interop calls and keeps UI updates fast.
'            recycleBinNode.ImageKey = RecycleBinKey
'            recycleBinNode.SelectedImageKey = RecycleBinKey
'        End If

'    End Sub

'    Public Sub SetListViewItemIconForFile(item As ListViewItem, fi As FileInfo)

'        Dim ext = fi.Extension.ToLowerInvariant()

'        ' Files with no extension get a generic file icon,
'        ' grouped under the FileKey to avoid duplicates.
'        If ext = "" Then
'            If Not IconCache.Images.ContainsKey(FileKey) Then
'                Dim fileIcon = IconLibrary.GenericFile(iconSizePix)
'                If fileIcon IsNot Nothing Then
'                    IconCache.Images.Add(FileKey, fileIcon.ToBitmap())
'                    ' use generic file icon for no-extension files
'                    item.ImageKey = FileKey
'                Else
'                    item.ImageKey = FallbackFileKey
'                End If
'            Else
'                item.ImageKey = FileKey
'            End If
'            ' Files with extensions get icons based on their extension,
'            ' grouped by extension to avoid duplicates.
'        ElseIf Not ext = ".exe" Then
'            If Not IconCache.Images.ContainsKey(ext) Then
'                Dim extensionIcon = ShellInterop.GetIconForPath(fi.FullName, iconSizePix)
'                If extensionIcon IsNot Nothing Then
'                    IconCache.Images.Add(ext, extensionIcon.ToBitmap())
'                    ' use extension as key to group same-type files
'                    item.ImageKey = ext
'                Else
'                    item.ImageKey = FileKey
'                End If
'            Else
'                item.ImageKey = ext
'            End If
'            ' Executable files get icons based on their full path to preserve unique icons
'            ' for different executables, since many .exe files have custom icons.
'        Else
'            If Not IconCache.Images.ContainsKey(fi.FullName) Then
'                Dim executableIcon = ShellInterop.GetIconForPath(fi.FullName, iconSizePix)
'                If executableIcon IsNot Nothing Then
'                    IconCache.Images.Add(fi.FullName, executableIcon.ToBitmap())
'                    ' use full path as key to preserve unique icons for different executables
'                    item.ImageKey = fi.FullName
'                Else
'                    item.ImageKey = ExecutableKey
'                End If
'            Else
'                item.ImageKey = fi.FullName
'            End If
'        End If


'    End Sub

'    'Public Sub SetListViewItemIconForDirectory(item As ListViewItem, di As DirectoryInfo)
'    '    ' Set icon, with caching in the image list to avoid duplicates and improve performance

'    '    If Not IconCache.Images.ContainsKey(di.FullName) Then

'    '        If Not IconCache.Images.ContainsKey(FolderKey) Then

'    '            Dim folderIcon = IconLibrary.GenericFolder(iconSizePix)

'    '            If folderIcon IsNot Nothing Then
'    '                IconCache.Images.Add(FolderKey, folderIcon.ToBitmap())
'    '                item.ImageKey = FolderKey
'    '            Else
'    '                item.ImageKey = FallbackFolderKey
'    '            End If

'    '        Else
'    '            item.ImageKey = FolderKey
'    '        End If

'    '    Else
'    '        item.ImageKey = di.FullName
'    '    End If

'    'End Sub


'    'Public Sub SetListViewItemIconForDirectory(item As ListViewItem, di As DirectoryInfo)

'    '    Dim key As String = di.FullName

'    '    ' If we already have the icon cached, use it immediately
'    '    If IconCache.Images.ContainsKey(key) Then
'    '        item.ImageKey = key
'    '        Return
'    '    End If

'    '    ' Try to get the real Explorer folder icon
'    '    Dim folderIcon = ShellInterop.GetIconForPath(di.FullName, iconSizePix)

'    '    If folderIcon IsNot Nothing Then
'    '        IconCache.Images.Add(key, folderIcon.ToBitmap())
'    '        item.ImageKey = key
'    '        Return
'    '    End If

'    '    ' Fallback: generic folder icon
'    '    If Not IconCache.Images.ContainsKey(FolderKey) Then
'    '        Dim genericIcon = IconLibrary.GenericFolder(iconSizePix)
'    '        If genericIcon IsNot Nothing Then
'    '            IconCache.Images.Add(FolderKey, genericIcon.ToBitmap())
'    '        End If
'    '    End If

'    '    item.ImageKey = If(IconCache.Images.ContainsKey(FolderKey), FolderKey, FallbackFolderKey)

'    'End Sub


'    Public Sub SetListViewItemIconForDirectory(item As ListViewItem, di As DirectoryInfo)

'        Dim key As String = di.FullName

'        ' If we already have the icon cached, use it immediately
'        If IconCache.Images.ContainsKey(key) Then
'            item.ImageKey = key
'            Return
'        End If

'        ' Try to get the real Explorer folder icon
'        Dim folderIcon = ShellInterop.GetIconForPath(di.FullName, iconSizePix)

'        If folderIcon IsNot Nothing Then
'            IconCache.Images.Add(key, folderIcon.ToBitmap())
'            item.ImageKey = key
'            Return
'        End If

'        ' Fallback: generic folder icon, but store it under THIS folder's key
'        Dim genericIcon As Icon = Nothing

'        ' Ensure the shared generic icon exists
'        If Not IconCache.Images.ContainsKey(FolderKey) Then
'            genericIcon = IconLibrary.GenericFolder(iconSizePix)
'            If genericIcon IsNot Nothing Then
'                IconCache.Images.Add(FolderKey, genericIcon.ToBitmap())
'            End If
'        End If

'        ' Now assign the generic icon to THIS folder's key
'        If Not IconCache.Images.ContainsKey(key) Then
'            If genericIcon Is Nothing Then
'                genericIcon = IconLibrary.GenericFolder(iconSizePix)
'            End If

'            If genericIcon IsNot Nothing Then
'                IconCache.Images.Add(key, genericIcon.ToBitmap())
'            End If
'        End If

'        item.ImageKey = If(IconCache.Images.ContainsKey(key), key, FallbackFolderKey)

'    End Sub




'    Public Sub UpdateIconSize()
'        ' Updates the icon size based on the current DPI scaling of the form.
'        ' This should be called when the form's DPI changes to ensure icons are rendered at the correct size.
'        iconSizePix = GetScaledIconSize(UIForm)
'        IconCache.ImageSize = New Size(iconSizePix, iconSizePix)
'        ClearCache() ' Clear cache to ensure icons are reloaded at the new size.
'    End Sub

'    Private Sub ClearCache()
'        ' Clears the icon cache. This can be used to free memory or to refresh icons if needed.
'        IconCache.Images.Clear()
'    End Sub

'    Public Function GetIconForPath(folderPath As String) As Bitmap
'        ' Retrieves an icon for a given folder path, using the cache if available.
'        ' This is a convenience method that combines caching logic with ShellInterop retrieval.
'        If Not IconCache.Images.ContainsKey(folderPath) Then
'            Dim icon = ShellInterop.GetIconForPath(folderPath, iconSizePix)
'            If icon IsNot Nothing Then
'                IconCache.Images.Add(folderPath, icon.ToBitmap())
'            End If
'        End If
'        Return If(IconCache.Images.ContainsKey(folderPath), IconCache.Images(folderPath), Nothing)
'    End Function

'    Public Function GetIconForVirtualFolder(folderGUID As String) As Bitmap
'        ' Retrieves an icon for a given virtual folder GUID, using the cache if available.
'        ' This is a convenience method that combines caching logic with ShellInterop retrieval for virtual folders.
'        If Not IconCache.Images.ContainsKey(folderGUID) Then
'            Dim icon = ShellInterop.GetIconForVirtualFolder(folderGUID, iconSizePix)
'            If icon IsNot Nothing Then
'                IconCache.Images.Add(folderGUID, icon.ToBitmap())
'            End If
'        End If
'        Return If(IconCache.Images.ContainsKey(folderGUID), IconCache.Images(folderGUID), Nothing)
'    End Function

'    Public Sub AddFallbackIcon(key As String, icon As Bitmap)
'        ' Allows adding a fallback icon to the cache under a specific key.
'        ' This can be used to ensure that there is always an icon available for certain keys, even if ShellInterop retrieval fails.
'        If Not IconCache.Images.ContainsKey(key) Then
'            IconCache.Images.Add(key, icon)
'        End If
'    End Sub

'    Public Function GetIconCache() As ImageList
'        ' Provides access to the shared ImageList cache for icons.
'        ' This allows other parts of the application to retrieve cached icons by key.
'        Return IconCache
'    End Function

'    Public Function GetScaledIconSize(form As Form) As Integer
'        Dim scale As Double = form.DeviceDpi / 96.0
'        Dim pixelSize As Integer = CInt(16 * scale)



'        'If pixelSize <= 16 Then
'        '    Return 16
'        'End If

'        Return pixelSize

'    End Function


'End Class



Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Threading
Imports File_Explorer.Explorer.Interop.Shell

Public Class IconEngine

    ' --- Canonical keys and constants ---------------------------------------

    Private Const ThisPCGUID As String = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    Private Const ThisPCKey As String = "ThisPC"
    Private Const FallbackThisPCKey As String = "FallbackThisPC"

    Private Const RecycleBinGUID As String = "shell:::{645FF040-5081-101B-9F08-00AA002F954E}"
    Private Const RecycleBinKey As String = "RecycleBin"
    Private Const FallbackRecycleBinKey As String = "Recycle"

    Private Const OpticalKey As String = "Optical"
    Private Const DriveKey As String = "Drive"

    Private Const FolderKey As String = "Folder"
    Private Const FallbackFolderKey As String = "FallbackFolder"

    Private Const FileKey As String = "File"
    Private Const FallbackFileKey As String = "FallbackFile"

    Private Const ExecutableKey As String = "FallbackExecutable"

    Private ReadOnly IconCache As New ImageList()
    Private ReadOnly UIForm As Form
    Private iconSizePix As Integer

    ' --- Construction / DPI -------------------------------------------------

    Public Sub New(form As Form)
        UIForm = form
        iconSizePix = GetScaledIconSize(UIForm)

        IconCache.ColorDepth = ColorDepth.Depth32Bit
        IconCache.ImageSize = New Size(iconSizePix, iconSizePix)
    End Sub

    Public Sub UpdateIconSize()
        iconSizePix = GetScaledIconSize(UIForm)
        IconCache.ImageSize = New Size(iconSizePix, iconSizePix)
        ClearCache()
    End Sub

    Private Sub ClearCache()
        IconCache.Images.Clear()
    End Sub

    Public Function GetIconCache() As ImageList
        Return IconCache
    End Function

    Public Function GetScaledIconSize(form As Form) As Integer
        Dim scale As Double = form.DeviceDpi / 96.0
        Dim pixelSize As Integer = CInt(16 * scale)
        Return pixelSize
    End Function

    ' --- Core declarative helpers -------------------------------------------

    Private Function EnsureIcon(key As String, factory As Func(Of Icon), fallbackKey As String) As String
        ' Try to ensure an icon exists under "key"; if factory fails, fall back to fallbackKey.
        If IconCache.Images.ContainsKey(key) Then
            Return key
        End If

        Dim ic As Icon = factory()
        If ic IsNot Nothing Then
            IconCache.Images.Add(key, ic.ToBitmap())
            Return key
        End If

        Return fallbackKey
    End Function

    Private Function EnsureSharedIcon(sharedKey As String,
                                      factory As Func(Of Icon),
                                      fallbackKey As String) As String
        ' Ensure a shared icon exists under sharedKey; if factory fails, return fallbackKey.
        If IconCache.Images.ContainsKey(sharedKey) Then
            Return sharedKey
        End If

        Dim ic As Icon = factory()
        If ic IsNot Nothing Then
            IconCache.Images.Add(sharedKey, ic.ToBitmap())
            Return sharedKey
        End If

        Return fallbackKey
    End Function

    ' --- Folder icon helpers ------------------------------------------------

    Private Function GetFolderIconKey(folderPath As String) As String
        ' Real folder icon per path; fallback is shared generic folder icon.
        If IconCache.Images.ContainsKey(folderPath) Then
            Return folderPath
        End If

        Dim realIcon = ShellInterop.GetIconForPath(folderPath, iconSizePix)
        If realIcon IsNot Nothing Then
            IconCache.Images.Add(folderPath, realIcon.ToBitmap())
            Return folderPath
        End If

        ' Shared generic folder icon
        Return EnsureSharedIcon(
            FolderKey,
            Function() IconLibrary.GenericFolder(iconSizePix),
            FallbackFolderKey
        )
    End Function

    Private Function GetVirtualFolderIconKey(cacheKey As String,
                                             folderGuid As String,
                                             fallbackKey As String) As String
        Return EnsureIcon(
            cacheKey,
            Function() ShellInterop.GetIconForVirtualFolder(folderGuid, iconSizePix),
            fallbackKey
        )
    End Function

    Private Function GetSpecialFolderIconKey(folderPath As String,
                                             fallbackKey As String) As String
        ' Same semantics as normal folder, but with a caller‑supplied fallback key.
        If IconCache.Images.ContainsKey(folderPath) Then
            Return folderPath
        End If

        Dim ic = ShellInterop.GetIconForPath(folderPath, iconSizePix)
        If ic IsNot Nothing Then
            IconCache.Images.Add(folderPath, ic.ToBitmap())
            Return folderPath
        End If

        Return fallbackKey
    End Function

    ' --- Drive icon helper --------------------------------------------------

    Private Function GetDriveIconKey(di As DriveInfo) As String
        Dim key As String = di.RootDirectory.FullName

        If IconCache.Images.ContainsKey(key) Then
            Return key
        End If

        Dim ic = ShellInterop.GetIconForPath(key, iconSizePix)
        If ic IsNot Nothing Then
            IconCache.Images.Add(key, ic.ToBitmap())
            Return key
        End If

        ' Fallback: optical vs generic drive
        If di.DriveType = DriveType.CDRom Then
            Return OpticalKey
        Else
            Return DriveKey
        End If
    End Function

    ' --- File icon helper ---------------------------------------------------

    Private Function GetFileIconKey(fi As FileInfo) As String
        Dim ext = fi.Extension.ToLowerInvariant()

        ' No extension → shared generic file icon
        If ext = "" Then
            Return EnsureSharedIcon(
                FileKey,
                Function() IconLibrary.GenericFile(iconSizePix),
                FallbackFileKey
            )
        End If

        ' Non‑exe: group by extension
        If ext <> ".exe" Then
            If IconCache.Images.ContainsKey(ext) Then
                Return ext
            End If

            Dim ic = ShellInterop.GetIconForPath(fi.FullName, iconSizePix)
            If ic IsNot Nothing Then
                IconCache.Images.Add(ext, ic.ToBitmap())
                Return ext
            End If

            ' Fallback to generic file icon
            Return EnsureSharedIcon(
                FileKey,
                Function() IconLibrary.GenericFile(iconSizePix),
                FallbackFileKey
            )
        End If

        ' Executables: per‑file icon
        Dim exeKey As String = fi.FullName
        If IconCache.Images.ContainsKey(exeKey) Then
            Return exeKey
        End If

        Dim exeIcon = ShellInterop.GetIconForPath(fi.FullName, iconSizePix)
        If exeIcon IsNot Nothing Then
            IconCache.Images.Add(exeKey, exeIcon.ToBitmap())
            Return exeKey
        End If

        Return ExecutableKey
    End Function

    ' --- Public node / item APIs -------------------------------------------

    Public Sub SetSpecialFolderNodeIcon(node As TreeNode,
                                        specialFolderPath As String,
                                        fallbackKey As String)
        Dim key = GetSpecialFolderIconKey(specialFolderPath, fallbackKey)
        node.ImageKey = key
        node.SelectedImageKey = key
    End Sub

    Public Sub SetEasyAccessUserEntryNodeIcon(node As TreeNode,
                                              folderPath As String)
        Dim key = GetFolderIconKey(folderPath)
        node.ImageKey = key
        node.SelectedImageKey = key
    End Sub

    Public Sub SetChildNodeIcon(child As TreeNode,
                                folderPath As String)
        Dim key = GetFolderIconKey(folderPath)
        child.ImageKey = key
        child.SelectedImageKey = key
    End Sub

    Public Sub SetThisPCNodeIcon(node As TreeNode)
        Dim key = GetVirtualFolderIconKey(ThisPCKey, ThisPCGUID, FallbackThisPCKey)
        node.ImageKey = key
        node.SelectedImageKey = key
    End Sub

    Public Sub SetDriveNodeIcon(node As TreeNode,
                                di As DriveInfo)
        Dim key = GetDriveIconKey(di)
        node.ImageKey = key
        node.SelectedImageKey = key
    End Sub

    Public Sub SetRecycleNodeIcon(node As TreeNode)
        Dim key = GetVirtualFolderIconKey(RecycleBinKey, RecycleBinGUID, FallbackRecycleBinKey)
        node.ImageKey = key
        node.SelectedImageKey = key
    End Sub

    Public Sub SetListViewItemIconForFile(item As ListViewItem,
                                          fi As FileInfo)
        Dim key = GetFileIconKey(fi)
        item.ImageKey = key
    End Sub

    Public Sub SetListViewItemIconForDirectory(item As ListViewItem,
                                               di As DirectoryInfo)
        Dim key = GetFolderIconKey(di.FullName)
        item.ImageKey = key
    End Sub

    ' --- Convenience bitmap accessors --------------------------------------

    Public Function GetIconForPath(path As String) As Bitmap
        Dim key = GetFolderIconKey(path)
        If IconCache.Images.ContainsKey(key) Then
            Return IconCache.Images(key)
        End If
        Return Nothing
    End Function

    Public Function GetIconForVirtualFolder(folderGuid As String,
                                            cacheKey As String) As Bitmap
        Dim key = GetVirtualFolderIconKey(cacheKey, folderGuid, cacheKey)
        If IconCache.Images.ContainsKey(key) Then
            Return IconCache.Images(key)
        End If
        Return Nothing
    End Function

    Public Sub AddFallbackIcon(key As String, icon As Bitmap)
        If Not IconCache.Images.ContainsKey(key) Then
            IconCache.Images.Add(key, icon)
        End If
    End Sub

End Class



