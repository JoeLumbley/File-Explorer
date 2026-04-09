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

        LoadFallbackIcons()

    End Sub

    Public Sub UpdateIconSize()
        ' Updates the icon size based on the current DPI scaling of the form.
        ' This should be called when the form's DPI changes to ensure icons are rendered at the correct size.
        iconSizePix = GetScaledIconSize(UIForm)
        IconCache.ImageSize = New Size(iconSizePix, iconSizePix)
        ClearCache()
    End Sub

    Private Sub ClearCache()
        IconCache.Images.Clear()
    End Sub

    Public Function GetIconCache() As ImageList
        ' Provides access to the shared ImageList cache for icons.
        ' This allows other parts of the application to retrieve cached icons by key.
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
        'exeIcon = Nothing ' Uncomment to test fallback behavior for executables

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

    Public Function GetIconForFolder(folderPath As String) As Bitmap
        ' Retrieves an icon for a given folder path, using the cache if available.
        ' This is a convenience method that combines caching logic with ShellInterop retrieval.
        Dim key = GetFolderIconKey(folderPath)
        If IconCache.Images.ContainsKey(key) Then
            Return IconCache.Images(key)
        End If
        Return Nothing
    End Function

    Public Function GetIconForVirtualFolder(folderGuid As String,
                                            cacheKey As String) As Bitmap
        ' Retrieves an icon for a given virtual folder GUID, using the cache if
        ' available. This is a convenience method that combines caching logic
        ' with ShellInterop retrieval for virtual folders.
        Dim key = GetVirtualFolderIconKey(cacheKey, folderGuid, cacheKey)
        If IconCache.Images.ContainsKey(key) Then
            Return IconCache.Images(key)
        End If
        Return Nothing
    End Function

    Public Sub AddFallbackIcon(key As String, icon As Bitmap)
        ' Allows adding a fallback icon to the cache under a specific key.
        ' This can be used to ensure that there is always an icon available for
        ' certain keys, even if ShellInterop retrieval fails.
        If Not IconCache.Images.ContainsKey(key) Then
            IconCache.Images.Add(key, icon)
        End If
    End Sub



    ' Loading fallback icons at initialization
    Private Sub LoadFallbackIcons()
        ' Loads fallback icons into the cache. This should be called during
        ' application initialization to ensure that fallback icons are available
        ' for use when ShellInterop retrieval fails.

        AddFallbackIcon(FallbackFolderKey, GetScaledBitmap(My.Resources.Resource1.Folder_16X16))
        AddFallbackIcon(FallbackFileKey, GetScaledBitmap(My.Resources.Resource1.Documents_16X16))
        AddFallbackIcon(OpticalKey, GetScaledBitmap(My.Resources.Resource1.Optical_16X16))
        AddFallbackIcon(DriveKey, GetScaledBitmap(My.Resources.Resource1.Drive_16X16))

        AddFallbackIcon("Documents", GetScaledBitmap(My.Resources.Resource1.Documents_16X16))
        AddFallbackIcon("Downloads", GetScaledBitmap(My.Resources.Resource1.Downloads_16X16))
        AddFallbackIcon("Desktop", GetScaledBitmap(My.Resources.Resource1.Desktop_16X16))
        AddFallbackIcon("Pictures", GetScaledBitmap(My.Resources.Resource1.Pictures_16X16))
        AddFallbackIcon("Music", GetScaledBitmap(My.Resources.Resource1.Music_16X16))
        AddFallbackIcon("Videos", GetScaledBitmap(My.Resources.Resource1.Videos_16X16))
        AddFallbackIcon("Shortcuts", GetScaledBitmap(My.Resources.Resource1.Shortcut_16X16))
        AddFallbackIcon("EasyAccess", GetScaledBitmap(My.Resources.Resource1.Easy_Access_16X16))

        AddFallbackIcon(FallbackThisPCKey, GetScaledBitmap(My.Resources.Resource1.Computer_16X16))
        AddFallbackIcon(FallbackRecycleBinKey, GetScaledBitmap(My.Resources.Resource1.Recycle_16X16))

        AddFallbackIcon(ExecutableKey, GetScaledBitmap(My.Resources.Resource1.Executable_16X16))

    End Sub

    Private Function GetScaledBitmap(src As Bitmap) As Bitmap

        If iconSizePix <= 16 Then
            Return src
        End If

        Dim baseIcon = src

        If baseIcon Is Nothing Then Return Nothing

        Return ResizeBitmapHighQuality(baseIcon, iconSizePix)

    End Function

    Private Shared Function ResizeBitmapHighQuality(src As Bitmap, size As Integer) As Bitmap
        Dim bmp As New Bitmap(size, size)
        Using g = Graphics.FromImage(bmp)
            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            g.DrawImage(src, New Rectangle(0, 0, size, size),
                        New Rectangle(0, 0, src.Width, src.Height), GraphicsUnit.Pixel)
        End Using
        Return bmp
    End Function


End Class



