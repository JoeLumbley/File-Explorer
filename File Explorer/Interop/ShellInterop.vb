Imports System.Runtime.InteropServices

Namespace Explorer.Interop.Shell

    ' ============================================================
    '   CORE ENUMS
    ' ============================================================
    Public Enum IconSize
        Small
        Large
        ExtraLarge
    End Enum

    <Flags>
    Public Enum FileOperationFlags As UInteger
        FOF_NOCONFIRMATION = &H10UI
        FOF_NOCONFIRMMKDIR = &H200UI
        FOFX_NOCOPYSECURITYATTRIBS = &H800UI
        FOFX_REQUIREELEVATION = &H200000UI
        FOFX_SHOWELEVATIONPROMPT = &H40000UI
    End Enum

    Public Enum SIGDN As UInteger
        SIGDN_NORMALDISPLAY = &H0UI
        SIGDN_PARENTRELATIVEPARSING = &H80018001UI
        SIGDN_DESKTOPABSOLUTEPARSING = &H80028000UI
        SIGDN_PARENTRELATIVEEDITING = &H80031001UI
        SIGDN_DESKTOPABSOLUTEEDITING = &H8004C000UI
        SIGDN_FILESYSPATH = &H80058000UI
        SIGDN_URL = &H80068000UI
    End Enum


    ' ============================================================
    '   SHELL INTERFACES
    ' ============================================================
    <ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IShellItem
        Sub BindToHandler(pbc As IntPtr, ByRef bhid As Guid, ByRef riid As Guid, ByRef ppv As IntPtr)
        Sub GetParent(ByRef ppsi As IShellItem)
        Sub GetDisplayName(sigdnName As SIGDN, ByRef ppszName As IntPtr)
        Sub GetAttributes(sfgaoMask As UInteger, ByRef psfgaoAttribs As UInteger)
        Sub Compare(psi As IShellItem, hint As UInteger, ByRef piOrder As Integer)
    End Interface


    ' ============================================================
    '   SHELLIDLISTARRAY (for "Shell IDList Array" drag-drop)
    ' ============================================================
    Public NotInheritable Class ShellIDListArray

        Public ReadOnly Property Items As List(Of IntPtr)

        Private Sub New(list As List(Of IntPtr))
            Me.Items = list
        End Sub

        Public Shared Function FromDataObject(data As IDataObject) As ShellIDListArray
            Dim raw = TryCast(data.GetData("Shell IDList Array"), Byte())
            If raw Is Nothing Then
                Return New ShellIDListArray(New List(Of IntPtr))
            End If

            Dim list As New List(Of IntPtr)
            Dim offset As Integer = 0

            ' First 4 bytes = count
            Dim count As Integer = BitConverter.ToInt32(raw, offset)
            offset += 4

            For i = 0 To count - 1
                Dim pidlSize As Integer = BitConverter.ToInt32(raw, offset)
                offset += 4

                Dim pidlPtr As IntPtr = Marshal.AllocHGlobal(pidlSize)
                Marshal.Copy(raw, offset, pidlPtr, pidlSize)
                offset += pidlSize

                list.Add(pidlPtr)
            Next

            Return New ShellIDListArray(list)
        End Function

    End Class


    ' ============================================================
    '   SHELL INTEROP
    ' ============================================================
    Public NotInheritable Class ShellInterop

        Private Sub New()
        End Sub


        ' ------------------------------------------------------------
        '   SHCreateItemFromParsingName (path → IShellItem)
        ' ------------------------------------------------------------
        <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
        Private Shared Function SHCreateItemFromParsingName(
            pszPath As String,
            pbc As IntPtr,
            ByRef riid As Guid,
            <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object) As Integer
        End Function

        Public Shared Function CreateShellItemFromPath(path As String) As IShellItem
            Dim iid As New Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")
            Dim obj As Object = Nothing
            SHCreateItemFromParsingName(path, IntPtr.Zero, iid, obj)
            Return CType(obj, IShellItem)
        End Function


        ' ------------------------------------------------------------
        '   SHCreateItemFromIDList (PIDL → IShellItem)
        ' ------------------------------------------------------------
        <DllImport("shell32.dll")>
        Private Shared Function SHCreateItemFromIDList(
            pidl As IntPtr,
            ByRef riid As Guid,
            <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object) As Integer
        End Function

        Public Shared Function CreateShellItemFromPIDL(pidl As IntPtr) As IShellItem
            Dim iid As New Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")
            Dim obj As Object = Nothing
            SHCreateItemFromIDList(pidl, iid, obj)
            Return CType(obj, IShellItem)
        End Function


        ' ------------------------------------------------------------
        '   DISPLAY NAME FROM IShellItem
        ' ------------------------------------------------------------
        Public Shared Function GetDisplayName(item As IShellItem, sigdn As SIGDN) As String
            Dim psz As IntPtr = IntPtr.Zero
            item.GetDisplayName(sigdn, psz)
            If psz = IntPtr.Zero Then Return Nothing
            Dim result As String = Marshal.PtrToStringUni(psz)
            CoTaskMemFree(psz)
            Return result
        End Function


        ' ------------------------------------------------------------
        '   COPY VIA SHFileOperation (Explorer-compatible)
        ' ------------------------------------------------------------
        Public Shared Sub CopyShellItem(source As IShellItem, destinationFolder As IShellItem)
            Dim srcPath As String = GetDisplayName(source, SIGDN.SIGDN_FILESYSPATH)
            Dim destPath As String = GetDisplayName(destinationFolder, SIGDN.SIGDN_FILESYSPATH)

            If String.IsNullOrEmpty(srcPath) OrElse String.IsNullOrEmpty(destPath) Then
                Exit Sub
            End If

            Dim op As New SHFILEOPSTRUCT With {
                .hwnd = IntPtr.Zero,
                .wFunc = FO_COPY,
                .pFrom = srcPath & vbNullChar & vbNullChar,
                .pTo = destPath & vbNullChar & vbNullChar,
                .fFlags = CShort(FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR)
            }

            SHFileOperation(op)
        End Sub


        ' ------------------------------------------------------------
        '   ICON EXTRACTION (path + PIDL)
        ' ------------------------------------------------------------
        <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function SHGetFileInfo(
            pszPath As String,
            dwFileAttributes As UInteger,
            ByRef psfi As SHFILEINFO,
            cbFileInfo As UInteger,
            uFlags As UInteger) As IntPtr
        End Function

        <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function SHGetFileInfo(
            pidl As IntPtr,
            dwFileAttributes As UInteger,
            ByRef psfi As SHFILEINFO,
            cbFileInfo As UInteger,
            uFlags As UInteger) As IntPtr
        End Function

        <DllImport("user32.dll")>
        Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
        End Function

        Public Shared Function GetIconForPath(path As String, iconSize As Integer) As Icon

            If iconSize <= 16 Then
                Return GetShellIcon(path, Shell.IconSize.Small)
            End If

            Dim baseIcon = GetShellIcon(path, Shell.IconSize.Small)

            If baseIcon Is Nothing Then Return Nothing

            Return ResizeIconHighQuality(baseIcon, iconSize)

        End Function


        Public Shared Function GetIconForVirtualFolder(guidString As String, iconSize As Integer) As Icon
            Dim pidl As IntPtr = ParseDisplayName(guidString)
            If pidl = IntPtr.Zero Then Return Nothing

            Try

                If iconSize <= 16 Then
                    Return GetIconForPIDL(pidl, Shell.IconSize.Small)
                End If

                Dim baseIcon = GetIconForPIDL(pidl, Shell.IconSize.Small)

                If baseIcon Is Nothing Then Return Nothing

                Return ResizeIconHighQuality(baseIcon, iconSize)

            Finally
                CoTaskMemFree(pidl)
            End Try
        End Function

        Private Shared Function ResizeIconHighQuality(src As Icon, size As Integer) As Icon
            Using bmp As New Bitmap(size, size)
                Using g = Graphics.FromImage(bmp)
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawIcon(src, New Rectangle(0, 0, size, size))
                End Using
                Return Icon.FromHandle(bmp.GetHicon())
            End Using
        End Function

        ' ------------------------------------------------------------
        '   INTERNAL HELPERS FOR ICONS & PIDLS
        ' ------------------------------------------------------------
        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function SHParseDisplayName(
            pszName As String,
            pbc As IntPtr,
            ByRef ppidl As IntPtr,
            sfgaoIn As UInteger,
            ByRef psfgaoOut As UInteger) As Integer
        End Function

        Public Shared Function ParseDisplayName(path As String) As IntPtr
            Dim pidl As IntPtr = IntPtr.Zero
            Dim attrs As UInteger = 0UI
            Dim hr = SHParseDisplayName(path, IntPtr.Zero, pidl, 0UI, attrs)
            If hr <> 0 Then Return IntPtr.Zero
            Return pidl
        End Function

        <DllImport("shell32.dll")>
        Private Shared Function SHGetSpecialFolderLocation(
            hwndOwner As IntPtr,
            nFolder As Integer,
            ByRef ppidl As IntPtr) As Integer
        End Function

        <DllImport("ole32.dll")>
        Private Shared Sub CoTaskMemFree(pv As IntPtr)
        End Sub

        Private Const CSIDL_BITBUCKET As Integer = &HA

        Private Shared Function GetShellIcon(path As String, size As IconSize) As Icon
            Dim flags As UInteger = SHGFI_ICON

            If size = IconSize.Small Then
                flags = flags Or SHGFI_SMALLICON
            Else
                flags = flags Or SHGFI_LARGEICON
            End If

            Dim shinfo As New SHFILEINFO()

            Dim attrs As UInteger

            If IO.File.Exists(path) OrElse IO.Directory.Exists(path) Then
                ' Real file → do NOT use SHGFI_USEFILEATTRIBUTES
                attrs = 0UI
            Else
                ' Missing file → fall back to extension icon
                flags = flags Or SHGFI_USEFILEATTRIBUTES
                attrs = FILE_ATTRIBUTE_NORMAL
            End If

            Dim result = SHGetFileInfo(
        path,
        attrs,
        shinfo,
        CUInt(Marshal.SizeOf(shinfo)),
        flags)

            If result = IntPtr.Zero Then Return Nothing

            Dim rawIcon As Icon = Icon.FromHandle(shinfo.hIcon)
            Dim cloned As Icon = CType(rawIcon.Clone(), Icon)
            DestroyIcon(shinfo.hIcon)
            Return cloned
        End Function

        Private Shared Function GetIconForPIDL(pidl As IntPtr, size As IconSize) As Icon
            Dim flags As UInteger = SHGFI_ICON Or SHGFI_PIDL

            If size = IconSize.Small Then
                flags = flags Or SHGFI_SMALLICON
            Else
                flags = flags Or SHGFI_LARGEICON
            End If

            Dim shinfo As New SHFILEINFO()

            Dim result = SHGetFileInfo(
                pidl,
                0UI,
                shinfo,
                CUInt(Marshal.SizeOf(shinfo)),
                flags)

            If result = IntPtr.Zero Then Return Nothing

            Dim rawIcon As Icon = Icon.FromHandle(shinfo.hIcon)
            Dim cloned As Icon = CType(rawIcon.Clone(), Icon)
            DestroyIcon(shinfo.hIcon)
            Return cloned
        End Function


        ' ------------------------------------------------------------
        '   SHFILEOPERATION (Copy/Move/Delete)
        ' ------------------------------------------------------------
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Private Structure SHFILEOPSTRUCT
            Public hwnd As IntPtr
            Public wFunc As UInteger
            Public pFrom As String
            Public pTo As String
            Public fFlags As UShort
            Public fAnyOperationsAborted As Boolean
            Public hNameMappings As IntPtr
            Public lpszProgressTitle As String
        End Structure

        Private Const FO_COPY As UInteger = &H2UI
        Private Const FO_MOVE As UInteger = &H1UI
        Private Const FO_DELETE As UInteger = &H3UI
        Private Const FO_RENAME As UInteger = &H4UI

        Private Const FOF_SILENT As UShort = &H4US
        Private Const FOF_NOCONFIRMATION As UShort = &H10US
        Private Const FOF_NOCONFIRMMKDIR As UShort = &H200US
        Private Const FOF_ALLOWUNDO As UShort = &H40US

        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function SHFileOperation(ByRef lpFileOp As SHFILEOPSTRUCT) As Integer
        End Function


        ' ------------------------------------------------------------
        '   SHFILEINFO STRUCT & CONSTANTS
        ' ------------------------------------------------------------
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
        Private Structure SHFILEINFO
            Public hIcon As IntPtr
            Public iIcon As Integer
            Public dwAttributes As UInteger
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
            Public szDisplayName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
            Public szTypeName As String
        End Structure

        Private Const SHGFI_ICON As UInteger = &H100UI
        Private Const SHGFI_PIDL As UInteger = &H8UI
        Private Const SHGFI_SMALLICON As UInteger = &H1UI
        Private Const SHGFI_LARGEICON As UInteger = &H0UI

        Private Const SHGFI_USEFILEATTRIBUTES As UInteger = &H10UI

        Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80UI
        Private Const FILE_ATTRIBUTE_DIRECTORY As UInteger = &H10UI

    End Class

End Namespace