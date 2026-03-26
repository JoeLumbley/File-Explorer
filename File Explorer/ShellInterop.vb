'Imports System
'Imports System.ComponentModel
'Imports System.Diagnostics
'Imports System.Drawing
'Imports System.IO
'Imports System.Runtime.InteropServices
'Imports System.Runtime.Versioning
'Imports System.Security
'Imports System.Text
'Imports File_Explorer.Explorer.Navigation
'Imports Microsoft.VisualBasic

'Namespace Explorer.Interop

'    <Flags>
'    Public Enum FileOperationFlags As UInteger
'        FOF_SILENT = &H4UI
'        FOF_NOCONFIRMATION = &H10UI
'        FOF_ALLOWUNDO = &H40UI
'        FOF_NOCONFIRMMKDIR = &H200UI
'        FOF_NOERRORUI = &H400UI

'        ' Extended flags (FOFX)
'        FOFX_NOCONFIRMMKDIR = &H200UI
'        FOFX_SHOWELEVATIONPROMPT = &H40000UI
'        FOFX_RECYCLEONDELETE = &H80000UI
'    End Enum

'    <ComImport(), Guid("3AD05575-8857-4850-9277-11B85BDB8E09"),
'     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'    Public Interface IFileOperation
'        Sub Advise(pfops As IFileOperationProgressSink, ByRef pdwCookie As UInteger)
'        Sub Unadvise(dwcookie As UInteger)
'        Sub SetOperationFlags(dwOperationFlags As FileOperationFlags)
'        Sub SetOwnerWindow(hwndOwner As IntPtr)

'        ' We only need DeleteItem + PerformOperations for this subsystem
'        Sub DeleteItem(psiItem As IShellItem, pfopsItem As IFileOperationProgressSink)
'        Sub PerformOperations()
'        Function GetAnyOperationsAborted() As Boolean
'    End Interface

'    <ComImport(), Guid("04b0f1a7-9490-44bc-96e1-4296a31252e2"),
'     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'    Public Interface IFileOperationProgressSink
'        ' We don’t need to implement methods for basic behavior; can be empty stub.
'    End Interface

'    <ComImport(), Guid("000214E6-0000-0000-C000-000000000046"),
'     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'    Public Interface IShellItem
'    End Interface

'    <ComImport(), Guid("3AD05575-8857-4850-9277-11B85BDB8E09")>
'    Public Class FileOperation
'    End Class

'    <ComImport(), Guid("3AD05575-8857-4850-9277-11B85BDB8E09"),
'    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'    Public Interface IFileOperation
'        Sub Advise(pfops As IFileOperationProgressSink, ByRef pdwCookie As UInteger)
'        Sub Unadvise(dwCookie As UInteger)

'        Sub SetOperationFlags(dwOperationFlags As FileOperationFlags)

'        Sub CopyItem(
'        psiItem As IShellItem,
'        psiDestinationFolder As IShellItem,
'        pszNewName As String,
'        pfopsItem As IFileOperationProgressSink)

'        Sub PerformOperations()
'    End Interface

'    Friend NotInheritable Class ShellInterop

'        Private Sub New()
'        End Sub

'        <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
'        Private Shared Sub SHCreateItemFromParsingName(
'            <MarshalAs(UnmanagedType.LPWStr)> pszPath As String,
'            pbc As IntPtr,
'            ByRef riid As Guid,
'            <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object)
'        End Sub

'        Public Shared Function CreateShellItemFromPath(path As String) As IShellItem
'            Dim iidShellItem As New Guid("000214E6-0000-0000-C000-000000000046")
'            Dim obj As Object = Nothing
'            SHCreateItemFromParsingName(path, IntPtr.Zero, iidShellItem, obj)
'            Return CType(obj, IShellItem)
'        End Function

'    End Class

'End Namespace



'Namespace Explorer.Interop.Shell
'    Public NotInheritable Class ShellInterop

'        'Public Class ShellInterop

'        Public Shared Function GetPathsFromDragDrop(data As IDataObject) As List(Of String)
'            Dim results As New List(Of String)

'            ' 1. Try normal FileDrop first
'            If data.GetDataPresent(DataFormats.FileDrop) Then
'                results.AddRange(CType(data.GetData(DataFormats.FileDrop), String()))
'                Return results
'            End If

'            ' 2. Try Shell IDList Array (Android devices, Libraries, This PC, etc.)
'            If data.GetDataPresent("Shell IDList Array") Then
'                Dim obj = data.GetData("Shell IDList Array")
'                Dim pidls = CType(obj, ShellIDListArray)

'                For Each pidl In pidls.Items
'                    Dim path = pidl.GetDisplayName(SIGDN.SIGDN_FILESYSPATH)

'                    If String.IsNullOrEmpty(path) Then
'                        ' Fallback: get the shell display name
'                        path = pidl.GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY)
'                    End If

'                    results.Add(path)
'                Next
'            End If

'            Return results
'        End Function
'        ' -------------------------------
'        '  PUBLIC API (UI-agnostic)
'        ' -------------------------------

'        Public Shared Function GetIconForPath(path As String, size As IconSize) As Icon
'            Return GetShellIcon(path, size)
'        End Function

'        Public Shared Function GetIconForPath(path As String, pixelSize As Integer) As Icon
'            ' Native sizes
'            If pixelSize <= 16 Then
'                Return GetShellIcon(path, IconSize.Small)
'            ElseIf pixelSize >= 32 Then
'                Return GetShellIcon(path, IconSize.Large)
'            End If

'            ' Scale from 32x32
'            Dim baseIcon = GetShellIcon(path, IconSize.Large)
'            If baseIcon Is Nothing Then Return Nothing

'            Return New Icon(baseIcon, pixelSize, pixelSize)
'        End Function


'        Public Shared Function GetIconForVirtualFolder(guidString As String, pixelSize As Integer) As Icon
'            Dim pidl As IntPtr = ParseDisplayName(guidString)
'            If pidl = IntPtr.Zero Then Return Nothing

'            Try
'                If pixelSize <= 16 Then
'                    Return GetIconForPIDL(pidl, IconSize.Small)
'                ElseIf pixelSize >= 32 Then
'                    Return GetIconForPIDL(pidl, IconSize.Large)
'                End If

'                Dim baseIcon = GetIconForPIDL(pidl, IconSize.Large)
'                If baseIcon Is Nothing Then Return Nothing

'                Return New Icon(baseIcon, pixelSize, pixelSize)
'            Finally
'                CoTaskMemFree(pidl)
'            End Try
'        End Function


'        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
'        Private Shared Function SHParseDisplayName(
'            pszName As String,
'            pbc As IntPtr,
'            ByRef ppidl As IntPtr,
'            sfgaoIn As UInteger,
'            ByRef psfgaoOut As UInteger
'        ) As Integer
'        End Function

'        Public Shared Function ParseDisplayName(path As String) As IntPtr
'            Dim pidl As IntPtr = IntPtr.Zero
'            Dim attrs As UInteger = 0

'            Dim hr = SHParseDisplayName(path, IntPtr.Zero, pidl, 0UI, attrs)

'            If hr <> 0 Then
'                Return IntPtr.Zero
'            End If

'            Return pidl
'        End Function

'        Public Shared Function GetRecycleBinIcon(size As IconSize) As Icon
'            Dim pidl As IntPtr = IntPtr.Zero
'            Dim hr = SHGetSpecialFolderLocation(IntPtr.Zero, CSIDL_BITBUCKET, pidl)

'            If hr <> 0 OrElse pidl = IntPtr.Zero Then
'                Return Nothing
'            End If

'            Try
'                Return GetIconForPIDL(pidl, size)
'            Finally
'                CoTaskMemFree(pidl)
'            End Try
'        End Function

'        Public Shared Function GetRecycleBinIcon(pixelSize As Integer) As Icon
'            ' Native sizes
'            If pixelSize <= 16 Then
'                Return GetRecycleBinIcon(IconSize.Small)
'            ElseIf pixelSize >= 32 Then
'                Return GetRecycleBinIcon(IconSize.Large)
'            End If

'            ' Scale from 32x32
'            Dim baseIcon = GetRecycleBinIcon(IconSize.Large)
'            If baseIcon Is Nothing Then Return Nothing

'            Return New Icon(baseIcon, pixelSize, pixelSize)
'        End Function


'        ' -------------------------------
'        '  INTERNAL HELPERS
'        ' -------------------------------

'        Private Shared Function GetShellIcon(path As String, size As IconSize) As Icon
'            Dim flags As UInteger = SHGFI_ICON Or SHGFI_USEFILEATTRIBUTES

'            If size = IconSize.Small Then
'                flags = flags Or SHGFI_SMALLICON
'            Else
'                flags = flags Or SHGFI_LARGEICON
'            End If

'            Dim shinfo As New SHFILEINFO()

'            Dim attrs As UInteger =
'                If(Directory.Exists(path), FILE_ATTRIBUTE_DIRECTORY, FILE_ATTRIBUTE_NORMAL)

'            Dim result = SHGetFileInfo(
'                path,
'                attrs,
'                shinfo,
'                CUInt(Marshal.SizeOf(shinfo)),
'                flags
'            )

'            If result = IntPtr.Zero Then Return Nothing

'            Dim rawIcon As Icon = Icon.FromHandle(shinfo.hIcon)
'            Dim daIcon As Icon = CType(rawIcon.Clone(), Icon)

'            DestroyIcon(shinfo.hIcon)

'            Return daIcon
'        End Function

'        Private Shared Function GetIconForPIDL(pidl As IntPtr, size As IconSize) As Icon
'            Dim flags As UInteger = SHGFI_ICON Or SHGFI_PIDL

'            If size = IconSize.Small Then
'                flags = flags Or SHGFI_SMALLICON
'            Else
'                flags = flags Or SHGFI_LARGEICON
'            End If

'            Dim shinfo As New SHFILEINFO()

'            Dim result = SHGetFileInfo(
'                pidl,
'                0UI,
'                shinfo,
'                CUInt(Marshal.SizeOf(shinfo)),
'                flags
'            )

'            If result = IntPtr.Zero Then Return Nothing

'            Dim rawIcon As Icon = Icon.FromHandle(shinfo.hIcon)
'            Dim daIcon As Icon = CType(rawIcon.Clone(), Icon)

'            DestroyIcon(shinfo.hIcon)

'            Return daIcon
'        End Function


'        ' -------------------------------
'        '  ICON SIZE ENUM
'        ' -------------------------------
'        Public Enum IconSize
'            Small
'            Large
'        End Enum


'        ' -------------------------------
'        '  SHGetFileInfo P/Invoke
'        ' -------------------------------

'        ' Path-based overload (existing)
'        <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
'        Private Shared Function SHGetFileInfo(
'            pszPath As String,
'            dwFileAttributes As UInteger,
'            ByRef psfi As SHFILEINFO,
'            cbFileInfo As UInteger,
'            uFlags As UInteger
'        ) As IntPtr
'        End Function

'        ' PIDL-based overload (for virtual folders like Recycle Bin)
'        <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
'        Private Shared Function SHGetFileInfo(
'            pidl As IntPtr,
'            dwFileAttributes As UInteger,
'            ByRef psfi As SHFILEINFO,
'            cbFileInfo As UInteger,
'            uFlags As UInteger
'        ) As IntPtr
'        End Function

'        <DllImport("user32.dll", SetLastError:=True)>
'        Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
'        End Function

'        <DllImport("shell32.dll")>
'        Private Shared Function SHGetSpecialFolderLocation(
'            hwndOwner As IntPtr,
'            nFolder As Integer,
'            ByRef ppidl As IntPtr
'        ) As Integer
'        End Function

'        <DllImport("ole32.dll")>
'        Private Shared Sub CoTaskMemFree(pv As IntPtr)
'        End Sub

'        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
'        Private Shared Function SHFileOperation(ByRef lpFileOp As SHFILEOPSTRUCT) As Integer
'        End Function

'        <DllImport("shell32.dll")>
'        Private Shared Function ILCombine(pidl1 As IntPtr, pidl2 As IntPtr) As IntPtr
'        End Function

'        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
'        Private Shared Function SHGetNameFromIDList(
'            pidl As IntPtr,
'            sigdnName As UInteger,
'            ByRef ppszName As IntPtr
'        ) As Integer
'        End Function


'        <DllImport("shell32.dll")>
'        Private Shared Function SHCreateItemFromIDList(pidl As IntPtr,
'    ByRef riid As Guid,
'    <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object) As Integer
'        End Function

'        Public Shared Function CreateShellItemFromPIDL(pidl As IntPtr) As IShellItem
'            Dim iid As New Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE") ' IShellItem
'            Dim obj As Object = Nothing
'            SHCreateItemFromIDList(pidl, iid, obj)
'            Return CType(obj, IShellItem)
'        End Function


'        Public Sub CopyShellItem(source As IShellItem, destinationFolder As IShellItem)
'            Dim fo As IFileOperation = CType(Activator.CreateInstance(
'        Type.GetTypeFromCLSID(New Guid("3AD05575-8857-4850-9277-11B85BDB8E09"))), IFileOperation)

'            fo.SetOperationFlags(
'        FileOperationFlags.FOF_NOCONFIRMMKDIR Or
'        FileOperationFlags.FOF_NOCONFIRMATION Or
'        FileOperationFlags.FOFX_NOCOPYSECURITYATTRIBS Or
'        FileOperationFlags.FOFX_SHOWELEVATIONPROMPT Or
'        FileOperationFlags.FOFX_REQUIREELEVATION
'    )

'            fo.CopyItem(source, destinationFolder, Nothing, Nothing)
'            fo.PerformOperations()
'        End Sub


'        <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
'        Private Shared Function SHCreateItemFromParsingName(
'    pszPath As String,
'    pbc As IntPtr,
'    ByRef riid As Guid,
'    <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object) As Integer
'        End Function

'        <Flags>
'        Public Enum FileOperationFlags As UInteger
'            FOF_NOCONFIRMATION = &H10
'            FOF_NOCONFIRMMKDIR = &H200
'            FOFX_NOCOPYSECURITYATTRIBS = &H800
'            FOFX_REQUIREELEVATION = &H200000
'        End Enum

'        Public Shared Function CreateShellItemFromPath(path As String) As IShellItem
'            Dim iid As New Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE") ' IShellItem
'            Dim obj As Object = Nothing
'            SHCreateItemFromParsingName(path, IntPtr.Zero, iid, obj)
'            Return CType(obj, IShellItem)
'        End Function

'        Public Enum SIGDN As UInteger
'            SIGDN_NORMALDISPLAY = &H0
'            SIGDN_PARENTRELATIVEPARSING = &H80018001UI
'            SIGDN_DESKTOPABSOLUTEPARSING = &H80028000UI
'            SIGDN_PARENTRELATIVEEDITING = &H80031001UI
'            SIGDN_DESKTOPABSOLUTEEDITING = &H8004C000UI
'            SIGDN_FILESYSPATH = &H80058000UI
'            SIGDN_URL = &H80068000UI
'        End Enum

'        Private Const SIGDN_NORMALDISPLAY As UInteger = &H0UI

'        Private Shared Function GetDisplayNameFromPidl(pidl As IntPtr) As String
'            Dim psz As IntPtr = IntPtr.Zero
'            Dim hr = SHGetNameFromIDList(pidl, SIGDN_NORMALDISPLAY, psz)
'            If hr <> 0 OrElse psz = IntPtr.Zero Then Return String.Empty
'            Try
'                Return Marshal.PtrToStringUni(psz)
'            Finally
'                CoTaskMemFree(psz)
'            End Try
'        End Function

'        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
'        Private Structure SHFILEOPSTRUCT
'            Public hwnd As IntPtr
'            Public wFunc As UInteger
'            Public pFrom As String
'            Public pTo As String
'            Public fFlags As UShort
'            Public fAnyOperationsAborted As Boolean
'            Public hNameMappings As IntPtr
'            Public lpszProgressTitle As String
'        End Structure

'        Private Const FO_DELETE As UInteger = &H3
'        Private Const FOF_ALLOWUNDO As UShort = &H40
'        Private Const FOF_NOCONFIRMATION As UShort = &H10
'        Private Const FOF_SILENT As UShort = &H4


'        Public Shared Function SendToRecycleBin(path As String) As Boolean
'            Dim op As New SHFILEOPSTRUCT With {
'        .wFunc = FO_DELETE,
'        .pFrom = path & vbNullChar & vbNullChar, ' Double-null terminated
'        .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_SILENT
'    }

'            Dim result = SHFileOperation(op)
'            Return (result = 0 AndAlso Not op.fAnyOperationsAborted)
'        End Function

'        ' -------------------------------
'        '  SHFILEINFO STRUCT
'        ' -------------------------------
'        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
'        Private Structure SHFILEINFO
'            Public hIcon As IntPtr
'            Public iIcon As Integer
'            Public dwAttributes As UInteger

'            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
'            Public szDisplayName As String

'            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
'            Public szTypeName As String
'        End Structure


'        ' -------------------------------
'        '  FLAGS & CONSTANTS
'        ' -------------------------------
'        Private Const SHGFI_ICON As UInteger = &H100
'        Private Const SHGFI_DISPLAYNAME As UInteger = &H200
'        Private Const SHGFI_TYPENAME As UInteger = &H400
'        Private Const SHGFI_ATTRIBUTES As UInteger = &H800
'        Private Const SHGFI_ICONLOCATION As UInteger = &H1000
'        Private Const SHGFI_EXETYPE As UInteger = &H2000
'        Private Const SHGFI_SYSICONINDEX As UInteger = &H4000
'        Private Const SHGFI_LINKOVERLAY As UInteger = &H8000
'        Private Const SHGFI_SELECTED As UInteger = &H10000
'        Private Const SHGFI_ATTR_SPECIFIED As UInteger = &H20000
'        Private Const SHGFI_LARGEICON As UInteger = &H0
'        Private Const SHGFI_SMALLICON As UInteger = &H1
'        Private Const SHGFI_USEFILEATTRIBUTES As UInteger = &H10
'        Private Const SHGFI_PIDL As UInteger = &H8

'        Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80
'        Private Const FILE_ATTRIBUTE_DIRECTORY As UInteger = &H10

'        ' Recycle Bin CSIDL
'        Private Const CSIDL_BITBUCKET As Integer = &HA

'    End Class

'End Namespace



'Imports System
'Imports System.IO
'Imports System.Runtime.InteropServices
'Imports System.Text
'Imports File_Explorer.Explorer.Interop.DeleteInterop

'Namespace Explorer.Interop.Shell

'    ' ============================================================
'    '   CORE SHELL ENUMS
'    ' ============================================================
'    Public Enum IconSize
'        Small
'        Large
'    End Enum

'    <Flags>
'    Public Enum FileOperationFlags As UInteger
'        FOF_NOCONFIRMATION = &H10UI
'        FOF_NOCONFIRMMKDIR = &H200UI
'        FOFX_NOCOPYSECURITYATTRIBS = &H800UI
'        FOFX_REQUIREELEVATION = &H200000UI
'        FOFX_SHOWELEVATIONPROMPT = &H40000UI
'    End Enum

'    Public Enum SIGDN As UInteger
'        SIGDN_NORMALDISPLAY = &H0UI
'        SIGDN_PARENTRELATIVEPARSING = &H80018001UI
'        SIGDN_DESKTOPABSOLUTEPARSING = &H80028000UI
'        SIGDN_PARENTRELATIVEEDITING = &H80031001UI
'        SIGDN_DESKTOPABSOLUTEEDITING = &H8004C000UI
'        SIGDN_FILESYSPATH = &H80058000UI
'        SIGDN_URL = &H80068000UI
'    End Enum


'    ' ============================================================
'    '   SHELL INTERFACES
'    ' ============================================================
'    <ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"),
'     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'    Public Interface IShellItem
'    End Interface

'    <ComImport, Guid("04B0F1A7-9490-44BC-96E1-4296A31252E2"),
'     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'    Public Interface IFileOperationProgressSink
'    End Interface

'    <ComImport, Guid("3AD05575-8857-4850-9277-11B85BDB8E09"),
'     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'    Public Interface IFileOperation
'        Sub Advise(pSink As IFileOperationProgressSink, ByRef cookie As UInteger)
'        Sub Unadvise(cookie As UInteger)
'        Sub SetOperationFlags(flags As FileOperationFlags)
'        Sub SetOwnerWindow(hwnd As IntPtr)

'        Sub CopyItem(
'            psiItem As IShellItem,
'            psiDestinationFolder As IShellItem,
'            <MarshalAs(UnmanagedType.LPWStr)> pszNewName As String,
'            pSink As IFileOperationProgressSink)

'        Sub DeleteItem(psiItem As IShellItem, pSink As IFileOperationProgressSink)

'        Sub PerformOperations()
'        Function GetAnyOperationsAborted() As Boolean
'    End Interface

'    <ComImport, Guid("3AD05575-8857-4850-9277-11B85BDB8E09")>
'    Public Class FileOperation
'    End Class


'    ' ============================================================
'    '   SHELL INTEROP CLASS
'    ' ============================================================
'    Public NotInheritable Class ShellInterop

'        Private Sub New()
'        End Sub


'        ' ------------------------------------------------------------
'        '   SHCreateItemFromParsingName (path → IShellItem)
'        ' ------------------------------------------------------------
'        <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
'        Private Shared Function SHCreateItemFromParsingName(
'            pszPath As String,
'            pbc As IntPtr,
'            ByRef riid As Guid,
'            <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object) As Integer
'        End Function

'        Public Shared Function CreateShellItemFromPath(path As String) As IShellItem
'            Dim iid As New Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")
'            Dim obj As Object = Nothing
'            SHCreateItemFromParsingName(path, IntPtr.Zero, iid, obj)
'            Return CType(obj, IShellItem)
'        End Function


'        ' ------------------------------------------------------------
'        '   SHCreateItemFromIDList (PIDL → IShellItem)
'        ' ------------------------------------------------------------
'        <DllImport("shell32.dll")>
'        Private Shared Function SHCreateItemFromIDList(
'            pidl As IntPtr,
'            ByRef riid As Guid,
'            <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object) As Integer
'        End Function

'        Public Shared Function CreateShellItemFromPIDL(pidl As IntPtr) As IShellItem
'            Dim iid As New Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")
'            Dim obj As Object = Nothing
'            SHCreateItemFromIDList(pidl, iid, obj)
'            Return CType(obj, IShellItem)
'        End Function
'        Public Shared Sub CopyShellItem(source As IShellItem, destinationFolder As IShellItem)

'            Dim srcPath As String = GetDisplayName(source, SIGDN.SIGDN_FILESYSPATH)
'            Dim destPath As String = GetDisplayName(destinationFolder, SIGDN.SIGDN_FILESYSPATH)

'            If String.IsNullOrEmpty(srcPath) OrElse String.IsNullOrEmpty(destPath) Then
'                Exit Sub
'            End If

'            Dim op As New SHFILEOPSTRUCT With {
'        .wFunc = FO_COPY,
'        .pFrom = srcPath & vbNullChar & vbNullChar,
'        .pTo = destPath & vbNullChar & vbNullChar,
'        .fFlags = FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
'    }

'            SHFileOperation(op)
'        End Sub

'        Private Shared Function GetDisplayName(item As IShellItem, sigdn As SIGDN) As String
'            Dim psz As IntPtr = IntPtr.Zero
'            Dim hr = SHGetNameFromIDList(item, sigdn, psz)
'            If hr <> 0 OrElse psz = IntPtr.Zero Then Return Nothing
'            Dim result = Marshal.PtrToStringUni(psz)
'            CoTaskMemFree(psz)
'            Return result
'        End Function

'        ' ------------------------------------------------------------
'        '   ICON EXTRACTION (PIDL + path)
'        ' ------------------------------------------------------------
'        <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
'        Private Shared Function SHGetFileInfo(
'            pszPath As String,
'            dwFileAttributes As UInteger,
'            ByRef psfi As SHFILEINFO,
'            cbFileInfo As UInteger,
'            uFlags As UInteger) As IntPtr
'        End Function

'        <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
'        Private Shared Function SHGetFileInfo(
'            pidl As IntPtr,
'            dwFileAttributes As UInteger,
'            ByRef psfi As SHFILEINFO,
'            cbFileInfo As UInteger,
'            uFlags As UInteger) As IntPtr
'        End Function

'        <DllImport("user32.dll")>
'        Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
'        End Function


'        ' ------------------------------------------------------------
'        '  PUBLIC ICON API (Explorer-accurate)
'        ' ------------------------------------------------------------
'        Public Shared Function GetIconForPath(path As String, size As IconSize) As Icon
'            Return GetShellIcon(path, size)
'        End Function

'        Public Shared Function GetIconForVirtualFolder(guidString As String, pixelSize As Integer) As Icon
'            Dim pidl As IntPtr = ParseDisplayName(guidString)
'            If pidl = IntPtr.Zero Then Return Nothing

'            Try
'                If pixelSize <= 16 Then
'                    Return GetIconForPIDL(pidl, IconSize.Small)
'                ElseIf pixelSize >= 32 Then
'                    Return GetIconForPIDL(pidl, IconSize.Large)
'                End If

'                Dim baseIcon = GetIconForPIDL(pidl, IconSize.Large)
'                If baseIcon Is Nothing Then Return Nothing

'                Return New Icon(baseIcon, pixelSize, pixelSize)
'            Finally
'                CoTaskMemFree(pidl)
'            End Try
'        End Function

'        Public Shared Function GetRecycleBinIcon(size As IconSize) As Icon
'            Dim pidl As IntPtr = IntPtr.Zero
'            Dim hr = SHGetSpecialFolderLocation(IntPtr.Zero, CSIDL_BITBUCKET, pidl)

'            If hr <> 0 OrElse pidl = IntPtr.Zero Then Return Nothing

'            Try
'                Return GetIconForPIDL(pidl, size)
'            Finally
'                CoTaskMemFree(pidl)
'            End Try
'        End Function


'        ' ------------------------------------------------------------
'        '   INTERNAL HELPERS FOR ICONS & PIDLS
'        ' ------------------------------------------------------------

'        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
'        Private Shared Function SHParseDisplayName(
'    pszName As String,
'    pbc As IntPtr,
'    ByRef ppidl As IntPtr,
'    sfgaoIn As UInteger,
'    ByRef psfgaoOut As UInteger) As Integer
'        End Function

'        Public Shared Function ParseDisplayName(path As String) As IntPtr
'            Dim pidl As IntPtr = IntPtr.Zero
'            Dim attrs As UInteger = 0UI
'            Dim hr = SHParseDisplayName(path, IntPtr.Zero, pidl, 0UI, attrs)
'            If hr <> 0 Then Return IntPtr.Zero
'            Return pidl
'        End Function


'        ' ------------------------------------------------------------
'        '   MEMORY + SPECIAL FOLDER HELPERS
'        ' ------------------------------------------------------------
'        <DllImport("shell32.dll")>
'        Private Shared Function SHGetSpecialFolderLocation(
'    hwndOwner As IntPtr,
'    nFolder As Integer,
'    ByRef ppidl As IntPtr) As Integer
'        End Function

'        <DllImport("ole32.dll")>
'        Private Shared Sub CoTaskMemFree(pv As IntPtr)
'        End Sub

'        Private Const CSIDL_BITBUCKET As Integer = &HA


'        ' ------------------------------------------------------------
'        '   ICON HELPERS (path + PIDL)
'        ' ------------------------------------------------------------
'        Private Shared Function GetShellIcon(path As String, size As IconSize) As Icon
'            Dim flags As UInteger = SHGFI_ICON Or SHGFI_USEFILEATTRIBUTES

'            If size = IconSize.Small Then
'                flags = flags Or SHGFI_SMALLICON
'            Else
'                flags = flags Or SHGFI_LARGEICON
'            End If

'            Dim shinfo As New SHFILEINFO()
'            Dim attrs As UInteger =
'        If(Directory.Exists(path), FILE_ATTRIBUTE_DIRECTORY, FILE_ATTRIBUTE_NORMAL)

'            Dim result = SHGetFileInfo(
'        path,
'        attrs,
'        shinfo,
'        CUInt(Marshal.SizeOf(shinfo)),
'        flags)

'            If result = IntPtr.Zero Then Return Nothing

'            Dim rawIcon As Icon = Icon.FromHandle(shinfo.hIcon)
'            Dim cloned As Icon = CType(rawIcon.Clone(), Icon)
'            DestroyIcon(shinfo.hIcon)
'            Return cloned
'        End Function


'        Private Shared Function GetIconForPIDL(pidl As IntPtr, size As IconSize) As Icon
'            Dim flags As UInteger = SHGFI_ICON Or SHGFI_PIDL

'            If size = IconSize.Small Then
'                flags = flags Or SHGFI_SMALLICON
'            Else
'                flags = flags Or SHGFI_LARGEICON
'            End If

'            Dim shinfo As New SHFILEINFO()

'            Dim result = SHGetFileInfo(
'        pidl,
'        0UI,
'        shinfo,
'        CUInt(Marshal.SizeOf(shinfo)),
'        flags)

'            If result = IntPtr.Zero Then Return Nothing

'            Dim rawIcon As Icon = Icon.FromHandle(shinfo.hIcon)
'            Dim cloned As Icon = CType(rawIcon.Clone(), Icon)
'            DestroyIcon(shinfo.hIcon)
'            Return cloned
'        End Function


'        ' ============================================================
'        '   SHFILEOPERATION (Copy/Move/Delete) SUPPORT
'        ' ============================================================

'        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
'        Private Structure SHFILEOPSTRUCT
'            Public hwnd As IntPtr
'            Public wFunc As UInteger
'            Public pFrom As String
'            Public pTo As String
'            Public fFlags As UShort
'            Public fAnyOperationsAborted As Boolean
'            Public hNameMappings As IntPtr
'            Public lpszProgressTitle As String
'        End Structure

'        ' --- Operation types ---
'        Private Const FO_COPY As UInteger = &H2UI
'        Private Const FO_MOVE As UInteger = &H1UI
'        Private Const FO_DELETE As UInteger = &H3UI
'        Private Const FO_RENAME As UInteger = &H4UI

'        ' --- Flags ---
'        Private Const FOF_SILENT As UShort = &H4US
'        Private Const FOF_NOCONFIRMATION As UShort = &H10US
'        Private Const FOF_NOCONFIRMMKDIR As UShort = &H200US
'        Private Const FOF_ALLOWUNDO As UShort = &H40US

'        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
'        Private Shared Function SHFileOperation(ByRef lpFileOp As SHFILEOPSTRUCT) As Integer
'        End Function


'        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
'        Private Shared Function SHGetNameFromIDList(
'    pidl As IntPtr,
'    sigdnName As SIGDN,
'    ByRef ppszName As IntPtr) As Integer
'        End Function


'        Public Shared Function GetDisplayNameFromPIDL(pidl As IntPtr, sigdn As SIGDN) As String
'            Dim psz As IntPtr = IntPtr.Zero
'            Dim hr = SHGetNameFromIDList(pidl, sigdn, psz)
'            If hr <> 0 OrElse psz = IntPtr.Zero Then Return Nothing

'            Dim result = Marshal.PtrToStringUni(psz)
'            CoTaskMemFree(psz)
'            Return result
'        End Function


'        ' ------------------------------------------------------------
'        '   STRUCTS & CONSTANTS
'        ' ------------------------------------------------------------
'        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
'        Private Structure SHFILEINFO
'            Public hIcon As IntPtr
'            Public iIcon As Integer
'            Public dwAttributes As UInteger
'            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
'            Public szDisplayName As String
'            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
'            Public szTypeName As String
'        End Structure

'        Private Const SHGFI_ICON As UInteger = &H100UI
'        Private Const SHGFI_PIDL As UInteger = &H8UI
'        Private Const SHGFI_SMALLICON As UInteger = &H1UI
'        Private Const SHGFI_LARGEICON As UInteger = &H0UI
'        Private Const SHGFI_USEFILEATTRIBUTES As UInteger = &H10UI

'        Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80UI
'        Private Const FILE_ATTRIBUTE_DIRECTORY As UInteger = &H10UI

'    End Class

'End Namespace


Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Drawing

Namespace Explorer.Interop.Shell

    ' ============================================================
    '   CORE ENUMS
    ' ============================================================
    Public Enum IconSize
        Small
        Large
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


        Public Shared Function GetIconForPath(path As String, size As IconSize) As Icon
            Return GetShellIcon(path, size)
        End Function

        Public Shared Function GetIconForVirtualFolder(guidString As String, pixelSize As Integer) As Icon
            Dim pidl As IntPtr = ParseDisplayName(guidString)
            If pidl = IntPtr.Zero Then Return Nothing

            Try
                If pixelSize <= 16 Then
                    Return GetIconForPIDL(pidl, IconSize.Small)
                ElseIf pixelSize >= 32 Then
                    Return GetIconForPIDL(pidl, IconSize.Large)
                End If

                Dim baseIcon = GetIconForPIDL(pidl, IconSize.Large)
                If baseIcon Is Nothing Then Return Nothing

                Return New Icon(baseIcon, pixelSize, pixelSize)
            Finally
                CoTaskMemFree(pidl)
            End Try
        End Function

        Public Shared Function GetRecycleBinIcon(size As IconSize) As Icon
            Dim pidl As IntPtr = IntPtr.Zero
            Dim hr = SHGetSpecialFolderLocation(IntPtr.Zero, CSIDL_BITBUCKET, pidl)

            If hr <> 0 OrElse pidl = IntPtr.Zero Then Return Nothing

            Try
                Return GetIconForPIDL(pidl, size)
            Finally
                CoTaskMemFree(pidl)
            End Try
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
            Dim flags As UInteger = SHGFI_ICON Or SHGFI_USEFILEATTRIBUTES

            If size = IconSize.Small Then
                flags = flags Or SHGFI_SMALLICON
            Else
                flags = flags Or SHGFI_LARGEICON
            End If

            Dim shinfo As New SHFILEINFO()
            Dim attrs As UInteger =
                If(Directory.Exists(path), FILE_ATTRIBUTE_DIRECTORY, FILE_ATTRIBUTE_NORMAL)

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