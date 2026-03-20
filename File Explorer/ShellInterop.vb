'Imports System
'Imports System.IO
'Imports System.Drawing
'Imports System.Diagnostics
'Imports System.Runtime.InteropServices
'Imports Microsoft.VisualBasic

'Imports System.Text
'Imports System.ComponentModel
'Imports System.Runtime.Versioning
'Imports System.Security

'Namespace Explorer.Interop.Shell

'    Public NotInheritable Class ShellInterop

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


'        Public Shared Function GetRecycleBinIcon(size As IconSize) As Icon
'            Dim pidl As IntPtr = IntPtr.Zero
'            SHGetSpecialFolderLocation(IntPtr.Zero, CSIDL_BITBUCKET, pidl)
'            Return GetIconForPIDL(pidl, size)
'        End Function

'        Private Shared Function GetShellIcon(path As String, size As IconSize) As Icon
'            Dim flags As UInteger = SHGFI_ICON Or SHGFI_USEFILEATTRIBUTES

'            If size = IconSize.Small Then
'                flags = flags Or SHGFI_SMALLICON
'            Else
'                flags = flags Or SHGFI_LARGEICON
'            End If

'            Dim shinfo As New SHFILEINFO()

'            Dim attrs As UInteger =
'            If(Directory.Exists(path), FILE_ATTRIBUTE_DIRECTORY, FILE_ATTRIBUTE_NORMAL)

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
'        <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
'        Private Shared Function SHGetFileInfo(
'            pszPath As String,
'            dwFileAttributes As UInteger,
'            ByRef psfi As SHFILEINFO,
'            cbFileInfo As UInteger,
'            uFlags As UInteger
'        ) As IntPtr
'        End Function

'        <DllImport("user32.dll", SetLastError:=True)>
'        Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
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

'        Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80
'        Private Const FILE_ATTRIBUTE_DIRECTORY As UInteger = &H10

'    End Class

'End Namespace


Imports System
Imports System.IO
Imports System.Drawing
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic

Imports System.Text
Imports System.ComponentModel
Imports System.Runtime.Versioning
Imports System.Security


'Imports System.Runtime.InteropServices

Namespace Explorer.Interop

    <Flags>
    Public Enum FileOperationFlags As UInteger
        FOF_SILENT = &H4UI
        FOF_NOCONFIRMATION = &H10UI
        FOF_ALLOWUNDO = &H40UI
        FOF_NOCONFIRMMKDIR = &H200UI
        FOF_NOERRORUI = &H400UI

        ' Extended flags (FOFX)
        FOFX_NOCONFIRMMKDIR = &H200UI
        FOFX_SHOWELEVATIONPROMPT = &H40000UI
        FOFX_RECYCLEONDELETE = &H80000UI
    End Enum

    <ComImport(), Guid("3AD05575-8857-4850-9277-11B85BDB8E09"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IFileOperation
        Sub Advise(pfops As IFileOperationProgressSink, ByRef pdwCookie As UInteger)
        Sub Unadvise(dwcookie As UInteger)
        Sub SetOperationFlags(dwOperationFlags As FileOperationFlags)
        Sub SetOwnerWindow(hwndOwner As IntPtr)

        ' We only need DeleteItem + PerformOperations for this subsystem
        Sub DeleteItem(psiItem As IShellItem, pfopsItem As IFileOperationProgressSink)
        Sub PerformOperations()
        Function GetAnyOperationsAborted() As Boolean
    End Interface

    <ComImport(), Guid("04b0f1a7-9490-44bc-96e1-4296a31252e2"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IFileOperationProgressSink
        ' We don’t need to implement methods for basic behavior; can be empty stub.
    End Interface

    <ComImport(), Guid("000214E6-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IShellItem
    End Interface

    <ComImport(), Guid("3AD05575-8857-4850-9277-11B85BDB8E09")>
    Public Class FileOperation
    End Class

    Friend NotInheritable Class ShellInterop

        Private Sub New()
        End Sub

        <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
        Private Shared Sub SHCreateItemFromParsingName(
            <MarshalAs(UnmanagedType.LPWStr)> pszPath As String,
            pbc As IntPtr,
            ByRef riid As Guid,
            <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object)
        End Sub

        Public Shared Function CreateShellItemFromPath(path As String) As IShellItem
            Dim iidShellItem As New Guid("000214E6-0000-0000-C000-000000000046")
            Dim obj As Object = Nothing
            SHCreateItemFromParsingName(path, IntPtr.Zero, iidShellItem, obj)
            Return CType(obj, IShellItem)
        End Function

    End Class

End Namespace



Namespace Explorer.Interop.Shell

    Public NotInheritable Class ShellInterop

        ' -------------------------------
        '  PUBLIC API (UI-agnostic)
        ' -------------------------------

        Public Shared Function GetIconForPath(path As String, size As IconSize) As Icon
            Return GetShellIcon(path, size)
        End Function

        Public Shared Function GetIconForPath(path As String, pixelSize As Integer) As Icon
            ' Native sizes
            If pixelSize <= 16 Then
                Return GetShellIcon(path, IconSize.Small)
            ElseIf pixelSize >= 32 Then
                Return GetShellIcon(path, IconSize.Large)
            End If

            ' Scale from 32x32
            Dim baseIcon = GetShellIcon(path, IconSize.Large)
            If baseIcon Is Nothing Then Return Nothing

            Return New Icon(baseIcon, pixelSize, pixelSize)
        End Function

        Public Shared Function GetRecycleBinIcon(size As IconSize) As Icon
            Dim pidl As IntPtr = IntPtr.Zero
            Dim hr = SHGetSpecialFolderLocation(IntPtr.Zero, CSIDL_BITBUCKET, pidl)

            If hr <> 0 OrElse pidl = IntPtr.Zero Then
                Return Nothing
            End If

            Try
                Return GetIconForPIDL(pidl, size)
            Finally
                CoTaskMemFree(pidl)
            End Try
        End Function

        Public Shared Function GetRecycleBinIcon(pixelSize As Integer) As Icon
            ' Native sizes
            If pixelSize <= 16 Then
                Return GetRecycleBinIcon(IconSize.Small)
            ElseIf pixelSize >= 32 Then
                Return GetRecycleBinIcon(IconSize.Large)
            End If

            ' Scale from 32x32
            Dim baseIcon = GetRecycleBinIcon(IconSize.Large)
            If baseIcon Is Nothing Then Return Nothing

            Return New Icon(baseIcon, pixelSize, pixelSize)
        End Function


        ' -------------------------------
        '  INTERNAL HELPERS
        ' -------------------------------

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
                flags
            )

            If result = IntPtr.Zero Then Return Nothing

            Dim rawIcon As Icon = Icon.FromHandle(shinfo.hIcon)
            Dim daIcon As Icon = CType(rawIcon.Clone(), Icon)

            DestroyIcon(shinfo.hIcon)

            Return daIcon
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
                flags
            )

            If result = IntPtr.Zero Then Return Nothing

            Dim rawIcon As Icon = Icon.FromHandle(shinfo.hIcon)
            Dim daIcon As Icon = CType(rawIcon.Clone(), Icon)

            DestroyIcon(shinfo.hIcon)

            Return daIcon
        End Function


        ' -------------------------------
        '  ICON SIZE ENUM
        ' -------------------------------
        Public Enum IconSize
            Small
            Large
        End Enum


        ' -------------------------------
        '  SHGetFileInfo P/Invoke
        ' -------------------------------

        ' Path-based overload (existing)
        <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function SHGetFileInfo(
            pszPath As String,
            dwFileAttributes As UInteger,
            ByRef psfi As SHFILEINFO,
            cbFileInfo As UInteger,
            uFlags As UInteger
        ) As IntPtr
        End Function

        ' PIDL-based overload (for virtual folders like Recycle Bin)
        <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function SHGetFileInfo(
            pidl As IntPtr,
            dwFileAttributes As UInteger,
            ByRef psfi As SHFILEINFO,
            cbFileInfo As UInteger,
            uFlags As UInteger
        ) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
        End Function

        <DllImport("shell32.dll")>
        Private Shared Function SHGetSpecialFolderLocation(
            hwndOwner As IntPtr,
            nFolder As Integer,
            ByRef ppidl As IntPtr
        ) As Integer
        End Function

        <DllImport("ole32.dll")>
        Private Shared Sub CoTaskMemFree(pv As IntPtr)
        End Sub




        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function SHFileOperation(ByRef lpFileOp As SHFILEOPSTRUCT) As Integer
        End Function

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

        Private Const FO_DELETE As UInteger = &H3
        Private Const FOF_ALLOWUNDO As UShort = &H40
        Private Const FOF_NOCONFIRMATION As UShort = &H10
        Private Const FOF_SILENT As UShort = &H4


        Public Shared Function SendToRecycleBin(path As String) As Boolean
            Dim op As New SHFILEOPSTRUCT With {
        .wFunc = FO_DELETE,
        .pFrom = path & vbNullChar & vbNullChar, ' Double-null terminated
        .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_SILENT
    }

            Dim result = SHFileOperation(op)
            Return (result = 0 AndAlso Not op.fAnyOperationsAborted)
        End Function




        ' -------------------------------
        '  SHFILEINFO STRUCT
        ' -------------------------------
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


        ' -------------------------------
        '  FLAGS & CONSTANTS
        ' -------------------------------
        Private Const SHGFI_ICON As UInteger = &H100
        Private Const SHGFI_DISPLAYNAME As UInteger = &H200
        Private Const SHGFI_TYPENAME As UInteger = &H400
        Private Const SHGFI_ATTRIBUTES As UInteger = &H800
        Private Const SHGFI_ICONLOCATION As UInteger = &H1000
        Private Const SHGFI_EXETYPE As UInteger = &H2000
        Private Const SHGFI_SYSICONINDEX As UInteger = &H4000
        Private Const SHGFI_LINKOVERLAY As UInteger = &H8000
        Private Const SHGFI_SELECTED As UInteger = &H10000
        Private Const SHGFI_ATTR_SPECIFIED As UInteger = &H20000
        Private Const SHGFI_LARGEICON As UInteger = &H0
        Private Const SHGFI_SMALLICON As UInteger = &H1
        Private Const SHGFI_USEFILEATTRIBUTES As UInteger = &H10
        Private Const SHGFI_PIDL As UInteger = &H8

        Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80
        Private Const FILE_ATTRIBUTE_DIRECTORY As UInteger = &H10

        ' Recycle Bin CSIDL
        Private Const CSIDL_BITBUCKET As Integer = &HA

    End Class

End Namespace



