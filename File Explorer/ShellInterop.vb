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
        <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function SHGetFileInfo(
            pszPath As String,
            dwFileAttributes As UInteger,
            ByRef psfi As SHFILEINFO,
            cbFileInfo As UInteger,
            uFlags As UInteger
        ) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
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

        Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80
        Private Const FILE_ATTRIBUTE_DIRECTORY As UInteger = &H10

    End Class

End Namespace