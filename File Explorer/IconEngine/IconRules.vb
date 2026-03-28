Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports File_Explorer.Explorer.Interop.Shell

' A single rule: if Match(request) is True, use Handler(request) to get the icon.
' This keeps the decision logic declarative and easy to extend.
Public Class IconRule
    Public Property Match As Func(Of IconRequest, Boolean)
    Public Property Handler As Func(Of IconRequest, Icon)
End Class

' Declarative rule table for icon selection.
' This is where you encode “Explorer-like” behavior in a readable way.
Public Module IconRules

    ' Ordered list of rules; first match wins.
    Private ReadOnly _rules As List(Of IconRule)

    ' Static constructor: initialize rules once.
    'Sub New()
    '    _rules = New List(Of IconRule) From {

    '        ' Virtual Recycle Bin.
    '        New IconRule With {
    '            .Match = Function(r) r IsNot Nothing AndAlso r.IsVirtual AndAlso
    '                                   String.Equals(r.VirtualName, "RecycleBin",
    '                                                 StringComparison.OrdinalIgnoreCase),
    '            .Handler = AddressOf IconRuleHandlers.GetRecycleBinIcon
    '        },

    '        ' Any folder (real or virtual).
    '        New IconRule With {
    '            .Match = Function(r) r IsNot Nothing AndAlso r.IsFolder.GetValueOrDefault(),
    '            .Handler = AddressOf IconRuleHandlers.GetFolderIcon
    '        },

    '        ' Any file.
    '        New IconRule With {
    '            .Match = Function(r) r IsNot Nothing AndAlso Not r.IsFolder.GetValueOrDefault(),
    '            .Handler = AddressOf IconRuleHandlers.GetFileIcon
    '        }
    '    }
    'End Sub

    ' Returns the first matching handler for the given request.
    Public Function ResolveHandler(request As IconRequest) As Func(Of IconRequest, Icon)
        For Each rule In _rules
            If rule.Match(request) Then
                Return rule.Handler
            End If
        Next

        ' Fallback handler if nothing matches.
        Return Function(r) IconLibrary.GenericFile
    End Function

End Module

' Concrete handlers that call into ShellInterop or IconLibrary.
' These are the only places that know about ShellInterop.
Public Module IconRuleHandlers

    ' Recycle Bin icon via Shell.
    Public Function GetRecycleBinIcon(request As IconRequest) As Icon
        ' NOTE: using ShellInterop here because your file is ShellInterop.vb
        Return ShellInterop.GetRecycleBinIcon(request.PixelSize)
    End Function

    ' Folder icon (virtual or real).
    'Public Function GetFolderIcon(request As IconRequest) As Icon
    '    If request Is Nothing Then Return IconLibrary.GenericFolder

    '    ' Virtual folders can have their own icon; for now, generic.
    '    If request.IsVirtual Then
    '        Return IconLibrary.GenericVirtualFolder
    '    End If

    '    ' Real folder via Shell.
    '    Dim icon = ShellInterop.GetFolderIcon(request.FullPath, request.PixelSize)
    '    If icon Is Nothing Then
    '        icon = IconLibrary.GenericFolder
    '    End If

    '    ' Apply overlay if requested.
    '    If request.Overlay <> IconOverlayKind.None Then
    '        icon = IconOverlay.ApplyOverlay(icon, request.Overlay)
    '    End If

    '    Return icon
    'End Function

    ' File icon (by PIDL or path).
    'Public Function GetFileIcon(request As IconRequest) As Icon
    '    If request Is Nothing Then Return IconLibrary.GenericFile

    '    Dim icon As Icon = Nothing

    '    ' Prefer PIDL if available.
    '    If request.Pidl <> IntPtr.Zero Then
    '        icon = ShellInterop.GetIconForPidl(request.Pidl, request.PixelSize)
    '    ElseIf Not String.IsNullOrEmpty(request.FullPath) Then
    '        icon = ShellInterop.GetIconForPath(request.FullPath, request.PixelSize)
    '    End If

    '    If icon Is Nothing Then
    '        icon = IconLibrary.GenericFile
    '    End If

    '    ' Apply overlay if requested.
    '    If request.Overlay <> IconOverlayKind.None Then
    '        icon = IconOverlay.ApplyOverlay(icon, request.Overlay)
    '    End If

    '    Return icon
    'End Function

End Module