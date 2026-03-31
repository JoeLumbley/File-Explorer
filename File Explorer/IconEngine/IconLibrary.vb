Imports System.Drawing
Imports File_Explorer.Explorer.Interop.Shell

' Central place for generic icons used as fallbacks/placeholders.
Public Module IconLibrary

    ' TODO: Replace these with your real icons (resources, imagelist, etc.).
    Public ReadOnly Property GenericFile As Icon
        Get
            Return SystemIcons.WinLogo
        End Get
    End Property

    'Public ReadOnly Property GenericFolder As Icon
    '    Get
    '        Return SystemIcons.
    '    End Get
    'End Property
    Public ReadOnly Property GenericFolder(IconSize As Integer) As Icon
        Get
            'Return ShellInterop.GetIconForPath("C:\Windows\System32", 32)

            'Dim IconSize As Integer =  = Form.GetScaledIconSize(Form)
            Return ShellInterop.GetIconForPath("C:\Windows\System32", IconSize)
        End Get
    End Property
    Public ReadOnly Property GenericVirtualFolder As Icon
        Get
            Return SystemIcons.WinLogo
        End Get
    End Property

    Public ReadOnly Property ShortcutOverlay As Icon
        Get
            Return SystemIcons.WinLogo
        End Get
    End Property

    Public ReadOnly Property SharedOverlay As Icon
        Get
            Return SystemIcons.WinLogo
        End Get
    End Property

End Module