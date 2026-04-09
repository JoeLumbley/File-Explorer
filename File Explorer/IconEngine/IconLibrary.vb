Imports System.Drawing
Imports File_Explorer.Explorer.Interop.Shell

' Central place for generic icons used as fallbacks/placeholders.
Public Module IconLibrary

    Public ReadOnly Property GenericFile(iconSize As Integer) As Icon
        Get
            Return ShellInterop.GetIconForPath("*.unknown", iconSize)
        End Get
    End Property

    Public ReadOnly Property GenericFolder(iconSize As Integer) As Icon
        Get
            Return ShellInterop.GetIconForPath("C:\Windows\System", iconSize)
        End Get
    End Property

End Module