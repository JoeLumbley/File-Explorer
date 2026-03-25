Namespace Explorer.Navigation

    Public NotInheritable Class ShellNavigation

        Public Shared Sub OpenVirtualFolder(vf As VirtualFolder)
            ' Opening virtual folders like the Recycle Bin can be tricky
            ' because they don't have a standard file system path. However,
            ' Windows provides special shell commands that can be used to
            ' access these locations. The "shell:RecycleBinFolder" command is a
            ' special URI that tells Windows to open the Recycle Bin in Windows
            ' Explorer.

            Try
                Dim uri = VirtualFolderMap.GetUri(vf)
                Process.Start("explorer.exe", uri)
            Catch ex As Exception
                ' Surface this to your UI layer
                Throw New InvalidOperationException(
                    $"Failed to open virtual folder '{vf}': {ex.Message}", ex)
            End Try

        End Sub

    End Class

End Namespace