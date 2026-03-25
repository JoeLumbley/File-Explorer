Namespace Explorer.Navigation

    Public NotInheritable Class VirtualFolderMap

        Private Shared ReadOnly FolderUris As New Dictionary(Of VirtualFolder, String) From {
            {VirtualFolder.RecycleBin, "shell:RecycleBinFolder"},
            {VirtualFolder.Downloads, "shell:Downloads"},
            {VirtualFolder.ThisPC, "shell:MyComputerFolder"},
            {VirtualFolder.Network, "shell:NetworkPlacesFolder"},
            {VirtualFolder.ControlPanel, "shell:ControlPanelFolder"},
            {VirtualFolder.Desktop, "shell:Desktop"}
        }

        Public Shared Function GetUri(vf As VirtualFolder) As String
            Return FolderUris(vf)
        End Function

        'Public Shared Function ResolveCanonical(shellPath As String) As VirtualFolder?
        '    Dim lower = shellPath.ToLowerInvariant()

        '    For Each kvp In FolderUris
        '        If lower = kvp.Value.ToLowerInvariant() Then
        '            Return kvp.Key
        '        End If
        '    Next

        '    Return Nothing
        'End Function

        'Public Shared Function ResolveCanonical(shellPath As String) As VirtualFolder?
        '    Dim lower = shellPath.ToLowerInvariant()

        '    Select Case lower
        '        Case "shell:controlpanel", "shell:controlpanelfolder"
        '            Return VirtualFolder.ControlPanel
        '        Case "shell:recyclebin", "shell:recyclebinfolder"
        '            Return VirtualFolder.RecycleBin
        '        Case "shell:downloads", "shell:downloadsfolder"
        '            Return VirtualFolder.Downloads
        '        Case "shell:computer", "shell:mycomputer", "shell:mycomputerfolder", "shell:pc", "shell:thispc"
        '            Return VirtualFolder.ThisPC
        '        Case "shell:network", "shell:networkplaces", "shell:networkplacesfolder"
        '            Return VirtualFolder.Network
        '        Case "shell:desktop"
        '            Return VirtualFolder.Desktop
        '    End Select

        '    Return Nothing
        'End Function

        'Public Shared Function ResolveCanonical(shellPath As String) As VirtualFolder?
        '    Dim lower = shellPath.TrimEnd("\"c).ToLowerInvariant()

        '    Select Case lower
        '        Case "shell:controlpanel", "shell:control panel", "shell:controlpanelfolder"
        '            Return VirtualFolder.ControlPanel

        '        Case "shell:recycle", "shell:recyclebin", "shell:recyclebinfolder"
        '            Return VirtualFolder.RecycleBin

        '        Case "shell:downloads", "shell:downloadsfolder"
        '            Return VirtualFolder.Downloads

        '        Case "shell:computer", "shell:mycomputer", "shell:mycomputerfolder",
        '             "shell:pc", "shell:thispc"
        '            Return VirtualFolder.ThisPC

        '        Case "shell:network", "shell:networkplaces", "shell:networkplacesfolder"
        '            Return VirtualFolder.Network

        '        Case "shell:desktop"
        '            Return VirtualFolder.Desktop
        '    End Select

        '    Return Nothing
        'End Function

        'Public Shared Function ResolveCanonical(shellPath As String) As VirtualFolder?
        '    Dim lower = shellPath.TrimEnd("\"c).ToLowerInvariant()

        '    Select Case lower
        '        Case "shell:controlpanel", "shell:control panel", "shell:controlpanelfolder"
        '            Return VirtualFolder.ControlPanel

        '        Case "shell:recycle", "shell:recyclebin", "shell:recycle bin", "shell:recyclebinfolder"
        '            Return VirtualFolder.RecycleBin

        '        Case "shell:downloads", "shell:downloadsfolder"
        '            Return VirtualFolder.Downloads

        '        Case "shell:computer", "shell:mycomputer", "shell:mycomputerfolder",
        '             "shell:pc", "shell:thispc"
        '            Return VirtualFolder.ThisPC

        '        Case "shell:network", "shell:networkplaces", "shell:networkplacesfolder"
        '            Return VirtualFolder.Network

        '        Case "shell:desktop", "shell:desktopfolder"
        '            Return VirtualFolder.Desktop
        '    End Select

        '    Return Nothing
        'End Function


        Public Shared Function ResolveCanonical(shellPath As String) As VirtualFolder?
            Dim lower = shellPath.TrimEnd("\"c).ToLowerInvariant()

            Select Case lower
                Case "shell:controlpanel", "shell:control panel",
                     "shell:controlpanelfolder", "shell:control panel folder"
                    Return VirtualFolder.ControlPanel

                Case "shell:recycle",
                     "shell:recyclebin", "shell:recycle bin",
                     "shell:recyclebinfolder", "shell:recycle bin folder"
                    Return VirtualFolder.RecycleBin

                Case "shell:downloads",
                     "shell:downloadsfolder", "shell:downloads folder"
                    Return VirtualFolder.Downloads

                Case "shell:computer",
                     "shell:mycomputer", "shell:my computer",
                     "shell:mycomputerfolder", "shell:my computer folder",
                     "shell:pc",
                     "shell:thispc", "shell:this pc"
                    Return VirtualFolder.ThisPC

                Case "shell:network",
                     "shell:networkfolder", "shell:network folder",
                     "shell:networkplaces", "shell:network places",
                     "shell:networkplacesfolder", "shell:network places folder"
                    Return VirtualFolder.Network

                Case "shell:desktop",
                     "shell:desktopfolder", "shell:desktop folder"
                    Return VirtualFolder.Desktop
            End Select

            Return Nothing
        End Function




    End Class

End Namespace