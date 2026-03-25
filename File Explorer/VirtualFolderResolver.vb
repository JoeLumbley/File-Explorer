Namespace Explorer.Navigation

    Public NotInheritable Class VirtualFolderResolver

        Private Shared ReadOnly GuidMap As New Dictionary(Of String, VirtualFolder)(StringComparer.OrdinalIgnoreCase) From {
            {"645FF040-5081-101B-9F08-00AA002F954E", VirtualFolder.RecycleBin},
            {"374DE290-123F-4565-9164-39C4925E467B", VirtualFolder.Downloads},
            {"20D04FE0-3AEA-1069-A2D8-08002B30309D", VirtualFolder.ThisPC},
            {"208D2C60-3AEA-1069-A2D7-08002B30309D", VirtualFolder.Network},
            {"5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0", VirtualFolder.ControlPanel},
            {"B4BFCC3A-DB2C-424C-B029-7FE99A87C641", VirtualFolder.Desktop}
        }

        Public Shared Function Resolve(shellPath As String) As VirtualFolder?
            If shellPath.StartsWith("shell:::", StringComparison.OrdinalIgnoreCase) Then

                Dim raw = shellPath.Substring("shell:::".Length)

                ' Extract only the GUID portion before any trailing slash or segment
                Dim endBrace = raw.IndexOf("}"c)
                If endBrace >= 0 Then
                    raw = raw.Substring(0, endBrace + 1)
                End If

                Dim guid = raw.Trim("{"c, "}"c)

                If GuidMap.TryGetValue(guid, Nothing) Then
                    Return GuidMap(guid)
                End If
            End If

            Return Nothing
        End Function

    End Class

End Namespace