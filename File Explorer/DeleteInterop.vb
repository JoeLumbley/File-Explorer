Imports System.Runtime.InteropServices
Imports System.Text

Namespace Explorer.Interop

    Friend NotInheritable Class DeleteInterop

        Private Sub New()
        End Sub

        Public Const FO_DELETE As UInteger = &H3UI

        <Flags>
        Public Enum FOF As UShort
            FOF_SILENT = &H4
            FOF_NOCONFIRMATION = &H10
            FOF_ALLOWUNDO = &H40
            FOF_NOCONFIRMMKDIR = &H200
            FOF_NOERRORUI = &H400
            FOF_WANTNUKEWARNING = &H4000
        End Enum

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Public Structure SHFILEOPSTRUCT
            Public hwnd As IntPtr
            Public wFunc As UInteger
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pFrom As String
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pTo As String
            Public fFlags As FOF
            <MarshalAs(UnmanagedType.Bool)>
            Public fAnyOperationsAborted As Boolean
            Public hNameMappings As IntPtr
            <MarshalAs(UnmanagedType.LPWStr)>
            Public lpszProgressTitle As String
        End Structure

        <DllImport("shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function SHFileOperation(ByRef lpFileOp As SHFILEOPSTRUCT) As Integer
        End Function

        Public Shared Function DeleteWithShell(
            ownerHandle As IntPtr,
            paths As IEnumerable(Of String),
            allowUndo As Boolean,
            showUI As Boolean,
            ByRef anyAborted As Boolean,
            Optional progressTitle As String = Nothing
        ) As Integer

            Dim sb As New StringBuilder()
            For Each p In paths
                If Not String.IsNullOrWhiteSpace(p) Then
                    sb.Append(p)
                    sb.Append(ChrW(0))
                End If
            Next
            sb.Append(ChrW(0))

            Dim flags As FOF = FOF.FOF_NOCONFIRMMKDIR

            If allowUndo Then
                flags = flags Or FOF.FOF_ALLOWUNDO
            End If

            If Not showUI Then
                flags = flags Or FOF.FOF_SILENT Or FOF.FOF_NOCONFIRMATION Or FOF.FOF_NOERRORUI
            End If

            Dim op As New SHFILEOPSTRUCT With {
                .hwnd = ownerHandle,
                .wFunc = FO_DELETE,
                .pFrom = sb.ToString(),
                .pTo = Nothing,
                .fFlags = flags,
                .lpszProgressTitle = progressTitle
            }

            Dim hr = SHFileOperation(op)
            anyAborted = op.fAnyOperationsAborted
            Return hr
        End Function

    End Class

End Namespace