'Imports System.IO
'Imports System.Threading
'Imports System.Threading.Tasks
''Imports Explorer.Interop
'Imports File_Explorer.Explorer.Interop

'Namespace Explorer.Engines

'    Public Class DeleteProgress
'        Public Property ItemsTotal As Integer
'        Public Property ItemsProcessed As Integer
'        Public Property AnyAborted As Boolean
'        Public Property LastError As String
'    End Class

'    Public Class DeleteEngine

'        Private ReadOnly _ownerHandle As IntPtr
'        Private ReadOnly _showStatus As Action(Of String)
'        Private ReadOnly _helpPanel As Control

'        Public Sub New(ownerHandle As IntPtr,
'                       helpPanel As Control,
'                       showStatus As Action(Of String))
'            _ownerHandle = ownerHandle
'            _helpPanel = helpPanel
'            _showStatus = showStatus
'        End Sub

'        Public Async Function DeleteAsync(
'            paths As List(Of String),
'            allowUndo As Boolean,
'            ct As CancellationToken,
'            Optional progress As DeleteProgress = Nothing
'        ) As Task(Of Boolean)

'            If paths Is Nothing OrElse paths.Count = 0 Then Return False

'            If progress IsNot Nothing Then
'                progress.ItemsTotal = paths.Count
'                progress.ItemsProcessed = 0
'                progress.AnyAborted = False
'                progress.LastError = Nothing
'            End If

'            '_helpPanel.Visible = True
'            '_helpPanel.BringToFront()

'            Try
'                ct.ThrowIfCancellationRequested()

'                Dim anyAborted As Boolean = False

'                Dim hr = Await Task.Run(
'                    Function()
'                        ct.ThrowIfCancellationRequested()
'                        Dim aborted As Boolean = False
'                        Dim r = DeleteInterop.DeleteWithShell(
'                            _ownerHandle,
'                            paths,
'                            allowUndo,
'                            showUI:=True,
'                            anyAborted:=aborted,
'                            progressTitle:="Deleting..."
'                        )
'                        anyAborted = aborted
'                        Return r
'                    End Function, ct)

'                If progress IsNot Nothing Then
'                    progress.ItemsProcessed = paths.Count
'                    progress.AnyAborted = anyAborted
'                End If

'                If ct.IsCancellationRequested OrElse anyAborted Then
'                    _showStatus("Delete canceled or aborted.")
'                    Return False
'                End If

'                If hr <> 0 Then
'                    If progress IsNot Nothing Then
'                        progress.LastError = "Shell delete failed. HRESULT=" & hr.ToString()
'                    End If
'                    _showStatus("Delete failed. One or more items may not have been removed.")
'                    Return False
'                End If

'                Return True

'            Catch ex As OperationCanceledException
'                _showStatus("Delete canceled.")
'                If progress IsNot Nothing Then
'                    progress.AnyAborted = True
'                    progress.LastError = "Canceled."
'                End If
'                Return False

'            Catch ex As Exception
'                _showStatus("Delete failed: " & ex.Message)
'                If progress IsNot Nothing Then
'                    progress.LastError = ex.Message
'                End If
'                Return False

'            Finally
'                '_helpPanel.Visible = False
'            End Try

'        End Function

'    End Class

'End Namespace



Imports System.Threading
Imports System.Threading.Tasks
'Imports Explorer.Interop
Imports File_Explorer.Explorer.Interop

Namespace Explorer.Engines

    Public Class DeleteProgress
        Public Property ItemsTotal As Integer
        Public Property ItemsProcessed As Integer
        Public Property AnyAborted As Boolean
        Public Property LastError As String
    End Class

    Public Class DeleteEngine

        Private ReadOnly _ownerHandle As IntPtr
        Private ReadOnly _showStatus As Action(Of String)
        Private ReadOnly _helpPanel As Control

        Public Sub New(ownerHandle As IntPtr,
                       helpPanel As Control,
                       showStatus As Action(Of String))
            _ownerHandle = ownerHandle
            _helpPanel = helpPanel
            _showStatus = showStatus
        End Sub

        Public Async Function DeleteAsync(
            paths As List(Of String),
            allowUndo As Boolean,
            ct As CancellationToken,
            Optional progress As DeleteProgress = Nothing
        ) As Task(Of Boolean)

            If paths Is Nothing OrElse paths.Count = 0 Then Return False

            If progress IsNot Nothing Then
                progress.ItemsTotal = paths.Count
                progress.ItemsProcessed = 0
                progress.AnyAborted = False
                progress.LastError = Nothing
            End If

            Try
                ct.ThrowIfCancellationRequested()

                Dim ok = Await Task.Run(
                    Function()
                        Return PerformShellDelete(paths, allowUndo, ct, progress)
                    End Function, ct)

                Return ok

            Catch ex As OperationCanceledException
                If progress IsNot Nothing Then
                    progress.AnyAborted = True
                    progress.LastError = "Canceled."
                End If
                _showStatus("Delete canceled.")
                Return False

            Catch ex As Exception
                If progress IsNot Nothing Then
                    progress.LastError = ex.Message
                End If
                _showStatus("Delete failed: " & ex.Message)
                Return False
            End Try

        End Function

        '    Private Function PerformShellDelete(
        '        paths As List(Of String),
        '        allowUndo As Boolean,
        '        ct As CancellationToken,
        '        progress As DeleteProgress
        '    ) As Boolean

        '        'Dim fo As IFileOperation = CType(New FileOperation(), IFileOperation)

        '        Dim CLSID_FileOperation As New Guid("3AD05575-8857-4850-9277-11B85BDB8E09")
        '        Dim fo As IFileOperation =
        'CType(Activator.CreateInstance(Type.GetTypeFromCLSID(CLSID_FileOperation)),
        '      IFileOperation)

        '        Dim flags As FileOperationFlags =
        '            FileOperationFlags.FOF_NOCONFIRMMKDIR Or
        '            FileOperationFlags.FOFX_SHOWELEVATIONPROMPT

        '        If allowUndo Then
        '            flags = flags Or
        '                     FileOperationFlags.FOF_ALLOWUNDO Or
        '                     FileOperationFlags.FOFX_RECYCLEONDELETE
        '        End If

        '        fo.SetOperationFlags(flags)
        '        fo.SetOwnerWindow(_ownerHandle)

        '        ' Optional: you could implement a real progress sink here.
        '        Dim cookie As UInteger
        '        fo.Advise(Nothing, cookie)

        '        For Each p In paths
        '            ct.ThrowIfCancellationRequested()

        '            Dim psi As IShellItem = ShellInterop.CreateShellItemFromPath(p)
        '            fo.DeleteItem(psi, Nothing)

        '            If progress IsNot Nothing Then
        '                progress.ItemsProcessed += 1
        '            End If
        '        Next

        '        fo.PerformOperations()

        '        Dim aborted = fo.GetAnyOperationsAborted()
        '        If progress IsNot Nothing Then
        '            progress.AnyAborted = aborted
        '        End If

        '        fo.Unadvise(cookie)

        '        If aborted Then
        '            _showStatus("Delete canceled or aborted.")
        '            Return False
        '        End If

        '        Return True
        '    End Function
        Private Function PerformShellDelete(
    paths As List(Of String),
    allowUndo As Boolean,
    ct As CancellationToken,
    progress As DeleteProgress
) As Boolean

            ct.ThrowIfCancellationRequested()

            Dim anyAborted As Boolean = False

            Dim hr = DeleteInterop.DeleteWithShell(
                ownerHandle:=_ownerHandle,
                paths:=paths,
                allowUndo:=allowUndo,
                showUI:=True,
                anyAborted:=anyAborted,
                progressTitle:="Deleting..."
            )

            If progress IsNot Nothing Then
                progress.ItemsProcessed = paths.Count
                progress.AnyAborted = anyAborted
                If hr <> 0 Then
                    progress.LastError = "Shell delete failed. HRESULT=" & hr.ToString()
                End If
            End If

            If anyAborted Then
                _showStatus("Delete canceled or aborted.")
                Return False
            End If

            If hr <> 0 Then
                _showStatus("Delete failed. One or more items may not have been removed.")
                Return False
            End If

            Return True
        End Function


    End Class

End Namespace