
Imports System.IO
Imports File_Explorer.Explorer.Navigation

Public Class SafeLaunchEngine

    Private ReadOnly _owner As Form1
    Private ReadOnly _recorder As LaunchRecorder   ' Optional for tests

    Public Sub New(owner As Form1, Optional recorder As LaunchRecorder = Nothing)
        _owner = owner
        _recorder = recorder
    End Sub

    Public Function SafeLaunch(input As String) As Boolean
        If String.IsNullOrWhiteSpace(input) Then
            Return False
        End If

        Dim path = input.Trim()

        ' 0. Virtual folder (shell::: or shell:)
        If path.StartsWith("shell:::", StringComparison.OrdinalIgnoreCase) Then

            Dim vf = VirtualFolderResolver.Resolve(path)
            If vf.HasValue Then
                Return LaunchVirtualFolder(vf.Value)
            End If

        ElseIf path.StartsWith("shell:", StringComparison.OrdinalIgnoreCase) Then

            Dim vf = VirtualFolderMap.ResolveCanonical(path)
            If vf.HasValue Then
                Return LaunchVirtualFolder(vf.Value)
            End If

        End If

        ' 1. URL
        If IsSafeUrl(path) Then
            Return LaunchUrl(path)
        End If

        ' 2. Folder
        If Directory.Exists(path) Then
            Return LaunchFolder(path)
        End If

        ' 3. File
        If File.Exists(path) Then
            Return LaunchFile(path)
        End If

        Return False
    End Function

    Private Function LaunchVirtualFolder(vf As VirtualFolder) As Boolean
        Try
            ShellNavigation.OpenVirtualFolder(vf)
            Return True
        Catch
            Return False
        End Try
    End Function

    ' ---------------------------------------------------------
    '  Launchers
    ' ---------------------------------------------------------

    Private Function LaunchUrl(url As String) As Boolean
        If _recorder IsNot Nothing Then
            _recorder.Record("Url", url)
            Return True
        End If

        Try
            Dim psi As New ProcessStartInfo(url) With {.UseShellExecute = True}
            Process.Start(psi)

            _owner.ShowStatus(_owner.StatusPad & _owner.IconOpen & $"  Opening URL: {url}")
            Return True

        Catch ex As Exception
            _owner.ShowStatus(_owner.StatusPad & _owner.IconError & " Cannot open URL: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function LaunchFolder(path As String) As Boolean
        If _recorder IsNot Nothing Then
            _recorder.Record("Folder", path)
            Return True
        End If

        _owner.NavigateTo(path)
        Return True
    End Function

    Private Function LaunchFile(filePath As String) As Boolean
        If _recorder IsNot Nothing Then
            _recorder.Record("File", filePath)
            Return True
        End If

        Try
            Dim psi As New ProcessStartInfo(filePath) With {.UseShellExecute = True}
            Process.Start(psi)

            Dim name = Path.GetFileNameWithoutExtension(filePath)
            _owner.ShowStatus(_owner.StatusPad & _owner.IconOpen & $"  Opened:   ""{name}""")
            Return True

        Catch ex As Exception
            _owner.ShowStatus(_owner.StatusPad & _owner.IconError & " Cannot open: " & ex.Message)
            Return False
        End Try
    End Function

    ' ---------------------------------------------------------
    '  URL Validation
    ' ---------------------------------------------------------
    Private Function IsSafeUrl(text As String) As Boolean
        Dim uri As Uri = Nothing

        If Not Uri.TryCreate(text, UriKind.Absolute, uri) Then
            Return False
        End If

        Select Case uri.Scheme
            Case Uri.UriSchemeHttp, Uri.UriSchemeHttps
                Return True
            Case Else
                Return False
        End Select
    End Function

End Class

