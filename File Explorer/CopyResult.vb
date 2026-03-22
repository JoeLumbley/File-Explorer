
Public Class CopyResult
    Public Property FilesCopied As Integer
    Public Property FilesSkipped As Integer
    Public Property DirectoriesCreated As Integer
    Public Property Errors As New List(Of String)
    Public Property CopiedFilePaths As New List(Of String)

    ' Progress pulse
    Public Property FilesProcessed As Integer
    Public Property TotalDirectories As Integer
    Public Property DirectoriesStarted As Integer

    ' NEW: Total files expected to be processed
    Public Property TotalFiles As Integer   ' ← Add this

    ' Cancellation flag
    Public Property WasCanceled As Boolean

    Public ReadOnly Property Success As Boolean
        Get
            Return Errors.Count = 0 AndAlso Not WasCanceled
        End Get
    End Property

    Public Sub Append(other As CopyResult)
        If other Is Nothing Then Exit Sub

        Me.FilesCopied += other.FilesCopied
        Me.FilesSkipped += other.FilesSkipped
        Me.DirectoriesCreated += other.DirectoriesCreated

        Me.FilesProcessed += other.FilesProcessed
        Me.TotalDirectories += other.TotalDirectories
        Me.DirectoriesStarted += other.DirectoriesStarted

        ' NEW: merge file totals
        Me.TotalFiles += other.TotalFiles

        If other.CopiedFilePaths IsNot Nothing Then
            Me.CopiedFilePaths.AddRange(other.CopiedFilePaths)
        End If

        If other.Errors IsNot Nothing AndAlso other.Errors.Count > 0 Then
            Me.Errors.AddRange(other.Errors)
        End If
    End Sub
End Class

