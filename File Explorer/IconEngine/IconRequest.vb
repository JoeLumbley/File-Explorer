'Imports System

'' Describes what icon is being requested.
'' This is the only thing callers need to construct.
'Public Class IconRequest

'    ' Full file or folder path (if applicable).
'    Public Property FullPath As String

'    ' Optional PIDL pointer for Shell items.
'    Public Property Pidl As IntPtr

'    ' Canonical name for virtual folders (e.g., "RecycleBin").
'    Public Property VirtualName As String

'    ' Desired logical icon size (e.g., 16, 32). Will be DPI-adjusted.
'    Public Property PixelSize As Integer

'    ' Whether the target is a folder. If Nothing, it will be inferred.
'    Public Property IsFolder As Boolean?

'    ' Whether this represents a virtual folder (not a real file system path).
'    Public Property IsVirtual As Boolean

'    ' Optional overlay to apply (shortcut arrow, shared, etc.).
'    Public Property Overlay As IconOverlayKind = IconOverlayKind.None

'    Public ReadOnly Property CacheKeyString As String
'        Get
'            Return IconCache.BuildKey(Me).ToString()
'        End Get
'    End Property

'    Public Overrides Function ToString() As String
'        If IsVirtual Then
'            Return $"Virtual: {VirtualName}, {PixelSize}px"
'        End If
'        Return $"{FullPath}, Folder={IsFolder}, {PixelSize}px"
'    End Function

'End Class