Imports System.Collections.Concurrent

' Key used to store/retrieve icons from the cache.
' Version ties entries to a specific DPI generation.
Public Structure IconCacheKey
    Public Property Kind As IconKind
    Public Property Identifier As String
    Public Property PixelSize As Integer
    Public Property Version As Integer

    Public Overrides Function ToString() As String
        Return $"{Kind}:{Identifier}:{PixelSize}@{Version}"
    End Function
End Structure

' Thread-safe cache for icons, keyed by IconCacheKey.
' This is intentionally simple: no eviction, no persistence.
Public Class IconCache

    ' Global icon cache shared across the app.
    Private Shared ReadOnly _cache As New ConcurrentDictionary(Of IconCacheKey, Icon)

    ' Builds a cache key from an IconRequest.
    ' This is where we decide whether to cache by extension, path, or virtual name.
    Public Shared Function BuildKey(request As IconRequest) As IconCacheKey
        If request Is Nothing Then Throw New ArgumentNullException(NameOf(request))

        Dim id As String
        Dim kind As IconKind

        If request.IsVirtual Then
            ' Virtual folders keyed by canonical name.
            kind = IconKind.Virtual
            id = request.VirtualName

        ElseIf request.IsFolder.GetValueOrDefault() Then
            ' Folder type keyed by full path for now.
            ' You can later refine this to a folder-type detector.
            kind = IconKind.FolderType
            id = request.FullPath

        Else
            ' Files: prefer extension-based caching when possible.
            Dim ext = IO.Path.GetExtension(request.FullPath)
            If String.IsNullOrEmpty(ext) Then
                kind = IconKind.Path
                id = request.FullPath
            Else
                kind = IconKind.FileExtension
                id = ext.ToLowerInvariant()
            End If
        End If

        Return New IconCacheKey With {
            .Kind = kind,
            .Identifier = id,
            .PixelSize = request.PixelSize,
            .Version = IconDpiManager.CacheVersion
        }
    End Function

    ' Attempts to retrieve an icon from the cache.
    Public Shared Function TryGet(key As IconCacheKey, ByRef icon As Icon) As Boolean
        icon = Nothing

        Dim tmp As Icon = Nothing
        If Not _cache.TryGetValue(key, tmp) Then Return False

        ' Ignore entries from a previous DPI version.
        If key.Version <> IconDpiManager.CacheVersion Then Return False

        icon = tmp
        Return True
    End Function

    ' Adds an icon to the cache or returns an existing one.
    ' The factory is only called if the key is not already present.
    'Public Shared Function AddOrGetExisting(key As IconCacheKey, factory As Func(Of Icon)) As Icon
    '    If factory Is Nothing Then Throw New ArgumentNullException(NameOf(factory))
    '    Return _cache.GetOrAdd(key, Function(_) factory())
    'End Function

    ' Clears the cache and disposes all icons.
    ' This is called on DPI change and can be called manually if needed.
    Public Shared Sub Clear()
        For Each kvp In _cache
            Try
                kvp.Value.Dispose()
            Catch
                ' Ignore disposal errors; cache is best-effort.
            End Try
        Next
        _cache.Clear()
    End Sub

End Class