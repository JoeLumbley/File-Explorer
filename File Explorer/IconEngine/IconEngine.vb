Imports System.Diagnostics
Imports System.Drawing
Imports System.Threading
Imports File_Explorer.Explorer.Interop.Shell

' High-level entry point for icon retrieval (sync + async).
' This is what the rest of the app should call.
'Public Class IconEngine

' Raised when an asynchronously loaded icon becomes available.
' UI layers (ListView/TreeView controllers) subscribe to this once.
'Public Shared Event IconAvailable(request As IconRequest, icon As Icon)

' Synchronous icon retrieval (cache + shell).
' Use this when you absolutely need the icon before continuing.
'Public Shared Function GetIcon(request As IconRequest) As IconResult
'    If request Is Nothing Then Throw New ArgumentNullException(NameOf(request))

'    ' Normalize request (folder detection, DPI scaling).
'    NormalizeRequest(request)

'    ' Build cache key and try cache first.
'    Dim key = IconCache.BuildKey(request)
'    Dim cached As Icon = Nothing

'    If IconCache.TryGet(key, cached) Then
'        Return New IconResult With {
'            .Icon = cached,
'            .Source = IconSourceKind.Cache
'        }
'    End If

'    ' Cache miss → resolve handler and load icon synchronously.
'    Dim handler = IconRules.ResolveHandler(request)
'    Dim icon = handler(request)

'    'If icon Is Nothing Then
'    '    icon = IconLibrary.GenericFile
'    'End If

'    '' Store in cache for future requests.
'    'IconCache.AddOrGetExisting(key, Function() icon)

'    If icon Is Nothing Then
'        ' Do NOT cache fallback icons
'        Return New IconResult With {
'            .Icon = IconLibrary.GenericFile,
'            .Source = IconSourceKind.Fallback
'        }
'    End If

'    IconCache.AddOrGetExisting(key, Function() icon)

'    Return New IconResult With {
'        .Icon = icon,
'        .Source = IconSourceKind.Shell
'    }
'End Function






'Public Shared Function GetIconForPath(path As String, pixelSize As Integer) As Icon

'    ' Native sizes
'    If pixelSize <= 16 Then
'        Return ShellInterop.GetIconForPath(path, IconSize.Small)
'    ElseIf pixelSize = 32 Then
'        Return ShellInterop.GetIconForPath(path, IconSize.Large)
'    End If

'    ' Scale from 32x32
'    Dim baseIcon = ShellInterop.GetIconForPath(path, IconSize.Large)
'    If baseIcon Is Nothing Then Return Nothing

'    Return New Icon(baseIcon, pixelSize, pixelSize)
'End Function
















' Asynchronous icon retrieval (placeholder + background load).
' Use this for ListView/TreeView population to keep the UI snappy.
'Public Shared Async Function GetIconAsync(
'request As IconRequest,
'ct As CancellationToken) As Task(Of IconResult)

'    'If request Is Nothing Then Throw New ArgumentNullException(NameOf(request))
'    If request Is Nothing Then ArgumentNullException.ThrowIfNull(NameOf(request))


'    NormalizeRequest(request)

'    Dim key = IconCache.BuildKey(request)

'    Dim cached As Icon = Nothing
'    If IconCache.TryGet(key, cached) Then
'        Return New IconResult With {.Icon = cached, .Source = IconSourceKind.Cache}
'    End If

'    ' Return placeholder immediately
'    Dim placeholder = GetPlaceholderIcon(request)

'    ' async load
'    Await LoadInBackgroundAsync(request, key, ct)

'    Return New IconResult With {.Icon = placeholder, .Source = IconSourceKind.Placeholder}
'End Function







'Private Shared Async Function LoadInBackgroundAsync(
'request As IconRequest,
'key As IconCacheKey,
'ct As CancellationToken) As Task

'    Try
'        Dim loaded = Await IconLoader.LoadIconInternalAsync(request, ct)
'        If loaded Is Nothing Then Exit Function

'        IconCache.AddOrGetExisting(key, Function() loaded)
'        RaiseIconAvailable(request, loaded)

'    Catch ex As OperationCanceledException
'    Catch ex As Exception
'        Debug.WriteLine("IconEngine async load error: " & ex.Message)
'    End Try
'End Function
















' Normalizes an IconRequest before use.
' This keeps callers simple and centralizes the “rules”.
'Private Shared Sub NormalizeRequest(request As IconRequest)
'    If request Is Nothing Then Throw New ArgumentNullException(NameOf(request))

'    ' Infer folder status if unknown and not virtual.
'    If Not request.IsFolder.HasValue AndAlso Not request.IsVirtual AndAlso
'       Not String.IsNullOrEmpty(request.FullPath) Then
'        request.IsFolder = IO.Directory.Exists(request.FullPath)
'    End If

'    ' Default logical size if not set.
'    If request.PixelSize <= 0 Then
'        request.PixelSize = 16
'    End If

'    ' Convert to DPI-aware pixel size.
'    request.PixelSize = IconDpiManager.ToPixelSize(request.PixelSize)
'End Sub

' Returns a placeholder icon appropriate for the request.
' This is what the UI shows while the real icon is loading.
'Private Shared Function GetPlaceholderIcon(request As IconRequest) As Icon
'    If request Is Nothing Then Return IconLibrary.GenericFile

'    If request.IsVirtual Then
'        Return IconLibrary.GenericVirtualFolder
'    End If

'    If request.IsFolder.GetValueOrDefault() Then
'        Return IconLibrary.GenericFolder
'    End If

'    Return IconLibrary.GenericFile
'End Function

' Raises the IconAvailable event.
' This is intentionally tiny so the event remains easy to reason about.
'Private Shared Sub RaiseIconAvailable(request As IconRequest, icon As Icon)
'    RaiseEvent IconAvailable(request, icon)
'End Sub

'End Class