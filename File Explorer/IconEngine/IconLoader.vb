Imports System.Threading

' Encapsulates background icon loading.
' This keeps async logic out of IconRules and ShellInterop.
Public Class IconLoader

    ' Loads an icon using the rule table on a background thread.
    Public Shared Async Function LoadIconInternalAsync(request As IconRequest,
                                                       ct As CancellationToken) As Task(Of Icon)
        If request Is Nothing Then Return Nothing

        ' Respect cancellation.
        ct.ThrowIfCancellationRequested()

        ' Resolve the appropriate handler for this request.
        Dim handler = IconRules.ResolveHandler(request)

        ' Run handler on a background thread.
        Return Await Task.Run(Function() handler(request), ct)
    End Function

End Class