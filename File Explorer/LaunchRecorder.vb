Public Class LaunchRecorder
    Public LastAction As String = Nothing
    Public LastTarget As String = Nothing

    Public Sub Record(action As String, target As String)
        LastAction = action
        LastTarget = target
    End Sub
End Class
