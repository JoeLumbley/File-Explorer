Namespace Explorer.Navigation

    Public Class VirtualItem
        Public Property Name As String
        Public Property Type As String
        Public Property Size As Long?
        Public Property Modified As Date?
        Public Property Pidl As IntPtr
        Public Property OriginalPath As String   ' For Recycle Bin
        Public Property IsFolder As Boolean
    End Class

End Namespace