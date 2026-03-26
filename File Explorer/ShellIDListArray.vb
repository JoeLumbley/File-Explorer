Imports System.Runtime.InteropServices

Public NotInheritable Class ShellIDListArray

    Public ReadOnly Property Items As List(Of IntPtr)

    Private Sub New(list As List(Of IntPtr))
        Items = list
    End Sub

    Public Shared Function FromDataObject(data As IDataObject) As ShellIDListArray
        Dim raw = TryCast(data.GetData("Shell IDList Array"), Byte())
        If raw Is Nothing Then Return New ShellIDListArray(New List(Of IntPtr))

        Dim list As New List(Of IntPtr)
        Dim offset As Integer = 0

        ' First 4 bytes = count
        Dim count As Integer = BitConverter.ToInt32(raw, offset)
        offset += 4

        For i = 0 To count - 1
            Dim pidlSize As Integer = BitConverter.ToInt32(raw, offset)
            offset += 4

            Dim pidlPtr As IntPtr = Marshal.AllocHGlobal(pidlSize)
            Marshal.Copy(raw, offset, pidlPtr, pidlSize)
            offset += pidlSize

            list.Add(pidlPtr)
        Next

        Return New ShellIDListArray(list)
    End Function

End Class