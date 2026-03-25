Imports System.IO

Public NotInheritable Class NaturalFileSystemComparer
    Implements IComparer(Of FileSystemInfo)

    Private Shared ReadOnly nat As New NaturalStringComparer()

    Public Function Compare(x As FileSystemInfo, y As FileSystemInfo) As Integer _
        Implements IComparer(Of FileSystemInfo).Compare

        Return nat.Compare(x.Name, y.Name)
    End Function
End Class