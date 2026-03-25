Imports System.Runtime.InteropServices

Public NotInheritable Class NaturalStringComparer
    Implements IComparer(Of String)

    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function StrCmpLogicalW(x As String, y As String) As Integer
    End Function

    Public Function Compare(x As String, y As String) As Integer _
        Implements IComparer(Of String).Compare

        Return StrCmpLogicalW(x, y)
    End Function
End Class