Imports System.IO

Public Module ListViewItemBuilder

    Private ReadOnly SizeUnits As (Unit As String, Factor As Long)() = {
        ("B", 1L), ' Bytes
        ("KB", 1024L), ' Kilobytes
        ("MB", 1024L ^ 2), ' Megabytes
        ("GB", 1024L ^ 3), ' Gigabytes
        ("TB", 1024L ^ 4), ' Terabytes
        ("PB", 1024L ^ 5), ' Petabytes
        ("EB", 1024L ^ 6)  ' Exabytes
    }

    Public Function ForDirectory(di As DirectoryInfo, icons As IconEngine) As ListViewItem
        Dim item As New ListViewItem(di.Name)

        item.SubItems.Add("Folder")
        item.SubItems.Add("")
        item.SubItems.Add(di.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
        item.Tag = di.FullName

        icons.SetListViewItemIconForDirectory(item, di)

        Return item
    End Function

    Public Function ForFile(fi As FileInfo,
                                fileTypeMap As Dictionary(Of String, String),
                                icons As IconEngine) As ListViewItem

        Dim item As New ListViewItem(fi.Name)

        Dim ext = fi.Extension.ToLowerInvariant()
        Dim category = fileTypeMap.GetValueOrDefault(ext, "Document")

        item.SubItems.Add(category)
        item.SubItems.Add(FormatSize(fi.Length))
        item.SubItems.Add(fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
        item.Tag = fi.FullName

        icons.SetListViewItemIconForFile(item, fi)

        Return item
    End Function

    Private Function FormatSize(bytes As Long) As String
        Dim absValue As Double = Math.Abs(bytes)

        ' Walk from largest to smallest unit
        For i As Integer = SizeUnits.Length - 1 To 0 Step -1
            Dim unit = SizeUnits(i)

            If absValue >= unit.Factor Then
                Dim value As Double = absValue / unit.Factor
                Dim formatted As String = $"{value:0.##} {unit.Unit}"
                Return If(bytes < 0, "-" & formatted, formatted)
            End If
        Next

        ' Fallback: smaller than 1 B
        Return $"{bytes} B"
    End Function

End Module




