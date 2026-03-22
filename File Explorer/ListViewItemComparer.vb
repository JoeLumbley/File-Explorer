
Public Class ListViewItemComparer
    Implements IComparer

    Private ReadOnly _column As Integer
    Private ReadOnly _order As SortOrder
    Private ReadOnly _columnTypes As Dictionary(Of Integer, ColumnDataType)

    Public Enum ColumnDataType
        Text
        Number
        DateValue
    End Enum

    Public Sub New(column As Integer, order As SortOrder,
                   Optional columnTypes As Dictionary(Of Integer, ColumnDataType) = Nothing)

        _column = column
        _order = order
        _columnTypes = If(columnTypes, New Dictionary(Of Integer, ColumnDataType))
    End Sub

    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Dim itemX As ListViewItem = CType(x, ListViewItem)
        Dim itemY As ListViewItem = CType(y, ListViewItem)

        Dim valX As String = itemX.SubItems(_column).Text
        Dim valY As String = itemY.SubItems(_column).Text

        Dim result As Integer

        Select Case GetColumnType(_column)
            Case ColumnDataType.Number
                result = CompareNumbers(valX, valY)

            Case ColumnDataType.DateValue
                result = CompareDates(valX, valY)

            Case Else
                result = String.Compare(valX, valY, StringComparison.OrdinalIgnoreCase)
        End Select

        If _order = SortOrder.Descending Then result = -result
        Return result
    End Function

    Private Function GetColumnType(col As Integer) As ColumnDataType
        If _columnTypes.ContainsKey(col) Then
            Return _columnTypes(col)
        End If
        Return ColumnDataType.Text
    End Function

    ' -----------------------------
    '   NUMBER SORTING (FILE SIZE)
    ' -----------------------------
    Private Function CompareNumbers(a As String, b As String) As Integer
        Dim n1 As Long = ParseSize(a)
        Dim n2 As Long = ParseSize(b)
        Return n1.CompareTo(n2)
    End Function

    ' -----------------------------
    '   DATE SORTING
    ' -----------------------------
    Private Function CompareDates(a As String, b As String) As Integer
        Dim d1, d2 As DateTime
        DateTime.TryParse(a, d1)
        DateTime.TryParse(b, d2)
        Return d1.CompareTo(d2)
    End Function

    Private Function ParseSize(input As String) As Long
        If String.IsNullOrWhiteSpace(input) Then Return 0

        ' Normalize
        Dim text = input.Trim().Replace(",", "").ToUpperInvariant()

        ' Extract sign
        Dim isNegative As Boolean = text.StartsWith("-")
        If isNegative Then text = text.Substring(1).Trim()

        ' Split into number + unit
        Dim parts = text.Split(" "c, StringSplitOptions.RemoveEmptyEntries)
        Dim numberPart As String = parts(0)
        Dim unitPart As String = If(parts.Length > 1, parts(1), "B")

        ' Parse numeric portion
        Dim value As Double
        If Not Double.TryParse(numberPart, Globalization.NumberStyles.Float,
                           Globalization.CultureInfo.InvariantCulture, value) Then
            Return 0
        End If

        ' Look up the factor from the shared SizeUnits array
        Dim factor As Long = 1
        For Each u In Form1.SizeUnits
            If u.Unit.Equals(unitPart, StringComparison.OrdinalIgnoreCase) Then
                factor = u.Factor
                Exit For
            End If
        Next

        ' Compute final byte count
        Dim bytes As Double = value * factor
        Dim result As Long = CLng(Math.Round(bytes))

        Return If(isNegative, -result, result)
    End Function

    ' TEMP: expose parser for testing
    Public Function Test_ParseSize(input As String) As Long
        Return ParseSize(input)
    End Function

End Class
