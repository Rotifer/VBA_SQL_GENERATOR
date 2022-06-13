Option Explicit

Public Sub ColumnSummary(rng As Range, columnIdx As Integer)
    Dim colRng As Range
    Dim rowCount As Long: rowCount = rng.Rows.Count
    Set colRng = Range(rng.Cells(1, 1), rng.Cells(rowCount, 1))
    colRng.Name = colRng.Cells(1, 1).Value
End Sub

Function CellType(cell As Range) As String
    If VBA.IsEmpty(cell) Then
        CellType = "NULL"
    ElseIf Application.IsText(cell) Then
        CellType = "TEXT"
    ElseIf Application.IsLogical(cell) Then
        CellType = "BOOLEAN"
    ElseIf Application.IsErr(cell) Then
        CellType = "ERROR"
    ElseIf VBA.IsDate(cell) Then
        CellType = "DATE"
    ElseIf VBA.InStr(1, cell.Text, ":") <> 0 Then
        CellType = "DATETIME"
    ElseIf VBA.IsNumeric(cell) Then
        If cell.Value = CLng(cell.Value) Then
            CellType = "INTEGER"
        Else
            CellType = "NUMERIC"
        End If
    End If
End Function


Sub run_ColumnSummary()
    Call ColumnSummary(ActiveSheet.UsedRange, 1)
End Sub

Sub test_cell_type()
    Dim rng As Range: Set rng = ActiveSheet.Range("A3:H3")
    Dim cell As Range
    For Each cell In rng.Cells
        Debug.Print (CellType(cell))
    Next cell
End Sub


Function CellValueType(cell As Range) As String
    Select Case VBA.VarType(cell.Value)
        Case 0: CellValueType = "NULL"
        Case 1: CellValueType = "NULL"
        Case 2: CellValueType = "INTEGER"
        Case 3: CellValueType = "INTEGER"
        Case 4: CellValueType = "NUMERIC"
        Case 5: CellValueType = "NUMERIC"
        Case 6: CellValueType = "CURRENCY"
        Case 7: CellValueType = "DATE"
        Case 8: CellValueType = "TEXT"
        Case 9: CellValueType = "OBJECT "
        Case 10: CellValueType = "ERROR"
    End Select
End Function

Sub test_CellValueType()
    Dim rng As Range: Set rng = ActiveSheet.Range("A3:H3")
    Dim cell As Range
    For Each cell In rng.Cells
        Debug.Print (CellValueType(cell))
    Next cell
End Sub
' https://excelmacromastery.com/
Function KeyExists(coll As Collection, key As String) As Boolean
    On Error GoTo EH
    VBA.IsObject (coll.item(key))
    KeyExists = True
EH:
End Function


Sub test()
    Dim col As Collection: Set col = New Collection
    Dim item As String: item = "A"
    Debug.Print KeyExists(col, item)
    col.Add "", item
    Debug.Print KeyExists(col, item)
End Sub

Function ColumnDataTypes(COLUMN As Range) As Collection
    Dim cell As Range
    Dim col As Collection: Set col = New Collection
    For Each cell In COLUMN.Cells
        If Not KeyExists(col, CellType(cell)) And CellType(cell) <> "NULL" Then
            col.Add CellType(cell), CellType(cell)
        End If
    Next cell
    Set ColumnDataTypes = col
End Function

Sub test_ColumnDataTypes()
    Dim COLUMN As Range: Set COLUMN = ActiveSheet.Range("D2:D101")
    Dim colDataTypes As Collection: Set colDataTypes = ColumnDataTypes(COLUMN)
    Dim i As Integer
    Dim typeCount As Integer: typeCount = colDataTypes.Count
    Debug.Print typeCount
    For i = 1 To typeCount
        Debug.Print colDataTypes(i)
    Next i
End Sub


Function AssignColumnDataType(colRng As Range) As String
    Dim colDataTypes As Collection: Set colDataTypes = ColumnDataTypes(colRng)
    If colDataTypes.Count = 1 Then
            AssignColumnDataType = colDataTypes(1)
    ElseIf colDataTypes.Count = 2 And (KeyExists(colDataTypes, "INTEGER") And KeyExists(colDataTypes, "NUMERIC")) Then
        AssignColumnDataType = "NUMERIC"
    Else
        AssignColumnDataType = "TEXT"
    End If
End Function

Sub testAssignColumnDataType()
    Dim rngCol As Range: Set rngCol = ActiveSheet.Range("D2:D101")
    MsgBox AssignColumnDataType(rngCol)
End Sub


Function GetRangeCellsAsArray(rng As Range) As String()
    Dim arr() As String
    Dim cell As Range
    Dim i As Long: i = 0
    For Each cell In rng.Cells
        ReDim Preserve arr(i)
        arr(i) = CStr(cell.Value)
        i = i + 1
    Next cell
    GetRangeCellsAsArray = arr
End Function

Function COLUMN_DEFINITIONS(colNamesRng As Range, colTypesRng As Range) As String
    Dim colNames() As String: colNames = GetRangeCellsAsArray(colNamesRng)
    Dim colTypes() As String: colTypes = GetRangeCellsAsArray(colTypesRng)
    Dim zippedPairs() As String
    Dim zippedPair As String
    Dim i As Long
    For i = 0 To UBound(colNames)
        zippedPair = "    " & colNames(i) & " " & colTypes(i)
        ReDim Preserve zippedPairs(i)
        zippedPairs(i) = zippedPair
    Next i
     COLUMN_DEFINITIONS = Join(zippedPairs, "," & vbCrLf)
End Function

Sub testing()
    Dim rng1 As Range: Set rng1 = ActiveSheet.Range("A1:H1")
    Dim rng2 As Range: Set rng2 = ActiveSheet.Range("A2:H2")
    MsgBox COLUMN_DEFINITIONS(rng1, rng2)
End Sub





























