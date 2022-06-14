Attribute VB_Name = "modSqlGenerator"
Option Explicit

' Name: modSqlGenerator
' Objective: Provide a set of functions for building SQL from a table of data in Excel sheets
' Contains functions to generate CREATE TABLE DDL and INSERT INTO SQL statements for that table.
' The output SQL strings are in PostgreSQL format but they should mostly work with SQLite too (except for the callto the 'to_date()' function
' Functions written in all uppercase and meant to be usedin the worksheets themselves asuser-defined functions while functions using capitalised
'  camel casing are helper or utility functions.
' To use: Import this module intoyour personal.xlsb workbook.
' See the README for more on usage and output
' Note: This code was developed on the Mac using Mac Excel so it relies on core VBA and cannot use the many Component Object Model (COM) features that
'  greatly with modern VBA. Because these enhancements are not available, I have to fall back on VBA's feture-poor containers arrays and collections instead of the
'  using Dictionary

'##################################################################################################################################
' Determine the data type of a given single cell range.
' The return values are meant to becompatible with PostgreSQL
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
' Not being used currently becauseit does distinguish NUMERIC and INTEGER
' Source: https://analystcave.com/vba-reference-functions/vba-conversion-functions/vba-vartype-function/
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
' VBA collections do not have a method for testing forthe presence of a key string.
' This function does that check on a given collection object andkey string returning True if the key value is in the collection.
'Idea taken from https://excelmacromastery.com/
Function KeyExists(coll As Collection, key As String) As Boolean
    On Error GoTo EH
    VBA.IsObject (coll.item(key))
    KeyExists = True
EH:
End Function
' Given a column range, build a collection of the data types as determined by' CellType' in the column
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

' Given a column range, get a collection of data types that it contains.
' Apply a set of rules to to assign a single type.
Function COLUMN_DATATYPE(COLUMN As Range) As String
    Dim colDataTypes As Collection: Set colDataTypes = ColumnDataTypes(COLUMN)
    If colDataTypes.Count = 1 Then
            COLUMN_DATATYPE = colDataTypes(1)
    ElseIf colDataTypes.Count = 2 And (KeyExists(colDataTypes, "INTEGER") And KeyExists(colDataTypes, "NUMERIC")) Then
        COLUMN_DATATYPE = "NUMERIC"
    Else
        COLUMN_DATATYPE = "TEXT"
    End If
End Function

' Utility function to return an array of strings for all the data values in the given range.
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

' Build a string of pairs of column names and column column data types to be used in a CREATE TABLE DDL statement
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
     COLUMN_DEFINITIONS = Join(zippedPairs, "," & vbCr)
End Function

' Return a CREATE TABLE DDL statement in PostgreSQL syntax (will also work in SQLite)
Function CREATE_TABLE_DDL(tableName As String, colNamesRng As Range, colTypesRng As Range) As String
    Dim columnsDefinition As String: columnsDefinition = VBA.LTrim(COLUMN_DEFINITIONS(colNamesRng, colTypesRng))
    Dim createTableStmt As String
    createTableStmt = "CREATE TABLE " & tableName & "(" & columnsDefinition & ");"
    CREATE_TABLE_DDL = createTableStmt
End Function


' Given a data value and a column data type, return the data value in a format appropriate for that type that
'  can be used in a PostgreSQL INSERT INTO SQL statement.
Function FormattedValueForInsertStmt(insertValue As String, dataType As String) As String
    If Len(insertValue) = 0 Then
        FormattedValueForInsertStmt = "NULL"
        Exit Function
    End If
    If dataType = "TEXT" Then
        FormattedValueForInsertStmt = "'" & insertValue & "'"
    ElseIf dataType = "DATE" Then
        FormattedValueForInsertStmt = "to_date('" & VBA.Format(CDate(insertValue), "YYYY-MM-DD") & "', 'YYYY-MM-DD')"
    Else
         FormattedValueForInsertStmt = insertValue
    End If
End Function

' Create a full INSERT INTO SQL statatement (PostgreSQL syntax) for a given table name, range of column names and range of column types
Function INSERT_VALUES_SQL(tableName As String, colNamesRng As Range, colTypesRng As Range, dataRowRng As Range) As String
    Dim insertTmpl As String
    Dim i As Integer
    Dim colTypes() As String: colTypes = GetRangeCellsAsArray(colTypesRng)
    Dim colType As String
    Dim formattedValues() As String
    Dim formattedValue As String
    Dim dataRowValues() As String: dataRowValues = GetRangeCellsAsArray(dataRowRng)
    Dim dataValue As String
    insertTmpl = "INSERT INTO  " & tableName & "(" & Join(GetRangeCellsAsArray(colNamesRng), ",") & ") VALUES(%COLUMN_VALUES%);"
    For i = 0 To UBound(colTypes)
        colType = colTypes(i)
        dataValue = dataRowValues(i)
        formattedValue = FormattedValueForInsertStmt(dataValue, colType)
        ReDim Preserve formattedValues(i)
        formattedValues(i) = formattedValue
    Next i
    INSERT_VALUES_SQL = Replace(insertTmpl, "%COLUMN_VALUES%", Join(formattedValues, ","))
End Function
































