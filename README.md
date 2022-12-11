# VBA Functions
```vb
Sub Main()
    On Error GoTo toEnd
    
    Dim firstRow As Long: firstRow = 2
    Dim lastRow As Long: lastRow = Cells(Rows.Count, "A").End(xlUp).row - 1
    Dim colCount As Long: colCount = Cells(firstRow, Columns.Count).End(xlToLeft).Column
    Dim formul As String: formul = "=A2&C2"
    Dim colOrdinal As Long: colOrdinal = 2
    lastRow = DeleteRows(firstRow, lastRow, colOrdinal)
    FunctionName 'without parameter
    Call FunctionName("with parameter")
    
toEnd:
End Sub
```
```vb
Function DeleteRowsBaseOnCellValue(ByVal firstRow As Long, lastRow As Long, colOrdinal As Long) As Long 'ByRef
    While (firstRow <= lastRow)
        If Mid(Cells(firstRow, colOrdinal), 1, 3) = "***" Or Mid(Cells(firstRow, colOrdinal), 1, 3) = "Ref" Then
            Rows(firstRow).Delete
            lastRow = lastRow - 1
        Else
            firstRow = firstRow + 1
        End If
    Wend
    DeleteRows = lastRow
End Function
```
```vb
Sub DeleteRowsRange(firstRow As Long, lastRow As Long)
    Rows(firstRow & ":" & lastRow).Select
    Selection.Delete Shift:=xlUp
End Sub
```
```vb
Sub ClearCell(firstRow As Long, lastRow As Long, colName As String)
    For Each r In Range(colName & firstRow & ":" & colName & lastRow)
        If r.Text = "#N/A" Or r.Text = "#DIV/0!" Then
            r.value = 0
        ElseIf (r.value < 0) Then
            r.value = 0
        End If
    Next
End Sub
```
```vb
Sub CreateColumnWithHeader(colName As String, colHead As String, position As Long)
    Columns(colName & ":" & colName).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(colName & position).Select
    ActiveCell.FormulaR1C1 = colHead
End Sub
```
```vb
Sub ChangeColumnOrdinal(currentPosition As String, targetPosition As String)
    Columns(currentPosition & ":" & currentPosition).Select
    Selection.Cut
    Columns(targetPosition & ":" & targetPosition).Select
    Selection.Insert Shift:=xlToRight
End Sub
```
```vb    
Sub CopyColumn(firstCol As String, lastCol As String, targetCol As String)
    Columns(firstCol & ":" & lastCol).Select
    Selection.Copy
    Columns(targetCol & ":" & targetCol).Select
    Selection.Insert Shift:=xlToRight
End Sub
```
```vb         
Sub ChangeColumnHead(colName As String, colHead As String, position as Long)
    Range(colName & position).Select
    ActiveCell.FormulaR1C1 = colHead
End Sub
```
```vb                 
Sub FormulaToColumn(firstRow As Long, lastRow As Long, colName As String, formul As String)
    Range(colName & firstRow).Formula = formul
    Range(colName & firstRow).AutoFill Destination:=Range(colName & firstRow & ":" & colName & lastRow)
End Sub
```
```vb
Sub KillFormulas(firstRow As Long, lastRow As Long, firstCol As String, lastCol As String)
    Range(firstCol & firstRow & ":" & lastCol & lastRow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub
```
```vb
Sub CreatePivot(pivotTable As String, sheetName As String, mainSheetName As String, firstRow As Long, lastRow As Long, colCount As Long)
    Sheets(mainSheetName).Select
    Application.CutCopyMode = False
    Sheets.Add.Name = sheetName
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "'" & mainSheetName & "'!R" & firstRow & "C1:R" & lastRow & "C" & colCount, Version:=6).CreatePivotTable TableDestination:= _
        "'" & sheetName & "'!R3C1", TableName:=pivotTable, DefaultVersion:=6
    Sheets(sheetName).Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables(pivotTable).PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables(pivotTable).RepeatAllLabels xlRepeatLabels
End Sub
```
```vb
Sub AddToPivotRows(pivotTable As String, colName As String, colPosition As Integer)
    With ActiveSheet.PivotTables(pivotTable).PivotFields(colName)
        .Orientation = xlRowField
        .Position = colPosition
    End With
End Sub
```
```vb                                                         
Sub FieldSettings(pivotTable As String, colName As String)
    ActiveSheet.PivotTables(pivotTable).PivotFields(colName).Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables(pivotTable).PivotFields(colName).LayoutForm = xlTabular
End Sub
```
```vb                        
Sub ValuesFilter(pivotTable As String, colName As String, showName As String)
    ActiveSheet.PivotTables(pivotTable).AddDataField ActiveSheet.PivotTables( _
        pivotTable).PivotFields(colName), showName, xlSum
End Sub
```
