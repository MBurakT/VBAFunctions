Sub Main()
    Dim firstRow As Long: firstRow = 2
    Dim lastRow As Long: lastRow = Cells(Rows.Count, "A").End(xlUp).row - 1
    Dim colCount As Long: colCount = Cells(firstRow, Columns.Count).End(xlToLeft).Column
    Dim formul As String: formul = "=A2&C2"
    lastRow = DeleteRows(firstRow, lastRow)
    Function 'without parameter
    Call Function("with parameter")
End Sub

Function DeleteRows(ByVal counter As Long, lastRow As Long) As Long 'ByRef
    While (counter <= lastRow)
        If Mid(Cells(counter, 2), 1, 3) = "***" Or Mid(Cells(counter, 2), 1, 3) = "Ref" Then
            Rows(counter).Delete
            lastRow = lastRow - 1
        Else
            counter = counter + 1
        End If
    Wend
    DeleteRows = lastRow
End Function

Sub DeleteRowsRange(firstRow As Long, lastRow As Long)
    Rows(firstRow & ":" & lastRow).Select
    Selection.Delete Shift:=xlUp
End Sub

Sub FormulaToColumn(firstRow As Long, lastRow As Long, colName As String, formul As String)
    Range(colName & firstRow).Formula = formul
    Range(colName & firstRow).AutoFill Destination:=Range(colName & firstRow & ":" & colName & lastRow)
End Sub

Sub KillFormulas(firstRow As Long, lastRow As Long, firstCol As String, lastCol As String)
    Range(firstCol & firstRow & ":" & lastCol & lastRow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

Sub CreateColumnWithHeader(colName As String, colHead As String, position As Long)
    Columns(colName & ":" & colName).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(colName & position).Select
    ActiveCell.FormulaR1C1 = colHead
End Sub
        
Sub ClearCell(firstRow As Long, lastRow As Long, colName As String)
    For Each r In Range(colName & firstRow & ":" & colName & lastRow)
        If r.Text = "#N/A" Or r.Text = "#DIV/0!" Then
            r.value = 0
        ElseIf (r.value < 0) Then
            r.value = 0
        End If
    Next
End Sub

Sub ChangeColumnOrdinal(currentPosition As String, targetPosition As String)
    Columns(currentPosition & ":" & currentPosition).Select
    Selection.Cut
    Columns(targetPosition & ":" & targetPosition).Select
    Selection.Insert Shift:=xlToRight
End Sub
            
Sub CopyColumn(firstCol As String, lastCol As String, targetCol As String)
    Columns(firstCol & ":" & lastCol).Select
    Selection.Copy
    Columns(targetCol & ":" & targetCol).Select
    Selection.Insert Shift:=xlToRight
End Sub
                    
Sub ChangeColumnHead(colName As String, colHead As String, position as Long)
    Range(colName & position).Select
    ActiveCell.FormulaR1C1 = colHead
End Sub
       
Sub CreatePivot(pivotTable As String, sheetName As String, firstRow As Long, lastRow As Long, colCount As Long)
    Sheets("Sheet1").Select
    Application.CutCopyMode = False
    Sheets.Add.Name = sheetName
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R" & firstRow & "C" & firstRow & ":R" & lastRow & "C" & colCount, Version:=6).CreatePivotTable TableDestination:= _
        "'" & sheetName & "'!R3C1", TableName:=pivotTable, DefaultVersion:=6
    Sheets(sheetName).Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables(pivotTable).PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables(pivotTable).RepeatAllLabels xlRepeatLabels
End Sub

Sub AddToPivotRows(pivotTable As String, colName As String, colPosition As Integer)
    With ActiveSheet.PivotTables(pivotTable).PivotFields(colName)
        .Orientation = xlRowField
        .Position = colPosition
    End With
End Sub
                                    
Sub ValuesFilter(pivotTable As String, colName As String, showName As String)
    ActiveSheet.PivotTables(pivotTable).AddDataField ActiveSheet.PivotTables( _
        pivotTable).PivotFields(colName), showName, xlSum
End Sub
                                    
Sub FieldSettings(pivotTable As String, colName As String)
    ActiveSheet.PivotTables(pivotTable).PivotFields(colName).Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables(pivotTable).PivotFields(colName).LayoutForm = xlTabular
End Sub
