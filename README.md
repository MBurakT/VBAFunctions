# Visual Basic For Applications
```vb
Sub Main()
    Dim firstRow As Long: firstRow = 2
    Dim lastRow As Long: lastRow = Cells(Rows.Count, "A").End(xlUp).row - 1
    Dim formul As String: formul = "=A2&C2"
    lastRow = DeleteRows(firstRow, lastRow)
    Call Main()
End Sub

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

Function DeleteRows(ByVal counter As Long, lastRow As Long) As Long
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

Sub CreateColumn(colName As String, colHead As String)
    Columns(colName & ":" & colName).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(colName & "1").Select
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
                    
Sub ChangeColumnHead(colName As String, colHead As String)
    Range(colName & "1").Select
    ActiveCell.FormulaR1C1 = colHead
End Sub

```
