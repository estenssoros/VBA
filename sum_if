Function get_spend(start_cell As Range, test_range As Range, sum_range As Range) As Double
Dim description As String
Dim row, col, test_col, sum_col As Long
Dim test_sht, sum_sht As String
Dim sum As Double

'----------------------------
If test_range.Columns.Count <> sum_range.Columns.Count Then
    MsgBox ("Ranges must be same length")
    Exit Function
End If
'----------------------------

description = start_cell.Value
test_sht = test_range.Worksheet.Name
sum_sht = sum_range.Worksheet.Name

'----------------------------
'find correct row for line item
row = 3
Do While Sheets(sum_sht).Cells(row, 1) <> ""
    If Sheets(sum_sht).Cells(row, 1) = description Then
        Exit Do
    End If
row = row + 1
Loop
'----------------------------
test_col = test_range.Column
sum_col = sum_range.Column

sum = 0
For i = test_col To test_range.Columns.Count + test_col
    If Sheets(test_sht).Cells(start_cell.row, i) <> "" Then
        sum = sum + Sheets(sum_sht).Cells(row, sum_col)
    End If
sum_col = sum_col + 1
Next i

get_spend = sum

End Function
