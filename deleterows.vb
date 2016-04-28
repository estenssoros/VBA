Sub remove_rows(sheet, start_row)
Dim row As Long
Dim sht As String

sht = sheet
row = start_row
Do While Sheets(sht).Range("C" & row) <> ""
    If Sheets(sht).Range("D" & row) = "" Then
        Sheets(sht).Rows(row).EntireRow.Delete
        row = row - 1
    End If
    
row = row + 1
Loop

End Sub