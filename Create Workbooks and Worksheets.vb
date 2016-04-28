Sub Try()
Dim row, row2, col As Long
Dim rate_sheet, category, description, company, service As String
Dim dest As String
Dim WS As Worksheet
Dim wb, newwb As Workbook

With Application.FileDialog(msoFileDialogFolderPicker)
    .Show
    dest = .SelectedItems(1)
End With


wb = ActiveWorkbook.Name
row = 3
Do While Sheets("LISTS").Range("F" & row) <> ""
    company = Sheets("LISTS").Range("F" & row)

    Workbooks.Add
    ActiveWorkbook.SaveAs (dest & "\" & company & ".xlsx")
        
    row2 = 3
    Do While Workbooks(wb).Sheets("LISTS").Range("G" & row2) <> ""
        service = Workbooks(wb).Sheets("LISTS").Range("G" & row2)
        Sheets.Add.Name = service
    row2 = row2 + 1
    Loop
    
    Call CreateRateSheet(company, sht, wb)
    Sheets("Sheet1").Delete
    Sheets("Sheet2").Delete
    Sheets("Sheet3").Delete
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
row = row + 1
Loop

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub