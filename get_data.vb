Sub FIND_DATA()
Dim customer, service, temp, supplier As String
Dim row, row2, col, col2, supplier_col As Long
Dim rng As Range
Application.ScreenUpdating = False
'clear old data
Sheets("MASTER").Columns("E:AT").ClearContents
Sheets("MASTER").Columns("E:AT").ColumnWidth = 2

customer = Sheets("MASTER").Range("C2")
service = Sheets("MASTER").Range("C3")

'populate customer rows in MASTER
row = 3
col = 3
Do While Sheets("Raw").Cells(row, col) <> ""
    'find customer column
    If Sheets("Raw").Cells(row, col) = customer Then
        'find supplier
        row = 4
        col2 = 5
        Do While Sheets("Raw").Cells(row, 3) <> ""
            If Sheets("Raw").Cells(row, col) <> "" And Sheets("raw").Cells(row, 4) = service Then
                
                'formatting
                Sheets("MASTER").Columns(col2).ColumnWidth = 20
                Sheets("MASTER").Columns(col2 + 1).ColumnWidth = 20
                
                temp = Sheets("Raw").Cells(row, 3)
                'contractor pricing
                Sheets("MASTER").Cells(5, col2) = temp
                'markups and travel allowances
                Sheets("MASTER").Cells(9, col2) = temp
                Sheets("MASTER").Cells(10, col2) = "Value ($ or %)"
                'straight time labor rate sheet
                Sheets("MASTER").Cells(15, col2) = temp
                Sheets("MASTER").Cells(16, col2) = "Hourly Rate ST ($/hr)"
                Sheets("MASTER").Cells(16, col2 + 1) = "Hourly Rate OT ($/hr)"
                'equipment rate sheet
                Sheets("MASTER").Cells(39, col2) = temp
                Sheets("MASTER").Cells(40, col2) = "Billable ST Rate ($/hr)"
                Sheets("MASTER").Cells(40, col2 + 1) = "Example Model (or equiv)"
                'material rate sheet
                Sheets("MASTER").Cells(105, col2) = temp
                Sheets("MASTER").Cells(106, col2) = "Price ($/unit)"
                'aditional scope sheet
                Sheets("MASTER").Cells(124, col2) = temp
                Sheets("MASTER").Cells(125, col2) = "Rate ($/unit)"
                
                col2 = col2 + 2
            End If
        row = row + 1
        Loop
    Exit Do
    End If
col = col + 1
Loop
On Error Resume Next
'Pull supplier data from Form E
col = 5
'loop through suppliers on master
Do While Sheets("MASTER").Cells(5, col) <> ""
    supplier = Sheets("MASTER").Cells(5, col)
    
    'find supplier column in FormE
    Set rng = Sheets("Form E").Range("B2:IS2").Find(supplier, lookat:=xlPart)
    supplier_col = rng.Column
    
    'date completed
    row = 6
    row2 = row - 3
    Sheets("MASTER").Cells(row, col) = Sheets("Form E").Cells(row2, supplier_col)
    
    'markups and travel allowances rate sheet
    row2 = 15
    For row = 11 To 12
        Sheets("MASTER").Cells(row, col) = Sheets("Form E").Cells(row2, supplier_col)
        row2 = row2 + 1
    Next row
    
    'straight time labor rate sheet
    Call row_loop("MASTER", "Form E", 17, 32, 12, col, supplier_col)
    
    'get comments
    Call row_loop("MASTER", "Form E", 35, 37, 12, col, supplier_col)
    
    'equipment rate sheet
    Call row_loop("MASTER", "Form E", 41, 103, 19, col, supplier_col)
    
    'material rate sheet
    Call row_loop("MASTER", "Form E", 107, 122, 40, col, supplier_col)
    
    'additional scope rate sheet
    Call row_loop("MASTER", "Form E", 126, 137, 48, col, supplier_col)
    
col = col + 2
Loop



Application.ScreenUpdating = True
End Sub


Sub row_loop(to_sheet As String, from_sheet As String, start_row As Long, stop_row As Long, off_set As Long, to_col As Long, from_col As Long)

    For row = start_row To stop_row
        Sheets(to_sheet).Cells(row, to_col) = Sheets(from_sheet).Cells(row + off_set, from_col)
        Sheets(to_sheet).Cells(row, to_col + 1) = Sheets(from_sheet).Cells(row + off_set, from_col + 1)
    
    Next row

End Sub
