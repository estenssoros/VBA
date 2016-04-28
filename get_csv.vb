Sub get_csv()

  Dim lCount As Long            ' Generic Counter/Loop
  Dim rLastCell As Range        ' Range for last cell used on worksheet
  Dim sht As Worksheet          ' Worksheet for looping
  Dim wbkToCopy As Workbook     ' Workbook to copy from
  Dim wbkCombined As Workbook   ' Workbook to copy to
  Dim vFileName As Variant      ' Variant for getting file name
  Dim vaFiles As Variant        ' Variant Array for storing selected csv files
  Dim lDefaultSheets            ' Original default sheets created

  With Application
    .ScreenUpdating = False
    lDefaultSheets = .SheetsInNewWorkbook
    .SheetsInNewWorkbook = 1
  End With


  ' Get CSV Files
    vaFiles = Application.GetOpenFilename(fileFilter:="CSV Files (*.csv),*.csv", Title:="Select files", MultiSelect:=True)

  ' Cycle through and copy worksheets over to new workbook
    If IsArray(vaFiles) Then
      Set wbkCombined = Workbooks.Add
      For lCount = LBound(vaFiles) To UBound(vaFiles)
        Set wbkToCopy = Workbooks.Open(FileName:=vaFiles(lCount))
          wbkToCopy.Sheets(1).Copy After:=Workbooks(wbkCombined.Name).Sheets(lCount)
          wbkToCopy.Close savechanges:=False
      Next lCount
      wbkCombined.Sheets(1).Delete

    ' Formatting
      For Each sht In wbkCombined.Sheets
        sht.Select
        If Cells(1, 2) = vbNullString Then Rows(1).Delete
        Rows("1:1").Font.Bold = True
        Rows("1:1").Font.Underline = xlUnderlineStyleSingle

        Set rLastCell = Cells.Find(What:="*", After:=Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        Range(Cells(1, 1), Cells(rLastCell.Row, rLastCell.Column)).Select
        With Selection
          .Columns.AutoFit
          .HorizontalAlignment = xlCenter
          With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
          End With
        End With
      Next sht

    ' Prompt for saving
      vFileName = Application.GetSaveAsFilename(wbkCombined.Name, "Excel files (*.xls),*.xls", 1, "Save Combined Report")
      If vFileName = False Then
        MsgBox "Save cancelled!" & vbNewLine & "You will need to save manually."
      Else
        wbkCombined.SaveAs FileName:=vFileName, FileFormat:=56
        wbkCombined.Close
      End If
    End If

  ' Cleanup
    Set rLastCell = Nothing
    Set sht = Nothing
    Set wbkToCopy = Nothing
    Set wbkCombined = Nothing
    Set vFileName = Nothing
    Set vaFiles = Nothing
    With Application
      .ScreenUpdating = True
      .SheetsInNewWorkbook = lDefaultSheets
      ThisWorkbook.Saved = True
      .Quit
    End With
End Sub