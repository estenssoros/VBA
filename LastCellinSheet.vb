Function LastCellInSheet(ByRef WhichSheet As Worksheet) _
    As Range

Dim TempRange As Range
Dim LastRow As Long
Dim LastColumn As Long
 
LastRow = 1
LastColumn = 1
 
Set TempRange = WhichSheet.Cells.Find("*", _
    , xlFormulas, xlPart, xlByRows, xlPrevious)
If Not TempRange Is Nothing Then LastRow = TempRange.Row

Set TempRange = WhichSheet.Cells.Find("*", _
    , xlFormulas, xlPart, xlByColumns, xlPrevious)
If Not TempRange Is Nothing Then LastColumn = TempRange.Column
 

Set LastCellInSheet = WhichSheet.Cells.Item(LastRow, LastColumn)
 
End Function