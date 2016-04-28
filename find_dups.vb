Sub find_dups(search_col, out_col)
'--------------------------------------
'generic sub to find duplicates in search column and output count
'--------------------------------------
Dim fnd As String, FirstFound As String
Dim FoundCell As Range, rng As Range
Dim myRange As Range, LastCell As Range

Set myRange = ActiveSheet.Range(search_col & "2:" & search_col & LastCellInSheet(ActiveSheet).Row)
Set LastCell = myRange.Cells(myRange.Cells.Count)

Call sub_bar("Searching Entries", LastCellInSheet(ActiveSheet).Row)
For r = 2 To LastCellInSheet(ActiveSheet).Row
    
    Range(out_col & r).ClearContents
    
    fnd = Range(search_col & r)
    If fnd = "" Then
        Range(out_col & r) = 0
        GoTo nextfor
        SubBar.NextAction
    End If
    
    Set FoundCell = myRange.Find(what:=fnd, lookat:=xlWhole, after:=LastCell)
    
    If Not FoundCell Is Nothing Then
      FirstFound = FoundCell.Address
    End If
    
    Set rng = FoundCell
    
    Do Until FoundCell Is Nothing
        Set FoundCell = myRange.FindNext(after:=FoundCell)
        Set rng = Union(rng, FoundCell)
        If FoundCell.Address = FirstFound Then Exit Do
    Loop
    
    Range(out_col & r) = rng.Cells.Count
    
    SubBar.NextAction
nextfor:
Next r

SubBar.Terminate
