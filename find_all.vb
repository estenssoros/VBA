Sub FindAll()
Dim fnd As String, FirstFound As String
Dim FoundCell As Range, rng As Range
Dim myRange As Range, LastCell As Range

Set myRange = ActiveSheet.UsedRange
Set LastCell = myRange.Cells(myRange.Cells.Count)

Application.ScreenUpdating = False

r = 2
Do While ActiveSheet.Range("A" & r) <> ""
    fnd = Range("N" & r)
    
    Set FoundCell = myRange.Find(what:=fnd, after:=LastCell)
    
    If Not FoundCell Is Nothing Then
      FirstFound = FoundCell.Address
    End If
    
    Set rng = FoundCell
    
    Do Until FoundCell Is Nothing
        Set FoundCell = myRange.FindNext(after:=FoundCell)
        Set rng = Union(rng, FoundCell)
        If FoundCell.Address = FirstFound Then Exit Do
    Loop
    
    Range("BN" & r) = rng.Cells.Count
    
r = r + 1
Application.StatusBar = r
Loop


Application.ScreenUpdating = True

End Sub