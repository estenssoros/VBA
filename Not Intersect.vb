Private Sub Worksheet_SelectionChange(ByVal Target As Range)
current_row = Target.Row

If Not Intersect(Target, Range("K2:K59")) Is Nothing Then
Application.ScreenUpdating = False
    With Range("K2:K59")
        .Interior.Pattern = xlNone
        .Font.Bold = False
    End With
    
    With Cells(Target.Row, Target.Column)
        .Interior.Pattern = xlSolid
        .Interior.Color = 49407
        .Font.Bold = True
    End With
    
    Range("H2") = Target.Value
    
    Call PositionDiagramm
Application.ScreenUpdating = True
End If
End Sub
