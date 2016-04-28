Sub PositionDiagramm()
    Dim x As Integer
    Dim y As Integer

    Dim height As Integer
    Dim width As Integer
    
    Dim my_left As Double
    Dim col As Integer

    With Range(ActiveWindow.ActivePane.VisibleRange.AddressLocal)
        x = .Left
        y = .Top
        width = .width
        height = .height
    End With
    
    my_left = 0
    col = ActiveCell.Column
    Do While col > 0
        my_left = my_left + Cells(1, col).width
        col = col - 1
    Loop
    

    With ActiveSheet.ChartObjects(1)
        .Top = y + ((height - .height) / 2)
        .Left = my_left + 10
        'x + ((width - .width) / 2)
    End With
End Sub
