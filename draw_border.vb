Sub draw_border()

Dim row, start_row, stop_row As Integer

start_row = 4
row = 4
For row = 4 To LastCellInSheet(Application.ActiveSheet).row
    If Range("A" & row) = "Comments" Then
        stop_row = row
        Range("A" & start_row & ":C" & stop_row).BorderAround LineStyle:=xlContinuous
        start_row = stop_row + 1
        
    End If
Next row

End Sub