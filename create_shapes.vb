Sub create_shapes()
Dim l, t, w, h As Long
Dim count As Integer
'creates and autosizes shape for each worksheet, applies macro "indiv_toggle" to shape

l = ActiveCell.Left
t = ActiveCell.Top + ActiveCell.Height

count = 0
For Each sht In ActiveWorkbook.Worksheets
    If sht.Name <> ActiveSheet.Name Then
        ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, l, t, 1, 1).Select
        
		With Selection
            With .ShapeRange
                .Line.Visible = msoFalse
                With .TextFrame2
                    .VerticalAnchor = msoAnchorMiddle
                    .TextRange.Characters.Text = sht.Name
                    .TextRange.Characters.ParagraphFormat.Alignment = msoAlignCenter
                    .WordWrap = msoFalse
                    .AutoSize = msoAutoSizeShapeToFitText
                End With
                
                If Sheets(sht.Name).Visible = True Then
                    .Fill.ForeColor.RGB = RGB(27, 72, 123)
                Else
                    .Fill.ForeColor.RGB = RGB(190, 30, 45)
                End If
                .Height = 15
                
                'arranges numeric sheet names in rows of 4
                If IsNumeric(sht.Name) Then
                    If count < 3 Then
                        .Left = l
                        l = l + .Width + 2
                        count = count + 1
                    Else
                        .Left = l
                        count = 0
                        t = t + .Height + 2
                        l = ActiveCell.Left
                    End If
                Else
                    t = t + .Height + 2
                    .Left = ActiveCell.Left
                End If
            End With
            .OnAction = "indiv_toggle"
        End With
    End If
Next sht

End Sub

Sub indiv_toggle()
'toggles shape color and visible property of tab

Dim shpname, shptext As String
Dim show As Boolean
Application.ScreenUpdating = False

shpname = Application.Caller
shptext = ActiveSheet.Shapes.Range(shpname).TextFrame2.TextRange.Text

If ActiveSheet.Shapes.Range(shpname).Fill.ForeColor.RGB = RGB(27, 72, 123) Then
    ActiveSheet.Shapes.Range(shpname).Fill.ForeColor.RGB = RGB(190, 30, 45)
    Call loop_ind(shptext, False)
    Exit Sub
End If

If ActiveSheet.Shapes.Range(shpname).Fill.ForeColor.RGB = RGB(190, 30, 45) Then
    ActiveSheet.Shapes.Range(shpname).Fill.ForeColor.RGB = RGB(27, 72, 123)
    Call loop_ind(shptext, True)
End If

End Sub

Sub loop_ws(str, bool)
'loops through each sheet and sets to visible property

For Each sht In ThisWorkbook.Worksheets
    If InStr(1, sht.Name, str, vbTextCompare) <> 0 Then
        sht.Visible = bool
    End If
Next sht

End Sub
