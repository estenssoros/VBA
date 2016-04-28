Public MainBar As ProgressBar
Public SubBar As ProgressBar

Sub main_bar(name, ttl_actions)
Set MainBar = New ProgressBar
With MainBar
    .Title = name
    .ExcelStatusBar = False
    .StartColour = rgbRed
    .EndColour = rgbGreen
    .TotalActions = ttl_actions
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .ShowBar
End With
End Sub


Sub sub_bar(name, ttl_actions)

Set SubBar = New ProgressBar
With SubBar
    .Title = name
    .ExcelStatusBar = False
    .StartColour = rgbRed
    .EndColour = rgbGreen
    .TotalActions = ttl_actions
    .StartUpPosition = 0
    .Top = MainBar.Top + MainBar.Height + 10
    .Left = MainBar.Left
    .ShowBar
End With

End Sub
