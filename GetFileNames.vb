Sub GetFileNames()
     
    Dim xRow As Long
    Dim xDirect$, xFname$, InitialFoldr$
     
    InitialFoldr$ = "G:\" '<<< Startup folder to begin searching from
     
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Please select a folder to list Files from"
        .InitialFileName = InitialFoldr$
        .Show
        If .SelectedItems.Count <> 0 Then
            xDirect$ = .SelectedItems(1) & "\"
            xFname$ = Dir(xDirect$, 7)
            Do While xFname$ <> ""
                ActiveCell.Offset(xRow) = xFname$
                xRow = xRow + 1
                xFname$ = Dir
            Loop
        End If
    End With
End Sub