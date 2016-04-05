Sub PullMaterials()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim row As Integer
Dim district, VS, CSV, TestDir As String

Dim MyProgressBar As ProgressBar
Set MyProgressBar = New ProgressBar

With MyProgressBar
    .title = "Pulling materials from drive."
    .ExcelStatusBar = True
    .StartColour = rgbRed
    .EndColour = rgbGreen
    .TotalActions = 15
    .StartUpPosition = 0
    .left = Application.left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .ShowBar
End With

TestDir = "\\somedir.csv"

On Error GoTo errhandler
If Len(Dir(TestDir)) = 0 Then
    MsgBox ("Not connected to network")
    MyProgressBar.Terminate
    Call PswrdLock
    Exit Sub
End If
    

district = Sheets(1).Range("G5")
Sheets("MATERIALS").Range("B1") = district

VS = ActiveWorkbook.Name

'1
MyProgressBar.NextAction "Downloading district .csv file"

'open materials.csv file for district
Workbooks.Open ("\\somedir".csv"), , _
                    , , , , , , , , Notify = True
'set .scv workbook name
CSV = ActiveWorkbook.Name

'=========================================================================
'find bottom row of materials
row = 3
Do While Workbooks(VS).Sheets("MATERIALS").Range("B" & row) <> ""
    row = row + 1
Loop
'2
MyProgressBar.NextAction "Clearing materials"
'clear materials
Workbooks(VS).Sheets("MATERIALS").Range("B3:F" & row - 1).ClearContents

'----------------------------
'find bottom row of rigs
row = 3
Do While Workbooks(VS).Sheets("MATERIALS").Range("H" & row) <> ""
    row = row + 1
Loop
'3
MyProgressBar.NextAction "Clearing rigs"
'clear rigs
Workbooks(VS).Sheets("MATERIALS").Range("H3:H" & row - 1).ClearContents

'----------------------------
'find bottom row of customers
row = 3
Do While Workbooks(VS).Sheets("MATERIALS").Range("J" & row) <> ""
    row = row + 1
Loop
'4
MyProgressBar.NextAction "Clearing customers"
'clear customers
Workbooks(VS).Sheets("MATERIALS").Range("J3:J" & row - 1).ClearContents

'----------------------------
'find bottom row for engineers
row = 3
Do While Workbooks(VS).Sheets("MATERIALS").Range("L" & row) <> ""
    row = row + 1
    If row = 3 Then
        row = 4
    End If
Loop
'5
MyProgressBar.NextAction "Clearing engineers"
'clear engineers
Workbooks(VS).Sheets("MATERIALS").Range("L3:L" & row - 1).ClearContents
'=========================================================================

'6
MyProgressBar.NextAction "Transferring material info"
'transfer material info
row = 3
Do While Workbooks(CSV).Sheets(1).Range("A" & row) <> ""
    Workbooks(VS).Sheets("MATERIALS").Range("B" & row) = Workbooks(CSV).Sheets(1).Range("A" & row) 'name
    Workbooks(VS).Sheets("MATERIALS").Range("C" & row) = Workbooks(CSV).Sheets(1).Range("B" & row) 'SAP#
    Workbooks(VS).Sheets("MATERIALS").Range("D" & row) = Workbooks(CSV).Sheets(1).Range("C" & row) 'SG
    Workbooks(VS).Sheets("MATERIALS").Range("E" & row) = Workbooks(CSV).Sheets(1).Range("D" & row) 'sack weight
    Workbooks(VS).Sheets("MATERIALS").Range("F" & row) = Workbooks(CSV).Sheets(1).Range("E" & row) 'bulk density
    row = row + 1
Loop
'7
MyProgressBar.NextAction "Transferring rigs"
'transfer rigs
row = 3
Do While Workbooks(CSV).Sheets(1).Range("F" & row) <> ""
    Workbooks(VS).Sheets("MATERIALS").Range("H" & row) = Workbooks(CSV).Sheets(1).Range("F" & row)
    row = row + 1
Loop

'8
MyProgressBar.NextAction "Transferring customers"
'transfer customers
row = 3
Do While Workbooks(CSV).Sheets(1).Range("G" & row) <> ""
    Workbooks(VS).Sheets("MATERIALS").Range("J" & row) = Workbooks(CSV).Sheets(1).Range("G" & row)
    row = row + 1
Loop
'9
MyProgressBar.NextAction "Transferring engineers"
'transfer engineers
row = 3
Do While Workbooks(CSV).Sheets(1).Range("H" & row) <> ""
    Workbooks(VS).Sheets("MATERIALS").Range("L" & row) = Workbooks(CSV).Sheets(1).Range("H" & row)
    row = row + 1
Loop

'10
MyProgressBar.NextAction "Closing .csv file"

Workbooks(CSV).Saved = True
Workbooks(CSV).Close

'11
MyProgressBar.NextAction "Updating additive validations"
Call AddValidation

'12
MyProgressBar.NextAction "Updating rig validation"
Call RigValidation

'13
MyProgressBar.NextAction "Updating customer validation"
Call CustomerValidation

'14
MyProgressBar.NextAction "Updating fluid material validations"
Call SpacerValidation

'15
MyProgressBar.NextAction "Calculate full rebuild"
Application.CalculateFullRebuild

MyProgressBar.Complete 2

Call PswrdLock
Exit Sub

errhandler:
MsgBox ("Excel cannot connect to Halliburton network. Sheet is using " & Sheets("MATERIALS").Range("B1") & " materials that have not been updated.")


MyProgressBar.Terminate
Call PswrdLock
Exit Sub

End Sub
