Sub CellsLock(bool)
'This procedure locks or unlocks all the cells that need to be
'adjusted by engineering. The main boolean input "bool"
'determines whether the .Locked key is true or false. Note that
'the PswrdUnlock/PswrdLock sub is included in this procedure.

Dim row, col, Sht As Integer
Dim Key As Boolean
Dim spcr As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'On Error GoTo ErrHandler

'xxxxxxxxxxxxxxxxxxxxxxxxxxxx
Call PswrdUnlock
'xxxxxxxxxxxxxxxxxxxxxxxxxxxx

If bool = "Lock" Or bool = "lock" Then
    Key = True
End If

If bool = "UnLock" Or bool = "unlock" Then
    Key = False
End If

'date, customer, job, district
For row = 2 To 5
    Sheets(1).Range("G" & row & ":H" & row).Locked = Key
    Sheets(1).Range("P" & row & ":Q" & row).Locked = Key
Next row

'so
Sheets(1).Range("J2").Locked = Key

'Slurry Count
Sheets(1).Range("Z8").Locked = Key

'Water SG
Sheets(1).Range("Z11").Locked = Key

'Additional Fluids
For row = 7 To 12
    Sheets(1).Range("AD" & row).Locked = Key
Next row

'XXXXXXXXXXXXXXXXXXXXXXXXXXXX
'Loop through loadsheets
'XXXXXXXXXXXXXXXXXXXXXXXXXXXX

For Sht = 1 To 4
    
    'Cement Names
    Sheets(Sht).Range("E7:J7").Locked = Key
    Sheets(Sht).Range("N7:S7").Locked = Key
    
    'lead cement
    For col = 7 To 8
        For row = 17 To 26
            Sheets(Sht).Cells(row, col).Locked = Key
        Next row
    Next col
    
    'lead additives
    For col = 7 To 9
        For row = 30 To 39
            Sheets(Sht).Cells(row, col).Locked = Key
        Next row
    Next col
    
    'lead blend/side
    
    For row = 30 To 39
        Sheets(Sht).Cells(row, 11).Locked = Key
    Next row
    
    'tail cement
    For col = 16 To 17
        For row = 17 To 26
            Sheets(Sht).Cells(row, col).Locked = Key
        Next row
    Next col
    
    'tail additives
    For col = 16 To 18
        For row = 30 To 39
            Sheets(Sht).Cells(row, col).Locked = Key
        Next row
    Next col
    
    'tail blend/side
    For row = 30 To 39
        Sheets(Sht).Cells(row, 20).Locked = Key
    Next row
    
    'density and sacks lead
    Sheets(Sht).Range("F10").Locked = Key
    Sheets(Sht).Range("J10").Locked = Key
    
    'density and sacks tail
    Sheets(Sht).Range("O10").Locked = Key
    Sheets(Sht).Range("S10").Locked = Key

    'type & ifacts - Lead
    Sheets(Sht).Range("F8").Locked = Key
    Sheets(Sht).Range("J8").Locked = Key
    
    'type & ifacts - Tail
    Sheets(Sht).Range("O8").Locked = Key
    Sheets(Sht).Range("S8").Locked = Key
    
    'Engineering Approval
    For row = 46 To 49
        Sheets(Sht).Cells(row, 5).Locked = Key
    Next row
    
    'Goal Seek Options
    Sheets(Sht).Range("X46").Locked = Key
    Sheets(Sht).Range("AG46").Locked = Key
    
Next Sht

'XXXXXXXXXXXXXXXXXXXXXXXXXXXX
'Additional Fluids
'XXXXXXXXXXXXXXXXXXXXXXXXXXXX

'----------------------------
'TUNED SPACER III
spcr = "TS III"

col = 3
Do While Sheets(spcr).Cells(2, col) <> ""
    
    'density
    Sheets(spcr).Cells(8, col + 1).Locked = Key
    'yp
    Sheets(spcr).Cells(9, col + 1).Locked = Key
    'volume
    Sheets(spcr).Cells(14, col + 1).Locked = Key

    'liquid adds
    For row = 24 To 27
        Sheets(spcr).Cells(row, col).Locked = Key
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    
    'dry adds
    For row = 31 To 33
        Sheets(spcr).Cells(row, col).Locked = Key
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    col = col + 13
Loop

'----------------------------
'CLEAN SPACER III
'----------------------------
spcr = "CS III"

col = 3
Do While Sheets(spcr).Cells(2, col) <> ""
    'density
    Sheets(spcr).Cells(8, col + 1).Locked = Key
    
    'volume
    Sheets(spcr).Cells(12, col + 1).Locked = Key
    
    'weighting agent selection
    Sheets(spcr).Cells(15, col).Locked = Key
    
    'SA-1015 loading
    Sheets(spcr).Cells(33, col + 1).Locked = Key
    
    'liquid additives
    For row = 22 To 25
        Sheets(spcr).Cells(row, col).Locked = Key
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    
    'dry additives
    For row = 28 To 30
        Sheets(spcr).Cells(row, col).Locked = Key
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    col = col + 11
Loop

'----------------------------
'WATER-MATERIAL
'----------------------------
spcr = "WATER-MATERIAL"

'name
Sheets(spcr).Range("C2:E2").Locked = Key
Sheets(spcr).Range("N2:P2").Locked = Key
Sheets(spcr).Range("Y2:AA2").Locked = Key

col = 3
Do While Sheets(spcr).Cells(2, col) <> ""
    'density
    Sheets(spcr).Cells(8, col + 1).Locked = Key
    'volume
    Sheets(spcr).Cells(10, col + 1).Locked = Key
    
    For row = 13 To 19
        Sheets(spcr).Cells(row, col).Locked = Key
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    For row = 23 To 29
        Sheets(spcr).Cells(row, col).Locked = Key
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row

    col = col + 11
Loop

'----------------------------
'TERGOVIS
'----------------------------
spcr = "TERGOVIS"

col = 3
Do While Sheets(spcr).Cells(2, col) <> ""
    'density
    Sheets(spcr).Cells(8, col + 1).Locked = Key
    'volume
    Sheets(spcr).Cells(12, col + 1).Locked = Key
    'enhancer
    Sheets(spcr).Cells(15, col).Locked = Key
  
    'materials
    For row = 14 To 17
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    
    'yield, wr
    For row = 19 To 20
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    
    'liquid add
    For row = 24 To 27
        Sheets(spcr).Cells(row, col).Locked = Key
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row

    'dry add
    For row = 30 To 32
        Sheets(spcr).Cells(row, col).Locked = Key
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    
    'lb/bbl
    For row = 35 To 38
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    
    'lab weigh up
    For row = 41 To 43
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row

    col = col + 6
Loop

'----------------------------
'POZ-SCAVENGER
'----------------------------
spcr = "POZ-SCAVENGER"

col = 3
Do While Sheets(spcr).Cells(2, col) <> ""
    'density
    Sheets(spcr).Cells(9, col + 1).Locked = Key
    
    'volume
    Sheets(spcr).Cells(13, col + 1).Locked = Key
    
    'materials
    For row = 15 To 18
        Sheets(spcr).Cells(row, col).Locked = Key
        Sheets(spcr).Cells(row, col + 1).Locked = Key
        Sheets(spcr).Cells(row, col + 2).Locked = Key
    Next row
    
    'lab weigh up
    For row = 32 To 34
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    
    col = col + 10
Loop

'----------------------------
'SWEEP
'----------------------------
spcr = "SWEEP"

col = 3
Do While Sheets(spcr).Cells(2, col) <> ""
    'density
    Sheets(spcr).Cells(9, col + 1).Locked = Key
    
    'volume
    Sheets(spcr).Cells(11, col + 1).Locked = Key
    
    For row = 14 To 20
        Sheets(spcr).Cells(row, col).Locked = Key
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    
    For row = 24 To 30
        Sheets(spcr).Cells(row, col).Locked = Key
        Sheets(spcr).Cells(row, col + 1).Locked = Key
    Next row
    
    col = col + 10
Loop

Call PswrdLock

End Sub
