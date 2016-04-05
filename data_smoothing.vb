Sub try2()
Dim row, lead, tail, stage, i As Integer
Dim lead_value, tail_value, var, test As Double
Dim out As Boolean

Dim avg, bbl, sum As Double
Dim count As Integer

i = 4
Do While Range("J" & i) <> ""
    Range("J" & i & ":N" & i).ClearContents
    i = i + 1
Loop

var = 0.5

'initial conditions
lead_value = Range("B4")
lead = 5
stage = 1

row = lead
Do While Range("B" & row) <> ""
    test = Range("B" & row)
    
'------------------------------------------------
    'determine if test value is outside variance
    out = False
    
    If test > (lead_value + var) Then
        out = True
    End If
    
    If test < (lead_value - var) Then
        out = True
    End If
'------------------------------------------------

If out = True Then
    
    tail = row - 1
    sum = 0
    count = 0
    
'find average comb pump rate
    For i = lead To tail
        sum = sum + Range("B" & i)
        count = count + 1
        
    Next i
    avg = sum / count
    
    
'find difference in bbl from lead-tail row range
    bbl = Range("C" & tail) - Range("C" & lead - 1)
    
    
'if row rate is unique then use instantaneous pump stage total
    If lead = tail Then
        bbl = Range("C" & lead) - Range("C" & lead - 1)
        count = 1
    End If

'populate events schedule
    Range("J" & stage + 3) = stage
    Range("K" & stage + 3) = avg
    Range("L" & stage + 3) = bbl
    Range("M" & stage + 3) = (tail - lead) / 60
    
'if avg pump rate > 0 and pump stage total = 0 then data point is an instantaneous change without volume
'need to overrite in event schedule
If avg > 0 And bbl = 0 Then
    stage = stage - 1
End If
    
'account for unique rate
    If lead = tail Then
    Range("M" & stage + 3) = 1 / 60
    End If
    
'show lead and tail data rows
    'Range("N" & stage + 3) = lead
    'Range("O" & stage + 3) = tail
    
'increase stage
    stage = stage + 1
    
'establish next lead value
    lead = tail + 1
    lead_value = Range("B" & lead)


End If
 
'move to next row
row = row + 1
Loop

'for final row in data stream
tail = row - 1
sum = 0
count = 0
    
        
For i = lead To tail
    sum = sum + Range("B" & i)
    count = count + 1
        
Next i
    
avg = sum / count
bbl = Range("C" & tail) - Range("C" & lead - 1)
    
If lead = tail Then
    bbl = Range("C" & lead) - Range("C" & lead - 1)
    count = 1
End If

Range("J" & stage + 3) = stage
Range("K" & stage + 3) = avg
Range("L" & stage + 3) = bbl
Range("M" & stage + 3) = 1 / 60
'Range("N" & stage + 3) = lead
'Range("O" & stage + 3) = tail



End Sub
