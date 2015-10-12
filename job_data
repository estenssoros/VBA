Sub emailINT()
Dim strto, body, subject As String
Dim name, number, rig As String
Dim i As Integer
Dim wb1 As Workbook
Dim TempFilePath As String
Dim TempFileName As String
Dim FileExtStr As String

name = Sheets("Intermediate").Range("C3")
number = Sheets("Intermediate").Range("C4")
rig = Sheets("Intermediate").Range("C5")

Range("O2") = Now()
If name = "" Then
    MsgBox ("Please enter well name.")
    UserForm1.Show
    Exit Sub
End If

If number = "" Then
    MsgBox ("Please enter well number.")
    UserForm1.Show
    Exit Sub
End If

If rig = "" Then
    MsgBox ("Please enter rig.")
    UserForm1.Show
    Exit Sub
End If

If Range("D9") = "" Then
    MsgBox ("Please enter surface depth.")
    UserForm2.Show
    Exit Sub
End If

If Range("D31") = "" Then
    MsgBox ("Please enter intermediate depth.")
    UserForm2.Show
    Exit Sub
End If

If Range("N16") = "" Then
    MsgBox ("Please enter hole excesss.")
    UserForm2.Show
    Exit Sub
End If

If Range("N26") = "" Then
    MsgBox ("Please enter anticipated mud density.")
    UserForm2.Show
    Exit Sub
End If

i = 2
Do While Sheets("LISTS").Range("D" & i) <> ""
    strcc = strcc & "; " & Sheets("LISTS").Range("D" & i)
    
    i = i + 1
Loop


i = 2
Do While Sheets("LISTS").Range("F" & i) <> ""
    If Sheets("LISTS").Range("F" & i) = rig Then
        strto = strto & "; " & Sheets("LISTS").Range("G" & i)
        strto = strto & "; " & Sheets("LISTS").Range("I" & i)
    End If
i = i + 1
Loop
        

'subject
subject = "Noble " & name & " " & number & " Intermediate Cement Job on " & rig

'Intro
body = "<FONT size = 4> All, <br><br>"
body = body & "Could you please confirm the attached information for the upcoming <b>" & name & " " & number & "</b> intermediate cement job on <b>" & rig & ". </b>"
body = body & "Please <b><u><FONT color = red>REPLY ALL</FONT></b></u> when finished."
body = body & "<br><br><br>"

'Surface
body = body & "<b><u>SURFACE CASING</u></b><ul>"
body = body & "<li>Depth: <FONT color = #1F497D>" & Range("D9") & "'</FONT></li>"
body = body & "<li>Size: <FONT color = #1F497D>9 5/8''</FONT></li>"
body = body & "<li>Casing Weight: <FONT color = #1F497D>36#</FONT></li>"
body = body & "<li>Casing Thread Type: <FONT color = #1F497D>P-110IC LTC/BTC</FONT></li></ul>"

'Intermediate
body = body & "<b><u>INTERMEDIATE CASING</u></b><ul>"
body = body & "<li>Estimated MD: <FONT color = #1F497D>" & Range("D31") & "'</FONT></li>"
body = body & "<li>Casing Size: <FONT color = #1F497D>7''</FONT></li>"
body = body & "<li>Casing Weight: <FONT color = #1F497D>26#</FONT></li>"
body = body & "<li>Casing Thread Type: <FONT color = #1F497D>P-110IC LTC/BTC</FONT></li>"
body = body & "<li>Shoe Track: <FONT color = #1F497D>42'</FONT></li></ul>"
    
'Hole
body = body & "<b><u>INTERMEDIATE OPEN HOLE</u></b><ul>"
body = body & "<li>Bit Size: <FONT color = #1F497D>" & Range("N10") & "''</FONT></li>"
body = body & "<li>Excess: <FONT color = #1F497D>" & Format(Range("N16"), "0%") & "</FONT></li></ul>"

'Mud
body = body & "<b><u>MUD @ Anticipated TD</u></b><ul>"
body = body & "<li>Weight: <FONT color = #1F497D>" & Range("N26") & "#/gal</FONT></li>"
body = body & "<li>PV/YP: <FONT color = #1F497D>" & Range("N27") & " cP / " & Range("N28") & " lb/100ft&sup2</FONT></li>"
body = body & "<li>Oil/Water Based: <FONT color = #1F497D>" & Range("N25") & "</FONT></li></ul>"

'Cement
body = body & "<b><u>CEMENT INFORMATION</u></b><ul>"
body = body & "<li>Primary: <FONT color = #1F497D><u><b>" & Range("I35") & "</b></u></FONT> sacks</u></li><ul>"
body = body & "<li>Density: <FONT color = #1F497D>" & Range("N20") & " </FONT>lb/gal</li>"
body = body & "<li>Yield: <FONT color = #1F497D>" & Range("N21") & " </FONT>ft3/sk</li>"
body = body & "<li>Water Requirement: <FONT color = #1F497D>" & Range("N22") & " </FONT>gal/sk</li>"
body = body & "<li>Top of Cement: <FONT color = #1F497D>" & Range("D11") & "</FONT>'</li></ul></ul>"

'Spacer
body = body & "<b><u>SPACER INFORMATION</u></b><ul>"
body = body & "<li>Amount: <FONT color = #1F497D>" & Range("D34") & "</FONT> bbl (Clean or Tuned): <FONT color = #1F497D>Tuned</FONT></li>"
body = body & "<li>Weight: <FONT color = #1F497D>12.5 #/gal</FONT></li>"
body = body & "<li>Poly-E-Flake: <FONT color = #1F497D> 0.5 #/bbl </FONT></li></ul><br>"

body = body & "Thank You, <br><br><b><i>Halliburton </b></i>"

Sheets("Production Liner").Visible = xlVeryHidden
Set wb1 = ActiveWorkbook

TempFilePath = Environ$("temp") & "\"
TempFileName = "Noble " & name & " " & number & " Intermediate Cement " & rig & ".xlsm"

wb1.SaveCopyAs TempFilePath & TempFileName

    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .to = strto
        .CC = strcc
        .BCC = " "
        .subject = subject
        .HTMLBody = body
        .attachments.Add TempFilePath & TempFileName
        .Display
    End With
   
    Kill TempFilePath & TempFileName
   
    Set OutMail = Nothing
    Set OutApp = Nothing


End Sub
