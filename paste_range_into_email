Sub Mail_Selectiong_Range_Outlook_Body()
Dim row As Integer
Dim rng As Range
Dim OutApp As Object
Dim OutMail As Object
Dim Subj, MailTo As String
Dim wb1 As Workbook

Set rng = Nothing

row = 2
Do While Sheets("UPDATE").Range("A" & row) <> ""
row = row + 1
Loop
row = row - 1

Set rng = Sheets("UPDATE").Range("A1:J" & row)
Subj = Sheets("Params").Range("D2")
Subj = Subj & " - " & Sheets("Params").Range("D3")
Subj = Subj & " - " & Sheets("Params").Range("D4")

'find mailto list
row = 6
Do While Sheets("Params").Range("D" & row) <> ""
MailTo = MailTo & Sheets("Params").Range("D" & row) & "; "
row = row + 1
Loop

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

'create temporary attachment file
If UCase(Sheets("Params").Range("D5")) = "YES" Then
    Set wb1 = ActiveWorkbook
    TempFilePath = Environ$("temp") & "\"
    TempFileName = Sheets("Params").Range("D2") & " " & Sheets("Params").Range("D3") & ".xlsm"
    wb1.SaveCopyAs TempFilePath & TempFileName

With OutMail
    .To = MailTo
    .CC = ""
    .BCC = ""
    .Subject = Subj
    .HTMLBody = RangetoHTML(rng)
    .Attachments.Add TempFilePath & TempFileName
    .Display
End With

Kill TempFilePath & TempFileName

Else

With OutMail
    .To = MailTo
    .CC = ""
    .BCC = ""
    .Subject = Subj
    .HTMLBody = RangetoHTML(rng)
    .Display
End With
End If

Set OutMail = Nothing
Set OutApp = Nothing

End Sub

Function RangetoHTML(rng As Range)
Dim fso As Object
Dim ts As Object
Dim TempFile As String
Dim TempWB As Workbook

TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm,yy hh-mm-ss") & ".htm"

rng.Copy
Set TempWB = Workbooks.Add(1)

With TempWB.Sheets(1)
    .Cells(1).PasteSpecial Paste:=8
    .Cells(1).PasteSpecial xlPasteValues, , False, False
    .Cells(1).PasteSpecial xlPasteFormats, , False, False
    .Cells(1).Select
    Application.CutCopyMode = False
    On Error Resume Next
    .DrawingObjects.Visible = True
    .DrawingObjects.Delete
    On Error GoTo 0
End With

With TempWB.PublishObjects.Add( _
    SourceType:=xlSourceRange, _
    Filename:=TempFile, _
    Sheet:=TempWB.Sheets(1).Name, _
    Source:=TempWB.Sheets(1).UsedRange.Address, _
    HtmlType:=xlHtmlStatic)
    .Publish (True)
End With
    
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
RangetoHTML = ts.ReadAll
ts.Close
RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
    "align=left x:publishsource=")

TempWB.Close savechanges:=False
Kill TempFile
    
Set ts = Nothing
Set fso = Nothing
Set TempWB = Nothing

End Function
