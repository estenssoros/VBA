Function GetNthWord(ByVal WhichText As String, _
                    ByVal WhichWord As Long, _
                    Optional ByVal Seperator As String = " ") As Variant

WhichText = Trim(WhichText)

Dim StrAns As String
 
If InStr(1, WhichText, Seperator) = 0 Then
    GetNthWord = CVErr(xlErrValue)
    Exit Function
End If
 
StrAns = Trim(VBA.Split(WhichText, Seperator, , _
                vbBinaryCompare)(WhichWord - 1))
 
If IsNumeric(StrAns) Then
    GetNthWord = CDbl(StrAns)
Else
    GetNthWord = StrAns
End If
 
End Function

Function NumberOfWords(ByVal WhichText As String, _
                    Optional ByVal Seperator As String = " ") As Variant

WhichText = Trim(WhichText)
 
If InStr(1, WhichText, Seperator) = 0 Then
    NumberOfWords = CVErr(xlErrValue)
    Exit Function
End If
 
'NumberOfWords = CLng(UBound(VBA.Split(WhichText, Seperator, _
                    , vbBinaryCompare)) + 1)

NumberOfWords = (Len(WhichText) - _
                    Len(Replace(WhichText, Seperator, vbNullString, 1, , _
                    vbBinaryCompare))) / Len(Seperator) + 1
 
End Function