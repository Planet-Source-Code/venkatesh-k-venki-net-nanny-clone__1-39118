Attribute VB_Name = "modWriteReadINI"
Public Function writeINI(inString As String)
    Open App.Path & "\" & "config.ini" For Output As #1
        Print #1, inString
    Close #1
End Function

Public Function readINI(mn As String, fst As String, snd As String, pos As Long) As String
    'searching the ini file for the string
    Dim st As Long
    Dim en As Long
    Dim res As String
    st = InStr(pos, mn, fst, vbBinaryCompare)
    en = InStr(st + Len(fst), mn, snd, vbBinaryCompare)
    
    If st <> 0 And en <> 0 Then
        pos = en
        res = Mid(mn, st + Len(fst), en - st - Len(fst))
        readINI = res
    Else
        pos = 0
    End If
    If st <> 0 And en = 0 Then
        pos = 0
        res = ""
        readINI = res
    End If
End Function

