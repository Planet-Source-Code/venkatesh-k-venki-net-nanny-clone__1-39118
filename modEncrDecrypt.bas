Attribute VB_Name = "modencrdecrypt"
'encrypting and decrypting module
Public Function DECR(dncr1 As String)
    Dim T1 As String
On Error Resume Next
For i = 1 To Len(dncr1)
    T1 = Mid$(dncr1, i, 1)
    Mid$(dncr1, i, 1) = Chr(Asc(T1) - 101)
Next i
    DECR = dncr1
End Function

Public Function ENCR(encr1 As String)
Dim T1 As String
On Error Resume Next
For i = 1 To Len(encr1)
    T1 = Mid$(encr1, i, 1)
    Mid$(encr1, i, 1) = Chr(Asc(T1) + 101)
Next i
    
    ENCR = encr1
End Function

