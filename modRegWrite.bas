Attribute VB_Name = "modRegWrite"
'module to create registry key

Public Sub createregistrykey(keystring As String, value As String)
    Dim obj As Object
    On Error Resume Next
    Set obj = CreateObject("wscript.shell")
    obj.RegWrite keystring, value
End Sub

