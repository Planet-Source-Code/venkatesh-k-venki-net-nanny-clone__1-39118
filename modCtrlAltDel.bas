Attribute VB_Name = "modCtrlAltDel"
'This module is to hide ur program from
'Ctrl+Alt+Del list
'works in win 98,95 etc
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0

Public Sub RemoveCtrlAltDel()
On Error Resume Next
    Dim lngPrID As Long
    Dim lngRn As Long
    
    lngPrID = GetCurrentProcessId()
    lngRn = RegisterServiceProcess(lngPrID, RSP_SIMPLE_SERVICE)
End Sub
