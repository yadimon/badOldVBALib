Attribute VB_Name = "mdlTapi"
Option Compare Database
Option Explicit

Declare Function TAPI_Make_Call Lib _
"tapi32.dll" Alias "tapiRequestMakeCall" _
(ByVal stNumber As String, _
ByVal stDummy1 As String, _
ByVal stDummy2 As String, _
ByVal stDummy3 As String) As Long



Public Function tapi_call(sNumber As Variant)
'    If IsNumeric(sNumber) Then
'        sNumber = CStr(sNumber)
'    Else
'        Debug.Print "tapi_call falsche Parameter"
'        Exit Function
'    End If
    TAPI_Make_Call CStr(sNumber), "", "", ""
End Function

