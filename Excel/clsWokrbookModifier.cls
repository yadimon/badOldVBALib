VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWokrbookModifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsWokrbookModifier
' Author    : Dmitry Gorelenkov
' Date      : 30.01.2013
' Purpose   : some modifications by any workbook
' Requires  :
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------

Option Explicit
Private f As clsFuncs
'Private bInitDone As Boolean
'
'Public Function init() As Boolean
'    Dim bResult As Boolean
'
'    bInitDone = True
'
'    init = bResult
'End Function

'einige werte ersetzen, vOldval - Array alte Werte, vNewVals - Array neue Werte, rngWhereToReplace - range wo ersetzt wird
Public Function replaceValues(vOldval As Variant, vNewVals As Variant, rngWhereToReplace As range)
    Dim rngCell As range
    Dim idx As Long
    
    For Each rngCell In rngWhereToReplace
        If f.isInArray(vOldval, rngCell.Value, True, True, True) Then
            idx = f.getIndexInArrayByValue(rngCell.Value, vOldval)
            If idx >= 0 Then
                rngCell.Value = vNewVals(idx)
            End If
        End If
    Next rngCell
End Function


Private Sub Class_Initialize()
    Set f = New clsFuncs
End Sub

Private Sub Class_Terminate()
    Set f = Nothing
End Sub

