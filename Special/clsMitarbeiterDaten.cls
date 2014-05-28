VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMitarbeiterDaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mclElements As New clsMyCollection 'inhalt key - value
Public sName As String
Private f As New clsFuncs

Public Function init(Optional sNameP As String = "")
    sName = sNameP
End Function

Public Function add(vValue As Variant, sKey As String, Optional stype As String = "Aufsummieren")
    'falls value schon vorhanden, addieren
    If mclElements.isInArray(sKey) Then
        If LCase(stype) = LCase("Aufzaehlen") Then
            Call mclElements.setValue(CLng(mclElements.getValue(sKey)) + 1, sKey)
        ElseIf LCase(stype) = LCase("Aufsummieren") Then
            Call mclElements.setValue(CLng(mclElements.getValue(sKey)) + CLng(vValue), sKey)
        End If
    Else ' sonst value erstellen
        Call mclElements.add(vValue, sKey)
    End If
End Function

Public Function getValue(sKey)
    getValue = mclElements.getValue(sKey)
End Function

Public Function setValue(vValue As Variant, sKey As String)
    Call mclElements.setValue(vValue, sKey)
End Function

Public Function getSize()
    getSize = mclElements.getSize()
End Function

Public Function getKeys()
    getKeys = mclElements.getKeys
End Function
