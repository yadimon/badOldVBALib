VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsPerformance
' Author    : Dmitry Gorelenkov
' Date      : 13.08.2012
' Purpose   : Ausfuerungszeit messen
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------


Option Explicit

Private timersCollection As New Collection

'zeit messen.
'Anwendung:
'performance("bemerkung", 0)
' %code%
'performance("bemerkung", 1)
Sub performance(what As String, bStartEnd As Boolean)
    On Error Resume Next
    'falls beginn
    If Not bStartEnd Then
        Call timersCollection.remove(what)
        Call timersCollection.add(Timer, what)
    'sonst
    Else
        Debug.Print "Ausführen: " & what & " -  " & (Timer - CDbl(timersCollection(what))) & " sekunden"
        Call timersCollection.remove(what)
    End If
End Sub
