VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsErrorHandler
' Author    : Dmitry Gorelenkov
' Date      : 13.08.2012
' Changed   : 29.10.2013
' Purpose   : Fehler abfangen und ausgeben
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------


Public Sub catchErr(Optional sProcName As String = "")
    Debug.Print "--------------------"
    Debug.Print Format(Now(), "dd.mm.yyyy, hh:nn:ss")
    Debug.Print "Fehler Source: " & err.Source
    
    If Len(Trim(sProcName)) > 0 Then
        Debug.Print "Fehler Source Proc: " & sProcName
    End If
    
    If Erl > 0 Then
        Debug.Print "Line: " & Erl
    End If
    
    Debug.Print "Fehler Beschreibung: " & err.Description
    Debug.Print "Fehler Nummer: " & err.Number
    Debug.Print "--------------------"
    'Debug.Print Error(Err)
End Sub
