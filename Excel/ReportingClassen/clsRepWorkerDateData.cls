VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRepWorkerDateData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsRepWorkerDateData
' Author    : Dmitry Gorelenkov
' Date      : 11.11.2013
' Changed   : 11.11.2013
' Purpose   : daten von einem mitarbeiter von bestimmter Datum
' Requires  :
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Explicit
'TODO properties, init()

Public Projekt_Stunden As Double
Public ThisDate As Date

Private Sub Class_Initialize()
    ThisDate = CDate(0)
End Sub

