VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRepWorkerSheetParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsRepWorkerSheetParser
' Author    : Dmitry Gorelenkov
' Date      : 12.11.2013
' Changed   : 22.11.2013
' Purpose   : fills worker object with default data from Worksheet
' Requires  : clsRepWorkerData, clsRepWorkerDateData, clsFuncsExl, clsFuncs
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Explicit

Private fXl As clsFuncsExl
Private f As clsFuncsExl

'fill data from worksheet of usual reporting to worker object
Public Function fillDataToWorker(ByRef wksWorkerWorksheet As Worksheet, ByRef wrkNewWorker As clsRepWorkerData)
    Dim i As Long
    Dim lFistRow As Long
    Dim lLastRow As Long
    Dim lProjektStundenCol As Long
    Dim sDateValue As String


    With wksWorkerWorksheet
        
        'create and add date objects
        lProjektStundenCol = fXl.LetterToCol("H")
        lFistRow = fXl.getFirstRow(wksWorkerWorksheet, 1)
        lLastRow = fXl.getLastRow(wksWorkerWorksheet, 1)
        'zeilen mit datum
        For i = lFistRow To lLastRow
            sDateValue = .Cells(i, 1).value
            'falls datum
            If f.isDateMy(sDateValue) Then
                Dim DateObject As clsRepWorkerDateData
                Set DateObject = New clsRepWorkerDateData
                DateObject.ThisDate = f.CdateTry(sDateValue)
                DateObject.Projekt_Stunden = f.ifNoValue(.Cells(i, lProjektStundenCol).value, 0)
                Call wrkNewWorker.addDateData(DateObject)
                
            End If
        Next i
    End With

End Function

Private Sub Class_Initialize()
    Set fXl = New clsFuncsExl
    Set f = New clsFuncs
End Sub

Private Sub Class_Terminate()
    Set fXl = Nothing
    Set f = Nothing
End Sub
