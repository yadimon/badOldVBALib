VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReportingData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsReportingData
' Author    : Dmitry Gorelenkov
' Date      : 11.11.2013
' Changed   : 11.11.2013
' Purpose   : Object that keeps data from reporting in the defined format
' Requires  : clsMyCollection, clsRepWorkerData
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Explicit
'TODO get, set, useful funcs

Public workers As clsMyCollection 'workers data saved in collection of clsRepWorkerData's
'Public totals() As clsMyCollection 'total worksheet data saved clsRepWorkerData's array

Public Function getWorkerByIndex(lIndex As Long) As clsRepWorkerData
    Set getWorkerByIndex = Nothing
    
    If workers.getSize > lIndex Then
        Set getWorkerByIndex = workers.getValue(lIndex)
    End If
End Function

Public Function getWorkerByName(sName As String) As clsRepWorkerData
    Set getWorkerByName = Nothing
    
    If Len(sName) > 0 Then
        Set getWorkerByName = workers.getValue(sName)
    End If
    
End Function
'add new worker
Public Function addWorker(wrkWorker As clsRepWorkerData)
    Call Me.workers.add(wrkWorker, wrkWorker.name)
End Function

Private Sub Class_Initialize()
    Set workers = New clsMyCollection
'    Set totals = New clsMyCollection
End Sub

