VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReportingHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsReportingHandler
' Author    : Dmitry Gorelenkov
' Date      : 08.11.2013
' Changed   : 08.11.2013
' Purpose   : handler for typical agency reporting
' Requires  : clsErrorHandler, IReportingHandler_AgencyFuncs, clsMyCollection, clsReportingData
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------


Option Explicit

Private m_bInitDone As Boolean
Private m_AgencyFunction As IReportingHandler_AgencyFuncs
Private m_mcDataCollection As clsMyCollection
Private ehError As clsErrorHandler
Private m_mcWorkers As clsMyCollection
Private m_rdReportingData As clsReportingData

Public Function init(agncAgency As IReportingHandler_AgencyFuncs) As Boolean
    On Error GoTo init_Error

    If m_bInitDone Then Exit Function
    
    Set m_AgencyFunction = agncAgency
    Set m_rdReportingData = m_AgencyFunction.getReportingData

    init = True
    m_bInitDone = True

    Exit Function
init_Error:
    ehError.catchErr ("init of Klassenmodul clsReportingHandler")
    init = False
    m_bInitDone = False
End Function


'get collection of workers, key as name, clsRepWorkerData as value
Public Function getWorkers() As clsMyCollection
    Set getWorkers = m_rdReportingData.workers
End Function

'"gesamt" als clsRepWorkerData zurueckgeben
Public Function getTotal() As clsRepWorkerData

End Function

'reload data
Public Function realoadData() As Boolean
    m_bInitDone = False
    realoadData = init(m_AgencyFunction)
End Function


Private Sub Class_Initialize()
    Set ehError = New clsErrorHandler
End Sub

Private Sub Class_Terminate()
    Set ehError = Nothing
    Set m_mcWorkers = Nothing
End Sub
