VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRepHandlerFuncsKikxxl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsReportingHandlerSubKikxxl
' Author    : Dmitry Gorelenkov
' Date      : 08.11.2013
' Changed   : 08.11.2013
' Purpose   : Kikxxl-agency extra function class, for clsReportingHandler
' Requires  : clsFuncsExl, clsReportingData, clsRepWorkerSheetParser, clsErrorHandler
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Explicit
Implements IReportingHandler_AgencyFuncs

Private Const m_sWorkerWksName As String = "MA 1"
Private Const m_sAgenturName As String = "Kikxxl"


Private m_wkbReporting As Workbook
Private m_sPathToWorkersFolder As String
Private fXl As clsFuncsExl
Private ehError As clsErrorHandler
Private m_worksheetParser As clsRepWorkerSheetParser

Private m_bInitDone As Boolean

Public Function init(wkbReporting As Workbook, sPathToWorkersFolder As String)
    If m_bInitDone Then Exit Function
    
    Set m_wkbReporting = wkbReporting
    m_sPathToWorkersFolder = sPathToWorkersFolder
    
    
    init = True
    m_bInitDone = True
End Function

Private Function IReportingHandler_AgencyFuncs_getReportingData() As clsReportingData
    Dim sFiles() As String
    Dim returnData As New clsReportingData
    Dim wkbWorkerWorkbook As Workbook
    Dim sPath As Variant
    On Error GoTo IReportingHandler_AgencyFuncs_getReportingData_Error

    sFiles = f.getDirList(m_sPathToWorkersFolder, "*.xl*")
    
    For Each sPath In sFiles
        'path file oeffnen
        Set wkbWorkerWorkbook = fXl.findOpenedDocOrOpen(CStr(sPath))
        'falls erfolgreich
        If Not wkbWorkerWorkbook Is Nothing Then
            Dim wrkWorker As clsRepWorkerData
            'worker erstellen
            Set wrkWorker = createWorkerByWorkbook(wkbWorkerWorkbook)
            'if parsed
            If Not wrkWorker Is Nothing Then
                'in reportingData hinzufuegen
                 Call returnData.addWorker(wrkWorker)
            Else
                f.dbg "Can not create Worker from " & wkbWorkerWorkbook.FullName
            End If
            
        End If
        
        wkbWorkerWorkbook.Close False
    Next sPath
    
    Set IReportingHandler_AgencyFuncs_getReportingData = returnData

    Exit Function
IReportingHandler_AgencyFuncs_getReportingData_Error:

    ehError.catchErr ("IReportingHandler_AgencyFuncs_getReportingData of Klassenmodul clsRepHandlerFuncsKikxxl")
End Function

'creates WorkerObject with data got from opended Workbook
Private Function createWorkerByWorkbook(wkbWorkerWorkbook As Workbook) As clsRepWorkerData
    Dim wrkNewWorker As New clsRepWorkerData
    Dim wksWorkerWorksheet As Worksheet
    Dim sWorkerFullName As String
    
    
    On Error GoTo createWorkerByWorkbook_Error

    Set wksWorkerWorksheet = wkbWorkerWorkbook.Sheets(m_sWorkerWksName)
    
    
    With wksWorkerWorksheet
        sWorkerFullName = .Range("A2").value & " " & .Range("A1").value
        'create worker object
        Call wrkNewWorker.init(sWorkerFullName, m_sAgenturName)
        'save data to worker
        Call m_worksheetParser.fillDataToWorker(wksWorkerWorksheet, wrkNewWorker)
        
    End With
    
    Set createWorkerByWorkbook = wrkNewWorker

    Exit Function
createWorkerByWorkbook_Error:
    Set createWorkerByWorkbook = Nothing
    ehError.catchErr ("createWorkerByWorkbook of Klassenmodul clsRepHandlerFuncsKikxxl")
End Function


Private Sub Class_Initialize()
    Set fXl = New clsFuncsExl
    Set ehError = New clsErrorHandler
    Set m_worksheetParser = New clsRepWorkerSheetParser
End Sub

Private Sub Class_Terminate()
    Set fXl = Nothing
    Set ehError = Nothing
    Set m_worksheetParser = Nothing
End Sub
