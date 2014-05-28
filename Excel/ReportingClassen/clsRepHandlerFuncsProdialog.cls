VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRepHandlerFuncsProdialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsRepHandlerFuncsProdialog
' Author    : Dmitry Gorelenkov
' Date      : 13.11.2013
' Changed   : 13.11.2013
' Purpose   : Prodialog-agency extra function class, for clsReportingHandler
' Requires  : clsFuncsExl, clsReportingData, clsRepWorkerSheetParser, clsErrorHandler
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------



Option Explicit
Implements IReportingHandler_AgencyFuncs

Private m_aNotWorkers As Variant
Private Const m_sStopIfName As String = "name"
Private Const m_sAgenturName As String = "Prodialog"
Private Const m_sNameAddress As String = "A2"

Private m_wkbReporting As Workbook

Private m_worksheetParser As clsRepWorkerSheetParser
Private fXl As clsFuncsExl
Private ehError As clsErrorHandler

Private m_bInitDone As Boolean

Public Function init(wkbReporting As Workbook)
    If m_bInitDone Then Exit Function
    
    Set m_wkbReporting = wkbReporting
    m_aNotWorkers = Array("Gesamt", "Gesamt PA", "Gesamt BS")
    
    init = True
    m_bInitDone = True
End Function

Private Function IReportingHandler_AgencyFuncs_getReportingData() As clsReportingData
    Dim sFiles() As String
    Dim returnData As New clsReportingData
    Dim wksWorkerWorksheet As Worksheet
    Dim sWorkerName As String
    
    
    On Error GoTo IReportingHandler_AgencyFuncs_getReportingData_Error

    For Each wksWorkerWorksheet In m_wkbReporting.Worksheets
        'falls nicht reservierter name
        If Not f.isInArray(m_aNotWorkers, wksWorkerWorksheet.name, True, True, True) Then
            'worker name
            sWorkerName = getWorkerNameByWorksheet(wksWorkerWorksheet)
            'falls den namen : "name" erreicht, stop
            If StrComp(sWorkerName, m_sStopIfName, vbTextCompare) = 0 Then
                Exit For
            End If
            
            'worker erstellen
            Dim wrkNewWorker As clsRepWorkerData
            Set wrkNewWorker = New clsRepWorkerData
            
            'initialize
            Call wrkNewWorker.init(sWorkerName, m_sAgenturName)
            'save data to worker
            Call m_worksheetParser.fillDataToWorker(wksWorkerWorksheet, wrkNewWorker)
            
            'if parsed
            If Not wrkNewWorker Is Nothing Then
                'in reportingData hinzufuegen
                 Call returnData.addWorker(wrkNewWorker)
            Else
                f.dbg "Can not create Worker from " & wksWorkerWorksheet.name & " worksheet"
            End If
        End If
        
    Next wksWorkerWorksheet
    
    Set IReportingHandler_AgencyFuncs_getReportingData = returnData

    Exit Function
IReportingHandler_AgencyFuncs_getReportingData_Error:

    ehError.catchErr ("IReportingHandler_AgencyFuncs_getReportingData of Klassenmodul clsRepHandlerFuncsProdialog")
End Function

'finden den MA namen
Private Function getWorkerNameByWorksheet(wksWorksheet As Worksheet) As String
    Dim sNameNotParsed As String
    Dim aNameVorname As Variant
    Dim sReturnName As String
    
    sNameNotParsed = wksWorksheet.Range(m_sNameAddress).value
    aNameVorname = Split(sNameNotParsed, ", ", 2, vbTextCompare)
    
    'vorname
    If f.arrayCheckIndex(aNameVorname, 1) Then
        sReturnName = aNameVorname(1)
    End If
    
    'nachname
    If f.arrayCheckIndex(aNameVorname, 0) Then
        sReturnName = sReturnName & " " & aNameVorname(0)
    End If
    
    getWorkerNameByWorksheet = Trim(sReturnName)
    
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


