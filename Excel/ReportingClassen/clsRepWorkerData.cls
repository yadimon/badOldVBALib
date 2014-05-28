VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRepWorkerData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsRepWorkerData
' Author    : Dmitry Gorelenkov
' Date      : 08.11.2013
' Changed   : 22.11.2013
' Purpose   : Klass fuer die Mitarbeiter bei Reporting
' Requires  : clsMyCollection, clsRepWorkerDateData, clsFuncs
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Explicit

Private m_sName As String 'name
Private m_lId As Long 'id
Private m_sAgency As String 'agentur von dem ma
Private m_mcData As clsMyCollection 'hier werden die clsRepWorkerDateData objects gespeichert. Key = datum in string

Private m_bInitDone As Boolean

Private f As clsFuncs


Public Property Get Data() As clsMyCollection
    Set Data = m_mcData
End Property

Public Property Get agency() As String
    agency = m_sAgency
End Property

Public Property Get id() As Long
    id = m_lId
End Property

Public Property Let id(lID As Long)
    m_lId = lID
End Property

Public Property Get name() As String
    name = m_sName
End Property

Public Function init(sName As String, sAgency As String, Optional mcDates As clsMyCollection, Optional lID As Long = -1)
    If m_bInitDone Then Exit Function
    
    If mcDates Is Nothing Then Set mcDates = New clsMyCollection
    
    m_sName = sName
    m_lId = lID
    m_sAgency = sAgency
    Set m_mcData = mcDates
    
    
    
    init = True
    m_bInitDone = True
End Function

'returns array with dates that the worker has
Public Function getWorkingDates() As Date()
    Dim sKey As Variant
    Dim dDatum As Date
    Dim aReturnDates() As Date

    For Each sKey In m_mcData.getKeys
        dDatum = f.CdateTry(sKey)
        Call f.arrayPush(aReturnDates, dDatum)
    Next sKey
    
    getWorkingDates = aReturnDates
End Function

'returns array of dates objects of the worker
Public Function getDateObjects() As clsRepWorkerDateData()
    Dim DateData As clsRepWorkerDateData
    Dim aReturnDatesArray() As clsRepWorkerDateData
    aReturnDatesArray = Array()
    Dim sKey As Variant
    
    For Each sKey In m_mcData.getKeys
        Set DateData = m_mcData(sKey)
        Call f.arrayPush(aReturnDatesArray, DateData)
    Next sKey
    
    getDateObjects = aReturnDatesArray
End Function

'returns array of dates objects of the worker
Public Function getDateObjectsAsCollection() As clsMyCollection
    Set getDateObjectsAsCollection = m_mcData.Copy
End Function

'return date object of this date
Public Function getWorkingDate(dDate As Date) As clsRepWorkerDateData
    Dim mcDataByDate As clsMyCollection
    Dim sKeyData As String
    
    Set getWorkingDate = Nothing
    
    sKeyData = dateToKey(dDate)
    
    If m_mcData.isInArray(sKeyData) Then
        Set getWorkingDate = m_mcData(sKeyData)
    End If
    
    
End Function

'adds new DateData object
Public Function addDateData(ByRef rwdDate As clsRepWorkerDateData)
    Call m_mcData.add(rwdDate, dateToKey(rwdDate.ThisDate))
End Function

'removes DateData object
Public Function removeDateData(rwdDate As clsRepWorkerDateData)
    Call m_mcData.remove(dateToKey(rwdDate.ThisDate))
End Function

'returns key for data collection, from date arg
Private Function dateToKey(dDate As Date)
    dateToKey = Format(dDate, "dd.mm.yyyy")
End Function

Private Sub Class_Initialize()
    Set m_mcData = New clsMyCollection
    Set f = New clsFuncs
End Sub

Private Sub Class_Terminate()
    Set m_mcData = Nothing
    Set f = Nothing
End Sub


