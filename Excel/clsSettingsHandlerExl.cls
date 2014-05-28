VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettingsHandlerExl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsSettingsHandlerExl
' Author    : Dmitry Gorelenkov
' Date      : 14.11.2013
' Changed   : 14.11.2013
' Purpose   : settings manager for excel
' Requires  : worksheet with the settings headers, or range with settings
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Explicit


Private m_bInitDone As Boolean
Private m_mcSettings As clsMyCollection
Private m_rngRangeGotFrom As Range

Public Function init(wksSettingsWorksheet As Worksheet, _
                    Optional sSettingNameColumn As String = "SettingName", _
                    Optional sSettingValueHeader As String = "SettingValue") As Boolean
    If m_bInitDone Then Exit Function
    
    Set m_mcSettings = New clsMyCollection
    
    'TODO
    init = True
    m_bInitDone = True
End Function

Public Function initByRange(rngSettingsRange As Range) As Boolean
    If m_bInitDone Then Exit Function
        
    If rngSettingsRange Is Nothing Then Exit Function
    
    Set m_rngRangeGotFrom = rngSettingsRange
    Set m_mcSettings = fXl.rangeToMyCollection(rngSettingsRange)
    
    
    initByRange = True
    m_bInitDone = True
End Function

Public Function getSetting(sSettingName As String) As String
    If Not test_init() Then Exit Function
    getSetting = m_mcSettings.getValue(sSettingName)
End Function

Public Function saveSetting(sSettingName As String, vValue As String) As Boolean
    If Not test_init() Then Exit Function
    'falls in settings vorhanden
    If m_mcSettings.isInArray(sSettingName) Then
        'save value
        Call m_mcSettings.setValue(vValue, sSettingName)
    End If
    
    'put collection to range
    saveSetting = fXl.myCollectionToRange(m_mcSettings, m_rngRangeGotFrom)
    
    'reloadData
End Function

'refresh data
Public Function reloadData() As Boolean
    If Not test_init() Then Exit Function
    m_bInitDone = False
    reloadData = initByRange(m_rngRangeGotFrom)
End Function

'test if initialized, if not, prints error
Private Function test_init() As Boolean
    
    If Not m_bInitDone Then
        Debug.Print "clsSettingsHandler: Not initialized!"
    End If
    
    test_init = m_bInitDone
End Function

