VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettingsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsSettingsHandler
' Author    : Dmitry Gorelenkov
' Date      : 27.01.2013
' Purpose   : settings get and set
' Requires  : clsFuncs, clsMyCollection
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
Implements ISettingsHandler
Private m_mcSettings As clsMyCollection
Private f As clsFuncs

Private Function ISettingsHandler_addSetting(sSettingName As String, sValue As String) As String
    ISettingsHandler_addSetting = "NOT IMPLEMENTED YET"
End Function

'get all setting names
Public Function ISettingsHandler_getSettingsNames() As Variant
    
    On Error GoTo getSettingsNames_Error

    ISettingsHandler_getSettingsNames = m_mcSettings.getKeys


    Exit Function

getSettingsNames_Error:
    Debug.Print "Error by getSettingsNames: " & Err.Number & vbCrLf & Err.Description
    ISettingsHandler_getSettingsNames = Array()
End Function

'einstellungen aus localer Tabelle laden / oder von dem schon vorgeladenen Array/collection
Public Function ISettingsHandler_getSetting(sSettingName As String) As String
    Dim dbTmp As Database
    Dim rsTmp As Recordset
    
    On Error GoTo getSetting_Error

    If m_mcSettings.isInArray(sSettingName) Then
        ISettingsHandler_getSetting = m_mcSettings(sSettingName)
        Exit Function
    End If

    'sont von DB laden
    
    Set dbTmp = CurrentDb()
    Set rsTmp = dbTmp.OpenRecordset("SELECT SettingValue FROM Settings WHERE SettingName = '" & sSettingName & "';")
    ISettingsHandler_getSetting = rsTmp.Fields(0).Value
    
    On Error GoTo 0
    Exit Function

getSetting_Error:
    Debug.Print "Error by getSetting: cant load Setting: " & sSettingName
    ISettingsHandler_getSetting = ""
    
End Function

'setzt setting wert ein. Wichtig, setting tabelle muss mit STRING vars gefuellt werden
Public Function ISettingsHandler_setSetting(sWhatParam As String, sWhatValue As String) As String
    
    On Error GoTo setSetting_Error

    
    Dim dbTmp As Database
    
    Set dbTmp = CurrentDb()
    
    dbTmp.Execute ("UPDATE Settings SET SettingValue = '" & sWhatValue & "' WHERE SettingName = '" & sWhatParam & "';")
    Call m_mcSettings.setValue(sWhatValue, sWhatParam) ' auch in settings collection speichern
    ISettingsHandler_setSetting = "OK"

    Exit Function

setSetting_Error:
    Debug.Print "Error by setSetting of Modul Variables"
    ISettingsHandler_setSetting = "0"

End Function

'werte in setting array speichern
Private Function loadSettingsArray(ByRef colSettingVar As clsMyCollection)
    Dim dbTmp As Database
    Dim rsTmp As Recordset
    

    
    Set dbTmp = CurrentDb()
    Set rsTmp = dbTmp.OpenRecordset("SELECT SettingValue, SettingName FROM Settings;")
    If Not rsTmp.EOF Then
        rsTmp.MoveFirst
        
        While Not rsTmp.EOF
            Call colSettingVar.add(rsTmp.Fields(0).Value, rsTmp.Fields(1).Value)
            rsTmp.MoveNext
        Wend
        
    End If
    
    Set f = Nothing
End Function

Private Sub Class_Initialize()
    Set m_mcSettings = New clsMyCollection
    Set f = New clsFuncs
    Call loadSettingsArray(m_mcSettings) 'setting werte in collection laden
End Sub


