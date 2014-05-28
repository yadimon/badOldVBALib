VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUserSettingsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsUserSettingsHandler
' Author    : Dmitry Gorelenkov
' Date      : 27.01.2013
' Changed   : 14.09.2013
' Purpose   : Setting Handler, fuer eigeloggten User
' Requires  : clsFuncs, clsErrorHandler
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------


Option Compare Database
Option Explicit
Implements ISettingsHandler
Private ehError As clsErrorHandler
Private f As clsFuncs
Private m_lCurrentUserID As Long
Private m_lCurrentUserGroupID As Long
Private m_stblSettings As String
Private m_stblUserSettings As String
Private m_stblGroupSettings As String

'bind userSettingsHandler Object to User
'if call init more the 1 time, not given args-setting will not be overridden
Public Function init(lCurrentUserID_p As Long, _
                     lCurrentUserGroupID_p As Long, _
                     stblSettings_p As String, _
                     stblUserSettings_p As String, _
                     stblGroupSettings_p As String)

    On Error GoTo init_Error
    
    'TODO test input
    m_lCurrentUserID = lCurrentUserID_p
    m_lCurrentUserGroupID = lCurrentUserGroupID_p
    m_stblSettings = stblSettings_p
    m_stblUserSettings = stblUserSettings_p
    m_stblGroupSettings = stblGroupSettings_p



    Exit Function

init_Error:
    ehError.catchErr ("init of Klassenmodul clsUserSettingsHandler")
    

End Function


'get all setting names
Public Function ISettingsHandler_getSettingsNames() As Variant
'no need?
    On Error GoTo getSettingsNames_Error
    Dim arrayRS As Variant
    Dim arrayOfSetting As Variant
    Dim arrayElem As Variant

    arrayRS = Nz(db.SQLArray(db.SQL("SELECT setsName FROM " & m_stblSettings)), Array())

    For Each arrayElem In arrayRS
        Call f.arrayPush(arrayOfSetting, arrayElem(0))

    Next arrayElem


    ISettingsHandler_getSettingsNames = arrayOfSetting

    Exit Function

getSettingsNames_Error:
    Debug.Print "Error by getSettingsNames: " & err.Number & vbCrLf & err.Description
    ISettingsHandler_getSettingsNames = Array()
End Function

'einstellungen aus localer Tabelle laden / oder von dem schon vorgeladenen Array/collection
Public Function ISettingsHandler_getSetting(sSettingName As String) As String
    ISettingsHandler_getSetting = getSetting(sSettingName)
End Function

'setzt setting wert ein. Wichtig, setting tabelle muss mit STRING vars gefuellt werden
Public Function ISettingsHandler_setSetting(sSettingName As String, sValue As String) As String
    ISettingsHandler_setSetting = setSetting(sSettingName, sValue)
End Function

'neue setting zu dem Benutzer hinzufuegen
Public Function ISettingsHandler_addSetting(sSettingName As String, sValue As String) As String
    ISettingsHandler_addSetting = addSetting(sSettingName, sValue)
End Function

'TODO mit Local Array einstellungen speichern (+performance)
'einstellungen aus localer Tabelle laden / oder von dem schon vorgeladenen Array/collection
Public Function getSetting(sSettingName As String) As String
'zuerst in Benutzer settings, danach in Group settings suchen
    Dim sResult As String
    Dim lAffected As Long
    Dim lSettingID As Long
    
    On Error GoTo getSetting_Error

    'falls keiner eingeloggt, fehler..
    If m_lCurrentUserID <= 0 Then
        getSetting = "%%NO_USER_SET%%"
        Exit Function
    End If

    'save SettingID from SettingName
    lSettingID = getSettingIdByName(sSettingName)

    'falls keine setting existiert..
    If lSettingID = -1 Then
        getSetting = "%%NO_SETTING_" & sSettingName & "_FOUND%%"
        Exit Function
    End If
    
    
    'in USER settings suchen..
    sResult = Nz(db.SQLValue("SELECT usrsetsValue FROM " & _
                             m_stblUserSettings & " WHERE usrsetsUser = " & m_lCurrentUserID & " AND usrsetsSetting = " & lSettingID, , , lAffected))

    'falls fehler oder keine Einstellung gefunden, weiter in Group einstellungen suchen
    If lAffected < 1 Then
        sResult = Nz(db.SQLValue("SELECT grpsetsValue FROM " & _
                                 m_stblGroupSettings & " WHERE grpsetsGroup = " & m_lCurrentUserGroupID & " AND grpsetsSetting = " & lSettingID, , , lAffected))
    End If

    If lAffected < 1 Then
getSetting_Error:
        If err.Number <> 0 Then ehError.catchErr ("getSetting of Klassenmodul clsUserSettingsHandler")
        sResult = "%%ERROR%%"
    End If

    getSetting = sResult

End Function

'setzt setting wert ein. Wichtig, setting tabelle muss mit STRING vars gefuellt werden
Public Function setSetting(sSettingName As String, sValue As String) As String
    Dim lSettingID As Long
    Dim sCurrentSettingValue As String
    Dim lAffected As Long
    On Error GoTo setSetting_Error

    If m_lCurrentUserID <= 0 Then
        setSetting = "No User set"
        Exit Function
    End If

    lSettingID = getSettingIdByName(sSettingName)

    'falls keine setting gefunden wurde..
    If lSettingID = -1 Then
        setSetting = "%% NO SETTING " & sSettingName & " FOUND%%"
        Exit Function
    End If

    'pruefen ob benutzer die Einstellung hat
    sCurrentSettingValue = Nz(db.SQLValue("SELECT usrsetsValue FROM " & m_stblUserSettings & _
                                          " WHERE usrsetsUser = " & m_lCurrentUserID & " AND usrsetsSetting = " & lSettingID, , , lAffected), "")

    'falls  mehr als 1 Setting gefunden, fehler..
    If lAffected > 1 Then
        setSetting = lAffected & " Records found"
        Exit Function
    End If

    'falls keine settings gefunden, setting erstellen
    If lAffected = 0 Then
        Dim sResult As String
        sResult = addSetting(sSettingName, sValue)
        If sResult <> "OK" Then
            setSetting = "Cannot add setting, error: " & sResult
            Exit Function
        End If

    Else    'setting gefunden
        'falls neues Wert, setzen
        If StrComp(sCurrentSettingValue, sValue, vbBinaryCompare) <> 0 Then
            Call db.SQL("UPDATE " & m_stblUserSettings & " SET usrsetsValue = '" & sValue & _
                        "' WHERE usrsetsUser = " & m_lCurrentUserID & " AND usrsetsSetting = " & lSettingID, , , lAffected)
            'falls nicht nur 1 aenderung, fehler..
            If lAffected <> 1 Then
                setSetting = lAffected & " Records changed"
                Exit Function
            End If
        End If
    End If




    setSetting = "OK"

    Exit Function

setSetting_Error:
    ehError.catchErr ("setSetting of Klassenmodul clsUserSettingsHandler")

End Function

'neue setting zu dem Benutzer hinzufuegen
Public Function addSetting(sSettingName As String, sValue As String) As String
    Dim lSettingID As Long
    Dim lAffected As Long
    On Error GoTo addSetting_Error


    lSettingID = getSettingIdByName(sSettingName)

    'falls keine setting gefunden wurde..
    If lSettingID = -1 Then
        addSetting = "%% NO SETTING " & sSettingName & " FOUND%%"
        Exit Function
    End If

    Call db.SQL("INSERT INTO " & m_stblUserSettings & " (usrsetsUser,usrsetsSetting, usrsetsValue) VALUES(" & _
                                                        m_lCurrentUserID & ", " & lSettingID & ", """ & sValue & """) ", , , lAffected)

    If lAffected <> 1 Then
        addSetting = "Add failed?, affected: " & lAffected & " records"
        Exit Function
    End If


    addSetting = "OK"

    Exit Function

addSetting_Error:
    ehError.catchErr ("addSetting of Klassenmodul clsUserSettingsHandler")

End Function

'get setting ID by settinName
Private Function getSettingIdByName(sSettingName As String) As Long
'todo performance: local array?
    Dim sResult As Variant
    sResult = db.SQLValue("SELECT setsID FROM " & m_stblSettings & " WHERE setsName = """ & sSettingName & """")
    If Not IsNumeric(sResult) Then sResult = -1
    
    getSettingIdByName = sResult
End Function

Private Sub Class_Initialize()
    Set ehError = New clsErrorHandler
    Set f = New clsFuncs
End Sub

Private Sub Class_Terminate()
    Set ehError = Nothing
    Set f = Nothing
End Sub
