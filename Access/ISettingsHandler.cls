VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISettingsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : ISettingsHandler
' Author    : Dmitry Gorelenkov
' Date      : 27.01.2013
' Purpose   : Settings Class Interface
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit


Public Function getSettingsNames() As Variant
    'get all setting names
End Function

Public Function getSetting(sSettingName As String) As String
    'get setting value as string
End Function

Public Function setSetting(sWhatParam As String, sWhatValue As String) As String
    'Gibt "OK" zuruck falls erfolgreicht, sonst fehlerstring.
End Function

Public Function addSetting(sSettingName As String, sValue As String) As String
    'neue setting hinzufuegen, return: "OK" sonst fehler
End Function

