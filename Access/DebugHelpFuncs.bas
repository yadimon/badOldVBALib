Attribute VB_Name = "DebugHelpFuncs"
'---------------------------------------------------------------------------------------
' Module    : DebugHelpFuncs
' Author    : Dmitry Gorelenkov
' Date      : 08.11.2013
' Changed   : 08.11.2013
' Purpose   : helper for debugging TMX
' Requires  :
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------


Option Compare Database
Option Explicit

Public Function closeAllForms()
    Dim frmForm As Variant
    For Each frmForm In Forms
        DoCmd.Close acForm, frmForm.Name, acSaveNo
    Next frmForm
End Function

'aktuelle version aktualisieren
Public Function updateVersion(sDescription As String)
    Dim dbase As Database
    Dim changelogTableRs As Recordset
    Dim sVersion As String
    
    Set dbase = CurrentDb()
'    dDate = Cdate()
    Set changelogTableRs = dbase.TableDefs("ChangeLog").OpenRecordset(dbOpenDynaset)
    sVersion = Format$(Date, "YY.MM.DD")
    
    With changelogTableRs
        .AddNew
'        .Edit
        !Version = sVersion
        !Datum = Now
        !Description = sDescription
        .Update
        .Close
    End With
    
    'current version update
    DoCmd.RunSQL ("UPDATE Settings SET SettingValue = '" & sVersion & "' WHERE SettingName = 'p_sCurrentVersion'")

'    changelogTableRs.Close
    dbase.Close
End Function

'alias for updateVersion
Public Function commit(sDescription As String)
    commit = updateVersion(sDescription)
End Function
