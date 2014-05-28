VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuestionnaireHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsQuestionnaireHandler
' Author    : Dmitry Gorelenkov
' Date      : 03.03.2013
' Changed   : 03.03.2013
' Purpose   : Handler, um questionnaire mit daten aus DB zu starten
' Requires  : clsQuestionnaire, clsQuestionnaireDB und weitere darin benutzte klassen
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
Private WithEvents qstnreQ As clsQuestionnaire
Attribute qstnreQ.VB_VarHelpID = -1
Private QuestionnaireDB As clsQuestionnaireDB
Private Const CM As Integer = 567

Public Function Start(lQstnreID As Long, lUserID As Long, lCustomerID As Long, _
                        Optional lWidthCM As Integer = 0, Optional lHeightCM As Integer = 0, _
                        Optional bRemoveAfter As Boolean = True, Optional sFormSaveName As String = "")
    Set QuestionnaireDB = New clsQuestionnaireDB
    Call QuestionnaireDB.init(lUserID, lCustomerID)
    
    Set qstnreQ = QuestionnaireDB.createQuestionnaire(lQstnreID)
    Call qstnreQ.openAndBuildIfNotExist(sFormSaveName, lHeightCM * CM, lWidthCM * CM, bRemoveAfter)
End Function



':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::EVENT HANDLERS::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'onsave events of questionnaire object
Private Sub qstnreQ_OnSave(ByRef bResult As Boolean)
    bResult = QuestionnaireDB.saveQuestionnaire(qstnreQ)

    If Not bResult Then
        MsgBox "Wahrscheinlich fehlerhaft gespeichert :/"
        'Todo LOG
    Else
        MsgBox "Gespeichert!"
        'TODO RaiseEvent "form_saved" todo: create event form_saved
    End If
    
    'in case there some new answers created, force to remove form after, and generate new on next start
    If QuestionnaireDB.CreatedNewAnswers Then
        Call qstnreQ.removeAfter
    End If
End Sub




