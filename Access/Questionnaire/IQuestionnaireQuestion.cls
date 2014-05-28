VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQuestionnaireQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : IQuestionnaireQuestion
' Author    : Dmitry Gorelenkov
' Date      : 08.02.2013
' Requires  : clsMyCollection, clsControlProperties
' Purpose   : Interface for Question of Questionnaire
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

'Private m_lId As Long
'Private m_qstnreParent As clsQuestionnaire 'parent, clsQuestionnaire reference
'Private m_bAnswered As Boolean
'Public Function init() - Its good if the questionObj has any initialize func. For ID number for example

Public Property Get answered() As Boolean
'    answered = m_bAnswered
End Property

Public Property Let answered(bValue As Boolean)
'change answered value, maybe clear question controls?
End Property

Public Property Get qstnreParent() As clsQuestionnaire
'    Set qstnreParent = m_qstnreParent
End Property

Public Property Set qstnreParent(qstnreParentQuestionnaire_p As clsQuestionnaire)
'    Set m_qstnreParent = qstnreParentQuestionnaire_p
End Property

Public Property Get id() As Long
'    Id = m_lId
End Property

Public Property Get qType() As String
'    get type of question object
End Property

Public Property Get Position() As Long
    'what position has the question on questionnaire
End Property

Public Property Let Position(lValue As Long)
    'what position has the question on questionnaire
End Property

Public Property Get NextQstPosition() As Long
    'position of next question to show
End Property

Public Function clearAnswer() As Boolean
    'reset all changes
End Function


Public Function getControlsRefs() As clsMyCollection
    'get collection of control references (must be already built on opened form)
End Function

Public Function getControlsProps() As clsMyCollection
    'returns props of controls
End Function

Public Function bindControlsByOpenedForm(ByRef frmForm_p As Access.Form)
    'save control references from opened built form to this Question object vars
End Function

Public Function getExtraSetting(sSettingName As String)
    'returns any extra setting/info/state of the question object
End Function

Public Function setExtraSetting(sSettingName As String, vValue As Variant)
    'add any extra setting/info/state of the question object
End Function

Public Function getAnswer() As clsMyCollection
    'get collection of answers
End Function


Public Function answeredHandler()
    'call parant eventHandler, in case the question ist answered.
    'you have to put this function to any event procedure of controls of this Question object,
    'to let parant Questionnaire object know this question is answered
    
'    m_bAnswered = True
'    If Not m_qstnreParent Is Nothing And Not m_bAnswered Then
'        Call m_qstnreParent.questionAnsweredHandler(Me)
'    End If
End Function


Private Sub Class_Terminate()
'    Set m_qstnreParent = Nothing
End Sub
