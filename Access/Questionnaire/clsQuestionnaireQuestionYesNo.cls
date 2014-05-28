VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuestionnaireQuestionYesNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IQuestionnaireQuestion

Private m_lID As Long 'id of the question
Private m_sType As String
Private m_qstnreParent As clsQuestionnaire 'parent, clsQuestionnaire reference
Private m_bAnswered As Boolean 'answered mark
Private m_mcControls As clsMyCollection 'collection with controls references
Private m_bInitDone As Boolean
Private m_lposition As Long
Private m_lNextQstPosition1 As Long
Private m_lNextQstPosition2 As Long
Private m_lNextQstPosition3 As Long

Private m_lAnswerIDYes As Long 'Answer ID to save by YES.
Private m_lAnswerIDNo As Long 'Answer ID to save by NO.
Private m_lAnswerIDNoAnswer As Long 'Answer ID to save by NO ANSWER.

Private m_mcSettings As clsMyCollection 'collection with any extra settings
Private m_mcCtrlsProps As clsMyCollection 'collection with properties of elements
Private fAccs As clsFuncsAccs

'CONTROLS THAT NEED TO REACT
Private WithEvents m_fraAnswers As Access.OptionGroup
Attribute m_fraAnswers.VB_VarHelpID = -1

Private Const LONGMAX As Long = 2 ^ 31 - 1


Private Sub Class_Initialize()
    Set fAccs = New clsFuncsAccs
    Set m_mcCtrlsProps = New clsMyCollection
    Set m_mcControls = New clsMyCollection
    Set m_mcSettings = New clsMyCollection
End Sub

Private Sub Class_Terminate()
    Set fAccs = Nothing
    Set m_mcCtrlsProps = Nothing
    Set m_mcControls = Nothing
    Set m_mcSettings = Nothing
End Sub

Private Property Let IQuestionnaireQuestion_answered(RHS As Boolean)
    answered = RHS
End Property

Private Function IQuestionnaireQuestion_clearAnswer() As Boolean
    IQuestionnaireQuestion_clearAnswer = clearAnswer
End Function

'only for interface...
Private Property Get IQuestionnaireQuestion_Id() As Long
    IQuestionnaireQuestion_Id = id
End Property

Private Property Get IQuestionnaireQuestion_qType() As String
    IQuestionnaireQuestion_qType = qType
End Property

Private Property Get IQuestionnaireQuestion_NextQstPosition() As Long
    IQuestionnaireQuestion_NextQstPosition = NextQstPosition
End Property

Private Property Let IQuestionnaireQuestion_Position(RHS As Long)
    Position = RHS
End Property

Private Property Get IQuestionnaireQuestion_Position() As Long
    IQuestionnaireQuestion_Position = Position
End Property

Private Property Get IQuestionnaireQuestion_answered() As Boolean
    IQuestionnaireQuestion_answered = answered
End Property

Private Function IQuestionnaireQuestion_answeredHandler() As Variant
    Call answeredHandler
End Function

Private Function IQuestionnaireQuestion_getAnswer() As clsMyCollection
    Set IQuestionnaireQuestion_getAnswer = getAnswer()
End Function

Private Function IQuestionnaireQuestion_getExtraSetting(sSettingName As String) As Variant
    IQuestionnaireQuestion_getExtraSetting = getExtraSetting(sSettingName)
End Function

Private Function IQuestionnaireQuestion_setExtraSetting(sSettingName As String, vValue As Variant) As Variant
    Call setExtraSetting(sSettingName, vValue)
End Function

Private Property Get IQuestionnaireQuestion_qstnreParent() As clsQuestionnaire
    Set IQuestionnaireQuestion_qstnreParent = qstnreParent
End Property

Private Property Set IQuestionnaireQuestion_qstnreParent(qstnreParentQuestionnaire_p As clsQuestionnaire)
    Set qstnreParent = qstnreParentQuestionnaire_p
End Property

Private Function IQuestionnaireQuestion_getControlsProps() As clsMyCollection
    Set IQuestionnaireQuestion_getControlsProps = getControlsProps
End Function

Private Function IQuestionnaireQuestion_getControlsRefs() As clsMyCollection
    Set IQuestionnaireQuestion_getControlsRefs = getControlsRefs
End Function

Private Function IQuestionnaireQuestion_bindControlsByOpenedForm(frmForm_p As Access.Form)
    IQuestionnaireQuestion_bindControlsByOpenedForm = bindControlsByOpenedForm(frmForm_p)
End Function




'real properties and funcs

Public Property Get answered() As Boolean
    answered = m_bAnswered
End Property

Private Property Let answered(bAnswered As Boolean)
    m_bAnswered = bAnswered
    'if set to "false", then clear answer
    If Not bAnswered Then
        Call clearAnswer
    End If
End Property

Public Property Get qstnreParent() As clsQuestionnaire
    Set qstnreParent = m_qstnreParent
End Property

Public Property Set qstnreParent(qstnreParentQuestionnaire_p As clsQuestionnaire)
    Set m_qstnreParent = qstnreParentQuestionnaire_p
End Property

Public Property Get id() As Long
    If Not m_bInitDone Then
        m_lID = CLng(Round(Rnd() * LONGMAX))
    End If
    
    id = m_lID
End Property

'type of Question
Public Property Get qType() As String
    qType = m_sType
End Property

Public Property Get Position() As Long
    Position = m_lposition
End Property

Public Property Let Position(lValue As Long)
    m_lposition = lValue
End Property

'position of next question to show
Public Property Get NextQstPosition() As Long
    Select Case m_fraAnswers.Value
        Case m_lAnswerIDYes 'case answered YES
            NextQstPosition = m_lNextQstPosition1
        Case m_lAnswerIDNo 'case NO
            NextQstPosition = m_lNextQstPosition2
        Case Else 'case else, "no Answer?"
            NextQstPosition = m_lNextQstPosition3
    End Select
End Property

'clears Answer controls
Public Function clearAnswer()
    m_fraAnswers.Value = Null
End Function

'call answered event handler of parent Questionnaire
'and set this QuestionObject property to "Answered"
Public Function answeredHandler()
    m_bAnswered = True
    Call m_qstnreParent.questionAnsweredHandler(Me)
End Function

Public Function init(lID_p As Long, lAnswerIDYes_p As Long, lAnswerIDNo_p As Long, lAnswerIDNoAnswer_p As Long, Optional lPosition As Long = -1, _
                        Optional lNextQuestionPosition1 As Long = -1, _
                        Optional lNextQuestionPosition2 As Long = -1, _
                        Optional lNextQuestionPosition3 As Long = -1, _
                        Optional sQuestionCaption_p As String = "Some Question", _
                        Optional sYesLabel As String = "Ja", Optional sNoLabel = "Nein", Optional sNoAnswerLabel = "Keine Antwort")
    If Not m_bInitDone Then m_lID = lID_p
    m_bInitDone = True
    m_lposition = lPosition
    m_lNextQstPosition1 = lNextQuestionPosition1
    m_lNextQstPosition2 = lNextQuestionPosition2
    m_lNextQstPosition3 = lNextQuestionPosition3
    m_sType = "YesNo"
    m_lAnswerIDYes = lAnswerIDYes_p
    m_lAnswerIDNo = lAnswerIDNo_p
    m_lAnswerIDNoAnswer = lAnswerIDNoAnswer_p
    Dim mcTmpControlDescription As clsMyCollection

    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_shpFrame" & id, "Name")
        Call .Add(acRectangle, "Type")
        Call .Add(8490, "Width")
        Call .Add(1806, "Height")
        Call .Add(0, "Top")
        Call .Add(0, "Left")
        Call .Add(2, "BorderWidth")
        Call .Add(0, "BorderColor")
        Call .Add(1, "BackStyle")
        Call .Add(-2147483607, "BackColor")
    End With
    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing
    '-----------------------------------------------------------------
    
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_lblCaption" & id, "Name")
        Call .Add(acLabel, "Type")
        Call .Add(sQuestionCaption_p, "Caption")
        Call .Add(6369, "Width")
        Call .Add(567, "Height")
        Call .Add(105, "Top")
        Call .Add(570, "Left")
        Call .Add("Futura Lt BT", "FontName")
        Call .Add(12, "FontSize")
        Call .Add(700, "FontWeight")
        Call .Add(False, "FontItalic")
        Call .Add(0, "ForeColor")
    End With
    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing
    '-----------------------------------------------------------------
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_fraAnswers" & id, "Name")
        Call .Add(acOptionGroup, "Type")
        Call .Add("[Event Procedure]", "AfterUpdate")
        Call .Add(7555, "Width")
        Call .Add(659, "Height")
        Call .Add(759, "Top")
        Call .Add(340, "Left")
    End With
    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing
    '-----------------------------------------------------------------
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_optYes" & id, "Name")
        Call .Add(acOptionButton, "Type")
        Call .Add(m_lAnswerIDYes, "OptionValue")
        Call .Add("m_fraAnswers" & id, "Parent")
        Call .Add(284, "Width")
        Call .Add(284, "Height")
        Call .Add(960, "Top")
        Call .Add(570, "Left")
    End With
    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing
    '-----------------------------------------------------------------
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_lblYes" & id, "Name")
        Call .Add(acLabel, "Type")
        Call .Add(sYesLabel, "Caption")
        Call .Add("m_optYes" & id, "Parent")

        Call .Add(1134, "Width")
        Call .Add(284, "Height")
        Call .Add(915, "Top")
        Call .Add(960, "Left")
        Call .Add("Futura Lt BT", "FontName")
        Call .Add(13, "FontSize")
        Call .Add(400, "FontWeight")
        Call .Add(False, "FontItalic")
        Call .Add(0, "ForeColor")
    End With
    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing
    
    '-----------------------------------------------------------------
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_optNo" & id, "Name")
        Call .Add(acOptionButton, "Type")
        Call .Add(m_lAnswerIDNo, "OptionValue")
        Call .Add("m_fraAnswers" & id, "Parent")
        Call .Add(284, "Width")
        Call .Add(284, "Height")
        Call .Add(960, "Top")
        Call .Add(2895, "Left")
    End With
    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing
    
    '-----------------------------------------------------------------
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_lblNo" & id, "Name")
        Call .Add(acLabel, "Type")
        Call .Add(sNoLabel, "Caption")
        Call .Add("m_optNo" & id, "Parent")
        
        Call .Add(1134, "Width")
        Call .Add(284, "Height")
        Call .Add(915, "Top")
        Call .Add(3240, "Left")
        Call .Add("Futura Lt BT", "FontName")
        Call .Add(13, "FontSize")
        Call .Add(400, "FontWeight")
        Call .Add(False, "FontItalic")
        Call .Add(0, "ForeColor")
    End With
    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing
        '-----------------------------------------------------------------
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_optNoAnswer" & id, "Name")
        Call .Add(acOptionButton, "Type")
        Call .Add(m_lAnswerIDNoAnswer, "OptionValue")
        Call .Add("m_fraAnswers" & id, "Parent")
        Call .Add(285, "Width")
        Call .Add(285, "Height")
        Call .Add(960, "Top")
        Call .Add(5535, "Left")
    End With
    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing
    
    '-----------------------------------------------------------------
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_lblNoAnswer" & id, "Name")
        Call .Add(acLabel, "Type")
        Call .Add(sNoAnswerLabel, "Caption")
        Call .Add("m_optNoAnswer" & id, "Parent")
        
        Call .Add(1650, "Width")
        Call .Add(315, "Height")
        Call .Add(885, "Top")
        Call .Add(5820, "Left")
        Call .Add("Futura Lt BT", "FontName")
        Call .Add(13, "FontSize")
        Call .Add(400, "FontWeight")
        Call .Add(False, "FontItalic")
        Call .Add(0, "ForeColor")
    End With
    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing
    '-----------------------------------------------------------------
    
    
    
    'referenzen auf real Controls zuerst mit null befuellen
    Dim cpropControl As Variant 'clsControlProperties
    For Each cpropControl In m_mcCtrlsProps.GetValues
        Call m_mcControls.Add(Null, CStr(cpropControl.getCollection("Name")))
    Next cpropControl
    
End Function

'add control by collection of propetries
Private Function addControlByCol(mcProps As clsMyCollection)
    Dim cpropsTmpCtrlProperty As clsControlProperties
    Set cpropsTmpCtrlProperty = New clsControlProperties
    
    Call cpropsTmpCtrlProperty.addControlPropsByCol(mcProps)
    Call m_mcCtrlsProps.Add(cpropsTmpCtrlProperty, cpropsTmpCtrlProperty.ControlName)
    
    Set cpropsTmpCtrlProperty = Nothing
End Function

Public Function getExtraSetting(sSettingName As String)
    'returns any extra setting/info/state of the question object
    getExtraSetting = m_mcSettings(sSettingName)
End Function

Public Function setExtraSetting(sSettingName As String, vValue As Variant)
    'add any extra setting/info/state of the question object
    Call m_mcSettings.Add(vValue, sSettingName)
End Function

Public Function getControlsRefs() As clsMyCollection
    Set getControlsRefs = m_mcControls
End Function

Public Function getControlsProps() As clsMyCollection
    Set getControlsProps = m_mcCtrlsProps
End Function

'alls selected answers, key=text of selection, value=ID of answer
Public Function getAnswer() As clsMyCollection
    Dim mcReturnCollection As New clsMyCollection
    
    Call mcReturnCollection.Add(m_fraAnswers.Value, "m_fraAnswers.value")

    Set getAnswer = mcReturnCollection
End Function


'alle controls die in dem Question object drin sein mussen, binden, wenn form geoeffnet ist, bei manchen auch events
Public Function bindControlsByOpenedForm(frmForm_p As Access.Form)
    Dim sCtrlName As Variant
    Dim tryCtrl As Control
    
    If frmForm_p Is Nothing Then Exit Function
    
    If fAccs.isFormLoaded(frmForm_p.Name) Then
    
        For Each sCtrlName In m_mcControls.getKeys
            Set tryCtrl = fAccs.tryGetControl(frmForm_p.Name, CStr(sCtrlName), Null)
            If Not tryCtrl Is Nothing Then
                Call m_mcControls.setValue(tryCtrl, sCtrlName)
                
                'bind events
                If Left(sCtrlName, Len("m_fraAnswers")) = "m_fraAnswers" Then
                    Set m_fraAnswers = tryCtrl
                End If
                
                Set tryCtrl = Nothing
                
            Else
                Debug.Print "bindControlsByOpenedForm: control " & sCtrlName & " not found on form"
            End If
        Next sCtrlName
    Else
        Debug.Print "called bindControlsByOpenedForm, with not loaded form as argument"
    End If
End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::EVENTS::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'case input changed, call "answered"
Private Sub m_fraAnswers_AfterUpdate()
    answeredHandler
End Sub



