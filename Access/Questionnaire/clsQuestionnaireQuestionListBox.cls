VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuestionnaireQuestionListBox"
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
Private m_lNextQstPosition As Long
Private m_bytMultiselect As Byte 'multiselect option

Private m_mcSettings As clsMyCollection 'collection with any extra settings
Private m_mcCtrlsProps As clsMyCollection 'collection with properties of elements
Private fAccs As clsFuncsAccs

'CONTROLS THAT NEED TO REACT
Private WithEvents m_txtInput1 As Access.TextBox
Attribute m_txtInput1.VB_VarHelpID = -1
Private WithEvents m_txtInput2 As Access.TextBox
Attribute m_txtInput2.VB_VarHelpID = -1
Private WithEvents m_txtInput3 As Access.TextBox
Attribute m_txtInput3.VB_VarHelpID = -1
Private aExtraTextBoxs As Variant

Private WithEvents m_lstAnswersList As Access.ListBox
Attribute m_lstAnswersList.VB_VarHelpID = -1
'Private m_shpFrame As Access.Rectangle

Private Const CM As Long = 567
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

Private Property Let IQuestionnaireQuestion_answered(RHS As Boolean)
    answered = RHS
End Property

Private Function IQuestionnaireQuestion_clearAnswer() As Boolean
    IQuestionnaireQuestion_clearAnswer = clearAnswer
End Function

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

'Private Function IQuestionnaireQuestion_init(lID_p As Long) As Variant
'    init (lID_p)
'End Function

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
    NextQstPosition = m_lNextQstPosition
End Property

'clears Answer controls
Public Function clearAnswer()
    Dim txtBoxTmp As Variant 'txtbox
    Call fAccs.unselectListBox(m_lstAnswersList)

    For Each txtBoxTmp In aExtraTextBoxs
        txtBoxTmp.Value = ""
    Next txtBoxTmp
End Function


'call answered event handler of parent Questionnaire
'and set this QuestionObject property to "Answered"
Public Function answeredHandler()
    Call m_qstnreParent.questionAnsweredHandler(Me)
    m_bAnswered = True
End Function

Public Function init(lID_p As Long, Optional lPosition As Long = -1, Optional lNextQuestionPosition As Long = -1, _
                     Optional sQuestionCaption_p As String = "Some Question", Optional mcValuesToFill_p As clsMyCollection, _
                     Optional iOwnAnswers_p As Integer = 0, Optional sOwnAnswersLabel As String = "", Optional bytMultiselect_p As Byte = 1)
    If Not m_bInitDone Then m_lID = lID_p
    m_bInitDone = True
    m_lposition = lPosition
    m_lNextQstPosition = lNextQuestionPosition
    m_sType = "ListBox"
    m_bytMultiselect = bytMultiselect_p
    aExtraTextBoxs = Array()

    Dim mcTmpControlDescription As clsMyCollection

    
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_shpFrame" & id, "Name")
        Call .Add(acRectangle, "Type")
        Call .Add(8490, "Width")
        Call .Add(4761, "Height")
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
        Call .Add(acLabel, "Type")
        Call .Add(7470, "Width")
        Call .Add(585, "Height")
        Call .Add(285, "Top")
        Call .Add(570, "Left")
        Call .Add("Futura Lt BT", "FontName")
        Call .Add(12, "FontSize")
        Call .Add(700, "FontWeight")
        Call .Add(False, "FontItalic")
        Call .Add(0, "ForeColor")
        Call .Add(sQuestionCaption_p, "Caption")
        Call .Add("m_lblCaption" & id, "Name")
    End With

    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing
    '-----------------------------------------------------------------
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_lblElseInput" & id, "Name")
        Call .Add(acLabel, "Type")
        Call .Add(sOwnAnswersLabel, "Caption")
        Call .Add(2295, "Width")
        Call .Add(345, "Height")
        Call .Add(1050, "Top")
        Call .Add(5685, "Left")
        Call .Add("Futura Lt BT", "FontName")
        Call .Add(13, "FontSize")
        Call .Add(700, "FontWeight")
        Call .Add(False, "FontItalic")
        Call .Add(0, "ForeColor")
    End With
    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing

    '-----------------------------------------------------------------
    'in case there is allowed 1 or more own answer fields..
    If iOwnAnswers_p >= 1 Then
        Set mcTmpControlDescription = New clsMyCollection
        With mcTmpControlDescription
            Call .Add("m_txtInput1" & id, "Name")
            Call .Add(acTextBox, "Type")
            Call .Add(2286, "Width")
            Call .Add(357, "Height")
            Call .Add(1590, "Top")
            Call .Add(5700, "Left")
            Call .Add("Futura Lt BT", "FontName")
            Call .Add(11, "FontSize")
            Call .Add(400, "FontWeight")
            Call .Add(False, "FontItalic")
            Call .Add(0, "ForeColor")
        End With
        Call addControlByCol(mcTmpControlDescription)
        Set mcTmpControlDescription = Nothing
    End If
    '-----------------------------------------------------------------
    If iOwnAnswers_p >= 2 Then
        Set mcTmpControlDescription = New clsMyCollection
        With mcTmpControlDescription
            Call .Add("m_txtInput2" & id, "Name")
            Call .Add(acTextBox, "Type")
            Call .Add(2286, "Width")
            Call .Add(357, "Height")
            Call .Add(2040, "Top")
            Call .Add(5700, "Left")
            Call .Add("Futura Lt BT", "FontName")
            Call .Add(11, "FontSize")
            Call .Add(400, "FontWeight")
            Call .Add(False, "FontItalic")
            Call .Add(0, "ForeColor")
        End With
        Call addControlByCol(mcTmpControlDescription)
        Set mcTmpControlDescription = Nothing
    End If
    '-----------------------------------------------------------------
    If iOwnAnswers_p >= 3 Then
        Set mcTmpControlDescription = New clsMyCollection
        With mcTmpControlDescription
            Call .Add("m_txtInput3" & id, "Name")
            Call .Add(acTextBox, "Type")
            Call .Add(2286, "Width")
            Call .Add(357, "Height")
            Call .Add(2475, "Top")
            Call .Add(5700, "Left")
            Call .Add("Futura Lt BT", "FontName")
            Call .Add(11, "FontSize")
            Call .Add(400, "FontWeight")
            Call .Add(False, "FontItalic")
            Call .Add(0, "ForeColor")
        End With
        Call addControlByCol(mcTmpControlDescription)
        Set mcTmpControlDescription = Nothing
    End If
    '-----------------------------------------------------------------
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add(acListBox, "Type")
        Call .Add(4521, "Width")
        Call .Add(3645, "Height")
        Call .Add(930, "Top")
        Call .Add(570, "Left")
        Call .Add("Futura Md BT", "FontName")
        Call .Add(12, "FontSize")
        Call .Add(400, "FontWeight")
        Call .Add(False, "FontItalic")
        Call .Add(0, "ForeColor")
        Call .Add("m_lstAnswersList" & id, "Name")
        Call .Add("Value List", "RowSourceType")
        Call .Add("1", "BoundColumn")
        Call .Add("2", "ColumnCount")
        Call .Add("0;" & 2 * CM, "ColumnWidths")
        Call .Add(m_bytMultiselect, "MultiSelect")
        Call .Add("-2147483606", "BackColor")
'        Call .add("[Event Procedure]", "OnClick")

        'values in the ListBox
        If Not mcValuesToFill_p Is Nothing Then
            Dim sKey As Variant
            Dim arrayKeyValue As Variant
            Dim sResultString As String
            arrayKeyValue = Array()
            For Each sKey In mcValuesToFill_p.getKeys
                'push key, index/ID of answer
                Call f.arrayPush(arrayKeyValue, sKey)
                'push visible text, selection to klick
                Call f.arrayPush(arrayKeyValue, """" & mcValuesToFill_p.getValue(sKey) & """")
            Next sKey
            sResultString = Join(arrayKeyValue, ";")
            Call .Add(sResultString, "RowSource")
        End If

    End With

    Call addControlByCol(mcTmpControlDescription)
    Set mcTmpControlDescription = Nothing
    '-----------------------------------------------------------------




    'referencen auf real Controls zuerst mit null befuellen
    Dim cpropControl As Variant    'clsControlProperties
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
    Dim idxRow As Variant
    Dim mcReturnCollection As New clsMyCollection
    
    With m_lstAnswersList
        For Each idxRow In .ItemsSelected
            Call mcReturnCollection.Add(.Column(0, idxRow), .Column(1, idxRow))
        Next idxRow
    End With
    
    Set getAnswer = mcReturnCollection
End Function

'new inserted Answers
Public Function getNewAnswers() As clsMyCollection
    Dim mcReturnCollection As New clsMyCollection
    Dim txtBoxTmp As Variant 'txtBox

    
    For Each txtBoxTmp In aExtraTextBoxs
        If Nz(txtBoxTmp.Value, "") <> vbNullString Then
            mcReturnCollection.Add (CStr(txtBoxTmp.Value))
        End If
    Next txtBoxTmp
    
    Set getNewAnswers = mcReturnCollection
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
                If Left(sCtrlName, Len("m_txtInput")) = "m_txtInput" Then
                    Select Case Mid(sCtrlName, Len("m_txtInput") + 1, 1)
                        Case "1"
                            Set m_txtInput1 = tryCtrl
                            m_txtInput1.OnChange = "[Event Procedure]"
                            Call f.arrayPush(aExtraTextBoxs, m_txtInput1)
                        Case "2"
                            Set m_txtInput2 = tryCtrl
                            m_txtInput2.OnChange = "[Event Procedure]"
                            Call f.arrayPush(aExtraTextBoxs, m_txtInput2)
                        Case "3"
                            Set m_txtInput3 = tryCtrl
                            m_txtInput3.OnChange = "[Event Procedure]"
                            Call f.arrayPush(aExtraTextBoxs, m_txtInput3)
                        End Select
                        
                ElseIf Left(sCtrlName, Len("m_lstAnswersList")) = "m_lstAnswersList" Then
                    Set m_lstAnswersList = tryCtrl
                    m_lstAnswersList.AfterUpdate = "[Event Procedure]"
'                    m_lstAnswersList.OnClick = "[Event Procedure]"
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

'access 2007 bugfix function. Events dont bind by Form_load
''NO NEED?
'Public Function refreshEventBindings()
'    Dim tmpControl As control
'    m_txtInput.OnChange = ""
'    m_txtInput.OnClick = ""
'
'    Set tmpControl = m_txtInput
'    Set m_txtInput = Nothing
'    Set m_txtInput = tmpControl
'    m_txtInput.OnChange = "[Event Procedure]"
'    m_txtInput.OnClick = "[Event Procedure]"
'End Function



'case input changed, call "answered"
Private Sub m_txtInput1_Change()
    answeredHandler
End Sub
Private Sub m_txtInput2_Change()
    answeredHandler
End Sub
Private Sub m_txtInput3_Change()
    answeredHandler
End Sub

'Private Sub m_lstAnswersList_Click()
'    answeredHandler
'End Sub

Private Sub m_lstAnswersList_AfterUpdate()
    answeredHandler
End Sub



