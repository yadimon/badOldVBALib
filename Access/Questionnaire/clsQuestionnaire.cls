VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuestionnaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsQuestionnaire
' Author    : Dmitry Gorelenkov
' Date      : 01.02.2013
' Changed   : 09.04.2013
' Purpose   : Create Questionnaire
' Requires  : IQuestionnaireQuestion, clsMyCollection
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
Private f As clsFuncs
Private fAccs As clsFuncsAccs
Private bInitDone As Boolean
Private m_sName As String
Public m_mcQuestions As clsMyCollection 'todo private
Private Const CM As Long = 567
Private m_frmQForm As Access.Form 'created form
Public m_sFormName As String
Private m_formHandler As clsAccsFormHandler
Private m_mcQuestionsAnswered As clsMyCollection 'answered questions to analyze
Private m_lQuestionsLeftPosTw As Long 'left TW coord of all shown control-blocks
Private m_lQuestionsTopPosTw As Long 'top TW coord of all shown control-blocks
Private m_lID As Long

Private m_mcLastCtrlsProps As clsMyCollection 'properties for Last (END) controls in the Questionnaire
Private m_mcLastCtrlsRefs As clsMyCollection 'Last (END) controls (refs) in the Questionnaire
Private m_mcFooterCtrlsProps As clsMyCollection 'properties for FOOT controls in the Questionnaire
Private m_mcFooterCtrlsRefs As clsMyCollection 'FOOT controls (refs) in the Questionnaire
Private m_mcAllNonStaticControls As clsMyCollection 'all active (changeble) controls

Private m_qstLastAnswered As IQuestionnaireQuestion 'last answered question

Private WithEvents m_cmdSaveForm As Access.CommandButton 'save button in Footer, with events
Attribute m_cmdSaveForm.VB_VarHelpID = -1
Private WithEvents m_cmdNextQuestion As Access.CommandButton 'next button on Questionnaire
Attribute m_cmdNextQuestion.VB_VarHelpID = -1
Private WithEvents m_cmdPrevQuestion As Access.CommandButton 'prev button on Questionnaire
Attribute m_cmdPrevQuestion.VB_VarHelpID = -1
Private m_txtFocusRemover As Access.TextBox 'just for remove focus from other elements
Private m_lblInfo As Access.Label 'info label z.B. anzeige: 4 von 15 beantwortet

Private m_bRemoveAfter As Boolean 'mark if the generated form have to be changed/again generated

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::EVENTS::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Event OnSave(ByRef bResult As Boolean) 'save event, trigger by Clicking on Save Button, bResult must contain a Result of saving procedure
'Public Event Saved(ByRef bResult) 'saved event with result of saving in bResult


Public Property Get Name() As String
    Name = m_sName
End Property

Public Property Get id() As Long
    id = m_lID
End Property

Public Function init(sName_p As String, lID_p As Long, Optional lLeftPosCoord As Long = 567, Optional lTopPosCoord As Long = 567) As Boolean
    Dim bResult As Boolean

    m_sName = sName_p
    m_lID = lID_p
    m_lQuestionsLeftPosTw = lLeftPosCoord
    m_lQuestionsTopPosTw = lTopPosCoord


    Dim mcTmpControlDescription As clsMyCollection
    Dim cpropsTmpCtrlProperty As clsControlProperties

    Set cpropsTmpCtrlProperty = New clsControlProperties
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_cmdSaveForm", "Name")
        Call .Add("[Event Procedure]", "OnClick")
        Call .Add(acCommandButton, "Type")
        Call .Add(4 * CM, "Left")
        Call .Add(1 * CM, "Top")
        Call .Add(4 * CM, "Width")
        Call .Add(1 * CM, "Height")
        Call .Add("Speichern", "Caption")
    End With

    Call cpropsTmpCtrlProperty.addControlPropsByCol(mcTmpControlDescription)
    Call m_mcLastCtrlsProps.Add(cpropsTmpCtrlProperty, cpropsTmpCtrlProperty.ControlName)
    Call m_mcLastCtrlsRefs.Add(Null, cpropsTmpCtrlProperty.ControlName)


    'weiter button
    Set cpropsTmpCtrlProperty = New clsControlProperties
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_cmdNextQuestion", "Name")
        Call .Add("[Event Procedure]", "OnClick")
        Call .Add(acCommandButton, "Type")
        Call .Add(2268, "Width")
        Call .Add(567, "Height")
        Call .Add(5000 + m_lQuestionsTopPosTw, "Top")
        Call .Add(6225 + m_lQuestionsLeftPosTw, "Left")
        Call .Add("Weiter", "Caption")
        Call .Add("Calibri", "FontName")
        Call .Add(11, "FontSize")
        Call .Add(400, "FontWeight")
        Call .Add(False, "FontItalic")
        Call .Add(-2147483630, "ForeColor")
    End With

    Call cpropsTmpCtrlProperty.addControlPropsByCol(mcTmpControlDescription)
    Call m_mcFooterCtrlsProps.Add(cpropsTmpCtrlProperty, cpropsTmpCtrlProperty.ControlName)
    Call m_mcFooterCtrlsRefs.Add(Null, cpropsTmpCtrlProperty.ControlName)

    'zuruck button
    Set cpropsTmpCtrlProperty = New clsControlProperties
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_cmdPrevQuestion", "Name")
        Call .Add("[Event Procedure]", "OnClick")
        Call .Add(acCommandButton, "Type")
        Call .Add(2268, "Width")
        Call .Add(567, "Height")
        'BUG? 4918 + m_lQuestionsTopPosTw is OK, by 4919 + m_lQuestionsTopPosTw, lists resetted :(
        'if set button "invisible" -> reset list.. why dont know
        'EDIT. have to set Section(acDetail).hight bigger then last controls that will moved..
        Call .Add(5000 + m_lQuestionsTopPosTw, "Top")
        Call .Add(0 + m_lQuestionsLeftPosTw, "Left")
        Call .Add("Zurück", "Caption")
        Call .Add("Calibri", "FontName")
        Call .Add(11, "FontSize")
        Call .Add(400, "FontWeight")
        Call .Add(False, "FontItalic")
        Call .Add(-2147483630, "ForeColor")
    End With
    Call cpropsTmpCtrlProperty.addControlPropsByCol(mcTmpControlDescription)
    Call m_mcFooterCtrlsProps.Add(cpropsTmpCtrlProperty, cpropsTmpCtrlProperty.ControlName)
    Call m_mcFooterCtrlsRefs.Add(Null, cpropsTmpCtrlProperty.ControlName)

    'focus remover
    Set cpropsTmpCtrlProperty = New clsControlProperties
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_txtFocusRemover", "Name")
        Call .Add(acTextBox, "Type")
        Call .Add(0, "Width")
        Call .Add(0, "Height")
        Call .Add(0, "Top")
        Call .Add(0, "Left")
    End With
    Call cpropsTmpCtrlProperty.addControlPropsByCol(mcTmpControlDescription)
    Call m_mcFooterCtrlsProps.Add(cpropsTmpCtrlProperty, cpropsTmpCtrlProperty.ControlName)
    Call m_mcFooterCtrlsRefs.Add(Null, cpropsTmpCtrlProperty.ControlName)

    'Info feld (3 von 15 ca. 3 mins ...)
    Set cpropsTmpCtrlProperty = New clsControlProperties
    Set mcTmpControlDescription = New clsMyCollection
    With mcTmpControlDescription
        Call .Add("m_lblInfo", "Name")
        Call .Add(acLabel, "Type")
        Call .Add(2775, "Width")
        Call .Add(465, "Height")
        Call .Add(5625, "Top")
        Call .Add(3150, "Left")
        Call .Add("Calibri", "FontName")
        Call .Add(13, "FontSize")
        Call .Add(400, "FontWeight")
        Call .Add(False, "FontItalic")
        Call .Add(0, "ForeColor")
        Call .Add(2, "TextAlign") 'center
    End With
    Call cpropsTmpCtrlProperty.addControlPropsByCol(mcTmpControlDescription)
    Call m_mcFooterCtrlsProps.Add(cpropsTmpCtrlProperty, cpropsTmpCtrlProperty.ControlName)
    Call m_mcFooterCtrlsRefs.Add(Null, cpropsTmpCtrlProperty.ControlName)
'    Call m_mcLastCtrlsRefs.Add(Null, cpropsTmpCtrlProperty.ControlName)

        '    'focus remover tmp
        '    Set cpropsTmpCtrlProperty = New clsControlProperties
        '    Set mcTmpControlDescription = New clsMyCollection
        '    With mcTmpControlDescription
        '        Call .Add("m_txtFocusRemover22", "Name")
        '        Call .Add(acTextBox, "Type")
        '        Call .Add(200, "Width")
        '        Call .Add(200, "Height")
        '        Call .Add(8000, "Top")
        '        Call .Add(1000, "Left")
        '    End With
        '    Call cpropsTmpCtrlProperty.addControlPropsByCol(mcTmpControlDescription)
        '    Call m_mcFooterCtrlsProps.Add(cpropsTmpCtrlProperty, cpropsTmpCtrlProperty.ControlName)
        '    Call m_mcFooterCtrlsRefs.Add(Null, cpropsTmpCtrlProperty.ControlName)



        bInitDone = True

        init = bResult
    End Function

'builds if needed, and opens form. Returns builded form name.
Public Function openAndBuildIfNotExist(Optional ByVal sFormSaveName As String = "", _
                                        Optional ByVal lHeightTW As Long = -1, Optional ByVal lWidthTW As Long = -1, _
                                        Optional bRemoveAfter As Boolean = False) As String
    If Not bInitDone Then Exit Function
    Dim sPostFix As String
    Dim MD5 As New MD5
'    On Error GoTo Final
'    Application.Echo False
    
    'generate sFormSaveName
    sPostFix = MD5.DigestStrToHexStr(CStr(m_lID))
    sFormSaveName = IIf(sFormSaveName = "", "Questionnaire_Generated_Form", sFormSaveName) & "_" & sPostFix
    
    
    'in case there is no Form, create one.
    If Not fAccs.formExists(sFormSaveName) Then
        Call buildForm(sFormSaveName, lHeightTW, lWidthTW)
    End If
    
    'open the Form
    Set m_frmQForm = OpenForm(sFormSaveName, lHeightTW, lWidthTW)
    
    'bind controls to Question Objects
    Call bindControls(m_frmQForm)
    
    'show first question
    Call showQuestion(0)
    
    'refreshMe
'final:
'    Application.Echo True
'    On Error GoTo 0
    
    'wait until form is closed 'TODO ready/done function?
    While fAccs.isFormLoaded(sFormSaveName)
        DoEvents
        f.Sleep (10)
    Wend
    
    'if something changed in form, or have to be rebuilded, delete form after routine
    'Important! Form module will be removed only after all VBA macro routine done!
    '  => cannot create form with the same name, after try to remove it, in the same macro routine:/
    If m_bRemoveAfter Or bRemoveAfter Then
        DoCmd.DeleteObject acForm, sFormSaveName
        sFormSaveName = ""
    End If
    
    
    'return created form name
    openAndBuildIfNotExist = sFormSaveName
    
End Function

'create and save Questionnaire Form
Public Function buildForm(sSaveName As String, Optional lHeightTW As Long = -1, Optional lWidthTW As Long = -1) As Access.Form
    

'    'change form name, in case there is one
'    Dim i As Integer
'    Dim sNameOrig As String
'    sNameOrig = sSaveName
'    For i = 1 To 10000
'        If Not fAccs.formExists(sSaveName) Then
'            Exit For
'        Else
'            sSaveName = sNameOrig & "_" & i
'        End If
'    Next i

    'create dialog form object
    Set m_frmQForm = CreateForm
    With m_frmQForm
        'save name
        m_sFormName = .Name
        .Section(acDetail).BackColor = -2147483606 'design 2
        .Caption = m_sName
        .PopUp = True
        .Modal = False
        .BorderStyle = acDialog
        .ControlBox = True
        .ShortcutMenu = False
        .RecordSelectors = False
        .NavigationButtons = False
        .FitToScreen = True
        .AutoResize = False
        .AutoCenter = True
'        .BorderStyle = acDialog
'        .Width = IIf(lWidthTW > 0, lWidthTW, 10 * CM)
'        .Section(acDetail).Height = IIf(lHeightTW > 0, lHeightTW, 10 * CM)
'        .InsideHeight = 23 * CM
'        .InsideWidth = 25 * CM
        .HasModule = True
        
    End With

    
    
    'build question controls on form
    Dim qstQuestion As Variant 'IQuestionnaireQuestion
    Dim mcNewControls As clsMyCollection
    
    For Each qstQuestion In m_mcQuestions.GetValues
        'create controls with the design
'        Set mcNewControls = m_formHandler.addControlsDown(m_frmQForm, qstQuestion.getControlsProps)
        Set mcNewControls = m_formHandler.addControls(m_frmQForm, qstQuestion.getControlsProps)
        'hide created controls
        Call fAccs.setPropertyForControls(mcNewControls, "visible", False)
    Next qstQuestion
    
    
    'add Last Page Controls ( save button etc.. )
    Set mcNewControls = m_formHandler.addControls(m_frmQForm, m_mcLastCtrlsProps)
    'must be invisible first
    Call fAccs.setPropertyForControls(mcNewControls, "visible", False)
    
    'add Footer Controls
    Set mcNewControls = m_formHandler.addControls(m_frmQForm, m_mcFooterCtrlsProps)

    
    
    'save and close
    DoCmd.save acForm, m_sFormName
    DoCmd.Close acForm, m_sFormName, acSaveYes

    'rename
    DoCmd.Rename sSaveName, acForm, m_sFormName
    m_sFormName = sSaveName
End Function

'opens form 'TODO add to fAccs?
Private Function OpenForm(sFormName As String, Optional lHeightTW As Long = -1, Optional lWidthTW As Long = -1) As Access.Form
    Dim frmRetForm As Access.Form

    DoCmd.OpenForm sFormName, acNormal, , , , acWindowNormal
    'fuer den fall, falls hauptfenster von Access versteckt ist. naja BAD :/
    Call f.ShowWindow(Forms(sFormName).hWnd, True)
    
    Set frmRetForm = Forms(sFormName)


    With frmRetForm
        'if no size args, autoresize
        If lHeightTW <= 0 Or lWidthTW <= 0 Then
            'set width and height of form, by controls of form
            .InsideWidth = IIf(lWidthTW > 0, lWidthTW, m_formHandler.findMostRightInForm(frmRetForm) + m_lQuestionsLeftPosTw)
            .InsideHeight = IIf(lHeightTW > 0, lHeightTW, m_formHandler.findMostDownInForm(frmRetForm) + m_lQuestionsTopPosTw)

            '        frmRetForm.AutoResize = True
        Else    'else set size

            '        With frmRetForm
            '            .InsideHeight = IIf(lHeightTW > 0, lHeightTW, frmRetForm.Section(acDetail).Height + 1 * CM)
            '            .InsideWidth = IIf(lWidthTW > 0, lWidthTW, frmRetForm.Width + 1 * CM)
            '        End With

            .InsideHeight = lHeightTW
            .InsideWidth = lWidthTW
        End If
        
        'section entsprechend einstellen. Wichtig!! damit andere controls passen..
        .Section(acDetail).Height = .InsideHeight
    End With
    Set OpenForm = frmRetForm
End Function

'bind controls from form:frmQForm_p to question objects
Private Function bindControls(frmQForm_p As Access.Form) As Boolean
    Dim sCtrlName As Variant
    Dim tryCtrl As Control


    If frmQForm_p Is Nothing Or IsNull(frmQForm_p) Then
        Debug.Print "bindControls: form arg is nothing"
        Exit Function
    End If

    If Not fAccs.isFormLoaded(frmQForm_p.Name) Then    'frmQForm_p.Name hm.. fehleranfaellig?
        Debug.Print "bindControls: wrong form, not loaded"
        Exit Function
    End If

    Dim qstQuestion As Variant    'IQuestionnaireQuestion
    'for each Question object
    For Each qstQuestion In m_mcQuestions.GetValues
        'try to find and bind controls on the form
        Call qstQuestion.bindControlsByOpenedForm(frmQForm_p)
    Next qstQuestion

    Dim allFormCtrls As New clsMyCollection
    Call allFormCtrls.merge(m_mcLastCtrlsRefs)
    Call allFormCtrls.merge(m_mcFooterCtrlsRefs)
    'All form controls
    If fAccs.isFormLoaded(m_frmQForm.Name) Then
        For Each sCtrlName In allFormCtrls.getKeys
            Set tryCtrl = fAccs.tryGetControl(m_frmQForm.Name, CStr(sCtrlName), Null)
            If Not tryCtrl Is Nothing Then

                'bind event controls
                Select Case sCtrlName 'TODO try eval?
                    'footer save button
                    Case "m_cmdSaveForm"
                        Set m_cmdSaveForm = tryCtrl
                        
                    'next button
                    Case "m_cmdNextQuestion"
                        Set m_cmdNextQuestion = tryCtrl
                        
                    'prev button
                    Case "m_cmdPrevQuestion"
                        Set m_cmdPrevQuestion = tryCtrl
                        
                    'focus remover
                    Case "m_txtFocusRemover"
                        Set m_txtFocusRemover = tryCtrl
                    
                    Case "m_lblInfo"
                        Set m_lblInfo = tryCtrl
                End Select


                'in collection binding
                If m_mcLastCtrlsRefs.isInArray(CStr(sCtrlName)) Then Call m_mcLastCtrlsRefs.setValue(tryCtrl, sCtrlName)
                If m_mcFooterCtrlsRefs.isInArray(CStr(sCtrlName)) Then Call m_mcFooterCtrlsRefs.setValue(tryCtrl, sCtrlName)

                Set tryCtrl = Nothing
            Else
                Debug.Print "bindControlsByOpenedForm: control " & sCtrlName & " not found on form"
            End If
        Next sCtrlName
        
        'save all active controls
        Call m_mcAllNonStaticControls.merge(m_mcFooterCtrlsRefs)
        Call m_mcAllNonStaticControls.merge(m_mcLastCtrlsRefs)
        For Each qstQuestion In m_mcQuestions.GetValues
            Call m_mcAllNonStaticControls.merge(qstQuestion.getControlsRefs)
        Next qstQuestion
        
        
        
    Else
        Debug.Print "called bindControlsByOpenedForm, with not loaded form as argument"
    End If

    bindControls = True
End Function

'add new Question to Questionnaire
Public Function addQuestion(ByRef qstQuestion As IQuestionnaireQuestion) ' As Variant)
    Dim lPosition As Long
    
    Set qstQuestion.qstnreParent = Me
    
    'falls position nicht definiert, abbrechen
    If qstQuestion.Position < 0 Then
        Exit Function
    Else
        lPosition = qstQuestion.Position
    End If
    
    Call m_mcQuestions.Add(qstQuestion, CStr(qstQuestion.id))
End Function

'returns question object with the lPosition position
Private Function getQuestionByPosition(lPosition As Long) As IQuestionnaireQuestion
    If lPosition < 0 Then Exit Function
    Dim qst As Variant 'IQuestionnaireQuestion
    For Each qst In m_mcQuestions.GetValues
        If qst.Position = lPosition Then
            Exit For
        End If
    Next qst
    
    Set getQuestionByPosition = qst
End Function


'handler in case any question is answered
Public Function questionAnsweredHandler(ByRef qstQuestion_p As IQuestionnaireQuestion)
    'TODO?
    Set m_qstLastAnswered = qstQuestion_p
    

End Function

'sets mark to remove generated form after routine
Public Function removeAfter()
    m_bRemoveAfter = True
End Function


'moves and shows the collection of controls of question with lPosition on the form, after next visible control
Private Function showQuestion(lPosition As Long)
    Dim questionToShow As IQuestionnaireQuestion
'    'define what question to show
    Set questionToShow = getQuestionByPosition(lPosition)
    If questionToShow Is Nothing Then Exit Function
    
    Call showNext(questionToShow.getControlsRefs)
        
End Function

'show next controls after last visible question
Private Function showNext(mcControls_p As clsMyCollection)
    Dim leftTw As Long
    Dim topTw As Long
    Dim allQstShownControls As New clsMyCollection
    'all answered question controls
    Dim qst As Variant
    For Each qst In m_mcQuestionsAnswered.GetValues
        Call allQstShownControls.merge(qst.getControlsRefs)
    Next qst
    
    
    'top and left space
    leftTw = m_lQuestionsLeftPosTw
    topTw = m_lQuestionsTopPosTw
    
    'make question controls visible
    Call fAccs.setPropertyForControls(mcControls_p, "Visible", True)
    
    'move question controls to next point
    Call fAccs.moveControls(mcControls_p, leftTw, topTw)
    
    '* by offset first and then visible order, all checkboxes are deselected O_o... why? dont know.. bug or whatever
End Function

'show final controls of the Form, after all questions answered (like save button)
Private Function showFinalControls()
    Call showNext(m_mcLastCtrlsRefs)
    'set focus on SAVE BUTTON
    m_cmdSaveForm.SetFocus
'    Call fAccs.setPropertyForControls(m_mcFooterCtrlsRefs, "Visible", False)
End Function


''deprecated?
''hide all questions after lPosition
'Private Function hideAllAfter(lPosition As Long)
'    Dim question As IQuestionnaireQuestion
'    Dim sKey As Variant
'
'    For Each sKey In m_mcQuestions.getKeys
'        Set question = m_mcQuestions.getValue(sKey)
'        If question.Position > lPosition Then
'            Call hideQuestion(question)
'        End If
'    Next sKey
'
'    'hide Footer...
'    Call fAccs.setPropertyForControls(m_mcLastCtrlsRefs, "Visible", False)
'End Function
'hide one question
Private Function hideQuestion(qstQuestion As IQuestionnaireQuestion)
    Call fAccs.setPropertyForControls(qstQuestion.getControlsRefs, "Visible", False)
End Function

'try to find next not loaded question position after lPosition
Private Function getNextQuestionPosition(Optional lPosition As Long = -1)
    Dim qst As Variant
    Dim minPos As Long
    minPos = LONGMAX
    For Each qst In m_mcQuestions.GetValues
        'falls noch nicht beantwortet
        If Not m_mcQuestionsAnswered.isInArray(CStr(qst.id)) Then
            If qst.Position > lPosition Then
                minPos = f.min(minPos, qst.Position)
            End If
        End If
    Next qst
    
    If minPos = LONGMAX Then minPos = -1
    
    getNextQuestionPosition = minPos
End Function

'refresh form controls, depend on questionnaire status.
Private Function refreshMe()
    Dim qstLastSaved As IQuestionnaireQuestion
    Dim nextQuestionPos As Long
    Dim nextQst As IQuestionnaireQuestion
    
    If m_mcQuestionsAnswered.getSize <= 0 Then
        nextQuestionPos = 0 'fist question
    Else
        Set qstLastSaved = m_mcQuestionsAnswered(m_mcQuestionsAnswered.getSize - 1)
        nextQuestionPos = f.ClngTry(qstLastSaved.NextQstPosition, -1)
    End If
    
    
    If nextQuestionPos < 0 Then
        nextQuestionPos = getNextQuestionPosition(qstLastSaved.Position)
    End If
    
    
    'Anzeige aktualisieren
    Dim iAllQstns As Integer
    Dim sRestTime As String
    iAllQstns = m_mcQuestions.getSize
    sRestTime = f.FormatSeconds((iAllQstns - nextQuestionPos) * 20, "nn:ss")
    m_lblInfo.Caption = nextQuestionPos + 1 & " von " & iAllQstns & "   ca. " & sRestTime & " Min."
    
    
    'hide all active controls
    Call hideNonStaticControls
    
    
    'if there no question to show anymore, show final controls of questionnaire
    If nextQuestionPos = -1 Then
        Call showFinalControls
        m_cmdPrevQuestion.Visible = True
        m_lblInfo.Visible = True
        m_lblInfo.Caption = "Fertig!"
    Else
        Call showQuestion(nextQuestionPos)
        Call fAccs.setPropertyForControls(m_mcFooterCtrlsRefs, "Visible", True)
        
        'in case the question was already answered, mark him as the last answered
        Set nextQst = getQuestionByPosition(nextQuestionPos)
        If nextQst.answered Then
            Set m_qstLastAnswered = nextQst
        End If
    End If
End Function

'hides all active controls
Private Function hideNonStaticControls()
    m_txtFocusRemover.SetFocus
    Call fAccs.setPropertyForControls(m_mcAllNonStaticControls, "Visible", False)
End Function


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::EVENT HANDLERS::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Private Sub m_cmdSaveForm_Click()
    'TODO, user ID, customerid
    If Not DEBUGMODE Then On Error GoTo m_cmdSaveForm_Click_Error
    Dim bResult As Boolean
    bResult = True
    RaiseEvent OnSave(bResult)
    
'    Debug.Print "bResult in questionnaire by onsave = " & bResult
    
    DoCmd.Close acForm, m_frmQForm.Name, acSaveNo
    
    Exit Sub

m_cmdSaveForm_Click_Error:
    ehError.catchErr ("m_cmdSaveForm_Click of Klassenmodul clsQuestionnaire")
    
End Sub

'next button klick
Private Sub m_cmdNextQuestion_Click()
    If m_qstLastAnswered Is Nothing Then Exit Sub
    
    
    Call m_mcQuestionsAnswered.Add(m_qstLastAnswered, CStr(m_qstLastAnswered.id))
    Call refreshMe
    
'    'den Question verstecken
'    Call hideQuestion(m_qstLastAnswered)
'    nextQuestionPos = f.ClngTry(m_qstLastAnswered.NextQstPosition, -1)
'
'    If nextQuestionPos < 0 Then
'        nextQuestionPos = getNextQuestionPosition(m_qstLastAnswered.Position)
'    End If
'
'    'if there no question to show anymore, show final controls of questionnaire
'    If nextQuestionPos = -1 Then
'        Call showFinalControls
'    Else
'        Call showQuestion(nextQuestionPos)
'    End If
End Sub

'prev button Click
Private Sub m_cmdPrevQuestion_Click()
    Dim qstPrevious As IQuestionnaireQuestion
    
    If m_mcQuestionsAnswered.getSize > 0 Then
        Set qstPrevious = m_mcQuestionsAnswered(m_mcQuestionsAnswered.getSize - 1)
        Set m_qstLastAnswered = qstPrevious
        Call m_mcQuestionsAnswered.pop
        Call refreshMe
    End If

End Sub



Private Sub Class_Initialize()
    Set f = New clsFuncs
    Set m_mcQuestions = New clsMyCollection
    Set fAccs = New clsFuncsAccs
    Set m_formHandler = New clsAccsFormHandler
    Set m_mcQuestionsAnswered = New clsMyCollection
    Set m_mcLastCtrlsProps = New clsMyCollection
    Set m_mcLastCtrlsRefs = New clsMyCollection
    Set m_mcFooterCtrlsProps = New clsMyCollection
    Set m_mcFooterCtrlsRefs = New clsMyCollection
    Set m_mcAllNonStaticControls = New clsMyCollection
End Sub

Private Sub Class_Terminate()
    Set f = Nothing
    Set m_mcQuestions = Nothing
    Set fAccs = Nothing
    Set m_formHandler = Nothing
    Set m_mcQuestionsAnswered = Nothing
    Set m_mcLastCtrlsProps = Nothing
    Set m_mcLastCtrlsRefs = Nothing
    Set m_mcFooterCtrlsProps = Nothing
    Set m_mcFooterCtrlsRefs = Nothing
    Set m_mcAllNonStaticControls = Nothing
End Sub





