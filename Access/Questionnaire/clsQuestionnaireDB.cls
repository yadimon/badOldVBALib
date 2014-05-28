VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuestionnaireDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsQuestionnaireDB
' Author    : Dmitry Gorelenkov
' Date      : 20.02.2013
' Changed   : 27.02.2013
' Purpose   : Daten von Datenbank in Questions und Questionnaire Klasse laden
' Requires  : clsDB, clsFuncs, clsQuestionnaire, clsQuestionnaireQuestionListBox, ADODB library
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
Private DB As New clsDB
Private f As New clsFuncs

Private m_lUserID As Long 'Call Agent
Private m_lCustomerID As Long 'Angerufener Kunde
Private m_bCreatedNewAnswers As Boolean 'mark if content of questionnaire changed

Public Property Get CreatedNewAnswers() As Boolean
    CreatedNewAnswers = m_bCreatedNewAnswers
End Property

Public Function init(lUserID_p As Long, lCustomerID_p As Long)
    m_lUserID = lUserID_p
    m_lCustomerID = lCustomerID_p
End Function


'erstellen Questionnaire, mit daten aus DB
Public Function createQuestionnaire(lQuestionnaireID As Long) As clsQuestionnaire
    Dim qnreQstnre As clsQuestionnaire
    Dim tmpQuestion As IQuestionnaireQuestion
    Dim rsQuestions As ADODB.Recordset
    Dim formRS As ADODB.Recordset
    Dim sFormName As String


    Set rsQuestions = DB.SQL("SELECT qsnrqstQuestionIDRef, qsnrqstQuestionPos, qsnrqstNextQuestionPos1, qsnrqstNextQuestionPos2, qsnrqstNextQuestionPos3 " & _
                             "FROM tblQSTNRE_QuestionnaireQuestions WHERE qsnrqstQstnnreIDRef = " & lQuestionnaireID, bOffline:=True)

    'Form parameter rauslesen
    Set formRS = DB.SQL("SELECT qstnreName FROM tblQSTNRE_Questionnaire WHERE qstnreID = " & lQuestionnaireID, bOffline:=True)
    '    formRS.ActiveConnection = Nothing
    If formRS.RecordCount < 1 Then Exit Function

    formRS.MoveFirst
    sFormName = formRS.Fields("qstnreName").Value



    Set qnreQstnre = New clsQuestionnaire
    Call qnreQstnre.init(sFormName, lQuestionnaireID)

    While Not rsQuestions.EOF
        'question erstellen
        Set tmpQuestion = createQuestion(rsQuestions.Fields("qsnrqstQuestionIDRef").Value, _
                                         rsQuestions.Fields("qsnrqstQuestionPos").Value, _
                                         f.ClngTry(rsQuestions.Fields("qsnrqstNextQuestionPos1").Value, -1), _
                                         f.ClngTry(rsQuestions.Fields("qsnrqstNextQuestionPos2").Value, -1), _
                                         f.ClngTry(rsQuestions.Fields("qsnrqstNextQuestionPos3").Value, -1))
        'in Questionnaire hinzufuegen
        Call qnreQstnre.addQuestion(tmpQuestion)
        Set tmpQuestion = Nothing

        'neaechstes Datensatz
        rsQuestions.MoveNext
    Wend



    Set formRS = Nothing
    Set rsQuestions = Nothing


    Set createQuestionnaire = qnreQstnre
End Function

'question mit DB ID und Questionnaire positionen erstellen
Public Function createQuestion(lQuestionID As Long, lPos As Long, Optional lNextQ1 As Long = -1, _
                                                                Optional lNextQ2 As Long = -1, _
                                                                Optional lNextQ3 As Long = -1) As IQuestionnaireQuestion
    Dim rsQuestionParams As ADODB.Recordset
    Dim rsQuestionParamsClone As ADODB.Recordset
    Dim qstQuestion As IQuestionnaireQuestion
    Dim sQType As String
    Dim sQuestionCaption As String
    
    'main info ueber die Frage
    Set rsQuestionParams = DB.SQL("SELECT qstnParamID, qstnParam, qstnValueOfParam FROM tblQSTNRE_Questions WHERE qstnID = " & lQuestionID, bOffline:=True)
    If rsQuestionParams.RecordCount < 1 Then
        Debug.Print "createQuestion empty recordset rsQuestionParams"
        Exit Function
    End If
    
    'type herausfinden
    Set rsQuestionParamsClone = rsQuestionParams.Clone
    rsQuestionParamsClone.Find ("qstnParam = 'Type'")
    sQType = Nz(rsQuestionParamsClone.Fields("qstnValueOfParam").Value, "")

    If sQType = "" Then
        Debug.Print "createQuestion: sQType = '' "
        Exit Function
    End If

    'abhaengig von FragenTyp questionObject erstellen
    '::::::::ListBox::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    If StrComp(sQType, "ListBox", vbTextCompare) = 0 Then
        Dim qstQuestionListBox As clsQuestionnaireQuestionListBox
        Dim mcAnswers As New clsMyCollection
        Dim iOwnAnswers As Integer
        Dim sOwnAnswersLabel As String
        Dim paramValue As Variant
        Dim bytMultiselect As Byte
        
        On Error Resume Next
        With rsQuestionParams
            
            While Not .EOF
                paramValue = .Fields("qstnValueOfParam").Value
                Select Case UCase(.Fields("qstnParam").Value)
                    'question answers..
                    Case UCase("Answer")
                        Call mcAnswers.Add(paramValue, .Fields("qstnParamID").Value)
                    'question TEXT
                    Case UCase("Question")
                        sQuestionCaption = paramValue
                    'question max own answers
                    Case UCase("OwnAnswers")
                        iOwnAnswers = CInt(f.ClngTry(paramValue, 0))
                    Case UCase("OwnAnswersLabel")
                        sOwnAnswersLabel = CStr(paramValue)
                    Case UCase("multiselect")
                        bytMultiselect = CByte(paramValue)
                    Case Else
                        'TODO error?
                End Select
                
                If err.Number <> 0 Then
                    Debug.Print "error by createQuestion: " & err.Number & " descr: " & err.Description
                    err.Clear
                End If
                
                .MoveNext
            Wend
        End With
        On Error GoTo 0


        Set qstQuestionListBox = New clsQuestionnaireQuestionListBox
        Call qstQuestionListBox.init(lQuestionID, lPos, lNextQ1, sQuestionCaption, mcAnswers, iOwnAnswers, sOwnAnswersLabel, bytMultiselect)
        Set qstQuestion = qstQuestionListBox

    ':::::::::::::::::YesNo::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ElseIf StrComp(sQType, "YesNo", vbTextCompare) = 0 Then
        Dim qstQuestionYesNo As clsQuestionnaireQuestionYesNo
        Set qstQuestionYesNo = New clsQuestionnaireQuestionYesNo
        Dim sYesLabel As String
        Dim sNoLabel As String
        Dim sNoAnswerLabel As String
        Dim lAnswerIDYes As Long
        Dim lAnswerIDNo As Long
        Dim lAnswerIDNoAnswer As Long
        Dim vParamValue As Variant
        Dim lParamID As Long
        
        With rsQuestionParams
            While Not .EOF
                vParamValue = .Fields("qstnValueOfParam").Value
                lParamID = .Fields("qstnParamID").Value
                
                Select Case UCase(.Fields("qstnParam").Value)
                'question TEXT
                Case UCase("Question")
                    sQuestionCaption = Nz(vParamValue, "")

                'question YES label, and ID of answer
                Case UCase("YesLabel")
                    sYesLabel = Nz(vParamValue, "")
                    lAnswerIDYes = lParamID

                'question NO label, and ID of answer
                Case UCase("NoLabel")
                    sNoLabel = Nz(vParamValue, "")
                    lAnswerIDNo = lParamID
                    
                'question NOANSWER label, and ID of answer
                Case UCase("NoAnswerLabel")
                    sNoAnswerLabel = Nz(vParamValue, "")
                    lAnswerIDNoAnswer = lParamID
                End Select

                .MoveNext
            Wend
        End With

        Call qstQuestionYesNo.init(lQuestionID, lAnswerIDYes, lAnswerIDNo, lAnswerIDNoAnswer, lPos, _
                                    lNextQ1, lNextQ2, lNextQ3, sQuestionCaption, sYesLabel, sNoLabel, sNoAnswerLabel)
        Set qstQuestion = qstQuestionYesNo
    End If



    Set createQuestion = qstQuestion
End Function

'daten von einer questionnaire speichern
Public Function saveQuestionnaire(qnrQstnnaire_p As clsQuestionnaire)
    'save Questionnaire
    Dim aColumns As Variant
    Dim aValues As Variant
    Dim lAffected As Long
    Dim answeredQnreID As Long
    Dim bSavedAllQuestions As Boolean
    
    'create Column and Values array, to easy save later
    aColumns = Array("answqstnrQuestionnaireIDRef", "answqstnrDate", "answqstnrUserIDRef", "answqstnrCustomerIDRef")
    aValues = Array(qnrQstnnaire_p.id, f.SQLString_formatTimeToSQLString(Now()), m_lUserID, m_lCustomerID)
    
    Call DB.SQL("INSERT INTO tblQSTNRE_AnsweredQuestionnaire (" & Join(aColumns, ", ") & ")" & _
            " VALUES(" & Join(aValues, ", ") & ")", , , lAffected)
            
    'neu hinzugefuegte ID bestimmen. TODO bessere methode?
    answeredQnreID = getLastInsertedID("answqstnrID", "tblQSTNRE_AnsweredQuestionnaire", aColumns, aValues)
    
    If lAffected <> 1 Then
        f.dbg "saveQuestionnaire lAffected " & lAffected
        Exit Function
    End If
    
    'save each question
    Dim qstQuestion As IQuestionnaireQuestion
    Dim i As Long
    bSavedAllQuestions = True
    For i = 0 To qnrQstnnaire_p.m_mcQuestions.getSize - 1
        Set qstQuestion = qnrQstnnaire_p.m_mcQuestions.getValue(i)
        
        bSavedAllQuestions = bSavedAllQuestions And saveQuestion(qstQuestion, answeredQnreID)
        
        Set qstQuestion = Nothing
    Next i
    
    saveQuestionnaire = bSavedAllQuestions
End Function

'Question object speichern
Public Function saveQuestion(qstQuestion As IQuestionnaireQuestion, lAnsweredQnreID As Long) As Boolean
    Dim vAnswer As Variant
    Dim lAnswerID As Long
    Dim vNewAnswer As Variant
    Dim lNewAnswerID As Long
    Dim aNotSavedValues As Variant    'answers that cannot be saved
    aNotSavedValues = Array()
    saveQuestion = False
    
    
    For Each vAnswer In qstQuestion.getAnswer.GetValues
        lAnswerID = f.ClngTry(vAnswer, -1)
        If Not saveAnswerToDB(lAnsweredQnreID, lAnswerID) Then
            Call f.arrayPush(aNotSavedValues, lAnswerID)
        End If
    Next vAnswer
    
    
    'je nach Question Typ weiter behandeln
    Select Case UCase(qstQuestion.qType)
    Case UCase("ListBox")
        Dim clsQuestionListBox As clsQuestionnaireQuestionListBox
        Set clsQuestionListBox = qstQuestion

        'neue antworten
        For Each vNewAnswer In clsQuestionListBox.getNewAnswers.GetValues
            Debug.Assert VarType(vNewAnswer) = vbString
            'in DB hinzufuegen, zu dem Question
            lNewAnswerID = createAnswer(qstQuestion.id, CStr(vNewAnswer))
            'und als Antwort speichern
            If Not saveAnswerToDB(lAnsweredQnreID, lNewAnswerID) Then
                Call f.arrayPush(aNotSavedValues, vAnswer)
            Else
                'markieren, dass es neue inhalte gibt
                m_bCreatedNewAnswers = True
            End If
        Next vNewAnswer
        
    Case UCase("YesNo")
        'keine weitere behandlung
    End Select
    
    
    If UBound(aNotSavedValues) = -1 Then
        saveQuestion = True
    End If
End Function

'answer in DB speichern
Private Function saveAnswerToDB(lAnsweredQnreID As Long, lAnswer As Long) As Boolean
    Dim aColumns As Variant
    Dim aValues As Variant
    Dim lAffected As Long
    
    'pruefen IDs
    If lAnsweredQnreID < 0 Or lAnswer < 0 Then
        saveAnswerToDB = False
        Exit Function
    End If
    
    Call f.arrayPush(aColumns, "answAnswrsAnswrdQstnreIDRef")
    Call f.arrayPush(aValues, lAnsweredQnreID)

    Call f.arrayPush(aColumns, "answAnswrsAnswerRef")
    Call f.arrayPush(aValues, lAnswer)
        
    Call DB.SQL("INSERT INTO tblQSTNRE_AnsweredQuestionnaireAnswers (" & Join(aColumns, ", ") & ")" & _
            " VALUES(" & Join(aValues, ", ") & ")", , , lAffected)
            
    saveAnswerToDB = IIf((lAffected = 1), True, False)

End Function

'Answer zu Question hinzufuegen, und ID von dem neuem Answer zuruckgeben
Public Function createAnswer(lQuestionID As Long, sValue As String)
    Dim aValues As Variant
    Dim sValues As String
    Dim aColumns As Variant
    Dim sColumns As String
    Dim lNewAnswerID As Long
    
    aColumns = Array("qstnID", "qstnParam", "qstnValueOfParam")
    aValues = Array(lQuestionID, "'Answer'", "'" & sValue & "'")
    sColumns = Join(aColumns, ", ")
    sValues = Join(aValues, ", ")
    Call DB.SQL("INSERT INTO tblQSTNRE_Questions (" & sColumns & ")" & _
            " VALUES(" & sValues & ")")
            
    'neu hinzugefuegtte ID bestimmen.
    lNewAnswerID = getLastInsertedID("qstnParamID", "tblQSTNRE_Questions", aColumns, aValues)
                  
    createAnswer = lNewAnswerID
End Function

'letzte hinzugefuegte ID zuruckgeben.
'schlussel spalte sIDColumn, tabelle- sTableName, aColumns - hinzugefuegte spalten, aValues - hinzug. werte.
Private Function getLastInsertedID(sIDColumnName As String, sTableName As String, aColumns As Variant, aValues As Variant) As Long
    Dim i As Integer
    Dim sSQLGetNewID As String
    
    Debug.Assert IsArray(aColumns) And IsArray(aValues) And UBound(aColumns) = UBound(aValues)
        
    sSQLGetNewID = "SELECT MAX(" & sIDColumnName & ") FROM " & sTableName
    For i = 0 To UBound(aColumns)
        If i <= UBound(aValues) Then
            sSQLGetNewID = f.SQLString_addWhereClause(sSQLGetNewID, CStr(aColumns(i)) & "=" & CStr(aValues(i)))
        End If
    Next i
    
    getLastInsertedID = DB.SQLValue(sSQLGetNewID)
End Function




