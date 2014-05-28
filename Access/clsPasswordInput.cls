VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPasswordInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsPasswordInput
' Author    : Dmitry Gorelenkov
' Date      : 25.01.2013
' Purpose   : Shows Password-InputBox
' Requires  : clsFuncsAccs, clsErrorHandler
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------

'TOOO BUGGYYY :(

Option Compare Database
Option Explicit
Private fAccs As clsFuncsAccs
Private ehError As clsErrorHandler

Private WithEvents m_frmPasswordForm As Access.Form
Attribute m_frmPasswordForm.VB_VarHelpID = -1
Private WithEvents m_btnOK As Access.CommandButton
Attribute m_btnOK.VB_VarHelpID = -1
Private WithEvents m_btnCancel As Access.CommandButton
Attribute m_btnCancel.VB_VarHelpID = -1
Private m_txtInput As Access.TextBox
Private m_lblLabel As Access.Label
Private m_sFormPwInput As String
Private m_sFormName As String
Private m_sNewFormName As String


'public func, um pw abzufragen
Public Function askPassword() As String
'    On Error GoTo askPassword_Error
    
    'falls kein formular vorhanden, erstellen
    If Not fAccs.formExists(m_sFormName) Then
        buildForm
    End If
    
    If m_sFormName = "" Then m_sFormName = m_sNewFormName
    DoCmd.OpenForm m_sFormName, acNormal, , , , acDialog
    
    Set m_frmPasswordForm = Forms(m_sFormName)

    
    askPassword = Nz(m_frmPasswordForm.txtInput.Value, "")

    'dialog schliessen
    DoCmd.Close acForm, m_sFormName, acSaveNo
    
'    Set m_txtInput = Nothing
'    Set m_btnOK = Nothing
'    Set m_btnCancel = Nothing
'    Set m_frmPasswordForm = Nothing


    Exit Function

'askPassword_Error:
'    If Err.Number <> 0 Then Err.Raise
    

End Function

'set some settings if needed
Public Sub init(sFormName As String)
    'todo test sFormName for allowed chars
    m_sNewFormName = sFormName
End Sub

Private Sub Class_Initialize()
    Set fAccs = New clsFuncsAccs
    Set ehError = New clsErrorHandler
    m_sNewFormName = "frmTEMP_DialogForm_DELETE_IF_NO_NEED"
    
'    Debug.Print "initialization .. "
'    buildForm
End Sub

Private Sub Class_Terminate()
    Debug.Print "termination .. wtf"
    Set m_txtInput = Nothing
    Set m_btnOK = Nothing
    Set m_btnCancel = Nothing
    Set m_frmPasswordForm = Nothing
    Set fAccs = Nothing
    Set ehError = Nothing
End Sub

'form konstruiren
Public Sub buildForm()
    Dim lngLineOfProc As Long
    Dim mdlPw As Access.Module
    
    'falls kein form existiert, neu erstellen
    If Not fAccs.formExists(m_sNewFormName) Then

        'dialog form object erstellen
        Set m_frmPasswordForm = CreateForm
        'name speichern
        m_sFormName = m_frmPasswordForm.Name


        m_frmPasswordForm.Caption = "Bitte Passwort eingeben"
        m_frmPasswordForm.PopUp = True
        m_frmPasswordForm.Modal = True
        m_frmPasswordForm.ControlBox = False
        m_frmPasswordForm.ShortcutMenu = False
        m_frmPasswordForm.RecordSelectors = False
        m_frmPasswordForm.NavigationButtons = False
        m_frmPasswordForm.BorderStyle = acDialog
        m_frmPasswordForm.Width = 3900
        m_frmPasswordForm.Section(acDetail).Height = 2225

        ' Create command button on form. btnOk
        Set m_btnCancel = CreateControl(m_sFormName, acCommandButton, , , , 450, 1740, 1245, 300)
        m_btnCancel.Caption = "Abbrechen"
        m_btnCancel.Name = "btnCancel"
        m_btnCancel.OnClick = "[Event Procedure]"


        'create btnCancel
        Set m_btnOK = CreateControl(m_sFormName, acCommandButton, , , , 2430, 1740, 1245, 300)
        m_btnOK.Caption = "OK"
        m_btnOK.Name = "btnOK"
        m_btnOK.OnClick = "[Event Procedure]"


        'create txtInput
        Set m_txtInput = CreateControl(m_sFormName, acTextBox, , , , 720, 1005, 2609, 315)
        m_txtInput.Name = "txtInput"
        m_txtInput.InputMask = "Password"

        'create label over input
        Set m_lblLabel = CreateControl(m_sFormName, acLabel, , , , 555, 315, 2985, 495)
        m_lblLabel.Name = "lblLabel"



        Set mdlPw = Forms(m_sFormName).Module

        lngLineOfProc = mdlPw.CreateEventProc("Click", "btnOK")
        mdlPw.InsertLines lngLineOfProc + 1, "Me.Visible = False"

        lngLineOfProc = mdlPw.CreateEventProc("Click", "btnCancel")
        mdlPw.InsertLines lngLineOfProc + 1, "Me.Visible = False" & vbCrLf & "Me.txtInput.value = """""

        'speichern schliessen
        DoCmd.Save acForm, m_sFormName
        DoCmd.Close acForm, m_sFormName, acSaveYes
        'umbenenen

        DoCmd.Rename m_sNewFormName, acForm, m_sFormName
    End If

    m_sFormName = m_sNewFormName

    'VBE code windos, hide..
    Application.VBE.MainWindow.Visible = False
    Set mdlPw = Nothing
End Sub

'form loeschen
Public Sub removeForm()
    On Error Resume Next
    DoCmd.Close acForm, m_sFormName, acSaveNo
    'dialog object loeschen
    DoCmd.DeleteObject acForm, m_sFormName
    Err.Clear
End Sub



'Private Sub m_btnOk_Click()
'    m_sFormPwInput = Nz(m_txtInput.Value, "")
'    closeInput
'End Sub
'
'Private Sub m_btnCancel_Click()
'    m_sFormPwInput = ""
'End Sub
'
'Private Sub closeInput()
'    Set m_btnOK = Nothing
'    Set m_btnCancel = Nothing
'    Set m_txtInput = Nothing
'    DoCmd.Close acForm, m_sPwFormName, acSaveNo
'End Sub

