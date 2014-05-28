VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEasyPwInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsEasyPwInput
' Author    : Dmitry Gorelenkov
' Date      : 12.03.2013
' Changed   : 17.04.2013
' Purpose   : Asking Password..
' Requires  : clsFuncs, Form with password fields.
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_formName As String
Private WithEvents m_frmPwInput As Access.Form 'form to show
Private WithEvents m_cmdConfirm As Access.CommandButton 'ok button
Private WithEvents m_cmdCancel As Access.CommandButton 'cancel button
Attribute m_cmdCancel.VB_VarHelpID = -1
Private m_sInput As String 'input pw

Public Sub init(sPwInputFormName As String)
    m_formName = sPwInputFormName
End Sub

Public Function askPassword(Optional sFormCaption As String = "Bitte Passwort eingeben") As String
    Set m_frmPwInput = Form_frmPasswordInput
    m_frmPwInput.Visible = True
    m_frmPwInput.Caption = sFormCaption
    Set m_cmdConfirm = m_frmPwInput.btnOK
    m_cmdConfirm.OnClick = "[Event Procedure]"
    Set m_cmdCancel = m_frmPwInput.btnCancel
    m_cmdCancel.OnClick = "[Event Procedure]"
    
    'wait for input
    While Not m_frmPwInput Is Nothing
        DoEvents
        f.Sleep (10)
    Wend
    
    askPassword = m_sInput

End Function
'OK button click event
Private Sub m_cmdConfirm_Click()
    m_sInput = Nz(m_frmPwInput.txtInput.Value, "")
    closeForm
End Sub

'cancel button click event
Private Sub m_cmdCancel_Click()
    m_sInput = ""
    closeForm
End Sub
'close form event
Private Sub m_frmPwInput_Unload(Cancel As Integer)
    Set m_frmPwInput = Nothing
End Sub

Private Sub closeForm()
    If Not m_frmPwInput Is Nothing Then
        m_frmPwInput.txtInput.Value = ""
        m_frmPwInput.Visible = False
        Set m_frmPwInput = Nothing
    End If
End Sub
