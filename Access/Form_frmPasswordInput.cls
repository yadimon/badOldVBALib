VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPasswordInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnCancel_Click()
'p_sFormPwInput = ""
    Me.txtInput.value = ""
    Me.Visible = False
    
End Sub

Private Sub btnOK_Click()
'p_sFormPwInput = Nz(Me.txtInput.Value, "")
    Me.Visible = False
End Sub


Private Sub Form_Current()
    Me.txtInput.SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible Then
        Me.txtInput.value = ""
        Me.Visible = False
        Cancel = True
    End If
End Sub
