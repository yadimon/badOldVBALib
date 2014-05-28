VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'*TODO make all the handlers by CLASS

Private Sub btnCancel_Click()
    loadParams
    If Not (p_bDevelopmentMode) Then
        Me.Application.Quit
    End If
End Sub

Private Sub Form_LostFocus()
'    Me.SetFocus
'    Call f.SetForegroundWindow(Me.hWnd)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    loadParams
    'todo..
    If fAccs.isFormLoaded("frmKontakteMain") Then
        Cancel = True
    End If
End Sub
