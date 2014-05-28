VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub Form_Load()
    
    Set p_DynChoise = New clsDynChoise
    Call p_DynChoise.init(Me, "tblDynList", "tblDynListTableNames", "tblDynListValues")
    p_DynChoise.load
End Sub
