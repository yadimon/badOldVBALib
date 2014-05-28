VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettingLayoutColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsSettingLayoutColor
' Author    : Dmitry Gorelenkov
' Date      : 27.01.2013
' Purpose   : color setting Option calss
' Requires  : clsColorPicker, clsSettingsHandler
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit


Private m_bInitDone As Boolean
Private m_clrpickDlg As clsColorPicker
Private m_sName As String
Private m_lDefaultValue As Long
Private m_lCurrentChosenColor As Long
Private WithEvents m_ctrFiled As Access.TextBox 'field with color code
Attribute m_ctrFiled.VB_VarHelpID = -1
Private WithEvents m_shpColorPickShape  As Access.Rectangle 'pick rectangle
Attribute m_shpColorPickShape.VB_VarHelpID = -1
Private m_shSettingsHandler As ISettingsHandler

Private Sub Class_Terminate()
    Set m_clrpickDlg = Nothing
End Sub

'felder bei setting form fuellen
Public Function fillFormFields(Optional lColorNr As Long = -1) As Variant
    On Error GoTo fillFormFields_Error

    If Not m_bInitDone Then Exit Function
    If lColorNr = -1 Then lColorNr = getValue
    
    m_ctrFiled.Value = lColorNr
    m_shpColorPickShape.BackColor = lColorNr
    m_lCurrentChosenColor = lColorNr
    
    Exit Function

fillFormFields_Error:
    Debug.Print Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description
    
End Function

Public Function getValue() As Long
    If Not m_bInitDone Then Exit Function
    Dim sRetVal As String
    
    sRetVal = m_shSettingsHandler.getSetting(m_sName)
    If isOkForValue(sRetVal) Then
        sRetVal = CLng(sRetVal)
    Else
        sRetVal = getDefaultColor
    End If
    
    getValue = sRetVal
End Function

Public Function init(sName_p As String, lDefaultValue_p As Long, ByRef ctrFiled_p As Access.TextBox, _
                    ByRef shpColorPickShape_p As Access.Rectangle, ByRef shSettingsHandler_p As ISettingsHandler) As Boolean
    m_sName = sName_p
    m_lDefaultValue = lDefaultValue_p
    Set m_ctrFiled = ctrFiled_p
    Set m_shpColorPickShape = shpColorPickShape_p
    Set m_clrpickDlg = New clsColorPicker
    m_ctrFiled.OnChange = "[Event Procedure]"
    m_shpColorPickShape.OnClick = "[Event Procedure]"
    
    Set m_shSettingsHandler = shSettingsHandler_p
    
    m_bInitDone = True
End Function

Public Property Get Name() As String
    If Not m_bInitDone Then Exit Property
    
    Name = m_sName
End Property

'set color value
Public Function setValue(Optional ByVal vValue As Variant = Null) As String
    If Not m_bInitDone Then Exit Function
    
    If IsNull(vValue) Then
        vValue = m_lCurrentChosenColor
    End If
    
    If vValue = "" Then
        vValue = getDefaultColor
    End If
    
    If isOkForValue(CStr(vValue)) Then
        setValue = m_shSettingsHandler.setSetting(m_sName, CStr(vValue))
    Else
        Debug.Print "setValue: falsche value to set (" & vValue & ")"
    End If
    
End Function

'on change by field
Private Sub m_ctrFiled_Change()
    Dim sText As String
    On Error GoTo m_ctrFiled_Change_Error

    sText = m_ctrFiled.Text
    
    'fall leeres string, default laden
    If sText = vbNullString Then
        Call fillFormFields(getDefaultColor)
    End If
    
    If isOkForValue(sText) Then
        fillFormFields (sText)
    End If
    
    m_ctrFiled.SelStart = Len(sText)


    Exit Sub

m_ctrFiled_Change_Error:
    If Err.Number = 6 Then 'ueberlauf..
        Exit Sub
    Else
        Debug.Print "m_ctrFiled_Change: " & Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description
    End If
    
End Sub

'on click on rectangle
Private Sub m_shpColorPickShape_Click()

    fillFormFields (m_clrpickDlg.showColorPick())
    
End Sub


'default color fuer die setting zuruckgeben
Public Function getDefaultColor() As Long
    If Not m_bInitDone Then Exit Function
    getDefaultColor = m_lDefaultValue
End Function

'prueft ob wert sValue als geegnete value gespeichert werden kann
Private Function isOkForValue(sValue As String) As Boolean
    Dim testVar As Long
    Dim bRet As Boolean
    bRet = False
    
    If Not IsNull(sValue) And sValue <> vbNullString And IsNumeric(sValue) Then 'no need? :/
        On Error GoTo wrong
        testVar = CLng(sValue)
        bRet = True
    End If
wrong:
    isOkForValue = bRet
End Function

'test myself? todo?
Public Sub test()
    If Not m_bInitDone Then Exit Sub
    
End Sub
