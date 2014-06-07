Attribute VB_Name = "codingHelper"
'---------------------------------------------------------------------------------------
' Module    : codingHelper
' Author    : Dmitry Gorelenkov
' Date      : 30.01.2013
' Changed   : 22.12.2013
' Purpose   : VBA code erstellen, modifizieren
' Requires  : clsFuncs, clsMyCollection, Dmitry :)
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Explicit
Private f As New clsFuncs

Public Sub createCode()
'    Call createArrayCodeFromRange("myarray", 5)
    Dim aDefaults As Variant
    aDefaults = Array("Type", "Left", "Top", "width", "Height", "Name", "FontName", "FontSize", "FontWeight", "FontItalic", "ForeColor", "TextAlign")
    
    Call printControlForClsQuestionnaireQuestion(takeDesignOfControl(, aDefaults))
End Sub

'prints properties for clsQuestionnaireQuestion
Private Function printControlForClsQuestionnaireQuestion(mcPropsCol As clsMyCollection)
    Dim prop As Variant ''string

    Dim quot As String
    Dim valueToPrint As Variant
    
    For Each prop In mcPropsCol.getKeys
        If VarType(mcPropsCol(prop)) = vbString Then
            quot = """"
        Else
            quot = ""
        End If
        
        valueToPrint = mcPropsCol(prop)
        If VarType(valueToPrint) = vbBoolean Then
            valueToPrint = IIf(valueToPrint, "true", "false")
        End If
        
        
        Debug.Print "Call .add(" & quot & valueToPrint & quot & ", """ & prop & """)"
    Next prop
    
End Function

'return properties of all controls of sForm, as collection of collections
Private Function takeDesignOfForm(Optional sForm As String = "", Optional aOnlyThisProps As Variant) As clsMyCollection
    Dim mcMainCol As clsMyCollection
    Dim mcPropCol As clsMyCollection
    Dim ctrl As Control    ' Variant 'control

    If sForm = "" Then sForm = Screen.ActiveForm.name
    If aOnlyThisProps Is Nothing Then aOnlyThisProps = Array()
    Debug.Assert IsArray(aOnlyThisProps)

    Set mcMainCol = New clsMyCollection

    For Each ctrl In Forms(sForm).Controls
        Set mcPropCol = takeDesignOfControl(ctrl, aOnlyThisProps)
        Call mcMainCol.Add(mcPropCol, ctrl.name)
    Next ctrl

End Function

'returns collection with properties of control ctrl
Private Function takeDesignOfControl(Optional ctrl As Control, Optional aOnlyThisProps As Variant)
    Dim mcReturnCol As New clsMyCollection
    Dim prop As Variant
    Dim propValue As Variant
    On Error Resume Next
    
    If ctrl Is Nothing Then Set ctrl = Screen.ActiveControl
    If Not IsArray(aOnlyThisProps) Then aOnlyThisProps = Array()
    Debug.Assert IsArray(aOnlyThisProps)

    For Each prop In ctrl.Properties
        If f.isInArray(aOnlyThisProps, prop.name, True, True, True) Or UBound(aOnlyThisProps) = -1 Then
            propValue = ModifyByType(prop.value, prop.Type)
            Call mcReturnCol.Add(prop.value, prop.name)
        End If
    Next prop

    Set takeDesignOfControl = mcReturnCol
End Function
'modify value by type
Private Function ModifyByType(vValue As Variant, typeIdx As Integer) As Variant
    Select Case typeIdx
        Case vbBoolean
            ModifyByType = CBool(vValue)
        Case vbByte
            ModifyByType = CByte(vValue)
        Case vbCurrency
            ModifyByType = CCur(vValue)
        Case vbInteger
            ModifyByType = CInt(vValue)
        Case vbLong
            ModifyByType = CLng(vValue)
        Case vbDate
            ModifyByType = CDate(vValue)
        Case Else
            ModifyByType = CStr(vValue)
    End Select
End Function


'Private Sub createArrayCodeFromRange(Optional sToVar As String = "someArray", Optional widthInCode As Integer = 5, _
'                                     Optional ByRef rngRangewithValues As Excel.range, Optional bQuotation As Boolean = True)
'    If rngRangewithValues Is Nothing Then Set rngRangewithValues = Selection
'
'    Dim rngCell As range
'    Dim currValue As Variant
'    Dim counter As Long
'    Dim sResult As String
'    Dim sQuot As String
'
'    sQuot = IIf(bQuotation, """", "")
'    counter = 1
'
'    'falls 1 emelent
'    If rngRangewithValues.count <= 1 Then
'        Debug.Print "Array(" & sQuot & CStr(currValue) & sQuot & ")"
'        Exit Sub
'    End If
'
'
'    For Each rngCell In rngRangewithValues
'        'HIER FILTER z.B. if rngCell.Intersect.Color = 65554
'        currValue = CStr(rngCell.value)
'        sResult = sResult & sQuot & CStr(currValue) & sQuot & "," & IIf((counter Mod widthInCode) = 0, " _" & vbCrLf, "")
'        counter = counter + 1
'
'
'    Next rngCell
'
'    sResult = left(sResult, Len(sResult) - 4)    'remove last ", _ " & crlf, 4 symbols
'    sResult = sToVar & " = Array( _" & vbCrLf & sResult & " _" & vbCrLf & ")"
'    Debug.Print sResult
'End Sub

Sub rename_txtctrls_by_source()
    Dim selectedCtrl As Control
    Dim selControls As Controls
    
    Set selControls = Screen.ActiveForm.Controls
    For Each selectedCtrl In selControls
        If selectedCtrl.InSelection Then
'            Debug.Print selectedCtrl.ControlSource
           selectedCtrl.name = "txt" & Mid$(selectedCtrl.ControlSource, 4)
           
        End If
    
        
    Next selectedCtrl
End Sub

'TODO
Sub adjust_selected_controls_horiz()
    Const horAbstand = 100&
    Dim currentTop As Long
    
    Dim selectedCtrl As Control
    Dim selControls As Controls
    
    Set selControls = Screen.ActiveForm.Controls
    For Each selectedCtrl In selControls
        If selectedCtrl.InSelection Then
'            Debug.Print selectedCtrl.ControlSource
           selectedCtrl.name = "txt" & Mid$(selectedCtrl.ControlSource, 4)
           
        End If
    
        
    Next selectedCtrl
End Sub

