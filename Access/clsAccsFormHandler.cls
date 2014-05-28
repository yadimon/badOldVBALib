VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAccsFormHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsAccsFormHandler
' Author    : Dmitry Gorelenkov
' Date      : 08.02.2013
' Changed   : 09.04.2013
' Purpose   : Handle dynamic/static created form (create elements, change properties of them)
' Requires  : clsControlProperties, clsMyCollection, clsFuncs
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
Private f As clsFuncs
Private Const LONGMAX As Long = 2 ^ 31 - 1






'add controls at the end (down)
Public Function addControlsDown(frmForm As Access.Form, ByVal mcCtrlProps As clsMyCollection, _
                                    Optional lLeftAbstandTw As Long = 567, Optional lTopAbstandTw As Long = 567) As clsMyCollection
    
    Dim mcCtrlPropsMoved As clsMyCollection
    
    Set mcCtrlPropsMoved = moveAllPropsControls(mcCtrlProps, CLng(lLeftAbstandTw), CLng(lTopAbstandTw + findMostDownInForm(frmForm)))
    
    Set addControlsDown = addControls(frmForm, mcCtrlProps)
End Function

'add controls right
Public Function addControlsRight(frmForm As Access.Form, ByVal mcCtrlProps As clsMyCollection, _
                                    Optional lLeftAbstandTw As Long = 1, Optional lTopAbstandTw As Long = 1) As clsMyCollection
    
    Dim mcCtrlPropsMoved As clsMyCollection
    
    Set mcCtrlPropsMoved = moveAllPropsControls(mcCtrlProps, CLng(lLeftAbstandTw + findMostRightInForm(frmForm)), CLng(lTopAbstandTw))
    
    Set addControlsRight = addControls(frmForm, mcCtrlProps)
End Function

Public Function addControlsLeft()
    Debug.Print "addControlsLeft of Klassenmodul clsAccsFormHandler is not implemented yet"
End Function

Public Function addControlsTop()
    Debug.Print "addControlsTop of Klassenmodul clsAccsFormHandler is not implemented yet" 'TODO
End Function

'modify position by all controls-props in mcCtrlProps
Private Function moveAllPropsControls(ByVal mcCtrlProps As clsMyCollection, _
                                    Optional lLeftAbstandTw As Long = 567, Optional lTopAbstandTw As Long = 567) As clsMyCollection
    Dim prpProps As Variant 'clsControlProperties
    Dim mcCtrlPropsCopy As clsMyCollection
    
    'copy collection values
    Set mcCtrlPropsCopy = mcCtrlProps.Copy
    
    'modify top and left of every control
    For Each prpProps In mcCtrlPropsCopy.GetValues
        With prpProps.getCollection
            Call .setValue(lTopAbstandTw + .getValue("Top"), "Top")
            Call .setValue(lLeftAbstandTw + .getValue("Left"), "Left")
        End With
    Next prpProps
                                        
    Set moveAllPropsControls = mcCtrlPropsCopy
End Function


'noch ein controlsset hinzufuegen
'versetzte position Left -> lLeftOffset, Top-> lTopOffset
Public Function addControls(frmForm As Access.Form, mcCtrlProps As clsMyCollection) As clsMyCollection
    Dim tmpCtrl As Control
    Dim mcReturnControls As New clsMyCollection
    Dim ctlProps As clsControlProperties
    Dim sKey As Variant
    
    For Each sKey In mcCtrlProps.getKeys()
        Set ctlProps = mcCtrlProps.getValue(sKey)
        Set tmpCtrl = addControl(frmForm, ctlProps)
        Call mcReturnControls.Add(tmpCtrl, tmpCtrl.Name)
        Set tmpCtrl = Nothing
    Next sKey
    
    'wegen OptionsGroups noch mal setProperties. TODO besser realisieren?
    For Each sKey In mcCtrlProps.getKeys()
         Set ctlProps = mcCtrlProps.getValue(sKey)
         If ctlProps.ControlType = acOptionGroup Then
            Call setPropertiesToControl(frmForm.Controls(ctlProps.ControlName), ctlProps)
         End If
    Next sKey
    
    Set addControls = mcReturnControls
End Function

'add single control element
Public Function addControl(frmForm As Access.Form, ctlProps As clsControlProperties) As Control
    Dim tmpCtrl As Control
    Dim sParent As String
    'falls control schon vorhanden
'    If Not IsNull(frmForm.Controls(ctlProps.ControlName)) Then
    If fAccs.isControlExists(frmForm.Name, ctlProps.ControlName) Then
'    If fAccs.tmpDel And False Then
'        Dim bA As Boolean
'        bA = fAccs.isFormLoaded(frmForm.name)
        Set tmpCtrl = frmForm.Controls(ctlProps.ControlName)
    Else
        If ctlProps.getCollection.isInArray("Parent") Then
            sParent = ctlProps.getCollection("Parent")
            Set tmpCtrl = CreateControl(frmForm.Name, ctlProps.getCollection("Type"), , sParent, , _
                                                                Nz(ctlProps.getCollection("Left"), 0), _
                                                                Nz(ctlProps.getCollection("Top"), 0), _
                                                                Nz(ctlProps.getCollection("Width"), 0), _
                                                                Nz(ctlProps.getCollection("Height"), 0))
        Else
        'make missing? howto?
            Set tmpCtrl = CreateControl(frmForm.Name, ctlProps.getCollection("Type"), , , , _
                                                                Nz(ctlProps.getCollection("Left"), 0), _
                                                                Nz(ctlProps.getCollection("Top"), 0), _
                                                                Nz(ctlProps.getCollection("Width"), 0), _
                                                                Nz(ctlProps.getCollection("Height"), 0))
        End If
    End If
    
                                
    Call setPropertiesToControl(tmpCtrl, ctlProps)
    
    Set addControl = tmpCtrl
End Function

'try to set properties to control by clsControlProperties
Public Function setPropertiesToControl(ctrl As Control, ctlProps As clsControlProperties)
    On Error Resume Next
    Dim sPropName As Variant 'String
    
    For Each sPropName In ctlProps.getCollection.getKeys
        ctrl.Properties.item(sPropName) = ctlProps.getCollection.getValue(sPropName)
    Next sPropName
        
    
    err.Clear
End Function

'letzes controlelement, das unterste punkt
Public Function findMostDownInForm(frmForm As Access.Form) As Long
    Dim ctrl As Variant
    Dim lastTop As Long
    Dim mcTmpCol As New clsMyCollection
    
    For Each ctrl In frmForm.Controls
        Call mcTmpCol.Add(ctrl)
    Next ctrl
    
    lastTop = findMostDownOfCtrls(mcTmpCol)
    
    findMostDownInForm = lastTop
End Function

'letzes (nach rechts) controlelement, das rechteste punkt
Public Function findMostRightInForm(frmForm As Access.Form) As Long
    Dim ctrl As Variant
    Dim lastLeft As Long
    Dim mcTmpCol As New clsMyCollection
    
    For Each ctrl In frmForm.Controls
        Call mcTmpCol.Add(ctrl)
    Next ctrl
    
    lastLeft = findMostRightOfCtrls(mcTmpCol)
    
    findMostRightInForm = lastLeft
End Function

'letzes (nach links) controlelement, das linkeste punkt in form
Public Function findMostLeftInForm(frmForm As Access.Form) As Long
    Dim ctrl As Variant
    Dim lastLeft As Long
    Dim mcTmpCol As New clsMyCollection
    
    For Each ctrl In frmForm.Controls
        Call mcTmpCol.Add(ctrl)
    Next ctrl
    
    lastLeft = findMostLeftOfCtrls(mcTmpCol)
    
    findMostLeftInForm = lastLeft
End Function

'letzes (nach rechts) controlelement, das rechteste punkt
Public Function findMostRightOfCtrls(mcControls As clsMyCollection) As Long
    Dim ctrl As Variant
    Dim lastLeft As Long
    On Error Resume Next
    
    For Each ctrl In mcControls.GetValues
        lastLeft = f.max(lastLeft, ctrl.Left + ctrl.Width)
    Next ctrl
    
    
    err.Clear
    
    findMostRightOfCtrls = lastLeft
End Function

'letzes (nach links) controlelement, das linkeste punkt
Public Function findMostLeftOfCtrls(mcControls As clsMyCollection) As Long
    Dim ctrl As Variant
    Dim lastLeft As Long
    lastLeft = LONGMAX
    On Error Resume Next
    
    For Each ctrl In mcControls.GetValues
        lastLeft = f.min(lastLeft, ctrl.Left)
    Next ctrl
    
    
    err.Clear
    
    findMostLeftOfCtrls = lastLeft
End Function

'letzes (nach unten) controlelement, das unterste punkt
Public Function findMostDownOfCtrls(mcControls As clsMyCollection) As Long
    Dim ctrl As Variant
    Dim lastDown As Long
    On Error Resume Next
    
    For Each ctrl In mcControls.GetValues
        lastDown = f.max(lastDown, ctrl.Top + ctrl.Height)
    Next ctrl
    
    
    err.Clear
    
    findMostDownOfCtrls = lastDown
End Function

'letzes (nach oben) controlelement, das oberste punkt
Public Function findMostTopOfCtrls(mcControls As clsMyCollection) As Long
    Dim ctrl As Variant
    Dim lastTop As Long
    lastTop = LONGMAX
    On Error Resume Next
    
    For Each ctrl In mcControls.GetValues
        lastTop = f.min(lastTop, ctrl.Top)
    Next ctrl
    
    
    err.Clear
    
    findMostTopOfCtrls = lastTop
End Function

Private Sub Class_Initialize()
    Set f = New clsFuncs
    Set fAccs = New clsFuncsAccs
End Sub

Private Sub Class_Terminate()
    Set f = Nothing
    Set fAccs = Nothing
End Sub



