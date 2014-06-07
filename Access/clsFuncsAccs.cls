VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuncsAccs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFuncsAccs
' Author    : Dmitry Gorelenkov
' Date      : 08.2012
' Changed   : 06.06.2014
' Purpose   : Funktionssammlung - Access
' Requires  : clsFuncs, clsErrorHandler, clsMyCollection
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------




Option Compare Database
Option Explicit
Private ehError As clsErrorHandler
Private f As clsFuncs
Private Const LONGMAX As Long = 2 ^ 31 - 1

Private Sub Class_Initialize()
    Set ehError = New clsErrorHandler
    Set f = New clsFuncs
End Sub

Private Sub Class_Terminate()
    Set ehError = Nothing
    Set f = Nothing
End Sub

'sucht nach einem Tag, der mit komma getrennt in "Marke" drin ist
Public Function getCtrlsByTag(ByRef frmForm As Form, ByVal sTag As String) As clsMyCollection
    On Error GoTo ErrorHandler
    
    If frmForm Is Nothing Then Exit Function
    
    Dim ctrControl As Object
    Dim aArrTags As Variant
    Dim retArrayCol As New clsMyCollection
    
    For Each ctrControl In frmForm.Controls
        aArrTags = Split(ctrControl.Tag, ",", , vbTextCompare)
        If f.isInArray(aArrTags, sTag, True, True, True) Or sTag = vbNullString Then
            Call retArrayCol.Add(ctrControl, ctrControl.name)
        End If
        
    Next ctrControl
    
    Set getCtrlsByTag = retArrayCol
    
    Exit Function
ErrorHandler:
    ehError.catchErr ("getCtrlsByTag")
End Function

'suchtein control, der den source sSource hat
Public Function getCtrlBySource(ByRef frmForm As Form, ByVal sSource As String) As Access.Control
    On Error GoTo ErrorHandler
    
    If frmForm Is Nothing Then Exit Function
    
    Dim ctrControl As Object ' control
    Dim prop As Property
    
    For Each ctrControl In frmForm.Controls
        For Each prop In ctrControl.Properties
            If prop.name = "ControlSource" Then
                If StrComp(prop.value, sSource, vbTextCompare) = 0 Then
                    'control gefunden
                    Set getCtrlBySource = ctrControl
                    Exit Function
                End If
            End If
        Next prop
    Next ctrControl
    
    Set getCtrlBySource = Nothing
    
    Exit Function
ErrorHandler:
    ehError.catchErr ("getCtrlBySource")
End Function

'add tooltip mit steuerelement-inhalten
Sub addToolTips(frmAktForm As Form)
    Dim ctlInForm As Control
    On Error Resume Next
    
    For Each ctlInForm In frmAktForm.Controls
      With ctlInForm
        If .ControlType = acTextBox Then
          .ControlTipText = IIf(IsNull(.value), "", .value)
        ElseIf .ControlType = acComboBox Or _
          .ControlType = acListBox Then
          If .ColumnCount > 1 Then
            .ControlTipText = IIf(IsNull(.value), "", .Column(.BoundColumn))
          Else
            .ControlTipText = IIf(IsNull(.value), "", .value)
          End If
        End If
      End With
    Next ctlInForm

End Sub

'sets properties back to default
'frmForm are controls (but a Form will be as argument)
Public Function setDefaultFormValues(ByRef frmForm As Variant)
    Dim ctl As Control
    On Error Resume Next
    For Each ctl In frmForm
        Select Case ctl.ControlType
            Case acComboBox, acListBox, acOptionGroup, acTextBox, acCheckBox    ' add other types if necessary
                If Not ctl.Locked Then
                    ctl.value = quotTrim(ctl.DefaultValue)
                End If
        End Select
    Next
End Function

'removes quotation at beginn and end of the string, if exists
Private Function quotTrim(ByVal vValue As Variant)

    If VarType(vValue) = vbString Then
        If Left(vValue, 1) = """" Then
            vValue = Right(vValue, Len(vValue) - 1)
        End If
        
        If Right(vValue, 1) = """" Then
            vValue = Left(vValue, Len(vValue) - 1)
        End If
    End If
    
    quotTrim = vValue
    
End Function


Public Function modifyAllPossible(ByRef frmForm As Variant, sWhatToModify As String, Optional bWhatValue As Boolean = True)
    Dim ctl As Variant
    On Error Resume Next
    
    For Each ctl In frmForm
        Select Case LCase(sWhatToModify)
            Case "value"
                ctl.value = ""

            Case "controlsource"
                ctl.ControlSource = ""
           
            Case "locked"
                ctl.Locked = bWhatValue
            
            Case "enabled"
                ctl.Enabled = bWhatValue
                
            Case "visible"
                ctl.Visible = bWhatValue
        
        End Select

        
    Next
End Function

'select 1 listeneintrag mit index
Public Function setListIndex(ByRef ctlList As Variant, index As Integer) As Boolean
    setListIndex = False
    
    If TypeName(ctlList) <> "ComboBox" And TypeName(ctlList) <> "ListBox" Then
        Debug.Assert False
        Exit Function
    End If
    
    With ctlList
        'wenn genug eintraege fuer den gesetzten wert da sind
        If .ListCount > index Then
            'falls multiselect bei listbox
            If TypeName(ctlList) = "ListBox" Then
                If .MultiSelect = 1 Then
                    .Selected(index) = True
                    setListIndex = True
                    Exit Function
                End If
            End If
            
            'sonst mit ItemData versuchen
            .value = .itemData(index)
            setListIndex = True
        End If
    End With
End Function

'list eintrag markieren bei Wert sValue, in spalte iColumn
Public Function selectListByValue(ByRef lstList As Access.ListBox, sValue As String, Optional iColumn As Integer = -1) As Boolean
    Dim i As Integer
    Dim tmpValue As Variant
    selectListByValue = False
    
    'fuer jeden eintrag
    For i = 0 To lstList.ListCount
        'falls keine Spalte eingegeben, gebundete Spalte lesen
        If iColumn = -1 Then
            tmpValue = lstList.itemData(i)
        ElseIf iColumn < lstList.ColumnCount - 1 Then
            tmpValue = lstList.Column(iColumn, i)
        End If
        'falls uebereinstimmt, markieren
        If StrComp(CStr(tmpValue), sValue, vbTextCompare) = 0 Then
            selectListByValue = setListIndex(lstList, i)
            Exit For
        End If
    Next i
End Function

'mehrere zeilen in ListBox selektieren, bei Werten in aValues array
Public Function selectListByValues(ByRef lstList As Access.ListBox, aValues As Variant, Optional iColumn As Integer = -1)
    Dim bReturnResult As Boolean
    bReturnResult = False
    If Not IsArray(aValues) Then
        Debug.Assert False
        Exit Function
    End If
    
    If lstList.MultiSelect <> 1 Then
        Debug.Assert False
        Exit Function
    End If
    
    Dim vValue As Variant
    bReturnResult = True
    For Each vValue In aValues
        bReturnResult = bReturnResult & selectListByValue(lstList, CStr(vValue), iColumn)
    Next vValue
    
    selectListByValues = bReturnResult
End Function

'deselect all items in ListBox.
'http://allenbrowne.com/func-12.html  modified
Function unselectListBox(lst As Access.ListBox) As Boolean
On Error GoTo Err_ClearList
    'Purpose:   Unselect all items in the listbox.
    'Return:    True if successful
    'Author:    Allen Browne. http://allenbrowne.com  June, 2006.
    Dim varItem As Variant

    If lst.MultiSelect = 0 Then
        lst = Null
    Else
        For Each varItem In lst.ItemsSelected
            lst.Selected(varItem) = False
        Next
    End If

    unselectListBox = True

    Exit Function

Err_ClearList:
    Debug.Print "error by clsFuncsAccs.unselectListBox"
End Function

'try to scale formular
Public Function resizeFormular(ByRef frmForm As Variant, dFaktor As Double)
    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In frmForm
        If Not ctrl Is Nothing Then
            ctrl.Height = ctrl.Height * dFaktor
            ctrl.Width = ctrl.Width * dFaktor
            ctrl.Top = ctrl.Top * dFaktor
            ctrl.Left = ctrl.Left * dFaktor
            ctrl.FontSize = ctrl.FontSize * dFaktor
            
            'for subform add recursive scale
            If ctrl.ControlType = acSubform Then
                Call resizeFormular(ctrl, dFaktor)
            End If
        End If
    Next ctrl

End Function

'datenbank pfad fuer alle (bzw mit bestimmten prefix) verknuepfte tabellen aendern
Public Function changeSourceDB(sPath As String, Optional sTblPrefix As String = vbNullString) As Integer
    Dim DB As Database
    Dim sPrefLen As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set DB = CurrentDb()
    changeSourceDB = 0
    
    
    
    sPrefLen = Len(sTblPrefix)
    
    For i = 0 To DB.TableDefs.Count - 1
        If DB.TableDefs(i).Attributes = dbAttachedTable Then '.connect <> ""
            If Left(DB.TableDefs(i).name, sPrefLen) = sTblPrefix Then
                If Mid(DB.TableDefs(i).connect, 11) <> sPath Then
                    DB.TableDefs(i).connect = ";database=" & sPath
                    DB.TableDefs(i).RefreshLink
                    changeSourceDB = 1 'set changed
                End If
            End If
        End If
    Next i
    
    Exit Function
ErrorHandler:
    ehError.catchErr ("changeSourceDB")
    changeSourceDB = -1
End Function

'select datenbank dialog
'sDefaultPath vordefiniertes Pfad
Public Function selectDBDialog(sPrefTbl As String, Optional sDefaultPath As String, Optional bSuccessMessage As Boolean = True) As String
    Dim sDBPath As String
    
    If Not f.fileExists(sDefaultPath) Then
        sDefaultPath = ""
    End If
    
    sDBPath = f.DialogGetPath(True, "BackEnd Datenbank Pfad", "Access DB Files", "*.accd*", sDefaultPath, , "Auswählen")
    If Not f.hatWert(sDBPath) Then Exit Function
    
    If changeSourceDB(sDBPath, sPrefTbl) <> -1 Then
        If bSuccessMessage Then
            MsgBox "Erfolgreich geändert"
        End If
    Else
        MsgBox "Fehler aufgetreten"
    End If
    
    selectDBDialog = sDBPath
End Function

'get DBpath of all attached tables
Public Function getBackEndPath(Optional dbDatabase As Database) As String
    Dim DB As Database
    Dim i As Integer
    Dim sFoundPath As String


    If Not DEBUGMODE Then On Error GoTo getBackEndPath_Error

    If dbDatabase Is Nothing Then Set DB = CurrentDb()

    For i = 0 To DB.TableDefs.Count - 1
        If DB.TableDefs(i).Attributes = dbAttachedTable Then
            If sFoundPath = "" Then
                sFoundPath = Mid(DB.TableDefs(i).connect, 11)
            End If
            'falls andere path gefunden, error, return 0
            If Mid(DB.TableDefs(i).connect, 11) <> sFoundPath Then
                getBackEndPath = "0"
            End If
        End If
    Next i

    getBackEndPath = sFoundPath
    Exit Function



    Exit Function

getBackEndPath_Error:
    ehError.catchErr ("getBackEndPath of Klassenmodul clsFuncsAccs")
    

End Function

'prueft alle verknuepfte mit BackEnd Tabellen, ob die BackEnd DB zugreifbar ist
'sehr langsam!!!
Public Function checkDBTblLinks(Optional dbDatabase As Database) As Boolean
    Dim DB As Database
    Dim i As Integer
    
    On Error GoTo checkTblLinks_Error

    If dbDatabase Is Nothing Then Set DB = CurrentDb()
    
    For i = 0 To DB.TableDefs.Count - 1
        If DB.TableDefs(i).Attributes = dbAttachedTable Then
            DB.TableDefs(i).RefreshLink
        End If
    Next i
    

    checkDBTblLinks = True
    
    Exit Function

checkTblLinks_Error:
    If err.Number = 3024 Then 'cannot refreshlink
        checkDBTblLinks = False
    Else
        ehError.catchErr ("checkTblLinks of Klassenmodul clsFuncsAccs")
    End If
End Function

'prueft ob eine tabelle existiert
'from http://www.office-loesung.de/ftopic1371_0_0_asc.php
Public Function tableExists(mytable As String) As Boolean
    On Error GoTo Fehlerbehandlung
    Dim td As DAO.TableDef
    Dim DB As Database
    Set DB = CurrentDb
    
    Set td = DB.TableDefs(mytable)
    tableExists = True
    Exit Function
Fehlerbehandlung:
    tableExists = False
End Function

'prueft ob in tabelle sTable sColName spalte existiert
'http://www.office-loesung.de/ftopic24463_0_0_asc.php  - modified
Public Function columnExists(sTable As String, sColName As String)
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim DB As Database
    Set DB = CurrentDb
    
    
    'Fehlerbehandlung abschalten
    On Error Resume Next
    
    Set tdf = DB.TableDefs(sTable)

    
    columnExists = True
    
    Set fld = tdf.Fields(sColName)
    
    If err <> 0 Then
        columnExists = False
    End If
    
    On Error GoTo 0 'no need?
End Function


'Recordset nach Excel exportieren
'http://www.ms-office-forum.net/forum/showthread.php?t=242723  - modified
Public Function Recordset2Excel(ByRef rstData As Object, _
                        Optional ByVal StartRowNr As Long = 1, _
                        Optional ByVal StartColNr As Long = 1, _
                        Optional ByVal bWithRecordsetHeaders As Boolean = True)
                        
    'rstData = ADODB oder DAO-Recordset-Referenz
    
    'tests
    If StartRowNr < 1 Or StartColNr < 1 Or (LCase(TypeName(rstData)) <> LCase("Recordset") And LCase(TypeName(rstData)) <> LCase("Recordset2")) Then
        f.dbg "Recordset2Excel: falsche parameter"
        Exit Function
    End If
    
    
    Dim lngRstDataCols As Long
    Dim strRstDataColsName() As String
    Dim i As Long

    Dim xlApp As Object   '.. = Excel.Application
    Dim xlBook As Object  '.. = Excel.Workbook
    Dim xlSheet As Object  '.. = Excel.WorkSheet

   On Error GoTo Recordset2Excel_Error

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    xlApp.Visible = True
    Set xlSheet = xlBook.Worksheets(1)
    
    'Überschriften
    If bWithRecordsetHeaders Then
        With rstData
            lngRstDataCols = .Fields.Count
            ReDim strRstDataColsName(lngRstDataCols - 1)
            For i = 0 To (lngRstDataCols - 1)
                strRstDataColsName(i) = .Fields(i).name
            Next i
            xlSheet.range(xlSheet.Cells(StartRowNr, StartColNr), _
                          xlSheet.Cells(StartRowNr, StartColNr + lngRstDataCols - 1)).value = strRstDataColsName
            StartRowNr = StartRowNr + 1
        End With
    End If

    'Daten
    xlSheet.Cells(StartRowNr, StartColNr).CopyFromRecordset rstData
    
    
   On Error GoTo 0
   Exit Function

Recordset2Excel_Error:
    ehError.catchErr ("Recordset2Excel of Klassenmodul clsFuncsAccs")
    

End Function

'moved to clsFuncs
'simple add WHERE sClause clause, todo: if ORDER BY/Gro ...  is in string
Public Function addWhereClause(ByVal sSQLString As String, sWhatToAdd As String, Optional sOperator As String = "AND")
    On Error GoTo addWhereClause_Error
    
    addWhereClause = f.SQLString_addWhereClause(sSQLString, sWhatToAdd, sOperator)

    Exit Function

addWhereClause_Error:
    ehError.catchErr ("addWhereClause of Klassenmodul clsFuncsAccs")
    
End Function


'aktuelle datensatz nach bestimmte bedingungen umstellen/suchen
Sub formMoveToRecord(ByRef frmForm As Access.Form, sFindFirst As String)
    Dim rsRS As Recordset
    Set rsRS = frmForm.RecordsetClone
    rsRS.FindFirst (sFindFirst)
    If Not rsRS.EOF Then frmForm.Bookmark = rsRS.Bookmark
    Set rsRS = Nothing
End Sub


'refresh form, and try to set pointer to last position
Function formReload(sFormName As String, Optional sByWhatField As String = "0", Optional bStayByIndex = False)
    Dim sFieldData As String
    
    On Error GoTo formReload_Error
    formReload = False
    
    With Forms(sFormName)
        'falls numerisch, versuchen spaltenname zu bekommen
        If IsNumeric(sByWhatField) Then
            sByWhatField = .Recordset.Fields(Val(sByWhatField)).name
        End If
        
        'key spalte speichern
        On Error Resume Next 'falls gelöscht oder sonst irgendwas, ignorieren
        sFieldData = f.ifNoValue(.Recordset.Fields(sByWhatField).value, "")
        .Requery
        'zu id springen, falls moeglich
        Call formMoveToRecord(Forms(sFormName), sByWhatField & " = " & sFieldData)
        .Recalc
        On Error GoTo formReload_Error
    End With

    formReload = True
    Exit Function

formReload_Error:
    ehError.catchErr ("formReload of Modul clFuncsAccs")
    
End Function

'prueft ob form sFormName geladen ist
Function isFormLoaded(ByVal strFormName As String) As Boolean
    isFormLoaded = CurrentProject.AllForms(strFormName).IsLoaded
'    On Error Resume Next
'    Dim frm As Access.Form
'
'    Set frm = Forms(strFormName)
'
'    isFormLoaded = Not frm Is Nothing
'    err.Clear


''http://www.aadconsulting.com/aadtips.html
'    Const conObjStateClosed = 0
'    Const conDesignView = 0
'
'    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) > conObjStateClosed Then
'        If Forms(strFormName).CurrentView > conDesignView Then
'            isFormLoaded = True
'        End If
'    End If
End Function

    'prueft ob form sFormName existiert
Function formExists(sFormName As String) As Boolean
    On Error Resume Next
    Dim frmDummy As Variant
    If sFormName = vbNullString Then Exit Function
    
    Set frmDummy = Application.CurrentProject.AllForms(sFormName)
    formExists = (err.Number = 0)
    err.Clear
End Function

'try to get control reference from form
Public Function tryGetControl(sFrmForm_p As String, sCtrlName As String, Optional inFailCase As Variant = Nothing) As Control
'    On Error Resume Next
'    Dim retCtrl As Variant
'    Set retCtrl = Access.Forms(sFrmForm_p).Controls(sCtrlName)
'    If err.Number <> 0 Then
'        Set retCtrl = inFailCase
'    End If
'
'    err.Clear
    
    Dim ctrl As Control
    Dim bFound As Boolean
    Dim retCtrl As Variant
    
    For Each ctrl In Access.Forms(sFrmForm_p).Controls
        If ctrl.name = sCtrlName Then
            Set retCtrl = ctrl
            bFound = True
            Exit For
        End If
    Next ctrl
    
    If Not bFound Then
        Set retCtrl = inFailCase
    End If
    
    Set tryGetControl = retCtrl
End Function

'set value vValue to property sProp_p, for every Control in Controls Set (clsMyCollection).
Public Function setPropertyForControls(mcControlsSet As clsMyCollection, sProp_p As String, vValue As Variant)
    Dim ctrl As Variant 'control
    On Error Resume Next
    For Each ctrl In mcControlsSet.GetValues
        ctrl.Properties(sProp_p) = vValue
    Next ctrl
    
    err.Clear
End Function

'check if control exists on a form
Public Function isControlExists(frmForm_p As String, sCtrlName As String) As Boolean
    isControlExists = Not (tryGetControl(frmForm_p, sCtrlName, Nothing) Is Nothing)
End Function

'moves controls in collection as Block of controls to coord left, top.
Public Function moveControls(mcControls As clsMyCollection, lLeftTw As Long, lTopTw As Long)
    Dim ctrl As Variant
    Dim lastLeft As Long
    Dim lastTop As Long
    Dim diffLeft As Long
    Dim diffTop As Long
    lastLeft = LONGMAX
    lastTop = LONGMAX
    
    On Error Resume Next
    
    For Each ctrl In mcControls.GetValues
        lastTop = f.min(lastTop, ctrl.Top)
        lastLeft = f.min(lastLeft, ctrl.Left)
    Next ctrl
    
    'difference between controlsSet most left-top point and new point to move to
    diffLeft = lLeftTw - lastLeft
    diffTop = lTopTw - lastTop
    
    'offset each control
    Call offsetControls(mcControls, diffLeft, diffTop)
    
    'offset each control
'    For Each ctrl In mcControls.GetValues
'        ctrl.top = ctrl.top + diffTop
'        ctrl.left = ctrl.left + diffLeft
'    Next ctrl
    
    If err.Number <> 0 Then
        Debug.Print "Fehler bei moveControls " & vbCrLf & _
        "Err.num:" & err.Number & " Err.Description: " & err.Description
    End If
    
    err.Clear
End Function

'change all controls properties LEFT and TOP to LEFT+lLeftTw, TOP+lTopTw
Public Function offsetControls(mcControls As clsMyCollection, lLeftTw As Long, lTopTw As Long)
    Dim ctrl As Variant
    On Error Resume Next
    
    Dim arrayOfParents As Variant
    arrayOfParents = Array()
    Dim aCtrlAndSize As Variant
    
    'OptionGroup Fix  'TODO all possible parent controls?
    'alle elemente durchsuchen, OptionGroup suchen
    For Each ctrl In mcControls.GetValues
        If ctrl.ControlType = acOptionGroup Then
            'original Width und Height parameter speichern
            Call f.arrayPush(arrayOfParents, Array(ctrl, ctrl.Width, ctrl.Height, ctrl.Left, ctrl.Top))
        End If
    Next ctrl

    
    'offset each control
    For Each ctrl In mcControls.GetValues
            ctrl.Top = ctrl.Top + lTopTw
            ctrl.Left = ctrl.Left + lLeftTw
    Next ctrl
    
    
    'fuer optionGroup versuchen groesse wiederherstellen
    For Each aCtrlAndSize In arrayOfParents
        'verschen richtige position und groesse nun zuweisen
        aCtrlAndSize(0).Left = aCtrlAndSize(3) + lLeftTw
        aCtrlAndSize(0).Width = aCtrlAndSize(1)
        
        aCtrlAndSize(0).Top = aCtrlAndSize(4) + lTopTw
        aCtrlAndSize(0).Height = aCtrlAndSize(2)
    Next aCtrlAndSize
    
    
    
    If err.Number <> 0 Then
        Debug.Print "Fehler bei offsetControls " & vbCrLf & _
        "Err.num:" & err.Number & " Err.Description: " & err.Description
    End If
    
    err.Clear
End Function

'alle empty strings in einer Tabelle in NULL wert umwandeln
'todo, genuere spalten eingabe
Function modifyTableEmptyStringToNull(sTableName As String)
    Dim rs As DAO.Recordset
    Dim cDB As Database
    
    On Error GoTo modifyTableStringToNull_Error
    
    Set cDB = CurrentDb()
    Set rs = cDB.OpenRecordset(sTableName)
    Dim intAnzahlfelder As Integer
    Dim stsql As String
    For intAnzahlfelder = 0 To rs.Fields.Count - 1
        'Check ob Feld Text oder Memo ist
        Select Case rs.Fields(intAnzahlfelder).Type
        'Text oder Memo
        Case dbMemo, dbText

            'SQL String für aktuelle Tabelle und aktuelles Feld erstellen
            stsql = "UPDATE " & sTableName & " SET " & sTableName & ".[" & _
                    rs.Fields(intAnzahlfelder).name & "] = NULL WHERE " & sTableName & ".[" & _
                    rs.Fields(intAnzahlfelder).name & "] ="""";"
            'Warnmeldungen AUs, wenn nicht abgeschaltet
            DoCmd.SetWarnings False
            
            DoCmd.RunSQL stsql
            
            'Hier einfach Warnmeldungen immer einschalten
            DoCmd.SetWarnings True
            
        End Select
    Next intAnzahlfelder
    
    rs.Close
    
    modifyTableEmptyStringToNull = True
    On Error GoTo 0
    Exit Function

modifyTableStringToNull_Error:

    ehError.catchErr ("modifyTableStringToNull")
    modifyTableEmptyStringToNull = False
End Function

'test if current recordset is locked
Public Function isCurrentRecordLocked(frm As Form, Optional ByRef sError As String)
    ' Display a message box that says either:
    '  -No user has the current record locked, or
    '  -The user & machine name of the user who
    '   who has locked the current record.
    
    Dim rst As DAO.Recordset
    
    On Error GoTo HandleErr
    
    ' Default message
    sError = "Record is not locked by another user."
    
    ' Clone the form's recordset and synch up to the
    ' form's current record
    Set rst = frm.RecordsetClone
    rst.Bookmark = frm.Bookmark
    
    ' If the current record is locked, then the next
    ' statement should produce an error that we will trap
    rst.Edit
    isCurrentRecordLocked = False
    
ExitHere:
    If Not rst Is Nothing Then rst.Close
    Exit Function
    
HandleErr:
    isCurrentRecordLocked = True
    ' Pass the error to acbGetUserAndMachine which will attempt
    ' to parse out the user and machine from the error message
    If err.Number = 3188 Then
        ' Locked on this machine.
        sError = "Some other part of this application " _
         & "on this machine has locked this record."
    ElseIf err.Number = 3260 Or err.Number = 3197 Then
        sError = err.Description
    End If
    Resume ExitHere
End Function

