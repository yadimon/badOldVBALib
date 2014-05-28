VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDynChoise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsDynChoise
' Author    : Dmitry Gorelenkov
' Date      : 01.10.2012
' Changed   : 29.10.2013
' Purpose   : Listenfelder, fuer ein bestimmtes Auswahl
'           : 'btnOK button in Form muss existieren, und lstList1, .... lstListX versteckt sein
' Requires  : clsFuncs, clsDB, clsMyCollection, clsFuncsAccs
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
Private WithEvents m_frmWorkForm As Form
Attribute m_frmWorkForm.VB_VarHelpID = -1
Private WithEvents m_btnOK As CommandButton
Attribute m_btnOK.VB_VarHelpID = -1
Private m_rsRules As ADODB.Recordset
Private m_sRulesTableName As String
Private f As clsFuncs
Private DB As clsDB
Private fAccs As clsFuncsAccs
Private bInitDone As Boolean
Private m_iMaxLists As Integer
Private m_iStartId As Integer
Private m_rsTables As ADODB.Recordset
Private m_sField As String
Private m_iOperation As Integer
Private m_sOperField As String
Private m_lblRglNr As Access.Label
Private m_aCurrentSelectionCodes As Variant
Private m_txtFocusTaker As Access.TextBox


'later: bind dinamic eventhandler (extra class)
Private WithEvents m_lstList1 As ListBox
Attribute m_lstList1.VB_VarHelpID = -1
Private WithEvents m_lstList2 As ListBox
Attribute m_lstList2.VB_VarHelpID = -1
Private WithEvents m_lstList3 As ListBox
Attribute m_lstList3.VB_VarHelpID = -1
Private WithEvents m_lstList4 As ListBox
Attribute m_lstList4.VB_VarHelpID = -1
Private WithEvents m_lstList5 As ListBox
Attribute m_lstList5.VB_VarHelpID = -1

Private Type dataSet
    sTableName As String
    sDBTableName As String
    lstListControl As Access.ListBox
    lblListLabel As Access.Label
    mcValues As clsMyCollection

End Type

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::EVENTS::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Event operationSelected(iOperationNr As Integer)


Private m_aMainArray() As dataSet


'Initialisation
Public Function init(ByRef frmFormPlaceholder As Access.Form, sRulesTableName As String, _
                     Optional sTablesTableName As String, _
                     Optional combtnBtn As Access.CommandButton) As Boolean

    Dim i As Integer
    Dim rsDBTable As ADODB.Recordset
    Dim mcTmpCollection As clsMyCollection

    '==========VARIABLEN=====================================================================================
    'array mit folge von selection codes
    On Error GoTo init_Error

    m_aCurrentSelectionCodes = Array()
    'maximale lists anzahl
    m_iMaxLists = 5
    'start id
    m_iStartId = 1
    'tabellenname mit regeln
    m_sRulesTableName = sRulesTableName
    'fieldprefix
    m_sField = "dlStep"
    'fieldname operation
    m_sOperField = "dlOper"


    'label fuer datensatz nr 'remove?
    Set m_lblRglNr = frmFormPlaceholder.lblRglNr

    'focus taker element
    Set m_txtFocusTaker = frmFormPlaceholder.txtFocusTaker


    ReDim m_aMainArray(m_iStartId To m_iMaxLists)

    'OK button setzen
    If combtnBtn Is Nothing Then
        On Error Resume Next
        Set m_btnOK = frmFormPlaceholder.btnOK
        If err.Number <> 0 Then
            f.dbg "OK-Button nicht gefunden"
            Exit Function
        End If
    Else
        Set m_btnOK = combtnBtn
    End If
    m_btnOK.OnClick = "[Event Procedure]"

    'lists mit events binden
    With frmFormPlaceholder
        Set m_lstList1 = .lstList1
        Set m_lstList2 = .lstList2
        Set m_lstList3 = .lstList3
        Set m_lstList4 = .lstList4
        Set m_lstList5 = .lstList5
    End With


    '==========DATEN=============================================================================================
    Set m_frmWorkForm = frmFormPlaceholder
    Set m_rsRules = DB.SQL("SELECT * FROM " & m_sRulesTableName & " WHERE dlDeaktiviert = FALSE", , , , True)
    Set m_rsTables = DB.SQL(sTablesTableName, , , , True)

    'main collection mit data erstellen
    For i = m_iStartId To m_iMaxLists
        Dim dataSet As dataSet
        Set mcTmpCollection = New clsMyCollection

        'form lists speichern
        Set dataSet.lstListControl = m_frmWorkForm.Controls("lstList" & i)
        Set dataSet.lblListLabel = m_frmWorkForm.Controls("lblList" & i)
        dataSet.lstListControl.AfterUpdate = "[Event Procedure]"


        m_rsTables.Filter = "dlvID = " & i
        '        m_rsTables.OpenRecordset
        'test ob tabelle vorhanden
        If DB.isEmptyRS(m_rsTables) Then
            MsgBox "kann DynChoise nicht initialisieren: tabelle ID: " & i & " nicht gefunden!"
        End If

        dataSet.sTableName = m_rsTables.Fields("dlvValue")
        dataSet.sDBTableName = m_rsTables.Fields("dlvRefTableName")

        'daten aus der zugehoeriger tabelle laden und in tmp collection speichern
        Set rsDBTable = DB.SQL(dataSet.sDBTableName, , , , True)
        rsDBTable.MoveFirst
        'fuer jeden wert in tabelle
        Do Until rsDBTable.EOF
            'tmp collection fuellen. key = ID, value = Name
            Call mcTmpCollection.Add(rsDBTable.Fields(1).Value, rsDBTable.Fields(0).Value)
            rsDBTable.MoveNext
        Loop

        'in dataset speichern
        Set dataSet.mcValues = mcTmpCollection
        m_aMainArray(i) = dataSet

    Next i

    '
    bInitDone = True

    init = True


    Exit Function

init_Error:
    f.dbg "init of Klassenmodul clsDynChoise2: Kann nicht initialisieren"


End Function

'auswahl zurucksetzen
Public Sub reset()
    m_aCurrentSelectionCodes = Array()
    Call load
End Sub

'lists refreshen
Public Sub load()
    If Not bInitDone Then Exit Sub
    Dim iCodesLength As Integer
    Dim iElementsLoaded As Integer
    Dim dsNextDataset As dataSet
    Dim aValuesToPut As Variant

    On Error Resume Next
    'auf button focus setzen, damit restliche elemente ausgeblendet werden koennen
'        m_txtFocusTaker.SetFocus
    On Error GoTo 0


    'anzahl schon markierter/gespeicherter lists
    iCodesLength = f.aLength(m_aCurrentSelectionCodes)
    

    'falls noch nicht die letzte liste
    If iCodesLength < m_iMaxLists Then
        'nur vorherige listen anzeigen
        notReadyForOper
        showUntil (iCodesLength)

        'nexte liste und values zu zeigen
        dsNextDataset = m_aMainArray(iCodesLength + 1)
        aValuesToPut = getNextValues(m_aCurrentSelectionCodes)
        'elements geladen
        iElementsLoaded = f.aLength(aValuesToPut)

        'falls keine elements mehr, dann readyForOper
        If iElementsLoaded <= 0 Then
            Call readyForOper
            'sonst laden
        Else
            Call loadList(dsNextDataset, aValuesToPut)
        End If


        'sonst, falls letzte liste schon ausgewaehlt, dann readyForOper
    Else
        Call readyForOper
    End If



    'falls nur 1 element geladen, click auf den emulieren
    If iElementsLoaded = 1 Then
        Call fAccs.setListIndex(dsNextDataset.lstListControl, 0)
        Call listSelected(dsNextDataset.lstListControl)
    End If

End Sub

'liste laden/anzeigen und mit aValues befuellen, return anzahl von geladenen elementen
Private Function loadList(dsDataSet As dataSet, aValues As Variant) As Integer

    With dsDataSet
        .lblListLabel.Caption = dsDataSet.sTableName
        .lblListLabel.Visible = True
        .lstListControl.Visible = True
    End With

    If f.aLength(aValues) = 0 Then
        loadList = 0
        Exit Function
    End If

    '-------fuer listbox------
    Dim i As Integer
    Dim arrayValue As Variant
    Dim vElem As Access.ListBox

    Set vElem = dsDataSet.lstListControl

    'alles loeschen
    For i = 0 To vElem.ListCount - 1
        vElem.Value = Null
        vElem.RemoveItem (0)
    Next i

    'neue elemente einfuegen
    For Each arrayValue In aValues
        Call vElem.AddItem(arrayValue & ";" & dsDataSet.mcValues.getValue(CStr(arrayValue)))
    Next arrayValue

    loadList = f.aLength(aValues)
End Function

'zeigt nur die listen bis iUntilNr
Private Function showUntil(iUntilNr As Integer)
    Dim i As Integer
    'case hide all, focus on tmp element
    If iUntilNr = 0 Then
        m_txtFocusTaker.SetFocus
    End If
    
    For i = LBound(m_aMainArray) To UBound(m_aMainArray)
        With m_aMainArray(i)
            'falls kleiner, anzeigen
            If i <= iUntilNr Then
                If Not .lblListLabel.Visible Then .lblListLabel.Visible = True
                If Not .lstListControl.Visible Then .lstListControl.Visible = True

                'sonst verstecken
            Else
                If .lblListLabel.Visible Then .lblListLabel.Visible = False
                If .lstListControl.Visible Then .lstListControl.Visible = False
            End If
        End With
    Next i
End Function

'nechste werte laden, abhaengig von dem gefuellten aArrayOfSelectedIds
Private Function getNextValues(aArrayOfSelectedIds As Variant) As Variant
    Dim rsFilteredRs As ADODB.Recordset
    Dim retArrayTmp As Variant
    Dim i As Integer
    Dim vValue As Variant

    retArrayTmp = Array()

    'darf nicht groesser, oder gleich sein, wenn gleich, dann ist es letzte, also sowieso keine werte mehr
    If f.aLength(aArrayOfSelectedIds) >= m_iMaxLists Then GoTo Final

    Set rsFilteredRs = m_rsRules.Clone
    m_rsRules.Filter = ""

    'filtrieren
    With rsFilteredRs
        .Filter = ""
        For i = LBound(aArrayOfSelectedIds) To UBound(aArrayOfSelectedIds)
            .Filter = IIf(.Filter <> 0, .Filter & " AND ", "") & .Fields(m_sField & (i)).name & "=" & f.ClngTry(aArrayOfSelectedIds(i)) & ""
        Next i
    End With


    'falls keine datensaetze
    If DB.isEmptyRS(rsFilteredRs) Then
        GoTo Final
    End If

    'array mit naechsten werten aufbauen
    rsFilteredRs.MoveFirst
    Do Until rsFilteredRs.EOF
        'moegliches Wert
        vValue = rsFilteredRs.Fields(m_sField & (f.aLength(aArrayOfSelectedIds) + 1)).Value
        'falls nicht null
        If f.hatWert(vValue) Then
            Call f.arrayPush(retArrayTmp, vValue)
        End If

        rsFilteredRs.MoveNext
    Loop


Final:
    getNextValues = f.arrayRemoveDublicates(retArrayTmp)

End Function

Function getOperation(aArrayOfSelectedIds As Variant) As Integer
    Dim retValue As Integer
    Dim rsFilteredRs As ADODB.Recordset
    Dim i As Integer

    retValue = -1

    Set rsFilteredRs = m_rsRules.Clone
    m_rsRules.Filter = ""

    'filtrieren
    With rsFilteredRs
        .Filter = ""
        For i = LBound(aArrayOfSelectedIds) To UBound(aArrayOfSelectedIds)
            .Filter = IIf(.Filter <> 0, .Filter & " AND ", "") & .Fields(m_sField & (i)).name & "=" & aArrayOfSelectedIds(i) & ""
        Next i
    End With

    'falls keine datensaetze
    If DB.isEmptyRS(rsFilteredRs) Then
        f.dbg "keine operation gefunden"
        Resume Final
    End If
    'naechsten Operation auslesen
    rsFilteredRs.MoveFirst
    retValue = rsFilteredRs.Fields(m_sOperField).Value

    'debug zwecken
    If rsFilteredRs.RecordCount > 1 Then
        f.dbg "mehr als 1 recordset, nach dem full filter for operation"
    End If

Final:
    getOperation = retValue
End Function

'ausfuehren wenn nicht bereit fuer operation
Private Function notReadyForOper()
    m_iOperation = -1
    m_btnOK.Enabled = False
End Function

'ausfuehren wenn bereit fuer operation
Private Function readyForOper()
    m_iOperation = getOperation(m_aCurrentSelectionCodes)
    m_btnOK.Enabled = True
End Function

'action das ausgefuehrt wird, wenn listbox selected
Private Function listSelected(lstList As Access.ListBox)
    Dim i As Integer
    Dim iListNum As Integer
    Dim tmpArray As Variant


    iListNum = getListNumber(lstList)

    'vorherige werte uebertragen, falls vorhanden
    For i = m_iStartId To iListNum - 1
        Call f.arrayPush(tmpArray, m_aCurrentSelectionCodes(i))
    Next i
    'neueste selected wert hinzufuegen
    Call f.arrayPush(tmpArray, lstList.Value)

    'array ersetzen
    m_aCurrentSelectionCodes = tmpArray
    Call f.arrayChangeBounds(m_aCurrentSelectionCodes, m_iStartId)

    'refresh
    Call load
End Function

'get number of the list (position)
Private Function getListNumber(lstList As Access.ListBox)
    Dim i As Integer

    For i = LBound(m_aMainArray) To UBound(m_aMainArray)
        If m_aMainArray(i).lstListControl Is lstList Then
            getListNumber = i
            Exit Function
        End If
    Next i

    getListNumber = -1
End Function

'zur zeit selektierte element-ID by tabellenname zuruckgeben
Public Function getActiveValueByListNr(iListNr As Integer) As Long
    Dim lReturn As Long
    On Error GoTo Final

    lReturn = -1

    If LBound(m_aMainArray) > iListNr Or UBound(m_aMainArray) < iListNr Then
        GoTo Final
    End If
    
    If f.arrayCheckIndex(m_aCurrentSelectionCodes, CLng(iListNr)) Then
        lReturn = m_aCurrentSelectionCodes(iListNr)
    End If
    
'    If m_aMainArray(iListNr).lstListControl.Visible Then
'        lReturn = f.ClngTry(Nz(m_aMainArray(iListNr).lstListControl.Value, -1), -1)
'    End If

Final:
    getActiveValueByListNr = lReturn
End Function


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::EVENT HANDLERS::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Private Sub m_btnOk_Click()
'MsgBox "do operation: " & m_iOperation
    On Error GoTo m_btnOK_Click_Error

    RaiseEvent operationSelected(m_iOperation)

    Call reset

    Exit Sub

m_btnOK_Click_Error:
    f.dbg "Ein Fehler ist aufgetreten: " & vbCrLf & _
          "Klassenmodul clsDynChoise" & vbCrLf & _
          "Prozedur:  m_btnOK_Click" & vbCrLf & _
          "Fehler-Nr.: " & err.Number & vbCrLf & _
          err.Description

End Sub

Private Sub m_lstList1_AfterUpdate()
    On Error Resume Next

    Call listSelected(m_lstList1)

    If err.Number <> 0 Then
        f.dbg "Ein Fehler ist aufgetreten: " & vbCrLf & _
              "Klassenmodul clsDynChoise" & vbCrLf & _
              "Prozedur:  m_lstList1_AfterUpdate" & vbCrLf & _
              "Fehler-Nr.: " & err.Number & vbCrLf & _
              err.Description
    End If
End Sub

Private Sub m_lstList2_AfterUpdate()
    On Error Resume Next

    Call listSelected(m_lstList2)

    If err.Number <> 0 Then
        f.dbg "Ein Fehler ist aufgetreten: " & vbCrLf & _
              "Klassenmodul clsDynChoise" & vbCrLf & _
              "Prozedur:  m_lstList2_AfterUpdate" & vbCrLf & _
              "Fehler-Nr.: " & err.Number & vbCrLf & _
              err.Description
    End If
End Sub

Private Sub m_lstList3_AfterUpdate()
    On Error Resume Next

    Call listSelected(m_lstList3)

    If err.Number <> 0 Then
        f.dbg "Ein Fehler ist aufgetreten: " & vbCrLf & _
              "Klassenmodul clsDynChoise" & vbCrLf & _
              "Prozedur:  m_lstList3_AfterUpdate" & vbCrLf & _
              "Fehler-Nr.: " & err.Number & vbCrLf & _
              err.Description
    End If
End Sub

Private Sub m_lstList4_AfterUpdate()
    On Error Resume Next

    Call listSelected(m_lstList4)

    If err.Number <> 0 Then
        f.dbg "Ein Fehler ist aufgetreten: " & vbCrLf & _
              "Klassenmodul clsDynChoise" & vbCrLf & _
              "Prozedur:  m_lstList4_AfterUpdate" & vbCrLf & _
              "Fehler-Nr.: " & err.Number & vbCrLf & _
              err.Description
    End If
End Sub

Private Sub m_lstList5_AfterUpdate()
    On Error Resume Next
    '    Application.Echo (False)
    '    m_frmWorkForm.Painting = False
    Call listSelected(m_lstList5)

    '    Application.Echo (True)
    '    m_frmWorkForm.Painting = True

    If err.Number <> 0 Then
        f.dbg "Ein Fehler ist aufgetreten: " & vbCrLf & _
              "Klassenmodul clsDynChoise" & vbCrLf & _
              "Prozedur:  m_lstList5_AfterUpdate" & vbCrLf & _
              "Fehler-Nr.: " & err.Number & vbCrLf & _
              err.Description
    End If


End Sub




'############################
'Default init, terminate
'############################
Private Sub Class_Initialize()
    Set f = New clsFuncs
    Set DB = New clsDB
    Set fAccs = New clsFuncsAccs
End Sub

Private Sub Class_Terminate()
    Set m_frmWorkForm = Nothing
    Set m_rsRules = Nothing
    Set f = Nothing
    Set DB = Nothing
    Set fAccs = Nothing
End Sub
