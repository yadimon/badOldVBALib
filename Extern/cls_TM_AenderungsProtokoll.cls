VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_TM_AenderungsProtokoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'############################
' ToDo's:
' -
'
' Langfristige Wünsche:
' - Änderungen auf Wunsch rückgängig
'   machen (???)


'############################
' API's
'############################
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SafeArrayGetDim Lib "oleaut32" (pArray() As Any) As Long

'############################
' Enumerationen
'############################

Public Enum tmEnum_DatensatzNeu
    tmc_NurDatensatzNeu = 0
    tmc_DatensatzNeuUndEinzelwerte = 1
End Enum

Public Enum tmEnum_DatensatzGelöscht
    tmc_DatensatzGelöschtSpeichern = 0
    tmc_AlleEinträgeImProtokollLöschen = 1
End Enum

'############################
' Typen
'############################
Private Type TM_Änderung
    DSNr As String
    Wert_Name As String
    Wert_Alt As String
    Wert_Neu As String
    Datum As String
    Benutzer As String
    UserName As String
    ComputerName As String
End Type

'############################
' Membervariablen der Klasse
'############################

Dim WithEvents mp_frmForm As Access.Form
Attribute mp_frmForm.VB_VarHelpID = -1
Dim WithEvents mp_cmdButtonProtokollÖffnen As Access.CommandButton
Attribute mp_cmdButtonProtokollÖffnen.VB_VarHelpID = -1
Dim mp_txtPrimaryKeyControl As Access.TextBox
Dim mp_strProtokollTab As String
Dim mp_strFormName As String
Dim mp_intDatensatzNeu As Integer
Dim mp_intDatensatzGelöscht As Integer
Dim mp_fLöschenProtokollieren As Boolean
Dim m_strIDGelöscht As String
Dim m_aÄnderung() As TM_Änderung
Dim m_db As DAO.Database

'############################
' Properties der Klasse
'############################

Public Property Set Formular(ByRef frmFormular As Access.Form)
'Fehlerbehandlung definieren
On Error GoTo Err_Formular

    Set mp_frmForm = frmFormular
    mp_strFormName = mp_frmForm.Name
    mp_frmForm.BeforeUpdate = "[Event Procedure]"
    mp_frmForm.AfterUpdate = "[Event Procedure]"
    mp_frmForm.AfterUpdate = "[Event Procedure]"
    mp_frmForm.OnDelete = "[Event Procedure]"
    mp_frmForm.AfterDelConfirm = "[Event Procedure]"

'Ende
Exit_Formular:
    On Error Resume Next
    Exit Property

'Fehlerbehandlung
Err_Formular:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_ÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  Formular" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_Formular
    End Select

End Property

Public Property Set ButtonProtokollÖffnen(ByRef cmdNewButtonProtokollÖffnen As Access.CommandButton)
'Fehlerbehandlung definieren
On Error GoTo Err_ButtonProtokollÖffnen

    Set mp_cmdButtonProtokollÖffnen = cmdNewButtonProtokollÖffnen
    mp_cmdButtonProtokollÖffnen.OnClick = "[Event Procedure]"

'Ende
Exit_ButtonProtokollÖffnen:
    On Error Resume Next
    Exit Property

'Fehlerbehandlung
Err_ButtonProtokollÖffnen:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_AenderungsProtokoll" & vbCrLf & _
                   "Prozedur:  ButtonProtokollÖffnen" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_ButtonProtokollÖffnen
    End Select

End Property

Public Property Let ProtokollTabelle(ByVal strProtokollTabelle As String)
'Fehlerbehandlung definieren
On Error GoTo Err_ProtokollTabelle
    
    'Variablen deklarieren
    Dim dbs As Object
    Dim obj As AccessObject
    Dim fTabVorhanden As Boolean
    Dim strTitel As String
    Dim strHinweis As String
    
    'Prüfen, ob Tabelle vorhanden
    Set dbs = Application.CurrentData
    fTabVorhanden = False
    For Each obj In dbs.AllTables
        If obj.Name = strProtokollTabelle Then
            fTabVorhanden = True
        End If
    Next obj

    'Werte zuweisen
    If fTabVorhanden = True Then
        mp_strProtokollTab = strProtokollTabelle
    Else
        strTitel = "Tabelle nicht vorhanden"
        strHinweis = "Sie haben für das TM-ÄnderungsProtokoll folgende" & vbCrLf & _
                     "Einstellungen festgelegt:" & vbCrLf & vbCrLf & _
                     "ProtokollTabelle: " & strProtokollTabelle & vbCrLf & vbCrLf & _
                     "Die angegebene Tabelle existiert nicht."
        MsgBox strHinweis, vbCritical, strTitel
    End If

'Ende
Exit_ProtokollTabelle:
    On Error Resume Next
    Set dbs = Nothing
    Exit Property

'Fehlerbehandlung
Err_ProtokollTabelle:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_ÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  ProtokollTabelle" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_ProtokollTabelle
    End Select

End Property

Public Property Let FormularName(ByVal strFormName As String)
'Fehlerbehandlung definieren
On Error GoTo Error_Handler

    mp_strFormName = strFormName

'Ende
Exit_Here:
    On Error Resume Next
    Exit Property

'Fehlerbehandlung
Error_Handler:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_AenderungsProtokoll " & vbCrLf & _
                   "Prozedur: FormName" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_Here
    End Select

End Property

Public Property Set PrimaryKeyControl(ByRef txtNewPrimaryKeyControl As Access.TextBox)
'Fehlerbehandlung definieren
On Error GoTo Err_PrimaryKeyControl
    
    Set mp_txtPrimaryKeyControl = txtNewPrimaryKeyControl

'Ende
Exit_PrimaryKeyControl:
    On Error Resume Next
    Exit Property

'Fehlerbehandlung
Err_PrimaryKeyControl:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_ÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  PrimaryKeyControl" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_PrimaryKeyControl
    End Select

End Property

Public Property Let DatensatzNeu(ByVal tmDatensatzNeu As tmEnum_DatensatzNeu)
'Fehlerbehandlung definieren
On Error GoTo Err_DatensatzNeu

    mp_intDatensatzNeu = tmDatensatzNeu

'Ende
Exit_DatensatzNeu:
    On Error Resume Next
    Exit Property

'Fehlerbehandlung
Err_DatensatzNeu:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_ÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  DatensatzNeu" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_DatensatzNeu
    End Select

End Property

Public Property Let DatensatzGelöscht(ByVal tmDatensatzGelöscht As tmEnum_DatensatzGelöscht)
'Fehlerbehandlung definieren
On Error GoTo Err_DatensatzGelöscht

    mp_intDatensatzGelöscht = tmDatensatzGelöscht

'Ende
Exit_DatensatzGelöscht:
    On Error Resume Next
    Exit Property

'Fehlerbehandlung
Err_DatensatzGelöscht:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_ÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  DatensatzGelöscht" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_DatensatzGelöscht
    End Select

End Property

Public Property Let LöschenProtokollieren(ByVal fLöschenProtokollieren As Boolean)
'Fehlerbehandlung definieren
On Error GoTo Err_LöschenProtokollieren

    mp_fLöschenProtokollieren = fLöschenProtokollieren

'Ende
Exit_LöschenProtokollieren:
    On Error Resume Next
    Exit Property

'Fehlerbehandlung
Err_LöschenProtokollieren:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_tmÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  LöschenProtokollieren" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_LöschenProtokollieren
    End Select

End Property

'############################
' Methoden der Klasse
'############################

Public Sub Init(Optional frmFormular As Form, _
                Optional cmdButtonProtokollÖffnen As Variant, _
                Optional strProtokollTabelle As Variant, _
                Optional strFormularName As Variant, _
                Optional txtPrimaryKeyControl As Variant, _
                Optional intDatensatzNeu As Variant, _
                Optional intDatensatzGelöscht As Variant, _
                Optional fLöschenProtokollieren As Variant)
'Fehlerbehandlung definieren
On Error GoTo Err_Init

    If Not IsEmpty(frmFormular) Then
        Set Formular = frmFormular
    End If

    If Not IsEmpty(cmdButtonProtokollÖffnen) Then
        Set ButtonProtokollÖffnen = cmdButtonProtokollÖffnen
    End If
    
    If Not IsMissing(strProtokollTabelle) Then
        Let ProtokollTabelle = strProtokollTabelle
    End If
    
    If Not IsMissing(strFormularName) Then
        Let FormularName = strFormularName
    End If
    
    If Not IsEmpty(txtPrimaryKeyControl) Then
        Set PrimaryKeyControl = txtPrimaryKeyControl
    End If
    
    If Not IsMissing(intDatensatzNeu) Then
        Let DatensatzNeu = intDatensatzNeu
    End If
    
    If Not IsMissing(intDatensatzGelöscht) Then
        Let DatensatzGelöscht = intDatensatzGelöscht
    End If
    
    If Not IsMissing(fLöschenProtokollieren) Then
        Let LöschenProtokollieren = fLöschenProtokollieren
    End If

'Ende
Exit_Init:
    On Error Resume Next
    Exit Sub

'Fehlerbehandlung
Err_Init:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_ÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  Init" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_Init
    End Select

End Sub

'############################
' Klasse verwalten
'############################

Private Sub Class_Initialize()
'Fehlerbehandlung definieren
On Error GoTo Err_Class_Initialize

    'Initialisieren
    Set m_db = CurrentDb
    
    'Standardwerte setzen
    mp_intDatensatzNeu = 0
    mp_intDatensatzGelöscht = 0
    mp_fLöschenProtokollieren = True
    mp_strProtokollTab = "tbl_AendungsProtokoll"

'Ende
Exit_Class_Initialize:
    On Error Resume Next
    Exit Sub

'Fehlerbehandlung
Err_Class_Initialize:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_ÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  Class_Initialize" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_Class_Initialize
    End Select

End Sub

Private Sub Class_Terminate()
'Fehlerbehandlung definieren
On Error GoTo Err_Class_Terminate

    'Aufräumen
    Set m_db = Nothing
    Set mp_frmForm = Nothing

'Ende
Exit_Class_Terminate:
    On Error Resume Next
    Exit Sub

'Fehlerbehandlung
Err_Class_Terminate:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_ÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  Class_Terminate" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_Class_Terminate
    End Select

End Sub

'############################
' Auf Ereignisse reagieren
'############################

Private Sub mp_frmForm_BeforeUpdate(Cancel As Integer)
'Fehlerbehandlung definieren
On Error GoTo Err_mp_frmForm_BeforeUpdate

    'Variablen deklarieren
    Dim ctlSteuerElem As Control    'Verweis auf Steuerelemente im frm
    Dim ctl2 As Control             '2. Verweis auf Steuerelemente im frm
    Dim strAlterWert As String      'Inhalt des alten Wertes
    Dim strNeuerWert As String      'Inhalt des neuen Wertes
    Dim strDatensatzNr As String    'Inhalt Feld Primärschlüssel
    Dim strDatum As String          'Datum der Änderung
    Dim strName As String           'Name des geänderten Feldes
    Dim strName2 As String          '2. Name des geänderten Feldes
                                    '(für Optionsgruppen benötigt)
    Dim strBenutzer As String       'Name des Benutzers
    Dim strUserName As String       'UserName aus Anmeldung
    Dim strComputerName As String   'Name des Computers
    Dim fSpeichern As Boolean       'Sind Änderungen zu speichern?
    Dim intAnz As Integer
    Dim fPrüfen As Boolean
    
    'Initialisierung
    strDatum = Now
    strBenutzer = CurrentUser
    strDatensatzNr = mp_txtPrimaryKeyControl.Value
    strUserName = ErmittleUserName
    strComputerName = ErmittleComputerName
    Erase m_aÄnderung()
    intAnz = 0
    
    'Neuer Datensatz
    If mp_frmForm.NewRecord Then
        intAnz = 1
        ReDim m_aÄnderung(1)
        m_aÄnderung(1).DSNr = strDatensatzNr
        m_aÄnderung(1).Wert_Name = "Datensatz angelegt"
        m_aÄnderung(1).Wert_Alt = ""
        m_aÄnderung(1).Wert_Neu = ""
        m_aÄnderung(1).Datum = strDatum
        m_aÄnderung(1).Benutzer = strBenutzer
        m_aÄnderung(1).UserName = strUserName
        m_aÄnderung(1).ComputerName = strComputerName
        'Umgang mit neuen Datensätzen prüfen
        If mp_intDatensatzNeu = 0 Then
            Exit Sub
        End If
    End If

    fSpeichern = False
    For Each ctlSteuerElem In mp_frmForm.Controls
        Select Case ctlSteuerElem.ControlType
            Case acTextBox
                strAlterWert = Nz(ctlSteuerElem.OldValue, "LEER")
                strNeuerWert = Nz(ctlSteuerElem.Value, "LEER")
                If StrComp(strAlterWert, strNeuerWert, vbBinaryCompare) <> 0 Then
                    If ctlSteuerElem.Controls.Count > 0 Then
                        strName = ctlSteuerElem.Controls(0).Caption
                    End If
                    fSpeichern = True
                End If
            'Bei Optionsgruppen
            Case acOptionGroup
                'Alten Wert ermitteln
                strAlterWert = Nz(ctlSteuerElem.OldValue, "LEER")
                If strAlterWert = "-1" Then
                    strAlterWert = "LEER"
                End If
                'Neuen Wert ermitteln
                strNeuerWert = Nz(ctlSteuerElem.Value, "LEER")
                'Wenn Werte unterschiedlich
                If strAlterWert <> strNeuerWert Then
                    If ctlSteuerElem.Controls.Count > 0 Then
                        strName = ctlSteuerElem.Controls(0).Caption
                    End If
                    For Each ctl2 In ctlSteuerElem.Controls
                        If ctl2.ControlType <> acLabel Then
                        Select Case ctl2.ControlType
                            'Beschriftung Alt und Neu für Umschaltflächen
                            Case acToggleButton
                                If ctl2.OptionValue = strNeuerWert Then
                                    strNeuerWert = ctl2.Caption
                                End If
                                If ctl2.OptionValue = strAlterWert Then
                                    strAlterWert = ctl2.Caption
                                End If
                            'Beschriftung Alt und Neu für OptionButton und CheckBox
                            Case Else
                            'Beschriftung "Neuer Wert" ermittteln
                            If ctl2.OptionValue = strNeuerWert Then
                                strName2 = ctl2.Name
                                If ctl2.Controls.Count > 0 Then
                                    strNeuerWert = ctl2.Controls(0).Caption
                                End If
                            End If
                            'Beschriftung "Alter Wert" ermitteln
                            If ctl2.OptionValue = strAlterWert Then
                                strName2 = ctl2.Name
                                If ctl2.Controls.Count > 0 Then
                                    strAlterWert = ctl2.Controls(0).Caption
                                End If
                            End If
                        End Select
                        End If
                    Next
                    fSpeichern = True
                End If
            Case acComboBox
                strAlterWert = Nz(ctlSteuerElem.OldValue, "LEER")
                strNeuerWert = Nz(ctlSteuerElem.Value, "LEER")
                If strAlterWert <> strNeuerWert Then
                    'Bei mehr als 1 Spalte
                    If ctlSteuerElem.ColumnCount > 1 Then
                        Call ComboDetails(ctlSteuerElem, strAlterWert, strNeuerWert)
                    End If
                    If ctlSteuerElem.Controls.Count > 0 Then
                        strName = ctlSteuerElem.Controls(0).Caption
                    End If
                    fSpeichern = True
                End If
            Case acListBox
                If ctlSteuerElem.MultiSelect = 0 Then
                    strAlterWert = Nz(ctlSteuerElem.OldValue, "LEER")
                    strNeuerWert = Nz(ctlSteuerElem.Value, "LEER")
                    'Bei mehr als 1 Spalte
                    If ctlSteuerElem.ColumnCount > 1 Then
                        Call ListDetails(ctlSteuerElem, strAlterWert, strNeuerWert)
                    End If
                    If strAlterWert <> strNeuerWert Then
                        If ctlSteuerElem.Controls.Count > 0 Then
                            strName = ctlSteuerElem.Controls(0).Caption
                        End If
                        fSpeichern = True
                    End If
                End If
            Case acOptionButton
                fPrüfen = False
                If ctlSteuerElem.Parent.Name = mp_frmForm.Name Then
                    fPrüfen = True
                Else
                    If ctlSteuerElem.Parent.ControlType <> acOptionGroup Then
                        fPrüfen = True
                    End If
                End If
                If fPrüfen = True Then
                    strAlterWert = Nz(ctlSteuerElem.OldValue, "LEER")
                    strAlterWert = IIf(strAlterWert = "0", "Nein", "Ja")
                    strNeuerWert = Nz(ctlSteuerElem.Value, "LEER")
                    strNeuerWert = IIf(strNeuerWert = "0", "Nein", "Ja")
                    If strAlterWert <> strNeuerWert Then
                        If ctlSteuerElem.Controls.Count > 0 Then
                            strName = ctlSteuerElem.Controls(0).Caption
                        End If
                        fSpeichern = True
                    End If
                End If
            Case acCheckBox
                fPrüfen = False
                If ctlSteuerElem.Parent.Name = mp_frmForm.Name Then
                    fPrüfen = True
                Else
                    If ctlSteuerElem.Parent.ControlType <> acOptionGroup Then
                        fPrüfen = True
                    End If
                End If
                If fPrüfen = True Then
                    strAlterWert = Nz(ctlSteuerElem.OldValue, "LEER")
                    strAlterWert = IIf(strAlterWert = "0", "Nein", "Ja")
                    strNeuerWert = Nz(ctlSteuerElem.Value, "LEER")
                    strNeuerWert = IIf(strNeuerWert = "0", "Nein", "Ja")
                    If strAlterWert <> strNeuerWert Then
                        If ctlSteuerElem.Controls.Count > 0 Then
                            strName = ctlSteuerElem.Controls(0).Caption
                        End If
                        fSpeichern = True
                    End If
                End If
            Case acToggleButton
                fPrüfen = False
                If ctlSteuerElem.Parent.Name = mp_frmForm.Name Then
                    fPrüfen = True
                Else
                    If ctlSteuerElem.Parent.ControlType <> acOptionGroup Then
                        fPrüfen = True
                    End If
                End If
                If fPrüfen = True Then
                    strAlterWert = Nz(ctlSteuerElem.OldValue, "LEER")
                    strAlterWert = IIf(strAlterWert = "0", "Nein", "Ja")
                    strNeuerWert = Nz(ctlSteuerElem.Value, "LEER")
                    strNeuerWert = IIf(strNeuerWert = "0", "Nein", "Ja")
                    If strAlterWert <> strNeuerWert Then
                        strName = ctlSteuerElem.Caption
                        fSpeichern = True
                    End If
                End If
        End Select
        'Ergebnis speichern
        If fSpeichern = True Then
            'Prüfen, ob Label vorhanden
            If strName = "" Then
                strName = "???"
            End If
            intAnz = intAnz + 1
            ReDim Preserve m_aÄnderung(intAnz)
            m_aÄnderung(intAnz).DSNr = strDatensatzNr
            m_aÄnderung(intAnz).Wert_Name = strName
            m_aÄnderung(intAnz).Wert_Alt = Left$(strAlterWert, 255)
            m_aÄnderung(intAnz).Wert_Neu = Left$(strNeuerWert, 255)
            m_aÄnderung(intAnz).Datum = strDatum
            m_aÄnderung(intAnz).Benutzer = strBenutzer
            m_aÄnderung(intAnz).UserName = strUserName
            m_aÄnderung(intAnz).ComputerName = strComputerName
            fSpeichern = False
        End If
    Next

'Ende
Exit_mp_frmForm_BeforeUpdate:
    On Error Resume Next
    Exit Sub

'Fehlerbehandlung
Err_mp_frmForm_BeforeUpdate:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_tmÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  mp_frmForm_BeforeUpdate" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_mp_frmForm_BeforeUpdate
    End Select

End Sub

Private Sub mp_frmForm_AfterUpdate()
'Fehlerbehandlung definieren
On Error GoTo Err_mp_frmForm_AfterUpdate

    'Variablen deklarieren
    Dim rst As DAO.Recordset
    Dim intI As Integer

    'Abbruch, wenn kein Wert gespeichert
    If SafeArrayGetDim(m_aÄnderung) = 0 Then
        Exit Sub
    End If
    
    'Initialisieren
    Set rst = m_db.OpenRecordset(mp_strProtokollTab)
    
    'Zwischengespeicherte Änderungen seichern
    With rst
        For intI = 1 To UBound(m_aÄnderung)
            .AddNew
            !FormName = mp_strFormName
            !Datensatznr = m_aÄnderung(intI).DSNr
            !Wert_Name = m_aÄnderung(intI).Wert_Name
            If Len(Trim$(m_aÄnderung(intI).Wert_Alt)) > 0 Then
                !Wert_Alt = m_aÄnderung(intI).Wert_Alt
            End If
            If Len(Trim$(m_aÄnderung(intI).Wert_Neu)) > 0 Then
                !Wert_Neu = m_aÄnderung(intI).Wert_Neu
            End If
            !Datum = m_aÄnderung(intI).Datum
            !Benutzer = m_aÄnderung(intI).Benutzer
            !UserName = m_aÄnderung(intI).UserName
            !ComputerName = m_aÄnderung(intI).ComputerName
            .Update
        Next intI
    End With

'Ende
Exit_mp_frmForm_AfterUpdate:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Sub

'Fehlerbehandlung
Err_mp_frmForm_AfterUpdate:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_AenderungsProtokoll" & vbCrLf & _
                   "Prozedur:  mp_frmForm_AfterUpdate" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_mp_frmForm_AfterUpdate
    End Select

End Sub

Private Sub mp_frmForm_Delete(Cancel As Integer)
'Fehlerbehandlung definieren
On Error GoTo Err_mp_frmForm_Delete

    If mp_fLöschenProtokollieren = True Then
        m_strIDGelöscht = mp_txtPrimaryKeyControl.Value
    End If

'Ende
Exit_mp_frmForm_Delete:
    On Error Resume Next
    Exit Sub

'Fehlerbehandlung
Err_mp_frmForm_Delete:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_tmÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  mp_frmForm_Delete" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_mp_frmForm_Delete
    End Select

End Sub

Private Sub mp_frmForm_AfterDelConfirm(Status As Integer)
'Fehlerbehandlung definieren
On Error GoTo Err_mp_frmForm_AfterDelConfirm

    'Variablen deklarieren
    Dim rst As DAO.Recordset
    Dim strDatum As String          'Datum der Löschung
    Dim strBenutzer As String       'Name des Benutzers
    Dim strSQL As String            'SQL-String für Löschen der Datensätze
    Dim strUserName As String       'UserName aus Anmeldung
    Dim strComputerName As String   'Name des Computers
    
    If mp_fLöschenProtokollieren = True Then

        If Status = acDeleteOK Then
            Select Case mp_intDatensatzGelöscht
                Case 0  '>Datensatz gelöscht< speichern
                    Set rst = m_db.OpenRecordset(mp_strProtokollTab)
                    strDatum = Now
                    strBenutzer = CurrentUser
                    strUserName = ErmittleUserName
                    strComputerName = ErmittleComputerName
                    With rst
                        .AddNew
                        !FormName = mp_strFormName
                        !Datensatznr = m_strIDGelöscht
                        !Wert_Name = "Datensatz gelöscht"
                        !Datum = strDatum
                        !Benutzer = strBenutzer
                        !UserName = strUserName
                        !ComputerName = strComputerName
                        .Update
                    End With
                Case 1  'Alle Einträge für DS im Änderungsprotokoll löschen
                    strSQL = "DELETE * FROM " & mp_strProtokollTab & " WHERE Datensatznr='" & m_strIDGelöscht & "';"
                    m_db.Execute strSQL
            End Select
        End If
    End If

'Ende
Exit_mp_frmForm_AfterDelConfirm:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Sub

'Fehlerbehandlung
Err_mp_frmForm_AfterDelConfirm:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_tmÄnderungsProtokoll" & vbCrLf & _
                   "Prozedur:  mp_frmForm_AfterDelConfirm" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_mp_frmForm_AfterDelConfirm
    End Select

End Sub

Private Sub mp_cmdButtonProtokollÖffnen_Click()
'Fehlerbehandlung definieren
On Error GoTo Err_mp_cmdButtonProtokollÖffnen_Click

    'Variablen deklarieren
    Dim strDocName As String
    Dim strOpenArgs As String
    
    strDocName = "zfrmAenderungsprotokoll"
    strOpenArgs = mp_strProtokollTab & "," & mp_strFormName & "," & mp_txtPrimaryKeyControl.Value
    
    DoCmd.OpenForm strDocName, , , , , , strOpenArgs

'Ende
Exit_mp_cmdButtonProtokollÖffnen_Click:
    On Error Resume Next
    Exit Sub

'Fehlerbehandlung
Err_mp_cmdButtonProtokollÖffnen_Click:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & _
                   "Klassenmodul cls_TM_AenderungsProtokoll" & vbCrLf & _
                   "Prozedur:  mp_cmdButtonProtokollÖffnen_Click" & vbCrLf & _
                   "Fehler-Nr.: " & Err.Number & vbCrLf & _
                   Err.Description, vbCritical
            Resume Exit_mp_cmdButtonProtokollÖffnen_Click
    End Select

End Sub

'############################
' Hilfsfunktionen
'############################

Private Function ErmittleComputerName() As String
'
'    Name: ErmittleComputerName
'   Zweck: Ermittelt den Computernamen durch API-Call
'
'   Autor: Thomas Möller
'          Access@Team-Moeller.de
'
'Erstellt: 01.03.2003
'  Update: 01.03.2003
' Version: 1.00
'
'   Input: -
'
'  Output: - Ermittelter Computername
'
'Benötigt: - Deklaration von "GetComputerName"
'
'  Anmer-: -
'  kungen:
'

'Fehlerbehandlung definieren
On Error GoTo Fehlerbehandlung

    'Variablen deklarieren
    Dim strBuffer As String
    Dim lSize As Long
    
    'Buffer bereitstellen
    lSize = 32
    strBuffer = String(lSize, Chr$(0))
    
    'Compternamen ermitteln
    GetComputerName strBuffer, lSize
    
    'Computernamen übergeben
    ErmittleComputerName = Left(strBuffer, lSize)

'Ende
Ausstieg:
    On Error Resume Next
    Exit Function

'Fehlerbehandlung
Fehlerbehandlung:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox Err.Description, vbCritical, Err.Number
            Resume Ausstieg
    End Select
    
End Function

Private Function ErmittleUserName() As String
'
'    Name: ErmittleUserName
'   Zweck: Ermittelt den Usernamen durch API-Call
'
'   Autor: Thomas Möller
'          Access@Team-Moeller.de
'
'Erstellt: 01.03.2003
'  Update: 01.03.2003
' Version: 1.00
'
'   Input: -
'
'  Output: - Ermittelter Username
'
'Benötigt: - Deklaration von "GetUserName"
'
'  Anmer-: -
'  kungen:
'

'Fehlerbehandlung definieren
On Error GoTo Fehlerbehandlung

    'Variablen deklarieren
    Dim strBuffer As String
    
    'Buffer bereitstellen
    strBuffer = String(100, Chr$(0))
    
    'Compternamen ermitteln
    GetUserName strBuffer, 100
    
    'Computernamen übergeben
    ErmittleUserName = Left$(strBuffer, InStr(strBuffer, Chr$(0)) - 1)

'Ende
Ausstieg:
    On Error Resume Next
    Exit Function

'Fehlerbehandlung
Fehlerbehandlung:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox Err.Description, vbCritical, Err.Number
            Resume Ausstieg
    End Select
    
End Function

Private Sub ComboDetails(ctlCombo As Control, ByRef strAlterWert As String, ByRef strNeuerWert As String)
'
'    Name: ComboDetails
'   Zweck: Ermittelt für Kombinationsfelder den sichtbaren
'          bisherigen Wert und den sichtbaren neuen Wert
'
'   Autor: Thomas Möller
'          Access@Team-Moeller.de
'
'Erstellt: 02.05.2003
'  Update: 04.05.2003
' Version: 1.50
'
'   Input: - KombinationsfeldSteuerelement
'          - Alter Wert
'          - Neuer Wert
'
'  Output: - strAlterWert wird ggfs. geändert
'          - strNeuerWert wird ggfs. geändert
'
'Benötigt: -
'
'  Anmer-: - Funktion nicht für "Benutzerdefinierte Funktion"
'  kungen:   als Datenquelle verfügbar
'

'Fehlerbehandlung definieren
On Error GoTo Fehlerbehandlung
    
    'Variablen deklarieren
    Dim intAnzSpalten As Integer
    Dim intGebSpalte As Integer
    Dim bytErsteSichtSp As Byte
    Dim strSpaltenBreiten As String
    Dim intSpalte As Integer
    Dim intPos As Integer
    Dim intCount As Integer
    Dim acboWerte As Variant
    Dim rst As DAO.Recordset

    'Daten ermitteln
    intAnzSpalten = ctlCombo.ColumnCount
    intGebSpalte = ctlCombo.BoundColumn
    
    'Erste sichtbare Spalte ermitteln
    strSpaltenBreiten = ctlCombo.ColumnWidths
    If strSpaltenBreiten = "" Then
        bytErsteSichtSp = 1
    Else
        If strSpaltenBreiten = "0" Then
            bytErsteSichtSp = 2
        Else
            intPos = InStr(strSpaltenBreiten, ";")
            intCount = 1
            Do Until intPos = 1 Or intPos = 0
                If CDbl(Left$(strSpaltenBreiten, intPos - 1)) <> 0 Then
                    Exit Do
                End If
                strSpaltenBreiten = Right$(strSpaltenBreiten, Len(strSpaltenBreiten) - intPos)
                intPos = InStr(strSpaltenBreiten, ";")
                intCount = intCount + 1
                If strSpaltenBreiten = "0" Then
                    intCount = intCount + 1
                    Exit Do
                End If
            Loop
            bytErsteSichtSp = intCount
        End If
    End If
    
    'Gebundene Spalte sichtbar?
    If bytErsteSichtSp = intGebSpalte Then
        Exit Sub
    End If
    
    'Werte in sichtbarer Spalte ermitteln
    If strAlterWert = "LEER" Then
        strNeuerWert = Nz(ctlCombo.Column(bytErsteSichtSp - 1), "LEER") & " (" & strNeuerWert & ")"
    Else
        Select Case ctlCombo.RowSourceType
            Case "Table/Query"      'Tabelle/Abfrage
                Set rst = m_db.OpenRecordset(ctlCombo.RowSource)
                rst.MoveFirst
                Do While Not rst.EOF
                    If rst(intGebSpalte - 1) = strAlterWert Then
                        strAlterWert = rst(bytErsteSichtSp - 1) & " (" & strAlterWert & ")"
                        strNeuerWert = Nz(ctlCombo.Column(bytErsteSichtSp - 1), "LEER") & " (" & strNeuerWert & ")"
                        Exit Do
                    End If
                    rst.MoveNext
                Loop
            Case "Value List"       'Werteliste
                'Array erstellen
                acboWerte = Split(ctlCombo.RowSource, ";")
                'Nach altem Wert suchen
                For intPos = intGebSpalte - 1 To UBound(acboWerte) Step intAnzSpalten
                    If acboWerte(intPos) = strAlterWert Then
                        strAlterWert = acboWerte(intPos - intGebSpalte + bytErsteSichtSp) & " (" & strAlterWert & ")"
                        strNeuerWert = Nz(ctlCombo.Column(bytErsteSichtSp - 1), "LEER") & " (" & strNeuerWert & ")"
                        Exit For
                    End If
                Next
            Case "Field List"       'Feldliste
                '
                'Kein besonderer Code notwendig. Es gibt eigentlich nur eine Spalte.
                'Damit ist die erste sichtbare Spalte immer gleich der gebundenen Spalte.
                '
            Case Else               'Benutzerdefinierte Funktion
                '
                'Wird derzeit nicht unterstützt.
                '
        End Select
    End If
    
'Ende
Ausstieg:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Sub
    
'Fehlerbehandlung
Fehlerbehandlung:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox Err.Description, vbCritical, Err.Number
            Resume Ausstieg
    End Select

End Sub
Private Sub ListDetails(ctlList As Control, ByRef strAlterWert As String, ByRef strNeuerWert As String)
'
'    Name: ListDetails
'   Zweck: Ermittelt für Listenfelder die sichtbaren
'          bisherigen Werte und die sichtbaren neuen Werte
'
'   Autor: Thomas Möller
'          Access@Team-Moeller.de
'
'Erstellt: 11.05.2003
'  Update: 01.06.2003
' Version: 1.55
'
'   Input: - ListenfeldSteuerelement
'          - Alter Wert
'          - Neuer Wert
'
'  Output: - strAlterWert wird ggfs. geändert
'          - strNeuerWert wird ggfs. geändert
'
'Benötigt: -
'
'  Anmer-: - Funktion nicht für "Benutzerdefinierte Funktion"
'  kungen:   als Datenquelle verfügbar
'

'Fehlerbehandlung definieren
On Error GoTo Fehlerbehandlung
    
    'Variablen deklarieren
    Dim strAlterWertBisher As String
    Dim strNeuerWertBisher As String
    Dim intAnzSpalten As Integer
    Dim intGebSpalte As Integer
    Dim fSpalteSichtb() As Boolean
    Dim strSpaltenBreiten As String
    Dim intI As Integer
    Dim intCount As Integer
    Dim intPos As Integer
    Dim acboWerte As Variant
    Dim rst As DAO.Recordset
    Dim intSpalte As Integer

    'Daten ermitteln
    strAlterWertBisher = strAlterWert
    strNeuerWertBisher = strNeuerWert
    intAnzSpalten = ctlList.ColumnCount
    ReDim fSpalteSichtb(intAnzSpalten)
    intGebSpalte = ctlList.BoundColumn
    
    'Sichtbare Spalten ermitteln
    strSpaltenBreiten = ctlList.ColumnWidths
    If strSpaltenBreiten = "" Then
        For intI = 1 To intAnzSpalten
            fSpalteSichtb(intI) = True
        Next
    Else
        If strSpaltenBreiten = "0" Then
            fSpalteSichtb(1) = False
            For intI = 2 To intAnzSpalten
                fSpalteSichtb(intI) = True
            Next
        Else
            intCount = 1
            intPos = InStr(strSpaltenBreiten, ";")
            Do Until intPos = 0
                If intPos = 1 Then
                    fSpalteSichtb(intCount) = True
                Else
                    If Left$(strSpaltenBreiten, intPos - 1) = 0 Then
                        fSpalteSichtb(intCount) = False
                    Else
                        fSpalteSichtb(intCount) = True
                    End If
                End If
                strSpaltenBreiten = Right$(strSpaltenBreiten, Len(strSpaltenBreiten) - intPos)
                intPos = InStr(strSpaltenBreiten, ";")
                intCount = intCount + 1
            Loop
            If Len(strSpaltenBreiten) > 0 Then
                If strSpaltenBreiten = 0 Then
                    fSpalteSichtb(intCount) = False
                Else
                    fSpalteSichtb(intCount) = True
                End If
                intCount = intCount + 1
            End If
            For intI = intCount To intAnzSpalten
                fSpalteSichtb(intCount) = True
            Next
        End If
    End If
    
    'Werte in sichtbaren Spalten ermitteln
    If strAlterWert = "LEER" Then
        strNeuerWert = ""
        For intI = 1 To intAnzSpalten
            If fSpalteSichtb(intI) Then
                strNeuerWert = strNeuerWert & ctlList.Column(intI - 1) & ", "
            End If
        Next
        strNeuerWert = Left$(strNeuerWert, Len(strNeuerWert) - 2) & " (" & strNeuerWertBisher & ")"
    Else
        Select Case ctlList.RowSourceType
            Case "Table/Query"      'Tabelle/Abfrage
                Set rst = m_db.OpenRecordset(ctlList.RowSource)
                rst.MoveFirst
                Do While Not rst.EOF
                    If rst(intGebSpalte - 1) = strAlterWert Then
                        'Alten Wert ermitteln
                        strAlterWert = ""
                        For intI = 1 To intAnzSpalten
                            If fSpalteSichtb(intI) Then
                                strAlterWert = strAlterWert & rst(intI - 1) & ", "
                            End If
                        Next
                        strAlterWert = Left$(strAlterWert, Len(strAlterWert) - 2) & " (" & strAlterWertBisher & ")"
                        'Neuen Wert ermitteln
                        strNeuerWert = ""
                        For intI = 1 To intAnzSpalten
                            If fSpalteSichtb(intI) Then
                                strNeuerWert = strNeuerWert & ctlList.Column(intI - 1) & ", "
                            End If
                        Next
                        strNeuerWert = Left$(strNeuerWert, Len(strNeuerWert) - 2) & " (" & strNeuerWertBisher & ")"
                        'Fertig
                        Exit Do
                    End If
                    rst.MoveNext
                Loop
            Case "Value List"       'Werteliste
                'Array erstellen
                acboWerte = Split(ctlList.RowSource, ";")
                'Nach altem Wert suchen
                For intPos = intGebSpalte - 1 To UBound(acboWerte) Step intAnzSpalten
                    If acboWerte(intPos) = strAlterWert Then
                        'Alten Wert ermitteln
                        strAlterWert = ""
                        For intI = 1 To intAnzSpalten
                            If fSpalteSichtb(intI) Then
                                strAlterWert = strAlterWert & acboWerte(intPos - intGebSpalte + intI) & ", "
                            End If
                        Next
                        strAlterWert = Left$(strAlterWert, Len(strAlterWert) - 2) & " (" & strAlterWertBisher & ")"
                        'Neuen Wert ermitteln
                        strNeuerWert = ""
                        For intI = 1 To intAnzSpalten
                            If fSpalteSichtb(intI) Then
                                strNeuerWert = strNeuerWert & ctlList.Column(intI - 1) & ", "
                            End If
                        Next
                        strNeuerWert = Left$(strNeuerWert, Len(strNeuerWert) - 2) & " (" & strNeuerWertBisher & ")"
                        'Fertig
                        Exit For
                    End If
                Next
            Case "Field List"       'Feldliste
                '
                'Kein besonderer Code notwendig. Es gibt eigentlich nur eine Spalte.
                'Damit ist die erste sichtbare Spalte immer gleich der gebundenen Spalte.
                '
            Case Else               'Benutzerdefinierte Funktion
                '
                'Wird derzeit nicht unterstützt.
                '
        End Select
    End If
    
'Ende
Ausstieg:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Sub
    
'Fehlerbehandlung
Fehlerbehandlung:
    Select Case Err.Number
        Case 0
            Resume Next
        Case Else
            MsgBox Err.Description, vbCritical, Err.Number
            Resume Ausstieg
    End Select

End Sub

