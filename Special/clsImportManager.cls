VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private f As New clsFuncs
Private aMitarbeiter As Variant 'mitarbeiter namen /abkuerzungen
Private dDate As Date
Private colMitarbeiterDaten As New clsMyCollection 'collection wo daten sind, Key = Mitarbeiter Value MitarbeiterDatenSet Class (wie array key - value)
Private init_done As Boolean
Private mycolEinstellungen As New clsMyCollection 'array mit einstellungen, key - value
Private errclErrorCollector As New clsErrorCollector

'initialisieren
Public Function init(ByVal dDateArg As Date, ByVal mycolEinstellungenP As clsMyCollection, Optional ByVal aMitarbeiterArg As Variant = Empty)
    
    dDate = dDateArg
    aMitarbeiter = aMitarbeiterArg
    
    Set mycolEinstellungen = mycolEinstellungenP
    
    
    
    init_done = True
End Function

'mitarbeiter array setzen
'aMitarbeiter
Public Property Get mitarbeiter() As Variant
    mitarbeiter = aMitarbeiter
End Property

Public Property Let mitarbeiter(ByVal value As Variant)
    aMitarbeiter = value
End Property


Public Property Get errors() As clsErrorCollector
    Set errors = errclErrorCollector
End Property

Public Property Let errors(ByVal value As clsErrorCollector)
    errclErrorCollector = value
End Property


'default data loader
Public Function getACD(wkbAcdExport As Workbook, acdMuster As Variant)
    getACD = False
    Dim i As Integer
    Dim lAllDataIter As Long
    Dim lAllDataIterSpalte As Long
    Dim aAllData As Variant
    Dim sDataSpalte As String
    Dim wksWorkingSheet As Worksheet
    Dim sSaveSpalte As String
    Dim sZeile As String
    Dim sWhatBlatt As String
    Dim mtrbBlattDaten As clsMitarbeiterDaten
    
    
    On Error GoTo ErrorHandler
    
    'daten sammeln und in array speichern
    Set wksWorkingSheet = wkbAcdExport.Worksheets(1) 'muss existieren, 1.
    aAllData = wksWorkingSheet.UsedRange.value2
    
    'spalte wo die daten bei ACDFILE drin sind
    'sDataSpalte = f.LetterToCol(mycolEinstellungen.getValue("ACDAnrCodDataSpalte"))

    
    'fuer jedes Element  in muster
    For i = 1 To UBound(acdMuster)
        On Error GoTo nexteZeileAusMuster
         'welches blatt
        sWhatBlatt = acdMuster(i, 4)
        'falls irgendwas in muster fehlt, naechster datensatz
        If Not f.hatWertArray(f.array2dToArrOfArrs(acdMuster)(i)) Then GoTo nexteZeileAusMuster

        
        If Left(sWhatBlatt, 1) = "!" Then 'wenn Blatt name aus einstellungen gelesen werden muss
            sWhatBlatt = Right(sWhatBlatt, Len(sWhatBlatt) - 1) '! loeschen
            sWhatBlatt = mycolEinstellungen.getValue(sWhatBlatt) 'aus einstellungen lesen
        End If
        
        'falls das Blatt schon in collection drin ist, mit dem arbeiten, sonst neu erstellen
        If colMitarbeiterDaten.isInArray(sWhatBlatt) Then
            Set mtrbBlattDaten = colMitarbeiterDaten.getValue(sWhatBlatt)
        Else
            Set mtrbBlattDaten = New clsMitarbeiterDaten
            mtrbBlattDaten.init (sWhatBlatt)
        End If
        

        
        'jede zeile in ext. dokument durchsuchen (gearbeitet wird mit dataArray)
        For lAllDataIter = LBound(aAllData) To UBound(aAllData)
            For lAllDataIterSpalte = LBound(aAllData, 2) To UBound(aAllData, 2)
                'falls richtigeZeile (stimmt mit musterwert ueberein)
                If UCase(aAllData(lAllDataIter, lAllDataIterSpalte)) = UCase(acdMuster(i, 1)) Then
                    'spalte wo die daten bei ACDFILE drin sind, abhaengig von gesuchten wert zelle position
                    sDataSpalte = f.LetterToCol(acdMuster(i, 2)) + lAllDataIterSpalte - 1
                    'gesuchte Zeile in export
                    sZeile = lAllDataIter
                    'spalte (key) wo die daten bei MA 99 gespeichert werden
                    sSaveSpalte = acdMuster(i, 3)
                    If IsNumeric(f.LetterToCol(sSaveSpalte)) Then
                        If IsNumeric(sDataSpalte) Then
                            Call mtrbBlattDaten.add(aAllData(CLng(sZeile), sDataSpalte), sSaveSpalte)
                        Else
                            Call errclErrorCollector.add("Spalte fuer Datenauslese aus Ext. Datei """ & sDataSpalte & """ kann nicht erkannt werden." & vbCrLf & " ACD Daten Tabelle pruefen!.")
                        End If
                    Else
                        Call errclErrorCollector.add("Spalte fuer Speichern in Bericht: """ & sSaveSpalte & """ kann nicht erkannt werden." & vbCrLf & " ACD Daten Tabelle pruefen!.")
                    End If
                End If
            Next lAllDataIterSpalte
        Next lAllDataIter
        
        
        Call colMitarbeiterDaten.setValue(mtrbBlattDaten, mtrbBlattDaten.sName)
        
nexteZeileAusMuster:
    Next i
    
    
    
    getACD = True
    Exit Function
ErrorHandler:
On Error Resume Next
    Dim param1
    Dim param2
    param1 = wkbAcdExport.Name
    param2 = acdMuster
    
    Call errclErrorCollector.add("Fehler bei getACD function aufgetreten" & vbCrLf & "Parameter: " & param1 & ", " & param2)
    
End Function

'Daten von Cobra Export laden

Public Function getCobraData(wkbCobraExport As Workbook, cobraMuster As Variant)
    
    getCobraData = False
    If IsEmpty(aMitarbeiter) Then aMitarbeiter = getMitarbeiterVonCobraExport(wkbCobraExport)
    Dim wksWorkingSheet As Worksheet
    Dim aAllData As Variant
    Dim aRetArray As Variant
    Dim sCurrentAgent As Variant
    Dim i As Integer
    Dim j As Integer
    Dim iSpalte As Integer
    Dim lBstStart As Long
    Dim lBstEnd As Long
    
    
    lBstStart = CLng(f.LetterToCol(mycolEinstellungen.getValue("BestellungenSpalteVon")))
    lBstEnd = CLng(f.LetterToCol(mycolEinstellungen.getValue("BestellungenSpalteBis")))
    
    'worksheet bestimmen
    If Not f.checkIfWksExists(wkbCobraExport, Format(dDate, "yymmdd")) Then
        Call errclErrorCollector.add("Tabelle " & Format(dDate, "yymmdd") & " wurde in """ & wkbCobraExport.FullName & """ nicht gefunden!")
        Exit Function
    End If
    Set wksWorkingSheet = wkbCobraExport.Worksheets(Format(dDate, "yymmdd"))
    
    
    aAllData = wksWorkingSheet.UsedRange.value2
    
    'fuer jeden mitarbeiter
    For Each sCurrentAgent In aMitarbeiter
    
    On Error GoTo nextAgent
        Dim mtrbClassInstance As clsMitarbeiterDaten
        Set mtrbClassInstance = New clsMitarbeiterDaten
        mtrbClassInstance.init (sCurrentAgent)
        
        Dim mtrbBSDATEN As clsMitarbeiterDaten
        Set mtrbBSDATEN = New clsMitarbeiterDaten
        mtrbBSDATEN.init (mtrbClassInstance.sName) 'geschachtelte klasse mit dem gleichen Namen speichern
        Call mtrbClassInstance.add(mtrbBSDATEN, mycolEinstellungen.getValue("BSInhalteBezeichnung"))
        
        
        'fuer jeden datensatz
        For i = LBound(aAllData) To UBound(aAllData)
        On Error GoTo nextDatenSatz:
            'falls richtiger agent, daten sammeln
            If UCase(aAllData(i, 1)) = UCase(mtrbClassInstance.sName) Then
            
                'fuer jedes element in muster
                For j = LBound(cobraMuster) To UBound(cobraMuster)
                On Error GoTo nechsteSpalteAusMuster
                
                    Dim sSpalte As String
                    'gesuchte spalte in export
                    sSpalte = cobraMuster(j, 1)
                    'falls spalte nicht leer ist
                    If sSpalte <> "" Then
                        'falls wert in spalte des UsedRanges uebereinstimt mit wert/spalte aus muster
                        Dim vWert As Variant
                        Dim vBerechnungsWert As Variant
                        
                        vWert = CheckArray(aAllData, i, f.LetterToCol(sSpalte))
                        vBerechnungsWert = CheckArray(aAllData, i, f.LetterToCol(CStr(cobraMuster(j, 3))))
                        'falls wert in spalte stimmt, oder "" angegeben dann weiterrechnen
                        If cobraMuster(j, 2) = "" Or LCase(CStr(vWert)) = LCase(CStr(cobraMuster(j, 2))) Then
                        
                            'falls typ = "Aufzaehlen"
                            'dann in mitarbeiter datenset bei entsprechende(von muster genommene) spalte addieren
                            If LCase(cobraMuster(j, 5)) = LCase("Aufzaehlen") Then
                                    Call mtrbClassInstance.add(vBerechnungsWert, CStr(cobraMuster(j, 4)), "Aufzaehlen")
                            'falls typ aufsummieren,
                            'dann in mitarbeiter datenset bei entsprechende(von muster genommene) spalte mit dem vorherigen wert aufsummieren
                            ElseIf cobraMuster(j, 5) = "Aufsummieren" Then
                                If f.hatWert(vBerechnungsWert) Then
                                    Call mtrbClassInstance.add(vBerechnungsWert, CStr(cobraMuster(j, 4)), "Aufsummieren")
                                End If
                            End If
                        End If
                    End If
nechsteSpalteAusMuster:
                'moegliche fehler aufsammeln
                Call errclErrorCollector.checkForError
                Next j
                On Error GoTo nextDatenSatz:
                'alle verkaufte artikel aufzaehlen/aufsummieren, speziel speichern als "!BSDATEN" also mitarbeiterklass in mitarbeiterklass speichern
                'fuer jede spalte
                For iSpalte = lBstStart To lBstEnd
                    vWert = CheckArray(aAllData, i, iSpalte)
                    'falls irgnedwas verkauft
                    If f.hatWert(vWert) And IsNumeric(vWert) Then
                        'speichern under beschreibungNamen (BSM174 z.B.) in "!BSDATEN"
                        Call mtrbClassInstance.getValue(mycolEinstellungen.getValue("BSInhalteBezeichnung")).add(CLng(vWert), CStr(CheckArray(aAllData, 1, iSpalte)), "Aufsummieren")
                    End If
                Next iSpalte
                
                'if "falls richtiger agent" ende
            End If
            
nextDatenSatz:
        'moegliche fehler aufsammeln
        Call errclErrorCollector.checkForError
        Next i
        On Error GoTo nextAgent
        Call colMitarbeiterDaten.add(mtrbClassInstance, mtrbClassInstance.sName)

        
nextAgent:
    'moegliche fehler aufsammeln
    Call errclErrorCollector.checkForError
    Next sCurrentAgent
    
    
    getCobraData = True
    
    
End Function

'prueft ob zugriff auf 2d array moeglich ist
Private Function CheckArray(ByRef aArray2d, iRow, iCol)
    CheckArray = ""
    
    If Not IsArray(aArray2d) Then Exit Function
    
    If UBound(aArray2d) < iRow Then
        Call errclErrorCollector.add("Zeile " & iRow & "in Cobra export nicht vorhanden")
        Exit Function
    End If
    
    If UBound(aArray2d, 2) < iCol Then
        Call errclErrorCollector.add("Spalte " & f.ColToLetter(iCol) & "in Cobra export nicht vorhanden")
        Exit Function
    End If
    
    CheckArray = aArray2d(iRow, iCol)
End Function

'returns array of agents got from CobraExport file
Public Function getMitarbeiterVonCobraExport(ByRef wkbCobraExport As Workbook) As Variant
    Dim wksWhereToSearch As Worksheet
    Dim sSheetName As String
    sSheetName = Format(dDate, "yymmdd")
    If f.checkIfWksExists(wkbCobraExport, sSheetName) Then
        Set wksWhereToSearch = wkbCobraExport.Sheets(sSheetName)
    Else
        Set wksWhereToSearch = wkbCobraExport.Sheets(wkbCobraExport.Sheets.Count)
    End If
    
    Dim retArray As Variant
    retArray = f.arrayRemoveDublicates(f.getArrayOfColumn(1, 2, , wksWhereToSearch))
    If f.isInArray(retArray, "", True) Then
        Call errclErrorCollector.add("""Stichwort"" in Cobra-Datei:" & vbCrLf & wkbCobraExport.FullName & vbCrLf & "fehlt!")
        Call f.arrayRemoveElement(retArray, "")
    End If
    
    getMitarbeiterVonCobraExport = retArray
    
End Function


'gesammelte daten in wkbTargetReporting datei setzen
Public Function setData(wkbTargetReporting As Workbook)
    Dim wksTargetTableBS As Worksheet
    Dim wksTargetTableDefault As Worksheet
    Dim sZielTabelleBsName As String
    Dim sZielTabelle As String
    Dim MA As clsMitarbeiterDaten
    Dim sBlatt As Variant
    Dim mtrbBlatt As clsMitarbeiterDaten
    
    'Call f.blattschutz(False, "2012", wkbTargetReporting)
        
    For Each sBlatt In colMitarbeiterDaten.getKeys
        Set mtrbBlatt = colMitarbeiterDaten.getValue(sBlatt)
        'blattname = mtrbBlatt.sName
        'find blatt, if exists .....
        'set values NOT "!"
        'dann extra fuer BS, if "!BSDATA" exists, dann check if blatt exists, dann setBS
        
        'blatt wo die daten hin muessen
        sZielTabelle = Replace(mycolEinstellungen.getValue("ZielTabelleSmpl"), "<MA>", mtrbBlatt.sName)
        
        'falls datenblatt nicht existiert, fehelr und weiter
        If Not f.checkIfWksExists(wkbTargetReporting, sZielTabelle) Then
            Call errclErrorCollector.add("Tabelle " & sZielTabelle & " wurde in """ & wkbTargetReporting.FullName & """ nicht gefunden!")
            
        'sonst versuchen daten zu setzen
        Else
            Set wksTargetTableDefault = wkbTargetReporting.Worksheets(sZielTabelle)
            Call setDataDefault(wksTargetTableDefault, mtrbBlatt)
        End If
        
        
        'versuchen BS daten zu setzen, falls BS eintrag vorhanden ist
        If IsObject(mtrbBlatt.getValue(mycolEinstellungen.getValue("BSInhalteBezeichnung"))) Then
            Dim bsMitarbeiter As clsMitarbeiterDaten
            Set bsMitarbeiter = mtrbBlatt.getValue(mycolEinstellungen.getValue("BSInhalteBezeichnung"))
            
            'zieltabellen name, <ma> bs tabelle
            sZielTabelleBsName = Replace(mycolEinstellungen.getValue("ZielTabelleBS"), "<MA>", bsMitarbeiter.sName)
            
            'falls tabelle nicht existiert, feher, und weiter
            If Not f.checkIfWksExists(wkbTargetReporting, sZielTabelleBsName) Then
                Call errclErrorCollector.add("Tabelle " & sZielTabelleBsName & " wurde in """ & wkbTargetReporting.FullName & """ nicht gefunden!")
                
            'sonst versuchen daten zu setzen
            Else
                Set wksTargetTableBS = wkbTargetReporting.Worksheets(sZielTabelleBsName)
                Call setDataBs(wksTargetTableBS, bsMitarbeiter)
            End If
        End If
        
    Next sBlatt
    
    'Call f.blattschutz(True, "2012", wkbTargetReporting)
    
End Function

'BS data setzen
Private Function setDataBs(wksTargetReporting As Worksheet, MA As clsMitarbeiterDaten)
    Dim arrayOfCols As Variant
    Dim arrayOfArtikel As Variant
    Dim artikelBuchstabe As String
    Dim iCol As Integer
    Dim i As Integer
    Dim zeile As Long
    Dim lValue As Long
    
    arrayOfArtikel = f.getArrayOfColumn(mycolEinstellungen.getValue("ArtikelNrSpalte"), 1, , wksTargetReporting)
    arrayOfCols = f.getArrayOfRow(1, 1, , wksTargetReporting)
    artikelBuchstabe = mycolEinstellungen.getValue("ArtikelBuchstaben") 'das was am anfang bei art.bezeichnung steht. z.B. "BSM"
    
    'find where to set
    For i = 1 To UBound(arrayOfCols)
        If f.isDateMy(arrayOfCols(i)) Then
            If arrayOfCols(i) = dDate Then
                iCol = i
                Exit For
            End If
        End If
    Next i
    
    'alternative: not that stable
    'iCol = f.getIndexInArrayByValue(CLng(dDate), arrayOfCols)
    
    'falls column gefunden
    If iCol > 0 Then
        'fuer jeden artikel, ab startRow in einstellungen
        For i = mycolEinstellungen.getValue("StartRowBS") To UBound(arrayOfArtikel)
            'falls der Artikel in mitarbeiterdaten vorhanden ist, also irgendwas davon verkauft
            lValue = MA.getValue(artikelBuchstabe & arrayOfArtikel(i))
            If lValue > 0 Then
                'in entsprechende Zelle den Wert einsetzen
                If wksTargetReporting.Cells(i, iCol).Locked Then
                    Debug.Print "schreibgeschutzt"
                    Debug.Print wksTargetReporting.Name
                    Debug.Print wksTargetReporting.Cells(i, iCol).Address
                Else
                    wksTargetReporting.Cells(i, iCol).value = lValue
                End If
            'falls nichts davon verkauft, LEER setzen
            Else
                If wksTargetReporting.Cells(i, iCol).Locked Then
                    Debug.Print "schreibgeschutzt"
                    Debug.Print wksTargetReporting.Name
                    Debug.Print wksTargetReporting.Cells(i, iCol).Address
                Else
                    wksTargetReporting.Cells(i, iCol).value = ""
                End If

            End If
        Next i
    Else
        Call errclErrorCollector.add("Kann in Tabelle " & wksTargetReporting.Name & _
        " keine Spalte um Artikeln hinzuzufuegen finden (Gesucht nach Datum: " & CStr(dDate) & ")" & vbCrLf & _
        "In datei: " & wksTargetReporting.Parent.FullName)
    End If
    
    
End Function

'daten setzen, default
Private Function setDataDefault(wksTargetReporting As Worksheet, MA As clsMitarbeiterDaten)
    Dim lCurrentRow As Long
'    Dim datumsArray As Variant
'    Dim i As Long
'
'
'    datumsArray = f.getArrayOfColumn(1, 1, , wksTargetReporting)
'    'zeile mit dem richtigen Datum finden
'    For i = 1 To UBound(datumsArray)
'        'falls numeric und length = 5 (datum format in zellen) oder einfach IsDate
'        If f.isDateMy(datumsArray(i)) Then
'            'falls datums uebereinstimmen
'            If CDate(datumsArray(i)) = dDate Then
'                lCurrentRow = i
'                Exit For
'            End If
'        End If
'    Next i
    
    lCurrentRow = findRowByDataInReporting(wksTargetReporting)
    
    'falls nicht gefunden exit
    If lCurrentRow = 0 Then Exit Function
    
    Dim key As Variant
    Dim keys As Variant
    Dim vWert As Variant
    keys = MA.getKeys()
    
    For Each key In keys
        'ausrufezeichen ignorieren, da keine blaetter
        If Left(key, 1) <> "!" Then
            vWert = MA.getValue(key)
            'falls wert alphanumerisch ist
            If f.IsAlphaNumeric(vWert) Then
                'falls key als column betrachtet werden kann
                If IsNumeric(f.LetterToCol(key)) Then
                    'in bestimmte zelle den wert setzen, falls nicht schreibgeschuetzt
                    With wksTargetReporting.Cells(lCurrentRow, key)
                        If Not .Locked Then
                            .value = vWert
                        Else
                            Debug.Print "schreibgeschutzt" & .Address & " in " & wksTargetReporting.Name
                        End If
                    End With
                End If
            End If
        End If
    Next key
    
End Function

Public Function findRowByDataInReporting(wksTargetReporting As Worksheet)
    Dim lCurrentRow As Long
    Dim datumsArray As Variant
    Dim i As Long
    
    
    datumsArray = f.getArrayOfColumn(1, 1, , wksTargetReporting)
    'zeile mit dem richtigen Datum finden
    For i = 1 To UBound(datumsArray)
        'falls numeric und length = 5 (datum format in zellen) oder einfach IsDate
        If f.isDateMy(datumsArray(i)) Then
            'falls datums uebereinstimmen
            If CDate(datumsArray(i)) = dDate Then
                lCurrentRow = i
                Exit For
            End If
        End If
    Next i
    
    
    findRowByDataInReporting = lCurrentRow
End Function

'modify FormatS in wkbTargetReporting, by Settings array
Public Function modifyFormats(wkbTargetReporting, aSettings)
    Dim sWks As String
    Dim sSpalte As String
    Dim sFormat As String
    Dim i As Integer
    
    For i = 1 To UBound(aSettings)
        sWks = aSettings(i, 4)
        sSpalte = aSettings(i, 2)
        sFormat = aSettings(i, 1)
        If f.hatWertArray(Array(sWks, sSpalte, sFormat)) Then
            Call modifyFormat(wkbTargetReporting.Worksheets(sWks), sSpalte, sFormat)
        End If
    Next i
End Function


'modify format in wksTargetReporting with sFormat at column sSpalte
Public Function modifyFormat(wksTargetReporting As Worksheet, sSpalte As String, sFormat As String)
    Dim lCurrentRow As Long
    lCurrentRow = findRowByDataInReporting(wksTargetReporting)
    With wksTargetReporting.Range(sSpalte & lCurrentRow)
        .value = f.FormatSeconds(.value, sFormat)
    End With
End Function


'zeigt ergebnisse, fehler die errorCollector gesammelt hat
Public Function showResult()
    If errclErrorCollector.hasError Then
        MsgBox "Folgende Fehler sind aufgetretten: " & vbCrLf & errclErrorCollector.getErrors
        Debug.Print errclErrorCollector.getErrors
    Else
        MsgBox "Erfolgreich beendet"
    End If
End Function


