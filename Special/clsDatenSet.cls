VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatenSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colDaten As Collection
Private mUsedMuster As clsMuster
Private f As clsFuncs
Private lg_lStartRow As Long

Private Sub Class_Initialize()
    Set f = New clsFuncs
    Set colDaten = New Collection
End Sub

Private Sub Class_Terminate()
    Set f = Nothing
    Set colDaten = Nothing
    Set mUsedMuster = Nothing
End Sub

'daten von einer tabelle laden
'oTabelle - woher laden, oMuster nach welchem Muster, g_lStartRow ab welche Zeile, lCutAtBegin wieviele datensaetze abschneiden
'todo add Zellenprüfer als object?
Sub setDatenByTable(ByRef oTabelle As Worksheet, ByRef oMuster As clsMuster, _
                    Optional lStartRow As Variant = 1, Optional lCutAtBegin As Variant = -1)
    Dim column() As Variant
    Dim sSpaltenName As Variant
    Dim lEndRow As Long
    'zur sicherheit colDaten leeren
    Set colDaten = New Collection
    lg_lStartRow = lStartRow
    lEndRow = oMuster.rangeOfMuster.Rows.Count
    
    'nur fuer richtiges versetzen
    Dim aZwischenWerteNichtBerechnen As Variant
    aZwischenWerteNichtBerechnen = Array("B", "C")
    
    For Each sSpaltenName In oMuster.valuesAsArray
        column = fXl.getArrayOfColumn(sSpaltenName, lg_lStartRow, lEndRow, oTabelle)
        
        'falls einige saetze ausgeschnitten werden muessen, bei gesamt und gesamtPA ist der fall.
        
        If lCutAtBegin > 0 And Not f.isInArray(aZwischenWerteNichtBerechnen, sSpaltenName) Then
            Call cutXelems(column, lCutAtBegin)
        End If
        
        Call colDaten.add(column, sSpaltenName)
    Next sSpaltenName
    
    Set mUsedMuster = oMuster
End Sub

'HARD CODED spezialfall
'nur woechentlich daten lesen, woechentliche zwischenergebnisse pro woche einfuegen und gesamt berechnen und einfuegen
'todo? vllt gesamt gar nicht laden, und als formel (summe) von muster laden?
Private Sub cutXelems(ByRef aArray, iElems)
    'wir haben 7 wochen
    Dim i As Integer, j As Integer
    Dim dSummeW As Double, dSummeG As Double
    Dim tmpArray As Variant
    Dim vWert As Variant



    
    'alle wirkliche werte sammeln, also wochenergebnisse und rest rausfiltern
    
    For i = 1 + iElems To UBound(aArray)
        'mod 7 wochentlich filtern, <49 rest filtern
        If i Mod 7 <> 0 And i <= 49 Then
            'falls in neuem Array 7. element kommt (woechentliche ergebniss muss hier rein), den mit 0 fuellen,
            'und erst in naechste stelle das richtige wert einfuegen
            If (f.aLength(tmpArray) + 1) Mod 7 = 0 And f.aLength(tmpArray) > -1 Then
                Call f.arrayPush(tmpArray, 0)
            End If
            Call f.arrayPush(tmpArray, aArray(i))
        End If
    Next i
    
    'array redimensionieren, mit 1 muss er anfangen
    Call f.arrayChangeBounds(tmpArray, 1)
    
    'ausgeschnittene fehlende elements mit 0 nachfuellen. Am ende des arrays hinzufuegen, da summe auch 0 ist, und nachtraeglich berechnet wird
    For i = UBound(tmpArray) + 1 To UBound(aArray)
        Call f.arrayPush(tmpArray, 0)
    Next i
    
    
    'ergebniss zuweisen
    aArray = tmpArray
    
    'aArray hat nun nur die tagesergebnisse, jetzt woechentliche und gesamt berechnen
    'fuer jede woche..
    For i = 1 To 7
        'woechentliche summe berechnen, und gleich gesamtsumme berechnen
        dSummeW = 0
        For j = 1 To 6
            vWert = aArray(j + (i - 1) * 7)
            If IsNumeric(vWert) Then
                dSummeW = dSummeW + CDbl(vWert)
            End If
        Next j
        
        tmpArray(7 * i) = dSummeW
        
        dSummeG = dSummeG + dSummeW
    Next i
    'gesamt einsetzen am ende
    tmpArray(UBound(tmpArray)) = dSummeG
    
    'dem aArray zuweisen
    aArray = tmpArray
'
'    Dim i As Long
'    aArray = f.arrayGetSubArray(aArray, iElems, , 1)
'    For i = 1 To iElems
'        Call InsertElementIntoArray(aArray, UBound(aArray) - 1, 0)
'    Next i
End Sub


Sub setDatenByTableBS(ByRef oTabelle As Worksheet, ByRef oMuster As clsMuster, _
                    Optional g_lStartRow As Variant = 1, Optional columnToCampare As Variant = 1, Optional tmpWksL As Worksheet)
    
    Dim newTmpTabelle As Worksheet
    'tabelle mit dem muster ausgleichen, temporaere tabelle benutzen
    Set newTmpTabelle = ausgleichen(oMuster.tabelleOfMuster, oTabelle, columnToCampare, tmpWksL)
    'wie gewoehnlich daten speichern
    Call setDatenByTable(newTmpTabelle, oMuster, g_lStartRow)
    
    'temporaere tabelle leeren
    newTmpTabelle.UsedRange.Clear
End Sub

'andere tabelleToChange in tabelleOrig-Format anpassen
Private Function ausgleichen(ByRef tabelleOrig As Worksheet, ByRef tabelleToChange As Worksheet, _
                                Optional byWhatColumn As Variant = 1, Optional tmpWksL As Worksheet) As Worksheet
    Dim tmpWks As Worksheet
    On Error GoTo errorHandler:
    'falls keine temp tabelle angegeben ist, eine erstellen
    'Vorsicht, danach muss man irgendwann die Tabelle loeschen...
    If tmpWksL Is Nothing Then
        Set tmpWks = tabelleOrig.Parent.Worksheets.add(After:=ActiveWorkbook.Worksheets(Sheets.Count))
        'tmpWks.Visible = False
    Else
        Set tmpWks = tmpWksL
    End If
    'column nummerisch auf jeden fall haben
    byWhatColumn = fXl.LetterToCol(byWhatColumn)
        
    'in temp tabelle kopieren
    tabelleToChange.UsedRange.Copy Destination:=tmpWks.Cells(1, 1)
        
    'eigentlicher ausgleich
    Dim lastrow As Long, firstRow As Long, i As Long
    Dim origArray As Variant, currentTmpWert As Variant
    'firstRow = f.getFirstRow(tabelleOrig, byWhatColumn)
    firstRow = tabelleOrig.UsedRange.row 'hard, todo?
    lastrow = fXl.getLastRow(tabelleOrig, byWhatColumn)
    
    'array mit werten vorbereiten (performance)
    With tabelleOrig
        origArray = f.array2DimTo1Dim(.Range(.Cells(firstRow, byWhatColumn), .Cells(lastrow, byWhatColumn)).value2)
    End With
    For i = firstRow To lastrow
        currentTmpWert = tmpWks.Cells(i, byWhatColumn)
        'falls die werte bei dem Row nicht gleich sind
        If origArray(i) <> currentTmpWert Then
            'falls der fehlerhafte Eintrag in einer tabelleOrig werten drin ist -> etwas fehlt
            If f.isInArray(origArray, currentTmpWert) Then
                'zeile einfuegen
                tmpWks.Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                'mit dem wert fuellen
                tmpWks.Cells(i, byWhatColumn).Value = origArray(i)
            'sonnst etwas zu viel, rausloeschen, noch mal mit dem wert aus origArray (gleiche i) probieren
            Else
                tmpWks.Rows(i).delete
                i = i - 1
            End If
            
        End If
    Next i
        
    'tempTabelle zuruckgeben
    Set ausgleichen = tmpWks
        
        
    Exit Function
errorHandler:
'falls temp tabelle angelegt ist, die löschen
    If tmpWksL Is Nothing Then
        If Not tmpWks Is Nothing Then
            Application.DisplayAlerts = False
            tmpWks.delete
            Application.DisplayAlerts = True
        End If
    End If
End Function


Public Property Get daten() As Collection
    Set daten = colDaten
End Property

Public Property Get usedMuster() As clsMuster
    Set usedMuster = mUsedMuster
End Property

Public Property Get g_lStartRow() As Long
    g_lStartRow = lg_lStartRow
End Property






