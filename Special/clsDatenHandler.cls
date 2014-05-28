VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatenHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private gesamtLength As Integer
Private wbWorkbook As Workbook
'Private wsSaveTabelle As Worksheet
'Private sStartColBySave As String
'Private col_range_Places As New Collection
'Private ig_lStartRow As Integer
'Private mMuster As clsMuster
Private bInitDone As Boolean
Private ig_iMaxSlots As Integer
Private f As clsFuncs


Public Function Import(dsDatenSet As clsDatenSet, ByRef wsSaveTabelle As Worksheet, sStartColBySave As Variant, Optional iStelle As Integer = 1, Optional ig_lStartRow As Variant)
    Dim tempMuster As clsMuster
    Set tempMuster = dsDatenSet.usedMuster
    If tempMuster Is Nothing Then Exit Function
    Dim rangeWhereToCopy As Range
    Dim rangeFromCopy As Range
    Dim daten As Collection
    Dim col_range_Places As Collection
    Dim sSpalte As Variant
    Dim toRow As Integer
    Dim f As New clsFuncs
    sStartColBySave = fXl.ColToLetter(sStartColBySave)
    
    If IsMissing(ig_lStartRow) Then
        ig_lStartRow = dsDatenSet.g_lStartRow
    End If
    If ig_lStartRow = Empty Then ig_lStartRow = 4 'default value in case even dsDatenSet doesnt have the value
    
    Set col_range_Places = setRangesCollection(tempMuster, wsSaveTabelle, CStr(sStartColBySave))
    
    Set daten = dsDatenSet.daten
    Set rangeWhereToCopy = col_range_Places(CStr(iStelle))
    Set rangeFromCopy = tempMuster.rangeOfMuster
    
    'muster kopieren
    rangeWhereToCopy.Clear
    rangeFromCopy.Copy Destination:=rangeWhereToCopy
    'werte einsetzen
    toRow = rangeWhereToCopy.Rows.Count
    'feur jedes freien Wert im Muster, aus dsDatenSet array auslesen, an neue Stelle hinzufuegen
    For Each sSpalte In tempMuster.valuesAsArray
        Dim tmpSaveArray As Variant
        Dim here As Integer
        tmpSaveArray = daten(sSpalte)
        'If fXl.LetterToCol(sSpalte) < 35 Then GoTo nex 'fuer hauptformeln ladevorgang speichern, mb todo
        here = fXl.LetterToCol(sSpalte) + (iStelle) * (rangeWhereToCopy.Columns.Count + 1) + fXl.LetterToCol(sStartColBySave)
        Call fXl.fillColumnByArray(here, tmpSaveArray, ig_lStartRow, toRow, wsSaveTabelle)
'nex:        'fuer hauptformeln ladevorgang speichern, mb todo
    Next sSpalte

End Function

'einen von 12(oder mehr) gespeicherten datenSets loeschen
Public Function removeSet(ByRef wsSaveTabelle As Worksheet, mMuster As clsMuster, iStelle As Variant, _
                    sStartColBySave As Variant) As Boolean
    On Error GoTo errorHandler
    Dim f As New clsFuncs
    Dim col_range_Places As Collection
    
    sStartColBySave = fXl.ColToLetter(sStartColBySave)
    
    Set col_range_Places = setRangesCollection(mMuster, wsSaveTabelle, CStr(sStartColBySave))
    
    col_range_Places(CStr(iStelle)).Clear
    
    removeSet = True
    
    Exit Function
errorHandler:
    removeSet = False
    
End Function
'muster mit formeln neu einfuegen
Public Sub refreshMuster(ByRef wsSaveTabelle As Worksheet, mMuster As clsMuster, sStartColBySave As Variant, Optional g_lStartRow As Variant = 1, Optional aDontSetFormulaHere As Variant)
    'alte place(0) muster loeschen (gesamt daten sammler)
    'If Not removeSet(wsSaveTabelle, mMuster, 0, sStartColBySave) Then Exit Sub
    Dim rangeWhereToCopy As Range, rangeFromCopy As Range
    Dim col_range_Places As Collection
    Dim k As Long
    Dim SpalteWoWirDenArrrayEinfuegen As String, valueFormula As String
    Dim letter As Variant
    Dim formelnArray As Variant
    Dim f As New clsFuncs
    Set col_range_Places = setRangesCollection(mMuster, wsSaveTabelle, CStr(sStartColBySave))
    
    'muster kopieren
    Set rangeWhereToCopy = col_range_Places(CStr(0))
    Set rangeFromCopy = mMuster.rangeOfMuster
    
    rangeWhereToCopy.Clear
    rangeFromCopy.Copy Destination:=rangeWhereToCopy
    
    'formeln setzen
    'fuer jede Spalte die mit formeln ausgefullt werden muss
    For Each letter In mMuster.valuesAsArray
        'falls ausnahmearray angegeben wurde
        If Not IsMissing(aDontSetFormulaHere) Then
            If f.isInArray(aDontSetFormulaHere, letter) Then GoTo nextletter
        End If
        
        formelnArray = Array()
        valueFormula = ""
        'Fuer jede Zelle die die formeln enthalten muss
        For k = g_lStartRow To mMuster.rangeOfMuster.Rows.Count
            valueFormula = getMainFormula(mMuster.rangeOfMuster.Columns.Count + 1, col_range_Places.Count - 1)
            Call f.arrayPush(formelnArray, valueFormula)
        Next k
        
        'array in bestimmte spalte einfuegen
        SpalteWoWirDenArrrayEinfuegen = fXl.LetterToCol(letter) + fXl.LetterToCol(sStartColBySave)
        'werte einfuegen
        Call fXl.fillColumnByArray(SpalteWoWirDenArrrayEinfuegen, formelnArray, g_lStartRow, , wsSaveTabelle, False) 'keine lokale formulas
nextletter:
    Next letter
    
End Sub

Function refreshDruckTabelle(ByRef wsSavePrintTabelle As Worksheet, ByRef wsSaveTabelle As Worksheet, ByRef mMuster As clsMuster, sStartColBySave As String)
    Dim col_range_Places As Collection
    Dim rngCell As Range
    Dim rngFromRng As Range
    Dim colDiff As Long
    Dim rowDiff As Long
        
    mMuster.rangeOfMuster.Copy
    
    With wsSavePrintTabelle
        .UsedRange.Clear
        mMuster.rangeOfMuster.Copy Destination:=.Range("A1")
        Application.CutCopyMode = False

        'ranges set erstellen
        Set col_range_Places = setRangesCollection(mMuster, wsSaveTabelle, CStr(sStartColBySave))
        'gesamt ergebnisse (erstes set) kopieren
        'col_range_Places(CStr(0)).Copy
        '.Paste Destination:=.Range("A1"), Link:=True
        
        'paste link:=true emulieren, wegen speicher begrenzung
        Set rngFromRng = col_range_Places(CStr(0))

        'einfach Link:=True aehnliche formulas setzen
        colDiff = 1 - rngFromRng.Cells(1, 1).column
        rowDiff = 1 - rngFromRng.Cells(1, 1).row
        For Each rngCell In rngFromRng
            .Cells(rngCell.row + rowDiff, rngCell.column + colDiff).Formula = "=" & rngFromRng.Parent.Name & "!" & f.ColToLetter(rngCell.column) & rngCell.row
        Next rngCell

    End With
    Set col_range_Places = Nothing
End Function

'erstellt mega formula (nicht lokale) fuer jede zelle des gesamt DatenSummen
Private Function getMainFormula(setLength, SetsAmount)
    Dim i As Long
    ReDim arrayWithFormulas(SetsAmount - 1) As Variant
    
    'array mit jedem element als formula fuer jeden setwert erstellen
    For i = 1 To SetsAmount
        arrayWithFormulas(i - 1) = "OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),," & setLength & "*" & i & ",1,1)"
    Next i
    
    'mit "+" als Zeichenkette zusammensetzen, und zuruckgeben
    getMainFormula = "=" & Join(arrayWithFormulas, "+")
End Function

Public Sub init(ByRef oWorkbook As Workbook, g_iMaxSlots As Integer)
    If bInitDone Then Exit Sub
    Set f = New clsFuncs
    Set wbWorkbook = oWorkbook
    'Set wsSaveTabelle = oSaveTabelle
    'sStartColBySave = startColBySave
    'Set mMuster = oMuster
    ig_iMaxSlots = g_iMaxSlots
    'ig_lStartRow = g_lStartRow
    
    
    
    bInitDone = True
End Sub

Private Function setRangesCollection(mMuster As clsMuster, wsSaveTabelle As Worksheet, sStartColBySave As String)
Dim i As Integer
    Dim tmpCol As Integer
    Dim musterHight As Integer
    Dim musterWidth As Integer
    Dim col_range_Places As New Collection
    musterHight = mMuster.rangeOfMuster.Rows.Count
    musterWidth = mMuster.rangeOfMuster.Columns.Count
    tmpCol = fXl.LetterToCol(sStartColBySave) + 1
    
    'ranges (slots) in collection laden
    For i = 0 To ig_iMaxSlots
        With wsSaveTabelle
            Dim tmpRange As Range
            
            Set tmpRange = .Range(.Cells(1, tmpCol), .Cells(musterHight, tmpCol + musterWidth - 1))
           ' Debug.Print "range tmp in collection:" & i & "  " & tmpRange.Address
           ' Debug.Print "range tmp in collection:" & i & "  columns:" & tmpRange.Columns.Count
            tmpCol = tmpCol + 1 + musterWidth
           ' Debug.Print "tmpCol:" & fXl.ColToLetter(tmpCol)
        End With
        Call col_range_Places.add(tmpRange, CStr(i))
        Set tmpRange = Nothing
    Next i
    Set setRangesCollection = col_range_Places
End Function
