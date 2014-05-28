VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuncsExl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFuncsExl
' Author    : Dmitry Gorelenkov
' Date      : 05.2012
' Changed   : 02.12.2013
' Purpose   : Excel function sammlung
' Requires  : Excel Objects Loaded, clsFuncs, clsMyCollection
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------

Option Explicit

Private f As clsFuncs

' Sekunden nach hh:mm:ss umrechnen
Public Function FormatSeconds(ByVal nSeconds As Long, _
     Optional ByVal sFormat As String = "hh:nn:ss") As String
    
     ' Falls anstelle nn mm im Ausgabeformat angegeben wurde
     sFormat = Replace(sFormat, "mm", "nn")
     
    
     FormatSeconds = Format$(DateAdd("s", nSeconds, CDate("00:00:00")), sFormat)
End Function


'returns element from aElements with hightest priority.
'aPrioElements elements with aPriority priority mark.
'if aPriority is nothing, than last element of array has hightest prio, first lowest
Public Function getProirityElement(ByVal aElements As Variant, ByVal aPrioElements As Variant, Optional ByVal aPriority As Variant)
    Dim element As Variant
    Dim lIndex As Long, lBestPrioIndex As Long
    
    If IsEmpty(aElements) Or Not IsArray(aElements) Then Exit Function
    If IsEmpty(aPrioElements) Or Not IsArray(aPrioElements) Then Exit Function
    
    
    'falls priority array nicht uebergeben wurde, ein erstellen mit werten 0,1,2,3,...
    If IsMissing(aPriority) Then
        ReDim aPriority(LBound(aPrioElements) To UBound(aPrioElements))
        Dim i As Long
        For i = 0 To UBound(aPriority)
            aPriority(i) = i
        Next i
    End If
    
    'priority array muss mindestens gleiche groesse haben
    If f.aLength(aPriority) < f.aLength(aPrioElements) Then Exit Function
        
    lBestPrioIndex = -1
    'elements travisieren, hoechste prio finden
    For Each element In aElements
        lIndex = f.getIndexInArrayByValue(element, aPrioElements)
        If lIndex = -1 Then GoTo nextElem
        If lBestPrioIndex > -1 Then
            If aPriority(lBestPrioIndex) < aPriority(lIndex) Then
                lBestPrioIndex = lIndex
            End If
        Else 'erste schrumpf
            lBestPrioIndex = lIndex
        End If
nextElem:
    Next element
    
    
    If lBestPrioIndex = -1 Then
        getProirityElement = ""
    Else
        getProirityElement = aPrioElements(lBestPrioIndex)
    End If
End Function

Public Function getDirAndFile(pfad As String)
    Dim tmp, file, dir
    tmp = Split(pfad, "\")
    file = tmp(UBound(tmp))
    tmp(UBound(tmp)) = ""
    dir = Join(tmp, "\")
    getDirAndFile = Array(dir, file)
End Function

'paste an 2d Array to specified range (1 cell)
Public Function arrayPasteToRange(aValues As Variant, rngPasteRange As Range, Optional bSetHochkomma As Boolean = False)
    On Error GoTo Err:
    If IsArray(aValues) Then
        If UBound(aValues) - LBound(aValues) > 0 Then
            If rngPasteRange.count = 1 Then
                Dim offsetRows As Long
                Dim offsetCols As Integer
                Dim i As Long
                Dim i2 As Long
                Dim lbound1 As Long
                Dim ubound1 As Long
                Dim lbound2 As Long
                Dim ubound2 As Long
                offsetRows = UBound(aValues) - LBound(aValues)
                offsetCols = UBound(aValues, 2) - LBound(aValues, 2)
                lbound1 = LBound(aValues, 1)
                ubound1 = UBound(aValues, 1)
                lbound2 = LBound(aValues, 2)
                ubound2 = UBound(aValues, 2)
                
                If bSetHochkomma Then
                    For i = lbound1 To ubound1
                        For i2 = lbound2 To ubound2
                            If VarType(aValues(i, i2)) = vbString Then
                                aValues(i, i2) = "'" & aValues(i, i2)
                            End If
                        Next i2
                    Next i
                End If
                rngPasteRange.Parent.Range(rngPasteRange, rngPasteRange.Offset(offsetRows, offsetCols)).value = aValues
                arrayPasteToRange = True
                Exit Function
            End If
        End If
    End If
    
Err:
    arrayPasteToRange = False
End Function

'Public Function f.array2DimTo1Dim(arr)
'    Dim tmpArray()
'    ReDim tmpArray(LBound(arr) To UBound(arr))
'    Dim i As Long
'    For i = LBound(arr) To UBound(arr)
'        tmpArray(i) = arr(i, 1)
'    Next i
'f.array2DimTo1Dim = tmpArray
'End Function

''push each element in X-dimensional array or collection to new 1-dim array
'Public Function f.arrayXDimTo1Dim(ByVal arr As Variant)
'    Dim tmpArray()
'    Dim element As Variant
'    tmpArray = Array()
'
'    For Each element In arr
'        Call f.arrayPush(tmpArray, element)
'    Next element
'
'    f.arrayXDimTo1Dim = tmpArray
'End Function

''return a copy of array
'Public Function arrayCopy(ByVal whatToCopy As Variant)
'    If Not IsArray(whatToCopy) Then Exit Function
'    Set arrayCopy = whatToCopy
'End Function

''return a  copy of object
'Public Function objCopy(ByVal whatToCopy As Object)
'    If Not IsObject(whatToCopy) Then Exit Function
'    Set objCopy = whatToCopy
'End Function

'prueft ob sSheetName worksheet in wkbBook vorhanden ist
Public Function checkIfWksExists(wkbBook As Workbook, sSheetName As String) As Boolean
    Dim sheet As Variant
    checkIfWksExists = False
    For Each sheet In wkbBook.Sheets
        If UCase(sheet.name) = UCase(sSheetName) Then
            checkIfWksExists = True
            Exit Function
        End If
    Next sheet
End Function

'letter to column from clsFuncs
Public Function LetterToCol(ByVal letter As String)
    LetterToCol = f.LetterToCol(letter)
End Function
'column to letter from clsFuncs
Public Function ColToLetter(ByVal col As Variant)
    ColToLetter = f.ColToLetter(col)
End Function

''prueft ob der Wert datum (als text) oder numerische Datum ist
'Public Function isDateMy(vWert)
'    On Error Resume Next
'    isDateMy = (IsNumeric(vWert) And Len(vWert) = 5) Or IsDate(vWert)
'End Function

''wandelt eine Buchstabe z.B. "B" in eine SpaltenNr also z.B. "2" um
'Public Function f.LetterToCol(ByVal letter As String) As Variant
'    If IsNumeric(letter) Then
'        LetterToCol = CInt(letter)
'        Exit Function
'    End If
'
'    Dim num As Integer
'    Dim letnum As Integer
'    Dim i As Integer
'    Dim currentLetter As String
'    letnum = Len(letter)
'
'    For i = 1 To letnum
'        currentLetter = Mid(letter, i, 1)
'        'falls nummer vorkommt, das input zuruckgeben
'        If Not IsAlphaBetical(currentLetter) Or IsNumeric(currentLetter) Then
'            LetterToCol = letter
'            Exit Function
'        End If
'        num = num * 26 + Asc(UCase(currentLetter)) - 64
'    Next i
'    LetterToCol = num
'End Function

''wandelt SpaltenNr in eine Buchstabe um
'Public Function f.ColToLetter(ByVal col As Variant) As String
'    If Not IsNumeric(col) Then
'        f.ColToLetter = CStr(col)
'        Exit Function
'    End If
'
'    Dim columnString
'    Dim currentLetterNumber
'        While (col > 0)
'            currentLetterNumber = (col - 1) Mod 26
'            columnString = Chr(currentLetterNumber + 65) & columnString
'            col = (col - (currentLetterNumber + 1)) / 26
'        Wend
'    f.ColToLetter = columnString
'End Function

'gibt array von werten zurueck die in eine Spalte "column" (nummer oder buchstabe) sich befinden.
'Ab zeile "fromFow" bis "toRow" in eine "sheet" tabelle
Public Function getArrayOfColumn(ByVal column As Variant, Optional fromRow As Variant = -1, Optional toRow As Variant = -1, Optional sheet As Worksheet)
    Dim aArray
    If sheet Is Nothing Then
        Set sheet = ActiveSheet
    End If
    
    If fromRow = -1 Then fromRow = getFirstRow(sheet, f.LetterToCol(column)) 'sheet.UsedRange.row
    If toRow = -1 Then toRow = getLastRow(sheet, f.LetterToCol(column)) 'sheet.UsedRange.Rows.Count
    
    If IsNumeric(column) Then column = f.ColToLetter(column)
    
    If fromRow > toRow Then toRow = fromRow
    
    'falls erste und leer, dann leere array uebergeben
    If fromRow = toRow And sheet.Cells(fromRow, f.LetterToCol(column)).value = vbNullString Then
        getArrayOfColumn = Array()
    Else
        aArray = sheet.Range(column & fromRow & ":" & column & toRow).value2
        If Not IsArray(aArray) Then
            Dim tmpret(1 To 1, 1 To 1)
            tmpret(1, 1) = aArray
            aArray = tmpret
        End If
        getArrayOfColumn = f.array2DimTo1Dim(aArray)
    End If
    
    'getArrayOfColumn = f.array2DimTo1Dim(aArray)
End Function


'gibt array von werten zurueck die in eine Zeile "row" sich befinden.
'Ab spalte "fromCol" bis "toCol" in eine "sheet" tabelle
Public Function getArrayOfRow(ByVal row As Variant, Optional fromCol As Variant = -1, Optional toCol As Variant = -1, Optional sheet As Worksheet)
    Dim aArray
    If sheet Is Nothing Then
        Set sheet = ActiveSheet
    End If
    
    
    fromCol = f.LetterToCol(fromCol)
    toCol = f.LetterToCol(toCol)
    
    'row und cols muessen nun numerisch sein
    If Not (IsNumeric(row) And IsNumeric(fromCol) And IsNumeric(toCol)) Then Exit Function
    
    If fromCol = -1 Then fromCol = getFirstCol(sheet)
    If toCol = -1 Then toCol = getLastCol(sheet)
    With sheet
        aArray = .Range(.Cells(row, fromCol), .Cells(row, toCol)).value2
    End With
    
    getArrayOfRow = f.arrayChangeBounds(f.arrayXDimTo1Dim(aArray), 1)
End Function

'fuellt ein Spalte mit werten von dem aArray, ab fromRow bis toRow
Public Function fillColumnByArray(ByVal column As Variant, aArray As Variant, Optional fromRow As Variant = 1, Optional toRow As Variant = -1, Optional sheet As Worksheet, Optional frmlaLocal = True)
    'Dim aArrayOut
    If toRow = -1 Then toRow = fromRow + UBound(aArray) - LBound(aArray)
    If sheet Is Nothing Then
        Set sheet = ActiveSheet
    End If
    If IsNumeric(column) Then column = f.ColToLetter(column)
    
    With sheet
    If frmlaLocal Then
        .Range(column & fromRow & ":" & column & toRow).FormulaLocal = WorksheetFunction.Transpose(aArray)
    Else
        '.range(column & fromRow & ":" & column & toRow).FormulaArray = WorksheetFunction.Transpose(aArray)
        'maximum 255 zeichen bei FromulaArray .. deswegen ..
        Dim i As Long
        For i = fromRow To toRow
            .Cells(i, column).Formula = aArray(i - fromRow)
        Next i
    End If
    End With
    
End Function

''topX elemente von einem Array
'Public Function getTopX(ByVal aValues As Variant, topWhat)
'    Dim aRet()
'    Dim iBegin As Integer
'    ReDim aRet(UBound(aValues))
'    Dim i As Integer
'    Call QuickSort(aValues, LBound(aValues), UBound(aValues))
'    iBegin = UBound(aValues) - topWhat + 1
'    For i = iBegin To UBound(aValues)
'        aRet(i - iBegin) = aValues(i)
'    Next i
'
'    ReDim Preserve aRet(topWhat - 1)
'
'    getTopX = aRet
'End Function

'vergleicht werte beides arrays, return 2te werte des assoziativen, falls 1. wert gleich ist.
'z.b. Array(2,3) und Array({2,"yea"}, {1, "zz"}, {3, "fff3"}) als parameter, wird zuruckgegeben Array("yea","fff3")
'falls nicht gefunden, wird "" als wert gesetzt
'Public Function getAssotiations(arrayKeys As Variant, cCollectionOfAssotiations As Collection)
'    Dim i As Integer
'    Dim retArray, test
'    ReDim retArray(LBound(arrayKeys) To UBound(arrayKeys))
'    For i = LBound(arrayKeys) To UBound(arrayKeys)
'        On Error GoTo nextOne
'            test = cCollectionOfAssotiations(CStr(arrayKeys(i)))
'            retArray(i) = test
'nextOne:
'    Next i
'    getAssotiations = retArray
'End Function

'gibt collection mit assotiationen, value - key von arrays entsprechend zuruck
'Public Function makeAssotiativeCollection(arrayValues, arrayKeys)
'    Dim cReturn As New Collection
'    Dim i As Integer
'
'    On Error Resume Next
'    For i = LBound(arrayKeys) To UBound(arrayKeys)
'        Call cReturn.add(arrayValues(i), CStr(arrayKeys(i)))
'    Next i
'
'    Set makeAssotiativeCollection = cReturn
'End Function

'no need? :(
Public Function modifyValuesToPercentString(ByVal aArray)
    Dim i As Integer
    Dim newArr() As String
    ReDim newArr(LBound(aArray) To UBound(aArray))
    For i = LBound(aArray) To UBound(aArray)
        newArr(i) = CStr(Format(CStr(aArray(i)), "0.0%"))
    Next i
    
    modifyValuesToPercentString = newArr
End Function



'Public Function reverse2dArrayValueKey(ByRef aArray As Variant)
'    Dim tmpArray
'    Dim iNewArray As Integer
'    Dim i As Integer
'    ReDim tmpArray(LBound(aArray) To UBound(aArray), 1)
'
'    iNewArray = LBound(aArray)
'    For i = UBound(aArray) To LBound(aArray) Step -1
'        tmpArray(iNewArray, 0) = aArray(i, 0)
'        tmpArray(iNewArray, 1) = aArray(i, 1)
'        iNewArray = iNewArray + 1
'    Next i
'
'    reverse2dArrayValueKey = tmpArray
'End Function
'
'Public Function reverse2dArrayValueKeyInPlace(ByRef InputArray As Variant)
'    Dim temp As Variant
'    Dim temp2 As Variant
'    Dim Ndx As Integer
'    Dim Ndx2 As Integer
'
'    Ndx2 = UBound(InputArray)
'    ''''''''''''''''''''''''''''''''''''''
'    ' loop from the LBound of InputArray to
'    ' the midpoint of InputArray
'    ''''''''''''''''''''''''''''''''''''''
'    For Ndx = LBound(InputArray) To ((UBound(InputArray) - LBound(InputArray) + 1) \ 2)
'        'swap the elements
'        temp = InputArray(Ndx, 0)
'        temp2 = InputArray(Ndx, 1)
'        InputArray(Ndx, 0) = InputArray(Ndx2, 0)
'        InputArray(Ndx, 1) = InputArray(Ndx2, 1)
'        InputArray(Ndx2, 0) = temp
'        InputArray(Ndx2, 1) = temp2
'        ' decrement the upper index
'        Ndx2 = Ndx2 - 1
'    Next Ndx
'End Function

'Public Function make2dArrayValueKey(ByRef aArrayValue As Variant, ByRef aArrayKey As Variant)
'    Dim retArray As Variant
'    Dim shorterArray As Variant
'    Dim longerArray As Variant
'    Dim iLengthNewArray As Integer
'    Dim i As Integer
'
'
'    If f.aLength(aArrayValue) > f.aLength(aArrayKey) Then
'        shorterArray = aArrayKey
'        longerArray = aArrayValue
'    Else
'        shorterArray = aArrayValue
'        longerArray = aArrayKey
'    End If
'
'    iLengthNewArray = f.aLength(shorterArray)
'
'    ReDim retArray(iLengthNewArray - 1, 1)
'    For i = 0 To iLengthNewArray - 1
'        retArray(i, 0) = aArrayValue(LBound(aArrayValue) + i)
'        retArray(i, 1) = aArrayKey(LBound(aArrayKey) + i)
'    Next i
'
'    make2dArrayValueKey = retArray
'End Function
'
''returns length of array
'Public Function f.aLength(ByRef aArray)
'
'    If IsEmpty(aArray) Or Not IsArray(aArray) Then
'        f.aLength = -1
'        Exit Function
'    End If
'
'    f.aLength = UBound(aArray) - LBound(aArray) + 1
'End Function

'letzte Spalte in bestimmter zeile "fromRow"
Public Function getLastCol(Optional ByRef oTabelle As Worksheet, Optional fromRow As Variant = 1)
    If oTabelle Is Nothing Then
        Set oTabelle = ActiveSheet
    End If
    
    With oTabelle
        getLastCol = .Cells(fromRow, .Columns.count).End(xlToLeft).column
    End With
End Function


'letzte Zeile in bestimmte Spalte "fromCol"
Public Function getLastRow(Optional ByRef oTabelle As Worksheet, Optional fromCol As Variant = 1)
    If oTabelle Is Nothing Then
        Set oTabelle = ActiveSheet
    End If
    
    If Not IsNumeric(fromCol) Then fromCol = LetterToCol(fromCol)
    
    With oTabelle
        getLastRow = .Cells(.Rows.count, fromCol).End(xlUp).row
    End With
End Function

'erste Zeile in bestimmte Spalte "fromCol" 'TODO bessere methode?
Public Function getFirstRow(Optional ByRef oTabelle As Worksheet, Optional fromCol As Variant = 1)
    Dim i As Long
    If oTabelle Is Nothing Then
        Set oTabelle = ActiveSheet
    End If
    
    For i = 1 To getLastRow(oTabelle, fromCol)
        If oTabelle.Cells(i, fromCol) <> vbNullString Then
            getFirstRow = i
            Exit For
        End If
    Next i
    
    getFirstRow = i
End Function

'erste Spalte in bestimmter zeile "fromRow"
Public Function getFirstCol(Optional ByRef oTabelle As Worksheet, Optional fromRow As Long = 1)
    Dim i As Long
    If oTabelle Is Nothing Then
        Set oTabelle = ActiveSheet
    End If
    
    For i = 1 To getLastCol(oTabelle, fromRow)
        If oTabelle.Cells(fromRow, i) <> "" Then
            getFirstCol = i
            Exit For
        End If
    Next i
    
    getFirstCol = i
End Function

''array vergroessern und value hinzufuegen
'Function f.arrayPush(ByRef aArray As Variant, Value As Variant)
'    Dim newLength As Integer
'    Dim minLength As Integer
'
'    If IsEmpty(aArray) Or Not IsArray(aArray) Then
'        ReDim aArray(0)
'    ElseIf UBound(aArray) = -1 Then
'        ReDim aArray(0)
'    Else
'        newLength = UBound(aArray) + 1
'        minLength = LBound(aArray)
'        ReDim Preserve aArray(minLength To newLength)
'    End If
'    aArray(UBound(aArray)) = Value
'End Function
'
''array um 1 verkleinern letztes Element wird zuruckgegeben
'Function arrayPop(ByRef aArray As Variant)
'    Dim newLength As Integer
'    Dim minLength As Integer
'
'    If IsEmpty(aArray) Or Not IsArray(aArray) Then
'        arrayPop = False
'    ElseIf UBound(aArray) = -1 Then
'        arrayPop = False
'    Else
'        newLength = UBound(aArray) - 1
'        minLength = LBound(aArray)
'        arrayPop = aArray(UBound(aArray))
'        ReDim Preserve aArray(minLength To newLength)
'    End If
'End Function
'
''value loeschen und array verkleinern
'Function arrayRemoveElement(ByRef aArray, Value As Variant) As Variant
'    Dim iPos As Integer
'
'    arrayRemoveElement = False
'
'    If IsEmpty(aArray) Or Not IsArray(aArray) Then 'falls nicht array
'        arrayRemoveElement = False
'    ElseIf UBound(aArray) = -1 Then 'falls keine elements
'        arrayRemoveElement = False
'    ElseIf UBound(aArray) = 0 Then 'falls nur 1 element drin
'        aArray = Array()
'        arrayRemoveElement = True
'    Else
'        iPos = getIndexInArrayByValue(Value, aArray)
'        If iPos > -1 Then
'            Dim i As Integer
'            For i = iPos To UBound(aArray) - 1
'                aArray(i) = aArray(i + 1)
'            Next i
'
'            ReDim Preserve aArray(LBound(aArray) To UBound(aArray) - 1)
'
'            arrayRemoveElement = True
'        End If
'
'    End If
'End Function
'
''eleminiert Dublikate in einem array
'Public Function arrayRemoveDublicates(ByVal arr)
'    Dim myCollection As New Collection
'    Dim element As Variant
'    Dim i As Long
'
'    On Error Resume Next
'    For Each element In arr
'        Call myCollection.add(element, CStr(element))
'    Next element
'    On Error GoTo 0
'
'    ReDim arr(LBound(arr) To LBound(arr) + myCollection.count - 1)
'    For i = 1 To myCollection.count
'        arr(LBound(arr) + i - 1) = myCollection(i)
'    Next i
'
'    arrayRemoveDublicates = arr
'End Function
'
''return subArray of aArray, with lbound = 0
''lFrom from what element (for example from second = 1), lTo to what element, for example to 4. element = 3
'Function arrayGetSubArray(ByVal aArray As Variant, Optional ByVal lFrom As Variant = -1, Optional ByVal lTo As Variant = -1, Optional ByVal lStartBound As Variant = -1)
'    Dim lBnd As Long, ubnd As Long
'    lBnd = LBound(aArray)
'    ubnd = UBound(aArray)
'
'
'
'    If lFrom = -1 Then
'        lFrom = lBnd
'    Else
'        lFrom = lBnd + lFrom
'    End If
'
'
'    If lTo = -1 Then
'        lTo = ubnd
'    Else
'        lTo = lBnd + lTo
'    End If
'
'
'    If IsEmpty(aArray) Or Not IsArray(aArray) Then
'        arrayGetSubArray = aArray
'        Exit Function
'    End If
'
'
'    If lFrom >= lBnd And lTo <= ubnd And lFrom <= lTo Then
'        Dim tmpArray As Variant
'        Dim i As Long
'        ReDim tmpArray(lFrom To lTo)
'
'        For i = lFrom To lTo
'            tmpArray(i) = aArray(i)
'        Next i
'
'        aArray = tmpArray
'    End If
'
'    'startbound setzen
'    If lStartBound >= 0 Then
'        Call f.arrayChangeBounds(aArray, lStartBound)
'    End If
'
'
'    arrayGetSubArray = aArray
'End Function
'
''changes bounds of array.
''todo optional vars
'Function f.arrayChangeBounds(ByRef aArray As Variant, lStartBound As Variant, Optional lEndBound As Variant = -1)
'    If Not IsArray(aArray) Then
'        f.arrayChangeBounds = aArray
'        Exit Function
'    End If
'
'    If lEndBound = -1 Then lEndBound = lStartBound + f.aLength(aArray) - 1
'
'    If lStartBound < 0 Or lEndBound < 0 Or lStartBound > lEndBound Then
'        f.arrayChangeBounds = aArray
'        Exit Function
'    End If
'
'
'    If IsEmpty(aArray) Then
'        ReDim Preserve aArray(lStartBound To lEndBound)
'        f.arrayChangeBounds = aArray
'        Exit Function
'    End If
'
'    Dim i As Long
'    Dim lBnd As Long
'    Dim tmpArray As Variant
'    ReDim tmpArray(lStartBound To lEndBound)
'
'    'werte nacheinanden kopieren
'    lBnd = LBound(aArray)
'    For i = lBnd To UBound(aArray)
'        'nur falls es noch in tmpArray platz gibt
'        If (i - lBnd <= lEndBound - lStartBound) Then
'            tmpArray(lStartBound + i - lBnd) = aArray(i)
'        End If
'    Next i
'
'    aArray = tmpArray
'    f.arrayChangeBounds = tmpArray
'End Function
'
'
''array in eine Collection umwandeln
'Function arrayToCollection(aArray As Variant) As Collection
'    If aArray = Empty Then Exit Function
'    Dim colReturn As New Collection
'    Dim wert As Variant
'    For Each wert In aArray
'        colReturn.add (wert)
'    Next wert
'    Set arrayToCollection = colReturn
'End Function
'
''Collection in ein array umwandeln
'Function collectionToArray(colCollection As Collection) As Variant
'    If colCollection Is Nothing Then Exit Function
'    Dim wert As Variant
'    Dim aReturn() As Variant
'    aReturn = Array()
'    For Each wert In colCollection
'        Call f.arrayPush(aReturn, wert)
'    Next wert
'    collectionToArray = aReturn
'End Function
'
''modify 2dArray to Array Of Arrays
'Public Function array2dToArrOfArrs(aArray As Variant)
'    Dim i As Long, j As Long
'    Dim retArray As Variant
'    Dim lbound1 As Long
'    Dim ubound1 As Long
'    Dim lbound2 As Long
'    Dim ubound2 As Long
'
'
'    If Not IsArray(aArray) Then
'        array2dToArrOfArrs = aArray
'        Exit Function
'    End If
'
'    If f.aLength(aArray) = 0 Then
'        array2dToArrOfArrs = Array()
'        Exit Function
'    End If
'
'    lbound1 = LBound(aArray, 1)
'    ubound1 = UBound(aArray, 1)
'
'    ReDim retArray(lbound1 To ubound1)
'    For i = lbound1 To ubound1
'        Dim tmpArray As Variant
'        lbound2 = LBound(aArray, 2)
'        ubound2 = UBound(aArray, 2)
'        ReDim tmpArray(lbound2 To ubound2)
'        For j = lbound2 To ubound2
'            tmpArray(j) = aArray(i, j)
'        Next j
'
'        retArray(i) = tmpArray
'
'    Next i
'
'    array2dToArrOfArrs = retArray
'End Function

'todo better own func ?
'modify Array of Array to 2dArray. Beginns with array(1,1)
Function ArrOfArrsToArray2d(aArray As Variant)
    Dim retArray As Variant
    
    If Not IsArray(aArray) Then
        ArrOfArrsToArray2d = aArray
        Exit Function
    End If
    
    If f.aLength(aArray) = 0 Then
        ArrOfArrsToArray2d = Array()
        Exit Function
    End If
    
    retArray = WorksheetFunction.index(aArray, 0, 0)
    
    ArrOfArrsToArray2d = retArray
End Function

    
    

''max of both values
'Function max(value1 As Variant, value2 As Variant) As Integer
'    If value1 = Empty Then
'        max = value2
'        Exit Function
'    ElseIf value2 = Empty Then
'        max = value1
'        Exit Function
'    Else
'        If IsNumeric(value1) And IsNumeric(value2) Then
'            If value1 > value2 Then
'                max = value1
'            Else
'                max = value2
'            End If
'        End If
'    End If
'End Function

''check if value is in Array, bAnyTypeCompare = true ? then check any Types (modify with Cstr before compare)
'Function isInArray(aArray, vValue, Optional bAnyTypeCompare As Boolean = False) As Boolean
'    Dim Value As Variant
'    Dim bIsEqual As Boolean
'
'    isInArray = False
'    For Each Value In aArray
'        If bAnyTypeCompare Then
'            bIsEqual = (CStr(Value) = CStr(vValue))
'        Else
'            bIsEqual = (Value = vValue)
'        End If
'
'        If bIsEqual Then
'            isInArray = True
'            Exit For
'        End If
'    Next Value
'End Function


'workbook finden, falls schon geoeffnet, oder oeffnen falls nicht geoeffnet
Function findOpenedDocOrOpen(filePath As String) As Workbook
    Dim wrkReturn As Workbook
    On Error GoTo findOpenedDocOrOpen_Error

    Set wrkReturn = findOpenedDoc(filePath)
    If wrkReturn Is Nothing Then
        Set findOpenedDocOrOpen = Workbooks.Open(filePath)
    Else
        Set findOpenedDocOrOpen = wrkReturn
    End If

    Exit Function
findOpenedDocOrOpen_Error:

    f.dbg "error in findOpenedDocOrOpen of Klassenmodul clsFuncsExl"
    Set findOpenedDocOrOpen = Nothing
End Function

'sucht geoeffnete datei mit dem Pfad, oder mit dem Namen
Function findOpenedDoc(filePath As String, Optional bByName As Boolean = False, Optional bLikeName As Boolean = False) As Workbook
    Dim curWkb As Workbook
    If bByName Then
        Set findOpenedDoc = findOpenedDocByName(filePath, bLikeName)
    Else
        For Each curWkb In Workbooks
            If StrComp(curWkb.FullName, filePath, vbTextCompare) = 0 Then
                Set findOpenedDoc = curWkb
                Exit For
            End If
        Next curWkb
    End If
End Function

'geoeffnetes Dok finden, dessen name ist gleich "sName" ODER falls bLikeName true ist, dessen name sName drin hat.
Function findOpenedDocByName(sName As String, Optional bLikeName As Boolean = False) As Workbook
    Dim curWkb As Workbook
    For Each curWkb In Workbooks
        If curWkb.name = sName Then
            Set findOpenedDocByName = curWkb
            Exit For
        ElseIf bLikeName And InStr(1, curWkb.name, sName) > 0 Then
            Set findOpenedDocByName = curWkb
            Exit For
        End If
    Next curWkb
End Function

'screenUpdating und Formulasberechnung Ein/Ausschalten
Public Function screenBerechnungSchutz(einAus As Boolean, Optional sSchutzPw As String = "2012")
    'reinfolge beim einschalten, ausschalten 'no need? Oo
    If einAus Then
        If IsEmpty(sSchutzPw) Then
            Call blattschutz(einAus)
        Else
            Call blattschutz(einAus, sSchutzPw)
        End If
        berechnung (einAus)
        screen (einAus)

    Else
        screen (einAus)
        berechnung (einAus)
        If IsEmpty(sSchutzPw) Then
            Call blattschutz(einAus)
        Else
            Call blattschutz(einAus, sSchutzPw)
        End If
    End If
    
End Function

'screen und berechnung ausschalten. oefter benoetigt..
Public Function berechnungAndScreen(einAus As Boolean)
    screen (einAus)
    berechnung (einAus)
End Function

'screenUpdating Ein/Ausschalten
Public Function screen(einAus As Boolean)
    If (CBool(einAus)) Then
        Excel.Application.ScreenUpdating = True
    Else
        Excel.Application.ScreenUpdating = False
    End If
End Function

'Formulasberechnung Ein/Ausschalten
Function berechnung(einAus As Boolean)
    If Excel.Application.Workbooks.count <= 0 Then Exit Function ' falls nichts geoeffnet, exit
    
    If (CBool(einAus)) Then
        Excel.Application.Calculation = xlCalculationAutomatic
    Else
        Excel.Application.Calculation = xlCalculationManual
    End If
End Function

'open dialog to pick a file, return "" on error.
Public Function getFilePathDialog(sFileFilter As String, sTitle As String) As String
'Pfad zur Datei bestimmen
    'On Error Resume Next
    On Error GoTo ErrorHandler

    getFilePathDialog = Excel.Application.GetOpenFilename(FileFilter:=sFileFilter, Title:=sTitle)
    'todo? international?
    If getFilePathDialog = "Falsch" Then getFilePathDialog = ""
    
    Exit Function
ErrorHandler:
    getFilePathDialog = ""
End Function

'filePicker or FolderPicker dialog. 'todo miltiselect option ?
Public Function DialogGetPath(Optional bFileOrFolder As Boolean = True, Optional sTitle As String, Optional filterTyp As String, Optional filterExt As String, Optional sInitial As String, Optional bMulti As Boolean = False, Optional sButtonName As String)
    Dim fd As FileDialog

    
        'if true -> file
        If bFileOrFolder Then
            Set fd = Excel.Application.FileDialog(msoFileDialogFilePicker)
            
            'falls filter werde vorhanden
            If f.hatWert(filterTyp) And f.hatWert(filterExt) Then
                fd.Filters.add filterTyp, filterExt
            End If
            'extra button title
            If f.hatWert(sButtonName) Then
                fd.buttonName = sButtonName
            End If
        
        'folder
        Else
            Set fd = Excel.Application.FileDialog(msoFileDialogFolderPicker)
        End If
        
        'anfangspfad
        If f.hatWert(sInitial) Then fd.InitialFileName = sInitial
        
        'if multiselected
        If bMulti Then
            fd.AllowMultiSelect = True
        End If
        
        If f.hatWert(sTitle) Then
            fd.Title = sTitle
        End If
        
        
        
        Dim vrtResult As Variant
        Dim vrtSelectedItem As Variant
        
        If fd.Show = -1 Then
    
            'Step through each string in the FileDialogSelectedItems collection.
            For Each vrtSelectedItem In fd.SelectedItems
                'extra slash fuer ordner
                If Not bFileOrFolder Then
                    If Right(vrtSelectedItem, 1) <> "\" Then vrtSelectedItem = vrtSelectedItem & "\"
                End If
            
                'falls multidatei, dann als Array zuruckgeben
                If bMulti Then
                    Call f.arrayPush(vrtResult, vrtSelectedItem)
                Else
                    vrtResult = vrtSelectedItem
                    Exit For 'zur sicherheit
                End If
                
                
                
            Next vrtSelectedItem
        'The user pressed Cancel.
        Else
            vrtResult = ""
        End If

        
    DialogGetPath = vrtResult
End Function

'blattschutz aufheben, oder wieder einschalten
Public Sub blattschutz(bYesNo As Boolean, Optional sSchutzPw As String = "2012", Optional ByRef wkbWorkbook As Workbook)
    On Error Resume Next
    Dim i As Integer
    
    If wkbWorkbook Is Nothing Then Set wkbWorkbook = ThisWorkbook
    
    With wkbWorkbook
    If bYesNo Then
        For i = 1 To .Worksheets.count
            With Worksheets(i)
            .Protect Password:=sSchutzPw, UserInterfaceOnly:=True
            '.EnableSelection = xlUnlockedCells
            '.EnableSelection = xlNoRestrictions
            End With
        Next i
    Else
        For i = 1 To .Worksheets.count
            With Worksheets(i)
            .Unprotect Password:=sSchutzPw
            '.EnableSelection = xlNoRestrictions
            End With
        Next i
    End If
    End With
End Sub

'prueft ob txtcheck ein wert hat, also ob es nicht 0 und nicht "" ist.
'bTrim ob auch trim eingesetz werden muss /default ja
'Public Function f.hatWert(Optional ByVal txtcheck, Optional bTrim As Boolean = True)
'    If Not IsMissing(txtcheck) Then
'
'        If bTrim Then
'            txtcheck = Trim(txtcheck)
'        End If
'
'        If Not (CStr(txtcheck) = "" Or CStr(txtcheck) = "0" Or Len(CStr(txtcheck)) = 0) Then
'            f.hatWert = True
'            Exit Function
'        End If
'        f.hatWert = False
'    End If
'End Function
'
''prueft f.hatWert fuer alle stings in Array
'Public Function f.hatWertArray(ByVal aTxtCheck As Variant, Optional bTrim As Boolean = True)
'    Dim element As Variant
'    f.hatWertArray = True
'    For Each element In aTxtCheck
'        If Not f.hatWert(element, bTrim) Then
'            f.hatWertArray = False
'            Exit For
'        End If
'    Next element
'End Function

'calculates all sheet in Workbook
Public Function refreshAllSheets(Optional wkbWorkbook As Workbook)
    Dim sheet As Variant
    If wkbWorkbook Is Nothing Then Set wkbWorkbook = ThisWorkbook
    
    For Each sheet In wkbWorkbook.Sheets
        sheet.Calculate
    Next sheet
End Function

'calculate array of tables objects
Public Function refreshTables(aTables As Variant)
    Dim wksTable As Variant
    For Each wksTable In aTables
        wksTable.Calculate
    Next wksTable
End Function

''returns index of element first foun din array
'Function getIndexInArrayByValue(vWert As Variant, aArray As Variant)
'    Dim i As Integer
'
'    getIndexInArrayByValue = -1
'
'    For i = LBound(aArray) To UBound(aArray)
'        If UCase(aArray(i)) = UCase(vWert) Then
'            getIndexInArrayByValue = i
'            Exit For
'        End If
'    Next i
'
'End Function

''check if file exists
'Public Function fileExists(sPath As String)
'    On Error Resume Next
'    fileExists = True
'    If dir(sPath) = "" Then
'        fileExists = False
'    End If
'End Function
'
''check if all files exists
'Public Function fileExistsArray(aPaths As Variant)
'    On Error GoTo ErrorHandler
'    Dim sPath As Variant
'
'    If IsArray(aPaths) Then
'        fileExistsArray = True
'        For Each sPath In aPaths
'            If Not fileExists(CStr(sPath)) Then
'                fileExistsArray = False
'                Exit For
'            End If
'        Next sPath
'    End If
'
'    Exit Function
'ErrorHandler:
'    fileExistsArray = False
'End Function

''quicksort, wegen kaskadenrekursion nur fuer kleine array benutzen..
'Sub QuickSort(ByRef sArray As Variant, Optional ByVal MinElem As Long = -1, Optional MaxElem As Long = -1)
''
'' QuickSort()
''
'' Sortieren eines Arrays mit dem QuickSort-Algorithmus, dem wohl schnellsten
'' Sortieralgorithmus von Welt.
''
'' IN:   sArray      Array das sortiert werden soll
''       MinElem     erstes Element des Arrays (oder Teil-Arrays) default: LBound(sArray)
''       MaxElem     letztes Element des Arrays (oder Teil-Arrays) default: UBound(sArray)
''
'
'If IsMissing(MinElem) Or MinElem = -1 Then MinElem = LBound(sArray)
'If IsMissing(MaxElem) Or MaxElem = -1 Then MaxElem = UBound(sArray)
'
'Dim Mitte As Long
'Dim vDummy As Variant
'Dim vMitte As Variant
'Dim i As Long, j As Long
'    '
'    ' Abbruchbedingung der Rekursion prüfen
'    '
'    If MinElem > MaxElem Then
'        '
'        ' Rekursion beenden
'        '
'        Exit Sub
'    End If
'    '
'    ' Ermittlung der Mitte des Arrays
'    '
'    Mitte = (MinElem + MaxElem) \ 2
'    vMitte = sArray(Mitte)
'    '
'    ' Für die Prüfung der linken und rechten
'    ' Seite die Zähler initialisieren
'    '
'    i = MinElem
'    j = MaxElem
'    Do
'        '
'        ' Von links bis zur Mitte prüfen
'        '
'        Do While sArray(i) < vMitte
'            i = i + 1
'        Loop
'        '
'        ' Von rechts bis zur Mitte prüfen
'        '
'        Do While sArray(j) > vMitte
'            j = j - 1
'        Loop
'
'        If i <= j Then
'            '
'            ' Die beiden gefundenen, falsch einsortierten
'            ' Elemente vertauschen
'            '
'            vDummy = sArray(j)
'            sArray(j) = sArray(i)
'            sArray(i) = vDummy
'            '
'            ' Prüfung bei der nächsten Position fortsetzen
'            '
'            i = i + 1
'            j = j - 1
'        End If
'
'    Loop Until i > j
'    '
'    ' Rekursiver Aufruf mit den Teil-Arrays
'    '
'    QuickSort sArray, MinElem, j
'    QuickSort sArray, i, MaxElem
'End Sub
'
''sorts first elements of multidim. array(x,1) 0,1
'Sub QuickSort2d(ByRef sArray As Variant, ByVal MinElem As Long, MaxElem As Long)
''
'' QuickSort()
''
'' Sortieren eines Arrays mit dem QuickSort-Algorithmus, dem wohl schnellsten
'' Sortieralgorithmus von Welt.
''
'' IN:   sArray      Array das sortiert werden soll
''       MinElem     erstes Element des Arrays (oder Teil-Arrays)
''       MaxElem     letztes Element des Arrays (oder Teil-Arrays)
''
'Dim Mitte As Long
'Dim vDummy As Variant
'Dim vDummy2 As Variant
'Dim vMitte As Variant
'Dim i As Long, j As Long
'    '
'    ' Abbruchbedingung der Rekursion prüfen
'    '
'    If MinElem > MaxElem Then
'        '
'        ' Rekursion beenden
'        '
'        Exit Sub
'    End If
'    '
'    ' Ermittlung der Mitte des Arrays
'    '
'    Mitte = (MinElem + MaxElem) \ 2
'    vMitte = sArray(Mitte, 0)
'    '
'    ' Für die Prüfung der linken und rechten
'    ' Seite die Zähler initialisieren
'    '
'    i = MinElem
'    j = MaxElem
'    Do
'        '
'        ' Von links bis zur Mitte prüfen
'        '
'        Do While sArray(i, 0) < vMitte
'            i = i + 1
'        Loop
'        '
'        ' Von rechts bis zur Mitte prüfen
'        '
'        Do While sArray(j, 0) > vMitte
'            j = j - 1
'        Loop
'
'        If i <= j Then
'            '
'            ' Die beiden gefundenen, falsch einsortierten
'            ' Elemente vertauschen
'            '
'            vDummy = sArray(j, 0)
'            vDummy2 = sArray(j, 1)
'            sArray(j, 0) = sArray(i, 0)
'            sArray(j, 1) = sArray(i, 1)
'            sArray(i, 0) = vDummy
'            sArray(i, 1) = vDummy2
'            '
'            ' Prüfung bei der nächsten Position fortsetzen
'            '
'            i = i + 1
'            j = j - 1
'        End If
'
'    Loop Until i > j
'    '
'    ' Rekursiver Aufruf mit den Teil-Arrays
'    '
'    QuickSort2d sArray, MinElem, j
'    QuickSort2d sArray, i, MaxElem
'End Sub
'
''Heap Sort. 'http://www.vbforums.com/showpost.php?p=2909250&postcount=7
'Public Function HeapSort(ByRef pvarArray As Variant)
'        Dim i As Long
'        Dim iMin As Long
'        Dim iMax As Long
'        Dim varSwap As Variant
'
'        iMin = LBound(pvarArray)
'        iMax = UBound(pvarArray)
'        For i = (iMax + iMin) \ 2 To iMin Step -1
'            HeapSortHelper pvarArray, i, iMin, iMax
'        Next i
'        For i = iMax To iMin + 1 Step -1
'            varSwap = pvarArray(i)
'            pvarArray(i) = pvarArray(iMin)
'            pvarArray(iMin) = varSwap
'            HeapSortHelper pvarArray, iMin, iMin, i - 1
'        Next i
'        HeapSort = pvarArray
'    End Function
'    Private Function HeapSortHelper(ByRef pvarArray As Variant, ByVal i As Long, iMin As Long, iMax As Long)
'        Dim lngLeaf As Long
'        Dim varSwap As Variant
'
'        Do
'            lngLeaf = i + i - (iMin - 1)
'            Select Case lngLeaf
'                Case Is > iMax: Exit Do
'                Case Is < iMax: If pvarArray(lngLeaf + 1) > pvarArray(lngLeaf) Then lngLeaf = lngLeaf + 1
'            End Select
'            If pvarArray(i) > pvarArray(lngLeaf) Then Exit Do
'            varSwap = pvarArray(i)
'            pvarArray(i) = pvarArray(lngLeaf)
'            pvarArray(lngLeaf) = varSwap
'            i = lngLeaf
'        Loop
'    End Function
'
''http://www.freevbcode.com/ShowCode.asp?ID=1086
'Public Function IsAlphaBetical(TestString As String) As Boolean
'
'    Dim sTemp As String
'    Dim iLen As Integer
'    Dim iCtr As Integer
'    Dim sChar As String
'
'    'returns true if all characters in a string are alphabetical
'    'returns false otherwise or for empty string
'
'    sTemp = TestString
'    iLen = Len(sTemp)
'    If iLen > 0 Then
'        For iCtr = 1 To iLen
'            sChar = Mid(sTemp, iCtr, 1)
'            If Not sChar Like "[A-Za-z]" Then Exit Function
'        Next
'
'    IsAlphaBetical = True
'    End If
'
'End Function


'worksheet hinzufuegen, am ende. TODO : index wohin.
Public Function addWorksheet(Optional sWksName As String, Optional wkbWorkbook As Workbook)
    If wkbWorkbook Is Nothing Then
        Set wkbWorkbook = ThisWorkbook
    End If
    
    If checkIfWksExists(wkbWorkbook, sWksName) Then
        addWorksheet = Null
        Exit Function
    Else
        Set addWorksheet = wkbWorkbook.Worksheets.add(After:=wkbWorkbook.Worksheets(wkbWorkbook.Worksheets.count))
        If sWksName <> vbNullString Then
            addWorksheet.name = sWksName
        End If
    End If
    
End Function


'array von worksheets ausdrucken, worksheets mussen in einer arbeitsmappe sein!!
'bEinzeln - einzeln ausdrucken
'bFirstPrint - drucker auswahlfenster ja/nein
'bOrdered ob alle sheets in richtige reihenfolge gedruckt werden sollen
Public Function printSheets(aWorksheets As Variant, Optional bEinzeln As Boolean = False, _
                            Optional bFirstPrint As Boolean = True, Optional bOrdered As Boolean = True, Optional wkbWorkbook As Workbook)
    On Error GoTo PrintSheets_Error

    If Not IsArray(aWorksheets) Then Exit Function
    Dim wksFromArray As Variant
    Dim bFirstSelect As Boolean
    Dim bDisplayPageBreaksB4 As Boolean
    If wkbWorkbook Is Nothing Then Set wkbWorkbook = ThisWorkbook
    Dim aReihnfolgeB4 As Variant 'array of worksheets refs
    
    bDisplayPageBreaksB4 = aWorksheets(0).DisplayPageBreaks 'todo save for each page?
    
    'case must be ordered, save original order, and sort by new array
    If bOrdered Then
        aReihnfolgeB4 = getWksArray(wkbWorkbook)
        worksheetsOrder (aWorksheets)
    End If
    
    If Not bEinzeln Then 'print all merged
        
        bFirstSelect = True 'to prevent print of selections b4
        For Each wksFromArray In aWorksheets
            wksFromArray.Select bFirstSelect
            bFirstSelect = False
        Next wksFromArray
        
        If bFirstPrint Then
            Application.Dialogs(xlDialogPrint).Show
        Else
            ActiveWindow.SelectedSheets.PrintOut
        End If
        
    Else 'print separate
        For Each wksFromArray In aWorksheets
            wksFromArray.Select
            
            If bFirstPrint Then
                Application.Dialogs(xlDialogPrint).Show
            Else
                ActiveWindow.SelectedSheets.PrintOut
            End If
            
            bFirstPrint = False
        Next wksFromArray
    End If
    
    printSheets = True
    
Final:
    'case must be ordered, order original
    If bOrdered Then
        worksheetsOrder (aReihnfolgeB4)
    End If
    
    'pagebreaks, hide?
    If IsArray(aWorksheets) Then
        For Each wksFromArray In aWorksheets
            wksFromArray.DisplayPageBreaks = bDisplayPageBreaksB4
        Next wksFromArray
    End If
    
    Exit Function

PrintSheets_Error:
    printSheets = False
    Debug.Print "PrintSheets - ERROR"
    GoTo Final
End Function

'returns array of refs of worksheets in the wkbWorkbook
Function getWksArray(Optional wkbWorkbook As Workbook)
    Dim wksSheet As Worksheet
    Dim aReturnArray As Variant
    If wkbWorkbook Is Nothing Then Set wkbWorkbook = ThisWorkbook
    
    For Each wksSheet In wkbWorkbook.Worksheets
        Call f.arrayPush(aReturnArray, wksSheet)
    Next wksSheet
    
    getWksArray = aReturnArray
End Function

'orders worksheets as order of array
Function worksheetsOrder(orderedArrayOfWks As Variant)
    Dim i As Integer
    Dim iArrLbnd As Integer
    Dim iArrUbnd As Integer
    
    worksheetsOrder = False
    If Not IsArray(orderedArrayOfWks) Then Exit Function
    If f.aLength(orderedArrayOfWks) <= 1 Then Exit Function 'no need to sort 1 element
    
    iArrLbnd = LBound(orderedArrayOfWks)
    iArrUbnd = UBound(orderedArrayOfWks)
    
    For i = iArrLbnd + 1 To iArrUbnd
        orderedArrayOfWks(i).Move After:=orderedArrayOfWks(i - 1)
    Next i
    
    worksheetsOrder = True
End Function

'reads 2 column range, to collection, as key - value
Function rangeToMyCollection(ByRef rngKeyValue As Range) As clsMyCollection
    Dim retCollection As New clsMyCollection
    Dim aValues As Variant
    Dim i As Long
    Dim sKey As String
    Dim sValue As String
    
    'falls nicht 2 dimensional
    If rngKeyValue.Columns.count <> 2 Then
        GoTo ReturnValue
    End If
    
    aValues = rngKeyValue.value
    For i = LBound(aValues) To UBound(aValues)
        sKey = CStr(aValues(i, 1))
        sValue = CStr(aValues(i, 2))
        
        If Len(sKey) > 0 Then
            Call retCollection.add(sValue, sKey)
        End If
        
    Next i
    
ReturnValue:
    Set rangeToMyCollection = retCollection
End Function

'saves data from myCollection to range, range must be 2dim and has enoght rows
Function myCollectionToRange(ByRef mcCollection As clsMyCollection, ByRef rngKeyValue As Range) As Boolean
    Dim i As Long
    
    'falls nicht 2 dimensional
    If rngKeyValue.Columns.count <> 2 Then
        myCollectionToRange = False
        Exit Function
    End If
    
    'falls nicht genuegend zeilen
    If rngKeyValue.Rows.count < mcCollection.getSize Then
        myCollectionToRange = False
        Exit Function
    End If
    
    'save values
    For i = 0 To mcCollection.getSize - 1
        rngKeyValue.Cells(i + 1, 1).value = mcCollection.getKeys(i)
        rngKeyValue.Cells(i + 1, 2).value = mcCollection.GetValues(i)
    Next i
    
    myCollectionToRange = True
End Function

Private Sub Class_Initialize()
    Set f = New clsFuncs
End Sub

Private Sub Class_Terminate()
    Set f = Nothing
End Sub
