VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuncs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFuncs
' Author    : Dmitry Gorelenkov
' Date      : 05.2012
' Changed   : 07.06.2014
' Purpose   : Functionssammlung - Allgemein
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------


Option Explicit

'API showWindows, nCmdShow = 0 => ausblenden, = 1 => anzeigen
Private Declare Function APIShowWindow Lib "User32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Sub APISleep Lib "kernel32" Alias "Sleep" (ByVal lngMilliseconds As Long)
Private Declare Function APISetForegroundWindow Lib "User32" Alias "SetForegroundWindow" (ByVal hwnd As Long) As Long
Private Declare Function APIGetActiveWindow Lib "User32" Alias "GetActiveWindow" () As Long

Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" _
(ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Long, ByVal fuWinIni As Long) As Long

Private Declare Function SetWindowPos Lib "User32" _
 (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
 ByVal x As Long, ByVal y As Long, _
 ByVal cx As Long, ByVal cy As Long, _
 ByVal wFlags As Long) As Long
' ================
' SetWindowPos Info
' ================
Private Enum SWPValues
    SWP_SHOWWINDOW = &H40
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
End Enum
 
Private Const SPI_GETSCREENSAVERRUNNING As Long = &H72

'ein window zu Xpos, Ypos positionieren
Public Function MoveWindow(ByVal hwnd As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Call SetWindowPos( _
     hwnd, hWndInsertAfter:=0, x:=Xpos, y:=Ypos, cx:=0, cy:=0, _
     wFlags:=SWP_NOSIZE Or SWP_NOZORDER)
End Function

'checks if screensaver is running
'http://www.mrexcel.com/forum/excel-questions/523494-visual-basic-applications-determine-whether-screen-saver-running.html
Public Function IsScrSaverRunning() As Boolean

    Dim lScrSrvRunning As Long
            
    SystemParametersInfo _
    SPI_GETSCREENSAVERRUNNING, vbNull, lScrSrvRunning, vbNull
    
    IsScrSaverRunning = CBool(lScrSrvRunning)
           
End Function

Public Function GetActiveWindow() As Long
    GetActiveWindow = APIGetActiveWindow
End Function

'window auf vordergrund setzen, bei hWnd
Public Function SetForegroundWindow(ByVal hwnd As Long)
    Call APISetForegroundWindow(hwnd)
End Function

'sleep in MS
Public Function Sleep(lngMs As Long) As Boolean
    If lngMs > 0 Then
        Call APISleep(lngMs)
        Sleep = True
    End If
End Function

'fenster anzeigen / verstecken
Public Function ShowWindow(lHandleWnd As Long, Optional bShow As Boolean = True)
    On Error GoTo ErrorHandler
    If bShow Then
        ShowWindow = APIShowWindow(lHandleWnd, 1)
    Else
        ShowWindow = APIShowWindow(lHandleWnd, 0)
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print "Fehler bei clsFuncs.ShowWindow"
End Function


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
    If aLength(aPriority) < aLength(aPrioElements) Then Exit Function
        
    lBestPrioIndex = -1
    'elements travisieren, hoechste prio finden
    For Each element In aElements
        lIndex = getIndexInArrayByValue(element, aPrioElements)
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

'prints an Array if bToDbg not defined, and returns printed string. bHoriz for horizontal printing
Public Function printArray(ByRef aArray, Optional bHoriz As Boolean = False, Optional bToDbg As Boolean = True) As String
    Dim elem As Variant
    Dim elemStr As String
    Dim counter As Long
    Dim retString As String
    
    If arrayGetDimension(aArray) > 1 Then
        retString = "This array has " & arrayGetDimension(aArray) & " dimensions!!" & vbCrLf
    End If
    
    counter = LBound(aArray)
    On Error Resume Next
    retString = retString & "Array(" & IIf(bHoriz, "", vbCrLf)
    For Each elem In aArray
        If IsArray(elem) Then
            elemStr = printArray(elem, True, False)
        ElseIf IsNull(elem) Then
            elemStr = "%<NULL>%"
        ElseIf IsEmpty(elem) Then
            elemStr = "%<Empty>%"
        ElseIf IsObject(elem) Then
            elemStr = "%<Object>%"
        Else
            elemStr = CStr(elem)
            If err.Number <> 0 Then
                elemStr = "%<Error>%"
            End If
        End If
        retString = retString & counter & " => " & elemStr & IIf(bHoriz, "; ", vbCrLf)
        counter = counter + 1
    Next elem
    retString = retString & ")"
    
    If bToDbg Then Debug.Print retString
    printArray = retString
End Function

'return dimension of array
Function arrayGetDimension(aArray As Variant) As Integer
    On Error GoTo err:
        Dim i As Integer
        Dim tmp As Integer
        i = 0
        Do While True:
            i = i + 1
            tmp = UBound(aArray, i)
        Loop
err:
        arrayGetDimension = i - 1
End Function

Public Function array2DimTo1Dim(arr, Optional elementNum As Variant = 1)
    Dim tmpArray()
    ReDim tmpArray(LBound(arr) To UBound(arr))
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        tmpArray(i) = arr(i, elementNum)
    Next i
array2DimTo1Dim = tmpArray
End Function

'push each element from X-dimensional array or collection to new 1-dim array
Public Function arrayXDimTo1Dim(ByVal arr As Variant)
    Dim tmpArray()
    Dim element As Variant
    Dim tmpArrayOfElem As Variant
    Dim element2 As Variant
    tmpArray = Array()
    
    For Each element In arr
        If IsArray(element) Then 'falls wieder array, weiter zerlegen
            tmpArrayOfElem = arrayXDimTo1Dim(element)
            For Each element2 In tmpArrayOfElem
                Call arrayPush(tmpArray, element2)
            Next element2
        Else
            Call arrayPush(tmpArray, element)
        End If
        
    Next element
    
    arrayXDimTo1Dim = tmpArray
End Function


'return a copy of array
Public Function arrayCopy(ByVal whatToCopy As Variant)
    If Not IsArray(whatToCopy) Then Exit Function
    arrayCopy = whatToCopy
End Function

'return a  copy of object
Public Function objCopy(ByVal whatToCopy As Object)
    If Not IsObject(whatToCopy) Then Exit Function
    Set objCopy = whatToCopy
End Function

''prueft ob sSheetName worksheet in wkbBook vorhanden ist
'Public Function checkIfWksExists(wkbBook As Workbook, sSheetName As String) As Boolean
'    Dim sheet As Variant
'    checkIfWksExists = False
'    For Each sheet In wkbBook.Sheets
'        If UCase(sheet.Name) = UCase(sSheetName) Then
'            checkIfWksExists = True
'            Exit Function
'        End If
'    Next sheet
'End Function

'prueft ob der Wert datum (als text) oder numerische Datum ist
Public Function isDateMy(vWert)
    On Error Resume Next
    isDateMy = (IsNumeric(vWert) And Len(vWert) = 5) Or IsDate(vWert)
End Function

'wandelt eine Buchstabe z.B. "B" in eine SpaltenNr also z.B. "2" um
Public Function LetterToCol(ByVal letter As String) As Variant
    If IsNumeric(letter) Then
        LetterToCol = CInt(letter)
        Exit Function
    End If
    
    Dim num As Integer
    Dim letnum As Integer
    Dim i As Integer
    Dim currentLetter As String
    letnum = Len(letter)
    
    For i = 1 To letnum
        currentLetter = Mid(letter, i, 1)
        'falls nummer vorkommt, das input zuruckgeben
        If Not IsAlphaBetical(currentLetter) Or IsNumeric(currentLetter) Then
            LetterToCol = letter
            Exit Function
        End If
        num = num * 26 + Asc(UCase(currentLetter)) - 64
    Next i
    LetterToCol = num
End Function

'wandelt SpaltenNr in eine Buchstabe um
Public Function ColToLetter(ByVal col As Variant) As String
    If Not IsNumeric(col) Then
        ColToLetter = CStr(col)
        Exit Function
    End If
    
    Dim columnString
    Dim currentLetterNumber
        While (col > 0)
            currentLetterNumber = (col - 1) Mod 26
            columnString = Chr(currentLetterNumber + 65) & columnString
            col = (col - (currentLetterNumber + 1)) / 26
        Wend
    ColToLetter = columnString
End Function

''gibt array von werten zurueck die in eine Spalte "column" (nummer oder buchstabe) sich befinden.
''Ab zeile "fromFow" bis "toRow" in eine "sheet" tabelle
'Public Function getArrayOfColumn(ByVal column As Variant, Optional fromRow As Variant = -1, Optional toRow As Variant = -1, Optional sheet As Worksheet)
'    Dim aArray
'    If sheet Is Nothing Then
'        Set sheet = ActiveSheet
'    End If
'
'    If fromRow = -1 Then fromRow = getFirstRow(sheet) 'sheet.UsedRange.row
'    If toRow = -1 Then toRow = getLastRow(sheet) 'sheet.UsedRange.Rows.Count
'
'    If IsNumeric(column) Then column = ColToLetter(column)
'
'    aArray = sheet.Range(column & fromRow & ":" & column & toRow).value2
'    getArrayOfColumn = Array2DimTo1Dim(aArray)
'End Function


''gibt array von werten zurueck die in eine Zeile "row" sich befinden.
''Ab spalte "fromCol" bis "toCol" in eine "sheet" tabelle
'Public Function getArrayOfRow(ByVal row As Variant, Optional fromCol As Variant = -1, Optional toCol As Variant = -1, Optional sheet As Worksheet)
'    Dim aArray
'    If sheet Is Nothing Then
'        Set sheet = ActiveSheet
'    End If
'
'
'    fromCol = LetterToCol(fromCol)
'    toCol = LetterToCol(toCol)
'
'    'row und cols muessen nun numerisch sein
'    If Not (IsNumeric(row) And IsNumeric(fromCol) And IsNumeric(toCol)) Then Exit Function
'
'    If fromCol = -1 Then fromCol = getFirstCol(sheet)
'    If toCol = -1 Then toCol = getLastCol(sheet)
'    With sheet
'        aArray = .Range(.Cells(row, fromCol), .Cells(row, toCol)).value2
'    End With
'
'    getArrayOfRow = arrayChangeBounds(ArrayXDimTo1Dim(aArray), 1)
'End Function

''fuellt ein Spalte mit werten von dem aArray, ab fromRow bis toRow
'Function fillColumnByArray(ByVal column As Variant, aArray As Variant, Optional fromRow As Variant = 1, Optional toRow As Variant = -1, Optional sheet As Worksheet, Optional frmlaLocal = True)
'    'Dim aArrayOut
'    If toRow = -1 Then toRow = fromRow + UBound(aArray) - LBound(aArray)
'    If sheet Is Nothing Then
'        Set sheet = ActiveSheet
'    End If
'    If IsNumeric(column) Then column = ColToLetter(column)
'
'    With sheet
'    If frmlaLocal Then
'        .Range(column & fromRow & ":" & column & toRow).FormulaLocal = WorksheetFunction.Transpose(aArray)
'    Else
'        '.range(column & fromRow & ":" & column & toRow).FormulaArray = WorksheetFunction.Transpose(aArray)
'        'maximum 255 zeichen bei FromulaArray .. deswegen ..
'        Dim i As Long
'        For i = fromRow To toRow
'            .Cells(i, column).Formula = aArray(i - fromRow)
'        Next i
'    End If
'    End With
'
'End Function

'topX elemente von einem Array
Function getTopX(ByVal aValues As Variant, topWhat)
    Dim aRet()
    Dim iBegin As Integer
    ReDim aRet(UBound(aValues))
    Dim i As Long
    Call QuickSort(aValues, LBound(aValues), UBound(aValues))
    iBegin = UBound(aValues) - topWhat + 1
    For i = iBegin To UBound(aValues)
        aRet(i - iBegin) = aValues(i)
    Next i
    
    ReDim Preserve aRet(topWhat - 1)
    
    getTopX = aRet
End Function

'vergleicht werte beides arrays, return 2te werte des assoziativen, falls 1. wert gleich ist.
'z.b. Array(2,3) und Array({2,"yea"}, {1, "zz"}, {3, "fff3"}) als parameter, wird zuruckgegeben Array("yea","fff3")
'falls nicht gefunden, wird "" als wert gesetzt
Function getAssotiations(arrayKeys As Variant, cCollectionOfAssotiations As Collection)
    Dim i As Long
    Dim retArray, test
    ReDim retArray(LBound(arrayKeys) To UBound(arrayKeys))
    For i = LBound(arrayKeys) To UBound(arrayKeys)
        On Error GoTo nextOne
            test = cCollectionOfAssotiations(CStr(arrayKeys(i)))
            retArray(i) = test
nextOne:
    Next i
    getAssotiations = retArray
End Function

'gibt collection mit assotiationen, value - key von arrays entsprechend zuruck
Function makeAssotiativeCollection(arrayValues, arrayKeys)
    Dim cReturn As New Collection
    Dim i As Long
    
    On Error Resume Next
    For i = LBound(arrayKeys) To UBound(arrayKeys)
        Call cReturn.Add(arrayValues(i), CStr(arrayKeys(i)))
    Next i
    
    Set makeAssotiativeCollection = cReturn
End Function

''no need? :(
'Function modifyValuesToPercentString(ByVal aArray)
'    Dim i As Integer
'    Dim newArr() As String
'    ReDim newArr(LBound(aArray) To UBound(aArray))
'    For i = LBound(aArray) To UBound(aArray)
'        newArr(i) = CStr(Format(CStr(aArray(i)), "0.0%"))
'    Next i
'
'    modifyValuesToPercentString = newArr
'End Function



Function reverse2dArrayValueKey(ByRef aArray As Variant)
    Dim tmpArray
    Dim iNewArray As Integer
    Dim i As Long
    ReDim tmpArray(LBound(aArray) To UBound(aArray), 1)
    
    iNewArray = LBound(aArray)
    For i = UBound(aArray) To LBound(aArray) Step -1
        tmpArray(iNewArray, 0) = aArray(i, 0)
        tmpArray(iNewArray, 1) = aArray(i, 1)
        iNewArray = iNewArray + 1
    Next i
    
    reverse2dArrayValueKey = tmpArray
End Function

Function reverse2dArrayValueKeyInPlace(ByRef InputArray As Variant)
    Dim temp As Variant
    Dim temp2 As Variant
    Dim Ndx As Integer
    Dim Ndx2 As Integer
    
    Ndx2 = UBound(InputArray)
    ''''''''''''''''''''''''''''''''''''''
    ' loop from the LBound of InputArray to
    ' the midpoint of InputArray
    ''''''''''''''''''''''''''''''''''''''
    For Ndx = LBound(InputArray) To ((UBound(InputArray) - LBound(InputArray) + 1) \ 2)
        'swap the elements
        temp = InputArray(Ndx, 0)
        temp2 = InputArray(Ndx, 1)
        InputArray(Ndx, 0) = InputArray(Ndx2, 0)
        InputArray(Ndx, 1) = InputArray(Ndx2, 1)
        InputArray(Ndx2, 0) = temp
        InputArray(Ndx2, 1) = temp2
        ' decrement the upper index
        Ndx2 = Ndx2 - 1
    Next Ndx
End Function

Function make2dArrayValueKey(ByRef aArrayValue As Variant, ByRef aArrayKey As Variant)
    Dim retArray As Variant
    Dim shorterArray As Variant
    Dim longerArray As Variant
    Dim iLengthNewArray As Integer
    Dim i As Long
    

    If aLength(aArrayValue) > aLength(aArrayKey) Then
        shorterArray = aArrayKey
        longerArray = aArrayValue
    Else
        shorterArray = aArrayValue
        longerArray = aArrayKey
    End If
        
    iLengthNewArray = aLength(shorterArray)
    
    ReDim retArray(iLengthNewArray - 1, 1)
    For i = 0 To iLengthNewArray - 1
        retArray(i, 0) = aArrayValue(LBound(aArrayValue) + i)
        retArray(i, 1) = aArrayKey(LBound(aArrayKey) + i)
    Next i
        
    make2dArrayValueKey = retArray
End Function

'returns length of array
Function aLength(ByRef aArray)

    If IsEmpty(aArray) Or Not IsArray(aArray) Then
        aLength = -1
        Exit Function
    End If
    
    aLength = UBound(aArray) - LBound(aArray) + 1
End Function
'
''letzte Spalte in bestimmter zeile "fromRow"
'Public Function getLastCol(Optional ByRef oTabelle As Worksheet, Optional fromRow As Variant = 1)
'    If oTabelle Is Nothing Then
'        Set oTabelle = ActiveSheet
'    End If
'
'    With oTabelle
'        getLastCol = .Cells(fromRow, .Columns.Count).End(xlToLeft).column
'    End With
'End Function

'
''letzte Zeile in bestimmte Spalte "fromCol"
'Public Function getLastRow(Optional ByRef oTabelle As Worksheet, Optional fromCol As Variant = 1)
'    If oTabelle Is Nothing Then
'        Set oTabelle = ActiveSheet
'    End If
'
'    With oTabelle
'        getLastRow = .Cells(.Rows.Count, fromCol).End(xlUp).row
'    End With
'End Function

''erste Zeile in bestimmte Spalte "fromCol" 'TODO bessere methode?
'Public Function getFirstRow(Optional ByRef oTabelle As Worksheet, Optional fromCol As Variant = 1)
'    Dim i As Long
'    If oTabelle Is Nothing Then
'        Set oTabelle = ActiveSheet
'    End If
'
'    For i = 1 To getLastRow(oTabelle, fromCol)
'        If oTabelle.Cells(i, fromCol) <> "" Then
'            getFirstRow = i
'            Exit For
'        End If
'    Next i
'
'    getFirstRow = i
'End Function

''erste Spalte in bestimmter zeile "fromRow"
'Public Function getFirstCol(Optional ByRef oTabelle As Worksheet, Optional fromRow As Long = 1)
'    Dim i As Long
'    If oTabelle Is Nothing Then
'        Set oTabelle = ActiveSheet
'    End If
'
'    For i = 1 To getLastCol(oTabelle, fromRow)
'        If oTabelle.Cells(fromRow, i) <> "" Then
'            getFirstCol = i
'            Exit For
'        End If
'    Next i
'
'    getFirstCol = i
'End Function

'array vergroessern und value hinzufuegen
Function arrayPush(ByRef aArray As Variant, value As Variant)
    Dim newLength As Long
    Dim minLength As Long
    
    'nacheinander pruefen, wegen fehlers
    If arrayIsEmpty(aArray) Then
        ReDim aArray(0)
    Else
        newLength = UBound(aArray) + 1
        minLength = LBound(aArray)
        ReDim Preserve aArray(minLength To newLength)
    End If
    
    If IsObject(value) Then
        Set aArray(UBound(aArray)) = value
    Else
        aArray(UBound(aArray)) = value
    End If
End Function

'test if an array is empty
'http://stackoverflow.com/questions/9874086/vba-dont-go-into-loop-when-array-is-empty
Function arrayIsEmpty(aArray As Variant)
    If IsArray(aArray) = False Then arrayIsEmpty = True
    
    On Error Resume Next
    If UBound(aArray) < LBound(aArray) Then
        arrayIsEmpty = True
        Exit Function
    Else
        arrayIsEmpty = False
    End If
End Function

'array um 1 verkleinern letztes Element wird zuruckgegeben
Function arrayPop(ByRef aArray As Variant)
    Dim newLength As Long
    Dim minLength As Long
    
    If IsEmpty(aArray) Or Not IsArray(aArray) Then
        arrayPop = False
    ElseIf UBound(aArray) = -1 Then
        arrayPop = False
    Else
        newLength = UBound(aArray) - 1
        minLength = LBound(aArray)
        arrayPop = aArray(UBound(aArray))
        ReDim Preserve aArray(minLength To newLength)
    End If
End Function

'value loeschen und array verkleinern
Function arrayRemoveElement(ByRef aArray, value As Variant) As Variant
    Dim iPos As Long
    
    arrayRemoveElement = False
    
    If IsEmpty(aArray) Or Not IsArray(aArray) Then 'falls nicht array
        arrayRemoveElement = False
    ElseIf UBound(aArray) = -1 Then 'falls keine elements
        arrayRemoveElement = False
    ElseIf aLength(aArray) = 1 And aArray(LBound(aArray)) = value Then 'falls nur 1 element.
        aArray = Array()
    Else
        iPos = getIndexInArrayByValue(value, aArray)
        If iPos > -1 Then
            Dim i As Long
            For i = iPos To UBound(aArray) - 1
                aArray(i) = aArray(i + 1)
            Next i
            
            ReDim Preserve aArray(LBound(aArray) To UBound(aArray) - 1)
            
            arrayRemoveElement = True
        End If
        
    End If
End Function

'eleminiert Dublikate in einem array
Public Function arrayRemoveDublicates(ByVal arr)
    Dim element As Variant

    
    If aLength(arr) <= 0 Then
        arrayRemoveDublicates = arr
        Exit Function
    End If
    
    Dim newUniqArray As Variant
    newUniqArray = Array()
    
    For Each element In arr
        If Not isInArray(newUniqArray, element) Then
            Call arrayPush(newUniqArray, element)
        End If
    Next element
    
    arrayRemoveDublicates = newUniqArray
End Function

'return subArray of aArray, with lbound = 0
'lFrom from what element (for example from second = 1), lTo to what element, for example to 4. element = 3
Function arrayGetSubArray(ByVal aArray As Variant, Optional ByVal lFrom As Variant = -1, Optional ByVal lTo As Variant = -1, Optional ByVal lStartBound As Variant = -1)
    Dim lBnd As Long, ubnd As Long
    lBnd = LBound(aArray)
    ubnd = UBound(aArray)
    
    
    
    If lFrom = -1 Then
        lFrom = lBnd
    Else
        lFrom = lBnd + lFrom
    End If
    
    
    If lTo = -1 Then
        lTo = ubnd
    Else
        lTo = lBnd + lTo
    End If
    
    
    If IsEmpty(aArray) Or Not IsArray(aArray) Then
        arrayGetSubArray = aArray
        Exit Function
    End If
    
    
    If lFrom >= lBnd And lTo <= ubnd And lFrom <= lTo Then
        Dim tmpArray As Variant
        Dim i As Long
        ReDim tmpArray(lFrom To lTo)
        
        For i = lFrom To lTo
            tmpArray(i) = aArray(i)
        Next i
        
        aArray = tmpArray
    End If
    
    'startbound setzen
    If lStartBound >= 0 Then
        Call arrayChangeBounds(aArray, lStartBound)
    End If
    
    
    arrayGetSubArray = aArray
End Function

'changes bounds of array.
'todo optional vars
Function arrayChangeBounds(ByRef aArray As Variant, lStartBound As Variant, Optional lEndBound As Variant = -1)
    If Not IsArray(aArray) Then
        arrayChangeBounds = aArray
        Exit Function
    End If
        
    If lEndBound = -1 Then lEndBound = lStartBound + aLength(aArray) - 1
        
    If lStartBound < 0 Or lEndBound < 0 Or lStartBound > lEndBound Then
        arrayChangeBounds = aArray
        Exit Function
    End If
    
    
    If IsEmpty(aArray) Then
        ReDim Preserve aArray(lStartBound To lEndBound)
        arrayChangeBounds = aArray
        Exit Function
    End If
    
    Dim i As Long
    Dim lBnd As Long
    Dim tmpArray As Variant
    ReDim tmpArray(lStartBound To lEndBound)
    
    'werte nacheinanden kopieren
    lBnd = LBound(aArray)
    For i = lBnd To UBound(aArray)
        'nur falls es noch in tmpArray platz gibt
        If (i - lBnd <= lEndBound - lStartBound) Then
            If IsObject(aArray(i)) Then
                Set tmpArray(lStartBound + i - lBnd) = aArray(i)
            Else
                tmpArray(lStartBound + i - lBnd) = aArray(i)
            End If
            
            
        End If
    Next i
    
    aArray = tmpArray
    arrayChangeBounds = tmpArray
End Function


'array in eine Collection umwandeln
Function arrayToCollection(aArray As Variant) As Collection
    If aArray = Empty Then Exit Function
    Dim colReturn As New Collection
    Dim wert As Variant
    For Each wert In aArray
        colReturn.Add (wert)
    Next wert
    Set arrayToCollection = colReturn
End Function

'Collection in ein array umwandeln
Function collectionToArray(colCollection As Collection) As Variant
    If colCollection Is Nothing Then Exit Function
    Dim wert As Variant
    Dim aReturn() As Variant
    aReturn = Array()
    For Each wert In colCollection
        Call arrayPush(aReturn, wert)
    Next wert
    collectionToArray = aReturn
End Function

'modify 2dArray to Array Of Arrays
Public Function array2dToArrOfArrs(aArray As Variant)
    Dim i As Long, j As Long
    Dim retArray As Variant
    Dim lbound1 As Long
    Dim ubound1 As Long
    Dim lbound2 As Long
    Dim ubound2 As Long
    
    
    If Not IsArray(aArray) Then
        array2dToArrOfArrs = aArray
        Exit Function
    End If
    
    If aLength(aArray) = 0 Then
        array2dToArrOfArrs = Array()
        Exit Function
    End If
    
    lbound1 = LBound(aArray, 1)
    ubound1 = UBound(aArray, 1)
    
    ReDim retArray(lbound1 To ubound1)
    For i = lbound1 To ubound1
        Dim tmpArray As Variant
        lbound2 = LBound(aArray, 2)
        ubound2 = UBound(aArray, 2)
        ReDim tmpArray(lbound2 To ubound2)
        For j = lbound2 To ubound2
            tmpArray(j) = aArray(i, j)
        Next j

        retArray(i) = tmpArray
        
    Next i
    
    array2dToArrOfArrs = retArray
End Function

'concatenates 2 arrays, array1 wird groesser, und kopie wird zuruckgegeben 'TODO kopie nur bei einem optional param zurueckgeben?
Public Function arrayConcat(ByRef array1 As Variant, ByRef array2 As Variant) As Variant
    Dim vResult As Variant
    Dim idx As Long
    Dim elem As Variant
    
    vResult = Array()
    
    On Error GoTo arrayConcat_Error
    
    'falls nicht arrays exit
    If Not (IsArray(array1) And IsArray(array2)) Then
        dbg "Falsche argumente bei arrayConcat"
        GoTo Final
    End If
    
    'falls einer der beiden ist leer, return den anderen
    If aLength(array1) = 0 Then
        vResult = arrayCopy(array2)
        GoTo Final
    ElseIf aLength(array2) = 0 Then
        vResult = arrayCopy(array1)
        GoTo Final
    End If
    
    vResult = arrayCopy(array1)
    
    
    idx = UBound(vResult)
    
    'array auf neue geoesse einstellen
    ReDim Preserve vResult(LBound(array1) To idx + aLength(array2))
    
    'jedes elemen von array2 einfuegen
    idx = idx + 1
    For Each elem In array2
        If IsObject(elem) Then
            Set vResult(idx) = elem
        Else
            vResult(idx) = elem
        End If
        idx = idx + 1
    Next elem
    
    
    
Final:
    arrayConcat = vResult
    
    On Error GoTo 0
    Exit Function
arrayConcat_Error:
    dbg "arrayConcat of Klassenmodul clsFuncs"
    Resume Final
End Function


'array must be sorted!!!
'binary search in array. In case nothing found returns -1
Public Function arrayBinarySearch(strArray As Variant, strSearch As String) As Long
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim lngMiddle As Long
    Dim bolInverseOrder As Boolean
    If Not IsArray(strArray) Then Exit Function
    lngFirst = LBound(strArray)
    lngLast = UBound(strArray)
    bolInverseOrder = (strArray(lngFirst) > strArray(lngLast))
    arrayBinarySearch = -1 'lngFirst - 1
    Do
        lngMiddle = (lngFirst + lngLast) \ 2
        If strArray(lngMiddle) = strSearch Then
            arrayBinarySearch = lngMiddle
            Exit Do
        ElseIf ((strArray(lngMiddle) < strSearch) Xor bolInverseOrder) Then
            lngFirst = lngMiddle + 1
        Else
            lngLast = lngMiddle - 1
        End If
    Loop Until lngFirst > lngLast
End Function

''todo better own func ?
''modify Array of Array to 2dArray. Beginns with array(1,1)
'Function ArrOfArrsToArray2d(aArray As Variant)
'    Dim retArray As Variant
'
'    If Not IsArray(aArray) Then
'        ArrOfArrsToArray2d = aArray
'        Exit Function
'    End If
'
'    If aLength(aArray) = 0 Then
'        ArrOfArrsToArray2d = Array()
'        Exit Function
'    End If
'
'    retArray = WorksheetFunction.index(aArray, 0, 0)
'
'    ArrOfArrsToArray2d = retArray
'End Function

    
    

'max of both values
Function max(value1 As Variant, value2 As Variant) As Variant
    If IsEmpty(value1) Or IsNull(value1) Then
        max = value2
        Exit Function
    ElseIf IsEmpty(value2) Or IsNull(value2) Then
        max = value1
        Exit Function
    Else
        If IsNumeric(value1) And IsNumeric(value2) Then
            If value1 > value2 Then
                max = value1
            Else
                max = value2
            End If
        End If
    End If
End Function

'min of both values
Function min(value1 As Variant, value2 As Variant) As Variant
    If IsEmpty(value1) Or IsNull(value1) Then
        min = value2
        Exit Function
    ElseIf IsEmpty(value2) Or IsNull(value2) Then
        min = value1
        Exit Function
    Else
        If IsNumeric(value1) And IsNumeric(value2) Then
            If value1 < value2 Then
                min = value1
            Else
                min = value2
            End If
        End If
    End If
End Function

'check if value is in Array, bAnyTypeCompare = true ? then check any Types (modify with Cstr before compare)
Function isInArray(aArray As Variant, vValue As Variant, Optional bAnyTypeCompare As Boolean = False, Optional bTrim As Boolean, Optional bNocase As Boolean) As Boolean
    Dim value As Variant
    Dim tmpVal As Variant
    
    isInArray = False
    
    For Each value In aArray
        If IsObject(vValue) And IsObject(value) Then
            If vValue Is value Then
                isInArray = True
                Exit Function
            End If
        End If
        
        If Not IsObject(vValue) And Not IsObject(value) Then
            tmpVal = vValue 'dont change vValue
            ' number 5 = string "5"
            If bAnyTypeCompare Then
                tmpVal = CStr(tmpVal)
                value = CStr(value)
            End If
            ' " text " = "text"
            If bTrim Then
                tmpVal = Trim(tmpVal)
                value = Trim(value)
            End If
            ' "TeXt" = "texT"
            If bNocase Then
                tmpVal = LCase(tmpVal)
                value = LCase(value)
            End If
            
            If value = tmpVal Then
                isInArray = True
                Exit For
            End If
            
        End If
    Next value
End Function

'prueft ob index gueltig ist
Function arrayCheckIndex(vArray As Variant, lIndex As Long)
    
    If lIndex <= UBound(vArray) And lIndex >= LBound(vArray) Then
        arrayCheckIndex = True
    Else
        arrayCheckIndex = False
    End If
    
End Function


''workbook finden, falls schon geoeffnet, oder oeffnen falls nicht geoeffnet
'Function findOpenedDocOrOpen(filePath As String) As Workbook
'    Dim wrkReturn As Workbook
'    Set wrkReturn = findOpenedDoc(filePath)
'    If wrkReturn Is Nothing Then
'        Set findOpenedDocOrOpen = Workbooks.Open(filePath)
'    Else
'        Set findOpenedDocOrOpen = wrkReturn
'    End If
'End Function

''sucht geoeffnete datei mit dem pfad
'Function findOpenedDoc(filePath As String, Optional bByName As Boolean, Optional bLikeName As Boolean) As Workbook
'    Dim curWkb As Workbook
'    For Each curWkb In Workbooks
'        If curWkb.FullName = filePath Then
'            Set findOpenedDoc = curWkb
'            Exit For
'        End If
'    Next curWkb
'End Function

''geoeffnetes Dok finden, dessen name ist gleich "sName" ODER falls bLikeName true ist, dessen name sName drin hat.
'Function findOpenedDocByName(sName As String, Optional bLikeName As Boolean) As Workbook
'    Dim curWkb As Workbook
'    For Each curWkb In Workbooks
'        If curWkb.Name = sName Then
'            Set findOpenedDocByName = curWkb
'            Exit For
'        ElseIf bLikeName And InStr(1, curWkb.Name, sName) > 0 Then
'            Set findOpenedDocByName = curWkb
'            Exit For
'        End If
'    Next curWkb
'End Function

''screenUpdating und Formulasberechnung Ein/Ausschalten
'Function screenBerechnungSchutz(einAus As Boolean, Optional sSchutzPw As String = Empty)
'    'reinfolge beim einschalten, ausschalten 'no need? Oo
'    If einAus Then
'        If IsEmpty(sSchutzPw) Then
'            Call blattschutz(einAus)
'        Else
'            Call blattschutz(einAus, sSchutzPw)
'        End If
'        berechnung (einAus)
'        screen (einAus)
'
'    Else
'        screen (einAus)
'        berechnung (einAus)
'        If IsEmpty(sSchutzPw) Then
'            Call blattschutz(einAus)
'        Else
'            Call blattschutz(einAus, sSchutzPw)
'        End If
'    End If
'
'End Function
'
''screenUpdating Ein/Ausschalten
'Function screen(einAus As Boolean)
'    If (CBool(einAus)) Then
'        Application.ScreenUpdating = True
'    Else
'        Application.ScreenUpdating = False
'    End If
'End Function
'
''Formulasberechnung Ein/Ausschalten
'Function berechnung(einAus As Boolean)
'    If (CBool(einAus)) Then
'        Application.Calculation = xlCalculationAutomatic
'    Else
'        Application.Calculation = xlCalculationManual
'    End If
'End Function

''open dialog to pick a file, return "" on error. 'deprecated?
'Function getFilePathDialog(sFileFilter As String, sTitle As String) As String
''Pfad zur Datei bestimmen
'    'On Error Resume Next
'    On Error GoTo ErrorHandler
'
'    getFilePathDialog = Application.GetOpenFilename(FileFilter:=sFileFilter, Title:=sTitle)
'    'todo? international?
'    If getFilePathDialog = "Falsch" Then getFilePathDialog = ""
'
'    Exit Function
'ErrorHandler:
'    getFilePathDialog = ""
'End Function

'filePicker or FolderPicker dialog. 'todo miltiselect option ?
'required Office 11? Object Library
Function DialogGetPath(Optional bFileOrFolder As Boolean = True, Optional sTitle As String, Optional filterTyp As String, _
                        Optional filterExt As String, Optional sInitial As String, Optional bMulti As Boolean = False, Optional sButtonName As String)
                        
    Dim fd As FileDialog

    'if true -> file
    If bFileOrFolder Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        'falls filter werde vorhanden
        If hatWert(filterTyp) And hatWert(filterExt) Then
            fd.Filters.Add filterTyp, filterExt
        End If
    
    'folder
    Else
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    End If
    
    'anfangspfad
    If hatWert(sInitial) Then fd.InitialFileName = sInitial
    
    'if multiselected
    If bMulti Then
        fd.AllowMultiSelect = True
    End If
    
    If hatWert(sTitle) Then
        fd.Title = sTitle
    End If
    
    'extra button title
    If hatWert(sButtonName) Then
        fd.buttonName = sButtonName
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
                Call arrayPush(vrtResult, vrtSelectedItem)
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

''blattschutz aufheben, oder wieder einschalten
'Public Sub blattschutz(bYesNo As Boolean, Optional sSchutzPw As String = "2012", Optional ByRef wkbWorkbook As Workbook)
'    On Error Resume Next
'    Dim i As Integer
'
'    If wkbWorkbook Is Nothing Then Set wkbWorkbook = ThisWorkbook
'
'    With wkbWorkbook
'    If bYesNo Then
'        For i = 1 To .Worksheets.Count
'            With Worksheets(i)
'            .Protect Password:=sSchutzPw, UserInterfaceOnly:=True
'            '.EnableSelection = xlUnlockedCells
'            '.EnableSelection = xlNoRestrictions
'            End With
'        Next i
'    Else
'        For i = 1 To .Worksheets.Count
'            With Worksheets(i)
'            .Unprotect Password:=sSchutzPw
'            '.EnableSelection = xlNoRestrictions
'            End With
'        Next i
'    End If
'    End With
'End Sub

'prueft ob txtcheck ein wert hat, also ob es nicht 0 und nicht "" ist.
'bTrim ob auch trim eingesetz werden muss /default ja
Public Function hatWert(Optional ByVal vCheck As Variant, Optional bTrim As Boolean = True)
    If Not IsMissing(vCheck) Then
        
        If IsNull(vCheck) Then
            hatWert = False
            Exit Function
        End If
        
        If IsObject(vCheck) Then
            If vCheck Is Nothing Then
                hatWert = False
            Else
                hatWert = True
            End If
            Exit Function
        End If
        
        If IsArray(vCheck) Then
            hatWert = True
            Exit Function
        End If
        
        
        If IsEmpty(vCheck) Then
            hatWert = False
            Exit Function
        End If
        
        If bTrim Then
            vCheck = Trim(vCheck)
        End If
        
        If Not (CStr(vCheck) = "" Or CStr(vCheck) = "0" Or Len(CStr(vCheck)) = 0) Then
            hatWert = True
            Exit Function
        End If
        hatWert = False
    End If
End Function

'prueft hatWert fuer alle stings in Array
Public Function hatWertArray(ByVal aTxtCheck As Variant, Optional bTrim As Boolean = True)
    Dim element As Variant
    hatWertArray = True
    For Each element In aTxtCheck
        If Not hatWert(element, bTrim) Then
            hatWertArray = False
            Exit For
        End If
    Next element
End Function

'wenn value keinen wert hat dann arg2 zuruckgeben, sonst das value (arg1)
Function ifNoValue(vValue As Variant, vInCaseNoValue As Variant)
    If hatWert(vValue) Then
        If IsObject(vValue) Then
            Set ifNoValue = vValue
        Else
            ifNoValue = vValue
        End If
    Else
        If IsObject(vInCaseNoValue) Then
            Set ifNoValue = vInCaseNoValue
        Else
            ifNoValue = vInCaseNoValue
        End If
        
    End If
End Function


''calculates all sheet in Workbook
'Public Function refreshAllSheets(Optional wkbWorkbook As Workbook)
'    Dim sheet As Variant
'    If wkbWorkbook Is Nothing Then Set wkbWorkbook = ThisWorkbook
'
'    For Each sheet In wkbWorkbook.Sheets
'        sheet.Calculate
'    Next sheet
'End Function

''calculate array of tables objects
'Public Function refreshTables(aTables As Variant)
'    Dim wksTable As Variant
'    For Each wksTable In aTables
'        wksTable.Calculate
'    Next wksTable
'End Function

'returns index of element first found in array
Function getIndexInArrayByValue(vWert As Variant, aArray As Variant)
    Dim i As Long
    
    getIndexInArrayByValue = -1
    For i = LBound(aArray) To UBound(aArray)
        If IsObject(vWert) And IsObject(aArray(i)) Then
            If aArray(i) Is vWert Then
                getIndexInArrayByValue = i
                Exit For
            End If
        ElseIf Not IsObject(vWert) And Not IsObject(aArray(i)) Then
            If UCase(aArray(i)) = UCase(vWert) Then
                getIndexInArrayByValue = i
                Exit For
            End If
        End If
        
    Next i
End Function

'returns array of indexes of element found in array
Public Function getAllIndexesInArrayByValue(vWert As Variant, aArray As Variant, _
    Optional bAnyTypeCompare As Boolean = False, Optional bTrim As Boolean = False, Optional bNocase As Boolean = True)
    Dim i As Long
    Dim retArray As Variant
    Dim value1 As Variant
    Dim value2 As Variant
    
    retArray = Array()
    
    For i = LBound(aArray) To UBound(aArray)
        If IsObject(vWert) And IsObject(aArray(i)) Then
            If aArray(i) Is vWert Then
                Call arrayPush(retArray, i)
            End If
        ElseIf Not IsObject(vWert) And Not IsObject(aArray(i)) Then
            
            value1 = vWert
            value2 = aArray(i)
            
            ' number 5 = string "5"
            If bAnyTypeCompare Then
                value1 = CStr(value1)
                value2 = CStr(value2)
            End If
            ' " text " = "text"
            If bTrim Then
                value1 = Trim(value1)
                value2 = Trim(value2)
            End If
            ' "TeXt" = "texT"
            If bNocase Then
                value1 = LCase(value1)
                value2 = LCase(value2)
            End If
            
            If value1 = value2 Then

                Call arrayPush(retArray, i)
            End If
        End If
        
    Next i
    getAllIndexesInArrayByValue = retArray
End Function


'check if file exists or directory. For dir: c:\xx\ need slash at the end!
Public Function fileExists(sPath As String)
    If sPath = vbNullString Then
        fileExists = False
        Exit Function
    End If
    
    If dir(sPath, vbDirectory) <> vbNullString Then
        fileExists = True
    End If
    
    Exit Function
EarlyExit:
    fileExists = False
    On Error GoTo 0
End Function

'check if all files exists
Public Function fileExistsArray(aPaths As Variant)
    On Error GoTo ErrorHandler
    Dim sPath As Variant
    
    If IsArray(aPaths) Then
        fileExistsArray = True
        For Each sPath In aPaths
            If Not fileExists(CStr(sPath)) Then
                fileExistsArray = False
                Exit For
            End If
        Next sPath
    End If
    
    Exit Function
ErrorHandler:
    fileExistsArray = False
End Function

'quicksort, wegen kaskadenrekursion nur fuer kleine array benutzen..
Sub QuickSort(ByRef sArray As Variant, Optional ByVal MinElem As Long = -1, Optional MaxElem As Long = -1)
'
' QuickSort()
'
' Sortieren eines Arrays mit dem QuickSort-Algorithmus, dem wohl schnellsten
' Sortieralgorithmus von Welt.
'
' IN:   sArray      Array das sortiert werden soll
'       MinElem     erstes Element des Arrays (oder Teil-Arrays) default: LBound(sArray)
'       MaxElem     letztes Element des Arrays (oder Teil-Arrays) default: UBound(sArray)
'

If IsMissing(MinElem) Or MinElem = -1 Then MinElem = LBound(sArray)
If IsMissing(MaxElem) Or MaxElem = -1 Then MaxElem = UBound(sArray)

Dim Mitte As Long
Dim vDummy As Variant
Dim vMitte As Variant
Dim i As Long, j As Long
    '
    ' Abbruchbedingung der Rekursion prüfen
    '
    If MinElem > MaxElem Then
        '
        ' Rekursion beenden
        '
        Exit Sub
    End If
    '
    ' Ermittlung der Mitte des Arrays
    '
    Mitte = (MinElem + MaxElem) \ 2
    vMitte = sArray(Mitte)
    '
    ' Für die Prüfung der linken und rechten
    ' Seite die Zähler initialisieren
    '
    i = MinElem
    j = MaxElem
    Do
        '
        ' Von links bis zur Mitte prüfen
        '
        Do While sArray(i) < vMitte
            i = i + 1
        Loop
        '
        ' Von rechts bis zur Mitte prüfen
        '
        Do While sArray(j) > vMitte
            j = j - 1
        Loop
        
        If i <= j Then
            '
            ' Die beiden gefundenen, falsch einsortierten
            ' Elemente vertauschen
            '
            vDummy = sArray(j)
            sArray(j) = sArray(i)
            sArray(i) = vDummy
            '
            ' Prüfung bei der nächsten Position fortsetzen
            '
            i = i + 1
            j = j - 1
        End If
        
    Loop Until i > j
    '
    ' Rekursiver Aufruf mit den Teil-Arrays
    '
    QuickSort sArray, MinElem, j
    QuickSort sArray, i, MaxElem
End Sub

'sorts first elements of multidim. array(x,1) 0,1
Sub QuickSort2d(ByRef sArray As Variant, ByVal MinElem As Long, MaxElem As Long)
'
' QuickSort()
'
' Sortieren eines Arrays mit dem QuickSort-Algorithmus, dem wohl schnellsten
' Sortieralgorithmus von Welt.
'
' IN:   sArray      Array das sortiert werden soll
'       MinElem     erstes Element des Arrays (oder Teil-Arrays)
'       MaxElem     letztes Element des Arrays (oder Teil-Arrays)
'
Dim Mitte As Long
Dim vDummy As Variant
Dim vDummy2 As Variant
Dim vMitte As Variant
Dim i As Long, j As Long
    '
    ' Abbruchbedingung der Rekursion prüfen
    '
    If MinElem > MaxElem Then
        '
        ' Rekursion beenden
        '
        Exit Sub
    End If
    '
    ' Ermittlung der Mitte des Arrays
    '
    Mitte = (MinElem + MaxElem) \ 2
    vMitte = sArray(Mitte, 0)
    '
    ' Für die Prüfung der linken und rechten
    ' Seite die Zähler initialisieren
    '
    i = MinElem
    j = MaxElem
    Do
        '
        ' Von links bis zur Mitte prüfen
        '
        Do While sArray(i, 0) < vMitte
            i = i + 1
        Loop
        '
        ' Von rechts bis zur Mitte prüfen
        '
        Do While sArray(j, 0) > vMitte
            j = j - 1
        Loop
        
        If i <= j Then
            '
            ' Die beiden gefundenen, falsch einsortierten
            ' Elemente vertauschen
            '
            vDummy = sArray(j, 0)
            vDummy2 = sArray(j, 1)
            sArray(j, 0) = sArray(i, 0)
            sArray(j, 1) = sArray(i, 1)
            sArray(i, 0) = vDummy
            sArray(i, 1) = vDummy2
            '
            ' Prüfung bei der nächsten Position fortsetzen
            '
            i = i + 1
            j = j - 1
        End If
        
    Loop Until i > j
    '
    ' Rekursiver Aufruf mit den Teil-Arrays
    '
    QuickSort2d sArray, MinElem, j
    QuickSort2d sArray, i, MaxElem
End Sub

'Heap Sort. 'http://www.vbforums.com/showpost.php?p=2909250&postcount=7
Public Function HeapSort(ByRef pvarArray As Variant)
        Dim i As Long
        Dim iMin As Long
        Dim iMax As Long
        Dim varSwap As Variant
       
        iMin = LBound(pvarArray)
        iMax = UBound(pvarArray)
        For i = (iMax + iMin) \ 2 To iMin Step -1
            HeapSortHelper pvarArray, i, iMin, iMax
        Next i
        For i = iMax To iMin + 1 Step -1
            varSwap = pvarArray(i)
            pvarArray(i) = pvarArray(iMin)
            pvarArray(iMin) = varSwap
            HeapSortHelper pvarArray, iMin, iMin, i - 1
        Next i
        HeapSort = pvarArray
    End Function
    Private Function HeapSortHelper(ByRef pvarArray As Variant, ByVal i As Long, iMin As Long, iMax As Long)
        Dim lngLeaf As Long
        Dim varSwap As Variant
       
        Do
            lngLeaf = i + i - (iMin - 1)
            Select Case lngLeaf
                Case Is > iMax: Exit Do
                Case Is < iMax: If pvarArray(lngLeaf + 1) > pvarArray(lngLeaf) Then lngLeaf = lngLeaf + 1
            End Select
            If pvarArray(i) > pvarArray(lngLeaf) Then Exit Do
            varSwap = pvarArray(i)
            pvarArray(i) = pvarArray(lngLeaf)
            pvarArray(lngLeaf) = varSwap
            i = lngLeaf
        Loop
    End Function
    
'http://www.freevbcode.com/ShowCode.asp?ID=1086
Public Function IsAlphaBetical(TestString As String) As Boolean
    
    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
    
    'returns true if all characters in a string are alphabetical
    'returns false otherwise or for empty string
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[A-Za-z]" Then Exit Function
        Next
    
    IsAlphaBetical = True
    End If
    
End Function
'http://www.freevbcode.com/ShowCode.asp?ID=1086
Public Function IsAlphaNumeric(TestString As String) As Boolean

    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
    
    'returns true if all characters in a string are alphabetical
    '   or numeric
    'returns false otherwise or for empty string
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9A-Za-z]" Then Exit Function
        Next
    
    IsAlphaNumeric = True
    End If
    
End Function
'http://www.freevbcode.com/ShowCode.asp?ID=1086
Public Function IsNumericOnly(TestString As String) As Boolean
    
    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
    
    'returns true if all characters in string are numeric
    'returns false otherwise or for empty string
    
    'this is different than VB's isNumeric
    'isNumeric returns true for something like 90.09
    'This function will return false
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9]" Then Exit Function
        Next
    
    IsNumericOnly = True
    End If
    
End Function

'numeric type BEGIN http://www.ms-office-forum.net/forum/showthread.php?t=129184 ---------
'modified by me
Public Function isDouble(var As Variant) As Boolean
   If IsNull(var) Then
      isDouble = False
   Else
      isDouble = ((VBA.VarType(var) = vbDouble) Or Me.isSingle(var))
   End If
End Function

Public Function isSingle(var As Variant) As Boolean
   If IsNull(var) Then
      isSingle = False
   Else
      isSingle = ((VBA.VarType(var) = vbSingle) Or Me.isLong(var))
   End If
End Function

Public Function isLong(var As Variant) As Boolean
   If IsNull(var) Then
      isLong = False
   Else
      isLong = ((VBA.VarType(var) = vbLong) Or Me.isInteger(var))
   End If
End Function

Public Function isInteger(var As Variant) As Boolean
   If IsNull(var) Then
      isInteger = False
   Else
      isInteger = (VBA.VarType(Val(var)) = vbInteger Or Me.isByte(var))
   End If
End Function

Public Function isByte(var As Variant) As Boolean
   If IsNull(var) Then
      isByte = False
   Else
      isByte = ((VBA.VarType(Val(var)) = vbByte))
   End If
End Function
'numeric type END ------------------------------------------------------------------------

Public Function rndRange(num1 As Variant, num2 As Variant)
    
    If IsNumeric(num1) And IsNumeric(num2) Then
        rndRange = Int(((num2 - num1 + 1) * Rnd) + num1)
    Else
        Debug.Print "f.rndRange, nicht numerische arguments"
    End If
    
End Function

'tries to modify to date, if failed returns input, or 2nd arg of input
Public Function CdateTry(something As Variant, Optional inFailCaseReturn As Variant) As Variant
    Dim sReplaced As String
    Dim delimiter As Variant 'string
    Dim aSplit As Variant
    Dim sDay As String
    
    If TypeName(something) = "Date" Or IsDate(something) Then
        CdateTry = CDate(something)
        Exit Function
    Else
        On Error Resume Next
        sReplaced = CStr(something)
        For Each delimiter In Array(".", ",", "'")
            sReplaced = Replace(sReplaced, delimiter, "/")
        Next delimiter
        
        'replace month with day, for formats dd.mm.yyyy
        aSplit = Split(sReplaced, "/")
        sDay = aSplit(0)
        aSplit(0) = aSplit(1)
        aSplit(1) = sDay
        sReplaced = Join(aSplit, "/")
        
        'try to parse date now
        CdateTry = CDate(sReplaced)
        
        If IsDate(CdateTry) Then Exit Function
        On Error GoTo 0
    End If

    If IsMissing(inFailCaseReturn) Then
        If IsObject(something) Then
            Set CdateTry = something
        Else
            CdateTry = something
        End If
    Else
        If IsObject(inFailCaseReturn) Then
            Set CdateTry = inFailCaseReturn
        Else
            CdateTry = inFailCaseReturn
        End If
    End If
End Function

'tries to modify to Long, if failed returns input, or 2nd arg of input
Public Function ClngTry(something As Variant, Optional inFailCaseReturn As Variant) As Variant
    If isLong(something) Then
        ClngTry = CLng(something)
        Exit Function
    ElseIf Not IsEmpty(something) And Not IsNull(something) Then
        On Error Resume Next
        ClngTry = CLng(something)
        If isLong(ClngTry) Then Exit Function
        On Error GoTo 0
    End If
    
    
     If IsMissing(inFailCaseReturn) Then
        If IsObject(something) Then
            Set ClngTry = something
        Else
            ClngTry = something
        End If
    Else
        If IsObject(inFailCaseReturn) Then
            Set ClngTry = inFailCaseReturn
        Else
            ClngTry = inFailCaseReturn
        End If
    End If
End Function

'bestaetigungs frage
Public Function ask(sString As String, Optional sTitle = " ") As Boolean
  
    If MsgBox(sString, vbYesNo, sTitle) = vbYes Then
        ask = True
    Else
        ask = False
    End If
End Function


'prints some info in debug window, with date
Public Sub dbg(sInfo As String)
   Debug.Print Format(Now, "hh:mm:ss") & " - " & sInfo
End Sub

' Umlaute und Sonderzeichen verschlüsseln
'
' UsePlusRatherThanHexForSpace:
'  False:  Leerzeichen als %32 verschlüsseln
'  True :  Leerzeichen als + verschlüsseln
'http://www.vbarchiv.net/tipps/tipp_139-urlencode-und-urldecode.html  :: Matthias Zirngibl
Public Function URLEncode(StringToEncode As String, Optional _
  UsePlusRatherThanHexForSpace As Boolean = False) As String
 
  Dim TempAns As String
  Dim CurChr As Integer
 
  CurChr = 1
  Do Until CurChr - 1 = Len(StringToEncode)
    Select Case Asc(Mid$(StringToEncode, CurChr, 1))
      Case 48 To 57, 65 To 90, 97 To 122
        TempAns = TempAns & Mid$(StringToEncode, CurChr, 1)
      Case 32
        If UsePlusRatherThanHexForSpace = True Then
          TempAns = TempAns & "+"
        Else
          TempAns = TempAns & "%" & Hex(32)
        End If
      Case Else
        TempAns = TempAns & "%" & Hex(Asc(Mid$(StringToEncode, _
          CurChr, 1)))
    End Select
    CurChr = CurChr + 1
  Loop
  URLEncode = TempAns
End Function
 
' Entschlüsseln von mit URLEncode
' verschlüsselten Zeichenketten
'http://www.vbarchiv.net/tipps/tipp_139-urlencode-und-urldecode.html :: Matthias Zirngibl
Public Function URLDecode(StringToDecode As String) As String
  Dim TempAns As String
  Dim CurChr As Integer
 
  CurChr = 1
  Do Until CurChr - 1 = Len(StringToDecode)
    Select Case Mid$(StringToDecode, CurChr, 1)
      Case "+"
        TempAns = TempAns & " "
      Case "%"
        TempAns = TempAns & Chr(Val("&h" & _
          Mid$(StringToDecode, CurChr + 1, 2)))
        CurChr = CurChr + 2
      Case Else
        TempAns = TempAns & Mid$(StringToDecode, CurChr, 1)
    End Select
    CurChr = CurChr + 1
  Loop
  URLDecode = TempAns
End Function

'get path of string
Public Function getPath(sPath As String)
    getPath = Left$(sPath, InStrRev(sPath, "\"))
End Function
'get filename of sPath path
Public Function getFilename(sPath As String)
    getFilename = Mid$(sPath, InStrRev(sPath, "\") + 1)
End Function

'zeit als string oder Datum in SQL save string modifizieren
Public Function SQLString_formatTimeToSQLString(dDate As Variant) As String
    Dim saveDate As Date
    Dim testDate As Variant
    testDate = CdateTry(dDate, -1)
    
    If testDate <> -1 Then
        saveDate = CDate(testDate)
        SQLString_formatTimeToSQLString = Format(saveDate, "\#yyyy-mm-dd HH\:MM\:SS\#")
        Exit Function
    End If
    
    SQLString_formatTimeToSQLString = ""
End Function

'formats number for SQL readable format
Public Function SQLString_formatFloatNumber(num As Variant) As String
    If Not IsNumeric(num) Then
        SQLString_formatFloatNumber = CStr(num)
    End If
    
    SQLString_formatFloatNumber = Replace(CStr(num), ",", ".")
End Function

'simple add WHERE sClause clause, todo: if ORDER BY/Gro ...  is in string
Public Function SQLString_addWhereClause(ByVal sSQLString As String, sWhatToAdd As String, Optional sOperator As String = "AND")
    Dim iWherePosInSqlString As String
    
    'pruefen ob Operator richtig ist
    If StrComp(sOperator, "AND", vbTextCompare) <> 0 And StrComp(sOperator, "OR", vbTextCompare) <> 0 Then
        SQLString_addWhereClause = sSQLString
            dbg "SQLString_addWhereClause, falsche Operator? : " & sOperator
        Exit Function
    End If
    
    sSQLString = Trim(sSQLString)
    
    'todo check in case "select * from tblXx where"
    'Falls " WHERE " nicht vorhanden
    iWherePosInSqlString = InStr(1, sSQLString, " WHERE ", vbTextCompare)
    If iWherePosInSqlString = 0 Then
        sSQLString = sSQLString & " WHERE " & sWhatToAdd
    'falls vorhanden
    Else
        sSQLString = sSQLString & " " & UCase(sOperator) & " " & sWhatToAdd
    End If
    
    
    SQLString_addWhereClause = sSQLString
End Function

'filters all letters except numeric
Public Function strFilterNumeric(sString As String) As String
    Dim letterPos As Long
    Dim retString As String
    Dim letter As String
    
    For letterPos = 1 To Len(sString)
        letter = Mid(sString, letterPos, 1)
        If IsNumeric(letter) Then
            retString = retString & letter
        End If
    Next letterPos
    
    strFilterNumeric = retString
End Function

'returns file content
Function getFileContent(name As String) As String
    Dim intUnit As Integer
     
    On Error GoTo ErrGetFileContent
    intUnit = FreeFile
    Open name For Input As intUnit
    getFileContent = Input(LOF(intUnit), intUnit)
ErrGetFileContent:
    Close intUnit
    Exit Function
End Function


'returns string that is matched by sSearchRegexin sContent text
Function regEx_getStringMatched(sSearchRegex As String, sContent As String)
    Dim objRegex As Object
    Dim objRegM As Object
    Set objRegex = CreateObject("vbscript.regexp")
    
    With objRegex
        .IgnoreCase = True
        .Pattern = sSearchRegex
        If .test(sContent) Then
            Set objRegM = .Execute(sContent)
            regEx_getStringMatched = objRegM(0).submatches(0)
        Else
            regEx_getStringMatched = ""
        End If
    End With
End Function


'return array of strings, with file paths
'bOnlyNames if you need names of files only
Function getDirList(ByVal sFolder As String, _
                    Optional sFilePattern As String = "*.*", _
                    Optional faFileAttr As VbFileAttribute = vbNormal, _
                    Optional bOnlyNames = False) As String()
    Dim sNewFile As String
    Dim retArray() As String

    
    
    If Right(sFolder, 1) <> "\" Then
        sFolder = sFolder & "\"
    End If
    
    sNewFile = dir(sFolder & sFilePattern, faFileAttr)
    
    While sNewFile <> ""
        If sNewFile <> "." And sNewFile <> ".." Then
            
            ' Mit bit-weisem Vergleich sicherstellen, daß Name1 ein
            ' Verzeichnis ist.
            If (GetAttr(sFolder & sNewFile) And faFileAttr) = faFileAttr Then
                If bOnlyNames Then
                    Call Me.arrayPush(retArray, sNewFile)
                Else
                    Call Me.arrayPush(retArray, sFolder & sNewFile)
                End If
            End If
        
            
        End If
        sNewFile = dir()
    Wend
    
    getDirList = retArray
End Function

