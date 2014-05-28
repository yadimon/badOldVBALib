VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTableHandler
' Author    : Dmitry Gorelenkov
' Date      : 08.09.2012
' Purpose   : Tabellen erstellen, mit Daten fuellen
' Requires  : clsMyCollection, clsErrorCollector, clsFuncs, clsFuncsExl, clsErrorHandler
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------

Option Explicit
Private m_rngMaMuster As Range
Private m_mcSettings As clsMyCollection
Private m_mcMA As clsMyCollection
Private m_wksWorkingSheet As Worksheet
Private m_ecErrorCollector As clsErrorCollector
Private m_aMaMusterValues As Variant
Private m_wkbReporting As Workbook
Private f As clsFuncs
Private fExl As clsFuncsExl
Private ehError As clsErrorHandler
Private m_mcCurrentSettings As clsMyCollection
Private m_sReportingFuncPrefix As String

'reporting file
Public Property Set reporting(objreporting As Workbook)
    Set m_wkbReporting = objreporting
End Property


'muster ref. und wertenArray speichern
Property Set muster(ByRef rngMuster As Range)
    Set m_rngMaMuster = rngMuster
    m_aMaMusterValues = rngMuster.value2
End Property

Property Set settings(ByRef vSettingsP As Variant)
    'falls schon als clsMyCollection uebergeben
    If TypeOf vSettingsP Is clsMyCollection Then
        Set m_mcSettings = f.objCopy(vSettingsP)
    'falls range, von range auslesen
    ElseIf TypeOf vSettingsP Is Range Then
        Dim castToRangeVar As Range
        Set castToRangeVar = vSettingsP
        Call m_mcSettings.loadFromRange(castToRangeVar)
    End If
End Property

Property Get settings()
    Set settings = m_mcSettings
End Property


'mitarbeiter mycollection speichern
Property Set mitarbeiter(ByRef vMA As Variant)

    If TypeOf vMA Is Range Then
        vMA = f.ArrayXDimTo1Dim(vMA.value2)
    End If
    
    If IsArray(vMA) Then
        Dim MA As Variant
        For Each MA In vMA
            Call m_mcMA.add(MA, CStr(MA))
        Next MA
    Else
        f.dbg "mitarbeiter koennten nicht gespeichert werden"
    End If
End Property

Property Set WorkingSheet(ByRef wksSheet As Worksheet)
    Set m_wksWorkingSheet = wksSheet
End Property

Public Property Get WorkingSheet() As Worksheet
    Set WorkingSheet = m_wksWorkingSheet
End Property


Property Set ErrorCollector(ByRef m_ecErrorCollectorP As clsErrorCollector)
        Set m_ecErrorCollector = m_ecErrorCollectorP
End Property

'alles erstellen, farben formate und static werte einfuegen
Public Function createTable(Optional wksTable As Worksheet, Optional anzahlKamp As Integer = -1, Optional bFillLinks As Boolean = False)
    If anzahlKamp = -1 Then
        anzahlKamp = m_mcSettings.getValue("Kampagnen")
    End If
    Dim i As Integer
    For i = 1 To anzahlKamp
        Call createKampagne(i, bFillLinks)
    Next i
    
End Function

'nur eine kampagne (spalte) einfuegen
Public Function createKampagne(iKampNr As Integer, Optional bFillLinks As Boolean = False)
    Dim aMA As Variant
    Dim MA As Variant
    
    aMA = m_mcMA.getKeys()
    For Each MA In aMA
        Call createMA(iKampNr, CStr(MA), bFillLinks)
    Next MA
    
    
'        If f.aLength(aMA) > 0 Then
'            Dim rngThisMaRange As Range
'            Set rngThisMaRange = getMaRange(iKampNr, CStr(aMA(0)))
'            rngThisMaRange.Cells(1, rngThisMaRange.Columns.Count - 2).Value = "Status"
'            rngThisMaRange.Cells(1, rngThisMaRange.Columns.Count - 1).Value = Format(Now(), "dd.mm.yyyy hh:nn")
'        End If
    
End Function

'nur einen mitarbeiter einfuegen
Public Function createMA(iKampNr As Integer, sMA As String, Optional bFillLinks As Boolean = False)
    Dim rngThisMaRange As Range
    Set rngThisMaRange = getMaRange(iKampNr, sMA)
    
    'set current placeholder values
    m_mcCurrentSettings("<MA>") = sMA
    m_mcCurrentSettings("<K>") = iKampNr
    m_mcCurrentSettings("<rngCurrentMaRange>") = rngThisMaRange
    
    
    m_rngMaMuster.Copy
    rngThisMaRange.PasteSpecial Paste:=xlPasteFormats
    rngThisMaRange.PasteSpecial Paste:=xlPasteColumnWidths
    
    'copy row heights
    Dim iRow As Long
    With m_rngMaMuster
        For iRow = 1 To .Rows.Count
            rngThisMaRange.Rows(iRow).RowHeight = .Rows(iRow).RowHeight
        Next iRow
    End With

    
    Dim aValuesWOreferences As Variant
    
    
    aValuesWOreferences = fillMaValuesArray(m_aMaMusterValues, bFillLinks)
    'aValuesWOreferences = replacePlaceHolder(aValuesWOreferences)
    
    rngThisMaRange.value2 = aValuesWOreferences
    

'    'array nur mit wirklichen werten kopieren
'    ReDim ValuesWOreferences(1 To UBound(m_aMaMusterValues), 1 To UBound(m_aMaMusterValues, 2))
'
'
'
'    ValuesWOreferences = removeReferences(m_aMaMusterValues)
End Function

'fill MA Values Array with data, from reporting. or with "" depending bFillFromReporting
Public Function fillMaValuesArray(ByVal aArray2d As Variant, bFillFromReporting As Boolean, Optional ByRef wkbReporting As Workbook)
    Dim i As Integer
    Dim i2 As Integer
    Dim retArray As Variant
    Dim sMusterString As String
    retArray = Array()
    
    'falls keine reporting uebergeben und in der klasse schon ein reporting gespeichert ist,
    'den nehmen
    If wkbReporting Is Nothing And Not m_wkbReporting Is Nothing Then
        Set wkbReporting = m_wkbReporting
    End If
    
    Dim aCurrentMaArray As Variant
    aCurrentMaArray = m_mcCurrentSettings("<rngCurrentMaRange>").value2
    
    ReDim retArray(1 To UBound(m_aMaMusterValues), 1 To UBound(m_aMaMusterValues, 2))
    
    For i = 1 To UBound(m_aMaMusterValues)
        For i2 = 1 To UBound(m_aMaMusterValues, 2)
            sMusterString = replacePlaceHolder(CStr(m_aMaMusterValues(i, i2)))
            
            'if link
            If checkIfLink(sMusterString) Then
                'if need to get data from reporting
                If bFillFromReporting And Not wkbReporting Is Nothing Then
                    retArray(i, i2) = loadFromReporting(sMusterString, wkbReporting)
                'if dont need, than set empty
                Else
                    'retArray(i, i2) = ""
                    'retArray(i, i2) = sMusterString
                    'vorhandene werte drin lassen
                    retArray(i, i2) = aCurrentMaArray(i, i2)
                End If
            'if simple value, copy
            Else
                retArray(i, i2) = sMusterString
            End If
        Next i2
    Next i
    
    fillMaValuesArray = retArray
End Function

'prueft ob sLink als link erkannt werden kann
Public Function checkIfLink(sLink As String)
    Dim aValues As Variant
    On Error GoTo checkIfLink_Error
    Dim maxRows As Long
    Dim maxCols As Long
    Dim colnum As String
    Dim rownum As String
    
    checkIfLink = False
    If Left(sLink, 20) = m_sReportingFuncPrefix Then
        checkIfLink = True
        Exit Function
    End If
    
    aValues = Split(sLink, ",")
    If f.aLength(aValues) = 3 Then
        maxRows = ThisWorkbook.Sheets(1).Rows.Count
        maxCols = ThisWorkbook.Sheets(1).Columns.Count
        
        
        colnum = fExl.LetterToCol(Trim(aValues(1)))
        rownum = Trim(aValues(2))
        
        If IsNumeric(colnum) And IsNumeric(rownum) Then
            If colnum <= maxCols And rownum <= maxRows Then
                checkIfLink = True
            End If
        End If
    End If
    
    On Error GoTo 0
    Exit Function

checkIfLink_Error:
    ehError.catchErr ("checkIfLink of Klassenmodul clsTableHandler")
    checkIfLink = False
End Function

'ersetzt bestimmte platzhalter mit uebergebenen werten werten
Public Function replacePlaceHolder(sString As String) As String
    Dim sPlaceholder As Variant
    
    'todo load from list/range if in it.
    m_mcCurrentSettings("<LASTDAY>") = m_sReportingFuncPrefix & "<LASTDAY>"
    m_mcCurrentSettings("<DATUM>") = m_sReportingFuncPrefix & "<DATUM>"
    
    For Each sPlaceholder In m_mcCurrentSettings.getKeys
        If InStr(1, sString, sPlaceholder, vbTextCompare) Then
            sString = Replace(sString, sPlaceholder, m_mcCurrentSettings(sPlaceholder))
        End If
    Next sPlaceholder
    
    
    replacePlaceHolder = sString
    
End Function

'erkennt link, und ladet daten from wkbReporting
Public Function loadFromReporting(sString As String, Optional wkbReporting As Workbook)
    Dim aParams As Variant
    Dim sWorksheet As String
    Dim lRow As Long
    Dim lCol As Long
    
    If wkbReporting Is Nothing And Not m_wkbReporting Is Nothing Then
        Set wkbReporting = m_wkbReporting
    End If
    
    'custom functions
    
    'HARD CODED
    'datum, heute und uhrzeit
    Dim sFuncNow As String
    sFuncNow = m_sReportingFuncPrefix & "<DATUM>"
    If InStr(1, sString, sFuncNow, vbTextCompare) Then
        loadFromReporting = Replace(sString, sFuncNow, Format(Now(), "dd.mm.yyyy hh:nn"))
        Exit Function
    End If
    
    'letzer arbeitstag herausfinden
    Dim sFuncLastDay As String
    sFuncLastDay = m_sReportingFuncPrefix & "<LASTDAY>"
    If InStr(1, sString, sFuncLastDay, vbTextCompare) Then
        Dim aDatum As Variant
        Dim aStunden As Variant
        Dim i As Long
        Dim lastStundenIndex As Long
        Dim sLastDate As String
        
        aDatum = fExl.getArrayOfColumn("A", 4, 53, wkbReporting.Worksheets(m_mcCurrentSettings("<MA>")))
        aStunden = fExl.getArrayOfColumn("H", 4, 53, wkbReporting.Worksheets(m_mcCurrentSettings("<MA>")))
        
        For i = 1 To UBound(aStunden)
            'falls stunden eingegeben und kien Wochenergebnis, dann als letzter tag merken
            If aStunden(i) <> 0 And i Mod 7 <> 0 Then
                lastStundenIndex = i
            End If
        Next i
        'letzter tag, ntsprechend bestimmen
        If lastStundenIndex <> 0 Then
            sLastDate = aDatum(lastStundenIndex)
        End If
        
        'werte ersetzen
        loadFromReporting = Replace(sString, sFuncLastDay, sLastDate)
        Exit Function
    End If
    
    '--------------------------------------------------
    'main - load from reporting
    aParams = Split(sString, ",")
    sWorksheet = Trim(aParams(0))
    lCol = CLng(fExl.LetterToCol(Trim(aParams(1))))
    lRow = CLng(Trim(aParams(2)))
    
    loadFromReporting = wkbReporting.Worksheets(sWorksheet).Cells(lRow, lCol).Value

    
    'params aufteilen, nach tabelle, spalte, zeile...
    'auslese, zuruckgeben
End Function


'gibt reference auf RANGE fuer einen bestimmten mitarbeiter, bestimmter kampagne
Public Function getMaRange(iKampNr As Integer, sMA As String, Optional wksZielDok As Worksheet)
    Dim startRow As Long
    Dim startColumn As Long
    Dim endRow As Long
    Dim endColumn As Long
    Dim rangeHeight As Long
    Dim rangeWidth As Long
    Dim fromTopPosition As Long
    
    On Error GoTo getMaRange_Error
    
    'zielworksheet muss existieren
    If wksZielDok Is Nothing Then
        If Not m_wksWorkingSheet Is Nothing Then
            Set wksZielDok = m_wksWorkingSheet
        Else
            Exit Function
        End If
    End If
    
    rangeHeight = m_rngMaMuster.Rows.Count
    rangeWidth = m_rngMaMuster.Columns.Count
    
    '1,2,3...
    fromTopPosition = f.getIndexInArrayByValue(sMA, m_mcMA.getValues) + 1
    
    endRow = fromTopPosition * rangeHeight
    startRow = endRow - rangeHeight + 1
    
    endColumn = iKampNr * rangeWidth
    startColumn = endColumn - rangeWidth + 1
    
    With wksZielDok
        Set getMaRange = .Range(.Cells(startRow, startColumn), .Cells(endRow, endColumn))
    End With

    On Error GoTo 0
    Exit Function

getMaRange_Error:
    ehError.catchErr ("getMaRange of Klassenmodul clsTableHandler")
    
    
End Function

'class mit parametern initialisieren
Public Function init(ByRef rngMuster As Range, vMA As Variant, vSettingsP As Variant, Optional ByRef wksWorkingTable As Worksheet, Optional ByRef m_ecErrorCollector As clsErrorCollector)
    On Error GoTo init_Error

    Set muster = rngMuster
    Set mitarbeiter = vMA
    Set settings = vSettingsP
    
    'falls keine Tabelle uebergeben
    If wksWorkingTable Is Nothing Then
        Dim sNewTableName As String
        sNewTableName = m_mcSettings("TableName")
        
        'falls tabelle existiert, die nehmen
        If fExl.checkIfWksExists(ThisWorkbook, sNewTableName) Then
            Set WorkingSheet = ThisWorkbook.Worksheets(sNewTableName)
        Else
            Set WorkingSheet = fExl.addWorksheet(sNewTableName)
            
            'falls fehler beim addWorksheet..
            If IsNull(WorkingSheet) Or WorkingSheet Is Nothing Then
                f.dbg "Init: keine Referenz auf Zieltabelle angelegt"
            End If
        End If
    Else
        Set WorkingSheet = wksWorkingTable
    End If
    
    
    If Not m_ecErrorCollector Is Nothing Then
        Set ErrorCollector = m_ecErrorCollector
    End If
    

    On Error GoTo 0
    Exit Function

init_Error:
    ehError.catchErr ("init of Klassenmodul clsTableHandler")
    
    
End Function

Private Sub Class_Initialize()
    Set m_mcSettings = New clsMyCollection
    Set m_mcMA = New clsMyCollection
    Set m_ecErrorCollector = New clsErrorCollector
    Set ehError = New clsErrorHandler
    Set f = New clsFuncs
    Set fExl = New clsFuncsExl
    Set m_mcCurrentSettings = New clsMyCollection
    
    m_sReportingFuncPrefix = "<reportringFunction>"
End Sub

Private Sub Class_Terminate()
    Set m_mcSettings = Nothing
    Set m_mcMA = Nothing
    Set m_ecErrorCollector = Nothing
    Set ehError = Nothing
    Set f = Nothing
    Set fExl = Nothing
    Set m_wkbReporting = Nothing
    Set m_mcCurrentSettings = Nothing
End Sub
