VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXlDataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsEExlDataSheet
' Author    : Dmitry Gorelenkov
' Date      : 17.11.2013
' Changed   : 02.12.2013
' Purpose   : helper for Data-Worksheet in excel
' Requires  : excel, clsFuncsExl, clsFuncs
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Explicit

Private m_wksDataWorksheet As Worksheet
Private m_rngHeaderRange As Range
Private m_bInitDone As Boolean
Private m_aHeaders As Variant
Private fXl As clsFuncsExl
Private f As clsFuncs


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::PROPERTIES::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'getter for range of headers
Public Property Get RangeOfHeader() As Range
    Set RangeOfHeader = m_rngHeaderRange
End Property


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::FUNCTIONS:::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Function init(wksDataWorksheet As Worksheet) As Boolean
    On Error GoTo init_Error

    If m_bInitDone Then Exit Function

    Set m_wksDataWorksheet = wksDataWorksheet
    Set m_rngHeaderRange = m_wksDataWorksheet.Range("A1:" & fXl.ColToLetter(fXl.getLastCol(m_wksDataWorksheet, 1)) & "1") 'TODO falls header wo anderes ist
    
    m_aHeaders = f.arrayXDimTo1Dim(m_rngHeaderRange.value)
    
    init = True
    m_bInitDone = True
    
    On Error GoTo 0
    Exit Function
init_Error:

    f.dbg "Error: init of Klassenmodul clsXlDataSheet"
    init = False
    m_bInitDone = False
    On Error GoTo 0
End Function

'every header in table, must have formatnumber for the column
'modify formatnumber of column by header
Public Function reformatTable(Optional wksDataWorksheet As Worksheet, Optional rngHeaderRange As Range)
    If Not test_init() Then Exit Function

    Dim i As Long

    If wksDataWorksheet Is Nothing Then Set wksDataWorksheet = m_wksDataWorksheet
    If rngHeaderRange Is Nothing Then Set rngHeaderRange = m_rngHeaderRange
    'for every header column
    For i = 1 To rngHeaderRange.Columns.Count
        'set columnt format = header cell format
        wksDataWorksheet.Columns.Item(i).NumberFormat = rngHeaderRange.Cells(1, 1).Offset(0, i - 1).NumberFormat
    Next i
End Function

'ersetzt werte in spalten die als Numeric definiert sind, aber string enthalten, duch numerische werte.
'aArrayOfHeaders_p array von header bei denen es ersetzt werden muss, falls nicht angegeben, wird bei allen numerischen spalten ersetzt
'bHardModify bestimmt ob werte modifiziert werden und texte geloescht werden muessen. also z.B. "text12" -> "12" umgewandelt werden muss
'bFillIfEmpty bestimmt of leere felder mit "0" befuellt werden muessen
Public Function modifyToNumber(Optional aArrayOfHeaders_p As Variant, _
                                Optional bHardModify As Boolean = False, _
                                Optional bFillIfEmpty As Boolean = False)
    Dim aArrayOfHeaders As Variant
    Dim sHeader As Variant 'string
    Dim rngColumn As Range
    Dim rngCell As Range
    'falls nichts uebergeben, alle numerische spalten nehmen
    If IsArray(aArrayOfHeaders_p) Then
        aArrayOfHeaders = aArrayOfHeaders_p
    Else
        aArrayOfHeaders = getHeadersByNumberFormat("0")
    End If
    
    For Each sHeader In aArrayOfHeaders
        Set rngColumn = getRangeOfColumnByHeader(CStr(sHeader))
        If Not rngColumn Is Nothing Then
        
            On Error Resume Next
            For Each rngCell In rngColumn
                If rngCell.value = vbNullString And bFillIfEmpty Then
                    rngCell.value = 0
                End If
                
                If rngCell.value <> vbNullString Then
                    If bHardModify Then
                        rngCell.value = CDec(Val(rngCell.value))
                    Else
                        rngCell.value = CDec(rngCell.value)
                    End If
                End If
            Next rngCell
            On Error GoTo 0
        End If
    Next sHeader

End Function

'ersetzt werte in spalten die als Datum definiert sind, aber string enthalten, duch datums werte.
'aArrayOfHeaders_p array von header bei denen es ersetzt werden muss, falls nicht angegeben, wird bei allen datums spalten ersetzt
'mit bHardModify werden die werte die nicht passt, als Cdate(0) datums werte gesetzt
Public Function modifyToDate(Optional aArrayOfHeaders_p As Variant, Optional bHardModify As Boolean = False)
    Dim aArrayOfHeaders As Variant
    Dim sHeader As Variant 'string
    Dim rngColumn As Range
    Dim rngCell As Range
    'falls nichts uebergeben, alle datums spalten nehmen
    If IsArray(aArrayOfHeaders_p) Then
        aArrayOfHeaders = aArrayOfHeaders_p
    Else
        aArrayOfHeaders = getHeadersByNumberFormat("m/d/yyyy") 'schwachstelle... wenn anders formatiert.. TODO
    End If
    
    For Each sHeader In aArrayOfHeaders
        Set rngColumn = getRangeOfColumnByHeader(CStr(sHeader))
        If Not rngColumn Is Nothing Then
        
            On Error Resume Next
            For Each rngCell In rngColumn
                If (bHardModify) Then
                    rngCell.value = f.CdateTry(rngCell.value, CDate(0))
                    If Err.Number <> 0 Then
                        rngCell.value = f.CdateTry(rngCell.Text, CDate(0))
                        Err.Clear
                    End If
                Else
                    rngCell.value = f.CdateTry(rngCell.value)
                End If
                
            Next rngCell
            On Error GoTo 0
        End If
    Next sHeader

End Function

'returns range of column by header name
Public Function getRangeOfColumnByHeader(sHeader As String) As Range
    Dim rngHeaderRange As Range
    Set getRangeOfColumnByHeader = Nothing
    Set rngHeaderRange = getRangeOfHeaderByName(sHeader)
    
    If Not rngHeaderRange Is Nothing Then
        With m_wksDataWorksheet
            Set getRangeOfColumnByHeader = .Range(.Cells(rngHeaderRange.Row + 1, rngHeaderRange.column), .Cells(getLastRow(), rngHeaderRange.column))
        End With
    End If
End Function

'returns array with headers with the sFormat .NumberFormat property
Public Function getHeadersByNumberFormat(sFormat As String)
    Dim cellHeader As Range
    Dim aArrayOfHeaders As Variant

    For Each cellHeader In Me.RangeOfHeader
        'falls numerisch
        If cellHeader.NumberFormat = sFormat Then
            Call f.arrayPush(aArrayOfHeaders, cellHeader.value)
        End If
    Next cellHeader

    getHeadersByNumberFormat = aArrayOfHeaders
End Function

'returns range of header, with this name
Public Function getRangeOfHeaderByName(sHeaderName As String) As Range
    Dim rngCell As Range

    For Each rngCell In m_rngHeaderRange
        If rngCell.value = sHeaderName Then
            Set getRangeOfHeaderByName = rngCell
            Exit For
        End If
    Next rngCell
End Function

'letzte zeile. Recht aufwendige funktion..
Public Function getLastRow() As Long
    Dim rngHeader As Range
    Dim lLastRow As Long
    Dim lResult As Long
    lResult = 1
    For Each rngHeader In m_rngHeaderRange
        lLastRow = fXl.getLastRow(m_wksDataWorksheet, rngHeader.column)
        lResult = f.max(lResult, lLastRow)
    Next rngHeader
    
    getLastRow = lResult
End Function

'last column of data range
Public Function getLastColumn() As Long
    getLastColumn = m_rngHeaderRange.Cells(1, m_rngHeaderRange.Columns.Count).column
End Function

'haengt daten von einem DataSheet zu anderem 'TODO merge, column anhaengig..
Public Function addDataFromDataSheet(dsDataSheet As clsXlDataSheet, Optional ByRef lLinesAdded) As Boolean
    Dim rngStartCell As Range
    Dim rngEndCell As Range
    Dim rngRangeFrom As Range
    
    If Not Me.equals(dsDataSheet) Then
        f.dbg "addDataFromDataSheet: headers ungleich"
        addDataFromDataSheet = False
        Exit Function
    End If
    
    Set rngRangeFrom = dsDataSheet.getDataRange()
    
    'start cell of range where to save
    Set rngStartCell = m_wksDataWorksheet.Cells(Me.getLastRow + 1, m_rngHeaderRange.Cells(1, 1).column)
    'end cell of range where to save
    Set rngEndCell = rngStartCell.Offset(rngRangeFrom.Rows.Count - 1, rngRangeFrom.Columns.Count - 1)
    
    m_wksDataWorksheet.Range(rngStartCell, rngEndCell).value = rngRangeFrom.value2

    
    lLinesAdded = rngRangeFrom.Rows.Count
    
    addDataFromDataSheet = True
End Function

'returns range of data (used range minus headers)
Public Function getDataRange() As Range
    Dim rngStartCell As Range
    Dim rngEndCell As Range
    
    Set rngStartCell = m_rngHeaderRange.Cells(1, 1).Offset(1, 0)
    Set rngEndCell = m_wksDataWorksheet.Cells(Me.getLastRow(), Me.getLastColumn())
    
    Set getDataRange = m_wksDataWorksheet.Range(rngStartCell, rngEndCell)
    
End Function

'tests if DataSheets (headers) are the same
Public Function equals(DataSheet As clsXlDataSheet) As Boolean
    Dim i As Long
    equals = True
    For i = 1 To Me.RangeOfHeader.Columns.Count
        If UCase(DataSheet.RangeOfHeader.Cells(1, i).value) <> UCase(Me.RangeOfHeader.Cells(1, i)) Then
            equals = False
            Exit For
        End If
    Next i
End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::PRIVATE FUNCTIONS:::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::



'test if initialized, if not, prints error
Private Function test_init() As Boolean

    If Not m_bInitDone Then
        Debug.Print "clsEExlDataSheet: Not initialized!"
    End If

    test_init = m_bInitDone
End Function

Private Sub Class_Initialize()
    Set fXl = New clsFuncsExl
    Set f = New clsFuncs
End Sub

Private Sub Class_Terminate()
    Set fXl = Nothing
    Set f = Nothing
End Sub
