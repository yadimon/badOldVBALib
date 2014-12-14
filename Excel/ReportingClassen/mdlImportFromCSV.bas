Attribute VB_Name = "mdlImportFromCSV"
Option Explicit
Private module

Public Function GetReporting(sPathToFile As String) As Workbook
    If IsCSV(sPathToFile) Then
        Set GetReporting = CopyCSVToNewWkb(sPathToFile)
    Else
        Set GetReporting = Workbooks.Open(sPathToFile)
    End If
End Function

Public Function IsCSV(sPathToCheck As String) As Boolean
    IsCSV = Right$(sPathToCheck, 4) = ".csv"
End Function

'some code from http://stackoverflow.com/questions/10269366/open-csv-file-with-correct-data-format-for-each-column-using-textfilecolumndatat
Public Function CopyCSVToNewWkb(sPath As String) As Workbook
'todo improve delimiter etc?
    Dim wkbNew As Workbook
    Dim wksDataSheet As Worksheet

    Dim aMyData As String, strData() As String, TempAr() As String
    Dim ArCol() As Long, i As Long
    Dim FileNr As Long


    Set wkbNew = Excel.Application.Workbooks.add(xlWBATWorksheet)    'with 1 sheet
    Set wksDataSheet = wkbNew.Worksheets(1)


    '~~> Open the text file in one go
    FileNr = FreeFile()
    Open sPath For Binary As #FileNr
    Line Input #FileNr, aMyData
    Close #FileNr

    strData() = Split(aMyData, """;""")    'strings, headers, delim: ";"

    '~~> Create our Array for TEXT
    ReDim ArCol(1 To UBound(strData))
    For i = 1 To UBound(ArCol)
        ArCol(i) = 2
    Next i

    With wksDataSheet.QueryTables.add(Connection:= _
                                      "TEXT;" & sPath, Destination:=Range("$A$1") _
                                                                    )
        .name = "Output"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote    '<=== "" values
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True    '<=== semikolon!
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = ArCol
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

    Call wksDataSheet.QueryTables(1).Refresh(False)
    Call wksDataSheet.QueryTables(1).Delete

    Set CopyCSVToNewWkb = wkbNew
End Function

