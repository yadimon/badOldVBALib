VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVorlagenHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsVorlagenHandler
' Author    : Dmitry Gorelenkov
' Date      : 01.10.2012
' Changed   : 23.04.2013
' Purpose   : Erstellt Vorlagen pro kampagne in zieldokument
' Requires  :
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------

Option Explicit
'abhaengig von einigen funcs von clsFuncs, globals auch benutzt

Private p_aAgenturenNamen As Variant
Private p_aVorlagenSheets As Variant
Private p_aVorlagenAddNamen As Variant
Private p_aVorlagenUniqSheets As Variant

Function init(ByVal aAgenturen, ByVal aVorlagenSheets, ByVal aVorlagenUniqSheets, ByVal aVorlagenAddNamen)
    p_aAgenturenNamen = aAgenturen
    p_aVorlagenSheets = aVorlagenSheets
    p_aVorlagenAddNamen = aVorlagenAddNamen
    p_aVorlagenUniqSheets = aVorlagenUniqSheets
End Function

'macht aus einer datei leeres Fehlerbericht, und hinzugefuegt alle vorlagen. LOCAL
Function addAlleVorlagenToFile(sPfad As String, Optional sKampageParam As String)
    loadParams
    Dim wkbFile As Workbook
    Dim i As Integer
    Dim j As Integer
    Dim wksNewFehlerBerichtTabelle As Worksheet
    
    addAlleVorlagenToFile = False
    
    'datei oeffnen
    Set wkbFile = fXl.findOpenedDocOrOpen(sPfad)
    If wkbFile Is Nothing Then Exit Function
    'falls arrays pruefen
    If Not IsArray(p_aAgenturenNamen) Then Exit Function
    If Not IsArray(p_aVorlagenSheets) Then Exit Function
    If Not IsArray(p_aVorlagenAddNamen) Then Exit Function
    If Not IsArray(p_aVorlagenUniqSheets) Then Exit Function
    'array mit sheets und namen mussen gleich gross sein
    If f.aLength(p_aVorlagenSheets) <> f.aLength(p_aVorlagenAddNamen) Then Exit Function
    
    
        
    'datei leeren
    If clearFile(wkbFile) Then
    
        'einzigartige vorlagen zuerst
        Dim wksUniqTmp As Worksheet
        For i = LBound(p_aVorlagenUniqSheets) To UBound(p_aVorlagenUniqSheets)
            Set wksUniqTmp = p_aVorlagenUniqSheets(i)
            Set wksNewFehlerBerichtTabelle = addVorlageToFile(wkbFile, wksUniqTmp)
            If wksNewFehlerBerichtTabelle Is Nothing Then
                GoTo ErrorHandler
            End If
            
            'suchen nach "<KAMPAGNE>" und durch aktuelle kampagne ersetzen
            Call wksNewFehlerBerichtTabelle.UsedRange.Replace("<KAMPAGNE>", sKampageParam)
            
        Next i
        
        
        'fuer jede agentur
        For i = LBound(p_aAgenturenNamen) To UBound(p_aAgenturenNamen)
        
            'fuer alle vorlagen
            For j = LBound(p_aVorlagenSheets) To UBound(p_aVorlagenSheets)
            
                'vorlage hinzufuegen, falls fehler, GoTo ErrorHandler
                Dim wksTmp As Worksheet
                Set wksTmp = p_aVorlagenSheets(j)
                Set wksNewFehlerBerichtTabelle = addVorlageToFile(wkbFile, wksTmp, p_aAgenturenNamen(i) & p_aVorlagenAddNamen(j))
                If wksNewFehlerBerichtTabelle Is Nothing Then
                    GoTo ErrorHandler
                End If
                
                'suchen nach "<AGENTUR>" und durch aktuelle agenturname ersetzen
                Call wksNewFehlerBerichtTabelle.UsedRange.Replace("<AGENTUR>", p_aAgenturenNamen(i))
                
            Next j
            
        Next i

    End If
    
    'erste worksheet loeschen
    Application.DisplayAlerts = False
    wkbFile.Sheets(1).Delete
    Application.DisplayAlerts = True
    'datei schliessen
    wkbFile.Close (True)
    
    addAlleVorlagenToFile = True
    
    
    
    Exit Function
ErrorHandler:
    On Error Resume Next
    addAlleVorlagenToFile = False
    wkbFile.Close (False)
    Debug.Print "Fehler bei addAlleVorlagenToFile, class: clsVorlagenHandler"
End Function


'haengt vorlage tabelle als letzte tabelle an in wkbFile datei, falls keine tabelle mit dem Namen schon existiert
Function addVorlageToFile(wkbFile As Workbook, wksVorlage As Worksheet, Optional sname As String) As Worksheet
    On Error GoTo ErrorHandler
    Dim origVorlageName As String
    Set addVorlageToFile = Nothing
    
    'falls parameter null sind
    If wkbFile Is Nothing Or wksVorlage Is Nothing Then Exit Function
    
    
    If IsMissing(sname) Or IsEmpty(sname) Or sname = vbNullString Then sname = wksVorlage.name
    'falls so eine Tabelle nicht existiert
    If Not fXl.checkIfWksExists(wkbFile, sname) Then
        'temporaer vorlage umbenennen
        origVorlageName = wksVorlage.name
        wksVorlage.name = sname
        
        
        'am ende anhaengen
        wksVorlage.Copy After:=wkbFile.Sheets(wkbFile.Sheets.count)
        'vorlage auf originalname umbenennen
        wksVorlage.name = origVorlageName
        
        Set addVorlageToFile = wkbFile.Sheets(wkbFile.Sheets.count)
    End If
    
    
    Exit Function
ErrorHandler:
    On Error Resume Next
    Set addVorlageToFile = Nothing
    wksVorlage.name = origVorlageName
    Debug.Print "Fehler bei addVorlageToFile, class: clsVorlagenHandler"
End Function

'remove all worksheets until first, and clears content of first worksheet
Function clearFile(wkbFile As Workbook)
    Dim i As Integer
    clearFile = False
    If wkbFile Is Nothing Then Exit Function
    
    For i = 2 To wkbFile.Sheets.count
        Application.DisplayAlerts = False
        wkbFile.Sheets(2).Delete
        Application.DisplayAlerts = True
    Next i
    If wkbFile.Sheets(1).ProtectContents Then
        Debug.Print wkbFile.Sheets(1).name & "ist schreibgeschützt" & vbCrLf & "File: " & wkbFile.FullName
        Exit Function
    End If
    
    wkbFile.Sheets(1).UsedRange.Delete
    wkbFile.Sheets(1).name = "3415215xSomeBadNamex614351"
    clearFile = True
End Function

