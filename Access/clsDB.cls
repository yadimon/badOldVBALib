VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsDB
' Author    : Dmitry Gorelenkov
' Date      : 08.10.2012
' Changed   : 01.11.2013
' Purpose   : datenbank handler, um connection herzustellen, SQL anweisungen auszufuehren
' Requires  : clsErrorHandler, clsFuncs
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
Private m_lAbfragenCounter As Long
Private rs As ADODB.Recordset 'no need?
Private connection As ADODB.connection
Private ehError As clsErrorHandler
Private f As New clsFuncs

'verbinden, falls connection nicht mehr vorhanden ist
Public Function connect(Optional CON As ADODB.connection) As Boolean
    connect = False
    If connection.State > 0 Then
        GoTo done:
    End If
    
    If CON Is Nothing Then Set CON = CurrentProject.connection
    
    Set connection = CON
    connection.CursorLocation = adUseClient 'langsamer? dafuer recordCount funktionsfaehig
    
done:
    connect = True
End Function

'abfrage / update durchfuehren
'lAffected entweder beim update/insert anzahl der Affected rows, oder bei einer Abfrage (select ..) anzahl von daten saetzen
Public Function SQL(sBefehl As String, Optional ByVal aParams As Variant, Optional bQry As Boolean = False, _
                    Optional ByRef lAffected As Long, Optional bOffline As Boolean = False) As ADODB.Recordset
    If Not connect Then Exit Function
    On Error GoTo ErrorHandler
    Dim bUseParamsArray As Boolean
    bUseParamsArray = True
    
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = connection
    
    cmd.CommandText = sBefehl
    
    'falls vorbereitete anfrage
    If bQry Then
        cmd.Properties("Jet OLEDB:Stored Query") = True
    End If
    
    
    If Not IsMissing(aParams) Then
        'single parameter
        If VarType(aParams) = vbString Or IsNumeric(aParams) Then
            aParams = Array(CStr(aParams))
        'if myCollection
        ElseIf TypeName(aParams) = "clsMyCollection" Then
            bUseParamsArray = False 'keine parameters nun extra als array uebergeben
            Dim sParamPure As String
            Dim par As ADODB.Parameter
            
            'fuer jeden parameter der erwartet wurde
            For Each par In cmd.Parameters
                
                sParamPure = Mid(par.name, 2, Len(par.name) - 2) 'ohne anfangs und end klammern "[", "]"
                
                If aParams.isInArray(sParamPure) Then
                    'correct parameter size
                    Dim valueLength As Long
                    valueLength = Len(CStr(Nz(aParams.getValue(sParamPure), "")))
                    If valueLength > par.Size Then
                        par.Size = valueLength
                    End If

                    par.Value = aParams.getValue(sParamPure)
                    
                Else
                    Debug.Print "Nicht genug parameters mit mycollection uebergeben. bei clsDB.SQL"
                    Exit Function
                End If
                
            Next par
        ElseIf Not IsArray(aParams) Then
            Debug.Print "Falshes Parameteruebergabe bei clsDB.SQL"
            Exit Function
        End If
    End If
    
    'Call p.performance("db query: " & cmd.CommandText, 0)
    
    If bUseParamsArray Then
        Set SQL = cmd.Execute(lAffected, Parameters:=aParams)
    Else
        Set SQL = cmd.Execute(lAffected)
    End If
    
    'falls bOffline, disconect dn RS
    If bOffline Then Set SQL.ActiveConnection = Nothing
    
    
    If lAffected = 0 And TypeOf SQL Is ADODB.Recordset Then
        On Error Resume Next
        lAffected = SQL.RecordCount 'todo wie kann man anzahl herausfinden? ohne schleifen.. ADOX?
        On Error GoTo ErrorHandler
    End If
    
    'Call p.performance("db query: " & cmd.CommandText, 1)
    'abfragen aufzaehlen
    countSQL
    
    Exit Function
ErrorHandler:
    lAffected = -1
    ehError.catchErr ("Fehler bei clsDB.SQL, befehl: " & sBefehl)
End Function

'nur 1 wert, falls vorhanden zuruckgeben
'lAffected entweder beim update/insert anzahl der Affected rows, oder bei einer Abfrage (select ..) anzahl von daten saetzen
Public Function SQLValue(sBefehl As String, Optional ByVal aParams As Variant, Optional bQry As Boolean = False, Optional ByRef lAffected As Long)
    On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Set rs = SQL(sBefehl, aParams, bQry, lAffected)
    'falls nicht leer
    If Not isEmptyRS(rs) Then
        rs.MoveFirst
        SQLValue = rs.Fields(0).Value
    Else
        SQLValue = vbNullString
    End If
    
    Exit Function
ErrorHandler:
    ehError.catchErr
    SQLValue = Empty
End Function

'deprecated
'aus recordset bzw abfrage ein array machen, array(0).getValue("ID") heisst id aus 1. row
Public Function SQLArrayOfCol(params As Variant)
    On Error GoTo ErrorHandler
    Dim fld As Variant
    Dim retArray As Variant
    Dim retMyColTmp As clsMyCollection
    
    If TypeOf params Is Recordset Then
        'changes?
    ElseIf TypeOf params Is Recordset2 Then
        'changes?
    ElseIf TypeOf params Is ADODB.Recordset Then
        'changes ?
    ElseIf TypeName(params) = "String" Then
        Set params = SQL(CStr(params))
    Else
        Debug.Print "SQLArray: uebergebene Argument type wird nicht unterstuetzt"
        Exit Function
   End If

    params.MoveFirst
    retArray = Array()
    Do Until params.EOF
        Set retMyColTmp = New clsMyCollection
        For Each fld In params.Fields
            Call retMyColTmp.Add(fld.Value, fld.name)
        Next fld
        Call f.arrayPush(retArray, retMyColTmp)
        params.MoveNext
    Loop
    
    SQLArrayOfCol = retArray
    
    Exit Function
ErrorHandler:
    Debug.Print "Fehler bei clsDB.SQLArray"
    
    ehError.catchErr
End Function

'aus recordset bzw abfrage ein array machen, array(0)(0) heisst erstes Field aus 1. row
'optional array of collections machen array(0).getValue("ID") heisst Field "ID" aus 1. row
Public Function SQLArray(rs As Variant, Optional bCollection As Boolean = False)
    On Error GoTo ErrorHandler
    Dim fld As Variant
    Dim retArray As Variant
    retArray = Array()
    
    
    If TypeName(rs) = "Recordset" Then
        'changes?
    ElseIf TypeName(rs) = "Recordset2" Then
        'changes?
    ElseIf TypeName(rs) = "ADODB.Recordset" Then
        'changes ?
    ElseIf TypeName(rs) = "String" Then
        Set rs = SQL(CStr(rs))
    Else
        Debug.Print "SQLArray: uebergebene Argument type wird nicht unterstuetzt"
        Exit Function
    End If
    
    
    
    If isEmptyRS(rs) Then
        GoTo Final
    End If
    
    rs.MoveFirst
    If bCollection Then 'if need array of collections
    Dim retMyColTmp As clsMyCollection
        Do Until rs.EOF
            Set retMyColTmp = New clsMyCollection
            For Each fld In rs.Fields
                Call retMyColTmp.Add(fld.Value, fld.name)
            Next fld
            Call f.arrayPush(retArray, retMyColTmp)
            rs.MoveNext
        Loop
    Else 'for simple array
    Dim retArrayTmp As Variant
        Do Until rs.EOF
            retArrayTmp = Array()
            For Each fld In rs.Fields
                Call f.arrayPush(retArrayTmp, fld.Value)
            Next fld
            Call f.arrayPush(retArray, retArrayTmp)
            rs.MoveNext
        Loop
    End If
    
    
Final:
    SQLArray = retArray
    
    Exit Function
ErrorHandler:
    ehError.catchErr ("clsDB.SQLArray")
End Function

Public Function isEmptyRS(rs As Variant)

    If TypeName(rs) = "Recordset" Or TypeName(rs) = "Recordset2" Or TypeName(rs) = "ADODB.Recordset" Then
        If rs.EOF And rs.BOF Then
            isEmptyRS = True
        Else
            isEmptyRS = False
        End If
    Else
        Debug.Print "isEmptyRS: uebergebene Argument type wird nicht unterstuetzt"
    End If
    
End Function

Public Function printRs(ByVal rsRS As Variant)
    On Error GoTo ErrorHandler
    Dim sRes As String
    Dim fld As Variant
    
    If TypeName(rsRS) = "Recordset" Then
        'changes?
    ElseIf TypeName(rsRS) = "Recordset2" Then
        'changes?
    ElseIf TypeName(rsRS) = "ADODB.Recordset" Then
        'changes ?
    Else
        Debug.Print "printRs:uebergebene Argument type """ & TypeName(rsRS) & """ wird nicht unterstuetzt"
        Exit Function
    End If
    
    rsRS.MoveFirst
    
    Do Until rsRS.EOF
        sRes = sRes & "|"
        For Each fld In rsRS.Fields
            sRes = sRes & fld.Value & "|"
        Next fld
        sRes = sRes & vbCrLf
        rsRS.MoveNext
    Loop
    
    Debug.Print sRes
    
    Exit Function
ErrorHandler:
    ehError.catchErr ("clsDB.printRS")
End Function

'requests anzahl zuruckgeben
Property Get requests() As Long
    requests = m_lAbfragenCounter
End Property

'request aufzaehlen
Private Function countSQL()
    m_lAbfragenCounter = m_lAbfragenCounter + 1
End Function


Private Sub Class_Initialize()
    Set connection = New ADODB.connection
    Set ehError = New clsErrorHandler
    m_lAbfragenCounter = 0
End Sub

Private Sub Class_Terminate()
    Set connection = Nothing
    Set ehError = Nothing
End Sub

