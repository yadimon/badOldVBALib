VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXlDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsXlDB
' Author    : Dmitry Gorelenkov
' Date      : 15.11.2013
' Changed   : 15.11.2013
' Purpose   : xl datenbank functions
' Requires  : ADODB
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------


'returns connection to the worksheet
'options like "IMEX=1;" sPath is ThisWorkbook per default
Public Function getConnectionXl(Optional sPath As String = "", Optional sOptions As String) As ADODB.connection
    Dim con As New ADODB.connection
    Dim sConString As String
    
    If Len(sPath) <= 0 Then
        sPath = ThisWorkbook.FullName
    End If
    
    If Not f.fileExists(sPath) Then
        Exit Function
    End If
    
    sConString = "Data Source=" & sPath _
        & ";Extended Properties=""Excel 8.0;HDR=Yes;" & sOptions & """;"
    'setup the connection
    With con
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = sConString
        .Open
    End With
    
    
    Set getConnectionXl = con
End Function

'optional wkbWorkbook if vWorksheetOrRange is worksheet name only
'bOffline = true means, the recordset will be disconnected
'modified from: http://stackoverflow.com/questions/2484516/vba-create-adodb-recordset-from-the-contents-of-a-spreadsheet
Public Function getRSFromXl(sSQL As String, Optional wkbWorkbook As Workbook, Optional bOffline As Boolean = True) As ADODB.Recordset
    Dim sStrForConnection As String
    Dim tableAndRangeForConnection As String
    Dim sConString As String
    Dim rst As New ADODB.Recordset
    Dim cnx As New ADODB.connection
    Dim cmd As New ADODB.Command
    
    If wkbWorkbook Is Nothing Then Set wkbWorkbook = ThisWorkbook
    
    
    Set cnx = getConnectionXl(wkbWorkbook.FullName)

    'setup the command
    Set cmd.ActiveConnection = cnx
    cmd.CommandType = adCmdText
    cmd.CommandText = sSQL
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenDynamic
    rst.LockType = adLockOptimistic

    'open the connection
    rst.Open cmd
    
    'if need offline
    If bOffline Then
        'disconnect the recordset
        Set rst.ActiveConnection = Nothing
    
        'cleanup
        If CBool(cmd.State And adStateOpen) = True Then
            Set cmd = Nothing
        End If
    
        If CBool(cnx.State And adStateOpen) = True Then cnx.Close
        Set cnx = Nothing
    End If
    

    '"return" the recordset object
    Set getRSFromXl = rst
End Function

'makes recordset from worksheet.
'vWorksheet can be a worksheet or a worksheet name
'optional wkbWorkbook if vWorksheet is worksheet name only
'bOffline = true means, the recordset will be disconnected
Public Function getRSFromWorksheet(vWorksheet As Variant, Optional wkbWorkbook As Workbook, Optional bOffline As Boolean = False) As ADODB.Recordset
    Dim tableAndRangeForConnection As String
    Dim rst As New ADODB.Recordset
    
    If wkbWorkbook Is Nothing Then Set wkbWorkbook = ThisWorkbook
    
    'test type
    Select Case TypeName(vWorksheet)
        Case "Worksheet":
            Set wkbWorkbook = vWorksheet.Parent
            tableAndRangeForConnection = vWorksheet.name & "$"
        Case "String":
            tableAndRangeForConnection = vWorksheet & "$"
        Case Else:
    End Select
    
    Set rst = getRSFromXl("SELECT * FROM [" & tableAndRangeForConnection & "]", wkbWorkbook, bOffline)

    '"return" the recordset object
    Set getRSFromWorksheet = rst
End Function

