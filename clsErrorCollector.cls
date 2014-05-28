VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_hasError As Boolean
Private colErrors As New Collection
Public p_sDelimiter As String

Private Sub Class_Initialize()
    p_sDelimiter = "------"
End Sub

Public Function add(ByVal vError As Variant)
    On Error Resume Next
    
    'in case vError is a string, simply add
    If VarType(vError) = vbString Then
        colErrors.add (Array(vError, 0)) 'maybe add more later
        
        
        'Err object uebergeben? WIE?
'    ElseIf VarType(vError) = vbError Then ' in case vError is an ErrObject
'        Dim vErrorText As String
'        vErrorText = "Fehler Source: " & vError.Source & vbCrLf
'        vErrorText = vErrorText & "Fehler Beschreibung: " & vError.Description & vbCrLf
'        vErrorText = vErrorText & "Fehler Nummer: " & vError.Number & vbCrLf
'        colErrors.add (Array(vErrorText, 1))

    p_hasError = True
    End If
    
End Function

'prueft ob error existiert, falls ja, information darueber speichern
Public Function checkForError()
    If Err.Number <> 0 Then
        Dim vErrorText As String
        vErrorText = "Fehler Source: " & Err.Source & vbCrLf
        vErrorText = vErrorText & "Fehler Beschreibung: " & Err.Description & vbCrLf
        vErrorText = vErrorText & "Fehler Nummer: " & Err.Number & vbCrLf
        add (vErrorText)
    End If
End Function

Public Function getErrors(Optional errLevel As Variant = 1) As String
    On Error Resume Next
    Dim element As Variant
    Dim resString As String
    For Each element In colErrors
        If element(1) <= errLevel Then
            resString = resString & element(0) & vbCrLf & _
            p_sDelimiter & vbCrLf
        End If
    Next element
    
    getErrors = resString
End Function

Public Property Get hasError() As Boolean
    hasError = p_hasError
End Property

Public Property Let delimiter(vValue As Variant)
    On Error GoTo ErrorHandler
    p_sDelimiter = CStr(vValue)
    
ErrorHandler:
    Debug.Print Err.Description
    Debug.Print Err.Number
    Debug.Print Err.Source
    Debug.Print Err.HelpContext
    Debug.Print Err.LastDllError
End Property

