VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParseCallNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsParseCallNumbers
' Author    : Dmitry Gorelenkov
' Date      : 23.08.2012
' Purpose   : Tries to get phonenumbers from a string with different delimiters
' Requires  : clsFuncs, clsErrorHandler
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
Private f As clsFuncs
Private ehError As New clsErrorHandler
Private m_minNumbs
Public m_maxNumbs
'returns array of possible call numbers in the sNumberString
Public Function tryToGet(ByVal sNumberString As String) As Variant
    Dim restArray As Variant
    Dim returnArray As Variant
    Dim i As Integer
    Dim j As Integer
    Dim sFirstPart As String
    On Error GoTo tryToGet_Error

    returnArray = Array()
    
    If Len(sNumberString) = 0 Then
        GoTo final
    End If
    
    
    'for every symbol
    For i = 1 To Len(sNumberString)
        'in case its not a number, or last number
        If Not IsNumeric(Mid(sNumberString, i, 1)) Or Len(sNumberString) = i Then
            'if first part is a number,
            sFirstPart = Mid(sNumberString, 1, i)
            If canBeAPhoneNumber(sFirstPart) Then
                'in case the full string is a one number
                If Len(sNumberString) = i Then
                    returnArray = Array(modifyToPhoneNumber(sNumberString))
                Else
                    'if there is second part that can be parsed as numbers (rekursive)
                    restArray = tryToGet(Mid(sNumberString, i + 1, Len(sNumberString)))
                    If UBound(restArray) > -1 Then
                        'concatenate first part and remained numbers as parsed array
                        returnArray = f.arrayConcat(Array(modifyToPhoneNumber(sFirstPart)), restArray)
                        Exit For
                    End If
                End If
                
            End If
         End If
    Next i
final:
    
    'add 0 at beginn if missing
    If f.aLength(returnArray) > 0 Then
        For j = LBound(returnArray) To UBound(returnArray)
            If Left(returnArray(j), 1) <> "0" Then
                returnArray(j) = "0" & returnArray(j)
            End If
        Next j
    End If
    tryToGet = returnArray

    On Error GoTo 0
    
    Exit Function

tryToGet_Error:
    ehError.catchErr ("tryToGet of Klassenmodul clsParseCallNumbers")
    
End Function

'checks if the string can be parsed as a phonenumber
Private Function canBeAPhoneNumber(ByVal sString As String) As Boolean
    Dim sNumsOnlyString As String
    
    canBeAPhoneNumber = False
    
    sNumsOnlyString = modifyToPhoneNumber(sString)
    
    If Len(sNumsOnlyString) >= m_minNumbs And Len(sNumsOnlyString) <= m_maxNumbs Then
        canBeAPhoneNumber = True
    End If
    
    
End Function

Private Function modifyToPhoneNumber(sString As String)
    Dim sNumsOnlyString As String
    Dim sCurLetter As String
    Dim i As Integer
    

    Call Replace(sString, "+", "00", 1, 1)

    
    For i = 1 To Len(sString)
        sCurLetter = Mid(sString, i, 1)
        If IsNumeric(sCurLetter) Then
            sNumsOnlyString = sNumsOnlyString & sCurLetter
        End If
    Next i
    
    modifyToPhoneNumber = sNumsOnlyString
End Function
Private Sub Class_Initialize()
    Set f = New clsFuncs
    m_minNumbs = 8
    m_maxNumbs = 16
End Sub

Private Sub Class_Terminate()
    Set Me!f = Nothing
    Set Me!ehError = Nothing
End Sub
