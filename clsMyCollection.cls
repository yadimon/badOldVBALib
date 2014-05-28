VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMyCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsMyCollection
' Author    : Dmitry Gorelenkov
' Date      : 05.2012
' Changed   : 03.03.2013
' Purpose   : Schnellere Collection, arbeitet mit arrays
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------


Option Explicit

Private aInhalt As Variant


'Public Property Get NewEnum() As IUnknown
''here:
''Attribute NewEnum.VB_UserMemId = -4
'   Set NewEnum = someKeys.[_NewEnum]
'End Property


'default value
Property Get value(sKey As Variant) As Variant
Attribute value.VB_UserMemId = 0
    'here:
    'Attribute Value.VB_UserMemId = 0
    If IsObject(getValue(sKey)) Then
        Set value = getValue(sKey)
    Else
        value = getValue(sKey)
    End If
End Property

'default value
Property Let value(sKey As Variant, vValue As Variant)
Attribute value.VB_UserMemId = 0
    'here:
    'Attribute Value.VB_UserMemId = 0
    Call setValue(vValue, sKey)
'    If IsObject(getValue(key)) Then
'
'        Set Value = getValue(sKey)
'    Else
'        Value = getValue(sKey)
'    End If
End Property

'default value
Property Set value(sKey As Variant, vValue As Variant)
Attribute value.VB_UserMemId = 0
    'here:
    'Attribute Value.VB_UserMemId = 0
    Call setValue(vValue, sKey)
'    If IsObject(getValue(key)) Then
'        Set Value = getValue(sKey)
'    Else
'        Value = getValue(sKey)
'    End If
End Property

Public Function getKeys() As Variant
    On Error GoTo errorHandler
    Dim retArray As Variant
    Dim element As Variant
    retArray = Array()
    
    For Each element In aInhalt
        ReDim Preserve retArray(UBound(retArray) + 1)
        retArray(UBound(retArray)) = element(0)
    Next element
    
    getKeys = retArray
    
    Exit Function
errorHandler:
    Debug.Print "Fehler bei getKeys"
End Function


Public Function GetValues() As Variant
    On Error GoTo errorHandler
    Dim element As Variant
    Dim retArray As Variant
    retArray = Array()
    
    For Each element In aInhalt
        ReDim Preserve retArray(UBound(retArray) + 1)
        If IsObject(element(1)) Then
            Set retArray(UBound(retArray)) = element(1)
        Else
            retArray(UBound(retArray)) = element(1)
        End If
        
    Next element
    
    GetValues = retArray
    
    Exit Function
errorHandler:
    Debug.Print "Fehler bei getValues"
End Function

Public Function getKey(index As Variant)
    On Error GoTo errorHandler
    If Not IsEmpty(index) And IsNumeric(index) And Val(index) = index And index <= UBound(aInhalt) And index >= 0 Then
        getKey = aInhalt(index)(0)
    Else
        getKey = -1
    End If
    
    Exit Function
errorHandler:
    Debug.Print "Fehler bei getKey, index = " & index
End Function

Public Function getValue(Optional sKey As Variant = "")
    On Error GoTo errorHandler
    Dim lIndex As Long
    If sKey = "" Then Exit Function
    
    
    If isIndex(sKey) Then
        lIndex = CLng(sKey)
    Else
        'try to get by name
        lIndex = getPosInArray(CStr(sKey))
    End If
    
    'falls nichts gefunden
    If lIndex = -1 Then
        Debug.Print "Fehler bei getValue, sKey = " & sKey & " nichts gefunden"
        Exit Function
    End If
    
    If IsObject(aInhalt(lIndex)(1)) Then
        Set getValue = aInhalt(lIndex)(1)
    Else
        getValue = aInhalt(lIndex)(1)
    End If
    
    Exit Function
errorHandler:
    Debug.Print "Fehler bei getValue, sKey = " & sKey
End Function


Public Function getKeyByValue(vValue As Variant)
    On Error GoTo errorHandler
    Dim bMatch As Boolean
    Dim i As Long
    getKeyByValue = -1
    bMatch = False
    
    For i = LBound(aInhalt) To UBound(aInhalt)
        If IsObject(vValue) Then
            If aInhalt(i)(1) Is vValue Then
                bMatch = True
            End If
        Else
            If aInhalt(i)(1) = vValue Then
                bMatch = True
            End If
        End If
        
        If bMatch Then
            getKeyByValue = aInhalt(i)(0)
            Exit For
        End If
        
    Next i
    
    
    Exit Function
errorHandler:
    Debug.Print "Fehler bei getKeyByValue, vValue = " & vValue
End Function

Public Function add(ByRef vValue As Variant, Optional sKey As String = "")
    On Error GoTo errorHandler
    Dim arrayMax As Long
    Dim valueElement As Variant
    
    arrayMax = UBound(aInhalt)
    
    If Not IsObject(vValue) Then
        valueElement = vValue
    Else
        Set valueElement = vValue
    End If
    
    If Not isInArray(sKey) Or sKey = "" Then
        ReDim Preserve aInhalt(arrayMax + 1)
        aInhalt(arrayMax + 1) = Array(sKey, valueElement)
    End If
    
    
    Exit Function
errorHandler:
    Debug.Print "Fehler bei add, vValue = " & vValue & " sKey = " & sKey
End Function

Public Function setValue(vValue As Variant, sKey As Variant)
    On Error GoTo errorHandler
    Dim Position As Long
    Position = getPosInArray(CStr(sKey))

    If Position <= -1 Then
        'sKey is index?
        If isIndex(sKey) Then
            Position = CLng(sKey)
        Else
            Call add(vValue, CStr(sKey))
            Exit Function
        End If
    End If
    
    If IsObject(vValue) Then
        Set aInhalt(Position)(1) = vValue
    Else
        aInhalt(Position)(1) = vValue
    End If
    
    Exit Function
errorHandler:
    Debug.Print "Fehler bei setValue, vValue = " & vValue & " sKey = " & sKey
End Function

Public Function isInArray(sKey As String)
    On Error GoTo errorHandler
    isInArray = False
    If getPosInArray(sKey) > -1 Then isInArray = True
    
    Exit Function
errorHandler:
    Debug.Print "Fehler bei isInArray, sKey = " & sKey
End Function

Public Function remove(sKey As Variant)
    On Error GoTo errorHandler
    Dim i As Long
    Dim idxToRemove As Long
    Dim ubnd As Long
    ubnd = UBound(aInhalt)
    
    'index to remove bestimmen
    If isIndex(sKey) Then
        idxToRemove = CLng(sKey)
    Else
        idxToRemove = getPosInArray(CStr(sKey))
    End If
    
    'falls nichts gefunden
    If Not isIndex(idxToRemove) Or UBound(aInhalt) < idxToRemove Or idxToRemove < 0 Then Exit Function
    
    If getSize = 1 Then
        aInhalt = Array()
    Else
        For i = idxToRemove + 1 To ubnd
            aInhalt(i - 1) = aInhalt(i)
        Next i
        
        ReDim Preserve aInhalt(ubnd - 1)
    End If
    Exit Function
errorHandler:
    Debug.Print "Fehler bei remove, sKey = " & sKey
End Function

''werte aus range laden, werden nur 1 - Key und 2 - Value werde beruecksichtigt
'Public Function loadFromRange(rngKeyValueRange As Range)
'    Dim i As Long
'    Dim sKey As String
'    Dim sValue As Variant
'    Dim aKeyValueRange As Variant
'
'    On Error GoTo loadFromRange_Error
'
'    aKeyValueRange = rngKeyValueRange.value2
'
'    For i = LBound(aKeyValueRange) To UBound(aKeyValueRange)
'        sKey = CStr(aKeyValueRange(i, 1))
'        sValue = CStr(aKeyValueRange(i, 2))
'        If sKey <> "" Or sValue <> "" Then
'            Call add(sValue, sKey)
'        End If
'    Next i
'
'    Exit Function
'
'loadFromRange_Error:
'    On Error Resume Next
'    Debug.Print "Fehler bei loadFromRange, rngKeyValueRange = " & rngKeyValueRange.Address
'    On Error GoTo 0
'End Function

Public Function getSize()
    On Error GoTo getSize_Error

    getSize = UBound(aInhalt) - LBound(aInhalt) + 1

    On Error GoTo 0
    Exit Function

getSize_Error:
    Debug.Print "Fehler bei getSize"
    
End Function

'get copy of current collection !! not the objects in it!
Public Function Copy() As clsMyCollection
    Dim newCol As New clsMyCollection
    Dim arrayMaxIdx As Long
    Dim aCurrentKeyValue As Variant
    Dim i As Long
    arrayMaxIdx = UBound(aInhalt)
    
    For i = LBound(aInhalt) To arrayMaxIdx
        aCurrentKeyValue = aInhalt(i)
        Call newCol.add(aCurrentKeyValue(1), CStr(aCurrentKeyValue(0)))
    Next i
    
    Set Copy = newCol
End Function

'merge this collection with mcCollectionToMerge
Public Function merge(mcCollectionToMerge As clsMyCollection)
    Dim i As Long
    
    For i = 0 To mcCollectionToMerge.getSize - 1
        Call add(mcCollectionToMerge.getValue(i), mcCollectionToMerge.getKey(i))
    Next i
End Function

'remove last element
Public Function pop()
    On Error GoTo pop_Error

    Call remove(CLng(getSize - 1))


    Exit Function

pop_Error:
    Debug.Print "Error by pop of Klassenmodul clsMyCollection"
    
End Function

Private Function getPosInArray(sKey As String)
    On Error GoTo errorHandler
    Dim i As Long
    getPosInArray = -1
    If sKey = "" Then Exit Function
    For i = LBound(aInhalt) To UBound(aInhalt)
        If StrComp(aInhalt(i)(0), sKey, vbTextCompare) = 0 Then
            getPosInArray = i
            Exit For
        End If
    Next i
    
    Exit Function
errorHandler:
    Debug.Print "Fehler bei getPosInArray, sKey = " & sKey
End Function

'pruefen ob "key" ein index ist
Private Function isIndex(key As Variant) As Boolean
    isIndex = False
    If IsNumeric(key) And VarType(key) <> vbString And Val(key) = key Then
        isIndex = True
    End If
End Function


Private Sub Class_Initialize()
    aInhalt = Array()
End Sub

Private Sub Class_Terminate()
    aInhalt = Empty
End Sub



'TESTS
'Sub test_pop()
'    Dim testCol As New clsMyCollection
'    Dim testNewCol As New clsMyCollection
'
'    Call testCol.add(100, "first")
'    Call testCol.add(200, "same")
'    Call testCol.add(500, "xxx")
'
'    Debug.Print testNewCol("first")
'    Debug.Print testNewCol("same")
'
'    Call testCol.pop
'    Call testCol.pop
'
'    Debug.Print ("first") '100
'    Debug.Print "Size: " & testCol.getSize '1
'    Call testCol.pop
'    Debug.Print "Size: " & testCol.getSize '0
'End Sub
'Sub test_Copy()
'    Dim testCol As New clsMyCollection
'    Dim testNewCol As New clsMyCollection
'
'    Call testCol.add(100, "first")
'    Call testCol.add(100, "same")
'    Set testNewCol = testCol.Copy
'    Call testNewCol.setValue(100000, "first")
'    Debug.Print testCol("first")
'    Debug.Print testCol("same")
'    Debug.Print testNewCol("first")
'    Debug.Print testNewCol("same")
'End Sub

'Sub test_addToSomeIndex()
'    Dim testCol As New clsMyCollection
'    Dim testNewCol As New clsMyCollection
'
'    Call testCol.add(125555, 5)
'    Call testCol.add(121111, 1)
'    Call testCol.add("text pos str(2)", 1)
'
'End Sub


'Sub test_merge()
'    Dim a As New clsMyCollection
'    Dim b As New clsMyCollection
'
'    Call a.add(124, "124")
'    Call a.add(125, "125")
'    Call a.add(126, "126")
'
'    Call b.add(124, "124")
'    Call b.add(128, "128")
'    Call b.add(127, "127")
'
'    Call a.merge(b)
'    Debug.Print "size 5 = " & a.getSize
'    Debug.Print "a(0) value key = 124, "; 124; " => " & a.getValue(0) & a.getKey(0)
'    Debug.Print "a(4) value key = 127, "; 127; " => " & a.getValue(4) & a.getKey(4)
'End Sub
