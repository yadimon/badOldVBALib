Attribute VB_Name = "VersionTester"
'---------------------------------------------------------------------------------------
' Module    : VersionTester
' Author    : Dmitry Gorelenkov
' Date      : 05.11.2013
' Changed   : 08.11.2013
' Purpose   : test if local modules are the same as exported files
' Requires  :
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------


Option Compare Database
Option Explicit


  
Private Const sPathToFiles As String = "X:\FILEZ\Classen"
'TODO test for differences
'TODO import file, save module backup

Sub testFileVersions()
    Dim lLine As Long
    lLine = 1
    Dim i As Long
    Dim sFile As String
    Dim sFileContent As String
    Dim sFileDate As String    'file date remote
    Dim dFileChangeDate As Date
    Dim dModuleChangeDate As Date

    Dim countNotMatchedDate As Long
    Dim countNoDateInFileFound As Long
    Dim countNoFileFound As Long
    Dim countNoComment As Long
    Dim countAll As Long


    Dim mdl As Module
    '    Set mdl = Modules.item("clsAccsFormHandler")

    For i = 0 To Modules.Count - 1
        Set mdl = Modules.item(i)

        If mdl.Find("'*Changed*:", lLine, 1, 10, 100, False, False, True) Then
            '            Debug.Print mdl.Name
            dModuleChangeDate = CDate(Right$(mdl.Lines(lLine, 1), 10))

            sFile = FindFile(mdl.Name, sPathToFiles, True)

            If sFile <> vbNullString Then
                sFileContent = GetFileContent(sFile)
                sFileDate = getStringMatched("'.{0,4}Changed.*:\s{0,4}([0-9]{2}\.[0-9]{2}\.[0-9]{2,4})", sFileContent)
                If sFileDate <> vbNullString Then
                    dFileChangeDate = CDate(sFileDate)

                    If dModuleChangeDate <> dFileChangeDate Then
                        Debug.Print "CHANGE DATE DOES NOT MATCH!!!!"
                        Debug.Print "MODULE: " & mdl.Name
                        Debug.Print "REMOTE FILE: " & sFile
                        Debug.Print "LOCAL:  " & dModuleChangeDate
                        Debug.Print "REMOTE: " & dFileChangeDate
                        If dModuleChangeDate > dFileChangeDate Then
                            Debug.Print "Please export your module"
                        Else
                            Debug.Print "Please import the file"
                        End If
                        Debug.Print "-------------------------------" & vbLf
                        countNotMatchedDate = countNotMatchedDate + 1
                    End If
                Else
                    countNoDateInFileFound = countNoDateInFileFound + 1
                    Debug.Print "NO 'Change' comment in File: " & sFile & " found!!"
                End If
            Else
                countNoFileFound = countNoFileFound + 1
'                Debug.Print "'Changed' found, file not found, by module: " & mdl.Name
            End If

        Else
'            Debug.Print mdl.Name & ": no 'Changed' Comment found"
            countNoComment = countNoComment + 1
        End If

        countAll = countAll + 1

    Next i
    
    'REPORT:
    Debug.Print "Modules tested: " & countAll
    Debug.Print "Modules without 'Changed' comment: " & countNoComment
    Debug.Print "Modules with 'Changed' comment: " & countAll - countNoComment
    Debug.Print "No files found of commented modules: " & countNoFileFound
    Debug.Print "Files found: " & countAll - countNoComment - countNoFileFound
    Debug.Print "No 'Changed' comment in founded files: " & countNoDateInFileFound
    Debug.Print "Files with wrong date: " & countNotMatchedDate
    
    
End Sub

Function getStringMatched(sSearchRegex As String, sContent As String)
    Dim objRegex As Object
    Dim objRegM As Object
    Set objRegex = CreateObject("vbscript.regexp")
    
    With objRegex
        .ignorecase = True
        .Pattern = sSearchRegex
        If .test(sContent) Then
            Set objRegM = .Execute(sContent)
            getStringMatched = objRegM(0).submatches(0)
        Else
            getStringMatched = ""
        End If
    End With
End Function

Function FindFile(sFilename As String, SourceFolderName As String, IncludeSubfolders As Boolean) As String

    Dim sResult As String
    Dim files As New Collection
    Dim sFilePath As Variant
    Call ListFilesInFolder(files, SourceFolderName, IncludeSubfolders)
    
    For Each sFilePath In files
        If InStr(1, sFilePath, sFilename & ".cls", vbTextCompare) Or _
        InStr(1, sFilePath, sFilename & ".bas", vbTextCompare) Then
            FindFile = sFilePath
            Exit Function
        End If
    Next sFilePath
End Function


Sub ListFilesInFolder(Daten As Collection, SourceFolderName As String, _
                      IncludeSubfolders As Boolean)


    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim SourceFolder, SubFolder, FileItem
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
    
    For Each FileItem In SourceFolder.files
        Daten.Add FileItem.Path
    Next FileItem
    
    If IncludeSubfolders Then
        For Each SubFolder In SourceFolder.SubFolders
            ListFilesInFolder Daten, SubFolder.Path, True
        Next SubFolder
    End If

    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
End Sub



Function GetFileContent(Name As String) As String
    Dim intUnit As Integer
     
    On Error GoTo ErrGetFileContent
    intUnit = FreeFile
    Open Name For Input As intUnit
    GetFileContent = Input(LOF(intUnit), intUnit)
ErrGetFileContent:
    Close intUnit
    Exit Function
End Function

