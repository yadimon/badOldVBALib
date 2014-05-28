VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuncsOtlk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFuncsOtlk
' Author    : Dmitry Gorelenkov
' Date      : 30.11.2012
' Changed   : 30.04.2013
' Purpose   : Outlook functions
' Requires  : clsFuncs, Outlook Object Library?
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------

Option Explicit
Private f As clsFuncs
Private ehError As clsErrorHandler

'einen termin hinzufuegen
Public Function addTermin(sSubject As String, sDatum As String, Optional vKalender As Variant = Nothing, _
                          Optional sBody As String, Optional sLocation As String, Optional iDuration As Integer = -1, _
                          Optional iReminderMinutesBeforeStart As Integer = -1, Optional bReminderPlaySound As Boolean, Optional bReminderSet As Boolean)

    Dim olCalenderTermin As AppointmentItem    'calendar item?

    On Error GoTo addTermin_Error


    If Not f.hatWert(vKalender) Then
        Set olCalenderTermin = Outlook.Application.CreateItem(olAppointmentItem)
    Else
        'falls calender object
        If TypeOf vKalender Is AppointmentItem Then
            Set olCalenderTermin = vKalender
            'sonst falls string (pfad?)
        ElseIf StrComp(TypeName(vKalender), "String", vbTextCompare) = 0 Then
            Set olCalenderTermin = OpenMAPIFolder(vKalender).Items.add
        Else
            Exit Function
            f.dbg "Parameter: vKalender kann nicht verwendet werden,  TypeName = " & TypeName(vKalender)
        End If

        If f.CdateTry(sDatum, -1) = -1 Then
            Exit Function
            f.dbg "Kann Termindatum nicht erkennen"
        End If
    End If

    'termin object fuellen, und speichern
    With olCalenderTermin
        .Start = f.CdateTry(sDatum)    'datum des Termines
        .Subject = sSubject    'Betreff für Termin
        If f.hatWert(sBody) Then .Body = sBody    'Zusätzlicher Text im Termininfo
        If f.hatWert(sLocation) Then .Location = sLocation  'Ort des Termines
        If iDuration > -1 Then .Duration = iDuration    'Dauer. Angabe ist jeweils in ganzen Minuten zu setzen
        If iReminderMinutesBeforeStart > -1 Then .ReminderMinutesBeforeStart = iReminderMinutesBeforeStart    'Erinnerung in Minuten
        If f.hatWert(bReminderPlaySound) Then .ReminderPlaySound = bReminderPlaySound    'mit oder ohne Sound
        If f.hatWert(bReminderSet) Then .ReminderSet = bReminderSet    'Erinnerung
        .Save    'Termin speichern
    End With

    addTermin = 1


    Exit Function

addTermin_Error:
    f.dbg "Error: addTermin of Klassenmodul clsFuncsOtlk"
    f.dbg Err.Description
End Function



'got from http://www.pcreview.co.uk/forums/getting-outlook-folder-using-path-t2634615.html
'modified by me
'******************************************************************************
'Custom procedure: OpenMAPIFolder(ByVal strPath)
'Purpose: Return a MAPIFolder from Path argument
'Returns: MAPIFolder object
'******************************************************************************
Public Function OpenMAPIFolder(ByVal strPath) As Outlook.MAPIFolder
    Dim objFldr As MAPIFolder
    Dim strDir As String
    Dim strName As String
    Dim m_olApp As Outlook.Application
    Dim i As Integer
    On Error Resume Next
    Set m_olApp = CreateObject("Outlook.Application")
    If Left(strPath, Len("\")) = "\" Then
        strPath = Mid(strPath, Len("\") + 1)
    Else
        Set objFldr = m_olApp.ActiveExplorer.CurrentFolder
    End If
    While strPath <> ""
        i = InStr(strPath, "\")
        If i Then
            strDir = Left(strPath, i - 1)
            strPath = Mid(strPath, i + Len("\"))
        Else
            strDir = strPath
            strPath = ""
        End If
        If objFldr Is Nothing Then
            Set objFldr = m_olApp.GetNamespace("MAPI").Folders(strDir)
            On Error GoTo 0
        Else
            Set objFldr = objFldr.Folders(strDir)
        End If
    Wend

    Set OpenMAPIFolder = objFldr
End Function


'---------------------------------------------------------------------------------------
' Procedure : createMail
' Author    : Dmitry Gorelenkov
' Date      : 30.11.2012
' Purpose   : Creates Mail
'---------------------------------------------------------------------------------------
'vorlage von http://snipplr.com/view/10041/
Public Function createMail(Optional sRecipient As String = "", Optional sSubject As String = "", _
                            Optional sBody As String = "", Optional sCC As String = "", Optional sAttSource As String = "") As Outlook.MailItem

   
    On Error GoTo createMail_Error

    Dim OutApp As Outlook.Application
    Dim objResult As Outlook.MailItem
    
    Set OutApp = tryToOpenOutlook()
    If OutApp Is Nothing Then
        Exit Function
    End If


    Set objResult = OutApp.CreateItem(Outlook.olMailItem)
    With objResult
        ' Add the To recipient(s) to the message.
        If sRecipient <> "" Then .Recipients.add (sRecipient)
        If sSubject <> "" Then .Subject = sSubject
        If sBody <> "" Then .Body = sBody
        If sCC <> "" Then .CC = sCC
        If sAttSource <> "" Then Call .Attachments.add(sAttSource, olByValue)
        .Display

        '        ' Add the CC recipient(s) to the message.
        '        Set objOutlookRecip = .Recipients.add("Michael Suyama")
        '        objOutlookRecip.Type = olCC
        '
        '        ' Add the BCC recipient(s) to the message.
        '        Set objOutlookRecip = .Recipients.add("Andrew Fuller")
        '        objOutlookRecip.Type = olBCC

        ' Set the Subject, Body, and Importance of the message.

        '        .Importance = olImportanceHigh 'High importance

        '        ' Add attachments to the message.
        '        If Not IsMissing(AttachmentPath) Then
        '        Set objOutlookAttach = .Attachments.add(AttachmentPath)
        '        End If

        ' Resolve each Recipient's name.
        '        For Each objOutlookRecip In .Recipients
        '        objOutlookRecip.Resolve
        '        Next

        '        ' Should we display the message before sending?
        '        If DisplayMsg Then
        '            .Display
        '        Else
        '        .Save
        '        .Send
        '        End If
    End With


    Set createMail = objResult
    Set objResult = Nothing
    Set OutApp = Nothing

    Exit Function

createMail_Error:
    ehError.catchErr ("createMail of Klassenmodul clsFuncsOtlk")


End Function

'try to open outlook, if not opened, ask user to open, and try again, if canceled, return nothing
Private Function tryToOpenOutlook() As Outlook.Application
    
    On Error Resume Next
    Dim OutApp As Outlook.Application
    Dim MsgResult As VbMsgBoxResult

    
    While True
        Set OutApp = GetObject(, "Outlook.Application")
        If Not OutApp Is Nothing Then
            GoTo Final
        End If
        
        MsgResult = MsgBox("Bitte Outlook öffnen!", vbOKCancel, "Outlook nicht gefunden")
        If MsgResult <> vbOK Then
            Set OutApp = Nothing
            GoTo Final
        End If
    Wend
    
    
Final:
    On Error GoTo 0
    
    Set tryToOpenOutlook = OutApp
End Function

Private Sub Class_Initialize()
    Set f = New clsFuncs
    Set ehError = New clsErrorHandler
End Sub

Private Sub Class_Terminate()
    Set f = Nothing
    Set ehError = Nothing
End Sub
