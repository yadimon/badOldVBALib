Attribute VB_Name = "mdl_TM-PasswortPruefen"
Option Compare Database
Option Explicit

Public Function PasswortPrüfen(txtHinweistext As String, txtPassworttext As String) As Boolean
'
'    Name: PasswortPrüfen
'   Zweck: Funktion zur Prüfung eines Passwortes
'
'   Autor: Thomas Möller
'          Th.Moeller@T-Online.de
'
'Erstellt: 07.09.1999
'  Update: 24.11.1999
' Version: 1.2
'
'   Input: 1. txtHinweistext: Text der im Formular erscheint
'          2. txtPassworttext: Passwort, auf das geprüft werden soll
'
'  Output: True, wenn Passwort richtig
'          False, wenn Passwort falsch oder abgebrochen
'
'Benötigt: frm_TM_PasswortPruefen
'
'Komment.: Das Formular wird geöffnet, der Hinweistext angezeigt.
'          Nach Klick auf die Buttons, wird Formular ausgeblendet
'          und eingegebens Passwort mit vorgegebenem verglichen.
    
    'Variablen deklarieren
    Dim txtdummy As String
    
    DoCmd.OpenForm "frm_TM_PasswortPruefen", , , , , acDialog, txtHinweistext

    If Forms!frm_TM_PasswortPruefen.Tag = True Then
        Select Case Forms!frm_TM_PasswortPruefen.txtPasswort
            Case txtPassworttext
                PasswortPrüfen = True
            Case Else
                txtdummy = MsgBox("Passwort falsch!", vbCritical)
                PasswortPrüfen = False
        End Select
    Else
        PasswortPrüfen = False
    End If
    
    DoCmd.Close acForm, "frm_TM_PasswortPruefen"

End Function

