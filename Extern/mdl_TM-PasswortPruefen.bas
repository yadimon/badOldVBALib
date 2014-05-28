Attribute VB_Name = "mdl_TM-PasswortPruefen"
Option Compare Database
Option Explicit

Public Function PasswortPr�fen(txtHinweistext As String, txtPassworttext As String) As Boolean
'
'    Name: PasswortPr�fen
'   Zweck: Funktion zur Pr�fung eines Passwortes
'
'   Autor: Thomas M�ller
'          Th.Moeller@T-Online.de
'
'Erstellt: 07.09.1999
'  Update: 24.11.1999
' Version: 1.2
'
'   Input: 1. txtHinweistext: Text der im Formular erscheint
'          2. txtPassworttext: Passwort, auf das gepr�ft werden soll
'
'  Output: True, wenn Passwort richtig
'          False, wenn Passwort falsch oder abgebrochen
'
'Ben�tigt: frm_TM_PasswortPruefen
'
'Komment.: Das Formular wird ge�ffnet, der Hinweistext angezeigt.
'          Nach Klick auf die Buttons, wird Formular ausgeblendet
'          und eingegebens Passwort mit vorgegebenem verglichen.
    
    'Variablen deklarieren
    Dim txtdummy As String
    
    DoCmd.OpenForm "frm_TM_PasswortPruefen", , , , , acDialog, txtHinweistext

    If Forms!frm_TM_PasswortPruefen.Tag = True Then
        Select Case Forms!frm_TM_PasswortPruefen.txtPasswort
            Case txtPassworttext
                PasswortPr�fen = True
            Case Else
                txtdummy = MsgBox("Passwort falsch!", vbCritical)
                PasswortPr�fen = False
        End Select
    Else
        PasswortPr�fen = False
    End If
    
    DoCmd.Close acForm, "frm_TM_PasswortPruefen"

End Function

