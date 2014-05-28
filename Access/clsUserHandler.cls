VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUserHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsUserHandler
' Author    : Dmitry Gorelenkov
' Date      : 27.01.2013
' Changed   : 11.11.2013
' Purpose   : Benutzerverwaltung
' Requires  : dazugehoeriges dialog fuer login, clsDB, clsFuncsAccs,  ISettingsHandler - > clsUserSettingsHandler
' Info      : Maybe low quality code :/
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
Private DB As clsDB
Private fAccs As clsFuncsAccs

Private m_stblUsersInGroup As String
Private m_stblGroups As String
Private m_stblUsers As String
Private m_stblUserSettings As String
Private m_stblGroupSettings As String
Private m_stblSettings As String
Private WithEvents m_frmLoginDlg As Access.Form
Attribute m_frmLoginDlg.VB_VarHelpID = -1
Private WithEvents m_btnLogin As Access.CommandButton
Attribute m_btnLogin.VB_VarHelpID = -1
Private WithEvents m_btnCancel As Access.CommandButton
Attribute m_btnCancel.VB_VarHelpID = -1
Private m_txtPassword As Access.TextBox
Private m_txtLogin As Access.TextBox
Private m_sCurrentUser As String
Private m_bInit_done As Boolean
Private m_sCurrentUserGroup As String
Private m_lCurrentUserLvl As Long
Private m_lCurrentUserID As Long
Private m_lCurrentUserGroupID As Long
Private m_sLoginFormName As String
Private m_shUserSettingsHandler As ISettingsHandler
Private m_ushUserSettingsHandler As clsUserSettingsHandler
Private m_bAllowCloseForm As Boolean

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::EVENTS::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Event loggedIn(lUserID As Long)
Public Event loggedOut(lUserID As Long)
Public Event cancelLogin()

Public Property Get UserSettingsHandler() As ISettingsHandler
    Set UserSettingsHandler = m_shUserSettingsHandler
End Property

Public Property Get LoginFormName() As String
    LoginFormName = m_sLoginFormName
End Property

Public Property Get CurrentUserGroup() As String
    CurrentUserGroup = m_sCurrentUserGroup
End Property

Private Property Let CurrentUserGroup(ByVal sCurrentUserGroup As String)
    m_sCurrentUserGroup = sCurrentUserGroup
End Property

Public Property Get CurrentUserLvl() As Long
    CurrentUserLvl = m_lCurrentUserLvl
End Property

Private Property Let CurrentUserLvl(ByVal lCurrentUserLvl As Long)
    m_lCurrentUserLvl = lCurrentUserLvl
End Property

Public Property Get CurrentUser() As String
    CurrentUser = m_sCurrentUser
End Property

Private Property Let CurrentUser(ByVal sCurrentUser As String)
    m_sCurrentUser = sCurrentUser
End Property
Public Property Get CurrentUserID() As Long
    CurrentUserID = m_lCurrentUserID
End Property

Private Property Let CurrentUserID(ByVal lCurrentUserID As Long)
    m_lCurrentUserID = lCurrentUserID
End Property

Public Property Get CurrentUserGroupID() As Long
    CurrentUserGroupID = m_lCurrentUserGroupID
End Property

Private Property Let CurrentUserGroupID(ByVal lCurrentUserGroupID As Long)
    m_lCurrentUserGroupID = lCurrentUserGroupID
End Property

Public Function init(Optional tblUsersInGroup As String, Optional tblGroups As String, Optional tblUsers As String, _
                    Optional tblUserSettings As String, Optional tblGroupSettings As String, Optional tblSettings As String, _
                    Optional ByRef frmLoginDlg As Form = Nothing)
    'TODO dynamic form?
    'TODO create empty tables, and ref. them
    
    'load defaults if no args
    m_stblUsersInGroup = IIf(tblUsersInGroup = "", "tblLoginUsersInGroups", tblUsersInGroup)
    m_stblGroups = IIf(tblGroups = "", "tblLoginGroups", tblGroups)
    m_stblUsers = IIf(tblUsers = "", "tblLoginUsers", tblUsers)
    m_stblUserSettings = IIf(tblUserSettings = "", "tblLoginUserSettings", tblUserSettings)
    m_stblGroupSettings = IIf(tblGroupSettings = "", "tblLoginGroupSettings", tblGroupSettings)
    m_stblSettings = IIf(tblGroupSettings = "", "tblLoginSettings", tblSettings)
    
    m_bAllowCloseForm = False
    
    m_bInit_done = True
    
End Function

'benutzer bei settingsHandler einstellen
Private Function initSettingsHandler(lUserID As Long, lUserGroupID As Long)
    'original classe benutzen
    Call m_ushUserSettingsHandler.init(lUserID, lUserGroupID, _
                        m_stblSettings, m_stblUserSettings, m_stblGroupSettings)
End Function

'pruefen ob irgendein User eingeloggt ist
Public Function testLogin()
    If Me.CurrentUser = "" Then
        requestLogin
    End If
End Function

'shows dialog, for login
Public Function requestLogin() As Boolean
'    loadFormIfNotLoaded
    showLoginForm
End Function

'pruefen ob login und passwort passen
Public Function checkLogin(sUserName As String, sUserPw As String) As String
    Dim lAffected As Long
    Dim sGotPw As String
    Dim sMsg As String
    On Error GoTo Login_Error

    If sUserName = "" Then Exit Function
    
    sGotPw = Nz(DB.SQLValue("SELECT usrPwd FROM " & m_stblUsers & " WHERE usrLogin = """ & sUserName & """", , , lAffected), "")
    If (lAffected = 0) Then
        sMsg = "Benutzer nicht gefunden"
    ElseIf (lAffected > 1) Then
        sMsg = "Fehler, mehrere Benutzer gefunden?"
    Else 'normalfall
        'falls passwort passt
        If StrComp(sGotPw, sUserPw, vbBinaryCompare) = 0 Then
            sMsg = "OK"
        Else
            sMsg = "Passwort ist falsch"
        End If
    End If
    
Final:
    checkLogin = sMsg


    Exit Function

Login_Error:
    ehError.catchErr ("checkLogin of Klassenmodul clsUserHandler")
    sMsg = "Fehler..."
    Resume Final
End Function

'aktiv benutzerdaten loeschen
Public Function logOut() As Boolean
    Dim tmpCurrentUserID As Long
    
    tmpCurrentUserID = CurrentUserID
    
    setCurrentUser ("")
    RaiseEvent loggedOut(tmpCurrentUserID)
    
    showLoginForm
End Function

'aktuellen benutzer mit eigenen variablen setzen,
'bei "" eingabe, aktivebenutzer daten loeschen
Public Function setCurrentUser(sUserName As String)
    Dim tCurrentUser As String
    Dim tCurrentUserID As Long
    Dim tCurrentUserGroup As String
    Dim tCurrentUserGroupID As Long
    Dim tCurrentUserLvl As Long
    
    'falls username vorhanden
    If sUserName <> vbNullString Then
        tCurrentUser = sUserName
        Dim rsAllInfo As ADODB.Recordset
        Set rsAllInfo = DB.SQL("SELECT u.usrLogin, u.usrID, g.grpID, g.grpName, g.grpAccesLevel " & _
                                "FROM tblLoginGroups as g LEFT JOIN (SELECT * FROM tblLoginUsers as u LEFT JOIN tblLoginUsersInGroups as ug ON u.usrID = ug.usringrpUserIDRef) as t ON g.grpID = t.usringrpGroupRef " & _
                                "WHERE u.usrLogin=""" & sUserName & """")
                                
        tCurrentUserID = CLng(Nz(rsAllInfo.Fields(1).Value, -1))
        tCurrentUserGroup = Nz(rsAllInfo.Fields(3).Value, "")
        tCurrentUserGroupID = CLng(Nz(rsAllInfo.Fields(2).Value, -1))
        tCurrentUserLvl = CLng(Nz(rsAllInfo.Fields(4).Value, -1))
        
    'sonst daten loeschen
    Else
        tCurrentUser = ""
        tCurrentUserID = -1
        tCurrentUserGroup = ""
        tCurrentUserGroupID = -1
        tCurrentUserLvl = -1
    End If
    
    CurrentUser = tCurrentUser
    CurrentUserID = tCurrentUserID
    CurrentUserGroup = tCurrentUserGroup
    CurrentUserGroupID = tCurrentUserGroupID
    CurrentUserLvl = tCurrentUserLvl
    
    Call initSettingsHandler(CurrentUserID, CurrentUserGroupID)
End Function

'aktiveBenutzer setting auslesen 'todo optional username?
Public Function getSetting(sSettingString As String) As String

    'falls keiner eingeloggt, fehler..
    On Error GoTo getSetting_Error

'    If CurrentUserID <= 0 Then
'        getSetting = "No User LoggedOn"
'        Exit Function
'    End If
    '-----------------------------------------
    getSetting = m_shUserSettingsHandler.getSetting(sSettingString)


    Exit Function

getSetting_Error:
    ehError.catchErr ("getSetting of Klassenmodul clsUserHandler")
    

End Function

'set value to user setting, if there no setting, create one
Public Function setSetting(sSettingName As String, sValue As String) As String
    On Error GoTo setSetting_Error
    
    If CurrentUserID <= 0 Then
        setSetting = "No User LoggedOn"
        Exit Function
    End If
    
    
    setSetting = m_shUserSettingsHandler.setSetting(sSettingName, sValue)
  
  
    Exit Function
setSetting_Error:
    ehError.catchErr ("setSetting of Klassenmodul clsUserHandler")
    
End Function

'neue setting zu dem Benutzer hinzufuegen
Public Function addSetting(sSettingName As String, sValue As String) As String
    On Error GoTo addSetting_Error
    
    If CurrentUserID <= 0 Then
        addSetting = "No User LoggedOn"
        Exit Function
    End If
    
    
    addSetting = m_shUserSettingsHandler.addSetting(sSettingName, sValue)
    
    
    
   Exit Function
addSetting_Error:
    ehError.catchErr ("addSetting of Klassenmodul clsUserHandler")
    
End Function


'change password sNewPassword by sUser
Public Function setPassword(sUser As String, sNewPassword As String) As Boolean
    f.dbg "setPassword of clsUserHandler: NOT IMPLEMENTED YET"
End Function

Public Function removeUser(sName As String)
    f.dbg "removeUser of clsUserHandler: NOT IMPLEMENTED YET"
End Function

Public Function addUser(sName As String)
    f.dbg "addUser of clsUserHandler: NOT IMPLEMENTED YET"
End Function

'return user name by ID
Public Function getUserNameByID(lUserID As Long)
    getUserNameByID = f.ifNoValue(DB.SQLValue("SELECT usrRealName FROM tblLoginUsers WHERE usrID = " & lUserID), "")
End Function

Private Sub Class_Initialize()
    Set DB = New clsDB
    Set fAccs = New clsFuncsAccs
    Set m_ushUserSettingsHandler = New clsUserSettingsHandler
    Set m_shUserSettingsHandler = m_ushUserSettingsHandler
End Sub

Private Sub Class_Terminate()
'    Set m_stblUsersInGroup = Nothing
'    Set m_stblGroups = Nothing
'    Set m_stblUsers = Nothing
'    Set m_stblUserSettings = Nothing
'    Set m_stblGroupSettings = Nothing
'    Set m_stblSettings = Nothing
    On Error Resume Next
    
    closeLoginForm
    
    unbindControlVars
    Set m_frmLoginDlg = Nothing
    

    Set DB = Nothing
    Set fAccs = Nothing
    Set m_ushUserSettingsHandler = Nothing
    
    Set m_shUserSettingsHandler = Nothing
    err.Clear
End Sub

'onclick event func, fuer login button
Private Sub m_btnLogin_Click()
    Dim sResult As String
    'pruefen die eingabedaten
    sResult = checkLogin(Nz(m_frmLoginDlg.txtLogin, ""), Nz(m_frmLoginDlg.txtPassword, ""))
    If sResult = "OK" Then
        m_sCurrentUser = Nz(m_frmLoginDlg.txtLogin, "")
        'active benutzer setzen
        setCurrentUser (m_sCurrentUser) 'mb extra login func?
        closeLoginForm
        RaiseEvent loggedIn(CurrentUserID)
        
        'TMP TODO how to ?
'        Call doLogActivity(enm_Log_LoggedIn, CurrentUserID)
'        Call doLogActivity(enm_Log_Working, CurrentUserID)
    Else
        'login fehlgeschlagen
        loginFormPrintError (sResult)
    End If
    
    
End Sub

'loginform schliessen
Public Sub closeLoginForm()
    m_bAllowCloseForm = True
    
    If Not m_frmLoginDlg Is Nothing Then
        m_frmLoginDlg.Visible = False
        Set m_frmLoginDlg = Nothing
    End If
    
    Call unbindControlVars
End Sub

'unbind controls from form
Private Function unbindControlVars()
    Set m_btnLogin = Nothing
    Set m_btnCancel = Nothing
    Set m_txtLogin = Nothing
    Set m_txtPassword = Nothing
End Function



'loginform anzeigen
Private Sub showLoginForm()
    Set m_frmLoginDlg = New Form_frmLogin
    m_bAllowCloseForm = False
    m_frmLoginDlg.Visible = True
    m_frmLoginDlg.OnUnload = "[Event Procedure]"
    
    'set local vars of dialog controls
    If Not m_frmLoginDlg Is Nothing Then
        'set controls to vars
        Call unbindControlVars
        
        Set m_btnLogin = m_frmLoginDlg.Controls("btnLogin")
        m_btnLogin.OnClick = "[Event Procedure]"
        Set m_btnCancel = m_frmLoginDlg.Controls("btnCancel")
        m_btnCancel.OnClick = "[Event Procedure]"
        Set m_txtLogin = m_frmLoginDlg.txtLogin
        Set m_txtPassword = m_frmLoginDlg.txtPassword
    End If
    
    
    While Not m_frmLoginDlg Is Nothing
'        If m_frmLoginDlg.name <> Screen.ActiveForm.name Then
            m_frmLoginDlg.SetFocus
'        End If
        DoEvents
        f.Sleep (10)
    Wend

End Sub

'Fehler in Loginform anzeigen
Private Function loginFormPrintError(sError As String)
    m_frmLoginDlg.txtErrorOut.Value = sError
End Function

Private Sub m_frmLoginDlg_Unload(Cancel As Integer)
    If Not m_bAllowCloseForm Then
        Cancel = True
    End If
End Sub

'on click on cancel button
Private Sub m_btnCancel_Click()
    RaiseEvent cancelLogin
End Sub


