VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMultiUserHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsMultiUserHandler
' Author    : Dmitry Gorelenkov
' Date      : 30.09.2013
' Changed   : 11.11.2013
' Purpose   : Handling of multiuser lock routine. Based on Database entities
' Requires  : Table with (long Primarykey) ID field and (long) LOCK field. clsDB, clsErrorHandler
'---------------------------------------------------------------------------------------


Option Compare Database
Option Explicit


Private m_bInitDone As Boolean 'initialising mark
Private m_sDataTableName As String 'Tabellenname
Private m_sIdFieldName As String 'Feldname mit ID, am besten KEY-Field
Private m_sLockFieldName As String 'Fieldname mit Lock-Mark
Private m_sDateFieldName As String 'fieldname with datetime of record
Private DB As clsDB 'datenbank class
Private ehError As clsErrorHandler 'errorhandler

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::PROPERTIES::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Property Get IdFieldName() As String
    IdFieldName = m_sIdFieldName
End Property

Public Property Get DataTableName() As String
    DataTableName = m_sDataTableName
End Property

Public Property Get LockFieldName() As String
    LockFieldName = m_sLockFieldName
End Property

Public Property Get DateFieldName() As String
    DateFieldName = m_sDateFieldName
End Property

Public Property Get Initialized() As Boolean
    Initialized = m_bInitDone
End Property



':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::FUNCTIONS:::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Function init(sDataTableName As String, sIdFiledName As String, sLockFieldName As String, sDateFieldName As String) As Boolean
    If m_bInitDone Then Exit Function
    
    m_sDataTableName = sDataTableName
    m_sIdFieldName = sIdFiledName
    m_sLockFieldName = sLockFieldName
    m_sDateFieldName = sDateFieldName
    
    m_bInitDone = True
    init = m_bInitDone
End Function


'get all UserIDs that locking anything now
Public Function getAllLockingUsers() As Variant
    On Error GoTo getAllLockingUsers_Error

    If Not m_bInitDone Then:: Call no_init_error:: Exit Function
    
    getAllLockingUsers = DB.SQLArray(DB.SQL(replaceSQL("SELECT DISTINCT :m_sLockFieldName FROM :m_sDataTableName")))


    Exit Function
getAllLockingUsers_Error:
    ehError.catchErr ("getAllLockingUsers of Klassenmodul clsMultiUserHandler")
End Function


'get all IDs, that are locked
Public Function getAllLockedIDs() As Variant
    On Error GoTo getAllLockedIDs_Error

    If Not m_bInitDone Then:: Call no_init_error:: Exit Function
    
    getAllLockedIDs = DB.SQLArray(DB.SQL(replaceSQL("SELECT :m_sIdFieldName FROM :m_sDataTableName")))


    Exit Function
getAllLockedIDs_Error:
    ehError.catchErr ("getAllLockedIDs of Klassenmodul clsMultiUserHandler")
End Function


'returns records that are locked by the user
Public Function getAllLockedBy(lUserID As Long) As Variant
    On Error GoTo getAllLockedBy_Error

    If Not m_bInitDone Then:: Call no_init_error:: Exit Function
    
    getAllLockedBy = DB.SQLArray(DB.SQL(replaceSQL("SELECT DISTINCT :m_sIdFieldName FROM :m_sDataTableName WHERE :m_sLockFieldName = :lUserID", lUserID)))


    Exit Function
getAllLockedBy_Error:
    ehError.catchErr ("getAllLockedBy of Klassenmodul clsMultiUserHandler")
End Function

'returns userId that locking this record
Public Function getLockingUser(lRecordID As Long) As Long
    On Error GoTo getLockingUser_Error

    If Not m_bInitDone Then:: Call no_init_error:: Exit Function
    
    getLockingUser = f.ifNoValue( _
                    DB.SQLValue(replaceSQL("SELECT DISTINCT :m_sLockFieldName FROM :m_sDataTableName WHERE :m_sIdFieldName = :lRecordID", , lRecordID)) _
                    , -1)


    Exit Function
getLockingUser_Error:
    ehError.catchErr ("getLockingUser of Klassenmodul clsMultiUserHandler")
End Function

'test if RecordID is locked
Public Function isLocked(lRecordID As Long)
    isLocked = (getLockingUser(lRecordID) > -1)
End Function

'lock by user
Public Function lockByUser(lUserID As Long, lRecordID As Long) As Boolean
    On Error GoTo lockByUser_Error

    If Not m_bInitDone Then:: Call no_init_error:: Exit Function
    
    If Not isLocked(lRecordID) Then
        lockByUser = doSQL("INSERT INTO :m_sDataTableName (:m_sLockFieldName, :m_sIdFieldName, :m_sDateFieldName) VALUES (:lUserID, :lRecordID, Now())", lUserID, lRecordID)
    End If


    Exit Function
lockByUser_Error:
    ehError.catchErr ("lockByUser of Klassenmodul clsMultiUserHandler")
End Function

'lock by user, with unlock other IDs locked by this user before
Public Function lockByUserExclusive(lUserID As Long, lRecordID As Long) As Boolean
    On Error GoTo lockByUserExclusive_Error

    If Not m_bInitDone Then:: Call no_init_error:: Exit Function
    
    lockByUserExclusive = unlockAllByUser(lUserID) And _
                            lockByUser(lUserID, lRecordID)


    Exit Function
lockByUserExclusive_Error:
    ehError.catchErr ("lockByUserExclusive of Klassenmodul clsMultiUserHandler")
End Function

'unlock all records
'optional only expiried records that are older then lSeconds
Public Function unlockAll(Optional lSeconds As Long = -1) As Boolean
    On Error GoTo unlockAll_Error

    If Not m_bInitDone Then:: Call no_init_error:: Exit Function
    Dim sSQL As String
    
    sSQL = "DELETE FROM :m_sDataTableName"
    
    'falls seconds angegeben, nur die datensaetze loeschen, die aelter sind.
    If lSeconds >= -1 Then
        sSQL = sSQL & " WHERE DateDiff('s', :m_sDateFieldName, Now()) > " & CStr(lSeconds)
    End If
    
    
    unlockAll = doSQL(sSQL)


    Exit Function
unlockAll_Error:
    ehError.catchErr ("unlockAll of Klassenmodul clsMultiUserHandler")
End Function



'unlock record: lRecordID
Public Function unlockRecord(lRecordID As Long) As Boolean
    On Error GoTo unlockRecord_Error

    If Not m_bInitDone Then:: Call no_init_error:: Exit Function
    
    unlockRecord = doSQL("DELETE FROM :m_sDataTableName WHERE :m_sIdFieldName = :lRecordID", , lRecordID)


    Exit Function
unlockRecord_Error:
    ehError.catchErr ("unlockRecord of Klassenmodul clsMultiUserHandler")
End Function


'unlock all records by user
Public Function unlockAllByUser(lUserID As Long) As Boolean
    On Error GoTo unlockAllByUser_Error

    If Not m_bInitDone Then:: Call no_init_error:: Exit Function
    
    unlockAllByUser = doSQL("DELETE FROM :m_sDataTableName WHERE :m_sLockFieldName = :lUserID", lUserID)


    Exit Function
unlockAllByUser_Error:
    ehError.catchErr ("unlockAllByUser of Klassenmodul clsMultiUserHandler")
End Function


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::PRIVATE FUNCTIONS:::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'helper func, to run sql
Private Function doSQL(sSQL_p As String, Optional lUserID As Long = -1, Optional lRecordID As Long = -1) As Boolean
    Dim sSQL As String
    Dim lAffected As Long
    
    sSQL = replaceSQL(sSQL_p, lUserID, lRecordID)
    
    Call DB.SQL(sSQL, , False, lAffected)
    If lAffected = -1 Then
        doSQL = False
    Else
        doSQL = True
    End If
'    Debug.Print "call SQL: " & sSQL
    
End Function

'helper func for sql string replacement
Private Function replaceSQL(sSQL As String, Optional lUserID As Long = -1, Optional lRecordID As Long = -1) As String
    Dim arraySQLFind, arraySQLReplace As Variant
    Dim i As Long
    
    arraySQLFind = Array(":m_sDataTableName", ":m_sLockFieldName", ":m_sIdFieldName", ":lUserID", ":lRecordID", ":m_sDateFieldName")
    arraySQLReplace = Array(m_sDataTableName, m_sLockFieldName, m_sIdFieldName, CStr(lUserID), CStr(lRecordID), CStr(m_sDateFieldName))
    
    
    For i = LBound(arraySQLFind) To UBound(arraySQLFind)
        sSQL = Replace(sSQL, arraySQLFind(i), arraySQLReplace(i), , , vbTextCompare)
    Next i
    
    replaceSQL = sSQL
End Function


'called if any function called, without initialization function
Private Function no_init_error()
    Debug.Print "Class is not initialized yet"
End Function

Private Sub Class_Initialize()
    Set DB = New clsDB
    Set ehError = New clsErrorHandler
End Sub

Private Sub Class_Terminate()
    Set DB = Nothing
    Set ehError = Nothing
End Sub


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::TESTS:::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'Public Sub clsMultiUserHandler_Test()
''TODO, assertions
'    Dim f As New clsFuncs
'    Dim muh As New clsMultiUserHandler
'    Call muh.init("tblKontakteLocks", "klKontaktIDRef", "klUserID", "klDate")
'    Dim aInput As Variant
'    Dim aOneEntity As Variant
'    Dim i As Long
'
'    aInput = Array( _
'                    Array(30786, 1), _
'                    Array(30793, 2), _
'                    Array(30793, 1), _
'                    Array(30786, 1), _
'                    Array(30801, 3), _
'                    Array(30786, 2), _
'                    Array(30801, 1), _
'                    Array(30786, 4), _
'                    Array(30787, 4), _
'                    Array(30788, 4) _
'                    )
'
'    For i = LBound(aInput) To UBound(aInput)
'        aOneEntity = aInput(i)
'        f.dbg "lock ID: " & CStr(aOneEntity(0)) & ", by User: " & CStr(aOneEntity(1))
'        f.dbg CStr(muh.lockByUser(CLng(aOneEntity(0)), CLng(aOneEntity(1))))
'    Next i
'
'
'    f.dbg "all locked IDs: "
'    f.printArray (muh.getAllLockedIDs)
'
'    f.dbg "all locking UserIDs: "
'    f.printArray (muh.getAllLockingUsers)
'
'    f.dbg "user locking ID:30786 "
'    f.dbg CStr(muh.getLockingUser(30786))
'
'    f.dbg "all records locked by user 4"
'    f.printArray (muh.getAllLockedBy(4))
'
'    f.dbg "unlock all by user 4"
'    muh.unlockAllByUser (4)
'
'    f.dbg "unlock all by id 1"
'    muh.unlockRecord (1)
'
'    f.dbg "unlock all"
'    muh.unlockAll
'End Sub
