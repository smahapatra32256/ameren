VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cuLockManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrMODULENAME As String = "utilLockManager.cuLockManager"

Public Function AddLock( _
                    ByVal v_lngUserId As Long, _
                    ByVal v_lngWindowHandle As Long, _
                    ByVal v_varLockTableNames As Variant, _
                    ByVal v_lngDuration As Long, _
                    ByRef r_rsLockedEmployees As Recordset, _
                    Optional ByVal v_blnRaiseError As Boolean = False) _
                    As Boolean
    Dim l_rsLocksInEffect As ADODB.Recordset
    Dim l_rsExpiredLocks As ADODB.Recordset
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_intNumOfEmp As Integer
    

    On Error GoTo eh

    
    AddLock = True

    ' Check for Locks in Effect
    Set l_rsLocksInEffect = LocksInEffect(v_varLockTableNames)
    If Not l_rsLocksInEffect Is Nothing Then
        '  There were employees that had Locks in Effect
        Set r_rsLockedEmployees = l_rsLocksInEffect
        Set l_rsLocksInEffect = Nothing
        Set l_rsExpiredLocks = Nothing
        
        AddLock = False
        Exit Function
    End If
    
    '  No Locks in Effect
    '  Check to see if Employees have Expired Locks on TTR_LOCKS
    
    Set l_rsExpiredLocks = ExpiredLocks(v_varLockTableNames)
    
    Set l_objCN = OpenConnection
    l_intNumOfEmp = 0
    For l_intNumOfEmp = LBound(v_varLockTableNames) To UBound(v_varLockTableNames)
        If IsLockExpired(v_varLockTableNames(l_intNumOfEmp), l_rsExpiredLocks) Then
            '  Expired Locks: Update these Employees
            l_strSQL = "UPDATE      METHDBA.TTR_LOCKS" & _
                        " SET       EMP_NUM = " & v_lngUserId & _
                        "           , WINDOW_HANDLE = " & v_lngWindowHandle & _
                        "           , KEY_TO_LOCK = '" & CStr(v_varLockTableNames(l_intNumOfEmp)) & "'" & _
                        "           , EXP_TIMESTAMP = '" & Format(DateAdd("n", v_lngDuration, Now()), "yyyy-mm-dd-hh.mm") & "'" & _
                        "           , TRIS32_LOCK_IND = 'Y'" & _
                        " WHERE     KEY_TO_LOCK = '" & CStr(v_varLockTableNames(l_intNumOfEmp)) & "'"
                        'ABOVE LINE (TRIS32_LOCK_IND) ADDED BY PREM ON AUG. 24, 2000
        Else
            '  No Expired Locks: Insert the remaining Employees
            l_strSQL = "INSERT INTO METHDBA.TTR_LOCKS " & _
                "(EMP_NUM, WINDOW_HANDLE, KEY_TO_LOCK, EXP_TIMESTAMP, TRIS32_LOCK_IND) " & _
                "VALUES " & _
                "(" & v_lngUserId & ", " & v_lngWindowHandle & ", " & _
                "'" & CStr(v_varLockTableNames(l_intNumOfEmp)) & "', " & _
                "'" & Format(DateAdd("n", v_lngDuration, Now()), "yyyy-mm-dd-hh.mm") & "', 'Y')"
                'ABOVE LINE (TRIS32_LOCK_IND) MODIFIED BY PREM ON AUG. 24, 2000
        End If
'    open and release connection immediately to avoid conflict with
'    connection requirement of ExpiredLocks.
'        Set l_objCN = OpenConnection
        l_objCN.Execute l_strSQL
'        Set l_objCN = Nothing
        l_strSQL = ""
    Next

    Set l_objCN = Nothing
    Set l_rsLocksInEffect = Nothing
    Set l_rsExpiredLocks = Nothing
    GetObjectContext.SetComplete
    
    Exit Function
    
eh:
    Set r_rsLockedEmployees = Nothing
    
    GetObjectContext.SetAbort
    Set l_objCN = Nothing
    Set l_rsLocksInEffect = Nothing
    Set l_rsExpiredLocks = Nothing
    
    
    If Not v_blnRaiseError Then
        AddLock = False
        Exit Function
    End If
    

    TrsRaiseError mstrMODULENAME, "AddLock", trErrutAddLock, True, , DescArgsFrom(l_strSQL), True


End Function
                
Public Function ReleaseUser( _
        ByVal v_lngUserEmployeeNumber As Long, _
        Optional ByVal v_lngWindowHandle As Long, _
        Optional ByVal v_blnRaiseError As Boolean = False) As Boolean
'Answer true if successful, otherwise false.
'Determine how severe it is to attempt to release without releasing.
'e.g. Is is possible for an attempt to release to fail because the
'lock expired and was already released by another user getting a lock?
'If so, just answer a boolean.  If it should never happen, raise an error.
    Dim cn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim recsAffected As Long
    

    On Error GoTo eh
TraceMessage "", ""
TraceMessage "", "Entering cuLockManager.ReleaseUser for " & CStr(v_lngUserEmployeeNumber) & ", hwnd: " & CStr(v_lngWindowHandle)

TraceMessage "", "Begin GetObjectContext"
    
TraceMessage "", "End GetObjectContext"

TraceMessage "", "Begin OpenConnection"
    Set cn = OpenConnection
TraceMessage "", "End OpenConnection"

    If v_lngWindowHandle = 0 Then
        sql = "DELETE FROM  METHDBA.TTR_LOCKS" & _
            " WHERE  EMP_NUM = " & CStr(v_lngUserEmployeeNumber) & _
            " AND TRIS32_LOCK_IND = 'Y'"
    Else
        sql = "DELETE FROM  METHDBA.TTR_LOCKS" & _
            " WHERE  EMP_NUM = " & CStr(v_lngUserEmployeeNumber) & _
            " AND    WINDOW_HANDLE = " & CStr(v_lngWindowHandle) & _
            " AND TRIS32_LOCK_IND = 'Y'"
    End If

TraceMessage "", "Begin Execute"
    cn.Execute sql, recsAffected
TraceMessage "", "End Execute"

TraceMessage "", "Number records affected: " & CStr(recsAffected)
    If recsAffected = 0 Then        'no records deleted
        ReleaseUser = False
    Else
        ReleaseUser = True
    End If
    
    Set cn = Nothing
    Set rs = Nothing
    
TraceMessage "", "Begin SetComplete"
    GetObjectContext.SetComplete
TraceMessage "", "End SetComplete"
    
TraceMessage "", "Exiting cuLockManager.ReleaseUser"
    
    Exit Function
    
eh:
' Any application errors that were raised by components invoked by this
' method may simply be raised.  Or, this component may choose to translate
' the error into 1 of its own application error.  Any other run-time errors
' need to be raised appropriately.  Errors should be logged by the first handler.

TraceMessage "", "Entering cuLockManager.ReleaseUser eh"
    GetObjectContext.SetAbort
    Set cn = Nothing
    Set rs = Nothing
    

    If Not v_blnRaiseError Then
        ReleaseUser = False
        Exit Function
    End If
           
    TrsRaiseError mstrMODULENAME, "ReleaseUser", trErrutReleaseLock, True, , DescArgsFrom(sql), True
               
End Function

Public Function ReleaseKeys( _
        ByVal v_lngUserEmployeeNumber As Long, _
        ByVal v_varEmployees As Variant, _
        Optional ByVal v_lngWindowHandle As Long, _
        Optional ByVal v_blnRaiseError As Boolean = False) As Boolean
' v_varEmployees should be an array of longs or strings
    Dim cn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim recsAffected As Long
    

    On Error GoTo eh
TraceMessage "", ""
TraceMessage "", "Entering cuLockManager.ReleaseKeys for " & CStr(v_lngUserEmployeeNumber) & ", hwnd: " & CStr(v_lngWindowHandle)

TraceMessage "", "Begin GetObjectContext"
    
TraceMessage "", "End GetObjectContext"

TraceMessage "", "Begin OpenConnection"
    Set cn = OpenConnection
TraceMessage "", "End OpenConnection"

    If v_lngWindowHandle = 0 Then
        sql = "DELETE FROM  METHDBA.TTR_LOCKS" & _
            " WHERE  EMP_NUM = " & CStr(v_lngUserEmployeeNumber) & _
            " AND " & BuildInClause("KEY_TO_LOCK", v_varEmployees, True) & _
            " AND TRIS32_LOCK_IND = 'Y'"
    Else
        sql = "DELETE FROM  METHDBA.TTR_LOCKS" & _
            " WHERE  EMP_NUM = " & CStr(v_lngUserEmployeeNumber) & _
            " AND    WINDOW_HANDLE = " & CStr(v_lngWindowHandle) & _
            " AND " & BuildInClause("KEY_TO_LOCK", v_varEmployees, True) & _
            " AND TRIS32_LOCK_IND = 'Y'"
    End If

TraceMessage "", "Begin Execute"
    cn.Execute sql, recsAffected
TraceMessage "", "End Execute"

TraceMessage "", "Number records affected: " & CStr(recsAffected)
    If recsAffected = 0 Then        'no records deleted
        ReleaseKeys = False
    Else
        ReleaseKeys = True
    End If
    
    Set cn = Nothing
    Set rs = Nothing
    
TraceMessage "", "Begin SetComplete"
    GetObjectContext.SetComplete
TraceMessage "", "End SetComplete"
    
TraceMessage "", "Exiting cuLockManager.ReleaseKeys"
    

    Exit Function
    
eh:
' Any application errors that were raised by components invoked by this
' method may simply be raised.  Or, this component may choose to translate
' the error into 1 of its own application error.  Any other run-time errors
' need to be raised appropriately.  Errors should be logged by the first handler.

TraceMessage "", "Entering cuLockManager.ReleaseKeys eh"
    GetObjectContext.SetAbort
    Set cn = Nothing
    Set rs = Nothing
    
    If Not v_blnRaiseError Then
        ReleaseKeys = False
        Exit Function
    End If
           
    TrsRaiseError mstrMODULENAME, "ReleaseKeys", trErrutReleaseLock, True, , DescArgsFrom(sql), True
               
End Function

Public Function ExtendLock( _
                    ByVal v_lngUserEmployeeNumber As Long, _
                    ByVal v_lngWindowHandle As Long, _
                    ByVal v_varLockKeys As Variant, _
                    ByVal v_lngDuration As Long, _
                    Optional ByVal v_blnRaiseError As Boolean = False) As Boolean
' Any unsuccessful key request causes the entire request to abort.
' Answer True when successful, otherwise answer false.
    Dim l_intKeyIndex As Integer
    Dim l_strSQL As String
    Dim recsAffected As Long
    Dim l_objCN As ADODB.Connection
    

    On Error GoTo eh

    
    Set l_objCN = OpenConnection

    For l_intKeyIndex = LBound(v_varLockKeys) To UBound(v_varLockKeys)
        l_strSQL = "UPDATE METHDBA.TTR_LOCKS " _
            & "SET EXP_TIMESTAMP = '" & Format(DateAdd("n", v_lngDuration, Now()), "yyyy-mm-dd-hh.mm") & "' " _
            & "WHERE EMP_NUM = " & CStr(v_lngUserEmployeeNumber) _
            & " AND KEY_TO_LOCK = '" & v_varLockKeys(l_intKeyIndex) & "'" _
            & " AND WINDOW_HANDLE = " & v_lngWindowHandle _
            & " AND TRIS32_LOCK_IND = 'Y'"
        l_objCN.Execute l_strSQL, recsAffected
    Next
    Set l_objCN = Nothing
    If recsAffected = 0 Then
        ExtendLock = False
    Else
        ExtendLock = True
    End If
    GetObjectContext.SetComplete
    
    
    
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    Set l_objCN = Nothing
    
    If Not v_blnRaiseError Then
        ExtendLock = False
        Exit Function
    End If
            
    TrsRaiseError mstrMODULENAME, "ExtendLock", trErrutExtendLock, True, , DescArgsFrom(l_strSQL), True
        
End Function

Private Function OpenConnection() As ADODB.Connection
  
    Dim cm As New util_DBConnectMgr.clsuDBConnectMgr
    Dim cn As ADODB.Connection
    
    On Error GoTo OpenConnection_Eh
    
    Set cn = New ADODB.Connection
    With cn
        .CursorLocation = adUseClient
        .ConnectionString = cm.GetConnectionString("SHADOWDIRECT")
        .Open
    End With
    
    Set OpenConnection = cn
    Set cm = Nothing
    Set cn = Nothing
    Exit Function
    
OpenConnection_Eh:
    Set cm = Nothing
    Set cn = Nothing
    
    TrsRaiseError mstrMODULENAME, "OpenConnection", trErrutOpenConnection, True
 
End Function

Private Function LocksInEffect(ByVal v_varLockTableNames As Variant) As Recordset
'  This function will determine if there are Locks in Effect for the given list of Employees
    
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_intNumOfEmp As Integer

    On Error GoTo eh

    Set l_objCN = OpenConnection
    Set l_objRS = New ADODB.Recordset
    
    l_strSQL = "SELECT  A.KEY_TO_LOCK, A.EMP_NUM, B.EMP_NAME, A.EXP_TIMESTAMP " _
             & "FROM    METHDBA.TTR_LOCKS A, METHDBA.TTR_EMPLOYEE B " _
             & "WHERE   A.EXP_TIMESTAMP > '" & Format(Now(), "yyyy-mm-dd-hh.mm") & "' " _
             & "AND     A.EMP_NUM = B.EMP_NUM " _
             & "AND     KEY_TO_LOCK IN ("

    For l_intNumOfEmp = LBound(v_varLockTableNames) To UBound(v_varLockTableNames)
        If l_intNumOfEmp <> UBound(v_varLockTableNames) Then
            l_strSQL = l_strSQL & "'" & v_varLockTableNames(l_intNumOfEmp) & "', "
        Else
            l_strSQL = l_strSQL & "'" & v_varLockTableNames(l_intNumOfEmp) & "')"
        End If
    Next
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    
    If l_objRS.BOF And l_objRS.EOF Then
        ' No Locks in Effect
        Set LocksInEffect = Nothing
    Else
        ' Locks in Effect
        Set LocksInEffect = l_objRS
    End If
    
    Set l_objRS = Nothing

    Exit Function

eh:
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError mstrMODULENAME, "LocksInEffect", , True, , DescArgsFrom(l_strSQL), True

End Function

Private Function IsLockExpired( _
                    ByVal v_strKey As String, _
                    r_rsExpiredLocks As ADODB.Recordset) As Boolean
' answer true when a record is found in r_rsExpiredLocks.
    r_rsExpiredLocks.Filter = "KEY_TO_LOCK = '" & v_strKey & "'"
    IsLockExpired = Not (r_rsExpiredLocks.BOF And r_rsExpiredLocks.EOF)

End Function

Private Function ExpiredLocks(ByVal v_varLockTableNames As Variant) As Recordset
'  This function will retrieve expired locks for each key in v_varLockTableNames.
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As clsuDBConnectMgr
    Dim l_strSQL As String
    Dim l_intNumOfEmp As Integer

    On Error GoTo eh
    
    Set l_objCN = OpenConnection
    Set l_objRS = New ADODB.Recordset
    
    l_strSQL = "SELECT  KEY_TO_LOCK, EXP_TIMESTAMP " _
             & "FROM    METHDBA.TTR_LOCKS " _
             & "WHERE   EXP_TIMESTAMP <= '" & Format(Now(), "yyyy-mm-dd-hh.mm") & "' " _
             & "AND     KEY_TO_LOCK IN ("

    For l_intNumOfEmp = LBound(v_varLockTableNames) To UBound(v_varLockTableNames)
        If l_intNumOfEmp <> UBound(v_varLockTableNames) Then
            l_strSQL = l_strSQL & "'" & v_varLockTableNames(l_intNumOfEmp) & "', "
        Else
            l_strSQL = l_strSQL & "'" & v_varLockTableNames(l_intNumOfEmp) & "')"
        End If
    Next
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set ExpiredLocks = l_objRS
    
    Set l_objCM = Nothing
    Set l_objRS = Nothing
    
    Exit Function

eh:
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    Set l_objCM = Nothing
    
    TrsRaiseError mstrMODULENAME, "ExpiredLocks", , True, , DescArgsFrom(l_strSQL), True

End Function
