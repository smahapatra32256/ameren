VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdDepartment"

' Static data constants
Private Const CRECORDNOTFOUND = "RECORDNOTFOUND"
Private Const CDEFAULTDATE As String = "01/01/0100"

' Department constants
Private Const CDP_GROUPNAME = "DEPARTMENTCACHE"
' index for record dynamic array
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const CDP_DEPT_NUM = 1
Private Const CDP_DEPT_NAME = 2
Private Const CDP_INTERFACE_TO_CD = 3
Private Const CDP_CMPY_CD = 4
Private Const CDP_PAYROLL_GROUP_CD = 5
Private Const CDP_NON_PROD_LOAD_IND = 6
Private Const CDP_ACCT_TYPE = 7
Private Const CDP_ACCT_NUM = 8
Private Const CDP_ACCT_DIV_NUM = 9
Private Const CDP_OT_REASON_CD = 10
Private Const CDP_WORK_REAS_IND = 11
Private Const CDP_LOCK_STATUS = 12
Private Const CDP_BUDGET_NUM = 13
Private Const CDP_EC_DRAW_DTLS_IND = 14
Private Const CDP_UNION_CD = 15
Private Const CDP_LAST_SLOT = 16

'*********************************************************
' Begin Public procedures
'*********************************************************

Public Function GetDepartment(ByVal v_strDeptNum As String) As ADODB.Recordset

    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetDepartment_EH

    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
   
    l_strSQL = "SELECT DEPT_NUM, DEPT_NAME, INTERFACE_TO_CD, ACCT_TYPE, ACCT_DIV_NUM, "
    l_strSQL = l_strSQL & "NON_PROD_LOAD_IND, ACCT_NUM, UNION_CD, CMPY_CD, OT_REASON_CD "
    l_strSQL = l_strSQL & ", PAYROLL_GROUP_CD, LOCK_STATUS  "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_DEPT "
    l_strSQL = l_strSQL & "WHERE DEPT_NUM = '" & v_strDeptNum & "'"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDepartment = l_objRS
    
    Set l_objRS = Nothing
    

    Exit Function

GetDepartment_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDepartment", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetDepartments( _
        ByRef v_colDeptNums As Collection) As ADODB.Recordset

    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
   
    l_strSQL = _
        "SELECT DEPT_NUM, DEPT_NAME, INTERFACE_TO_CD, " _
        & "ACCT_TYPE, NON_PROD_LOAD_IND, ACCT_NUM, UNION_CD, " _
        & "CMPY_CD, OT_REASON_CD, PAYROLL_GROUP_CD, LOCK_STATUS  " _
        & "FROM METHDBA.TTR_DEPT " _
        & "WHERE DEPT_NUM IN " _
        & AsDelimitedString(v_colDeptNums)

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDepartments = l_objRS
       
    Set l_objRS = Nothing
    

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDepartments", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetAllDepartments() As ADODB.Recordset

    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetAllDepartments_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    l_objCN.Open
   
'Tnv
    '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
   
    l_strSQL = "SELECT CAST(LTRIM(RTRIM(DEPT_NUM)) AS VARCHAR(3)) AS DEPT_NUM, "
    l_strSQL = l_strSQL & "DEPT_NAME, INTERFACE_TO_CD, ACCT_TYPE, "
    l_strSQL = l_strSQL & "NON_PROD_LOAD_IND, ACCT_NUM, CMPY_CD, "
    l_strSQL = l_strSQL & "ACCT_DIV_NUM, OT_REASON_CD, "
    l_strSQL = l_strSQL & "PAYROLL_GROUP_CD, LOCK_STATUS, BUDGET_NUM, "
    l_strSQL = l_strSQL & "WRK_REASON_IND , EC_DRAW_DTLS_IND, "
    l_strSQL = l_strSQL & "INTERFACE_IND "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_DEPT "
    l_strSQL = l_strSQL & "ORDER BY DEPT_NUM "
    l_strSQL = l_strSQL & "FOR FETCH ONLY WITH UR"
      
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    Set GetAllDepartments = l_objRS
    Set l_objRS = Nothing
    Exit Function

GetAllDepartments_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetAllDepartments", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetDepartmentLockInfo( _
        ByVal v_lngEmpNum As Long, ByVal v_datChargeDate As Date) As ADODB.Recordset
    
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetDepartmentLockInfo_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "select A.EMP_NUM, B.DEPT_NUM, B.LOCK_STATUS, "
    l_strSQL = l_strSQL & "C.PAYROLL_GROUP_CD, C.PAY_PER_START_DATE, C.PAY_PER_END_DATE, C.PAY_PROCESSED_IND "
    l_strSQL = l_strSQL & "from METHDBA.TTR_EMPLOYEE A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_DEPT B, "
    l_strSQL = l_strSQL & "METHDBA.TTR_COLL_PERIODS C "
    l_strSQL = l_strSQL & "Where A.EMP_NUM = " & v_lngEmpNum
    l_strSQL = l_strSQL & " and ((A.DEPT_EFF_DATE > '" & Format(v_datChargeDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " and B.DEPT_NUM = A.PREV_DEPT_NUM) "
    l_strSQL = l_strSQL & " or A.DEPT_EFF_DATE <= '" & Format(v_datChargeDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " and B.DEPT_NUM = A.DEPT_NUM) "
    l_strSQL = l_strSQL & " and C.PAYROLL_GROUP_CD = B.PAYROLL_GROUP_CD "
    l_strSQL = l_strSQL & " and C.PAY_PER_START_DATE = "
    l_strSQL = l_strSQL & " (Select MAX(D.PAY_PER_START_DATE)"
    l_strSQL = l_strSQL & " from METHDBA.TTR_COLL_PERIODS D "
    l_strSQL = l_strSQL & " where D.PAYROLL_GROUP_CD = B.PAYROLL_GROUP_CD "
    l_strSQL = l_strSQL & " and D.PAY_PER_START_DATE <= '" & Format(v_datChargeDate, "mm/dd/yyyy") & "')"
    
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDepartmentLockInfo = l_objRS
        
    Set l_objRS = Nothing
    

    Exit Function

GetDepartmentLockInfo_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDepartmentLockInfo", , , , DescArgsFrom(l_strSQL), True

End Function
Public Function GetInterfaceInd(ByVal v_strDeptNum As String) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetInterfaceInd_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
   
    l_strSQL = "SELECT INTERFACE_IND "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_DEPT "
    l_strSQL = l_strSQL & "WHERE DEPT_NUM = '" & v_strDeptNum & "'"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    If l_objRS.EOF Then
        Exit Function
    Else
        Set GetInterfaceInd = l_objRS
    End If
        
    Set l_objRS = Nothing
    

    Exit Function

GetInterfaceInd_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetInterfaceInd", , , , DescArgsFrom(l_strSQL), True

End Function


'*********************************************************
' Begin Public getters
'*********************************************************
Public Function PayrollGroupCodeFor( _
            ByVal v_strDeptNum As String) As String
' Backward compatability only
' Remove this after callers changed; cuDepartmentMgr
    PayrollGroupCodeFor = GetPayrollGroup(v_strDeptNum)

End Function

Public Function IsDepartmentValid( _
            ByVal v_strDeptNum As String) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    IsDepartmentValid = Not IsEmpty(l_varRecord)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsDepartmentValid"
End Function

Public Function GetDepartmentName( _
            ByVal v_strDeptNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetDepartmentName = l_varRecord(CDP_DEPT_NAME)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetDepartmentName"
End Function

Public Function GetSystemInterfaceCode( _
            ByVal v_strDeptNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetSystemInterfaceCode = l_varRecord(CDP_INTERFACE_TO_CD)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetSystemInterfaceCode"
End Function

Public Function GetCompanyCode( _
            ByVal v_strDeptNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetCompanyCode = l_varRecord(CDP_CMPY_CD)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetCompanyCode"
End Function

Public Function GetPayrollGroup( _
            ByVal v_strDeptNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetPayrollGroup = l_varRecord(CDP_PAYROLL_GROUP_CD)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetPayrollGroup"
End Function

Public Function GetUnionCode(ByVal v_strDeptNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetUnionCode = l_varRecord(CDP_UNION_CD)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetUnionCode"
End Function

Public Function GetNonProdLoadInd( _
            ByVal v_strDeptNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetNonProdLoadInd = l_varRecord(CDP_NON_PROD_LOAD_IND)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetNonProdLoadInd"
End Function

Public Function GetAccountType( _
            ByVal v_strDeptNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetAccountType = l_varRecord(CDP_ACCT_TYPE)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetAccountType"
End Function

Public Function GetAccountNumber( _
            ByVal v_strDeptNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetAccountNumber = l_varRecord(CDP_ACCT_NUM)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetAccountNumber"
End Function

Public Function GetAccountDivision( _
            ByVal v_strDeptNum As String) As String
'BLUE069
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetAccountDivision = l_varRecord(CDP_ACCT_DIV_NUM)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetAccountDivision"
End Function

Public Function GetOvertimeReasonCode( _
            ByVal v_strDeptNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetOvertimeReasonCode = l_varRecord(CDP_OT_REASON_CD)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetOvertimeReasonCode"
End Function

Public Function GetECWorkReasonInd( _
            ByVal v_strDeptNum As String) As String
TrsRaiseError m_strCMODULENAME, "GetECWorkReasonInd", trErrObsoleteMethod, , , , True

'' redirect to preserve binary compatibility
'
'On Error GoTo eh
'    GetECWorkReasonInd = GetWorkReasonInd(v_strDeptNum)
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetECWorkReasonInd"
End Function

Public Function GetECDrawDetailsInd( _
            ByVal v_strDeptNum As String) As String
TrsRaiseError m_strCMODULENAME, "GetECDrawDetailsInd", trErrObsoleteMethod, , , , True
'    Dim l_varRecord As Variant
'    Dim lblnExists As Boolean
'
'On Error GoTo eh
'    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
'    If IsEmpty(l_varRecord) Then
'        Exit Function
'    End If
'    GetECDrawDetailsInd = l_varRecord(CDP_EC_DRAW_DTLS_IND)
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetECDrawDetailsInd"
End Function

Public Function GetWorkReasonInd( _
            ByVal v_strDeptNum As String) As String
TrsRaiseError m_strCMODULENAME, "GetWorkReasonInd", trErrObsoleteMethod, , , , True
'    Dim l_varRecord As Variant
'    Dim lblnExists As Boolean
'
'On Error GoTo eh
'    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
'    If IsEmpty(l_varRecord) Then
'        Exit Function
'    End If
'    GetWorkReasonInd = l_varRecord(CDP_WORK_REAS_IND)
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetWorkReasonInd"
'
End Function

Public Function GetDepartmentLockStatus( _
            ByVal v_strDeptNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetDepartmentLockStatus = l_varRecord(CDP_LOCK_STATUS)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetDepartmentLockStatus"
End Function

Public Function GetBudgetNumber( _
            ByVal v_strDeptNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strDeptNum, CDP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetBudgetNumber = l_varRecord(CDP_BUDGET_NUM)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetBudgetNumber"
End Function

'*********************************************************
' End Public getters
'*********************************************************


'*********************************************************
' Begin Public Cache management methods
'*********************************************************

Public Function GetGroupRecords( _
                    ByVal v_strGroupName As String, _
                    ByVal v_varKeys As Variant, _
                    ByRef r_varKeyArray As Variant, _
                    ByRef r_varColumnArray As Variant) As ADODB.Recordset
    Dim l_rsRecords As ADODB.Recordset
    Dim l_colErrArgs As Collection
On Error GoTo eh

    Select Case v_strGroupName
        Case CDP_GROUPNAME
            Set l_rsRecords = GetDepartmentCacheRecords(v_varKeys)
            r_varKeyArray = Array("DEPT_NUM")
            r_varColumnArray = Empty
    End Select
    Set GetGroupRecords = l_rsRecords

    Set l_rsRecords = Nothing
    Exit Function
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    l_colErrArgs.Add "v_varKeys = " & v_varKeys
    l_colErrArgs.Add "l_rsRecords.Source = " & l_rsRecords.Source
    
    Set l_rsRecords = Nothing
    TrsRaiseError m_strCMODULENAME, "GetGroupRecords", , , , l_colErrArgs, True
End Function

Private Function GetDepartmentCacheRecords(Optional ByVal v_varKeys As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String

On Error GoTo eh
    
    If IsMissing(v_varKeys) Or IsEmpty(v_varKeys) Then
        l_strWhereClause = ""
    ElseIf IsArray(v_varKeys) Then
        l_strWhereClause = " WHERE " & BuildInClause("DEPT_NUM", v_varKeys, True)
    Else
        l_strWhereClause = " WHERE DEPT_NUM = '" & Trim(v_varKeys) & "'"
    End If
    
    l_strSQL = ""
    l_strSQL = l_strSQL & "SELECT DEPT_NUM, DEPT_NAME, INTERFACE_TO_CD, CMPY_CD, "
    l_strSQL = l_strSQL & " PAYROLL_GROUP_CD, NON_PROD_LOAD_IND, ACCT_TYPE, RTRIM(ACCT_NUM) as ACCT_NUM, "
    l_strSQL = l_strSQL & " LEFT(ACCT_DIV_NUM, 2) as ACCT_DIV_NUM, OT_REASON_CD, WRK_REASON_IND, LOCK_STATUS, "
    l_strSQL = l_strSQL & " BUDGET_NUM, EC_DRAW_DTLS_IND, UNION_CD"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_DEPT "
    l_strSQL = l_strSQL & l_strWhereClause
    l_strSQL = l_strSQL & " ORDER BY DEPT_NUM "
    
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetDepartmentCacheRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing

    TrsRaiseError m_strCMODULENAME, "GetDepartmentCacheRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'*********************************************************
' Begin Private Caching methods
'*********************************************************

Private Function GetRecord( _
            ByVal v_varKey As Variant, _
            ByVal v_strGroupName As String, _
            ByRef r_blnExists As Boolean) As Variant
' Answer an variant array, representing a record, or empty.
' This method should not invoke InitializeGroup unless the
' group is setup for bulk initialization.  If setup for lazy
' initialization it should only get the single record and add it
' to the cache.  If the record is not found, using lazy, it should
' add a value to the cache to denote invalid so that it does not
' attempt to retrieve it from the database again.  Currently, a string
' constant is added for record not found and the spmCache component
' will answer that it exists, but the property value is a
' RECORDNOTFOUND string.
'
    Dim recCache As utilRecordCache.cuRecordCache
    Dim l_colErrArgs As Collection
           
On Error GoTo eh

'//--------------------------------------------------------------------------------------------------------------//
'//-----http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cossdk/htm/imtxas_2bs6.asp
'//-----Do not use the SafeRef method in COM+. In MTS, SafeRef was used by an object to obtain
'//-----a reference to itself that could be passed outside its context wrapper.
'//-----With contexts and interception fully integrated into COM+, this is no longer necessary.
'//-----
'//-----[C/C++]
'//-----Note   The C++ function SafeRef should be used instead of this method.
'//--------------------------------------------------------------------------------------------------------------//

        Set recCache = CreateObject("utilRecordCache.cuRecordCache")
    
        If IsArray(v_varKey) Then
            GetRecord = recCache.GetRecord2(v_varKey, v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                r_blnExists, Me)
        Else
            GetRecord = recCache.GetRecord(CStr(v_varKey), v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                r_blnExists, Me)
        End If
        
        Set recCache = Nothing
    Exit Function
eh:
    Set recCache = Nothing

    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & CStr(v_varKey)
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    
    TrsRaiseError m_strCMODULENAME, "GetRecord", , , , l_colErrArgs, True
End Function

Public Sub InvalidateDepartmentCache(ByVal v_strDeptNum As String)

On Error GoTo eh
   
    InvalidateCache Array(v_strDeptNum), CDP_GROUPNAME
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateDepartmentCache"
End Sub

'//-----code has been factored into utilRecordCache-----//
Private Sub InvalidateCache( _
                ByVal v_varKey As Variant, _
                ByVal v_strGroupName As String)

    Dim recCache As utilRecordCache.cuRecordCache
    Dim l_colErrArgs As Collection
    Dim l_blnExists As Boolean
On Error GoTo eh

    Set recCache = CreateObject("utilRecordCache.cuRecordCache")
    
    If IsArray(v_varKey) Then '//-----uses metadata to format key, may need to retrieve data to get metadata
        recCache.InvalidateCache2 v_varKey, v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                l_blnExists, Me
    Else
        recCache.InvalidateCache CStr(v_varKey), v_strGroupName
    End If
        
    Set recCache = Nothing

    Exit Sub
eh:
    Set recCache = Nothing
    
    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & CStr(v_varKey)
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    
    TrsRaiseError m_strCMODULENAME, "InvalidateCache", , , , l_colErrArgs, True
End Sub

'//-----used by RecordCache-----//

Private Function IsLazyInitialized(ByVal v_strGroupName As String) As Boolean
' If group is not lazy initialized, then it is bulk initialized.
On Error GoTo eh
    Select Case Trim(v_strGroupName)
    Case CDP_GROUPNAME
        IsLazyInitialized = False
    Case Else
        IsLazyInitialized = True
    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsLazyInitialized"
End Function


Private Function GroupSupportsLoadRecord(ByVal v_strGroupName As String) As Boolean
'All group caches that support Invalidate must use parameterized SQL in corresponding GetxxxRecords (GetGroupRecords)
On Error GoTo eh
    Select Case Trim(v_strGroupName)
    Case CDP_GROUPNAME
        GroupSupportsLoadRecord = True
    Case Else
        GroupSupportsLoadRecord = False
    End Select
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GroupSupportsLoadRecord"
End Function


Private Function OneRowForEachKey(ByVal v_strGroupName As String) As Boolean
'//-----normally each key/row is cached into a 1-dimensional array that represents each column-----//
'//-----union components will have multiple rows per key and will be cached into a 2-dimensional array-----//
On Error GoTo eh
    Select Case Trim(v_strGroupName)
    Case CDP_GROUPNAME
        OneRowForEachKey = True
    Case Else
        OneRowForEachKey = False
    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OneRowForEachKey"
End Function

Private Function GetCacheRecords(ByVal v_strSQL As String, Optional eLockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long
    
On Error GoTo eh
    
    #If TRDEBUG Then
        stTime = GetTickCount()
    #End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open v_strSQL, l_objCN, adOpenStatic, eLockType
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    
    If l_objRS.BOF And l_objRS.EOF Then
        Set l_objRS = Nothing
    End If
    
    Set GetCacheRecords = l_objRS
    
    Set l_objRS = Nothing
    Set l_objCN = Nothing
    Set l_objCM = Nothing
    
    #If TRDEBUG Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        Debug.Print "GetCacheRecords: " & CStr(dfTime / 1000)
    #End If
    
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCacheRecords", , , , DescArgsFrom(v_strSQL), True
End Function

Public Function GetEmpsinDept( _
        ByVal v_strDeptNum As String, ByVal v_dtChargeDate As Date) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetEmpsinDept_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "select DEPT_NUM, E.EMP_NUM, E.EMP_NAME, EWG.EMP_NUM, EWG.WG_NUM, "
    l_strSQL = l_strSQL & "EWG.EFF_START_DATE, EWG.EFF_END_DATE, WG.WG_NUM, "
    l_strSQL = l_strSQL & "WG.WG_DESC, DEMAND_IND "
    l_strSQL = l_strSQL & "From METHDBA.TTR_EMPLOYEE E, "
    l_strSQL = l_strSQL & "METHDBA.TTR_EMP_WG EWG, "
    l_strSQL = l_strSQL & "METHDBA.TTR_WORK_GROUP WG "
    l_strSQL = l_strSQL & "Where DEPT_NUM = '" & v_strDeptNum & "'"
    l_strSQL = l_strSQL & " and (E.EMP_STATUS = 'A' "
    l_strSQL = l_strSQL & " or (E.EMP_STATUS IN ('T','L') "
    l_strSQL = l_strSQL & " and E.STATUS_EFF_DATE > '" & Format(v_dtChargeDate, "mm/dd/yyyy") & "'))"
    l_strSQL = l_strSQL & " and EWG.EMP_NUM = E.EMP_NUM "
    l_strSQL = l_strSQL & " and WG.WG_NUM = EWG.WG_NUM "
    l_strSQL = l_strSQL & " and '" & Format(v_dtChargeDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " between EWG.EFF_START_DATE AND EWG.EFF_END_DATE "
    l_strSQL = l_strSQL & " Union "
    l_strSQL = l_strSQL & " SELECT DEPT_NUM, E.EMP_NUM, E.EMP_NAME, A.EMP_NUM, "
    l_strSQL = l_strSQL & " A.WG_NUM, A.START_DATE, A.END_DATE, WG.WG_NUM, "
    l_strSQL = l_strSQL & " WG.WG_DESC, DEMAND_IND "
    l_strSQL = l_strSQL & " From METHDBA.TTR_EMPLOYEE E, "
    l_strSQL = l_strSQL & " METHDBA.TTR_LINE_ASGNM A, "
    l_strSQL = l_strSQL & " METHDBA.TTR_WORK_GROUP WG "
    l_strSQL = l_strSQL & " Where DEPT_NUM = '" & v_strDeptNum & "'"
    l_strSQL = l_strSQL & " And (E.EMP_STATUS = 'A' "
    l_strSQL = l_strSQL & " Or (E.EMP_STATUS IN ('T','L') "
    l_strSQL = l_strSQL & " And E.STATUS_EFF_DATE > '" & Format(v_dtChargeDate, "mm/dd/yyyy") & "'))"
    l_strSQL = l_strSQL & " And A.EMP_NUM = E.EMP_NUM "
    l_strSQL = l_strSQL & " And WG.WG_NUM = A.WG_NUM "
    l_strSQL = l_strSQL & " And '" & Format(v_dtChargeDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " between A.START_DATE AND A.END_DATE "
    l_strSQL = l_strSQL & " Order BY 5,4"
    

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmpsinDept = l_objRS
        
    Set l_objRS = Nothing

    Exit Function
    
GetEmpsinDept_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmpsinDept", , , , DescArgsFrom(l_strSQL), True


End Function
Public Function GetEmpsinDept2( _
        ByVal v_strDeptNum As String, ByVal vdtStartDate As Date, _
                                      ByVal vdtEndDate As Date) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetEmpsinDept2_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "select DEPT_NUM, E.EMP_NUM, E.EMP_NAME, EWG.EMP_NUM, EWG.WG_NUM, "
    l_strSQL = l_strSQL & "EWG.EFF_START_DATE, EWG.EFF_END_DATE, WG.WG_NUM, "
    l_strSQL = l_strSQL & "WG.WG_DESC, DEMAND_IND "
    l_strSQL = l_strSQL & "From METHDBA.TTR_EMPLOYEE E, "
    l_strSQL = l_strSQL & "METHDBA.TTR_EMP_WG EWG, "
    l_strSQL = l_strSQL & "METHDBA.TTR_WORK_GROUP WG "
    l_strSQL = l_strSQL & "Where DEPT_NUM = '" & v_strDeptNum & "'"
    l_strSQL = l_strSQL & " and (E.EMP_STATUS = 'A' "
    l_strSQL = l_strSQL & " or (E.EMP_STATUS IN ('T','L') "
    l_strSQL = l_strSQL & " and E.STATUS_EFF_DATE >= '" & Format(vdtEndDate, "mm/dd/yyyy") & "'))"
    l_strSQL = l_strSQL & " and EWG.EMP_NUM = E.EMP_NUM "
    l_strSQL = l_strSQL & " and WG.WG_NUM = EWG.WG_NUM "
    l_strSQL = l_strSQL & " AND ('" & Format(vdtStartDate, "mm/dd/yyyy") & "' >= EWG.EFF_START_DATE "
    l_strSQL = l_strSQL & " OR EWG.EFF_START_DATE BETWEEN '" & Format(vdtStartDate, "mm/dd/yyyy")
    l_strSQL = l_strSQL & "' AND '" & Format(vdtEndDate, "mm/dd/yyyy") & "') AND ('"
    l_strSQL = l_strSQL & Format(vdtEndDate, "mm/dd/yyyy") & "' <= EWG.EFF_END_DATE "
    l_strSQL = l_strSQL & " OR EWG.EFF_END_DATE BETWEEN '" & Format(vdtStartDate, "mm/dd/yyyy")
    l_strSQL = l_strSQL & "' AND '" & Format(vdtEndDate, "mm/dd/yyyy") & "')"
    l_strSQL = l_strSQL & " Union "
    l_strSQL = l_strSQL & " SELECT DEPT_NUM, E.EMP_NUM, E.EMP_NAME, A.EMP_NUM, "
    l_strSQL = l_strSQL & " A.WG_NUM, A.START_DATE, A.END_DATE, WG.WG_NUM, "
    l_strSQL = l_strSQL & " WG.WG_DESC, DEMAND_IND "
    l_strSQL = l_strSQL & " From METHDBA.TTR_EMPLOYEE E, "
    l_strSQL = l_strSQL & " METHDBA.TTR_LINE_ASGNM A, "
    l_strSQL = l_strSQL & " METHDBA.TTR_WORK_GROUP WG "
    l_strSQL = l_strSQL & " Where DEPT_NUM = '" & v_strDeptNum & "'"
    l_strSQL = l_strSQL & " And (E.EMP_STATUS = 'A' "
    l_strSQL = l_strSQL & " Or (E.EMP_STATUS IN ('T','L') "
    l_strSQL = l_strSQL & " And E.STATUS_EFF_DATE >= '" & Format(vdtEndDate, "mm/dd/yyyy") & "'))"
    l_strSQL = l_strSQL & " And A.EMP_NUM = E.EMP_NUM "
    l_strSQL = l_strSQL & " And WG.WG_NUM = A.WG_NUM "
    l_strSQL = l_strSQL & " AND ('" & Format(vdtStartDate, "mm/dd/yyyy") & "' >= A.START_DATE "
    l_strSQL = l_strSQL & " OR A.START_DATE BETWEEN '" & Format(vdtStartDate, "mm/dd/yyyy")
    l_strSQL = l_strSQL & "' AND '" & Format(vdtEndDate, "mm/dd/yyyy") & "') and ('"
    l_strSQL = l_strSQL & Format(vdtEndDate, "mm/dd/yyyy") & "' <= A.END_DATE "
    l_strSQL = l_strSQL & " OR A.END_DATE BETWEEN '" & Format(vdtStartDate, "mm/dd/yyyy")
    l_strSQL = l_strSQL & "' AND '" & Format(vdtEndDate, "mm/dd/yyyy") & "')"
    l_strSQL = l_strSQL & " Order BY 5,4"
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmpsinDept2 = l_objRS
        
    Set l_objRS = Nothing

    Exit Function
    
GetEmpsinDept2_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmpsinDept2", , , , DescArgsFrom(l_strSQL), True
    
End Function


