VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdSchedChgPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cdSchedChgPay"

Public Function GetSchedChgPay(ByVal v_aryEmpNum As Variant, _
                                ByVal v_dtStartDate As Date, _
                                ByVal v_dtEndDate As Date) As ADODB.Recordset

On Error GoTo eh

    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClauseEMP As String
    
    Dim l_aryEmpNum As Variant
    
    'Get Connection
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing

    'Open Connection
    l_objCN.Open

    'Format SQL Statement
    If IsArray(v_aryEmpNum) Then
        l_strInClauseEMP = BuildInClause("SCP.EMP_NUM", v_aryEmpNum, False)
    Else
        l_strInClauseEMP = "SCP.EMP_NUM = " & str(CLng(v_aryEmpNum))
    End If
    
    l_strSQL = "SELECT SCP.EMP_NUM, EMP_NAME, CHARGE_DATE, BASELINE_SCHED_CD, EXCEPTION_CD, PAYMENT_TYPE"
    l_strSQL = l_strSQL & ", CASE WHEN PAYMENT_TYPE = 'F' THEN 'Full Shift'"
    l_strSQL = l_strSQL & " WHEN PAYMENT_TYPE = 'D' THEN 'Difference'"
    l_strSQL = l_strSQL & " ELSE  'Not Found' END AS PAYMENT_TYPE_DESC"
    l_strSQL = l_strSQL & ", CHAR('Y') AS ORPHAN_IND"
'//    l_strSQL = l_strSQL & " LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP, CHAR(INDEX_TIME) AS INDEX_TIME"
    
    l_strSQL = l_strSQL & ", SCP.CREATED_EMP_NUM, SCP.LAST_MOD_EMP_NUM"
    l_strSQL = l_strSQL & ", CHAR(SCP.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(SCP.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP"
    l_strSQL = l_strSQL & ", SCP.EMP_NUM AS KEY_EMP_NUM, CHARGE_DATE AS KEY_CHARGE_DATE"
    l_strSQL = l_strSQL & ", 'L' AS RECORD_STATUS"
    
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_SCHED_CHG_PMTS SCP, METHDBA.TTR_EMPLOYEE E"
    l_strSQL = l_strSQL & " WHERE " & l_strInClauseEMP
    l_strSQL = l_strSQL & " AND CHARGE_DATE BETWEEN " & dbFormat(v_dtStartDate) & " AND " & dbFormat(v_dtEndDate)
    l_strSQL = l_strSQL & " AND SCP.EMP_NUM =  E.EMP_NUM"
    l_strSQL = l_strSQL & " ORDER BY SCP.EMP_NUM, CHARGE_DATE"
    
    Set l_objRS = New ADODB.Recordset
    With l_objRS
        .CursorLocation = adUseClient
        .Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
        If Not l_objRS Is Nothing Then
           Set GetSchedChgPay = l_objRS
        Else
            Set GetSchedChgPay = Nothing
        End If
    End With

    Set l_objCN = Nothing
    Set l_objRS = Nothing

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing

    TrsRaiseError m_strCMODULENAME, "GetSchedChgPay", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetScope(ByVal v_rsScopeParms As ADODB.Recordset) As ADODB.Recordset

On Error GoTo eh

    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strWhereClause As String
        
    'Get Connection
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing

    'Open Connection
    l_objCN.Open

    'Format SQL Statement
    If RecordsExist(v_rsScopeParms) Then
        With v_rsScopeParms
            .MoveFirst
            Do While Not .EOF
                l_strWhereClause = l_strWhereClause & "(EMP_NUM = " & !EMP_NUM & " AND CHARGE_DATE BETWEEN " & _
                    dbFormat(!START_DATE) & " AND " & dbFormat(!END_DATE) & ") OR "
                .MoveNext
            Loop
            l_strWhereClause = Left(l_strWhereClause, Len(l_strWhereClause) - 4)
        End With
    End If
        
    l_strSQL = "SELECT EMP_NUM, CHARGE_DATE, BASELINE_SCHED_CD, EXCEPTION_CD, PAYMENT_TYPE"
    l_strSQL = l_strSQL & ", CREATED_EMP_NUM, LAST_MOD_EMP_NUM"
    l_strSQL = l_strSQL & ", CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_SCHED_CHG_PMTS"
    l_strSQL = l_strSQL & " WHERE " & l_strWhereClause
    l_strSQL = l_strSQL & " ORDER BY EMP_NUM, CHARGE_DATE"
    
    Set l_objRS = New ADODB.Recordset
    With l_objRS
        .CursorLocation = adUseClient
        .Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
        If Not l_objRS Is Nothing Then
           Set GetScope = l_objRS
        Else
            Set GetScope = Nothing
        End If
    End With

    Set l_objCN = Nothing
    Set l_objRS = Nothing

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing

    TrsRaiseError m_strCMODULENAME, "GetScope", , , , DescArgsFrom(l_strSQL), True
End Function


'Public Function ExistenceCheckSchedChgPay _
'    (ByVal v_aryEmployees As Variant, _
'    ByVal v_dtStartDate As Date, _
'    ByVal v_dtEndDate As Date) As Boolean
'
'On Error GoTo ExistenceCheckSchedChgPay_EH
'
'    Dim l_strSQL As String
'    Dim l_objCN As ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'
'    'Get Connection
'    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
'    Set l_objCN = New ADODB.Connection
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'    Set l_objCM = Nothing
'
'    'Open Connection
'    l_objCN.Open
'
'    'Format SQL Statement
'    l_strSQL = "SELECT COUNT(*) " & _
'                    " FROM methdba.ttr_sched_chg_pmts" & _
'                    " WHERE " & BuildInClause("emp_num", v_aryEmployees, False) & _
'                    " AND charge_date BETWEEN " & _
'                    "'" & Format(v_dtStartDate, "mm/dd/yyyy") & "'  AND '" & Format(v_dtEndDate, "mm/dd/yyyy") & "' "
'
'    Set l_objRS = New ADODB.Recordset
'    With l_objRS
'        .CursorLocation = adUseClient
'        .Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'        If Not (.BOF) And Not (.EOF) Then
'            'If the count is > 0 then schedule changes exist, return true
'            If .Fields(0) > 0 Then
'                ExistenceCheckSchedChgPay = True
'            Else
'                ExistenceCheckSchedChgPay = False
'            End If
'        Else
'            ExistenceCheckSchedChgPay = False
'        End If
'    End With
'
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'
'    Exit Function
'
'ExistenceCheckSchedChgPay_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "ExistenceCheckSchedChgPay", , , , DescArgsFrom(l_strSQL), True
'
'End Function

Private Function dbFormat(ByVal v_varData As Variant) As String
    Dim l_strFormated As String, fieldType As Integer
    
    On Error GoTo eh
    
    fieldType = VarType(v_varData)
    Select Case fieldType
        Case vbString
            l_strFormated = "'" & Replace(v_varData, "'", "''") & "'"
        Case vbInteger, vbLong
            l_strFormated = CStr(v_varData)
        Case vbDate
            l_strFormated = "'" & Format(v_varData, "mm/dd/yyyy") & "'"
        Case Else
            l_strFormated = Format(v_varData)
    End Select
    dbFormat = l_strFormated
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "dbFormat", , , , DescArgsFrom(CStr(v_varData)), True
End Function
