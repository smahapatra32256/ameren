VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdSchedTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME = "cdSchedTrade"

Public Function SaveTrade( _
                    ByVal v_rsTrade As ADODB.Recordset, _
                    ByVal v_rsTradeDetails As ADODB.Recordset, _
                    ByVal v_rsEmpTradeSchedule As ADODB.Recordset, _
                    ByVal v_rsLineTradeSchedule As ADODB.Recordset, _
                    ByVal v_lngLastModEmpNum As Long, _
                    Optional ByVal v_varOtherSQL As Variant _
                    ) As Boolean
       
    Dim l_objCN As New ADODB.Connection
    Dim l_strSQL As String
    Dim l_strTimeStamp As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim i As Integer
    Dim l_lngRecsAffected As Long
    
    'v_rsTrade and v_rsTradeDetails will only contain inserts
    'v_rsEmpTradeSchedule and v_rsLineTradeSchedule will contain
    '   only inserts and deletes (no updates)
    '   The deletes do not need processed before inserts
    '   This is because created_timestamp makes all rows unique and
    '   it is part of the primary key.
    
    'The code is designed to handle update requests, eventhough the
    ' business process will not pass any.
    '
    'When provided v_varOtherSQL should be a string or an array of strings
    ' in the form of SQL statements.
    '

    On Error GoTo SaveTrade_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    v_rsTrade.Filter = adFilterPendingRecords
    If RecordsExist(v_rsTrade) Then
        v_rsTrade.MoveFirst
        While Not v_rsTrade.EOF
            l_strSQL = GetSQLTradeHeader(v_rsTrade)
            l_objCN.Execute l_strSQL
            v_rsTrade.MoveNext
        Wend
    End If
    
    v_rsTradeDetails.Filter = adFilterPendingRecords
    If RecordsExist(v_rsTradeDetails) Then
        v_rsTradeDetails.MoveFirst
        While Not v_rsTradeDetails.EOF
            l_strSQL = GetSQLTradeDetails(v_rsTradeDetails)
            l_objCN.Execute l_strSQL
            v_rsTradeDetails.MoveNext
        Wend
    End If
    
    v_rsEmpTradeSchedule.Filter = adFilterPendingRecords
    If RecordsExist(v_rsEmpTradeSchedule) Then
        v_rsEmpTradeSchedule.MoveFirst
        While Not v_rsEmpTradeSchedule.EOF
            l_strSQL = GetSQLEmpTradeSchedule(v_rsEmpTradeSchedule)
            l_objCN.Execute l_strSQL
            v_rsEmpTradeSchedule.MoveNext
        Wend
    End If
    
    v_rsLineTradeSchedule.Filter = adFilterPendingRecords
    If RecordsExist(v_rsLineTradeSchedule) Then
        v_rsLineTradeSchedule.MoveFirst
        While Not v_rsLineTradeSchedule.EOF
            l_strSQL = GetSQLLineTradeSchedule(v_rsLineTradeSchedule)
            l_objCN.Execute l_strSQL
            v_rsLineTradeSchedule.MoveNext
        Wend
    End If
    
    If Not IsMissing(v_varOtherSQL) And Not IsEmpty(v_varOtherSQL) Then
        If IsArray(v_varOtherSQL) Then
            For i = LBound(v_varOtherSQL) To UBound(v_varOtherSQL)
                If Not Trim(v_varOtherSQL(i)) = "" Then
                    l_objCN.Execute v_varOtherSQL(i), l_lngRecsAffected
                    If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                        Err.Raise trErrAppCascadeCommitConflict
                    End If
                End If
            Next
        Else
            If Not Trim(v_varOtherSQL) = "" Then
                l_objCN.Execute v_varOtherSQL, l_lngRecsAffected
                If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                    Err.Raise trErrAppCascadeCommitConflict
                End If
            End If
        End If
    End If
    
    GetObjectContext.SetComplete
    l_objCN.Close
    Set l_objCN = Nothing
    SaveTrade = True
    
    
    Exit Function

SaveTrade_EH:

    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    SaveTrade = False

' 6/19/02 RSG
' There is no optimistic locking on trades since they are only added or deleted.
' The trade tables do not support updates, in that there is no LAST_MOD_TIMESTAMP
' defined.  If optimistic locking becomes a requirement here, either add a timestamp
' or change the WHERE clause to include original values of all the attributes.
' Therefore, the only way a commit conflict error can be raised is while processing
' v_varOtherSQL, which is cascade exception updates/deletes.
' Raise commit conflict error with default values.
' Could parse EMP_NUM, CHARGE_DATE, SHIFT_NUM, EXCEPTION_CD from the SQL statement
' later, if necessary, or change the process so that the exception recordset is passed
' and SQL built here.
    Select Case Err.Number
        Case trErrAppCascadeCommitConflict
            Err.Source = _
                Join( _
                    Array( _
                        CStr(TrType_CA), _
                        "0", _
                        "", _
                        "0", _
                        "") _
                    , ";")
            TrsRaiseError m_strCMODULENAME, "SaveTrade", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveTrade", , True, , DescArgsFrom(l_strSQL), True
    End Select

End Function

Private Function GetSQLTradeHeader( _
            ByVal v_rsTrade As ADODB.Recordset _
            ) As String
            
    Dim l_aryPrimaryKey(0) As String
    Dim l_varIgnore(0) As String
    Dim l_strSQL As String
    Dim l_strTimeStamp As String
            
    l_strSQL = ""
    l_aryPrimaryKey(0) = "TRADE_NUM"
    l_varIgnore(0) = "CREATED_TIMESTAMP"

    Select Case v_rsTrade.Status
       Case adRecModified
           l_strSQL = UpdateStatement(v_rsTrade, "METHDBA.TTR_TRADE", l_aryPrimaryKey, l_varIgnore, True)
       Case adRecNew
           l_strSQL = InsertStatement(v_rsTrade, "METHDBA.TTR_TRADE", l_varIgnore)
       Case adRecDeleted
           l_strSQL = DeleteStatement(v_rsTrade, "METHDBA.TTR_TRADE", l_aryPrimaryKey)
        Case Else
           l_strSQL = ""
    End Select

    GetSQLTradeHeader = l_strSQL
            
End Function


Private Function GetSQLTradeDetails( _
            ByVal v_rsTradeDetails As ADODB.Recordset _
            ) As String
            
    Dim l_aryPrimaryKey(2) As String
    Dim l_varIgnore(0) As String
    Dim l_strSQL As String
    Dim l_strTimeStamp As String
                
    l_strSQL = ""
    l_aryPrimaryKey(0) = "TRADE_NUM"
    l_aryPrimaryKey(1) = "CHARGE_DATE"
    l_aryPrimaryKey(2) = "SHIFT_NUM"
    l_varIgnore(0) = "CREATED_TIMESTAMP"

    Select Case v_rsTradeDetails.Status
       Case adRecModified
           l_strSQL = UpdateStatement(v_rsTradeDetails, "METHDBA.TTR_TRADE_DETAIL", l_aryPrimaryKey, l_varIgnore, True)
       Case adRecNew
           l_strSQL = InsertStatement(v_rsTradeDetails, "METHDBA.TTR_TRADE_DETAIL", l_varIgnore)
       Case adRecDeleted
           l_strSQL = DeleteStatement(v_rsTradeDetails, "METHDBA.TTR_TRADE_DETAIL", l_aryPrimaryKey)
        Case Else
           l_strSQL = ""
    End Select

    GetSQLTradeDetails = l_strSQL
            
End Function

Private Function GetSQLEmpTradeSchedule( _
            ByVal v_rsEmpTradeSchedule As ADODB.Recordset _
            ) As String
            
    Dim l_aryPrimaryKey(2) As String
    Dim l_varIgnore(0) As String
    Dim l_strSQL As String
    Dim l_strTimeStamp As String
            
    l_strSQL = ""
    l_aryPrimaryKey(0) = "EMP_NUM"
    l_aryPrimaryKey(1) = "CHARGE_DATE"
    l_aryPrimaryKey(2) = "SHIFT_NUM"
    l_varIgnore(0) = "CREATED_TIMESTAMP"

    Select Case v_rsEmpTradeSchedule.Status
       Case adRecModified
           l_strSQL = UpdateStatement(v_rsEmpTradeSchedule, "METHDBA.TTR_EMP_TRADE", l_aryPrimaryKey, l_varIgnore, True)
       Case adRecNew
           l_strSQL = InsertStatement(v_rsEmpTradeSchedule, "METHDBA.TTR_EMP_TRADE", l_varIgnore)
       Case adRecDeleted
           l_strSQL = DeleteStatement(v_rsEmpTradeSchedule, "METHDBA.TTR_EMP_TRADE", l_aryPrimaryKey)
        Case Else
           l_strSQL = ""
    End Select

    GetSQLEmpTradeSchedule = l_strSQL
            
End Function

Private Function GetSQLLineTradeSchedule( _
            ByVal v_rsLineTradeSchedule As ADODB.Recordset _
            ) As String
            
    Dim l_aryPrimaryKey(3) As String
    Dim l_varIgnore(0) As String
    Dim l_strSQL As String
    Dim l_strTimeStamp As String
            
    l_strSQL = ""
    l_aryPrimaryKey(0) = "WG_NUM"
    l_aryPrimaryKey(1) = "LINE_NUM"
    l_aryPrimaryKey(2) = "CHARGE_DATE"
    l_aryPrimaryKey(3) = "SHIFT_NUM"
    l_varIgnore(0) = "CREATED_TIMESTAMP"

    Select Case v_rsLineTradeSchedule.Status
       Case adRecModified
           l_strSQL = UpdateStatement(v_rsLineTradeSchedule, "METHDBA.TTR_LINE_TRADE", l_aryPrimaryKey, l_varIgnore, True)
       Case adRecNew
           l_strSQL = InsertStatement(v_rsLineTradeSchedule, "METHDBA.TTR_LINE_TRADE", l_varIgnore)
       Case adRecDeleted
           l_strSQL = DeleteStatement(v_rsLineTradeSchedule, "METHDBA.TTR_LINE_TRADE", l_aryPrimaryKey)
        Case Else
           l_strSQL = ""
    End Select

    GetSQLLineTradeSchedule = l_strSQL
            
End Function

Public Function GetDemandLineTrades( _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_aryLineNums As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
        
'Employee is assigned to a Demand Line and therefore get the schedule
' and attributes of the line
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strStartDate As String
    Dim l_strEndDate As String
    Dim l_strWorkGroupSelection As String
    Dim l_strLineNumSelection As String

    On Error GoTo GetDemandLineTrades_EH
    
    If IsArray(v_aryWorkGroups) Then
        l_strWorkGroupSelection = BuildInClause("WG_NUM", v_aryWorkGroups, False) & " "
    Else
        l_strWorkGroupSelection = "WG_NUM = " & str(CLng(v_aryWorkGroups)) & " "
    End If
    If IsArray(v_aryLineNums) Then
        l_strLineNumSelection = BuildInClause("LINE_NUM", v_aryLineNums, False) & " "
    Else
        l_strLineNumSelection = "LINE_NUM = " & str(CLng(v_aryLineNums)) & " "
    End If
    l_strStartDate = "'" & Format(v_datStartDate, "yyyy-mm-dd") & "'"
    l_strEndDate = "'" & Format(v_datEndDate, "yyyy-mm-dd") & "'"
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT WG_NUM, LINE_NUM, CHARGE_DATE, SHIFT_NUM, SCHED_CD " & _
                    ",TRADE_NUM, CREATED_EMP_NUM, CREATED_TIMESTAMP  " & _
                "From METHDBA.TTR_LINE_TRADE " & _
                "WHERE " & l_strWorkGroupSelection & _
                "  AND " & l_strLineNumSelection & _
                "  AND CHARGE_DATE >= " & l_strStartDate & _
                "  AND CHARGE_DATE <= " & l_strEndDate & _
                "  order by WG_NUM, LINE_NUM , CHARGE_DATE, SHIFT_NUM DESC"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDemandLineTrades = l_objRS
    
    Set l_objRS = Nothing

    GetObjectContext.SetComplete

    Exit Function

GetDemandLineTrades_EH:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineTrades", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetEmployeeTrades( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo GetEmployeeTrades_EH
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("emp_num", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

'Need to add     EXCEPTION_CD  for trade exception code

    l_strSQL = "SELECT EMP_NUM, CHARGE_DATE, SHIFT_NUM, SCHED_CD " & _
                    ",TRADE_NUM, CREATED_EMP_NUM, CREATED_TIMESTAMP  " & _
       " FROM METHDBA.TTR_EMP_TRADE" & _
       " WHERE " & l_strEmployeeSelection & _
       " AND (CHARGE_DATE" & _
       " between '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" & _
       " AND '" & Format(v_datEndDate, "yyyy-mm-dd") & "')" & _
       "  ORDER BY  EMP_NUM, CHARGE_DATE, SHIFT_NUM "


    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeTrades = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

GetEmployeeTrades_EH:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeTrades", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetTradeHeader( _
    ByVal v_aryTradeNums As Variant _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strTradeSelection As String

    On Error GoTo GetTradeHeader_EH
    
    If IsArray(v_aryTradeNums) Then
        l_strTradeSelection = BuildInClause("TRADE_NUM", v_aryTradeNums, False) & " "
    Else
        l_strTradeSelection = "TRADE_NUM= " & str(CLng(v_aryTradeNums)) & " "
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT TRADE_NUM, TRADE_DESC, APPROVED_EMP_NUM, " & _
                    "CREATED_EMP_NUM, CREATED_TIMESTAMP  " & _
       " FROM METHDBA.TTR_TRADE" & _
       " WHERE " & l_strTradeSelection

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTradeHeader = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

GetTradeHeader_EH:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTradeHeader", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetTradeDetails( _
    ByVal v_aryTradeNums As Variant _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strTradeSelection As String

    On Error GoTo GetTradeDetails_EH
    
    If IsArray(v_aryTradeNums) Then
        l_strTradeSelection = BuildInClause("TRADE_NUM", v_aryTradeNums, False) & " "
    Else
        l_strTradeSelection = "TRADE_NUM= " & str(CLng(v_aryTradeNums)) & " "
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open


'The first part of this SQL gets Employee Trades
    l_strSQL = _
        "SELECT A.TRADE_NUM, A.CHARGE_DATE, A.WG_NUM," & _
                "A.LINE_NUM, A.EMP_NUM, A.SHIFT_NUM," & _
                "A.SCHED_CD,A.ROLE_TYPE_CD, A.WEEK_TRADE_IND    " & _
        "FROM METHDBA.TTR_TRADE_DETAIL A  " & _
        "WHERE " & l_strTradeSelection & _
        "  AND a.wg_num = 0 "
'The Second part of this SQL gets Line Trades
'The employee assigned to the line will be returned
'If no employee is assigned to the line emp_num will be zero
    l_strSQL = l_strSQL & _
        "Union    " & _
        "SELECT X.TRADE_NUM, X.CHARGE_DATE, X.WG_NUM," & _
                "X.LINE_NUM, VALUE(Y.EMP_NUM,0) AS EMP_NUM," & _
                "X.SHIFT_NUM, X.SCHED_CD," & _
                "X.ROLE_TYPE_CD , X.WEEK_TRADE_IND    " & _
        "FROM   "
    l_strSQL = l_strSQL & _
            "(SELECT WG_NUM, LINE_NUM, CHARGE_DATE," & _
                    "TRADE_NUM,EMP_NUM, SHIFT_NUM, SCHED_CD," & _
                    "ROLE_TYPE_CD , WEEK_TRADE_IND    " & _
            "From METHDBA.TTR_TRADE_DETAIL " & _
            "WHERE " & l_strTradeSelection & _
            "  AND WG_NUM > 0  " & _
            "GROUP BY WG_NUM, LINE_NUM, CHARGE_DATE," & _
                "TRADE_NUM,EMP_NUM, SHIFT_NUM, SCHED_CD," & _
                "ROLE_TYPE_CD , WEEK_TRADE_IND " & _
            ") X  "
'This matches the Trade with the employee assigned to the line
    l_strSQL = l_strSQL & _
        "LEFT OUTER JOIN " & _
            "METHDBA.TTR_LINE_ASGNM Y " & _
            "ON Y.WG_NUM = X.WG_NUM  " & _
            "AND Y.LINE_NUM = X.LINE_NUM  " & _
            "AND Y.START_DATE <= X.CHARGE_DATE  " & _
            "AND Y.END_DATE >= X.CHARGE_DATE  "


    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTradeDetails = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

GetTradeDetails_EH:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTradeDetails", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetTradesByEmployee( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) _
            As ADODB.Recordset

'This gets all trades that is on the selected
' employee's schedule.  In the case of a demand line it is
' the employee that is honoring the trade.
'This may be different than the employee that requested the trade.
'Also if an employee made an Emp Trade prior to being placed on
' a demand line, that trade will not be retrieved

    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String
    Dim l_strEmployeeSelection2 As String
    Dim l_strTradeSubSelects As String

    On Error GoTo GetTradesByEmployee_EH
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("X.emp_num", v_aryEmployees, False) & " "
        l_strEmployeeSelection2 = BuildInClause("Y.emp_num", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "X.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
        l_strEmployeeSelection2 = "Y.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strTradeSubSelects = " ( a.trade_num in " & _
                "(select distinct x.trade_num " & _
                " from methdba.ttr_line_asgnm y " & _
                    ", methdba.ttr_trade_detail x " & _
                " Where " & l_strEmployeeSelection2 & _
                "   and y.START_DATE <= '" & Format(v_datEndDate, "yyyy-mm-dd") & "'" & _
                "   AND y.END_DATE >= '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" & _
                "   AND x.CHARGE_DATE >= y.START_DATE " & _
                "   and x.CHARGE_DATE <= y.END_DATE " & _
                "   and x.CHARGE_DATE >= '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" & _
                "   and x.CHARGE_DATE <= '" & Format(v_datEndDate, "yyyy-mm-dd") & "'" & _
                "   and x.wg_num = y.wg_num " & _
                "   and x.line_num = y.line_num) " & _
                "OR a.trade_num in " & _
                "(select distinct x.trade_num " & _
                "from methdba.ttr_trade_detail x " & _
                "where x.CHARGE_DATE >= '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" & _
                "  and x.CHARGE_DATE <= '" & Format(v_datEndDate, "yyyy-mm-dd") & "'" & _
                "  AND " & l_strEmployeeSelection & " and x.wg_num = 0)) "

'    l_strSQL = "SELECT C.EMP_NAME,A.EMP_NUM, A.CHARGE_DATE, A.SHIFT_NUM, A.SCHED_CD " & _
'                    ",A.TRADE_NUM, B.TRADE_DESC, B.APPROVED_EMP_NUM " & _
'                    ",B.CREATED_EMP_NUM, B.CREATED_TIMESTAMP " & _
'                    ",A.ROLE_TYPE_CD, A.WEEK_TRADE_IND " & _
'                    ",A.WG_NUM, A.LINE_NUM " & _
'       " FROM METHDBA.TTR_TRADE_DETAIL A, METHDBA.TTR_TRADE B " & _
'       ", METHDBA.TTR_EMPLOYEE C " & _
'       " WHERE " & l_strEmployeeSelection & _
'       " AND (A.CHARGE_DATE" & _
'       " between '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" & _
'       " AND '" & Format(v_datEndDate, "yyyy-mm-dd") & "')" & _
'       " AND A.EMP_NUM = C.EMP_NUM " & _
'       " AND A.TRADE_NUM = B.TRADE_NUM " & _
'       "  ORDER BY  A.EMP_NUM, A.CHARGE_DATE, A.SHIFT_NUM, A.TRADE_NUM DESC "

l_strSQL = _
    "SELECT DISTINCT " & _
        "NE1.wg_num, ne1.LINE_NUM, ne1.CHARGE_DATE," & _
        "ne1.TRADE_NUM, ne1.TRADE_DESC, ne1.SHIFT_NUM, ne1.EMP_NUM," & _
        "coalesce(ne3.EMP_NAME,'Unknown Employee') as emp_name" & _
        ", ne1.SCHED_CD, ne1.ROLE_TYPE_CD, ne1.WEEK_TRADE_IND" & _
        ", ne1.APPROVED_EMP_NUM, ne1.CREATED_EMP_NUM, ne1.CREATED_TIMESTAMP  " & _
    "From "
l_strSQL = l_strSQL & _
        "(SELECT DISTINCT " & _
            "A.wg_num, A.LINE_NUM, A.TRADE_NUM, B.TRADE_DESC,A.CHARGE_DATE," & _
            "A.SHIFT_NUM, A.EMP_NUM, A.SCHED_CD, A.ROLE_TYPE_CD," & _
            "A.WEEK_TRADE_IND, B.APPROVED_EMP_NUM, B.CREATED_EMP_NUM," & _
            "B.CREATED_TIMESTAMP  " & _
        "FROM METHDBA.TTR_TRADE_DETAIL A " & _
            ", METHDBA.TTR_TRADE B " & _
        "WHERE " & l_strTradeSubSelects & _
        "  and a.wg_num = 0 " & _
        "  AND b.TRADE_NUM = a.TRADE_NUM " & _
        "  AND A.CHARGE_DATE >= '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" & _
        "  AND A.CHARGE_DATE <= '" & Format(v_datEndDate, "yyyy-mm-dd") & "'" & _
        ") ne1 "
        
    l_strSQL = l_strSQL & _
        "LEFT OUTER JOIN " & _
            "METHDBA.TTR_EMPLOYEE ne3 " & _
        "ON  NE3.EMP_NUM = NE1.emp_num "
        
    l_strSQL = l_strSQL & _
        "Union  "
        
    l_strSQL = l_strSQL & _
        "SELECT DISTINCT NE1.wg_num, NE1.LINE_NUM,NE1.CHARGE_DATE," & _
            "NE1.TRADE_NUM,NE1.TRADE_DESC,NE1.SHIFT_NUM," & _
            "COALESCE(NE2.EMP_NUM,0) AS EMP_NUM," & _
            "COALESCE(NE3.EMP_NAME,'UNASSIGNED LINE') AS EMP_NAME," & _
            "NE1.SCHED_CD, NE1.ROLE_TYPE_CD, NE1.WEEK_TRADE_IND," & _
            "NE1.APPROVED_EMP_NUM, NE1.CREATED_EMP_NUM, NE1.CREATED_TIMESTAMP  " & _
        "From "
        
    l_strSQL = l_strSQL & _
        "(SELECT DISTINCT " & _
            "A.wg_num, A.LINE_NUM, A.CHARGE_DATE, A.TRADE_NUM,B.TRADE_DESC," & _
            "A.SHIFT_NUM, A.SCHED_CD, A.ROLE_TYPE_CD, A.WEEK_TRADE_IND," & _
            "B.APPROVED_EMP_NUM, B.CREATED_EMP_NUM, B.CREATED_TIMESTAMP  " & _
        " FROM METHDBA.TTR_TRADE_DETAIL A " & _
            ", METHDBA.TTR_TRADE B " & _
        " WHERE " & l_strTradeSubSelects & _
        "   and a.wg_num > 0 " & _
        "   AND b.TRADE_NUM = a.TRADE_NUM " & _
        "   AND A.CHARGE_DATE >= '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" & _
        "   AND A.CHARGE_DATE <= '" & Format(v_datEndDate, "yyyy-mm-dd") & "'" & _
        ") NE1 "
      
    l_strSQL = l_strSQL & _
        "LEFT OUTER JOIN " & _
            "METHDBA.TTR_LINE_ASGNM NE2 " & _
            " ON NE2.wg_num  = NE1.wg_num " & _
            "AND NE2.line_num = NE1.line_num " & _
            "AND  NE2.start_date <= NE1.charge_date " & _
            "AND NE2.end_date >= NE1.charge_date "
            
    l_strSQL = l_strSQL & _
        "LEFT OUTER JOIN " & _
            "METHDBA.TTR_EMPLOYEE ne3 " & _
            " ON NE3.EMP_NUM = NE2.emp_num "
            
    l_strSQL = l_strSQL & _
        " order by 1,2,3,4 "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTradesByEmployee = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete


    Exit Function
    
GetTradesByEmployee_EH:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTradesByEmployee", , , , DescArgsFrom(l_strSQL), True


End Function

Public Function GetTradesByEmployeeAndLine( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_aryWorkGroups As Variant, _
    ByVal v_aryLineNums As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date _
    ) As ADODB.Recordset
    
'This function retrieves all trades initiated by the employee
' or from an employee on the line.
'In the case of a demand line changing employees, the employee
' that requested the trade will be displayed.
'The employee currenly on the line will not be displayed

    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String
    Dim l_strWorkGroupSelection As String
    Dim l_strLineNumSelection As String

    On Error GoTo GetTradesByEmployeeAndLine_EH
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.emp_num", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    If IsArray(v_aryWorkGroups) Then
        l_strWorkGroupSelection = BuildInClause("A.WG_NUM", v_aryWorkGroups, False) & " "
    Else
        l_strWorkGroupSelection = "A.WG_NUM = " & str(CLng(v_aryWorkGroups)) & " "
    End If
    If IsArray(v_aryLineNums) Then
        l_strLineNumSelection = BuildInClause("A.LINE_NUM", v_aryLineNums, False) & " "
    Else
        l_strLineNumSelection = "A.LINE_NUM = " & str(CLng(v_aryLineNums)) & " "
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT C.EMP_NAME,A.EMP_NUM, A.CHARGE_DATE, A.SHIFT_NUM, A.SCHED_CD " & _
                    ",A.TRADE_NUM, B.TRADE_DESC, B.APPROVED_EMP_NUM " & _
                    ",B.CREATED_EMP_NUM, B.CREATED_TIMESTAMP " & _
                    ",A.ROLE_TYPE_CD, A.WEEK_TRADE_IND " & _
                    ",A.WG_NUM, A.LINE_NUM " & _
       " FROM METHDBA.TTR_TRADE_DETAIL A, METHDBA.TTR_TRADE B " & _
       ", METHDBA.TTR_EMPLOYEE C " & _
       " WHERE " & l_strEmployeeSelection & _
       " AND A.CHARGE_DATE" & _
       " between '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" & _
       " AND '" & Format(v_datEndDate, "yyyy-mm-dd") & "'" & _
       " AND A.EMP_NUM = C.EMP_NUM " & _
       " AND A.TRADE_NUM = B.TRADE_NUM " & _
       " AND A.WG_NUM = 0 " & _
       " AND A.LINE_NUM = 0 "
       
    l_strSQL = l_strSQL & _
        " UNION " & _
        "SELECT C.EMP_NAME,A.EMP_NUM, A.CHARGE_DATE, A.SHIFT_NUM, A.SCHED_CD " & _
                    ",A.TRADE_NUM, B.TRADE_DESC, B.APPROVED_EMP_NUM " & _
                    ",B.CREATED_EMP_NUM, B.CREATED_TIMESTAMP " & _
                    ",A.ROLE_TYPE_CD, A.WEEK_TRADE_IND " & _
                    ",A.WG_NUM, A.LINE_NUM " & _
       " FROM METHDBA.TTR_TRADE_DETAIL A, METHDBA.TTR_TRADE B " & _
       ", METHDBA.TTR_EMPLOYEE C " & _
       " WHERE " & l_strWorkGroupSelection & _
       " AND " & l_strLineNumSelection & _
       " AND A.CHARGE_DATE" & _
       " between '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" & _
       " AND '" & Format(v_datEndDate, "yyyy-mm-dd") & "'" & _
       " AND A.EMP_NUM = C.EMP_NUM " & _
       " AND A.TRADE_NUM = B.TRADE_NUM "
    l_strSQL = l_strSQL & _
       "  ORDER BY  2, 3, 4, 6 DESC "


    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTradesByEmployeeAndLine = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete


    Exit Function
    
GetTradesByEmployeeAndLine_EH:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTradesByEmployeeAndLine", , , , DescArgsFrom(l_strSQL), True
  
    
    
End Function

Public Function GetTradesByLine( _
    ByVal v_aryWorkGroups As Variant, _
    ByVal v_aryLineNums As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date _
    ) As ADODB.Recordset
    
'This function retrieves all trades initiated from an employee
' on the line.
'In the case of line changes employees, the employee
' that requested the trade will be displayed.
'The employee currenly on the line will not be displayed

    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strWorkGroupSelection As String
    Dim l_strLineNumSelection As String

    On Error GoTo GetTradesByLine_EH
    
    If IsArray(v_aryWorkGroups) Then
        l_strWorkGroupSelection = BuildInClause("A.WG_NUM", v_aryWorkGroups, False) & " "
    Else
        l_strWorkGroupSelection = "A.WG_NUM = " & str(CLng(v_aryWorkGroups)) & " "
    End If
    If IsArray(v_aryLineNums) Then
        l_strLineNumSelection = BuildInClause("A.LINE_NUM", v_aryLineNums, False) & " "
    Else
        l_strLineNumSelection = "A.LINE_NUM = " & str(CLng(v_aryLineNums)) & " "
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "" & _
        "SELECT C.EMP_NAME,A.EMP_NUM, A.CHARGE_DATE, A.SHIFT_NUM, A.SCHED_CD " & _
                    ",A.TRADE_NUM, B.TRADE_DESC, B.APPROVED_EMP_NUM " & _
                    ",B.CREATED_EMP_NUM, B.CREATED_TIMESTAMP " & _
                    ",A.ROLE_TYPE_CD, A.WEEK_TRADE_IND " & _
                    ",A.WG_NUM, A.LINE_NUM " & _
       " FROM METHDBA.TTR_TRADE_DETAIL A, METHDBA.TTR_TRADE B " & _
       ", METHDBA.TTR_EMPLOYEE C " & _
       " WHERE " & l_strWorkGroupSelection & _
       " AND " & l_strLineNumSelection & _
       " AND A.CHARGE_DATE" & _
       " between '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" & _
       " AND '" & Format(v_datEndDate, "yyyy-mm-dd") & "'" & _
       " AND A.EMP_NUM = C.EMP_NUM " & _
       " AND A.TRADE_NUM = B.TRADE_NUM "
    l_strSQL = l_strSQL & _
       "  ORDER BY  A.EMP_NUM, A.CHARGE_DATE, A.SHIFT_NUM, A.TRADE_NUM DESC "


    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTradesByLine = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete


    Exit Function
    
GetTradesByLine_EH:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTradesByLine", , , , DescArgsFrom(l_strSQL), True
  
    
    
End Function



