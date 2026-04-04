VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ScheduleOption
    AllExceptions = 1
    OnlyPendingExceptions = 2
End Enum

Private Const m_strCMODULENAME As String = "cdSchedule"


Public Function GetScheduleChangePayments( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

On Error GoTo eh

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

    l_strSQL = "select EMP_NUM, BASELINE_SCHED_CD, " _
        & "EXCEPTION_CD, PAYMENT_TYPE , CHARGE_DATE " _
        & "From methdba.ttr_sched_chg_pmts " _
        & "where " & l_strEmployeeSelection _
        & " and (charge_date " _
        & "between " & DateAsEnclosedString(v_datStartDate) _
        & " AND " & DateAsEnclosedString(v_datEndDate) & ") " _
        & "order by EMP_NUM, charge_date"
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetScheduleChangePayments = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduleChangePayments", , , , DescArgsFrom(l_strSQL), True

End Function








Public Function GetAvailSchedInfo(ByVal v_colScheduleCodes As Collection) As ADODB.Recordset

'code assumes at least One Schedule to retrieve

    Dim l_intIndex As Integer
    Dim l_strSQLinClause As String
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetBaseSchedule_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQLinClause = "("
    For l_intIndex = 1 To v_colScheduleCodes.Count
        If l_intIndex = 1 Then
            l_strSQLinClause = "(" & str(v_colScheduleCodes.Item(l_intIndex))
        Else
            l_strSQLinClause = l_strSQLinClause & "," & str(v_colScheduleCodes.Item(l_intIndex))
        End If
            
    Next
    l_strSQLinClause = l_strSQLinClause & ")"
    
    l_strSQL = "select * from methdba.ttr_avail_sched " & _
                "where sched_cd in " & l_strSQLinClause & " " & _
                "order by sched_cd"
                                   
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetAvailSchedInfo = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetBaseSchedule_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAvailSchedInfo", , , , DescArgsFrom(l_strSQL), True

    
End Function

Public Function GetEmployeePrimaryWGBaseSchedule( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset
        
'Employee's primary work group(s) base schedule

    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strStartDate As String
    Dim l_strEndDate As String
    Dim l_strEmployeeSelection As String

    On Error GoTo GetEmployeePrimaryWGBaseSchedule_EH
    
    l_strStartDate = "'" & Format(v_datStartDate, "yyyy-mm-dd") & "'"
    l_strEndDate = "'" & Format(v_datEndDate, "yyyy-mm-dd") & "'"
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("emp_num", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT A.WG_NUM, A.EFF_START_DATE, A.EFF_END_DATE, B.SCHED_CD, " & _
                "      B.EFF_DATE, A.EMP_NUM, B.ROT_CD, B.ROT_WEEK,B.END_DATE " & _
                "From METHDBA.TTR_EMP_WG A,METHDBA.TTR_WG_BASE_SCHED B " & _
                " WHERE " & l_strEmployeeSelection & _
                "  AND A.WG_NUM = B.WG_NUM " & _
                "  AND A.PRIMARY_IND = 'Y' " & _
                "  AND A.EFF_START_DATE <= " & l_strEndDate & " " & _
                "  AND A.EFF_END_DATE >= " & l_strStartDate & _
                "  AND B.EFF_DATE <= " & l_strEndDate & " " & _
                "  AND B.END_DATE >= " & l_strStartDate & " " & _
                "ORDER BY EMP_NUM, A.EFF_START_DATE, B.EFF_DATE DESC "
                
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, v_adLockType

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeePrimaryWGBaseSchedule = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetEmployeePrimaryWGBaseSchedule_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeePrimaryWGBaseSchedule", , , , DescArgsFrom(l_strSQL), True
    
End Function


Public Function GetEmployeeScheduledExceptions( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date _
        ) As ADODB.Recordset
        
           
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strSchedTypeforSQL As String
    Dim l_intIndex As Integer
    Dim l_strEmployeeSelection As String
    

    On Error GoTo GetEmployeeScheduledExceptions_EH
    
        
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

    l_strSQL = _
        "select EMP_NUM, CHARGE_DATE, SHIFT_NUM, HOURS_OFFSET_NUM, " _
            & " HOURS, EXCEPTION_CD, OVER_EXC_CD, FMLA_IND," _
            & " CASE_NUM, WG_NUM, JOB_NUM, RELIEF_COMP," _
            & " COMP_CD, SHIFT_COMP, SCHED_SRC_CD, SCHED_CD " _
                    & ",CREATED_EMP_NUM, LAST_MOD_EMP_NUM " _
        & ",char(CREATED_TIMESTAMP) as CREATED_TIMESTAMP " _
        & ",char(LAST_MOD_TIMESTAMP) as LAST_MOD_TIMESTAMP " & _
        "from methdba.ttr_emp_sched_exc  " & _
        "where " & l_strEmployeeSelection _
        & " AND (charge_date " _
        & " between '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" _
        & " AND '" & Format(v_datEndDate, "yyyy-mm-dd") & "')  " & _
        "ORDER BY  EMP_NUM, CHARGE_DATE, SHIFT_NUM DESC, HOURS_OFFSET_NUM"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeScheduledExceptions = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetEmployeeScheduledExceptions_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeScheduledExceptions", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function GetRotationSchedules( _
        ByVal v_colRotationCodes As Collection _
        ) As ADODB.Recordset

'This function is similar to GetSchedulesInRotation except
' GetSchedulesInRotation does the following
'   1. Needed for only one Rotation Code
'   2. an extra coloumn CURRENT_SCHED is returned
'   3. the recordset is updateable and not read only
'The seperation is needed to conserve database resources
'GetRotationSchedules is called when getting the schedule for the employee
'GetSchedulesInRotation is called to display the schedule codes
'  and the current schedule on a date

    Dim l_intIndex As Integer
    Dim l_strSQLinClause As String
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetRotationSchedules_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQLinClause = "(-1"  'Seed with a dummy value. An empty recordset is
                              'returned if there are no rotations to retrieve
    For l_intIndex = 1 To v_colRotationCodes.Count
        l_strSQLinClause = l_strSQLinClause & "," & str(v_colRotationCodes.Item(l_intIndex))
    Next
    l_strSQLinClause = l_strSQLinClause & ")"
    
    l_strSQL = "select rot_cd, rot_week, sched_cd " & _
                "from methdba.ttr_rotation_sched " & _
                "where rot_cd in " & l_strSQLinClause & " " & _
                "order by rot_cd, rot_week"
                                   
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetRotationSchedules = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetRotationSchedules_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetRotationSchedules", , , , DescArgsFrom(l_strSQL), True

    
End Function

Public Function GetWGBaseSchedules( _
        ByVal v_colWorkGroups As Collection, _
        ByVal v_datDate As Date _
        ) As ADODB.Recordset

    Dim l_intIndex As Integer
    Dim l_strSQLinClause As String
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetWGBaseSchedules_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQLinClause = "(-1"
    For l_intIndex = 1 To v_colWorkGroups.Count
        l_strSQLinClause = l_strSQLinClause & "," & str(v_colWorkGroups.Item(l_intIndex))
    Next
    l_strSQLinClause = l_strSQLinClause & ")"
    
    l_strSQL = "select WG_NUM , EFF_DATE as START_DATE, ROT_CD, ROT_WEEK, SCHED_CD " & _
                " , CREATED_EMP_NUM, LAST_MOD_EMP_NUM " & _
                " , Char(CREATED_TIMESTAMP) as CREATED_TIMESTAMP " & _
                " , Char(LAST_MOD_TIMESTAMP) as LAST_MOD_TIMESTAMP " & _
                " , 'L' as REC_STATUS, END_DATE " & _
                "from methdba.ttr_wg_base_sched " & _
                "where wg_num in " & l_strSQLinClause & " " & _
                " and END_DATE >= '" & Format(v_datDate, "yyyy-mm-dd") & "' " & _
                "order by wg_num, eff_date desc"
                
'  " and '" & Format(v_datDate, "yyyy-mm-dd") & "' between EFF_DATE and END_DATE "
' JIMB Why did I have this above


    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWGBaseSchedules = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetWGBaseSchedules_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWGBaseSchedules", , , , DescArgsFrom(l_strSQL), True

    
End Function

Public Function GetEmployeePersonalSchedule( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset
        
'Employee Personal Schedule used to be called Non Demand Schedule
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strStartDate As String
    Dim l_strEndDate As String
    Dim l_strEmployeeSelection As String

    On Error GoTo GetEmployeePersonalSchedule_EH

    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("emp_num", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    l_strStartDate = "'" & Format(v_datStartDate, "yyyy-mm-dd") & "'"
    l_strEndDate = "'" & Format(v_datEndDate, "yyyy-mm-dd") & "'"
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT EMP_NUM, START_DATE, END_DATE, " & _
                 "  ROT_CD, ROT_WEEK , SCHED_CD " & _
                "From METHDBA.TTR_EMP_SCHED " & _
                "WHERE " & l_strEmployeeSelection & _
                "  AND START_DATE <= " & l_strEndDate & _
                " AND END_DATE >= " & l_strStartDate & " " & _
                "ORDER BY EMP_NUM, START_DATE, END_DATE "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, v_adLockType

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeePersonalSchedule = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetEmployeePersonalSchedule_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeePersonalSchedule", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function GetEmployeeLineAssignments( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset
        
'Employee is assigned to a Demand Line and therefore get the schedule
' and attributes of the line
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strStartDate As String
    Dim l_strEndDate As String
    Dim l_strEmployeeSelection As String

    On Error GoTo GetEmployeeLineAssignments_EH

    l_strStartDate = "'" & Format(v_datStartDate, "yyyy-mm-dd") & "'"
    l_strEndDate = "'" & Format(v_datEndDate, "yyyy-mm-dd") & "'"
        
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

    l_strSQL = "SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE, EMP_NUM, " & _
                "RELIEF_COMP, char(LAST_MOD_TIMESTAMP) as LAST_MOD_TIMESTAMP " & _
                "From METHDBA.TTR_LINE_ASGNM " & _
                " WHERE " & l_strEmployeeSelection & _
                "  AND START_DATE <= " & l_strEndDate & " " & _
                "  AND END_DATE >= " & l_strStartDate
                
                
    l_strSQL = l_strSQL & "  order by EMP_NUM, START_DATE DESC"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, v_adLockType

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeLineAssignments = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetEmployeeLineAssignments_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeLineAssignments", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetDemandLineSchedules( _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_aryLineNums As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset
        
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

    On Error GoTo GetDemandLineSchedules_EH
    
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

    l_strSQL = "SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE, " & _
                " ROT_CD, ROT_WEEK, SCHED_CD," & _
                " JOB_NUM, REQ_LINE_IND," & _
                " CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
                " FROM METHDBA.TTR_WG_LINE_SCHED" & _
                " WHERE " & l_strWorkGroupSelection & _
                "  AND " & l_strLineNumSelection & _
                "  AND START_DATE <= " & l_strEndDate & " " & _
                "  AND END_DATE >= " & l_strStartDate
    l_strSQL = l_strSQL & "  order by WG_NUM, LINE_NUM , START_DATE"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, v_adLockType

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDemandLineSchedules = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetDemandLineSchedules_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineSchedules", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function GetWGDemandLineBaseSchedules( _
        ByVal v_colWorkGroups As Collection, _
        ByVal v_datDate As Date _
        ) As ADODB.Recordset
    
    Dim l_intIndex As Integer
    Dim l_strSQLinClause As String
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetWGDemandLineBaseSchedules_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQLinClause = "(-1"
    For l_intIndex = 1 To v_colWorkGroups.Count
        l_strSQLinClause = l_strSQLinClause & "," & str(v_colWorkGroups.Item(l_intIndex))
    Next
    l_strSQLinClause = l_strSQLinClause & ")"
    
    l_strSQL = "select WG_NUM, LINE_NUM, START_DATE, END_DATE, ROT_CD, ROT_WEEK, SCHED_CD  " & _
                " , JOB_NUM, REQ_LINE_IND " & _
                " , CREATED_EMP_NUM, LAST_MOD_EMP_NUM " & _
                " , Char(CREATED_TIMESTAMP) as CREATED_TIMESTAMP " & _
                " , Char(LAST_MOD_TIMESTAMP) as LAST_MOD_TIMESTAMP " & _
                " , 'L' as REC_STATUS, END_DATE " & _
                "from methdba.ttr_wg_line_sched " & _
                "where wg_num in " & l_strSQLinClause & " " & _
                " and END_DATE >= '" & Format(v_datDate, "yyyy-mm-dd") & "'  " & _
                "order by wg_num, line_num, start_date desc"
                
''    " and '" & Format(v_datDate, "yyyy-mm-dd") & "' between EFF_DATE and END_DATE "
'Why did JIMB have the above

                                   
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWGDemandLineBaseSchedules = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetWGDemandLineBaseSchedules_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWGDemandLineBaseSchedules", , , , DescArgsFrom(l_strSQL), True

    
End Function


Public Function GetDemandLineTrades( _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_aryLineNums As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset
        
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
        l_strWorkGroupSelection = BuildInClause("A.WG_NUM", v_aryWorkGroups, False) & " "
    Else
        l_strWorkGroupSelection = "A.WG_NUM = " & str(CLng(v_aryWorkGroups)) & " "
    End If
    If IsArray(v_aryLineNums) Then
        l_strLineNumSelection = BuildInClause("A.LINE_NUM", v_aryLineNums, False) & " "
    Else
        l_strLineNumSelection = "A.LINE_NUM = " & str(CLng(v_aryLineNums)) & " "
    End If
    l_strStartDate = "'" & Format(v_datStartDate, "yyyy-mm-dd") & "'"
    l_strEndDate = "'" & Format(v_datEndDate, "yyyy-mm-dd") & "'"
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT A.WG_NUM, A.LINE_NUM, A.CHARGE_DATE, A.SHIFT_NUM," & _
                " A.SCHED_CD, A.TRADE_NUM, B.ROLE_TYPE_CD, B.WEEK_TRADE_IND " & _
                "From METHDBA.TTR_LINE_TRADE A, METHDBA.TTR_TRADE_DETAIL B " & _
                "WHERE " & l_strWorkGroupSelection & _
                "  AND " & l_strLineNumSelection & _
                "  AND A.CHARGE_DATE >= " & l_strStartDate & _
                "  AND A.CHARGE_DATE <= " & l_strEndDate & _
                "  AND B.TRADE_NUM = A.TRADE_NUM " & _
                "  AND B.CHARGE_DATE = A.CHARGE_DATE " & _
                "  AND B.WG_NUM = A.WG_NUM " & _
                "  AND B.LINE_NUM = A.LINE_NUM " & _
                "  AND B.SHIFT_NUM = A.SHIFT_NUM" & _
                "  order by A.WG_NUM, A.LINE_NUM , A.CHARGE_DATE, A.SHIFT_NUM DESC"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, v_adLockType

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDemandLineTrades = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetDemandLineTrades_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineTrades", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function GetEmployeeTrades( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo GetEmployeeTrades_EH
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

'Need to add     EXCEPTION_CD  for trade exception code

    l_strSQL = "SELECT A.EMP_NUM, A.CHARGE_DATE, A.SHIFT_NUM, " & _
       " A.SCHED_CD, A.TRADE_NUM, B.ROLE_TYPE_CD, B.WEEK_TRADE_IND, C.EMP_NAME " & _
       " FROM METHDBA.TTR_EMP_TRADE A, METHDBA.TTR_TRADE_DETAIL B, METHDBA.TTR_EMPLOYEE C" & _
       " WHERE " & l_strEmployeeSelection & _
       " AND (A.CHARGE_DATE" & _
       " BETWEEN '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" & _
       " AND '" & Format(v_datEndDate, "yyyy-mm-dd") & "')" & _
       " AND B.TRADE_NUM = A.TRADE_NUM " & _
       " AND B.EMP_NUM = A.EMP_NUM " & _
       " AND B.CHARGE_DATE = A.CHARGE_DATE" & _
       " AND B.SHIFT_NUM = A.SHIFT_NUM" & _
       " AND C.EMP_NUM = A.EMP_NUM" & _
       " ORDER BY  A.EMP_NUM, A.CHARGE_DATE, A.SHIFT_NUM DESC"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, v_adLockType

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeTrades = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function

GetEmployeeTrades_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeTrades", , , , DescArgsFrom(l_strSQL), True

End Function



Public Function GetEmployeeScheduledOvertime( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date _
        ) As ADODB.Recordset
        
           
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_intIndex As Integer
    Dim l_strEmployeeSelection As String
    

    On Error GoTo GetEmployeeScheduledOvertime_EH
    
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

    l_strSQL = "select a.DATE, a.wg_num, a.job_num, a.SCHED_CD, " _
        & "a.SCHED_TYPE, a.START_TIME, a.END_TIME, a.EXCEPTION_CD,  " _
        & "a.OVER_EXC_CD, a.relief_comp, a.ot_opp_num, a.case_num, " _
        & "a.loc_cd, a.comp_cd, a.shift_comp, a.fmla_ind, a.trav_pay_hrs,  " _
        & "a.charge_date, a.subsys_owner, a.pending_status, b.exception_desc, a.EMP_NUM  " _
        & ", a.LAST_MOD_EMP_NUM " _
        & ",char(a.TIME_STAMP) as LAST_MOD_TIMESTAMP " _
        & " from methdba.ttr_base_hrs_exc A, " _
        & " methdba.ttr_exception_cds B " _
        & " where " & l_strEmployeeSelection _
        & " AND (a.charge_date " _
        & " between '" & Format(v_datStartDate, "yyyy-mm-dd") & "'" _
        & " AND '" & Format(v_datEndDate, "yyyy-mm-dd") & "')" _
        & " AND SCHED_TYPE = 120 " _
        & " AND EXCEPTION_TYPE <> 'S' " _
        & " and ((a.OVER_EXC_CD <> '' " _
        & " and a.OVER_EXC_CD = b.exception_cd) " _
        & " or (a.OVER_EXC_CD = '' " _
        & " and a.EXCEPTION_CD = b.exception_cd)) "
'        & " AND PENDING_STATUS = 'T' "
'JIMB 12/23/2002
' Removed PENDING_STATUS check from where clause
' This is to allow Canvasses and Non Canvassed OT to show in View Schedule
' This has no effect on Time Collection, because the client is looking
' for a SYSTEM_EXC_CD = "SSO".  SSO is only for OT on base schedule
' The long term goal is for overtime to be scheduled and the time reported
' after it is worked.
' TRIS is not designed for this yet, but the data is retrieved to help
' the process
' The client will need to look at the PENDING_STATUS flag, that will need
' added to the schedule it receives.
        
    l_strSQL = l_strSQL & "  ORDER BY  A.EMP_NUM, DATE, SCHED_TYPE"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeScheduledOvertime = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetEmployeeScheduledOvertime_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeScheduledOvertime", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function GetSchedulesInRotation( _
        ByVal v_intRotationCode As Integer _
        ) As ADODB.Recordset

'This function is similar to GetRotationSchedules except
'   1. Needed for only one Rotation Code
'   2. an extra coloumn CURRENT_SCHED is returned
'   3. the recordset is updateable and not read only
'The seperation is needed to conserve database resources
'GetRotationSchedules is called when getting the schedule for the employee
'GetSchedulesInRotation is called to display the schedule codes
'  and the current schedule on a date


    Dim l_intIndex As Integer
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "select rot_cd, rot_week, sched_cd, 'N' as CURRENT_SCHED " & _
                "from methdba.ttr_rotation_sched " & _
                "where rot_cd = " & str(v_intRotationCode) & " " & _
                "order by rot_cd, rot_week"
                                   
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetSchedulesInRotation = l_objRS
    
    Set l_objRS = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetSchedulesInRotation", , , , DescArgsFrom(l_strSQL), True

    
End Function

Public Function GetWorkGroupLineAssignments( _
        ByVal v_aryWorkGroups As Variant, _
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

    On Error GoTo GetWorkGroupLineAssignments_EH

    l_strStartDate = "'" & Format(v_datStartDate, "yyyy-mm-dd") & "'"
    l_strEndDate = "'" & Format(v_datEndDate, "yyyy-mm-dd") & "'"
        
    If IsArray(v_aryWorkGroups) Then
        l_strWorkGroupSelection = BuildInClause("WG_NUM", v_aryWorkGroups, False) & " "
    Else
        l_strWorkGroupSelection = "WG_NUM = " & str(CLng(v_aryWorkGroups)) & " "
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE, EMP_NUM, " & _
                "RELIEF_COMP " & _
                "From METHDBA.TTR_LINE_ASGNM " & _
                " WHERE " & l_strWorkGroupSelection & _
                "  AND START_DATE <= " & l_strEndDate & " " & _
                "  AND END_DATE >= " & l_strStartDate
                
    l_strSQL = l_strSQL & "  order by EMP_NUM, START_DATE DESC"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWorkGroupLineAssignments = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetWorkGroupLineAssignments_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupLineAssignments", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function GetWorkGroupLineBaseSched( _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
        

    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strStartDate As String
    Dim l_strEndDate As String
    Dim l_strWorkGroupSelection As String

    On Error GoTo GetWorkGroupLineBaseSched_EH
    
    If IsArray(v_aryWorkGroups) Then
        l_strWorkGroupSelection = BuildInClause("WG_NUM", v_aryWorkGroups, False) & " "
    Else
        l_strWorkGroupSelection = "WG_NUM = " & str(CLng(v_aryWorkGroups)) & " "
    End If
    l_strStartDate = "'" & Format(v_datStartDate, "yyyy-mm-dd") & "'"
    l_strEndDate = "'" & Format(v_datEndDate, "yyyy-mm-dd") & "'"
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE, " & _
                " ROT_CD, ROT_WEEK, SCHED_CD " & _
                " , JOB_NUM, REQ_LINE_IND " & _
                "From METHDBA.TTR_WG_LINE_SCHED " & _
                "WHERE " & l_strWorkGroupSelection & _
                "  AND START_DATE <= " & l_strEndDate & " " & _
                "  AND END_DATE >= " & l_strStartDate
    l_strSQL = l_strSQL & "  order by WG_NUM, LINE_NUM , START_DATE"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWorkGroupLineBaseSched = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetWorkGroupLineBaseSched_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupLineBaseSched", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetWorkGroupLineTrades( _
        ByVal v_aryWorkGroups As Variant, _
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

    On Error GoTo GetWorkGroupLineTrades_EH
    
    If IsArray(v_aryWorkGroups) Then
        l_strWorkGroupSelection = BuildInClause("A.WG_NUM", v_aryWorkGroups, False) & " "
    Else
        l_strWorkGroupSelection = "A.WG_NUM = " & str(CLng(v_aryWorkGroups)) & " "
    End If
    l_strStartDate = "'" & Format(v_datStartDate, "yyyy-mm-dd") & "'"
    l_strEndDate = "'" & Format(v_datEndDate, "yyyy-mm-dd") & "'"
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT A.WG_NUM, A.LINE_NUM, A.CHARGE_DATE, A.SHIFT_NUM," & _
                " A.SCHED_CD, A.TRADE_NUM, B.ROLE_TYPE_CD, B.WEEK_TRADE_IND " & _
                "From METHDBA.TTR_LINE_TRADE A, METHDBA.TTR_TRADE_DETAIL B " & _
                "WHERE " & l_strWorkGroupSelection & _
                "  AND A.CHARGE_DATE >= " & l_strStartDate & _
                "  AND A.CHARGE_DATE <= " & l_strEndDate & _
                "  AND B.TRADE_NUM = A.TRADE_NUM " & _
                "  AND B.CHARGE_DATE = A.CHARGE_DATE " & _
                "  AND B.WG_NUM = A.WG_NUM " & _
                "  AND B.LINE_NUM = A.LINE_NUM " & _
                "  AND B.SHIFT_NUM = A.SHIFT_NUM" & _
                "  order by A.WG_NUM, A.LINE_NUM , A.CHARGE_DATE, A.SHIFT_NUM DESC"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWorkGroupLineTrades = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetWorkGroupLineTrades_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupLineTrades", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function GetWorkGroupAssignments( _
        ByVal v_aryWorkGroups As Variant, _
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

    On Error GoTo GetWorkGroupAssignments_EH

    l_strStartDate = "'" & Format(v_datStartDate, "yyyy-mm-dd") & "'"
    l_strEndDate = "'" & Format(v_datEndDate, "yyyy-mm-dd") & "'"
        
    If IsArray(v_aryWorkGroups) Then
        l_strWorkGroupSelection = BuildInClause("WG_NUM", v_aryWorkGroups, False) & " "
    Else
        l_strWorkGroupSelection = "WG_NUM = " & str(CLng(v_aryWorkGroups)) & " "
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT WG_NUM, EMP_NUM, PRIMARY_IND," & _
                "EFF_START_DATE, EFF_END_DATE " & _
                "From METHDBA.TTR_EMP_WG " & _
                " WHERE " & l_strWorkGroupSelection & _
                "  AND EFF_START_DATE <= " & l_strEndDate & " " & _
                "  AND EFF_END_DATE >= " & l_strStartDate
                
    l_strSQL = l_strSQL & "  order by EMP_NUM, EFF_START_DATE "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWorkGroupAssignments = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetWorkGroupAssignments_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupAssignments", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function GetPrimaryWorkGroupAssignments( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset
        
'RSG CASCADE PENDING CHANGES
        
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strStartDate As String
    Dim l_strEndDate As String
    Dim l_strEmployeeSelection As String


On Error GoTo eh

    l_strStartDate = "'" & Format(v_datStartDate, "yyyy-mm-dd") & "'"
    l_strEndDate = "'" & Format(v_datEndDate, "yyyy-mm-dd") & "'"
        
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    l_objCN.Open
    
    l_strSQL = "SELECT WG_NUM, EMP_NUM, PRIMARY_IND," & _
                " EFF_START_DATE, EFF_END_DATE, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
                " FROM METHDBA.TTR_EMP_WG" & _
                " WHERE " & l_strEmployeeSelection & _
                " AND PRIMARY_IND = 'Y' " & _
                " AND EFF_START_DATE <= " & l_strEndDate & _
                " AND EFF_END_DATE >= " & l_strStartDate & _
                " ORDER BY EMP_NUM, EFF_START_DATE "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, v_adLockType

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetPrimaryWorkGroupAssignments = l_objRS
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPrimaryWorkGroupAssignments", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function GetWorkGroupBaseSchedules( _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset
        
'RSG CASCADE PENDING CHANGES
        
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strStartDate As String
    Dim l_strEndDate As String
    Dim l_strWorkGroupSelection As String


On Error GoTo eh

    l_strStartDate = "'" & Format(v_datStartDate, "yyyy-mm-dd") & "'"
    l_strEndDate = "'" & Format(v_datEndDate, "yyyy-mm-dd") & "'"
        
    If IsArray(v_aryWorkGroups) Then
        l_strWorkGroupSelection = BuildInClause("WG_NUM", v_aryWorkGroups, False) & " "
    Else
        l_strWorkGroupSelection = "WG_NUM = " & str(CLng(v_aryWorkGroups)) & " "
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    l_objCN.Open
    
    l_strSQL = "SELECT WG_NUM, EFF_DATE, END_DATE, ROT_CD, ROT_WEEK, SCHED_CD," & _
                " CHAR(CREATED_TIMESTAMP) as CREATED_TIMESTAMP" & _
                " FROM METHDBA.TTR_WG_BASE_SCHED" & _
                " WHERE " & l_strWorkGroupSelection & _
                " AND EFF_DATE <= " & l_strEndDate & _
                " AND END_DATE >= " & l_strStartDate & _
                " ORDER BY WG_NUM, EFF_DATE DESC"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, v_adLockType

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetWorkGroupBaseSchedules = l_objRS
    Set l_objRS = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupBaseSchedules", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetScheduleTimeStamps( _
        ByVal v_aryEmployees As Variant _
        ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

On Error GoTo eh

    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "select EMP_NUM, char(LAST_MOD_SCHED_TM) as LAST_MOD_SCHED_TM " _
        & "From methdba.TTR_EMP_MOD " _
        & "where " & l_strEmployeeSelection _
        & "order by EMP_NUM "
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetScheduleTimeStamps = l_objRS
    
    Set l_objRS = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduleTimeStamps", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetJobPosition(ByVal vlngWGNum As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset

On Error GoTo eh

    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    l_objCN.Open
    
    l_strSQL = "SELECT WG_NUM, CHARGE_DATE, LINE_NUM, JOB_POSITION " & _
                " FROM METHDBA.TTR_DEMAND_JOB_POS " & _
                "WHERE WG_NUM = " & vlngWGNum & _
                " AND CHARGE_DATE BETWEEN '" & Format(vdtStartDate, "mm/dd/yyyy") & "'" & _
                " AND '" & Format(vdtEndDate, "mm/dd/yyyy") & "'"
                
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetJobPosition = l_objRS
    Set l_objRS = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetJobPosition", , , , DescArgsFrom(l_strSQL), True



End Function

Public Function GetReportedOTHolAndVacs( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date _
        ) As ADODB.Recordset
        
           
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strSchedTypeforSQL As String
    Dim l_intIndex As Integer
    Dim l_strEmployeeSelection As String
    

    On Error GoTo eh
    
        
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "select EMP_NUM, CHARGE_DATE, DATE_WORKED, HOURS, "
    l_strSQL = l_strSQL & "START_TIME, END_TIME, A.EXCEPTION_CD, "
    l_strSQL = l_strSQL & "SCHED_CD, OT_OPP_NUM, TRAVEL_PAY_HRS "
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_ACTUAL_TIME A, METHDBA.TTR_EXCEPTION_CDS B "
    l_strSQL = l_strSQL & " WHERE " & l_strEmployeeSelection
    l_strSQL = l_strSQL & " AND CHARGE_DATE BETWEEN " & DateAsEnclosedString(v_datStartDate)
    l_strSQL = l_strSQL & " AND " & DateAsEnclosedString(v_datEndDate)
    l_strSQL = l_strSQL & " AND A.EXCEPTION_CD = B.EXCEPTION_CD "
    l_strSQL = l_strSQL & " AND ( GROUP_IND IN ('H','V') "
     l_strSQL = l_strSQL & "  OR OT_OPP_NUM = 90000000 ) "
    l_strSQL = l_strSQL & " ORDER BY EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetReportedOTHolAndVacs = l_objRS
    
    Set l_objRS = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetReportedOTHolAndVacs", , , , DescArgsFrom(l_strSQL), True

End Function



Public Function GetFlexSchedInfo(ByVal v_strStartTime As String) As ADODB.Recordset

'code assumes at least One Schedule to retrieve


    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetFlexSchedInfo_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    
    l_strSQL = "select * from methdba.ttr_avail_sched " & _
                "WHERE SCHED_CD BETWEEN 20000 AND 20099 " & _
                " and MON_START_TIME = '" & v_strStartTime & "' " & _
                "order by sched_cd"
                                   
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetFlexSchedInfo = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetFlexSchedInfo_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetFlexSchedInfo", , , , DescArgsFrom(l_strSQL), True

    
End Function

