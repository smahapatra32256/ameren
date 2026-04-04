VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdSchedCodeMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrCMODULENAME As String = "cbSchedCodeMaint"
Public Function SchedCodeInUse(ByVal v_lSchedCode As Long) As Boolean

On Error GoTo SchedCodeInUse_EH

    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    'Get Connection
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    
    'Open Connection
    l_objCN.Open
    
    'Format SQL Statement
    l_strSQL = "SELECT 1 AS EXISTS_SW " & _
                   " From SYSIBM.SYSDUMMY1 " & _
                   " WHERE Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_WG_LINE_SCHED " & _
                   " WHERE SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_BASE_HRS_EXC " & _
                   " WHERE SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_BASE_HRS_EXC " & _
                   " WHERE PREV_SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_BSE_HRS_EXC_HS " & _
                   " WHERE SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_BSE_HRS_EXC_HS " & _
                   " WHERE PREV_SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_BASE_WORK_HRS " & _
                   " WHERE SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_RELIEF " & _
                   " WHERE SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_RELIEF " & _
                   " WHERE REL_SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_WG_BASE_SCHED " & _
                   " WHERE SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_SCHED_CHG_PMTS " & _
                   " WHERE BASELINE_SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_ROTATION_SCHED " & _
                   " WHERE SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_EMP_SCHED" & _
                   " WHERE SCHED_CD = " & v_lSchedCode & ")"
    l_strSQL = l_strSQL & " OR Exists " & _
                    " (SELECT * " & _
                   " From METHDBA.TTR_TRADE_DETAIL" & _
                   " WHERE SCHED_CD = " & v_lSchedCode & ")"
                                        
    SchedCodeInUse = False
    Set l_objRS = New ADODB.Recordset
    
    With l_objRS
        .CursorLocation = adUseClient
        .Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
        If IsRecordsetPopulated(l_objRS) Then
            SchedCodeInUse = True
        End If
    End With
    
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
    
SchedCodeInUse_EH:

    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
   
    TrsRaiseError mstrCMODULENAME, "SchedCodeInUse", , , , , True
    
End Function

Public Function SaveSchedCode _
    (ByRef r_adrsSchedCodeEntries As ADODB.Recordset _
    , ByRef r_adrsBRV As ADODB.Recordset _
    , ByVal v_lTRISUser As Long) As Boolean
                                
On Error GoTo SaveSchedCode_EH

    Dim l_bReturn As Boolean
    Dim l_sFilterString As String
    
    SaveSchedCode = True
    
    If IsRecordsetPopulated(r_adrsSchedCodeEntries) Then
        With r_adrsSchedCodeEntries
            .MoveFirst
            'Process Delete records
            .Filter = adFilterNone
            l_sFilterString = "RECORD_STATUS = 'D'"
            .Filter = MDAC25FilterFix(l_sFilterString)
            Do While Not .EOF
                If Not DeleteSchedCode(!KEY_SCHED_CD, !TIME_STAMP, r_adrsBRV, v_lTRISUser) Then
                    SaveSchedCode = False
                    .Filter = adFilterNone
                    GoTo SaveSchedCode_EH
                End If
                .MoveNext
            Loop
            .Filter = adFilterNone
            l_sFilterString = "RECORD_STATUS = 'I'"
            .Filter = MDAC25FilterFix(l_sFilterString)
            Do While Not .EOF
                If Not InsertSchedCode(r_adrsSchedCodeEntries, r_adrsBRV, v_lTRISUser) Then
                    SaveSchedCode = False
                    .Filter = adFilterNone
                    GoTo SaveSchedCode_EH
                End If
                .MoveNext
            Loop
            .Filter = adFilterNone
        End With
    End If
    
    GetObjectContext.SetComplete
    
    Exit Function
    
SaveSchedCode_EH:
    
    GetObjectContext.SetAbort
    
    TrsRaiseError mstrCMODULENAME, "SaveSchedCode", , , , , True

End Function
Private Function UpdateSchedCode( _
    ByRef r_adrsSchedCodeEntries As ADODB.Recordset _
    , ByRef r_adrsBRV As ADODB.Recordset _
    , ByVal v_lTRISUser As Long) As Boolean
    
On Error GoTo UpdateSchedCode_EH

    Dim l_sSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_lDayCounter As Long, l_lNumberOfDays As Long
    
    'Get Connection
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    'Open Connection
    l_objCN.Open
    With r_adrsSchedCodeEntries
        l_sSQL = " UPDATE methdba.ttr_avail_sched " & _
                    " SET sched_cd = " & !SCHED_CD & ", " & _
                    " sched_desc = '" & Trim(Replace(!SCHED_DESC, "'", "''")) & "', "
        l_sSQL = l_sSQL & _
                    " sun_act_day = '" & Trim(!SUN_ACT_DAY) & "'," & _
                    " sun_shift_cd = '" & Trim(!SUN_SHIFT_CD) & "'," & _
                    " sun_start_time = '" & Trim(!SUN_START_TIME) & "', " & _
                    " sun_end_time = '" & Trim(!SUN_END_TIME) & "', " & _
                    " sun_lunch_start = '" & Trim(!SUN_LUNCH_START) & "', " & _
                    " sun_lunch_end = '" & Trim(!SUN_LUNCH_END) & "', " & _
                    " sun_str_hrs = " & !SUN_STR_HRS & ", " & _
                    " sun_ot_start_time = '" & Trim(!SUN_OT_START_TIME) & "', " & _
                    " sun_ot_end_time = '" & Trim(!SUN_OT_END_TIME) & "', " & _
                    " sun_ot_hours = " & !SUN_OT_HOURS & ", " & _
                    " sun_comp_shift_cd = '" & Trim(!SUN_COMP_SHIFT_CD) & "', "
        l_sSQL = l_sSQL & _
                    " mon_act_day = '" & Trim(!MON_ACT_DAY) & "'," & _
                    " mon_shift_cd = '" & Trim(!MON_SHIFT_CD) & "'," & _
                    " mon_start_time = '" & Trim(!MON_START_TIME) & "', " & _
                    " mon_end_time = '" & Trim(!MON_END_TIME) & "', " & _
                    " mon_lunch_start = '" & Trim(!MON_LUNCH_START) & "', " & _
                    " mon_lunch_end = '" & Trim(!MON_LUNCH_END) & "', " & _
                    " mon_str_hrs = " & !MON_STR_HRS & ", " & _
                    " mon_ot_start_time = '" & Trim(!MON_OT_START_TIME) & "', " & _
                    " mon_ot_end_time = '" & Trim(!mon_ot_end_time) & "', " & _
                    " mon_ot_hours = " & !MON_OT_HOURS & ", " & _
                    " mon_comp_shift_cd = '" & Trim(!MON_COMP_SHIFT_CD) & "', "
            l_sSQL = l_sSQL & _
                    " tue_act_day = '" & Trim(!TUE_ACT_DAY) & "'," & _
                    " tue_shift_cd = '" & Trim(!TUE_SHIFT_CD) & "'," & _
                    " tue_start_time = '" & Trim(!TUE_START_TIME) & "', " & _
                    " tue_end_time = '" & Trim(!TUE_END_TIME) & "', " & _
                    " tue_lunch_start = '" & Trim(!TUE_LUNCH_START) & "', " & _
                    " tue_lunch_end = '" & Trim(!TUE_LUNCH_END) & "', " & _
                    " tue_str_hrs = " & !TUE_STR_HRS & ", " & _
                    " tue_ot_start_time = '" & Trim(!TUE_OT_START_TIME) & "', " & _
                    " tue_ot_end_time = '" & Trim(!TUE_OT_END_TIME) & "', " & _
                    " tue_ot_hours = " & !TUE_OT_HOURS & ", " & _
                    " tue_comp_shift_cd = '" & Trim(!TUE_COMP_SHIFT_CD) & "', "
            l_sSQL = l_sSQL & _
                    " wed_act_day = '" & Trim(!WED_ACT_DAY) & "'," & _
                    " wed_shift_cd = '" & Trim(!WED_SHIFT_CD) & "'," & _
                    " wed_start_time = '" & Trim(!WED_START_TIME) & "', " & _
                    " wed_end_time = '" & Trim(!WED_END_TIME) & "', " & _
                    " wed_lunch_start = '" & Trim(!WED_LUNCH_START) & "', " & _
                    " wed_lunch_end = '" & Trim(!WED_LUNCH_END) & "', " & _
                    " wed_str_hrs = " & !WED_STR_HRS & ", " & _
                    " wed_ot_start_time = '" & Trim(!WED_OT_START_TIME) & "', " & _
                    " wed_ot_end_time = '" & Trim(!WED_OT_END_TIME) & "', " & _
                    " wed_ot_hours = " & !WED_OT_HOURS & ", " & _
                    " wed_comp_shift_cd = '" & Trim(!WED_COMP_SHIFT_CD) & "', "
            l_sSQL = l_sSQL & _
                    " thu_act_day = '" & Trim(!THU_ACT_DAY) & "'," & _
                    " thu_shift_cd = '" & Trim(!THU_SHIFT_CD) & "'," & _
                    " thu_start_time = '" & Trim(!THU_START_TIME) & "', " & _
                    " thu_end_time = '" & Trim(!THU_END_TIME) & "', " & _
                    " thu_lunch_start = '" & Trim(!THU_LUNCH_START) & "', " & _
                    " thu_lunch_end = '" & Trim(!THU_LUNCH_END) & "', " & _
                    " thu_str_hrs = " & !THU_STR_HRS & ", " & _
                    " thu_ot_start_time = '" & Trim(!THU_OT_START_TIME) & "', " & _
                    " thu_ot_end_time = '" & Trim(!THU_OT_END_TIME) & "', " & _
                    " thu_ot_hours = " & !THU_OT_HOURS & ", " & _
                    " thu_comp_shift_cd = '" & Trim(!THU_COMP_SHIFT_CD) & "', "
            l_sSQL = l_sSQL & _
                    " fri_act_day = '" & Trim(!FRI_ACT_DAY) & "'," & _
                    " fri_shift_cd = '" & Trim(!FRI_SHIFT_CD) & "'," & _
                    " fri_start_time = '" & Trim(!FRI_START_TIME) & "', " & _
                    " fri_end_time = '" & Trim(!FRI_END_TIME) & "', " & _
                    " fri_lunch_start = '" & Trim(!FRI_LUNCH_START) & "', " & _
                    " fri_lunch_end = '" & Trim(!FRI_LUNCH_END) & "', " & _
                    " fri_str_hrs = " & !FRI_STR_HRS & ", " & _
                    " fri_ot_start_time = '" & Trim(!FRI_OT_START_TIME) & "', " & _
                    " fri_ot_end_time = '" & Trim(!FRI_OT_END_TIME) & "', " & _
                    " fri_ot_hours = " & !FRI_OT_HOURS & ", " & _
                    " fri_comp_shift_cd = '" & Trim(!FRI_COMP_SHIFT_CD) & "', "
            l_sSQL = l_sSQL & _
                    " sat_act_day = '" & Trim(!SAT_ACT_DAY) & "'," & _
                    " sat_shift_cd = '" & Trim(!SAT_SHIFT_CD) & "'," & _
                    " sat_start_time = '" & Trim(!SAT_START_TIME) & "', " & _
                    " sat_end_time = '" & Trim(!SAT_END_TIME) & "', " & _
                    " sat_lunch_start = '" & Trim(!SAT_LUNCH_START) & "', " & _
                    " sat_lunch_end = '" & Trim(!SAT_LUNCH_END) & "', " & _
                    " sat_str_hrs = " & !SAT_STR_HRS & ", " & _
                    " sat_ot_start_time = '" & Trim(!SAT_OT_START_TIME) & "', " & _
                    " sat_ot_end_time = '" & Trim(!SAT_OT_END_TIME) & "', " & _
                    " sat_ot_hours = " & !SAT_OT_HOURS & ", " & _
                    " sat_comp_shift_cd = '" & Trim(!SAT_COMP_SHIFT_CD) & "', "
        
        l_sSQL = l_sSQL & _
                    " last_mod_emp_num = " & !LAST_MOD_EMP_NUM
        l_sSQL = l_sSQL & _
                    " WHERE sched_cd = " & !SCHED_CD & _
                    " AND TIME_STAMP = '" & !TIME_STAMP & "'"
    End With
    
    l_objCN.Execute (l_sSQL)
        
    Set l_objCN = Nothing
    
    UpdateSchedCode = True
    
    Exit Function
    
UpdateSchedCode_EH:
    UpdateSchedCode = False
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    TrsRaiseError mstrCMODULENAME, "UpdateSchedCode", , , , DescArgsFrom(l_sSQL), True
    
End Function
Private Function UpdateSchedCodeDesc( _
    ByRef r_adrsSchedCodeEntries As ADODB.Recordset _
    , ByRef v_adrsBRV As ADODB.Recordset _
    , ByVal v_lTRISUser As Long) As Boolean
    
On Error GoTo UpdateSchedCodeDesc_EH

    Dim l_sSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_lDayCounter As Long, l_lNumberOfDays As Long
    
    'Get Connection
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    'Open Connection
    l_objCN.Open
 
    l_sSQL = " UPDATE methdba.ttr_avail_sched " & _
                " SET sched_desc = '" & Trim(r_adrsSchedCodeEntries!SCHED_DESC) & "', " & _
                " last_mod_emp_num = " & r_adrsSchedCodeEntries!LAST_MOD_EMP_NUM
    l_sSQL = l_sSQL & _
                " WHERE sched_cd = " & r_adrsSchedCodeEntries!SCHED_CD & _
                " AND TIME_STAMP = '" & Trim(r_adrsSchedCodeEntries!TIME_STAMP) & "'"
           
    l_objCN.Execute (l_sSQL)
    
    Set l_objCN = Nothing
    
    UpdateSchedCodeDesc = True
    
    Exit Function
    
UpdateSchedCodeDesc_EH:
    UpdateSchedCodeDesc = False
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    TrsRaiseError mstrCMODULENAME, "UpdateSchedCodeDesc", , , , DescArgsFrom(l_sSQL), True
    
End Function

Private Function InsertSchedCode( _
    ByRef r_adrsSchedCodeEntries As ADODB.Recordset _
    , ByRef v_adrsBRV As ADODB.Recordset _
    , ByVal v_lTRISUser As Long) As Boolean
    
On Error GoTo InsertSchedCode_EH

    Dim l_sSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_lNewSchedCode As Long
    
    
    
    'Get Connection
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    'Open Connection
    l_objCN.Open
    With r_adrsSchedCodeEntries
        If !SCHED_CD > 0 Then
            l_lNewSchedCode = !SCHED_CD
        Else
            l_lNewSchedCode = GetNewSchedCode
        End If
        l_sSQL = "INSERT INTO methdba.ttr_avail_sched " & _
                    " (sched_cd, sched_desc, " & _
                    " sun_act_day,sun_shift_cd, sun_start_time,sun_end_time, " & _
                    " sun_lunch_start, sun_lunch_end, sun_str_hrs, " & _
                    " sun_ot_start_time, sun_ot_end_time, sun_ot_hours, " & _
                    " sun_comp_shift_cd, "
        l_sSQL = l_sSQL & " mon_act_day,mon_shift_cd, mon_start_time,mon_end_time, " & _
                    " mon_lunch_start, mon_lunch_end, mon_str_hrs, " & _
                    " mon_ot_start_time, mon_ot_end_time, mon_ot_hours, " & _
                    " mon_comp_shift_cd, "
        l_sSQL = l_sSQL & " tue_act_day,tue_shift_cd, tue_start_time,tue_end_time, " & _
                    " tue_lunch_start, tue_lunch_end, tue_str_hrs, " & _
                    " tue_ot_start_time, tue_ot_end_time, tue_ot_hours, " & _
                    " tue_comp_shift_cd, "
        l_sSQL = l_sSQL & " wed_act_day,wed_shift_cd, wed_start_time,wed_end_time, " & _
                    " wed_lunch_start, wed_lunch_end, wed_str_hrs, " & _
                    " wed_ot_start_time, wed_ot_end_time, wed_ot_hours, " & _
                    " wed_comp_shift_cd, "
        l_sSQL = l_sSQL & " thu_act_day,thu_shift_cd, thu_start_time,thu_end_time, " & _
                    " thu_lunch_start, thu_lunch_end, thu_str_hrs, " & _
                    " thu_ot_start_time, thu_ot_end_time, thu_ot_hours, " & _
                    " thu_comp_shift_cd, "
        l_sSQL = l_sSQL & " fri_act_day,fri_shift_cd, fri_start_time,fri_end_time, " & _
                    " fri_lunch_start, fri_lunch_end, fri_str_hrs, " & _
                    " fri_ot_start_time, fri_ot_end_time, fri_ot_hours, " & _
                    " fri_comp_shift_cd, "
        l_sSQL = l_sSQL & " sat_act_day,sat_shift_cd, sat_start_time,sat_end_time, " & _
                    " sat_lunch_start, sat_lunch_end, sat_str_hrs, " & _
                    " sat_ot_start_time, sat_ot_end_time, sat_ot_hours, " & _
                    " sat_comp_shift_cd, "
        l_sSQL = l_sSQL & " last_mod_emp_num)  "
        l_sSQL = l_sSQL & " VALUES(" & l_lNewSchedCode & " ,'" & Trim(Replace(!SCHED_DESC, "'", "''")) & "', " & _
                    "'" & !SUN_ACT_DAY & "', '" & !SUN_SHIFT_CD & "', '" & !SUN_START_TIME & "', '" & !SUN_END_TIME & "', " & _
                    "'" & !SUN_LUNCH_START & "', '" & !SUN_LUNCH_END & "', " & !SUN_STR_HRS & ", " & _
                    "'" & !SUN_OT_START_TIME & "', '" & !SUN_OT_END_TIME & "', " & !SUN_OT_HOURS & ", " & _
                    "'" & !SUN_COMP_SHIFT_CD & "', "
        l_sSQL = l_sSQL & "'" & !MON_ACT_DAY & "', '" & !MON_SHIFT_CD & "', '" & !MON_START_TIME & "', '" & !MON_END_TIME & "', " & _
                    "'" & !MON_LUNCH_START & "', '" & !MON_LUNCH_END & "', " & !MON_STR_HRS & ", " & _
                    "'" & !MON_OT_START_TIME & "', '" & !mon_ot_end_time & "', " & !MON_OT_HOURS & ", " & _
                    "'" & !MON_COMP_SHIFT_CD & "', "
        l_sSQL = l_sSQL & "'" & !TUE_ACT_DAY & "', '" & !TUE_SHIFT_CD & "', '" & !TUE_START_TIME & "', '" & !TUE_END_TIME & "', " & _
                    "'" & !TUE_LUNCH_START & "', '" & !TUE_LUNCH_END & "', " & !TUE_STR_HRS & ", " & _
                    "'" & !TUE_OT_START_TIME & "', '" & !TUE_OT_END_TIME & "', " & !TUE_OT_HOURS & ", " & _
                    "'" & !TUE_COMP_SHIFT_CD & "', "
        l_sSQL = l_sSQL & "'" & !WED_ACT_DAY & "', '" & !WED_SHIFT_CD & "', '" & !WED_START_TIME & "', '" & !WED_END_TIME & "', " & _
                    "'" & !WED_LUNCH_START & "', '" & !WED_LUNCH_END & "', " & !WED_STR_HRS & ", " & _
                    "'" & !WED_OT_START_TIME & "', '" & !WED_OT_END_TIME & "', " & !WED_OT_HOURS & ", " & _
                    "'" & !WED_COMP_SHIFT_CD & "', "
        l_sSQL = l_sSQL & "'" & !THU_ACT_DAY & "', '" & !THU_SHIFT_CD & "', '" & !THU_START_TIME & "', '" & !THU_END_TIME & "', " & _
                    "'" & !THU_LUNCH_START & "', '" & !THU_LUNCH_END & "', " & !THU_STR_HRS & ", " & _
                    "'" & !THU_OT_START_TIME & "', '" & !THU_OT_END_TIME & "', " & !THU_OT_HOURS & ", " & _
                    "'" & !THU_COMP_SHIFT_CD & "', "
        l_sSQL = l_sSQL & "'" & !FRI_ACT_DAY & "', '" & !FRI_SHIFT_CD & "', '" & !FRI_START_TIME & "', '" & !FRI_END_TIME & "', " & _
                    "'" & !FRI_LUNCH_START & "', '" & !FRI_LUNCH_END & "', " & !FRI_STR_HRS & ", " & _
                    "'" & !FRI_OT_START_TIME & "', '" & !FRI_OT_END_TIME & "', " & !FRI_OT_HOURS & ", " & _
                    "'" & !FRI_COMP_SHIFT_CD & "', "
        l_sSQL = l_sSQL & "'" & !SAT_ACT_DAY & "', '" & !SAT_SHIFT_CD & "', '" & !SAT_START_TIME & "', '" & !SAT_END_TIME & "', " & _
                    "'" & !SAT_LUNCH_START & "', '" & !SAT_LUNCH_END & "', " & !SAT_STR_HRS & ", " & _
                    "'" & !SAT_OT_START_TIME & "', '" & !SAT_OT_END_TIME & "', " & !SAT_OT_HOURS & ", " & _
                    "'" & !SAT_COMP_SHIFT_CD & "', "
        l_sSQL = l_sSQL & v_lTRISUser & ")"
        !SCHED_CD = l_lNewSchedCode
    End With
    
    l_objCN.Execute (l_sSQL)
    
    Set l_objCN = Nothing
    InsertSchedCode = True
    
    Exit Function
    
InsertSchedCode_EH:
    GetObjectContext.SetAbort
    InsertSchedCode = False
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    TrsRaiseError mstrCMODULENAME, "InsertSchedCode", , , , DescArgsFrom(l_sSQL), True
    
End Function
Private Function GetNewSchedCode() As Long
    On Error GoTo GetNewSchedCode_Err
    
    Dim l_objutilDBDataServices As utilDBDataServices.cuDBDataServices
    Dim l_lSequenceNumber As Long
    
    'GetNextSequenceNumber will create a new transaction
    Set l_objutilDBDataServices = CreateObject("utilDBDataServices.cuDBDataServices")
    
    l_lSequenceNumber = l_objutilDBDataServices.GetNextSequenceNumber("WS")
    If l_lSequenceNumber > 0 Then
        GetNewSchedCode = l_lSequenceNumber
    Else
        GoTo GetNewSchedCode_Err
    End If
 
    Set l_objutilDBDataServices = Nothing
    
    Exit Function

GetNewSchedCode_Err:

    GetObjectContext.SetAbort
    Set l_objutilDBDataServices = Nothing
    
    TrsRaiseError mstrCMODULENAME, "GetNewSchedCode ", , , , , True

End Function

Private Function InsertPhysicalSched(ByRef r_adrsSchedCodeEntries As ADODB.Recordset _
    , ByRef v_adrsBRV As ADODB.Recordset _
    , ByVal v_lTRISUser As Long) As Boolean
    
'
'On Error GoTo InsertSchedCode_EH
'
'    Dim l_sSQL As String
'    Dim l_objCN As ADODB.Connection
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'
'    'Get Connection
'    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
'    Set l_objCN = New ADODB.Connection
'    l_objCN.ConnectionString = l_objCM.GetConnectionString
'    Set l_objCM = Nothing
'    'Open Connection
'    l_objCN.Open
'
'    l_sSQL = "INSERT INTO methdba.ttr_avail_sched " & _
'                " (sched_cd, sched_desc, " & _
'                " sun_act_day,sun_shift_cd, sun_start_time,sun_end_time, " & _
'                " sun_lunch_start, sun_lunch_end, sun_str_hrs, " & _
'                " sun_ot_start_time, sun_ot_end_time, sun_ot_hours, " & _
'                " sun_comp_shift_cd, "
'    l_sSQL = l_sSQL & " mon_act_day,mon_shift_cd, mon_start_time,mon_end_time, " & _
'                " mon_lunch_start, mon_lunch_end, mon_str_hrs, " & _
'                " mon_ot_start_time, mon_ot_end_time, mon_ot_hours, " & _
'                " mon_comp_shift_cd, "
'
'    l_objCN.Execute (l_sSQL)
'
'    Set l_objCN = Nothing
'    InsertSchedCode = True
'
'    Exit Function
'
'InsertSchedCode_EH:
'    GetObjectContext.SetAbort
'    InsertSchedCode = False
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    TrsRaiseError mstrCMODULENAME, "InsertSchedCode", , , , DescArgsFrom(l_sSQL), True
'
'

End Function
Private Function DeleteSchedCode( _
    ByVal v_iSchedCode As Integer, _
    ByVal v_sTimeStamp As String, _
    ByRef r_adrsBRV As ADODB.Recordset, _
    ByVal v_lTRISUser As Long)

On Error GoTo DeleteSchedCode_EH

    Dim l_sSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    'Get Connection
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    
    'Open Connection
    l_objCN.Open
    
    l_sSQL = "DELETE FROM methdba.ttr_avail_sched " & _
                        "WHERE sched_cd = " & v_iSchedCode & _
                        " AND TIME_STAMP = '" & v_sTimeStamp & "'"
    
    l_objCN.Execute (l_sSQL)

    Set l_objCN = Nothing
    DeleteSchedCode = True

    Exit Function

DeleteSchedCode_EH:
    GetObjectContext.SetAbort
    DeleteSchedCode = False
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    TrsRaiseError mstrCMODULENAME, "DeleteSchedCode", , , , DescArgsFrom(l_sSQL), True

End Function

Private Function IsRecordsetPopulated(ByRef r_adrsRecordset As ADODB.Recordset) As Boolean

    If Not r_adrsRecordset Is Nothing Then
        If Not (r_adrsRecordset.EOF And r_adrsRecordset.BOF) Then
            IsRecordsetPopulated = True
        Else
            IsRecordsetPopulated = False
        End If
    Else
        IsRecordsetPopulated = False
    End If
    
End Function
Public Function GetScheduleDetail(ByVal v_lSchedCode As Long) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long
    
On Error GoTo GetScheduleDetail_EH
    
    #If TRDEBUG Then
        stTime = GetTickCount()
    #End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT  SCHED_CD, SCHED_DESC, SCHED_SHIFT_CD "
    l_strSQL = l_strSQL & _
                ", SUN_ACT_DAY, SUN_SHIFT_CD, SUN_START_TIME" & _
                ", SUN_END_TIME, SUN_LUNCH_START, SUN_LUNCH_END" & _
                ", SUN_STR_HRS, SUN_OT_START_TIME, SUN_OT_END_TIME" & _
                ", SUN_OT_HOURS, SUN_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", MON_ACT_DAY, MON_SHIFT_CD, MON_START_TIME" & _
                ", MON_END_TIME, MON_LUNCH_START, MON_LUNCH_END" & _
                ", MON_STR_HRS, MON_OT_START_TIME, MON_OT_END_TIME" & _
                ", MON_OT_HOURS, MON_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", TUE_ACT_DAY, TUE_SHIFT_CD, TUE_START_TIME" & _
                ", TUE_END_TIME, TUE_LUNCH_START, TUE_LUNCH_END" & _
                ", TUE_STR_HRS, TUE_OT_START_TIME, TUE_OT_END_TIME" & _
                ", TUE_OT_HOURS, TUE_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", WED_ACT_DAY, WED_SHIFT_CD, WED_START_TIME" & _
                ", WED_END_TIME, WED_LUNCH_START, WED_LUNCH_END" & _
                ", WED_STR_HRS, WED_OT_START_TIME, WED_OT_END_TIME" & _
                ", WED_OT_HOURS, WED_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", THU_ACT_DAY, THU_SHIFT_CD, THU_START_TIME" & _
                ", THU_END_TIME, THU_LUNCH_START, THU_LUNCH_END" & _
                ", THU_STR_HRS, THU_OT_START_TIME, THU_OT_END_TIME" & _
                ", THU_OT_HOURS, THU_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", FRI_ACT_DAY, FRI_SHIFT_CD, FRI_START_TIME" & _
                ", FRI_END_TIME, FRI_LUNCH_START, FRI_LUNCH_END" & _
                ", FRI_STR_HRS, FRI_OT_START_TIME, FRI_OT_END_TIME" & _
                ", FRI_OT_HOURS, FRI_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", SAT_ACT_DAY, SAT_SHIFT_CD, SAT_START_TIME" & _
                ", SAT_END_TIME, SAT_LUNCH_START, SAT_LUNCH_END" & _
                ", SAT_STR_HRS, SAT_OT_START_TIME, SAT_OT_END_TIME" & _
                ", SAT_OT_HOURS, SAT_COMP_SHIFT_CD, LAST_MOD_EMP_NUM,Char(TIME_STAMP) as TIME_STAMP "
    l_strSQL = l_strSQL & _
                " FROM       METHDBA.TTR_AVAIL_SCHED" & _
                " WHERE      SCHED_CD = " & v_lSchedCode & _
                " ORDER BY   SCHED_CD"


    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetScheduleDetail = Nothing
    Else
        Set GetScheduleDetail = l_objRS
    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing
    
    #If TRDEBUG Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        Debug.Print "GetScheduleDetail: " & CStr(dfTime / 1000)
    #End If
    
    Exit Function

GetScheduleDetail_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "GetScheduleDetail", , , , DescArgsFrom(l_strSQL), True
End Function


