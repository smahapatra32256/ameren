VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdAdjustments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Conversion Phase 2/3
'Authors:   Jim Barnett (TRIS Conversion Phase 2)
'           Prem Chander (TRIS Conversion Phase 3)
'Advisor:   John Stricklin
'Notes:     Method GetAdjustmentDetails was completely modified to
'retrieve information from new tables TTR_ADJ_RQT; TTR_ADJ_TIME
'This method is used by busSchedExec.cbsSchedExec to get information
'for client business function ViewCalendar. -- Prem Nov. 2004

Option Explicit

Private Const MODULENAME As String = "dbAdjustments.cdAdjustments"
Private Const m_DBSCRIPTSROOT As String = "C:\TRIS SOLUTION\dbScripts\TRIS_db2"
Private Const m_DBSCRIPTSPATH As String = "\dbAdjustments\cdAdjustments"
Private m_ResourceMgr As utilADOHelper.cuResourceMgr

Public Function GetAdjustmentDetails(ByVal v_aryEmployees As Variant _
                , ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                ) As ADODB.Recordset
On Error GoTo GetAdjustmentDetails_EH
'SQL Statement Containers
Dim lsSQLStatement As String, lsWHEREClauseNewTable As String
Dim lsWHEREClauseOldTable As String

    Set GetAdjustmentDetails = Nothing

    If (IsArray(v_aryEmployees)) Then
        lsWHEREClauseNewTable = BuildInClause("B.EMP_NUM", v_aryEmployees, False)
        lsWHEREClauseOldTable = Replace(lsWHEREClauseNewTable, "B.", "")
    Else
        lsWHEREClauseNewTable = "B.EMP_NUM = " & CStr(CLng(v_aryEmployees))
        lsWHEREClauseOldTable = "EMP_NUM = " & CStr(CLng(v_aryEmployees))
    End If

    'Check for New Tbl Vs. Old Tbl cutoff (Jan. 26, 2005)
    If (v_dtStartDate < #1/27/2005#) Then
        'Get from new and old Tables
        lsSQLStatement = "SELECT  A.ADJ_PREV_PAID_FL, A.ADJ_SEQ_NUM AS SEQ_NUM" & _
            "       , CASE WHEN (A.HOURS < 0) OR (A.AMOUNT < 0) THEN 'A'" & _
            "       ELSE 'S' END AS ADJ_TYPE" & _
            "       , A.AMOUNT" & _
            "       , A.APPR_BY_EMP_NUM AS APPR_BY_MANAGER" & _
            "       , VALUE(NULLIF(A.APPR_DATE,'0001-01-01'), DATE('0100-01-01')) AS APPR_DATE" & _
            "       , A.BANK_YEAR, A.CASE_NUM" & _
            "       , VALUE(NULLIF(A.CHARGE_DATE,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , A.CHEM_MTL_TSTR_IND, A.COMP_CD" & _
            "       , A.COMP_SHIFT_CD AS SHIFT_COMP, COALESCE(B.COMPLETED_BY_EMP, 0) AS COMPLETED_BY_EMP" & _
            "       , VALUE(NULLIF(DATE(B.COMPLETED_TM_STMP),'0001-01-01'), DATE('0100-01-01')) AS COMPLETED_DATE" & _
            "       , CHAR(B.CREATED_TIMESTAMP) AS TIME_STAMP" & _
            "       , VALUE(NULLIF(A.DATE_WORKED,'0001-01-01'), DATE('0100-01-01')) AS DATE_WORKED" & _
            "       , A.DAY_NAME, A.DEPT_NUM" & _
            "       , 'Y' AS DETAIL_IND, A.DISTR_DEPT_NUM" & _
            "       , A.DISTRIBUTION_ID, A.DISTRIBUTION_TYPE" & _
            "       , A.EARNINGS_CD, B.EMP_NUM" & _
            "       , A.END_LOC_CD, A.END_TIME" & _
            "       , A.EXCEPTION_CD, A.FAMILY_LEAVE_IND" & _
            "       , VALUE(NULLIF(DATE(B.ATHRZ_TM_STMP),'0001-01-01'), DATE('0100-01-01')) AS FINAL_APPR_DATE" & _
            "       , COALESCE(B.ATHRZ_EMP_NUM, 0) AS FINAL_APPROVER, A.HALF_TRIP_IND" & _
            "       , A.HOURLY_RATE, A.HOURS" & _
            "       , A.INTERFACE_IND, CASE (A.COMP_CD)" & _
            "       WHEN 'IC ' THEN 'IC '" & _
            "       ELSE '   ' END AS INTERPLANT_CD"
        lsSQLStatement = lsSQLStatement & _
            "       , A.JOB_NUM, B.LAST_MOD_EMP_NUM" & _
            "       , CHAR(B.LAST_MOD_TIMESTAMP) AS INDEX_TIME, A.LOC_CD" & _
            "       , CASE (B.FINAL_CYCLE_NUM)" & _
            "       WHEN 0 THEN 'Y'" & _
            "       ELSE 'N' END AS PROCESSED_IND" & _
            "       , VALUE(NULLIF(A.OT_BANK_CHRG_DT,'0001-01-01'), DATE('0100-01-01')) AS OT_BANK_CHRG_DT" & _
            "       , A.OT_OPP_DESC" & _
            "       , A.OT_OPP_NUM, A.OT_REASON_CD" & _
            "       , A.OVER_EXC_CD, A.REL_CD" & _
            "       , B.CREATED_EMP_NUM AS REQ_BY_EMP_NUM" & _
            "       , VALUE(NULLIF(DATE(B.CREATED_TIMESTAMP),'0001-01-01'), DATE('0100-01-01')) AS REQ_DATE" & _
            "       , COALESCE(B.REQ_REASON, '') AS REQ_REASON" & _
            "       , B.SPECIAL_CHECK_IND AS SPECIAL_CHECK" & _
            "       , A.SCHED_CD, A.SCHED_CHG_EXC_CD" & _
            "       , A.SCHED_SRC_CD, A.START_LOC_CD" & _
            "       , A.START_TIME, A.TRADE_EXC_CD" & _
            "       , A.TRAVEL_PAY_HRS, A.UNION_CD" & _
            "       , A.WG_NUM, A.ADJ_RQT_NUM" & _
            " FROM  METHDBA.TTR_ADJ_TIME A" & _
            "       , METHDBA.TTR_ADJ_RQT B" & _
            " WHERE " & lsWHEREClauseNewTable & _
            " AND   A.CHARGE_DATE BETWEEN '" & CStr(v_dtStartDate) & "'" & _
            " AND   '" & CStr(v_dtEndDate) & "'" & _
            " AND   A.ADJ_RQT_NUM = B.ADJ_RQT_NUM"

        lsSQLStatement = lsSQLStatement & _
            " UNION" & _
            " SELECT    ADJ_PREV_PAID_FL, SEQ_NUM" & _
            "           , ADJ_TYPE" & _
            "           , CASE WHEN (ADJ_TYPE = 'A') THEN (AMOUNT * -1)" & _
            "           ELSE AMOUNT END AS AMOUNT" & _
            "           , APPR_BY_MANAGER" & _
            "           , VALUE(NULLIF(APPR_DATE,'0001-01-01'), DATE('0100-01-01')) AS APPR_DATE" & _
            "           , 0 AS BANK_YEAR, CASE_NUM" & _
            "           , DATE('0100-01-01') AS CHARGE_DATE, '' AS CHEM_MTL_TSTR_IND" & _
            "           , '' AS COMP_CD, '' AS SHIFT_COMP" & _
            "           , COALESCE(COMPLETED_BY_EMP, 0) AS COMPLETED_BY_EMP" & _
            "           , VALUE(NULLIF(COMPLETED_DATE,'0001-01-01'), DATE('0100-01-01')) AS COMPLETED_DATE" & _
            "           , CHAR(TIME_STAMP) AS TIME_STAMP" & _
            "           , VALUE(NULLIF(DATE_WORKED,'0001-01-01'), DATE('0100-01-01')) AS DATE_WORKED" & _
            "           , '' AS DAY_NAME, DEPT_NUM" & _
            "           , DETAIL_IND, '' AS DISTR_DEPT_NUM" & _
            "           , '' AS DISTRIBUTION_ID, '' AS DISTRIBUTION_TYPE" & _
            "           , EARNINGS_CD, EMP_NUM" & _
            "           , 0 AS END_LOC_CD, END_TIME" & _
            "           , EXCEPTION_CD, FAMILY_LEAVE_IND" & _
            "           , VALUE(NULLIF(FINAL_APPR_DATE,'0001-01-01'), DATE('0100-01-01')) AS FINAL_APPR_DATE" & _
            "           , COALESCE(FINAL_APPROVER, 0) AS FINAL_APPROVER, '' AS HALF_TRIP_IND" & _
            "           , HOURLY_RATE"
        lsSQLStatement = lsSQLStatement & _
            "           , CASE WHEN (ADJ_TYPE = 'A') THEN (HOURS * -1)" & _
            "           ELSE HOURS END AS HOURS" & _
            "           , '' AS INTERFACE_IND, INTERPLANT_CD"
        lsSQLStatement = lsSQLStatement & _
            "           , JOB_NUM, LAST_MOD_EMP_NUM" & _
            "           , CHAR(INDEX_TIME) AS INDEX_TIME, 0 AS LOC_CD" & _
            "           , PROCESSED_IND, DATE('0100-01-01') AS OT_BANK_CHRG_DT" & _
            "           , '' AS OT_OPP_DESC, 0 AS OT_OPP_NUM" & _
            "           , '' AS OT_REASON_CD, '' AS OVER_EXC_CD" & _
            "           , REL_CD, REQ_BY_EMP_NUM" & _
            "           , VALUE(NULLIF(REQ_DATE,'0001-01-01'), DATE('0100-01-01')) AS REQ_DATE" & _
            "           , COALESCE(REQ_REASON, '') AS REQ_REASON" & _
            "           , SPECIAL_CHECK, 0 AS SCHED_CD" & _
            "           , '' AS SCHED_CHG_EXC_CD, 0 AS SCHED_SRC_CD" & _
            "           , 0 AS START_LOC_CD, START_TIME" & _
            "           , '' AS TRADE_EXC_CD, 0 AS TRAVEL_PAY_HRS" & _
            "           , UNION_CD, 0 AS WG_NUM" & _
            "           , -1 AS ADJ_RQT_NUM" & _
            " FROM      METHDBA.TTR_ADJUSTMENTS" & _
            " WHERE     " & lsWHEREClauseOldTable & _
            " AND       DATE_WORKED BETWEEN '" & CStr(v_dtStartDate) & "'" & _
            " AND       '" & CStr(v_dtEndDate) & "'"
    Else
        'Get from new table alone
        lsSQLStatement = "SELECT  A.ADJ_PREV_PAID_FL, A.ADJ_SEQ_NUM AS SEQ_NUM" & _
            "       , CASE WHEN (A.HOURS < 0) OR (A.AMOUNT < 0) THEN 'A'" & _
            "       ELSE 'S' END AS ADJ_TYPE" & _
            "       , A.AMOUNT, A.APPR_BY_EMP_NUM AS APPR_BY_MANAGER" & _
            "       , VALUE(NULLIF(A.APPR_DATE,'0001-01-01'), DATE('0100-01-01')) AS APPR_DATE" & _
            "       , A.BANK_YEAR, A.CASE_NUM" & _
            "       , VALUE(NULLIF(A.CHARGE_DATE,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , A.CHEM_MTL_TSTR_IND, A.COMP_CD" & _
            "       , A.COMP_SHIFT_CD AS SHIFT_COMP, COALESCE(B.COMPLETED_BY_EMP, 0) AS COMPLETED_BY_EMP" & _
            "       , VALUE(NULLIF(DATE(B.COMPLETED_TM_STMP),'0001-01-01'), DATE('0100-01-01')) AS COMPLETED_DATE" & _
            "       , CHAR(B.CREATED_TIMESTAMP) AS TIME_STAMP" & _
            "       , VALUE(NULLIF(A.DATE_WORKED,'0001-01-01'), DATE('0100-01-01')) AS DATE_WORKED" & _
            "       , A.DAY_NAME, A.DEPT_NUM" & _
            "       , 'Y' AS DETAIL_IND, A.DISTR_DEPT_NUM" & _
            "       , A.DISTRIBUTION_ID, A.DISTRIBUTION_TYPE" & _
            "       , A.EARNINGS_CD, B.EMP_NUM" & _
            "       , A.END_LOC_CD, A.END_TIME" & _
            "       , A.EXCEPTION_CD, A.FAMILY_LEAVE_IND" & _
            "       , VALUE(NULLIF(DATE(B.ATHRZ_TM_STMP),'0001-01-01'), DATE('0100-01-01')) AS FINAL_APPR_DATE" & _
            "       , COALESCE(B.ATHRZ_EMP_NUM, 0) AS FINAL_APPROVER, A.HALF_TRIP_IND" & _
            "       , A.HOURLY_RATE, A.HOURS" & _
            "       , A.INTERFACE_IND, CASE (A.COMP_CD)" & _
            "       WHEN 'IC ' THEN 'IC '" & _
            "       ELSE '   ' END AS INTERPLANT_CD"
        lsSQLStatement = lsSQLStatement & _
            "       , A.JOB_NUM, B.LAST_MOD_EMP_NUM" & _
            "       , CHAR(B.LAST_MOD_TIMESTAMP) AS INDEX_TIME, A.LOC_CD" & _
            "       , CASE (B.FINAL_CYCLE_NUM)" & _
            "       WHEN 0 THEN 'Y'" & _
            "       ELSE 'N' END AS PROCESSED_IND" & _
            "       , VALUE(NULLIF(A.OT_BANK_CHRG_DT,'0001-01-01'), DATE('0100-01-01')) AS OT_BANK_CHRG_DT" & _
            "       , A.OT_OPP_DESC" & _
            "       , A.OT_OPP_NUM, A.OT_REASON_CD" & _
            "       , A.OVER_EXC_CD, A.REL_CD" & _
            "       , B.CREATED_EMP_NUM AS REQ_BY_EMP_NUM" & _
            "       , VALUE(NULLIF(DATE(B.CREATED_TIMESTAMP),'0001-01-01'), DATE('0100-01-01')) AS REQ_DATE" & _
            "       , COALESCE(B.REQ_REASON, '') AS REQ_REASON" & _
            "       , B.SPECIAL_CHECK_IND AS SPECIAL_CHECK" & _
            "       , A.SCHED_CD, A.SCHED_CHG_EXC_CD" & _
            "       , A.SCHED_SRC_CD, A.START_LOC_CD" & _
            "       , A.START_TIME, A.TRADE_EXC_CD" & _
            "       , A.TRAVEL_PAY_HRS, A.UNION_CD" & _
            "       , A.WG_NUM, A.ADJ_RQT_NUM" & _
            " FROM  METHDBA.TTR_ADJ_TIME A" & _
            "       , METHDBA.TTR_ADJ_RQT B" & _
            " WHERE " & lsWHEREClauseNewTable & _
            " AND   A.CHARGE_DATE BETWEEN '" & CStr(v_dtStartDate) & "'" & _
            " AND   '" & CStr(v_dtEndDate) & "'" & _
            " AND   A.ADJ_RQT_NUM = B.ADJ_RQT_NUM"
    End If

    'Call ExecuteSELECTSQL and retrieve information
    Set GetAdjustmentDetails = ExecuteSQL(lsSQLStatement _
                                , adOpenStatic _
                                , adLockBatchOptimistic)

    GetObjectContext.SetComplete

    Exit Function
GetAdjustmentDetails_EH:
    GetObjectContext.SetAbort

    TrsRaiseError MODULENAME _
        , "GetAdjustmentDetails" _
        , , , , DescArgsFrom(lsSQLStatement) _
        , True

End Function
Public Function GetDistributionAdjustmentDetails(ByVal v_aryEmployees As Variant _
                , ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_adrsDistributions As ADODB.Recordset _
                ) As ADODB.Recordset
On Error GoTo GetDistributionAdjustmentDetails_EH
'SQL Statement Containers
Dim lsSQLStatement As String, lsWHEREClauseNewTable As String
Dim lsWHEREClauseOldTable As String

    Set GetDistributionAdjustmentDetails = Nothing

    If (IsArray(v_aryEmployees)) Then
        lsWHEREClauseNewTable = BuildInClause("B.EMP_NUM", v_aryEmployees, False)
        lsWHEREClauseOldTable = Replace(lsWHEREClauseNewTable, "B.", "")
    Else
        lsWHEREClauseNewTable = "B.EMP_NUM = " & CStr(CLng(v_aryEmployees))
        lsWHEREClauseOldTable = "EMP_NUM = " & CStr(CLng(v_aryEmployees))
    End If
    

    'Check for New Tbl Vs. Old Tbl cutoff (Jan. 26, 2005)
    If (v_dtStartDate < #1/27/2005#) Then
        'Get from new and old Tables
        lsSQLStatement = "SELECT  A.ADJ_PREV_PAID_FL, A.ADJ_SEQ_NUM AS SEQ_NUM" & _
            "       , CASE WHEN (A.HOURS < 0) OR (A.AMOUNT < 0) THEN 'A'" & _
            "       ELSE 'S' END AS ADJ_TYPE" & _
            "       , A.AMOUNT" & _
            "       , A.APPR_BY_EMP_NUM AS APPR_BY_MANAGER" & _
            "       , VALUE(NULLIF(A.APPR_DATE,'0001-01-01'), DATE('0100-01-01')) AS APPR_DATE" & _
            "       , A.BANK_YEAR, A.CASE_NUM" & _
            "       , VALUE(NULLIF(A.CHARGE_DATE,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , A.CHEM_MTL_TSTR_IND, A.COMP_CD" & _
            "       , A.COMP_SHIFT_CD AS SHIFT_COMP, COALESCE(B.COMPLETED_BY_EMP, 0) AS COMPLETED_BY_EMP" & _
            "       , VALUE(NULLIF(DATE(B.COMPLETED_TM_STMP),'0001-01-01'), DATE('0100-01-01')) AS COMPLETED_DATE" & _
            "       , CHAR(B.CREATED_TIMESTAMP) AS TIME_STAMP" & _
            "       , VALUE(NULLIF(A.DATE_WORKED,'0001-01-01'), DATE('0100-01-01')) AS DATE_WORKED" & _
            "       , A.DAY_NAME, A.DEPT_NUM" & _
            "       , 'Y' AS DETAIL_IND, A.DISTR_DEPT_NUM" & _
            "       , A.DISTRIBUTION_ID, A.DISTRIBUTION_TYPE" & _
            "       , A.EARNINGS_CD, B.EMP_NUM" & _
            "       , A.END_LOC_CD, A.END_TIME" & _
            "       , A.EXCEPTION_CD, A.FAMILY_LEAVE_IND" & _
            "       , VALUE(NULLIF(DATE(B.ATHRZ_TM_STMP),'0001-01-01'), DATE('0100-01-01')) AS FINAL_APPR_DATE" & _
            "       , COALESCE(B.ATHRZ_EMP_NUM, 0) AS FINAL_APPROVER, A.HALF_TRIP_IND" & _
            "       , A.HOURLY_RATE, A.HOURS" & _
            "       , A.INTERFACE_IND, CASE (A.COMP_CD)" & _
            "       WHEN 'IC ' THEN 'IC '" & _
            "       ELSE '   ' END AS INTERPLANT_CD"
        lsSQLStatement = lsSQLStatement & _
            "       , A.JOB_NUM, B.LAST_MOD_EMP_NUM" & _
            "       , CHAR(B.LAST_MOD_TIMESTAMP) AS INDEX_TIME, A.LOC_CD" & _
            "       , CASE (B.FINAL_CYCLE_NUM)" & _
            "       WHEN 0 THEN 'Y'" & _
            "       ELSE 'N' END AS PROCESSED_IND" & _
            "       , VALUE(NULLIF(A.OT_BANK_CHRG_DT,'0001-01-01'), DATE('0100-01-01')) AS OT_BANK_CHRG_DT" & _
            "       , A.OT_OPP_DESC" & _
            "       , A.OT_OPP_NUM, A.OT_REASON_CD" & _
            "       , A.OVER_EXC_CD, A.REL_CD" & _
            "       , B.CREATED_EMP_NUM AS REQ_BY_EMP_NUM" & _
            "       , VALUE(NULLIF(DATE(B.CREATED_TIMESTAMP),'0001-01-01'), DATE('0100-01-01')) AS REQ_DATE" & _
            "       , COALESCE(B.REQ_REASON, '') AS REQ_REASON" & _
            "       , B.SPECIAL_CHECK_IND AS SPECIAL_CHECK" & _
            "       , A.SCHED_CD, A.SCHED_CHG_EXC_CD" & _
            "       , A.SCHED_SRC_CD, A.START_LOC_CD" & _
            "       , A.START_TIME, A.TRADE_EXC_CD" & _
            "       , A.TRAVEL_PAY_HRS, A.UNION_CD" & _
            "       , A.WG_NUM, A.ADJ_RQT_NUM" & _
            " FROM  METHDBA.TTR_ADJ_TIME A" & _
            "       , METHDBA.TTR_ADJ_RQT B" & _
            " WHERE " & lsWHEREClauseNewTable & _
            " AND   A.CHARGE_DATE BETWEEN '" & CStr(v_dtStartDate) & "'" & _
            " AND   '" & CStr(v_dtEndDate) & "'" & _
            " AND   A.ADJ_RQT_NUM = B.ADJ_RQT_NUM"
            
        lsSQLStatement = lsSQLStatement & _
            " AND (" & BuildDistributionSelection("A", v_adrsDistributions) & ") "

        lsSQLStatement = lsSQLStatement & _
            " UNION" & _
            " SELECT    ADJ_PREV_PAID_FL, SEQ_NUM" & _
            "           , ADJ_TYPE" & _
            "           , CASE WHEN (ADJ_TYPE = 'A') THEN (AMOUNT * -1)" & _
            "           ELSE AMOUNT END AS AMOUNT" & _
            "           , APPR_BY_MANAGER" & _
            "           , VALUE(NULLIF(APPR_DATE,'0001-01-01'), DATE('0100-01-01')) AS APPR_DATE" & _
            "           , 0 AS BANK_YEAR, CASE_NUM" & _
            "           , DATE('0100-01-01') AS CHARGE_DATE, '' AS CHEM_MTL_TSTR_IND" & _
            "           , '' AS COMP_CD, '' AS SHIFT_COMP" & _
            "           , COALESCE(COMPLETED_BY_EMP, 0) AS COMPLETED_BY_EMP" & _
            "           , VALUE(NULLIF(COMPLETED_DATE,'0001-01-01'), DATE('0100-01-01')) AS COMPLETED_DATE" & _
            "           , CHAR(TIME_STAMP) AS TIME_STAMP" & _
            "           , VALUE(NULLIF(DATE_WORKED,'0001-01-01'), DATE('0100-01-01')) AS DATE_WORKED" & _
            "           , '' AS DAY_NAME, DEPT_NUM" & _
            "           , DETAIL_IND, '' AS DISTR_DEPT_NUM" & _
            "           , '' AS DISTRIBUTION_ID, '' AS DISTRIBUTION_TYPE" & _
            "           , EARNINGS_CD, EMP_NUM" & _
            "           , 0 AS END_LOC_CD, END_TIME" & _
            "           , EXCEPTION_CD, FAMILY_LEAVE_IND" & _
            "           , VALUE(NULLIF(FINAL_APPR_DATE,'0001-01-01'), DATE('0100-01-01')) AS FINAL_APPR_DATE" & _
            "           , COALESCE(FINAL_APPROVER, 0) AS FINAL_APPROVER, '' AS HALF_TRIP_IND" & _
            "           , HOURLY_RATE"
        lsSQLStatement = lsSQLStatement & _
            "           , CASE WHEN (ADJ_TYPE = 'A') THEN (HOURS * -1)" & _
            "           ELSE HOURS END AS HOURS" & _
            "           , '' AS INTERFACE_IND, INTERPLANT_CD"
        lsSQLStatement = lsSQLStatement & _
            "           , JOB_NUM, LAST_MOD_EMP_NUM" & _
            "           , CHAR(INDEX_TIME) AS INDEX_TIME, 0 AS LOC_CD" & _
            "           , PROCESSED_IND, DATE('0100-01-01') AS OT_BANK_CHRG_DT" & _
            "           , '' AS OT_OPP_DESC, 0 AS OT_OPP_NUM" & _
            "           , '' AS OT_REASON_CD, '' AS OVER_EXC_CD" & _
            "           , REL_CD, REQ_BY_EMP_NUM" & _
            "           , VALUE(NULLIF(REQ_DATE,'0001-01-01'), DATE('0100-01-01')) AS REQ_DATE" & _
            "           , COALESCE(REQ_REASON, '') AS REQ_REASON" & _
            "           , SPECIAL_CHECK, 0 AS SCHED_CD" & _
            "           , '' AS SCHED_CHG_EXC_CD, 0 AS SCHED_SRC_CD" & _
            "           , 0 AS START_LOC_CD, START_TIME" & _
            "           , '' AS TRADE_EXC_CD, 0 AS TRAVEL_PAY_HRS" & _
            "           , UNION_CD, 0 AS WG_NUM" & _
            "           , -1 AS ADJ_RQT_NUM" & _
            " FROM      METHDBA.TTR_ADJUSTMENTS" & _
            " WHERE     " & lsWHEREClauseOldTable & _
            " AND       DATE_WORKED BETWEEN '" & CStr(v_dtStartDate) & "'" & _
            " AND       '" & CStr(v_dtEndDate) & "'"
    Else
        'Get from new table alone
        lsSQLStatement = "SELECT  A.ADJ_PREV_PAID_FL, A.ADJ_SEQ_NUM AS SEQ_NUM" & _
            "       , CASE WHEN (A.HOURS < 0) OR (A.AMOUNT < 0) THEN 'A'" & _
            "       ELSE 'S' END AS ADJ_TYPE" & _
            "       , A.AMOUNT, A.APPR_BY_EMP_NUM AS APPR_BY_MANAGER" & _
            "       , VALUE(NULLIF(A.APPR_DATE,'0001-01-01'), DATE('0100-01-01')) AS APPR_DATE" & _
            "       , A.BANK_YEAR, A.CASE_NUM" & _
            "       , VALUE(NULLIF(A.CHARGE_DATE,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , A.CHEM_MTL_TSTR_IND, A.COMP_CD" & _
            "       , A.COMP_SHIFT_CD AS SHIFT_COMP, COALESCE(B.COMPLETED_BY_EMP, 0) AS COMPLETED_BY_EMP" & _
            "       , VALUE(NULLIF(DATE(B.COMPLETED_TM_STMP),'0001-01-01'), DATE('0100-01-01')) AS COMPLETED_DATE" & _
            "       , CHAR(B.CREATED_TIMESTAMP) AS TIME_STAMP" & _
            "       , VALUE(NULLIF(A.DATE_WORKED,'0001-01-01'), DATE('0100-01-01')) AS DATE_WORKED" & _
            "       , A.DAY_NAME, A.DEPT_NUM" & _
            "       , 'Y' AS DETAIL_IND, A.DISTR_DEPT_NUM" & _
            "       , A.DISTRIBUTION_ID, A.DISTRIBUTION_TYPE" & _
            "       , A.EARNINGS_CD, B.EMP_NUM" & _
            "       , A.END_LOC_CD, A.END_TIME" & _
            "       , A.EXCEPTION_CD, A.FAMILY_LEAVE_IND" & _
            "       , VALUE(NULLIF(DATE(B.ATHRZ_TM_STMP),'0001-01-01'), DATE('0100-01-01')) AS FINAL_APPR_DATE" & _
            "       , COALESCE(B.ATHRZ_EMP_NUM, 0) AS FINAL_APPROVER, A.HALF_TRIP_IND" & _
            "       , A.HOURLY_RATE, A.HOURS" & _
            "       , A.INTERFACE_IND, CASE (A.COMP_CD)" & _
            "       WHEN 'IC ' THEN 'IC '" & _
            "       ELSE '   ' END AS INTERPLANT_CD"
        lsSQLStatement = lsSQLStatement & _
            "       , A.JOB_NUM, B.LAST_MOD_EMP_NUM" & _
            "       , CHAR(B.LAST_MOD_TIMESTAMP) AS INDEX_TIME, A.LOC_CD" & _
            "       , CASE (B.FINAL_CYCLE_NUM)" & _
            "       WHEN 0 THEN 'Y'" & _
            "       ELSE 'N' END AS PROCESSED_IND" & _
            "       , VALUE(NULLIF(A.OT_BANK_CHRG_DT,'0001-01-01'), DATE('0100-01-01')) AS OT_BANK_CHRG_DT" & _
            "       , A.OT_OPP_DESC" & _
            "       , A.OT_OPP_NUM, A.OT_REASON_CD" & _
            "       , A.OVER_EXC_CD, A.REL_CD" & _
            "       , B.CREATED_EMP_NUM AS REQ_BY_EMP_NUM" & _
            "       , VALUE(NULLIF(DATE(B.CREATED_TIMESTAMP),'0001-01-01'), DATE('0100-01-01')) AS REQ_DATE" & _
            "       , COALESCE(B.REQ_REASON, '') AS REQ_REASON" & _
            "       , B.SPECIAL_CHECK_IND AS SPECIAL_CHECK" & _
            "       , A.SCHED_CD, A.SCHED_CHG_EXC_CD" & _
            "       , A.SCHED_SRC_CD, A.START_LOC_CD" & _
            "       , A.START_TIME, A.TRADE_EXC_CD" & _
            "       , A.TRAVEL_PAY_HRS, A.UNION_CD" & _
            "       , A.WG_NUM, A.ADJ_RQT_NUM" & _
            " FROM  METHDBA.TTR_ADJ_TIME A" & _
            "       , METHDBA.TTR_ADJ_RQT B" & _
            " WHERE " & lsWHEREClauseNewTable & _
            " AND   A.CHARGE_DATE BETWEEN '" & CStr(v_dtStartDate) & "'" & _
            " AND   '" & CStr(v_dtEndDate) & "'" & _
            " AND   A.ADJ_RQT_NUM = B.ADJ_RQT_NUM"
            
        lsSQLStatement = lsSQLStatement & _
            " AND (" & BuildDistributionSelection("A", v_adrsDistributions) & ") "
            
    End If

    'Call ExecuteSELECTSQL and retrieve information
    Set GetDistributionAdjustmentDetails = ExecuteSQL(lsSQLStatement _
                                , adOpenStatic _
                                , adLockBatchOptimistic)

    GetObjectContext.SetComplete

    Exit Function
GetDistributionAdjustmentDetails_EH:
    GetObjectContext.SetAbort

    TrsRaiseError MODULENAME _
        , "GetDistributionAdjustmentDetails" _
        , , , , DescArgsFrom(lsSQLStatement) _
        , True

End Function

Private Function BuildDistributionSelection( _
            ByVal v_sTablePrefix As String, _
            ByVal v_adrsDistributions As ADODB.Recordset) As String
            
    Dim l_strSelection As String
    Dim l_strDistribution As String
    Dim i As Integer

    l_strDistribution = ""
    
    If RecordsExist(v_adrsDistributions) Then
        With v_adrsDistributions
            .MoveFirst
            l_strDistribution = Trim(UCase(!DISTRIBUTION_ID))
            For i = 1 To Len(l_strDistribution)
                If Mid(l_strDistribution, i, 1) = " " Then
                    Mid(l_strDistribution, i, 1) = "%"
                End If
            Next i
            l_strSelection = v_sTablePrefix & ".DISTRIBUTION_ID LIKE '%" & l_strDistribution & "%' "
            .MoveNext
            Do While Not .EOF
                l_strDistribution = Trim(UCase(!DISTRIBUTION_ID))
                For i = 1 To Len(l_strDistribution)
                    If Mid(l_strDistribution, i, 1) = " " Then
                        Mid(l_strDistribution, i, 1) = "%"
                    End If
                Next i
                l_strSelection = l_strSelection & " OR "
                l_strSelection = l_strSelection & v_sTablePrefix & ".DISTRIBUTION_ID like '%" & l_strDistribution & "%' "
                .MoveNext
            Loop
        End With
    Else
        l_strSelection = ""
    End If
    
    BuildDistributionSelection = l_strSelection
            
End Function

'Returns a boolean indicating whether the recordset is empty or not
Public Function RecordsExist( _
                r_rsAnyRecordset As ADODB.Recordset) As Boolean

On Error GoTo eh

    If Not r_rsAnyRecordset Is Nothing Then
        If Not (r_rsAnyRecordset.BOF And r_rsAnyRecordset.EOF) Then
            RecordsExist = True
        End If
    End If
    Exit Function

eh:
    Err.Raise Err.Number, "mADOExtensions.RecordsExist"

End Function

'Begin Prem TRIS 5.8
Public Function GetAdjustmentsStatusByScope(ByVal vlTRISUserNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetAdjustmentsStatusByScope_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    lsSQL = "SELECT A.ADJ_RQT_NUM, LTRIM(RTRIM(B.EMP_NAME)) || ' [' || LTRIM(RTRIM(CHAR(A.EMP_NUM))) || ']' AS EMPLOYEE" & _
        "       , A.DEPT_NUM" & _
        "       , CASE WHEN (A.FINAL_CYCLE_NUM IS NOT NULL) THEN 5" & _
        "       ELSE CASE WHEN (A.ATHRZ_EMP_NUM > 0) THEN 4" & _
        "       ELSE CASE WHEN (A.COMPLETED_BY_EMP > 0) THEN 3" & _
        "       ELSE CASE WHEN (A.APPROVED_EMP_NUM > 0) THEN 2" & _
        "       ELSE 1" & _
        "       END END END END AS ADJ_STATUS" & _
        "       , CASE WHEN (A.FINAL_CYCLE_NUM IS NOT NULL) THEN '5 - Payroll processed'" & _
        "       ELSE CASE WHEN (A.ATHRZ_EMP_NUM > 0) THEN '4 - Payroll authorized'" & _
        "       ELSE CASE WHEN (A.COMPLETED_BY_EMP > 0) THEN '3 - Payroll completed'" & _
        "       ELSE CASE WHEN (A.APPROVED_EMP_NUM > 0) THEN '2 - Department head approved'" & _
        "       ELSE '1 - Requested'" & _
        "       END END END END AS ADJ_STATUS_TXT" & _
        "       , A.REQ_REASON, CASE (UCASE(A.SPECIAL_CHECK_IND))" & _
        "       WHEN 'Y' THEN 1" & _
        "       ELSE 0 END AS SPECIAL_CHECK_IND"
    lsSQL = lsSQL & _
        "       , LTRIM(RTRIM(COALESCE(C.EMP_NAME, ''))) || ' [' || LTRIM(RTRIM(CHAR(COALESCE(A.APPROVED_EMP_NUM, 0)))) || ']' AS APPROVED_EMP" & _
        "       , COALESCE(CHAR(A.APPROVED_TM_STMP), '') AS APPROVED_TM_STMP" & _
        "       , LTRIM(RTRIM(COALESCE(D.EMP_NAME, ''))) || ' [' || LTRIM(RTRIM(CHAR(COALESCE(A.COMPLETED_BY_EMP, 0)))) || ']' AS COMPLETED_BY_EMP" & _
        "       , COALESCE(CHAR(A.COMPLETED_TM_STMP), '') AS COMPLETED_TM_STMP" & _
        "       , LTRIM(RTRIM(COALESCE(E.EMP_NAME, ''))) || ' [' || LTRIM(RTRIM(CHAR(COALESCE(A.ATHRZ_EMP_NUM, 0)))) || ']' AS ATHRZ_EMP" & _
        "       , COALESCE(CHAR(A.ATHRZ_TM_STMP), '') AS ATHRZ_TM_STMP" & _
        "       , COALESCE(A.FINAL_CYCLE_NUM, 0) AS FINAL_CYCLE_NUM" & _
        "       , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP"
    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_ADJ_RQT A" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON B.EMP_NUM = A.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE C" & _
        "       ON C.EMP_NUM = A.APPROVED_EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE D" & _
        "       ON D.EMP_NUM = A.COMPLETED_BY_EMP" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE E" & _
        "       ON E.EMP_NUM = A.ATHRZ_EMP_NUM" & _
        " WHERE A.EMP_NUM IN (SELECT    EMP_NUM" & _
        "                   FROM        METHDBA.TTR_DS_EMP_DTL" & _
        "                   WHERE       EFF_EMP_NUM = ?)" & _
        " AND   DATE(A.CREATED_TIMESTAMP) BETWEEN ?" & _
        " AND   ?" & _
        " ORDER BY A.ADJ_RQT_NUM DESC" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlTRISUserNumber
        .Update
        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update
        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    Set GetAdjustmentsStatusByScope = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenStatic _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetAdjustmentsStatusByScope_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsStatusByScope"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsStatusByScope" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsStatusByScope" _
                , , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetAdjustmentsDetailsByScope(ByVal vlTRISUserNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetAdjustmentsDetailsByScope_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

        lsSQL = "SELECT A.ADJ_RQT_NUM, COALESCE(C.ADJ_SEQ_NUM, 0) AS ADJ_SEQ_NUM" & _
            "       , LTRIM(RTRIM(B.EMP_NAME)) || ' [' || LTRIM(RTRIM(CHAR(A.EMP_NUM))) || ']' AS EMPLOYEE" & _
            "       , CASE WHEN (COALESCE(C.HOURS, 0) < 0) OR (COALESCE(C.AMOUNT, 0) < 0) THEN 'As paid'" & _
            "       ELSE 'Should have been paid' END AS ADJ_TYPE" & _
            "       , COALESCE(C.CHARGE_DATE, DATE('0100-01-01')) AS CHARGE_DATE, COALESCE(C.DATE_WORKED, DATE('0100-01-01')) AS DATE_WORKED" & _
            "       , LTRIM(RTRIM(COALESCE(K.DESCRIPTION, ''))) || ' [' || COALESCE(C.DISTRIBUTION_ID, '') || '] {' || COALESCE(C.DISTRIBUTION_TYPE, '') || '}' AS DISTRIBUTION" & _
            "       , COALESCE(C.HOURS, 0.0) AS HOURS, COALESCE(C.START_TIME, '') AS START_TIME, COALESCE(C.END_TIME, '') AS END_TIME" & _
            "       , LTRIM(RTRIM(COALESCE(D.EXCEPTION_DESC, ''))) || ' [' || COALESCE(C.EXCEPTION_CD, '') || ']' AS EXCEPTION" & _
            "       , LTRIM(RTRIM(COALESCE(E.EXCEPTION_DESC, ''))) || ' [' || COALESCE(C.OVER_EXC_CD, '') || ']' AS OVERRIDE" & _
            "       , COALESCE(C.EARNINGS_CD, '') AS EARNINGS_CD" & _
            "       , LTRIM(RTRIM(COALESCE(F.JOB_DESC, ''))) || ' [' || COALESCE(C.JOB_NUM, '') || ']' AS JOB_CLASS" & _
            "       , LTRIM(RTRIM(COALESCE(G.OT_REASON_DESC, ''))) || ' [' || COALESCE(C.OT_REASON_CD, '') || ']' AS OT_REASON" & _
            "       , COALESCE(C.OT_OPP_DESC, '') AS OT_OPP_DESC" & _
            "       , CASE WHEN (COALESCE(C.TRAVEL_PAY_HRS, 0.0) > 0.0) THEN 1" & _
            "       ELSE 0 END AS TRAVEL_PAY_IND" & _
            "       , CASE (B.EXEC_CONTRACT_IND) WHEN 'E' THEN 0.0" & _
            "       ELSE COALESCE(C.HOURLY_RATE, 0.0) END AS HOURLY_RATE" & _
            "       , CASE (B.EXEC_CONTRACT_IND) WHEN 'E' THEN CASE (SUBSTR(C.START_TIME, 1, 2))" & _
            "           WHEN '99' THEN COALESCE(C.AMOUNT, 0.0)" & _
            "           ELSE 0.0 END" & _
            "       ELSE COALESCE(C.AMOUNT, 0.0) END AS AMOUNT" & _
            "       , COALESCE(C.WG_NUM, 0) AS WG_NUM, COALESCE(C.CASE_NUM, 0) AS CASE_NUM"
        lsSQL = lsSQL & _
            "       , LTRIM(RTRIM(COALESCE(H.COMP_DESC, ''))) || ' [' || COALESCE(C.COMP_CD, '') || ']' AS COMP_OTHER" & _
            "       , CASE (B.EXEC_CONTRACT_IND)" & _
            "       WHEN 'E' THEN 0.0" & _
            "       ELSE COALESCE(C.COMP_AMT, 0.0) END AS COMP_AMT" & _
            "       , LTRIM(RTRIM(COALESCE(I.COMP_DESC, ''))) || ' [' || COALESCE(C.COMP_SHIFT_CD, '') || ']' AS COMP_SHIFT" & _
            "       , CASE (B.EXEC_CONTRACT_IND)" & _
            "       WHEN 'E' THEN 0.0" & _
            "       ELSE COALESCE(C.COMP_SHIFT_CD_AMT, 0.0) END AS COMP_SHIFT_CD_AMT" & _
            "       , LTRIM(RTRIM(COALESCE(J.COMP_DESC, ''))) || ' [' || COALESCE(C.REL_CD, '') || ']' AS COMP_RELIEF" & _
            "       , COALESCE(C.SCHED_CHG_EXC_CD, '') AS SCHED_CHG_EXC_CD, COALESCE(C.SCHED_CHG_EARN_CD, '') AS SCHED_CHG_EARN_CD" & _
            "       , COALESCE(C.START_LOC_CD, 0) AS START_LOC_CD, COALESCE(C.END_LOC_CD, 0) AS END_LOC_CD" & _
            "       , CASE COALESCE(C.CHEM_MTL_TSTR_IND, '') WHEN 'Y' THEN 1" & _
            "       ELSE 0 END AS CHEM_MTL_TSTR_IND" & _
            "       , CASE COALESCE(C.HALF_TRIP_IND, '') WHEN 'Y' THEN 1" & _
            "       ELSE 0 END AS HALF_TRIP_IND"
        lsSQL = lsSQL & _
            "       , CASE (COALESCE(C.APPR_IND, '')) WHEN 'Y' THEN 1" & _
            "       ELSE 0 END AS APPR_IND" & _
            "       , COALESCE(C.APPR_BY_EMP_NUM, 0) AS APPR_BY_EMP_NUM" & _
            "       , VALUE(NULLIF(C.APPR_DATE,'0001-01-01'), DATE('0100-01-01')) AS APPR_DATE" & _
            "       , COALESCE(C.OT_OPP_NUM, 0) AS OT_OPP_NUM" & _
            "       , VALUE(NULLIF(C.OT_BANK_CHRG_DT,'0001-01-01'), DATE('0100-01-01')) AS OT_BANK_CHRG_DT" & _
            "       , COALESCE(C.DISTR_DEPT_NUM, '') AS DISTR_DEPT_NUM, COALESCE(C.LOC_CD, 0) AS LOC_CD" & _
            "       , COALESCE(C.TRADE_EXC_CD, '') AS TRADE_EXC_CD" & _
            "       , CASE COALESCE(C.ADJ_PREV_PAID_FL, '') WHEN 'Y' THEN 1" & _
            "       ELSE 0 END AS ADJ_PREV_PAID_FL" & _
            "       , CASE COALESCE(C.PROCESSED_IND, '') WHEN 'Y' THEN 1" & _
            "       ELSE 0 END AS PROCESSED_IND" & _
            "       , COALESCE(C.CREATED_EMP_NUM, 0) AS CREATED_EMP_NUM, COALESCE(CHAR(C.CREATED_TIMESTAMP), '') AS CREATED_TIMESTAMP" & _
            "       , COALESCE(C.LAST_MOD_EMP_NUM, 0) AS LAST_MOD_EMP_NUM, COALESCE(CHAR(C.LAST_MOD_TIMESTAMP), '') AS LAST_MOD_TIMESTAMP"
        lsSQL = lsSQL & _
            " FROM  METHDBA.TTR_ADJ_RQT A" & _
            "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
            "       ON B.EMP_NUM = A.EMP_NUM" & _
            "       LEFT OUTER JOIN METHDBA.TTR_ADJ_TIME C" & _
            "       ON A.ADJ_RQT_NUM = C.ADJ_RQT_NUM" & _
            "       LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS D" & _
            "       ON D.EXCEPTION_CD = C.EXCEPTION_CD" & _
            "       LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS E" & _
            "       ON E.EXCEPTION_CD = C.OVER_EXC_CD" & _
            "       LEFT OUTER JOIN METHDBA.TTR_JOBS F" & _
            "       ON F.JOB_NUM = C.JOB_NUM" & _
            "       LEFT OUTER JOIN METHDBA.TTR_OT_REASONS G" & _
            "       ON G.OT_REASON_CD = C.OT_REASON_CD" & _
            "       LEFT OUTER JOIN METHDBA.TTR_COMP_CD H" & _
            "       ON H.COMP_CD = C.COMP_CD" & _
            "       LEFT OUTER JOIN METHDBA.TTR_COMP_CD I" & _
            "       ON I.COMP_CD = C.COMP_SHIFT_CD" & _
            "       LEFT OUTER JOIN METHDBA.TTR_COMP_CD J" & _
            "       ON J.COMP_CD = C.REL_CD" & _
            "       LEFT OUTER JOIN METHDBA.TTR_ACCT_ASSIGNMTS K" & _
            "       ON K.WG_NUM = C.WG_NUM" & _
            "       AND K.DISTRIBUTION_ID = C.DISTRIBUTION_ID"
        lsSQL = lsSQL & _
            " WHERE A.EMP_NUM IN (SELECT    EMP_NUM" & _
            "                   FROM        METHDBA.TTR_DS_EMP_DTL" & _
            "                   WHERE       EFF_EMP_NUM = ?)" & _
            " AND   DATE(A.CREATED_TIMESTAMP) BETWEEN ?" & _
            " AND   ?" & _
            " ORDER BY A.ADJ_RQT_NUM DESC, C.ADJ_SEQ_NUM" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlTRISUserNumber
        .Update
        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update
        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    Set GetAdjustmentsDetailsByScope = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenStatic _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetAdjustmentsDetailsByScope_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsDetailsByScope"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsDetailsByScope" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsDetailsByScope" _
                , , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
'End Prem TRIS 5.8

Public Function GetNotificationListForApproval(ByVal vlTRISUserNumber As Long) As ADODB.Recordset
On Error GoTo eh
    Dim SQL As String
    Dim oADOHelper As utilADOHelper.cuADOHelper
    Dim rs As ADODB.Recordset
    
    'get the SQL
    SQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetNotificationListForApproval.sql")
    
    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelper
    oADOHelper.ADO.CreateNumericParameter "EmpNum", adInteger, 4, 0, vlTRISUserNumber
    oADOHelper.ADO.CreateNumericParameter "EmpNum", adInteger, 4, 0, vlTRISUserNumber
    
    'execute sql
    Set rs = oADOHelper.ExecuteSelect(SQL, adOpenStatic, adLockBatchOptimistic)

    Set oADOHelper = Nothing
    Set GetNotificationListForApproval = rs
    Exit Function
eh:
    Set oADOHelper = Nothing
    Set rs = Nothing
    
    TrsRaiseError MODULENAME, "GetNotificationListForApproval", , , , DescArgsFrom(SQL, CStr(vlTRISUserNumber)), True
End Function

Private Sub Class_Initialize()
    Set m_ResourceMgr = New utilADOHelper.cuResourceMgr
End Sub

Private Sub Class_Terminate()
    Set m_ResourceMgr = Nothing
End Sub

