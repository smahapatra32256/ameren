VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdSchedException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cdSchedException"
Dim madrsExceptions As ADODB.Recordset
Dim madrsOverExcep As ADODB.Recordset
Dim madrsSchedChgExcep As ADODB.Recordset

'BEGIN PREM TRIS CONVERSION PHASE 3
Public Function GetAdjustments( _
                ByVal v_aryEmployees As Variant _
                , ByVal v_datStartDate As Date _
                , ByVal v_datEndDate As Date _
                ) As ADODB.Recordset
On Error GoTo GetAdjustments_EH
Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
Dim ladcnCN As ADODB.Connection
Dim ladrs As ADODB.Recordset
Dim lsSQL As String, lsWHEREClause As String

    'Prepare Where Clause
    If (IsArray(v_aryEmployees)) Then
        lsWHEREClause = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        lsWHEREClause = "A.EMP_NUM = " & CStr(CLng(v_aryEmployees)) & " "
    End If

    If (v_datStartDate < #1/27/2005#) Then
        'Get from old and new tables
        lsSQL = "SELECT     A.EMP_NUM" & _
            "       , VALUE(NULLIF(B.CHARGE_DATE,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , CASE WHEN (B.HOURS) > 0 THEN 'S'" & _
            "       ELSE 'A' END AS ADJ_TYPE" & _
            "       , B.EXCEPTION_CD" & _
            "       , C.CASE_GROUP" & _
            "       , LEFT(C.EXCEPTION_BANK_CD, 1) AS BANK_CD" & _
            "       , 'N' AS FULL_SHIFT_IND" & _
            "       , SUM(B.HOURS) AS HOURS" & _
            "       , C.EXCEPTION_TYPE" & _
            "       , C.SYMBOL_CD" & _
            "       , D.EG_SUB_CD" & _
            "       , B.OVER_EXC_CD " & _
            " FROM  METHDBA.TTR_ADJ_RQT A" & _
            "       , METHDBA.TTR_ADJ_TIME B" & _
            "       , METHDBA.TTR_EXCEPTION_CDS C" & _
            "       LEFT OUTER JOIN METHDBA.TTR_EXC_GROUP D" & _
            "       ON C.EXCEPTION_CD = D.EXCEPTION_CD"
        lsSQL = lsSQL & _
            " WHERE " & lsWHEREClause & _
            " AND   A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND   B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND   '" & CStr(v_datEndDate) & "'" & _
            " AND   B.HOURS <> 0" & _
            " AND   B.EXCEPTION_CD = C.EXCEPTION_CD" & _
            " AND   D.EG_CD = 2" & _
            " AND   D.EXC_SEQ_NUM = 1" & _
            " GROUP BY A.EMP_NUM, B.CHARGE_DATE" & _
            "       , B.HOURS, B.EXCEPTION_CD" & _
            "       , C.CASE_GROUP, C.EXCEPTION_BANK_CD" & _
            "       , C.EXCEPTION_TYPE, C.SYMBOL_CD" & _
            "       , D.EG_SUB_CD" & _
            "       , B.OVER_EXC_CD "

        lsSQL = lsSQL & " UNION ALL " & _
            " SELECT    A.EMP_NUM" & _
            "       , VALUE(NULLIF(A.DATE_WORKED,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , A.ADJ_TYPE" & _
            "       , A.EXCEPTION_CD" & _
            "       , B.CASE_GROUP" & _
            "       , LEFT(B.EXCEPTION_BANK_CD, 1) AS BANK_CD" & _
            "       , 'N' AS FULL_SHIFT_IND" & _
            "       , SUM(A.HOURS) AS HOURS" & _
            "       , B.EXCEPTION_TYPE" & _
            "       , B.SYMBOL_CD" & _
            "       , C.EG_SUB_CD" & _
            "       , space(3) as OVER_EXC_CD " & _
            " FROM  METHDBA.TTR_ADJUSTMENTS A" & _
            "       , METHDBA.TTR_EXCEPTION_CDS B" & _
            "       LEFT OUTER JOIN METHDBA.TTR_EXC_GROUP C" & _
            "       ON B.EXCEPTION_CD = C.EXCEPTION_CD"
        lsSQL = lsSQL & _
            " WHERE " & lsWHEREClause & _
            " AND   A.DATE_WORKED BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND   '" & CStr(v_datEndDate) & "'" & _
            " AND   A.HOURS <> 0" & _
            " AND   A.ADJ_TYPE IN ('A','S')" & _
            " AND   B.EXCEPTION_CD = A.EXCEPTION_CD" & _
            " AND   C.EG_CD = 2" & _
            " AND   C.EXC_SEQ_NUM = 1" & _
            " GROUP BY A.EMP_NUM, DATE_WORKED" & _
            "       , A.ADJ_TYPE, A.EXCEPTION_CD" & _
            "       , B.CASE_GROUP, B.EXCEPTION_BANK_CD" & _
            "       , B.EXCEPTION_TYPE, B.SYMBOL_CD" & _
            "       , C.EG_SUB_CD"
    Else
        'get from new table
        lsSQL = "SELECT     A.EMP_NUM" & _
            "       , VALUE(NULLIF(B.CHARGE_DATE,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , CASE WHEN (B.HOURS) > 0 THEN 'S'" & _
            "       ELSE 'A' END AS ADJ_TYPE" & _
            "       , B.EXCEPTION_CD" & _
            "       , C.CASE_GROUP" & _
            "       , LEFT(C.EXCEPTION_BANK_CD, 1) AS BANK_CD" & _
            "       , 'N' AS FULL_SHIFT_IND" & _
            "       , SUM(B.HOURS) AS HOURS" & _
            "       , C.EXCEPTION_TYPE" & _
            "       , C.SYMBOL_CD" & _
            "       , D.EG_SUB_CD" & _
            "       , B.OVER_EXC_CD " & _
            " FROM  METHDBA.TTR_ADJ_RQT A" & _
            "       , METHDBA.TTR_ADJ_TIME B" & _
            "       , METHDBA.TTR_EXCEPTION_CDS C" & _
            "       LEFT OUTER JOIN METHDBA.TTR_EXC_GROUP D" & _
            "       ON C.EXCEPTION_CD = D.EXCEPTION_CD"
        lsSQL = lsSQL & _
            " WHERE " & lsWHEREClause & _
            " AND   A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND   B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND   '" & CStr(v_datEndDate) & "'" & _
            " AND   B.HOURS <> 0" & _
            " AND   B.EXCEPTION_CD = C.EXCEPTION_CD" & _
            " AND   D.EG_CD = 2" & _
            " AND   D.EXC_SEQ_NUM = 1" & _
            " GROUP BY A.EMP_NUM, B.CHARGE_DATE" & _
            "       , B.HOURS, B.EXCEPTION_CD" & _
            "       , C.CASE_GROUP, C.EXCEPTION_BANK_CD" & _
            "       , C.EXCEPTION_TYPE, C.SYMBOL_CD" & _
            "       , D.EG_SUB_CD " & _
            "       , B.OVER_EXC_CD "
    End If

    lsSQL = lsSQL & _
        " ORDER BY EMP_NUM, ADJ_TYPE, CHARGE_DATE"

    'Create Connection
    Set ladcnCN = New ADODB.Connection

    'Get Connection String
    Set lobjCM = New clsuDBConnectMgr
    ladcnCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
    Set lobjCM = Nothing

    'Open Connection
    ladcnCN.Open

    'Create RS
    Set ladrs = New ADODB.Recordset
    ladrs.CursorLocation = adUseClient
    'Open (Execute)
    ladrs.Open lsSQL, ladcnCN _
        , adOpenStatic, adLockReadOnly
    ladrs.ActiveConnection = Nothing

    'Close and Destroy
    ladcnCN.Close
    Set ladcnCN = Nothing

    'Set Return
    Set GetAdjustments = ladrs

    Set ladrs = Nothing

    Exit Function
GetAdjustments_EH:
    Set lobjCM = Nothing
    Set ladcnCN = Nothing
    Set ladrs = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetAdjustments" _
        , , , , DescArgsFrom(lsSQL) _
        , True

End Function 'END PREM TRIS CONVERSION PHASE 3
Public Function GetExceptions(ByVal v_aryEmployees As Variant _
                , ByVal v_datStartDate As Date _
                , ByVal v_datEndDate As Date _
                , Optional v_datLastClosedPP As Date = CDate("1980-01-01") _
                , Optional v_datEndOfOpenPP As Date = CDate("1980-01-01") _
                ) As ADODB.Recordset
On Error GoTo eh
Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
Dim lobjCN As ADODB.Connection
Dim lobjRS As ADODB.Recordset
Dim lsSQL As String
Dim lsEmployeeSelection As String
Dim ldtHistoryEndDate As Date
Dim ldtActualTimeStartDate As Date
Dim ldtActualTimeEndDate As Date
Dim ldtSchedExcStartDate As Date
Dim lsInClausePartValueKey As String
Dim liActualTimeTables As TrActualTimeTable
Dim l_strEmpSchedSQL As String
Dim l_strSchedOTSQL As String
Dim l_strActTimeSQL As String
Dim l_strSchedOTandActTimeSQL As String
Dim l_strSumSchedOTandActTimeSQL As String


    If (IsArray(v_aryEmployees)) Then
        lsEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        lsEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    liActualTimeTables = DetermineActualTimeTables(v_datStartDate _
                            , v_datEndDate _
                            , ldtHistoryEndDate _
                            , ldtActualTimeStartDate)
    If (v_datLastClosedPP > ldtActualTimeStartDate) Then
        ldtSchedExcStartDate = DateAdd("d", 1, v_datLastClosedPP)
    Else
        ldtSchedExcStartDate = ldtActualTimeStartDate
    End If
    If (v_datEndOfOpenPP > ldtActualTimeStartDate) Then
        ldtActualTimeEndDate = v_datEndOfOpenPP
    Else
        ldtActualTimeEndDate = v_datEndDate
    End If
    
    l_strEmpSchedSQL = GetExceptionsEmpSchedSQL( _
        lsEmployeeSelection, ldtSchedExcStartDate, v_datEndDate)
        
    l_strSchedOTSQL = GetExceptionsSchedOTSQL(lsEmployeeSelection, ldtSchedExcStartDate, v_datEndDate)
    Select Case liActualTimeTables
        Case TrTable_ActualAndHistory
            l_strActTimeSQL = GetExceptionsActTimeSQL(TrTable_Actual, lsEmployeeSelection, ldtActualTimeStartDate, ldtActualTimeEndDate) _
                & " UNION " & GetExceptionsActTimeSQL(TrTable_History, lsEmployeeSelection, v_datStartDate, ldtHistoryEndDate)
        Case TrTable_Actual
            l_strActTimeSQL = GetExceptionsActTimeSQL(TrTable_Actual, lsEmployeeSelection, ldtActualTimeStartDate, ldtActualTimeEndDate)
        Case TrTable_History
            l_strActTimeSQL = GetExceptionsActTimeSQL(TrTable_History, lsEmployeeSelection, v_datStartDate, ldtHistoryEndDate)
        Case Else
            l_strActTimeSQL = ""
    End Select

    If l_strActTimeSQL = "" Then
        l_strSchedOTandActTimeSQL = l_strSchedOTSQL
    Else
        l_strSchedOTandActTimeSQL = l_strSchedOTSQL & " UNION " & l_strActTimeSQL
    End If
    
    l_strSumSchedOTandActTimeSQL = _
            "SELECT NE1.emp_num, NE1.charge_date, NE1.exception_cd, NE1.over_exc_cd," _
            & "NE1.case_group, NE1.bank_cd, NE1.FULL_SHIFT_EXC_IND,SUM(NE1.hours) AS hours, " _
            & "NE1.shift_num, NE1.symbol_cd, NE1.shift_code, NE1.sched_change_pmt, NE1.eg_sub_cd " _
            & "from  ( " & l_strSchedOTandActTimeSQL _
            & " ) as NE1 " _
            & " GROUP BY NE1.emp_num, NE1.charge_date, NE1.exception_cd,NE1.over_exc_cd, " _
            & "NE1.case_group, NE1.bank_cd, NE1.FULL_SHIFT_EXC_IND, " _
            & "NE1.shift_num, NE1.symbol_cd, NE1.shift_code, NE1.sched_change_pmt, NE1.eg_sub_cd "
            
    lsSQL = l_strEmpSchedSQL _
        & " UNION ALL " & l_strSumSchedOTandActTimeSQL _
        & " ORDER BY EMP_NUM,CHARGE_DATE"
        
    Set lobjCN = New ADODB.Connection
    Set lobjCM = New clsuDBConnectMgr
    lobjCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
    Set lobjCM = Nothing
    lobjCN.Open

    Set lobjRS = New ADODB.Recordset
    lobjRS.CursorLocation = adUseClient
    lobjRS.Open lsSQL, lobjCN, adOpenStatic, adLockBatchOptimistic
    lobjRS.ActiveConnection = Nothing

    lobjCN.Close
    Set lobjCN = Nothing

    Set GetExceptions = lobjRS
    Set lobjRS = Nothing

    Exit Function
eh:
    Set lobjCM = Nothing
    Set lobjCN = Nothing
    Set lobjRS = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetExceptions" _
        , , , , DescArgsFrom(lsSQL) _
        , True

End Function
Public Function GetScheduledExceptions( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date _
        ) As ADODB.Recordset
On Error GoTo eh
Dim l_objCN As ADODB.Connection
Dim l_objRS As ADODB.Recordset
Dim l_strSQL As String
Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
Dim l_strEmployeeSelection As String

    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    l_strSQL = "select EMP_NUM, CHARGE_DATE, SHIFT_NUM, " _
        & "HOURS_OFFSET_NUM, HOURS, " _
        & "EXCEPTION_CD, OVER_EXC_CD, FMLA_IND, " _
        & "CASE_NUM, WG_NUM, JOB_NUM, RELIEF_COMP, " _
        & "COMP_CD, SHIFT_COMP, SCHED_CD, SCHED_SRC_CD, " _
        & "CREATED_EMP_NUM, LAST_MOD_EMP_NUM,  " _
        & "CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, " _
        & "CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP " _
        & "From METHDBA.TTR_EMP_SCHED_EXC  " _
        & "where " & l_strEmployeeSelection _
        & "  and CHARGE_DATE >= " & DateAsEnclosedString(v_datStartDate) _
        & "  and CHARGE_DATE <= " & DateAsEnclosedString(v_datEndDate) _

    l_strSQL = l_strSQL & " order by EMP_NUM,CHARGE_DATE"

    Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing

    l_objCN.Open

    Set l_objRS = New ADODB.Recordset
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN _
        , adOpenStatic, adLockBatchOptimistic
    l_objRS.ActiveConnection = Nothing

    l_objCN.Close
    Set l_objCN = Nothing

    Set GetScheduledExceptions = l_objRS
    Set l_objRS = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduledExceptions", , , , DescArgsFrom(l_strSQL), True

End Function
Public Function SaveExceptions( _
                ByVal v_rsExceptions As ADODB.Recordset _
                , ByVal v_lngUserNum As Long _
                ) As Boolean
                
TrsRaiseError m_strCMODULENAME, "This method has been moved to cdSchedExceptionTrans.", , , , , True

End Function
Public Function GetActualTimeSchedChangePay( _
                ByVal v_aryEmployees As Variant _
                , ByVal v_datStartDate As Date _
                , ByVal v_datEndDate As Date _
                , Optional v_datLastClosedPP As Date = CDate("1980-01-01") _
                , Optional v_datEndOfOpenPP As Date = CDate("1980-01-01") _
                ) As ADODB.Recordset
On Error GoTo eh
Dim l_objCN As ADODB.Connection
Dim l_objRS As ADODB.Recordset
Dim l_strSQL As String
Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
Dim l_strEmployeeSelection As String
Dim l_datHistoryEndDate As Date
Dim l_datActualTimeStartDate As Date
Dim l_datActualTimeEndDate As Date
Dim l_datSchedExcStartDate As Date
Dim l_strInClausePartValueKey As String
Dim l_lngActualTimeTables As Long

    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    
    l_lngActualTimeTables = DetermineActualTimeTables(v_datStartDate, v_datEndDate, l_datHistoryEndDate, l_datActualTimeStartDate)
    
    If v_datLastClosedPP > l_datActualTimeStartDate Then
        l_datSchedExcStartDate = DateAdd("d", 1, v_datLastClosedPP)
    Else
        l_datSchedExcStartDate = l_datActualTimeStartDate
    End If
    If v_datEndOfOpenPP > l_datActualTimeStartDate Then
        l_datActualTimeEndDate = v_datEndOfOpenPP
    Else
        l_datActualTimeEndDate = v_datEndDate
    End If

    If l_lngActualTimeTables = TrTable_Actual Or _
        l_lngActualTimeTables = TrTable_ActualAndHistory Then
        l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, l_datHistoryEndDate), True)
        l_strSQL = "select EMP_NUM, CHARGE_DATE , SCHED_CD , " _
            & " SCHED_CHG_EXC_CD,sum(HOURS) as hours, 'SCP' as RECORD_TYPE " _
            & " from  METHDBA.TTR_ACTUAL_TIME  " _
            & " where " & l_strEmployeeSelection _
            & "  and CHARGE_DATE >= " & DateAsEnclosedString(l_datActualTimeStartDate) _
            & "  and CHARGE_DATE <= " & DateAsEnclosedString(l_datActualTimeEndDate) _
            & "  and SCHED_CHG_EXC_CD <> '' " _
            & " group by emp_num, charge_date, sched_cd, sched_chg_exc_cd "
    End If
    
    If l_lngActualTimeTables = TrTable_History Or _
       l_lngActualTimeTables = TrTable_ActualAndHistory Then
       l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, l_datHistoryEndDate), True)
       If Trim(l_strSQL) <> "" Then
            l_strSQL = l_strSQL & " UNION ALL " _
                 & "select EMP_NUM, CHARGE_DATE , SCHED_CD , " _
                 & " SCHED_CHG_EXC_CD,sum(HOURS) as hours,'SCP' as RECORD_TYPE " _
                 & " from  METHDBA.TTR_ACTUAL_TIME_HS  " _
                 & " where " & l_strEmployeeSelection _
                 & "  and CHARGE_DATE >= " & DateAsEnclosedString(v_datStartDate) _
                 & "  and CHARGE_DATE <= " & DateAsEnclosedString(l_datHistoryEndDate) _
                 & "  and SCHED_CHG_EXC_CD <> '' " _
                 & " group by emp_num, charge_date, sched_cd, sched_chg_exc_cd "
        Else
            l_strSQL = "select EMP_NUM, CHARGE_DATE , SCHED_CD , " _
                 & " SCHED_CHG_EXC_CD,sum(HOURS) as hours,'SCP' as RECORD_TYPE " _
                 & " from  METHDBA.TTR_ACTUAL_TIME_HS  " _
                 & " where " & l_strEmployeeSelection _
                 & "  and CHARGE_DATE >= " & DateAsEnclosedString(v_datStartDate) _
                 & "  and CHARGE_DATE <= " & DateAsEnclosedString(l_datHistoryEndDate) _
                 & "  and SCHED_CHG_EXC_CD <> '' " _
                 & " group by emp_num, charge_date, sched_cd, sched_chg_exc_cd "
        End If
    End If
    
    l_strSQL = l_strSQL & " order by EMP_NUM,CHARGE_DATE"

    Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing

    l_objCN.Open
    
    Set l_objRS = New ADODB.Recordset
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    l_objRS.ActiveConnection = Nothing

    l_objCN.Close
    Set l_objCN = Nothing

    Set GetActualTimeSchedChangePay = l_objRS
    Set l_objRS = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetActualTimeSchedChangePay", , , , DescArgsFrom(l_strSQL), True

End Function

'--sql 1 scheduled exceptions (exception cd or override exception cd)
'--sql 2, 3, 4, 5 old adjustments (2,3,4 TTR_ADJ_RQT) (5 TTR_ADJUSTMENTS (old)) (If (v_datStartDate < #1/27/2005#) Then)
'--sql 10, 11, 12 new adjustments (Else) (10, 11, 12 TTR_ADJ_RQT)
'--sql 2, 5, 10 & use adjustments & exception cd
'--sql 3, 11 & use adjustments & override exception cd
'--sql 4, 12 & use adjustments & schedule change exception cd

'--sql 16 - use actual time & exception cd
'--sql 17 - use actual time & override exception cd
'--sql 18 - use actual time & schedule change exception cd
'--sql 19 - use actual time history & exception cd
'--sql 20 - use actual time history & override exception cd
'--sql 21 - use actual time history & schedule change exception cd
'--lsSQLALL (after it is all 'UNION'ed together) Group by ... including CASE_GROUP

Public Function GetEmpExceptions( _
                ByVal v_aryEmployees As Variant _
                , ByVal v_datStartDate As Date _
                , ByVal v_datEndDate As Date _
                , ByVal v_adrsSelectedExceptions As ADODB.Recordset _
                , Optional v_datLastClosedPP As Date = CDate("1980-01-01") _
                , Optional v_datEndOfOpenPP As Date = CDate("1980-01-01") _
                ) As ADODB.Recordset
On Error GoTo eh
Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
Dim ladCN As ADODB.Connection
Dim ladrs As ADODB.Recordset
Dim lsEmpSelection As String
Dim ldtHistoryEndDate As Date, ldtSchedExcStartDate As Date
Dim ldtActualTimeStartDate As Date, ldtActualTimeEndDate As Date
Dim lsInClausePartValueKey As String, lsExcepInClause As String
Dim lsOvrExcepInClause As String
Dim lsSchedChgExcepInClause As String
Dim lsAllExcepInClause As String
Dim lsAllOvrExcepInClause As String
Dim lsAllSchedChgExcepInClause As String
Dim lsAllExcepandOvrInClause As String
Dim lsExcepandOvrInClause As String
Dim liActualTimeTables As TrActualTimeTable, liRecCounter As Integer
Dim lblnValidExcep As Boolean
Dim lblnFirstExcep As Boolean
Dim lsSQL1 As String
Dim lsSQL2 As String
Dim lsSQL3 As String
Dim lsSQL4 As String
Dim lsSQL5 As String
Dim lsSQL10 As String
Dim lsSQL11 As String
Dim lsSQL12 As String
Dim lsSQL16 As String
Dim lsSQL17 As String
Dim lsSQL18 As String
Dim lsSQL19 As String
Dim lsSQL20 As String
Dim lsSQL21 As String
Dim lsSQLALL As String
Dim lsSQL As String
Dim lblnUnion As Boolean

    If (IsArray(v_aryEmployees)) Then
        lsEmpSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        lsEmpSelection = "A.EMP_NUM = " & CStr(CLng(v_aryEmployees)) & " "
    End If

    liActualTimeTables = DetermineActualTimeTables(v_datStartDate _
                            , v_datEndDate _
                            , ldtHistoryEndDate _
                            , ldtActualTimeStartDate)
                            
    If v_datLastClosedPP = #1/1/1980# Then
        ldtSchedExcStartDate = v_datStartDate
    Else
        If (v_datLastClosedPP > ldtActualTimeStartDate) Then
            ldtSchedExcStartDate = DateAdd("d", 1, v_datLastClosedPP)
        Else
            ldtSchedExcStartDate = ldtActualTimeStartDate
        End If
    End If

    If v_datEndOfOpenPP = #1/1/1980# Then
        ldtActualTimeEndDate = v_datEndDate
    Else
        If (v_datEndOfOpenPP > ldtActualTimeStartDate) Then
            ldtActualTimeEndDate = v_datEndOfOpenPP
        Else
            ldtActualTimeEndDate = v_datEndDate
        End If
    End If

    GetExceptionCodes
    GetOverExcepCodes
    GetSchedChgExcepCodes
    
    If (RecordsExist(v_adrsSelectedExceptions)) Then
        
        With v_adrsSelectedExceptions
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsExceptionCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsExcepInClause = lsExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsExcepInClause = lsExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If

                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsExcepInClause = " A.EXCEPTION_CD IN ( " & lsExcepInClause & ")"
            End If
            
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsOverExcepCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsOvrExcepInClause = lsOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsOvrExcepInClause = lsOvrExcepInClause & _
                            ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If
                
                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsOvrExcepInClause & _
                ")"
            End If
            
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsSchedChgExcepCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
                        "'" & !EXCEPTION_CD & "'"
                    Else
                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If
                
                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
                lsSchedChgExcepInClause & ")"
            End If
            
        End With
        
    Else
    
        If RecordsExist(madrsExceptions) Then
            With madrsExceptions
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllExcepInClause = lsAllExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllExcepInClause = lsAllExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
                
                lsAllExcepInClause = " A.EXCEPTION_CD IN ( " & lsAllExcepInClause & ")"
            End With
        End If
        
        If RecordsExist(madrsOverExcep) Then
            With madrsOverExcep
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & _
                            ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
                
                lsAllOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsAllOvrExcepInClause & ")"
            End With
        End If
        
        If RecordsExist(madrsSchedChgExcep) Then
            With madrsSchedChgExcep
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
                        "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
            
                lsAllSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
                lsAllSchedChgExcepInClause & ")"
            End With
        End If
        
    End If

'*******************************************************
'*  Retrieve Scheduled Exceptions
'*******************************************************

    lsSQL1 = lsSQL1 & "SELECT a.emp_num, d.emp_name,  a.job_num, a.charge_date, " _
        & "'00:00' as start_time, '00:00' as end_time, a.exception_cd, " _
        & "a.over_exc_cd, '' as sched_chg_exc_cd, a.case_num, " _
        & " 'C' as adj_type, '1' AS processed_ind, '' as OT_OPP_DESC, " _
        & " '' as distr_dept_num, 0 as amount, 0 as travel_pay_hrs, hours " _
        & "FROM methdba.ttr_emp_sched_exc A, " _
        & " methdba.ttr_exception_cds B, " _
        & " methdba.ttr_employee D " _
        & " WHERE " & lsEmpSelection _
        & " AND A.charge_date between " & DateAsEnclosedString(ldtSchedExcStartDate) & " AND " _
                & DateAsEnclosedString(v_datEndDate) _
        & " AND A.emp_num = D.emp_num " _
        & " AND B.exception_cd = A.exception_cd "
        
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) > 0 Then
                lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & " OR " & _
                lsOvrExcepInClause & ") "
            Else
                If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) = 0 Then
                    lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & ") "
                Else
                    If Len(lsExcepInClause) = 0 And Len(lsOvrExcepInClause) > 0 Then
                        lsSQL1 = lsSQL1 & " AND (" & lsOvrExcepInClause & ") "
                    Else
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                        " OR A.OVER_EXC_CD <> '') "
                    End If
                End If
            End If
        Else
            If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) > 0 Then
                lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & " OR " & _
                    lsAllOvrExcepInClause & ") "
                lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
            Else
                If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) = 0 Then
                    lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & ") "
                    lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                Else
                    If Len(lsAllExcepInClause) = 0 And Len(lsAllOvrExcepInClause) > 0 Then
                        lsSQL1 = lsSQL1 & " AND (" & lsAllOvrExcepInClause & ") "
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                    Else
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                    End If
                End If
            End If
        End If

        lsSQL1 = lsSQL1 _
        & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
        & " AND A.exception_cd NOT IN ( SELECT exception_cd  " _
                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5 ) " _
        & " AND A.over_exc_cd NOT IN ( SELECT exception_cd  " _
                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5) "
                    
        lsSQL1 = lsSQL1 & " UNION "

    If (v_datStartDate < #1/27/2005#) Then
        'GET INFO FROM OLD AND NEW TABLES
        'FOR SHOULD HAVE BEEN & AS PAID
        lsSQL2 = lsSQL2 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'S' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
        
        lsSQL2 = lsSQL2 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMPLOYEE E"
       
       lsSQL2 = lsSQL2 & _
            " WHERE " & lsEmpSelection & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND B.EXCEPTION_CD = C.EXCEPTION_CD"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsExcepInClause) > 0) Then
                lsSQL2 = lsSQL2 & _
                    " AND " & Replace(lsExcepInClause, "A.", "B.")
            Else
                lsSQL2 = lsSQL2 & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
            End If
        Else
            lsSQL2 = lsSQL2 & _
                " AND " & Replace(lsAllExcepInClause, "A.", "B.")
                lsSQL2 = lsSQL2 & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
        End If
        
        lsSQL2 = lsSQL2 & _
            " AND B.EXCEPTION_CD <> ''" & _
            " AND B.HOURS <> 0"

        lsSQL2 = lsSQL2 & " UNION "
        
        lsSQL3 = lsSQL3 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
                    
        lsSQL3 = lsSQL3 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMPLOYEE E"
            
        lsSQL3 = lsSQL3 & _
            " WHERE " & lsEmpSelection & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND B.OVER_EXC_CD = C.EXCEPTION_CD"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsOvrExcepInClause) > 0) Then
                lsSQL3 = lsSQL3 & _
                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL3 = lsSQL3 & _
                " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
        End If
        
        lsSQL3 = lsSQL3 & _
            " AND B.OVER_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL3 = lsSQL3 & " UNION "
        
        lsSQL4 = lsSQL4 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL4 = lsSQL4 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMPLOYEE E"
            
        lsSQL4 = lsSQL4 & _
            " WHERE " & lsEmpSelection & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD"
        
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsSchedChgExcepInClause) > 0) Then
                lsSQL4 = lsSQL4 & _
                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL4 = lsSQL4 & _
                " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
        End If
        
        lsSQL4 = lsSQL4 & _
            " AND B.SCHED_CHG_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL4 = lsSQL4 & " UNION "

        lsSQL5 = lsSQL5 & " SELECT A.EMP_NUM, D.EMP_NAME" & _
            " , A.JOB_NUM, A.DATE_WORKED AS CHARGE_DATE" & _
            " , A.START_TIME, A.END_TIME" & _
            " , A.EXCEPTION_CD, '' AS OVER_EXC_CD" & _
            " , '' AS SCHED_CHG_EXC_CD" & _
            " , A.CASE_NUM" & _
            " , A.ADJ_TYPE" & _
            " , CASE WHEN (A.PROCESSED_IND) = 'Y' THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , '' AS OT_OPP_DESC, '' AS DISTR_DEPT_NUM" & _
            " , 0 AS AMOUNT, 0 AS TRAVEL_PAY_HRS" & _
            " , CASE WHEN ADJ_TYPE = 'A' THEN (A.HOURS * -1) ELSE A.HOURS END AS HOURS" & _
            " FROM METHDBA.TTR_ADJUSTMENTS A" & _
            " , METHDBA.TTR_EXCEPTION_CDS B" & _
            " , METHDBA.TTR_EMPLOYEE D"
        lsSQL5 = lsSQL5 & _
            " WHERE " & lsEmpSelection & _
            " AND A.EMP_NUM = D.EMP_NUM" & _
            " AND A.DATE_WORKED BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.HOURS <> 0" & _
            " AND A.EXCEPTION_CD = B.EXCEPTION_CD"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsExcepInClause) > 0) Then
                lsSQL5 = lsSQL5 & _
                    " AND " & lsExcepInClause
            Else
                lsSQL5 = lsSQL5 & _
                    " AND A.EXCEPTION_CD <> 'ST'"
            End If
        Else
            lsSQL5 = lsSQL5 & _
                " AND " & lsAllExcepInClause
            lsSQL5 = lsSQL5 & _
                    " AND A.EXCEPTION_CD <> 'ST'"
        End If
        
        lsSQL5 = lsSQL5 & _
            " AND A.EXCEPTION_CD NOT IN ('ET','AT')" & _
            " AND A.EXCEPTION_CD <> ''"
            
        lsSQL5 = lsSQL5 & " UNION "

    Else
        'GET ADJ FROM NEW TABLE FOR SHOULD HAVE BEEN & AS PAID
        lsSQL10 = lsSQL10 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL10 = lsSQL10 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMPLOYEE E"
                    
        lsSQL10 = lsSQL10 & _
            " WHERE " & lsEmpSelection & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND B.EXCEPTION_CD = C.EXCEPTION_CD"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsExcepInClause) > 0) Then
                lsSQL10 = lsSQL10 & _
                    " AND " & Replace(lsExcepInClause, "A.", "B.")
            Else
                lsSQL10 = lsSQL10 & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
            End If
        Else
            lsSQL10 = lsSQL10 & _
                " AND " & Replace(lsAllExcepInClause, "A.", "B.")
            lsSQL10 = lsSQL10 & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
        End If
        
        lsSQL10 = lsSQL10 & _
            " AND B.EXCEPTION_CD <> ''" & _
            " AND B.HOURS <> 0"

        lsSQL10 = lsSQL10 & " UNION "

        lsSQL11 = lsSQL11 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL11 = lsSQL11 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMPLOYEE E"
                    
        lsSQL11 = lsSQL11 & _
            " WHERE " & lsEmpSelection & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND B.OVER_EXC_CD = C.EXCEPTION_CD"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsOvrExcepInClause) > 0) Then
                lsSQL11 = lsSQL11 & _
                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL11 = lsSQL11 & _
                " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
        End If
        
        lsSQL11 = lsSQL11 & _
            " AND B.OVER_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL11 = lsSQL11 & " UNION "

        lsSQL12 = lsSQL12 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL12 = lsSQL12 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMPLOYEE E"
            
        lsSQL12 = lsSQL12 & _
            " WHERE " & lsEmpSelection & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsSchedChgExcepInClause) > 0) Then
                lsSQL12 = lsSQL12 & _
                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL12 = lsSQL12 & _
                " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
        End If
        
        lsSQL12 = lsSQL12 & _
            " AND B.SCHED_CHG_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL12 = lsSQL12 & " UNION "

    End If
'END PREM TRIS CONVERSION PHASE 3

'*******************************************************
'*  Retrieve Exceptions from Actual Time
'*******************************************************
    If liActualTimeTables = TrTable_Actual Or _
        liActualTimeTables = TrTable_ActualAndHistory Then
        lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
        lsSQL16 = lsSQL16 _
            & "SELECT A.emp_num, d.emp_name, A.job_num, A.charge_date, A.start_time, " _
            & "A.end_time, A.exception_cd, A.over_exc_cd, A.sched_chg_exc_cd, " _
            & "A.case_num, ' ' as adj_type," _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM," _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL16 = lsSQL16 & "WHERE " & lsEmpSelection _
            & " AND A.charge_date between " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.exception_cd "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsExcepInClause) > 0 Then
                    lsSQL16 = lsSQL16 & " AND " & lsExcepInClause
                Else
                    lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
                End If
            Else
                lsSQL16 = lsSQL16 & " AND " & lsAllExcepInClause _
                & " AND A.EXCEPTION_CD <> 'ST'" _
                & " AND A.OVER_EXC_CD = ''" _
                & " AND A.SCHED_CHG_EXC_CD = ''"
            End If
            
        lsSQL16 = lsSQL16 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.EXCEPTION_CD <> ''"
            
        lsSQL16 = lsSQL16 & " UNION "

        lsSQL17 = lsSQL17 _
            & "SELECT A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
            & " A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, " _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM," _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL17 = lsSQL17 & "WHERE " & lsEmpSelection _
            & " AND A.charge_date between " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.OVER_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsOvrExcepInClause) > 0 Then
                    lsSQL17 = lsSQL17 & " AND " & lsOvrExcepInClause
                End If
            Else
                lsSQL17 = lsSQL17 & " AND " & lsAllOvrExcepInClause
            End If
            
        lsSQL17 = lsSQL17 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.OVER_EXC_CD <> ''"
            
        lsSQL17 = lsSQL17 & " UNION "

        lsSQL18 = lsSQL18 _
            & "SELECT A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
            & " A.start_time, A.end_time, A.exception_cd, " _
            & " A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & " ' ' as adj_type, " _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL18 = lsSQL18 & "WHERE " & lsEmpSelection _
            & " AND A.charge_date between " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsSchedChgExcepInClause) > 0 Then
                    lsSQL18 = lsSQL18 & " AND " & lsSchedChgExcepInClause
                End If
            Else
                lsSQL18 = lsSQL18 & " AND " & lsAllSchedChgExcepInClause
            End If
            
        lsSQL18 = lsSQL18 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.SCHED_CHG_EXC_CD <> ''"
            
        lsSQL18 = lsSQL18 & " UNION "
        
    End If
    
'*******************************************************
'*  Retrieve Exceptions from Actual Time History
'*******************************************************
    If liActualTimeTables = TrTable_History Or _
       liActualTimeTables = TrTable_ActualAndHistory Then
       lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
       lsSQL19 = lsSQL19 _
            & "SELECT A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
            & "A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, '3' AS processed_ind, " _
            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee D "
            
        lsSQL19 = lsSQL19 & " WHERE " & lsEmpSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.exception_cd "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsExcepInClause) > 0 Then
                    lsSQL19 = lsSQL19 & " AND " & lsExcepInClause
                Else
                    lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
                End If
            Else
                lsSQL19 = lsSQL19 & " AND " & lsAllExcepInClause _
                & " AND A.EXCEPTION_CD <> 'ST'" _
                & " AND A.OVER_EXC_CD = ''" _
                & " AND A.SCHED_CHG_EXC_CD = ''"
            End If
        
        lsSQL19 = lsSQL19 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.EXCEPTION_CD <> ''"
            
        lsSQL19 = lsSQL19 & " UNION "

        lsSQL20 = lsSQL20 _
            & "SELECT A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
            & "A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, '3' AS processed_ind, " _
            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee D "
            
        lsSQL20 = lsSQL20 & " WHERE " & lsEmpSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.OVER_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsOvrExcepInClause) > 0 Then
                    lsSQL20 = lsSQL20 & " AND " & lsOvrExcepInClause
                End If
            Else
                lsSQL20 = lsSQL20 & " AND " & lsAllOvrExcepInClause
            End If
        
        lsSQL20 = lsSQL20 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.OVER_EXC_CD <> ''"
            
        lsSQL20 = lsSQL20 & " UNION "

        lsSQL21 = lsSQL21 _
            & "SELECT A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
            & "A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, '3' AS processed_ind, " _
            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee D "
            
        lsSQL21 = lsSQL21 & " WHERE " & lsEmpSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsSchedChgExcepInClause) > 0 Then
                    lsSQL21 = lsSQL21 & " AND " & lsSchedChgExcepInClause
                End If
            Else
                lsSQL21 = lsSQL21 & " AND " & lsAllSchedChgExcepInClause
            End If
        
        lsSQL21 = lsSQL21 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.SCHED_CHG_EXC_CD <> ''"
            
        lsSQL21 = lsSQL21 & " UNION "

    End If
    
    If (RecordsExist(v_adrsSelectedExceptions)) Then
        If Len(lsExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL2 & lsSQL5 & lsSQL10 & lsSQL16 & lsSQL19
        End If
        If Len(lsOvrExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL3 & lsSQL11 & lsSQL17 & lsSQL20
        End If
        If Len(lsSchedChgExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL4 & lsSQL12 & lsSQL18 & lsSQL21
        End If
        If Len(lsOvrExcepInClause) > 0 Or Len(lsExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL1
        End If
    Else
        lsSQLALL = lsSQLALL & lsSQL1 & lsSQL2 & lsSQL3 & lsSQL4 & lsSQL5 & lsSQL10 & _
            lsSQL11 & lsSQL12 & lsSQL16 & lsSQL17 & lsSQL18 & lsSQL19 & lsSQL20 & lsSQL21
    End If
    
    If Trim(Right(lsSQLALL, 6)) = "UNION" Then
        lsSQLALL = Left(lsSQLALL, Len(lsSQLALL) - 6)
    End If
    
    '//try to elimiate duplicates when selected exceptions are 'E' and 'FCL'
    'sql 16 - use actual time & exception cd 'E'
    'sql 17 - use actual time & override exception cd 'FCL'
    'since these were UNION ALL, there would be 2 rows (of the same actual time) only differing by CASE_GROUP
    '//take CASE_GROUP out of SQL detail and group by, use UNION rather than UNION ALL, then use 3 joins to get the possible CASE_GROUP's
    If Len(lsSQLALL) > 0 Then
        lsSQL = "SELECT EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, NTE_GROUP.EXCEPTION_CD, " _
            & " OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, COALESCE(NullIf(EXC1.CASE_GROUP, ''), NullIf(EXC2.CASE_GROUP, ''), NullIf(EXC3.CASE_GROUP, ''), '    ') AS CASE_GROUP, ADJ_TYPE, " _
            & " PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, AMOUNT, " _
            & " TRAVEL_PAY_HRS, HOURS FROM (" & vbCrLf
    
        lsSQL = lsSQL & "SELECT EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, EXCEPTION_CD, " _
            & " OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, ADJ_TYPE, " _
            & " PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, SUM(AMOUNT) AS AMOUNT, " _
            & " SUM(TRAVEL_PAY_HRS) AS TRAVEL_PAY_HRS, SUM(HOURS) AS HOURS FROM (" _
            & lsSQLALL & ") NTE " & vbCrLf
            
        lsSQL = lsSQL & " GROUP BY EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, EXCEPTION_CD, " _
            & " OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, ADJ_TYPE, " _
            & " PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM ) NTE_GROUP " _
            & " " & vbCrLf _
            & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC1" _
            & " ON NTE_GROUP.EXCEPTION_CD = EXC1.EXCEPTION_CD " _
            & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC2 " _
            & " ON NTE_GROUP.OVER_EXC_CD = EXC2.EXCEPTION_CD " _
            & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC3" _
            & " ON NTE_GROUP.SCHED_CHG_EXC_CD = EXC3.EXCEPTION_CD "
        lsSQL = lsSQL & " ORDER BY emp_num, processed_ind, charge_date"
    End If

    Set ladCN = New ADODB.Connection
    Set lobjCM = New clsuDBConnectMgr
    ladCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
    Set lobjCM = Nothing
    ladCN.Open

    Set ladrs = New ADODB.Recordset
    ladrs.CursorLocation = adUseClient
    ladrs.Open lsSQL, ladCN _
        , adOpenStatic, adLockBatchOptimistic
    ladrs.ActiveConnection = Nothing

    ladCN.Close
    Set ladCN = Nothing

    Set GetEmpExceptions = ladrs
    Set ladrs = Nothing
   
    '//-----clean up class variables used by madrsExceptions/GetExceptionCodes/IsExceptionCode-----//
    DestroyState
    
    Exit Function
eh:
    Set lobjCM = Nothing
    Set ladCN = Nothing
    Set ladrs = Nothing

    DestroyState

    TrsRaiseError m_strCMODULENAME, "GetEmpExceptions" _
        , , , , DescArgsFrom(lsSQL) _
        , True

End Function

'Public Function GetEmpExceptionsOLD( _
'                ByVal v_aryEmployees As Variant _
'                , ByVal v_datStartDate As Date _
'                , ByVal v_datEndDate As Date _
'                , ByVal v_adrsSelectedExceptions As ADODB.Recordset _
'                , Optional v_datLastClosedPP As Date = CDate("1980-01-01") _
'                , Optional v_datEndOfOpenPP As Date = CDate("1980-01-01") _
'                ) As ADODB.Recordset
'On Error GoTo eh
'Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
'Dim ladCN As ADODB.Connection
'Dim ladrs As ADODB.Recordset
'Dim lsEmpSelection As String
'Dim ldtHistoryEndDate As Date, ldtSchedExcStartDate As Date
'Dim ldtActualTimeStartDate As Date, ldtActualTimeEndDate As Date
'Dim lsInClausePartValueKey As String, lsExcepInClause As String
'Dim lsOvrExcepInClause As String
'Dim lsSchedChgExcepInClause As String
'Dim lsAllExcepInClause As String
'Dim lsAllOvrExcepInClause As String
'Dim lsAllSchedChgExcepInClause As String
'Dim lsAllExcepandOvrInClause As String
'Dim lsExcepandOvrInClause As String
'Dim liActualTimeTables As TrActualTimeTable, liRecCounter As Integer
'Dim lblnValidExcep As Boolean
'Dim lblnFirstExcep As Boolean
'Dim lsSQL1 As String
'Dim lsSQL2 As String
'Dim lsSQL3 As String
'Dim lsSQL4 As String
'Dim lsSQL5 As String
'Dim lsSQL10 As String
'Dim lsSQL11 As String
'Dim lsSQL12 As String
'Dim lsSQL16 As String
'Dim lsSQL17 As String
'Dim lsSQL18 As String
'Dim lsSQL19 As String
'Dim lsSQL20 As String
'Dim lsSQL21 As String
'Dim lsSQLALL As String
'Dim lsSQL As String
'Dim lblnUnion As Boolean
'
'    If (IsArray(v_aryEmployees)) Then
'        lsEmpSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
'    Else
'        lsEmpSelection = "A.EMP_NUM = " & CStr(CLng(v_aryEmployees)) & " "
'    End If
'
'    liActualTimeTables = DetermineActualTimeTables(v_datStartDate _
'                            , v_datEndDate _
'                            , ldtHistoryEndDate _
'                            , ldtActualTimeStartDate)
'
'    If v_datLastClosedPP = #1/1/1980# Then
'        ldtSchedExcStartDate = v_datStartDate
'    Else
'        If (v_datLastClosedPP > ldtActualTimeStartDate) Then
'            ldtSchedExcStartDate = DateAdd("d", 1, v_datLastClosedPP)
'        Else
'            ldtSchedExcStartDate = ldtActualTimeStartDate
'        End If
'    End If
'
'    If v_datEndOfOpenPP = #1/1/1980# Then
'        ldtActualTimeEndDate = v_datEndDate
'    Else
'        If (v_datEndOfOpenPP > ldtActualTimeStartDate) Then
'            ldtActualTimeEndDate = v_datEndOfOpenPP
'        Else
'            ldtActualTimeEndDate = v_datEndDate
'        End If
'    End If
'
'    GetExceptionCodes
'    GetOverExcepCodes
'    GetSchedChgExcepCodes
'
'    If (RecordsExist(v_adrsSelectedExceptions)) Then
'
'        With v_adrsSelectedExceptions
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsExceptionCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsExcepInClause = lsExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsExcepInClause = lsExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsExcepInClause = " A.EXCEPTION_CD IN ( " & lsExcepInClause & ")"
'            End If
'
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsOverExcepCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsOvrExcepInClause = lsOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsOvrExcepInClause = lsOvrExcepInClause & _
'                            ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsOvrExcepInClause & _
'                ")"
'            End If
'
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsSchedChgExcepCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
'                        "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
'                lsSchedChgExcepInClause & ")"
'            End If
'
'        End With
'
'    Else
'
'        If RecordsExist(madrsExceptions) Then
'            With madrsExceptions
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllExcepInClause = lsAllExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllExcepInClause = lsAllExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllExcepInClause = " A.EXCEPTION_CD IN ( " & lsAllExcepInClause & ")"
'            End With
'        End If
'
'        If RecordsExist(madrsOverExcep) Then
'            With madrsOverExcep
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & _
'                            ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsAllOvrExcepInClause & ")"
'            End With
'        End If
'
'        If RecordsExist(madrsSchedChgExcep) Then
'            With madrsSchedChgExcep
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
'                        "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
'                lsAllSchedChgExcepInClause & ")"
'            End With
'        End If
'
'    End If
'
''*******************************************************
''*  Retrieve Scheduled Exceptions
''*******************************************************
'
'    lsSQL1 = lsSQL1 & "SELECT a.emp_num, d.emp_name,  a.job_num, a.charge_date, " _
'        & "'00:00' as start_time, '00:00' as end_time, a.exception_cd, " _
'        & "a.over_exc_cd, '' as sched_chg_exc_cd, a.case_num, b.case_group, " _
'        & " 'C' as adj_type, '1' AS processed_ind, '' as OT_OPP_DESC, " _
'        & " '' as distr_dept_num, 0 as amount, 0 as travel_pay_hrs, hours " _
'        & "FROM methdba.ttr_emp_sched_exc A, " _
'        & " methdba.ttr_exception_cds B, " _
'        & " methdba.ttr_employee D " _
'        & " WHERE " & lsEmpSelection _
'        & " AND A.charge_date between " & DateAsEnclosedString(ldtSchedExcStartDate) & " AND " _
'                & DateAsEnclosedString(v_datEndDate) _
'        & " AND A.emp_num = D.emp_num " _
'        & " AND B.exception_cd = A.exception_cd "
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) > 0 Then
'                lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & " OR " & _
'                lsOvrExcepInClause & ") "
'            Else
'                If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) = 0 Then
'                    lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & ") "
'                Else
'                    If Len(lsExcepInClause) = 0 And Len(lsOvrExcepInClause) > 0 Then
'                        lsSQL1 = lsSQL1 & " AND (" & lsOvrExcepInClause & ") "
'                    Else
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                        " OR A.OVER_EXC_CD <> '') "
'                    End If
'                End If
'            End If
'        Else
'            If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) > 0 Then
'                lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & " OR " & _
'                    lsAllOvrExcepInClause & ") "
'                lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'            Else
'                If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) = 0 Then
'                    lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & ") "
'                    lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                Else
'                    If Len(lsAllExcepInClause) = 0 And Len(lsAllOvrExcepInClause) > 0 Then
'                        lsSQL1 = lsSQL1 & " AND (" & lsAllOvrExcepInClause & ") "
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                    Else
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                    End If
'                End If
'            End If
'        End If
'
'        lsSQL1 = lsSQL1 _
'        & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'        & " AND A.exception_cd NOT IN ( SELECT exception_cd  " _
'                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5 ) " _
'        & " AND A.over_exc_cd NOT IN ( SELECT exception_cd  " _
'                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5) "
'
'        lsSQL1 = lsSQL1 & " UNION ALL "
'
'    If (v_datStartDate < #1/27/2005#) Then
'        'GET INFO FROM OLD AND NEW TABLES
'        'FOR SHOULD HAVE BEEN & AS PAID
'        lsSQL2 = lsSQL2 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'S' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL2 = lsSQL2 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMPLOYEE E"
'
'       lsSQL2 = lsSQL2 & _
'            " WHERE " & lsEmpSelection & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND B.EXCEPTION_CD = C.EXCEPTION_CD"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsExcepInClause) > 0) Then
'                lsSQL2 = lsSQL2 & _
'                    " AND " & Replace(lsExcepInClause, "A.", "B.")
'            Else
'                lsSQL2 = lsSQL2 & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'            End If
'        Else
'            lsSQL2 = lsSQL2 & _
'                " AND " & Replace(lsAllExcepInClause, "A.", "B.")
'                lsSQL2 = lsSQL2 & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'        End If
'
'        lsSQL2 = lsSQL2 & _
'            " AND B.EXCEPTION_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL2 = lsSQL2 & " UNION ALL "
'
'        lsSQL3 = lsSQL3 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL3 = lsSQL3 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMPLOYEE E"
'
'        lsSQL3 = lsSQL3 & _
'            " WHERE " & lsEmpSelection & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND B.OVER_EXC_CD = C.EXCEPTION_CD"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsOvrExcepInClause) > 0) Then
'                lsSQL3 = lsSQL3 & _
'                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL3 = lsSQL3 & _
'                " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL3 = lsSQL3 & _
'            " AND B.OVER_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL3 = lsSQL3 & " UNION ALL "
'
'        lsSQL4 = lsSQL4 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL4 = lsSQL4 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMPLOYEE E"
'
'        lsSQL4 = lsSQL4 & _
'            " WHERE " & lsEmpSelection & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsSchedChgExcepInClause) > 0) Then
'                lsSQL4 = lsSQL4 & _
'                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL4 = lsSQL4 & _
'                " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL4 = lsSQL4 & _
'            " AND B.SCHED_CHG_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL4 = lsSQL4 & " UNION ALL "
'
'        lsSQL5 = lsSQL5 & " SELECT A.EMP_NUM, D.EMP_NAME" & _
'            " , A.JOB_NUM, A.DATE_WORKED AS CHARGE_DATE" & _
'            " , A.START_TIME, A.END_TIME" & _
'            " , A.EXCEPTION_CD, '' AS OVER_EXC_CD" & _
'            " , '' AS SCHED_CHG_EXC_CD" & _
'            " , A.CASE_NUM, B.CASE_GROUP" & _
'            " , A.ADJ_TYPE" & _
'            " , CASE WHEN (A.PROCESSED_IND) = 'Y' THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , '' AS OT_OPP_DESC, '' AS DISTR_DEPT_NUM" & _
'            " , 0 AS AMOUNT, 0 AS TRAVEL_PAY_HRS" & _
'            " , CASE WHEN ADJ_TYPE = 'A' THEN (A.HOURS * -1) ELSE A.HOURS END AS HOURS" & _
'            " FROM METHDBA.TTR_ADJUSTMENTS A" & _
'            " , METHDBA.TTR_EXCEPTION_CDS B" & _
'            " , METHDBA.TTR_EMPLOYEE D"
'        lsSQL5 = lsSQL5 & _
'            " WHERE " & lsEmpSelection & _
'            " AND A.EMP_NUM = D.EMP_NUM" & _
'            " AND A.DATE_WORKED BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.HOURS <> 0" & _
'            " AND A.EXCEPTION_CD = B.EXCEPTION_CD"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsExcepInClause) > 0) Then
'                lsSQL5 = lsSQL5 & _
'                    " AND " & lsExcepInClause
'            Else
'                lsSQL5 = lsSQL5 & _
'                    " AND A.EXCEPTION_CD <> 'ST'"
'            End If
'        Else
'            lsSQL5 = lsSQL5 & _
'                " AND " & lsAllExcepInClause
'            lsSQL5 = lsSQL5 & _
'                    " AND A.EXCEPTION_CD <> 'ST'"
'        End If
'
'        lsSQL5 = lsSQL5 & _
'            " AND A.EXCEPTION_CD NOT IN ('ET','AT')" & _
'            " AND A.EXCEPTION_CD <> ''"
'
'        lsSQL5 = lsSQL5 & " UNION ALL "
'
'    Else
'        'GET ADJ FROM NEW TABLE FOR SHOULD HAVE BEEN & AS PAID
'        lsSQL10 = lsSQL10 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL10 = lsSQL10 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMPLOYEE E"
'
'        lsSQL10 = lsSQL10 & _
'            " WHERE " & lsEmpSelection & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND B.EXCEPTION_CD = C.EXCEPTION_CD"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsExcepInClause) > 0) Then
'                lsSQL10 = lsSQL10 & _
'                    " AND " & Replace(lsExcepInClause, "A.", "B.")
'            Else
'                lsSQL10 = lsSQL10 & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'            End If
'        Else
'            lsSQL10 = lsSQL10 & _
'                " AND " & Replace(lsAllExcepInClause, "A.", "B.")
'            lsSQL10 = lsSQL10 & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'        End If
'
'        lsSQL10 = lsSQL10 & _
'            " AND B.EXCEPTION_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL10 = lsSQL10 & " UNION ALL "
'
'        lsSQL11 = lsSQL11 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL11 = lsSQL11 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMPLOYEE E"
'
'        lsSQL11 = lsSQL11 & _
'            " WHERE " & lsEmpSelection & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND B.OVER_EXC_CD = C.EXCEPTION_CD"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsOvrExcepInClause) > 0) Then
'                lsSQL11 = lsSQL11 & _
'                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL11 = lsSQL11 & _
'                " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL11 = lsSQL11 & _
'            " AND B.OVER_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL11 = lsSQL11 & " UNION ALL "
'
'        lsSQL12 = lsSQL12 & " SELECT A.EMP_NUM, E.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL12 = lsSQL12 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMPLOYEE E"
'
'        lsSQL12 = lsSQL12 & _
'            " WHERE " & lsEmpSelection & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsSchedChgExcepInClause) > 0) Then
'                lsSQL12 = lsSQL12 & _
'                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL12 = lsSQL12 & _
'                " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL12 = lsSQL12 & _
'            " AND B.SCHED_CHG_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL12 = lsSQL12 & " UNION ALL "
'
'    End If
''END PREM TRIS CONVERSION PHASE 3
'
''*******************************************************
''*  Retrieve Exceptions from Actual Time
''*******************************************************
'    If liActualTimeTables = TrTable_Actual Or _
'        liActualTimeTables = TrTable_ActualAndHistory Then
'        lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
'        lsSQL16 = lsSQL16 _
'            & "SELECT A.emp_num, d.emp_name, A.job_num, A.charge_date, A.start_time, " _
'            & "A.end_time, A.exception_cd, A.over_exc_cd, A.sched_chg_exc_cd, " _
'            & "A.case_num, B.case_group, ' ' as adj_type," _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM," _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL16 = lsSQL16 & "WHERE " & lsEmpSelection _
'            & " AND A.charge_date between " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.exception_cd "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsExcepInClause) > 0 Then
'                    lsSQL16 = lsSQL16 & " AND " & lsExcepInClause
'                Else
'                    lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'                End If
'            Else
'                lsSQL16 = lsSQL16 & " AND " & lsAllExcepInClause _
'                & " AND A.EXCEPTION_CD <> 'ST'" _
'                & " AND A.OVER_EXC_CD = ''" _
'                & " AND A.SCHED_CHG_EXC_CD = ''"
'            End If
'
'        lsSQL16 = lsSQL16 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.EXCEPTION_CD <> ''"
'
'        lsSQL16 = lsSQL16 & " UNION ALL "
'
'        lsSQL17 = lsSQL17 _
'            & "SELECT A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
'            & " A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, " _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM," _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL17 = lsSQL17 & "WHERE " & lsEmpSelection _
'            & " AND A.charge_date between " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.OVER_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsOvrExcepInClause) > 0 Then
'                    lsSQL17 = lsSQL17 & " AND " & lsOvrExcepInClause
'                End If
'            Else
'                lsSQL17 = lsSQL17 & " AND " & lsAllOvrExcepInClause
'            End If
'
'        lsSQL17 = lsSQL17 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.OVER_EXC_CD <> ''"
'
'        lsSQL17 = lsSQL17 & " UNION ALL "
'
'        lsSQL18 = lsSQL18 _
'            & "SELECT A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
'            & " A.start_time, A.end_time, A.exception_cd, " _
'            & " A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & " ' ' as adj_type, " _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL18 = lsSQL18 & "WHERE " & lsEmpSelection _
'            & " AND A.charge_date between " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsSchedChgExcepInClause) > 0 Then
'                    lsSQL18 = lsSQL18 & " AND " & lsSchedChgExcepInClause
'                End If
'            Else
'                lsSQL18 = lsSQL18 & " AND " & lsAllSchedChgExcepInClause
'            End If
'
'        lsSQL18 = lsSQL18 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.SCHED_CHG_EXC_CD <> ''"
'
'        lsSQL18 = lsSQL18 & " UNION ALL "
'
'    End If
'
''*******************************************************
''*  Retrieve Exceptions from Actual Time History
''*******************************************************
'    If liActualTimeTables = TrTable_History Or _
'       liActualTimeTables = TrTable_ActualAndHistory Then
'       lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
'       lsSQL19 = lsSQL19 _
'            & "SELECT A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
'            & "A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, '3' AS processed_ind, " _
'            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL19 = lsSQL19 & " WHERE " & lsEmpSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.exception_cd "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsExcepInClause) > 0 Then
'                    lsSQL19 = lsSQL19 & " AND " & lsExcepInClause
'                Else
'                    lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'                End If
'            Else
'                lsSQL19 = lsSQL19 & " AND " & lsAllExcepInClause _
'                & " AND A.EXCEPTION_CD <> 'ST'" _
'                & " AND A.OVER_EXC_CD = ''" _
'                & " AND A.SCHED_CHG_EXC_CD = ''"
'            End If
'
'        lsSQL19 = lsSQL19 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.EXCEPTION_CD <> ''"
'
'        lsSQL19 = lsSQL19 & " UNION ALL "
'
'        lsSQL20 = lsSQL20 _
'            & "SELECT A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
'            & "A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, '3' AS processed_ind, " _
'            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL20 = lsSQL20 & " WHERE " & lsEmpSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.OVER_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsOvrExcepInClause) > 0 Then
'                    lsSQL20 = lsSQL20 & " AND " & lsOvrExcepInClause
'                End If
'            Else
'                lsSQL20 = lsSQL20 & " AND " & lsAllOvrExcepInClause
'            End If
'
'        lsSQL20 = lsSQL20 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.OVER_EXC_CD <> ''"
'
'        lsSQL20 = lsSQL20 & " UNION ALL "
'
'        lsSQL21 = lsSQL21 _
'            & "SELECT A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
'            & "A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, '3' AS processed_ind, " _
'            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL21 = lsSQL21 & " WHERE " & lsEmpSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsSchedChgExcepInClause) > 0 Then
'                    lsSQL21 = lsSQL21 & " AND " & lsSchedChgExcepInClause
'                End If
'            Else
'                lsSQL21 = lsSQL21 & " AND " & lsAllSchedChgExcepInClause
'            End If
'
'        lsSQL21 = lsSQL21 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.SCHED_CHG_EXC_CD <> ''"
'
'        lsSQL21 = lsSQL21 & " UNION ALL "
'
'    End If
'
'    If (RecordsExist(v_adrsSelectedExceptions)) Then
'        If Len(lsExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL2 & lsSQL5 & lsSQL10 & lsSQL16 & lsSQL19
'        End If
'        If Len(lsOvrExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL3 & lsSQL11 & lsSQL17 & lsSQL20
'        End If
'        If Len(lsSchedChgExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL4 & lsSQL12 & lsSQL18 & lsSQL21
'        End If
'        If Len(lsOvrExcepInClause) > 0 Or Len(lsExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL1
'        End If
'    Else
'        lsSQLALL = lsSQLALL & lsSQL1 & lsSQL2 & lsSQL3 & lsSQL4 & lsSQL5 & lsSQL10 & _
'            lsSQL11 & lsSQL12 & lsSQL16 & lsSQL17 & lsSQL18 & lsSQL19 & lsSQL20 & lsSQL21
'    End If
'
'    If Trim(Right(lsSQLALL, 10)) = "UNION ALL" Then
'        lsSQLALL = Left(lsSQLALL, Len(lsSQLALL) - 10)
'    End If
'
'    If Len(lsSQLALL) > 0 Then
'        lsSQL = "SELECT EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, EXCEPTION_CD, " _
'            & " OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM , CASE_GROUP, ADJ_TYPE, " _
'            & " PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, SUM(AMOUNT) AS AMOUNT, " _
'            & " SUM(TRAVEL_PAY_HRS) AS TRAVEL_PAY_HRS, SUM(HOURS) AS HOURS FROM (" _
'            & lsSQLALL & ") NTE "
'        lsSQL = lsSQL & " GROUP BY EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, EXCEPTION_CD, " _
'            & " OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, CASE_GROUP, ADJ_TYPE, " _
'            & " PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM "
'        lsSQL = lsSQL & " ORDER BY emp_num, processed_ind, charge_date"
'    End If
'
'    Set ladCN = New ADODB.Connection
'    Set lobjCM = New clsuDBConnectMgr
'    ladCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
'    Set lobjCM = Nothing
'    ladCN.Open
'
'    Set ladrs = New ADODB.Recordset
'    ladrs.CursorLocation = adUseClient
'    ladrs.Open lsSQL, ladCN _
'        , adOpenStatic, adLockBatchOptimistic
'    ladrs.ActiveConnection = Nothing
'
'    ladCN.Close
'    Set ladCN = Nothing
'
'    Set GetEmpExceptionsOLD = ladrs
'    Set ladrs = Nothing
'
'    '//-----clean up class variables used by madrsExceptions/GetExceptionCodes/IsExceptionCode-----//
'    DestroyState
'
'    Exit Function
'eh:
'    Set lobjCM = Nothing
'    Set ladCN = Nothing
'    Set ladrs = Nothing
'
'    DestroyState
'
'    TrsRaiseError m_strCMODULENAME, "GetEmpExceptionsOLD" _
'        , , , , DescArgsFrom(lsSQL) _
'        , True
'
'End Function

Public Function GetEmpVacationsForYear( _
                ByVal v_aryEmployees As Variant _
                , ByVal v_lYear As Long _
                ) As ADODB.Recordset
On Error GoTo eh
Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
Dim lobjCN As ADODB.Connection
Dim lobjRS As ADODB.Recordset
Dim lsSQL As String

Dim lsEmpSelection As String, lsInClausePartValueKey As String
Dim ldtHistoryEndDate As Date, ldtActualTimeStartDate As Date
Dim liActualTimeTables As TrActualTimeTable, l_sExceptionInClause As String
Dim ldtStartDate As Date, ldtEndDate As Date
Dim ldtHistoryStart As Date, ldtHistoryEnd As Date
Dim ldtActualStart As Date, ldtActualEnd As Date
Dim ldtScheduleStart As Date, ldtScheduleEnd As Date

    If (IsArray(v_aryEmployees)) Then
        lsEmpSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        lsEmpSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    ldtStartDate = CDate(CStr(v_lYear) & "-01-01")
    ldtEndDate = CDate(CStr(v_lYear) & "-12-31")

    liActualTimeTables = DetermineActualTimeTables(ldtStartDate _
                            , ldtEndDate _
                            , ldtHistoryEndDate _
                            , ldtActualTimeStartDate)

    If (liActualTimeTables = TrTable_History) Or _
        (liActualTimeTables = TrTable_ActualAndHistory) Then
        ldtHistoryStart = ldtStartDate
        ldtHistoryEnd = ldtHistoryEnd
    End If

    If (liActualTimeTables = TrTable_Actual) Or _
        (liActualTimeTables = TrTable_ActualAndHistory) Then
        If (liActualTimeTables = TrTable_History) Or _
            (liActualTimeTables = TrTable_ActualAndHistory) Then
            ldtActualStart = DateAdd("d", 1, ldtHistoryEndDate)
            ldtActualEnd = Format(Now(), "mm/dd/yyyy")
        Else
            ldtActualStart = ldtStartDate
            ldtActualEnd = ldtEndDate
        End If
    End If

    ldtScheduleStart = ldtActualStart
    ldtScheduleEnd = ldtEndDate

'*******************************************************
'*  Retrieve Vacation  Scheduled Exceptions
'*******************************************************
    lsSQL = "SELECT a.emp_num,  a.charge_date, a.exception_cd, ' ' as adj_type, '1' as processed_ind,  hours   " _
        & "FROM methdba.ttr_emp_sched_exc A " _
        & " WHERE " & lsEmpSelection _
        & "  AND A.charge_date between " & DateAsEnclosedString(ldtScheduleStart) & " AND " _
                & DateAsEnclosedString(ldtScheduleEnd) _
        & " AND A.exception_cd  IN ( SELECT rg_value_cd FROM methdba.ttr_rg_group WHERE rg_cd = 99 ) "

'*******************************************************
'*  Retrieve Vacation  "Should be paid" Adjustments
'*******************************************************
'BEGIN PREM TRIS CONVERSION PHASE 3

    lsSQL = lsSQL & " UNION ALL "

    'NOTE JAN 26, 2005 IS THE CUTOFF FOR OLD TABLE Vs. NEW TABLE
    If (v_lYear < 2006) Then
        'Get info from old and new tables
        ' FOR AS AND SHOULD HAVE BEEN PAID
        lsSQL = lsSQL & " SELECT    A.EMP_NUM" & _
            "       , VALUE(NULLIF(B.CHARGE_DATE,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , B.EXCEPTION_CD, 'S' AS ADJ_TYPE" & _
            "       , '2' AS PROCESSED_IND" & _
            "       , SUM(B.HOURS) AS HOURS" & _
            " FROM  METHDBA.TTR_ADJ_RQT A" & _
            "       , METHDBA.TTR_ADJ_TIME B" & _
            " WHERE " & lsEmpSelection & _
            " AND   A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND   B.CHARGE_DATE BETWEEN '" & ldtActualStart & "'" & _
            " AND   '" & ldtActualEnd & "'" & _
            " AND   B.HOURS > 0" & _
            " AND   B.EXCEPTION_CD IN (SELECT   RG_VALUE_CD" & _
            "                       FROM        METHDBA.TTR_RG_GROUP" & _
            "                       WHERE       RG_CD = 99)" & _
            " GROUP BY A.EMP_NUM, B.CHARGE_DATE, B.EXCEPTION_CD"

        lsSQL = lsSQL & " UNION ALL"

        lsSQL = lsSQL & " SELECT    EMP_NUM" & _
            "       , VALUE(NULLIF(DATE_WORKED,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , EXCEPTION_CD, ADJ_TYPE" & _
            "       , '2' AS PROCESSED_IND" & _
            "       , SUM(HOURS) AS HOURS" & _
            " FROM  METHDBA.TTR_ADJUSTMENTS" & _
            " WHERE " & lsEmpSelection & _
            " AND   DATE_WORKED BETWEEN '" & ldtActualStart & "'" & _
            " AND   '" & ldtActualEnd & "'" & _
            " AND   HOURS <> 0" & _
            " AND   ADJ_TYPE = 'S'" & _
            " AND   EXCEPTION_CD IN (SELECT   RG_VALUE_CD" & _
            "                       FROM        METHDBA.TTR_RG_GROUP" & _
            "                       WHERE       RG_CD = 99)" & _
            " GROUP BY EMP_NUM, DATE_WORKED, EXCEPTION_CD, ADJ_TYPE"

        lsSQL = lsSQL & " UNION ALL"

        lsSQL = lsSQL & " SELECT    A.EMP_NUM" & _
            "       , VALUE(NULLIF(B.CHARGE_DATE,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , B.EXCEPTION_CD, 'A' AS ADJ_TYPE" & _
            "       , '2' AS PROCESSED_IND" & _
            "       , SUM(B.HOURS) AS HOURS" & _
            " FROM  METHDBA.TTR_ADJ_RQT A" & _
            "       , METHDBA.TTR_ADJ_TIME B" & _
            " WHERE " & lsEmpSelection & _
            " AND   A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND   B.CHARGE_DATE BETWEEN '" & ldtActualStart & "'" & _
            " AND   '" & ldtActualEnd & "'" & _
            " AND   B.HOURS < 0" & _
            " AND   B.EXCEPTION_CD IN (SELECT   RG_VALUE_CD" & _
            "                       FROM        METHDBA.TTR_RG_GROUP" & _
            "                       WHERE       RG_CD = 99)" & _
            " GROUP BY A.EMP_NUM, B.CHARGE_DATE, B.EXCEPTION_CD"

        lsSQL = lsSQL & " UNION ALL"
'PLEASE NOTE HOURS WERE NOT SIGNED (NEGATIVE) FOR AS PAID IN TTR_ADJUSTMENTS
        lsSQL = lsSQL & " SELECT    EMP_NUM" & _
            "       , VALUE(NULLIF(DATE_WORKED,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , EXCEPTION_CD, ADJ_TYPE" & _
            "       , '2' AS PROCESSED_IND" & _
            "       , SUM(HOURS * -1) AS HOURS" & _
            " FROM  METHDBA.TTR_ADJUSTMENTS" & _
            " WHERE " & lsEmpSelection & _
            " AND   DATE_WORKED BETWEEN '" & ldtActualStart & "'" & _
            " AND   '" & ldtActualEnd & "'" & _
            " AND   HOURS <> 0" & _
            " AND   ADJ_TYPE = 'A'" & _
            " AND   EXCEPTION_CD IN (SELECT   RG_VALUE_CD" & _
            "                       FROM        METHDBA.TTR_RG_GROUP" & _
            "                       WHERE       RG_CD = 99)" & _
            " GROUP BY EMP_NUM, DATE_WORKED, EXCEPTION_CD, ADJ_TYPE"
    Else
        'Get info from new table
        lsSQL = lsSQL & " SELECT    A.EMP_NUM" & _
            "       , VALUE(NULLIF(B.CHARGE_DATE,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , B.EXCEPTION_CD, 'S' AS ADJ_TYPE" & _
            "       , '2' AS PROCESSED_IND" & _
            "       , SUM(B.HOURS) AS HOURS" & _
            " FROM  METHDBA.TTR_ADJ_RQT A" & _
            "       , METHDBA.TTR_ADJ_TIME B" & _
            " WHERE " & lsEmpSelection & _
            " AND   A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND   B.CHARGE_DATE BETWEEN '" & ldtActualStart & "'" & _
            " AND   '" & ldtActualEnd & "'" & _
            " AND   B.HOURS > 0" & _
            " AND   B.EXCEPTION_CD IN (SELECT   RG_VALUE_CD" & _
            "                       FROM        METHDBA.TTR_RG_GROUP" & _
            "                       WHERE       RG_CD = 99)" & _
            " GROUP BY A.EMP_NUM, B.CHARGE_DATE, B.EXCEPTION_CD"

        lsSQL = lsSQL & " UNION ALL"

        lsSQL = lsSQL & " SELECT    A.EMP_NUM" & _
            "       , VALUE(NULLIF(B.CHARGE_DATE,'0001-01-01'), DATE('0100-01-01')) AS CHARGE_DATE" & _
            "       , B.EXCEPTION_CD, 'A' AS ADJ_TYPE" & _
            "       , '2' AS PROCESSED_IND" & _
            "       , SUM(B.HOURS) AS HOURS" & _
            " FROM  METHDBA.TTR_ADJ_RQT A" & _
            "       , METHDBA.TTR_ADJ_TIME B" & _
            " WHERE " & lsEmpSelection & _
            " AND   A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND   B.CHARGE_DATE BETWEEN '" & ldtActualStart & "'" & _
            " AND   '" & ldtActualEnd & "'" & _
            " AND   B.HOURS < 0" & _
            " AND   B.EXCEPTION_CD IN (SELECT   RG_VALUE_CD" & _
            "                       FROM        METHDBA.TTR_RG_GROUP" & _
            "                       WHERE       RG_CD = 99)" & _
            " GROUP BY A.EMP_NUM, B.CHARGE_DATE, B.EXCEPTION_CD"
    End If
'END PREM TRIS CONVERSION PHASE 3

'*******************************************************
'*  Retrieve Vacation Exceptions from Actual Time
'*******************************************************
    If liActualTimeTables = TrTable_Actual Or _
        liActualTimeTables = TrTable_ActualAndHistory Then
        lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(ldtStartDate, ldtHistoryEndDate), True)
        lsSQL = lsSQL & " UNION ALL " _
            & "SELECT A.emp_num,   A.charge_date, A.exception_cd,' ' as adj_type,'3' as processed_ind   , SUM(hours) AS hours" _
            & " FROM METHDBA.TTR_ACTUAL_TIME A " _
            & "WHERE " & lsEmpSelection _
            & "  AND charge_date between  " & DateAsEnclosedString(ldtActualStart) & " AND " _
                    & DateAsEnclosedString(ldtActualEnd) _
            & " AND A.exception_cd  IN ( SELECT rg_value_cd FROM methdba.ttr_rg_group WHERE rg_cd = 99 ) " _
            & "GROUP BY a.emp_num, a.charge_date, a.exception_cd, processed_ind "
    End If
    
'*******************************************************
'*  Retrieve Exceptions from Actual Time History
'*******************************************************
    If liActualTimeTables = TrTable_History Or _
       liActualTimeTables = TrTable_ActualAndHistory Then
       lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(ldtStartDate, ldtHistoryEndDate), True)
       lsSQL = lsSQL & " UNION ALL " _
            & "SELECT A.emp_num,   A.charge_date, A.exception_cd,' ' as adj_type,'3' as processed_ind, SUM(hours) AS hours   " _
            & "from  METHDBA.TTR_ACTUAL_TIME_HS A " _
            & " WHERE " & lsEmpSelection & "  AND " & lsInClausePartValueKey _
            & "  AND charge_date between " & DateAsEnclosedString(ldtHistoryStart) & " AND " _
                    & DateAsEnclosedString(ldtHistoryEnd) _
            & " AND A.exception_cd  IN ( SELECT rg_value_cd FROM methdba.ttr_rg_group WHERE rg_cd = 99 ) " _
            & "GROUP BY a.emp_num, a.charge_date, a.exception_cd, processed_ind "
    End If
    
    lsSQL = lsSQL & " ORDER BY emp_num, charge_date"
        
        
    Set lobjCN = New ADODB.Connection
    Set lobjCM = New clsuDBConnectMgr
    lobjCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
    Set lobjCM = Nothing
    lobjCN.Open

    Set lobjRS = New ADODB.Recordset
    lobjRS.CursorLocation = adUseClient
    lobjRS.Open lsSQL, lobjCN _
        , adOpenStatic, adLockBatchOptimistic
    lobjRS.ActiveConnection = Nothing

    lobjCN.Close
    Set lobjCN = Nothing

    Set GetEmpVacationsForYear = lobjRS
    Set lobjRS = Nothing

    Exit Function
eh:
    Set lobjCM = Nothing
    Set lobjCN = Nothing
    Set lobjRS = Nothing
    TrsRaiseError m_strCMODULENAME _
        , "GetEmpExceptionsForYear" _
        , , , , DescArgsFrom(lsSQL) _
        , True

End Function

Public Function GetWGExceptions(ByVal v_aryWorkgroups As Variant _
                , ByVal v_datStartDate As Date _
                , ByVal v_datEndDate As Date _
                , ByVal v_adrsSelectedExceptions As ADODB.Recordset _
                , Optional v_datLastClosedPP As Date = CDate("1980-01-01") _
                , Optional v_datEndOfOpenPP As Date = CDate("1980-01-01") _
                ) As ADODB.Recordset
On Error GoTo eh
Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
Dim lobjCN As ADODB.Connection
Dim lobjRS As ADODB.Recordset
Dim lsSQL As String
Dim lsWGSelection As String, lsWGSelection2 As String
Dim ldtHistoryEndDate As Date, ldtSchedExcStartDate As Date
Dim ldtActualTimeStartDate As Date, ldtActualTimeEndDate As Date
Dim lsInClausePartValueKey As String, liActualTimeTables As TrActualTimeTable
Dim lsExcepInClause As String, liRecCounter As Integer
Dim lsOvrExcepInClause As String
Dim lsSchedChgExcepInClause As String
Dim lsExcepandOvrInClause As String
Dim lsAllExcepInClause As String
Dim lsAllOvrExcepInClause As String
Dim lsAllSchedChgExcepInClause As String
Dim lsAllExcepandOvrInClause As String
Dim lblnValidExcep As Boolean
Dim lblnFirstExcep As Boolean
Dim lsSQL1 As String
Dim lsSQL2 As String
Dim lsSQL3 As String
Dim lsSQL4 As String
Dim lsSQL5 As String
Dim lsSQL10 As String
Dim lsSQL11 As String
Dim lsSQL12 As String
Dim lsSQL16 As String
Dim lsSQL17 As String
Dim lsSQL18 As String
Dim lsSQL19 As String
Dim lsSQL20 As String
Dim lsSQL21 As String
Dim lsSQLALL As String

    If (IsArray(v_aryWorkgroups)) Then
        lsWGSelection = BuildInClause("A.WG_NUM", v_aryWorkgroups, False) & " "
        lsWGSelection2 = Replace(lsWGSelection, "A.WG_NUM", "D.WG_NUM")
    Else
        lsWGSelection = "A.WG_NUM = " & CStr(CLng(v_aryWorkgroups)) & " "
        lsWGSelection2 = Replace(lsWGSelection, "A.WG_NUM", "D.WG_NUM")
    End If

    liActualTimeTables = DetermineActualTimeTables(v_datStartDate _
                            , v_datEndDate _
                            , ldtHistoryEndDate _
                            , ldtActualTimeStartDate)
    
    If v_datLastClosedPP = #1/1/1980# Then
        ldtSchedExcStartDate = v_datStartDate
    Else
        If (v_datLastClosedPP > ldtActualTimeStartDate) Then
            ldtSchedExcStartDate = DateAdd("d", 1, v_datLastClosedPP)
        Else
            ldtSchedExcStartDate = ldtActualTimeStartDate
        End If
    End If
    If v_datEndOfOpenPP = #1/1/1980# Then
        ldtActualTimeEndDate = v_datEndDate
    Else
        If (v_datEndOfOpenPP > ldtActualTimeStartDate) Then
            ldtActualTimeEndDate = v_datEndOfOpenPP
        Else
            ldtActualTimeEndDate = v_datEndDate
        End If
    End If

    GetExceptionCodes
    GetOverExcepCodes
    GetSchedChgExcepCodes
    
    If (RecordsExist(v_adrsSelectedExceptions)) Then
        
        With v_adrsSelectedExceptions
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsExceptionCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsExcepInClause = lsExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsExcepInClause = lsExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If

                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsExcepInClause = " A.EXCEPTION_CD IN ( " & lsExcepInClause & ")"
            End If
            
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsOverExcepCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsOvrExcepInClause = lsOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsOvrExcepInClause = lsOvrExcepInClause & _
                            ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If
                
                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsOvrExcepInClause & _
                ")"
            End If
            
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsSchedChgExcepCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
                        "'" & !EXCEPTION_CD & "'"
                    Else
                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If
                
                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
                lsSchedChgExcepInClause & ")"
            End If
            
        End With
        
    Else
    
        If RecordsExist(madrsExceptions) Then
            With madrsExceptions
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllExcepInClause = lsAllExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllExcepInClause = lsAllExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
                
                lsAllExcepInClause = " A.EXCEPTION_CD IN ( " & lsAllExcepInClause & ")"
            End With
        End If
        
        If RecordsExist(madrsOverExcep) Then
            With madrsOverExcep
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & _
                            ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
                
                lsAllOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsAllOvrExcepInClause & ")"
            End With
        End If
        
        If RecordsExist(madrsSchedChgExcep) Then
            With madrsSchedChgExcep
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
                        "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
            
                lsAllSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
                lsAllSchedChgExcepInClause & ")"
            End With
        End If
        
    End If
    
'*******************************************************
'*  Retrieve Scheduled Exceptions
'*******************************************************
        
    lsSQL1 = lsSQL1 & "SELECT d.wg_num, a.emp_num, e.emp_name, a.job_num, " _
        & "a.charge_date, '00:00' as start_time, '00:00' as end_time, a.exception_cd, " _
        & "a.over_exc_cd, '' as sched_chg_exc_cd, a.case_num, " _
        & "'C' as adj_type, '1' AS processed_ind, '' AS OT_OPP_DESC, " _
        & " '' as distr_dept_num, 0 as amount, 0 as travel_pay_hrs, hours " _
        & "FROM methdba.ttr_emp_sched_exc A, " _
        & " methdba.ttr_exception_cds B, " _
        & " methdba.ttr_emp_wg D , methdba.ttr_employee E " _
        & " WHERE " & lsWGSelection2 _
        & " AND A.charge_date between " & DateAsEnclosedString(ldtSchedExcStartDate) & " AND " _
                & DateAsEnclosedString(v_datEndDate) _
        & " AND A.emp_num = D.emp_num " _
        & " AND A.emp_num = E.emp_num " _
        & " AND A.charge_date between D.eff_start_date AND D.eff_end_date " _
        & " AND B.exception_cd = A.exception_cd " _
        & " AND D.PRIMARY_IND = 'Y' "
        
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) > 0 Then
                lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & " OR " & _
                lsOvrExcepInClause & ") "
            Else
                If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) = 0 Then
                    lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & ") "
                Else
                    If Len(lsExcepInClause) = 0 And Len(lsOvrExcepInClause) > 0 Then
                        lsSQL1 = lsSQL1 & " AND (" & lsOvrExcepInClause & ") "
                    Else
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                        " OR A.OVER_EXC_CD <> '') "
                    End If
                End If
            End If
        Else
            If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) > 0 Then
                lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & " OR " & _
                    lsAllOvrExcepInClause & ") "
                lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
            Else
                If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) = 0 Then
                    lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & ") "
                    lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                Else
                    If Len(lsAllExcepInClause) = 0 And Len(lsAllOvrExcepInClause) > 0 Then
                        lsSQL1 = lsSQL1 & " AND (" & lsAllOvrExcepInClause & ") "
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                    Else
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                    End If
                End If
            End If
        End If
        
        lsSQL1 = lsSQL1 _
        & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
        & " AND A.exception_cd NOT IN ( SELECT exception_cd  " _
                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5 ) " _
        & " AND A.over_exc_cd NOT IN ( SELECT exception_cd  " _
                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5) "
                    
'BEGIN PREM TRIS CONVERSION PHASE 3
    lsSQL1 = lsSQL1 & " UNION "

    If (v_datStartDate < #1/27/2005#) Then
        'Get from Old and New Tables
        lsSQL2 = lsSQL2 & " SELECT E.WG_NUM" & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
                   
        lsSQL2 = lsSQL2 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsWGSelection2, "D.", "E.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.EXCEPTION_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsExcepInClause) > 0) Then
                lsSQL2 = lsSQL2 & _
                    " AND " & Replace(lsExcepInClause, "A.", "B.")
            Else
                lsSQL2 = lsSQL2 & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
            End If
        Else
            lsSQL2 = lsSQL2 & _
                    " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
        End If
        
        lsSQL2 = lsSQL2 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.EXCEPTION_CD <> ''" & _
            " AND B.HOURS <> 0"
                    
        lsSQL2 = lsSQL2 & " UNION "
        
        lsSQL3 = lsSQL3 & " SELECT E.WG_NUM" & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
                    
        lsSQL3 = lsSQL3 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsWGSelection2, "D.", "E.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.OVER_EXC_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsOvrExcepInClause) > 0) Then
                lsSQL3 = lsSQL3 & _
                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL3 = lsSQL3 & _
                    " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
        End If
        
        lsSQL3 = lsSQL3 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.OVER_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL3 = lsSQL3 & " UNION "
        
        lsSQL4 = lsSQL4 & " SELECT E.WG_NUM" & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL4 = lsSQL4 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsWGSelection2, "D.", "E.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsSchedChgExcepInClause) > 0) Then
                lsSQL4 = lsSQL4 & _
                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL4 = lsSQL4 & _
                    " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
        End If
        
        lsSQL4 = lsSQL4 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.SCHED_CHG_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL4 = lsSQL4 & " UNION "

        lsSQL5 = lsSQL5 & " SELECT D.WG_NUM" & _
            " , A.EMP_NUM, E.EMP_NAME" & _
            " , A.JOB_NUM, A.DATE_WORKED AS CHARGE_DATE" & _
            " , A.START_TIME, A.END_TIME" & _
            " , A.EXCEPTION_CD, '' AS OVER_EXC_CD" & _
            " , '' AS SCHED_CHG_EXC_CD" & _
            " , A.CASE_NUM" & _
            " , A.ADJ_TYPE" & _
            " , CASE WHEN (A.PROCESSED_IND) = 'Y' THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , '' AS OT_OPP_DESC, '' AS DISTR_DEPT_NUM" & _
            " , 0 AS AMOUNT, 0 AS TRAVEL_PAY_HRS, " & _
            " CASE WHEN A.ADJ_TYPE = 'A' THEN (A.HOURS * -1) ELSE A.HOURS END AS HOURS" & _
            " FROM METHDBA.TTR_ADJUSTMENTS A" & _
            " , METHDBA.TTR_EXCEPTION_CDS B" & _
            " , METHDBA.TTR_EMP_WG D, METHDBA.TTR_EMPLOYEE E" & _
            " WHERE " & lsWGSelection2 & _
            " AND A.EMP_NUM = D.EMP_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EXCEPTION_CD = B.EXCEPTION_CD" & _
            " AND A.DATE_WORKED BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.DATE_WORKED BETWEEN D.EFF_START_DATE" & _
            " AND D.EFF_END_DATE"
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If (Len(lsExcepInClause) > 0) Then
                    lsSQL5 = lsSQL5 & _
                        " AND " & lsExcepInClause
                Else
                    lsSQL5 = lsSQL5 & _
                        " AND A.EXCEPTION_CD <> 'ST'"
                End If
            Else
                lsSQL5 = lsSQL5 & _
                        " AND " & lsAllExcepInClause & _
                        " AND A.EXCEPTION_CD <> 'ST'"
            End If

        lsSQL5 = lsSQL5 & _
            " AND D.PRIMARY_IND = 'Y'" & _
            " AND A.HOURS <> 0" & _
            " AND A.EXCEPTION_CD <> ''" & _
            " AND A.EXCEPTION_CD NOT IN ('ET','AT')"
            
        lsSQL5 = lsSQL5 & " UNION "
                
    Else
        'Get from New Table
        lsSQL10 = lsSQL10 & " SELECT E.WG_NUM" & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL10 = lsSQL10 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " INNER JOIN METHDBA.TTR_ADJ_TIME B " & _
            " ON A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND B.EXCEPTION_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsExcepInClause) > 0) Then
                lsSQL10 = lsSQL10 & _
                    " AND " & Replace(lsExcepInClause, "A.", "B.")
            Else
                lsSQL10 = lsSQL10 & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
            End If
        Else
            lsSQL10 = lsSQL10 & _
                    " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
        End If
            
        lsSQL10 = lsSQL10 & " LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
            " ON B.EXCEPTION_CD = C.EXCEPTION_CD" & _
            " LEFT OUTER JOIN METHDBA.TTR_EMP_WG E" & _
            " ON A.EMP_NUM = E.EMP_NUM" & _
            " AND " & Replace(lsWGSelection2, "D.", "E.") & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE" & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE F" & _
            " ON A.EMP_NUM = F.EMP_NUM" & _
            " WHERE COALESCE(E.WG_NUM, 0) <> 0"

        lsSQL10 = lsSQL10 & " UNION "
        
        lsSQL11 = lsSQL11 & " SELECT E.WG_NUM" & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL11 = lsSQL11 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " INNER JOIN METHDBA.TTR_ADJ_TIME B " & _
            " ON A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND B.OVER_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"

        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsOvrExcepInClause) > 0) Then
                lsSQL11 = lsSQL11 & _
                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL11 = lsSQL11 & _
                    " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
        End If
        
        lsSQL11 = lsSQL11 & " LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
            " ON B.OVER_EXC_CD = C.EXCEPTION_CD" & _
            " LEFT OUTER JOIN METHDBA.TTR_EMP_WG E" & _
            " ON A.EMP_NUM = E.EMP_NUM" & _
            " AND " & Replace(lsWGSelection2, "D.", "E.") & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE" & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE F" & _
            " ON A.EMP_NUM = F.EMP_NUM" & _
            " WHERE COALESCE(E.WG_NUM, 0) <> 0"

        lsSQL11 = lsSQL11 & " UNION "
        
        lsSQL12 = lsSQL12 & " SELECT E.WG_NUM" & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL12 = lsSQL12 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " INNER JOIN METHDBA.TTR_ADJ_TIME B " & _
            " ON A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND B.SCHED_CHG_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"

        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsSchedChgExcepInClause) > 0) Then
                lsSQL12 = lsSQL12 & _
                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL12 = lsSQL12 & _
                    " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
        End If
            
        lsSQL12 = lsSQL12 & " LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
            " ON B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
            " LEFT OUTER JOIN METHDBA.TTR_EMP_WG E" & _
            " ON A.EMP_NUM = E.EMP_NUM" & _
            " AND " & Replace(lsWGSelection2, "D.", "E.") & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE" & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE F" & _
            " ON A.EMP_NUM = F.EMP_NUM" & _
            " WHERE COALESCE(E.WG_NUM, 0) <> 0"

        lsSQL12 = lsSQL12 & " UNION "
        
    End If
'END PREM TRIS CONVERSION PHASE 3

'*******************************************************
'*  Retrieve Exceptions from Actual Time
'*******************************************************
    If liActualTimeTables = TrTable_Actual Or _
        liActualTimeTables = TrTable_ActualAndHistory Then
        lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
        lsSQL16 = lsSQL16 _
            & "SELECT A.wg_num, A.emp_num,  d.emp_name, A.job_num, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, " _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL16 = lsSQL16 & "WHERE " & lsWGSelection _
            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.exception_cd "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsExcepInClause) > 0 Then
                    lsSQL16 = lsSQL16 & " AND " & lsExcepInClause
                Else
                    lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
                End If
            Else
                    lsSQL16 = lsSQL16 & " AND " & lsAllExcepInClause
                    lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
            End If
            
        lsSQL16 = lsSQL16 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.EXCEPTION_CD <> ''"
            
        lsSQL16 = lsSQL16 & " UNION "
        
        lsSQL17 = lsSQL17 _
            & "SELECT A.wg_num, A.emp_num, d.emp_name, A.job_num, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, " _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL17 = lsSQL17 & "WHERE " & lsWGSelection _
            & " AND A.charge_date between " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.OVER_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsOvrExcepInClause) > 0 Then
                    lsSQL17 = lsSQL17 & " AND " & lsOvrExcepInClause
                End If
            Else
                lsSQL17 = lsSQL17 & " AND " & lsAllOvrExcepInClause
            End If
            
        lsSQL17 = lsSQL17 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.OVER_EXC_CD <> ''"
            
        lsSQL17 = lsSQL17 & " UNION "
        
        lsSQL18 = lsSQL18 _
            & "SELECT A.wg_num, A.emp_num, d.emp_name, A.job_num, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, " _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL18 = lsSQL18 & "WHERE " & lsWGSelection _
            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsSchedChgExcepInClause) > 0 Then
                    lsSQL18 = lsSQL18 & " AND " & lsSchedChgExcepInClause
                End If
            Else
                lsSQL18 = lsSQL18 & " AND " & lsAllSchedChgExcepInClause
            End If
            
        lsSQL18 = lsSQL18 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.SCHED_CHG_EXC_CD <> ''"
            
        lsSQL18 = lsSQL18 & " UNION "

    End If
    
'*******************************************************
'*  Retrieve Exceptions from Actual Time History
'*******************************************************
    If liActualTimeTables = TrTable_History Or _
       liActualTimeTables = TrTable_ActualAndHistory Then
       lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
       lsSQL19 = lsSQL19 _
            & "SELECT A.wg_num, A.emp_num, d.emp_name, A.job_num, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, '3' AS processed_ind, " _
            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee d "
            
        lsSQL19 = lsSQL19 & " WHERE " & lsWGSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.exception_cd "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsExcepInClause) > 0 Then
                    lsSQL19 = lsSQL19 & " AND " & lsExcepInClause
                Else
                    lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
                End If
            Else
                lsSQL19 = lsSQL19 & " AND " & lsAllExcepInClause
                lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
            End If
        
        lsSQL19 = lsSQL19 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.EXCEPTION_CD <> ''"
            
        lsSQL19 = lsSQL19 & " UNION "
        
        lsSQL20 = lsSQL20 _
            & "SELECT A.wg_num, A.emp_num, d.emp_name, A.job_num, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, '3' AS processed_ind, " _
            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee d "
            
        lsSQL20 = lsSQL20 & " WHERE " & lsWGSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.OVER_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsOvrExcepInClause) > 0 Then
                    lsSQL20 = lsSQL20 & " AND " & lsOvrExcepInClause
                End If
            Else
                lsSQL20 = lsSQL20 & " AND " & lsAllOvrExcepInClause
            End If
        
        lsSQL20 = lsSQL20 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.OVER_EXC_CD <> ''"
            
        lsSQL20 = lsSQL20 & " UNION "
        
        lsSQL21 = lsSQL21 _
            & "SELECT A.wg_num, A.emp_num, d.emp_name, A.job_num, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, '3' AS processed_ind, " _
            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee d "
            
        lsSQL21 = lsSQL21 & " WHERE " & lsWGSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsSchedChgExcepInClause) > 0 Then
                    lsSQL21 = lsSQL21 & " AND " & lsSchedChgExcepInClause
                End If
            Else
                lsSQL21 = lsSQL21 & " AND " & lsAllSchedChgExcepInClause
            End If
                
        
        lsSQL21 = lsSQL21 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.SCHED_CHG_EXC_CD <> ''"
            
        lsSQL21 = lsSQL21 & " UNION "
        
    End If
    
    If (RecordsExist(v_adrsSelectedExceptions)) Then
        If Len(lsExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL2 & lsSQL5 & lsSQL10 & lsSQL16 & lsSQL19
        End If
        If Len(lsOvrExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL3 & lsSQL11 & lsSQL17 & lsSQL20
        End If
        If Len(lsSchedChgExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL4 & lsSQL12 & lsSQL18 & lsSQL21
        End If
        If Len(lsOvrExcepInClause) > 0 Or Len(lsExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL1
        End If
    Else
        lsSQLALL = lsSQLALL & lsSQL1 & lsSQL2 & lsSQL3 & lsSQL4 & lsSQL5 & lsSQL10 & _
            lsSQL11 & lsSQL12 & lsSQL16 & lsSQL17 & lsSQL18 & lsSQL19 & lsSQL20 & lsSQL21
    End If
    
    If Trim(Right(lsSQLALL, 6)) = "UNION" Then
        lsSQLALL = Left(lsSQLALL, Len(lsSQLALL) - 6)
    End If
    
    '//try to elimiate duplicates when selected exceptions are 'E' and 'FCL'
    'sql 16 - use actual time & exception cd 'E'
    'sql 17 - use actual time & override exception cd 'FCL'
    'since these were UNION ALL, there would be 2 rows (of the same actual time) only differing by CASE_GROUP
    '//take CASE_GROUP out of SQL detail and group by, use UNION rather than UNION ALL, then use 3 joins to get the possible CASE_GROUP's
    If Len(lsSQLALL) > 0 Then
        lsSQL = "SELECT WG_NUM, EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, " _
        & " NTE_GROUP.EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, " _
        & " COALESCE(NullIf(EXC1.CASE_GROUP, ''), NullIf(EXC2.CASE_GROUP, ''), NullIf(EXC3.CASE_GROUP, ''), '    ') AS CASE_GROUP, " _
        & " ADJ_TYPE, PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, " _
        & " AMOUNT, TRAVEL_PAY_HRS, " _
        & " HOURS FROM (" & vbCrLf '//NTE_GROUP

        lsSQL = lsSQL & "SELECT WG_NUM, EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, " _
        & " EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, " _
        & " ADJ_TYPE, PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, " _
        & " SUM(AMOUNT) AS AMOUNT, SUM(TRAVEL_PAY_HRS) AS TRAVEL_PAY_HRS, " _
        & " SUM(HOURS) AS HOURS FROM (" _
        & lsSQLALL & ") NTE " & vbCrLf '//detail
        
        lsSQL = lsSQL & " GROUP BY WG_NUM, EMP_NUM, EMP_NAME, JOB_NUM, " _
        & " CHARGE_DATE, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, " _
        & " ADJ_TYPE, PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM ) NTE_GROUP" _
        & " " & vbCrLf _
        & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC1" _
        & " ON NTE_GROUP.EXCEPTION_CD = EXC1.EXCEPTION_CD " _
        & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC2 " _
        & " ON NTE_GROUP.OVER_EXC_CD = EXC2.EXCEPTION_CD " _
        & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC3" _
        & " ON NTE_GROUP.SCHED_CHG_EXC_CD = EXC3.EXCEPTION_CD "
            
        lsSQL = lsSQL & " ORDER BY wg_num, emp_num, processed_ind, charge_date"
    End If
        
    Set lobjCN = New ADODB.Connection
    Set lobjCM = New clsuDBConnectMgr
    lobjCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
    Set lobjCM = Nothing
    lobjCN.Open

    Set lobjRS = New ADODB.Recordset
    lobjRS.CursorLocation = adUseClient
    lobjRS.Open lsSQL, lobjCN, adOpenStatic, adLockBatchOptimistic
    lobjRS.ActiveConnection = Nothing
    
    lobjCN.Close
    Set lobjCN = Nothing

    Set GetWGExceptions = lobjRS
    Set lobjRS = Nothing

    '//-----clean up class variables used by madrsExceptions/GetExceptionCodes/IsExceptionCode-----//
    DestroyState

    Exit Function
eh:
    Set lobjCM = Nothing
    Set lobjCN = Nothing
    Set lobjRS = Nothing

    DestroyState
    
    TrsRaiseError m_strCMODULENAME _
        , "GetWGExceptions" _
        , , , , DescArgsFrom(lsSQL) _
        , True

End Function

'Public Function GetWGExceptionsOLD(ByVal v_aryWorkgroups As Variant _
'                , ByVal v_datStartDate As Date _
'                , ByVal v_datEndDate As Date _
'                , ByVal v_adrsSelectedExceptions As ADODB.Recordset _
'                , Optional v_datLastClosedPP As Date = CDate("1980-01-01") _
'                , Optional v_datEndOfOpenPP As Date = CDate("1980-01-01") _
'                ) As ADODB.Recordset
'On Error GoTo eh
'Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
'Dim lobjCN As ADODB.Connection
'Dim lobjRS As ADODB.Recordset
'Dim lsSQL As String
'Dim lsWGSelection As String, lsWGSelection2 As String
'Dim ldtHistoryEndDate As Date, ldtSchedExcStartDate As Date
'Dim ldtActualTimeStartDate As Date, ldtActualTimeEndDate As Date
'Dim lsInClausePartValueKey As String, liActualTimeTables As TrActualTimeTable
'Dim lsExcepInClause As String, liRecCounter As Integer
'Dim lsOvrExcepInClause As String
'Dim lsSchedChgExcepInClause As String
'Dim lsExcepandOvrInClause As String
'Dim lsAllExcepInClause As String
'Dim lsAllOvrExcepInClause As String
'Dim lsAllSchedChgExcepInClause As String
'Dim lsAllExcepandOvrInClause As String
'Dim lblnValidExcep As Boolean
'Dim lblnFirstExcep As Boolean
'Dim lsSQL1 As String
'Dim lsSQL2 As String
'Dim lsSQL3 As String
'Dim lsSQL4 As String
'Dim lsSQL5 As String
'Dim lsSQL10 As String
'Dim lsSQL11 As String
'Dim lsSQL12 As String
'Dim lsSQL16 As String
'Dim lsSQL17 As String
'Dim lsSQL18 As String
'Dim lsSQL19 As String
'Dim lsSQL20 As String
'Dim lsSQL21 As String
'Dim lsSQLALL As String
'
'    If (IsArray(v_aryWorkgroups)) Then
'        lsWGSelection = BuildInClause("A.WG_NUM", v_aryWorkgroups, False) & " "
'        lsWGSelection2 = Replace(lsWGSelection, "A.WG_NUM", "D.WG_NUM")
'    Else
'        lsWGSelection = "A.WG_NUM = " & CStr(CLng(v_aryWorkgroups)) & " "
'        lsWGSelection2 = Replace(lsWGSelection, "A.WG_NUM", "D.WG_NUM")
'    End If
'
'    liActualTimeTables = DetermineActualTimeTables(v_datStartDate _
'                            , v_datEndDate _
'                            , ldtHistoryEndDate _
'                            , ldtActualTimeStartDate)
'
'    If v_datLastClosedPP = #1/1/1980# Then
'        ldtSchedExcStartDate = v_datStartDate
'    Else
'        If (v_datLastClosedPP > ldtActualTimeStartDate) Then
'            ldtSchedExcStartDate = DateAdd("d", 1, v_datLastClosedPP)
'        Else
'            ldtSchedExcStartDate = ldtActualTimeStartDate
'        End If
'    End If
'    If v_datEndOfOpenPP = #1/1/1980# Then
'        ldtActualTimeEndDate = v_datEndDate
'    Else
'        If (v_datEndOfOpenPP > ldtActualTimeStartDate) Then
'            ldtActualTimeEndDate = v_datEndOfOpenPP
'        Else
'            ldtActualTimeEndDate = v_datEndDate
'        End If
'    End If
'
'    GetExceptionCodes
'    GetOverExcepCodes
'    GetSchedChgExcepCodes
'
'    If (RecordsExist(v_adrsSelectedExceptions)) Then
'
'        With v_adrsSelectedExceptions
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsExceptionCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsExcepInClause = lsExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsExcepInClause = lsExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsExcepInClause = " A.EXCEPTION_CD IN ( " & lsExcepInClause & ")"
'            End If
'
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsOverExcepCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsOvrExcepInClause = lsOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsOvrExcepInClause = lsOvrExcepInClause & _
'                            ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsOvrExcepInClause & _
'                ")"
'            End If
'
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsSchedChgExcepCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
'                        "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
'                lsSchedChgExcepInClause & ")"
'            End If
'
'        End With
'
'    Else
'
'        If RecordsExist(madrsExceptions) Then
'            With madrsExceptions
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllExcepInClause = lsAllExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllExcepInClause = lsAllExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllExcepInClause = " A.EXCEPTION_CD IN ( " & lsAllExcepInClause & ")"
'            End With
'        End If
'
'        If RecordsExist(madrsOverExcep) Then
'            With madrsOverExcep
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & _
'                            ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsAllOvrExcepInClause & ")"
'            End With
'        End If
'
'        If RecordsExist(madrsSchedChgExcep) Then
'            With madrsSchedChgExcep
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
'                        "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
'                lsAllSchedChgExcepInClause & ")"
'            End With
'        End If
'
'    End If
'
''*******************************************************
''*  Retrieve Scheduled Exceptions
''*******************************************************
'
'    lsSQL1 = lsSQL1 & "SELECT d.wg_num, a.emp_num, e.emp_name, a.job_num, " _
'        & "a.charge_date, '00:00' as start_time, '00:00' as end_time, a.exception_cd, " _
'        & "a.over_exc_cd, '' as sched_chg_exc_cd, a.case_num, b.case_group, " _
'        & "'C' as adj_type, '1' AS processed_ind, '' AS OT_OPP_DESC, " _
'        & " '' as distr_dept_num, 0 as amount, 0 as travel_pay_hrs, hours " _
'        & "FROM methdba.ttr_emp_sched_exc A, " _
'        & " methdba.ttr_exception_cds B, " _
'        & " methdba.ttr_emp_wg D , methdba.ttr_employee E " _
'        & " WHERE " & lsWGSelection2 _
'        & " AND A.charge_date between " & DateAsEnclosedString(ldtSchedExcStartDate) & " AND " _
'                & DateAsEnclosedString(v_datEndDate) _
'        & " AND A.emp_num = D.emp_num " _
'        & " AND A.emp_num = E.emp_num " _
'        & " AND A.charge_date between D.eff_start_date AND D.eff_end_date " _
'        & " AND B.exception_cd = A.exception_cd " _
'        & " AND D.PRIMARY_IND = 'Y' "
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) > 0 Then
'                lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & " OR " & _
'                lsOvrExcepInClause & ") "
'            Else
'                If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) = 0 Then
'                    lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & ") "
'                Else
'                    If Len(lsExcepInClause) = 0 And Len(lsOvrExcepInClause) > 0 Then
'                        lsSQL1 = lsSQL1 & " AND (" & lsOvrExcepInClause & ") "
'                    Else
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                        " OR A.OVER_EXC_CD <> '') "
'                    End If
'                End If
'            End If
'        Else
'            If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) > 0 Then
'                lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & " OR " & _
'                    lsAllOvrExcepInClause & ") "
'                lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'            Else
'                If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) = 0 Then
'                    lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & ") "
'                    lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                Else
'                    If Len(lsAllExcepInClause) = 0 And Len(lsAllOvrExcepInClause) > 0 Then
'                        lsSQL1 = lsSQL1 & " AND (" & lsAllOvrExcepInClause & ") "
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                    Else
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                    End If
'                End If
'            End If
'        End If
'
'        lsSQL1 = lsSQL1 _
'        & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'        & " AND A.exception_cd NOT IN ( SELECT exception_cd  " _
'                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5 ) " _
'        & " AND A.over_exc_cd NOT IN ( SELECT exception_cd  " _
'                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5) "
'
''BEGIN PREM TRIS CONVERSION PHASE 3
'    lsSQL1 = lsSQL1 & " UNION ALL "
'
'    If (v_datStartDate < #1/27/2005#) Then
'        'Get from Old and New Tables
'        lsSQL2 = lsSQL2 & " SELECT E.WG_NUM" & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL2 = lsSQL2 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsWGSelection2, "D.", "E.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.EXCEPTION_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsExcepInClause) > 0) Then
'                lsSQL2 = lsSQL2 & _
'                    " AND " & Replace(lsExcepInClause, "A.", "B.")
'            Else
'                lsSQL2 = lsSQL2 & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'            End If
'        Else
'            lsSQL2 = lsSQL2 & _
'                    " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'        End If
'
'        lsSQL2 = lsSQL2 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.EXCEPTION_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL2 = lsSQL2 & " UNION ALL "
'
'        lsSQL3 = lsSQL3 & " SELECT E.WG_NUM" & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL3 = lsSQL3 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsWGSelection2, "D.", "E.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.OVER_EXC_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsOvrExcepInClause) > 0) Then
'                lsSQL3 = lsSQL3 & _
'                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL3 = lsSQL3 & _
'                    " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL3 = lsSQL3 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.OVER_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL3 = lsSQL3 & " UNION ALL "
'
'        lsSQL4 = lsSQL4 & " SELECT E.WG_NUM" & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL4 = lsSQL4 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsWGSelection2, "D.", "E.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsSchedChgExcepInClause) > 0) Then
'                lsSQL4 = lsSQL4 & _
'                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL4 = lsSQL4 & _
'                    " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL4 = lsSQL4 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.SCHED_CHG_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL4 = lsSQL4 & " UNION ALL "
'
'        lsSQL5 = lsSQL5 & " SELECT D.WG_NUM" & _
'            " , A.EMP_NUM, E.EMP_NAME" & _
'            " , A.JOB_NUM, A.DATE_WORKED AS CHARGE_DATE" & _
'            " , A.START_TIME, A.END_TIME" & _
'            " , A.EXCEPTION_CD, '' AS OVER_EXC_CD" & _
'            " , '' AS SCHED_CHG_EXC_CD" & _
'            " , A.CASE_NUM, B.CASE_GROUP" & _
'            " , A.ADJ_TYPE" & _
'            " , CASE WHEN (A.PROCESSED_IND) = 'Y' THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , '' AS OT_OPP_DESC, '' AS DISTR_DEPT_NUM" & _
'            " , 0 AS AMOUNT, 0 AS TRAVEL_PAY_HRS, " & _
'            " CASE WHEN A.ADJ_TYPE = 'A' THEN (A.HOURS * -1) ELSE A.HOURS END AS HOURS" & _
'            " FROM METHDBA.TTR_ADJUSTMENTS A" & _
'            " , METHDBA.TTR_EXCEPTION_CDS B" & _
'            " , METHDBA.TTR_EMP_WG D, METHDBA.TTR_EMPLOYEE E" & _
'            " WHERE " & lsWGSelection2 & _
'            " AND A.EMP_NUM = D.EMP_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EXCEPTION_CD = B.EXCEPTION_CD" & _
'            " AND A.DATE_WORKED BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.DATE_WORKED BETWEEN D.EFF_START_DATE" & _
'            " AND D.EFF_END_DATE"
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If (Len(lsExcepInClause) > 0) Then
'                    lsSQL5 = lsSQL5 & _
'                        " AND " & lsExcepInClause
'                Else
'                    lsSQL5 = lsSQL5 & _
'                        " AND A.EXCEPTION_CD <> 'ST'"
'                End If
'            Else
'                lsSQL5 = lsSQL5 & _
'                        " AND " & lsAllExcepInClause & _
'                        " AND A.EXCEPTION_CD <> 'ST'"
'            End If
'
'        lsSQL5 = lsSQL5 & _
'            " AND D.PRIMARY_IND = 'Y'" & _
'            " AND A.HOURS <> 0" & _
'            " AND A.EXCEPTION_CD <> ''" & _
'            " AND A.EXCEPTION_CD NOT IN ('ET','AT')"
'
'        lsSQL5 = lsSQL5 & " UNION ALL "
'
'    Else
'        'Get from New Table
'        lsSQL10 = lsSQL10 & " SELECT E.WG_NUM" & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL10 = lsSQL10 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " INNER JOIN METHDBA.TTR_ADJ_TIME B " & _
'            " ON A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND B.EXCEPTION_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsExcepInClause) > 0) Then
'                lsSQL10 = lsSQL10 & _
'                    " AND " & Replace(lsExcepInClause, "A.", "B.")
'            Else
'                lsSQL10 = lsSQL10 & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'            End If
'        Else
'            lsSQL10 = lsSQL10 & _
'                    " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'        End If
'
'        lsSQL10 = lsSQL10 & " LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
'            " ON B.EXCEPTION_CD = C.EXCEPTION_CD" & _
'            " LEFT OUTER JOIN METHDBA.TTR_EMP_WG E" & _
'            " ON A.EMP_NUM = E.EMP_NUM" & _
'            " AND " & Replace(lsWGSelection2, "D.", "E.") & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE" & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE F" & _
'            " ON A.EMP_NUM = F.EMP_NUM" & _
'            " WHERE COALESCE(E.WG_NUM, 0) <> 0"
'
'        lsSQL10 = lsSQL10 & " UNION ALL "
'
'        lsSQL11 = lsSQL11 & " SELECT E.WG_NUM" & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL11 = lsSQL11 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " INNER JOIN METHDBA.TTR_ADJ_TIME B " & _
'            " ON A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND B.OVER_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsOvrExcepInClause) > 0) Then
'                lsSQL11 = lsSQL11 & _
'                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL11 = lsSQL11 & _
'                    " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL11 = lsSQL11 & " LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
'            " ON B.OVER_EXC_CD = C.EXCEPTION_CD" & _
'            " LEFT OUTER JOIN METHDBA.TTR_EMP_WG E" & _
'            " ON A.EMP_NUM = E.EMP_NUM" & _
'            " AND " & Replace(lsWGSelection2, "D.", "E.") & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE" & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE F" & _
'            " ON A.EMP_NUM = F.EMP_NUM" & _
'            " WHERE COALESCE(E.WG_NUM, 0) <> 0"
'
'        lsSQL11 = lsSQL11 & " UNION ALL "
'
'        lsSQL12 = lsSQL12 & " SELECT E.WG_NUM" & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL12 = lsSQL12 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " INNER JOIN METHDBA.TTR_ADJ_TIME B " & _
'            " ON A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND B.SCHED_CHG_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsSchedChgExcepInClause) > 0) Then
'                lsSQL12 = lsSQL12 & _
'                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL12 = lsSQL12 & _
'                    " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL12 = lsSQL12 & " LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
'            " ON B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
'            " LEFT OUTER JOIN METHDBA.TTR_EMP_WG E" & _
'            " ON A.EMP_NUM = E.EMP_NUM" & _
'            " AND " & Replace(lsWGSelection2, "D.", "E.") & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE" & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE F" & _
'            " ON A.EMP_NUM = F.EMP_NUM" & _
'            " WHERE COALESCE(E.WG_NUM, 0) <> 0"
'
'        lsSQL12 = lsSQL12 & " UNION ALL "
'
'    End If
''END PREM TRIS CONVERSION PHASE 3
'
''*******************************************************
''*  Retrieve Exceptions from Actual Time
''*******************************************************
'    If liActualTimeTables = TrTable_Actual Or _
'        liActualTimeTables = TrTable_ActualAndHistory Then
'        lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
'        lsSQL16 = lsSQL16 _
'            & "SELECT A.wg_num, A.emp_num,  d.emp_name, A.job_num, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, " _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL16 = lsSQL16 & "WHERE " & lsWGSelection _
'            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.exception_cd "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsExcepInClause) > 0 Then
'                    lsSQL16 = lsSQL16 & " AND " & lsExcepInClause
'                Else
'                    lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'                End If
'            Else
'                    lsSQL16 = lsSQL16 & " AND " & lsAllExcepInClause
'                    lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'            End If
'
'        lsSQL16 = lsSQL16 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.EXCEPTION_CD <> ''"
'
'        lsSQL16 = lsSQL16 & " UNION ALL "
'
'        lsSQL17 = lsSQL17 _
'            & "SELECT A.wg_num, A.emp_num, d.emp_name, A.job_num, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, " _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL17 = lsSQL17 & "WHERE " & lsWGSelection _
'            & " AND A.charge_date between " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.OVER_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsOvrExcepInClause) > 0 Then
'                    lsSQL17 = lsSQL17 & " AND " & lsOvrExcepInClause
'                End If
'            Else
'                lsSQL17 = lsSQL17 & " AND " & lsAllOvrExcepInClause
'            End If
'
'        lsSQL17 = lsSQL17 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.OVER_EXC_CD <> ''"
'
'        lsSQL17 = lsSQL17 & " UNION ALL "
'
'        lsSQL18 = lsSQL18 _
'            & "SELECT A.wg_num, A.emp_num, d.emp_name, A.job_num, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, " _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL18 = lsSQL18 & "WHERE " & lsWGSelection _
'            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsSchedChgExcepInClause) > 0 Then
'                    lsSQL18 = lsSQL18 & " AND " & lsSchedChgExcepInClause
'                End If
'            Else
'                lsSQL18 = lsSQL18 & " AND " & lsAllSchedChgExcepInClause
'            End If
'
'        lsSQL18 = lsSQL18 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.SCHED_CHG_EXC_CD <> ''"
'
'        lsSQL18 = lsSQL18 & " UNION ALL "
'
'    End If
'
''*******************************************************
''*  Retrieve Exceptions from Actual Time History
''*******************************************************
'    If liActualTimeTables = TrTable_History Or _
'       liActualTimeTables = TrTable_ActualAndHistory Then
'       lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
'       lsSQL19 = lsSQL19 _
'            & "SELECT A.wg_num, A.emp_num, d.emp_name, A.job_num, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, '3' AS processed_ind, " _
'            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee d "
'
'        lsSQL19 = lsSQL19 & " WHERE " & lsWGSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.exception_cd "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsExcepInClause) > 0 Then
'                    lsSQL19 = lsSQL19 & " AND " & lsExcepInClause
'                Else
'                    lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'                End If
'            Else
'                lsSQL19 = lsSQL19 & " AND " & lsAllExcepInClause
'                lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'            End If
'
'        lsSQL19 = lsSQL19 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.EXCEPTION_CD <> ''"
'
'        lsSQL19 = lsSQL19 & " UNION ALL "
'
'        lsSQL20 = lsSQL20 _
'            & "SELECT A.wg_num, A.emp_num, d.emp_name, A.job_num, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, '3' AS processed_ind, " _
'            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee d "
'
'        lsSQL20 = lsSQL20 & " WHERE " & lsWGSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.OVER_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsOvrExcepInClause) > 0 Then
'                    lsSQL20 = lsSQL20 & " AND " & lsOvrExcepInClause
'                End If
'            Else
'                lsSQL20 = lsSQL20 & " AND " & lsAllOvrExcepInClause
'            End If
'
'        lsSQL20 = lsSQL20 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.OVER_EXC_CD <> ''"
'
'        lsSQL20 = lsSQL20 & " UNION ALL "
'
'        lsSQL21 = lsSQL21 _
'            & "SELECT A.wg_num, A.emp_num, d.emp_name, A.job_num, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, '3' AS processed_ind, " _
'            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee d "
'
'        lsSQL21 = lsSQL21 & " WHERE " & lsWGSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsSchedChgExcepInClause) > 0 Then
'                    lsSQL21 = lsSQL21 & " AND " & lsSchedChgExcepInClause
'                End If
'            Else
'                lsSQL21 = lsSQL21 & " AND " & lsAllSchedChgExcepInClause
'            End If
'
'
'        lsSQL21 = lsSQL21 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.SCHED_CHG_EXC_CD <> ''"
'
'        lsSQL21 = lsSQL21 & " UNION ALL "
'
'    End If
'
'    If (RecordsExist(v_adrsSelectedExceptions)) Then
'        If Len(lsExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL2 & lsSQL5 & lsSQL10 & lsSQL16 & lsSQL19
'        End If
'        If Len(lsOvrExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL3 & lsSQL11 & lsSQL17 & lsSQL20
'        End If
'        If Len(lsSchedChgExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL4 & lsSQL12 & lsSQL18 & lsSQL21
'        End If
'        If Len(lsOvrExcepInClause) > 0 Or Len(lsExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL1
'        End If
'    Else
'        lsSQLALL = lsSQLALL & lsSQL1 & lsSQL2 & lsSQL3 & lsSQL4 & lsSQL5 & lsSQL10 & _
'            lsSQL11 & lsSQL12 & lsSQL16 & lsSQL17 & lsSQL18 & lsSQL19 & lsSQL20 & lsSQL21
'    End If
'
'    If Trim(Right(lsSQLALL, 10)) = "UNION ALL" Then
'        lsSQLALL = Left(lsSQLALL, Len(lsSQLALL) - 10)
'    End If
'
'    If Len(lsSQLALL) > 0 Then
'        lsSQL = "SELECT WG_NUM, EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, " _
'        & " EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM , CASE_GROUP, " _
'        & " ADJ_TYPE, PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, " _
'        & " SUM(AMOUNT) AS AMOUNT, SUM(TRAVEL_PAY_HRS) AS TRAVEL_PAY_HRS, " _
'        & " SUM(HOURS) AS HOURS FROM (" _
'        & lsSQLALL & ") NTE "
'        lsSQL = lsSQL & " GROUP BY WG_NUM, EMP_NUM, EMP_NAME, JOB_NUM, " _
'        & " CHARGE_DATE, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, " _
'        & " CASE_GROUP, ADJ_TYPE, PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM "
'        lsSQL = lsSQL & " ORDER BY wg_num, emp_num, processed_ind, charge_date"
'    End If
'
'    Set lobjCN = New ADODB.Connection
'    Set lobjCM = New clsuDBConnectMgr
'    lobjCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
'    Set lobjCM = Nothing
'    lobjCN.Open
'
'    Set lobjRS = New ADODB.Recordset
'    lobjRS.CursorLocation = adUseClient
'    lobjRS.Open lsSQL, lobjCN, adOpenStatic, adLockBatchOptimistic
'    lobjRS.ActiveConnection = Nothing
'
'    lobjCN.Close
'    Set lobjCN = Nothing
'
'    Set GetWGExceptionsOLD = lobjRS
'    Set lobjRS = Nothing
'
'    '//-----clean up class variables used by madrsExceptions/GetExceptionCodes/IsExceptionCode-----//
'    DestroyState
'
'    Exit Function
'eh:
'    Set lobjCM = Nothing
'    Set lobjCN = Nothing
'    Set lobjRS = Nothing
'
'    DestroyState
'
'    TrsRaiseError m_strCMODULENAME _
'        , "GetWGExceptionsOLD" _
'        , , , , DescArgsFrom(lsSQL) _
'        , True
'
'End Function


Public Function GetDeptExceptions(ByVal v_aryDepartments As Variant _
                , ByVal v_datStartDate As Date _
                , ByVal v_datEndDate As Date _
                , ByVal v_adrsSelectedExceptions As ADODB.Recordset _
                , Optional v_datLastClosedPP As Date = CDate("1980-01-01") _
                , Optional v_datEndOfOpenPP As Date = CDate("1980-01-01") _
                ) As ADODB.Recordset
On Error GoTo eh
Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
Dim lobjCN As ADODB.Connection
Dim lobjRS As ADODB.Recordset
Dim lsSQL As String
Dim lsDeptSelection As String, lsDeptSelection2 As String
Dim ldtHistoryEndDate As Date, ldtSchedExcStartDate As Date
Dim ldtActualTimeStartDate As Date, ldtActualTimeEndDate As Date
Dim lsInClausePartValueKey As String, lsExcepInClause As String
Dim lsOvrExcepInClause As String
Dim lsSchedChgExcepInClause As String
Dim lsExcepandOvrInClause As String
Dim lsAllExcepInClause As String
Dim lsAllOvrExcepInClause As String
Dim lsAllSchedChgExcepInClause As String
Dim lsAllExcepandOvrInClause As String
Dim liActualTimeTables As TrActualTimeTable, liRecCounter As Integer
Dim lblnValidExcep As Boolean
Dim lblnFirstExcep As Boolean
Dim lsSQL1 As String
Dim lsSQL2 As String
Dim lsSQL3 As String
Dim lsSQL4 As String
Dim lsSQL5 As String
Dim lsSQL10 As String
Dim lsSQL11 As String
Dim lsSQL12 As String
Dim lsSQL16 As String
Dim lsSQL17 As String
Dim lsSQL18 As String
Dim lsSQL19 As String
Dim lsSQL20 As String
Dim lsSQL21 As String
Dim lsSQLALL As String

    If (IsArray(v_aryDepartments)) Then
        lsDeptSelection = BuildInClause("A.DEPT_NUM", v_aryDepartments, True) & " "
        lsDeptSelection2 = Replace(lsDeptSelection, "A.DEPT_NUM", "D.DEPT_NUM")
    Else
        lsDeptSelection = "A.DEPT_NUM = '" & CStr(CLng(v_aryDepartments)) & "'"
        lsDeptSelection2 = Replace(lsDeptSelection, "A.DEPT_NUM", "D.DEPT_NUM")
    End If

    liActualTimeTables = DetermineActualTimeTables(v_datStartDate _
                            , v_datEndDate _
                            , ldtHistoryEndDate _
                            , ldtActualTimeStartDate)
                            
    If v_datLastClosedPP = #1/1/1980# Then
        ldtSchedExcStartDate = v_datStartDate
    Else
        If (v_datLastClosedPP > ldtActualTimeStartDate) Then
            ldtSchedExcStartDate = DateAdd("d", 1, v_datLastClosedPP)
        Else
            ldtSchedExcStartDate = ldtActualTimeStartDate
        End If
    End If
    If v_datEndOfOpenPP = #1/1/1980# Then
        ldtActualTimeEndDate = v_datEndDate
    Else
        If (v_datEndOfOpenPP > ldtActualTimeStartDate) Then
            ldtActualTimeEndDate = v_datEndOfOpenPP
        Else
            ldtActualTimeEndDate = v_datEndDate
        End If
    End If

    GetExceptionCodes
    GetOverExcepCodes
    GetSchedChgExcepCodes
    
    If (RecordsExist(v_adrsSelectedExceptions)) Then
        With v_adrsSelectedExceptions
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsExceptionCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsExcepInClause = lsExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsExcepInClause = lsExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If

                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsExcepInClause = " A.EXCEPTION_CD IN ( " & lsExcepInClause & ")"
            End If
            
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsOverExcepCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsOvrExcepInClause = lsOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsOvrExcepInClause = lsOvrExcepInClause & _
                            ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If
                
                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsOvrExcepInClause & _
                ")"
            End If
            
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsSchedChgExcepCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
                        "'" & !EXCEPTION_CD & "'"
                    Else
                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If
                
                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
                lsSchedChgExcepInClause & ")"
            End If
            
        End With
        
    Else
    
        If RecordsExist(madrsExceptions) Then
            With madrsExceptions
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllExcepInClause = lsAllExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllExcepInClause = lsAllExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
                
                lsAllExcepInClause = " A.EXCEPTION_CD IN ( " & lsAllExcepInClause & ")"
            End With
        End If
        
        If RecordsExist(madrsOverExcep) Then
            With madrsOverExcep
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & _
                            ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
                
                lsAllOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsAllOvrExcepInClause & ")"
            End With
        End If
        
        If RecordsExist(madrsSchedChgExcep) Then
            With madrsSchedChgExcep
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
                        "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
            
                lsAllSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
                lsAllSchedChgExcepInClause & ")"
            End With
        End If
        
    End If

'*******************************************************
'*  Retrieve Scheduled Exceptions
'*******************************************************
        
    lsSQL1 = lsSQL1 & "SELECT D.DEPT_NUM, a.emp_num, d.emp_name,a.job_num, a.charge_date, " _
        & " '00:00' as start_time, '00:00' as end_time, a.exception_cd, a.over_exc_cd, " _
        & " '' as sched_chg_exc_cd, a.case_num, 'C' as adj_type, " _
        & " '1' AS processed_ind, '' AS OT_OPP_DESC,'' as distr_dept_num, " _
        & " 0 as amount, 0 as travel_pay_hrs, hours  " _
        & " FROM methdba.ttr_emp_sched_exc A, " _
        & " methdba.ttr_exception_cds B, " _
        & " methdba.ttr_employee D, methdba.ttr_emp_wg E " _
        & " WHERE " & lsDeptSelection2 _
        & " AND A.charge_date between " & DateAsEnclosedString(ldtSchedExcStartDate) & " AND " _
                & DateAsEnclosedString(v_datEndDate) _
        & " AND A.emp_num = D.emp_num " _
        & " AND B.exception_cd = A.exception_cd " _
        & " AND A.emp_num = E.emp_num " _
        & " AND E.PRIMARY_IND = 'Y' " _
        & " AND A.charge_date between e.eff_start_date AND e.eff_end_date "
        
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) > 0 Then
                lsSQL1 = lsSQL1 & "AND (" & lsExcepInClause & " OR " & _
                lsOvrExcepInClause & ") "
            Else
                If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) = 0 Then
                    lsSQL1 = lsSQL1 & "AND (" & lsExcepInClause & ") "
                Else
                    If Len(lsExcepInClause) = 0 And Len(lsOvrExcepInClause) > 0 Then
                        lsSQL1 = lsSQL1 & "AND (" & lsOvrExcepInClause & ") "
                    Else
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                        " OR A.OVER_EXC_CD <> '') "
                    End If
                End If
            End If
        Else
            If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) > 0 Then
                lsSQL1 = lsSQL1 & "AND (" & lsAllExcepInClause & " OR " & _
                    lsAllOvrExcepInClause & ") "
                lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
            Else
                If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) = 0 Then
                    lsSQL1 = lsSQL1 & "AND (" & lsAllExcepInClause & ") "
                    lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                Else
                    If Len(lsAllExcepInClause) = 0 And Len(lsAllOvrExcepInClause) > 0 Then
                        lsSQL1 = lsSQL1 & "AND (" & lsAllOvrExcepInClause & ") "
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                    Else
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                    End If
                End If
            End If
        End If

        
    lsSQL1 = lsSQL1 _
        & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
        & " AND (A.OVER_EXC_CD <> '' OR A.EXCEPTION_CD <> 'ST') "
'BEGIN PREM TRIS CONVERSION PHASE 3
    lsSQL1 = lsSQL1 & " UNION "

    If (v_datStartDate < #1/27/2005#) Then
        'Get from Old and New Tables
        lsSQL2 = lsSQL2 & " SELECT    F.DEPT_NUM " & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC, B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL2 = lsSQL2 & " FROM  METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM " & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.EXCEPTION_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
        
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsExcepInClause) > 0) Then
                lsSQL2 = lsSQL2 & _
                    " AND " & Replace(lsExcepInClause, "A.", "B.")
            Else
                lsSQL2 = lsSQL2 & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
            End If
        Else
            lsSQL2 = lsSQL2 & _
                " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
                " AND B.EXCEPTION_CD <> 'ST'" & _
                " AND B.OVER_EXC_CD = ''" & _
                " AND B.SCHED_CHG_EXC_CD = ''"
        End If
        
        lsSQL2 = lsSQL2 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.EXCEPTION_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL2 = lsSQL2 & " UNION "
        
        lsSQL3 = lsSQL3 & " SELECT F.DEPT_NUM " & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
        
        lsSQL3 = lsSQL3 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM " & _
            " , METHDBA.TTR_EXCEPTION_CDS C " & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.OVER_EXC_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsOvrExcepInClause) > 0) Then
                lsSQL3 = lsSQL3 & _
                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL3 = lsSQL3 & _
                " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
        End If
        
        lsSQL3 = lsSQL3 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.OVER_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL3 = lsSQL3 & " UNION "
    
        lsSQL4 = lsSQL4 & " SELECT F.DEPT_NUM " & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL4 = lsSQL4 & " FROM METHDBA.TTR_ADJ_RQT A " & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C " & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsSchedChgExcepInClause) > 0) Then
                lsSQL4 = lsSQL4 & _
                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL4 = lsSQL4 & _
                " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
        End If
        
        lsSQL4 = lsSQL4 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.SCHED_CHG_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL4 = lsSQL4 & " UNION "

        lsSQL5 = lsSQL5 & " SELECT E.DEPT_NUM " & _
            " , A.EMP_NUM, E.EMP_NAME" & _
            " , A.JOB_NUM, A.DATE_WORKED AS CHARGE_DATE" & _
            " , A.START_TIME, A.END_TIME" & _
            " , A.EXCEPTION_CD, '' AS OVER_EXC_CD" & _
            " , '' AS SCHED_CHG_EXC_CD" & _
            " , A.CASE_NUM" & _
            " , A.ADJ_TYPE" & _
            " , CASE WHEN (A.PROCESSED_IND) = 'Y' THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , '' AS OT_OPP_DESC, '' AS DISTR_DEPT_NUM" & _
            " , 0 AS AMOUNT, 0 AS TRAVEL_PAY_HRS" & _
            " , CASE WHEN A.ADJ_TYPE = 'A' THEN (A.HOURS * -1) ELSE A.HOURS END AS HOURS" & _
            " FROM  METHDBA.TTR_ADJUSTMENTS A" & _
            " , METHDBA.TTR_EXCEPTION_CDS B" & _
            " , METHDBA.TTR_EMP_WG D, METHDBA.TTR_EMPLOYEE E" & _
            " WHERE " & Replace(lsDeptSelection2, "D.", "E.") & _
            " AND A.EMP_NUM = D.EMP_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EXCEPTION_CD = B.EXCEPTION_CD" & _
            " AND A.DATE_WORKED BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.DATE_WORKED BETWEEN D.EFF_START_DATE" & _
            " AND D.EFF_END_DATE"
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If (Len(lsExcepInClause) > 0) Then
                    lsSQL5 = lsSQL5 & _
                        " AND " & lsExcepInClause
                Else
                    lsSQL5 = lsSQL5 & _
                        " AND A.EXCEPTION_CD <> 'ST'"
                End If
            Else
                lsSQL5 = lsSQL5 & _
                    " AND " & lsAllExcepInClause & _
                    " AND A.EXCEPTION_CD <> 'ST'"
            End If

        lsSQL5 = lsSQL5 & _
            " AND D.PRIMARY_IND = 'Y'" & _
            " AND A.HOURS <> 0" & _
            " AND A.EXCEPTION_CD NOT IN ('ET','AT')" & _
            " AND A.EXCEPTION_CD <> ''"
            
        lsSQL5 = lsSQL5 & " UNION "
        
    Else
        'Get from New Table
        lsSQL10 = lsSQL10 & " SELECT F.DEPT_NUM " & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL10 = lsSQL10 & " FROM  METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C " & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.EXCEPTION_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsExcepInClause) > 0) Then
                lsSQL10 = lsSQL10 & _
                    " AND " & Replace(lsExcepInClause, "A.", "B.")
            Else
                lsSQL10 = lsSQL10 & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
            End If
        Else
            lsSQL10 = lsSQL10 & _
                    " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
        End If
        
        lsSQL10 = lsSQL10 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.EXCEPTION_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL10 = lsSQL10 & " UNION "
        
        lsSQL11 = lsSQL11 & " SELECT F.DEPT_NUM " & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
                   
        lsSQL11 = lsSQL11 & " FROM  METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C " & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.OVER_EXC_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsOvrExcepInClause) > 0) Then
                lsSQL11 = lsSQL11 & _
                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL11 = lsSQL11 & _
                    " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
        End If
        
        lsSQL11 = lsSQL11 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.OVER_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL11 = lsSQL11 & " UNION "
        
        lsSQL12 = lsSQL12 & " SELECT F.DEPT_NUM " & _
            " , A.EMP_NUM, F.EMP_NAME" & _
            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL12 = lsSQL12 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C " & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsSchedChgExcepInClause) > 0) Then
                lsSQL12 = lsSQL12 & _
                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL12 = lsSQL12 & _
                    " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
        End If
        
        lsSQL12 = lsSQL12 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.SCHED_CHG_EXC_CD <> ''" & _
            " AND B.HOURS <> 0"
            
        lsSQL12 = lsSQL12 & " UNION "
        
    End If
'END PREM TRIS CONVERSION PHASE 3
'*******************************************************
'*  Retrieve Exceptions from Actual Time
'*******************************************************
    If liActualTimeTables = TrTable_Actual Or _
        liActualTimeTables = TrTable_ActualAndHistory Then
        lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
        lsSQL16 = lsSQL16 _
            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
            & "A.sched_chg_exc_cd, A.case_num, ' ' as adj_type, " _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(o.ot_opp_desc,a.ot_opp_desc) AS ot_opp_desc, distr_dept_num,  " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.ttr_actual_time A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL16 = lsSQL16 & "WHERE " & lsDeptSelection _
            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.exception_cd "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsExcepInClause) > 0 Then
                    lsSQL16 = lsSQL16 & " AND " & lsExcepInClause
                Else
                    lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
                End If
            Else
                lsSQL16 = lsSQL16 & "AND " & lsAllExcepInClause _
                & " AND A.EXCEPTION_CD <> 'ST'" _
                & " AND A.OVER_EXC_CD = ''" _
                & " AND A.SCHED_CHG_EXC_CD = ''"
            End If
            
            lsSQL16 = lsSQL16 _
                & " AND a.emp_num = d.emp_num " _
                & " AND A.EXCEPTION_CD <> '' " _
                & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) "
                
        lsSQL16 = lsSQL16 & " UNION "
        
        lsSQL17 = lsSQL17 _
            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
            & "A.sched_chg_exc_cd, A.case_num, ' ' as adj_type, " _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num,  " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.ttr_actual_time A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL17 = lsSQL17 & "WHERE " & lsDeptSelection _
            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.OVER_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsOvrExcepInClause) > 0 Then
                    lsSQL17 = lsSQL17 & " AND " & lsOvrExcepInClause
                End If
            Else
                lsSQL17 = lsSQL17 & " AND " & lsAllOvrExcepInClause
            End If
            
        lsSQL17 = lsSQL17 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.OVER_EXC_CD <> ''"
            
        lsSQL17 = lsSQL17 & " UNION "
        
        lsSQL18 = lsSQL18 _
            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
            & "A.sched_chg_exc_cd, A.case_num, ' ' as adj_type, " _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num,  " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.ttr_actual_time A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL18 = lsSQL18 & "WHERE " & lsDeptSelection _
            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsSchedChgExcepInClause) > 0 Then
                    lsSQL18 = lsSQL18 & " AND " & lsSchedChgExcepInClause
                End If
            Else
                lsSQL18 = lsSQL18 & " AND " & lsAllSchedChgExcepInClause
            End If
            
        lsSQL18 = lsSQL18 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.SCHED_CHG_EXC_CD <> ''"
            
        lsSQL18 = lsSQL18 & " UNION "
        
    End If
    
'*******************************************************
'*  Retrieve Exceptions from Actual Time History
'*******************************************************
    If liActualTimeTables = TrTable_History Or _
       liActualTimeTables = TrTable_ActualAndHistory Then
       lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
       lsSQL19 = lsSQL19 _
            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
            & "A.sched_chg_exc_cd, A.case_num, ' ' as adj_type, " _
            & "'3' AS processed_ind, COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee d "
            
        lsSQL19 = lsSQL19 & " WHERE " & lsDeptSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.exception_cd "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsExcepInClause) > 0 Then
                    lsSQL19 = lsSQL19 & " AND " & lsExcepInClause
                Else
                    lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
                End If
            Else
                lsSQL19 = lsSQL19 & " AND " & lsAllExcepInClause _
                & " AND A.EXCEPTION_CD <> 'ST'" _
                & " AND A.OVER_EXC_CD = ''" _
                & " AND A.SCHED_CHG_EXC_CD = ''"
            End If
        
        lsSQL19 = lsSQL19 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.EXCEPTION_CD <> '' "
            
        lsSQL19 = lsSQL19 & " UNION "
        
        lsSQL20 = lsSQL20 _
            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
            & "A.sched_chg_exc_cd, A.case_num, ' ' as adj_type, " _
            & "'3' AS processed_ind, COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee d "
            
        lsSQL20 = lsSQL20 & " WHERE " & lsDeptSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.OVER_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsOvrExcepInClause) > 0 Then
                    lsSQL20 = lsSQL20 & " AND " & lsOvrExcepInClause
                End If
            Else
                lsSQL20 = lsSQL20 & " AND " & lsAllOvrExcepInClause
            End If
        
        lsSQL20 = lsSQL20 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.OVER_EXC_CD <> ''"
            
        lsSQL20 = lsSQL20 & " UNION "
        
        lsSQL21 = lsSQL21 _
            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
            & "A.sched_chg_exc_cd, A.case_num, ' ' as adj_type, " _
            & "'3' AS processed_ind, COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee d "
            
        lsSQL21 = lsSQL21 & " WHERE " & lsDeptSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsSchedChgExcepInClause) > 0 Then
                    lsSQL21 = lsSQL21 & " AND " & lsSchedChgExcepInClause
                End If
            Else
                lsSQL21 = lsSQL21 & " AND " & lsAllSchedChgExcepInClause
            End If
        
        lsSQL21 = lsSQL21 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.SCHED_CHG_EXC_CD <> ''"
            
        lsSQL21 = lsSQL21 & " UNION "
        
    End If
    
    If (RecordsExist(v_adrsSelectedExceptions)) Then
        If Len(lsExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL2 & lsSQL5 & lsSQL10 & lsSQL16 & lsSQL19
        End If
        If Len(lsOvrExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL3 & lsSQL11 & lsSQL17 & lsSQL20
        End If
        If Len(lsSchedChgExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL4 & lsSQL12 & lsSQL18 & lsSQL21
        End If
        If Len(lsOvrExcepInClause) > 0 Or Len(lsExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL1
        End If
    Else
        lsSQLALL = lsSQLALL & lsSQL1 & lsSQL2 & lsSQL3 & lsSQL4 & lsSQL5 & lsSQL10 & _
            lsSQL11 & lsSQL12 & lsSQL16 & lsSQL17 & lsSQL18 & lsSQL19 & lsSQL20 & lsSQL21
    End If
    
    If Trim(Right(lsSQLALL, 6)) = "UNION" Then
        lsSQLALL = Left(lsSQLALL, Len(lsSQLALL) - 6)
    End If
    
    If Len(lsSQLALL) > 0 Then
        lsSQL = "SELECT DEPT_NUM, EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, NTE_GROUP.EXCEPTION_CD, " _
            & " OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, COALESCE(NullIf(EXC1.CASE_GROUP, ''), NullIf(EXC2.CASE_GROUP, ''), NullIf(EXC3.CASE_GROUP, ''), '    ') AS CASE_GROUP, ADJ_TYPE, " _
            & " PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, " _
            & " AMOUNT, TRAVEL_PAY_HRS, " _
            & " HOURS FROM (" & vbCrLf

        lsSQL = lsSQL & "SELECT DEPT_NUM, EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, EXCEPTION_CD, " _
            & " OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, ADJ_TYPE, " _
            & " PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, " _
            & " SUM(AMOUNT) AS AMOUNT, SUM(TRAVEL_PAY_HRS) AS TRAVEL_PAY_HRS, " _
            & " SUM(HOURS) AS HOURS FROM (" _
            & lsSQLALL & ") NTE " & vbCrLf
            
        lsSQL = lsSQL & " GROUP BY DEPT_NUM, EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, EXCEPTION_CD, " _
            & " OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, ADJ_TYPE, " _
            & " PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM) NTE_GROUP" _
            & " " & vbCrLf _
            & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC1" _
            & " ON NTE_GROUP.EXCEPTION_CD = EXC1.EXCEPTION_CD " _
            & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC2 " _
            & " ON NTE_GROUP.OVER_EXC_CD = EXC2.EXCEPTION_CD " _
            & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC3" _
            & " ON NTE_GROUP.SCHED_CHG_EXC_CD = EXC3.EXCEPTION_CD "
        lsSQL = lsSQL & " ORDER BY DEPT_NUM, emp_num,  processed_ind, charge_date"
    End If

    Set lobjCN = New ADODB.Connection
    Set lobjCM = New clsuDBConnectMgr
    lobjCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
    Set lobjCM = Nothing

    lobjCN.Open

    Set lobjRS = New ADODB.Recordset
    lobjRS.CursorLocation = adUseClient
    lobjRS.Open lsSQL, lobjCN, adOpenStatic, adLockBatchOptimistic
    lobjRS.ActiveConnection = Nothing

    lobjCN.Close
    Set lobjCN = Nothing

    Set GetDeptExceptions = lobjRS
    Set lobjRS = Nothing

    '//-----clean up class variables used by madrsExceptions/GetExceptionCodes/IsExceptionCode-----//
    DestroyState

    Exit Function
eh:
    Set lobjCM = Nothing
    Set lobjCN = Nothing
    Set lobjRS = Nothing

    DestroyState

    TrsRaiseError m_strCMODULENAME _
        , "GetDeptExceptions" _
        , , , , DescArgsFrom(lsSQL) _
        , True

End Function

'Public Function GetDeptExceptionsOLD(ByVal v_aryDepartments As Variant _
'                , ByVal v_datStartDate As Date _
'                , ByVal v_datEndDate As Date _
'                , ByVal v_adrsSelectedExceptions As ADODB.Recordset _
'                , Optional v_datLastClosedPP As Date = CDate("1980-01-01") _
'                , Optional v_datEndOfOpenPP As Date = CDate("1980-01-01") _
'                ) As ADODB.Recordset
'On Error GoTo eh
'Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
'Dim lobjCN As ADODB.Connection
'Dim lobjRS As ADODB.Recordset
'Dim lsSQL As String
'Dim lsDeptSelection As String, lsDeptSelection2 As String
'Dim ldtHistoryEndDate As Date, ldtSchedExcStartDate As Date
'Dim ldtActualTimeStartDate As Date, ldtActualTimeEndDate As Date
'Dim lsInClausePartValueKey As String, lsExcepInClause As String
'Dim lsOvrExcepInClause As String
'Dim lsSchedChgExcepInClause As String
'Dim lsExcepandOvrInClause As String
'Dim lsAllExcepInClause As String
'Dim lsAllOvrExcepInClause As String
'Dim lsAllSchedChgExcepInClause As String
'Dim lsAllExcepandOvrInClause As String
'Dim liActualTimeTables As TrActualTimeTable, liRecCounter As Integer
'Dim lblnValidExcep As Boolean
'Dim lblnFirstExcep As Boolean
'Dim lsSQL1 As String
'Dim lsSQL2 As String
'Dim lsSQL3 As String
'Dim lsSQL4 As String
'Dim lsSQL5 As String
'Dim lsSQL10 As String
'Dim lsSQL11 As String
'Dim lsSQL12 As String
'Dim lsSQL16 As String
'Dim lsSQL17 As String
'Dim lsSQL18 As String
'Dim lsSQL19 As String
'Dim lsSQL20 As String
'Dim lsSQL21 As String
'Dim lsSQLALL As String
'
'    If (IsArray(v_aryDepartments)) Then
'        lsDeptSelection = BuildInClause("A.DEPT_NUM", v_aryDepartments, True) & " "
'        lsDeptSelection2 = Replace(lsDeptSelection, "A.DEPT_NUM", "D.DEPT_NUM")
'    Else
'        lsDeptSelection = "A.DEPT_NUM = '" & CStr(CLng(v_aryDepartments)) & "'"
'        lsDeptSelection2 = Replace(lsDeptSelection, "A.DEPT_NUM", "D.DEPT_NUM")
'    End If
'
'    liActualTimeTables = DetermineActualTimeTables(v_datStartDate _
'                            , v_datEndDate _
'                            , ldtHistoryEndDate _
'                            , ldtActualTimeStartDate)
'
'    If v_datLastClosedPP = #1/1/1980# Then
'        ldtSchedExcStartDate = v_datStartDate
'    Else
'        If (v_datLastClosedPP > ldtActualTimeStartDate) Then
'            ldtSchedExcStartDate = DateAdd("d", 1, v_datLastClosedPP)
'        Else
'            ldtSchedExcStartDate = ldtActualTimeStartDate
'        End If
'    End If
'    If v_datEndOfOpenPP = #1/1/1980# Then
'        ldtActualTimeEndDate = v_datEndDate
'    Else
'        If (v_datEndOfOpenPP > ldtActualTimeStartDate) Then
'            ldtActualTimeEndDate = v_datEndOfOpenPP
'        Else
'            ldtActualTimeEndDate = v_datEndDate
'        End If
'    End If
'
'    GetExceptionCodes
'    GetOverExcepCodes
'    GetSchedChgExcepCodes
'
'    If (RecordsExist(v_adrsSelectedExceptions)) Then
'        With v_adrsSelectedExceptions
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsExceptionCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsExcepInClause = lsExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsExcepInClause = lsExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsExcepInClause = " A.EXCEPTION_CD IN ( " & lsExcepInClause & ")"
'            End If
'
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsOverExcepCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsOvrExcepInClause = lsOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsOvrExcepInClause = lsOvrExcepInClause & _
'                            ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsOvrExcepInClause & _
'                ")"
'            End If
'
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsSchedChgExcepCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
'                        "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
'                lsSchedChgExcepInClause & ")"
'            End If
'
'        End With
'
'    Else
'
'        If RecordsExist(madrsExceptions) Then
'            With madrsExceptions
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllExcepInClause = lsAllExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllExcepInClause = lsAllExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllExcepInClause = " A.EXCEPTION_CD IN ( " & lsAllExcepInClause & ")"
'            End With
'        End If
'
'        If RecordsExist(madrsOverExcep) Then
'            With madrsOverExcep
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & _
'                            ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsAllOvrExcepInClause & ")"
'            End With
'        End If
'
'        If RecordsExist(madrsSchedChgExcep) Then
'            With madrsSchedChgExcep
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
'                        "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
'                lsAllSchedChgExcepInClause & ")"
'            End With
'        End If
'
'    End If
'
''*******************************************************
''*  Retrieve Scheduled Exceptions
''*******************************************************
'
'    lsSQL1 = lsSQL1 & "SELECT D.DEPT_NUM, a.emp_num, d.emp_name,a.job_num, a.charge_date, " _
'        & " '00:00' as start_time, '00:00' as end_time, a.exception_cd, a.over_exc_cd, " _
'        & " '' as sched_chg_exc_cd, a.case_num, b.case_group, 'C' as adj_type, " _
'        & " '1' AS processed_ind, '' AS OT_OPP_DESC,'' as distr_dept_num, " _
'        & " 0 as amount, 0 as travel_pay_hrs, hours  " _
'        & " FROM methdba.ttr_emp_sched_exc A, " _
'        & " methdba.ttr_exception_cds B, " _
'        & " methdba.ttr_employee D, methdba.ttr_emp_wg E " _
'        & " WHERE " & lsDeptSelection2 _
'        & " AND A.charge_date between " & DateAsEnclosedString(ldtSchedExcStartDate) & " AND " _
'                & DateAsEnclosedString(v_datEndDate) _
'        & " AND A.emp_num = D.emp_num " _
'        & " AND B.exception_cd = A.exception_cd " _
'        & " AND A.emp_num = E.emp_num " _
'        & " AND E.PRIMARY_IND = 'Y' " _
'        & " AND A.charge_date between e.eff_start_date AND e.eff_end_date "
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) > 0 Then
'                lsSQL1 = lsSQL1 & "AND (" & lsExcepInClause & " OR " & _
'                lsOvrExcepInClause & ") "
'            Else
'                If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) = 0 Then
'                    lsSQL1 = lsSQL1 & "AND (" & lsExcepInClause & ") "
'                Else
'                    If Len(lsExcepInClause) = 0 And Len(lsOvrExcepInClause) > 0 Then
'                        lsSQL1 = lsSQL1 & "AND (" & lsOvrExcepInClause & ") "
'                    Else
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                        " OR A.OVER_EXC_CD <> '') "
'                    End If
'                End If
'            End If
'        Else
'            If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) > 0 Then
'                lsSQL1 = lsSQL1 & "AND (" & lsAllExcepInClause & " OR " & _
'                    lsAllOvrExcepInClause & ") "
'                lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'            Else
'                If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) = 0 Then
'                    lsSQL1 = lsSQL1 & "AND (" & lsAllExcepInClause & ") "
'                    lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                Else
'                    If Len(lsAllExcepInClause) = 0 And Len(lsAllOvrExcepInClause) > 0 Then
'                        lsSQL1 = lsSQL1 & "AND (" & lsAllOvrExcepInClause & ") "
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                    Else
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                    End If
'                End If
'            End If
'        End If
'
'
'    lsSQL1 = lsSQL1 _
'        & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'        & " AND (A.OVER_EXC_CD <> '' OR A.EXCEPTION_CD <> 'ST') "
''BEGIN PREM TRIS CONVERSION PHASE 3
'    lsSQL1 = lsSQL1 & " UNION ALL "
'
'    If (v_datStartDate < #1/27/2005#) Then
'        'Get from Old and New Tables
'        lsSQL2 = lsSQL2 & " SELECT    F.DEPT_NUM " & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC, B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL2 = lsSQL2 & " FROM  METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM " & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.EXCEPTION_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsExcepInClause) > 0) Then
'                lsSQL2 = lsSQL2 & _
'                    " AND " & Replace(lsExcepInClause, "A.", "B.")
'            Else
'                lsSQL2 = lsSQL2 & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'            End If
'        Else
'            lsSQL2 = lsSQL2 & _
'                " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
'                " AND B.EXCEPTION_CD <> 'ST'" & _
'                " AND B.OVER_EXC_CD = ''" & _
'                " AND B.SCHED_CHG_EXC_CD = ''"
'        End If
'
'        lsSQL2 = lsSQL2 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.EXCEPTION_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL2 = lsSQL2 & " UNION ALL "
'
'        lsSQL3 = lsSQL3 & " SELECT F.DEPT_NUM " & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL3 = lsSQL3 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM " & _
'            " , METHDBA.TTR_EXCEPTION_CDS C " & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.OVER_EXC_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsOvrExcepInClause) > 0) Then
'                lsSQL3 = lsSQL3 & _
'                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL3 = lsSQL3 & _
'                " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL3 = lsSQL3 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.OVER_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL3 = lsSQL3 & " UNION ALL "
'
'        lsSQL4 = lsSQL4 & " SELECT F.DEPT_NUM " & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL4 = lsSQL4 & " FROM METHDBA.TTR_ADJ_RQT A " & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C " & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsSchedChgExcepInClause) > 0) Then
'                lsSQL4 = lsSQL4 & _
'                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL4 = lsSQL4 & _
'                " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL4 = lsSQL4 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.SCHED_CHG_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL4 = lsSQL4 & " UNION ALL "
'
'        lsSQL5 = lsSQL5 & " SELECT E.DEPT_NUM " & _
'            " , A.EMP_NUM, E.EMP_NAME" & _
'            " , A.JOB_NUM, A.DATE_WORKED AS CHARGE_DATE" & _
'            " , A.START_TIME, A.END_TIME" & _
'            " , A.EXCEPTION_CD, '' AS OVER_EXC_CD" & _
'            " , '' AS SCHED_CHG_EXC_CD" & _
'            " , A.CASE_NUM, B.CASE_GROUP" & _
'            " , A.ADJ_TYPE" & _
'            " , CASE WHEN (A.PROCESSED_IND) = 'Y' THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , '' AS OT_OPP_DESC, '' AS DISTR_DEPT_NUM" & _
'            " , 0 AS AMOUNT, 0 AS TRAVEL_PAY_HRS" & _
'            " , CASE WHEN A.ADJ_TYPE = 'A' THEN (A.HOURS * -1) ELSE A.HOURS END AS HOURS" & _
'            " FROM  METHDBA.TTR_ADJUSTMENTS A" & _
'            " , METHDBA.TTR_EXCEPTION_CDS B" & _
'            " , METHDBA.TTR_EMP_WG D, METHDBA.TTR_EMPLOYEE E" & _
'            " WHERE " & Replace(lsDeptSelection2, "D.", "E.") & _
'            " AND A.EMP_NUM = D.EMP_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EXCEPTION_CD = B.EXCEPTION_CD" & _
'            " AND A.DATE_WORKED BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.DATE_WORKED BETWEEN D.EFF_START_DATE" & _
'            " AND D.EFF_END_DATE"
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If (Len(lsExcepInClause) > 0) Then
'                    lsSQL5 = lsSQL5 & _
'                        " AND " & lsExcepInClause
'                Else
'                    lsSQL5 = lsSQL5 & _
'                        " AND A.EXCEPTION_CD <> 'ST'"
'                End If
'            Else
'                lsSQL5 = lsSQL5 & _
'                    " AND " & lsAllExcepInClause & _
'                    " AND A.EXCEPTION_CD <> 'ST'"
'            End If
'
'        lsSQL5 = lsSQL5 & _
'            " AND D.PRIMARY_IND = 'Y'" & _
'            " AND A.HOURS <> 0" & _
'            " AND A.EXCEPTION_CD NOT IN ('ET','AT')" & _
'            " AND A.EXCEPTION_CD <> ''"
'
'        lsSQL5 = lsSQL5 & " UNION ALL "
'
'    Else
'        'Get from New Table
'        lsSQL10 = lsSQL10 & " SELECT F.DEPT_NUM " & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL10 = lsSQL10 & " FROM  METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C " & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.EXCEPTION_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsExcepInClause) > 0) Then
'                lsSQL10 = lsSQL10 & _
'                    " AND " & Replace(lsExcepInClause, "A.", "B.")
'            Else
'                lsSQL10 = lsSQL10 & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'            End If
'        Else
'            lsSQL10 = lsSQL10 & _
'                    " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'        End If
'
'        lsSQL10 = lsSQL10 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.EXCEPTION_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL10 = lsSQL10 & " UNION ALL "
'
'        lsSQL11 = lsSQL11 & " SELECT F.DEPT_NUM " & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL11 = lsSQL11 & " FROM  METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C " & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.OVER_EXC_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsOvrExcepInClause) > 0) Then
'                lsSQL11 = lsSQL11 & _
'                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL11 = lsSQL11 & _
'                    " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL11 = lsSQL11 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.OVER_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL11 = lsSQL11 & " UNION ALL "
'
'        lsSQL12 = lsSQL12 & " SELECT F.DEPT_NUM " & _
'            " , A.EMP_NUM, F.EMP_NAME" & _
'            " , B.JOB_NUM, B.CHARGE_DATE, B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL12 = lsSQL12 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C " & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsDeptSelection2, "D.", "F.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsSchedChgExcepInClause) > 0) Then
'                lsSQL12 = lsSQL12 & _
'                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL12 = lsSQL12 & _
'                    " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL12 = lsSQL12 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.SCHED_CHG_EXC_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL12 = lsSQL12 & " UNION ALL "
'
'    End If
''END PREM TRIS CONVERSION PHASE 3
''*******************************************************
''*  Retrieve Exceptions from Actual Time
''*******************************************************
'    If liActualTimeTables = TrTable_Actual Or _
'        liActualTimeTables = TrTable_ActualAndHistory Then
'        lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
'        lsSQL16 = lsSQL16 _
'            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
'            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
'            & "A.sched_chg_exc_cd, A.case_num, B.case_group, ' ' as adj_type, " _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(o.ot_opp_desc,a.ot_opp_desc) AS ot_opp_desc, distr_dept_num,  " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.ttr_actual_time A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL16 = lsSQL16 & "WHERE " & lsDeptSelection _
'            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.exception_cd "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsExcepInClause) > 0 Then
'                    lsSQL16 = lsSQL16 & " AND " & lsExcepInClause
'                Else
'                    lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'                End If
'            Else
'                lsSQL16 = lsSQL16 & "AND " & lsAllExcepInClause _
'                & " AND A.EXCEPTION_CD <> 'ST'" _
'                & " AND A.OVER_EXC_CD = ''" _
'                & " AND A.SCHED_CHG_EXC_CD = ''"
'            End If
'
'            lsSQL16 = lsSQL16 _
'                & " AND a.emp_num = d.emp_num " _
'                & " AND A.EXCEPTION_CD <> '' " _
'                & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) "
'
'        lsSQL16 = lsSQL16 & " UNION ALL "
'
'        lsSQL17 = lsSQL17 _
'            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
'            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
'            & "A.sched_chg_exc_cd, A.case_num, B.case_group, ' ' as adj_type, " _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num,  " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.ttr_actual_time A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL17 = lsSQL17 & "WHERE " & lsDeptSelection _
'            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.OVER_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsOvrExcepInClause) > 0 Then
'                    lsSQL17 = lsSQL17 & " AND " & lsOvrExcepInClause
'                End If
'            Else
'                lsSQL17 = lsSQL17 & " AND " & lsAllOvrExcepInClause
'            End If
'
'        lsSQL17 = lsSQL17 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.OVER_EXC_CD <> ''"
'
'        lsSQL17 = lsSQL17 & " UNION ALL "
'
'        lsSQL18 = lsSQL18 _
'            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
'            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
'            & "A.sched_chg_exc_cd, A.case_num, B.case_group, ' ' as adj_type, " _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num,  " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.ttr_actual_time A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL18 = lsSQL18 & "WHERE " & lsDeptSelection _
'            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsSchedChgExcepInClause) > 0 Then
'                    lsSQL18 = lsSQL18 & " AND " & lsSchedChgExcepInClause
'                End If
'            Else
'                lsSQL18 = lsSQL18 & " AND " & lsAllSchedChgExcepInClause
'            End If
'
'        lsSQL18 = lsSQL18 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.SCHED_CHG_EXC_CD <> ''"
'
'        lsSQL18 = lsSQL18 & " UNION ALL "
'
'    End If
'
''*******************************************************
''*  Retrieve Exceptions from Actual Time History
''*******************************************************
'    If liActualTimeTables = TrTable_History Or _
'       liActualTimeTables = TrTable_ActualAndHistory Then
'       lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
'       lsSQL19 = lsSQL19 _
'            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
'            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
'            & "A.sched_chg_exc_cd, A.case_num, B.case_group, ' ' as adj_type, " _
'            & "'3' AS processed_ind, COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee d "
'
'        lsSQL19 = lsSQL19 & " WHERE " & lsDeptSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.exception_cd "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsExcepInClause) > 0 Then
'                    lsSQL19 = lsSQL19 & " AND " & lsExcepInClause
'                Else
'                    lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'                End If
'            Else
'                lsSQL19 = lsSQL19 & " AND " & lsAllExcepInClause _
'                & " AND A.EXCEPTION_CD <> 'ST'" _
'                & " AND A.OVER_EXC_CD = ''" _
'                & " AND A.SCHED_CHG_EXC_CD = ''"
'            End If
'
'        lsSQL19 = lsSQL19 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.EXCEPTION_CD <> '' "
'
'        lsSQL19 = lsSQL19 & " UNION ALL "
'
'        lsSQL20 = lsSQL20 _
'            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
'            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
'            & "A.sched_chg_exc_cd, A.case_num, B.case_group, ' ' as adj_type, " _
'            & "'3' AS processed_ind, COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee d "
'
'        lsSQL20 = lsSQL20 & " WHERE " & lsDeptSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.OVER_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsOvrExcepInClause) > 0 Then
'                    lsSQL20 = lsSQL20 & " AND " & lsOvrExcepInClause
'                End If
'            Else
'                lsSQL20 = lsSQL20 & " AND " & lsAllOvrExcepInClause
'            End If
'
'        lsSQL20 = lsSQL20 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.OVER_EXC_CD <> ''"
'
'        lsSQL20 = lsSQL20 & " UNION ALL "
'
'        lsSQL21 = lsSQL21 _
'            & "SELECT A.dept_num, A.emp_num,  d.emp_name, A.job_num, A.charge_date, " _
'            & "A.start_time, A.end_time, A.exception_cd, A.over_exc_cd, " _
'            & "A.sched_chg_exc_cd, A.case_num, B.case_group, ' ' as adj_type, " _
'            & "'3' AS processed_ind, COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee d "
'
'        lsSQL21 = lsSQL21 & " WHERE " & lsDeptSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsSchedChgExcepInClause) > 0 Then
'                    lsSQL21 = lsSQL21 & " AND " & lsSchedChgExcepInClause
'                End If
'            Else
'                lsSQL21 = lsSQL21 & " AND " & lsAllSchedChgExcepInClause
'            End If
'
'        lsSQL21 = lsSQL21 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.SCHED_CHG_EXC_CD <> ''"
'
'        lsSQL21 = lsSQL21 & " UNION ALL "
'
'    End If
'
'    If (RecordsExist(v_adrsSelectedExceptions)) Then
'        If Len(lsExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL2 & lsSQL5 & lsSQL10 & lsSQL16 & lsSQL19
'        End If
'        If Len(lsOvrExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL3 & lsSQL11 & lsSQL17 & lsSQL20
'        End If
'        If Len(lsSchedChgExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL4 & lsSQL12 & lsSQL18 & lsSQL21
'        End If
'        If Len(lsOvrExcepInClause) > 0 Or Len(lsExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL1
'        End If
'    Else
'        lsSQLALL = lsSQLALL & lsSQL1 & lsSQL2 & lsSQL3 & lsSQL4 & lsSQL5 & lsSQL10 & _
'            lsSQL11 & lsSQL12 & lsSQL16 & lsSQL17 & lsSQL18 & lsSQL19 & lsSQL20 & lsSQL21
'    End If
'
'    If Trim(Right(lsSQLALL, 10)) = "UNION ALL" Then
'        lsSQLALL = Left(lsSQLALL, Len(lsSQLALL) - 10)
'    End If
'
'    If Len(lsSQLALL) > 0 Then
'        lsSQL = "SELECT DEPT_NUM, EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, EXCEPTION_CD, " _
'            & " OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM , CASE_GROUP, ADJ_TYPE, " _
'            & " PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, " _
'            & " SUM(AMOUNT) AS AMOUNT, SUM(TRAVEL_PAY_HRS) AS TRAVEL_PAY_HRS, " _
'            & " SUM(HOURS) AS HOURS FROM (" _
'            & lsSQLALL & ") NTE "
'        lsSQL = lsSQL & " GROUP BY DEPT_NUM, EMP_NUM, EMP_NAME, JOB_NUM, CHARGE_DATE, EXCEPTION_CD, " _
'            & " OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, CASE_GROUP, ADJ_TYPE, " _
'            & " PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM"
'        lsSQL = lsSQL & " ORDER BY DEPT_NUM, emp_num,  processed_ind, charge_date"
'    End If
'
'    Set lobjCN = New ADODB.Connection
'    Set lobjCM = New clsuDBConnectMgr
'    lobjCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
'    Set lobjCM = Nothing
'
'    lobjCN.Open
'
'    Set lobjRS = New ADODB.Recordset
'    lobjRS.CursorLocation = adUseClient
'    lobjRS.Open lsSQL, lobjCN, adOpenStatic, adLockBatchOptimistic
'    lobjRS.ActiveConnection = Nothing
'
'    lobjCN.Close
'    Set lobjCN = Nothing
'
'    Set GetDeptExceptionsOLD = lobjRS
'    Set lobjRS = Nothing
'
'    '//-----clean up class variables used by madrsExceptions/GetExceptionCodes/IsExceptionCode-----//
'    DestroyState
'
'    Exit Function
'eh:
'    Set lobjCM = Nothing
'    Set lobjCN = Nothing
'    Set lobjRS = Nothing
'
'    DestroyState
'
'    TrsRaiseError m_strCMODULENAME _
'        , "GetDeptExceptionsOLD" _
'        , , , , DescArgsFrom(lsSQL) _
'        , True
'
'End Function

Public Function GetJobExceptions(ByVal v_aryJobs As Variant _
                , ByVal v_datStartDate As Date _
                , ByVal v_datEndDate As Date _
                , ByVal v_adrsSelectedExceptions As ADODB.Recordset _
                , Optional v_datLastClosedPP As Date = CDate("1980-01-01") _
                , Optional v_datEndOfOpenPP As Date = CDate("1980-01-01") _
                ) As ADODB.Recordset
                
'Code has been moved to GetJobExceptions2

End Function

Public Function GetJobExceptions2(ByVal v_aryJobs As Variant _
                , ByVal v_datStartDate As Date _
                , ByVal v_datEndDate As Date _
                , ByVal v_adrsSelectedExceptions As ADODB.Recordset _
                , ByVal v_lLocation As Long _
                , Optional v_datLastClosedPP As Date = CDate("1980-01-01") _
                , Optional v_datEndOfOpenPP As Date = CDate("1980-01-01") _
                ) As ADODB.Recordset
On Error GoTo eh
Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
Dim lobjCN As ADODB.Connection
Dim lobjRS As ADODB.Recordset
Dim lsSQL As String
Dim lsJobSelection As String, lsJobSelection2 As String
Dim ldtHistoryEndDate As Date, ldtSchedExcStartDate As Date
Dim ldtActualTimeStartDate As Date, ldtActualTimeEndDate As Date
Dim llngLocation As Long
Dim lsInClausePartValueKey As String, lsExcepInClause As String
Dim lsOvrExcepInClause As String
Dim lsSchedChgExcepInClause
Dim lsExcepandOvrInClause As String
Dim lsAllExcepInClause As String
Dim lsAllOvrExcepInClause As String
Dim lsAllSchedChgExcepInClause As String
Dim lsAllExcepandOvrInClause As String
Dim liActualTimeTables As TrActualTimeTable, liRecCounter As Integer
Dim lblnValidExcep As Boolean
Dim lblnFirstExcep As Boolean
Dim lsSQL1 As String
Dim lsSQL2 As String
Dim lsSQL3 As String
Dim lsSQL4 As String
Dim lsSQL5 As String
Dim lsSQL10 As String
Dim lsSQL11 As String
Dim lsSQL12 As String
Dim lsSQL16 As String
Dim lsSQL17 As String
Dim lsSQL18 As String
Dim lsSQL19 As String
Dim lsSQL20 As String
Dim lsSQL21 As String
Dim lsSQLALL As String

    If (IsArray(v_aryJobs)) Then
        lsJobSelection = BuildInClause("A.JOB_NUM", v_aryJobs, True) & " "
        lsJobSelection2 = "(((A.JOB_NUM = '') AND (" & Replace(lsJobSelection, "A.JOB_NUM", "E.JOB_NUM") & "))" & _
            " OR ((A.JOB_NUM <> '') AND (" & lsJobSelection & "))) "
    Else
        lsJobSelection = "A.JOB_NUM = '" & CStr(CLng(v_aryJobs)) & "' "
        lsJobSelection2 = "(((A.JOB_NUM = '') AND (" & Replace(lsJobSelection, "A.JOB_NUM", "E.JOB_NUM") & "))" & _
            " OR ((A.JOB_NUM <> '') AND (" & lsJobSelection & "))) "
    End If

    liActualTimeTables = DetermineActualTimeTables(v_datStartDate _
                            , v_datEndDate _
                            , ldtHistoryEndDate _
                            , ldtActualTimeStartDate)
    
    If v_datLastClosedPP = #1/1/1980# Then
        ldtSchedExcStartDate = v_datStartDate
    Else
        If (v_datLastClosedPP > ldtActualTimeStartDate) Then
            ldtSchedExcStartDate = DateAdd("d", 1, v_datLastClosedPP)
        Else
            ldtSchedExcStartDate = ldtActualTimeStartDate
        End If
    End If
        
    If v_datEndOfOpenPP = #1/1/1980# Then
        ldtActualTimeEndDate = v_datEndDate
    Else
        If (v_datEndOfOpenPP > ldtActualTimeStartDate) Then
            ldtActualTimeEndDate = v_datEndOfOpenPP
        Else
            ldtActualTimeEndDate = v_datEndDate
        End If
    End If
    
    llngLocation = v_lLocation

    GetExceptionCodes
    GetOverExcepCodes
    GetSchedChgExcepCodes
    
If (RecordsExist(v_adrsSelectedExceptions)) Then
        
        With v_adrsSelectedExceptions
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsExceptionCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsExcepInClause = lsExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsExcepInClause = lsExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If

                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsExcepInClause = " A.EXCEPTION_CD IN ( " & lsExcepInClause & ")"
            End If
            
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsOverExcepCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsOvrExcepInClause = lsOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsOvrExcepInClause = lsOvrExcepInClause & _
                            ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If
                
                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsOvrExcepInClause & _
                ")"
            End If
            
            .MoveFirst
            lblnValidExcep = False
            lblnFirstExcep = True
            
            For liRecCounter = 1 To .RecordCount
                If IsSchedChgExcepCode(!EXCEPTION_CD) Then
                    lblnValidExcep = True
                    If lblnFirstExcep Then
                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
                        "'" & !EXCEPTION_CD & "'"
                    Else
                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                End If
                
                .MoveNext
            Next liRecCounter
            
            If lblnValidExcep Then
                lsSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
                lsSchedChgExcepInClause & ")"
            End If
            
        End With
        
    Else
    
        If RecordsExist(madrsExceptions) Then
            With madrsExceptions
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllExcepInClause = lsAllExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllExcepInClause = lsAllExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
                
                lsAllExcepInClause = " A.EXCEPTION_CD IN ( " & lsAllExcepInClause & ")"
            End With
        End If
        
        If RecordsExist(madrsOverExcep) Then
            With madrsOverExcep
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & _
                            ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
                
                lsAllOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsAllOvrExcepInClause & ")"
            End With
        End If
        
        If RecordsExist(madrsSchedChgExcep) Then
            With madrsSchedChgExcep
                .MoveFirst
                lblnFirstExcep = True
                For liRecCounter = 1 To .RecordCount
                    If lblnFirstExcep Then
                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
                        "'" & !EXCEPTION_CD & "'"
                    Else
                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
                        ",'" & !EXCEPTION_CD & "'"
                    End If
                    lblnFirstExcep = False
                    
                    .MoveNext
                Next liRecCounter
            
                lsAllSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
                lsAllSchedChgExcepInClause & ")"
            End With
        End If
        
    End If

'*******************************************************
'*  Retrieve Scheduled Exceptions
'*******************************************************
        
    lsSQL1 = lsSQL1 & "SELECT CASE WHEN A.JOB_NUM = '' THEN E.JOB_NUM ELSE a.job_num END AS JOB_NUM, " _
        & "a.emp_num, e.emp_name, a.charge_date, " _
        & "'00:00' as start_time, '00:00' as end_time, a.exception_cd, " _
        & "a.over_exc_cd, '' as sched_chg_exc_cd, a.case_num, " _
        & "'C' as adj_type, '1' AS processed_ind, '' AS OT_OPP_DESC, " _
        & " '' as distr_dept_num, 0 as amount, 0 as travel_pay_hrs, hours " _
        & "FROM methdba.ttr_emp_sched_exc A, " _
        & " methdba.ttr_exception_cds B, " _
        & " methdba.ttr_emp_wg D , methdba.ttr_employee E " _
        & " WHERE " & lsJobSelection2 _
        & " AND A.charge_date between " & DateAsEnclosedString(ldtSchedExcStartDate) & " AND " _
                & DateAsEnclosedString(v_datEndDate) _
        & " AND A.emp_num = D.emp_num " _
        & " AND A.emp_num = E.emp_num " _
        & " AND D.PRIMARY_IND = 'Y'" _
        & " AND A.charge_date between D.eff_start_date AND D.eff_end_date " _
        & " AND B.exception_cd = A.exception_cd "
        
    If (RecordsExist(v_adrsSelectedExceptions)) Then
            If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) > 0 Then
                lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & " OR " & _
                lsOvrExcepInClause & ") "
            Else
                If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) = 0 Then
                    lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & ") "
                Else
                    If Len(lsExcepInClause) = 0 And Len(lsOvrExcepInClause) > 0 Then
                        lsSQL1 = lsSQL1 & " AND (" & lsOvrExcepInClause & ") "
                    Else
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                        " OR A.OVER_EXC_CD <> '') "
                    End If
                End If
            End If
        Else
            If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) > 0 Then
                lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & " OR " & _
                    lsAllOvrExcepInClause & ") "
                lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
            Else
                If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) = 0 Then
                    lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & ") "
                    lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                Else
                    If Len(lsAllExcepInClause) = 0 And Len(lsAllOvrExcepInClause) > 0 Then
                        lsSQL1 = lsSQL1 & " AND (" & lsAllOvrExcepInClause & ") "
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                    Else
                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
                            " OR A.OVER_EXC_CD <> '') "
                    End If
                End If
            End If
        End If
        
        lsSQL1 = lsSQL1 _
        & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
        & " AND A.exception_cd NOT IN ( SELECT exception_cd " _
                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5 ) " _
        & " AND A.over_exc_cd NOT IN ( SELECT exception_cd " _
                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5 ) "
'BEGIN TRIS CONVERSION PHASE 3
    lsSQL1 = lsSQL1 & " UNION "

    If (v_datStartDate < #1/27/2005#) Then
        'Get from Old and New Tables
        lsSQL2 = lsSQL2 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
            " , F.EMP_NAME, B.CHARGE_DATE " & _
            " , B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL2 = lsSQL2 & " FROM METHDBA.TTR_ADJ_RQT A " & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C " & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.EXCEPTION_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If llngLocation > 0 Then
            lsSQL2 = lsSQL2 & _
                " AND B.LOC_CD = " & llngLocation
        End If
        
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsExcepInClause) > 0) Then
                lsSQL2 = lsSQL2 & _
                    " AND " & Replace(lsExcepInClause, "A.", "B.")
            Else
                lsSQL2 = lsSQL2 & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
            End If
        Else
            lsSQL2 = lsSQL2 & _
            " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
            " AND B.EXCEPTION_CD <> 'ST'" & _
            " AND B.OVER_EXC_CD = ''" & _
            " AND B.SCHED_CHG_EXC_CD = ''"
        End If
        
        lsSQL2 = lsSQL2 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.EXCEPTION_CD <> ''" & _
            " AND B.HOURS <> 0"
                    
        lsSQL2 = lsSQL2 & " UNION "
        
        lsSQL3 = lsSQL3 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
            " , F.EMP_NAME, B.CHARGE_DATE" & _
            " , B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL3 = lsSQL3 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C " & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.OVER_EXC_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If llngLocation > 0 Then
            lsSQL3 = lsSQL3 & _
                " AND B.LOC_CD = " & llngLocation
        End If
        
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsOvrExcepInClause) > 0) Then
                lsSQL3 = lsSQL3 & _
                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL3 = lsSQL3 & _
                    " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
        End If
        
        lsSQL3 = lsSQL3 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.HOURS <> 0" & _
            " AND B.OVER_EXC_CD <> ''"
            
        lsSQL3 = lsSQL3 & " UNION "
        
        lsSQL4 = lsSQL4 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
            " , F.EMP_NAME, B.CHARGE_DATE" & _
            " , B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
        
        lsSQL4 = lsSQL4 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C " & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If llngLocation > 0 Then
            lsSQL4 = lsSQL4 & _
                " AND B.LOC_CD = " & llngLocation
        End If
        
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsSchedChgExcepInClause) > 0) Then
                lsSQL4 = lsSQL4 & _
                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL4 = lsSQL4 & _
                    " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
        End If
        
        lsSQL4 = lsSQL4 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.HOURS <> 0" & _
            " AND B.SCHED_CHG_EXC_CD <> ''"
            
        lsSQL4 = lsSQL4 & " UNION "

        lsSQL5 = lsSQL5 & " SELECT A.JOB_NUM, A.EMP_NUM" & _
            " , E.EMP_NAME, A.DATE_WORKED AS CHARGE_DATE" & _
            " , A.START_TIME, A.END_TIME" & _
            " , A.EXCEPTION_CD, '' AS OVER_EXC_CD" & _
            " , '' AS SCHED_CHG_EXC_CD" & _
            " , A.CASE_NUM" & _
            " , A.ADJ_TYPE" & _
            " , CASE WHEN (A.PROCESSED_IND) = 'Y' THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , '' AS OT_OPP_DESC, '' AS DISTR_DEPT_NUM" & _
            " , 0 AS AMOUNT, 0 AS TRAVEL_PAY_HRS, " & _
            " CASE WHEN A.ADJ_TYPE = 'A' THEN (A.HOURS * -1) ELSE A.HOURS END AS HOURS" & _
            " FROM  METHDBA.TTR_ADJUSTMENTS A" & _
            " , METHDBA.TTR_EXCEPTION_CDS B" & _
            " , METHDBA.TTR_EMP_WG D, METHDBA.TTR_EMPLOYEE E" & _
            " WHERE " & lsJobSelection & _
            " AND A.EMP_NUM = D.EMP_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EXCEPTION_CD = B.EXCEPTION_CD" & _
            " AND A.DATE_WORKED BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.DATE_WORKED BETWEEN D.EFF_START_DATE" & _
            " AND D.EFF_END_DATE"
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If (Len(lsExcepInClause) > 0) Then
                    lsSQL5 = lsSQL5 & _
                        " AND " & lsExcepInClause
                Else
                    lsSQL5 = lsSQL5 & _
                        " AND A.EXCEPTION_CD <> 'ST'"
                End If
            Else
                lsSQL5 = lsSQL5 & _
                        " AND " & lsAllExcepInClause & _
                        " AND A.EXCEPTION_CD <> 'ST'"
            End If

        lsSQL5 = lsSQL5 & _
            " AND D.PRIMARY_IND = 'Y'" & _
            " AND A.EXCEPTION_CD <> ''" & _
            " AND A.HOURS <> 0" & _
            " AND A.EXCEPTION_CD NOT IN ('ET','AT')"
            
        lsSQL5 = lsSQL5 & " UNION "
        
    Else
        'Get from Old and New Tables
        lsSQL10 = lsSQL10 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
            " , F.EMP_NAME, B.CHARGE_DATE" & _
            " , B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL10 = lsSQL10 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.EXCEPTION_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
                    
        If llngLocation > 0 Then
            lsSQL10 = lsSQL10 & _
                " AND B.LOC_CD = " & llngLocation
        End If
        
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsExcepInClause) > 0) Then
                lsSQL10 = lsSQL10 & _
                    " AND " & Replace(lsExcepInClause, "A.", "B.")
            Else
                lsSQL10 = lsSQL10 & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
            End If
        Else
            lsSQL10 = lsSQL10 & _
                    " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
                    " AND B.EXCEPTION_CD <> 'ST'" & _
                    " AND B.OVER_EXC_CD = ''" & _
                    " AND B.SCHED_CHG_EXC_CD = ''"
        End If
        
        lsSQL10 = lsSQL10 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.HOURS <> 0" & _
            " AND B.EXCEPTION_CD <> ''"
            
        lsSQL10 = lsSQL10 & " UNION "
        
        lsSQL11 = lsSQL11 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
            " , F.EMP_NAME, B.CHARGE_DATE" & _
            " , B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL11 = lsSQL11 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.OVER_EXC_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
            
        If llngLocation > 0 Then
            lsSQL11 = lsSQL11 & _
                " AND B.LOC_CD = " & llngLocation
        End If
        
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsOvrExcepInClause) > 0) Then
                lsSQL11 = lsSQL11 & _
                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL11 = lsSQL11 & _
                    " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
        End If
        
        lsSQL11 = lsSQL11 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.HOURS <> 0" & _
            " AND B.OVER_EXC_CD <> ''"
                    
        lsSQL11 = lsSQL11 & " UNION "

        lsSQL12 = lsSQL12 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
            " , F.EMP_NAME, B.CHARGE_DATE" & _
            " , B.START_TIME, B.END_TIME" & _
            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
            " , B.SCHED_CHG_EXC_CD" & _
            " , B.CASE_NUM" & _
            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
            " ELSE '2' END AS PROCESSED_IND" & _
            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
            " , B.DISTR_DEPT_NUM" & _
            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
            
        lsSQL12 = lsSQL12 & " FROM METHDBA.TTR_ADJ_RQT A" & _
            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
            " , METHDBA.TTR_EXCEPTION_CDS C" & _
            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
            " AND '" & CStr(v_datEndDate) & "'" & _
            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
            " AND A.EMP_NUM = E.EMP_NUM" & _
            " AND A.EMP_NUM = F.EMP_NUM" & _
            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
            " AND E.EFF_END_DATE"
        
        If llngLocation > 0 Then
            lsSQL12 = lsSQL12 & _
                " AND B.LOC_CD = " & llngLocation
        End If
        
        If (RecordsExist(v_adrsSelectedExceptions)) Then
            If (Len(lsSchedChgExcepInClause) > 0) Then
                lsSQL12 = lsSQL12 & _
                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
            End If
        Else
            lsSQL12 = lsSQL12 & _
                    " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
        End If
        
        lsSQL12 = lsSQL12 & _
            " AND E.PRIMARY_IND = 'Y'" & _
            " AND B.HOURS <> 0" & _
            " AND B.SCHED_CHG_EXC_CD <> ''"
            
        lsSQL12 = lsSQL12 & " UNION "
    
    End If

'END TRIS CONVERSION PHASE 3
'*******************************************************
'*  Retrieve Exceptions from Actual Time
'*******************************************************
    If liActualTimeTables = TrTable_Actual Or _
        liActualTimeTables = TrTable_ActualAndHistory Then
        lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
        lsSQL16 = lsSQL16 _
            & "SELECT  A.job_num, A.emp_num,  d.emp_name, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, " _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL16 = lsSQL16 & "WHERE " & lsJobSelection _
            & " AND A.charge_date between " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.exception_cd "
            
            If llngLocation > 0 Then
                lsSQL16 = lsSQL16 & " AND A.LOC_CD = " & llngLocation
            End If
                    
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsExcepInClause) > 0 Then
                    lsSQL16 = lsSQL16 & " AND " & lsExcepInClause
                Else
                    lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
                End If
            Else
                lsSQL16 = lsSQL16 & " AND " & lsAllExcepInClause
                lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
            End If
            
        lsSQL16 = lsSQL16 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.EXCEPTION_CD <> ''"
            
        lsSQL16 = lsSQL16 & " UNION "
        
        lsSQL17 = lsSQL17 _
            & "SELECT  A.job_num, A.emp_num,  d.emp_name, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, " _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL17 = lsSQL17 & "WHERE " & lsJobSelection _
            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.OVER_EXC_CD "
            
            If llngLocation > 0 Then
                lsSQL17 = lsSQL17 & " AND A.LOC_CD = " & llngLocation
            End If
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsOvrExcepInClause) > 0 Then
                    lsSQL17 = lsSQL17 & " AND " & lsOvrExcepInClause
                End If
            Else
                lsSQL17 = lsSQL17 & " AND " & lsAllOvrExcepInClause
            End If
            
        lsSQL17 = lsSQL17 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.OVER_EXC_CD <> ''"
            
        lsSQL17 = lsSQL17 & " UNION "
        
        lsSQL18 = lsSQL18 _
            & "SELECT  A.job_num, A.emp_num,  d.emp_name, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, " _
            & " CASE processed_ind " _
            & " WHEN 'Y' THEN '3' " _
            & " ELSE '2' END AS processed_ind, " _
            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
            & " amount, travel_pay_hrs, hours " _
            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cdS B, " _
            & " methdba.ttr_employee D "
            
        lsSQL18 = lsSQL18 & "WHERE " & lsJobSelection _
            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
            & DateAsEnclosedString(ldtActualTimeEndDate) _
            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
            
            If llngLocation > 0 Then
                lsSQL18 = lsSQL18 & " AND A.LOC_CD = " & llngLocation
            End If
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsSchedChgExcepInClause) > 0 Then
                    lsSQL18 = lsSQL18 & " AND " & lsSchedChgExcepInClause
                End If
            Else
                lsSQL18 = lsSQL18 & " AND " & lsAllSchedChgExcepInClause
            End If
            
        lsSQL18 = lsSQL18 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.SCHED_CHG_EXC_CD <> ''"
            
        lsSQL18 = lsSQL18 & " UNION "

    End If
    
'*******************************************************
'*  Retrieve Exceptions from Actual Time History
'*******************************************************
    If liActualTimeTables = TrTable_History Or _
       liActualTimeTables = TrTable_ActualAndHistory Then
       lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
       lsSQL19 = lsSQL19 _
            & "SELECT A.job_num, A.emp_num,  d.emp_name, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, '3' AS processed_ind, " _
            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee d "
            
        lsSQL19 = lsSQL19 & " WHERE " & lsJobSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.exception_cd "
            
            If llngLocation > 0 Then
                lsSQL19 = lsSQL19 & " AND A.LOC_CD = " & llngLocation
            End If
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsExcepInClause) > 0 Then
                    lsSQL19 = lsSQL19 & " AND " & lsExcepInClause
                Else
                    lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
                End If
            Else
                lsSQL19 = lsSQL19 & " AND " & lsAllExcepInClause
                lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
                    & " AND A.OVER_EXC_CD = ''" _
                    & " AND A.SCHED_CHG_EXC_CD = ''"
            End If
        
        lsSQL19 = lsSQL19 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.EXCEPTION_CD <> ''"
            
        lsSQL19 = lsSQL19 & " UNION "
        
        lsSQL20 = lsSQL20 _
            & "SELECT A.job_num, A.emp_num,  d.emp_name, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, '3' AS processed_ind, " _
            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee d "
            
        lsSQL20 = lsSQL20 & " WHERE " & lsJobSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.OVER_EXC_CD "
            
            If llngLocation > 0 Then
                lsSQL20 = lsSQL20 & " AND A.LOC_CD = " & llngLocation
            End If
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsOvrExcepInClause) > 0 Then
                    lsSQL20 = lsSQL20 & " AND " & lsOvrExcepInClause
                End If
            Else
                lsSQL20 = lsSQL20 & " AND " & lsAllOvrExcepInClause
            End If
        
        lsSQL20 = lsSQL20 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.OVER_EXC_CD <> ''"
            
        lsSQL20 = lsSQL20 & " UNION "
        
        lsSQL21 = lsSQL21 _
            & "SELECT A.job_num, A.emp_num,  d.emp_name, " _
            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, " _
            & "' ' as adj_type, '3' AS processed_ind, " _
            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
            & "distr_dept_num, amount, travel_pay_hrs, hours " _
            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
            & " METHDBA.TTR_exception_cds B, " _
            & " methdba.ttr_employee d "
            
        lsSQL21 = lsSQL21 & " WHERE " & lsJobSelection & " AND " & lsInClausePartValueKey _
            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
            & DateAsEnclosedString(ldtHistoryEndDate) _
            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
            
            If llngLocation > 0 Then
                lsSQL21 = lsSQL21 & " AND A.LOC_CD = " & llngLocation
            End If
            
            If (RecordsExist(v_adrsSelectedExceptions)) Then
                If Len(lsSchedChgExcepInClause) > 0 Then
                    lsSQL21 = lsSQL21 & " AND " & lsSchedChgExcepInClause
                End If
            Else
                lsSQL21 = lsSQL21 & " AND " & lsAllSchedChgExcepInClause
            End If
        
        lsSQL21 = lsSQL21 _
            & " AND a.emp_num = d.emp_num " _
            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
            & " AND A.SCHED_CHG_EXC_CD <> ''"
            
        lsSQL21 = lsSQL21 & " UNION "

    End If
    
    If (RecordsExist(v_adrsSelectedExceptions)) Then
        If Len(lsExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL2 & lsSQL5 & lsSQL10 & lsSQL16 & lsSQL19
        End If
        If Len(lsOvrExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL3 & lsSQL11 & lsSQL17 & lsSQL20
        End If
        If Len(lsSchedChgExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL4 & lsSQL12 & lsSQL18 & lsSQL21
        End If
        If Len(lsOvrExcepInClause) > 0 Or Len(lsExcepInClause) > 0 Then
            lsSQLALL = lsSQLALL & lsSQL1
        End If
    Else
        lsSQLALL = lsSQLALL & lsSQL1 & lsSQL2 & lsSQL3 & lsSQL4 & lsSQL5 & lsSQL10 & _
            lsSQL11 & lsSQL12 & lsSQL16 & lsSQL17 & lsSQL18 & lsSQL19 & lsSQL20 & lsSQL21
    End If
    
    If Trim(Right(lsSQLALL, 6)) = "UNION" Then
        lsSQLALL = Left(lsSQLALL, Len(lsSQLALL) - 6)
    End If
    
    '//try to elimiate duplicates when selected exceptions are 'E' and 'FCL'
    'sql 16 - use actual time & exception cd 'E'
    'sql 17 - use actual time & override exception cd 'FCL'
    'since these were UNION ALL, there would be 2 rows (of the same actual time) only differing by CASE_GROUP
    '//take CASE_GROUP out of SQL detail and group by, use UNION rather than UNION ALL, then use 3 joins to get the possible CASE_GROUP's
    If Len(lsSQLALL) > 0 Then
        lsSQL = "SELECT JOB_NUM, EMP_NUM, EMP_NAME, CHARGE_DATE, " _
        & " NTE_GROUP.EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM , " _
        & " COALESCE(NullIf(EXC1.CASE_GROUP, ''), NullIf(EXC2.CASE_GROUP, ''), NullIf(EXC3.CASE_GROUP, ''), '    ') AS CASE_GROUP, " _
        & " ADJ_TYPE, PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, " _
        & " AMOUNT, TRAVEL_PAY_HRS, " _
        & " HOURS FROM (" & vbCrLf
        
        lsSQL = lsSQL & "SELECT JOB_NUM, EMP_NUM, EMP_NAME, CHARGE_DATE, " _
        & " EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM , " _
        & " ADJ_TYPE, PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, " _
        & " SUM(AMOUNT) AS AMOUNT, SUM(TRAVEL_PAY_HRS) AS TRAVEL_PAY_HRS, " _
        & " SUM(HOURS) AS HOURS FROM (" _
        & lsSQLALL & ") NTE " & vbCrLf
        
        lsSQL = lsSQL & " GROUP BY JOB_NUM, EMP_NUM, EMP_NAME, " _
        & " CHARGE_DATE, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, " _
        & " ADJ_TYPE, PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM) NTE_GROUP " _
        & " " & vbCrLf _
        & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC1" _
        & " ON NTE_GROUP.EXCEPTION_CD = EXC1.EXCEPTION_CD " _
        & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC2 " _
        & " ON NTE_GROUP.OVER_EXC_CD = EXC2.EXCEPTION_CD " _
        & " LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS EXC3" _
        & " ON NTE_GROUP.SCHED_CHG_EXC_CD = EXC3.EXCEPTION_CD "
            
        lsSQL = lsSQL & " ORDER BY job_num, emp_num,  processed_ind, charge_date"
    End If
    
      
    Set lobjCN = New ADODB.Connection
    Set lobjCM = New clsuDBConnectMgr
    lobjCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
    Set lobjCM = Nothing

    lobjCN.Open

    Set lobjRS = New ADODB.Recordset
    lobjRS.CursorLocation = adUseClient
    lobjRS.Open lsSQL, lobjCN, adOpenStatic, adLockBatchOptimistic
    lobjRS.ActiveConnection = Nothing

    lobjCN.Close
    Set lobjCN = Nothing

    Set GetJobExceptions2 = lobjRS
    Set lobjRS = Nothing

    '//-----clean up class variables used by madrsExceptions/GetExceptionCodes/IsExceptionCode-----//
    DestroyState

    Exit Function
eh:
    Set lobjCM = Nothing
    Set lobjCN = Nothing
    Set lobjRS = Nothing

    DestroyState

    TrsRaiseError m_strCMODULENAME _
        , "GetJobExceptions2" _
        , , , , DescArgsFrom(lsSQL) _
        , True

End Function

'Public Function GetJobExceptions2OLD(ByVal v_aryJobs As Variant _
'                , ByVal v_datStartDate As Date _
'                , ByVal v_datEndDate As Date _
'                , ByVal v_adrsSelectedExceptions As ADODB.Recordset _
'                , ByVal v_lLocation As Long _
'                , Optional v_datLastClosedPP As Date = CDate("1980-01-01") _
'                , Optional v_datEndOfOpenPP As Date = CDate("1980-01-01") _
'                ) As ADODB.Recordset
'On Error GoTo eh
'Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
'Dim lobjCN As ADODB.Connection
'Dim lobjRS As ADODB.Recordset
'Dim lsSQL As String
'Dim lsJobSelection As String, lsJobSelection2 As String
'Dim ldtHistoryEndDate As Date, ldtSchedExcStartDate As Date
'Dim ldtActualTimeStartDate As Date, ldtActualTimeEndDate As Date
'Dim llngLocation As Long
'Dim lsInClausePartValueKey As String, lsExcepInClause As String
'Dim lsOvrExcepInClause As String
'Dim lsSchedChgExcepInClause
'Dim lsExcepandOvrInClause As String
'Dim lsAllExcepInClause As String
'Dim lsAllOvrExcepInClause As String
'Dim lsAllSchedChgExcepInClause As String
'Dim lsAllExcepandOvrInClause As String
'Dim liActualTimeTables As TrActualTimeTable, liRecCounter As Integer
'Dim lblnValidExcep As Boolean
'Dim lblnFirstExcep As Boolean
'Dim lsSQL1 As String
'Dim lsSQL2 As String
'Dim lsSQL3 As String
'Dim lsSQL4 As String
'Dim lsSQL5 As String
'Dim lsSQL10 As String
'Dim lsSQL11 As String
'Dim lsSQL12 As String
'Dim lsSQL16 As String
'Dim lsSQL17 As String
'Dim lsSQL18 As String
'Dim lsSQL19 As String
'Dim lsSQL20 As String
'Dim lsSQL21 As String
'Dim lsSQLALL As String
'
'    If (IsArray(v_aryJobs)) Then
'        lsJobSelection = BuildInClause("A.JOB_NUM", v_aryJobs, True) & " "
'        lsJobSelection2 = "(((A.JOB_NUM = '') AND (" & Replace(lsJobSelection, "A.JOB_NUM", "E.JOB_NUM") & "))" & _
'            " OR ((A.JOB_NUM <> '') AND (" & lsJobSelection & "))) "
'    Else
'        lsJobSelection = "A.JOB_NUM = '" & CStr(CLng(v_aryJobs)) & "' "
'        lsJobSelection2 = "(((A.JOB_NUM = '') AND (" & Replace(lsJobSelection, "A.JOB_NUM", "E.JOB_NUM") & "))" & _
'            " OR ((A.JOB_NUM <> '') AND (" & lsJobSelection & "))) "
'    End If
'
'    liActualTimeTables = DetermineActualTimeTables(v_datStartDate _
'                            , v_datEndDate _
'                            , ldtHistoryEndDate _
'                            , ldtActualTimeStartDate)
'
'    If v_datLastClosedPP = #1/1/1980# Then
'        ldtSchedExcStartDate = v_datStartDate
'    Else
'        If (v_datLastClosedPP > ldtActualTimeStartDate) Then
'            ldtSchedExcStartDate = DateAdd("d", 1, v_datLastClosedPP)
'        Else
'            ldtSchedExcStartDate = ldtActualTimeStartDate
'        End If
'    End If
'
'    If v_datEndOfOpenPP = #1/1/1980# Then
'        ldtActualTimeEndDate = v_datEndDate
'    Else
'        If (v_datEndOfOpenPP > ldtActualTimeStartDate) Then
'            ldtActualTimeEndDate = v_datEndOfOpenPP
'        Else
'            ldtActualTimeEndDate = v_datEndDate
'        End If
'    End If
'
'    llngLocation = v_lLocation
'
'    GetExceptionCodes
'    GetOverExcepCodes
'    GetSchedChgExcepCodes
'
'If (RecordsExist(v_adrsSelectedExceptions)) Then
'
'        With v_adrsSelectedExceptions
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsExceptionCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsExcepInClause = lsExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsExcepInClause = lsExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsExcepInClause = " A.EXCEPTION_CD IN ( " & lsExcepInClause & ")"
'            End If
'
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsOverExcepCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsOvrExcepInClause = lsOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsOvrExcepInClause = lsOvrExcepInClause & _
'                            ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsOvrExcepInClause & _
'                ")"
'            End If
'
'            .MoveFirst
'            lblnValidExcep = False
'            lblnFirstExcep = True
'
'            For liRecCounter = 1 To .RecordCount
'                If IsSchedChgExcepCode(!EXCEPTION_CD) Then
'                    lblnValidExcep = True
'                    If lblnFirstExcep Then
'                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
'                        "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsSchedChgExcepInClause = lsSchedChgExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'                End If
'
'                .MoveNext
'            Next liRecCounter
'
'            If lblnValidExcep Then
'                lsSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
'                lsSchedChgExcepInClause & ")"
'            End If
'
'        End With
'
'    Else
'
'        If RecordsExist(madrsExceptions) Then
'            With madrsExceptions
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllExcepInClause = lsAllExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllExcepInClause = lsAllExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllExcepInClause = " A.EXCEPTION_CD IN ( " & lsAllExcepInClause & ")"
'            End With
'        End If
'
'        If RecordsExist(madrsOverExcep) Then
'            With madrsOverExcep
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllOvrExcepInClause = lsAllOvrExcepInClause & _
'                            ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllOvrExcepInClause = " A.OVER_EXC_CD IN ( " & lsAllOvrExcepInClause & ")"
'            End With
'        End If
'
'        If RecordsExist(madrsSchedChgExcep) Then
'            With madrsSchedChgExcep
'                .MoveFirst
'                lblnFirstExcep = True
'                For liRecCounter = 1 To .RecordCount
'                    If lblnFirstExcep Then
'                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
'                        "'" & !EXCEPTION_CD & "'"
'                    Else
'                        lsAllSchedChgExcepInClause = lsAllSchedChgExcepInClause & _
'                        ",'" & !EXCEPTION_CD & "'"
'                    End If
'                    lblnFirstExcep = False
'
'                    .MoveNext
'                Next liRecCounter
'
'                lsAllSchedChgExcepInClause = " A.SCHED_CHG_EXC_CD IN ( " & _
'                lsAllSchedChgExcepInClause & ")"
'            End With
'        End If
'
'    End If
'
''*******************************************************
''*  Retrieve Scheduled Exceptions
''*******************************************************
'
'    lsSQL1 = lsSQL1 & "SELECT CASE WHEN A.JOB_NUM = '' THEN E.JOB_NUM ELSE a.job_num END AS JOB_NUM, " _
'        & "a.emp_num, e.emp_name, a.charge_date, " _
'        & "'00:00' as start_time, '00:00' as end_time, a.exception_cd, " _
'        & "a.over_exc_cd, '' as sched_chg_exc_cd, a.case_num, b.case_group, " _
'        & "'C' as adj_type, '1' AS processed_ind, '' AS OT_OPP_DESC, " _
'        & " '' as distr_dept_num, 0 as amount, 0 as travel_pay_hrs, hours " _
'        & "FROM methdba.ttr_emp_sched_exc A, " _
'        & " methdba.ttr_exception_cds B, " _
'        & " methdba.ttr_emp_wg D , methdba.ttr_employee E " _
'        & " WHERE " & lsJobSelection2 _
'        & " AND A.charge_date between " & DateAsEnclosedString(ldtSchedExcStartDate) & " AND " _
'                & DateAsEnclosedString(v_datEndDate) _
'        & " AND A.emp_num = D.emp_num " _
'        & " AND A.emp_num = E.emp_num " _
'        & " AND D.PRIMARY_IND = 'Y'" _
'        & " AND A.charge_date between D.eff_start_date AND D.eff_end_date " _
'        & " AND B.exception_cd = A.exception_cd "
'
'    If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) > 0 Then
'                lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & " OR " & _
'                lsOvrExcepInClause & ") "
'            Else
'                If Len(lsExcepInClause) > 0 And Len(lsOvrExcepInClause) = 0 Then
'                    lsSQL1 = lsSQL1 & " AND (" & lsExcepInClause & ") "
'                Else
'                    If Len(lsExcepInClause) = 0 And Len(lsOvrExcepInClause) > 0 Then
'                        lsSQL1 = lsSQL1 & " AND (" & lsOvrExcepInClause & ") "
'                    Else
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                        " OR A.OVER_EXC_CD <> '') "
'                    End If
'                End If
'            End If
'        Else
'            If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) > 0 Then
'                lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & " OR " & _
'                    lsAllOvrExcepInClause & ") "
'                lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'            Else
'                If Len(lsAllExcepInClause) > 0 And Len(lsAllOvrExcepInClause) = 0 Then
'                    lsSQL1 = lsSQL1 & " AND (" & lsAllExcepInClause & ") "
'                    lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                Else
'                    If Len(lsAllExcepInClause) = 0 And Len(lsAllOvrExcepInClause) > 0 Then
'                        lsSQL1 = lsSQL1 & " AND (" & lsAllOvrExcepInClause & ") "
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                    Else
'                        lsSQL1 = lsSQL1 & " AND (A.EXCEPTION_CD <> 'ST' " & _
'                            " OR A.OVER_EXC_CD <> '') "
'                    End If
'                End If
'            End If
'        End If
'
'        lsSQL1 = lsSQL1 _
'        & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'        & " AND A.exception_cd NOT IN ( SELECT exception_cd " _
'                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5 ) " _
'        & " AND A.over_exc_cd NOT IN ( SELECT exception_cd " _
'                    & " FROM methdba.ttr_exc_group WHERE eg_cd = 5 ) "
''BEGIN TRIS CONVERSION PHASE 3
'    lsSQL1 = lsSQL1 & " UNION ALL "
'
'    If (v_datStartDate < #1/27/2005#) Then
'        'Get from Old and New Tables
'        lsSQL2 = lsSQL2 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
'            " , F.EMP_NAME, B.CHARGE_DATE " & _
'            " , B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL2 = lsSQL2 & " FROM METHDBA.TTR_ADJ_RQT A " & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C " & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.EXCEPTION_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If llngLocation > 0 Then
'            lsSQL2 = lsSQL2 & _
'                " AND B.LOC_CD = " & llngLocation
'        End If
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsExcepInClause) > 0) Then
'                lsSQL2 = lsSQL2 & _
'                    " AND " & Replace(lsExcepInClause, "A.", "B.")
'            Else
'                lsSQL2 = lsSQL2 & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'            End If
'        Else
'            lsSQL2 = lsSQL2 & _
'            " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
'            " AND B.EXCEPTION_CD <> 'ST'" & _
'            " AND B.OVER_EXC_CD = ''" & _
'            " AND B.SCHED_CHG_EXC_CD = ''"
'        End If
'
'        lsSQL2 = lsSQL2 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.EXCEPTION_CD <> ''" & _
'            " AND B.HOURS <> 0"
'
'        lsSQL2 = lsSQL2 & " UNION ALL "
'
'        lsSQL3 = lsSQL3 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
'            " , F.EMP_NAME, B.CHARGE_DATE" & _
'            " , B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL3 = lsSQL3 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C " & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.OVER_EXC_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If llngLocation > 0 Then
'            lsSQL3 = lsSQL3 & _
'                " AND B.LOC_CD = " & llngLocation
'        End If
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsOvrExcepInClause) > 0) Then
'                lsSQL3 = lsSQL3 & _
'                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL3 = lsSQL3 & _
'                    " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL3 = lsSQL3 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.HOURS <> 0" & _
'            " AND B.OVER_EXC_CD <> ''"
'
'        lsSQL3 = lsSQL3 & " UNION ALL "
'
'        lsSQL4 = lsSQL4 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
'            " , F.EMP_NAME, B.CHARGE_DATE" & _
'            " , B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL4 = lsSQL4 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C " & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If llngLocation > 0 Then
'            lsSQL4 = lsSQL4 & _
'                " AND B.LOC_CD = " & llngLocation
'        End If
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsSchedChgExcepInClause) > 0) Then
'                lsSQL4 = lsSQL4 & _
'                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL4 = lsSQL4 & _
'                    " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL4 = lsSQL4 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.HOURS <> 0" & _
'            " AND B.SCHED_CHG_EXC_CD <> ''"
'
'        lsSQL4 = lsSQL4 & " UNION ALL "
'
'        lsSQL5 = lsSQL5 & " SELECT A.JOB_NUM, A.EMP_NUM" & _
'            " , E.EMP_NAME, A.DATE_WORKED AS CHARGE_DATE" & _
'            " , A.START_TIME, A.END_TIME" & _
'            " , A.EXCEPTION_CD, '' AS OVER_EXC_CD" & _
'            " , '' AS SCHED_CHG_EXC_CD" & _
'            " , A.CASE_NUM, B.CASE_GROUP" & _
'            " , A.ADJ_TYPE" & _
'            " , CASE WHEN (A.PROCESSED_IND) = 'Y' THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , '' AS OT_OPP_DESC, '' AS DISTR_DEPT_NUM" & _
'            " , 0 AS AMOUNT, 0 AS TRAVEL_PAY_HRS, " & _
'            " CASE WHEN A.ADJ_TYPE = 'A' THEN (A.HOURS * -1) ELSE A.HOURS END AS HOURS" & _
'            " FROM  METHDBA.TTR_ADJUSTMENTS A" & _
'            " , METHDBA.TTR_EXCEPTION_CDS B" & _
'            " , METHDBA.TTR_EMP_WG D, METHDBA.TTR_EMPLOYEE E" & _
'            " WHERE " & lsJobSelection & _
'            " AND A.EMP_NUM = D.EMP_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EXCEPTION_CD = B.EXCEPTION_CD" & _
'            " AND A.DATE_WORKED BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.DATE_WORKED BETWEEN D.EFF_START_DATE" & _
'            " AND D.EFF_END_DATE"
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If (Len(lsExcepInClause) > 0) Then
'                    lsSQL5 = lsSQL5 & _
'                        " AND " & lsExcepInClause
'                Else
'                    lsSQL5 = lsSQL5 & _
'                        " AND A.EXCEPTION_CD <> 'ST'"
'                End If
'            Else
'                lsSQL5 = lsSQL5 & _
'                        " AND " & lsAllExcepInClause & _
'                        " AND A.EXCEPTION_CD <> 'ST'"
'            End If
'
'        lsSQL5 = lsSQL5 & _
'            " AND D.PRIMARY_IND = 'Y'" & _
'            " AND A.EXCEPTION_CD <> ''" & _
'            " AND A.HOURS <> 0" & _
'            " AND A.EXCEPTION_CD NOT IN ('ET','AT')"
'
'        lsSQL5 = lsSQL5 & " UNION ALL "
'
'    Else
'        'Get from Old and New Tables
'        lsSQL10 = lsSQL10 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
'            " , F.EMP_NAME, B.CHARGE_DATE" & _
'            " , B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL10 = lsSQL10 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.EXCEPTION_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If llngLocation > 0 Then
'            lsSQL10 = lsSQL10 & _
'                " AND B.LOC_CD = " & llngLocation
'        End If
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsExcepInClause) > 0) Then
'                lsSQL10 = lsSQL10 & _
'                    " AND " & Replace(lsExcepInClause, "A.", "B.")
'            Else
'                lsSQL10 = lsSQL10 & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'            End If
'        Else
'            lsSQL10 = lsSQL10 & _
'                    " AND " & Replace(lsAllExcepInClause, "A.", "B.") & _
'                    " AND B.EXCEPTION_CD <> 'ST'" & _
'                    " AND B.OVER_EXC_CD = ''" & _
'                    " AND B.SCHED_CHG_EXC_CD = ''"
'        End If
'
'        lsSQL10 = lsSQL10 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.HOURS <> 0" & _
'            " AND B.EXCEPTION_CD <> ''"
'
'        lsSQL10 = lsSQL10 & " UNION ALL "
'
'        lsSQL11 = lsSQL11 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
'            " , F.EMP_NAME, B.CHARGE_DATE" & _
'            " , B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL11 = lsSQL11 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.OVER_EXC_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If llngLocation > 0 Then
'            lsSQL11 = lsSQL11 & _
'                " AND B.LOC_CD = " & llngLocation
'        End If
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsOvrExcepInClause) > 0) Then
'                lsSQL11 = lsSQL11 & _
'                    " AND " & Replace(lsOvrExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL11 = lsSQL11 & _
'                    " AND " & Replace(lsAllOvrExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL11 = lsSQL11 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.HOURS <> 0" & _
'            " AND B.OVER_EXC_CD <> ''"
'
'        lsSQL11 = lsSQL11 & " UNION ALL "
'
'        lsSQL12 = lsSQL12 & " SELECT B.JOB_NUM, A.EMP_NUM" & _
'            " , F.EMP_NAME, B.CHARGE_DATE" & _
'            " , B.START_TIME, B.END_TIME" & _
'            " , B.EXCEPTION_CD, B.OVER_EXC_CD" & _
'            " , B.SCHED_CHG_EXC_CD" & _
'            " , B.CASE_NUM, C.CASE_GROUP" & _
'            " , CASE WHEN B.HOURS > 0 THEN 'S' ELSE 'A' END AS ADJ_TYPE" & _
'            " , CASE WHEN (A.FINAL_CYCLE_NUM) > 0 THEN '3'" & _
'            " ELSE '2' END AS PROCESSED_IND" & _
'            " , COALESCE(O.OT_OPP_DESC,B.OT_OPP_DESC) AS OT_OPP_DESC" & _
'            " , B.DISTR_DEPT_NUM" & _
'            " , B.AMOUNT, B.TRAVEL_PAY_HRS, B.HOURS"
'
'        lsSQL12 = lsSQL12 & " FROM METHDBA.TTR_ADJ_RQT A" & _
'            " , METHDBA.TTR_ADJ_TIME B LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O" & _
'            " ON B.OT_OPP_NUM = O.OT_OPP_NUM" & _
'            " , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " , METHDBA.TTR_EMP_WG E, METHDBA.TTR_EMPLOYEE F" & _
'            " WHERE " & Replace(lsJobSelection, "A.", "B.") & _
'            " AND B.CHARGE_DATE BETWEEN '" & CStr(v_datStartDate) & "'" & _
'            " AND '" & CStr(v_datEndDate) & "'" & _
'            " AND A.ADJ_RQT_NUM = B.ADJ_RQT_NUM" & _
'            " AND A.EMP_NUM = E.EMP_NUM" & _
'            " AND A.EMP_NUM = F.EMP_NUM" & _
'            " AND B.SCHED_CHG_EXC_CD = C.EXCEPTION_CD" & _
'            " AND B.CHARGE_DATE BETWEEN E.EFF_START_DATE" & _
'            " AND E.EFF_END_DATE"
'
'        If llngLocation > 0 Then
'            lsSQL12 = lsSQL12 & _
'                " AND B.LOC_CD = " & llngLocation
'        End If
'
'        If (RecordsExist(v_adrsSelectedExceptions)) Then
'            If (Len(lsSchedChgExcepInClause) > 0) Then
'                lsSQL12 = lsSQL12 & _
'                    " AND " & Replace(lsSchedChgExcepInClause, "A.", "B.")
'            End If
'        Else
'            lsSQL12 = lsSQL12 & _
'                    " AND " & Replace(lsAllSchedChgExcepInClause, "A.", "B.")
'        End If
'
'        lsSQL12 = lsSQL12 & _
'            " AND E.PRIMARY_IND = 'Y'" & _
'            " AND B.HOURS <> 0" & _
'            " AND B.SCHED_CHG_EXC_CD <> ''"
'
'        lsSQL12 = lsSQL12 & " UNION ALL "
'
'    End If
'
''END TRIS CONVERSION PHASE 3
''*******************************************************
''*  Retrieve Exceptions from Actual Time
''*******************************************************
'    If liActualTimeTables = TrTable_Actual Or _
'        liActualTimeTables = TrTable_ActualAndHistory Then
'        lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
'        lsSQL16 = lsSQL16 _
'            & "SELECT  A.job_num, A.emp_num,  d.emp_name, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, " _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL16 = lsSQL16 & "WHERE " & lsJobSelection _
'            & " AND A.charge_date between " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.exception_cd "
'
'            If llngLocation > 0 Then
'                lsSQL16 = lsSQL16 & " AND A.LOC_CD = " & llngLocation
'            End If
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsExcepInClause) > 0 Then
'                    lsSQL16 = lsSQL16 & " AND " & lsExcepInClause
'                Else
'                    lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'                End If
'            Else
'                lsSQL16 = lsSQL16 & " AND " & lsAllExcepInClause
'                lsSQL16 = lsSQL16 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'            End If
'
'        lsSQL16 = lsSQL16 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.EXCEPTION_CD <> ''"
'
'        lsSQL16 = lsSQL16 & " UNION ALL "
'
'        lsSQL17 = lsSQL17 _
'            & "SELECT  A.job_num, A.emp_num,  d.emp_name, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, " _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL17 = lsSQL17 & "WHERE " & lsJobSelection _
'            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.OVER_EXC_CD "
'
'            If llngLocation > 0 Then
'                lsSQL17 = lsSQL17 & " AND A.LOC_CD = " & llngLocation
'            End If
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsOvrExcepInClause) > 0 Then
'                    lsSQL17 = lsSQL17 & " AND " & lsOvrExcepInClause
'                End If
'            Else
'                lsSQL17 = lsSQL17 & " AND " & lsAllOvrExcepInClause
'            End If
'
'        lsSQL17 = lsSQL17 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.OVER_EXC_CD <> ''"
'
'        lsSQL17 = lsSQL17 & " UNION ALL "
'
'        lsSQL18 = lsSQL18 _
'            & "SELECT  A.job_num, A.emp_num,  d.emp_name, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, " _
'            & " CASE processed_ind " _
'            & " WHEN 'Y' THEN '3' " _
'            & " ELSE '2' END AS processed_ind, " _
'            & " COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, distr_dept_num, " _
'            & " amount, travel_pay_hrs, hours " _
'            & " FROM METHDBA.TTR_ACTUAL_TIME A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cdS B, " _
'            & " methdba.ttr_employee D "
'
'        lsSQL18 = lsSQL18 & "WHERE " & lsJobSelection _
'            & " AND A.charge_date between  " & DateAsEnclosedString(ldtActualTimeStartDate) & " AND " _
'            & DateAsEnclosedString(ldtActualTimeEndDate) _
'            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
'
'            If llngLocation > 0 Then
'                lsSQL18 = lsSQL18 & " AND A.LOC_CD = " & llngLocation
'            End If
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsSchedChgExcepInClause) > 0 Then
'                    lsSQL18 = lsSQL18 & " AND " & lsSchedChgExcepInClause
'                End If
'            Else
'                lsSQL18 = lsSQL18 & " AND " & lsAllSchedChgExcepInClause
'            End If
'
'        lsSQL18 = lsSQL18 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.SCHED_CHG_EXC_CD <> ''"
'
'        lsSQL18 = lsSQL18 & " UNION ALL "
'
'    End If
'
''*******************************************************
''*  Retrieve Exceptions from Actual Time History
''*******************************************************
'    If liActualTimeTables = TrTable_History Or _
'       liActualTimeTables = TrTable_ActualAndHistory Then
'       lsInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, ldtHistoryEndDate), True)
'       lsSQL19 = lsSQL19 _
'            & "SELECT A.job_num, A.emp_num,  d.emp_name, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, '3' AS processed_ind, " _
'            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee d "
'
'        lsSQL19 = lsSQL19 & " WHERE " & lsJobSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.exception_cd "
'
'            If llngLocation > 0 Then
'                lsSQL19 = lsSQL19 & " AND A.LOC_CD = " & llngLocation
'            End If
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsExcepInClause) > 0 Then
'                    lsSQL19 = lsSQL19 & " AND " & lsExcepInClause
'                Else
'                    lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'                End If
'            Else
'                lsSQL19 = lsSQL19 & " AND " & lsAllExcepInClause
'                lsSQL19 = lsSQL19 & " AND A.EXCEPTION_CD <> 'ST'" _
'                    & " AND A.OVER_EXC_CD = ''" _
'                    & " AND A.SCHED_CHG_EXC_CD = ''"
'            End If
'
'        lsSQL19 = lsSQL19 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.EXCEPTION_CD <> ''"
'
'        lsSQL19 = lsSQL19 & " UNION ALL "
'
'        lsSQL20 = lsSQL20 _
'            & "SELECT A.job_num, A.emp_num,  d.emp_name, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, '3' AS processed_ind, " _
'            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee d "
'
'        lsSQL20 = lsSQL20 & " WHERE " & lsJobSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.OVER_EXC_CD "
'
'            If llngLocation > 0 Then
'                lsSQL20 = lsSQL20 & " AND A.LOC_CD = " & llngLocation
'            End If
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsOvrExcepInClause) > 0 Then
'                    lsSQL20 = lsSQL20 & " AND " & lsOvrExcepInClause
'                End If
'            Else
'                lsSQL20 = lsSQL20 & " AND " & lsAllOvrExcepInClause
'            End If
'
'        lsSQL20 = lsSQL20 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.OVER_EXC_CD <> ''"
'
'        lsSQL20 = lsSQL20 & " UNION ALL "
'
'        lsSQL21 = lsSQL21 _
'            & "SELECT A.job_num, A.emp_num,  d.emp_name, " _
'            & "A.charge_date, A.start_time, A.end_time, A.exception_cd, " _
'            & "A.over_exc_cd, A.sched_chg_exc_cd, A.case_num, B.case_group, " _
'            & "' ' as adj_type, '3' AS processed_ind, " _
'            & "COALESCE(O.OT_OPP_DESC,A.OT_OPP_DESC) AS ot_opp_desc, " _
'            & "distr_dept_num, amount, travel_pay_hrs, hours " _
'            & "from  METHDBA.TTR_ACTUAL_TIME_HS A LEFT OUTER JOIN METHDBA.TTR_OT_OPPORTUNITY O " _
'            & " ON A.OT_OPP_NUM = O.OT_OPP_NUM, " _
'            & " METHDBA.TTR_exception_cds B, " _
'            & " methdba.ttr_employee d "
'
'        lsSQL21 = lsSQL21 & " WHERE " & lsJobSelection & " AND " & lsInClausePartValueKey _
'            & " AND A.charge_date between " & DateAsEnclosedString(v_datStartDate) & " AND " _
'            & DateAsEnclosedString(ldtHistoryEndDate) _
'            & " AND B.exception_cd = A.SCHED_CHG_EXC_CD "
'
'            If llngLocation > 0 Then
'                lsSQL21 = lsSQL21 & " AND A.LOC_CD = " & llngLocation
'            End If
'
'            If (RecordsExist(v_adrsSelectedExceptions)) Then
'                If Len(lsSchedChgExcepInClause) > 0 Then
'                    lsSQL21 = lsSQL21 & " AND " & lsSchedChgExcepInClause
'                End If
'            Else
'                lsSQL21 = lsSQL21 & " AND " & lsAllSchedChgExcepInClause
'            End If
'
'        lsSQL21 = lsSQL21 _
'            & " AND a.emp_num = d.emp_num " _
'            & " AND ((B.EXCEPTION_TYPE IN ('P', 'M'))) " _
'            & " AND A.SCHED_CHG_EXC_CD <> ''"
'
'        lsSQL21 = lsSQL21 & " UNION ALL "
'
'    End If
'
'    If (RecordsExist(v_adrsSelectedExceptions)) Then
'        If Len(lsExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL2 & lsSQL5 & lsSQL10 & lsSQL16 & lsSQL19
'        End If
'        If Len(lsOvrExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL3 & lsSQL11 & lsSQL17 & lsSQL20
'        End If
'        If Len(lsSchedChgExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL4 & lsSQL12 & lsSQL18 & lsSQL21
'        End If
'        If Len(lsOvrExcepInClause) > 0 Or Len(lsExcepInClause) > 0 Then
'            lsSQLALL = lsSQLALL & lsSQL1
'        End If
'    Else
'        lsSQLALL = lsSQLALL & lsSQL1 & lsSQL2 & lsSQL3 & lsSQL4 & lsSQL5 & lsSQL10 & _
'            lsSQL11 & lsSQL12 & lsSQL16 & lsSQL17 & lsSQL18 & lsSQL19 & lsSQL20 & lsSQL21
'    End If
'
'    If Trim(Right(lsSQLALL, 10)) = "UNION ALL" Then
'        lsSQLALL = Left(lsSQLALL, Len(lsSQLALL) - 10)
'    End If
'
'    If Len(lsSQLALL) > 0 Then
'        lsSQL = "SELECT JOB_NUM, EMP_NUM, EMP_NAME, CHARGE_DATE, " _
'        & " EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM , CASE_GROUP, " _
'        & " ADJ_TYPE, PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM, " _
'        & " SUM(AMOUNT) AS AMOUNT, SUM(TRAVEL_PAY_HRS) AS TRAVEL_PAY_HRS, " _
'        & " SUM(HOURS) AS HOURS FROM (" _
'        & lsSQLALL & ") NTE "
'        lsSQL = lsSQL & " GROUP BY JOB_NUM, EMP_NUM, EMP_NAME, " _
'        & " CHARGE_DATE, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CASE_NUM, " _
'        & " CASE_GROUP, ADJ_TYPE, PROCESSED_IND, OT_OPP_DESC, DISTR_DEPT_NUM"
'        lsSQL = lsSQL & " ORDER BY job_num, emp_num,  processed_ind, charge_date"
'    End If
'
'
'    Set lobjCN = New ADODB.Connection
'    Set lobjCM = New clsuDBConnectMgr
'    lobjCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
'    Set lobjCM = Nothing
'
'    lobjCN.Open
'
'    Set lobjRS = New ADODB.Recordset
'    lobjRS.CursorLocation = adUseClient
'    lobjRS.Open lsSQL, lobjCN, adOpenStatic, adLockBatchOptimistic
'    lobjRS.ActiveConnection = Nothing
'
'    lobjCN.Close
'    Set lobjCN = Nothing
'
'    Set GetJobExceptions2OLD = lobjRS
'    Set lobjRS = Nothing
'
'    '//-----clean up class variables used by madrsExceptions/GetExceptionCodes/IsExceptionCode-----//
'    DestroyState
'
'    Exit Function
'eh:
'    Set lobjCM = Nothing
'    Set lobjCN = Nothing
'    Set lobjRS = Nothing
'
'    DestroyState
'
'    TrsRaiseError m_strCMODULENAME _
'        , "GetJobExceptions2OLD" _
'        , , , , DescArgsFrom(lsSQL) _
'        , True
'
'End Function

Private Function DetermineActualTimeTables( _
                ByVal v_datStartDate As Date _
                , ByVal v_datEndDate As Date _
                , ByRef r_datHistoryEndDate As Date _
                , ByRef r_datActualTimeStartDate As Date _
                ) As TrActualTimeTable
On Error GoTo eh
Dim l_datCutOffChargeDate As Date

    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))

    If (v_datStartDate > l_datCutOffChargeDate) Then
        DetermineActualTimeTables = TrTable_Actual

        r_datActualTimeStartDate = v_datStartDate
    ElseIf (v_datEndDate <= l_datCutOffChargeDate) Then
        DetermineActualTimeTables = TrTable_History

        r_datHistoryEndDate = v_datEndDate
    Else
        DetermineActualTimeTables = TrTable_ActualAndHistory

        r_datHistoryEndDate = l_datCutOffChargeDate
        r_datActualTimeStartDate = DateAdd("d", 1, l_datCutOffChargeDate)
    End If

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineActualTimeTables"
End Function
Private Function GetActualTimeHSCutoffDate()
On Error GoTo GetActualTimeHSCutoffDate_EH
Dim l_objCodeTables As dbCodeTables.cdCodeTables
Dim l_varHSCutOffDate As Variant

    Set l_objCodeTables = CreateObject( _
                            "dbCodeTables.cdCodeTables")
    l_varHSCutOffDate = l_objCodeTables.GetActualTimeHSCutoffDate( _
                            CKEYCUTOFFDATE)
    If (IsDate(l_varHSCutOffDate)) Then
        GetActualTimeHSCutoffDate = CStr(l_varHSCutOffDate)
    Else
        Err.Raise trErrdb_GetHSCutOffDateInfo
    End If
 
    Set l_objCodeTables = Nothing
    Exit Function
GetActualTimeHSCutoffDate_EH:
    Set l_objCodeTables = Nothing

    Select Case (Err.Number)
        Case trErrdb_GetHSCutOffDateInfo
            TrsRaiseError m_strCMODULENAME _
                , "GetActualTimeHSCutoffDate" _
                , trErrdb_GetHSCutOffDateInfo _
                , , , DescArgsFrom(CStr(l_varHSCutOffDate)) _
                , True
    Case Else
            TrsRaiseError m_strCMODULENAME _
                , "GetActualTimeHSCutoffDate" _
                , , , , , True
    End Select

End Function

Private Function GetExceptionsEmpSchedSQL( _
    ByVal v_strEmployeeSelection As String, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date _
    ) As String
    
    Dim l_strSQL As String
    
 On Error GoTo eh
 
             
'*******************************************************
'*  Retrieve Scheduled Exceptions
'*******************************************************
    l_strSQL = "select EMP_NUM, CHARGE_DATE, A.EXCEPTION_CD, " _
        & "OVER_EXC_CD, B.CASE_GROUP, " _
        & "Left(B.EXCEPTION_BANK_CD,1) as BANK_CD, " _
        & "'N' AS FULL_SHIFT_EXC_IND,HOURS, SHIFT_NUM, B.SYMBOL_CD,' ' AS SHIFT_CODE " _
        & ", 'N' as SCHED_CHANGE_PMT , C.EG_SUB_CD " _
        & "From METHDBA.TTR_EMP_SCHED_EXC A, " _
        & "     METHDBA.TTR_EXCEPTION_CDS B LEFT OUTER JOIN METHDBA.TTR_EXC_GROUP C ON B.EXCEPTION_CD = C.EXCEPTION_CD " _
        & "where " & v_strEmployeeSelection _
        & "  and CHARGE_DATE >= " & DateAsEnclosedString(v_datStartDate) _
        & "  and CHARGE_DATE <= " & DateAsEnclosedString(v_datEndDate) _
        & "  and B.EXCEPTION_CD = A.EXCEPTION_CD " _
        & "  and ((B.EXCEPTION_TYPE = 'P') OR (OVER_EXC_CD <> '')) " _
        & "  and A.EXCEPTION_CD NOT IN ( SELECT EXCEPTION_CD  " _
                    & " FROM METHDBA.TTR_EXC_GROUP where EG_CD = 5 ) " _
        & "  and OVER_EXC_CD NOT IN ( SELECT EXCEPTION_CD  " _
                    & " FROM METHDBA.TTR_EXC_GROUP where EG_CD = 5  ) " _
        & " and C.EG_CD = 2" _
        & " and C.EXC_SEQ_NUM = 1 "

        
    GetExceptionsEmpSchedSQL = l_strSQL
 
     Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetExceptionsEmpSchedSQL"
    
    
    
End Function

Private Function GetExceptionsSchedOTSQL( _
    ByVal v_strEmployeeSelection As String, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date _
    ) As String
    
    Dim l_strSQL As String
    Dim l_strHoursSQL As String
    
 On Error GoTo eh
 
 'Get Scheduled Overtime
    l_strHoursSQL = _
        "case " & _
            "WHEN END_TIME = '00:00' or END_TIME = '24:00' then " & _
                "DECIMAL( (DECIMAL(24,3,1) " & _
                "- DECIMAL(Left(START_TIME,2),3,1) " & _
                "- DECIMAL(Right(START_TIME,2),3,1) / 60),3,1) " & _
            "Else " & _
                "DECIMAL( (DECIMAL(Left(END_TIME,2),3,1) " & _
                "+ DECIMAL(Right(END_TIME,2),3,1) / 60 " & _
                "- DECIMAL(Left(START_TIME,2),3,1) " & _
                "- DECIMAL(Right(START_TIME,2),3,1) / 60),3,1) " & _
             "End "
             
    l_strSQL = _
          "select EMP_NUM, CHARGE_DATE, A.EXCEPTION_CD, DATE, START_TIME, " _
        & " case END_TIME when '00:00' then '24:00' " _
            & " Else END_TIME end as END_TIME, " _
        & "OVER_EXC_CD, B.CASE_GROUP, " _
        & "Left(B.EXCEPTION_BANK_CD,1) as BANK_CD, " _
        & "'N' AS FULL_SHIFT_EXC_IND, " _
        & l_strHoursSQL & " as HOURS, " _
        & " 0 as SHIFT_NUM, B.SYMBOL_CD,' ' AS SHIFT_CODE " _
        & ", 'N' as SCHED_CHANGE_PMT , C.EG_SUB_CD " _
        & "From METHDBA.TTR_BASE_HRS_EXC A, " _
        & "     METHDBA.TTR_EXCEPTION_CDS B LEFT OUTER JOIN METHDBA.TTR_EXC_GROUP C ON B.EXCEPTION_CD = C.EXCEPTION_CD " _
        & "where " & v_strEmployeeSelection _
        & "  and CHARGE_DATE >= " & DateAsEnclosedString(v_datStartDate) _
        & "  and CHARGE_DATE <= " & DateAsEnclosedString(v_datEndDate) _
        & "  and B.EXCEPTION_CD = A.EXCEPTION_CD " _
        & "  and ((B.EXCEPTION_TYPE = 'P') OR (OVER_EXC_CD <> '') ) " _
        & "  and A.EXCEPTION_CD NOT IN ( SELECT EXCEPTION_CD  " _
                    & " FROM METHDBA.TTR_EXC_GROUP where EG_CD = 5 ) " _
        & "  and OVER_EXC_CD NOT IN ( SELECT EXCEPTION_CD  " _
                    & " FROM METHDBA.TTR_EXC_GROUP where EG_CD = 5  ) " _
        & " and C.EG_CD = 2" _
        & " and C.EXC_SEQ_NUM = 1 "
        
    GetExceptionsSchedOTSQL = l_strSQL
 
     Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetExceptionsSchedOTSQL"

    
End Function

Private Function GetExceptionsActTimeSQL( _
    ByVal v_ActTables As TrActualTimeTable, _
    ByVal v_strEmployeeSelection As String, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date _
    ) As String

    
    Dim l_strSQL As String
    Dim l_strActTable As String
    Dim lsInClausePartValueKey As String
    
    Select Case v_ActTables
        Case TrTable_Actual
            l_strActTable = "METHDBA.TTR_ACTUAL_TIME A, "
            lsInClausePartValueKey = ""
        Case TrTable_History
            l_strActTable = "METHDBA.TTR_ACTUAL_TIME_HS A, "
            lsInClausePartValueKey = " AND " & _
                BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
        Case Else
            GetExceptionsActTimeSQL = ""
            Exit Function
    End Select
    
    l_strSQL = _
             "SELECT A.emp_num, A.charge_date, A.exception_cd, DATE_WORKED, START_TIME, " _
            & " case END_TIME when '00:00' then '24:00' " _
            & " Else END_TIME end as END_TIME, " _
            & "over_exc_cd, B.case_group, " _
            & "Left(B.exception_bank_cd,1) AS bank_cd, " _
            & "'N' AS FULL_SHIFT_EXC_IND, hours, " _
            & "0 as shift_num, B.symbol_cd,' ' AS shift_code " _
            & ", 'N' as sched_change_pmt  , C.eg_sub_cd " _
            & "from " & l_strActTable _
            & "  METHDBA.TTR_EXCEPTION_CDS B LEFT OUTER JOIN METHDBA.TTR_EXC_GROUP C ON B.EXCEPTION_CD = C.EXCEPTION_CD "

    l_strSQL = l_strSQL & "where " & v_strEmployeeSelection & lsInClausePartValueKey _
            & "  and CHARGE_DATE >= " & DateAsEnclosedString(v_datStartDate) _
            & "  and CHARGE_DATE <= " & DateAsEnclosedString(v_datEndDate) _
            & "  and B.EXCEPTION_CD = A.EXCEPTION_CD " _
            & "  and ( (B.EXCEPTION_TYPE = 'P') OR (OVER_EXC_CD <> '') ) " _
            & "  and OTHER_PAY_SEQ_NUM = 0  " _
            & "  and A.EXCEPTION_CD NOT IN " _
            & "        ( SELECT EXCEPTION_CD FROM METHDBA.TTR_EXC_GROUP where EG_CD = 5 ) " _
            & "  and OVER_EXC_CD NOT IN  " _
            & "        ( SELECT EXCEPTION_CD FROM METHDBA.TTR_EXC_GROUP where EG_CD = 5 ) " _
            & " and C.EG_CD = 2" _
            & " and C.EXC_SEQ_NUM = 1 "
            
        
    
 On Error GoTo eh
 
    GetExceptionsActTimeSQL = l_strSQL
 
     Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetExceptionsActTimeSQL"
    
    
End Function

Private Function GetExceptionCodes()
On Error GoTo GetExceptionCodes_eh
    
    Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim lobjCN As ADODB.Connection
    Dim lobjRS As ADODB.Recordset
    Dim lsSQL As String

    
    lsSQL = "SELECT EXCEPTION_CD " _
            & "FROM METHDBA.TTR_EXCEPTION_CDS " _
            & "  WHERE EXCEPTION_TYPE <> 'S' " _
            & " AND SCHED_CHG_IND <> 'S' AND (CONTRACT_EARN_CD <> '' OR EXEC_EARN_CD <> '')"
            
    Set lobjCN = New ADODB.Connection
    Set lobjCM = New clsuDBConnectMgr
    lobjCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
    Set lobjCM = Nothing

    lobjCN.Open

    Set lobjRS = New ADODB.Recordset
    lobjRS.CursorLocation = adUseClient
    lobjRS.Open lsSQL, lobjCN, adOpenStatic, adLockBatchOptimistic
    lobjRS.ActiveConnection = Nothing

    lobjCN.Close
    Set lobjCN = Nothing

    Set madrsExceptions = lobjRS
    Set lobjRS = Nothing

    Exit Function
GetExceptionCodes_eh:
    Set lobjCM = Nothing
    Set lobjCN = Nothing
    Set lobjRS = Nothing
    Set madrsExceptions = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetExceptionsCodes" _
        , , , , DescArgsFrom(lsSQL) _
        , True

End Function
Private Function GetOverExcepCodes()
On Error GoTo GetOverExcepCodes_eh
    
    Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim lobjCN As ADODB.Connection
    Dim lobjRS As ADODB.Recordset
    Dim lsSQL As String

    
    lsSQL = "SELECT EXCEPTION_CD " _
            & "FROM METHDBA.TTR_EXCEPTION_CDS " _
            & "  WHERE CONTRACT_EARN_CD = '' AND EXEC_EARN_CD = '' "
            
    Set lobjCN = New ADODB.Connection
    Set lobjCM = New clsuDBConnectMgr
    lobjCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
    Set lobjCM = Nothing

    lobjCN.Open

    Set lobjRS = New ADODB.Recordset
    lobjRS.CursorLocation = adUseClient
    lobjRS.Open lsSQL, lobjCN, adOpenStatic, adLockBatchOptimistic
    lobjRS.ActiveConnection = Nothing

    lobjCN.Close
    Set lobjCN = Nothing

    Set madrsOverExcep = lobjRS
    Set lobjRS = Nothing

    Exit Function
GetOverExcepCodes_eh:
    Set lobjCM = Nothing
    Set lobjCN = Nothing
    Set lobjRS = Nothing
    Set madrsOverExcep = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetOverExcepCodes" _
        , , , , DescArgsFrom(lsSQL) _
        , True

End Function

Private Function GetSchedChgExcepCodes()
On Error GoTo GetSchedChgExcepCodes_eh
    
    Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim lobjCN As ADODB.Connection
    Dim lobjRS As ADODB.Recordset
    Dim lsSQL As String

    
    lsSQL = "SELECT EXCEPTION_CD " _
            & "FROM METHDBA.TTR_EXCEPTION_CDS " _
            & "  WHERE SCHED_CHG_IND = 'S' "
            
    Set lobjCN = New ADODB.Connection
    Set lobjCM = New clsuDBConnectMgr
    lobjCN.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
    Set lobjCM = Nothing

    lobjCN.Open

    Set lobjRS = New ADODB.Recordset
    lobjRS.CursorLocation = adUseClient
    lobjRS.Open lsSQL, lobjCN, adOpenStatic, adLockBatchOptimistic
    lobjRS.ActiveConnection = Nothing

    lobjCN.Close
    Set lobjCN = Nothing

    Set madrsSchedChgExcep = lobjRS
    Set lobjRS = Nothing

    Exit Function
GetSchedChgExcepCodes_eh:
    Set lobjCM = Nothing
    Set lobjCN = Nothing
    Set lobjRS = Nothing
    Set madrsSchedChgExcep = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetSchedChgExcepCodes" _
        , , , , DescArgsFrom(lsSQL) _
        , True

End Function
Private Function IsExceptionCode(ByVal v_exception_code As String) As Boolean
On Error GoTo IsExceptionCode_eh
    
    IsExceptionCode = False
    
    madrsExceptions.Filter = "EXCEPTION_CD = '" & v_exception_code & "'"
    If mADOExtensions.RecordsExist(madrsExceptions) Then
        IsExceptionCode = True
    Else
        IsExceptionCode = False
    End If
    
    madrsExceptions.Filter = adFilterNone
    
    Exit Function

IsExceptionCode_eh:
    
    TrsRaiseError m_strCMODULENAME, "IsExceptionCode"
    
End Function
Private Function IsOverExcepCode(ByVal v_exception_code As String) As Boolean
On Error GoTo IsOverExcepCode_eh
    
    IsOverExcepCode = False
    
    madrsOverExcep.Filter = "EXCEPTION_CD = '" & v_exception_code & "'"
    If mADOExtensions.RecordsExist(madrsOverExcep) Then
        IsOverExcepCode = True
    Else
        IsOverExcepCode = False
    End If
    
    madrsOverExcep.Filter = adFilterNone
    
    Exit Function

IsOverExcepCode_eh:
    
    TrsRaiseError m_strCMODULENAME, "IsOverExcepCode"
    
End Function


Private Function IsSchedChgExcepCode(ByVal v_exception_code As String) As Boolean
On Error GoTo IsSchedChgExcepCode_eh
    
    IsSchedChgExcepCode = False
    
    madrsSchedChgExcep.Filter = "EXCEPTION_CD = '" & v_exception_code & "'"
    If mADOExtensions.RecordsExist(madrsSchedChgExcep) Then
        IsSchedChgExcepCode = True
    Else
        IsSchedChgExcepCode = False
    End If
    
    Exit Function

IsSchedChgExcepCode_eh:
    
    TrsRaiseError m_strCMODULENAME, "IsSchedChgExcepCode"
    
End Function

Private Sub DestroyState()
    Set madrsExceptions = Nothing
    Set madrsOverExcep = Nothing
    Set madrsSchedChgExcep = Nothing
End Sub

