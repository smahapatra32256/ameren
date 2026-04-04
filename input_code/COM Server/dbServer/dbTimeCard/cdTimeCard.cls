VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdTimeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdTimecard"




Public Function GetAccountOverrideDescription(ByVal v_colOverExccd As Collection) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetAccountOverrideDescription_EH
    
    Dim l_strOverExccdforSQL As String
    If v_colOverExccd.Count < 1 Then
        l_strOverExccdforSQL = "('')"
        Err.Raise trErrdbGetEmptyColOverExccd
    Else
        l_strOverExccdforSQL = AsDelimitedString(v_colOverExccd, True)
    End If
    
   
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT EXCEPTION_CD, CAST(EXCEPTION_DESC AS CHAR(60)) AS EXCEPTION_DESC From METHDBA.TTR_EXCEPTION_CDS "
    l_strSQL = l_strSQL & "Where EXCEPTION_CD in " & l_strOverExccdforSQL
    l_strSQL = l_strSQL & " group by EXCEPTION_CD, EXCEPTION_DESC "
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetAccountOverrideDescription = l_objrs
    
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

GetAccountOverrideDescription_EH:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    Select Case Err.Number
        Case trErrdbGetEmptyColOverExccd
            TrsRaiseError m_strCMODULENAME, "GetAccountOverrideDescription", trErrdbGetEmptyColOverExccd, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetAccountOverrideDescription", , , , DescArgsFrom(l_strSQL), True
    End Select
    
End Function

'Rule WHITE046 - Please delete if no longer used
Public Function GetCaseNumbers( _
        ByVal v_lngEmpNum As Long, _
        ByVal v_strCaseGroup As String, _
        ByVal v_datStartDate As Date) As ADODB.Recordset
TrsRaiseError m_strCMODULENAME, "GetCaseNumbers", trErrObsoleteMethod, , , , True
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objrs As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'    Dim l_datStartDate As Date
'    Dim l_datEndDate As Date
'
'    On Error GoTo GetCaseNumbers_EH
'
'    l_datStartDate = DateSerial(Year(v_datStartDate), 1, 1)
'    l_datEndDate = DateSerial(Year(v_datStartDate), 12, 31)
'
'    Set l_objCN = New ADODB.Connection
'    Set l_objrs = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
'
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open
'
'    l_strSQL = "SELECT REFERENCE_NUM, START_DATE FROM METHDBA.TTR_ALL_TIME_CASES "
'    l_strSQL = l_strSQL & "WHERE EMP_NUM = " & v_lngEmpNum & " "
'    l_strSQL = l_strSQL & "AND CASE_GROUP = '" & Trim(v_strCaseGroup) & "'"
'    l_strSQL = l_strSQL & "AND START_DATE BETWEEN '" & Format(l_datStartDate, "mm/dd/yyyy") & "' "
'    l_strSQL = l_strSQL & "AND '" & Format(l_datEndDate, "mm/dd/yyyy") & "' "
'
'    l_objrs.CursorLocation = adUseClient
'    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objrs.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'
'    Set GetCaseNumbers = l_objrs
'
'    Set l_objrs = Nothing
'    GetObjectContext.SetComplete
'    Exit Function
'
'GetCaseNumbers_EH:
'    GetObjectContext.SetAbort
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objrs = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetCaseNumbers", , , , DescArgsFrom(l_strSQL), True

End Function
Public Function Get_OtherPay( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
        
    Dim l_lngEmpNum As Long
    Dim l_intStart As Integer
    Dim l_intEnd As Integer
    Dim i As Integer
    Dim l_datCutOffChargeDate As Date

On Error GoTo eh

    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set Get_OtherPay = Get_OtherPayACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_Actual)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set Get_OtherPay = Get_OtherPayACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_History)
        Else
           Set Get_OtherPay = Get_OtherPayACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_ActualAndHistory) 'use both current and history data
        End If
    End If


    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "Get_OtherPay", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "Get_OtherPay", , , , , True
    End Select

End Function

Private Function Get_OtherPayACTData( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_TrActualTimeTable As TrActualTimeTable _
            ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strInClausePartValueKey As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
    If v_TrActualTimeTable = TrTable_History Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT  EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, "
        l_strSQL = l_strSQL & "JOB_NUM, WG_NUM, DEPT_NUM, UNION_CD, HOURS, AMOUNT, "
        l_strSQL = l_strSQL & "START_LOC_CD, END_LOC_CD, LOC_CD, EARNINGS_CD, "
        l_strSQL = l_strSQL & "COMP_SHIFT_CD, COMP_SHIFT_CD_AMT, DISTRIBUTION_ID, "
        l_strSQL = l_strSQL & "CHEM_MTL_TSTR_IND, HALF_TRIP_IND, OT_OPP_NUM, "
        l_strSQL = l_strSQL & "DISTRIBUTION_TYPE, DISTR_DEPT_NUM, EXCEPTION_CD, APPR_IND, APPR_BY_EMP_NUM, "
        l_strSQL = l_strSQL & "VALUE(NULLIF(APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE, "
        l_strSQL = l_strSQL & "CHAR(INDEX_TIME) AS INDEX_TIME, DAY_NAME, ' ' AS RECORD_STATUS, LAST_MOD_EMP_NUM, ROW_ID "
        l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACTUAL_TIME_HS "
        l_strSQL = l_strSQL & "WHERE " & l_strEmployeeSelection & " "
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND START_TIME LIKE '99%' "
    End If
    If v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & " UNION ALL "
    End If
    If v_TrActualTimeTable = TrTable_Actual Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT  EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, "
        l_strSQL = l_strSQL & "JOB_NUM, WG_NUM, DEPT_NUM, UNION_CD, HOURS, AMOUNT, "
        l_strSQL = l_strSQL & "START_LOC_CD, END_LOC_CD, LOC_CD, EARNINGS_CD, "
        l_strSQL = l_strSQL & "COMP_SHIFT_CD, COMP_SHIFT_CD_AMT, DISTRIBUTION_ID, "
        l_strSQL = l_strSQL & "CHEM_MTL_TSTR_IND, HALF_TRIP_IND, OT_OPP_NUM, "
        l_strSQL = l_strSQL & "DISTRIBUTION_TYPE, DISTR_DEPT_NUM, EXCEPTION_CD, APPR_IND, APPR_BY_EMP_NUM, "
        l_strSQL = l_strSQL & "VALUE(NULLIF(APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE, "
        l_strSQL = l_strSQL & "CHAR(INDEX_TIME) AS INDEX_TIME, DAY_NAME, ' ' AS RECORD_STATUS, LAST_MOD_EMP_NUM, ROW_ID "
        l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACTUAL_TIME "
        l_strSQL = l_strSQL & "WHERE " & l_strEmployeeSelection & " "
        l_strSQL = l_strSQL & "AND CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND START_TIME LIKE '99%' "
    End If
    l_strSQL = l_strSQL & "ORDER BY EMP_NUM, CHARGE_DATE, START_TIME"

              
             
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing

    '  Will send back an empty recordset if no data retrieved
    Set Get_OtherPayACTData = l_objrs
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "Get_OtherPayACTData", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetOtherPayEntries( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset

    Dim l_lngEmpNum As Long
    Dim l_intStart As Integer
    Dim l_intEnd As Integer
    Dim i As Integer
    Dim l_datCutOffChargeDate As Date
    
On Error GoTo eh

    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set GetOtherPayEntries = GetOtherPayEntriesACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_Actual)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set GetOtherPayEntries = GetOtherPayEntriesACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_History)
        Else
           Set GetOtherPayEntries = GetOtherPayEntriesACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_ActualAndHistory) 'use both current and history data
        End If
    End If
    
    GetObjectContext.SetComplete
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetOtherPayEntries", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetOtherPayEntries", , , , , True
    End Select
    
End Function

Private Function GetOtherPayEntriesACTData( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_TrActualTimeTable As TrActualTimeTable _
            ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClausePartValueKey As String
    Dim l_strEmployeeSelection As String

    On Error GoTo eh
   
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
   

    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
    If v_TrActualTimeTable = TrTable_History Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, A.DISTRIBUTION_ID, A.HOURS "
        l_strSQL = l_strSQL & ", A.START_LOC_CD, A.END_LOC_CD, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", B.EXCEPTION_DESC AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, A.JOB_NUM, A.DISTR_DEPT_NUM, A.COMP_SHIFT_CD "
        l_strSQL = l_strSQL & ", A.CHEM_MTL_TSTR_IND, A.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(A.INDEX_TIME) As INDEX_TIME, A.AMOUNT, CHAR(A.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS A "
        l_strSQL = l_strSQL & ", METHDBA.TTR_EXCEPTION_CDS B "
    '  Call data service for in clause
        l_strSQL = l_strSQL & "WHERE " & l_strEmployeeSelection & " "
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND   A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND START_TIME LIKE '99%' "
        l_strSQL = l_strSQL & "AND   A.EXCEPTION_CD = B.EXCEPTION_CD "
    End If
    If v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & " UNION ALL "
    End If
    If v_TrActualTimeTable = TrTable_Actual Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, A.DISTRIBUTION_ID, A.HOURS "
        l_strSQL = l_strSQL & ", A.START_LOC_CD, A.END_LOC_CD, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", B.EXCEPTION_DESC AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, A.JOB_NUM, A.DISTR_DEPT_NUM, A.COMP_SHIFT_CD "
        l_strSQL = l_strSQL & ", A.CHEM_MTL_TSTR_IND, A.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(A.INDEX_TIME) As INDEX_TIME, A.AMOUNT, CHAR(A.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME A "
        l_strSQL = l_strSQL & ", METHDBA.TTR_EXCEPTION_CDS B "
    '  Call data service for in clause
        l_strSQL = l_strSQL & "WHERE " & l_strEmployeeSelection & " "
        l_strSQL = l_strSQL & "AND   A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND START_TIME LIKE '99%' "
        l_strSQL = l_strSQL & "AND   A.EXCEPTION_CD = B.EXCEPTION_CD "
    End If
    l_strSQL = l_strSQL & "ORDER BY EMP_NUM, CHARGE_DATE, START_TIME "


    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetOtherPayEntriesACTData = l_objrs
    
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetOtherPayEntriesACTData", , , , DescArgsFrom(l_strSQL), True
    
End Function

Public Function Get_TimeCard( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
        
    Dim l_datCutOffChargeDate As Date

    On Error GoTo eh

    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set Get_TimeCard = Get_TimeCardACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_Actual)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set Get_TimeCard = Get_TimeCardACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_History)
        Else
           Set Get_TimeCard = Get_TimeCardACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_ActualAndHistory) 'use both current and history data
        End If
    End If

    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort

    TrsRaiseError m_strCMODULENAME, "Get_TimeCard", , , , , True

End Function

Private Function Get_TimeCardACTData( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_TrActualTimeTable As TrActualTimeTable _
            ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClausePartValueKey As String
    Dim l_strEmployeeSelection As String

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
    If v_TrActualTimeTable = TrTable_History Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT EMP_NUM, CHARGE_DATE, START_TIME, CHAR(INDEX_TIME) AS INDEX_TIME, " _
            & "DATE_WORKED, END_TIME, DEPT_NUM, DISTR_DEPT_NUM, " _
            & "LOC_CD, WG_NUM, HOURS, " _
            & "JOB_NUM, EARNINGS_CD, " _
            & "UNION_CD , COMP_SHIFT_CD, COMP_SHIFT_CD_AMT, " _
            & "EXCEPTION_CD, DISTRIBUTION_ID, COMP_CD, COMP_AMT, " _
            & "VALUE(NULLIF(APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE, " _
            & "APPR_BY_EMP_NUM, CASE_NUM, " _
            & "OVER_EXC_CD, LAST_MOD_EMP_NUM, " _
            & "OT_REASON_CD, DAY_NAME, APPR_IND, " _
            & "REL_CD, FAMILY_LEAVE_IND, " _
            & "TRAVEL_PAY_HRS, DISTRIBUTION_TYPE, " _
            & "SCHED_CHG_EXC_CD , SCHED_CHG_EARN_CD, TRADE_EXC_CD, " _
            & "' ' AS RECORD_STATUS, ' ' AS SCHED_TIME_IND, " _
            & " DATE('0100-01-01') AS OT_BANK_CHRG_DT, OT_OPP_NUM, OT_OPP_DESC " _
            & " , SPACE(30) AS OT_ADJUST_DESC, CHAR(TIME_STAMP) AS TIME_STAMP  " _
            & " , INTERFACE_IND, SCHED_CD, SCHED_SRC_CD "
            
        l_strSQL = l_strSQL & " , JOB_NBR, JOBTASK_NBR, JOBTASK_STEP_NBR, EMPRV_DATABASE_ID, PREARRANGED_OT_IND, ROW_ID " & _
            " FROM METHDBA.TTR_ACTUAL_TIME_HS" & _
           " WHERE " & l_strEmployeeSelection & _
           " AND " & l_strInClausePartValueKey & _
           " AND START_TIME NOT LIKE '99%'" & _
           " AND CHARGE_DATE" & _
           " BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "'" & _
           " AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "'" & _
           " AND OT_OPP_NUM = 0 "
           
    ' The OT_OPP_NUM = 0 and the SQL may go away after Jim Barnett
    ' talks with the DBA
    ' Currently the OT_BANK_CHRG_DT defualts to the current system
    ' date if not provided in the insert statement.
    ' Please note: that the ORDER BY clause must stay
    
        l_strSQL = l_strSQL & " UNION ALL " _
            & "SELECT EMP_NUM, CHARGE_DATE, START_TIME, CHAR(INDEX_TIME) AS INDEX_TIME, " _
            & "DATE_WORKED, END_TIME, DEPT_NUM, DISTR_DEPT_NUM, " _
            & "LOC_CD, WG_NUM, HOURS, " _
            & "JOB_NUM, EARNINGS_CD, " _
            & "UNION_CD , COMP_SHIFT_CD, COMP_SHIFT_CD_AMT, " _
            & "EXCEPTION_CD, DISTRIBUTION_ID, COMP_CD, COMP_AMT, " _
            & "VALUE(NULLIF(APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE, " _
            & "APPR_BY_EMP_NUM, CASE_NUM, " _
            & "OVER_EXC_CD, LAST_MOD_EMP_NUM, " _
            & "OT_REASON_CD, DAY_NAME, APPR_IND, " _
            & "REL_CD, FAMILY_LEAVE_IND, " _
            & "TRAVEL_PAY_HRS, DISTRIBUTION_TYPE, " _
            & "SCHED_CHG_EXC_CD , SCHED_CHG_EARN_CD, TRADE_EXC_CD, " _
            & "' ' AS RECORD_STATUS, ' ' AS SCHED_TIME_IND, " _
            & " VALUE(NULLIF(OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT, OT_OPP_NUM, OT_OPP_DESC " _
            & " , SPACE(30) AS OT_ADJUST_DESC, CHAR(TIME_STAMP) AS TIME_STAMP  " _
            & " , INTERFACE_IND, SCHED_CD, SCHED_SRC_CD "
        l_strSQL = l_strSQL _
            & " , JOB_NBR, JOBTASK_NBR, JOBTASK_STEP_NBR, EMPRV_DATABASE_ID, PREARRANGED_OT_IND, ROW_ID " & _
            " FROM METHDBA.TTR_ACTUAL_TIME_HS" & _
           " WHERE " & l_strEmployeeSelection & _
           " AND " & l_strInClausePartValueKey & _
           " AND START_TIME NOT LIKE '99%'" & _
           " AND CHARGE_DATE" & _
           " BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "'" & _
           " AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "'" & _
           " AND OT_OPP_NUM > 0 "
    End If
    If v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & " UNION ALL "
    End If
    If v_TrActualTimeTable = TrTable_Actual Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT EMP_NUM, CHARGE_DATE, START_TIME, CHAR(INDEX_TIME) AS INDEX_TIME, " _
            & "DATE_WORKED, END_TIME, DEPT_NUM, DISTR_DEPT_NUM, " _
            & "LOC_CD, WG_NUM, HOURS, " _
            & "JOB_NUM, EARNINGS_CD, " _
            & "UNION_CD , COMP_SHIFT_CD, COMP_SHIFT_CD_AMT, " _
            & "EXCEPTION_CD, DISTRIBUTION_ID, COMP_CD, COMP_AMT, " _
            & "VALUE(NULLIF(APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE, " _
            & "APPR_BY_EMP_NUM, CASE_NUM, " _
            & "OVER_EXC_CD, LAST_MOD_EMP_NUM, " _
            & "OT_REASON_CD, DAY_NAME, APPR_IND, " _
            & "REL_CD, FAMILY_LEAVE_IND, " _
            & "TRAVEL_PAY_HRS, DISTRIBUTION_TYPE, " _
            & "SCHED_CHG_EXC_CD , SCHED_CHG_EARN_CD, TRADE_EXC_CD, " _
            & "' ' AS RECORD_STATUS, ' ' AS SCHED_TIME_IND, " _
            & " DATE('0100-01-01') AS OT_BANK_CHRG_DT, OT_OPP_NUM, OT_OPP_DESC " _
            & " , SPACE(30) AS OT_ADJUST_DESC, CHAR(TIME_STAMP) AS TIME_STAMP  " _
            & " , INTERFACE_IND, SCHED_CD, SCHED_SRC_CD " _
            & " , JOB_NBR, JOBTASK_NBR, JOBTASK_STEP_NBR, EMPRV_DATABASE_ID, PREARRANGED_OT_IND, ROW_ID " _
            & " FROM METHDBA.TTR_ACTUAL_TIME" & _
           " WHERE " & l_strEmployeeSelection & _
           " AND START_TIME NOT LIKE '99%'" & _
           " AND CHARGE_DATE" & _
           " BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "'" & _
           " AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "'" & _
           " AND OT_OPP_NUM = 0 "
           
    ' The OT_OPP_NUM = 0 and the SQL may go away after Jim Barnett
    ' talks with the DBA
    ' Currently the OT_BANK_CHRG_DT defualts to the current system
    ' date if not provided in the insert statement.
    ' Please note: that the ORDER BY clause must stay
    
        l_strSQL = l_strSQL & " UNION ALL " _
            & "SELECT EMP_NUM, CHARGE_DATE, START_TIME, CHAR(INDEX_TIME) AS INDEX_TIME, " _
            & "DATE_WORKED, END_TIME, DEPT_NUM, DISTR_DEPT_NUM, " _
            & "LOC_CD, WG_NUM, HOURS, " _
            & "JOB_NUM, EARNINGS_CD, " _
            & "UNION_CD , COMP_SHIFT_CD, COMP_SHIFT_CD_AMT, " _
            & "EXCEPTION_CD, DISTRIBUTION_ID, COMP_CD, COMP_AMT, " _
            & "VALUE(NULLIF(APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE, " _
            & "APPR_BY_EMP_NUM, CASE_NUM, " _
            & "OVER_EXC_CD, LAST_MOD_EMP_NUM, " _
            & "OT_REASON_CD, DAY_NAME, APPR_IND, " _
            & "REL_CD, FAMILY_LEAVE_IND, " _
            & "TRAVEL_PAY_HRS, DISTRIBUTION_TYPE, " _
            & "SCHED_CHG_EXC_CD , SCHED_CHG_EARN_CD, TRADE_EXC_CD, " _
            & "' ' AS RECORD_STATUS, ' ' AS SCHED_TIME_IND, " _
            & " VALUE(NULLIF(OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT, OT_OPP_NUM, OT_OPP_DESC " _
            & " , SPACE(30) AS OT_ADJUST_DESC, CHAR(TIME_STAMP) AS TIME_STAMP  " _
            & " , INTERFACE_IND, SCHED_CD, SCHED_SRC_CD "
            
        l_strSQL = l_strSQL _
            & " , JOB_NBR, JOBTASK_NBR, JOBTASK_STEP_NBR, EMPRV_DATABASE_ID, PREARRANGED_OT_IND, ROW_ID " & _
            " FROM METHDBA.TTR_ACTUAL_TIME" & _
           " WHERE " & l_strEmployeeSelection & _
           " AND START_TIME NOT LIKE '99%'" & _
           " AND CHARGE_DATE" & _
           " BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "'" & _
           " AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "'" & _
           " AND OT_OPP_NUM > 0 "
    End If
    
    l_strSQL = l_strSQL & " ORDER BY EMP_NUM, CHARGE_DATE, START_TIME"

              
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing

    Set Get_TimeCardACTData = l_objrs
    
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing

    TrsRaiseError m_strCMODULENAME, "Get_TimeCardACTData", , , , DescArgsFrom(l_strSQL), True

End Function

'//for older clients 6.9.4, 6.9.5
Public Function GetActualTimeEntries( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
    On Error GoTo eh
    
    Set GetActualTimeEntries = GetActualTimeEntries2(v_aryEmployees, v_datStartDate, v_datEndDate, 30)

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTimeEntries", , , , , True
End Function

'//for newer clients 6.9.7 +
Public Function GetActualTimeEntries2( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_intDescriptionLength As Integer) As ADODB.Recordset
' RED100
    Dim l_datCutOffChargeDate As Date

    On Error GoTo eh
    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set GetActualTimeEntries2 = GetActualTimeEntriesACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_Actual, v_intDescriptionLength)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set GetActualTimeEntries2 = GetActualTimeEntriesACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_History, v_intDescriptionLength)
        Else
           Set GetActualTimeEntries2 = GetActualTimeEntriesACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_ActualAndHistory, v_intDescriptionLength) 'use both current and history data
        End If
    End If

    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetActualTimeEntries2", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetActualTimeEntries2", , , , , True
    End Select

End Function

Private Function GetActualTimeEntriesACTData( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_TrActualTimeTable As TrActualTimeTable, _
        Optional ByVal v_intDescriptionLength As Integer = 39 _
            ) As ADODB.Recordset
' RED100
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClausePartValueKey As String

    On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
    If v_TrActualTimeTable = TrTable_History Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT       A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, C.OT_REASON_DESC, B.OT_OPP_DESC, A.JOB_NUM, A.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM, A.FAMILY_LEAVE_IND,  A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.CHEM_MTL_TSTR_IND, A.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD, B.OT_OPP_NUM, A.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(A.INDEX_TIME) As INDEX_TIME, A.AMOUNT, A.DAY_NAME "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS TIME_STAMP, A.SCHED_CD, A.SCHED_SRC_CD "
        l_strSQL = l_strSQL & ", CASE WHEN A.JOB_NBR <> ' ' THEN A.JOB_NBR || A.JOBTASK_NBR ELSE ' ' END AS JOB_TASK"
        l_strSQL = l_strSQL & " , A.JOBTASK_STEP_NBR, A.EMPRV_DATABASE_ID, A.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS A, METHDBA.TTR_OT_OPPORTUNITY B "
        l_strSQL = l_strSQL & ", METHDBA.TTR_OT_REASONS C, METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("A.EMP_NUM", v_aryEmployees, False)
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND   A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND  ( (A.OT_OPP_NUM >= 300000 AND A.OT_OPP_NUM < 32000000) or "
        l_strSQL = l_strSQL & "       (A.OT_OPP_NUM >= 32200000 AND A.OT_OPP_NUM < 65000000) ) "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM = B.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   A.OT_REASON_CD = C.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   A.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT       A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, C.OT_REASON_DESC, B.OT_OPP_DESC, A.JOB_NUM, A.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM, A.FAMILY_LEAVE_IND,  A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.CHEM_MTL_TSTR_IND, A.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD, B.OT_OPP_NUM, Z.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(A.INDEX_TIME) As INDEX_TIME, A.AMOUNT, A.DAY_NAME "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS TIME_STAMP, A.SCHED_CD, A.SCHED_SRC_CD "
        l_strSQL = l_strSQL & ", CASE WHEN A.JOB_NBR <> ' ' THEN A.JOB_NBR || A.JOBTASK_NBR ELSE ' ' END AS JOB_TASK"
        l_strSQL = l_strSQL & " , A.JOBTASK_STEP_NBR, A.EMPRV_DATABASE_ID, A.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS A, METHDBA.TTR_OT_OPPORTUNITY B "
        l_strSQL = l_strSQL & ", METHDBA.TTR_OT_REASONS C, METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & ", METHDBA.TTR_CANV_SHEETS Z "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("A.EMP_NUM", v_aryEmployees, False)
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND   A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM > 1 "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM <= 299999 "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM = B.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   B.OT_OPP_NUM = Z.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   Z.OT_REASON_CD = C.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   A.EMP_NUM = Z.EMP_NUM "
        l_strSQL = l_strSQL & "AND   A.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT D.EMP_NUM, D.CHARGE_DATE, D.DATE_WORKED, D.DISTRIBUTION_ID, D.HOURS, D.UNION_CD "
        l_strSQL = l_strSQL & ", D.HOURLY_RATE, D.START_LOC_CD, D.END_LOC_CD, D.INTERFACE_IND, D.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", D.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", D.START_TIME, D.END_TIME, D.EXCEPTION_CD, D.OVER_EXC_CD, D.EARNINGS_CD "
        l_strSQL = l_strSQL & ", D.WG_NUM, F.OT_REASON_DESC, E.OT_OPP_DESC, D.JOB_NUM , D.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", D.COMP_SHIFT_CD, D.COMP_SHIFT_CD_AMT, D.COMP_CD, D.COMP_AMT, D.REL_CD "
        l_strSQL = l_strSQL & ", D.CASE_NUM, D.FAMILY_LEAVE_IND, D.SCHED_CHG_EARN_CD, D.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", D.CHEM_MTL_TSTR_IND, D.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", D.DEPT_NUM, D.LOC_CD, D.OT_OPP_NUM, F.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(D.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", D.APPR_BY_EMP_NUM, D.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(D.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", D.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(D.INDEX_TIME) As INDEX_TIME, D.AMOUNT, D.DAY_NAME "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", CHAR(D.TIME_STAMP) AS TIME_STAMP, D.SCHED_CD, D.SCHED_SRC_CD "
        l_strSQL = l_strSQL & ", CASE WHEN D.JOB_NBR <> ' ' THEN D.JOB_NBR || D.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", D.JOBTASK_STEP_NBR, D.EMPRV_DATABASE_ID, D.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS D, METHDBA.TTR_OT_NOCAN_OPP E, METHDBA.TTR_OT_REASONS F "
        l_strSQL = l_strSQL & ", METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("D.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND   D.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM > 31999999 "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM < 32200000 "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM = E.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   E.OT_REASON_CD = F.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   D.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT D.EMP_NUM, D.CHARGE_DATE, D.DATE_WORKED, D.DISTRIBUTION_ID, D.HOURS, D.UNION_CD "
        l_strSQL = l_strSQL & ", D.HOURLY_RATE, D.START_LOC_CD, D.END_LOC_CD, D.INTERFACE_IND, D.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", D.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", D.START_TIME, D.END_TIME, D.EXCEPTION_CD, D.OVER_EXC_CD, D.EARNINGS_CD "
        l_strSQL = l_strSQL & ", D.WG_NUM, F.OT_REASON_DESC, D.OT_OPP_DESC, D.JOB_NUM , D.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", D.COMP_SHIFT_CD, D.COMP_SHIFT_CD_AMT, D.COMP_CD, D.COMP_AMT, D.REL_CD "
        l_strSQL = l_strSQL & ", D.CASE_NUM, D.FAMILY_LEAVE_IND, D.SCHED_CHG_EARN_CD, D.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", D.CHEM_MTL_TSTR_IND, D.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", D.DEPT_NUM, D.LOC_CD, D.OT_OPP_NUM, D.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(D.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", D.APPR_BY_EMP_NUM, D.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(D.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", D.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(D.INDEX_TIME) As INDEX_TIME, D.AMOUNT, D.DAY_NAME "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", CHAR(D.TIME_STAMP) AS TIME_STAMP, D.SCHED_CD, D.SCHED_SRC_CD "
        l_strSQL = l_strSQL & ", CASE WHEN D.JOB_NBR <> ' ' THEN D.JOB_NBR || D.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", D.JOBTASK_STEP_NBR, D.EMPRV_DATABASE_ID, D.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS D, METHDBA.TTR_OT_REASONS F "
        l_strSQL = l_strSQL & ", METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("D.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND   D.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM = 90000000 "
        l_strSQL = l_strSQL & "AND   D.OT_REASON_CD = F.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   D.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT G.EMP_NUM, G.CHARGE_DATE, G.DATE_WORKED, G.DISTRIBUTION_ID, G.HOURS, G.UNION_CD "
        l_strSQL = l_strSQL & ", G.HOURLY_RATE, G.START_LOC_CD, G.END_LOC_CD, G.INTERFACE_IND, G.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", G.TRADE_EXC_CD, H.EXCEPTION_BANK_CD, CAST(H.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", G.START_TIME, G.END_TIME, G.EXCEPTION_CD, G.OVER_EXC_CD, G.EARNINGS_CD "
        l_strSQL = l_strSQL & ", G.WG_NUM, '' AS OT_REASON_DESC, G.OT_OPP_DESC, G.JOB_NUM, G.DISTR_DEPT_NUM, G.COMP_SHIFT_CD, G.COMP_SHIFT_CD_AMT "
        l_strSQL = l_strSQL & ", G.COMP_CD, G.COMP_AMT, G.REL_CD, G.CASE_NUM, G.FAMILY_LEAVE_IND, G.SCHED_CHG_EARN_CD "
        l_strSQL = l_strSQL & ", G.SCHED_CHG_EXC_CD, G.CHEM_MTL_TSTR_IND, G.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", G.DEPT_NUM, G.LOC_CD, G.OT_OPP_NUM, G.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(G.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", G.APPR_BY_EMP_NUM, G.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(G.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", G.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(G.INDEX_TIME) As INDEX_TIME, G.AMOUNT, G.DAY_NAME "
        ' The Line commented out below may replace the line immediately below it
        ' Jim Barnett needs to talk to the DBA
        ' OT_BANL_CHRG_DT defaults to the current date if not supplied
        ' in an insert statement
        'l_strSQL = l_strSQL & ", VALUE(NULLIF(G.OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", DATE('0100-01-01') as OT_BANK_CHRG_DT, CHAR(G.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", G.SCHED_CD, G.SCHED_SRC_CD "
        l_strSQL = l_strSQL & ", CASE WHEN G.JOB_NBR <> ' ' THEN G.JOB_NBR || G.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", G.JOBTASK_STEP_NBR, G.EMPRV_DATABASE_ID, G.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACTUAL_TIME_HS G, METHDBA.TTR_EXCEPTION_CDS H "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("G.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND   CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   OT_OPP_NUM = 0 "
        l_strSQL = l_strSQL & "AND   G.EXCEPTION_CD = H.EXCEPTION_CD "
    End If
    If v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & " UNION ALL "
    End If
    If v_TrActualTimeTable = TrTable_Actual Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
       
        l_strSQL = l_strSQL & "SELECT       A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, C.OT_REASON_DESC, B.OT_OPP_DESC, A.JOB_NUM, A.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM, A.FAMILY_LEAVE_IND,  A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.CHEM_MTL_TSTR_IND, A.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD, B.OT_OPP_NUM, A.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(A.INDEX_TIME) As INDEX_TIME, A.AMOUNT, A.DAY_NAME "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS TIME_STAMP, A.SCHED_CD, A.SCHED_SRC_CD "
        l_strSQL = l_strSQL & " , CASE WHEN A.JOB_NBR <> ' ' THEN A.JOB_NBR || A.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", A.JOBTASK_STEP_NBR, A.EMPRV_DATABASE_ID, A.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME A, METHDBA.TTR_OT_OPPORTUNITY B "
        l_strSQL = l_strSQL & ", METHDBA.TTR_OT_REASONS C, METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("A.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & "AND   A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND  ( (A.OT_OPP_NUM >= 300000 AND A.OT_OPP_NUM < 32000000) or "
        l_strSQL = l_strSQL & "       (A.OT_OPP_NUM >= 32200000 AND A.OT_OPP_NUM < 65000000) ) "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM = B.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   A.OT_REASON_CD = C.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   A.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT       A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, C.OT_REASON_DESC, B.OT_OPP_DESC, A.JOB_NUM, A.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM, A.FAMILY_LEAVE_IND,  A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.CHEM_MTL_TSTR_IND, A.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD, B.OT_OPP_NUM, Z.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(A.INDEX_TIME) As INDEX_TIME, A.AMOUNT, A.DAY_NAME "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS TIME_STAMP, A.SCHED_CD, A.SCHED_SRC_CD "
        l_strSQL = l_strSQL & ", CASE WHEN A.JOB_NBR <> ' ' THEN A.JOB_NBR || A.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", A.JOBTASK_STEP_NBR, A.EMPRV_DATABASE_ID, A.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME A, METHDBA.TTR_OT_OPPORTUNITY B "
        l_strSQL = l_strSQL & ", METHDBA.TTR_OT_REASONS C, METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & ", METHDBA.TTR_CANV_SHEETS Z "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("A.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & "AND   A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM > 1 "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM <= 299999 "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM = B.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   B.OT_OPP_NUM = Z.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   Z.OT_REASON_CD = C.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   A.EMP_NUM = Z.EMP_NUM "
        l_strSQL = l_strSQL & "AND   A.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT D.EMP_NUM, D.CHARGE_DATE, D.DATE_WORKED, D.DISTRIBUTION_ID, D.HOURS, D.UNION_CD "
        l_strSQL = l_strSQL & ", D.HOURLY_RATE, D.START_LOC_CD, D.END_LOC_CD, D.INTERFACE_IND, D.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", D.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", D.START_TIME, D.END_TIME, D.EXCEPTION_CD, D.OVER_EXC_CD, D.EARNINGS_CD "
        l_strSQL = l_strSQL & ", D.WG_NUM, F.OT_REASON_DESC, E.OT_OPP_DESC, D.JOB_NUM , D.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", D.COMP_SHIFT_CD, D.COMP_SHIFT_CD_AMT, D.COMP_CD, D.COMP_AMT, D.REL_CD "
        l_strSQL = l_strSQL & ", D.CASE_NUM, D.FAMILY_LEAVE_IND, D.SCHED_CHG_EARN_CD, D.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", D.CHEM_MTL_TSTR_IND, D.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", D.DEPT_NUM, D.LOC_CD, D.OT_OPP_NUM, F.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(D.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", D.APPR_BY_EMP_NUM, D.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(D.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", D.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(D.INDEX_TIME) As INDEX_TIME, D.AMOUNT, D.DAY_NAME "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", CHAR(D.TIME_STAMP) AS TIME_STAMP, D.SCHED_CD, D.SCHED_SRC_CD "
        l_strSQL = l_strSQL & ", CASE WHEN D.JOB_NBR <> ' ' THEN D.JOB_NBR || D.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", D.JOBTASK_STEP_NBR, D.EMPRV_DATABASE_ID, D.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME D, METHDBA.TTR_OT_NOCAN_OPP E, METHDBA.TTR_OT_REASONS F "
        l_strSQL = l_strSQL & ", METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("D.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & "AND   D.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM > 31999999 "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM < 32200000 "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM = E.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   E.OT_REASON_CD = F.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   D.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT D.EMP_NUM, D.CHARGE_DATE, D.DATE_WORKED, D.DISTRIBUTION_ID, D.HOURS, D.UNION_CD "
        l_strSQL = l_strSQL & ", D.HOURLY_RATE, D.START_LOC_CD, D.END_LOC_CD, D.INTERFACE_IND, D.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", D.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", D.START_TIME, D.END_TIME, D.EXCEPTION_CD, D.OVER_EXC_CD, D.EARNINGS_CD "
        l_strSQL = l_strSQL & ", D.WG_NUM, F.OT_REASON_DESC, D.OT_OPP_DESC, D.JOB_NUM , D.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", D.COMP_SHIFT_CD, D.COMP_SHIFT_CD_AMT, D.COMP_CD, D.COMP_AMT, D.REL_CD "
        l_strSQL = l_strSQL & ", D.CASE_NUM, D.FAMILY_LEAVE_IND, D.SCHED_CHG_EARN_CD, D.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", D.CHEM_MTL_TSTR_IND, D.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", D.DEPT_NUM, D.LOC_CD, D.OT_OPP_NUM, D.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(D.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", D.APPR_BY_EMP_NUM, D.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(D.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", D.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(D.INDEX_TIME) As INDEX_TIME, D.AMOUNT, D.DAY_NAME "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", CHAR(D.TIME_STAMP) AS TIME_STAMP, D.SCHED_CD, D.SCHED_SRC_CD "
        l_strSQL = l_strSQL & " , CASE WHEN D.JOB_NBR <> ' ' THEN D.JOB_NBR || D.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", D.JOBTASK_STEP_NBR, D.EMPRV_DATABASE_ID, D.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME D, METHDBA.TTR_OT_REASONS F "
        l_strSQL = l_strSQL & ", METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("D.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & "AND   D.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM = 90000000 "
        l_strSQL = l_strSQL & "AND   D.OT_REASON_CD = F.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   D.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT G.EMP_NUM, G.CHARGE_DATE, G.DATE_WORKED, G.DISTRIBUTION_ID, G.HOURS, G.UNION_CD "
        l_strSQL = l_strSQL & ", G.HOURLY_RATE, G.START_LOC_CD, G.END_LOC_CD, G.INTERFACE_IND, G.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", G.TRADE_EXC_CD, H.EXCEPTION_BANK_CD, CAST(H.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", G.START_TIME, G.END_TIME, G.EXCEPTION_CD, G.OVER_EXC_CD, G.EARNINGS_CD "
        l_strSQL = l_strSQL & ", G.WG_NUM, '' AS OT_REASON_DESC, G.OT_OPP_DESC, G.JOB_NUM, G.DISTR_DEPT_NUM, G.COMP_SHIFT_CD, G.COMP_SHIFT_CD_AMT "
        l_strSQL = l_strSQL & ", G.COMP_CD, G.COMP_AMT, G.REL_CD, G.CASE_NUM, G.FAMILY_LEAVE_IND, G.SCHED_CHG_EARN_CD "
        l_strSQL = l_strSQL & ", G.SCHED_CHG_EXC_CD, G.CHEM_MTL_TSTR_IND, G.HALF_TRIP_IND "
        l_strSQL = l_strSQL & ", G.DEPT_NUM, G.LOC_CD, G.OT_OPP_NUM, G.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(G.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", G.APPR_BY_EMP_NUM, G.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(G.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", G.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(G.INDEX_TIME) As INDEX_TIME, G.AMOUNT, G.DAY_NAME "
        ' The Line commented out below may replace the line immediately below it
        ' Jim Barnett needs to talk to the DBA
        ' OT_BANL_CHRG_DT defaults to the current date if not supplied
        ' in an insert statement
        'l_strSQL = l_strSQL & ", VALUE(NULLIF(G.OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", DATE('0100-01-01') as OT_BANK_CHRG_DT, CHAR(G.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", G.SCHED_CD, G.SCHED_SRC_CD "
        l_strSQL = l_strSQL & " , CASE WHEN G.JOB_NBR <> ' ' THEN G.JOB_NBR || G.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", G.JOBTASK_STEP_NBR, G.EMPRV_DATABASE_ID, G.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACTUAL_TIME G, METHDBA.TTR_EXCEPTION_CDS H "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("G.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & "AND   CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   OT_OPP_NUM = 0 "
        l_strSQL = l_strSQL & "AND   G.EXCEPTION_CD = H.EXCEPTION_CD "
    End If
    l_strSQL = l_strSQL & "ORDER BY EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME "



    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetActualTimeEntriesACTData = l_objrs
    
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetActualTimeEntriesACTData", , , , DescArgsFrom(l_strSQL), True

End Function

''//-----The following was created to allow the Callaway WebService to run faster
'Public Function GetActualTimeEntriesForDept( _
'        ByVal v_aryDepartments As Variant, _
'        ByVal v_datStartDate As Date, _
'        ByVal v_datEndDate As Date) As ADODB.Recordset
'' RED100
'    Dim l_datCutOffChargeDate As Date
'
'    On Error GoTo eh
'
'    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
'    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
'        Set GetActualTimeEntriesForDept = GetActualTimeEntriesACTDataForDept(v_aryDepartments, v_datStartDate, v_datEndDate, TrTable_Actual)
'    Else
'        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
'           Set GetActualTimeEntriesForDept = GetActualTimeEntriesACTDataForDept(v_aryDepartments, v_datStartDate, v_datEndDate, TrTable_History)
'        Else
'           Set GetActualTimeEntriesForDept = GetActualTimeEntriesACTDataForDept(v_aryDepartments, v_datStartDate, v_datEndDate, TrTable_ActualAndHistory) 'use both current and history data
'        End If
'    End If
'
'    GetObjectContext.SetComplete
'    Exit Function
'
'eh:
'    GetObjectContext.SetAbort
'
'    Select Case Err.Number
'        Case trErrdbInvalidZerolngEmpNum
'            TrsRaiseError m_strCMODULENAME, "GetActualTimeEntriesForDept", trErrdbInvalidZerolngEmpNum, , , , True
'        Case Else
'            TrsRaiseError m_strCMODULENAME, "GetActualTimeEntriesForDept", , , , , True
'    End Select
'
'End Function
'
'Private Function GetActualTimeEntriesACTDataForDept( _
'        ByVal v_aryDepartments As Variant, _
'        ByVal v_datStartDate As Date, _
'        ByVal v_datEndDate As Date, _
'        ByVal v_TrActualTimeTable As TrActualTimeTable _
'            ) As ADODB.Recordset
'' RED100
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objrs As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'    Dim l_strInClausePartValueKey As String
'
'    On Error GoTo eh
'
'    Set l_objCN = New ADODB.Connection
'    Set l_objrs = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
'
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open
'
'    l_strSQL = l_strSQL & "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_SELECTION_00 "
'    l_strSQL = l_strSQL & "(DEPT_NUM CHAR(3) NOT NULL "
'    l_strSQL = l_strSQL & ",START_DATE DATE NOT NULL "
'    l_strSQL = l_strSQL & ",END_DATE DATE NOT NULL "
'    l_strSQL = l_strSQL & ") "
'    l_strSQL = l_strSQL & "CCSID EBCDIC "
'    l_strSQL = l_strSQL & "ON COMMIT PRESERVE ROWS "
'    l_strSQL = l_strSQL & "; "
'    l_strSQL = l_strSQL & "INSERT INTO SESSION.GTT_SELECTION_00 "
'    l_strSQL = l_strSQL & "(DEPT_NUM,START_DATE,END_DATE) "
'    l_strSQL = l_strSQL & "SELECT "
'    l_strSQL = l_strSQL & "DEPT_NUM "
'    l_strSQL = l_strSQL & ",'" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
'    l_strSQL = l_strSQL & ",'" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
'    l_strSQL = l_strSQL & "From "
'    l_strSQL = l_strSQL & "METHDBA.TTR_DEPT "
'    l_strSQL = l_strSQL & "WHERE " & BuildSelection("DEPT_NUM", v_aryDepartments, True) & " "
'    l_strSQL = l_strSQL & "Group By "
'    l_strSQL = l_strSQL & "DEPT_NUM "
'    l_strSQL = l_strSQL & "; "
'    l_objCN.Execute l_strSQL
'
'    l_strSQL = ""
'    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
'    If v_TrActualTimeTable = TrTable_History Or _
'       v_TrActualTimeTable = TrTable_ActualAndHistory Then
'        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, E.EMP_NAME, A.DATE_WORKED, A.CHARGE_DATE, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
'        l_strSQL = l_strSQL & ", I.EXCEPTION_DESC AS DISTRIBUTION_DESC, A.EXCEPTION_CD, A.OVER_EXC_CD, A.SCHED_CHG_EXC_CD, A.WG_NUM, A.JOB_NUM, A.DISTR_DEPT_NUM "
'        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.DEPT_NUM, A.LOC_CD, A.SCHED_CD "
'        l_strSQL = l_strSQL & ", DISTRIBUTION_TYPE, A.START_TIME, '0' AS TIME_ENTRY_TYPE, EXCEPTION_BANK_CD, APPR_IND "
'        l_strSQL = l_strSQL & "From "
'        l_strSQL = l_strSQL & "Session.GTT_SELECTION_00 Z "
'        l_strSQL = l_strSQL & " WHERE " & l_strInClausePartValueKey
'        l_strSQL = l_strSQL & "Inner Join "
'        l_strSQL = l_strSQL & "METHDBA.TTR_ACTUAL_TIME A "
'        l_strSQL = l_strSQL & "ON A.DEPT_NUM = Z.DEPT_NUM "
'        l_strSQL = l_strSQL & "AND A.CHARGE_DATE BETWEEN Z.START_DATE AND Z.END_DATE "
'        l_strSQL = l_strSQL & "Inner Join "
'        l_strSQL = l_strSQL & "METHDBA.TTR_EMPLOYEE E "
'        l_strSQL = l_strSQL & "ON A.EMP_NUM = E.EMP_NUM "
'        l_strSQL = l_strSQL & "Inner Join "
'        l_strSQL = l_strSQL & "METHDBA.TTR_EXCEPTION_CDS i "
'        l_strSQL = l_strSQL & "ON A.EXCEPTION_CD = I.EXCEPTION_CD "
'    End If
'    If v_TrActualTimeTable = TrTable_ActualAndHistory Then
'        l_strSQL = l_strSQL & " UNION ALL "
'    End If
'    If v_TrActualTimeTable = TrTable_Actual Or _
'       v_TrActualTimeTable = TrTable_ActualAndHistory Then
'        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, E.EMP_NAME, A.DATE_WORKED, A.CHARGE_DATE, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
'        l_strSQL = l_strSQL & ", I.EXCEPTION_DESC AS DISTRIBUTION_DESC, A.EXCEPTION_CD, A.OVER_EXC_CD, A.SCHED_CHG_EXC_CD, A.WG_NUM, A.JOB_NUM, A.DISTR_DEPT_NUM "
'        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.DEPT_NUM, A.LOC_CD, A.SCHED_CD "
'        l_strSQL = l_strSQL & ", DISTRIBUTION_TYPE, A.START_TIME, '0' AS TIME_ENTRY_TYPE, EXCEPTION_BANK_CD, APPR_IND "
'        l_strSQL = l_strSQL & "From "
'        l_strSQL = l_strSQL & "Session.GTT_SELECTION_00 Z "
'        l_strSQL = l_strSQL & "Inner Join "
'        l_strSQL = l_strSQL & "METHDBA.TTR_ACTUAL_TIME A "
'        l_strSQL = l_strSQL & "ON A.DEPT_NUM = Z.DEPT_NUM "
'        l_strSQL = l_strSQL & "AND A.CHARGE_DATE BETWEEN Z.START_DATE AND Z.END_DATE "
'        l_strSQL = l_strSQL & "Inner Join "
'        l_strSQL = l_strSQL & "METHDBA.TTR_EMPLOYEE E "
'        l_strSQL = l_strSQL & "ON A.EMP_NUM = E.EMP_NUM "
'        l_strSQL = l_strSQL & "Inner Join "
'        l_strSQL = l_strSQL & "METHDBA.TTR_EXCEPTION_CDS i "
'        l_strSQL = l_strSQL & "ON A.EXCEPTION_CD = I.EXCEPTION_CD "
'    End If
'
'    l_strSQL = l_strSQL & " ORDER BY A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.START_TIME, A.END_TIME, A.DISTRIBUTION_ID; "
'    l_strSQL = l_strSQL & "DROP TABLE SESSION.GTT_SELECTION_00; "
'
'    l_objrs.CursorLocation = adUseClient
'    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objrs.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'
'    Set GetActualTimeEntriesACTDataForDept = l_objrs
'
'    Set l_objrs = Nothing
'    GetObjectContext.SetComplete
'    Exit Function
'
'eh:
'    GetObjectContext.SetAbort
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objrs = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetActualTimeEntriesACTDataForDept", , , , DescArgsFrom(l_strSQL), True
'
'End Function
'
''Private Function GetActualTimeEntriesACTDataForDept( _
''        ByVal v_aryDepartments As Variant, _
''        ByVal v_datStartDate As Date, _
''        ByVal v_datEndDate As Date, _
''        ByVal v_TrActualTimeTable As TrActualTimeTable _
''            ) As ADODB.Recordset
''' RED100
''    Dim l_objCN As New ADODB.Connection
''    Dim l_objrs As ADODB.Recordset
''    Dim l_strSQL As String
''    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
''    Dim l_strInClausePartValueKey As String
''
''    On Error GoTo eh
''
''    Set l_objCN = New ADODB.Connection
''    Set l_objrs = New ADODB.Recordset
''    Set l_objCM = New clsuDBConnectMgr
''
''    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
''
''    Set l_objCM = Nothing
''
''    l_objCN.Open
''
''    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
''    If v_TrActualTimeTable = TrTable_History Or _
''       v_TrActualTimeTable = TrTable_ActualAndHistory Then
''        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, E.EMP_NAME, A.DATE_WORKED, A.CHARGE_DATE, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
''        l_strSQL = l_strSQL & ", I.EXCEPTION_DESC AS DISTRIBUTION_DESC, A.EXCEPTION_CD, A.OVER_EXC_CD, A.SCHED_CHG_EXC_CD, A.WG_NUM, A.JOB_NUM, A.DISTR_DEPT_NUM "
''        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.DEPT_NUM, A.LOC_CD, A.SCHED_CD "
''        l_strSQL = l_strSQL & ", DISTRIBUTION_TYPE, A.START_TIME, '0' AS TIME_ENTRY_TYPE, EXCEPTION_BANK_CD, APPR_IND "
''        l_strSQL = l_strSQL & "From "
''        l_strSQL = l_strSQL & "(SELECT * "
''        l_strSQL = l_strSQL & " From METHDBA.TTR_ACTUAL_TIME_HS "
''        'l_strSQL = l_strSQL & "  where dept_num in ({0})"
''        l_strSQL = l_strSQL & "  WHERE " & BuildSelection("DEPT_NUM", v_aryDepartments, True) & " "
''        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
''        'l_strSQL = l_strSQL & "  AND   CHARGE_DATE BETWEEN {1} and {2}"
''        l_strSQL = l_strSQL & "  AND   CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
''        l_strSQL = l_strSQL & ") A "
''        l_strSQL = l_strSQL & "inner join METHDBA.TTR_EMPLOYEE E "
''        l_strSQL = l_strSQL & "ON A.EMP_NUM = E.EMP_NUM "
''        l_strSQL = l_strSQL & "inner join METHDBA.TTR_EXCEPTION_CDS I "
''        l_strSQL = l_strSQL & "on A.EXCEPTION_CD = I.EXCEPTION_CD "
''        l_strSQL = l_strSQL & "left outer join METHDBA.TTR_OT_REASONS C"
''        l_strSQL = l_strSQL & "on a.OT_REASON_CD = c.OT_REASON_CD "
''        l_strSQL = l_strSQL & "left outer join METHDBA.TTR_OT_OPPORTUNITY B "
''        l_strSQL = l_strSQL & "on A.ot_opp_num between 300000 and 65000000 "
''        l_strSQL = l_strSQL & "and a.ot_opp_num = b.ot_opp_num "
''    End If
''    If v_TrActualTimeTable = TrTable_ActualAndHistory Then
''        l_strSQL = l_strSQL & " UNION ALL "
''    End If
''    If v_TrActualTimeTable = TrTable_Actual Or _
''       v_TrActualTimeTable = TrTable_ActualAndHistory Then
''        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, E.EMP_NAME, A.DATE_WORKED, A.CHARGE_DATE, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
''        l_strSQL = l_strSQL & ", I.EXCEPTION_DESC AS DISTRIBUTION_DESC, A.EXCEPTION_CD, A.OVER_EXC_CD, A.SCHED_CHG_EXC_CD, A.WG_NUM, A.JOB_NUM, A.DISTR_DEPT_NUM "
''        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.DEPT_NUM, A.LOC_CD, A.SCHED_CD "
''        l_strSQL = l_strSQL & ", DISTRIBUTION_TYPE, A.START_TIME, '0' AS TIME_ENTRY_TYPE, EXCEPTION_BANK_CD, APPR_IND "
''        l_strSQL = l_strSQL & "From "
''        l_strSQL = l_strSQL & "(SELECT * "
''        l_strSQL = l_strSQL & "From METHDBA.TTR_ACTUAL_TIME "
''        'l_strSQL = l_strSQL & "  where dept_num in ({0})"
''        l_strSQL = l_strSQL & "  WHERE " & BuildSelection("DEPT_NUM", v_aryDepartments, True) & " "
''        'l_strSQL = l_strSQL & "  AND   CHARGE_DATE BETWEEN {1} and {2}"
''        l_strSQL = l_strSQL & "  AND   CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
''        l_strSQL = l_strSQL & ") A "
''        l_strSQL = l_strSQL & "inner join METHDBA.TTR_EMPLOYEE E "
''        l_strSQL = l_strSQL & "ON A.EMP_NUM = E.EMP_NUM "
''        l_strSQL = l_strSQL & "inner join METHDBA.TTR_EXCEPTION_CDS I "
''        l_strSQL = l_strSQL & "on A.EXCEPTION_CD = I.EXCEPTION_CD "
''        l_strSQL = l_strSQL & "left outer join METHDBA.TTR_OT_REASONS C "
''        l_strSQL = l_strSQL & "on a.OT_REASON_CD = c.OT_REASON_CD "
''        l_strSQL = l_strSQL & "left outer join METHDBA.TTR_OT_OPPORTUNITY B "
''        l_strSQL = l_strSQL & "on A.ot_opp_num between 300000 and 65000000 "
''        l_strSQL = l_strSQL & "and a.ot_opp_num = b.ot_opp_num "
''    End If
''    l_strSQL = l_strSQL & " ORDER BY A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.START_TIME, A.END_TIME "
''
''
''    l_objrs.CursorLocation = adUseClient
''    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
''
''    l_objrs.ActiveConnection = Nothing
''    l_objCN.Close
''
''    Set l_objCN = Nothing
''
''    Set GetActualTimeEntriesACTDataForDept = l_objrs
''
''    Set l_objrs = Nothing
''    GetObjectContext.SetComplete
''    Exit Function
''
''eh:
''    GetObjectContext.SetAbort
''    Set l_objCM = Nothing
''    Set l_objCN = Nothing
''    Set l_objrs = Nothing
''
''    TrsRaiseError m_strCMODULENAME, "GetActualTimeEntriesACTDataForDept", , , , DescArgsFrom(l_strSQL), True
''
''End Function

'//-----
'//for older clients 6.9.4, 6.9.5
Public Function GetActualTimeEntriesWithoutOtherPay( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
    On Error GoTo eh
    
    Set GetActualTimeEntriesWithoutOtherPay = GetActualTimeEntriesWithoutOtherPay2(v_aryEmployees, v_datStartDate, v_datEndDate, 30)

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTimeEntriesWithoutOtherPay", , , , , True
End Function

'//for newer clients 6.9.7 +
Public Function GetActualTimeEntriesWithoutOtherPay2( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_intDescriptionLength As Integer) As ADODB.Recordset
' RED100
    Dim l_datCutOffChargeDate As Date

    On Error GoTo eh
    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set GetActualTimeEntriesWithoutOtherPay2 = GetActualTimeEntriesWithoutOtherPayACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_Actual, v_intDescriptionLength)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set GetActualTimeEntriesWithoutOtherPay2 = GetActualTimeEntriesWithoutOtherPayACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_History, v_intDescriptionLength)
        Else
           Set GetActualTimeEntriesWithoutOtherPay2 = GetActualTimeEntriesWithoutOtherPayACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_ActualAndHistory, v_intDescriptionLength) 'use both current and history data
        End If
    End If

    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetActualTimeEntriesWithoutOtherPay2", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetActualTimeEntriesWithoutOtherPay2", , , , , True
    End Select

End Function

Private Function GetActualTimeEntriesWithoutOtherPayACTData( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_TrActualTimeTable As TrActualTimeTable, _
        Optional ByVal v_intDescriptionLength As Integer = 39 _
            ) As ADODB.Recordset
    
' RED100
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClausePartValueKey As String
    
    On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
    If v_TrActualTimeTable = TrTable_History Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, C.OT_REASON_DESC, B.OT_OPP_DESC, A.JOB_NUM, A.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM, A.FAMILY_LEAVE_IND,  A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD, B.OT_OPP_NUM, A.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(A.INDEX_TIME) As INDEX_TIME, A.AMOUNT, A.DAY_NAME "
        l_strSQL = l_strSQL & ", NULLIF(A.OT_BANK_CHRG_DT,'0001-01-01') AS OT_BANK_CHRG_DT, CHAR(A.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", CASE WHEN A.JOB_NBR <> ' ' THEN A.JOB_NBR || A.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", A.JOBTASK_STEP_NBR, A.EMPRV_DATABASE_ID, A.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS A, METHDBA.TTR_OT_OPPORTUNITY B "
        l_strSQL = l_strSQL & ", METHDBA.TTR_OT_REASONS C, METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("A.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND   A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND  ( (A.OT_OPP_NUM >= 300000 AND A.OT_OPP_NUM < 32000000) or "
        l_strSQL = l_strSQL & "       (A.OT_OPP_NUM >= 32200000 AND A.OT_OPP_NUM < 65000000) ) "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM = B.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   A.OT_REASON_CD = C.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   A.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & "AND   A.START_TIME NOT LIKE '99%'"
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, C.OT_REASON_DESC, B.OT_OPP_DESC, A.JOB_NUM, A.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM, A.FAMILY_LEAVE_IND,  A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD, B.OT_OPP_NUM, Z.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(A.INDEX_TIME) As INDEX_TIME, A.AMOUNT, A.DAY_NAME "
        l_strSQL = l_strSQL & ", NULLIF(A.OT_BANK_CHRG_DT,'0001-01-01') AS OT_BANK_CHRG_DT, CHAR(A.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", CASE WHEN A.JOB_NBR <> ' ' THEN A.JOB_NBR || A.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", A.JOBTASK_STEP_NBR, A.EMPRV_DATABASE_ID, A.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS A, METHDBA.TTR_OT_OPPORTUNITY B "
        l_strSQL = l_strSQL & ", METHDBA.TTR_OT_REASONS C, METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & ", METHDBA.TTR_CANV_SHEETS Z "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("A.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND   A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM > 1 "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM < 300000 "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM = B.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   B.OT_OPP_NUM = Z.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   Z.OT_REASON_CD = C.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   A.EMP_NUM = Z.EMP_NUM "
        l_strSQL = l_strSQL & "AND   A.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & "AND   A.START_TIME NOT LIKE '99%'"
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT D.EMP_NUM, D.CHARGE_DATE, D.DATE_WORKED, D.DISTRIBUTION_ID, D.HOURS, D.UNION_CD "
        l_strSQL = l_strSQL & ", D.HOURLY_RATE, D.START_LOC_CD, D.END_LOC_CD, D.INTERFACE_IND, D.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", D.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", D.START_TIME, D.END_TIME, D.EXCEPTION_CD, D.OVER_EXC_CD, D.EARNINGS_CD "
        l_strSQL = l_strSQL & ", D.WG_NUM, F.OT_REASON_DESC, E.OT_OPP_DESC, D.JOB_NUM , D.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", D.COMP_SHIFT_CD, D.COMP_SHIFT_CD_AMT, D.COMP_CD, D.COMP_AMT, D.REL_CD "
        l_strSQL = l_strSQL & ", D.CASE_NUM, D.FAMILY_LEAVE_IND, D.SCHED_CHG_EARN_CD, D.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", D.DEPT_NUM, D.LOC_CD, D.OT_OPP_NUM, F.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(D.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", D.APPR_BY_EMP_NUM, D.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(D.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", D.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(D.INDEX_TIME) As INDEX_TIME, D.AMOUNT, D.DAY_NAME "
        l_strSQL = l_strSQL & ", NULLIF(D.OT_BANK_CHRG_DT,'0001-01-01') AS OT_BANK_CHRG_DT, CHAR(D.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", CASE WHEN D.JOB_NBR <> ' ' THEN D.JOB_NBR || D.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", D.JOBTASK_STEP_NBR, D.EMPRV_DATABASE_ID, D.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS D, METHDBA.TTR_OT_NOCAN_OPP E, METHDBA.TTR_OT_REASONS F "
        l_strSQL = l_strSQL & ", METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("D.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND   D.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM > 31999999 "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM < 32200000 "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM = E.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   E.OT_REASON_CD = F.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   D.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & "AND   D.START_TIME NOT LIKE '99%'"
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT D.EMP_NUM, D.CHARGE_DATE, D.DATE_WORKED, D.DISTRIBUTION_ID, D.HOURS, D.UNION_CD "
        l_strSQL = l_strSQL & ", D.HOURLY_RATE, D.START_LOC_CD, D.END_LOC_CD, D.INTERFACE_IND, D.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", D.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", D.START_TIME, D.END_TIME, D.EXCEPTION_CD, D.OVER_EXC_CD, D.EARNINGS_CD "
        l_strSQL = l_strSQL & ", D.WG_NUM, F.OT_REASON_DESC, D.OT_OPP_DESC, D.JOB_NUM , D.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", D.COMP_SHIFT_CD, D.COMP_SHIFT_CD_AMT, D.COMP_CD, D.COMP_AMT, D.REL_CD "
        l_strSQL = l_strSQL & ", D.CASE_NUM, D.FAMILY_LEAVE_IND, D.SCHED_CHG_EARN_CD, D.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", D.DEPT_NUM, D.LOC_CD, D.OT_OPP_NUM, D.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(D.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", D.APPR_BY_EMP_NUM, D.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(D.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", D.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(D.INDEX_TIME) As INDEX_TIME, D.AMOUNT, D.DAY_NAME "
        l_strSQL = l_strSQL & ", NULLIF(D.OT_BANK_CHRG_DT,'0001-01-01') AS OT_BANK_CHRG_DT, CHAR(D.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", CASE WHEN D.JOB_NBR <> ' ' THEN D.JOB_NBR || D.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", D.JOBTASK_STEP_NBR, D.EMPRV_DATABASE_ID, D.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS D, METHDBA.TTR_OT_REASONS F "
        l_strSQL = l_strSQL & ", METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("D.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND   D.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM = 90000000 "
        l_strSQL = l_strSQL & "AND   D.OT_REASON_CD = F.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   D.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & "AND   D.START_TIME NOT LIKE '99%'"
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT G.EMP_NUM , G.CHARGE_DATE, G.DATE_WORKED, G.DISTRIBUTION_ID, G.HOURS, G.UNION_CD "
        l_strSQL = l_strSQL & ", G.HOURLY_RATE, G.START_LOC_CD, G.END_LOC_CD, G.INTERFACE_IND, G.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", G.TRADE_EXC_CD, H.EXCEPTION_BANK_CD, CAST(H.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", G.START_TIME, G.END_TIME, G.EXCEPTION_CD, G.OVER_EXC_CD, G.EARNINGS_CD "
        l_strSQL = l_strSQL & ", G.WG_NUM, '' AS OT_REASON_DESC, G.OT_OPP_DESC, G.JOB_NUM, G.DISTR_DEPT_NUM, G.COMP_SHIFT_CD, G.COMP_SHIFT_CD_AMT "
        l_strSQL = l_strSQL & ", G.COMP_CD, G.COMP_AMT, G.REL_CD, G.CASE_NUM, G.FAMILY_LEAVE_IND, G.SCHED_CHG_EARN_CD "
        l_strSQL = l_strSQL & ", G.SCHED_CHG_EXC_CD, G.DEPT_NUM, G.LOC_CD, G.OT_OPP_NUM, G.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(G.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", G.APPR_BY_EMP_NUM, G.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(G.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", G.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(G.INDEX_TIME) As INDEX_TIME, G.AMOUNT, G.DAY_NAME "
        ' The Line commented out below may replace the line immediately below it
        ' Jim Barnett needs to talk to the DBA
        ' OT_BANL_CHRG_DT defaults to the current date if not supplied
        ' in an insert statement
        'l_strSQL = l_strSQL & ", VALUE(NULLIF(G.OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", NULLIF(DATE('0100-01-01'),'0100-01-01') as OT_BANK_CHRG_DT, CHAR(G.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", CASE WHEN G.JOB_NBR <> ' ' THEN G.JOB_NBR || G.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", G.JOBTASK_STEP_NBR, G.EMPRV_DATABASE_ID, G.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACTUAL_TIME_HS G, METHDBA.TTR_EXCEPTION_CDS H "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("G.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & "AND   CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   OT_OPP_NUM = 0 "
        l_strSQL = l_strSQL & "AND   G.EXCEPTION_CD = H.EXCEPTION_CD "
        l_strSQL = l_strSQL & "AND   G.START_TIME NOT LIKE '99%'"
    End If
    If v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & " UNION ALL "
    End If
    If v_TrActualTimeTable = TrTable_Actual Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, C.OT_REASON_DESC, B.OT_OPP_DESC, A.JOB_NUM, A.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM, A.FAMILY_LEAVE_IND,  A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD, B.OT_OPP_NUM, A.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(A.INDEX_TIME) As INDEX_TIME, A.AMOUNT, A.DAY_NAME "
        l_strSQL = l_strSQL & ", NULLIF(A.OT_BANK_CHRG_DT,'0001-01-01') AS OT_BANK_CHRG_DT, CHAR(A.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", CASE WHEN A.JOB_NBR <> ' ' THEN A.JOB_NBR || A.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", A.JOBTASK_STEP_NBR, A.EMPRV_DATABASE_ID, A.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME A, METHDBA.TTR_OT_OPPORTUNITY B "
        l_strSQL = l_strSQL & ", METHDBA.TTR_OT_REASONS C, METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("A.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & "AND   A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND  ( (A.OT_OPP_NUM >= 300000 AND A.OT_OPP_NUM < 32000000) or "
        l_strSQL = l_strSQL & "       (A.OT_OPP_NUM >= 32200000 AND A.OT_OPP_NUM < 65000000) ) "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM = B.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   A.OT_REASON_CD = C.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   A.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & "AND   A.START_TIME NOT LIKE '99%'"
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, C.OT_REASON_DESC, B.OT_OPP_DESC, A.JOB_NUM, A.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM, A.FAMILY_LEAVE_IND,  A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD, B.OT_OPP_NUM, Z.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(A.INDEX_TIME) As INDEX_TIME, A.AMOUNT, A.DAY_NAME "
        l_strSQL = l_strSQL & ", NULLIF(A.OT_BANK_CHRG_DT,'0001-01-01') AS OT_BANK_CHRG_DT, CHAR(A.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", CASE WHEN A.JOB_NBR <> ' ' THEN A.JOB_NBR || A.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", A.JOBTASK_STEP_NBR, A.EMPRV_DATABASE_ID, A.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME A, METHDBA.TTR_OT_OPPORTUNITY B "
        l_strSQL = l_strSQL & ", METHDBA.TTR_OT_REASONS C, METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & ", METHDBA.TTR_CANV_SHEETS Z "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("A.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & "AND   A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM > 1 "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM < 300000 "
        l_strSQL = l_strSQL & "AND   A.OT_OPP_NUM = B.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   B.OT_OPP_NUM = Z.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   Z.OT_REASON_CD = C.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   A.EMP_NUM = Z.EMP_NUM "
        l_strSQL = l_strSQL & "AND   A.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & "AND   A.START_TIME NOT LIKE '99%'"
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT D.EMP_NUM, D.CHARGE_DATE, D.DATE_WORKED, D.DISTRIBUTION_ID, D.HOURS, D.UNION_CD "
        l_strSQL = l_strSQL & ", D.HOURLY_RATE, D.START_LOC_CD, D.END_LOC_CD, D.INTERFACE_IND, D.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", D.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", D.START_TIME, D.END_TIME, D.EXCEPTION_CD, D.OVER_EXC_CD, D.EARNINGS_CD "
        l_strSQL = l_strSQL & ", D.WG_NUM, F.OT_REASON_DESC, E.OT_OPP_DESC, D.JOB_NUM , D.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", D.COMP_SHIFT_CD, D.COMP_SHIFT_CD_AMT, D.COMP_CD, D.COMP_AMT, D.REL_CD "
        l_strSQL = l_strSQL & ", D.CASE_NUM, D.FAMILY_LEAVE_IND, D.SCHED_CHG_EARN_CD, D.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", D.DEPT_NUM, D.LOC_CD, D.OT_OPP_NUM, F.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(D.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", D.APPR_BY_EMP_NUM, D.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(D.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", D.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(D.INDEX_TIME) As INDEX_TIME, D.AMOUNT, D.DAY_NAME "
        l_strSQL = l_strSQL & ", NULLIF(D.OT_BANK_CHRG_DT,'0001-01-01') AS OT_BANK_CHRG_DT, CHAR(D.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", CASE WHEN D.JOB_NBR <> ' ' THEN D.JOB_NBR || D.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", D.JOBTASK_STEP_NBR, D.EMPRV_DATABASE_ID, D.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME D, METHDBA.TTR_OT_NOCAN_OPP E, METHDBA.TTR_OT_REASONS F "
        l_strSQL = l_strSQL & ", METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("D.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & "AND   D.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM > 31999999 "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM < 32200000 "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM = E.OT_OPP_NUM "
        l_strSQL = l_strSQL & "AND   E.OT_REASON_CD = F.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   D.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & "AND   D.START_TIME NOT LIKE '99%'"
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT D.EMP_NUM, D.CHARGE_DATE, D.DATE_WORKED, D.DISTRIBUTION_ID, D.HOURS, D.UNION_CD "
        l_strSQL = l_strSQL & ", D.HOURLY_RATE, D.START_LOC_CD, D.END_LOC_CD, D.INTERFACE_IND, D.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", D.TRADE_EXC_CD, I.EXCEPTION_BANK_CD, CAST(I.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", D.START_TIME, D.END_TIME, D.EXCEPTION_CD, D.OVER_EXC_CD, D.EARNINGS_CD "
        l_strSQL = l_strSQL & ", D.WG_NUM, F.OT_REASON_DESC, D.OT_OPP_DESC, D.JOB_NUM , D.DISTR_DEPT_NUM "
        l_strSQL = l_strSQL & ", D.COMP_SHIFT_CD, D.COMP_SHIFT_CD_AMT, D.COMP_CD, D.COMP_AMT, D.REL_CD "
        l_strSQL = l_strSQL & ", D.CASE_NUM, D.FAMILY_LEAVE_IND, D.SCHED_CHG_EARN_CD, D.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", D.DEPT_NUM, D.LOC_CD, D.OT_OPP_NUM, D.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(D.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", D.APPR_BY_EMP_NUM, D.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(D.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", D.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(D.INDEX_TIME) As INDEX_TIME, D.AMOUNT, D.DAY_NAME "
        l_strSQL = l_strSQL & ", NULLIF(D.OT_BANK_CHRG_DT,'0001-01-01') AS OT_BANK_CHRG_DT, CHAR(D.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", CASE WHEN D.JOB_NBR <> ' ' THEN D.JOB_NBR || D.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", D.JOBTASK_STEP_NBR, D.EMPRV_DATABASE_ID, D.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME D, METHDBA.TTR_OT_REASONS F "
        l_strSQL = l_strSQL & ", METHDBA.TTR_EXCEPTION_CDS I "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("D.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & "AND   D.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   D.OT_OPP_NUM = 90000000 "
        l_strSQL = l_strSQL & "AND   D.OT_REASON_CD = F.OT_REASON_CD "
        l_strSQL = l_strSQL & "AND   D.EXCEPTION_CD = I.EXCEPTION_CD "
        l_strSQL = l_strSQL & "AND   D.START_TIME NOT LIKE '99%'"
        l_strSQL = l_strSQL & " UNION ALL "
        l_strSQL = l_strSQL & "SELECT G.EMP_NUM , G.CHARGE_DATE, G.DATE_WORKED, G.DISTRIBUTION_ID, G.HOURS, G.UNION_CD "
        l_strSQL = l_strSQL & ", G.HOURLY_RATE, G.START_LOC_CD, G.END_LOC_CD, G.INTERFACE_IND, G.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", G.TRADE_EXC_CD, H.EXCEPTION_BANK_CD, CAST(H.EXCEPTION_DESC AS CHAR(" & v_intDescriptionLength & ")) AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", G.START_TIME, G.END_TIME, G.EXCEPTION_CD, G.OVER_EXC_CD, G.EARNINGS_CD "
        l_strSQL = l_strSQL & ", G.WG_NUM, '' AS OT_REASON_DESC, G.OT_OPP_DESC, G.JOB_NUM, G.DISTR_DEPT_NUM, G.COMP_SHIFT_CD, G.COMP_SHIFT_CD_AMT "
        l_strSQL = l_strSQL & ", G.COMP_CD, G.COMP_AMT, G.REL_CD, G.CASE_NUM, G.FAMILY_LEAVE_IND, G.SCHED_CHG_EARN_CD "
        l_strSQL = l_strSQL & ", G.SCHED_CHG_EXC_CD, G.DEPT_NUM, G.LOC_CD, G.OT_OPP_NUM, G.OT_REASON_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(G.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", G.APPR_BY_EMP_NUM, G.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(G.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", G.APPR_IND, '0' AS TIME_ENTRY_TYPE, '' AS RECORD_STATUS "
        l_strSQL = l_strSQL & ", CHAR(G.INDEX_TIME) As INDEX_TIME, G.AMOUNT, G.DAY_NAME "
        ' The Line commented out below may replace the line immediately below it
        ' Jim Barnett needs to talk to the DBA
        ' OT_BANL_CHRG_DT defaults to the current date if not supplied
        ' in an insert statement
        'l_strSQL = l_strSQL & ", VALUE(NULLIF(G.OT_BANK_CHRG_DT,'0001-01-01'),DATE('0100-01-01')) AS OT_BANK_CHRG_DT "
        l_strSQL = l_strSQL & ", NULLIF(DATE('0100-01-01'),'0100-01-01') as OT_BANK_CHRG_DT, CHAR(G.TIME_STAMP) AS TIME_STAMP "
        l_strSQL = l_strSQL & ", CASE WHEN G.JOB_NBR <> ' ' THEN G.JOB_NBR || G.JOBTASK_NBR ELSE ' ' END AS JOB_TASK "
        l_strSQL = l_strSQL & ", G.JOBTASK_STEP_NBR, G.EMPRV_DATABASE_ID, G.PREARRANGED_OT_IND "
        l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACTUAL_TIME G, METHDBA.TTR_EXCEPTION_CDS H "
        l_strSQL = l_strSQL & "WHERE " & BuildSelection("G.EMP_NUM", v_aryEmployees, False) & " "
        l_strSQL = l_strSQL & "AND   CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & "AND   OT_OPP_NUM = 0 "
        l_strSQL = l_strSQL & "AND   G.EXCEPTION_CD = H.EXCEPTION_CD "
        l_strSQL = l_strSQL & "AND   G.START_TIME NOT LIKE '99%'"
    End If
    l_strSQL = l_strSQL & "ORDER BY EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME "




    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetActualTimeEntriesWithoutOtherPayACTData = l_objrs
    
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetActualTimeEntriesWithoutOtherPayACTData", , , , DescArgsFrom(l_strSQL), True

End Function

'This function provides time for all employees passed including the wg number charged to
'this component is used exclusively for Approve Workgroup Time
'the business component busTimecard.cbsTimecard will process only those Work Group numbers selected using the filter method


Public Function GetApproveTimeSummaryByEmployee( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
        
    Dim l_lngStart As Long
    Dim l_datCutOffChargeDate As Date
    
On Error GoTo GetApproveTimeSummaryByEmployee_EH
    
    l_lngStart = LBound(v_aryEmployees)
    
    If Val(v_aryEmployees(l_lngStart)) < 1 Then
        Err.Raise trErrdbInvalidZerolngEmpNum
    Else
    End If
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then
        Set GetApproveTimeSummaryByEmployee = GetApproveTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_Actual)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then
           Set GetApproveTimeSummaryByEmployee = GetApproveTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_History)
        Else
           Set GetApproveTimeSummaryByEmployee = GetApproveTimeSummaryALLData(v_aryEmployees, v_datStartDate, v_datEndDate)
        End If
    End If
    
    GetObjectContext.SetComplete
    Exit Function
    
GetApproveTimeSummaryByEmployee_EH:
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetApproveTimeSummaryByEmployee", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetApproveTimeSummaryByEmployee", , , , , True
    End Select
    
End Function


Private Function GetApproveTimeSummaryACTData( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_TrActualTimeTable As TrActualTimeTable) As ADODB.Recordset
        
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetApproveTimeSummaryACTData_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = GetTimeSummarySql(v_aryEmployees, v_datStartDate, v_datEndDate, v_TrActualTimeTable, 1)

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetApproveTimeSummaryACTData = l_objrs
    
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function
    
GetApproveTimeSummaryACTData_EH:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetApproveTimeSummaryACTData", , , , DescArgsFrom(l_strSQL), True
    
End Function



Private Function GetApproveTimeSummaryALLData( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
        
    Dim l_decST_HRS_APPROVED As Variant
    Dim l_decST_HRS_UNAPPROVED As Variant
    Dim l_decOT_HRS_APPROVED As Variant
    Dim l_decOT_HRS_UNAPPROVED As Variant
    Dim l_decSYM_HRS_APPROVED As Variant
    Dim l_decSYM_HRS_UNAPPROVED As Variant
    Dim l_decTOT_HRS_UNAPPROVED As Variant
    Dim l_decTOT_HRS_APPROVED As Variant
    Dim l_decOTHERPAYIND_APP As Variant
    Dim l_decOTHERPAYIND_UNAPP As Variant

    Dim l_objrsACT As ADODB.Recordset
    Dim l_objrsHIS As ADODB.Recordset
    Dim l_strfilteredEmployee As String
    
On Error GoTo GetApproveTimeSummaryALLData_EH
    
    
    Set l_objrsACT = GetApproveTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_Actual)
    
    Set l_objrsHIS = GetApproveTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_History)
'need to summarize merge this two recordsets
    If RecordsExist(l_objrsACT) Then
        l_objrsACT.MoveFirst
        While Not l_objrsACT.EOF
            l_strfilteredEmployee = "EMP_NUM = " & Trim(str(l_objrsACT!EMP_NUM)) & " AND WG_NUM = " & Trim(str(l_objrsACT!WG_NUM)) & " AND EMP_NAME = " & "'" & Replace(Trim(l_objrsACT!EMP_NAME), "'", "''") & "'"
            l_strfilteredEmployee = MDAC25FilterFix(l_strfilteredEmployee)
            l_objrsHIS.Filter = l_strfilteredEmployee
            If l_objrsHIS.EOF And l_objrsHIS.BOF Then
            'bypass employee
            Else 'summarize employee data
                l_objrsACT!OTHERPAYIND_APP = CDec(l_objrsACT!OTHERPAYIND_APP) + CDec(l_objrsHIS!OTHERPAYIND_APP)
                l_objrsACT!OTHERPAYIND_UNAPP = CDec(l_objrsACT!OTHERPAYIND_UNAPP) + CDec(l_objrsHIS!OTHERPAYIND_UNAPP)
                l_objrsACT!ST_HRS_APPROVED = CDec(l_objrsACT!ST_HRS_APPROVED) + CDec(l_objrsHIS!ST_HRS_APPROVED)
                l_objrsACT!ST_HRS_UNAPPROVED = CDec(l_objrsACT!ST_HRS_UNAPPROVED) + CDec(l_objrsHIS!ST_HRS_UNAPPROVED)
                l_objrsACT!OT_HRS_APPROVED = CDec(l_objrsACT!OT_HRS_APPROVED) + CDec(l_objrsHIS!OT_HRS_APPROVED)
                l_objrsACT!OT_HRS_UNAPPROVED = CDec(l_objrsACT!OT_HRS_UNAPPROVED) + CDec(l_objrsHIS!OT_HRS_UNAPPROVED)
                l_objrsACT!SYM_HRS_APPROVED = CDec(l_objrsACT!SYM_HRS_APPROVED) + CDec(l_objrsHIS!SYM_HRS_APPROVED)
                l_objrsACT!SYM_HRS_UNAPPROVED = CDec(l_objrsACT!SYM_HRS_UNAPPROVED) + CDec(l_objrsHIS!SYM_HRS_UNAPPROVED)
                l_objrsACT!TOT_HRS_UNAPPROVED = CDec(l_objrsACT!TOT_HRS_UNAPPROVED) + CDec(l_objrsHIS!TOT_HRS_UNAPPROVED)
                l_objrsACT!TOT_HRS_APPROVED = CDec(l_objrsACT!TOT_HRS_APPROVED) + CDec(l_objrsHIS!TOT_HRS_APPROVED)
                l_objrsACT.Update
            End If
            l_objrsACT.MoveNext
        Wend
    Else
        Set l_objrsACT = l_objrsHIS
    End If
    
    If RecordsExist(l_objrsACT) Then
        l_objrsACT.MoveFirst
    End If
    
    Set GetApproveTimeSummaryALLData = l_objrsACT
    
    Set l_objrsACT = Nothing
    Set l_objrsHIS = Nothing
    Exit Function
    
GetApproveTimeSummaryALLData_EH:
    Set l_objrsACT = Nothing
    Set l_objrsHIS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetApproveTimeSummaryALLData", , , , , True
    
End Function



'This public method is exclusively used for View Time (by employee or by department or by WorkGroup
'time for the employee, including time charged to any other WorkGroups for this employee will be included in the result set

Public Function GetTimeSummaryByEmployee( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
        
    Dim l_lngStart As Long
    Dim l_datCutOffChargeDate As Date
    
On Error GoTo GetTimeSummaryByEmployee_EH
    
    l_lngStart = LBound(v_aryEmployees)
    
    If Val(v_aryEmployees(l_lngStart)) < 1 Then
        Err.Raise trErrdbInvalidZerolngEmpNum
    Else
    End If
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set GetTimeSummaryByEmployee = GetTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_Actual)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set GetTimeSummaryByEmployee = GetTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_History)
        Else
           Set GetTimeSummaryByEmployee = GetTimeSummaryALLData(v_aryEmployees, v_datStartDate, v_datEndDate) 'use both current and history data
        End If
    End If
    
    GetObjectContext.SetComplete
    Exit Function
    
GetTimeSummaryByEmployee_EH:
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetTimeSummaryByEmployee", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetTimeSummaryByEmployee", , , , , True
    End Select
    
End Function


Private Function GetTimeSummaryACTData( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_TrActualTimeTable As TrActualTimeTable) As ADODB.Recordset
        
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetTimeSummaryACTData_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = GetTimeSummarySql(v_aryEmployees, v_datStartDate, v_datEndDate, v_TrActualTimeTable)

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTimeSummaryACTData = l_objrs
    
    Set l_objrs = Nothing
    Exit Function
    
GetTimeSummaryACTData_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetTimeSummaryACTData", , , , DescArgsFrom(l_strSQL), True
    
End Function


Private Function GetTimeSummaryALLData( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
        
    Dim l_decST_HRS_APPROVED As Variant
    Dim l_decST_HRS_UNAPPROVED As Variant
    Dim l_decOT_HRS_APPROVED As Variant
    Dim l_decOT_HRS_UNAPPROVED As Variant
    Dim l_decSYM_HRS_APPROVED As Variant
    Dim l_decSYM_HRS_UNAPPROVED As Variant
    Dim l_decTOT_HRS_UNAPPROVED As Variant
    Dim l_decTOT_HRS_APPROVED As Variant
    Dim l_decOTHERPAYIND_APP As Variant
    Dim l_decOTHERPAYIND_UNAPP As Variant

    Dim l_objrsACT As ADODB.Recordset
    Dim l_objrsHIS As ADODB.Recordset
    Dim l_strfilteredEmployee As String
    
On Error GoTo GetTimeSummaryALLData_EH
    
    
    Set l_objrsACT = GetTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_Actual)
    
    Set l_objrsHIS = GetTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_History)
' to summarize merge this two recordsets
    If RecordsExist(l_objrsACT) Then
        l_objrsACT.MoveFirst
        While Not l_objrsACT.EOF
            l_strfilteredEmployee = "EMP_NUM = " & Trim(str(l_objrsACT!EMP_NUM))
            l_objrsHIS.Filter = l_strfilteredEmployee
            If l_objrsHIS.EOF And l_objrsHIS.BOF Then
            'bypass employee
            Else 'summarize employee data
                l_objrsACT!OTHERPAYIND_APP = CDec(l_objrsACT!OTHERPAYIND_APP) + CDec(l_objrsHIS!OTHERPAYIND_APP)
                l_objrsACT!OTHERPAYIND_UNAPP = CDec(l_objrsACT!OTHERPAYIND_UNAPP) + CDec(l_objrsHIS!OTHERPAYIND_UNAPP)
                l_objrsACT!ST_HRS_APPROVED = CDec(l_objrsACT!ST_HRS_APPROVED) + CDec(l_objrsHIS!ST_HRS_APPROVED)
                l_objrsACT!ST_HRS_UNAPPROVED = CDec(l_objrsACT!ST_HRS_UNAPPROVED) + CDec(l_objrsHIS!ST_HRS_UNAPPROVED)
                l_objrsACT!OT_HRS_APPROVED = CDec(l_objrsACT!OT_HRS_APPROVED) + CDec(l_objrsHIS!OT_HRS_APPROVED)
                l_objrsACT!OT_HRS_UNAPPROVED = CDec(l_objrsACT!OT_HRS_UNAPPROVED) + CDec(l_objrsHIS!OT_HRS_UNAPPROVED)
                l_objrsACT!SYM_HRS_APPROVED = CDec(l_objrsACT!SYM_HRS_APPROVED) + CDec(l_objrsHIS!SYM_HRS_APPROVED)
                l_objrsACT!SYM_HRS_UNAPPROVED = CDec(l_objrsACT!SYM_HRS_UNAPPROVED) + CDec(l_objrsHIS!SYM_HRS_UNAPPROVED)
                l_objrsACT!TOT_HRS_UNAPPROVED = CDec(l_objrsACT!TOT_HRS_UNAPPROVED) + CDec(l_objrsHIS!TOT_HRS_UNAPPROVED)
                l_objrsACT!TOT_HRS_APPROVED = CDec(l_objrsACT!TOT_HRS_APPROVED) + CDec(l_objrsHIS!TOT_HRS_APPROVED)
                l_objrsACT.Update
            End If
            l_objrsACT.MoveNext
        Wend
    Else
        Set l_objrsACT = l_objrsHIS
    End If
    
    If RecordsExist(l_objrsACT) Then
        l_objrsACT.MoveFirst
    End If
    
    Set GetTimeSummaryALLData = l_objrsACT
    
    Set l_objrsACT = Nothing
    Set l_objrsHIS = Nothing
    Exit Function
    
GetTimeSummaryALLData_EH:
    Set l_objrsACT = Nothing
    Set l_objrsHIS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetTimeSummaryALLData", , , , , True
    
End Function




'Get TTR_ACTUAL_TIME_HS cutoff Charge_Dt
Private Function GetActualTimeHSCutoffDate() As String
    
    Dim l_objCodeTables As dbCodeTables.cdCodeTables
    Dim l_varHSCutOffDate As Variant
    
    On Error GoTo GetActualTimeHSCutoffDate_EH
    
    Set l_objCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_varHSCutOffDate = l_objCodeTables.GetActualTimeHSCutoffDate(CKEYCUTOFFDATE)
    If IsDate(l_varHSCutOffDate) Then
        GetActualTimeHSCutoffDate = CStr(l_varHSCutOffDate)
    Else
        Err.Raise trErrdb_GetHSCutOffDateInfo
    End If
 
    Set l_objCodeTables = Nothing

    GetObjectContext.SetComplete
    Exit Function
    
GetActualTimeHSCutoffDate_EH:
    GetObjectContext.SetAbort
    
    Set l_objCodeTables = Nothing
    
    Select Case Err.Number
        Case trErrdb_GetHSCutOffDateInfo
            TrsRaiseError m_strCMODULENAME, "GetActualTimeHSCutoffDate", trErrdb_GetHSCutOffDateInfo, , , DescArgsFrom(CStr(l_varHSCutOffDate)), True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetActualTimeHSCutoffDate", , , , , True
    End Select

End Function

Public Function GetDeptTimeSummaryByEmployee( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
        
    Dim l_lngStart As Long
    Dim l_datCutOffChargeDate As Date
    
On Error GoTo GetDeptTimeSummaryByEmployee_EH
    
    l_lngStart = LBound(v_aryEmployees)
    
    If Val(v_aryEmployees(l_lngStart)) < 1 Then
        Err.Raise trErrdbInvalidZerolngEmpNum
    Else
    End If
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then
        Set GetDeptTimeSummaryByEmployee = GetDeptTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_Actual)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then
           Set GetDeptTimeSummaryByEmployee = GetDeptTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_History)
        Else
           Set GetDeptTimeSummaryByEmployee = GetDeptTimeSummaryALLData(v_aryEmployees, v_datStartDate, v_datEndDate)
        End If
    End If
    
    GetObjectContext.SetComplete
    Exit Function
    
GetDeptTimeSummaryByEmployee_EH:
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetDeptTimeSummaryByEmployee", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetDeptTimeSummaryByEmployee", , , , , True
    End Select
    
End Function


Private Function GetDeptTimeSummaryACTData( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_TrActualTimeTable As TrActualTimeTable) As ADODB.Recordset
        
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetDeptTimeSummaryACTData_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = GetTimeSummarySql(v_aryEmployees, v_datStartDate, v_datEndDate, v_TrActualTimeTable, 2)

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDeptTimeSummaryACTData = l_objrs
    
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function
    
GetDeptTimeSummaryACTData_EH:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDeptTimeSummaryACTData", , , , DescArgsFrom(l_strSQL), True
    
End Function



Private Function GetDeptTimeSummaryALLData( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
        
    Dim l_decST_HRS_APPROVED As Variant
    Dim l_decST_HRS_UNAPPROVED As Variant
    Dim l_decOT_HRS_APPROVED As Variant
    Dim l_decOT_HRS_UNAPPROVED As Variant
    Dim l_decSYM_HRS_APPROVED As Variant
    Dim l_decSYM_HRS_UNAPPROVED As Variant
    Dim l_decTOT_HRS_UNAPPROVED As Variant
    Dim l_decTOT_HRS_APPROVED As Variant
    Dim l_decOTHERPAYIND_APP As Variant
    Dim l_decOTHERPAYIND_UNAPP As Variant

    Dim l_objrsACT As ADODB.Recordset
    Dim l_objrsHIS As ADODB.Recordset
    Dim l_strfilteredEmployee As String
    
On Error GoTo GetDeptTimeSummaryALLData_EH
    
    
    Set l_objrsACT = GetDeptTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_Actual)
    
    Set l_objrsHIS = GetDeptTimeSummaryACTData(v_aryEmployees, v_datStartDate, v_datEndDate, TrTable_History)
'need to summarize merge this two recordsets
    If RecordsExist(l_objrsACT) Then
        l_objrsACT.MoveFirst
        While Not l_objrsACT.EOF
            l_strfilteredEmployee = "EMP_NUM = " & Trim(str(l_objrsACT!EMP_NUM)) & " AND DEPT_NUM = " & "'" & Trim(l_objrsACT!DEPT_NUM) & "'" & " AND EMP_NAME = " & "'" & Replace(Trim(l_objrsACT!EMP_NAME), "'", "''") & "'"
            l_strfilteredEmployee = MDAC25FilterFix(l_strfilteredEmployee)
            l_objrsHIS.Filter = l_strfilteredEmployee
            If l_objrsHIS.EOF And l_objrsHIS.BOF Then
            'bypass employee
            Else 'summarize employee data
                l_objrsACT!OTHERPAYIND_APP = CDec(l_objrsACT!OTHERPAYIND_APP) + CDec(l_objrsHIS!OTHERPAYIND_APP)
                l_objrsACT!OTHERPAYIND_UNAPP = CDec(l_objrsACT!OTHERPAYIND_UNAPP) + CDec(l_objrsHIS!OTHERPAYIND_UNAPP)
                l_objrsACT!ST_HRS_APPROVED = CDec(l_objrsACT!ST_HRS_APPROVED) + CDec(l_objrsHIS!ST_HRS_APPROVED)
                l_objrsACT!ST_HRS_UNAPPROVED = CDec(l_objrsACT!ST_HRS_UNAPPROVED) + CDec(l_objrsHIS!ST_HRS_UNAPPROVED)
                l_objrsACT!OT_HRS_APPROVED = CDec(l_objrsACT!OT_HRS_APPROVED) + CDec(l_objrsHIS!OT_HRS_APPROVED)
                l_objrsACT!OT_HRS_UNAPPROVED = CDec(l_objrsACT!OT_HRS_UNAPPROVED) + CDec(l_objrsHIS!OT_HRS_UNAPPROVED)
                l_objrsACT!SYM_HRS_APPROVED = CDec(l_objrsACT!SYM_HRS_APPROVED) + CDec(l_objrsHIS!SYM_HRS_APPROVED)
                l_objrsACT!SYM_HRS_UNAPPROVED = CDec(l_objrsACT!SYM_HRS_UNAPPROVED) + CDec(l_objrsHIS!SYM_HRS_UNAPPROVED)
                l_objrsACT!TOT_HRS_UNAPPROVED = CDec(l_objrsACT!TOT_HRS_UNAPPROVED) + CDec(l_objrsHIS!TOT_HRS_UNAPPROVED)
                l_objrsACT!TOT_HRS_APPROVED = CDec(l_objrsACT!TOT_HRS_APPROVED) + CDec(l_objrsHIS!TOT_HRS_APPROVED)
                l_objrsACT.Update
            End If
            l_objrsACT.MoveNext
        Wend
    Else
        Set l_objrsACT = l_objrsHIS
    End If
    
    If RecordsExist(l_objrsACT) Then
        l_objrsACT.MoveFirst
    End If
    
    Set GetDeptTimeSummaryALLData = l_objrsACT
    
    Set l_objrsACT = Nothing
    Set l_objrsHIS = Nothing
    Exit Function
    
GetDeptTimeSummaryALLData_EH:
    Set l_objrsACT = Nothing
    Set l_objrsHIS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDeptTimeSummaryALLData", , , , , True
    
End Function




Public Function GetActualTimeCascadeExists( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
' v_aryEmployees should be an array of longs.
' return time entries only, no overtime or other pay
' RED278
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClausePartValueKey As String
    Dim l_lngActualTimeTables As Long
    Dim l_datHistoryEndDate As Date
    Dim l_datActualTimeStartDate As Date

On Error GoTo eh
    
    l_lngActualTimeTables = DetermineActualTimeTables(v_datStartDate, v_datEndDate, l_datHistoryEndDate, l_datActualTimeStartDate)
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, l_datHistoryEndDate), True)
    If l_lngActualTimeTables = TrTable_History Or _
       l_lngActualTimeTables = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, A.START_TIME, A.END_TIME, B.EMP_NAME"
        l_strSQL = l_strSQL & " FROM  METHDBA.TTR_ACTUAL_TIME_HS A, METHDBA.TTR_EMPLOYEE B"
        l_strSQL = l_strSQL & " WHERE " & BuildInClause("A.EMP_NUM", v_aryEmployees, False)
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & " AND   A.CHARGE_DATE BETWEEN " & DateAsEnclosedString(v_datStartDate) & " AND " & DateAsEnclosedString(l_datHistoryEndDate)
        l_strSQL = l_strSQL & " AND   A.OT_OPP_NUM = 0"         'no overtime
        l_strSQL = l_strSQL & " AND   A.START_TIME < '24:00'"   'no other pay
        l_strSQL = l_strSQL & " AND   A.EMP_NUM = B.EMP_NUM"
    End If
    If l_lngActualTimeTables = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & " UNION ALL "
    End If
    If l_lngActualTimeTables = TrTable_Actual Or _
       l_lngActualTimeTables = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, A.START_TIME, A.END_TIME, B.EMP_NAME"
        l_strSQL = l_strSQL & " FROM  METHDBA.TTR_ACTUAL_TIME A, METHDBA.TTR_EMPLOYEE B"
        l_strSQL = l_strSQL & " WHERE " & BuildInClause("A.EMP_NUM", v_aryEmployees, False)
        l_strSQL = l_strSQL & " AND   A.CHARGE_DATE BETWEEN " & DateAsEnclosedString(l_datActualTimeStartDate) & " AND " & DateAsEnclosedString(v_datEndDate)
        l_strSQL = l_strSQL & " AND   A.OT_OPP_NUM = 0"
        l_strSQL = l_strSQL & " AND   A.START_TIME < '24:00'"
        l_strSQL = l_strSQL & " AND   A.EMP_NUM = B.EMP_NUM"
    End If
    l_strSQL = l_strSQL & " ORDER BY EMP_NUM, CHARGE_DATE"

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing

    Set GetActualTimeCascadeExists = l_objrs
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetActualTimeCascadeExists", , , , DescArgsFrom(l_strSQL), True
End Function

Private Function DetermineActualTimeTables( _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date, _
                    ByRef r_datHistoryEndDate As Date, _
                    ByRef r_datActualTimeStartDate As Date _
                    ) As TrActualTimeTable
                                            
    Dim l_datCutOffChargeDate As Date

On Error GoTo eh
    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then
        DetermineActualTimeTables = TrTable_Actual
        r_datActualTimeStartDate = v_datStartDate
    ElseIf v_datEndDate <= l_datCutOffChargeDate Then
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


Public Function GetActualTimeDatesWorked( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
' v_aryEmployees should be an array of longs.
' or a long for one employee
'

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClausePartValueKey As String
    Dim l_lngActualTimeTables As Long
    Dim l_strEmployeeSelection As String
    Dim l_datHistoryEndDate As Date
    Dim l_datActualTimeStartDate As Date

On Error GoTo eh

    l_lngActualTimeTables = DetermineActualTimeTables(v_datStartDate, v_datEndDate, l_datHistoryEndDate, l_datActualTimeStartDate)
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False)
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(v_aryEmployees)
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing

    l_objCN.Open


'May need to alter SQL to SUM by Shift using SCHED_CD
'Currently SCHED_CD is unavailable

    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, l_datHistoryEndDate), True)
    If l_lngActualTimeTables = TrTable_History Or _
       l_lngActualTimeTables = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, SUM(HOURS) AS TOT_HRS"
        l_strSQL = l_strSQL & " ,SUBSTR('A',1,1) as TIME_TYPE "
        l_strSQL = l_strSQL & " FROM  METHDBA.TTR_ACTUAL_TIME_HS A"
        l_strSQL = l_strSQL & " WHERE " & l_strEmployeeSelection
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & " AND   A.CHARGE_DATE BETWEEN " & DateAsEnclosedString(v_datStartDate) & " AND " & DateAsEnclosedString(l_datHistoryEndDate)
        l_strSQL = l_strSQL & " AND   A.START_TIME < '24:00'"   'no other pay
        l_strSQL = l_strSQL & " GROUP BY A.EMP_NUM , A.CHARGE_DATE "
    End If
    If l_lngActualTimeTables = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & " UNION ALL "
    End If
    If l_lngActualTimeTables = TrTable_Actual Or _
       l_lngActualTimeTables = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, SUM(HOURS) AS TOT_HRS"
        l_strSQL = l_strSQL & " ,SUBSTR('A',1,1) as TIME_TYPE "
        l_strSQL = l_strSQL & " FROM  METHDBA.TTR_ACTUAL_TIME A"
        l_strSQL = l_strSQL & " WHERE " & l_strEmployeeSelection
        l_strSQL = l_strSQL & " AND   A.CHARGE_DATE BETWEEN " & DateAsEnclosedString(l_datActualTimeStartDate) & " AND " & DateAsEnclosedString(v_datEndDate)
        l_strSQL = l_strSQL & " AND   A.START_TIME < '24:00'"
        l_strSQL = l_strSQL & " GROUP BY A.EMP_NUM , A.CHARGE_DATE "
    End If
    l_strSQL = l_strSQL & " ORDER BY 1,2"

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing

    Set GetActualTimeDatesWorked = l_objrs
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetActualTimeDatesWorked", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetExceptionSchedTimeEntries( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
' v_aryEmployees should be an array of longs.
' return time entries only, no overtime or other pay
' RED278
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClausePartValueKey As String
    Dim l_lngActualTimeTables As Long
    Dim l_datHistoryEndDate As Date
    Dim l_datActualTimeStartDate As Date
    Dim l_strEmployeeSelection As String

On Error GoTo eh
    
    l_lngActualTimeTables = DetermineActualTimeTables(v_datStartDate, v_datEndDate, l_datHistoryEndDate, l_datActualTimeStartDate)
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False)
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(v_aryEmployees)
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, l_datHistoryEndDate), True)
    If l_lngActualTimeTables = TrTable_History Or _
       l_lngActualTimeTables = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, A.HOURS, A.START_TIME, A.END_TIME"
        l_strSQL = l_strSQL & ",A.DATE_WORKED, A.EXCEPTION_CD, A.FAMILY_LEAVE_IND"
        l_strSQL = l_strSQL & ",A.CASE_NUM, WG_NUM, APPR_IND, APPR_BY_EMP_NUM, PROCESSED_IND"
        l_strSQL = l_strSQL & ",OVER_EXC_CD, FAMILY_LEAVE_IND, CHAR(A.INDEX_TIME) AS INDEX_TIME"
        l_strSQL = l_strSQL & " FROM  METHDBA.TTR_ACTUAL_TIME_HS A"
        l_strSQL = l_strSQL & " WHERE " & l_strEmployeeSelection
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & " AND   A.CHARGE_DATE BETWEEN " & DateAsEnclosedString(v_datStartDate) & " AND " & DateAsEnclosedString(l_datHistoryEndDate)
        l_strSQL = l_strSQL & " AND   A.OT_OPP_NUM = 0"         'no overtime
        l_strSQL = l_strSQL & " AND   A.START_TIME < '24:00'"   'no other pay
    End If
    If l_lngActualTimeTables = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & " UNION ALL "
    End If
    If l_lngActualTimeTables = TrTable_Actual Or _
       l_lngActualTimeTables = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT A.EMP_NUM, A.CHARGE_DATE, A.HOURS, A.START_TIME, A.END_TIME"
        l_strSQL = l_strSQL & ",A.DATE_WORKED, A.EXCEPTION_CD, A.FAMILY_LEAVE_IND"
        l_strSQL = l_strSQL & ",A.CASE_NUM,WG_NUM, APPR_IND, APPR_BY_EMP_NUM, PROCESSED_IND"
        l_strSQL = l_strSQL & ",OVER_EXC_CD, FAMILY_LEAVE_IND, CHAR(A.INDEX_TIME) AS INDEX_TIME"
        l_strSQL = l_strSQL & " FROM  METHDBA.TTR_ACTUAL_TIME A"
        l_strSQL = l_strSQL & " WHERE " & l_strEmployeeSelection
        l_strSQL = l_strSQL & " AND   A.CHARGE_DATE BETWEEN " & DateAsEnclosedString(l_datActualTimeStartDate) & " AND " & DateAsEnclosedString(v_datEndDate)
        l_strSQL = l_strSQL & " AND   A.OT_OPP_NUM = 0"
        l_strSQL = l_strSQL & " AND   A.START_TIME < '24:00'"
    End If
    l_strSQL = l_strSQL & " ORDER BY EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME"

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing

    Set GetExceptionSchedTimeEntries = l_objrs
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetExceptionSchedTimeEntries", , , , DescArgsFrom(l_strSQL), True
End Function
Public Function GetDistributionActualTimeEntries( _
        ByVal v_adrsDistributions As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_bActiveEmployees As Boolean) As ADODB.Recordset

    Dim l_datCutOffChargeDate As Date

    On Error GoTo eh
    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set GetDistributionActualTimeEntries = GetDistributionActualTimeEntriesACTData(v_adrsDistributions, v_datStartDate, v_datEndDate, v_bActiveEmployees, TrTable_Actual)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set GetDistributionActualTimeEntries = GetDistributionActualTimeEntriesACTData(v_adrsDistributions, v_datStartDate, v_datEndDate, v_bActiveEmployees, TrTable_History)
        Else
           Set GetDistributionActualTimeEntries = GetDistributionActualTimeEntriesACTData(v_adrsDistributions, v_datStartDate, v_datEndDate, v_bActiveEmployees, TrTable_ActualAndHistory) 'use both current and history data
        End If
    End If

    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetDistributionActualTimeEntries", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetDistributionActualTimeEntries", , , , , True
    End Select

End Function
Public Function GetDistributionActualTimeEntries2( _
        ByVal v_adrsDistributions As ADODB.Recordset, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_aryEmployees As Variant) As ADODB.Recordset

    Dim l_datCutOffChargeDate As Date

    On Error GoTo eh
    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set GetDistributionActualTimeEntries2 = GetDistributionActualTimeEntriesACTData2(v_adrsDistributions, v_datStartDate, v_datEndDate, v_aryEmployees, TrTable_Actual)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set GetDistributionActualTimeEntries2 = GetDistributionActualTimeEntriesACTData2(v_adrsDistributions, v_datStartDate, v_datEndDate, v_aryEmployees, TrTable_History)
        Else
           Set GetDistributionActualTimeEntries2 = GetDistributionActualTimeEntriesACTData2(v_adrsDistributions, v_datStartDate, v_datEndDate, v_aryEmployees, TrTable_ActualAndHistory) 'use both current and history data
        End If
    End If

    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetDistributionActualTimeEntries2", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetDistributionActualTimeEntries2", , , , , True
    End Select

End Function
Private Function GetDistributionActualTimeEntriesACTData( _
        ByVal v_adrsDistributions As ADODB.Recordset, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_bActiveEmployees As Boolean, _
        ByVal v_TrActualTimeTable As TrActualTimeTable) As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClausePartValueKey As String
    Dim l_sDistributionInClause As String
    Dim l_sDistributionInClauseA As String
    
    On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_sDistributionInClauseA = BuildDistributionSelection("A", v_adrsDistributions)
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
    If v_TrActualTimeTable = TrTable_History Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = "SELECT       A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, B.EXCEPTION_BANK_CD, B.EXCEPTION_DESC AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM,  A.JOB_NUM, A.DISTR_DEPT_NUM, A.OT_OPP_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM, A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE  "
        l_strSQL = l_strSQL & ", A.AMOUNT "
        l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS TIME_STAMP, A.SCHED_CD, A.SCHED_SRC_CD "
'BEGIN PREM TRIS CONVERSION PHASE 3 (NEEDED FOR ACTUAL TIME REPORT BY DISTRIBUTION)
        l_strSQL = l_strSQL & ", A.OT_REASON_CD "
'END PREM TRIS CONVERSION PHASE 3 (NEEDED FOR ACTUAL TIME REPORT BY DISTRIBUTION)
        If v_bActiveEmployees Then
            l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS A,METHDBA.TTR_EXCEPTION_CDS B "
            l_strSQL = l_strSQL & " ,METHDBA.TTR_EMPLOYEE C "
        Else
            l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS A,METHDBA.TTR_EXCEPTION_CDS B "
        End If
        l_strSQL = l_strSQL & "WHERE (" & l_sDistributionInClauseA & ")"
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & " AND  A.EMP_NUM = C.EMP_NUM "
        If v_bActiveEmployees Then
            l_strSQL = l_strSQL & " AND  A.EMP_NUM = C.EMP_NUM "
            l_strSQL = l_strSQL & "AND C.EMP_STATUS = 'A' "
        End If
        l_strSQL = l_strSQL & " AND A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " AND A.EXCEPTION_CD = B.EXCEPTION_CD "
    End If
    If v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & " UNION ALL "
    End If
    If v_TrActualTimeTable = TrTable_Actual Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT       A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, B.EXCEPTION_BANK_CD, B.EXCEPTION_DESC AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, A.JOB_NUM, A.DISTR_DEPT_NUM, A.OT_OPP_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A. COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM,  A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE "
        l_strSQL = l_strSQL & ", A.AMOUNT "
        l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS TIME_STAMP, A.SCHED_CD, A.SCHED_SRC_CD "
'BEGIN PREM TRIS CONVERSION PHASE 3 (NEEDED FOR ACTUAL TIME REPORT BY DISTRIBUTION)
        l_strSQL = l_strSQL & ", A.OT_REASON_CD "
'END PREM TRIS CONVERSION PHASE 3 (NEEDED FOR ACTUAL TIME REPORT BY DISTRIBUTION)
        If v_bActiveEmployees Then
            l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME A,METHDBA.TTR_EXCEPTION_CDS B "
            l_strSQL = l_strSQL & " ,METHDBA.TTR_EMPLOYEE C "
        Else
            l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME A,METHDBA.TTR_EXCEPTION_CDS B "
        End If
        l_strSQL = l_strSQL & "WHERE (" & l_sDistributionInClauseA & ")"
        l_strSQL = l_strSQL & " AND A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " AND A.EXCEPTION_CD = B.EXCEPTION_CD "
        If v_bActiveEmployees Then
            l_strSQL = l_strSQL & " AND  A.EMP_NUM = C.EMP_NUM "
            l_strSQL = l_strSQL & " AND C.EMP_STATUS = 'A' "
        End If

    End If
    l_strSQL = l_strSQL & "ORDER BY DISTRIBUTION_ID, DISTRIBUTION_TYPE, EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME "
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDistributionActualTimeEntriesACTData = l_objrs
    
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDistributionActualTimeEntriesACTData", , , , DescArgsFrom(l_strSQL), True

End Function
Private Function GetDistributionActualTimeEntriesACTData2( _
        ByVal v_adrsDistributions As ADODB.Recordset, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_aryEmployees As Variant, _
        ByVal v_TrActualTimeTable As TrActualTimeTable) As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClausePartValueKey As String
    Dim l_sDistributionInClause As String
    Dim l_sDistributionInClauseA As String
    
    On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_sDistributionInClauseA = BuildDistributionSelection("A", v_adrsDistributions)
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
    If v_TrActualTimeTable = TrTable_History Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = "SELECT       A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, B.EXCEPTION_BANK_CD, B.EXCEPTION_DESC AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM,  A.JOB_NUM, A.DISTR_DEPT_NUM, A.OT_OPP_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A.COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM, A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE  "
        l_strSQL = l_strSQL & ", A.AMOUNT "
        l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS TIME_STAMP, A.SCHED_CD, A.SCHED_SRC_CD "
'BEGIN PREM TRIS CONVERSION PHASE 3 (NEEDED FOR ACTUAL TIME REPORT BY DISTRIBUTION)
        l_strSQL = l_strSQL & ", A.OT_REASON_CD "
'END PREM TRIS CONVERSION PHASE 3 (NEEDED FOR ACTUAL TIME REPORT BY DISTRIBUTION)
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME_HS A,METHDBA.TTR_EXCEPTION_CDS B "
        l_strSQL = l_strSQL & "WHERE (" & l_sDistributionInClauseA & ")"
        l_strSQL = l_strSQL & " AND " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & " AND " & BuildSelection("A.EMP_NUM", v_aryEmployees, False)
        l_strSQL = l_strSQL & " AND A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " AND A.EXCEPTION_CD = B.EXCEPTION_CD "
    End If
    If v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & " UNION ALL "
    End If
    If v_TrActualTimeTable = TrTable_Actual Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "SELECT       A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.DISTRIBUTION_ID, A.HOURS, A.UNION_CD "
        l_strSQL = l_strSQL & ", A.HOURLY_RATE, A.START_LOC_CD, A.END_LOC_CD, A.INTERFACE_IND, A.DISTRIBUTION_TYPE "
        l_strSQL = l_strSQL & ", A.TRADE_EXC_CD, B.EXCEPTION_BANK_CD, B.EXCEPTION_DESC AS DISTRIBUTION_DESC "
        l_strSQL = l_strSQL & ", A.START_TIME, A.END_TIME, A.EXCEPTION_CD, A.OVER_EXC_CD, A.EARNINGS_CD "
        l_strSQL = l_strSQL & ", A.WG_NUM, A.JOB_NUM, A.DISTR_DEPT_NUM, A.OT_OPP_NUM "
        l_strSQL = l_strSQL & ", A.COMP_SHIFT_CD, A.COMP_SHIFT_CD_AMT, A.COMP_CD, A.COMP_AMT, A.REL_CD "
        l_strSQL = l_strSQL & ", A.CASE_NUM,  A.SCHED_CHG_EARN_CD, A.SCHED_CHG_EXC_CD "
        l_strSQL = l_strSQL & ", A.DEPT_NUM, A.LOC_CD "
        l_strSQL = l_strSQL & ", Value(NullIF(Value(LEFT(CHAR(NullIf(A.TRAVEL_PAY_HRS,0)),1),'N'),''),'Y') as TRAVEL_PAY_IND "
        l_strSQL = l_strSQL & ", A.APPR_BY_EMP_NUM, A.LAST_MOD_EMP_NUM "
        l_strSQL = l_strSQL & ", VALUE(NULLIF(A.APPR_DATE,'0001-01-01'),DATE('0100-01-01')) AS APPR_DATE "
        l_strSQL = l_strSQL & ", A.APPR_IND, '0' AS TIME_ENTRY_TYPE "
        l_strSQL = l_strSQL & ", A.AMOUNT "
        l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS TIME_STAMP, A.SCHED_CD, A.SCHED_SRC_CD "
'BEGIN PREM TRIS CONVERSION PHASE 3 (NEEDED FOR ACTUAL TIME REPORT BY DISTRIBUTION)
        l_strSQL = l_strSQL & ", A.OT_REASON_CD "
'END PREM TRIS CONVERSION PHASE 3 (NEEDED FOR ACTUAL TIME REPORT BY DISTRIBUTION)
        l_strSQL = l_strSQL & "FROM  METHDBA.TTR_ACTUAL_TIME A,METHDBA.TTR_EXCEPTION_CDS B "
        l_strSQL = l_strSQL & "WHERE (" & l_sDistributionInClauseA & ")"
        l_strSQL = l_strSQL & " AND A.CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "' AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " AND A.EXCEPTION_CD = B.EXCEPTION_CD "
        l_strSQL = l_strSQL & " AND " & BuildSelection("A.EMP_NUM", v_aryEmployees, False)

    End If
    l_strSQL = l_strSQL & "ORDER BY DISTRIBUTION_ID, DISTRIBUTION_TYPE, EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME "
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDistributionActualTimeEntriesACTData2 = l_objrs
    
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDistributionActualTimeEntriesACTData2", , , , DescArgsFrom(l_strSQL), True

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
            'l_strDistribution = Trim(UCase(!DISTRIBUTION_ID))
            l_strDistribution = UCase(!DISTRIBUTION_ID)
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


Public Function GetOTShiftCode(ByVal v_lngOTOppNum As Long) As ADODB.Recordset
On Error GoTo GetOTShiftCode_EH
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = " SELECT SHIFT_CD FROM METHDBA.TTR_OT_OPPORTUNITY "
    l_strSQL = l_strSQL & " WHERE OT_OPP_NUM = " & v_lngOTOppNum

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetOTShiftCode = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete
    Exit Function
GetOTShiftCode_EH:
    GetObjectContext.SetAbort

    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetOTShiftCode", , , , DescArgsFrom(l_strSQL), True
    
End Function

Private Function GetTimeSummarySql( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_TrActualTimeTable As TrActualTimeTable, _
        Optional v_intSumBy As Integer = 0) As String

    Dim l_strSQL As String
    Dim l_strGroupBy As String
    Dim l_strEmpSel As String
    Dim l_strSQL1stLine As String
    Dim l_strActTimeTable As String
    Dim l_strInClausePartValueKey As String
    

    On Error GoTo eh
    
'Private Enum TrActualTimeTable
'    TrTable_Actual = 1
'    TrTable_History = 2
'End Enum

'v_intSumBy has 3 values
'  0 - None
'  1 - Work group number
'  2 - Deptartment number

    Select Case v_intSumBy
        Case 0
            l_strSQL1stLine = "SELECT EMP.EMP_NUM, EMP_NAME, "
            l_strGroupBy = "EMP_NUM "
        Case 1
            l_strSQL1stLine = "SELECT EMP.EMP_NUM, EMP_NAME,WG_NUM, "
            l_strGroupBy = "EMP_NUM,WG_NUM "
        Case 2
            l_strSQL1stLine = "SELECT EMP.EMP_NUM, EMP_NAME, SUMMARY1.DEPT_NUM, "
            l_strGroupBy = "EMP_NUM,DEPT_NUM "
        Case Else
            GetTimeSummarySql = ""
            Exit Function
    End Select
    
    Select Case v_TrActualTimeTable
        Case TrTable_Actual
            l_strActTimeTable = " methdba.ttr_actual_time a "
            l_strInClausePartValueKey = ""
        Case TrTable_History
            l_strActTimeTable = " methdba.ttr_actual_time_hs a "
            l_strInClausePartValueKey = " AND " & _
                BuildInClause("part_value_key", _
                    BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
    End Select
        
    
    l_strEmpSel = BuildSelection("emp_num", v_aryEmployees, False) & " "
    
    l_strSQL = l_strSQL1stLine
    l_strSQL = l_strSQL & "       ST_HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "       OT_HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "       SYM_HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "       OTHERPAYIND_UNAPP,"
    l_strSQL = l_strSQL & "       TOT_HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "       ST_HRS_APPROVED,"
    l_strSQL = l_strSQL & "       OT_HRS_APPROVED,"
    l_strSQL = l_strSQL & "       SYM_HRS_APPROVED,"
    l_strSQL = l_strSQL & "       OTHERPAYIND_APP,"
    l_strSQL = l_strSQL & "       TOT_HRS_APPROVED"
    l_strSQL = l_strSQL & " FROM"
    l_strSQL = l_strSQL & " (SELECT " & l_strGroupBy & " , "
    l_strSQL = l_strSQL & "       SUM(ST_HRS_UNAPPROVED) AS ST_HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "       SUM(OT_HRS_UNAPPROVED) AS OT_HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "       SUM(SYM_HRS_UNAPPROVED) AS SYM_HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "       SUM(OTHERPAYIND_UNAPP) AS OTHERPAYIND_UNAPP,"
    l_strSQL = l_strSQL & "       SUM(HRS_UNAPPROVED) AS TOT_HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "       SUM(ST_HRS_APPROVED) AS ST_HRS_APPROVED,"
    l_strSQL = l_strSQL & "       SUM(OT_HRS_APPROVED) AS OT_HRS_APPROVED,"
    l_strSQL = l_strSQL & "       SUM(SYM_HRS_APPROVED) AS SYM_HRS_APPROVED,"
    l_strSQL = l_strSQL & "       SUM(OTHERPAYIND_APP) AS OTHERPAYIND_APP,"
    l_strSQL = l_strSQL & "       SUM(HRS_APPROVED) AS TOT_HRS_APPROVED"
    l_strSQL = l_strSQL & " FROM"
    l_strSQL = l_strSQL & " (select " & l_strGroupBy & " , "
    l_strSQL = l_strSQL & "     case when appr_ind = 'Y' AND start_time not like '99%'"
    l_strSQL = l_strSQL & "          THEN (hours + travel_pay_hrs) ELSE 0 END AS HRS_APPROVED,"
    l_strSQL = l_strSQL & "     case when appr_ind = 'N' AND start_time not like '99%'"
    l_strSQL = l_strSQL & "           THEN (hours + travel_pay_hrs) ELSE 0 END AS HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "     case when ot_opp_num > 0 and appr_ind = 'Y'"
    l_strSQL = l_strSQL & "             then (hours + travel_pay_hrs) ELSE 0 end as OT_HRS_APPROVED,"
    l_strSQL = l_strSQL & "     case when ot_opp_num > 0 and appr_ind = 'N'"
    l_strSQL = l_strSQL & "             then (hours + travel_pay_hrs) ELSE 0 end as OT_HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "     case when ot_opp_num = 0 and appr_ind = 'Y' AND symbol_time_ind = 'N' AND start_time not like '99%'"
    l_strSQL = l_strSQL & "             then hours ELSE 0 end as ST_HRS_APPROVED,"
    l_strSQL = l_strSQL & "     case when ot_opp_num = 0 and appr_ind = 'N' AND symbol_time_ind = 'N' AND start_time not like '99%'"
    l_strSQL = l_strSQL & "             then hours ELSE 0 end as ST_HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "     case when ot_opp_num = 0 and appr_ind = 'Y' AND symbol_time_ind = 'Y'"
    l_strSQL = l_strSQL & "             then hours ELSE 0 end as SYM_HRS_APPROVED,"
    l_strSQL = l_strSQL & "     case when ot_opp_num = 0 and appr_ind = 'N' AND symbol_time_ind = 'Y'"
    l_strSQL = l_strSQL & "             then hours ELSE 0 end as SYM_HRS_UNAPPROVED,"
    l_strSQL = l_strSQL & "     case when start_time like '99%' and appr_ind = 'N'"
    l_strSQL = l_strSQL & "            then 1 else 0 end as OTHERPAYIND_UNAPP,"
    l_strSQL = l_strSQL & "     case when start_time like '99%' and appr_ind = 'Y'"
    l_strSQL = l_strSQL & "            then 1 else 0 end as OTHERPAYIND_APP"
    l_strSQL = l_strSQL & " from " & l_strActTimeTable
    l_strSQL = l_strSQL & "     ,methdba.ttr_exception_cds b "
    l_strSQL = l_strSQL & " where " & l_strEmpSel
    l_strSQL = l_strSQL & l_strInClausePartValueKey
    l_strSQL = l_strSQL & "   and charge_date  >=  '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & "   and charge_date  <=  '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & "   and a.exception_cd = b.exception_cd"
    l_strSQL = l_strSQL & "   AND a.EXCEPTION_CD NOT IN "
    l_strSQL = l_strSQL & " ( SELECT RG_VALUE_CD FROM METHDBA.TTR_RG_GROUP WHERE RG_CD = 189 ) "
    l_strSQL = l_strSQL & ") DETAILS "
    l_strSQL = l_strSQL & " GROUP BY " & l_strGroupBy & " ) SUMMARY1 "
    l_strSQL = l_strSQL & " inner join methdba.ttr_employee EMP"
    l_strSQL = l_strSQL & " on SUMMARY1.EMP_NUM = EMP.EMP_NUM"


    GetTimeSummarySql = l_strSQL
    
    Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "GetTimeSummarySql", , , , DescArgsFrom(l_strSQL), True
    

End Function

