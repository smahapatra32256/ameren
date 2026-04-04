VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdEAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   Post TRIS 5.5 Employee Absence Management
'Author:    Prem Chander
'Date:      August 2005
'Description:

'TO DO: ON HOLD FOR INIT ABS. MGMT
'1. ADD NEW ERROR CODES TO mstdErrorHandling
'2. ADD NEW ERROR CODES TO RESOURCE FILE

Option Explicit
Private Const MODULENAME As String = "dbEAM.cdEAM"
'BEGIN PREM
Public Function GetAbsenceEmployeeInformation(ByVal vlEmployeeNumber As Long _
                ) As ADODB.Recordset
TrsRaiseError MODULENAME, "GetAbsenceEmployeeInformation", trErrObsoleteMethod, , , , True
'On Error GoTo GetAbsenceEmployeeInformation_EH
'Dim ladrs As ADODB.Recordset
'Dim lsSQL As String
'
'    Set GetAbsenceEmployeeInformation = Nothing
'
'    lsSQL = "SELECT A.EMP_NUM, A.EMP_NAME" & _
'        "           , A.LOC_CD, B.LOC_DESC" & _
'        "           , COALESCE(INTEGER(C.SUPV_ID), 0) AS SUPERVISOR_NUM, COALESCE(D.EMP_NAME, '') AS SUPERVISOR_NAME" & _
'        "           , A.HOME_PHONE_NUM AS EMP_PHONE" & _
'        "           , A.DEPT_NUM, A.PREV_DEPT_NUM" & _
'        "           , A.DEPT_EFF_DATE, A.UNION_CD" & _
'        "           , A.CMPY_CD, A.EXEC_CONTRACT_IND" & _
'        "           , A.MGMT_BEN_PKG_IND, E.UNION_BEN_PKG_IND" & _
'        " FROM      METHDBA.TTR_EMPLOYEE A" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_LOCATIONS B" & _
'        "           ON B.LOC_CD = A.LOC_CD" & _
'        "           LEFT OUTER JOIN METHDBA.TAR_COMN_EMPL_TBL C" & _
'        "           ON CHAR(A.EMP_NUM) = C.EMPL_ID" & _
'        "           AND C.SUPV_ID <> ''" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE D" & _
'        "           ON D.EMP_NUM = INTEGER(C.SUPV_ID)" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_UNIONS E" & _
'        "           ON E.UNION_CD = A.UNION_CD" & _
'        " WHERE     A.EMP_NUM = ?;"
'
'    Set ladrs = CreateParameterRS
'    With ladrs
'        .AddNew
'        !PARM_ORDER = 1
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlEmployeeNumber
'        .Update
'    End With
'
'    Set GetAbsenceEmployeeInformation = NewExecuteSQL(lsSQL _
'                            , ladrs _
'                            , adOpenStatic _
'                            , adLockReadOnly)
'
'    Set ladrs = Nothing
'
'    Exit Function
'GetAbsenceEmployeeInformation_EH:
'    Set ladrs = Nothing
'
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME _
'                    , "GetAbsenceEmployeeInformation"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "GetAbsenceEmployeeInformation" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(lsSQL) _
'                    , True
'            Case Else
'                TrsRaiseError MODULENAME _
'                    , "GetAbsenceEmployeeInformation" _
'                    , trErrdbGetEAMEmployee _
'                    , , _
'                    , DescArgsFrom(lsSQL) _
'                    , True
'    End Select

End Function
Public Function GetEmployeeAbsences(ByVal vlEmployeeNumber As Long _
                , ByVal vdtStartDate As Date _
                , ByVal vdtEndDate As Date _
                , ByVal vlTRISUserNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetEmployeeAbsences_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetEmployeeAbsences = Nothing

    lsSQL = "SELECT A.EMP_NUM, C.EMP_NAME" & _
        "           , A.ABS_NUM, COALESCE(RTRIM(LTRIM(D.EXCEPTION_DESC)), '') AS ABS_DESC" & _
        "           , DATE(A.ABS_ST_TS) AS ABS_ST_DT, CHAR(TIME(A.ABS_ST_TS)) AS ABS_ST_TM" & _
        "           , DATE(A.ABS_END_TS) AS ABS_END_DT, CHAR(TIME(A.ABS_END_TS)) AS ABS_END_TM" & _
        "           , DATE(A.RTN_TO_WRK_TS) AS PROB_RTN_WRK_DT, CHAR(TIME(A.RTN_TO_WRK_TS)) AS PROB_RTN_WRK_TM" & _
        "           , COALESCE(DATE(B.RTN_TO_WRK_TS), DATE('01/01/0100')) AS TRUE_RTN_WRK_DT" & _
        "           , COALESCE(CHAR(TIME(B.RTN_TO_WRK_TS)), '') AS TRUE_RTN_WRK_TM" & _
        "           , A.ABS_ST_TS AS ABS_ST_TS, A.RTN_TO_WRK_TS AS PROB_RTN_WRK_TS" & _
        "           , COALESCE(CHAR(B.RTN_TO_WRK_TS), '') AS TRUE_RTN_WRK_TS" & _
        "           , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "           , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        "           , COALESCE(E.EMP_NAME, '') AS CREATED_EMP_NAME, COALESCE(F.EMP_NAME, '') AS LAST_MOD_EMP_NAME"
    lsSQL = lsSQL & _
        " FROM      METHDBA.TTR_ABS_INIT A" & _
        "           LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK B" & _
        "           ON A.ABS_NUM = B.ABS_NUM" & _
        "           LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE C" & _
        "           ON A.EMP_NUM = C.EMP_NUM" & _
        "           LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS D" & _
        "           ON A.EXCEPTION_CD = D.EXCEPTION_CD" & _
        "           LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE E" & _
        "           ON A.CREATED_EMP_NUM = E.EMP_NUM" & _
        "           LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE F" & _
        "           ON A.LAST_MOD_EMP_NUM = F.EMP_NUM" & _
        " WHERE     A.EMP_NUM = ?" & _
        " AND       DATE(A.ABS_ST_TS) BETWEEN ?" & _
        " AND       ?" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = Format(vdtStartDate, "yyyy-mm-dd")
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = Format(vdtEndDate, "yyyy-mm-dd")
        .Update
    End With

    Set GetEmployeeAbsences = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    Exit Function
GetEmployeeAbsences_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetEmployeeAbsences"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetEmployeeAbsences" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetEmployeeAbsences" _
                    , trErrdbGetEAM _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetInitialAbsenceByAbsence(ByVal vlAbsenceNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetInitialAbsenceByAbsence_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetInitialAbsenceByAbsence = Nothing

    lsSQL = "SELECT  A.ABS_NUM, '' AS ABS_DESC" & _
        "       , A.EMP_NUM, B.EMP_NAME" & _
        "       , DATE(A.ABS_CALLIN_TS) AS ABS_CALLIN_DT, CHAR(TIME(A.ABS_CALLIN_TS)) AS ABS_CALLIN_TM" & _
        "       , DATE(A.ABS_ST_TS) AS ABS_ST_DT, CHAR(TIME(A.ABS_ST_TS)) AS ABS_ST_TM" & _
        "       , DATE(A.ABS_END_TS) AS ABS_END_DT, CHAR(TIME(A.ABS_END_TS)) AS ABS_END_TM" & _
        "       , CHAR(A.ABS_RPT_TS) AS ABS_RPT_TS, A.EMP_PHONE" & _
        "       , A.ABS_LOC_CD, COALESCE(C.LOC_DESC, '') AS ABS_LOC_DESC" & _
        "       , A.ABS_RPTR_TYPE, COALESCE(A.ABS_RPTR_TXT, '') AS ABS_RPTR_TXT" & _
        "       , A.EMP_LOC_CD, COALESCE(D.LOC_DESC, '') AS EMP_LOC_DESC" & _
        "       , A.EXCEPTION_CD, COALESCE(A.OVER_EXC_CD, '') AS OVER_EXC_CD" & _
        "       , COALESCE(A.ADDL_TXT, '') AS ADDL_TXT" & _
        "       , CASE (A.FMLA_REASON_CD)" & _
        "       WHEN 'NA' THEN 'NOT APPLICABLE'" & _
        "       WHEN 'BA' THEN 'BIRTH OR ADOPTION'" & _
        "       WHEN 'CC' THEN 'CARE OF CHILD'" & _
        "       WHEN 'CP' THEN 'CARE OF PARENT'" & _
        "       WHEN 'CS' THEN 'CARE OF SPOUSE'" & _
        "       WHEN 'EI' THEN 'EMPLOYEE ILLNESS'" & _
        "       ELSE 'NOT APPLICABLE' END AS FMLA_REASON"
    lsSQL = lsSQL & _
        "       , COALESCE(A.FMLA_REASON_CD, '') AS FMLA_REASON_CD" & _
        "       , DATE(A.RTN_TO_WRK_TS) AS RTN_TO_WRK_DT, CHAR(TIME(A.RTN_TO_WRK_TS)) AS RTN_TO_WRK_TM" & _
        "       , '' AS RTN_TO_WRK_SFT" & _
        "       , COALESCE(A.FOLW_UP_DT, DATE('01/01/0100')) AS FOLW_UP_DT" & _
        "       , CASE (A.RTW_NOT_REQD_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS RTW_NOT_REQD_IND" & _
        "       , COALESCE(INTEGER(E.SUPV_ID), 0) AS SUPERVISOR_NUM, COALESCE(F.EMP_NAME, '') AS SUPERVISOR_NAME" & _
        "       , A.RPTNG_EMP_NUM, COALESCE(G.EMP_NAME, '') AS RPTNG_EMP_NAME" & _
        "       , A.ABS_PRIM_CONT_EMP, COALESCE(H.EMP_NAME, '') AS PRIM_CONT_EMP_NAME" & _
        "       , COALESCE(A.ABS_NOTIFY_WG_NUM, 0) AS ABS_NOTIFY_WG_NUM, COALESCE(W.WG_DESC, '') AS WG_DESC" & _
        "       , COALESCE(A.PRVDR_NM, '') AS PRVDR_NM, COALESCE(A.PRVDR_PHN_NBR, '') AS PRVDR_PHN_NBR" & _
        "       , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP"
    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_ABS_INIT A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON A.EMP_NUM = B.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_LOCATIONS C" & _
        "       ON A.ABS_LOC_CD = C.LOC_CD" & _
        "       LEFT OUTER JOIN METHDBA.TTR_LOCATIONS D" & _
        "       ON A.EMP_LOC_CD = D.LOC_CD" & _
        "       LEFT OUTER JOIN METHDBA.TAR_COMN_EMPL_TBL E" & _
        "       ON CHAR(A.EMP_NUM) = E.EMPL_ID" & _
        "       AND E.SUPV_ID <> ''" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE F" & _
        "       ON INTEGER(E.SUPV_ID) = F.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE G" & _
        "       ON A.RPTNG_EMP_NUM = G.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE H" & _
        "       ON A.ABS_PRIM_CONT_EMP = H.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_WORK_GROUP W" & _
        "       ON W.WG_NUM = A.ABS_NOTIFY_WG_NUM" & _
        " WHERE A.ABS_NUM = ?;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlAbsenceNumber
        .Update
    End With

    Set GetInitialAbsenceByAbsence = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockBatchOptimistic)
    Set ladrs = Nothing

    Exit Function
GetInitialAbsenceByAbsence_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetInitialAbsenceByAbsence"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetInitialAbsenceByAbsence" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetInitialAbsenceByAbsence" _
                    , trErrdbGetAbsenceInitial _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select
End Function
Public Function GetAbsenceDatesByAbsence(ByVal vlAbsenceNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetAbsenceDatesByAbsence_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetAbsenceDatesByAbsence = Nothing

    lsSQL = "SELECT  A.ABS_NUM, B.ABS_DATE" & _
        "       , B.CREATED_EMP_NUM, CHAR(B.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , B.LAST_MOD_EMP_NUM, CHAR(B.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        "       , 0 AS SHIFT_1_SCHED, '' AS SHIFT_1_SHIFT" & _
        "       , '' AS SHIFT_1_START, '' AS SHIFT_1_END" & _
        "       , 0 AS SHIFT_2_SCHED, '' AS SHIFT_2_SHIFT" & _
        "       , '' AS SHIFT_2_START, '' AS SHIFT_2_END" & _
        " FROM  METHDBA.TTR_ABS_INIT A, METHDBA.TTR_ABS_DATE B" & _
        " WHERE  A.ABS_NUM = ?" & _
        " AND   A.ABS_NUM = B.ABS_NUM;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlAbsenceNumber
        .Update
    End With

    Set GetAbsenceDatesByAbsence = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockBatchOptimistic)
    Set ladrs = Nothing

    Exit Function
GetAbsenceDatesByAbsence_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetAbsenceDatesByAbsence"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetAbsenceDatesByAbsence" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetAbsenceDatesByAbsence" _
                    , trErrdbGetAbsenceDates _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetAbsenceProviderByAbsence(ByVal vlAbsenceNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetAbsenceProviderByAbsence_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetAbsenceProviderByAbsence = Nothing

    lsSQL = "SELECT  A.ABS_NUM, A.PRVDR_NBR" & _
        "       , COALESCE(B.PRVDR_NM, '') AS PRVDR_NM" & _
        "       , DATE(A.VST_ST_ADM_TS) AS VST_ST_ADM_DT, CHAR(TIME(A.VST_ST_ADM_TS)) AS VST_ST_ADM_TM" & _
        "       , DATE(A.VST_END_DGCE_TS) AS VST_END_DGCE_DT, CHAR(TIME(A.VST_END_DGCE_TS)) AS VST_END_DGCE_TM" & _
        "       , CHAR(A.VST_ST_ADM_TS) AS VST_ST_ADM_TS, CHAR(A.VST_END_DGCE_TS) AS VST_END_DGCE_TS" & _
        "       , COALESCE(B.CLS_TXT, '') AS CLS_TXT, COALESCE(B.CTY_NM, '') AS CTY_NM" & _
        "       , COALESCE(B.LN1_ADDR, '') AS LN1_ADDR, COALESCE(B.LN2_ADDR, '') AS LN2_ADDR" & _
        "       , COALESCE(B.ST_CD, '') AS ST_CD, COALESCE(B.ZIP_CD, '') AS ZIP_CD" & _
        "       , COALESCE(B.PHN_NBR, '') AS PHN_NBR" & _
        "       , CASE (B.TYP_CD)" & _
        "       WHEN 'A' THEN 'PHYSICIAN ASSISTANT'" & _
        "       WHEN 'D' THEN 'DOCTOR'" & _
        "       WHEN 'H' THEN 'HOSPITAL'" & _
        "       WHEN 'N' THEN 'NURSE PRACTITIONER'" & _
        "       WHEN 'P' THEN 'PHYSICAL THERAPIST'" & _
        "       WHEN 'R' THEN 'REHAB. FACILITY'" & _
        "       WHEN 'S' THEN 'SPECIALIST'" & _
        "       WHEN 'T' THEN 'DENTIST'" & _
        "       ELSE '' END AS TYP_DESC" & _
        "       , CASE (A.RLSE_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS RLSE_IND"
    lsSQL = lsSQL & _
        "       , COALESCE(A.ADDL_TXT, '') AS ADDL_TXT" & _
        "       , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM  METHDBA.TTR_ABS_PROVIDER A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_PROVIDER B" & _
        "       ON A.PRVDR_NBR = B.PRVDR_NBR" & _
        " WHERE  A.ABS_NUM = ?;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlAbsenceNumber
        .Update
    End With

'    Set GetAbsenceProviderByAbsence = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetAbsenceProviderByAbsence = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    Exit Function
GetAbsenceProviderByAbsence_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetAbsenceProviderByAbsence"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetAbsenceProviderByAbsence" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetAbsenceProviderByAbsence" _
                    , trErrdbGetAbsenceProvider _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetContactByAbsence(ByVal vlAbsenceNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetContactByAbsence_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetContactByAbsence = Nothing

    lsSQL = "SELECT  A.ABS_NUM, A.ABS_CONT_NUM" & _
        "           , LTRIM(RTRIM(CHAR(A.CREATED_EMP_NUM))) || ' - ' ||  CHAR(ABS_CONT_TS) AS ABS_CONT_DESC" & _
        "           , DATE(A.ABS_CONT_TS) AS ABS_CONT_DT, CHAR(TIME(A.ABS_CONT_TS)) AS ABS_CONT_TM" & _
        "           , CHAR(A.ABS_CONT_TS) AS ABS_CONT_TS, COALESCE(A.ADDL_TXT, '') AS ADDL_TXT" & _
        "           , CASE (A.USR_INIT_IND)" & _
        "           WHEN 'Y' THEN '1'" & _
        "           WHEN '1' THEN '1'" & _
        "           ELSE '0' END AS USR_INIT_IND" & _
        "           , CASE (A.TALK_TO_EMP_IND)" & _
        "           WHEN 'Y' THEN '1'" & _
        "           WHEN '1' THEN '1'" & _
        "           ELSE '0' END AS TALK_TO_EMP_IND, COALESCE(A.TALK_TO_EMP_TXT, '') AS TALK_TO_EMP_TXT" & _
        "           , CASE (A.EMP_TO_DR_IND)" & _
        "           WHEN 'Y' THEN '1'" & _
        "           WHEN '1' THEN '1'" & _
        "           ELSE '0' END AS EMP_TO_DR_IND, COALESCE(A.EMP_TO_DR_TXT, '') AS EMP_TO_DR_TXT" & _
        "           , CASE (A.OTH_TRTMNT_IND)" & _
        "           WHEN 'Y' THEN '1'" & _
        "           WHEN '1' THEN '1'" & _
        "           ELSE '0' END AS OTH_TRTMNT_IND, COALESCE(A.OTH_TRTMNT_TXT, '') AS OTH_TRTMNT_TXT" & _
        "           , CASE (A.ADVS_SL_IND)" & _
        "           WHEN 'Y' THEN '1'" & _
        "           WHEN '1' THEN '1'" & _
        "           ELSE '0' END AS ADVS_SL_IND, COALESCE(A.ADVS_SL_TXT, '') AS ADVS_SL_TXT" & _
        "           , A.RPTNG_EMP_NUM, B.EMP_NAME AS RPTNG_EMP_NAME"
    lsSQL = lsSQL & _
        "           , COALESCE(A.PHONE_NUM, '') AS PHONE_NUM, COALESCE(A.PRVDR_PHN_NUM, '') AS PRVDR_PHN_NUM" & _
        "           , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "           , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM      METHDBA.TTR_ABS_CONTACT A" & _
        "           INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "           ON B.EMP_NUM = A.RPTNG_EMP_NUM" & _
        " WHERE     A.ABS_NUM = ?;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlAbsenceNumber
        .Update
    End With

    Set GetContactByAbsence = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)

    Set ladrs = Nothing

    Exit Function
GetContactByAbsence_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetContactByAbsence"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetContactByAbsence" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetContactByAbsence" _
                    , trErrdbGetAbsenceContact _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetReturnToWorkByAbsence(ByVal vlAbsenceNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetReturnToWorkByAbsence_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetReturnToWorkByAbsence = Nothing

    lsSQL = "SELECT  A.ABS_NUM, COALESCE(A.ABS_REASON_TXT, '') AS ABS_REASON_TXT" & _
        "           , DATE(A.RTN_TO_WRK_TS) AS RTN_TO_WRK_DT, CHAR(TIME(A.RTN_TO_WRK_TS)) AS RTN_TO_WRK_TM" & _
        "           , CASE (A.CONFRM_RSN_IND)" & _
        "           WHEN 'Y' THEN '1'" & _
        "           WHEN '1' THEN '1'" & _
        "           ELSE '0' END AS CONFRM_RSN_IND" & _
        "           , CASE (A.PRFM_DUTY_IND)" & _
        "           WHEN 'Y' THEN '1'" & _
        "           WHEN '1' THEN '1'" & _
        "           ELSE '0' END AS PRFM_DUTY_IND, COALESCE(A.PRFM_DUTY_TXT, '') AS PRFM_DUTY_TXT" & _
        "           , CASE (A.SIDE_EFF_IND)" & _
        "           WHEN 'Y' THEN '1'" & _
        "           WHEN '1' THEN '1'" & _
        "           ELSE '0' END AS SIDE_EFF_IND, COALESCE(A.SIDE_EFF_TXT, '') AS SIDE_EFF_TXT" & _
        "           , CASE (A.MEET_EMP_IND)" & _
        "           WHEN 'Y' THEN '1'" & _
        "           WHEN '1' THEN '1'" & _
        "           ELSE '0' END AS MEET_EMP_IND, COALESCE(A.MEET_EMP_TXT, '') AS MEET_EMP_TXT"
    lsSQL = lsSQL & _
        "           , COALESCE(A.ADDL_TXT, '') AS ADDL_TXT, A.RPT_DT" & _
        "           , A.RPTNG_EMP_NUM, B.EMP_NAME AS RPTNG_EMP_NAME" & _
        "           , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "           , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM      METHDBA.TTR_ABS_RET_WRK A" & _
        "           INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "           ON B.EMP_NUM = A.RPTNG_EMP_NUM" & _
        " WHERE     A.ABS_NUM = ?;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlAbsenceNumber
        .Update
    End With

    Set GetReturnToWorkByAbsence = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    Exit Function
GetReturnToWorkByAbsence_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorkByAbsence"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorkByAbsence" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorkByAbsence" _
                    , trErrdbGetAbsenceReturnToWork _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetContactCallLogByAbsence(ByVal vlAbsenceNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetContactCallLogByAbsence_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetContactCallLogByAbsence = Nothing

    lsSQL = "SELECT  A.ABS_NUM, A.CALL_NUM" & _
        "           , CHAR(A.CALL_ST_TS) AS CALL_ST_TS, CHAR(A.CALL_END_TS) AS CALL_END_TS" & _
        "           , A.CALLER_EMP_NUM, B.EMP_NAME AS CALLER_EMP_NAME" & _
        "           , A.PHONE_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        " FROM      METHDBA.TTR_ABS_CALL_LOG A" & _
        "           INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "           ON B.EMP_NUM = A.CALLER_EMP_NUM" & _
        " WHERE     A.ABS_NUM = ?;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlAbsenceNumber
        .Update
    End With

    Set GetContactCallLogByAbsence = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    Exit Function
GetContactCallLogByAbsence_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetContactCallLogByAbsence"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetContactCallLogByAbsence" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetContactCallLogByAbsence" _
                    , trErrdbGetAbsenceContactCallLog _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeePortfolios(ByVal vlEmployeeNumber As Long _
                , ByVal vdtStartDate As Date _
                , ByVal vdtEndDate As Date _
                , ByVal vlTRISUserNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetEmployeePortfolios_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetEmployeePortfolios = Nothing

    lsSQL = "SELECT  A.PRTFL_NUM, A.PRTFL_DESC" & _
        "           , CASE (A.SHARE_IND)" & _
        "           WHEN 'Y' THEN '1'" & _
        "           WHEN '1' THEN '1'" & _
        "           ELSE '0' END AS SHARE_IND" & _
        "           , A.CREATED_EMP_NUM, D.EMP_NAME AS CREATED_EMP_NAME" & _
        "           , A.LAST_MOD_EMP_NUM, E.EMP_NAME AS LAST_MOD_EMP_NAME" & _
        "           , CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM      METHDBA.TTR_ABS_PORTFOLIO A" & _
        "           LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE D" & _
        "           ON D.EMP_NUM = A.CREATED_EMP_NUM" & _
        "           LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE E" & _
        "           ON E.EMP_NUM = A.LAST_MOD_EMP_NUM" & _
        " WHERE     A.PRTFL_NUM IN (SELECT    DISTINCT C.PRTFL_NUM" & _
        "                       FROM        METHDBA.TTR_ABS_INIT B" & _
        "                                   , METHDBA.TTR_ABS_PRTFL_DTL C" & _
        "                       WHERE       B.EMP_NUM = ?" & _
        "                       AND         B.ABS_NUM = C.ABS_NUM" & _
        "                       AND         DATE(B.ABS_ST_TS) BETWEEN ?" & _
        "                       AND         ?"
    lsSQL = lsSQL & _
        " AND       ((A.SHARE_IND IN ('N', '0')" & _
        " AND       A.CREATED_EMP_NUM = ?)" & _
        " OR        A.SHARE_IND IN ('Y', '1')" & _
        " OR        (A.SHARE_IND = 'N'" & _
        " AND       EXISTS  (SELECT EMP_NUM" & _
        "                   FROM    METHDBA.TTR_EMPLOYEE" & _
        "                   WHERE   EMP_NUM = ?" & _
        "                   AND     SEC_LEVEL IN (9, 10, 11, 981)))))" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = Format(vdtStartDate, "yyyy-mm-dd")
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = Format(vdtEndDate, "yyyy-mm-dd")
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlTRISUserNumber
        .Update

        .AddNew
        !PARM_ORDER = 5
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlTRISUserNumber
        .Update
    End With

    Set GetEmployeePortfolios = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    Exit Function
GetEmployeePortfolios_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetEmployeePortfolios"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetEmployeePortfolios" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetEmployeePortfolios" _
                    , trErrdbGetAbsencePortfolio _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetPortfolio(ByVal vlPortfolioNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetPortfolio_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetPortfolio = Nothing

    lsSQL = "SELECT  PRTFL_NUM, PRTFL_DESC" & _
        "           , CASE (SHARE_IND)" & _
        "           WHEN 'Y' THEN '1'" & _
        "           WHEN '1' THEN '1'" & _
        "           ELSE '0' END AS SHARE_IND" & _
        "           , CREATED_EMP_NUM, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "           , LAST_MOD_EMP_NUM, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM      METHDBA.TTR_ABS_PORTFOLIO" & _
        " WHERE     PRTFL_NUM = ?;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlPortfolioNumber
        .Update
    End With

    Set GetPortfolio = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    Exit Function
GetPortfolio_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetPortfolio"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetPortfolio" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetPortfolio" _
                    , trErrdbGetEAM _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetPortfolioAbsences(ByVal vlPortfolioNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetPortfolioAbsences_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetPortfolioAbsences = Nothing

    lsSQL = "SELECT A.PRTFL_NUM, A.ABS_NUM" & _
        "   , COALESCE(RTRIM(LTRIM(C.EXCEPTION_DESC)), '') AS ABS_DESC" & _
        "   , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "   , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        "   , B.ABS_ST_TS, B.ABS_END_TS" & _
        "   , B.RTN_TO_WRK_TS AS PROB_RTN_WRK_TS" & _
        "   , COALESCE(CHAR(D.RTN_TO_WRK_TS), '') AS TRUE_RTN_WRK_TS" & _
        "   , COALESCE(DATE(D.RTN_TO_WRK_TS), DATE('01/01/0100')) AS TRUE_RTN_WRK_DT" & _
        "   , COALESCE(CHAR(TIME(D.RTN_TO_WRK_TS)), '') AS TRUE_RTN_WRK_TM" & _
        " FROM  METHDBA.TTR_ABS_PRTFL_DTL A, METHDBA.TTR_ABS_INIT B" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
        "       ON B.EXCEPTION_CD = C.EXCEPTION_CD" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK D" & _
        "       ON B.ABS_NUM = D.ABS_NUM" & _
        " WHERE A.PRTFL_NUM = ?" & _
        " AND   A.ABS_NUM = B.ABS_NUM" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlPortfolioNumber
        .Update
    End With

    Set GetPortfolioAbsences = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    Exit Function
GetPortfolioAbsences_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsences"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsences" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsences" _
                    , trErrdbGetEAM _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetPortfolioInitialAbsence(ByVal vlPortfolioNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetPortfolioInitialAbsence_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetPortfolioInitialAbsence = Nothing

    lsSQL = "SELECT  A.ABS_NUM, '' AS ABS_DESC" & _
        "       , B.EMP_NUM, C.EMP_NAME" & _
        "       , B.ABS_CALLIN_TS, B.ABS_ST_TS" & _
        "       , B.ABS_END_TS" & _
        "       , B.ABS_RPT_TS, B.EMP_PHONE" & _
        "       , B.ABS_LOC_CD, COALESCE(D.LOC_DESC, '') AS ABS_LOC_DESC" & _
        "       , B.ABS_RPTR_TYPE, COALESCE(B.ABS_RPTR_TXT, '') AS ABS_RPTR_TXT" & _
        "       , B.EMP_LOC_CD, COALESCE(E.LOC_DESC, '') AS EMP_LOC_DESC" & _
        "       , B.EXCEPTION_CD, COALESCE(B.OVER_EXC_CD, '') AS OVER_EXC_CD" & _
        "       , COALESCE(B.ADDL_TXT, '') AS ADDL_TXT" & _
        "       , CASE (B.FMLA_REASON_CD)" & _
        "       WHEN 'NA' THEN 'NOT APPLICABLE'" & _
        "       WHEN 'BA' THEN 'BIRTH OR ADOPTION'" & _
        "       WHEN 'CC' THEN 'CARE OF CHILD'" & _
        "       WHEN 'CP' THEN 'CARE OF PARENT'" & _
        "       WHEN 'CS' THEN 'CARE OF SPOUSE'" & _
        "       WHEN 'EI' THEN 'EMPLOYEE ILLNESS'" & _
        "       ELSE 'NOT APPLICABLE' END AS FMLA_REASON" & _
        "       , COALESCE(B.FMLA_REASON_CD, '') AS FMLA_REASON_CD" & _
        "       , B.RTN_TO_WRK_TS" & _
        "       , COALESCE(B.FOLW_UP_DT, DATE('01/01/0100')) AS FOLW_UP_DT" & _
        "       , COALESCE(INTEGER(F.SUPV_ID), 0) AS SUPERVISOR_NUM, COALESCE(G.EMP_NAME, '') AS SUPERVISOR_NAME" & _
        "       , B.RPTNG_EMP_NUM, COALESCE(H.EMP_NAME, '') AS RPTNG_EMP_NAME" & _
        "       , B.ABS_PRIM_CONT_EMP, COALESCE(I.EMP_NAME, '') AS PRIM_CONT_EMP_NAME"
    lsSQL = lsSQL & _
        "       , B.ABS_NOTIFY_WG_NUM, COALESCE(W.WG_DESC, '') AS WG_DESC" & _
        "       , COALESCE(B.ABS_NOTIFY_WG_NUM, 0) AS ABS_NOTIFY_WG_NUM, COALESCE(W.WG_DESC, '') AS WG_DESC" & _
        "       , COALESCE(B.PRVDR_NM, '') AS PRVDR_NM, COALESCE(B.PRVDR_PHN_NBR, '') AS PRVDR_PHN_NBR" & _
        "       , CASE (B.RTW_NOT_REQD_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS RTW_NOT_REQD_IND" & _
        "       , B.CREATED_EMP_NUM, CHAR(B.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , B.LAST_MOD_EMP_NUM, CHAR(B.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP"
    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_ABS_PRTFL_DTL A, METHDBA.TTR_ABS_INIT B" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE C" & _
        "       ON B.EMP_NUM = C.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_LOCATIONS D" & _
        "       ON B.ABS_LOC_CD = D.LOC_CD" & _
        "       LEFT OUTER JOIN METHDBA.TTR_LOCATIONS E" & _
        "       ON B.EMP_LOC_CD = E.LOC_CD" & _
        "       LEFT OUTER JOIN METHDBA.TAR_COMN_EMPL_TBL F" & _
        "       ON CHAR(B.EMP_NUM) = F.EMPL_ID" & _
        "       AND F.SUPV_ID <> ''" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE G" & _
        "       ON INTEGER(F.SUPV_ID) = G.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE H" & _
        "       ON B.RPTNG_EMP_NUM = H.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE I" & _
        "       ON B.ABS_PRIM_CONT_EMP = I.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_WORK_GROUP W" & _
        "       ON W.WG_NUM = B.ABS_NOTIFY_WG_NUM" & _
        " WHERE A.PRTFL_NUM = ?" & _
        " AND   A.ABS_NUM = B.ABS_NUM" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlPortfolioNumber
        .Update
    End With

    Set GetPortfolioInitialAbsence = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)

    Set ladrs = Nothing

    Exit Function
GetPortfolioInitialAbsence_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetPortfolioInitialAbsence"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetPortfolioInitialAbsence" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetPortfolioInitialAbsence" _
                    , trErrdbGetAbsenceInitial _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetPortfolioAbsenceProviders(ByVal vlPortfolioNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetPortfolioAbsenceProviders_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetPortfolioAbsenceProviders = Nothing

    lsSQL = "SELECT  A.ABS_NUM" & _
        "       , B.PRVDR_NBR, C.PRVDR_NM" & _
        "       , B.VST_ST_ADM_TS, B.VST_END_DGCE_TS" & _
        "       , COALESCE(C.CLS_TXT, '') AS CLS_TXT, COALESCE(C.CTY_NM, '') AS CTY_NM" & _
        "       , COALESCE(C.LN1_ADDR, '') AS LN1_ADDR, COALESCE(C.LN2_ADDR, '') AS LN2_ADDR" & _
        "       , COALESCE(C.ST_CD, '') AS ST_CD, COALESCE(C.ZIP_CD, '') AS ZIP_CD" & _
        "       , COALESCE(C.PHN_NBR, '') AS PHN_NBR, COALESCE(B.ADDL_TXT, '') AS ADDL_TXT" & _
        "       , CASE (C.TYP_CD)" & _
        "       WHEN 'A' THEN 'PHYSICIAN ASSISTANT'" & _
        "       WHEN 'D' THEN 'DOCTOR'" & _
        "       WHEN 'H' THEN 'HOSPITAL'" & _
        "       WHEN 'N' THEN 'NURSE PRACTITIONER'" & _
        "       WHEN 'P' THEN 'PHYSICAL THERAPIST'" & _
        "       WHEN 'R' THEN 'REHAB. FACILITY'" & _
        "       WHEN 'S' THEN 'SPECIALIST'" & _
        "       WHEN 'T' THEN 'DENTIST'" & _
        "       ELSE '' END AS TYP_DESC" & _
        "       , CASE (B.RLSE_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS RLSE_IND"
    lsSQL = lsSQL & _
        "       , B.CREATED_EMP_NUM, CHAR(B.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , B.LAST_MOD_EMP_NUM, CHAR(B.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM  METHDBA.TTR_ABS_PRTFL_DTL A, METHDBA.TTR_ABS_PROVIDER B" & _
        "       LEFT OUTER JOIN METHDBA.TTR_PROVIDER C" & _
        "       ON B.PRVDR_NBR = C.PRVDR_NBR" & _
        " WHERE A.PRTFL_NUM = ?" & _
        " AND   A.ABS_NUM = B.ABS_NUM" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlPortfolioNumber
        .Update
    End With

    Set GetPortfolioAbsenceProviders = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    Exit Function
GetPortfolioAbsenceProviders_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceProviders"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceProviders" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceProviders" _
                    , trErrdbGetAbsenceProvider _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetPortfolioAbsenceContacts(ByVal vlPortfolioNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetPortfolioAbsenceContacts_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetPortfolioAbsenceContacts = Nothing

    lsSQL = "SELECT  A.ABS_NUM, B.ABS_CONT_NUM" & _
        "       , LTRIM(RTRIM(CHAR(B.CREATED_EMP_NUM))) || ' - ' ||  CHAR(B.ABS_CONT_TS) AS ABS_CONT_DESC" & _
        "       , B.ABS_CONT_TS, COALESCE(B.ADDL_TXT, '') AS ADDL_TXT" & _
        "       , CASE (B.USR_INIT_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS USR_INIT_IND" & _
        "       , CASE (B.TALK_TO_EMP_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS TALK_TO_EMP_IND, COALESCE(B.TALK_TO_EMP_TXT, '') AS TALK_TO_EMP_TXT" & _
        "       , CASE (B.EMP_TO_DR_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS EMP_TO_DR_IND, COALESCE(B.EMP_TO_DR_TXT, '') AS EMP_TO_DR_TXT" & _
        "       , CASE (B.OTH_TRTMNT_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS OTH_TRTMNT_IND, COALESCE(B.OTH_TRTMNT_TXT, '') AS OTH_TRTMNT_TXT" & _
        "       , CASE (B.ADVS_SL_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS ADVS_SL_IND, COALESCE(B.ADVS_SL_TXT, '') AS ADVS_SL_TXT" & _
        "       , B.RPTNG_EMP_NUM, C.EMP_NAME AS RPTNG_EMP_NAME"
    lsSQL = lsSQL & _
        "       , COALESCE(B.PHONE_NUM, '') AS PHONE_NUM, COALESCE(B.PRVDR_PHN_NUM, '') AS PRVDR_PHN_NUM" & _
        "       , B.CREATED_EMP_NUM, CHAR(B.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , B.LAST_MOD_EMP_NUM, CHAR(B.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM      METHDBA.TTR_ABS_PRTFL_DTL A, METHDBA.TTR_ABS_CONTACT B" & _
        "           INNER JOIN METHDBA.TTR_EMPLOYEE C" & _
        "           ON C.EMP_NUM = B.RPTNG_EMP_NUM" & _
        " WHERE A.PRTFL_NUM = ?" & _
        " AND   A.ABS_NUM = B.ABS_NUM" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlPortfolioNumber
        .Update
    End With

    Set GetPortfolioAbsenceContacts = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)

    Set ladrs = Nothing

    Exit Function
GetPortfolioAbsenceContacts_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceContacts"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceContacts" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceContacts" _
                    , trErrdbGetAbsenceContact _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetPortfolioAbsenceRtnToWork(ByVal vlPortfolioNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetPortfolioAbsenceRtnToWork_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetPortfolioAbsenceRtnToWork = Nothing

    lsSQL = "SELECT  A.ABS_NUM, COALESCE(B.ABS_REASON_TXT, '') AS ABS_REASON_TXT" & _
        "       , B.RTN_TO_WRK_TS" & _
        "       , CASE (B.CONFRM_RSN_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS CONFRM_RSN_IND" & _
        "       , CASE (B.PRFM_DUTY_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS PRFM_DUTY_IND, COALESCE(B.PRFM_DUTY_TXT, '') AS PRFM_DUTY_TXT" & _
        "       , CASE (B.SIDE_EFF_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS SIDE_EFF_IND, COALESCE(B.SIDE_EFF_TXT, '') AS SIDE_EFF_TXT" & _
        "       , CASE (B.MEET_EMP_IND)" & _
        "       WHEN 'Y' THEN '1'" & _
        "       WHEN '1' THEN '1'" & _
        "       ELSE '0' END AS MEET_EMP_IND, COALESCE(B.MEET_EMP_TXT, '') AS MEET_EMP_TXT"
    lsSQL = lsSQL & _
        "       , COALESCE(B.ADDL_TXT, '') AS ADDL_TXT, B.RPT_DT" & _
        "       , B.RPTNG_EMP_NUM, C.EMP_NAME AS RPTNG_EMP_NAME" & _
        "       , B.CREATED_EMP_NUM, CHAR(B.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , B.LAST_MOD_EMP_NUM, CHAR(B.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM  METHDBA.TTR_ABS_PRTFL_DTL A, METHDBA.TTR_ABS_RET_WRK B" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE C" & _
        "       ON C.EMP_NUM = B.RPTNG_EMP_NUM" & _
        " WHERE A.PRTFL_NUM = ?" & _
        " AND   A.ABS_NUM = B.ABS_NUM" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlPortfolioNumber
        .Update
    End With

    Set GetPortfolioAbsenceRtnToWork = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    Exit Function
GetPortfolioAbsenceRtnToWork_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceRtnToWork"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceRtnToWork" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceRtnToWork" _
                    , trErrdbGetAbsenceReturnToWork _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetPortfolioAbsenceDates(ByVal vlPortfolioNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetPortfolioAbsenceDates_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetPortfolioAbsenceDates = Nothing

    lsSQL = "SELECT  A.ABS_NUM, B.ABS_DATE" & _
        "       , B.CREATED_EMP_NUM, CHAR(B.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , B.LAST_MOD_EMP_NUM, CHAR(B.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM  METHDBA.TTR_ABS_PRTFL_DTL A, METHDBA.TTR_ABS_DATE B" & _
        " WHERE A.PRTFL_NUM = ?" & _
        " AND   A.ABS_NUM = B.ABS_NUM" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlPortfolioNumber
        .Update
    End With

    Set GetPortfolioAbsenceDates = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    Exit Function
GetPortfolioAbsenceDates_EH:
    Set ladrs = Nothing
    
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceDates"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceDates" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceDates" _
                    , trErrdbGetAbsenceDates _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetPortfolioAbsenceCallLog(ByVal vlPortfolioNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetPortfolioAbsenceCallLog_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

    Set GetPortfolioAbsenceCallLog = Nothing

    lsSQL = "SELECT  A.ABS_NUM, B.CALL_NUM" & _
        "           , B.CALL_ST_TS, B.CALL_END_TS" & _
        "           , B.CALLER_EMP_NUM, C.EMP_NAME AS CALLER_EMP_NAME" & _
        "           , B.PHONE_NUM, CHAR(B.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        " FROM      METHDBA.TTR_ABS_PRTFL_DTL A, METHDBA.TTR_ABS_CALL_LOG B" & _
        "           , METHDBA.TTR_EMPLOYEE C" & _
        " WHERE     A.PRTFL_NUM = ?" & _
        " AND       A.ABS_NUM = B.ABS_NUM" & _
        " AND       B.CALLER_EMP_NUM = C.EMP_NUM" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlPortfolioNumber
        .Update
    End With

    Set GetPortfolioAbsenceCallLog = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    Exit Function
GetPortfolioAbsenceCallLog_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceCallLog"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceCallLog" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetPortfolioAbsenceCallLog" _
                    , trErrdbGetAbsenceDates _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function

'//-----------------------------------------------------------//
'//-----this is obsolete, use GetContactsDueByScope2 instead
'//-----------------------------------------------------------//
Public Function GetContactsDueByScope(ByVal vlTRISUserNumber As Long _
                , ByVal vbActiveInformationOnly As Boolean _
                ) As ADODB.Recordset
TrsRaiseError MODULENAME, "GetContactsDueByScope", trErrObsoleteMethod, , , , True
'On Error GoTo GetContactsDueByScope_EH
'Dim lsSQL1 As String, lsSQL2 As String
'Dim lsSQL3 As String, lsSQL4 As String
'Dim lsSQL5 As String, lsSQL6 As String
'Dim ladcn As ADODB.Connection, ladrsParms As ADODB.Recordset
'
'    Set GetContactsDueByScope = Nothing
'
'    'SQL 1
'    lsSQL1 = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_ABS_EMP" & _
'        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL" & _
'        ", CMPY_CD CHAR(4), DEPT_NUM CHAR(3) NOT NULL DEFAULT '') CCSID EBCDIC" & _
'        " ON COMMIT PRESERVE ROWS;"
'
'    'SQL2
'    lsSQL2 = lsSQL2 & _
'        " INSERT INTO SESSION.GTT_ABS_EMP (EMP_NUM, EMP_NAME" & _
'        ", CMPY_CD, DEPT_NUM)" & _
'        " SELECT DISTINCT B.EMP_NUM, B.EMP_NAME" & _
'        "       , B.CMPY_CD, B.DEPT_NUM" & _
'        " FROM  METHDBA.TTR_EMPLOYEE B, METHDBA.TTR_APPROVAL A" & _
'        " WHERE A.EMP_NUM = ?" & _
'        " AND   A.ACCESS_LEVEL <> 'V'" & _
'        " AND   ((B.DIVISION_NUM = RIGHT(DIGITS(A.GROUP_ID), 3)" & _
'        " AND   A.APPR_LEVEL_IND = 10)" & _
'        " OR    (B.FUNCTION_NUM = RIGHT(DIGITS(A.GROUP_ID), 3)" & _
'        " AND   A.APPR_LEVEL_IND = 20)" & _
'        " OR    (B.DEPT_NUM = A.DEPT_NUM" & _
'        " AND   A.APPR_LEVEL_IND = 30)" & _
'        " OR    (B.EMP_NUM = A.GROUP_ID" & _
'        " AND   A.APPR_LEVEL_IND = 60))"
'    If (vbActiveInformationOnly) Then
'        lsSQL2 = lsSQL2 & _
'            " AND   (B.EMP_STATUS <> 'T'" & _
'            " OR    (B.EMP_STATUS = 'T' " & _
'            " AND   B.STATUS_EFF_DATE > CURRENT DATE - 6 MONTHS))"
'    Else
'        lsSQL2 = lsSQL2 & _
'            " AND   B.EMP_STATUS IN ('A', 'L', 'T', 'I')"
'    End If
'    'NOTE THERE IS SOME BAD DATA IN EMPLOYEE TABLE
'    lsSQL2 = lsSQL2 & _
'            " AND   B.DIVISION_NUM <> ''" & _
'            " AND   B.FUNCTION_NUM <> ''" & _
'            " AND   B.DEPT_NUM <> ''"
'    lsSQL2 = lsSQL2 & _
'        " UNION " & _
'        "SELECT DISTINCT B.EMP_NUM, B.EMP_NAME" & _
'        "       , B.CMPY_CD, B.DEPT_NUM" & _
'        " FROM  METHDBA.TTR_EMPLOYEE B" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_EMP_WG C" & _
'        "       ON  (B.EMP_NUM = C.EMP_NUM" & _
'        "       AND CURRENT DATE BETWEEN C.EFF_START_DATE" & _
'        "       AND C.EFF_END_DATE)" & _
'        "       , METHDBA.TTR_APPROVAL A" & _
'        " WHERE A.EMP_NUM = ?" & _
'        " AND   A.GROUP_ID = C.WG_NUM" & _
'        " AND   A.APPR_LEVEL_IND = 50" & _
'        " AND   A.ACCESS_LEVEL <> 'V'"
'    If (vbActiveInformationOnly) Then
'        lsSQL2 = lsSQL2 & _
'            " AND   (B.EMP_STATUS <> 'T'" & _
'            " OR    (B.EMP_STATUS = 'T' " & _
'            " AND   B.STATUS_EFF_DATE > CURRENT DATE - 6 MONTHS))"
'    Else
'        lsSQL2 = lsSQL2 & _
'            " AND   B.EMP_STATUS IN ('A', 'L', 'T', 'I')"
'    End If
'    lsSQL2 = lsSQL2 & _
'        " GROUP BY B.EMP_NUM, B.EMP_NAME, B.CMPY_CD, B.DEPT_NUM;"
'
'    Set ladrsParms = CreateParameterRS
'    With ladrsParms
'        .AddNew
'        !PARM_ORDER = 1
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlTRISUserNumber
'        .Update
'
'        .AddNew
'        !PARM_ORDER = 2
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlTRISUserNumber
'        .Update
'    End With
'
'    'SQL 3
'    lsSQL3 = " DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_ABS_NUM" & _
'        " (ABS_NUM        INTEGER NOT NULL) CCSID EBCDIC" & _
'        " ON COMMIT PRESERVE ROWS;"
'
'    'SQL 4
'    lsSQL4 = lsSQL4 & _
'        " INSERT INTO SESSION.GTT_ABS_NUM (ABS_NUM)" & _
'        " SELECT    B.ABS_NUM" & _
'        " FROM      SESSION.GTT_ABS_EMP A" & _
'        "           INNER JOIN METHDBA.TTR_ABS_INIT B" & _
'        "           ON  B.EMP_NUM = A.EMP_NUM" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK C" & _
'        "           ON  C.ABS_NUM = B.ABS_NUM" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_RG_GROUP B1" & _
'        "           ON  C.ABS_NUM IS NULL" & _
'        "           AND B1.RG_CD = 300" & _
'        "           AND B1.BUS_TYPE_CD = 1" & _
'        "           AND B1.RG_VALUE_CD = B.EXCEPTION_CD" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_RG_GROUP C1" & _
'        "           ON  C.ABS_NUM IS NULL" & _
'        "           AND B1.RG_CD IS NULL" & _
'        "           AND B.FMLA_REASON_CD <> ''" & _
'        "           AND C1.RG_CD = 301" & _
'        "           AND C1.BUS_TYPE_CD = 1" & _
'        "           AND C1.RG_VALUE_CD = B.EXCEPTION_CD" & _
'        " WHERE     C.ABS_NUM IS NULL" & _
'        " AND       (B1.RG_CD IS NOT NULL" & _
'        "           OR C1.RG_CD IS NOT NULL)" & _
'        " GROUP BY  B.ABS_NUM;"
'
'    'SQL 5
'    lsSQL5 = "SELECT *   FROM (" & _
'        " SELECT    A.ABS_NUM, B.EMP_NUM" & _
'        "           , C.EMP_NAME, B.EXCEPTION_CD" & _
'        "           , D.EXCEPTION_DESC AS ABS_DESC" & _
'        "           , DATE(COALESCE(MAX(E.ABS_CONT_TS), B.ABS_ST_TS)) AS LAST_CONT_DT" & _
'        "           , DAYS(CURRENT DATE) - DAYS(DATE(COALESCE(MAX(E.ABS_CONT_TS), B.ABS_ST_TS))) AS ELAPSED_DAYS" & _
'        "           , CASE (C.CMPY_CD)" & _
'        "           WHEN 'ARG' THEN 'Y'" & _
'        "           WHEN 'CIL' THEN 'Y'" & _
'        "           WHEN 'CIS' THEN 'Y'" & _
'        "           WHEN 'MV1' THEN 'Y'" & _
'        "           WHEN 'MV3' THEN 'Y'" & _
'        "           WHEN 'AED' THEN 'Y'" & _
'        "           WHEN 'CIC' THEN 'Y'" & _
'        "           WHEN 'CIP' THEN 'Y'" & _
'        "           WHEN 'GEN' THEN 'Y'" & _
'        "           WHEN 'GMC' THEN 'Y'" & _
'        "           WHEN 'IHC' THEN 'Y'" & _
'        "           WHEN 'IMS' THEN 'Y'" & _
'        "           WHEN 'IPC' THEN 'Y'" & _
'        "           ELSE 'N' END AS ILLINOIS_CMPY"
'    lsSQL5 = lsSQL5 & _
'        "           , '' AS SL_POLICY_TXT, '' AS SL_POLICY_FROM" & _
'        "           , MAX(E.ABS_CONT_TS) AS ABS_CONT_TS, B.EMP_PHONE" & _
'        "           , C.DEPT_NUM, COALESCE(B.PRVDR_PHN_NBR, '') AS PRVDR_PHN_NUM" & _
'        " FROM      SESSION.GTT_ABS_NUM A" & _
'        "           INNER JOIN METHDBA.TTR_ABS_INIT B" & _
'        "           ON  B.ABS_NUM = A.ABS_NUM" & _
'        "           INNER JOIN SESSION.GTT_ABS_EMP C" & _
'        "           ON  C.EMP_NUM = B.EMP_NUM" & _
'        "           INNER JOIN METHDBA.TTR_EXCEPTION_CDS D" & _
'        "           ON  D.EXCEPTION_CD = B.EXCEPTION_CD" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_ABS_CONTACT E" & _
'        "           ON  E.ABS_NUM = A.ABS_NUM " & _
'        " GROUP BY  A.ABS_NUM, B.EMP_NUM" & _
'        "           , C.EMP_NAME, B.EXCEPTION_CD" & _
'        "           , D.EXCEPTION_DESC, C.CMPY_CD" & _
'        "           , B.ABS_ST_TS, ABS_CONT_TS" & _
'        "           , B.EMP_PHONE, C.DEPT_NUM, B.PRVDR_PHN_NBR) ONE" & _
'        " WHERE     ELAPSED_DAYS > 0" & _
'        " OR        (ELAPSED_DAYS = 0 AND ABS_CONT_TS IS NULL)" & _
'        " ORDER BY ELAPSED_DAYS DESC" & _
'        " FOR FETCH ONLY WITH UR;"
'
'    'SQL 6
'    lsSQL6 = "DROP TABLE SESSION.GTT_ABS_EMP; DROP TABLE SESSION.GTT_ABS_NUM;"
'
'    'Get connection
'    Set ladcn = OpenConnection
'
'    'Declare and Populate Global Temp Employee
'    ExecuteOTHERSQL lsSQL1 _
'        , ladcn _
'        , adCmdText
'    NewExecuteOtherSQL lsSQL2 _
'        , ladrsParms _
'        , ladcn _
'        , adCmdText
'    Set ladrsParms = Nothing
'
'    'Declare and Populate Global Temp Absence
'    ExecuteOTHERSQL lsSQL3 _
'        , ladcn _
'        , adCmdText
'    ExecuteOTHERSQL lsSQL4 _
'        , ladcn _
'        , adCmdText
'
'    'Get Contacts Due
'    Set GetContactsDueByScope = NewExecuteSQLWithConnection(lsSQL5 _
'            , Nothing _
'            , ladcn _
'            , adOpenStatic _
'            , adLockBatchOptimistic _
'            , adCmdText)
'
'    'Drop Global Temps
'    ExecuteOTHERSQL lsSQL6 _
'        , ladcn _
'        , adCmdText
'
'    'Close and Drop connection
'    If Not (ladcn Is Nothing) Then
'        ladcn.Close
'    End If
'    Set ladcn = Nothing
'
'    Exit Function
'GetContactsDueByScope_EH:
'    Set ladrsParms = Nothing
'
'    If Not (ladcn Is Nothing) Then
'        lsSQL6 = "DROP TABLE SESSION.GTT_ABS_EMP; DROP TABLE SESSION.GTT_ABS_NUM;"
'        ExecuteOTHERSQL lsSQL6 _
'            , ladcn _
'            , adCmdText
'    End If
'
'    Set ladcn = Nothing
'
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME _
'                    , "GetContactsDueByScope"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "GetContactsDueByScope" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(lsSQL1 & " " & lsSQL2 & _
'                                " " & lsSQL3 & " " & lsSQL4 & _
'                                " " & lsSQL5 & " " & lsSQL6) _
'                    , True
'            Case Else
'                TrsRaiseError MODULENAME _
'                    , "GetContactsDueByScope" _
'                    , trErrdbGetAbsenceContact _
'                    , , _
'                    , DescArgsFrom(lsSQL1 & " " & lsSQL2 & _
'                                " " & lsSQL3 & " " & lsSQL4 & _
'                                " " & lsSQL5 & " " & lsSQL6) _
'                    , True
'    End Select

End Function
'//---------------------------------------------------------------------------------------------------------//
'//not used, obsolete
'//---------------------------------------------------------------------------------------------------------//
Public Function GetContactsDueByEmployees(ByVal vsEmployeeWhereClause As String _
                ) As ADODB.Recordset
TrsRaiseError MODULENAME, "GetContactsDueByEmployees", trErrObsoleteMethod, , , , True
'On Error GoTo GetContactsDueByEmployees_EH
'Dim lsSQL1 As String, lsSQL2 As String
'Dim lsSQL3 As String, lsSQL4 As String
'Dim ladcn As ADODB.Connection
'
'    Set GetContactsDueByEmployees = Nothing
'
'    If (Trim(vsEmployeeWhereClause) <> "") Then
'        lsSQL1 = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_ABS_EMP" & _
'            " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL" & _
'            ", CMPY_CD CHAR(4), DEPT_NUM CHAR(3) NOT NULL DEFAULT '') CCSID EBCDIC" & _
'            " ON COMMIT PRESERVE ROWS;"
'        lsSQL1 = lsSQL1 & _
'            " INSERT INTO SESSION.GTT_ABS_EMP (EMP_NUM, EMP_NAME, CMPY_CD, DEPT_NUM)" & _
'            " SELECT A.EMP_NUM, A.EMP_NAME, A.CMPY_CD, A.DEPT_NUM FROM METHDBA.TTR_EMPLOYEE A" & _
'            " WHERE A.EMP_NUM IN (" & Trim(vsEmployeeWhereClause) & ")" & _
'            " GROUP BY A.EMP_NUM, A.EMP_NAME, A.CMPY_CD, A.DEPT_NUM;"
'
'        lsSQL2 = lsSQL2 & _
'            " DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_ABS_NUM" & _
'            " (ABS_NUM INTEGER NOT NULL) CCSID EBCDIC" & _
'            " ON COMMIT PRESERVE ROWS;"
'        lsSQL2 = lsSQL2 & _
'            " INSERT INTO SESSION.GTT_ABS_NUM" & _
'            " (ABS_NUM)" & _
'            " SELECT    B.ABS_NUM" & _
'            " FROM      SESSION.GTT_ABS_EMP A" & _
'            "           INNER JOIN METHDBA.TTR_ABS_INIT B" & _
'            "           ON  B.EMP_NUM = A.EMP_NUM" & _
'            "           LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK C" & _
'            "           ON  C.ABS_NUM = B.ABS_NUM" & _
'            "           LEFT OUTER JOIN METHDBA.TTR_RG_GROUP B1" & _
'            "           ON  C.ABS_NUM IS NULL" & _
'            "           AND B1.RG_CD = 300" & _
'            "           AND B1.BUS_TYPE_CD = 1" & _
'            "           AND B1.RG_VALUE_CD = B.EXCEPTION_CD" & _
'            "           LEFT OUTER JOIN METHDBA.TTR_RG_GROUP C1" & _
'            "           ON  C.ABS_NUM IS NULL" & _
'            "           AND B1.RG_CD IS NULL" & _
'            "           AND B.FMLA_REASON_CD <> ''" & _
'            "           AND C1.RG_CD = 301" & _
'            "           AND C1.BUS_TYPE_CD = 1" & _
'            "           AND C1.RG_VALUE_CD = B.EXCEPTION_CD" & _
'            " WHERE     C.ABS_NUM IS NULL" & _
'            " AND       (B1.RG_CD IS NOT NULL" & _
'            "           OR C1.RG_CD IS NOT NULL)" & _
'            " GROUP BY  B.ABS_NUM;"
'
'        lsSQL3 = "SELECT  DISTINCT ABS_NUM, EMP_NUM" & _
'            "       , EMP_NAME, EXCEPTION_CD" & _
'            "       , ABS_DESC, MAX(LAST_CONT_DT) AS LAST_CONT_DT" & _
'            "       , MIN(ELAPSED_DAYS) AS ELAPSED_DAYS, ILLINOIS_CMPY" & _
'            "       , SL_POLICY_TXT, SL_POLICY_FROM" & _
'            "       , MAX(ABS_CONT_TS) AS ABS_CONT_TS, EMP_PHONE" & _
'            "       , DEPT_NUM, PRVDR_PHN_NUM"
'        lsSQL3 = lsSQL3 & _
'            " FROM (SELECT    A.ABS_NUM, B.EMP_NUM" & _
'            "           , C.EMP_NAME, B.EXCEPTION_CD" & _
'            "           , D.EXCEPTION_DESC AS ABS_DESC" & _
'            "           , DATE(COALESCE(MAX(E.ABS_CONT_TS),B.ABS_ST_TS)) AS LAST_CONT_DT" & _
'            "           , DAYS(CURRENT DATE) - DAYS(DATE(COALESCE(MAX(E.ABS_CONT_TS),B.ABS_ST_TS))) AS ELAPSED_DAYS" & _
'            "           , CASE (C.CMPY_CD)" & _
'            "           WHEN 'ARG' THEN 'Y'" & _
'            "           WHEN 'CIL' THEN 'Y'" & _
'            "           WHEN 'CIS' THEN 'Y'" & _
'            "           WHEN 'MV1' THEN 'Y'" & _
'            "           WHEN 'MV3' THEN 'Y'" & _
'            "           WHEN 'AED' THEN 'Y'" & _
'            "           WHEN 'CIC' THEN 'Y'" & _
'            "           WHEN 'CIP' THEN 'Y'" & _
'            "           WHEN 'GEN' THEN 'Y'" & _
'            "           WHEN 'GMC' THEN 'Y'" & _
'            "           WHEN 'IHC' THEN 'Y'" & _
'            "           WHEN 'IMS' THEN 'Y'" & _
'            "           WHEN 'IPC' THEN 'Y'" & _
'            "           ELSE 'N' END AS ILLINOIS_CMPY"
'        lsSQL3 = lsSQL3 & _
'            "           , '' AS SL_POLICY_TXT, '' AS SL_POLICY_FROM" & _
'            "           , MAX(E.ABS_CONT_TS) AS ABS_CONT_TS, B.EMP_PHONE" & _
'            "           , C.DEPT_NUM, COALESCE(B.PRVDR_PHN_NBR, '') AS PRVDR_PHN_NUM" & _
'            " FROM      SESSION.GTT_ABS_NUM A" & _
'            "           INNER JOIN METHDBA.TTR_ABS_INIT B" & _
'            "           ON  B.ABS_NUM = A.ABS_NUM" & _
'            "           INNER JOIN SESSION.GTT_ABS_EMP C" & _
'            "           ON  C.EMP_NUM = B.EMP_NUM" & _
'            "           INNER JOIN METHDBA.TTR_EXCEPTION_CDS D" & _
'            "           ON  D.EXCEPTION_CD = B.EXCEPTION_CD" & _
'            "           LEFT OUTER JOIN METHDBA.TTR_ABS_CONTACT E" & _
'            "           ON  E.ABS_NUM = A.ABS_NUM " & _
'            " GROUP BY  A.ABS_NUM, B.EMP_NUM" & _
'            "           , C.EMP_NAME, B.EXCEPTION_CD" & _
'            "           , D.EXCEPTION_DESC, C.CMPY_CD" & _
'            "           , B.ABS_ST_TS, ABS_CONT_TS" & _
'            "           , B.EMP_PHONE, C.DEPT_NUM, B.PRVDR_PHN_NBR) ONE" & _
'            " WHERE     ELAPSED_DAYS > 0" & _
'            " OR        (ELAPSED_DAYS = 0 AND ABS_CONT_TS IS NULL)"
'        lsSQL3 = lsSQL3 & _
'            " GROUP BY  ABS_NUM" & _
'            "           , EMP_NUM, EMP_NAME" & _
'            "           , EXCEPTION_CD, ABS_DESC" & _
'            "           , ILLINOIS_CMPY, SL_POLICY_TXT" & _
'            "           , SL_POLICY_FROM, EMP_PHONE" & _
'            "           , DEPT_NUM, PRVDR_PHN_NUM" & _
'            " ORDER BY ELAPSED_DAYS DESC" & _
'            " FOR FETCH ONLY WITH UR;"
'
'        lsSQL4 = "DROP TABLE SESSION.GTT_ABS_EMP; DROP TABLE SESSION.GTT_ABS_NUM;"
'
'        'Get Connection
'        Set ladcn = OpenConnection
'
'        'Create and Populate Global Temps
'        ExecuteOTHERSQL lsSQL1 _
'            , ladcn _
'            , adCmdText
'        ExecuteOTHERSQL lsSQL2 _
'            , ladcn _
'            , adCmdText
'
'        'Get Contacts Due
'        Set GetContactsDueByEmployees = NewExecuteSQLWithConnection(lsSQL3 _
'                                        , Nothing _
'                                        , ladcn _
'                                        , adOpenStatic _
'                                        , adLockBatchOptimistic _
'                                        , adCmdText)
'
'        'Drop Global Temps
'        ExecuteOTHERSQL lsSQL4 _
'            , ladcn _
'            , adCmdText
'
'        'Close and Drop connection
'        If Not (ladcn Is Nothing) Then
'            ladcn.Close
'        End If
'        Set ladcn = Nothing
'    End If
'
'    Exit Function
'GetContactsDueByEmployees_EH:
'    If Not (ladcn Is Nothing) Then
'        lsSQL4 = "DROP TABLE SESSION.GTT_ABS_EMP; DROP TABLE SESSION.GTT_ABS_NUM;"
'        ExecuteOTHERSQL lsSQL4 _
'            , ladcn _
'            , adCmdText
'    End If
'
'    Set ladcn = Nothing
'
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME _
'                    , "GetContactsDueByEmployees"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "GetContactsDueByEmployees" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(lsSQL1 & " " & lsSQL2 & " " & lsSQL3 & " " & lsSQL4) _
'                    , True
'            Case Else
'                TrsRaiseError MODULENAME _
'                    , "GetContactsDueByEmployees" _
'                    , trErrdbGetAbsenceContact _
'                    , , _
'                    , DescArgsFrom(lsSQL1 & " " & lsSQL2 & " " & lsSQL3 & " " & lsSQL4) _
'                    , True
'    End Select

End Function
Public Function GetReturnToWorksDueByScope(ByVal vlTRISUserNumber As Long _
                , ByVal vbActiveInformationOnly As Boolean _
                ) As ADODB.Recordset
On Error GoTo GetReturnToWorksDueByScope_EH
Dim lsSQL1 As String, lsSQL2 As String
Dim lsSQL3 As String, lsSQL4 As String
Dim ladcn As ADODB.Connection, ladrsParms As ADODB.Recordset

    Set GetReturnToWorksDueByScope = Nothing

    'SQL 1
    lsSQL1 = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_ABS_EMP" & _
        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL, DEPT_NUM CHAR(3) NOT NULL DEFAULT '') CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"

    lsSQL2 = lsSQL2 & _
        " INSERT INTO SESSION.GTT_ABS_EMP (EMP_NUM, EMP_NAME, DEPT_NUM)" & _
        " SELECT DISTINCT B.EMP_NUM, B.EMP_NAME, B.DEPT_NUM" & _
        " FROM  METHDBA.TTR_EMPLOYEE B, METHDBA.TTR_APPROVAL A" & _
        " WHERE A.EMP_NUM = ?" & _
        " AND   A.ACCESS_LEVEL <> 'V'" & _
        " AND   ((B.DIVISION_NUM = RIGHT(DIGITS(A.GROUP_ID), 3)" & _
        " AND   A.APPR_LEVEL_IND = 10)" & _
        " OR    (B.FUNCTION_NUM = RIGHT(DIGITS(A.GROUP_ID), 3)" & _
        " AND   A.APPR_LEVEL_IND = 20)" & _
        " OR    (B.DEPT_NUM = A.DEPT_NUM" & _
        " AND   A.APPR_LEVEL_IND = 30)" & _
        " OR    (B.EMP_NUM = A.GROUP_ID" & _
        " AND   A.APPR_LEVEL_IND = 60))"
    If (vbActiveInformationOnly) Then
        lsSQL2 = lsSQL2 & _
            " AND   (B.EMP_STATUS <> 'T'" & _
            " OR    (B.EMP_STATUS = 'T' " & _
            " AND   B.STATUS_EFF_DATE > CURRENT DATE - 6 MONTHS))"
    Else
        lsSQL2 = lsSQL2 & _
            " AND   B.EMP_STATUS IN ('A', 'L', 'T', 'I')"
    End If
    'NOTE THERE IS SOME BAD DATA IN EMPLOYEE TABLE
    lsSQL2 = lsSQL2 & _
            " AND   B.DIVISION_NUM <> ''" & _
            " AND   B.FUNCTION_NUM <> ''" & _
            " AND   B.DEPT_NUM <> ''"
    lsSQL2 = lsSQL2 & _
        " UNION " & _
        "SELECT DISTINCT B.EMP_NUM, B.EMP_NAME, B.DEPT_NUM" & _
        " FROM  METHDBA.TTR_EMPLOYEE B" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMP_WG C" & _
        "       ON  (B.EMP_NUM = C.EMP_NUM" & _
        "       AND CURRENT DATE BETWEEN C.EFF_START_DATE" & _
        "       AND C.EFF_END_DATE)" & _
        "       , METHDBA.TTR_APPROVAL A" & _
        " WHERE A.EMP_NUM = ?" & _
        " AND   A.GROUP_ID = C.WG_NUM" & _
        " AND   A.APPR_LEVEL_IND = 50" & _
        " AND   A.ACCESS_LEVEL <> 'V'"
    If (vbActiveInformationOnly) Then
        lsSQL2 = lsSQL2 & _
            " AND   (B.EMP_STATUS <> 'T'" & _
            " OR    (B.EMP_STATUS = 'T' " & _
            " AND   B.STATUS_EFF_DATE > CURRENT DATE - 6 MONTHS))"
    Else
        lsSQL2 = lsSQL2 & _
            " AND   B.EMP_STATUS IN ('A', 'L', 'T', 'I')"
    End If
    lsSQL2 = lsSQL2 & _
        " GROUP BY B.EMP_NUM, B.EMP_NAME, B.DEPT_NUM;"

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
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlTRISUserNumber
        .Update
    End With

    lsSQL3 = "SELECT  B.ABS_NUM, C.EXCEPTION_DESC AS ABS_DESC" & _
        "           , A.EMP_NUM, A.EMP_NAME, A.DEPT_NUM" & _
        "           , DATE(B.RTN_TO_WRK_TS) AS PROB_RTN_WRK_DT, CHAR(TIME(B.RTN_TO_WRK_TS)) AS PROB_RTN_WRK_TM" & _
        "           , DAYS(CURRENT DATE) - DAYS(DATE(B.RTN_TO_WRK_TS)) AS ELAPSED_DAYS" & _
        " FROM  SESSION.GTT_ABS_EMP A" & _
        "       INNER JOIN METHDBA.TTR_ABS_INIT B" & _
        "       ON B.EMP_NUM = A.EMP_NUM" & _
        "       AND B.RTW_NOT_REQD_IND IN ('0', 'N')" & _
        "       INNER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
        "       ON C.EXCEPTION_CD = B.EXCEPTION_CD" & _
        " WHERE NOT EXISTS (SELECT  D.ABS_NUM" & _
        "                   FROM    METHDBA.TTR_ABS_RET_WRK D" & _
        "                   WHERE   D.ABS_NUM = B.ABS_NUM)" & _
        " ORDER BY B.RTN_TO_WRK_TS" & _
        " FOR FETCH ONLY WITH UR;"

    lsSQL4 = "DROP TABLE SESSION.GTT_ABS_EMP;"

    'Get Connection
    Set ladcn = OpenConnection

    'Create and Populate Global Temp Employee
    ExecuteOTHERSQL lsSQL1 _
        , ladcn _
        , adCmdText
    NewExecuteOtherSQL lsSQL2 _
        , ladrsParms _
        , ladcn _
        , adCmdText
    Set ladrsParms = Nothing

    'Get Return to Work
    Set GetReturnToWorksDueByScope = NewExecuteSQLWithConnection(lsSQL3 _
                                        , Nothing _
                                        , ladcn _
                                        , adOpenStatic _
                                        , adLockReadOnly _
                                        , adCmdText)

    'Drop Global Temps
    ExecuteOTHERSQL lsSQL4 _
        , ladcn _
        , adCmdText

    'Close and Drop connection
    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

    Exit Function
GetReturnToWorksDueByScope_EH:
    Set ladrsParms = Nothing

    If Not (ladcn Is Nothing) Then
        lsSQL4 = "DROP TABLE SESSION.GTT_ABS_EMP;"
        ExecuteOTHERSQL lsSQL4 _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorksDueByScope"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorksDueByScope" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL1 & " " & lsSQL2 & _
                                    " " & lsSQL3 & " " & lsSQL4) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorksDueByScope" _
                    , trErrdbGetAbsenceReturnToWork _
                    , , _
                    , DescArgsFrom(lsSQL1 & " " & lsSQL2 & _
                                    " " & lsSQL3 & " " & lsSQL4) _
                    , True
    End Select

End Function
Public Function GetReturnToWorksDueByEmployees(ByVal vsEmployeeWhereClause As String _
                ) As ADODB.Recordset
On Error GoTo GetReturnToWorksDueByEmployees_EH
Dim lsSQL1 As String, lsSQL2 As String
Dim lsSQL3 As String
Dim ladcn As ADODB.Connection

    Set GetReturnToWorksDueByEmployees = Nothing

    If (Trim(vsEmployeeWhereClause) <> "") Then
        lsSQL1 = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_ABS_EMP" & _
            " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL, DEPT_NUM CHAR(3) NOT NULL DEFAULT '') CCSID EBCDIC" & _
            " ON COMMIT PRESERVE ROWS;"
        lsSQL1 = lsSQL1 & _
            " INSERT INTO SESSION.GTT_ABS_EMP (EMP_NUM, EMP_NAME, DEPT_NUM)" & _
            " SELECT A.EMP_NUM, A.EMP_NAME, A.DEPT_NUM FROM METHDBA.TTR_EMPLOYEE A" & _
            " WHERE A.EMP_NUM IN (" & Trim(vsEmployeeWhereClause) & ")" & _
            " GROUP BY A.EMP_NUM, A.EMP_NAME, A.DEPT_NUM;"

        lsSQL2 = "SELECT  B.ABS_NUM, C.EXCEPTION_DESC AS ABS_DESC" & _
            "           , A.EMP_NUM, A.EMP_NAME, A.DEPT_NUM" & _
            "           , DATE(B.RTN_TO_WRK_TS) AS PROB_RTN_WRK_DT, CHAR(TIME(B.RTN_TO_WRK_TS)) AS PROB_RTN_WRK_TM" & _
            "           , DAYS(CURRENT DATE) - DAYS(DATE(B.RTN_TO_WRK_TS)) AS ELAPSED_DAYS" & _
            " FROM  SESSION.GTT_ABS_EMP A" & _
            "       INNER JOIN METHDBA.TTR_ABS_INIT B" & _
            "       AND B.RTW_NOT_REQD_IND IN ('0', 'N')" & _
            "       ON B.EMP_NUM = A.EMP_NUM" & _
            "       INNER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
            "       ON C.EXCEPTION_CD = B.EXCEPTION_CD" & _
            " WHERE NOT EXISTS (SELECT  D.ABS_NUM" & _
            "                   FROM    METHDBA.TTR_ABS_RET_WRK D" & _
            "                   WHERE   D.ABS_NUM = B.ABS_NUM)" & _
            " ORDER BY B.RTN_TO_WRK_TS" & _
            " FOR FETCH ONLY WITH UR;"

        lsSQL3 = "DROP TABLE SESSION.GTT_ABS_EMP;"

        'Get Connection
        Set ladcn = OpenConnection

        'Create and Populate Global Temp
        ExecuteOTHERSQL lsSQL1 _
            , ladcn _
            , adCmdText

        'Get Return to Work Due
        Set GetReturnToWorksDueByEmployees = NewExecuteSQLWithConnection(lsSQL2 _
                                            , Nothing _
                                            , ladcn _
                                            , adOpenStatic _
                                            , adLockReadOnly _
                                            , adCmdText)
        'Drop Global Temp
        ExecuteOTHERSQL lsSQL3 _
            , ladcn _
            , adCmdText

        If Not (ladcn Is Nothing) Then
            ladcn.Close
        End If
        Set ladcn = Nothing
    End If

    Exit Function
GetReturnToWorksDueByEmployees_EH:
    If Not (ladcn Is Nothing) Then
        lsSQL3 = "DROP TABLE SESSION.GTT_ABS_EMP;"
        ExecuteOTHERSQL lsSQL3 _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorksDueByEmployees"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorksDueByEmployees" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorksDueByEmployees" _
                    , trErrdbGetAbsenceReturnToWork _
                    , , _
                    , _
                    , True
    End Select

End Function
Public Function GetFMLAAbsences(ByVal vsEmployeeWhereClause As String _
                ) As ADODB.Recordset
On Error GoTo GetFMLAAbsences_EH
Dim lsSQL As String, llIndex As Long
Dim ladcn As ADODB.Connection

    Set GetFMLAAbsences = Nothing

    If (Trim(vsEmployeeWhereClause) <> "") Then
        lsSQL = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_ABS_EMP" & _
            " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL  ) CCSID EBCDIC" & _
            " ON COMMIT PRESERVE ROWS;"
        lsSQL = lsSQL & _
            " INSERT INTO SESSION.GTT_ABS_EMP (EMP_NUM, EMP_NAME)" & _
            " SELECT A.EMP_NUM, A.EMP_NAME FROM METHDBA.TTR_EMPLOYEE A" & _
            " WHERE A.EMP_NUM IN (" & Trim(vsEmployeeWhereClause) & ")" & _
            " GROUP BY A.EMP_NUM, A.EMP_NAME;"

        Set ladcn = OpenConnection

        ExecuteOTHERSQL lsSQL _
            , ladcn _
            , adCmdText

        lsSQL = "SELECT  B.ABS_NUM, C.EXCEPTION_DESC AS ABS_DESC" & _
            "           , A.EMP_NUM, A.EMP_NAME" & _
            " FROM      SESSION.GTT_ABS_EMP A" & _
            "           INNER JOIN METHDBA.TTR_ABS_INIT B" & _
            "           ON B.EMP_NUM = A.EMP_NUM" & _
            "           INNER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
            "           ON C.EXCEPTION_CD = B.EXCEPTION_CD" & _
            " WHERE     B.FMLA_REASON_CD IN ('BA', 'CC', 'CP', 'CS', 'EI')" & _
            " FOR FETCH ONLY WITH UR;"

        Set GetFMLAAbsences = NewExecuteSQLWithConnection(lsSQL _
                                , Nothing _
                                , ladcn _
                                , adOpenStatic _
                                , adLockReadOnly _
                                , adCmdText)

        lsSQL = "DROP TABLE SESSION.GTT_ABS_EMP;"
        ExecuteOTHERSQL lsSQL, ladcn, adCmdText

        If Not (ladcn Is Nothing) Then
            ladcn.Close
        End If
        Set ladcn = Nothing
    End If

    Exit Function
GetFMLAAbsences_EH:
    If Not (ladcn Is Nothing) Then
        lsSQL = "DROP TABLE SESSION.GTT_ABS_EMP"
        ExecuteOTHERSQL lsSQL, ladcn, adCmdText
    End If

    Set ladcn = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetFMLAAbsences"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetFMLAAbsences" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetFMLAAbsences" _
                    , trErrdbGetEAM _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select


End Function
Public Function GetDailyEmployeeAbsences(ByVal vlTRISUserNumber As Long _
                , ByVal vbActiveInformationOnly As Boolean _
                ) As ADODB.Recordset
TrsRaiseError MODULENAME, "GetDailyEmployeeAbsences", trErrObsoleteMethod, , , , True
'On Error GoTo GetDailyEmployeeAbsences_EH
''NOTE: THIS METHOD LOOKS AT TTR_EMP_SCHED_EXC AND TTR_ACTUAL_TIME TABLE
'Dim lsSQL1 As String, lsSQL2 As String
'Dim lsSQL3 As String, lsSQL4 As String
'Dim ladcn As ADODB.Connection
'Dim ladrsParms As ADODB.Recordset
'
'    Set GetDailyEmployeeAbsences = Nothing
'
'    'SQL 1
'    lsSQL1 = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_ABS_EMP" & _
'        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL" & _
'        ", DEPT_NUM CHAR(3) NOT NULL DEFAULT '') CCSID EBCDIC" & _
'        " ON COMMIT PRESERVE ROWS;"
'
'    'SQL2
'    lsSQL2 = lsSQL2 & _
'        " INSERT INTO SESSION.GTT_ABS_EMP (EMP_NUM, EMP_NAME, DEPT_NUM)" & _
'        " SELECT DISTINCT B.EMP_NUM, B.EMP_NAME, B.DEPT_NUM" & _
'        " FROM  METHDBA.TTR_EMPLOYEE B, METHDBA.TTR_APPROVAL A" & _
'        " WHERE A.EMP_NUM = ?" & _
'        " AND   A.ACCESS_LEVEL <> 'V'" & _
'        " AND   ((B.DIVISION_NUM = RIGHT(DIGITS(A.GROUP_ID), 3)" & _
'        " AND   A.APPR_LEVEL_IND = 10)" & _
'        " OR    (B.FUNCTION_NUM = RIGHT(DIGITS(A.GROUP_ID), 3)" & _
'        " AND   A.APPR_LEVEL_IND = 20)" & _
'        " OR    (B.DEPT_NUM = A.DEPT_NUM" & _
'        " AND   A.APPR_LEVEL_IND = 30)" & _
'        " OR    (B.EMP_NUM = A.GROUP_ID" & _
'        " AND   A.APPR_LEVEL_IND = 60))"
'    If (vbActiveInformationOnly) Then
'        lsSQL2 = lsSQL2 & _
'            " AND   (B.EMP_STATUS <> 'T'" & _
'            " OR    (B.EMP_STATUS = 'T' " & _
'            " AND   B.STATUS_EFF_DATE > CURRENT DATE - 6 MONTHS))"
'    Else
'        lsSQL2 = lsSQL2 & _
'            " AND   B.EMP_STATUS IN ('A', 'L', 'T', 'I')"
'    End If
'    'NOTE THERE IS SOME BAD DATA IN EMPLOYEE TABLE
'    lsSQL2 = lsSQL2 & _
'            " AND   B.DIVISION_NUM <> ''" & _
'            " AND   B.FUNCTION_NUM <> ''" & _
'            " AND   B.DEPT_NUM <> ''"
'    lsSQL2 = lsSQL2 & _
'        " UNION " & _
'        "SELECT DISTINCT B.EMP_NUM, B.EMP_NAME, B.DEPT_NUM" & _
'        " FROM  METHDBA.TTR_EMPLOYEE B" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_EMP_WG C" & _
'        "       ON  (B.EMP_NUM = C.EMP_NUM" & _
'        "       AND CURRENT DATE BETWEEN C.EFF_START_DATE" & _
'        "       AND C.EFF_END_DATE)" & _
'        "       , METHDBA.TTR_APPROVAL A" & _
'        " WHERE A.EMP_NUM = ?" & _
'        " AND   A.GROUP_ID = C.WG_NUM" & _
'        " AND   A.APPR_LEVEL_IND = 50" & _
'        " AND   A.ACCESS_LEVEL <> 'V'"
'    If (vbActiveInformationOnly) Then
'        lsSQL2 = lsSQL2 & _
'            " AND   (B.EMP_STATUS <> 'T'" & _
'            " OR    (B.EMP_STATUS = 'T' " & _
'            " AND   B.STATUS_EFF_DATE > CURRENT DATE - 6 MONTHS))"
'    Else
'        lsSQL2 = lsSQL2 & _
'            " AND   B.EMP_STATUS IN ('A', 'L', 'T', 'I')"
'    End If
'    lsSQL2 = lsSQL2 & _
'        " GROUP BY B.EMP_NUM, B.EMP_NAME, B.DEPT_NUM;"
'
'    lsSQL3 = lsSQL3 & _
'        " SELECT    A.EMP_NUM, A.EMP_NAME" & _
'        "           , COALESCE(B.EXCEPTION_CD, '') AS EXCEPTION_CD, COALESCE(C.EXCEPTION_DESC, '') AS EXCEPTION_DESC" & _
'        "           , COALESCE(B.OVER_EXC_CD, '') AS OVER_EXC_CD, COALESCE(D.EXCEPTION_DESC, '') AS OVER_EXC_DESC" & _
'        "           , B.HOURS, A.DEPT_NUM, COALESCE(E.DEPT_NAME, '') AS DEPT_NAME" & _
'        " FROM      SESSION.GTT_ABS_EMP A" & _
'        "           INNER JOIN METHDBA.TTR_EMP_SCHED_EXC B" & _
'        "           ON A.EMP_NUM = B.EMP_NUM" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
'        "           ON C.EXCEPTION_CD = B.EXCEPTION_CD" & _
'        "           AND C.SCHED_IND = 'Y'" & _
'        "           AND (C.EXCEPTION_BANK_CD <> 'O'" & _
'        "           OR (C.EXCEPTION_CD LIKE 'S%'" & _
'        "           AND C.WORK_REST_IND = 'R'))" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS D" & _
'        "           ON D.EXCEPTION_CD = B.OVER_EXC_CD" & _
'        "           AND D.SCHED_IND = 'Y'" & _
'        "           AND (D.EXCEPTION_BANK_CD <> 'O'" & _
'        "           OR (D.EXCEPTION_CD LIKE 'S%'" & _
'        "           AND D.WORK_REST_IND = 'R'))" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_DEPT E" & _
'        "           ON A.DEPT_NUM = E.DEPT_NUM" & _
'        " WHERE     B.CHARGE_DATE = CURRENT DATE"
'    lsSQL3 = lsSQL3 & _
'        " UNION " & _
'        " SELECT    A.EMP_NUM, A.EMP_NAME" & _
'        "           , COALESCE(B.EXCEPTION_CD, '') AS EXCEPTION_CD, COALESCE(C.EXCEPTION_DESC, '') AS EXCEPTION_DESC" & _
'        "           , COALESCE(B.OVER_EXC_CD, '') AS OVER_EXC_CD, COALESCE(D.EXCEPTION_DESC, '') AS OVER_EXC_DESC" & _
'        "           , B.HOURS, A.DEPT_NUM, COALESCE(E.DEPT_NAME, '') AS DEPT_NAME" & _
'        " FROM      SESSION.GTT_ABS_EMP A" & _
'        "           INNER JOIN METHDBA.TTR_ACTUAL_TIME B" & _
'        "           ON A.EMP_NUM = B.EMP_NUM" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
'        "           ON C.EXCEPTION_CD = B.EXCEPTION_CD" & _
'        "           AND C.SCHED_IND = 'Y'" & _
'        "           AND (C.EXCEPTION_BANK_CD <> 'O'" & _
'        "           OR (C.EXCEPTION_CD LIKE 'S%'" & _
'        "           AND C.WORK_REST_IND = 'R'))" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS D" & _
'        "           ON D.EXCEPTION_CD = B.OVER_EXC_CD" & _
'        "           AND D.SCHED_IND = 'Y'" & _
'        "           AND (D.EXCEPTION_BANK_CD <> 'O'" & _
'        "           OR (D.EXCEPTION_CD LIKE 'S%'" & _
'        "           AND D.WORK_REST_IND = 'R'))" & _
'        "           LEFT OUTER JOIN METHDBA.TTR_DEPT E" & _
'        "           ON A.DEPT_NUM = E.DEPT_NUM" & _
'        " WHERE     B.CHARGE_DATE = CURRENT DATE" & _
'        " FOR FETCH ONLY WITH UR;"
'
'    lsSQL4 = "DROP TABLE SESSION.GTT_ABS_EMP;"
'
'    Set ladrsParms = CreateParameterRS
'    With ladrsParms
'        .AddNew
'        !PARM_ORDER = 1
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlTRISUserNumber
'        .Update
'
'        .AddNew
'        !PARM_ORDER = 2
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlTRISUserNumber
'        .Update
'    End With
'
'    'Get connection
'    Set ladcn = OpenConnection
'
'    'Declare and Populate Global Temp Employee
'    ExecuteOTHERSQL lsSQL1 _
'        , ladcn _
'        , adCmdText
'    NewExecuteOtherSQL lsSQL2 _
'        , ladrsParms _
'        , ladcn _
'        , adCmdText
'    Set ladrsParms = Nothing
'
'    Set GetDailyEmployeeAbsences = NewExecuteSQLWithConnection(lsSQL3 _
'                            , Nothing _
'                            , ladcn _
'                            , adOpenStatic _
'                            , adLockReadOnly _
'                            , adCmdText)
'
'    ExecuteOTHERSQL lsSQL4, ladcn, adCmdText
'
'    If Not (ladcn Is Nothing) Then
'        ladcn.Close
'    End If
'    Set ladcn = Nothing
'
'    Exit Function
'GetDailyEmployeeAbsences_EH:
'    Set ladrsParms = Nothing
'
'    If Not (ladcn Is Nothing) Then
'        lsSQL4 = "DROP TABLE SESSION.GTT_ABS_EMP"
'        ExecuteOTHERSQL lsSQL4, ladcn, adCmdText
'    End If
'
'    Set ladcn = Nothing
'
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME _
'                    , "GetDailyEmployeeAbsences"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "GetDailyEmployeeAbsences" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(lsSQL1 & " " & lsSQL2 & " " & lsSQL3 & " " & lsSQL4) _
'                    , True
'            Case Else
'                TrsRaiseError MODULENAME _
'                    , "GetDailyEmployeeAbsences" _
'                    , trErrdbGetEAM _
'                    , , _
'                    , DescArgsFrom(lsSQL1 & " " & lsSQL2 & " " & lsSQL3 & " " & lsSQL4) _
'                    , True
'    End Select

End Function
Public Function GetListOfProviders() As ADODB.Recordset
On Error GoTo GetListOfProviders_EH
Dim lsSQL As String

    Set GetListOfProviders = Nothing

    lsSQL = "SELECT PRVDR_NBR, PRVDR_NM" & _
        "       , LN1_ADDR, LN2_ADDR" & _
        "       , CTY_NM, ST_CD" & _
        "       , ZIP_CD, PHN_NBR" & _
        "       , CLS_TXT, TYP_CD" & _
        "       , CASE (TYP_CD)" & _
        "       WHEN 'A' THEN 'PHYSICIAN ASSISTANT'" & _
        "       WHEN 'D' THEN 'DOCTOR'" & _
        "       WHEN 'H' THEN 'HOSPITAL'" & _
        "       WHEN 'N' THEN 'NURSE PRACTITIONER'" & _
        "       WHEN 'P' THEN 'PHYSICAL THERAPIST'" & _
        "       WHEN 'R' THEN 'REHAB. FACILITY'" & _
        "       WHEN 'S' THEN 'SPECIALIST'" & _
        "       WHEN 'T' THEN 'DENTIST'" & _
        "       ELSE '' END AS TYP_DESC" & _
        " FROM  METHDBA.TTR_PROVIDER" & _
        " ORDER BY PRVDR_NM" & _
        " FOR FETCH ONLY WITH UR;"

    Set GetListOfProviders = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    Exit Function
GetListOfProviders_EH:
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "IsAValidProvider"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "IsAValidProvider" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "IsAValidProvider" _
                    , trErrdbGetProviderCodeTable _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAValidAbsence(ByVal vlAbsenceNumber As Long _
                ) As Boolean
On Error GoTo IsAValidAbsence_EH
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParm As ADODB.Recordset

    IsAValidAbsence = False

    lsSQL = "SELECT  ABS_NUM FROM METHDBA.TTR_ABS_INIT" & _
        " WHERE ABS_NUM = ?" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParm = CreateParameterRS
    With ladrsParm
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlAbsenceNumber
        .Update
    End With

    Set ladrs = NewExecuteSQL(lsSQL _
                    , ladrsParm _
                    , adOpenForwardOnly _
                    , adLockReadOnly)
    Set ladrsParm = Nothing

    If (RecordsExist(ladrs)) Then
        IsAValidAbsence = True
    End If
    Set ladrs = Nothing

    Exit Function
IsAValidAbsence_EH:
    Set ladrsParm = Nothing
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "IsAValidAbsence"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "IsAValidAbsence" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "IsAValidAbsence" _
                    , trErrdbGetEAM _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAValidAbsenceContact(ByVal vlAbsenceContactNumber As Long _
                , ByVal vlAbsenceNumber As Long _
                ) As Boolean
On Error GoTo IsAValidAbsenceContact_EH
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParm As ADODB.Recordset
    IsAValidAbsenceContact = False

    lsSQL = "SELECT  ABS_CONT_NUM FROM METHDBA.TTR_ABS_CONTACT" & _
        " WHERE ABS_NUM = ?" & _
        " AND   ABS_CONT_NUM = ?" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParm = CreateParameterRS
    With ladrsParm
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlAbsenceNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlAbsenceContactNumber
        .Update
    End With

    Set ladrs = NewExecuteSQL(lsSQL _
                    , ladrsParm _
                    , adOpenForwardOnly _
                    , adLockReadOnly)
    Set ladrsParm = Nothing

    If (RecordsExist(ladrs)) Then
        IsAValidAbsenceContact = True
    End If
    Set ladrs = Nothing

    Exit Function
IsAValidAbsenceContact_EH:
    Set ladrsParm = Nothing
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "IsAValidAbsenceContact"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "IsAValidAbsenceContact" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "IsAValidAbsenceContact" _
                    , trErrdbGetAbsenceContact _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAValidProvider(ByVal vlProviderNumber As Long _
                ) As Boolean
On Error GoTo IsAValidProvider_EH
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParm As ADODB.Recordset

    IsAValidProvider = False

    lsSQL = "SELECT  PRVDR_NBR FROM METHDBA.TTR_PROVIDER" & _
        " WHERE PRVDR_NBR = ?" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParm = CreateParameterRS
    With ladrsParm
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlProviderNumber
        .Update
    End With

    Set ladrs = NewExecuteSQL(lsSQL _
                    , ladrsParm _
                    , adOpenForwardOnly _
                    , adLockReadOnly)
    Set ladrsParm = Nothing

    If (RecordsExist(ladrs)) Then
        IsAValidProvider = True
    End If
    Set ladrs = Nothing

    Exit Function
IsAValidProvider_EH:
    Set ladrsParm = Nothing
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "IsAValidProvider"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "IsAValidProvider" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "IsAValidProvider" _
                    , trErrdbGetProviderCodeTable _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAnAdministrator(ByVal vlTRISUserNumber As Long _
                ) As Boolean
On Error GoTo IsAnAdministrator_EH
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParm As ADODB.Recordset

    IsAnAdministrator = False

    lsSQL = "SELECT     EMP_NUM" & _
        " FROM      METHDBA.TTR_EMPLOYEE" & _
        " WHERE     EMP_NUM = ?" & _
        " AND       SEC_LEVEL IN (9, 10, 11, 981)" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParm = CreateParameterRS
    With ladrsParm
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlTRISUserNumber
        .Update
    End With

    Set ladrs = NewExecuteSQL(lsSQL _
                    , ladrsParm _
                    , adOpenForwardOnly _
                    , adLockReadOnly)
    Set ladrsParm = Nothing

    If (RecordsExist(ladrs)) Then
        IsAnAdministrator = True
    End If
    Set ladrs = Nothing

    Exit Function
IsAnAdministrator_EH:
    Set ladrsParm = Nothing
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "IsAnAdministrator"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "IsAnAdministrator" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "IsAnAdministrator" _
                    , trErrdb_CheckEmployee_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAValidPortfolio(ByVal vlPortfolioNumber As Long _
                ) As Boolean
On Error GoTo IsAValidPortfolio_EH
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParm As ADODB.Recordset

    IsAValidPortfolio = False

    lsSQL = "SELECT PRTFL_NUM" & _
        " FROM  METHDBA.TTR_ABS_PORTFOLIO" & _
        " WHERE PRTFL_NUM = ?" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParm = CreateParameterRS
    With ladrsParm
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlPortfolioNumber
        .Update
    End With

     Set ladrs = NewExecuteSQL(lsSQL _
                    , ladrsParm _
                    , adOpenForwardOnly _
                    , adLockReadOnly)
    Set ladrsParm = Nothing

    If (RecordsExist(ladrs)) Then
        IsAValidPortfolio = True
    End If
    Set ladrs = Nothing

    Exit Function
IsAValidPortfolio_EH:
    Set ladrsParm = Nothing
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "IsAValidPortfolio"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "IsAValidPortfolio" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "IsAValidPortfolio" _
                    , trErrdbGetEAM _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetReturnToWorksDueByScope2(ByVal vlTRISUserNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetReturnToWorksDueByScope2_EH
Dim lsSQL As String
Dim ladrsParms As ADODB.Recordset

    Set GetReturnToWorksDueByScope2 = Nothing

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
    End With

    lsSQL = "SELECT  A.ABS_NUM, C.EXCEPTION_DESC AS ABS_DESC" & _
        "           , B.EMP_NUM, B.EMP_NAME, B.DEPT_NUM" & _
        "           , DATE(A.RTN_TO_WRK_TS) AS PROB_RTN_WRK_DT, CHAR(TIME(A.RTN_TO_WRK_TS)) AS PROB_RTN_WRK_TM" & _
        "           , DAYS(CURRENT DATE) - DAYS(DATE(A.RTN_TO_WRK_TS)) AS ELAPSED_DAYS" & _
        " FROM  METHDBA.TTR_ABS_INIT A" & _
        "       INNER JOIN METHDBA.TTR_DS_EMP_DTL B" & _
        "       ON B.EFF_EMP_NUM = ?" & _
        "       AND A.EMP_NUM = B.EMP_NUM" & _
        "       AND A.RTW_NOT_REQD_IND IN ('0', 'N')" & _
        "       INNER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
        "       ON C.EXCEPTION_CD = A.EXCEPTION_CD" & _
        " WHERE NOT EXISTS (SELECT  D.ABS_NUM" & _
        "                   FROM    METHDBA.TTR_ABS_RET_WRK D" & _
        "                   WHERE   D.ABS_NUM = A.ABS_NUM)" & _
        " ORDER BY A.RTN_TO_WRK_TS" & _
        " FOR FETCH ONLY WITH UR;"

    'Get Return to Work
    Set GetReturnToWorksDueByScope2 = NewExecuteSQL(lsSQL _
                                        , ladrsParms _
                                        , adOpenStatic _
                                        , adLockReadOnly _
                                        , adCmdText)
    Set ladrsParms = Nothing

    Exit Function
GetReturnToWorksDueByScope2_EH:
    Set ladrsParms = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorksDueByScope2"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorksDueByScope2" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetReturnToWorksDueByScope2" _
                    , trErrdbGetAbsenceReturnToWork _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetContactsDueByScope2(ByVal vlTRISUserNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetContactsDueByScope2_EH
Dim lsSQL1 As String, lsSQL2 As String
Dim lsSQL3 As String, lsSQL4 As String
Dim lsSQL5 As String, lsSQL6 As String
Dim ladcn As ADODB.Connection, ladrsParms As ADODB.Recordset

    Set GetContactsDueByScope2 = Nothing

    'SQL 1
    lsSQL1 = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_ABS_EMP" & _
        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL" & _
        ", ILLINOIS_CMPY CHAR(1), DEPT_NUM CHAR(3) NOT NULL DEFAULT '') CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"

    'SQL2
    lsSQL2 = " INSERT INTO SESSION.GTT_ABS_EMP (EMP_NUM, EMP_NAME" & _
        ", ILLINOIS_CMPY, DEPT_NUM)" & _
        " SELECT DISTINCT A.EMP_NUM, A.EMP_NAME" & _
        "       , CASE (B.CMPY_CD)" & _
        "       WHEN 'AIC' THEN 'Y'" & _
        "       WHEN 'ARG' THEN 'Y'" & _
        "       WHEN 'CIL' THEN 'Y'" & _
        "       WHEN 'CIS' THEN 'Y'" & _
        "       WHEN 'MV1' THEN 'Y'" & _
        "       WHEN 'MV3' THEN 'Y'" & _
        "       WHEN 'AED' THEN 'Y'" & _
        "       WHEN 'CIC' THEN 'Y'" & _
        "       WHEN 'CIP' THEN 'Y'" & _
        "       WHEN 'GEN' THEN 'Y'" & _
        "       WHEN 'GMC' THEN 'Y'" & _
        "       WHEN 'IHC' THEN 'Y'" & _
        "       WHEN 'IMS' THEN 'Y'" & _
        "       WHEN 'IPC' THEN 'Y'" & _
        "       ELSE 'N' END AS ILLINOIS_CMPY" & _
        "       , A.DEPT_NUM" & _
        " FROM  METHDBA.TTR_DS_EMP_DTL A" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON A.EFF_EMP_NUM = ?" & _
        "       AND B.EMP_NUM = A.EMP_NUM" & _
        " GROUP BY A.EMP_NUM, A.EMP_NAME, B.CMPY_CD, A.DEPT_NUM;"

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
    End With

    'SQL 3
    lsSQL3 = " DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_ABS_NUM" & _
        " (ABS_NUM        INTEGER NOT NULL) CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"

    'SQL 4
    lsSQL4 = lsSQL4 & _
        " INSERT INTO SESSION.GTT_ABS_NUM (ABS_NUM)" & _
        " SELECT    B.ABS_NUM" & _
        " FROM      SESSION.GTT_ABS_EMP A" & _
        "           INNER JOIN METHDBA.TTR_ABS_INIT B" & _
        "           ON  B.EMP_NUM = A.EMP_NUM" & _
        "           LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK C" & _
        "           ON  C.ABS_NUM = B.ABS_NUM" & _
        "           LEFT OUTER JOIN METHDBA.TTR_RG_GROUP B1" & _
        "           ON  C.ABS_NUM IS NULL AND B.RTW_NOT_REQD_IND IN ('0', 'N')" & _
        "           AND B1.RG_CD = 300" & _
        "           AND B1.BUS_TYPE_CD = 1" & _
        "           AND B1.RG_VALUE_CD = B.EXCEPTION_CD" & _
        "           LEFT OUTER JOIN METHDBA.TTR_RG_GROUP C1" & _
        "           ON  C.ABS_NUM IS NULL AND B.RTW_NOT_REQD_IND IN ('0', 'N')" & _
        "           AND B1.RG_CD IS NULL" & _
        "           AND B.FMLA_REASON_CD <> ''" & _
        "           AND C1.RG_CD = 301" & _
        "           AND C1.BUS_TYPE_CD = 1" & _
        "           AND C1.RG_VALUE_CD = B.EXCEPTION_CD" & _
        " WHERE     C.ABS_NUM IS NULL AND B.RTW_NOT_REQD_IND IN ('0', 'N')" & _
        " AND       (B1.RG_CD IS NOT NULL" & _
        "           OR C1.RG_CD IS NOT NULL)" & _
        " GROUP BY  B.ABS_NUM;"

    'SQL 5
    lsSQL5 = "SELECT    DISTINCT ABS_NUM, EMP_NUM" & _
        "           , EMP_NAME, EXCEPTION_CD" & _
        "           , ABS_DESC, MAX(LAST_CONT_DT) AS LAST_CONT_DT" & _
        "           , MIN(ELAPSED_DAYS) AS ELAPSED_DAYS, ILLINOIS_CMPY" & _
        "           , SL_POLICY_TXT, SL_POLICY_FROM" & _
        "           , MAX(ABS_CONT_TS) AS ABS_CONT_TS, EMP_PHONE" & _
        "           , DEPT_NUM, PRVDR_PHN_NUM" & _
        " FROM      (SELECT     A.ABS_NUM, B.EMP_NUM" & _
        "                       , C.EMP_NAME, B.EXCEPTION_CD" & _
        "                       , D.EXCEPTION_DESC AS ABS_DESC" & _
        "                       , DATE(COALESCE(MAX(E.ABS_CONT_TS), B.ABS_ST_TS)) AS LAST_CONT_DT" & _
        "                       , DAYS(CURRENT DATE) - DAYS(DATE(COALESCE(MAX(E.ABS_CONT_TS), B.ABS_ST_TS))) AS ELAPSED_DAYS" & _
        "                       , C.ILLINOIS_CMPY"
    lsSQL5 = lsSQL5 & _
        "                       , '' AS SL_POLICY_TXT, '' AS SL_POLICY_FROM" & _
        "                       , MAX(E.ABS_CONT_TS) AS ABS_CONT_TS, B.EMP_PHONE" & _
        "                       , C.DEPT_NUM, COALESCE(B.PRVDR_PHN_NBR, '') AS PRVDR_PHN_NUM" & _
        "           FROM        SESSION.GTT_ABS_NUM A" & _
        "                       INNER JOIN METHDBA.TTR_ABS_INIT B" & _
        "                       ON  B.ABS_NUM = A.ABS_NUM" & _
        "                       INNER JOIN SESSION.GTT_ABS_EMP C" & _
        "                       ON  C.EMP_NUM = B.EMP_NUM" & _
        "                       INNER JOIN METHDBA.TTR_EXCEPTION_CDS D" & _
        "                       ON  D.EXCEPTION_CD = B.EXCEPTION_CD" & _
        "                       LEFT OUTER JOIN METHDBA.TTR_ABS_CONTACT E" & _
        "                       ON  E.ABS_NUM = A.ABS_NUM " & _
        "           GROUP BY    A.ABS_NUM, B.EMP_NUM" & _
        "                       , C.EMP_NAME, B.EXCEPTION_CD" & _
        "                       , D.EXCEPTION_DESC, C.ILLINOIS_CMPY" & _
        "                       , B.ABS_ST_TS, ABS_CONT_TS" & _
        "                       , B.EMP_PHONE, C.DEPT_NUM, B.PRVDR_PHN_NBR) ONE" & _
        " WHERE     ELAPSED_DAYS > 0" & _
        " OR        (ELAPSED_DAYS = 0 AND ABS_CONT_TS IS NULL)"
    lsSQL5 = lsSQL5 & _
        " GROUP BY  ABS_NUM" & _
        "           , EMP_NUM, EMP_NAME" & _
        "           , EXCEPTION_CD, ABS_DESC" & _
        "           , ILLINOIS_CMPY, SL_POLICY_TXT" & _
        "           , SL_POLICY_FROM, EMP_PHONE" & _
        "           , DEPT_NUM, PRVDR_PHN_NUM" & _
        " ORDER BY ELAPSED_DAYS DESC" & _
        " FOR FETCH ONLY WITH UR;"

    'SQL 6
    lsSQL6 = "DROP TABLE SESSION.GTT_ABS_EMP; DROP TABLE SESSION.GTT_ABS_NUM;"

    'Get connection
    Set ladcn = OpenConnection

    'Declare and Populate Global Temp Employee
    ExecuteOTHERSQL lsSQL1 _
        , ladcn _
        , adCmdText
    NewExecuteOtherSQL lsSQL2 _
        , ladrsParms _
        , ladcn _
        , adCmdText
    Set ladrsParms = Nothing

    'Declare and Populate Global Temp Absence
    ExecuteOTHERSQL lsSQL3 _
        , ladcn _
        , adCmdText
    ExecuteOTHERSQL lsSQL4 _
        , ladcn _
        , adCmdText

    'Get Contacts Due
    Set GetContactsDueByScope2 = NewExecuteSQLWithConnection(lsSQL5 _
            , Nothing _
            , ladcn _
            , adOpenStatic _
            , adLockBatchOptimistic _
            , adCmdText)

    'Drop Global Temps
    ExecuteOTHERSQL lsSQL6 _
        , ladcn _
        , adCmdText

    'Close and Drop connection
    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

    Exit Function
GetContactsDueByScope2_EH:
    Set ladrsParms = Nothing

    If Not (ladcn Is Nothing) Then
        lsSQL6 = "DROP TABLE SESSION.GTT_ABS_EMP; DROP TABLE SESSION.GTT_ABS_NUM;"
        ExecuteOTHERSQL lsSQL6 _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetContactsDueByScope2"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetContactsDueByScope2" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL1 & " " & lsSQL2 & _
                                " " & lsSQL3 & " " & lsSQL4 & _
                                " " & lsSQL5 & " " & lsSQL6) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetContactsDueByScope2" _
                    , trErrdbGetAbsenceContact _
                    , , _
                    , DescArgsFrom(lsSQL1 & " " & lsSQL2 & _
                                " " & lsSQL3 & " " & lsSQL4 & _
                                " " & lsSQL5 & " " & lsSQL6) _
                    , True
    End Select

End Function
Public Function GetDailyEmployeeAbsences2(ByVal vlTRISUserNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetDailyEmployeeAbsences2_EH
'NOTE: THIS METHOD LOOKS AT TTR_EMP_SCHED_EXC AND TTR_ACTUAL_TIME TABLE
Dim lsSQL As String
Dim ladrsParms As ADODB.Recordset

    Set GetDailyEmployeeAbsences2 = Nothing

    lsSQL = lsSQL & _
        " SELECT    DISTINCT A.EMP_NUM, A.EMP_NAME" & _
        "           , COALESCE(B.EXCEPTION_CD, '') AS EXCEPTION_CD, COALESCE(C.EXCEPTION_DESC, '') AS EXCEPTION_DESC" & _
        "           , COALESCE(B.OVER_EXC_CD, '') AS OVER_EXC_CD, COALESCE(D.EXCEPTION_DESC, '') AS OVER_EXC_DESC" & _
        "           , B.HOURS, A.DEPT_NUM, COALESCE(E.DEPT_NAME, '') AS DEPT_NAME" & _
        " FROM      METHDBA.TTR_DS_EMP_DTL A" & _
        "           INNER JOIN METHDBA.TTR_EMP_SCHED_EXC B" & _
        "           ON A.EMP_NUM = B.EMP_NUM" & _
        "           INNER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
        "           ON C.EXCEPTION_CD = B.EXCEPTION_CD" & _
        "           AND C.SCHED_IND = 'Y'" & _
        "           AND C.EXCEPTION_BANK_CD <> 'O'" & _
        "           AND C.WORK_REST_IND = 'R'" & _
        "           AND C.EXCEPTION_CD <> 'LU'" & _
        "           LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS D" & _
        "           ON D.EXCEPTION_CD = B.OVER_EXC_CD" & _
        "           LEFT OUTER JOIN METHDBA.TTR_DEPT E" & _
        "           ON A.DEPT_NUM = E.DEPT_NUM" & _
        " WHERE     A.EFF_EMP_NUM = ?" & _
        " AND       B.CHARGE_DATE = CURRENT DATE"
    lsSQL = lsSQL & _
        " UNION " & _
        " SELECT    DISTINCT A.EMP_NUM, A.EMP_NAME" & _
        "           , COALESCE(B.EXCEPTION_CD, '') AS EXCEPTION_CD, COALESCE(C.EXCEPTION_DESC, '') AS EXCEPTION_DESC" & _
        "           , COALESCE(B.OVER_EXC_CD, '') AS OVER_EXC_CD, COALESCE(D.EXCEPTION_DESC, '') AS OVER_EXC_DESC" & _
        "           , B.HOURS, A.DEPT_NUM, COALESCE(E.DEPT_NAME, '') AS DEPT_NAME" & _
        " FROM      METHDBA.TTR_DS_EMP_DTL A" & _
        "           INNER JOIN METHDBA.TTR_ACTUAL_TIME B" & _
        "           ON A.EMP_NUM = B.EMP_NUM" & _
        "           INNER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
        "           ON C.EXCEPTION_CD = B.EXCEPTION_CD" & _
        "           AND C.SCHED_IND = 'Y'" & _
        "           AND C.EXCEPTION_BANK_CD <> 'O'" & _
        "           AND C.WORK_REST_IND = 'R'" & _
        "           AND C.EXCEPTION_CD <> 'LU'" & _
        "           LEFT OUTER JOIN METHDBA.TTR_EXCEPTION_CDS D" & _
        "           ON D.EXCEPTION_CD = B.OVER_EXC_CD" & _
        "           LEFT OUTER JOIN METHDBA.TTR_DEPT E" & _
        "           ON A.DEPT_NUM = E.DEPT_NUM"
    lsSQL = lsSQL & _
        " WHERE     A.EFF_EMP_NUM = ?" & _
        " AND       B.CHARGE_DATE = CURRENT DATE" & _
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
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlTRISUserNumber
        .Update
    End With

    Set GetDailyEmployeeAbsences2 = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly _
                            , adCmdText)
    Set ladrsParms = Nothing

    Exit Function
GetDailyEmployeeAbsences2_EH:
    Set ladrsParms = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetDailyEmployeeAbsences2"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetDailyEmployeeAbsences2" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetDailyEmployeeAbsences2" _
                    , trErrdbGetEAM _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'BEGIN PREM TRIS 5.9
'THIS IS A MONSTER QUERY, GOOD KIND, IT BRINGS BACK RESULTS!!!!
Public Function GetAbsenceEmployeeInformation3(ByVal vlEmployeeNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetAbsenceEmployeeInformation3_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

'The following are present so that TRIS 5.8 clients
'may operate without problems until they are upgraded to TRIS 5.9
'Y' AS AUTO_CASE
'' AS AUTO_CASE_TYPE
'' AS AC_MM_DD_IND
'0 AS AUTO_CASE_DURATION

    Set GetAbsenceEmployeeInformation3 = Nothing

    lsSQL = "SELECT CMPY_CD, DEPT_NUM" & _
        "       , PREV_DEPT_NUM, DEPT_EFF_DATE" & _
        "       , EXEC_CONTRACT_IND" & _
        "       , HOME_ADDR, HOME_PHONE_NUM" & _
        "       , LOC_CD, LOC_DESC" & _
        "       , MGMT_BEN_PKG_IND, UNION_BEN_PKG_IND" & _
        "       , UNION_CD, PAYROLL_GROUP_CD, ILLINOIS_CMPY" & _
        "       , LAST_PP_PROC_DATE, TIME_INCREMENT" & _
        "       , SL_POLICY_TXT_1 || SL_POLICY_TXT_2 || SL_POLICY_TXT_3 || SL_POLICY_TXT_4 AS SL_POLICY_TXT" & _
        "       , SL_POLICY_FROM, PS_URL" & _
        "       , FORM457_URL" & _
        "       , SUPERVISOR_NUM, SUPERVISOR_NAME" & _
        "       , 'Y' AS AUTO_CASE, '' AS AUTO_CASE_TYPE" & _
        "       , '' AS AC_MM_DD_IND, 0 AS AUTO_CASE_DURATION" & _
        " FROM  (SELECT CMPY_CD, DEPT_NUM" & _
        "               , PREV_DEPT_NUM, DEPT_EFF_DATE" & _
        "               , EXEC_CONTRACT_IND" & _
        "               , HOME_ADDR, HOME_PHONE_NUM" & _
        "               , LOC_CD, LOC_DESC" & _
        "               , MGMT_BEN_PKG_IND, UNION_BEN_PKG_IND" & _
        "               , UNION_CD, PAYROLL_GROUP_CD, ILLINOIS_CMPY" & _
        "               , LAST_PP_PROC_DATE, TIME_INCREMENT"
    lsSQL = lsSQL & _
        "               , CASE (ILLINOIS_CMPY)" & _
        "               WHEN 'Y' THEN 'Employees on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment or obtaining medication.'" & _
        "               ELSE 'The Sick Leave policy may be discussed with the Employee on first follow-up call and as necessary thereafter.'" & _
        "               END AS SL_POLICY_TXT_1" & _
        "               , CASE (ILLINOIS_CMPY)" & _
        "               WHEN 'Y' THEN ''" & _
        "               ELSE '...Employee on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment, or obtaining medication.'" & _
        "               END AS SL_POLICY_TXT_2" & _
        "               , CASE (ILLINOIS_CMPY)" & _
        "               WHEN 'Y' THEN ''" & _
        "               ELSE '...Special cases which require that an employee on sick leave be away from his/her address of record for prescribed physical activity, therapy, etc. should be discussed with the employee''s supervisor in advance.'" & _
        "               END AS SL_POLICY_TXT_3" & _
        "               , CASE (ILLINOIS_CMPY)" & _
        "               WHEN 'Y' THEN ''" & _
        "               ELSE '...Violations of these requirements will be considered as abuse of sick leave and subject to disciplinary action, up to and including discharge.'" & _
        "               END AS SL_POLICY_TXT_4" & _
        "               , CASE (ILLINOIS_CMPY)" & _
        "               WHEN 'N' THEN 'From letter by D.A. Slagley dated 05/23/1989, Personal Activities While On Sick Leave'" & _
        "               ELSE '' END AS SL_POLICY_FROM" & _
        "               , 'https://hrprod1:7002/' AS PS_URL" & _
        "               , 'https://sunp9.ameren.com:80/idc/groups/public/@public015506/documents/document_other/gob019622.pdf' AS FORM457_URL" & _
        "               , SUPERVISOR_NUM, SUPERVISOR_NAME"
    lsSQL = lsSQL & _
        "       FROM    (SELECT A.CMPY_CD, A.DEPT_NUM" & _
        "                       , A.PREV_DEPT_NUM, A.DEPT_EFF_DATE" & _
        "                       , A.EXEC_CONTRACT_IND" & _
        "                       , RTRIM(LTRIM(F.EMPL_HOME_LINE1_AD)) || ' ' || RTRIM(LTRIM(F.EMPL_HOME_LINE2_AD)) || ', ' || RTRIM(LTRIM(F.CITY_AD)) || ', ' || RTRIM(LTRIM(F.STATE_AD)) || ' - ' || RTRIM(LTRIM(F.ZIP_CODE_AD)) AS HOME_ADDR" & _
        "                       , A.HOME_PHONE_NUM, A.LOC_CD, B.LOC_DESC" & _
        "                       , A.MGMT_BEN_PKG_IND, C.UNION_BEN_PKG_IND" & _
        "                       , A.UNION_CD, D.PAYROLL_GROUP_CD" & _
        "                       , CASE (A.CMPY_CD)" & _
        "                       WHEN 'AIC' THEN 'Y'" & _
        "                       WHEN 'ARG' THEN 'Y'" & _
        "                       WHEN 'CIL' THEN 'Y'" & _
        "                       WHEN 'CIS' THEN 'Y'" & _
        "                       WHEN 'MV1' THEN 'Y'" & _
        "                       WHEN 'MV3' THEN 'Y'" & _
        "                       WHEN 'AED' THEN 'Y'" & _
        "                       WHEN 'CIC' THEN 'Y'" & _
        "                       WHEN 'CIP' THEN 'Y'" & _
        "                       WHEN 'GEN' THEN 'Y'" & _
        "                       WHEN 'GMC' THEN 'Y'" & _
        "                       WHEN 'IHC' THEN 'Y'" & _
        "                       WHEN 'IMS' THEN 'Y'" & _
        "                       WHEN 'IPC' THEN 'Y'" & _
        "                       ELSE 'N' END AS ILLINOIS_CMPY"
    lsSQL = lsSQL & _
        "                       , MAX(E.PAY_PER_END_DATE) AS LAST_PP_PROC_DATE" & _
        "                       ,case    " & _
        "                                  when LOJ6.RG_VALUE_CD is not null then 3 " & _
        "                                  when LOJ5.RG_VALUE_CD is not null then 6 " & _
        "                                  when A.EXEC_CONTRACT_IND = 'E' then 30   " & _
        "                                 else 6 END AS TIME_INCREMENT  " & _
        "                       , INTEGER(F.SUPV_ID) AS SUPERVISOR_NUM, G.EMP_NAME AS SUPERVISOR_NAME"
    lsSQL = lsSQL & _
        "               FROM    METHDBA.TTR_EMPLOYEE A" & _
        "                       INNER JOIN METHDBA.TTR_LOCATIONS B" & _
        "                       ON B.LOC_CD = A.LOC_CD" & _
        "                       INNER JOIN METHDBA.TTR_UNIONS C" & _
        "                       ON C.UNION_CD = A.UNION_CD" & _
        "                       INNER JOIN METHDBA.TTR_DEPT D" & _
        "                       ON D.DEPT_NUM = A.DEPT_NUM" & _
        "                       INNER JOIN METHDBA.TTR_COLL_PERIODS E" & _
        "                       ON E.PAYROLL_GROUP_CD = D.PAYROLL_GROUP_CD" & _
        "                       AND E.PAY_PROCESSED_IND = 'Y'" & _
        "                       AND E.PAY_PROCESS_DATE <= CURRENT DATE" & _
        "                       INNER JOIN METHDBA.TAR_COMN_EMPL_TBL F" & _
        "                       ON F.EMPL_ID = CHAR(A.EMP_NUM)" & _
        "                       AND F.SUPV_ID <> ''" & _
        "                       INNER JOIN METHDBA.TTR_EMPLOYEE G" & _
        "                       ON G.EMP_NUM = INTEGER(F.SUPV_ID)" & _
        "               LEFT OUTER JOIN METHDBA.TTR_RG_GROUP LOJ5 " & _
        "               ON LOJ5.RG_CD = 165 " & _
        "                 AND LOJ5.RG_VALUE_CD = A.DEPT_NUM " & _
        "               LEFT OUTER JOIN METHDBA.TTR_RG_GROUP LOJ6 " & _
        "               ON LOJ6.RG_CD = 166  " & _
        "                 AND LOJ6.RG_VALUE_CD = A.DEPT_NUM " & _
        "               WHERE   A.EMP_NUM = ?"
    lsSQL = lsSQL & _
        "               GROUP BY A.CMPY_CD, A.DEPT_NUM" & _
        "                       , A.PREV_DEPT_NUM, A.DEPT_EFF_DATE" & _
        "                       , A.EXEC_CONTRACT_IND" & _
        "                       , F.EMPL_HOME_LINE1_AD, F.EMPL_HOME_LINE2_AD" & _
        "                       , F.CITY_AD, F.STATE_AD" & _
        "                       , F.ZIP_CODE_AD, A.HOME_PHONE_NUM" & _
        "                       , A.LOC_CD, B.LOC_DESC" & _
        "                       , A.MGMT_BEN_PKG_IND, C.UNION_BEN_PKG_IND" & _
        "                       , A.UNION_CD, D.PAYROLL_GROUP_CD, F.SUPV_ID" & _
        "                       , G.EMP_NAME, LOJ5.RG_VALUE_CD, LOJ6.RG_VALUE_CD) ONE) TWO" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update
    End With

    Set GetAbsenceEmployeeInformation3 = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockBatchOptimistic)

    Set ladrs = Nothing

    Exit Function
GetAbsenceEmployeeInformation3_EH:
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetAbsenceEmployeeInformation3"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetAbsenceEmployeeInformation3" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetAbsenceEmployeeInformation3" _
                    , trErrdbGetEAMEmployee _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'END PREM TRIS 5.9
