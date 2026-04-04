VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdEAMNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   Post TRIS 5.5 Employee Absence Management
'Author:    Prem Chander
'Date:      October 2005
'Description:

'//--------------------------------------------------------------------------------------------------------------------------------------//
'//all but GetEAMNotifyPreference is obsolete and has been moved to Notification or Notification2
'//--------------------------------------------------------------------------------------------------------------------------------------//
Option Explicit

Private Const MODULENAME As String = "dbEAM.cdEAMNotify"
Public Function GetPrimaryServer() As String
On Error GoTo GetPrimaryServer_EH
Dim lsSQL As String, ladrs As ADODB.Recordset

    GetPrimaryServer = ""

    lsSQL = " SELECT    PARM_VALUE" & _
            " FROM      METHDBA.TTR_SYSPARMS" & _
            " WHERE     PARM_NUM = " & _
            " AND       PARM_VALUE = " & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrs = ExecuteSQL(lsSQL _
                    , adOpenForwardOnly _
                    , adLockReadOnly _
                    , adCmdText)

    If (RecordsExist(ladrs)) Then
        GetPrimaryServer = Trim(ladrs!PARM_VALUE)
    End If

    Exit Function
GetPrimaryServer_EH:
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetPrimaryServer"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetPrimaryServer" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetPrimaryServer" _
                , trErrdbGetEAM _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetEAMNotifyPreference(ByVal vsTRISUserNumber As String _
                    ) As ADODB.Recordset
On Error GoTo GetEAMNotifyPreference_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetEAMNotifyPreference = Nothing

    lsSQL = " SELECT    RG_VALUE_CD, SUBSTR(RG_VALUE_CD, 1, 9) AS USER_NUM_AS_STR" & _
            "           , SUBSTR(RG_VALUE_CD, 11, 1) AS NOTIFY_IND, INTEGER(COALESCE(LTRIM(RTRIM(SUBSTR(RG_VALUE_CD, 13, 9))), '0')) AS ELAPSED_DAYS" & _
            "           , 'CNTCT' AS RECORD_TYPE" & _
            " FROM      METHDBA.TTR_RG_GROUP" & _
            " WHERE     RG_CD = 303" & _
            " AND       BUS_TYPE_CD = 40" & _
            " AND       SUBSTR(RG_VALUE_CD, 1, 9) = ?" & _
            " UNION " & _
            " SELECT    RG_VALUE_CD, SUBSTR(RG_VALUE_CD, 1, 9) AS USER_NUM_AS_STR" & _
            "           , SUBSTR(RG_VALUE_CD, 11, 1) AS NOTIFY_IND, INTEGER('0') AS ELAPSED_DAYS" & _
            "           , 'RTW' AS RECORD_TYPE" & _
            " FROM      METHDBA.TTR_RG_GROUP" & _
            " WHERE     RG_CD = 307" & _
            " AND       BUS_TYPE_CD = 41" & _
            " AND       SUBSTR(RG_VALUE_CD, 1, 9) = ?;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 9
        !PARM_VALUE = vsTRISUserNumber
        .Update
        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 9
        !PARM_VALUE = vsTRISUserNumber
        .Update
    End With

    Set GetEAMNotifyPreference = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenStatic _
                    , adLockBatchOptimistic _
                    , adCmdText)
    Set ladrsParms = Nothing

    Exit Function
GetEAMNotifyPreference_EH:
    Set ladrsParms = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetEAMNotifyPreference"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetEAMNotifyPreference" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetEAMNotifyPreference" _
                , trErrdbGetEAM _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetNotifyInformation() As ADODB.Recordset
On Error GoTo GetNotifyInformation_EH
Dim lsSQL As String

    Set GetNotifyInformation = Nothing
'//may be able to remove the group by since the join to TTR_ABS_CONTACT e, should include only 1 row
    lsSQL = "SELECT DISTINCT ABS_NUM, ABS_DESC" & _
        "       , EMP_NUM, EMP_NAME" & _
        "       , ABS_ST_TS, LAST_CONT_DT" & _
        "       , ELAPSED_DAYS, FU_NOTIFY_EMP_NUM, CREATED_EMP_NUM" & _
        "       , ABS_PRIM_CONT_EMP, ABS_NOTIFY_WG_NUM" & _
        " FROM  (SELECT A.ABS_NUM, C.EXCEPTION_DESC AS ABS_DESC" & _
        "               , A.EMP_NUM, B.EMP_NAME" & _
        "               , A.ABS_ST_TS" & _
        "               , DATE(COALESCE(E.ABS_CONT_TS, A.ABS_ST_TS)) AS LAST_CONT_DT" & _
        "               , DAYS(CURRENT DATE) - DAYS(DATE(COALESCE(E.ABS_CONT_TS, A.ABS_ST_TS))) AS ELAPSED_DAYS" & _
        "               , INTEGER(SUBSTR(D.RG_VALUE_CD, 13, 9)) AS FU_ELAPSED_DAYS" & _
        "               , E.ABS_CONT_TS" & _
        "               , INTEGER(SUBSTR(D.RG_VALUE_CD, 1, 9)) AS FU_NOTIFY_EMP_NUM" & _
        "               , A.CREATED_EMP_NUM, A.ABS_PRIM_CONT_EMP" & _
        "               , 0 AS ABS_NOTIFY_WG_NUM"
    lsSQL = lsSQL & _
        "       FROM    METHDBA.TTR_ABS_INIT A" & _
        "               INNER JOIN METHDBA.TTR_RG_GROUP D" & _
        "               ON (INTEGER(SUBSTR(D.RG_VALUE_CD, 1, 9)) = A.CREATED_EMP_NUM" & _
        "               OR INTEGER(SUBSTR(D.RG_VALUE_CD, 1, 9)) = A.ABS_PRIM_CONT_EMP)" & _
        "               AND D.RG_CD = 303" & _
        "               AND D.BUS_TYPE_CD = 40" & _
        "               AND SUBSTR(D.RG_VALUE_CD, 11, 1) = 'Y'" & _
        "               LEFT OUTER JOIN (SELECT ABS_NUM, MAX(ABS_CONT_TS) AS ABS_CONT_TS FROM METHDBA.TTR_ABS_CONTACT GROUP BY ABS_NUM) E" & _
        "               ON E.ABS_NUM = A.ABS_NUM" & _
        "               LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK F" & _
        "               ON F.ABS_NUM = A.ABS_NUM" & _
        "               LEFT OUTER JOIN METHDBA.TTR_RG_GROUP G" & _
        "               ON F.ABS_NUM IS NULL" & _
        "               AND G.RG_CD = 300" & _
        "               AND G.BUS_TYPE_CD = 1" & _
        "               AND G.RG_VALUE_CD = A.EXCEPTION_CD"
    lsSQL = lsSQL & _
        "               LEFT OUTER JOIN METHDBA.TTR_RG_GROUP H" & _
        "               ON F.ABS_NUM IS NULL" & _
        "               AND G.RG_CD IS NULL" & _
        "               AND A.FMLA_REASON_CD <> ''" & _
        "               AND H.RG_CD = 301" & _
        "               AND H.BUS_TYPE_CD = 1" & _
        "               AND H.RG_VALUE_CD = A.EXCEPTION_CD" & _
        "               INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "               ON B.EMP_NUM = A.EMP_NUM" & _
        "               INNER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
        "               ON C.EXCEPTION_CD = A.EXCEPTION_CD"
    lsSQL = lsSQL & _
        "       WHERE   F.ABS_NUM IS NULL" & _
        "       AND     (G.RG_CD IS NOT NULL" & _
        "       OR      H.RG_CD IS NOT NULL)" & _
        "       GROUP BY A.ABS_NUM, C.EXCEPTION_DESC" & _
        "               , A.EMP_NUM, B.EMP_NAME" & _
        "               , A.ABS_ST_TS, D.RG_VALUE_CD" & _
        "               , ABS_CONT_TS, A.CREATED_EMP_NUM" & _
        "               , A.ABS_PRIM_CONT_EMP, A.ABS_NOTIFY_WG_NUM" & _
        "       UNION "
    lsSQL = lsSQL & _
        "       SELECT  A.ABS_NUM, C.EXCEPTION_DESC AS ABS_DESC" & _
        "               , A.EMP_NUM, B.EMP_NAME" & _
        "               , A.ABS_ST_TS" & _
        "               , DATE(COALESCE(E.ABS_CONT_TS, A.ABS_ST_TS)) AS LAST_CONT_DT" & _
        "               , DAYS(CURRENT DATE) - DAYS(DATE(COALESCE(E.ABS_CONT_TS, A.ABS_ST_TS))) AS ELAPSED_DAYS" & _
        "               , 1 AS FU_ELAPSED_DAYS" & _
        "               , E.ABS_CONT_TS" & _
        "               , 0 AS FU_NOTIFY_EMP_NUM" & _
        "               , A.CREATED_EMP_NUM, A.ABS_PRIM_CONT_EMP" & _
        "               , A.ABS_NOTIFY_WG_NUM" & _
        "       FROM    METHDBA.TTR_ABS_INIT A"
    lsSQL = lsSQL & _
        "               LEFT OUTER JOIN (SELECT ABS_NUM, MAX(ABS_CONT_TS) AS ABS_CONT_TS FROM METHDBA.TTR_ABS_CONTACT GROUP BY ABS_NUM) E" & _
        "               ON E.ABS_NUM = A.ABS_NUM" & _
        "               LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK F" & _
        "               ON F.ABS_NUM = A.ABS_NUM" & _
        "               LEFT OUTER JOIN METHDBA.TTR_RG_GROUP G" & _
        "               ON F.ABS_NUM IS NULL" & _
        "               AND G.RG_CD = 300" & _
        "               AND G.BUS_TYPE_CD = 1" & _
        "               AND G.RG_VALUE_CD = A.EXCEPTION_CD" & _
        "               LEFT OUTER JOIN METHDBA.TTR_RG_GROUP H" & _
        "               ON  F.ABS_NUM IS NULL" & _
        "               AND G.RG_CD IS NULL" & _
        "               AND A.FMLA_REASON_CD <> ''" & _
        "               AND H.RG_CD = 301" & _
        "               AND H.BUS_TYPE_CD = 1" & _
        "               AND H.RG_VALUE_CD = A.EXCEPTION_CD" & _
        "               INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "               ON B.EMP_NUM = A.EMP_NUM" & _
        "               INNER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
        "               ON C.EXCEPTION_CD = A.EXCEPTION_CD"
    lsSQL = lsSQL & _
        "       WHERE   A.ABS_NOTIFY_WG_NUM > 0" & _
        "       AND     F.ABS_NUM IS NULL" & _
        "       AND     (G.RG_CD IS NOT NULL" & _
        "       OR      H.RG_CD IS NOT NULL)" & _
        "       GROUP BY    A.ABS_NUM, C.EXCEPTION_DESC" & _
        "                   , A.EMP_NUM, B.EMP_NAME" & _
        "                   , A.ABS_ST_TS, ABS_CONT_TS" & _
        "                   , A.CREATED_EMP_NUM, A.ABS_PRIM_CONT_EMP" & _
        "                   , A.ABS_NOTIFY_WG_NUM) ONE" & _
        " WHERE ELAPSED_DAYS >= FU_ELAPSED_DAYS" & _
        " OR    (ELAPSED_DAYS = 0" & _
        " AND   ABS_CONT_TS IS NULL" & _
        " AND   ((FU_ELAPSED_DAYS = 0" & _
        " AND   ABS_NOTIFY_WG_NUM <= 0)" & _
        " OR    (FU_ELAPSED_DAYS = 1" & _
        " AND   ABS_NOTIFY_WG_NUM > 0)))" & _
        " ORDER BY ELAPSED_DAYS DESC" & _
        " FOR FETCH ONLY WITH UR;"

    Set GetNotifyInformation = ExecuteSQL(lsSQL _
                    , adOpenStatic _
                    , adLockReadOnly _
                    , adCmdText)

    Exit Function
GetNotifyInformation_EH:
   Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetNotifyInformation"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetNotifyInformation" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetNotifyInformation" _
                , trErrdbGetEAM _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetContactsDueByPreference() As ADODB.Recordset
On Error GoTo GetContactsDueByPreference_EH
Dim lsSQL As String

    Set GetContactsDueByPreference = Nothing
'//may be able to remove the group by since the join to TTR_ABS_CONTACT e, should include only 1 row
    lsSQL = "SELECT DISTINCT ABS_NUM, ABS_DESC" & _
        "       , EMP_NUM, EMP_NAME" & _
        "       , ABS_ST_TS, LAST_CONT_DT" & _
        "       , ELAPSED_DAYS, FU_NOTIFY_EMP_NUM, CREATED_EMP_NUM" & _
        "       , ABS_PRIM_CONT_EMP, ABS_NOTIFY_WG_NUM" & _
        " FROM  (SELECT A.ABS_NUM, C.EXCEPTION_DESC AS ABS_DESC" & _
        "               , A.EMP_NUM, B.EMP_NAME" & _
        "               , A.ABS_ST_TS" & _
        "               , DATE(COALESCE(E.ABS_CONT_TS, A.ABS_ST_TS)) AS LAST_CONT_DT" & _
        "               , DAYS(CURRENT DATE) - DAYS(DATE(COALESCE(E.ABS_CONT_TS, A.ABS_ST_TS))) AS ELAPSED_DAYS" & _
        "               , INTEGER(SUBSTR(D.RG_VALUE_CD, 13, 9)) AS FU_ELAPSED_DAYS" & _
        "               , E.ABS_CONT_TS" & _
        "               , INTEGER(SUBSTR(D.RG_VALUE_CD, 1, 9)) AS FU_NOTIFY_EMP_NUM" & _
        "               , A.CREATED_EMP_NUM, A.ABS_PRIM_CONT_EMP" & _
        "               , 0 AS ABS_NOTIFY_WG_NUM"
    lsSQL = lsSQL & _
        "       FROM    METHDBA.TTR_ABS_INIT A" & _
        "               INNER JOIN METHDBA.TTR_RG_GROUP D" & _
        "               ON (INTEGER(SUBSTR(D.RG_VALUE_CD, 1, 9)) = A.CREATED_EMP_NUM" & _
        "               OR INTEGER(SUBSTR(D.RG_VALUE_CD, 1, 9)) = A.ABS_PRIM_CONT_EMP)" & _
        "               AND D.RG_CD = 303" & _
        "               AND D.BUS_TYPE_CD = 40" & _
        "               AND SUBSTR(D.RG_VALUE_CD, 11, 1) = 'Y'" & _
        "               LEFT OUTER JOIN (SELECT ABS_NUM, MAX(ABS_CONT_TS) AS ABS_CONT_TS FROM METHDBA.TTR_ABS_CONTACT GROUP BY ABS_NUM) E" & _
        "               ON E.ABS_NUM = A.ABS_NUM" & _
        "               LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK F" & _
        "               ON F.ABS_NUM = A.ABS_NUM" & _
        "               LEFT OUTER JOIN METHDBA.TTR_RG_GROUP G" & _
        "               ON F.ABS_NUM IS NULL" & _
        "               AND G.RG_CD = 300" & _
        "               AND G.BUS_TYPE_CD = 1" & _
        "               AND G.RG_VALUE_CD = A.EXCEPTION_CD"
    lsSQL = lsSQL & _
        "               LEFT OUTER JOIN METHDBA.TTR_RG_GROUP H" & _
        "               ON F.ABS_NUM IS NULL" & _
        "               AND G.RG_CD IS NULL" & _
        "               AND A.FMLA_REASON_CD <> ''" & _
        "               AND H.RG_CD = 301" & _
        "               AND H.BUS_TYPE_CD = 1" & _
        "               AND H.RG_VALUE_CD = A.EXCEPTION_CD" & _
        "               INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "               ON B.EMP_NUM = A.EMP_NUM" & _
        "               INNER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
        "               ON C.EXCEPTION_CD = A.EXCEPTION_CD"
    lsSQL = lsSQL & _
        "       WHERE   F.ABS_NUM IS NULL" & _
        "       AND     (G.RG_CD IS NOT NULL" & _
        "       OR      H.RG_CD IS NOT NULL)" & _
        "       GROUP BY A.ABS_NUM, C.EXCEPTION_DESC" & _
        "               , A.EMP_NUM, B.EMP_NAME" & _
        "               , A.ABS_ST_TS, D.RG_VALUE_CD" & _
        "               , ABS_CONT_TS, A.CREATED_EMP_NUM" & _
        "               , A.ABS_PRIM_CONT_EMP, A.ABS_NOTIFY_WG_NUM" & _
        "       UNION "
    lsSQL = lsSQL & _
        "       SELECT  A.ABS_NUM, C.EXCEPTION_DESC AS ABS_DESC" & _
        "               , A.EMP_NUM, B.EMP_NAME" & _
        "               , A.ABS_ST_TS" & _
        "               , DATE(COALESCE(E.ABS_CONT_TS, A.ABS_ST_TS)) AS LAST_CONT_DT" & _
        "               , DAYS(CURRENT DATE) - DAYS(DATE(COALESCE(E.ABS_CONT_TS, A.ABS_ST_TS))) AS ELAPSED_DAYS" & _
        "               , 1 AS FU_ELAPSED_DAYS" & _
        "               , E.ABS_CONT_TS" & _
        "               , 0 AS FU_NOTIFY_EMP_NUM" & _
        "               , A.CREATED_EMP_NUM, A.ABS_PRIM_CONT_EMP" & _
        "               , A.ABS_NOTIFY_WG_NUM" & _
        "       FROM    METHDBA.TTR_ABS_INIT A"
    lsSQL = lsSQL & _
        "               LEFT OUTER JOIN (SELECT ABS_NUM, MAX(ABS_CONT_TS) AS ABS_CONT_TS FROM METHDBA.TTR_ABS_CONTACT GROUP BY ABS_NUM) E" & _
        "               ON E.ABS_NUM = A.ABS_NUM" & _
        "               LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK F" & _
        "               ON F.ABS_NUM = A.ABS_NUM" & _
        "               LEFT OUTER JOIN METHDBA.TTR_RG_GROUP G" & _
        "               ON F.ABS_NUM IS NULL" & _
        "               AND G.RG_CD = 300" & _
        "               AND G.BUS_TYPE_CD = 1" & _
        "               AND G.RG_VALUE_CD = A.EXCEPTION_CD" & _
        "               LEFT OUTER JOIN METHDBA.TTR_RG_GROUP H" & _
        "               ON  F.ABS_NUM IS NULL" & _
        "               AND G.RG_CD IS NULL" & _
        "               AND A.FMLA_REASON_CD <> ''" & _
        "               AND H.RG_CD = 301" & _
        "               AND H.BUS_TYPE_CD = 1" & _
        "               AND H.RG_VALUE_CD = A.EXCEPTION_CD" & _
        "               INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "               ON B.EMP_NUM = A.EMP_NUM" & _
        "               INNER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
        "               ON C.EXCEPTION_CD = A.EXCEPTION_CD"
    lsSQL = lsSQL & _
        "       WHERE   A.ABS_NOTIFY_WG_NUM > 0" & _
        "       AND     F.ABS_NUM IS NULL" & _
        "       AND     (G.RG_CD IS NOT NULL" & _
        "       OR      H.RG_CD IS NOT NULL)" & _
        "       GROUP BY    A.ABS_NUM, C.EXCEPTION_DESC" & _
        "                   , A.EMP_NUM, B.EMP_NAME" & _
        "                   , A.ABS_ST_TS, ABS_CONT_TS" & _
        "                   , A.CREATED_EMP_NUM, A.ABS_PRIM_CONT_EMP" & _
        "                   , A.ABS_NOTIFY_WG_NUM) ONE" & _
        " WHERE ELAPSED_DAYS >= FU_ELAPSED_DAYS" & _
        " OR    (ELAPSED_DAYS = 0" & _
        " AND   ABS_CONT_TS IS NULL" & _
        " AND   ((FU_ELAPSED_DAYS = 0" & _
        " AND   ABS_NOTIFY_WG_NUM <= 0)" & _
        " OR    (FU_ELAPSED_DAYS = 1" & _
        " AND   ABS_NOTIFY_WG_NUM > 0)))" & _
        " ORDER BY ELAPSED_DAYS DESC" & _
        " FOR FETCH ONLY WITH UR;"

    Set GetContactsDueByPreference = ExecuteSQL(lsSQL _
                    , adOpenStatic _
                    , adLockReadOnly _
                    , adCmdText)

    Exit Function
GetContactsDueByPreference_EH:
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetContactsDueByPreference"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetContactsDueByPreference" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetContactsDueByPreference" _
                , trErrdbGetEAM _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetContactsDueByFUDate() As ADODB.Recordset
On Error GoTo GetContactsDueByFUDate_EH
Dim lsSQL As String

    Set GetContactsDueByFUDate = Nothing

    lsSQL = "SELECT A.ABS_NUM, B.EXCEPTION_DESC AS ABS_DESC" & _
        "       , A.EMP_NUM, C.EMP_NAME" & _
        "       , A.ABS_ST_TS, A.FOLW_UP_DT" & _
        "       , A.CREATED_EMP_NUM AS FU_NOTIFY_EMP_NUM, A.ABS_PRIM_CONT_EMP" & _
        "       , A.ABS_NOTIFY_WG_NUM" & _
        " FROM  METHDBA.TTR_ABS_INIT A" & _
        "       INNER JOIN METHDBA.TTR_EXCEPTION_CDS B" & _
        "       ON B.EXCEPTION_CD = A.EXCEPTION_CD" & _
        "       AND CURRENT DATE = A.FOLW_UP_DT" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE C" & _
        "       ON C.EMP_NUM = A.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK D" & _
        "       ON A.ABS_NUM = D.ABS_NUM" & _
        " WHERE D.ABS_NUM IS NULL" & _
        " ORDER BY A.ABS_NUM" & _
        " FOR FETCH ONLY WITH UR;"

    Set GetContactsDueByFUDate = ExecuteSQL(lsSQL _
                    , adOpenStatic _
                    , adLockReadOnly _
                    , adCmdText)

    Exit Function
GetContactsDueByFUDate_EH:
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetContactsDueByFUDate"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetContactsDueByFUDate" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetContactsDueByFUDate" _
                , trErrdbGetEAM _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetCurrentRtnToWrksDue() As ADODB.Recordset
On Error GoTo GetCurrentRtnToWrksDue_EH
Dim lsSQL As String

    Set GetCurrentRtnToWrksDue = Nothing

    lsSQL = "SELECT A.ABS_NUM, B.EXCEPTION_DESC AS ABS_DESC" & _
        "       , A.EMP_NUM, C.EMP_NAME" & _
        "       , A.ABS_ST_TS, A.RTN_TO_WRK_TS" & _
        "       , A.CREATED_EMP_NUM AS FU_NOTIFY_EMP_NUM, A.ABS_PRIM_CONT_EMP" & _
        "       , A.ABS_NOTIFY_WG_NUM" & _
        " FROM  METHDBA.TTR_ABS_INIT A" & _
        "       INNER JOIN METHDBA.TTR_EXCEPTION_CDS B" & _
        "       ON B.EXCEPTION_CD = A.EXCEPTION_CD" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE C" & _
        "       ON C.EMP_NUM = A.EMP_NUM" & _
        "       INNER JOIN METHDBA.TTR_RG_GROUP D" & _
        "       ON (INTEGER(SUBSTR(D.RG_VALUE_CD, 1, 9)) = A.CREATED_EMP_NUM" & _
        "       OR INTEGER(SUBSTR(D.RG_VALUE_CD, 1, 9)) = A.ABS_PRIM_CONT_EMP)" & _
        "       AND D.RG_CD = 307" & _
        "       AND D.BUS_TYPE_CD = 41" & _
        "       AND SUBSTR(D.RG_VALUE_CD, 11, 1) = 'Y'" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK E" & _
        "       ON A.ABS_NUM = E.ABS_NUM" & _
        " WHERE DATE(A.RTN_TO_WRK_TS) = CURRENT DATE + 1 DAY" & _
        " AND   E.ABS_NUM IS NULL" & _
        " ORDER BY A.ABS_NUM" & _
        " FOR FETCH ONLY WITH UR;"

    Set GetCurrentRtnToWrksDue = ExecuteSQL(lsSQL _
                    , adOpenStatic _
                    , adLockReadOnly _
                    , adCmdText)

    Exit Function
GetCurrentRtnToWrksDue_EH:
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetCurrentRtnToWrksDue"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetCurrentRtnToWrksDue" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetCurrentRtnToWrksDue" _
                , trErrdbGetEAM _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
''''    lsSQL = " SELECT    * FROM (" & _
''''        " SELECT    A.ABS_NUM, C.EXCEPTION_DESC AS ABS_DESC" & _
''''        "           , A.EMP_NUM, B.EMP_NAME" & _
''''        "           , A.EMP_PHONE, A.ABS_ST_TS" & _
''''        "           , DATE(COALESCE(MAX(E.ABS_CONT_TS), A.ABS_ST_TS)) AS LAST_CONT_DT" & _
''''        "           , DAYS(CURRENT_DATE) - DAYS(DATE(COALESCE(MAX(E.ABS_CONT_TS), A.ABS_ST_TS))) AS ELAPSED_DAYS" & _
''''        "           , INTEGER(SUBSTR(D.RG_VALUE_CD, 13, 9)) AS FU_ELAPSED_DAYS" & _
''''        "           , INTEGER(SUBSTR(D.RG_VALUE_CD, 1, 9)) AS NOTIFY_EMP_NUM" & _
''''        "           , MAX(E.ABS_CONT_TS) AS ABS_CONT_TS" & _
''''        "           , A.CREATED_EMP_NUM, A.ABS_PRIM_CONT_EMP" & _
''''        "           , A.ABS_NOTIFY_WG_NUM"
''''    lsSQL = lsSQL & _
''''        " FROM      METHDBA.TTR_ABS_INIT A" & _
''''        "           INNER JOIN METHDBA.TTR_RG_GROUP D" & _
''''        "           ON (INTEGER(SUBSTR(D.RG_VALUE_CD, 1, 9)) = A.CREATED_EMP_NUM" & _
''''        "           OR INTEGER(SUBSTR(D.RG_VALUE_CD, 1, 9)) = A.ABS_PRIM_CONT_EMP)" & _
''''        "           AND D.RG_CD = 303" & _
''''        "           AND D.BUS_TYPE_CD = 40" & _
''''        "           AND SUBSTR(D.RG_VALUE_CD, 11, 1) = 'Y'" & _
''''        "           LEFT OUTER JOIN METHDBA.TTR_ABS_CONTACT E" & _
''''        "           ON E.ABS_NUM = A.ABS_NUM" & _
''''        "           LEFT OUTER JOIN METHDBA.TTR_ABS_RET_WRK F" & _
''''        "           ON F.ABS_NUM = A.ABS_NUM" & _
''''        "           LEFT OUTER JOIN METHDBA.TTR_RG_GROUP G" & _
''''        "           ON F.ABS_NUM IS NULL" & _
''''        "           AND G.RG_CD = 300" & _
''''        "           AND G.BUS_TYPE_CD = 1" & _
''''        "           AND G.RG_VALUE_CD = A.EXCEPTION_CD"
''''    lsSQL = lsSQL & _
''''        "           LEFT OUTER JOIN METHDBA.TTR_RG_GROUP H" & _
''''        "           ON  F.ABS_NUM IS NULL" & _
''''        "           AND G.RG_CD IS NULL" & _
''''        "           AND A.FMLA_REASON_CD <> ''" & _
''''        "           AND H.RG_CD = 301" & _
''''        "           AND H.BUS_TYPE_CD = 1" & _
''''        "           AND H.RG_VALUE_CD = A.EXCEPTION_CD" & _
''''        "           INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
''''        "           ON B.EMP_NUM = A.EMP_NUM" & _
''''        "           INNER JOIN METHDBA.TTR_EXCEPTION_CDS C" & _
''''        "           ON C.EXCEPTION_CD = A.EXCEPTION_CD"
''''    lsSQL = lsSQL & _
''''        " WHERE     F.ABS_NUM IS NULL" & _
''''        " AND       (G.RG_CD IS NOT NULL" & _
''''        " OR        H.RG_CD IS NOT NULL)" & _
''''        " GROUP BY  A.ABS_NUM, C.EXCEPTION_DESC" & _
''''        "           , A.EMP_NUM, B.EMP_NAME" & _
''''        "           , A.EMP_PHONE, A.ABS_ST_TS" & _
''''        "           , D.RG_VALUE_CD, ABS_CONT_TS" & _
''''        "           , A.CREATED_EMP_NUM, A.ABS_PRIM_CONT_EMP" & _
''''        "           , A.ABS_NOTIFY_WG_NUM) ONE" & _
''''        " WHERE     ELAPSED_DAYS >= FU_ELAPSED_DAYS" & _
''''        " OR        (ELAPSED_DAYS = 0 AND ABS_CONT_TS IS NULL)" & _
''''        " ORDER BY  ELAPSED_DAYS DESC" & _
''''        " FOR FETCH ONLY WITH UR;"
