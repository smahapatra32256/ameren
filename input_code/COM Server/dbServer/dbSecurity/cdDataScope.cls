VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdDataScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Conversion Phase 1
'Author:    Prem M. Chander
'Date:      September 1999 (Firstcut)
'Description: This class module will reside in a Server Package
'under Microsoft Transaction Server (MTS). This module will be called by
'its associated business component. This module will access the database for
'pertinant information
'Notes: Mr. Eric Lynn suggested that I use ALL CAPS FOR MODULE LEVEL CONSTANTS,
'       Call Functions to GetConnectionString, OpenConnection from within the
'       ExecuteSQL function. He is okay with passing object context within module functions
'       however he suggests using GetObjectContext rather than a variable(and parameter).

Option Explicit

'Module Name
Private Const MODULENAME = "dbSecurity.cdDataScope"
Public Function db_Approval(ByVal v_lEmpNum As Long _
                    ) As ADODB.Recordset

    Set db_Approval = Nothing
    GetObjectContext.SetComplete

End Function
Public Function db_Employee(ByVal v_bActiveEmps As Boolean _
                    , ByVal v_iSQLID As Integer _
                    , ByVal v_sWHEREDATA As String _
                    , ByVal v_dtLoadFromDate As Date _
                    , ByVal v_lEmpNum As Long _
                    ) As ADODB.Recordset

    Set db_Employee = Nothing
    GetObjectContext.SetComplete

End Function
Public Function db_WorkGroup(ByVal v_iSQLID As Integer _
                    , ByVal v_sWHEREDATA As String _
                    , ByVal v_dtLoadFromDate As Date _
                    , ByVal v_lEmpNum As Long _
                    ) As ADODB.Recordset

    Set db_WorkGroup = Nothing
    GetObjectContext.SetComplete

End Function
Public Function db_Department(ByVal v_iSQLID As Integer _
                    , ByVal v_sWHEREDATA As String _
                    , ByVal v_lEmpNum As Long _
                    ) As ADODB.Recordset

    Set db_Department = Nothing
    GetObjectContext.SetComplete

End Function
Public Function db_Function(ByVal v_iSQLID As Integer _
                    , ByVal v_sWHEREDATA As String _
                    , ByVal v_lEmpNum As Long _
                    ) As ADODB.Recordset

    Set db_Function = Nothing
    GetObjectContext.SetComplete

End Function
Public Function db_Division(ByVal v_iSQLID As Integer _
                    , ByVal v_sWHEREDATA As String _
                    , ByVal v_lEmpNum As Long _
                    ) As ADODB.Recordset

    Set db_Division = Nothing
    GetObjectContext.SetComplete

End Function
Public Function db_EMPWG(ByVal v_lEmpNum As Long _
                    ) As ADODB.Recordset

    Set db_EMPWG = Nothing
    GetObjectContext.SetComplete

End Function
Public Function dbProcessSave(ByVal vadrsDataAccessEntries As ADODB.Recordset _
                , ByVal vlDataAccessEmployeeNumber As Long _
                , ByVal vlUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo dbProcessSave_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long, lsFilterString As String
'SQL
Dim lsSQL As String

    dbProcessSave = False

    lsSQL = ""

    'Check for existence!
    If Not (vadrsDataAccessEntries Is Nothing) Then
        With vadrsDataAccessEntries
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                'Check for deletes, if present process them first
                lsFilterString = "RECORD_STATUS = 'D'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        If (CLng(!APPR_LEVEL_IND) = 30) Then
                            lsSQL = lsSQL & _
                                " DELETE   FROM METHDBA.TTR_APPROVAL " & _
                                " WHERE     EMP_NUM = " & vlDataAccessEmployeeNumber & _
                                " AND       DEPT_NUM = '" & !KEY_DEPT_NUM & "'"
                        Else
                            lsSQL = lsSQL & _
                                " DELETE   FROM METHDBA.TTR_APPROVAL " & _
                                " WHERE     EMP_NUM = " & vlDataAccessEmployeeNumber & _
                                " AND       GROUP_ID = " & !KEY_GROUP_ID
                        End If

                        lsSQL = lsSQL & _
                                " AND       APPR_LEVEL_IND = " & CLng(!APPR_LEVEL_IND) & _
                                " AND       TIME_STAMP = '" & !APP_TIME_STAMP & "'" & _
                                " AND       LAST_MOD_EMP_NUM = " & !APP_MOD_EMP_NUM & ";"

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone
                'Check for Inserts, if present process them next
                lsFilterString = "RECORD_STATUS = 'I'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " INSERT   INTO METHDBA.TTR_APPROVAL " & _
                                " (EMP_NUM, APPR_LEVEL_IND " & _
                                " , GROUP_ID, LAST_MOD_EMP_NUM " & _
                                " , TIME_STAMP, ACCESS_LEVEL " & _
                                " , DEPT_NUM) " & _
                                " VALUES " & _
                                " (" & vlDataAccessEmployeeNumber & ", " & CLng(!APPR_LEVEL_IND) & _
                                " , " & !GROUP_ID & ", " & vlUserEmployeeNumber & _
                                " , CURRENT TIMESTAMP, '" & Trim(!ACCESS_LEVEL) & "'" & _
                                " , '" & !DEPT_NUM & "');"

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone
                'Check for Inserts, if present process them next
                lsFilterString = "RECORD_STATUS = 'C'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " UPDATE   METHDBA.TTR_APPROVAL " & _
                                " SET       ACCESS_LEVEL = '" & !ACCESS_LEVEL & "'" & _
                                "           , TIME_STAMP = CURRENT TIMESTAMP " & _
                                "           , LAST_MOD_EMP_NUM = " & vlUserEmployeeNumber & _
                                " WHERE     EMP_NUM = " & vlDataAccessEmployeeNumber & _
                                " AND       APPR_LEVEL_IND = " & CLng(!APPR_LEVEL_IND)
                        If (CLng(!APPR_LEVEL_IND) = 30) Then
                            lsSQL = lsSQL & _
                                    " AND       DEPT_NUM = '" & !KEY_DEPT_NUM & "'" & _
                                    " AND       TIME_STAMP = '" & !APP_TIME_STAMP & "';"
                        Else
                            lsSQL = lsSQL & _
                                    " AND       GROUP_ID = " & !KEY_GROUP_ID & _
                                    " AND       TIME_STAMP = '" & !APP_TIME_STAMP & "';"
                        End If

                        .MoveNext
                    Next llRecordIndex
                End If
            End If
        End With
    End If

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
        'close and destroy
    End If

    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessSave = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessSave_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing
    dbProcessSave = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbProcessSave"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbProcessSave" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
'            Case trErrdbInsIntoApproval, trErrdbUpdApproval _
'                , trErrdbDelFromApproval
'                TrsRaiseError MODULENAME _
'                    , "dbProcessSave" _
'                    , Err.Number _
'                    , , , _
'                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbProcessSave" _
                    , trErrdbSaveApproval _
                    , , , _
                    , True
    End Select

End Function
Public Function GetApprovalInformation(ByVal vlEmployee As Long _
                ) As ADODB.Recordset
On Error GoTo GetApprovalInformation_EH
Dim ladrs As ADODB.Recordset
'SQL Statement Container
Dim lsSQL As String

    Set GetApprovalInformation = Nothing

    'Set SQL Statement
    lsSQL = "SELECT     DISTINCT APPR_LEVEL_IND, GROUP_ID" & _
            "           , ACCESS_LEVEL, DEPT_NUM" & _
            " FROM      METHDBA.TTR_APPROVAL" & _
            " WHERE     EMP_NUM = ?" & _
            " ORDER BY  APPR_LEVEL_IND, ACCESS_LEVEL"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployee
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    'Set GetApprovalInformation = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set GetApprovalInformation = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetApprovalInformation_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetApprovalInformation"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetApprovalInformation" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetApprovalInformation" _
                    , trErrdb_Approval_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployees(ByVal vlEmployee As Long _
                , ByVal vbActiveInformationOnly As Boolean _
                ) As ADODB.Recordset
On Error GoTo GetEmployees_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

'" WHERE A.EMP_NUM = " & CStr(vlEmployee)
    Set GetEmployees = Nothing

    lsSQL = "SELECT DISTINCT B.EMP_NUM, B.EMP_NAME" & _
        "       , B.DIVISION_NUM, B.FUNCTION_NUM" & _
        "       , B.DEPT_NUM, B.SEC_LEVEL" & _
        "       , B.LOC_CD, 0 AS WG_NUM" & _
        "       , B.UNION_CD, B.JOB_NUM" & _
        "       , COALESCE(B.PREV_DEPT_NUM, '') AS PREV_DEPT_NUM, COALESCE(B.PRIOR_JOB_NUM, '') AS PRIOR_JOB_NUM" & _
        "       , CASE (COALESCE(B.JOB_EFF_DATE, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.JOB_EFF_DATE, DATE('0100-01-01')) END AS JOB_EFF_DATE" & _
        "       , CASE (COALESCE(B.DEPT_EFF_DATE, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.DEPT_EFF_DATE, DATE('0100-01-01')) END AS DEPT_EFF_DATE" & _
        "       , CASE (COALESCE(B.PRIOR_JOB_EFF_DT, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.PRIOR_JOB_EFF_DT, DATE('0100-01-01')) END AS PRIOR_JOB_EFF_DT" & _
        "       , CHAR(B.TIME_STAMP) AS TIME_STAMP" & _
        "       , CHAR(A.TIME_STAMP) AS APP_TIME_STAMP, A.LAST_MOD_EMP_NUM AS APP_MOD_EMP_NUM"

    lsSQL = lsSQL & _
        "       , CASE (A.APPR_LEVEL_IND)" & _
        "       WHEN 10 THEN 'DIVISION'" & _
        "       WHEN 20 THEN 'FUNCTION'" & _
        "       WHEN 30 THEN 'DEPARTMENT'" & _
        "       WHEN 60 THEN 'EMPLOYEE'" & _
        "       ELSE '' END AS ACCESS_SOURCE, A.APPR_LEVEL_IND" & _
        "       , CASE (A.ACCESS_LEVEL)" & _
        "       WHEN 'A' THEN 'APPROVAL'" & _
        "       WHEN 'E' THEN 'ENTRY'" & _
        "       WHEN 'V' THEN 'VIEW'" & _
        "       ELSE '' END AS APP_OR_ENT, A.ACCESS_LEVEL AS APP_ENT_AUTHORITY" & _
        " FROM  METHDBA.TTR_EMPLOYEE B, METHDBA.TTR_APPROVAL A" & _
        " WHERE A.EMP_NUM = ?" & _
        " AND   ((B.DIVISION_NUM = RIGHT(DIGITS(A.GROUP_ID), 3)" & _
        " AND   A.APPR_LEVEL_IND = 10)" & _
        " OR    (B.FUNCTION_NUM = RIGHT(DIGITS(A.GROUP_ID), 3)" & _
        " AND   A.APPR_LEVEL_IND = 20)" & _
        " OR    (B.DEPT_NUM = A.DEPT_NUM" & _
        " AND   A.APPR_LEVEL_IND = 30)" & _
        " OR    (B.EMP_NUM = A.GROUP_ID" & _
        " AND   A.APPR_LEVEL_IND = 60))"

    If (vbActiveInformationOnly) Then
        lsSQL = lsSQL & _
            " AND   (B.EMP_STATUS <> 'T'" & _
            " OR    (B.EMP_STATUS = 'T' " & _
            " AND   B.STATUS_EFF_DATE > CURRENT DATE - 6 MONTHS))"
    Else
        lsSQL = lsSQL & _
            " AND   B.EMP_STATUS IN ('A', 'L', 'T', 'I')"
    End If

    'NOTE THERE IS SOME BAD DATA IN EMPLOYEE TABLE
    lsSQL = lsSQL & _
            " AND   B.DIVISION_NUM <> ''" & _
            " AND   B.FUNCTION_NUM <> ''" & _
            " AND   B.DEPT_NUM <> ''"

    lsSQL = lsSQL & " UNION " & _
        "SELECT DISTINCT B.EMP_NUM, B.EMP_NAME" & _
        "       , B.DIVISION_NUM, B.FUNCTION_NUM" & _
        "       , B.DEPT_NUM, B.SEC_LEVEL" & _
        "       , B.LOC_CD, C.WG_NUM" & _
        "       , B.UNION_CD, B.JOB_NUM" & _
        "       , COALESCE(B.PREV_DEPT_NUM, '') AS PREV_DEPT_NUM, COALESCE(B.PRIOR_JOB_NUM, '') AS PRIOR_JOB_NUM" & _
        "       , CASE (COALESCE(B.JOB_EFF_DATE, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.JOB_EFF_DATE, DATE('0100-01-01')) END AS JOB_EFF_DATE" & _
        "       , CASE (COALESCE(B.DEPT_EFF_DATE, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.DEPT_EFF_DATE, DATE('0100-01-01')) END AS DEPT_EFF_DATE" & _
        "       , CASE (COALESCE(B.PRIOR_JOB_EFF_DT, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.PRIOR_JOB_EFF_DT, DATE('0100-01-01')) END AS PRIOR_JOB_EFF_DT" & _
        "       , CHAR(B.TIME_STAMP) AS TIME_STAMP" & _
        "       , CHAR(A.TIME_STAMP) AS APP_TIME_STAMP, A.LAST_MOD_EMP_NUM AS APP_MOD_EMP_NUM" & _
        "       , 'WORKGROUP' AS ACCESS_SOURCE, A.APPR_LEVEL_IND" & _
        "       , CASE (A.ACCESS_LEVEL)" & _
        "       WHEN 'A' THEN 'APPROVAL'" & _
        "       WHEN 'E' THEN 'ENTRY'" & _
        "       WHEN 'V' THEN 'VIEW'" & _
        "       ELSE '' END AS APP_OR_ENT, A.ACCESS_LEVEL AS APP_ENT_AUTHORITY"

    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_EMPLOYEE B" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMP_WG C" & _
        "       ON  (B.EMP_NUM = C.EMP_NUM" & _
        "       AND CURRENT DATE BETWEEN C.EFF_START_DATE" & _
        "       AND C.EFF_END_DATE)" & _
        "       , METHDBA.TTR_APPROVAL A" & _
        " WHERE A.EMP_NUM = ?" & _
        " AND   A.GROUP_ID = C.WG_NUM" & _
        " AND   A.APPR_LEVEL_IND = 50"

    If (vbActiveInformationOnly) Then
        lsSQL = lsSQL & _
            " AND   (B.EMP_STATUS <> 'T'" & _
            " OR    (B.EMP_STATUS = 'T' " & _
            " AND   B.STATUS_EFF_DATE > CURRENT DATE - 6 MONTHS))"
    Else
        lsSQL = lsSQL & _
            " AND   B.EMP_STATUS IN ('A', 'L', 'T', 'I')"
    End If

    lsSQL = lsSQL & _
        " ORDER BY EMP_NUM;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployee
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployee
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    'Set GetEmployees = ExecuteSQL(lsSQL _
                        , adOpenStatic _
                        , adLockBatchOptimistic)
    Set GetEmployees = NewExecuteSQL(lsSQL _
                        , ladrs _
                        , adOpenStatic _
                        , adLockBatchOptimistic)

    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetEmployees_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Raise the appropriate error message depending
    'on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetEmployees"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetEmployees" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetEmployees" _
                , trErrdb_Employee_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetAllWorkgroups(ByVal vlEmployee As Long) As ADODB.Recordset
On Error GoTo GetAllWorkgroups_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

'" WHERE D.EMP_NUM = " & CStr(vlEmployee)
'" WHERE B.EMP_NUM = " & CStr(vlEmployee)

    Set GetAllWorkgroups = Nothing

    lsSQL = "SELECT  DISTINCT A.WG_NUM, A.CREW_NUM" & _
        "       , A.WG_DESC, A.DEMAND_IND" & _
        "       , CASE (D.APPR_LEVEL_IND)" & _
        "       WHEN 10 THEN 'DIVISION'" & _
        "       WHEN 20 THEN 'FUNCTION'" & _
        "       WHEN 30 THEN 'DEPARTMENT'" & _
        "       WHEN 50 THEN 'WORKGROUP'" & _
        "       ELSE '' END AS ACCESS_SOURCE, D.APPR_LEVEL_IND" & _
        "       , CASE (D.ACCESS_LEVEL)" & _
        "       WHEN 'A' THEN 'APPROVAL'" & _
        "       WHEN 'E' THEN 'ENTRY'" & _
        "       WHEN 'V' THEN 'VIEW'" & _
        "       ELSE '' END AS APP_OR_ENT, D.ACCESS_LEVEL AS APP_ENT_AUTHORITY" & _
        "       , COALESCE(A.EFF_START_DATE, DATE('0100-01-01')) AS EFF_START_DATE" & _
        "       , COALESCE(A.EFF_END_DATE, DATE('0100-01-01')) AS EFF_END_DATE" & _
        "       , A.WG_SUPERVISOR, COALESCE(B.EMP_NAME, '') AS WG_SUPERVISOR_NAME" & _
        "       , A.PROJECT_ID, C.DEPT_NUM" & _
        "       , C.LOC_CD, COALESCE(E.PLANT_CD, '') AS PLANT_CD" & _
        "       , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        "       , C.HR_LOC, C.DIVISION_NUM AS ASSOC_DIV_NUM" & _
        "       , C.FUNCTION_NUM AS ASSOC_FUNC_NUM, C.DEPT_NUM AS ASSOC_DEPT_NUM" & _
        "       , CHAR(D.TIME_STAMP) AS APP_TIME_STAMP, D.LAST_MOD_EMP_NUM AS APP_MOD_EMP_NUM"
    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_WORK_GROUP A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON A.WG_SUPERVISOR = B.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS C" & _
        "       ON A.WG_NUM = C.WG_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_APPROVAL D" & _
        "       ON ((C.DIVISION_NUM= RIGHT(DIGITS(D.GROUP_ID), 3)" & _
        "       AND D.APPR_LEVEL_IND = 10)" & _
        "       OR (C.FUNCTION_NUM = RIGHT(DIGITS(D.GROUP_ID), 3)" & _
        "       AND D.APPR_LEVEL_IND = 20)" & _
        "       OR (C.DEPT_NUM = D.DEPT_NUM" & _
        "       AND D.APPR_LEVEL_IND = 30)" & _
        "       OR (C.WG_NUM = D.GROUP_ID" & _
        "       AND D.APPR_LEVEL_IND = 50))" & _
        "       LEFT OUTER JOIN METHDBA.TTR_LOCATIONS E" & _
        "       ON E.LOC_CD = C.LOC_CD" & _
        " WHERE D.EMP_NUM = ?" & _
        " AND   A.EFF_END_DATE >= CURRENT DATE - 6 MONTHS"

    lsSQL = lsSQL & " UNION "

    lsSQL = lsSQL & _
        " SELECT  DISTINCT A.WG_NUM, A.CREW_NUM" & _
        "       , A.WG_DESC, A.DEMAND_IND" & _
        "       , 'WORKGROUP' AS ACCESS_SOURCE, INTEGER('50') AS APPR_LEVEL_IND" & _
        "       , CASE (B.PRIMARY_IND)" & _
        "       WHEN 'Y' THEN 'PRIMARY'" & _
        "       ELSE 'SECONDARY' END AS APP_OR_ENT" & _
        "       , CASE (B.PRIMARY_IND)" & _
        "       WHEN 'Y' THEN 'P'" & _
        "       ELSE 'S' END AS APP_ENT_AUTHORITY" & _
        "       , COALESCE(B.EFF_START_DATE, DATE('0100-01-01')) AS EFF_START_DATE" & _
        "       , COALESCE(B.EFF_END_DATE, DATE('0100-01-01')) AS EFF_END_DATE" & _
        "       , A.WG_SUPERVISOR, COALESCE(E.EMP_NAME , '') AS WG_SUPERVISOR_NAME" & _
        "       , A.PROJECT_ID, C.DEPT_NUM" & _
        "       , C.LOC_CD, COALESCE(D.PLANT_CD, '') AS PLANT_CD" & _
        "       , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        "       , C.HR_LOC, C.DIVISION_NUM AS ASSOC_DIV_NUM" & _
        "       , C.FUNCTION_NUM AS ASSOC_FUNC_NUM, C.DEPT_NUM AS ASSOC_DEPT_NUM" & _
        "       , CHAR('0') AS APP_TIME_STAMP, 0 AS APP_MOD_EMP_NUM"
    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_WORK_GROUP A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMP_WG B" & _
        "       ON A.WG_NUM = B.WG_NUM" & _
        "       AND CURRENT DATE BETWEEN B.EFF_START_DATE" & _
        "       AND B.EFF_END_DATE" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS C" & _
        "       ON A.WG_NUM = C.WG_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_LOCATIONS D" & _
        "       ON D.LOC_CD = C.LOC_CD" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE E" & _
        "       ON A.WG_SUPERVISOR = E.EMP_NUM" & _
        " WHERE B.EMP_NUM = ?" & _
        " ORDER BY WG_NUM;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployee
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployee
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    'Set GetAllWorkgroups = ExecuteSQL(lsSQL _
                        , adOpenStatic _
                        , adLockBatchOptimistic)
    Set GetAllWorkgroups = NewExecuteSQL(lsSQL _
                        , ladrs _
                        , adOpenStatic _
                        , adLockBatchOptimistic)

    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetAllWorkgroups_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Raise the appropriate error message depending
    'on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetAllWorkgroups"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetAllWorkgroups" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetAllWorkgroups" _
                , trErrdb_WorkGroup_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetDepartments(ByVal vlEmployee As Long) As ADODB.Recordset
On Error GoTo GetDepartments_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

'" WHERE C.EMP_NUM = " & CStr(vlEmployee)

    Set GetDepartments = Nothing

    lsSQL = "SELECT  DISTINCT A.DEPT_NUM, A.DEPT_NAME" & _
        "       , CASE (C.APPR_LEVEL_IND)" & _
        "       WHEN 10 THEN 'DIVISION'" & _
        "       WHEN 20 THEN 'FUNCTION'" & _
        "       WHEN 30 THEN 'DEPARTMENT'" & _
        "       ELSE '' END AS ACCESS_SOURCE, C.APPR_LEVEL_IND" & _
        "       , CASE (C.ACCESS_LEVEL)" & _
        "       WHEN 'A' THEN 'APPROVAL'" & _
        "       WHEN 'E' THEN 'ENTRY'" & _
        "       WHEN 'V' THEN 'VIEW'" & _
        "       ELSE '' END AS APP_OR_ENT, C.ACCESS_LEVEL AS APP_ENT_AUTHORITY" & _
        "       , B.DIVISION_NUM AS ASSOC_DIV_NUM, B.FUNCTION_NUM AS ASSOC_FUNC_NUM" & _
        "       , CHAR(C.TIME_STAMP) AS APP_TIME_STAMP, C.LAST_MOD_EMP_NUM AS APP_MOD_EMP_NUM"
    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_DEPT A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS B" & _
        "       ON A.DEPT_NUM = B.DEPT_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_APPROVAL C" & _
        "       ON ((B.DIVISION_NUM = RIGHT(DIGITS(C.GROUP_ID), 3)" & _
        "       AND C.APPR_LEVEL_IND = 10)" & _
        "       OR (B.FUNCTION_NUM = RIGHT(DIGITS(C.GROUP_ID), 3)" & _
        "       AND C.APPR_LEVEL_IND = 20)" & _
        "       OR (B.DEPT_NUM = C.DEPT_NUM" & _
        "       AND C.APPR_LEVEL_IND = 30))" & _
        " WHERE C.EMP_NUM = ?" & _
        " ORDER BY A.DEPT_NUM;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployee
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    'Set GetDepartments = ExecuteSQL(lsSQL _
                        , adOpenStatic _
                        , adLockBatchOptimistic)

    Set GetDepartments = NewExecuteSQL(lsSQL _
                        , ladrs _
                        , adOpenStatic _
                        , adLockBatchOptimistic)
    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetDepartments_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Raise the appropriate error message depending
    'on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetDepartments"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetDepartments" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetDepartments" _
                , trErrdb_Department_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetFunctions(ByVal vlEmployee As Long) As ADODB.Recordset
On Error GoTo GetFunctions_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

'" WHERE C.EMP_NUM = " & CStr(vlEmployee)
    Set GetFunctions = Nothing

    lsSQL = "SELECT  DISTINCT A.FUNCTION_NUM, A.FUNCTION_NAME" & _
        "       , CASE (C.APPR_LEVEL_IND)" & _
        "       WHEN 10 THEN 'DIVISION'" & _
        "       WHEN 20 THEN 'FUNCTION'" & _
        "       ELSE '' END AS ACCESS_SOURCE, C.APPR_LEVEL_IND" & _
        "       , CASE (C.ACCESS_LEVEL)" & _
        "       WHEN 'A' THEN 'APPROVAL'" & _
        "       WHEN 'E' THEN 'ENTRY'" & _
        "       WHEN 'V' THEN 'VIEW'" & _
        "       ELSE '' END AS APP_OR_ENT, C.ACCESS_LEVEL AS APP_ENT_AUTHORITY" & _
        "       , B.DIVISION_NUM AS ASSOC_DIV_NUM" & _
        "       , CHAR(C.TIME_STAMP) AS APP_TIME_STAMP, C.LAST_MOD_EMP_NUM AS APP_MOD_EMP_NUM"
    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_FUNCTION A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS B" & _
        "       ON A.FUNCTION_NUM = B.FUNCTION_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_APPROVAL C" & _
        "       ON ((B.DIVISION_NUM = RIGHT(DIGITS(C.GROUP_ID), 3)" & _
        "       AND C.APPR_LEVEL_IND = 10)" & _
        "       OR (B.FUNCTION_NUM = RIGHT(DIGITS(C.GROUP_ID), 3)" & _
        "       AND C.APPR_LEVEL_IND = 20))" & _
        " WHERE C.EMP_NUM = ?" & _
        " ORDER BY A.FUNCTION_NUM;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployee
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    'Set GetFunctions = ExecuteSQL(lsSQL _
                        , adOpenStatic _
                        , adLockBatchOptimistic)
    Set GetFunctions = NewExecuteSQL(lsSQL _
                        , ladrs _
                        , adOpenStatic _
                        , adLockBatchOptimistic)

    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetFunctions_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Raise the appropriate error message depending
    'on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetFunctions"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetFunctions" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetFunctions" _
                , trErrdb_Function_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetDivisions(ByVal vlEmployee As Long) As ADODB.Recordset
On Error GoTo GetDivisions_EH
Dim ladrs As ADODB.Recordset
Dim lsSQL As String

'" WHERE C.EMP_NUM = " & CStr(vlEmployee)

    Set GetDivisions = Nothing

    lsSQL = "SELECT  DISTINCT A.DIVISION_NUM, A.DIVISION_NAME" & _
        "       , 'DIVISION' AS ACCESS_SOURCE, C.APPR_LEVEL_IND" & _
        "       , CASE (C.ACCESS_LEVEL)" & _
        "       WHEN 'A' THEN 'APPROVAL'" & _
        "       WHEN 'E' THEN 'ENTRY'" & _
        "       WHEN 'V' THEN 'VIEW'" & _
        "       ELSE '' END AS APP_OR_ENT, C.ACCESS_LEVEL AS APP_ENT_AUTHORITY" & _
        "       , CHAR(C.TIME_STAMP) AS APP_TIME_STAMP, C.LAST_MOD_EMP_NUM AS APP_MOD_EMP_NUM"
    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_DIVISION A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS B" & _
        "       ON A.DIVISION_NUM = B.DIVISION_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_APPROVAL C" & _
        "       ON B.DIVISION_NUM = RIGHT(DIGITS(C.GROUP_ID), 3)" & _
        "       AND C.APPR_LEVEL_IND = 10" & _
        " WHERE C.EMP_NUM = ?" & _
        " ORDER BY A.DIVISION_NUM;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployee
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    'Set GetDivisions = ExecuteSQL(lsSQL _
                        , adOpenStatic _
                        , adLockBatchOptimistic)
    Set GetDivisions = NewExecuteSQL(lsSQL _
                        , ladrs _
                        , adOpenStatic _
                        , adLockBatchOptimistic)

    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetDivisions_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Raise the appropriate error message depending
    'on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetDivisions"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetDivisions" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetDivisions" _
                , trErrdb_Division_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetEmployeeWorkgroup(ByVal vlEmployee As Long) As ADODB.Recordset
On Error GoTo GetEmployeeWorkgroup_EH
'SQL Statement Container
Dim lsSQL As String

    Set GetEmployeeWorkgroup = Nothing

    'Set SQL Statement
    lsSQL = "SELECT  DISTINCT A.WG_NUM, A.CREW_NUM" & _
        "       , A.WG_DESC, A.DEMAND_IND" & _
        "       , CASE (B.PRIMARY_IND)" & _
        "       WHEN 'Y' THEN 'PRIMARY'" & _
        "       ELSE 'SECONDARY' END AS APP_OR_ENT" & _
        "       , CASE (B.PRIMARY_IND)" & _
        "       WHEN 'Y' THEN 'P'" & _
        "       ELSE 'S' END AS APP_ENT_AUTHORITY, INTEGER('50') AS APPR_LEVEL_IND" & _
        "       , COALESCE(B.EFF_START_DATE, DATE('0100-01-01')) AS EFF_START_DATE" & _
        "       , COALESCE(B.EFF_END_DATE, DATE('0100-01-01')) AS EFF_END_DATE" & _
        "       , A.WG_SUPERVISOR, E.EMP_NAME AS WG_SUPERVISOR_NAME" & _
        "       , A.PROJECT_ID, C.DEPT_NUM" & _
        "       , C.LOC_CD, D.PLANT_CD" & _
        "       , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        "       , C.HR_LOC, 'WORKGROUP' AS ACCESS_SOURCE" & _
        "       , C.DIVISION_NUM AS ASSOC_DIV_NUM, C.FUNCTION_NUM AS ASSOC_FUNC_NUM" & _
        "       , C.DEPT_NUM AS ASSOC_DEPT_NUM" & _
        "       , CHAR('0') AS APP_TIME_STAMP, 0 AS APP_MOD_EMP_NUM"
    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_WORK_GROUP A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMP_WG B" & _
        "       ON A.WG_NUM = B.WG_NUM" & _
        "       AND CURRENT DATE BETWEEN B.EFF_START_DATE" & _
        "       AND B.EFF_END_DATE" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS C" & _
        "       ON A.WG_NUM = C.WG_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_LOCATIONS D" & _
        "       ON D.LOC_CD = C.LOC_CD" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE E" & _
        "       ON A.WG_SUPERVISOR = E.EMP_NUM" & _
        " WHERE B.EMP_NUM = " & CStr(vlEmployee) & _
        " ORDER BY A.WG_NUM;"

    'Call ExecuteSQL and retrieve information
    Set GetEmployeeWorkgroup = ExecuteSQL(lsSQL _
                        , adOpenStatic _
                        , adLockBatchOptimistic)

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetEmployeeWorkgroup_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetEmployeeWorkgroup"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetEmployeeWorkgroup" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL)
        Case Else
            TrsRaiseError MODULENAME _
                    , "GetEmployeeWorkgroup" _
                    , trErrdb_EmpWG_F _
                    , , _
                    , DescArgsFrom("Employee Number = " & CStr(vlEmployee) _
                        , lsSQL) _
                    , True
    End Select

End Function
Public Sub GetAllDSInformation(ByVal vlEmployee As Long _
                , ByVal vbActiveInformationOnly As Boolean _
                , ByRef radrsEmployee As ADODB.Recordset _
                , ByRef radrsWorkgroup As ADODB.Recordset _
                , ByRef radrsDepartment As ADODB.Recordset _
                , ByRef radrsFunction As ADODB.Recordset _
                , ByRef radrsDivision As ADODB.Recordset)
On Error GoTo GetAllDSInformation_EH
Dim ladcn As ADODB.Connection
Dim ladrsParms As ADODB.Recordset
Dim lsSQL1 As String, lsSQL2 As String
Dim lsSQL3 As String, lsSQL4 As String
Dim lsSQL5 As String, lsSQL6 As String
Dim lsSQL7 As String, lsSQL8 As String

    'Set SQL Statement
    lsSQL1 = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_EMP_APP" & _
        " (APPR_LEVEL_IND INTEGER NOT NULL DEFAULT 0, ACCESS_SOURCE VARCHAR(30) NOT NULL DEFAULT ''" & _
        ", APP_ENT_AUTHORITY CHAR(1) NOT NULL DEFAULT '', APP_OR_ENT VARCHAR(30) NOT NULL DEFAULT ''" & _
        ", GROUP_ID INTEGER NOT NULL DEFAULT 0, DEPT_NUM CHAR(3) NOT NULL DEFAULT ''" & _
        ", APP_TIME_STAMP CHAR(26) NOT NULL DEFAULT '', APP_MOD_EMP_NUM INTEGER NOT NULL DEFAULT 0" & _
        ") CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"

    lsSQL2 = "INSERT INTO SESSION.GTT_EMP_APP (APPR_LEVEL_IND, ACCESS_SOURCE" & _
        ", APP_ENT_AUTHORITY, APP_OR_ENT" & _
        ", GROUP_ID, DEPT_NUM" & _
        ", APP_TIME_STAMP, APP_MOD_EMP_NUM)" & _
        " SELECT    DISTINCT APPR_LEVEL_IND" & _
        "           , CASE (APPR_LEVEL_IND)" & _
        "           WHEN 10 THEN 'DIVISION'" & _
        "           WHEN 20 THEN 'FUNCTION'" & _
        "           WHEN 30 THEN 'DEPARTMENT'" & _
        "           WHEN 50 THEN 'WORKGROUP'" & _
        "           WHEN 60 THEN 'EMPLOYEE'" & _
        "           ELSE '' END AS ACCESS_SOURCE" & _
        "           , ACCESS_LEVEL AS APP_ENT_AUTHORITY" & _
        "           , CASE (ACCESS_LEVEL)" & _
        "           WHEN 'A' THEN 'APPROVAL'" & _
        "           WHEN 'E' THEN 'ENTRY'" & _
        "           WHEN 'V' THEN 'VIEW'" & _
        "           ELSE '' END AS APP_OR_ENT" & _
        "           , GROUP_ID, DEPT_NUM" & _
        "           , CHAR(TIME_STAMP) AS APP_TIME_STAMP, LAST_MOD_EMP_NUM AS APP_MOD_EMP_NUM" & _
        " FROM      METHDBA.TTR_APPROVAL" & _
        " WHERE     EMP_NUM = ?"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployee
        .Update
    End With

    lsSQL3 = "SELECT DISTINCT B.EMP_NUM, B.EMP_NAME" & _
        "       , B.DIVISION_NUM, B.FUNCTION_NUM" & _
        "       , B.DEPT_NUM, B.SEC_LEVEL" & _
        "       , B.LOC_CD, 0 AS WG_NUM" & _
        "       , B.UNION_CD, B.JOB_NUM" & _
        "       , COALESCE(B.PREV_DEPT_NUM, '') AS PREV_DEPT_NUM, COALESCE(B.PRIOR_JOB_NUM, '') AS PRIOR_JOB_NUM" & _
        "       , CASE (COALESCE(B.JOB_EFF_DATE, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.JOB_EFF_DATE, DATE('0100-01-01')) END AS JOB_EFF_DATE" & _
        "       , CASE (COALESCE(B.DEPT_EFF_DATE, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.DEPT_EFF_DATE, DATE('0100-01-01')) END AS DEPT_EFF_DATE" & _
        "       , CASE (COALESCE(B.PRIOR_JOB_EFF_DT, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.PRIOR_JOB_EFF_DT, DATE('0100-01-01')) END AS PRIOR_JOB_EFF_DT" & _
        "       , CHAR(B.TIME_STAMP) AS TIME_STAMP" & _
        "       , A.APP_TIME_STAMP, A.APP_MOD_EMP_NUM"
    lsSQL3 = lsSQL3 & _
        "       , A.ACCESS_SOURCE, A.APPR_LEVEL_IND" & _
        "       , A.APP_OR_ENT, A.APP_ENT_AUTHORITY" & _
        " FROM  SESSION.GTT_EMP_APP A, METHDBA.TTR_EMPLOYEE B" & _
        " WHERE ((RIGHT(DIGITS(A.GROUP_ID), 3) = B.DIVISION_NUM" & _
        " AND   A.APPR_LEVEL_IND = 10)" & _
        " OR    (RIGHT(DIGITS(A.GROUP_ID), 3) = B.FUNCTION_NUM" & _
        " AND   A.APPR_LEVEL_IND = 20)" & _
        " OR    (A.DEPT_NUM = B.DEPT_NUM" & _
        " AND   A.APPR_LEVEL_IND = 30)" & _
        " OR    (A.GROUP_ID = B.EMP_NUM" & _
        " AND   A.APPR_LEVEL_IND = 60))"

    If (vbActiveInformationOnly) Then
        lsSQL3 = lsSQL3 & _
            " AND   (B.EMP_STATUS <> 'T'" & _
            " OR    (B.EMP_STATUS = 'T' " & _
            " AND   B.STATUS_EFF_DATE > CURRENT DATE - 6 MONTHS))"
    Else
        lsSQL3 = lsSQL3 & _
            " AND   B.EMP_STATUS IN ('A', 'L', 'T', 'I')"
    End If

    'NOTE THERE IS SOME BAD DATA IN EMPLOYEE TABLE
    lsSQL3 = lsSQL3 & _
            " AND   B.DIVISION_NUM <> ''" & _
            " AND   B.FUNCTION_NUM <> ''" & _
            " AND   B.DEPT_NUM <> ''"

    lsSQL3 = lsSQL3 & " UNION " & _
        "SELECT DISTINCT B.EMP_NUM, B.EMP_NAME" & _
        "       , B.DIVISION_NUM, B.FUNCTION_NUM" & _
        "       , B.DEPT_NUM, B.SEC_LEVEL" & _
        "       , B.LOC_CD, C.WG_NUM" & _
        "       , B.UNION_CD, B.JOB_NUM" & _
        "       , COALESCE(B.PREV_DEPT_NUM, '') AS PREV_DEPT_NUM, COALESCE(B.PRIOR_JOB_NUM, '') AS PRIOR_JOB_NUM" & _
        "       , CASE (COALESCE(B.JOB_EFF_DATE, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.JOB_EFF_DATE, DATE('0100-01-01')) END AS JOB_EFF_DATE" & _
        "       , CASE (COALESCE(B.DEPT_EFF_DATE, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.DEPT_EFF_DATE, DATE('0100-01-01')) END AS DEPT_EFF_DATE" & _
        "       , CASE (COALESCE(B.PRIOR_JOB_EFF_DT, DATE('0100-01-01')))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(B.PRIOR_JOB_EFF_DT, DATE('0100-01-01')) END AS PRIOR_JOB_EFF_DT" & _
        "       , CHAR(B.TIME_STAMP) AS TIME_STAMP" & _
        "       , A.APP_TIME_STAMP, A.APP_MOD_EMP_NUM" & _
        "       , 'WORKGROUP' AS ACCESS_SOURCE, A.APPR_LEVEL_IND" & _
        "       , A.APP_OR_ENT, A.APP_ENT_AUTHORITY"

    lsSQL3 = lsSQL3 & _
        " FROM  SESSION.GTT_EMP_APP A, METHDBA.TTR_EMPLOYEE B" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMP_WG C" & _
        "       ON  (B.EMP_NUM = C.EMP_NUM" & _
        "       AND CURRENT DATE BETWEEN C.EFF_START_DATE" & _
        "       AND C.EFF_END_DATE)" & _
        " WHERE A.GROUP_ID = C.WG_NUM" & _
        " AND   A.APPR_LEVEL_IND = 50"

    If (vbActiveInformationOnly) Then
        lsSQL3 = lsSQL3 & _
            " AND   (B.EMP_STATUS <> 'T'" & _
            " OR    (B.EMP_STATUS = 'T' " & _
            " AND   B.STATUS_EFF_DATE > CURRENT DATE - 6 MONTHS))"
    Else
        lsSQL3 = lsSQL3 & _
            " AND   B.EMP_STATUS IN ('A', 'L', 'T', 'I')"
    End If

    lsSQL3 = lsSQL3 & _
        " ORDER BY EMP_NUM;"

    lsSQL4 = "SELECT  DISTINCT A.WG_NUM, A.CREW_NUM" & _
        "       , A.WG_DESC, A.DEMAND_IND" & _
        "       , D.ACCESS_SOURCE, D.APPR_LEVEL_IND" & _
        "       , D.APP_OR_ENT, D.APP_ENT_AUTHORITY" & _
        "       , COALESCE(A.EFF_START_DATE, DATE('0100-01-01')) AS EFF_START_DATE" & _
        "       , COALESCE(A.EFF_END_DATE, DATE('0100-01-01')) AS EFF_END_DATE" & _
        "       , A.WG_SUPERVISOR, COALESCE(B.EMP_NAME, '') AS WG_SUPERVISOR_NAME" & _
        "       , A.PROJECT_ID, C.DEPT_NUM" & _
        "       , C.LOC_CD, COALESCE(E.PLANT_CD, '') AS PLANT_CD" & _
        "       , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        "       , C.HR_LOC, C.DIVISION_NUM AS ASSOC_DIV_NUM" & _
        "       , C.FUNCTION_NUM AS ASSOC_FUNC_NUM, C.DEPT_NUM AS ASSOC_DEPT_NUM" & _
        "       , D.APP_TIME_STAMP, D.APP_MOD_EMP_NUM"
    lsSQL4 = lsSQL4 & _
        " FROM  SESSION.GTT_EMP_APP D" & _
        "       , METHDBA.TTR_WORK_GROUP A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON A.WG_SUPERVISOR = B.EMP_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS C" & _
        "       ON A.WG_NUM = C.WG_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_LOCATIONS E" & _
        "       ON E.LOC_CD = C.LOC_CD" & _
        " WHERE ((RIGHT(DIGITS(D.GROUP_ID), 3) = C.DIVISION_NUM" & _
        " AND   D.APPR_LEVEL_IND = 10)" & _
        " OR    (RIGHT(DIGITS(D.GROUP_ID), 3) = C.FUNCTION_NUM" & _
        " AND   D.APPR_LEVEL_IND = 20)" & _
        " OR    (D.DEPT_NUM = C.DEPT_NUM" & _
        " AND   D.APPR_LEVEL_IND = 30)" & _
        " OR    (D.GROUP_ID = C.WG_NUM" & _
        " AND   D.APPR_LEVEL_IND = 50))" & _
        " AND   A.EFF_END_DATE >= CURRENT DATE - 6 MONTHS"

    lsSQL4 = lsSQL4 & " UNION "

    lsSQL4 = lsSQL4 & _
        " SELECT  DISTINCT A.WG_NUM, A.CREW_NUM" & _
        "       , A.WG_DESC, A.DEMAND_IND" & _
        "       , 'WORKGROUP' AS ACCESS_SOURCE, INTEGER('50') AS APPR_LEVEL_IND" & _
        "       , CASE (B.PRIMARY_IND)" & _
        "       WHEN 'Y' THEN 'PRIMARY'" & _
        "       ELSE 'SECONDARY' END AS APP_OR_ENT" & _
        "       , CASE (B.PRIMARY_IND)" & _
        "       WHEN 'Y' THEN 'P'" & _
        "       ELSE 'S' END AS APP_ENT_AUTHORITY" & _
        "       , COALESCE(B.EFF_START_DATE, DATE('0100-01-01')) AS EFF_START_DATE" & _
        "       , COALESCE(B.EFF_END_DATE, DATE('0100-01-01')) AS EFF_END_DATE" & _
        "       , A.WG_SUPERVISOR, COALESCE(E.EMP_NAME , '') AS WG_SUPERVISOR_NAME" & _
        "       , A.PROJECT_ID, C.DEPT_NUM" & _
        "       , C.LOC_CD, COALESCE(D.PLANT_CD, '') AS PLANT_CD" & _
        "       , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        "       , C.HR_LOC, C.DIVISION_NUM AS ASSOC_DIV_NUM" & _
        "       , C.FUNCTION_NUM AS ASSOC_FUNC_NUM, C.DEPT_NUM AS ASSOC_DEPT_NUM" & _
        "       , CHAR('0') AS APP_TIME_STAMP, 0 AS APP_MOD_EMP_NUM"
    lsSQL4 = lsSQL4 & _
        " FROM  METHDBA.TTR_WORK_GROUP A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMP_WG B" & _
        "       ON A.WG_NUM = B.WG_NUM" & _
        "       AND CURRENT DATE BETWEEN B.EFF_START_DATE" & _
        "       AND B.EFF_END_DATE" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS C" & _
        "       ON A.WG_NUM = C.WG_NUM" & _
        "       LEFT OUTER JOIN METHDBA.TTR_LOCATIONS D" & _
        "       ON D.LOC_CD = C.LOC_CD" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE E" & _
        "       ON A.WG_SUPERVISOR = E.EMP_NUM" & _
        " WHERE B.EMP_NUM = ?" & _
        " ORDER BY WG_NUM;"

    lsSQL5 = "SELECT  DISTINCT A.DEPT_NUM, A.DEPT_NAME" & _
        "       , C.ACCESS_SOURCE, C.APPR_LEVEL_IND" & _
        "       , C.APP_OR_ENT, C.APP_ENT_AUTHORITY" & _
        "       , B.DIVISION_NUM AS ASSOC_DIV_NUM, B.FUNCTION_NUM AS ASSOC_FUNC_NUM" & _
        "       , C.APP_TIME_STAMP, C.APP_MOD_EMP_NUM"
    lsSQL5 = lsSQL5 & _
        " FROM  SESSION.GTT_EMP_APP C" & _
        "       , METHDBA.TTR_DEPT A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS B" & _
        "       ON A.DEPT_NUM = B.DEPT_NUM" & _
        " WHERE ((RIGHT(DIGITS(C.GROUP_ID), 3) = B.DIVISION_NUM" & _
        " AND   C.APPR_LEVEL_IND = 10)" & _
        " OR    (RIGHT(DIGITS(C.GROUP_ID), 3)  = B.FUNCTION_NUM" & _
        " AND   C.APPR_LEVEL_IND = 20)" & _
        " OR    (C.DEPT_NUM = B.DEPT_NUM" & _
        " AND   C.APPR_LEVEL_IND = 30))" & _
        " ORDER BY A.DEPT_NUM;"

    lsSQL6 = "SELECT  DISTINCT A.FUNCTION_NUM, A.FUNCTION_NAME" & _
        "       , C.ACCESS_SOURCE, C.APPR_LEVEL_IND" & _
        "       , C.APP_OR_ENT, C.APP_ENT_AUTHORITY" & _
        "       , B.DIVISION_NUM AS ASSOC_DIV_NUM" & _
        "       , C.APP_TIME_STAMP, C.APP_MOD_EMP_NUM"
    lsSQL6 = lsSQL6 & _
        " FROM  SESSION.GTT_EMP_APP C" & _
        "       , METHDBA.TTR_FUNCTION A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS B" & _
        "       ON A.FUNCTION_NUM = B.FUNCTION_NUM" & _
        " WHERE ((RIGHT(DIGITS(C.GROUP_ID), 3) = B.DIVISION_NUM" & _
        " AND   C.APPR_LEVEL_IND = 10)" & _
        " OR    (RIGHT(DIGITS(C.GROUP_ID), 3) = B.FUNCTION_NUM" & _
        " AND   C.APPR_LEVEL_IND = 20))" & _
        " ORDER BY A.FUNCTION_NUM;"

    lsSQL7 = "SELECT  DISTINCT A.DIVISION_NUM, A.DIVISION_NAME" & _
        "       , C.ACCESS_SOURCE, C.APPR_LEVEL_IND" & _
        "       , C.APP_OR_ENT, C.APP_ENT_AUTHORITY" & _
        "       , C.APP_TIME_STAMP, C.APP_MOD_EMP_NUM"
    lsSQL7 = lsSQL7 & _
        " FROM  SESSION.GTT_EMP_APP C" & _
        "       , METHDBA.TTR_DIVISION A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS B" & _
        "       ON A.DIVISION_NUM = B.DIVISION_NUM" & _
        " WHERE RIGHT(DIGITS(C.GROUP_ID), 3) = B.DIVISION_NUM" & _
        " AND   C.APPR_LEVEL_IND = 10" & _
        " ORDER BY A.DIVISION_NUM;"

    lsSQL8 = "DROP TABLE SESSION.GTT_EMP_APP;"

    'Get connection
    Set ladcn = OpenConnection

    ExecuteOTHERSQL lsSQL1 _
        , ladcn _
        , adCmdText
    lsSQL1 = ""

    NewExecuteOtherSQL lsSQL2 _
        , ladrsParms _
        , ladcn _
        , adCmdText
    lsSQL2 = ""

    Set radrsEmployee = NewExecuteSQLWithConnection(lsSQL3 _
                        , Nothing _
                        , ladcn _
                        , adOpenStatic _
                        , adLockBatchOptimistic _
                        , adCmdText)
    lsSQL3 = ""

    Set radrsWorkgroup = NewExecuteSQLWithConnection(lsSQL4 _
                        , ladrsParms _
                        , ladcn _
                        , adOpenStatic _
                        , adLockBatchOptimistic _
                        , adCmdText)
    Set ladrsParms = Nothing
    lsSQL4 = ""

    Set radrsDepartment = NewExecuteSQLWithConnection(lsSQL5 _
                        , Nothing _
                        , ladcn _
                        , adOpenStatic _
                        , adLockBatchOptimistic _
                        , adCmdText)
    lsSQL5 = ""

    Set radrsFunction = NewExecuteSQLWithConnection(lsSQL6 _
                        , Nothing _
                        , ladcn _
                        , adOpenStatic _
                        , adLockBatchOptimistic _
                        , adCmdText)
    lsSQL6 = ""

    Set radrsDivision = NewExecuteSQLWithConnection(lsSQL7 _
                        , Nothing _
                        , ladcn _
                        , adOpenStatic _
                        , adLockBatchOptimistic _
                        , adCmdText)
    lsSQL7 = ""

    ExecuteOTHERSQL lsSQL8 _
        , ladcn _
        , adCmdText
    lsSQL8 = ""

    'Close and Drop connection
    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

    GetObjectContext.SetComplete
    Exit Sub
GetAllDSInformation_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    If Not (ladcn Is Nothing) Then
        lsSQL8 = "DROP TABLE SESSION.GTT_EMP_APP;"
        ExecuteOTHERSQL lsSQL8, ladcn, adCmdText
    End If

    Set ladcn = Nothing

Dim lsSQL As String

    If (Trim(lsSQL1) <> "") Then
        lsSQL = lsSQL1
    Else
        If (Trim(lsSQL2) <> "") Then
            lsSQL = lsSQL2
        Else
            If (Trim(lsSQL3) <> "") Then
                lsSQL = lsSQL3
            Else
                If (Trim(lsSQL4) <> "") Then
                    lsSQL = lsSQL4
                Else
                    If (Trim(lsSQL5) <> "") Then
                        lsSQL = lsSQL5
                    Else
                        If (Trim(lsSQL6) <> "") Then
                            lsSQL = lsSQL6
                        Else
                            If (Trim(lsSQL7) <> "") Then
                                lsSQL = lsSQL7
                            Else
                                If (Trim(lsSQL8) <> "") Then
                                    lsSQL = lsSQL8
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetAllDSInformation"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetAllDSInformation" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL)
        Case Else
            TrsRaiseError MODULENAME _
                , "GetAllDSInformation" _
                , trErrdb_Approval_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Sub
