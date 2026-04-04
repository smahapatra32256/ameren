VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdRequestFirstAppr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Conversion Phase 1
'Author:    Prem M. Chander
'Date:      December 2000 (Firstcut)
'Description: This class module will reside in a Server Package
'under Microsoft Transaction Server (MTS). This module will be called by
'its associated business component. This module will access the database for
'pertinant information

'NOTE: Apply MDAC25FilterFix if necessary, always Format date
'to mm/dd/yyyy in INSERT, UPDATE SQL and WHERE Clauses of SQL.

Option Explicit

'Module Name
Private Const MODULENAME = "dbAdjustment.cdRequestFirstAppr"
Public Function db_GetAdjustments(ByVal v_iAdjustmentType As Integer _
                , Optional ByVal v_lRequestByEmployee As Long _
                , Optional ByVal v_sDepartmentInClause As String _
                , Optional ByVal v_sRequestEmployeeJob As String _
                ) As ADODB.Recordset
On Error GoTo db_GetAdjustments_EH
'SQL Statement Container
Dim lsSQL As String
Dim ladrsParm As ADODB.Recordset

    Set db_GetAdjustments = Nothing

    lsSQL = "SELECT DISTINCT A.ADJ_RQT_NUM" & _
        "       , A.EMP_NUM, B.EMP_NAME" & _
        "       , A.DEPT_NUM, A.SPECIAL_CHECK_IND" & _
        "       , A.CREATED_EMP_NUM, COALESCE(C.EMP_NAME, '') AS CREATED_EMP_NAME" & _
        "       , CASE COALESCE(DATE(A.CREATED_TIMESTAMP), DATE('0100-01-01'))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(DATE(A.CREATED_TIMESTAMP), DATE('0100-01-01')) END AS CREATED_DATE" & _
        "       , COALESCE('(' || SUBSTR(C.WORK_PHONE_NUM, 1, 3) || ') ' || SUBSTR(C.WORK_PHONE_NUM, 4, 3) || '-' || SUBSTR(C.WORK_PHONE_NUM, 7, 4), '') AS CREATED_EMP_PHONE" & _
        "       , CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , A.APPROVED_EMP_NUM, COALESCE(D.EMP_NAME, '') AS APPROVED_EMP_NAME" & _
        "       , CASE COALESCE(DATE(A.APPROVED_TM_STMP), DATE('0100-01-01'))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(DATE(A.APPROVED_TM_STMP), DATE('0100-01-01')) END AS APPROVED_DATE" & _
        "       , A.COMPLETED_BY_EMP, COALESCE(E.EMP_NAME, '') AS COMPLETED_EMP_NAME" & _
        "       , CASE COALESCE(DATE(A.COMPLETED_TM_STMP), DATE('0100-01-01'))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(DATE(A.COMPLETED_TM_STMP), DATE('0100-01-01')) END AS COMPLETED_DATE" & _
        "       , A.ATHRZ_EMP_NUM, COALESCE(F.EMP_NAME, '') AS ATHRZ_EMP_NAME" & _
        "       , CASE COALESCE(DATE(A.ATHRZ_TM_STMP), DATE('0100-01-01'))" & _
        "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
        "       ELSE COALESCE(DATE(A.ATHRZ_TM_STMP), DATE('0100-01-01')) END AS ATHRZ_DATE"
    lsSQL = lsSQL & _
        "       , A.REQ_REASON" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM  METHDBA.TTR_ADJ_RQT A"

    'Prepare Where Clause Statement
    Select Case (v_iAdjustmentType)
        Case ADJUSTMENTREQUESTFOR  'Request For
            lsSQL = lsSQL & _
                "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
                "       ON B.EMP_NUM = A.EMP_NUM" & _
                "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE C" & _
                "       ON C.EMP_NUM = A.CREATED_EMP_NUM" & _
                "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE D" & _
                "       ON D.EMP_NUM = A.APPROVED_EMP_NUM" & _
                "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE E" & _
                "       ON E.EMP_NUM = A.COMPLETED_BY_EMP" & _
                "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE F" & _
                "       ON F.EMP_NUM = A.ATHRZ_EMP_NUM" & _
                " WHERE A.CREATED_EMP_NUM = ?" & _
                " AND   A.APPROVED_TM_STMP IS NULL" & _
                " AND   A.COMPLETED_TM_STMP IS NULL" & _
                " AND   A.ATHRZ_TM_STMP IS NULL"

        Case ADJUSTMENTFIRSTAPPROVAL  'First Approval
            lsSQL = lsSQL & _
                "       INNER JOIN METHDBA.TTR_DS_DEPT_DTL G" & _
                "       ON G.EFF_EMP_NUM = ?" & _
                "       AND A.DEPT_NUM = G.DEPT_NUM" & _
                "       AND A.COMPLETED_TM_STMP IS NULL" & _
                "       AND A.ATHRZ_TM_STMP IS NULL" & _
                "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
                "       ON B.EMP_NUM = A.EMP_NUM" & _
                "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE C" & _
                "       ON C.EMP_NUM = A.CREATED_EMP_NUM" & _
                "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE D" & _
                "       ON D.EMP_NUM = A.APPROVED_EMP_NUM" & _
                "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE E" & _
                "       ON E.EMP_NUM = A.COMPLETED_BY_EMP" & _
                "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE F" & _
                "       ON F.EMP_NUM = A.ATHRZ_EMP_NUM"
    End Select

    lsSQL = lsSQL & _
        " ORDER BY A.DEPT_NUM, CREATED_DATE;"

    Set ladrsParm = CreateParameterRS
    With ladrsParm
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = v_lRequestByEmployee
        .Update
    End With

    Set db_GetAdjustments = NewExecuteSQL(lsSQL _
                            , ladrsParm _
                            , adOpenStatic _
                            , adLockBatchOptimistic)
    Set ladrsParm = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_GetAdjustments_EH:
    'Abort Transaction
    GetObjectContext.SetAbort
    Set ladrsParm = Nothing

Dim llMyErrNumber As Long

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (v_iAdjustmentType)
        Case ADJUSTMENTREQUESTFOR
            llMyErrNumber = trErrdbGetReqAdj
        Case ADJUSTMENTFIRSTAPPROVAL
            llMyErrNumber = trErrdbGetFirstAppAdj
        Case ADJUSTMENTPAYROLLCOMPLETE
            llMyErrNumber = trErrdbGetPayCompAdj
        Case ADJUSTMENTFINALAPPROVAL
            llMyErrNumber = trErrdbGetFinalAppAdj
    End Select

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "db_GetAdjustments"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "db_GetAdjustments" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "db_GetAdjustments" _
                , llMyErrNumber _
                , , , _
                , True
    End Select

End Function
Public Function db_InsertRequest( _
                ByVal v_adrsInsertRequest As ADODB.Recordset _
                , ByVal v_lUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo db_InsertRequest_EH
'SQL
Dim lsSQL As String
'ADO
Dim ladcnConnection As ADODB.Connection
Dim lsFilterString As String, llRecordIndex As Long
'Sequence Number
Dim lobjUtility As utilDBDataServices.cuDBDataSvcNoTxn
Dim llAdjustmentRequestNumber As Long

    db_InsertRequest = False

    lsSQL = ""

    With v_adrsInsertRequest
        .Filter = adFilterNone

        lsFilterString = "RECORD_STATUS = 'I'"
        .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
        lsFilterString = ""
        If Not (.EOF) Then
            .MoveFirst

            For llRecordIndex = 1 To .RecordCount
                llAdjustmentRequestNumber = -1

                Set lobjUtility = CreateObject("utilDBDataServices.cuDBDataSvcNoTxn")
                '//llAdjustmentRequestNumber = lobjUtility.GetNextSequenceNumber("AD")
                llAdjustmentRequestNumber = lobjUtility.GetNextAdjustmentSequenceNumber()
                Set lobjUtility = Nothing

                If (llAdjustmentRequestNumber > 0) Then
                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ADJ_RQT" & _
                        " (ADJ_RQT_NUM, EMP_NUM" & _
                        " , DEPT_NUM, SPECIAL_CHECK_IND" & _
                        " , CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                        " , LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP" & _
                        " , REQ_REASON)" & _
                        " VALUES" & _
                        " (" & CStr(llAdjustmentRequestNumber) & ", " & CStr(!EMP_NUM) & _
                        " , '" & Trim(!DEPT_NUM) & "', '" & Trim(!SPECIAL_CHECK_IND) & "'" & _
                        " , " & CStr(v_lUserEmployeeNumber) & ", CURRENT TIMESTAMP" & _
                        " , " & CStr(v_lUserEmployeeNumber) & ", CURRENT TIMESTAMP" & _
                        " , '" & Trim(!REQ_REASON) & "');"
                End If

                .MoveNext
            Next llRecordIndex

            Set lobjUtility = Nothing
        End If

        .Filter = adFilterNone
    End With

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    'close and destroy
    If Not (ladcnConnection Is Nothing) Then
        If (ladcnConnection.State = adStateOpen) Then
            ladcnConnection.Close
        End If

        Set ladcnConnection = Nothing
    End If

    db_InsertRequest = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_InsertRequest_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing
    Set lobjUtility = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "db_InsertRequest"

        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "db_InsertRequest" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True

        Case Else
            TrsRaiseError MODULENAME _
                , "db_InsertRequest" _
                , trErrdbInsReqAdj _
                , , , _
                , True
    End Select

End Function
Public Function db_UpdateRequest( _
                ByVal v_adrsUpdateRequest As ADODB.Recordset _
                , ByVal v_lUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo db_UpdateRequest_EH
'SQL
Dim lsSQL As String
'ADO
Dim ladcnConnection As ADODB.Connection
Dim l_sFilterString As String, l_lRecordIndex As Long
Dim lsDB2LowDate As String

    db_UpdateRequest = False

    lsSQL = ""

    With v_adrsUpdateRequest
        .Filter = adFilterNone
        .MoveFirst

        l_sFilterString = "RECORD_STATUS = 'C'"
        .Filter = stdmTRISCommon.MDAC25FilterFix(l_sFilterString)
        l_sFilterString = ""
        If Not (.EOF) Then
            .MoveFirst

            For l_lRecordIndex = 1 To .RecordCount
                'NOTE: CREATED_DATE IS EXCLUDED FROM THE WHERE CLAUSE AS IT IS
                'DERIEVED FROM THE CREATED_TIMESTAMP. ALSO MICROSOFT LOW DATE IS
                '01/01/0100 AND DB2 LOW DATE IS 01/01/0001
                lsSQL = lsSQL & _
                        " UPDATE METHDBA.TTR_ADJ_RQT" & _
                        " SET   DEPT_NUM = '" & Trim(!DEPT_NUM) & "'" & _
                        "       , SPECIAL_CHECK_IND = '" & Trim(!SPECIAL_CHECK_IND) & "'" & _
                        "       , REQ_REASON = '" & Trim(!REQ_REASON) & "'" & _
                        "       , LAST_MOD_EMP_NUM = " & CStr(v_lUserEmployeeNumber) & _
                        "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP " & _
                        " WHERE ADJ_RQT_NUM = " & CStr(!ADJ_RQT_NUM) & _
                        " AND   EMP_NUM = " & CStr(!EMP_NUM) & _
                        " AND   CREATED_EMP_NUM = " & CStr(!CREATED_EMP_NUM) & _
                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                .MoveNext
            Next l_lRecordIndex
        End If

        .Filter = adFilterNone
    End With

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    If Not (ladcnConnection Is Nothing) Then
        If (ladcnConnection.State = adStateOpen) Then
            ladcnConnection.Close
        End If

        Set ladcnConnection = Nothing
    End If

    db_UpdateRequest = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_UpdateRequest_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "db_UpdateRequest"

        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "db_UpdateRequest" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True

        Case Else
            TrsRaiseError MODULENAME _
                , "db_UpdateRequest" _
                , trErrdbUpdReqAdj _
                , , , _
                , True
    End Select

End Function
Public Function db_UpdateWithFirstApproval( _
                ByVal v_adrsUpdateWithFirstApproval As ADODB.Recordset _
                , ByVal v_lUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo db_UpdateWithFirstApproval_EH
'SQL
Dim lsSQL As String
'ADO
Dim ladcnConnection As ADODB.Connection
Dim lsFilterString As String, llRecordIndex As Long

    db_UpdateWithFirstApproval = False

    lsSQL = ""

    With v_adrsUpdateWithFirstApproval
        .Filter = adFilterNone
        .MoveFirst

        lsFilterString = "RECORD_STATUS = 'C'"
        .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
        lsFilterString = ""
        If Not (.EOF) Then
            .MoveFirst

            For llRecordIndex = 1 To .RecordCount
                'NOTE: CREATED_DATE IS EXCLUDED FROM THE WHERE CLAUSE AS IT IS
                'DERIEVED FROM THE CREATED_TIMESTAMP. ALSO MICROSOFT LOW DATE IS
                '01/01/0100 AND DB2 LOW DATE IS 01/01/0001
                Select Case (Trim(!APPROVED_IND))
                    Case "Y"
                        lsSQL = lsSQL & _
                                " UPDATE METHDBA.TTR_ADJ_RQT" & _
                                " SET   DEPT_NUM = '" & Trim(!DEPT_NUM) & "'" & _
                                "       , SPECIAL_CHECK_IND = '" & Trim(!SPECIAL_CHECK_IND) & "'" & _
                                "       , REQ_REASON = '" & Trim(!REQ_REASON) & "'" & _
                                "       , APPROVED_EMP_NUM = " & CStr(!APPROVED_EMP_NUM) & _
                                "       , APPROVED_TM_STMP = CURRENT TIMESTAMP" & _
                                "       , LAST_MOD_EMP_NUM = " & CStr(v_lUserEmployeeNumber) & _
                                "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
                                " WHERE ADJ_RQT_NUM = " & CStr(!ADJ_RQT_NUM) & _
                                " AND   EMP_NUM = " & CStr(!EMP_NUM) & _
                                " AND   CREATED_EMP_NUM = " & CStr(!CREATED_EMP_NUM) & _
                                " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                    Case "N"
                        lsSQL = lsSQL & _
                                " UPDATE METHDBA.TTR_ADJ_RQT" & _
                                " SET   DEPT_NUM = '" & Trim(!DEPT_NUM) & "'" & _
                                "       , SPECIAL_CHECK_IND = '" & Trim(!SPECIAL_CHECK_IND) & "'" & _
                                "       , REQ_REASON = '" & Trim(!REQ_REASON) & "'" & _
                                "       , LAST_MOD_EMP_NUM = " & CStr(v_lUserEmployeeNumber) & _
                                "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
                                "       , APPROVED_EMP_NUM = 0" & _
                                "       , APPROVED_TM_STMP = NULL" & _
                                " WHERE ADJ_RQT_NUM = " & CStr(!ADJ_RQT_NUM) & _
                                " AND   EMP_NUM = " & CStr(!EMP_NUM) & _
                                " AND   CREATED_EMP_NUM = " & CStr(!CREATED_EMP_NUM) & _
                                " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"
                End Select

                .MoveNext
            Next llRecordIndex
        End If

        .Filter = adFilterNone
    End With

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    If Not (ladcnConnection Is Nothing) Then
        If (ladcnConnection.State = adStateOpen) Then
            ladcnConnection.Close
        End If

        Set ladcnConnection = Nothing
    End If

    db_UpdateWithFirstApproval = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_UpdateWithFirstApproval_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "db_UpdateWithFirstApproval"

        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "db_UpdateWithFirstApproval" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True

        Case Else
            TrsRaiseError MODULENAME _
                , "db_UpdateWithFirstApproval" _
                , trErrdbUpdFirstAppAdj _
                , , , _
                , True
    End Select

End Function
Public Function db_DeleteReqFirstApp( _
                ByVal v_adrsDeleteReqFirstApp As ADODB.Recordset _
                , ByVal v_lUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo db_DeleteReqFirstApp_EH
'SQL
Dim lsSQL As String
'ADO
Dim ladcnConnection As ADODB.Connection
Dim lsFilterString As String, llRecordIndex As Long

    db_DeleteReqFirstApp = False

    lsSQL = ""

    With v_adrsDeleteReqFirstApp
        .Filter = adFilterNone
        .MoveFirst

        lsFilterString = "RECORD_STATUS = 'D'"
        .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
        lsFilterString = ""
        If Not (.EOF) Then
            .MoveFirst

            For llRecordIndex = 1 To .RecordCount
                lsSQL = lsSQL & _
                    " DELETE FROM METHDBA.TTR_ADJ_RQT" & _
                    " WHERE ADJ_RQT_NUM = " & CStr(!ADJ_RQT_NUM) & _
                    " AND   EMP_NUM = " & CStr(!EMP_NUM) & _
                    " AND   CREATED_EMP_NUM = " & CStr(!CREATED_EMP_NUM) & _
                    " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                .MoveNext
            Next llRecordIndex
        End If

        .Filter = adFilterNone
    End With

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    If Not (ladcnConnection Is Nothing) Then
        If (ladcnConnection.State = adStateOpen) Then
            ladcnConnection.Close
        End If

        Set ladcnConnection = Nothing
    End If

    db_DeleteReqFirstApp = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_DeleteReqFirstApp_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "db_DeleteReqFirstApp"

        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "db_DeleteReqFirstApp" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True

        Case Else
            TrsRaiseError MODULENAME _
                , "db_DeleteReqFirstApp" _
                , trErrdbDelAdj _
                , , , _
                , True
    End Select

End Function


