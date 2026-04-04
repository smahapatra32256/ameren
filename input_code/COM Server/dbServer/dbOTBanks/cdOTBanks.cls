VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdOTBanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'May 2004
'Author: Prem M. Chander
'Module Description: This class will provide data access
'to get/create/modify Employee Overtime Banks

Option Explicit

Private Const MODULENAME = "dbOTBanks.cdOTBanks"
Public Function GetEmployeeOvertimeBanks(ByVal vsEmployeeWhereClause As String _
                , ByVal viOTBankCode As Integer _
                ) As ADODB.Recordset
On Error GoTo GetEmployeeOvertimeBanks_EH
Dim lsSQL As String, ladcn As ADODB.Connection
Dim ladrsParms As ADODB.Recordset

    Set GetEmployeeOvertimeBanks = Nothing

    lsSQL = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_EMP" & _
        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL  ) CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"
    lsSQL = lsSQL & _
        " INSERT INTO SESSION.GTT_EMP (EMP_NUM, EMP_NAME)" & _
        " SELECT A.EMP_NUM, A.EMP_NAME FROM METHDBA.TTR_EMPLOYEE A" & _
        " WHERE A.EMP_NUM IN (" & Trim(vsEmployeeWhereClause) & ")" & _
        " GROUP BY A.EMP_NUM, A.EMP_NAME;"

    Set ladcn = OpenConnection
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    lsSQL = " SELECT  A.EMP_NUM, B.EMP_NAME" & _
        "       , A.OT_BANK_CD, C.OT_BANK_DESC" & _
        "       , A.TOT_OT_HRS, A.TOT_OT_HRS_REF" & _
        "       , (A.TOT_OT_HRS - A.TOT_OT_HRS_REF) AS TOT_OT_HRS_ACC" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
        " FROM  SESSION.GTT_EMP B" & _
        "       INNER JOIN METHDBA.TTR_OT_BANKS A" & _
        "       ON A.EMP_NUM = B.EMP_NUM" & _
        "       AND A.OT_BANK_CD = ?" & _
        "       INNER JOIN METHDBA.TTR_OT_BANK_CD C" & _
        "       ON C.OT_BANK_CD = A.OT_BANK_CD;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = viOTBankCode
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeOvertimeBanks = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeeOvertimeBanks = NewExecuteSQLWithConnection(lsSQL _
                            , ladrsParms _
                            , ladcn _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    lsSQL = "DROP TABLE SESSION.GTT_EMP;"
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetEmployeeOvertimeBanks_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    If Not (ladcn Is Nothing) Then
        lsSQL = "DROP TABLE SESSION.GTT_EMP;"
        ExecuteOTHERSQL lsSQL _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetEmployeeOvertimeBanks"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetEmployeeOvertimeBanks" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetEmployeeOvertimeBanks" _
                    , trErrdbGetOTBanks _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeeOvertimeBanks2(ByVal vsEmployeeWhereClause As String _
                ) As ADODB.Recordset
On Error GoTo GetEmployeeOvertimeBanks2_EH
Dim lsSQL As String, ladcn As ADODB.Connection

    Set GetEmployeeOvertimeBanks2 = Nothing

    lsSQL = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_EMP" & _
        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL  ) CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"
    lsSQL = lsSQL & _
        " INSERT INTO SESSION.GTT_EMP (EMP_NUM, EMP_NAME)" & _
        " SELECT A.EMP_NUM, A.EMP_NAME FROM METHDBA.TTR_EMPLOYEE A" & _
        " WHERE A.EMP_NUM IN (" & Trim(vsEmployeeWhereClause) & ")" & _
        " GROUP BY A.EMP_NUM, A.EMP_NAME;"

    Set ladcn = OpenConnection
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    lsSQL = " SELECT  A.EMP_NUM, B.EMP_NAME" & _
        "       , A.OT_BANK_CD, C.OT_BANK_DESC" & _
        "       , A.TOT_OT_HRS, A.TOT_OT_HRS_REF" & _
        "       , (A.TOT_OT_HRS - A.TOT_OT_HRS_REF) AS TOT_OT_HRS_ACC" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
        " FROM  SESSION.GTT_EMP B" & _
        "       INNER JOIN METHDBA.TTR_OT_BANKS A" & _
        "       ON A.EMP_NUM = B.EMP_NUM" & _
        "       INNER JOIN METHDBA.TTR_OT_BANK_CD C" & _
        "       ON C.OT_BANK_CD = A.OT_BANK_CD;"

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeOvertimeBanks2 = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeeOvertimeBanks2 = NewExecuteSQLWithConnection(lsSQL _
                            , Nothing _
                            , ladcn _
                            , adOpenStatic _
                            , adLockReadOnly)

    lsSQL = "DROP TABLE SESSION.GTT_EMP;"
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetEmployeeOvertimeBanks2_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    If Not (ladcn Is Nothing) Then
        lsSQL = "DROP TABLE SESSION.GTT_EMP;"
        ExecuteOTHERSQL lsSQL _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetEmployeeOvertimeBanks2"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetEmployeeOvertimeBanks2" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetEmployeeOvertimeBanks2" _
                    , trErrdbGetOTBanks _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'METHOD NO LONGER USED
Public Function GetEmployeeOvertimeBanksForRezeroing(ByVal vsEmployeeWhereClause As String _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeOvertimeBanksForRezeroing_EH

    Set GetEmployeeOvertimeBanksForRezeroing = Nothing
    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetEmployeeOvertimeBanksForRezeroing_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    TrsRaiseError MODULENAME _
        , "GetEmployeeOvertimeBanksForRezeroing" _
        , trErrdbGetOTBanks _
        , , , _
        , True

End Function
Public Function GetOTGroupEmployeesStandingBank(ByVal vlOTGroupNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetOTGroupEmployeesStandingBank_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetOTGroupEmployeesStandingBank = Nothing

    lsSQL = "SELECT  A.EMP_NUM, F.EMP_NAME" & _
        "       , A.TOT_OT_HRS" & _
        "       , (A.TOT_OT_HRS - A.TOT_OT_HRS_REF) AS TOT_OT_HRS_ACC" & _
        "       , A.TOT_OT_HRS_REF" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
        "       , A.OT_BANK_CD, E.OT_BANK_DESC" & _
        " FROM  METHDBA.TTR_OT_BANKS A" & _
        "       INNER JOIN METHDBA.TTR_OT_GROUP B" & _
        "       ON B.OT_STAND_BANK_CD = A.OT_BANK_CD" & _
        "       AND B.OT_GROUP_NUM = ?" & _
        "       INNER JOIN METHDBA.TTR_EMP_WG C" & _
        "       ON C.EMP_NUM = A.EMP_NUM" & _
        "       AND CURRENT DATE BETWEEN C.EFF_START_DATE" & _
        "       AND C.EFF_END_DATE" & _
        "       INNER JOIN METHDBA.TTR_OT_GROUP_WG D" & _
        "       ON D.WG_NUM = C.WG_NUM" & _
        "       AND D.OT_GROUP_NUM = ?" & _
        "       INNER JOIN METHDBA.TTR_OT_BANK_CD E" & _
        "       ON E.OT_BANK_CD = A.OT_BANK_CD" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE F" & _
        "       ON F.EMP_NUM = A.EMP_NUM" & _
        " ORDER BY A.TOT_OT_HRS;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOTGroupNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOTGroupNumber
        .Update
    End With

    Set GetOTGroupEmployeesStandingBank = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

'    "SELECT    A.EMP_NUM, B.EMP_NAME" & _
'        "           , A.TOT_OT_HRS" & _
'        "           , (A.TOT_OT_HRS - A.TOT_OT_HRS_REF) AS TOT_OT_HRS_ACC" & _
'        "           , A.TOT_OT_HRS_REF" & _
'        "           , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
'        "           , A.OT_BANK_CD, C.OT_BANK_DESC" & _
'        " FROM      METHDBA.TTR_OT_BANKS A, METHDBA.TTR_EMPLOYEE B" & _
'        "           , METHDBA.TTR_OT_BANK_CD C" & _
'        " WHERE     A.OT_BANK_CD = (SELECT OT_STAND_BANK_CD" & _
'        "                       FROM    METHDBA.TTR_OT_GROUP" & _
'        "                       WHERE   OT_GROUP_NUM = " & CStr(vlOTGroupNumber) & ")" & _
'        " AND       A.EMP_NUM IN (SELECT  DISTINCT EMP_NUM " & _
'        "                   FROM    METHDBA.TTR_EMP_WG " & _
'        "                   WHERE   '" & Format(Date, "mm/dd/yyyy") & "'" & _
'        "                           BETWEEN EFF_START_DATE AND EFF_END_DATE " & _
'        "                   AND     WG_NUM IN (SELECT   WG_NUM " & _
'        "                               FROM    METHDBA.TTR_OT_GROUP_WG " & _
'        "                               WHERE   OT_GROUP_NUM = " & CStr(vlOTGroupNumber) & ")) " & _
'        " AND       A.EMP_NUM = B.EMP_NUM " & _
'        " AND       A.OT_BANK_CD = C.OT_BANK_CD " & _
'        " ORDER BY A.TOT_OT_HRS"
    'Call ExecuteSQL and retrieve information
'    Set GetOTGroupEmployeesStandingBank = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetOTGroupEmployeesStandingBank_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetOTGroupEmployeesStandingBank"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetOTGroupEmployeesStandingBank" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetOTGroupEmployeesStandingBank" _
                    , trErrdbGetOTBanks _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetBanksForRezeroByGroupAndBank(ByVal viOTBankCode As Integer _
                    , ByVal vlOTGroupNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetBanksForRezeroByGroupAndBank_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetBanksForRezeroByGroupAndBank = Nothing

    'If Volunteer Bank Info then drop total ot hours, total hours refused &
    'accepted in the volunteer field then set the rest ot zero and viceversa.
    lsSQL = "SELECT  A.EMP_NUM, E.EMP_NAME" & _
        "       , A.TOT_OT_HRS" & _
        "       , (A.TOT_OT_HRS - A.TOT_OT_HRS_REF) AS TOT_OT_HRS_ACC" & _
        "       , A.TOT_OT_HRS_REF" & _
        "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
        "       , A.OT_BANK_CD, D.OT_BANK_DESC" & _
        " FROM  METHDBA.TTR_OT_BANKS A" & _
        "       INNER JOIN METHDBA.TTR_EMP_WG B" & _
        "       ON B.EMP_NUM = A.EMP_NUM" & _
        "       AND CURRENT DATE BETWEEN B.EFF_START_DATE" & _
        "       AND B.EFF_END_DATE" & _
        "       INNER JOIN METHDBA.TTR_OT_GROUP_WG C" & _
        "       ON C.WG_NUM = B.WG_NUM" & _
        "       AND C.OT_GROUP_NUM = ?" & _
        "       INNER JOIN METHDBA.TTR_OT_BANK_CD D" & _
        "       ON D.OT_BANK_CD = A.OT_BANK_CD" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE E" & _
        "       ON E.EMP_NUM = A.EMP_NUM" & _
        " WHERE A.OT_BANK_CD = ?" & _
        " ORDER BY A.TOT_OT_HRS;"
    
'    lsSQL = "SELECT    A.EMP_NUM, B.EMP_NAME" & _
'        "           , A.TOT_OT_HRS" & _
'        "           , (A.TOT_OT_HRS - A.TOT_OT_HRS_REF) AS TOT_OT_HRS_ACC" & _
'        "           , A.TOT_OT_HRS_REF" & _
'        "           , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
'        "           , A.OT_BANK_CD, C.OT_BANK_DESC" & _
'        " FROM      METHDBA.TTR_OT_BANKS A, METHDBA.TTR_EMPLOYEE B" & _
'        "           , METHDBA.TTR_OT_BANK_CD C" & _
'        " WHERE     A.OT_BANK_CD = " & CStr(viOTBankCode) & _
'        " AND       A.EMP_NUM IN (SELECT  DISTINCT EMP_NUM" & _
'        "                   FROM    METHDBA.TTR_EMP_WG" & _
'        "                   WHERE   '" & Format(Date, "mm/dd/yyyy") & "'" & _
'        "                           BETWEEN EFF_START_DATE AND EFF_END_DATE" & _
'        "                   AND     WG_NUM IN (SELECT   WG_NUM" & _
'        "                               FROM    METHDBA.TTR_OT_GROUP_WG" & _
'        "                               WHERE   OT_GROUP_NUM = " & CStr(vlOTGroupNumber) & "))" & _
'        " AND       A.EMP_NUM = B.EMP_NUM" & _
'        " AND       A.OT_BANK_CD = C.OT_BANK_CD" & _
'        " ORDER BY A.TOT_OT_HRS;"
'
    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOTGroupNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = viOTBankCode
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetBanksForRezeroByGroupAndBank = ExecuteSQL(lsSQL _
'                    , adOpenStatic _
'                    , adLockReadOnly)
    Set GetBanksForRezeroByGroupAndBank = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenStatic _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetBanksForRezeroByGroupAndBank_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetBanksForRezeroByGroupAndBank"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetBanksForRezeroByGroupAndBank" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetBanksForRezeroByGroupAndBank" _
                    , trErrdbGetOTBanks _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'This method will process delete, insert and update to
'the OT Banks table and also insert records in the
'OT Banks Audit Table.
Public Function dbProcessEmployeeOvertimeBanks( _
            ByVal vadrsOvertimeBankEntries As ADODB.Recordset _
            , ByVal vadrsOvertimeBankAuditEntries As ADODB.Recordset _
            , ByVal vlUserEmployeeNumber As Long _
            ) As Boolean
On Error GoTo dbProcessEmployeeOvertimeBanks_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

    dbProcessEmployeeOvertimeBanks = False
    lsSQL = ""

    'Check for existence!
    If Not (vadrsOvertimeBankEntries Is Nothing) Then
        With vadrsOvertimeBankEntries
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes, if present process them first
                    .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")

                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " DELETE    FROM    METHDBA.TTR_OT_BANKS " & _
                                    " WHERE     EMP_NUM = " & !KEY_EMP_NUM & _
                                    " AND       OT_BANK_CD = " & !KEY_OT_BANK_CD & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone

                    'Check for inserts
                    .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " INSERT    INTO METHDBA.TTR_OT_BANKS " & _
                                    " (EMP_NUM, OT_BANK_CD " & _
                                    " , TOT_OT_HRS , TOT_OT_HRS_REF " & _
                                    " , TIME_STAMP, LAST_MOD_EMP_NUM) " & _
                                    " VALUES " & _
                                    " (" & CStr(!EMP_NUM) & ", " & CStr(!OT_BANK_CD) & _
                                    " , " & CStr(!TOT_OT_HRS) & ", " & CStr(!TOT_OT_HRS_REF) & _
                                    " , CURRENT TIMESTAMP, " & CStr(vlUserEmployeeNumber) & ");"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone

                    'Check for deletes, if present process them first
                    .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " UPDATE    METHDBA.TTR_OT_BANKS " & _
                                    " SET       TOT_OT_HRS = " & CStr(!TOT_OT_HRS) & _
                                    "           , TOT_OT_HRS_REF = " & CStr(!TOT_OT_HRS_REF) & _
                                    "           , TIME_STAMP = CURRENT TIMESTAMP " & _
                                    "           , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                    " WHERE     EMP_NUM = " & CStr(!EMP_NUM) & _
                                    " AND       OT_BANK_CD = " & CStr(!KEY_OT_BANK_CD) & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
            End If
        End With
    End If

    'Process Audits
    If Not (vadrsOvertimeBankAuditEntries Is Nothing) Then
        With vadrsOvertimeBankAuditEntries
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone
', PAYROLL_PROC_IND WAS REMOVED FROM THE OT_BANK_ADJ TABLE
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " INSERT    INTO METHDBA.TTR_OT_BANK_ADJ " & _
                                " (EMP_NUM, OT_BANK_CD " & _
                                " , HRS_ADJ " & _
                                " , REASON, OT_OPP_NUM " & _
                                " , ACC_REF_IND " & _
                                " , TIME_STAMP, LAST_MOD_EMP_NUM) " & _
                                " VALUES " & _
                                " (" & CStr(!EMP_NUM) & ", " & CStr(!OT_BANK_CD) & _
                                " , " & CStr(!HRS_ADJ) & _
                                " , '" & Trim(!REASON) & "', " & CStr(!OT_OPP_NUM) & _
                                " , '" & Trim(!ACC_REF_IND) & "'" & _
                                " , CURRENT TIMESTAMP, " & CStr(vlUserEmployeeNumber) & ");"

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
    End If

    'close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessEmployeeOvertimeBanks = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessEmployeeOvertimeBanks_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbProcessEmployeeOvertimeBanks"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbProcessEmployeeOvertimeBanks" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
''            Case trErrdbDelFromOTBanks, trErrdbInsIntoOTBanks _
''                , trErrdbUpdOTBanks, trErrdbInsIntoOTBanksAudit
''                TrsRaiseError MODULENAME _
''                    , "dbProcessEmployeeOvertimeBanks" _
''                    , Err.Number _
''                    , , , _
''                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbProcessEmployeeOvertimeBanks" _
                    , trErrdbSaveOTBanks _
                    , , , _
                    , True
    End Select

End Function
