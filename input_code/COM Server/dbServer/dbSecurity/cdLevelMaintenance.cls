VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdLevelMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 2
'March 2002
'Author: Prem M. Chander
'Module Description: This class will retrieve information from
'the datastore.
Option Explicit

Private Const MODULENAME As String = "dbSecurity.cdLevelMaintenance"
'This function determines those TRIS menu items that are available
'for a given Security Level
Public Function db_GetLevels(ByVal v_iToGet As Integer _
                , Optional ByVal v_lSecurityLevel As Long _
                , Optional ByVal v_sDivision As String _
                , Optional ByVal vbSystemAdmin As Boolean = False _
                ) As ADODB.Recordset
On Error GoTo db_GetLevels_EH
Dim ladrs As ADODB.Recordset
'Declare a variable for SQL Statement
Dim lsSQL As String

    Set db_GetLevels = Nothing

    'Prepare SQL Statement
    lsSQL = "SELECT SEC_LEVEL, SEC_LEVEL_DESC" & _
            "       , DIVISION_NUM, LAST_MOD_EMP_NUM" & _
            "       , CHAR(TIME_STAMP) AS TIME_STAMP, SEC_LEVEL AS KEY_SEC_LEVEL" & _
            "       , DIVISION_NUM AS KEY_DIVISION_NUM" & _
            " FROM  METHDBA.TTR_SECURITY_LEVEL"

    'check for what to get
    Select Case (v_iToGet)
        Case GETALLSECLVLINFO
            'BEGIN TRIS SOX CHANGE JULY 2004
            If Not (vbSystemAdmin) Then
                lsSQL = lsSQL & _
                    " WHERE SEC_LEVEL NOT IN (9, 10, 11, 405, 981, 1045)"
            End If
            'END TRIS SOX CHANGE JULY 2004

        Case GETSECLVLINFOBYLEVEL
            'BEGIN TRIS SOX CHANGE JULY 2004
            If Not (vbSystemAdmin) Then
                Select Case (v_lSecurityLevel)
                    Case 9, 10, 11, 405, 981, 1045
                        GoTo PERFORMNOTASYSADMIN
                    Case Else
                        '
                End Select
            End If

            'lsSQL = lsSQL & " WHERE SEC_LEVEL = " & CStr(v_lSecurityLevel)
            lsSQL = lsSQL & " WHERE SEC_LEVEL = ?"

        Case GETALLSECLVLINFOBYDIVISION
            'lsSQL = lsSQL & " WHERE DIVISION_NUM = '" & CStr(v_sDivision) & "' "
            lsSQL = lsSQL & " WHERE DIVISION_NUM = ?"
            If Not (vbSystemAdmin) Then
                lsSQL = lsSQL & " AND SEC_LEVEL NOT IN (9, 10, 11, 405, 981, 1045)"
            End If
    End Select

    lsSQL = lsSQL & " ORDER BY  SEC_LEVEL;"

    Set ladrs = CreateParameterRS
    With ladrs
        Select Case (v_iToGet)
            Case GETSECLVLINFOBYLEVEL
                .AddNew
                !PARM_ORDER = 1
                !PARM_NAME = ""
                !PARM_TYPE = adInteger
                !PARM_DIRECTION = adParamInput
                !PARM_SIZE = 0
                !PARM_VALUE = v_lSecurityLevel
                .Update

            Case GETALLSECLVLINFOBYDIVISION
                .AddNew
                !PARM_ORDER = 1
                !PARM_NAME = ""
                !PARM_TYPE = adChar
                !PARM_DIRECTION = adParamInput
                !PARM_SIZE = 3
                !PARM_VALUE = v_sDivision
                .Update
        End Select
    End With

    'Execute
    'Set db_GetLevels = ExecuteSQL(lsSQL _
                        , adOpenStatic _
                        , adLockBatchOptimistic)
    Set db_GetLevels = NewExecuteSQL(lsSQL _
                        , ladrs _
                        , adOpenStatic _
                        , adLockBatchOptimistic)
PERFORMNOTASYSADMIN:
    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_GetLevels_EH:
    'Complete Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "db_GetLevels"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "db_GetLevels" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "db_GetLevels" _
                    , trErrdb_GetLevels_F _
                    , , , _
                    , True
    End Select

End Function
'This function retrieves employees associated to a Security Level
Public Function db_GetEmployeesForLevel(ByVal v_lSecurityLevel As Long _
                , ByVal v_bActiveOnly As Boolean _
                , ByVal v_dtLoadFromDate As Date _
                , Optional ByVal vbSystemAdmin As Boolean = False _
                ) As ADODB.Recordset
On Error GoTo db_GetEmployeesForLevel_EH
Dim ladrs As ADODB.Recordset
'Declare a variable for SQL Statement
Dim lsSQL As String

'" WHERE SEC_LEVEL = " & CStr(v_lSecurityLevel)

    Set db_GetEmployeesForLevel = Nothing

    'BEGIN TRIS SOX CHANGE JULY 2004
    Select Case (v_lSecurityLevel)
        Case 9, 10, 11, 405, 981, 1045
            If Not (vbSystemAdmin) Then
                GoTo PERFORMNOTASYSADMIN
            End If
        Case Else
            '
    End Select
    'END TRIS SOX CHANGE JULY 2004

    'Prepare SQL Statement
    lsSQL = "SELECT EMP_NUM, EMP_NAME" & _
            "       , DIVISION_NUM, LAST_MOD_EMP_NUM" & _
            "       , CHAR(TIME_STAMP) AS TIME_STAMP, EMP_NUM AS KEY_EMP_NUM" & _
            " FROM  METHDBA.TTR_EMPLOYEE" & _
            " WHERE SEC_LEVEL = ?"
    If (v_bActiveOnly) Then
        lsSQL = lsSQL & _
            " AND (EMP_STATUS <> 'T'" & _
            " OR  (EMP_STATUS = 'T'" & _
            " AND  STATUS_EFF_DATE > CURRENT DATE - 6 MONTHS))"
    End If
    lsSQL = lsSQL & " ORDER BY  EMP_NUM;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = v_lSecurityLevel
        .Update
    End With

    'Execute
    'Set db_GetEmployeesForLevel = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockBatchOptimistic)
    Set db_GetEmployeesForLevel = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockBatchOptimistic)
PERFORMNOTASYSADMIN:
    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_GetEmployeesForLevel_EH:
    'Complete Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "db_GetEmployeesForLevel"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "db_GetEmployeesForLevel" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "db_GetEmployeesForLevel" _
                    , trErrdb_GetEmployeesForLevel_F _
                    , , , _
                    , True
    End Select

End Function
Public Function db_GetAvailableBusFuncs(Optional ByVal vbSystemAdmin As Boolean = False _
                    ) As ADODB.Recordset
On Error GoTo db_GetAvailableBusFuncs_EH
'Declare a variable for SQL Statement
Dim lsSQL As String

    Set db_GetAvailableBusFuncs = Nothing

    'Prepare SQL Statement
    lsSQL = "(SELECT    A.BUS_FUNC_KEY, A.BUS_FUNC_NM" & _
        "       , A.BUS_FUNC_LBL, A.BUS_FUNC_PRNT_KEY" & _
        "       , B.BUS_FUNC_NM AS BUS_FUNC_PRNT_NM" & _
        "       , B.BUS_FUNC_LBL AS BUS_FUNC_PRNT_LBL" & _
        " FROM  METHDBA.TTR_BUS_FUNC A, METHDBA.TTR_BUS_FUNC B"
    'BEGIN TRIS SOX CHANGE JULY 2004
    If Not (vbSystemAdmin) Then
        lsSQL = lsSQL & _
            " WHERE   A.BUS_FUNC_PRNT_KEY = B.BUS_FUNC_KEY" & _
            " AND     A.BUS_FUNC_KEY NOT IN ('A61D', 'A62B'))"
    Else
        lsSQL = lsSQL & _
            " WHERE   A.BUS_FUNC_PRNT_KEY = B.BUS_FUNC_KEY)"
    End If
    'END TRIS SOX CHANGE JULY 2004

    lsSQL = lsSQL & _
        " UNION" & _
        " (SELECT   BUS_FUNC_KEY, BUS_FUNC_NM" & _
        "           , BUS_FUNC_LBL, '' AS BUS_FUNC_PRNT_KEY" & _
        "           , '' AS BUS_FUNC_PRNT_NM, '' AS BUS_FUNC_PRNT_LBL" & _
        " FROM      METHDBA.TTR_BUS_FUNC" & _
        " WHERE     BUS_FUNC_KEY = 'A0')" & _
        " ORDER BY   BUS_FUNC_KEY" & _
        " FOR FETCH ONLY WITH UR;"

    'Execute
    Set db_GetAvailableBusFuncs = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockBatchOptimistic)

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_GetAvailableBusFuncs_EH:
    'Complete Transaction
    GetObjectContext.SetAbort

    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "db_GetAvailableBusFuncs"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "db_GetAvailableBusFuncs" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "db_GetAvailableBusFuncs" _
                    , trErrdb_GetAvailableBusFuncs_F _
                    , , , _
                    , True
    End Select

End Function
Public Function db_GetAuthrzBusFuncsForLevel(ByVal v_lSecurityLevel As Long _
                    , Optional ByVal vbSystemAdmin As Boolean = False _
                    ) As ADODB.Recordset
On Error GoTo db_GetAuthrzBusFuncsForLevel_EH
Dim ladrs As ADODB.Recordset
'Declare a variable for SQL Statement
Dim lsSQL As String

'" WHERE     A.SEC_LEVEL = " & v_lSecurityLevel
'v_lSecurityLevel

    Set db_GetAuthrzBusFuncsForLevel = Nothing

    'BEGIN TRIS SOX CHANGE JULY 2004
    Select Case (v_lSecurityLevel)
        Case 9, 10, 11, 405, 981, 1045
            If Not (vbSystemAdmin) Then
                GoTo PERFORMNOTSYSADMIN
            End If
        Case Else
            '
    End Select
    'END TRIS SOX CHANGE JULY 2004

    'NOTE THE SEC_LEVEL IS -1 WHEN SELECTING BUS_FUNC_KEY A0

    'Prepare SQL Statement
    lsSQL = " SELECT    A.BUS_FUNC_KEY, B.BUS_FUNC_NM" & _
            "           , B.BUS_FUNC_LBL, B.BUS_FUNC_PRNT_KEY" & _
            "           , A.SEC_LEVEL" & _
            "           , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
            "           , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
            "           , 'L' AS RECORD_STATUS" & _
            " FROM      METHDBA.TTR_ATHRZ_BUS_FUNC A, METHDBA.TTR_BUS_FUNC B" & _
            " WHERE     A.SEC_LEVEL = ?" & _
            " AND       A.BUS_FUNC_KEY = B.BUS_FUNC_KEY" & _
            " UNION " & _
            " SELECT   BUS_FUNC_KEY, BUS_FUNC_NM" & _
            "           , BUS_FUNC_LBL, '' AS BUS_FUNC_PRNT_KEY" & _
            "           , -1 AS SEC_LEVEL" & _
            "           , 0 AS CREATED_EMP_NUM, '' AS CREATED_TIMESTAMP" & _
            "           , 0 AS LAST_MOD_EMP_NUM, '' AS LAST_MOD_TIMESTAMP" & _
            "           , 'L' AS RECORD_STATUS" & _
            " FROM      METHDBA.TTR_BUS_FUNC" & _
            " WHERE     BUS_FUNC_KEY = 'A0'" & _
            " ORDER BY  BUS_FUNC_KEY;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = v_lSecurityLevel
        .Update
    End With

    'Execute
    'Set db_GetAuthrzBusFuncsForLevel = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockBatchOptimistic)
    Set db_GetAuthrzBusFuncsForLevel = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenStatic _
                            , adLockBatchOptimistic)

PERFORMNOTSYSADMIN:
    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_GetAuthrzBusFuncsForLevel_EH:
    'Complete Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "db_GetAuthrzBusFuncsForLevel"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "db_GetAuthrzBusFuncsForLevel" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "db_GetAuthrzBusFuncsForLevel" _
                    , trErrdb_GetAuthrzBusFuncsForLevel_F _
                    , , , _
                    , True
    End Select

End Function
Public Function db_ProcessSave(ByRef r_lSecurityLevel As Long _
                    , ByVal v_adrsSecurityLevelInfo As ADODB.Recordset _
                    , ByVal v_adrsSecLvlEmployee As ADODB.Recordset _
                    , ByVal v_adrsSecLvlBusFunc As ADODB.Recordset _
                    , ByVal v_lUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo db_ProcessSave_EH
'ADODB Connection Object
Dim l_adcnConnection As ADODB.Connection
Dim l_lRecordIndex As Long, l_lSecurityLevel As Long
Dim l_sFilterString As String
Dim l_objUtility As utilDBDataServices.cuDBDataServices
Dim lsSQL As String, lsInsertSQL As String
Dim lsUpdateSQL As String

    db_ProcessSave = False

    'Initialize
    l_lSecurityLevel = -2
    lsSQL = ""

    If Not (v_adrsSecurityLevelInfo Is Nothing) Then
        With v_adrsSecurityLevelInfo
            If Not (.EOF And .BOF) Then
                If (l_lSecurityLevel = -2) Then
                    l_lSecurityLevel = !SEC_LEVEL
                End If

                Select Case (Trim(!RECORD_STATUS))
                        Case "I"
                            'Get the new security level number
                            Set l_objUtility = CreateObject( _
                                                    "utilDBDataServices.cuDBDataServices")
                            l_lSecurityLevel = l_objUtility.GetNextSequenceNumber("SL")
                            Set l_objUtility = Nothing

                            lsSQL = lsSQL & _
                                    " INSERT INTO METHDBA.TTR_SECURITY_LEVEL" & _
                                    " (SEC_LEVEL, SEC_LEVEL_DESC" & _
                                    ", DIVISION_NUM, LAST_MOD_EMP_NUM)" & _
                                    " VALUES " & _
                                    " (" & l_lSecurityLevel & ", '" & Trim(!SEC_LEVEL_DESC) & _
                                    "', '" & Trim(!DIVISION_NUM) & "', " & v_lUserEmployeeNumber & ");"
                        Case "C"
                            lsSQL = lsSQL & _
                                    " UPDATE METHDBA.TTR_SECURITY_LEVEL" & _
                                    " SET   SEC_LEVEL_DESC = '" & Trim(!SEC_LEVEL_DESC) & "'" & _
                                    "       , LAST_MOD_EMP_NUM = " & v_lUserEmployeeNumber & _
                                    "       , TIME_STAMP = CURRENT TIMESTAMP" & _
                                    " WHERE SEC_LEVEL = " & CStr(l_lSecurityLevel) & _
                                    " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"
                        Case "D"
                            'Delete authorized business functions for security level first
                            lsSQL = lsSQL & _
                                    " DELETE FROM METHDBA.TTR_ATHRZ_BUS_FUNC" & _
                                    " WHERE SEC_LEVEL = " & CStr(l_lSecurityLevel) & ";"

                            'Delete from security restrictions table
                            lsSQL = lsSQL & _
                                    " DELETE FROM METHDBA.TTR_SECURITY_RESTR" & _
                                    " WHERE SEC_LEVEL = " & CStr(l_lSecurityLevel) & ";"

                            'Update to employees to level 0
                            lsSQL = lsSQL & _
                                    " UPDATE METHDBA.TTR_EMPLOYEE" & _
                                    " SET   SEC_LEVEL = 0" & _
                                    " WHERE SEC_LEVEL = " & CStr(l_lSecurityLevel) & ";"

                            'Delete parent
                            lsSQL = lsSQL & _
                                    " DELETE FROM METHDBA.TTR_SECURITY_LEVEL" & _
                                    " WHERE SEC_LEVEL = " & CStr(l_lSecurityLevel) & _
                                    " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"
                            GoTo EXECUTEANDEXIT
                End Select
            End If
        End With
    End If

    If Not (v_adrsSecLvlEmployee Is Nothing) Then
        With v_adrsSecLvlEmployee
            If Not (.EOF And .BOF) Then
                If (l_lSecurityLevel = -2) Then
                    l_lSecurityLevel = !SEC_LEVEL
                End If

                .Filter = adFilterNone

                l_sFilterString = "RECORD_STATUS = 'D'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(l_sFilterString)
                l_sFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst

                    For l_lRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " UPDATE METHDBA.TTR_EMPLOYEE" & _
                                " SET   SEC_LEVEL = 0" & _
                                "       , TIME_STAMP = CURRENT TIMESTAMP" & _
                                "       , LAST_MOD_EMP_NUM = " & v_lUserEmployeeNumber & _
                                " WHERE EMP_NUM = " & !EMP_NUM & _
                                " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                        .MoveNext
                    Next l_lRecordIndex
                End If

                .Filter = adFilterNone
                l_sFilterString = "RECORD_STATUS = 'I'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(l_sFilterString)
                l_sFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst

                    For l_lRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " UPDATE METHDBA.TTR_EMPLOYEE" & _
                                " SET   SEC_LEVEL = " & l_lSecurityLevel & _
                                "       , TIME_STAMP = CURRENT TIMESTAMP" & _
                                "       , LAST_MOD_EMP_NUM = " & v_lUserEmployeeNumber & _
                                " WHERE EMP_NUM = " & !EMP_NUM & _
                                " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                        .MoveNext
                    Next l_lRecordIndex
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    If Not (v_adrsSecLvlBusFunc Is Nothing) Then
        With v_adrsSecLvlBusFunc
            If Not (.EOF And .BOF) Then
                If (l_lSecurityLevel = -2) Then
                    l_lSecurityLevel = !SEC_LEVEL
                End If

                .Filter = adFilterNone
                l_sFilterString = "RECORD_STATUS = 'D'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(l_sFilterString)
                l_sFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst
                    If (l_lSecurityLevel = -2) Then
                        l_lSecurityLevel = !SEC_LEVEL
                    End If

                    For l_lRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " DELETE FROM METHDBA.TTR_ATHRZ_BUS_FUNC" & _
                                " WHERE SEC_LEVEL = " & l_lSecurityLevel & _
                                " AND   BUS_FUNC_KEY = '" & !BUS_FUNC_KEY & "'" & _
                                " AND   CREATED_TIMESTAMP = '" & !CREATED_TIMESTAMP & "';"

                        .MoveNext
                    Next l_lRecordIndex
                End If

                .Filter = adFilterNone
                l_sFilterString = "RECORD_STATUS = 'I'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(l_sFilterString)
                l_sFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst

                    For l_lRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " INSERT INTO METHDBA.TTR_ATHRZ_BUS_FUNC" & _
                                " (SEC_LEVEL, BUS_FUNC_KEY" & _
                                ", CREATED_EMP_NUM, LAST_MOD_EMP_NUM" & _
                                ", CREATED_TIMESTAMP, LAST_MOD_TIMESTAMP)" & _
                                " VALUES" & _
                                " (" & l_lSecurityLevel & ", '" & !BUS_FUNC_KEY & "', " & _
                                v_lUserEmployeeNumber & ", " & v_lUserEmployeeNumber & _
                                ", CURRENT TIMESTAMP, CURRENT TIMESTAMP);"

                        .MoveNext
                    Next l_lRecordIndex
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

EXECUTEANDEXIT:
    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set l_adcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, l_adcnConnection
    End If

    'close and destroy
    If Not (l_adcnConnection Is Nothing) Then
        l_adcnConnection.Close
        Set l_adcnConnection = Nothing
    End If

    r_lSecurityLevel = l_lSecurityLevel
    db_ProcessSave = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_ProcessSave_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set l_adcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "db_ProcessSave"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "db_ProcessSave" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
'            Case trErrdb_InsSecLevel, trErrdb_UpdSecLevel _
'                , trErrdb_DelSecLevel, trErrdb_UpdEmpLevelAssignment _
'                , trErrdb_BusFuncLevelAssignment
'                TrsRaiseError MODULENAME _
'                    , "db_ProcessSave" _
'                    , Err.Number _
'                    , , , _
'                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "db_ProcessSave" _
                    , trErrdb_ProcessSave_F _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessSecurityLevelForEmployee( _
                ByVal vadrsEmployeeSecurityLevel As ADODB.Recordset _
                , ByVal vlUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo dbProcessSecurityLevelForEmployee_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection, lsSQL As String
Dim llRecordIndex As Long, lsFilterString As String

    dbProcessSecurityLevelForEmployee = False

    lsSQL = ""

    If Not (vadrsEmployeeSecurityLevel Is Nothing) Then
        With vadrsEmployeeSecurityLevel
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone
    
                lsFilterString = "RECORD_STATUS = 'D'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " UPDATE METHDBA.TTR_EMPLOYEE" & _
                                " SET   SEC_LEVEL = 0" & _
                                "       , TIME_STAMP = CURRENT TIMESTAMP" & _
                                "       , LAST_MOD_EMP_NUM = " & vlUserEmployeeNumber & _
                                " WHERE EMP_NUM = " & CStr(!EMP_NUM) & ";"

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone

                lsFilterString = "RECORD_STATUS = 'I'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " UPDATE METHDBA.TTR_EMPLOYEE" & _
                                " SET   SEC_LEVEL = " & CStr(!SEC_LEVEL) & _
                                "       , TIME_STAMP = CURRENT TIMESTAMP" & _
                                "       , LAST_MOD_EMP_NUM = " & vlUserEmployeeNumber & _
                                " WHERE EMP_NUM = " & CStr(!EMP_NUM) & ";"

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone
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

    dbProcessSecurityLevelForEmployee = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessSecurityLevelForEmployee_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbProcessSecurityLevelForEmployee"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbProcessSecurityLevelForEmployee" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
'            Case trErrdb_UpdEmpLevelAssignment
'                TrsRaiseError MODULENAME _
'                    , "dbProcessSecurityLevelForEmployee" _
'                    , Err.Number _
'                    , , , _
'                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbProcessSecurityLevelForEmployee" _
                    , trErrdb_ProcessSave_F _
                    , , , _
                    , True
    End Select

End Function
