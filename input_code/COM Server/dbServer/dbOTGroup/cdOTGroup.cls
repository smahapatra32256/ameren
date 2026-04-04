VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdOTGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'May 2004
'Author: Prem M. Chander
'Module Description: This class will provide data access
'to get/create/modify Overtime Groups

Option Explicit

Private Const MODULENAME = "dbOTGroup.cdOTGroup"
Public Function GetAllOvertimeGroups() As ADODB.Recordset
On Error GoTo GetAllOvertimeGroups_EH
Dim lsSQL As String

    Set GetAllOvertimeGroups = Nothing
'JIM B WHAT HELL ARE DOING WITH CHAR OF 34 ON OT GROUP DESCRIPTION...
'AND WHY DID YOU CHANGE THE ORDER BY CLAUSE. ASK THE USER TO SORT BY AS DESIRED??
    lsSQL = "SELECT  A.OT_GROUP_NUM" & _
        "       , CAST(LEFT(CHAR(A.LOC_CD), 3) || '-' || " & _
        "       A.OT_GROUP_DESC AS CHAR(34)) AS OT_GROUP_DESC" & _
        "       , A.LOC_CD, A.OT_STAND_BANK_CD" & _
        "       , A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM" & _
        "       , B.EMP_NAME" & _
        "       , CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM  METHDBA.TTR_OT_GROUP A" & _
        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON (A.LAST_MOD_EMP_NUM = B.EMP_NUM)" & _
        " ORDER BY 2" & _
        " FOR FETCH ONLY WITH UR;"

    'Call ExecuteSQL and retrieve information
    Set GetAllOvertimeGroups = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetAllOvertimeGroups_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetAllOvertimeGroups"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetAllOvertimeGroups" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetAllOvertimeGroups" _
                    , trErrdbGetOTGroup _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetOvertimeGroupBase(ByVal vlOTGroupNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetOvertimeGroupBase_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetOvertimeGroupBase = Nothing

    lsSQL = "SELECT  OT_GROUP_NUM, OT_GROUP_DESC " & _
        "       , LOC_CD, OT_STAND_BANK_CD " & _
        "       , CREATED_EMP_NUM, LAST_MOD_EMP_NUM " & _
        "       , CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP " & _
        "       , CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP " & _
        " FROM  METHDBA.TTR_OT_GROUP " & _
        " WHERE OT_GROUP_NUM = ?;"

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
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetOvertimeGroupBase = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetOvertimeGroupBase = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetOvertimeGroupBase_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetOvertimeGroupBase"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetOvertimeGroupBase" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetOvertimeGroupBase" _
                    , trErrdbGetOTGroup _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetOvertimeGroupBanks(ByVal vlOTGroupNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetOvertimeGroupBanks_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetOvertimeGroupBanks = Nothing

    lsSQL = "SELECT  A.OT_GROUP_NUM, A.OT_BANK_CD" & _
        "       , B.OT_BANK_DESC" & _
        "       , A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM" & _
        "       , CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM  METHDBA.TTR_OT_GROUP_BANK A" & _
        "       INNER JOIN METHDBA.TTR_OT_BANK_CD B" & _
        "       ON B.OT_BANK_CD = A.OT_BANK_CD" & _
        " WHERE A.OT_GROUP_NUM = ?" & _
        " ORDER BY A.OT_BANK_CD;"

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
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetOvertimeGroupBanks = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetOvertimeGroupBanks = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetOvertimeGroupBanks_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetOvertimeGroupBanks"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetOvertimeGroupBanks" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetOvertimeGroupBanks" _
                    , trErrdbGetOTGroup _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetOvertimeGroupWorkgroups(ByVal vlOTGroupNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetOvertimeGroupWorkgroups_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetOvertimeGroupWorkgroups = Nothing

    lsSQL = "SELECT  A.OT_GROUP_NUM, A.WG_NUM" & _
        "       , B.WG_DESC, A.ORDER_OF_CALL" & _
        "       , A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM" & _
        "       , CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "       , CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM  METHDBA.TTR_OT_GROUP_WG A" & _
        "       INNER JOIN METHDBA.TTR_WORK_GROUP B" & _
        "       ON B.WG_NUM = A.WG_NUM" & _
        " WHERE A.OT_GROUP_NUM = ?" & _
        " ORDER BY A.WG_NUM;"

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
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetOvertimeGroupWorkgroups = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetOvertimeGroupWorkgroups = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetOvertimeGroupWorkgroups_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetOvertimeGroupWorkgroups"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetOvertimeGroupWorkgroups" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetOvertimeGroupWorkgroups" _
                    , trErrdbGetOTGroup _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeesInOTGroup(ByVal vlOTGroupNumber As Long _
                , Optional ByVal vdtDate As Date = #1/1/1900# _
                ) As ADODB.Recordset
On Error GoTo GetEmployeesInOTGroup_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetEmployeesInOTGroup = Nothing

    If (vdtDate = #1/1/1900#) Then
        vdtDate = Date
    End If

    lsSQL = " SELECT  DISTINCT A.EMP_NUM, A.EMP_NAME, A.EMP_STATUS" & _
        " FROM  METHDBA.TTR_OT_GROUP_WG C " & _
        "       INNER JOIN METHDBA.TTR_EMP_WG B" & _
        "       ON C.OT_GROUP_NUM = ? " & _
        "       AND C.WG_NUM = B.WG_NUM" & _
        "       AND ? BETWEEN EFF_START_DATE AND EFF_END_DATE" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE A" & _
        "       ON B.EMP_NUM = A.EMP_NUM" & _
        " ORDER BY A.EMP_NUM" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtDate
        .Update

        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOTGroupNumber
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeesInOTGroup = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeesInOTGroup = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetEmployeesInOTGroup_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetEmployeesInOTGroup"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetEmployeesInOTGroup" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetEmployeesInOTGroup" _
                    , trErrdbGetOTGroup _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAValidOvertimeGroup(ByVal vlOTGroupNumber As Long _
                ) As Boolean
On Error GoTo IsAValidOvertimeGroup_EH
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    IsAValidOvertimeGroup = False

    lsSQL = "SELECT  OT_GROUP_NUM " & _
        " FROM  METHDBA.TTR_OT_GROUP " & _
        " WHERE OT_GROUP_NUM = ?" & _
        " FOR FETCH ONLY WITH UR;"

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
    End With

    'Call ExecuteSQL and retrieve information
'    Set ladrs = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenStatic _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    'Check if recordset is empty
    IsAValidOvertimeGroup = mADOExtensions.RecordsExist(ladrs)

    'Destroy Recordset
    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
IsAValidOvertimeGroup_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    'Destroy Recordset
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "IsAValidOvertimeGroup"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "IsAValidOvertimeGroup" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "IsAValidOvertimeGroup" _
                    , trErrdbGetOTGroup _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsWorkgroupAssociatedtoOTGroup(ByVal vlWorkgroup As Long _
                    ) As Boolean
On Error GoTo IsWorkgroupAssociatedtoOTGroup_EH
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    IsWorkgroupAssociatedtoOTGroup = False

    lsSQL = "SELECT  OT_GROUP_NUM, WG_NUM " & _
        " FROM  METHDBA.TTR_OT_GROUP_WG " & _
        " WHERE WG_NUM = ?" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlWorkgroup
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set ladrs = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
                , ladrsParms _
                , adOpenStatic _
                , adLockReadOnly)
    Set ladrsParms = Nothing

    'Check if recordset is empty
    IsWorkgroupAssociatedtoOTGroup = mADOExtensions.RecordsExist(ladrs)

    'Destroy Recordset
    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
IsWorkgroupAssociatedtoOTGroup_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    'Destroy Recordset
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "IsWorkgroupAssociatedtoOTGroup"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "IsWorkgroupAssociatedtoOTGroup" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "IsWorkgroupAssociatedtoOTGroup" _
                    , trErrdbGetOTGroup _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsOTGroupAssociatedToOpenOpportunity(ByVal vlOTGroup As Long _
                    ) As Boolean
On Error GoTo IsOTGroupAssociatedToOpenOpportunity_EH
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    IsOTGroupAssociatedToOpenOpportunity = False
'PER JIM CHECK FOR OPEN/POSTED CANVASSES IN THE LAST 30 DAYS.
'ANYTHING THAT IS OLDER THAN THAT AND IS STILL OPEN OR POSTED LET THEM DELETE
'THE OT GROUP - DEC. 15, 2004
    lsSQL = "SELECT  OT_OPP_NUM, OT_GROUP_NUM " & _
        " FROM  METHDBA.TTR_OT_OPPORTUNITY " & _
        " WHERE OT_GROUP_NUM = ?" & _
        " AND   OPP_STATUS <> 'V'" & _
        " AND   CHARGE_DATE > CURRENT DATE - 1 MONTH" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOTGroup
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set ladrs = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
                , ladrsParms _
                , adOpenStatic _
                , adLockReadOnly)
    Set ladrsParms = Nothing

    'Check if recordset is empty
    IsOTGroupAssociatedToOpenOpportunity = mADOExtensions.RecordsExist(ladrs)

    'Destroy Recordset
    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
IsOTGroupAssociatedToOpenOpportunity_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    'Destroy Recordset
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "IsOTGroupAssociatedToOpenOpportunity"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "IsOTGroupAssociatedToOpenOpportunity" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "IsOTGroupAssociatedToOpenOpportunity" _
                    , trErrdbGetOTGroup _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function dbProcessOvertimeGroup( _
            ByRef rlOTGroupNumber As Long _
            , ByVal vadrsOTGroupBase As ADODB.Recordset _
            , ByVal vadrsOTGroupBanks As ADODB.Recordset _
            , ByVal vadrsOTGroupWorkgroups As ADODB.Recordset _
            , ByVal vlUserEmployeeNumber As Long _
            ) As Boolean
On Error GoTo dbProcessOvertimeGroup_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String
'SEQ_NUM
Dim lobjUtility As utilDBDataServices.cuDBDataServices
Dim llPrevOTGroupNumber As Long

    dbProcessOvertimeGroup = False
    lsSQL = ""

    If (rlOTGroupNumber = -1) Then
        GoTo CREATENEW
    End If

    '1. If OT_GROUP is to be deleted then delete any children that may be present
    If Not (vadrsOTGroupBase Is Nothing) Then
        With vadrsOTGroupBase
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes
                    .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            'Delete all workgroups associated to the OT Group
                            lsSQL = lsSQL & _
                                " DELETE    FROM METHDBA.TTR_OT_GROUP_WG " & _
                                " WHERE     OT_GROUP_NUM = " & CStr(!KEY_OT_GROUP_NUM) & ";"

                            lsSQL = lsSQL & _
                                " DELETE    FROM METHDBA.TTR_OT_GROUP_BANK " & _
                                " WHERE     OT_GROUP_NUM = " & CStr(!KEY_OT_GROUP_NUM) & ";"

                            lsSQL = lsSQL & _
                                " DELETE    FROM    METHDBA.TTR_OT_GROUP " & _
                                " WHERE     OT_GROUP_NUM = " & CStr(!KEY_OT_GROUP_NUM) & _
                                " AND       CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                            .MoveNext
                        Next llRecordIndex

                        .Filter = adFilterNone
                        GoTo NORMALEXIT
                    End If

                    .Filter = adFilterNone
                End If
            End If
        End With
    End If

    '1. OT Group Workgroup
    If Not (vadrsOTGroupWorkgroups Is Nothing) Then
        With vadrsOTGroupWorkgroups
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes
                    .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                " DELETE    FROM METHDBA.TTR_OT_GROUP_WG " & _
                                " WHERE OT_GROUP_NUM = " & CStr(!KEY_OT_GROUP_NUM) & _
                                " AND   WG_NUM = " & CStr(!KEY_WG_NUM) & _
                                " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
            End If
        End With
    End If

    '2. Delete OT_GROUP_BANK
    If Not (vadrsOTGroupBanks Is Nothing) Then
        With vadrsOTGroupBanks
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes
                    .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                " DELETE    FROM METHDBA.TTR_OT_GROUP_BANK " & _
                                " WHERE OT_GROUP_NUM = " & CStr(!KEY_OT_GROUP_NUM) & _
                                " AND   OT_BANK_CD = " & CStr(!KEY_OT_BANK_CD) & _
                                " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
            End If
        End With
    End If

CREATENEW:
    '1. INSERT OT_GROUP_BASE
    If (rlOTGroupNumber = -1) Then
        Set lobjUtility = CreateObject( _
                            "utilDBDataServices.cuDBDataServices")
        rlOTGroupNumber = lobjUtility.GetNextSequenceNumber("OG")
        Set lobjUtility = Nothing
    End If

    If Not (vadrsOTGroupBase Is Nothing) Then
        With vadrsOTGroupBase
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for inserts
                    .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " INSERT INTO METHDBA.TTR_OT_GROUP" & _
                                    " (OT_GROUP_NUM, OT_GROUP_DESC" & _
                                    " , LOC_CD, OT_STAND_BANK_CD" & _
                                    " , CREATED_EMP_NUM, LAST_MOD_EMP_NUM" & _
                                    " , CREATED_TIMESTAMP, LAST_MOD_TIMESTAMP)" & _
                                    " VALUES " & _
                                    " (" & CStr(rlOTGroupNumber) & ", '" & Trim(!OT_GROUP_DESC) & "'" & _
                                    " , " & CStr(!LOC_CD) & ", " & CStr(!OT_STAND_BANK_CD) & _
                                    " , " & CStr(vlUserEmployeeNumber) & ", " & CStr(vlUserEmployeeNumber) & _
                                    " , CURRENT TIMESTAMP, CURRENT TIMESTAMP);"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
            End If
        End With
    End If

    '2. OT Group Bank
    If Not (vadrsOTGroupBanks Is Nothing) Then
        With vadrsOTGroupBanks
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for inserts
                    .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                " INSERT INTO METHDBA.TTR_OT_GROUP_BANK" & _
                                " (OT_GROUP_NUM, OT_BANK_CD" & _
                                " , CREATED_EMP_NUM, LAST_MOD_EMP_NUM" & _
                                " , CREATED_TIMESTAMP, LAST_MOD_TIMESTAMP)" & _
                                " VALUES " & _
                                " (" & CStr(rlOTGroupNumber) & ", " & CStr(!OT_BANK_CD) & _
                                " , " & CStr(vlUserEmployeeNumber) & ", " & CStr(vlUserEmployeeNumber) & _
                                " , CURRENT TIMESTAMP, CURRENT TIMESTAMP);"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
            End If
        End With
    End If

    '3. OT Group Workgroup
    If Not (vadrsOTGroupWorkgroups Is Nothing) Then
        With vadrsOTGroupWorkgroups
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for inserts
                    .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                " INSERT INTO METHDBA.TTR_OT_GROUP_WG" & _
                                " (OT_GROUP_NUM, WG_NUM" & _
                                " , ORDER_OF_CALL" & _
                                " , CREATED_EMP_NUM, LAST_MOD_EMP_NUM" & _
                                " , CREATED_TIMESTAMP, LAST_MOD_TIMESTAMP)" & _
                                " VALUES " & _
                                " (" & CStr(rlOTGroupNumber) & ", " & CStr(!WG_NUM) & _
                                " , " & CStr(!ORDER_OF_CALL) & _
                                " , " & CStr(vlUserEmployeeNumber) & ", " & CStr(vlUserEmployeeNumber) & _
                                " , CURRENT TIMESTAMP, CURRENT TIMESTAMP);"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
            End If
        End With
    End If

    'Check for updates
    '1. Base
    If Not (vadrsOTGroupBase Is Nothing) Then
        With vadrsOTGroupBase
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for changes
                    .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                " UPDATE    METHDBA.TTR_OT_GROUP" & _
                                " SET   OT_GROUP_DESC = '" & Trim(!OT_GROUP_DESC) & "' " & _
                                "       , LOC_CD = " & CStr(!LOC_CD) & _
                                "       , OT_STAND_BANK_CD = " & CStr(!OT_STAND_BANK_CD) & _
                                "       , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
                                " WHERE OT_GROUP_NUM = " & CStr(rlOTGroupNumber) & _
                                " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
            End If
        End With
    End If

    'NOTE: OT GROUP BANK INFORMATION CAN ONLY BE ADDED OR REMOVED NOT CHANGED.

    '2. OT Group Workgroup
    If Not (vadrsOTGroupWorkgroups Is Nothing) Then
        With vadrsOTGroupWorkgroups
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for changes
                    .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                " UPDATE    METHDBA.TTR_OT_GROUP_WG " & _
                                " SET   ORDER_OF_CALL = " & CStr(!ORDER_OF_CALL) & _
                                "       , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
                                " WHERE OT_GROUP_NUM = " & CStr(rlOTGroupNumber) & _
                                " AND   WG_NUM = " & CStr(!WG_NUM) & _
                                " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
            End If
        End With
    End If

NORMALEXIT:
    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    'Close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessOvertimeGroup = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessOvertimeGroup_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbProcessOvertimeGroup"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbProcessOvertimeGroup" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
''            Case trErrdbDelFromOTGroup, trErrdbInsIntoOTGroup _
''                , trErrdbUpdOTGroup, trErrdbDelFromOTGroupBank _
''                , trErrdbInsIntoOTGroupBank, trErrdbDelFromOTGroupWG _
''                , trErrdbInsIntoOTGroupWG, trErrdbUpdOTGroupWG
''                TrsRaiseError MODULENAME _
''                    , "dbProcessOvertimeGroup" _
''                    , Err.Number _
''                    , , , _
''                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbProcessOvertimeGroup" _
                    , trErrdbSaveOTGroup _
                    , , , _
                    , True
    End Select

End Function
Public Function ProcessOTGroupDelete(ByVal vlOTGroupNumber As Long _
                , ByVal vlUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo ProcessOTGroupDelete_EH
'NOTE: User Employee Number is passed to this method but it
'is not used. It may be used later for Audit.
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'SQL
Dim lsSQL As String

    'Delete all workgroups associated to the OT Group
    lsSQL = "DELETE FROM METHDBA.TTR_OT_GROUP_WG" & _
        " WHERE   OT_GROUP_NUM = " & CStr(vlOTGroupNumber) & ";"
    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    'Delete all banks associated to the OT Group
    lsSQL = "DELETE FROM METHDBA.TTR_OT_GROUP_BANK" & _
        " WHERE   OT_GROUP_NUM = " & CStr(vlOTGroupNumber) & ";"
    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    lsSQL = "DELETE    FROM METHDBA.TTR_OT_GROUP" & _
        " WHERE     OT_GROUP_NUM = " & CStr(vlOTGroupNumber) & ";"
    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    'Close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    ProcessOTGroupDelete = True

    'Commit Transaction
    GetObjectContext.SetComplete

    Exit Function
ProcessOTGroupDelete_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessOTGroupDelete"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "ProcessOTGroupDelete" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
''            Case trErrdbDelFromOTGroup, trErrdbDelFromOTGroupBank _
''                , trErrdbDelFromOTGroupWG, trErrdbSaveOTGroup _
''                , trErrdb_ExecuteSQL_F
''                TrsRaiseError MODULENAME _
''                    , "ProcessOTGroupDelete" _
''                    , Err.Number _
''                    , , , _
''                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessOTGroupDelete" _
                    , trErrdbSaveOTGroup _
                    , , , _
                    , True
    End Select

End Function
