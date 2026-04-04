VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdFunSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Conversion Phase 1
'Author:    Prem M. Chander
'Date:      December 1999 (Firstcut)
'Description: This class module will reside in a Server Package
'under Microsoft Transaction Server (MTS). This module will be called by
'its associated business component. This module will access the database for
'pertinant information
'Notes: Mr. Eric Lynn suggested that I use ALL CAPS FOR MODULE LEVEL CONSTANTS,
'       Call Functions to GetConnectionString, OpenConnection from within the
'       ExecuteSQL function. He is okay with passing object context within module functions
'       however he suggests using GetObjectContext rather than a variable(and parameter).

Option Explicit

'Declare Constant
Private Const MODULENAME As String = "dbSecurity.cdFunSecurity"
'Given an employee number defined as Long, the function will return whether the
'employee exist or not
Public Function db_CheckEmployee(ByVal v_lEmpNum As Long _
                ) As ADODB.Recordset
On Error GoTo db_CheckEmployee_EH
Dim ladrs As ADODB.Recordset
'Declare varaible for SQL Statement
Dim lsSQL As String

'" AND       A.EMP_NUM = " & CStr(v_lEmpNum)

    Set db_CheckEmployee = Nothing

    'Set SQL SELECT statment
    lsSQL = "SELECT A.EMP_NUM, A.EMP_NAME" & _
            "       , A.SEC_LEVEL, B.RELEASE" & _
            "       , B.RELEASE_EFF_DATE, A.DIVISION_NUM" & _
            "       , A.FUNCTION_NUM, A.DEPT_NUM" & _
            "       , A.LOC_CD, A.EMP_STATUS" & _
            "       , A.STATUS_EFF_DATE, A.UNION_CD" & _
            "       , A.PREV_DEPT_NUM, A.JOB_NUM" & _
            "       , A.PRIOR_JOB_NUM" & _
            "       , CASE (COALESCE(A.JOB_EFF_DATE, DATE('0100-01-01')))" & _
            "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
            "       ELSE COALESCE(A.JOB_EFF_DATE, DATE('0100-01-01')) END AS JOB_EFF_DATE" & _
            "       , CASE (COALESCE(A.DEPT_EFF_DATE, DATE('0100-01-01')))" & _
            "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
            "       ELSE COALESCE(A.DEPT_EFF_DATE, DATE('0100-01-01')) END AS DEPT_EFF_DATE" & _
            "       , CASE (COALESCE(A.PRIOR_JOB_EFF_DT, DATE('0100-01-01')))" & _
            "       WHEN DATE('0001-01-01') THEN DATE('0100-01-01')" & _
            "       ELSE COALESCE(A.PRIOR_JOB_EFF_DT, DATE('0100-01-01')) END AS PRIOR_JOB_EFF_DT" & _
            " FROM  METHDBA.TTR_EMPLOYEE A, METHDBA.TTR_DEPT B" & _
            " WHERE A.EMP_NUM = ?" & _
            " AND   A.DEPT_NUM = B.DEPT_NUM" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = v_lEmpNum
        .Update
    End With

    'Call Execute SQL
    'Set db_CheckEmployee = ExecuteSQL(lsSQL _
                            , adOpenForwardOnly _
                            , adLockReadOnly)
    Set db_CheckEmployee = NewExecuteSQL(lsSQL _
                            , ladrs _
                            , adOpenForwardOnly _
                            , adLockReadOnly)

    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_CheckEmployee_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "db_CheckEmployee"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "db_CheckEmployee" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "db_CheckEmployee" _
                    , trErrdb_CheckEmployee_F _
                    , , , _
                    , True
    End Select

End Function
'This function determines if TRIS Database is available for TRIS Users
Public Function db_TRISAvailable() As Boolean
On Error GoTo db_TRISAvailable_EH
'Declare a variable for SQL Statement
Dim lsSQL As String
'Declare an ADODB Recordset object
Dim ladrsTRISAvail As ADODB.Recordset

    db_TRISAvailable = False

    'Set SQL SELECT statment
    lsSQL = "SELECT PARM_VALUE" & _
            " FROM  METHDBA.TTR_SYSPARMS" & _
            " WHERE PARM_NUM = 8 AND PARM_SUFFIX = 0" & _
            " FOR FETCH ONLY WITH UR;"

    'Call Execute SQL
    Set ladrsTRISAvail = ExecuteSQL(lsSQL _
                            , adOpenForwardOnly _
                            , adLockReadOnly)

    'Check to see if any records are returned.
    If Not (ladrsTRISAvail Is Nothing) Then
        If Not (ladrsTRISAvail.BOF And _
            ladrsTRISAvail.EOF) Then
            If (Trim(ladrsTRISAvail!PARM_VALUE) = "N") Then
                'Return True (TRIS is available for the Users)
                db_TRISAvailable = True
            End If
        End If
    End If

    'Destroy Recordset
    Set ladrsTRISAvail = Nothing

    'Abort Transaction
    GetObjectContext.SetComplete

    Exit Function
db_TRISAvailable_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    'Destroy Recordset
    Set ladrsTRISAvail = Nothing

    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "db_TRISAvailable"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "db_TRISAvailable" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "db_TRISAvailable" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select

End Function
'This function determines those TRIS menu items that are available
'for a given Security Level
Public Function db_GetMenuItems(ByVal v_lSecLvl As Long _
                ) As ADODB.Recordset
On Error GoTo db_GetMenuItems_EH
Dim ladrs As ADODB.Recordset
'Declare a variable for SQL Statement
Dim lsSQL As String

'" WHERE A.SEC_LEVEL = " & v_lSecLvl

    Set db_GetMenuItems = Nothing

    lsSQL = "SELECT B.BUS_FUNC_NM, B.BUS_FUNC_LBL" & _
            "       , A.BUS_FUNC_KEY, B.BUS_FUNC_PRNT_KEY" & _
            "       , A.SEC_LEVEL" & _
            " FROM  METHDBA.TTR_ATHRZ_BUS_FUNC A, METHDBA.TTR_BUS_FUNC B" & _
            " WHERE A.SEC_LEVEL = ?" & _
            " AND   A.BUS_FUNC_KEY = B.BUS_FUNC_KEY" & _
            " ORDER BY  BUS_FUNC_KEY" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = v_lSecLvl
        .Update
    End With

    'Set db_GetMenuItems = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockBatchOptimistic)
    Set db_GetMenuItems = NewExecuteSQL(lsSQL _
                        , ladrs _
                        , adOpenStatic _
                        , adLockBatchOptimistic)
    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_GetMenuItems_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "db_GetMenuItems"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "db_GetMenuItems" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "db_GetMenuItems" _
                    , trErrdb_GetMenuItems_F _
                    , , , _
                    , True
    End Select

End Function
Public Function db_GetCustomerInfo(ByVal v_sCustomerID As String _
                ) As ADODB.Recordset
On Error GoTo db_GetCustomerInfo_EH
Dim ladrs As ADODB.Recordset
'Declare a variable for SQL Statement
Dim lsSQL As String
'" WHERE USER_ID = '" & v_sCustomerID & "'"

    Set db_GetCustomerInfo = Nothing

    'Set SQL SELECT statment
    lsSQL = "SELECT USER_NUM, CURR_MACH_ID" & _
            "       , CURR_RMT_SRVR_NM, CHAR(CURR_TIMESTAMP) AS CURR_TIMESTAMP" & _
            " FROM  METHDBA.TTR_TRIS_USER" & _
            " WHERE USER_ID = ?;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 18
        !PARM_VALUE = v_sCustomerID
        .Update
    End With

    'Set db_GetCustomerInfo = ExecuteSQL(lsSQL _
                                , adOpenStatic _
                                , adLockReadOnly)
    Set db_GetCustomerInfo = NewExecuteSQL(lsSQL _
                                , ladrs _
                                , adOpenStatic _
                                , adLockReadOnly)

    Set ladrs = Nothing

    'Commit Transaction
    GetObjectContext.SetComplete

    Exit Function
db_GetCustomerInfo_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "db_GetCustomerInfo"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "db_GetCustomerInfo" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL, "v_sCustomerID=" & v_sCustomerID) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "db_GetCustomerInfo" _
                    , trErrdb_GetCltInfo _
                    , , , _
                    , True
    End Select

End Function
Public Function db_InsCustomerInfo(ByVal v_sCustomerID As String _
            , ByVal v_lCustomerNumber As Long _
            , ByVal v_sClientMachineID As String _
            , ByVal v_sRemoteServerName As String _
            , ByVal v_sClientVersionNumber As String _
            , ByVal v_sClientEdition As String _
            ) As Boolean
On Error GoTo db_InsCustomerInfo_EH
Dim lsSQL As String, ladrs As ADODB.Recordset
'ADODB Connection Object
Dim ladcn As ADODB.Connection

'    lsSQL = "INSERT INTO METHDBA.TTR_TRIS_USER" & _
            " (USER_ID, USER_NUM" & _
            " , CURR_MACH_ID, CURR_RMT_SRVR_NM" & _
            " , CLNT_VER_NUM, CLNT_EDTN_CD)" & _
            " VALUES" & _
            " ('" & v_sCustomerID & "', " & v_lCustomerNumber & _
            " , '" & v_sClientMachineID & "', '" & v_sRemoteServerName & _
            " ', '" & v_sClientVersionNumber & "', '" & v_sClientEdition & "');"
    lsSQL = "INSERT INTO METHDBA.TTR_TRIS_USER" & _
            " (USER_ID, USER_NUM" & _
            " , CURR_MACH_ID, CURR_RMT_SRVR_NM" & _
            " , CLNT_VER_NUM, CLNT_EDTN_CD)" & _
            " VALUES" & _
            " (?, ?" & _
            " , ?, ?" & _
            " , ?, ?);"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 18
        !PARM_VALUE = v_sCustomerID
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = v_lCustomerNumber
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 18
        !PARM_VALUE = v_sClientMachineID
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 18
        !PARM_VALUE = v_sRemoteServerName
        .Update
    
        .AddNew
        !PARM_ORDER = 5
        !PARM_NAME = ""
        !PARM_TYPE = adVarChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 12
        !PARM_VALUE = v_sClientVersionNumber
        .Update

        .AddNew
        !PARM_ORDER = 6
        !PARM_NAME = ""
        !PARM_TYPE = adVarChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 5
        !PARM_VALUE = v_sClientEdition
        .Update
    End With

    'Get Connection
    Set ladcn = OpenConnection()
    'Call Routine to Execute SQL
'    If (ExecuteOTHERSQL(lsSQL, ladcn) = 0) Then
'        Err.Raise trErrdb_InsCltInfo
'    End If
    If (NewExecuteOtherSQL(lsSQL, ladrs _
        , ladcn, adCmdText) = 0) Then
        Err.Raise trErrdb_InsCltInfo
    End If

    Set ladrs = Nothing

    'close and destroy
    If Not (ladcn Is Nothing) Then
        ladcn.Close

        Set ladcn = Nothing
    End If

    db_InsCustomerInfo = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_InsCustomerInfo_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing
    Set ladcn = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "db_InsCustomerInfo"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "db_InsCustomerInfo" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "db_InsCustomerInfo" _
                    , trErrdb_InsCltInfo _
                    , , , _
                    , True
    End Select

End Function
Public Function db_UpdCustomerInfo(ByVal v_sCustomerID As String _
                , ByVal v_lCustomerNumber As Long _
                , ByVal v_sClientMachineID As String _
                , ByVal v_sRemoteServerName As String _
                , ByVal v_sPrevSessTimeStamp As String _
                , ByVal v_sPrevSessMachineID As String _
                , ByVal v_sPrevSessRSN As String _
                , ByVal v_sClientVersionNumber As String _
                , ByVal v_sClientEdition As String _
                ) As Boolean
On Error GoTo db_UpdCustomerInfo_EH
Dim lsSQL As String, ladrs As ADODB.Recordset
'ADODB Connection Object
Dim ladcn As ADODB.Connection

'    lsSQL = "UPDATE METHDBA.TTR_TRIS_USER" & _
            " SET   PREV_TIMESTAMP = '" & v_sPrevSessTimeStamp & "'" & _
            "       , PREV_MACH_ID = '" & v_sPrevSessMachineID & "'" & _
            "       , PREV_RMT_SRVR_NM = '" & v_sPrevSessRSN & "'" & _
            "       , USER_NUM = " & v_lCustomerNumber & _
            "       , CURR_MACH_ID = '" & v_sClientMachineID & "'" & _
            "       , CURR_RMT_SRVR_NM = '" & v_sRemoteServerName & "'" & _
            "       , CURR_TIMESTAMP = CURRENT TIMESTAMP" & _
            "       , CLNT_VER_NUM = '" & v_sClientVersionNumber & "'" & _
            "       , CLNT_EDTN_CD = '" & v_sClientEdition & "'" & _
            " WHERE USER_ID = '" & v_sCustomerID & "';"
    lsSQL = "UPDATE METHDBA.TTR_TRIS_USER" & _
            " SET   PREV_TIMESTAMP = ?" & _
            "       , PREV_MACH_ID = ?" & _
            "       , PREV_RMT_SRVR_NM = ?" & _
            "       , USER_NUM = ?" & _
            "       , CURR_MACH_ID = ?" & _
            "       , CURR_RMT_SRVR_NM = ?" & _
            "       , CURR_TIMESTAMP = CURRENT TIMESTAMP" & _
            "       , CLNT_VER_NUM = ?" & _
            "       , CLNT_EDTN_CD = ?" & _
            " WHERE USER_ID = ?;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 26
        !PARM_VALUE = v_sPrevSessTimeStamp
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 18
        !PARM_VALUE = v_sPrevSessMachineID
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 18
        !PARM_VALUE = v_sPrevSessRSN
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = v_lCustomerNumber
        .Update

        .AddNew
        !PARM_ORDER = 5
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 18
        !PARM_VALUE = v_sClientMachineID
        .Update

        .AddNew
        !PARM_ORDER = 6
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 18
        !PARM_VALUE = v_sRemoteServerName
        .Update
    
        .AddNew
        !PARM_ORDER = 7
        !PARM_NAME = ""
        !PARM_TYPE = adVarChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 12
        !PARM_VALUE = v_sClientVersionNumber
        .Update

        .AddNew
        !PARM_ORDER = 8
        !PARM_NAME = ""
        !PARM_TYPE = adVarChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 5
        !PARM_VALUE = v_sClientEdition
        .Update

        .AddNew
        !PARM_ORDER = 9
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 18
        !PARM_VALUE = v_sCustomerID
        .Update
    End With

    'Get Connection
    Set ladcn = OpenConnection()
    'Call Routine to Execute SQL
'    If (ExecuteOTHERSQL(lsSQL, ladcn) = 0) Then
'        Err.Raise trErrdb_UpdCltInfo
'    End If
    If (NewExecuteOtherSQL(lsSQL, ladrs _
        , ladcn, adCmdText) = 0) Then
        Err.Raise trErrdb_UpdCltInfo
    End If
    Set ladrs = Nothing

    'close and destroy
    If Not (ladcn Is Nothing) Then
        ladcn.Close

        Set ladcn = Nothing
    End If

    db_UpdCustomerInfo = True
    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
db_UpdCustomerInfo_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing
    Set ladcn = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "db_UpdCustomerInfo"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "db_UpdCustomerInfo" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "db_UpdCustomerInfo" _
                    , trErrdb_UpdCltInfo _
                    , , , _
                    , True
    End Select

End Function
Public Function GetAlias(ByVal v_lCustomerNumber As Long _
                ) As ADODB.Recordset
On Error GoTo eh
Dim ladrs As ADODB.Recordset
Dim lsSQL As String
'" AND   INTEGER(SUBSTR(RG_VALUE_CD,1,9)) = " & str(v_lCustomerNumber)

    'Set SQL SELECT statment
    lsSQL = " SELECT CASE" & _
            " WHEN SUBSTR(RG_VALUE_CD,11,19) BETWEEN '000000000' AND '999999999'" & _
            " THEN INTEGER(SUBSTR(RG_VALUE_CD,11,19))" & _
            " ELSE 0 END AS ALIAS_EMP_NUM" & _
            " FROM  METHDBA.TTR_RG_GROUP " & _
            " WHERE RG_CD = 7 " & _
            " AND   BUS_TYPE_CD = 2 " & _
            " AND   INTEGER(SUBSTR(RG_VALUE_CD,1,9)) = ?" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = v_lCustomerNumber
        .Update
    End With

    'Set GetAlias = ExecuteSQL(lsSQL _
                                , adOpenStatic _
                                , adLockReadOnly)
    Set GetAlias = NewExecuteSQL(lsSQL _
                    , ladrs _
                    , adOpenStatic _
                    , adLockReadOnly)

    Set ladrs = Nothing

    GetObjectContext.SetComplete

    Exit Function
eh:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladrs = Nothing

    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetAlias"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetAlias" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetAlias" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select

End Function
