VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdClientRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'PLEASE DON'T BREAK BINARY COMPATIBILITY
'PLEASE DON'T BREAK BINARY COMPATIBILITY
'PLEASE DON'T BREAK BINARY COMPATIBILITY
'PLEASE DON'T BREAK BINARY COMPATIBILITY
'PLEASE DON'T BREAK BINARY COMPATIBILITY

'Project:   TRIS Conversion Phase 2
'Author:    Prem M. Chander
'Date:      June 2002
'Description: This class module will reside in a Server Package
'under Microsoft Transaction Server (MTS). This module will be called by
'its associated business component. This module will access the database for
'pertinant information

Option Explicit

'Declare Constant
Private Const MODULENAME As String = "dbClientRelease.cdClientRelease"

'This function retrieves the Client Version Information
'PARM SUFFIX 1 IS THE VERSION NUMBER
'PARM_SUFFIX 2 IS THE VERSION DATE
'PARM_NUM 17 IS FOR TRIS 32 BIT CLIENT VERSION INFORMATION
Public Function db_GetClientVersionInfo() As ADODB.Recordset
On Error GoTo db_GetClientVersionInfo_EH
'Declare a variable for SQL Statement
Dim lsSQL As String

    Set db_GetClientVersionInfo = Nothing

    'Set SQL SELECT statment
    lsSQL = "SELECT PARM_SUFFIX, PARM_VALUE, PARM_DESC" & _
            "       , '' AS RMT_SVR_NM, '' AS TRIS_DB2_REGION" & _
            " FROM  METHDBA.TTR_SYSPARMS" & _
            " WHERE PARM_NUM = 17" & _
            " ORDER BY   PARM_SUFFIX" & _
            " FOR FETCH ONLY WITH UR;"

    Set db_GetClientVersionInfo = ExecuteSQL(lsSQL _
                                , adOpenStatic _
                                , adLockBatchOptimistic)
    Exit Function
db_GetClientVersionInfo_EH:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "db_GetClientVersionInfo"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "db_GetClientVersionInfo" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "db_GetClientVersionInfo" _
                , trErrdb_GetClientVersionInfo_F _
                , , , _
                , True
    End Select

End Function
