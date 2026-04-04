VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsClientRelease"
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
'the client. This module provides the client with availability information, user existence,
'authorized business functions and other system support information. This module will call its associated
'data access component to retrieve pertinant information from the database.

Option Explicit
'Constant
Private Const MODULENAME As String = "busClientRelease.cbsClientRelease"

'This function checks the Version Number/Date Sent by the Client
'against the one retrieved form the Database
Public Function GetClientVersion() As ADODB.Recordset
On Error GoTo GetClientVersion_EH
'Data Access Object
Dim lobjdb As dbClientRelease.cdClientRelease
Dim lobjCnMgr As util_DBConnectMgr.clsuDBConnectMgr
Dim lsDSN As String

    'Construct the Connection Manager Object
    Set lobjCnMgr = CreateObject( _
                            "util_DBConnectMgr.clsuDBConnectMgr")
    'Get/Set Connection String
    lsDSN = lobjCnMgr.GetDataSourceName
    'Destroy the Connection Manager Object
    Set lobjCnMgr = Nothing

    If (InStr(1, lsDSN, "LUW") > 0) Then
        lsDSN = Trim(Left(lsDSN, 7))
    Else
        lsDSN = Trim(Left(lsDSN, 4))
    End If

    'Construct the Data Access Object
    Set lobjdb = CreateObject("dbClientRelease.cdClientRelease")
    'Call Data Access Object Method and assign return
    Set GetClientVersion = lobjdb.db_GetClientVersionInfo()
    Set lobjdb = Nothing

    If Not (GetClientVersion Is Nothing) Then
        If Not (GetClientVersion.EOF And _
            GetClientVersion.BOF) Then
            GetClientVersion.MoveFirst

            GetClientVersion!RMT_SVR_NM = Trim(Environ("COMPUTERNAME"))
            GetClientVersion!TRIS_DB2_REGION = lsDSN
            GetClientVersion.Update
        End If
    End If

    Exit Function
GetClientVersion_EH:
    Set lobjCnMgr = Nothing
    Set lobjdb = Nothing

    'Raise Error depending on the Error Number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetClientVersion"

        Case trErrdb_ExecuteSQL_F, trErrdb_GetClientVersionInfo_F _
            , trErrGetClientVersion_F
            TrsRaiseError MODULENAME _
                , "GetClientVersion" _
                , Err.Number _
                , , , , _
                True

        Case Else
            TrsRaiseError MODULENAME _
                , "GetClientVersion" _
                , trErrGetClientVersion_F _
                , , , , _
                True
    End Select

End Function
