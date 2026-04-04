VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdSysAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   Post TRIS 5.5 Enhancement
'Author:    Prem Chander
'Date:      June 2005
'Description:   This module provides the ability to query and modify
'TRIS Data.

'IMPORTANT: When using adUseServer for Cursor Location, it seems like only the first two
'queries of batch SQL statement get executed!!!!

Option Explicit

Private Const MODULENAME = "dbSysAdmin.cdSysAdmin"
Public Function ExecuteSysAdminRequest(ByVal vsRequest As String _
                    , ByVal vlUserNumber As Long _
                    , ByRef rlRecordsAffected As Long _
                    , ByRef rsErrorString As String _
                    , Optional ByVal vsUserID As String = "" _
                    ) As ADODB.Recordset
On Error GoTo ExecuteSysAdminRequest_EH
Dim lobjConnectionMgr As util_DBConnectMgr.clsuDBConnectMgr
'To set lock/cursor types
Dim ladrs As New ADODB.Recordset
'To set Prepared/Command Text type
Dim ladcom As New ADODB.Command
'To Capture Errors
Dim ladcn As New ADODB.Connection
Dim llRecordsAffected As Long, lsConnectionString As String

    Set lobjConnectionMgr = CreateObject( _
                            "util_DBConnectMgr.clsuDBConnectMgr")
    'Set lobjConnectionMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
    lsConnectionString = lobjConnectionMgr.GetConnectionString
    Set lobjConnectionMgr = Nothing

    With ladcn
        .CursorLocation = adUseClient
        .ConnectionString = lsConnectionString
        .Open
    End With

    With ladcom
        .ActiveConnection = ladcn
        .CommandType = adCmdText
        .CommandText = vsRequest
        .CommandTimeout = 0
'Based on research by TNV. Turning prepared on causes warnings
'to stop execution of Bulk SQL statements
'        .Prepared = True
    End With

    With ladrs
        .CursorLocation = adUseClient
        .Open ladcom _
                , _
                , adOpenStatic, adLockReadOnly
    End With

    Set ExecuteSysAdminRequest = ladrs
    Set ladrs = Nothing
    Set ladcom = Nothing
    Set ladcn = Nothing

    WriteLogFile vsRequest, lsConnectionString _
        , vlUserNumber, True

    Exit Function
ExecuteSysAdminRequest_EH:
    'ladcom.ActiveConnection.RollbackTrans
    Set lobjConnectionMgr = Nothing

    Dim lobjErr As ADODB.Error
    Dim lcolErr As New Collection
    For Each lobjErr In ladcn.Errors
        rsErrorString = rsErrorString & "Error: " & lobjErr.Description & " ;" _
            & "Native Error: " & CStr(lobjErr.NativeError) & " ;" _
            & "SQLSTATE: " & lobjErr.SQLState

       lcolErr.Add "Error: " & lobjErr.Description & " ;" _
            & "Native Error: " & CStr(lobjErr.NativeError) & " ;" _
            & "SQLSTATE: " & lobjErr.SQLState
    Next

    Set ladrs = Nothing
    Set ladcom = Nothing
    Set ladcn = Nothing

    WriteLogFile vsRequest, lsConnectionString _
        , vlUserNumber, False _
        , rsErrorString

    TrsRaiseError MODULENAME _
        , "ExecuteSysAdminRequest" _
        , , True _
        , , lcolErr _
        , True, CStr(vlUserNumber)

    Set lcolErr = Nothing

End Function
Public Function ExecuteSysAdminSELECT(ByVal vsRequest As String _
                    , ByVal vlUserNumber As Long _
                    , ByRef rsErrorString As String _
                    , Optional ByVal vsUserID As String = "" _
                    ) As ADODB.Recordset
On Error GoTo ExecuteSysAdminSELECT_EH
Dim lobjConnectionMgr As util_DBConnectMgr.clsuDBConnectMgr
'To set lock/cursor types
Dim ladrs As New ADODB.Recordset
'To Capture Errors
Dim ladcn As New ADODB.Connection
Dim lsConnectionString As String

    Set ExecuteSysAdminSELECT = Nothing

    Set lobjConnectionMgr = CreateObject( _
                            "util_DBConnectMgr.clsuDBConnectMgr")
    'Set lobjConnectionMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
    lsConnectionString = lobjConnectionMgr.GetConnectionString
    Set lobjConnectionMgr = Nothing

    With ladcn
        .CursorLocation = adUseClient
        .ConnectionString = lsConnectionString
        .Open
    End With

    With ladrs
        .CursorLocation = adUseClient
        .Open vsRequest, ladcn _
                , adOpenStatic, adLockReadOnly _
                , adCmdText
    End With

    Set ExecuteSysAdminSELECT = ladrs
    Set ladrs = Nothing
    Set ladcn = Nothing

    WriteLogFile vsRequest, lsConnectionString _
        , vlUserNumber, True

    Exit Function
ExecuteSysAdminSELECT_EH:
    Set lobjConnectionMgr = Nothing

    Dim lobjErr As ADODB.Error
    Dim lcolErr As New Collection
    For Each lobjErr In ladcn.Errors
        rsErrorString = rsErrorString & "Error: " & lobjErr.Description & " ;" _
            & "Native Error: " & CStr(lobjErr.NativeError) & " ;" _
            & "SQLSTATE: " & lobjErr.SQLState

       lcolErr.Add "Error: " & lobjErr.Description & " ;" _
            & "Native Error: " & CStr(lobjErr.NativeError) & " ;" _
            & "SQLSTATE: " & lobjErr.SQLState
    Next

    Set ladrs = Nothing
    Set ladcn = Nothing

    WriteLogFile vsRequest, lsConnectionString _
        , vlUserNumber, False _
        , rsErrorString

    TrsRaiseError MODULENAME _
        , "ExecuteSysAdminSELECT" _
        , , True _
        , , lcolErr _
        , True, CStr(vlUserNumber)

    Set lcolErr = Nothing

End Function
Public Function ExecuteSysAdminINSUPDDEL(ByVal vsRequest As String _
                    , ByVal vlUserNumber As Long _
                    , ByRef rsErrorString As String _
                    , Optional ByVal vsUserID As String = "" _
                    ) As Long
On Error GoTo ExecuteSysAdminINSUPDDEL_EH
Dim lobjConnectionMgr As util_DBConnectMgr.clsuDBConnectMgr
Dim ladcom As New ADODB.Command
Dim lsConnectionString As String, llRecordsAffected As Long

    Set lobjConnectionMgr = CreateObject( _
                            "util_DBConnectMgr.clsuDBConnectMgr")
    'Set lobjConnectionMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
    lsConnectionString = lobjConnectionMgr.GetConnectionString
    Set lobjConnectionMgr = Nothing

    With ladcom
        .CommandType = adCmdText
        .CommandText = vsRequest
'Based on research by TNV. Turning prepared on causes warnings
'to stop execution of Bulk SQL statements
'        .Prepared = True
        .CommandTimeout = 0
        .ActiveConnection = lsConnectionString
        .ActiveConnection.CursorLocation = adUseClient
        .Execute llRecordsAffected
    End With

    Set ladcom = Nothing

    ExecuteSysAdminINSUPDDEL = llRecordsAffected

    WriteLogFile vsRequest, lsConnectionString _
        , vlUserNumber, True

    Exit Function
ExecuteSysAdminINSUPDDEL_EH:
    Set lobjConnectionMgr = Nothing

    Dim lobjErr As ADODB.Error
    Dim lcolErr As New Collection
    For Each lobjErr In ladcom.ActiveConnection.Errors
        rsErrorString = rsErrorString & "Error: " & lobjErr.Description & " ;" _
            & "Native Error: " & CStr(lobjErr.NativeError) & " ;" _
            & "SQLSTATE: " & lobjErr.SQLState

       lcolErr.Add "Error: " & lobjErr.Description & " ;" _
            & "Native Error: " & CStr(lobjErr.NativeError) & " ;" _
            & "SQLSTATE: " & lobjErr.SQLState
    Next

    Set ladcom = Nothing

    WriteLogFile vsRequest, lsConnectionString _
        , vlUserNumber, False _
        , rsErrorString

    TrsRaiseError MODULENAME _
        , "ExecuteSysAdminINSUPDDEL" _
        , , True _
        , , lcolErr _
        , True, CStr(vlUserNumber)

    Set lcolErr = Nothing

End Function
Private Function WriteLogFile(ByVal vsRequest As String _
                    , ByVal vsConnectionString As String _
                    , ByVal vlUserNumber As Long _
                    , ByVal vbSuccessfulExecution As Boolean _
                    , Optional ByVal vsErrorString As String = "")
On Error GoTo WriteLogFile_EH
Dim liFileHandle As Integer, lsNow As String
Dim lsDSN As String, lsUID As String
Dim llDelimiterPos1 As Long, llDelimiterPos2 As Long

    'Get DSN.e.g if connection string is say DSN=DB2ALOC then
    'get DB2ALOC
    llDelimiterPos1 = InStr(1, vsConnectionString, ";")
    If (llDelimiterPos1 > 0) Then
        lsDSN = Mid(vsConnectionString _
                , 5 _
                , (llDelimiterPos1 - 5))
    End If
    'Get UID. e.g if connection string is say DSN=DB2ALOC,UID=MYUID,PWD=MYPWD
    'get MYUID
    llDelimiterPos2 = InStr(llDelimiterPos1 + 1, vsConnectionString, ";")
    If (llDelimiterPos2 > 0) Then
        lsUID = Mid(vsConnectionString _
                , llDelimiterPos1 + 5 _
                , (llDelimiterPos2 - (llDelimiterPos1 + 5)))
    End If

    lsNow = " " & Format(Now, "mm-dd-yyyy hh.mm.ss ampm")

    liFileHandle = FreeFile

'    If (Dir("C:\TRIS SOLUTION\DATACHANGE") <> "") Then
'        MkDir "C:\TRIS SOLUTION\DATACHANGE"
'    End If
'
    Open "C:\TRIS SOLUTION\DataChange\SALOG" & _
        vlUserNumber & _
        lsNow & _
        ".txt" _
        For Output Access Write As #liFileHandle
        
'//-----need a domain account for running the COM+ process and setup access to network share
'//-----or run IIS with anonymous off and impersonate on
'    Open "\\Corp.dir.ameren.com\DFS\GOB\Development\TRIS\MIGRATE\SYSADMIN\SALOG" & _
'        vlUserNumber & _
'        lsNow & _
'        ".txt" _
'        For Output Access Write As #liFileHandle

    Print #liFileHandle, vsRequest

    Print #liFileHandle,
    Print #liFileHandle,

    Print #liFileHandle, "Performed by: " & CStr(vlUserNumber)
    Print #liFileHandle, ", at approximately: " & lsNow
    Print #liFileHandle, ", against datasource: " & lsDSN
    Print #liFileHandle, ", using ID: " & lsUID
    Print #liFileHandle,
    If (vbSuccessfulExecution) Then
        Print #liFileHandle, "SUCCESSFUL EXECUTION"
    Else
        Print #liFileHandle, "ERRORS ENCOUNTERED"
        Print #liFileHandle, vsErrorString
    End If

    Close #liFileHandle

    Exit Function
WriteLogFile_EH:
    Close #liFileHandle

    TrsRaiseError MODULENAME _
        , "WriteLogFile" _
        , , True _
        , , , True _
        , CStr(vlUserNumber)

End Function
'CATALOG INFORMATION
'SYSIBM.SYSTABLES
Public Function GetTableCatalogInformation() As ADODB.Recordset
On Error GoTo GetTableCatalogInformation_EH
Dim lobjConnectionMgr As util_DBConnectMgr.clsuDBConnectMgr
Dim ladrs As New ADODB.Recordset
Dim lsConnectionString As String, lsSQL As String

    lsSQL = "SELECT * FROM SYSIBM.SYSTABLES" & _
        " WHERE     NAME LIKE 'TTR_%'" & _
        " AND       CREATOR = 'METHDBA'" & _
        " ORDER BY NAME" & _
        " FOR FETCH ONLY WITH UR;"

    Set lobjConnectionMgr = CreateObject( _
                            "util_DBConnectMgr.clsuDBConnectMgr")
    'Set lobjConnectionMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
    lsConnectionString = lobjConnectionMgr.GetConnectionString
    Set lobjConnectionMgr = Nothing

    ladrs.CursorLocation = adUseClient
    ladrs.Open lsSQL, lsConnectionString _
        , adOpenStatic, adLockReadOnly _
        , adCmdText

    Set GetTableCatalogInformation = ladrs
    Set ladrs = Nothing

    Exit Function
GetTableCatalogInformation_EH:
    Set lobjConnectionMgr = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "GetTableCatalogInformation" _
        , , True _
        , , , True

End Function
'SYSIBM.SYSCOLUMNS
Public Function GetColumnCatalogInformation() As ADODB.Recordset
On Error GoTo GetColumnCatalogInformation_EH
Dim lobjConnectionMgr As util_DBConnectMgr.clsuDBConnectMgr
Dim ladrs As New ADODB.Recordset
Dim lsConnectionString As String, lsSQL As String

    lsSQL = "SELECT A.NAME, A.TBNAME" & _
        "           , A.TBCREATOR , A.COLNO" & _
        "           , A.COLTYPE , A.LENGTH" & _
        "           , A.SCALE, A.NULLS" & _
        "           , A.COLCARD , A.HIGH2KEY" & _
        "           , A.LOW2KEY , A.UPDATES" & _
        "           , A.IBMREQD , A.REMARKS" & _
        "           , A.DEFAULT , A.KEYSEQ" & _
        "           , A.FOREIGNKEY , A.FLDPROC" & _
        "           , A.LABEL , A.STATSTIME" & _
        "           , A.DEFAULTVALUE , A.COLCARDF" & _
        "           , A.COLSTATUS , A.LENGTH2" & _
        "           , A.DATATYPEID , A.SOURCETYPEID" & _
        "           , A.TYPESCHEMA , A.TYPENAME" & _
        "           , A.CREATEDTS" & _
        " FROM      SYSIBM.SYSCOLUMNS A" & _
        " WHERE A.TBNAME LIKE 'TTR_%'" & _
        "   AND A.TBCREATOR = 'METHDBA'" & _
        " ORDER BY A.TBNAME, A.NAME" & _
        " FOR FETCH ONLY WITH UR;"

    Set lobjConnectionMgr = CreateObject( _
                            "util_DBConnectMgr.clsuDBConnectMgr")
    'Set lobjConnectionMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
    lsConnectionString = lobjConnectionMgr.GetConnectionString
    Set lobjConnectionMgr = Nothing

    ladrs.CursorLocation = adUseClient
    ladrs.Open lsSQL, lsConnectionString _
        , adOpenStatic, adLockReadOnly _
        , adCmdText

    Set GetColumnCatalogInformation = ladrs
    Set ladrs = Nothing

    Exit Function
GetColumnCatalogInformation_EH:
    Set lobjConnectionMgr = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "GetColumnCatalogInformation" _
        , , True _
        , , , True

End Function
'SYSIBM.SYSINDEXES
Public Function GetIndexCatalogInformation() As ADODB.Recordset
On Error GoTo GetIndexCatalogInformation_EH
Dim lobjConnectionMgr As util_DBConnectMgr.clsuDBConnectMgr
Dim ladrs As New ADODB.Recordset
Dim lsConnectionString As String, lsSQL As String

    lsSQL = "SELECT A.NAME, A.CREATOR" & _
        "       , A.TBNAME, A.TBCREATOR" & _
        "       , A.UNIQUERULE, A.COLCOUNT" & _
        "       , A.CLUSTERING, A.CLUSTERED" & _
        "       , A.DBID, A.OBID" & _
        "       , A.ISOBID, A.DBNAME" & _
        "       , A.INDEXSPACE, A.FIRSTKEYCARD" & _
        "       , A.FULLKEYCARD, A.NLEAF" & _
        "       , A.NLEVELS, A.BPOOL" & _
        "       , A.PGSIZE, A.ERASERULE" & _
        "       , A.DSETPASS, A.CLOSERULE" & _
        "       , A.SPACE, A.IBMREQD" & _
        "       , A.CLUSTERRATIO, A.CREATEDBY" & _
        "       , A.IOFACTOR, A.PREFETCHFACTOR" & _
        "       , A.STATSTIME, A.INDEXTYPE" & _
        "       , A.FIRSTKEYCARDF, A.FULLKEYCARDF" & _
        "       , A.CREATEDTS, A.ALTEREDTS" & _
        "       , A.PIECESIZE, A.COPY" & _
        "       , A.COPYLRSN, A.CLUSTERRATIOF" & _
        "       , A.SPACEF, A.REMARKS"
    lsSQL = lsSQL & _
        " FROM  SYSIBM.SYSINDEXES A" & _
        " WHERE A.TBNAME LIKE 'TTR_%'" & _
        "   AND A.TBCREATOR = 'METHDBA'" & _
        " ORDER BY A.TBNAME, A.NAME" & _
        " FOR FETCH ONLY WITH UR;"

    Set lobjConnectionMgr = CreateObject( _
                            "util_DBConnectMgr.clsuDBConnectMgr")
    'Set lobjConnectionMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
    lsConnectionString = lobjConnectionMgr.GetConnectionString
    Set lobjConnectionMgr = Nothing

    ladrs.CursorLocation = adUseClient
    ladrs.Open lsSQL, lsConnectionString _
        , adOpenStatic, adLockReadOnly _
        , adCmdText

    Set GetIndexCatalogInformation = ladrs
    Set ladrs = Nothing

    Exit Function
GetIndexCatalogInformation_EH:
    Set lobjConnectionMgr = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "GetIndexCatalogInformation" _
        , , True _
        , , , True

End Function
'SYSIBM.SYSKEYS
Public Function GetIndexKeysCatalogInformation() As ADODB.Recordset
On Error GoTo GetIndexKeysCatalogInformation_EH
Dim lobjConnectionMgr As util_DBConnectMgr.clsuDBConnectMgr
Dim ladrs As New ADODB.Recordset
Dim lsConnectionString As String, lsSQL As String

    lsSQL = "SELECT A.IXNAME, A.IXCREATOR" & _
        "       , A.COLNAME, A.COLNO" & _
        "       , A.COLSEQ, A.ORDERING" & _
        "       , A.IBMREQD" & _
        " FROM  SYSIBM.SYSKEYS A" & _
        "       INNER JOIN (SELECT NAME FROM SYSIBM.SYSINDEXES" & _
        "                   WHERE TBCREATOR = 'METHDBA' AND TBNAME LIKE 'TTR_%') B" & _
        "       ON B.NAME = A.IXNAME" & _
        " WHERE A.IXCREATOR = 'METHDBA'" & _
        " ORDER BY A.IXNAME, A.COLSEQ" & _
        " FOR FETCH ONLY WITH UR;"

    Set lobjConnectionMgr = CreateObject( _
                            "util_DBConnectMgr.clsuDBConnectMgr")
    'Set lobjConnectionMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
    lsConnectionString = lobjConnectionMgr.GetConnectionString
    Set lobjConnectionMgr = Nothing

    ladrs.CursorLocation = adUseClient
    ladrs.Open lsSQL, lsConnectionString _
        , adOpenStatic, adLockReadOnly _
        , adCmdText

    Set GetIndexKeysCatalogInformation = ladrs
    Set ladrs = Nothing

    Exit Function
GetIndexKeysCatalogInformation_EH:
    Set lobjConnectionMgr = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "GetIndexKeysCatalogInformation" _
        , , True _
        , , , True

End Function
