VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdJCL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'October 2004
'Author:    Prem M. Chander
'Advisor:   John R. Stricklin
'Module Description: This class will provide the user the ability
'to retrieve/insert TRIS 16 bit report request information from/into
'TTR_JCL%% tables. It may also call TR427

Option Explicit

Private Const MODULENAME = "dbJCL.cdJCL"
Public Function dbProcessJCLRequest(ByVal vsProcedureName As String _
                    , ByVal vadrsPreparedParameters As ADODB.Recordset _
                    , ByVal vadrsPrinters As ADODB.Recordset _
                    , ByVal vlTRISUserNumber As Long _
                    , Optional ByVal viCopies As Integer = 1 _
                    ) As Long
On Error GoTo dbProcessJCLRequest_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String
'Sequence Number
Dim lobjUtility As utilDBDataServices.cuDBDataSvcNoTxn
Dim llJCLRequestNumber As Long

    dbProcessJCLRequest = -1
    lsSQL = ""

    'Get JCL Request Number
    Set lobjUtility = CreateObject("utilDBDataServices.cuDBDataSvcNoTxn")
    '//llJCLRequestNumber = lobjUtility.GetNextSequenceNumber("JC")
    llJCLRequestNumber = lobjUtility.GetNextJCLSequenceNumber()
    Set lobjUtility = Nothing

    lsSQL = lsSQL & _
            " INSERT INTO METHDBA.TTR_JCL_RQT" & _
            " (JCL_RQT_NUM, PROC_NAME" & _
            " , PROGRAM_ID, CREATED_EMP_NUM)" & _
            " VALUES " & _
            " (" & CStr(llJCLRequestNumber) & " , '" & Trim(vsProcedureName) & "'" & _
            " , 'TR427', " & CStr(vlTRISUserNumber) & ");"

    If Not (vadrsPrinters Is Nothing) Then
        With vadrsPrinters
            If Not (.EOF And .BOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    'Note: DEST_SEQ_NUM is zero based
                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_JCL_DEST" & _
                        " (JCL_RQT_NUM, DEST_SEQ_NUM" & _
                        " , DEST_ID, COPY_CNT " & _
                        " , PROGRAM_ID)" & _
                        " VALUES " & _
                        " (" & llJCLRequestNumber & ", " & CStr(llRecordIndex - 1) & _
                        " , '" & Trim(!PRINTER_ID) & "', " & CStr(viCopies) & _
                        " , 'TR427');"

                    .MoveNext
                Next llRecordIndex
            End If
        End With
    End If

    If Not (vadrsPreparedParameters Is Nothing) Then
        With vadrsPreparedParameters
            If Not (.EOF And .BOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    'Note: DATA_SEQ_NUM is zero based
                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_JCL_DATA" & _
                        " (JCL_RQT_NUM, DATA_SEQ_NUM" & _
                        " , DATA_VALUE, PROGRAM_ID )" & _
                        " VALUES " & _
                        " (" & llJCLRequestNumber & ", " & CStr(!DATA_SEQ_NUM) & _
                        " , '" & Trim(!DATA_VALUE) & "', 'TR427');"

                    .MoveNext
                Next llRecordIndex
            End If
        End With
    End If

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

    dbProcessJCLRequest = llJCLRequestNumber

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessJCLRequest_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing
    dbProcessJCLRequest = -1

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbProcessJCLRequest"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbProcessJCLRequest" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
''            Case trErrdbInsIntoJCLRQT, trErrdbInsIntoJCLDEST _
''                , trErrdbInsIntoJCLDATA
''                TrsRaiseError MODULENAME _
''                    , "dbProcessJCLRequest" _
''                    , Err.Number _
''                    , , , _
''                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbProcessJCLRequest" _
                    , trErrdbJCLRequest _
                    , , , _
                    , True
    End Select

End Function
Public Function IsAValidPrinter(ByVal vsPrinterID As String _
                    ) As Boolean
On Error GoTo IsAValidPrinter_EH
Dim lsSQLStatement As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    IsAValidPrinter = False
'PLEASE NOTE THAT PRINTERID PASSED TO THE FUNCTION
' IS NOT THE SAME AS PRINTER_ID IN THE CPF_TABLE
    lsSQLStatement = "SELECT  DISTINCT VPS_ID " & _
        " FROM  METHDBA.CPF_TABLE " & _
        " WHERE VPS_ID = ?" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 8
        !PARM_VALUE = vsPrinterID
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set ladrs = ExecuteSQL(lsSQLStatement _
'                    , adOpenForwardOnly _
'                    , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQLStatement _
                    , ladrsParms _
                    , adOpenForwardOnly _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    'Check if recordset is empty
    If Not (ladrs Is Nothing) Then
        If Not (ladrs.EOF And ladrs.BOF) Then
            IsAValidPrinter = True
        End If
    End If

    'Destroy Recordset
    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
IsAValidPrinter_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    'Destroy Recordset
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on
    'the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "IsAValidPrinter"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "IsAValidPrinter" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "IsAValidPrinter" _
                    , trErrdbGetPrinters _
                    , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
    End Select

End Function
Public Function AreValidPrinters(ByVal vsPrinterIDs As String _
                    ) As Integer
On Error GoTo AreValidPrinters_EH
Dim lsSQLStatement As String, ladrs As ADODB.Recordset

    AreValidPrinters = 0
'PLEASE NOTE THAT PRINTERID PASSED TO THE FUNCTION
' IS NOT THE SAME AS PRINTER_ID IN THE CPF_TABLE

    lsSQLStatement = "SELECT  DISTINCT VPS_ID " & _
        " FROM  METHDBA.CPF_TABLE " & _
        " WHERE VPS_ID IN (" & Trim(vsPrinterIDs) & ")" & _
        " FOR FETCH ONLY WITH UR;"

    'Call ExecuteSQL and retrieve information
    Set ladrs = ExecuteSQL(lsSQLStatement _
                    , adOpenForwardOnly _
                    , adLockReadOnly)

    'Check if recordset is empty
    If Not (ladrs Is Nothing) Then
        If Not (ladrs.EOF And ladrs.BOF) Then
            AreValidPrinters = ladrs.RecordCount
        End If
    End If

    'Destroy Recordset
    Set ladrs = Nothing

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
AreValidPrinters_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    'Destroy Recordset
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on
    'the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "AreValidPrinters"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "AreValidPrinters" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "AreValidPrinters" _
                    , trErrdbGetPrinters _
                    , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
    End Select

End Function
Public Function Start427() As Boolean
On Error GoTo Start427_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
Dim ladcomCommand As ADODB.Command

    Start427 = False

    'Get connection
    Set ladcnConnection = OpenConnection

    'Check if it exists
    If Not (ladcnConnection Is Nothing) Then
        'Create a command object
        Set ladcomCommand = New ADODB.Command

        'Execute
        With ladcomCommand
            .ActiveConnection = ladcnConnection
            .CommandType = adCmdStoredProc
            .CommandText = "METHDBA.TR427"
            .Prepared = True

            .Execute
        End With

        Set ladcomCommand = Nothing

        If (ladcnConnection.State = adStateOpen) Then
            ladcnConnection.Close
        End If
    End If

    Set ladcnConnection = Nothing

    Start427 = True

    GetObjectContext.SetComplete

    Exit Function
Start427_EH:
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing
    Set ladcomCommand = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on
    'the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "Start427"
            Case Else
                'check the status
                TrsRaiseError MODULENAME _
                    , "Start427" _
                    , trErrdb427 _
                    , , , _
                    , True
    End Select

End Function
Public Function CheckJCLStatus(ByVal vlRequestNumber As Long _
                    , ByVal vsProgramID As String _
                    ) As String
On Error GoTo CheckJCLStatus_EH
Dim lsSQLStatement As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset
'Status, Return, Error and SQL Codes

    CheckJCLStatus = ""

    lsSQLStatement = "SELECT    A.STATUS_CD, B.STATUS_DSC" & _
        "       , A.RETURN_CD, C.RETURN_DSC" & _
        "       , A.ERROR_CD, D.ERROR_DSC" & _
        "       , A.SQL_CD" & _
        " FROM  METHDBA.TTR_JCL_CONTROL A, METHDBA.TTR_STATUS_CD B" & _
        "       , METHDBA.TTR_RETURN_CD C, METHDBA.TTR_ERROR_CD D" & _
        " WHERE A.JCL_RQT_NUM = ?" & _
        " AND   A.PROGRAM_ID = ?" & _
        " AND   A.STATUS_CD = B.STATUS_CD" & _
        " AND   A.RETURN_CD = C.RETURN_CD" & _
        " AND   A.ERROR_CD = D.ERROR_CD" & _
        " FOR FETCH ONLY WITH UR;"
        '"for fetch only with ur -- read lock only

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlRequestNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 8
        !PARM_VALUE = vsProgramID
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set ladrs = ExecuteSQL(lsSQLStatement _
'                    , adOpenForwardOnly _
'                    , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQLStatement _
                    , ladrsParms _
                    , adOpenForwardOnly _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    'Check if recordset is empty
    If Not (ladrs Is Nothing) Then
        If Not (ladrs.EOF And ladrs.BOF) Then
            CheckJCLStatus = "Status: " & Trim(ladrs!STATUS_DSC) & " [" & CStr(ladrs!STATUS_CD) & "]; " & _
                "Return: " & Trim(ladrs!RETURN_DSC) & " [" & CStr(ladrs!RETURN_CD) & "]; " & _
                "Error: " & Trim(ladrs!ERROR_DSC) & " [" & ladrs!ERROR_CD & "]; SQL Code: " & _
                CStr(ladrs!SQL_CD) & ""
        End If
    End If

    'Destroy Recordset
    Set ladrs = Nothing

    GetObjectContext.SetComplete

    Exit Function
CheckJCLStatus_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on
    'the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "CheckJCLStatus"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "CheckJCLStatus" _
                    , trErrdb_ExecuteSQL_F _
                    , , , _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "CheckJCLStatus" _
                    , trErrdbCheckStatus _
                    , , , _
                    , True
    End Select

End Function
Public Function GetListOfPrinters() As ADODB.Recordset
On Error GoTo GetListOfPrinters_EH
Dim lsSQLStatement As String

    Set GetListOfPrinters = Nothing

    lsSQLStatement = "SELECT    DISTINCT NETNAME, VPS_ID, DESCRIPTION" & _
                    " FROM      METHDBA.CPF_TABLE" & _
                    " WHERE     VPS_ID <> ''" & _
                    " ORDER BY VPS_ID" & _
                    " FOR FETCH ONLY WITH UR;"

    'Call ExecuteSQL and retrieve information
    Set GetListOfPrinters = ExecuteSQL(lsSQLStatement _
                    , adOpenStatic _
                    , adLockReadOnly)

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
GetListOfPrinters_EH:
    GetObjectContext.SetAbort

    'Call Standard Module Function
    'Raise the appropriate error message depending on
    'the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetListOfPrinters"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetListOfPrinters" _
                    , trErrdb_ExecuteSQL_F _
                    , , , _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetListOfPrinters" _
                    , trErrdbGetPrinters _
                    , , , _
                    , True
    End Select

End Function
