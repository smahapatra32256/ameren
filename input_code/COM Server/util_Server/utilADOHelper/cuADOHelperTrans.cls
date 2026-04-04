VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cuADOHelperTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const MODULENAME As String = "cuADOHelperTrans"

Private m_ADO As utilADOHelper.cuADO

Public Property Set ADO(ByVal v_ADO As utilADOHelper.cuADO)
    Set m_ADO = v_ADO
End Property
Public Property Get ADO() As utilADOHelper.cuADO
    Set ADO = m_ADO
End Property

Public Function ExecuteSQL(ByVal vsSQL As String _
                , Optional ByVal vCommandType As ADODB.CommandTypeEnum = -1 _
                ) As Long
On Error GoTo eh
    Dim blnCreateConnection As Boolean
    Dim llIndex As Long, llRecordsAffected As Long

    '//-----set defaults for optional parms
    If vCommandType = -1 Then
        vCommandType = adCmdText
    End If
    
    If ADO.Connection Is Nothing Then
        ADO.CreateConnection
        blnCreateConnection = True
    Else
        blnCreateConnection = False
    End If
               
    'Command Object
    If ADO.Command Is Nothing Then
        ADO.CreateCommand
    End If
    
    'Set command object connection
    ADO.Command.Prepared = False
    ADO.Command.CommandText = vsSQL
    ADO.Command.ActiveConnection = ADO.Connection
        
    ADO.Command.Execute llRecordsAffected
    If llRecordsAffected = -1 Or llRecordsAffected = 0 Then
         Err.Raise trErrAppCommitConflict
    End If
    
    'close connection
    If blnCreateConnection Then
        'Destroy Connection Object
        ADO.DestroyConnection
    End If
    
    ExecuteSQL = llRecordsAffected

    Exit Function
eh:
    'close connection
    If blnCreateConnection Then
        'Destroy Connection Object
        ADO.DestroyConnection
    End If
    'Call the standard module function to Raise Error
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME, "ExecuteSQL"
        Case -2147467259, trErrAppCommitConflict '[IBM][CLI Driver][DB2] SQL0100W  No row was found for FETCH, UPDATE or DELETE; or the result of a query is an empty table.  SQLSTATE=02000
            Err.Raise trErrAppCommitConflict
        Case Else
            TrsRaiseError MODULENAME, "ExecuteSQL", trErrdb_ExecuteSQL_F, , , DescArgsFrom(ADO.DeriveSQL())
    End Select
End Function

Public Function ExecuteSelect(ByVal vsSQL As String _
                , Optional ByVal vCursorType As ADODB.CursorTypeEnum = -1 _
                , Optional ByVal vLockType As ADODB.LockTypeEnum = -1 _
                , Optional ByVal vCommandType As ADODB.CommandTypeEnum = -1 _
                ) As ADODB.Recordset
On Error GoTo eh
    Dim blnCreateConnection As Boolean
    '//Dim oStopWatch As TRISServerUtilityCOM.StopWatch
    '//Dim oSQLTrace As TRISServerUtilityCOM.Trace
    '//Dim dblElapsedTime As Double
    Dim sDeriveSQL As String
    Dim rc As Long
    
    '//Set oStopWatch = New TRISServerUtilityCOM.StopWatch
    '//Set oSQLTrace = New TRISServerUtilityCOM.Trace
        
    Set ExecuteSelect = Nothing
    '//-----set defaults for optional parms (changing defaults requires generating new guid, so set the default in code)
    If vCursorType = -1 Then
        vCursorType = adOpenStatic
    End If
    If vLockType = -1 Then
        vLockType = adLockBatchOptimistic
    End If
    If vCommandType = -1 Then
        vCommandType = adCmdText
    End If
   
    If ADO.Connection Is Nothing Then
        ADO.CreateConnection
        blnCreateConnection = True
    Else
        blnCreateConnection = False
    End If
        
    'Command Object
    If ADO.Command Is Nothing Then
        'Set m_ADO = New ADODB.Command
        ADO.CreateCommand
    End If
    
    ADO.Command.CommandType = vCommandType
    ADO.Command.CommandText = vsSQL
        
    Set ExecuteSelect = New ADODB.Recordset
    With ExecuteSelect
        .CursorLocation = adUseClient
        .CursorType = vCursorType
        .LockType = vLockType

        'Set command object connection
        ADO.Command.ActiveConnection = ADO.Connection
        '.ActiveConnection = m_ADOConnection

        '//oStopWatch.StartCount
        .Open ADO.Command
        
        '//oStopWatch.StopCount
        '//dblElapsedTime = oStopWatch.GetCount
        sDeriveSQL = ADO.DeriveSQL()
        
        '//oSQLTrace.WriteLineSQLTrace sDeriveSQL, dblElapsedTime, .RecordCount, ADO.Connection.ConnectionString
        'Datascope query using code in mstdDBOPERATIONS was not returning any results and no errors
        'select count(*) from TTR_APPROVAL
        'an error was raised when RecordCount was evaluated, this would population the ActiveConnection.Errors collection
        'TODO: What other errors are not being raised by db2?
        
'        If .RecordCount > 0 Then 'cause hidden errors like unavailable resource error to be thrown
'        End If
'        If .ActiveConnection.Errors.Count > 0 Then
'            Dim lcon As ADODB.Connection
'            Dim lerr As ADODB.Error
'
'            Set lcon = .ActiveConnection
'            Set lerr = lcon.Errors(0)
'            Err.Raise lerr.Number, "ExecuteSelect", lerr.Description
''            MsgBox lerr.Description & vbCrLf & lerr.Number
''            lcon.Errors.Item(0).Number
'        End If
        
        .ActiveConnection = Nothing
    End With

    'close connection
    If blnCreateConnection Then
        'Destroy Connection Object
        ADO.DestroyConnection
    End If
    
    '//Set oStopWatch = Nothing
    '//Set oSQLTrace = Nothing
    Exit Function
eh:
    'Call the standard module function to Raise Error
    '//Set oStopWatch = Nothing
    '//Set oSQLTrace = Nothing
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "ExecuteSelect"
            Case Else
                TrsRaiseError MODULENAME, "ExecuteSelect", trErrdb_ExecuteSQL_F, , , DescArgsFrom(ADO.DeriveSQL())
    End Select
End Function

Private Sub Class_Initialize()
    Set ADO = New utilADOHelper.cuADO
End Sub

Private Sub Class_Terminate()
    ADO.Destroy
    Set ADO = Nothing
End Sub

