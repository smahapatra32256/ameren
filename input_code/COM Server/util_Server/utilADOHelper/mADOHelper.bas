Attribute VB_Name = "mADOHelper"
Option Explicit

'Module Name
Private Const MODULENAME = "utilADOHelper.mADOHelper"

'This function will return an ADODB Connection
Public Function OpenConnection() As ADODB.Connection
On Error GoTo OpenConnection_EH
'Declare an ADODB Connection Object
Dim ladcnConn As ADODB.Connection

    'Construct Connection Object
    Set ladcnConn = New ADODB.Connection
    With ladcnConn
        'Get/Set Connection String
        .ConnectionString = ConnectionString
        'Set Cursor Location
        .CursorLocation = adUseClient
        'Open
        .Open
    End With

    Set OpenConnection = ladcnConn

    'Destroy Connection Object
    Set ladcnConn = Nothing

    Exit Function
OpenConnection_EH:
'juan the swan
    'build the error collection from ladcnConn
    Dim lobjErr As ADODB.Error
    Dim lcolErr As New Collection

    For Each lobjErr In ladcnConn.Errors
       lcolErr.Add "Error: " & lobjErr.Description & " ;" _
            & "Native Error: " & CStr(lobjErr.NativeError) & " ;" _
            & "SQLSTATE: " & lobjErr.SQLState
    Next

    'Destroy Connection Object
    Set ladcnConn = Nothing

    'Call the standard module function to Raise Error
    TrsRaiseError MODULENAME _
            , "OpenConnection" _
            , , , _
            , lcolErr

End Function

'This function will return the Connection Information (like User ID, Password and Datasource Name)
Public Function ConnectionString() As String
On Error GoTo ConnectionString_EH
'Declare the Connection Manager Object
Dim lobjConnString As util_DBConnectMgr.clsuDBConnectMgr

    'Construct the Connection Manager Object
    Set lobjConnString = CreateObject( _
                            "util_DBConnectMgr.clsuDBConnectMgr")
    'Get/Set Connection String
    ConnectionString = lobjConnString.GetConnectionString
    'Destroy the Connection Manager Object
    Set lobjConnString = Nothing

    Exit Function
ConnectionString_EH:
    'Destroy Connection Manager Object
    Set lobjConnString = Nothing

    'Call the standard module function
    TrsRaiseError MODULENAME _
            , "ConnectionString"

End Function



