VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdCallLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cdCallLog"


Public Function OTStartCall( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_strNumDialed As String, _
                    ByVal v_strUserID As String _
                ) As String
' Return a 26 char timestamp representing the server time the call was initiated.

    Dim l_lngLastCallNum As Long
    Dim l_strCurrentTS As String
    Dim l_objCN As New ADODB.Connection
    Dim l_objCmd As ADODB.Command
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strOptionalCriteria As String
    Dim l_lngRecordsAffected As Long
    Dim l_objCallInfo As cdCallLogNonTxn

    On Error GoTo eh

    Set l_objCallInfo = CreateObject("dbCallLog.cdCallLogNonTxn")
    l_lngLastCallNum = l_objCallInfo.GetLastCallNum(v_lngOTOppNum, v_lngEmpNum)
    l_strCurrentTS = l_objCallInfo.GetCurrentTimestamp
    Set l_objCallInfo = Nothing
    
    Set l_objCN = New ADODB.Connection
    Set l_objCmd = New ADODB.Command
    Set l_objCM = New clsuDBConnectMgr
    
    l_strSQL = "INSERT INTO METHDBA.TTR_OT_OPP_CALLOUT " & _
        "(OT_OPP_NUM, EMP_NUM, CALL_NUM," & _
        " CALL_ST_TIMESTAMP, CALL_END_TIMESTAMP," & _
        " DIALED_NUM, CREATED_EMP_NUM, CREATED_TIMESTAMP)" & _
        " VALUES (" & _
        CStr(v_lngOTOppNum) & ", " & CStr(v_lngEmpNum) & ", " & CStr(l_lngLastCallNum + 1) & _
        ", '" & l_strCurrentTS & "', '" & l_strCurrentTS & "', '" & v_strNumDialed & "', " & _
        v_strUserID & ", '" & l_strCurrentTS & "')"

    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open
    With l_objCmd
        .ActiveConnection = l_objCN
        .CommandText = l_strSQL
        .CommandType = adCmdUnknown
        .Execute l_lngRecordsAffected
    End With

    l_objCN.Close
    Set l_objCN = Nothing
    Set l_objCmd = Nothing
    GetObjectContext.SetComplete
    
    OTStartCall = l_strCurrentTS
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCallInfo = Nothing
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objCmd = Nothing
    TrsRaiseError m_strCMODULENAME, "OTStartCall", , , , DescArgsFrom(l_strSQL)

End Function


Public Sub OTEndCall( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_strCallStartID As String, _
                    ByVal v_strUserID As String)

    Dim l_objCN As New ADODB.Connection
    Dim l_objCmd As ADODB.Command
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strOptionalCriteria As String
    Dim l_lngRecordsAffected As Long

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objCmd = New ADODB.Command
    Set l_objCM = New clsuDBConnectMgr
    
    l_strSQL = "UPDATE METHDBA.TTR_OT_OPP_CALLOUT" & _
        " SET CALL_END_TIMESTAMP = CURRENT TIMESTAMP" & _
        " WHERE OT_OPP_NUM = " & v_lngOTOppNum & _
        " AND EMP_NUM = " & v_lngEmpNum & _
        " AND CALL_ST_TIMESTAMP = '" & v_strCallStartID & "'"

    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open
    With l_objCmd
        .ActiveConnection = l_objCN
        .CommandText = l_strSQL
        .CommandType = adCmdUnknown
        .Execute l_lngRecordsAffected
    End With

    l_objCN.Close
    Set l_objCN = Nothing
    Set l_objCmd = Nothing
    GetObjectContext.SetComplete
    
    Exit Sub

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objCmd = Nothing
    TrsRaiseError m_strCMODULENAME, "OTEndCall", , , , DescArgsFrom(l_strSQL)

End Sub



