VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdCallLogNonTxn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cdCallLogNonTxn"


Public Function GetLastCallNum( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_lngEmpNum As Long _
                ) As Long

    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT MAX(CALL_NUM) AS CALL_NUM FROM METHDBA.TTR_OT_OPP_CALLOUT" & _
        " WHERE OT_OPP_NUM = " & v_lngOTOppNum & _
        " AND EMP_NUM = " & v_lngEmpNum & _
        " GROUP BY OT_OPP_NUM, EMP_NUM"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    If l_objRS.EOF Then
        GetLastCallNum = 0
    Else
        GetLastCallNum = l_objRS!CALL_NUM
    End If
    
    Set l_objRS = Nothing
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetLastCallNum", , , , DescArgsFrom(l_strSQL)

End Function


Public Function GetCurrentTimestamp() As String

    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT CHAR(CURRENT TIMESTAMP) as CURRENT_TIMESTAMP" & _
                " FROM SYSIBM.SYSDUMMY1"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    If l_objRS.EOF Then
        Err.Raise 30500, m_strCMODULENAME, "GetCurrentTimestamp failed"
    Else
        GetCurrentTimestamp = l_objRS!CURRENT_TIMESTAMP
    End If
    
    Set l_objRS = Nothing
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCurrentTimestamp", , , , DescArgsFrom(l_strSQL)

End Function
