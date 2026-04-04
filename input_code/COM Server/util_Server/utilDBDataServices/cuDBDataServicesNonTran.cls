VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cuDBDataSvcNoTxn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "utilDBDataServices.cuDBDataSvcNoTxn"

Public Function GetTimeStamp() As String

    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetTimeStamp_eh
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    l_strSQL = "select CHAR (CURRENT TIMESTAMP) as CURRENT_TIMESTAMP " & _
                " from methdba.ttr_sysparms " & _
                " where parm_num = 2 and parm_suffix = 1 "
              
    Set l_objRS = New ADODB.Recordset
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    
    If l_objRS.EOF And l_objRS.BOF Then
        Err.Raise trErrutGetTimeStamp
    End If
        
    GetTimeStamp = l_objRS!CURRENT_TIMESTAMP
    
    l_objRS.ActiveConnection = Nothing
    Set l_objRS = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    
    
    Exit Function

GetTimeStamp_eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    
    Select Case Err.Number
        Case trErrutGetTimeStamp
            TrsRaiseError m_strCMODULENAME, "GetTimeStamp", trErrutGetTimeStamp, True, , DescArgsFrom(l_strSQL), True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetTimeStamp", , True, , DescArgsFrom(l_strSQL), True
    End Select
    
End Function


Public Function GetNextAdjustmentSequenceNumber() As Long
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    l_strSQL = "SELECT NEXT VALUE FOR METHDBA.TTR_ADJUST_TIME_SEQ AS SEQ_NUM " & _
                "FROM SYSIBM.SYSDUMMY1"
              
    Set l_objRS = New ADODB.Recordset
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    
'    If l_objRS.EOF And l_objRS.BOF Then
'        Err.Raise trErrutGetTimeStamp
'    End If
    
    GetNextAdjustmentSequenceNumber = l_objRS!SEQ_NUM
    
    l_objRS.ActiveConnection = Nothing
    Set l_objRS = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    
    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
       
    TrsRaiseError m_strCMODULENAME, "GetNextAdjustmentSequenceNumber", , True, , DescArgsFrom(l_strSQL), True
End Function

Public Function GetNextJCLSequenceNumber() As Long
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    l_strSQL = "SELECT NEXT VALUE FOR METHDBA.TTR_JCL_RQT_SEQ AS SEQ_NUM " & _
                "FROM SYSIBM.SYSDUMMY1"
              
    Set l_objRS = New ADODB.Recordset
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    
'    If l_objRS.EOF And l_objRS.BOF Then
'        Err.Raise trErrutGetTimeStamp
'    End If
        
    GetNextJCLSequenceNumber = l_objRS!SEQ_NUM
    
    l_objRS.ActiveConnection = Nothing
    Set l_objRS = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    
    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
       
    TrsRaiseError m_strCMODULENAME, "GetNextJCLSequenceNumber", , True, , DescArgsFrom(l_strSQL), True
End Function
