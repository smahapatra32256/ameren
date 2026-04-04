VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 4  'RequiresNewTransaction
END
Attribute VB_Name = "cuDBDataServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "utilDBDataServices.cuDBDataServices"


Public Function GetTimeStamp() As String

    Dim l_utilDBDataServicesNonTran As utilDBDataServices.cuDBDataSvcNoTxn
    
    On Error GoTo GetTimeStamp_eh
    
    Set l_utilDBDataServicesNonTran = New utilDBDataServices.cuDBDataSvcNoTxn
    
    GetTimeStamp = l_utilDBDataServicesNonTran.GetTimeStamp
    
    GetObjectContext.SetComplete
    
    Exit Function

GetTimeStamp_eh:
    GetObjectContext.SetAbort
    Set l_utilDBDataServicesNonTran = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimeStamp", , , , , True
    
End Function


Public Function GetNextSequenceNumber( _
        ByVal v_strSequenceType As String _
        ) As Long
        
'Need to create Error Number
'Piggy backing on trErrutGetTimeStamp for now
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_lngNextSeqNum As Long
    Dim l_lngRecsAffected As Long

    On Error GoTo GetNextSequenceNumber_eh
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
              
    l_strSQL = "UPDATE methdba.ttr_seq_numbers " & _
                    " SET seq_num_last_used = seq_num_last_used + 1 " & _
                    " WHERE seq_num_type = '" & v_strSequenceType & "'"
              
    l_objCN.Execute l_strSQL, l_lngRecsAffected
              
    If l_lngRecsAffected <> 1 Then
        Err.Raise trErrutGetTimeStamp
    End If
              
    l_strSQL = "select SEQ_NUM_LAST_USED " & _
                " from methdba.ttr_seq_numbers " & _
                " where SEQ_NUM_TYPE = '" & v_strSequenceType & "' "
                            
    Set l_objRS = New ADODB.Recordset
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    l_objRS.ActiveConnection = Nothing
    
    If l_objRS.EOF And l_objRS.BOF Then
        Err.Raise trErrutGetTimeStamp
    End If
    
    l_lngNextSeqNum = l_objRS!SEQ_NUM_LAST_USED

    Set l_objRS = Nothing
    l_objCN.Close
    Set l_objCN = Nothing

    GetNextSequenceNumber = l_lngNextSeqNum
    GetObjectContext.SetComplete
    
    Exit Function

GetNextSequenceNumber_eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrutGetTimeStamp
            TrsRaiseError m_strCMODULENAME, "GetNextSequenceNumber", trErrutGetTimeStamp, True, , DescArgsFrom(l_strSQL), True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetNextSequenceNumber", , True, , DescArgsFrom(l_strSQL), True
    End Select
    
End Function

