VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdAccrualBankMgrTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Private Const m_strCMODULENAME As String = "cdAccrualBankMgrTrans"



Public Function UpdateEmployeeBanks( _
                    ByVal v_rsEmployeeBanks As ADODB.Recordset, _
                    ByVal v_rsAuditTrail As ADODB.Recordset, _
                    ByVal v_lngLastModEmpNum As Long _
                    ) As Boolean

' Current App always does updates
' A batch job runs in September that creates banks for following
' The assumption is that will not change, so only updates will occur

    
       
    Dim l_objCN As New ADODB.Connection
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    

    On Error GoTo UpdateEmployeeBanks_EH
    
    
    
    If v_rsEmployeeBanks.EOF And v_rsEmployeeBanks.BOF Then
        GetObjectContext.SetComplete
        UpdateEmployeeBanks = True
        Exit Function
    End If
    
    
' This code below is used to eliminate the need to get a
' connection if no changes have been made
    v_rsEmployeeBanks.Filter = adFilterPendingRecords
    If v_rsEmployeeBanks.EOF And v_rsEmployeeBanks.BOF Then
        v_rsEmployeeBanks.Filter = ""
        GetObjectContext.SetComplete
        UpdateEmployeeBanks = True
        
        Exit Function
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    v_rsEmployeeBanks.MoveFirst
    l_strSQL = ""
    Do Until v_rsEmployeeBanks.EOF
        If v_rsEmployeeBanks.Status = adRecModified Then
            'BuildSQL, The neon driver can not batch updates
'            l_strSQL = l_strSQL & GetUpdateSQL(v_rsEmployeeBanks)
            l_strSQL = GetUpdateSQL(v_rsEmployeeBanks, v_lngLastModEmpNum)
            If l_strSQL <> "" Then
                l_objCN.Execute l_strSQL
            End If
        End If
        v_rsEmployeeBanks.MoveNext
    Loop
    
    'update banks
' The neon driver can not batch updates, the code below can be
' used to send all update statements in one call to the database
'    If Trim(l_strSQL) = "" Then
'    Else
'        l_objCN.Execute l_strSQL
'    End If
    
'Insert Audit Trail Records
    If v_rsAuditTrail.EOF And v_rsAuditTrail.BOF Then
        GetObjectContext.SetComplete
        l_objCN.Close
        Set l_objCN = Nothing
        
        Exit Function
    End If
    v_rsAuditTrail.MoveFirst
    While Not v_rsAuditTrail.EOF
        l_strSQL = InsertStatement(v_rsAuditTrail, "METHDBA.TTR_EMP_BANK_AUDIT")
'        l_strSQL = l_strSQL & InsertStatement(v_rsAuditTrail, "METHDBA.TTR_EMP_BANK_AUDIT") & " ; "
        l_objCN.Execute l_strSQL
        v_rsAuditTrail.MoveNext
    Wend
' The neon driver can not batch updates, the code below can be
' used to send all update statements in one call to the database
'    If Trim(l_strSQL) = "" Then
'    Else
'        l_objCN.Execute l_strSQL
'    End If
    
    GetObjectContext.SetComplete
    l_objCN.Close
    Set l_objCN = Nothing
    UpdateEmployeeBanks = True
    
    
    Exit Function

UpdateEmployeeBanks_EH:

    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    UpdateEmployeeBanks = False
    
    
    TrsRaiseError m_strCMODULENAME, "UpdateEmployeeBanks", , True, , DescArgsFrom(l_strSQL), True


End Function

'Private Sub UpdateAccrualBanks(ByVal v_rsEmployeeBanks As ADODB.Recordset, ByVal v_objCN As ADODB.Connection)
'
'    Dim l_lngEmpNum As Long
'    Dim l_lngEffYear As Long
'    Dim l_strSQL As String
'    Dim l_strSQLWhereClause As String
'    Dim l_strSQLSetStatement As String
'
'    On Error GoTo UpdateAccrualBanks_EH
'
'    l_lngEmpNum = v_rsEmployeeBanks!EMP_NUM
'    l_lngEffYear = v_rsEmployeeBanks!EFF_YEAR
'    l_strSQL = " Update methdba.ttr_emp_banks "
'    l_strSQLSetStatement = FormatSetStatement(v_rsEmployeeBanks)
'    l_strSQLWhereClause = " Where emp_num = " & Str(l_lngEmpNum) & _
'                            " and eff_year = " & Str(l_lngEffYear) & " "
'    l_strSQL = l_strSQL & l_strSQLSetStatement & l_strSQLWhereClause
'
'    If Trim(l_strSQLSetStatement) <> "" Then
'        v_objCN.Execute l_strSQL
'    End If
'    Exit Sub
'
'UpdateAccrualBanks_EH:
'
'    TrsRaiseError m_strCMODULENAME, "UpdateAccrualBanks", Err.Number, , , , True
'
'End Sub

Private Function GetUpdateSQL( _
                    ByVal v_rsEmployeeBanks As ADODB.Recordset, _
                    ByVal v_lngLastModEmpNum As Long _
                    ) As String

'This rountine builds the update SQL
' This will be needed when the Batch Updates (Multiple SQL in 1 call)
' is purchased from NEON
'Also in the mDataServiceLib module the UpdateStatement canbe used
' This was written prior to that module being available

' 7/18/02 RSG - Replaced v_strTimeStamp with CURRENT TIMESTAMP
    Dim l_lngEmpNum As Long
    Dim l_lngEffYear As Long
    Dim l_strSQL As String
    Dim l_strSQLWhereClause As String
    Dim l_strSQLSetStatement As String
    Dim l_strLastMod As String
    Dim l_strTimeStamp As String
    
    On Error GoTo GetUpdateSQL_EH
      

    l_lngEmpNum = v_rsEmployeeBanks!EMP_NUM
    l_lngEffYear = v_rsEmployeeBanks!EFF_YEAR
    l_strSQL = " Update methdba.ttr_emp_banks "
    l_strSQLSetStatement = FormatSetStatement(v_rsEmployeeBanks)
    l_strLastMod = " , LAST_MOD_EMP_NUM = " & str(v_lngLastModEmpNum)
    l_strTimeStamp = " , TIME_STAMP = CURRENT TIMESTAMP "
    l_strSQLWhereClause = " Where emp_num = " & str(l_lngEmpNum) & _
                            " and eff_year = " & str(l_lngEffYear) & " "
    l_strSQL = l_strSQL & l_strSQLSetStatement & l_strLastMod & l_strTimeStamp & l_strSQLWhereClause
    
    If Trim(l_strSQLSetStatement) <> "" Then
        GetUpdateSQL = " " & l_strSQL & " ; "
    Else
        GetUpdateSQL = ""
    End If
    Exit Function
                              
GetUpdateSQL_EH:

    TrsRaiseError m_strCMODULENAME, "GetUpdateSQL", , , , DescArgsFrom(l_strSQL), True

End Function

