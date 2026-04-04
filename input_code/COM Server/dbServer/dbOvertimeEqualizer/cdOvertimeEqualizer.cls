VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdOvertimeEqualizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cdOvertimeEqualizer"
Public Function UpdateEmployeeOvertimeBanks( _
                    ByVal v_rsEmployeeOTBanks As ADODB.Recordset, _
                    Optional ByVal v_rsAuditTrail As ADODB.Recordset = Nothing _
                    ) As Boolean
       
    Dim l_objCN As New ADODB.Connection
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    Dim l_aryPrimaryKey(1) As String
    Dim l_bolHaveAuditTrail As Boolean
    Dim l_colErrorDesc As Collection
    Dim l_lngRecsAffected As Long

    On Error GoTo UpdateEmployeeOvertimeBanks_EH
    
    
    
    If v_rsEmployeeOTBanks.EOF And v_rsEmployeeOTBanks.BOF Then
        GetObjectContext.SetComplete
        UpdateEmployeeOvertimeBanks = True
        Exit Function
    End If
    
    
' This code below is used to eliminate the need to get a
' connection if no changes have been made
    v_rsEmployeeOTBanks.Filter = adFilterPendingRecords
    If v_rsEmployeeOTBanks.EOF And v_rsEmployeeOTBanks.BOF Then
        v_rsEmployeeOTBanks.Filter = ""
        GetObjectContext.SetComplete
        UpdateEmployeeOvertimeBanks = True
        Exit Function
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    v_rsEmployeeOTBanks.MoveFirst
    l_aryPrimaryKey(0) = "EMP_NUM"
    l_aryPrimaryKey(1) = "OT_BANK_CD"
    Do Until v_rsEmployeeOTBanks.EOF
        Select Case v_rsEmployeeOTBanks.Status
            Case adRecModified
        'BuildSQL, The neon driver can not batch updates
        '       l_strSQL = l_strSQL & GetUpdateSQL(v_rsEmployeeOTBanks)
                l_strSQL = GetUpdateSQL(v_rsEmployeeOTBanks)
            Case adRecNew
'                l_strSQL = l_strSQL & InsertStatement(v_rsEmployeeOTBanks, "METHDBA.TTR_OT_BANKS")
                l_strSQL = InsertStatement(v_rsEmployeeOTBanks, "METHDBA.TTR_OT_BANKS")
            Case adRecDeleted
'                l_strSQL = l_strSQL & DeleteStatement(v_rsEmployeeOTBanks, "METHDBA.TTR_OT_BANKS", l_aryPrimaryKey)
                l_strSQL = DeleteStatement(v_rsEmployeeOTBanks, "METHDBA.TTR_OT_BANKS", l_aryPrimaryKey)
            Case Else
                l_strSQL = ""
        End Select
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsEmployeeOTBanks.MoveNext
    Loop
    
    'update banks
' The neon driver can not batch updates, the code below can be
' used to send all update statements in one call to the database
'    If Trim(l_strSQL) = "" Then
'    Else
'        l_objCN.Execute l_strSQL
'    End If
    
'Insert Audit Trail Records
    
    If v_rsAuditTrail Is Nothing Then
        l_bolHaveAuditTrail = False
    Else
        If v_rsAuditTrail.EOF And v_rsAuditTrail.BOF Then
            l_bolHaveAuditTrail = False
        Else
            l_bolHaveAuditTrail = True
        End If
    End If
    
    If Not l_bolHaveAuditTrail Then
        GetObjectContext.SetComplete
        l_objCN.Close
        Set l_objCN = Nothing
        UpdateEmployeeOvertimeBanks = True
        Exit Function
    End If
    v_rsAuditTrail.MoveFirst
    While Not v_rsAuditTrail.EOF
        l_strSQL = InsertStatement(v_rsAuditTrail, "METHDBA.TTR_OT_BANK_ADJ")
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
    UpdateEmployeeOvertimeBanks = True
    
    
    Exit Function

UpdateEmployeeOvertimeBanks_EH:

    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = l_strSQL
            TrsRaiseError m_strCMODULENAME, "UpdateEmployeeOvertimeBanks", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateEmployeeOvertimeBanks", , True, , DescArgsFrom(l_strSQL), True
    End Select

End Function

Public Function UpdateEmployeeOvertimeTempBanks( _
                    ByVal v_rsEmployeeOTTempBanks As ADODB.Recordset _
                    ) As Boolean
 
       
    Dim l_objCN As New ADODB.Connection
    Dim l_strSQL As String

    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    Dim l_aryPrimaryKey(3) As String
    Dim l_colErrorDesc As Collection
    Dim l_lngRecsAffected As Long

    On Error GoTo UpdateEmployeeOvertimeTempBanks_EH
    
    
    
    If v_rsEmployeeOTTempBanks.EOF And v_rsEmployeeOTTempBanks.BOF Then
        GetObjectContext.SetComplete
        UpdateEmployeeOvertimeTempBanks = True
        Exit Function
    End If
    
    
' This code below is used to eliminate the need to get a
' connection if no changes have been made
    v_rsEmployeeOTTempBanks.Filter = adFilterPendingRecords
    If v_rsEmployeeOTTempBanks.EOF And v_rsEmployeeOTTempBanks.BOF Then
        v_rsEmployeeOTTempBanks.Filter = ""
        GetObjectContext.SetComplete
        UpdateEmployeeOvertimeTempBanks = True
        Exit Function
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    v_rsEmployeeOTTempBanks.MoveFirst
    l_strSQL = ""
    l_aryPrimaryKey(0) = "EMP_NUM"
    l_aryPrimaryKey(1) = "OT_OPP_NUM"
    l_aryPrimaryKey(2) = "OT_BANK_CD"
    l_aryPrimaryKey(3) = "EFF_DATE"
    Do Until v_rsEmployeeOTTempBanks.EOF
        Select Case v_rsEmployeeOTTempBanks.Status
            Case adRecModified
        'BuildSQL, The neon driver can not batch updates
        '       l_strSQL = l_strSQL & GetUpdateSQLTempBanks(v_rsEmployeeOTTempBanks)
                l_strSQL = GetUpdateSQLTempBanks(v_rsEmployeeOTTempBanks)
            Case adRecNew
'                l_strSQL = l_strSQL & InsertStatement(v_rsEmployeeOTTempBanks, "METHDBA.TTR_OT_TEMP_BANKS")
                l_strSQL = InsertStatement(v_rsEmployeeOTTempBanks, "METHDBA.TTR_OT_TEMP_BANKS")
            Case adRecDeleted
'                l_strSQL = l_strSQL & DeleteStatement(v_rsEmployeeOTTempBanks, "METHDBA.TTR_OT_TEMP_BANKS", l_aryPrimaryKey)
                l_strSQL = DeleteStatement(v_rsEmployeeOTTempBanks, "METHDBA.TTR_OT_TEMP_BANKS", l_aryPrimaryKey)
            Case Else
                l_strSQL = ""
        End Select
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsEmployeeOTTempBanks.MoveNext
    Loop
    
    'update banks
' The neon driver can not batch updates, the code below can be
' used to send all update statements in one call to the database
'    If Trim(l_strSQL) = "" Then
'    Else
'        l_objCN.Execute l_strSQL
'    End If
    
    GetObjectContext.SetComplete
    l_objCN.Close
    Set l_objCN = Nothing
    UpdateEmployeeOvertimeTempBanks = True
    
    Exit Function

UpdateEmployeeOvertimeTempBanks_EH:

    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = l_strSQL
            TrsRaiseError m_strCMODULENAME, "UpdateEmployeeOvertimeTempBanks", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateEmployeeOvertimeTempBanks", , True, , DescArgsFrom(l_strSQL), True
    End Select
    
End Function

Public Function UpdateExecutiveOvertimeBanks( _
                    ByVal v_rsExcutiveOTBanks As ADODB.Recordset _
                    ) As Boolean
 
       
    Dim l_objCN As New ADODB.Connection
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    Dim l_aryPrimaryKey(1) As String
    Dim l_colErrorDesc As Collection

    On Error GoTo UpdateEmployeeOvertimeTempBanks_EH
    
    If v_rsExcutiveOTBanks.EOF And v_rsExcutiveOTBanks.BOF Then
        GetObjectContext.SetComplete
        UpdateExecutiveOvertimeBanks = True
        Exit Function
    End If
    
    
' This code below is used to eliminate the need to get a
' connection if no changes have been made
    v_rsExcutiveOTBanks.Filter = adFilterPendingRecords
    If v_rsExcutiveOTBanks.EOF And v_rsExcutiveOTBanks.BOF Then
        v_rsExcutiveOTBanks.Filter = ""
        GetObjectContext.SetComplete
        UpdateExecutiveOvertimeBanks = True
        Exit Function
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_aryPrimaryKey(0) = "EMP_NUM"
    l_aryPrimaryKey(1) = "EFF_YEAR"
    v_rsExcutiveOTBanks.MoveFirst
    l_strSQL = ""
    Do Until v_rsExcutiveOTBanks.EOF
        Select Case v_rsExcutiveOTBanks.Status
            Case adRecModified
        'BuildSQL, The neon driver can not batch updates
        '       l_strSQL = l_strSQL & GetUpdateSQL(v_rsExcutiveOTBanks)
                l_strSQL = UpdateStatement(v_rsExcutiveOTBanks, "METHDBA.TTR_OT_EXEC_HRS", l_aryPrimaryKey)
            Case adRecNew
'                l_strSQL = l_strSQL & InsertStatement(v_rsExcutiveOTBanks, "METHDBA.TTR_OT_EXEC_HRS")
                l_strSQL = InsertStatement(v_rsExcutiveOTBanks, "METHDBA.TTR_OT_EXEC_HRS")
            Case adRecDeleted
'                l_strSQL = l_strSQL & DeleteStatement(v_rsExcutiveOTBanks, "METHDBA.TTR_OT_EXEC_HRS", l_aryPrimaryKey)
                l_strSQL = DeleteStatement(v_rsExcutiveOTBanks, "METHDBA.TTR_OT_EXEC_HRS", l_aryPrimaryKey)
            Case Else
                l_strSQL = ""
        End Select
        If l_strSQL <> "" Then
                l_objCN.Execute l_strSQL
        End If
        v_rsExcutiveOTBanks.MoveNext
    Loop
    
    'update banks
' The neon driver can not batch updates, the code below can be
' used to send all update statements in one call to the database
'    If Trim(l_strSQL) = "" Then
'    Else
'        l_objCN.Execute l_strSQL
'    End If
    
    GetObjectContext.SetComplete
    l_objCN.Close
    Set l_objCN = Nothing
    UpdateExecutiveOvertimeBanks = True
    
    Exit Function

UpdateEmployeeOvertimeTempBanks_EH:

    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    TrsRaiseError m_strCMODULENAME, "UpdateEmployeeOvertimeTempBanks", , , , DescArgsFrom(l_strSQL), True

End Function


Private Function GetUpdateSQL( _
                    ByVal v_rsEmployeeOTBanks As ADODB.Recordset _
                    ) As String

'This rountine builds the update SQL
' This will be needed when the Batch Updates (Multiple SQL in 1 call)
' is purchased from NEON
'Also in the mDataServiceLib module the UpdateStatement canbe used
' This was written prior to that module being available


    Dim l_lngEmpNum As Long
    Dim l_strOTBank As String
    Dim l_strSQL As String
    Dim l_strSQLWhereClause As String
    Dim l_strSQLSetStatement As String
    Dim l_strTimeStamp As String
    Dim l_strSetTimeStamp As String
    
    On Error GoTo GetUpdateSQL_EH
      

    l_lngEmpNum = v_rsEmployeeOTBanks!EMP_NUM
    l_strOTBank = Trim(str(v_rsEmployeeOTBanks!OT_BANK_CD))
    l_strTimeStamp = "'" & v_rsEmployeeOTBanks("TIME_STAMP").OriginalValue & "'"
    l_strSQL = " Update methdba.ttr_ot_banks "
    l_strSQLSetStatement = FormatSetStatement(v_rsEmployeeOTBanks)
    l_strSetTimeStamp = " , TIME_STAMP = CURRENT TIMESTAMP "
    l_strSQLWhereClause = " Where emp_num = " & str(l_lngEmpNum) & _
                            " and OT_BANK_CD = " & l_strOTBank & _
                            " and TIME_STAMP = " & l_strTimeStamp
    l_strSQL = l_strSQL & l_strSQLSetStatement & l_strSetTimeStamp & l_strSQLWhereClause
    
    If Trim(l_strSQLSetStatement) <> "" Then
        GetUpdateSQL = " " & l_strSQL & " ; "
    Else
        GetUpdateSQL = ""
    End If
    Exit Function
                              
GetUpdateSQL_EH:

    TrsRaiseError m_strCMODULENAME, "GetUpdateSQL", , , , DescArgsFrom(l_strSQL)

End Function

Private Function GetUpdateSQLTempBanks( _
                    ByVal v_rsEmployeeOTBanks As ADODB.Recordset _
                    ) As String

'This rountine builds the update SQL
' This will be needed when the Batch Updates (Multiple SQL in 1 call)
' is purchased from NEON
'Also in the mDataServiceLib module the UpdateStatement canbe used
' This was written prior to that module being available

    Dim l_lngEmpNum As Long
    Dim l_strOTBank As String
    Dim l_strOTTempAssignInd As String
    Dim l_lngOTOppNum As Long
    Dim l_strSQL As String
    Dim l_strSQLWhereClause As String
    Dim l_strSQLSetStatement As String
    Dim l_strTimeStamp As String
    Dim l_datEffDate As Date
    Dim l_strSetTimeStamp As String
    
    On Error GoTo eh
      
    l_lngEmpNum = v_rsEmployeeOTBanks!EMP_NUM
    l_strOTBank = Trim(str(v_rsEmployeeOTBanks!OT_BANK_CD))
    l_lngOTOppNum = v_rsEmployeeOTBanks!OT_OPP_NUM
    l_datEffDate = v_rsEmployeeOTBanks!EFF_DATE
    l_strTimeStamp = "'" & v_rsEmployeeOTBanks("TIME_STAMP").OriginalValue & "'"
    l_strSQL = " Update methdba.TTR_OT_TEMP_BANKS "
    l_strSQLSetStatement = FormatSetStatement(v_rsEmployeeOTBanks)
    l_strSetTimeStamp = " , TIME_STAMP = CURRENT TIMESTAMP "
    l_strSQLWhereClause = " Where emp_num = " & str(l_lngEmpNum) & _
                            " and OT_OPP_NUM = " & str(l_lngOTOppNum) & _
                            " and OT_BANK_CD = " & l_strOTBank & _
                            " and EFF_DATE = '" & Format(l_datEffDate, "YYYY-MM-DD") & "' " & _
                            " and TIME_STAMP = " & l_strTimeStamp
    l_strSQL = l_strSQL & l_strSQLSetStatement & l_strSetTimeStamp & l_strSQLWhereClause
    
    If Trim(l_strSQLSetStatement) <> "" Then
        GetUpdateSQLTempBanks = " " & l_strSQL & " ; "
    Else
        GetUpdateSQLTempBanks = ""
    End If
    Exit Function
                              
eh:

    TrsRaiseError m_strCMODULENAME, "GetUpdateSQLTempBanks", , , , DescArgsFrom(l_strSQL)

End Function

Public Function GetOvertimeResponses() As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetOvertimeResponses_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    

    l_strSQL = "select OT_RESP_CD , RESPONSE_DESC, OT_CHRG_IND, OT_WORKED_IND " & _
                " , OT_CONTACTED_IND " & _
               " From methdba.ttr_ot_responses "
               
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    GetObjectContext.SetComplete
    Set GetOvertimeResponses = l_objRS
    
    Set l_objRS = Nothing
    

    Exit Function

GetOvertimeResponses_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetOvertimeResponses", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetEmployeeOvertimeBanks( _
        ByVal v_aryEmployees As Variant _
        ) As ADODB.Recordset
    
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo eh
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    

    l_strSQL = "select EMP_NUM , A.OT_BANK_CD, OT_BANK_DESC, " & _
                " TOT_OT_HRS, TOT_OT_HRS_REF, A.LAST_MOD_EMP_NUM " & _
               " From methdba.ttr_ot_banks A, " & _
               "      methdba.ttr_ot_bank_cd B " & _
               " where " & l_strEmployeeSelection & _
               "   and B.OT_BANK_CD = A.OT_BANK_CD " & _
               "  order by EMP_NUM, A.OT_BANK_CD "
               
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    GetObjectContext.SetComplete
    Set GetEmployeeOvertimeBanks = l_objRS
    
    Set l_objRS = Nothing
    

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetEmployeeOvertimeBanks", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetEmployeeOvertimeTempBanks( _
        ByVal v_aryEmployees As Variant _
        ) As ADODB.Recordset
    
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo eh
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    l_strSQL = "select EMP_NUM , OT_OPP_NUM, OT_BANK_CD, " & _
                " EFF_DATE, TOT_OT_HRS, TOT_OT_HRS_REF, LAST_MOD_EMP_NUM " & _
                " ,CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, HOURS, RATE " & _
                " ,CHAR(TIME_STAMP) as TIME_STAMP " & _
               " From methdba.ttr_ot_temp_banks " & _
               " where " & BuildInClause("EMP_NUM", v_aryEmployees, False) _
               & "  order by emp_num, eff_date, OT_BANK_CD"
               
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    GetObjectContext.SetComplete
    Set GetEmployeeOvertimeTempBanks = l_objRS
    
    Set l_objRS = Nothing
    

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetEmployeeOvertimeTempBanks", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetEmployeeQuarterlyOTBanks( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date _
        ) As ADODB.Recordset
TrsRaiseError m_strCMODULENAME, "GetEmployeeQuarterlyOTBanks", trErrObsoleteMethod, , , , True
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'    Dim l_lngCurrentYear As Long
'    Dim l_lngPriorYear As Long
'    Dim l_lngNextYear As Long
'    Dim l_strInClause As String
'    Dim l_strEmployeeSelection As String
'
'    On Error GoTo eh
'
'
'    Set l_objCN = New ADODB.Connection
'    Set l_objRS = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
'
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.CursorLocation = adUseClient
'    l_objCN.Open
'    If IsArray(v_aryEmployees) Then
'        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
'    Else
'        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
'    End If
'
'    l_lngCurrentYear = Year(v_datStartDate)
'    l_lngPriorYear = l_lngCurrentYear - 1
'    l_lngNextYear = l_lngCurrentYear + 1
'    l_strInClause = str(l_lngCurrentYear) & "," & str(l_lngPriorYear) & "," & str(l_lngNextYear)
'
'    l_strSQL = "SELECT EMP_NUM, EFF_YEAR, QTR_1_HRS, QTR_2_HRS, QTR_3_HRS, QTR_4_HRS "
'    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_EXEC_HRS "
'    l_strSQL = l_strSQL & "WHERE " & l_strEmployeeSelection
'    l_strSQL = l_strSQL & " AND EFF_YEAR IN (" & l_strInClause & ") "
'    l_strSQL = l_strSQL & " ORDER BY EMP_NUM, EFF_YEAR DESC"
'
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'    Set l_objCN = Nothing
'
'    Set GetEmployeeQuarterlyOTBanks = l_objRS
'
'    Set l_objRS = Nothing
'
'    GetObjectContext.SetComplete
'
'
'Exit Function
'eh:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'    GetObjectContext.SetAbort
'    TrsRaiseError m_strCMODULENAME, "GetEmployeeQuarterlyOTBanks", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetEmployeeOvertimeBanksForUpdate( _
        ByVal v_aryEmployees As Variant _
        ) As ADODB.Recordset
    
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo eh
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    

    l_strSQL = "select EMP_NUM , OT_BANK_CD, " & _
                " TOT_OT_HRS, TOT_OT_HRS_REF, LAST_MOD_EMP_NUM " & _
                " ,CHAR(TIME_STAMP) as TIME_STAMP " & _
               " From methdba.ttr_ot_banks  " & _
               " where " & l_strEmployeeSelection & _
               "  order by EMP_NUM, OT_BANK_CD "
               
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    GetObjectContext.SetComplete
    Set GetEmployeeOvertimeBanksForUpdate = l_objRS
    
    Set l_objRS = Nothing
    

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetEmployeeOvertimeBanksForUpdate", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetOTGroupBanks( _
    ByVal v_intOTGroup As Integer _
    ) As ADODB.Recordset

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
    

    l_strSQL = "select OT_GROUP_NUM, OT_BANK_CD " & _
               " from methdba.TTR_OT_GROUP_BANK " & _
               "where OT_GROUP_NUM = " & str(v_intOTGroup)
               
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    GetObjectContext.SetComplete
    Set GetOTGroupBanks = l_objRS
    
    Set l_objRS = Nothing
    

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetOTGroupBanks", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetOvertimeBankTimeCharges( _
    ByVal v_datChargeDate As Date, _
    ByVal v_strShift As String, _
    Optional v_lngOTOppNum As Long = 0 _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    If v_lngOTOppNum = 0 Then
        If Trim(v_strShift) = "" Then
            l_strWhereClause = "where CHARGE_DATE = '" & Format(v_datChargeDate, "YYYY-MM-DD") & "' "
        Else
            l_strWhereClause = "where CHARGE_DATE = '" & Format(v_datChargeDate, "YYYY-MM-DD") & "' " & _
                "  AND SHIFT_CD = '" & Trim(v_strShift) & "'"
        End If
    Else
        l_strWhereClause = "where OT_OPP_NUM = " & str(v_lngOTOppNum)
    End If
    l_strSQL = "select OT_OPP_NUM, EMP_NUM, CHARGE_DATE, DATE_WORKED, " & _
                    "START_TIME, END_TIME, HOURS, RATE, HOURS_CHARGED, " & _
                    "ACC_REF_IND, SHIFT_CD, CREATED_EMP_NUM, LAST_MOD_EMP_NUM, " & _
                    "CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, " & _
                    "CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP " & _
               " from methdba.TTR_OT_BANK_TM_CHG  "

'    l_strSQL = "select a.OT_OPP_NUM, a.EMP_NUM, a.CHARGE_DATE, a.DATE_WORKED, " & _
'                    "a.START_TIME, a.END_TIME, a.HOURS, a.RATE, a.HOURS_CHARGED, " & _
'                    "a.ACC_REF_IND, a.SHIFT_CD, a.CREATED_EMP_NUM, a.LAST_MOD_EMP_NUM, " & _
'                    "CHAR(a.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, " & _
'                    "CHAR(a.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP " & _
'                    ",b.EXCEPTION_CD, c.OT_RESPONSE " & _
'               " from methdba.TTR_OT_BANK_TM_CHG a "
'
'    l_strSQL = l_strSQL & _
'                "INNER JOIN methdba.ttr_ot_opp_tm_chrg b " & _
'                "  on a.ot_opp_num = b.ot_opp_num " & _
'                " and a.emp_num = b.emp_num " & _
'                " and a.charge_date = b.charge_date " & _
'                " and a.date_worked = b.date_worked " & _
'                " and a.start_time = b.start_time " & _
'                " and a.end_time = b.end_time "
'
'    l_strSQL = l_strSQL & _
'                "INNER JOIN " & _
'                " ( select ot_opp_num, emp_num, " & _
'                    " case when FORCE_RESP_CD > '' then FORCE_RESP_CD " & _
'                    " else CANVASS_RESP_CD END AS OT_RESPONSE " & _
'                "  From methdba.ttr_canv_sheets ) c " & _
'                " on a.ot_opp_num = c.ot_opp_num " & _
'                " and a.emp_num = c.emp_num "
                
    l_strSQL = l_strSQL & l_strWhereClause
               
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    GetObjectContext.SetComplete
    Set GetOvertimeBankTimeCharges = l_objRS
    
    Set l_objRS = Nothing
    

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetOvertimeBankTimeCharges", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function SaveOvertimeBankTimeCharges( _
                    ByVal v_rsOTBankCharges As ADODB.Recordset _
                    ) As Boolean
 
       
    Dim l_objCN As New ADODB.Connection
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_varIgnore(1) As String
    Dim l_aryPrimaryKey(2) As String
    Dim l_colErrorDesc As Collection
    Dim l_lngRecsAffected As Long

    On Error GoTo eh

' This code below is used to eliminate the need to get a
' connection if no changes have been made
    v_rsOTBankCharges.Filter = adFilterPendingRecords
    If v_rsOTBankCharges.EOF And v_rsOTBankCharges.BOF Then
        v_rsOTBankCharges.Filter = ""
        GetObjectContext.SetComplete
        SaveOvertimeBankTimeCharges = True
        Exit Function
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_aryPrimaryKey(0) = "OT_OPP_NUM"
    l_aryPrimaryKey(1) = "EMP_NUM"
    l_aryPrimaryKey(2) = "CREATED_TIMESTAMP"
    l_varIgnore(0) = "CREATED_TIMESTAMP"
    l_varIgnore(1) = "LAST_MOD_TIMESTAMP"
    v_rsOTBankCharges.MoveFirst
    l_strSQL = ""
    Do Until v_rsOTBankCharges.EOF
        Select Case v_rsOTBankCharges.Status
            Case adRecModified
                l_strSQL = UpdateStatement(v_rsOTBankCharges, "METHDBA.TTR_OT_BANK_TM_CHG", l_aryPrimaryKey, l_varIgnore, True)
            Case adRecNew
                l_strSQL = InsertStatement(v_rsOTBankCharges, "METHDBA.TTR_OT_BANK_TM_CHG", l_varIgnore)
            Case adRecDeleted
                l_strSQL = DeleteStatement(v_rsOTBankCharges, "METHDBA.TTR_OT_BANK_TM_CHG", l_aryPrimaryKey)
            Case Else
                l_strSQL = ""
        End Select
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsOTBankCharges.MoveNext
    Loop
    
    GetObjectContext.SetComplete
    l_objCN.Close
    Set l_objCN = Nothing
    SaveOvertimeBankTimeCharges = True
    
    Exit Function
    
eh:

    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = l_strSQL
            TrsRaiseError m_strCMODULENAME, "SaveOvertimeBankTimeCharges", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveOvertimeBankTimeCharges", , True, , DescArgsFrom(l_strSQL), True
    End Select

End Function


Public Function GetOppOTBankCharges( _
    ByVal v_lngOTOppNum As Long _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    l_strWhereClause = "where OT_OPP_NUM = " & str(v_lngOTOppNum)
    
    l_strSQL = _
        "select OT_OPP_NUM, EMP_NUM, CHARGE_DATE, DATE_WORKED, " & _
                    "START_TIME, END_TIME, HOURS, HOURS_CHARGED, ACC_REF_IND " & _
               " from methdba.TTR_OT_BANK_TM_CHG  " & _
               l_strWhereClause
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    GetObjectContext.SetComplete
    Set GetOppOTBankCharges = l_objRS
    
    Set l_objRS = Nothing
    

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetOppOTBankCharges", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetOTGroupCanvassBank( _
    ByVal v_intOTGroup As Integer _
    ) As ADODB.Recordset
        
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
    l_strSQL = "SELECT OT_STAND_BANK_CD " & _
                "FROM METHDBA.TTR_OT_GROUP " & _
                "WHERE OT_GROUP_NUM = " & str(v_intOTGroup)


    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetOTGroupCanvassBank = l_objRS

    Set l_objRS = Nothing

    GetObjectContext.SetComplete

    Exit Function


eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTGroupCanvassBank", , , , DescArgsFrom(l_strSQL)

    
    
End Function




Public Function GetHolidayRestExceptions( _
    ) As ADODB.Recordset
        
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

    l_strSQL = "SELECT EXCEPTION_CD " & _
                "FROM METHDBA.TTR_EXCEPTION_CDS " & _
                "WHERE GROUP_IND = 'H' " & _
                "  AND WORK_REST_IND = 'R' "


    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetHolidayRestExceptions = l_objRS

    Set l_objRS = Nothing

    GetObjectContext.SetComplete

    Exit Function


eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetHolidayRestExceptions", , , , DescArgsFrom(l_strSQL)

    
    
End Function

