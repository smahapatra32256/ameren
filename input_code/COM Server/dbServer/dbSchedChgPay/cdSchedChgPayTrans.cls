VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdSchedChgPayTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdSchedChgPayTrans"
Private Const m_datCOPTIONALSTARTDATE As Date = #1/1/100#
Private Const m_datCOPTIONALENDDATE As Date = #12/31/9999#
'//-----used the following for testing
Private Const m_strConnectStringReg As String = "SHADOWDIRECT"
'DB2TEST_UDB
'SHADOWDIRECT
       
Public Function SaveSchedChgPay(ByVal v_rsSchedChgPayRequest As ADODB.Recordset, ByVal v_lTRISUser As Long) As Boolean
    Dim cn As ADODB.Connection
    Dim rc As Boolean, bError As Boolean
On Error GoTo eh

    Set cn = OpenConnection

    With v_rsSchedChgPayRequest
        bError = False
        If Not (v_rsSchedChgPayRequest Is Nothing) Then
            '//-----The following code was moved to DeleteSchedChgPay so that it is done prior to validation-----//
        '  Perform Deletes
            .Filter = adFilterNone
            .MoveFirst
            '//-----clean up job pos and 21st Watch-----//
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
             Do While Not .EOF
                rc = DeleteSchedChgPay(v_rsSchedChgPayRequest, cn)
                If Not rc Then bError = True
                .MoveNext
            Loop

        '  Perform Inserts
            .Filter = adFilterNone
            .MoveFirst
            '//-----an insert with emp_num=0 means to unassign, so skip it here (AdjustForOverlap should do the rest)-----//
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
            Do While Not .EOF
                rc = InsertSchedChgPay(v_rsSchedChgPayRequest, v_lTRISUser, cn)
                If Not rc Then bError = True
                .MoveNext
            Loop

        '  Perform Changes
            .Filter = adFilterNone
            .MoveFirst
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
            Do While Not .EOF
                rc = UpdateSchedChgPay(v_rsSchedChgPayRequest, v_lTRISUser, cn)
                If Not rc Then bError = True
                .MoveNext
            Loop
        .Filter = adFilterNone
        End If
    End With

    cn.Close

    Set cn = Nothing
    GetObjectContext.SetComplete

    SaveSchedChgPay = Not bError

    Exit Function
eh:
    Dim adoError As ADODB.Error
    Dim sError As String
    
    GetObjectContext.SetAbort
    
    If Not v_rsSchedChgPayRequest Is Nothing Then
        sError = v_rsSchedChgPayRequest.GetString
    End If
    
    TrsRaiseError m_strCMODULENAME, "SaveSchedChgPay", , True, , DescArgsFrom(sError), True
End Function
    
Private Function DeleteSchedChgPay(v_rsSchedChgPayRequest As ADODB.Recordset, cn As ADODB.Connection) As Boolean
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long

    On Error GoTo eh

    With v_rsSchedChgPayRequest
        l_strSQL = "DELETE FROM METHDBA.TTR_SCHED_CHG_PMTS"
        l_strSQL = l_strSQL & " WHERE EMP_NUM = " & !KEY_EMP_NUM
        l_strSQL = l_strSQL & " AND CHARGE_DATE = " & dbFormat(!KEY_CHARGE_DATE)
        l_strSQL = l_strSQL & " AND LAST_MOD_TIMESTAMP = '" & !LAST_MOD_TIMESTAMP & "'"
    End With
    
    If l_strSQL <> "" Then
       cn.Execute l_strSQL, l_lngRecsAffected
       If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
            Err.Raise trErrAppCommitConflict
       End If
    End If
    
    DeleteSchedChgPay = True
Exit Function

eh:

    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = _
                Join( _
                    Array( _
                        "1", _
                        CStr(v_rsSchedChgPayRequest!KEY_EMP_NUM), _
                        Format(v_rsSchedChgPayRequest!KEY_CHARGE_DATE, "mm/dd/yyyy")) _
                    , ";")
            TrsRaiseError m_strCMODULENAME, "DeleteSchedChgPay", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "DeleteSchedChgPay", , True, , DescArgsFrom(l_strSQL), True
    End Select

End Function

Private Function InsertSchedChgPay(v_rsSchedChgPayRequest As ADODB.Recordset, ByVal v_lTRISUser As Long, cn As ADODB.Connection) As Boolean
    Dim l_strSQL As String

On Error GoTo eh

    With v_rsSchedChgPayRequest
        l_strSQL = "INSERT INTO METHDBA.TTR_SCHED_CHG_PMTS"
        l_strSQL = l_strSQL & " (EMP_NUM"
        l_strSQL = l_strSQL & ", CHARGE_DATE"
        
        If Not IsNull(!BASELINE_SCHED_CD) Then
            l_strSQL = l_strSQL & ", BASELINE_SCHED_CD"
        End If
        
        l_strSQL = l_strSQL & ", EXCEPTION_CD"
        l_strSQL = l_strSQL & ", PAYMENT_TYPE"
        
        l_strSQL = l_strSQL & ", CREATED_EMP_NUM"
        l_strSQL = l_strSQL & ", LAST_MOD_EMP_NUM"
        l_strSQL = l_strSQL & ", CREATED_TIMESTAMP"
        l_strSQL = l_strSQL & ", LAST_MOD_TIMESTAMP)"
        
        l_strSQL = l_strSQL & " VALUES(" & !EMP_NUM
        l_strSQL = l_strSQL & ", " & dbFormat(!CHARGE_DATE)
        
        If Not IsNull(!BASELINE_SCHED_CD) Then
            l_strSQL = l_strSQL & ", " & !BASELINE_SCHED_CD
        End If
        l_strSQL = l_strSQL & ", " & dbFormat(!EXCEPTION_CD)
        l_strSQL = l_strSQL & ", " & dbFormat(!PAYMENT_TYPE)

        If v_lTRISUser >= 0 Then
            '//-----because an Update ("C") request changes to a delete/insert,
            '//----- need to carry the created columns over-----//
            '//-----also needed for conversion of 21st Watch which will preserve the last_mod columns-----//
            If IsNull(!CREATED_EMP_NUM) Or !CREATED_EMP_NUM = 0 Then
                l_strSQL = l_strSQL & ", " & v_lTRISUser
            Else
                l_strSQL = l_strSQL & ", " & !CREATED_EMP_NUM '//-----not a new insert, but re-inserted-----//
            End If

            l_strSQL = l_strSQL & ", " & v_lTRISUser

            If IsNull(!CREATED_TIMESTAMP) Or Trim(!CREATED_TIMESTAMP) = "" Then
                l_strSQL = l_strSQL & ", CURRENT TIMESTAMP "
            Else
                l_strSQL = l_strSQL & ", " & dbFormat(!CREATED_TIMESTAMP) '//-----not a new insert, but re-inserted-----//
            End If

            l_strSQL = l_strSQL & ", CURRENT TIMESTAMP )"
        Else '//-----for conversion of legacy data (ttr_base_hours_exc) 21st Watch-----//
            l_strSQL = l_strSQL & ", " & !CREATED_EMP_NUM
            l_strSQL = l_strSQL & ", 999999999" '//& !LAST_MOD_EMP_NUM
            l_strSQL = l_strSQL & ", " & dbFormat(!CREATED_TIMESTAMP)
            l_strSQL = l_strSQL & ", CURRENT TIMESTAMP )" '// & dbFormat(!LAST_MOD_TIMESTAMP) & " )"
        End If
    End With

    If l_strSQL <> "" Then
       cn.Execute l_strSQL
    End If
    InsertSchedChgPay = True
Exit Function
eh:
    Dim adoError As ADODB.Error
    GetObjectContext.SetAbort
    For Each adoError In cn.Errors
         If adoError.NativeError = -803 Then '//-----db2 duplicate key-----//
            Err.Raise trErrbusSchedDemandDuplicateLineAsgnm
            Exit Function
         Else
            TrsRaiseError m_strCMODULENAME, "InsertSchedChgPay", , True, , DescArgsFrom(l_strSQL), True
            Exit Function
         End If
    Next

    InsertSchedChgPay = False
    TrsRaiseError m_strCMODULENAME, "InsertSchedChgPay", , True, , DescArgsFrom(l_strSQL), True
End Function

Private Function UpdateSchedChgPay(v_rsSchedChgPayRequest As ADODB.Recordset, ByVal v_lTRISUser As Long, cn As ADODB.Connection) As Boolean
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long

On Error GoTo eh
        
    With v_rsSchedChgPayRequest
        l_strSQL = "UPDATE METHDBA.TTR_SCHED_CHG_PMTS SET"
        l_strSQL = l_strSQL & " EMP_NUM = " & !EMP_NUM
        l_strSQL = l_strSQL & ", CHARGE_DATE = " & dbFormat(!CHARGE_DATE)
        
        If Not IsNull(!BASELINE_SCHED_CD) Then
            l_strSQL = l_strSQL & ", BASELINE_SCHED_CD = " & !BASELINE_SCHED_CD
        Else
            l_strSQL = l_strSQL & ", BASELINE_SCHED_CD = 0"
        End If
        l_strSQL = l_strSQL & ", EXCEPTION_CD = " & dbFormat(!EXCEPTION_CD)
        l_strSQL = l_strSQL & ", PAYMENT_TYPE = " & dbFormat(!PAYMENT_TYPE)
        l_strSQL = l_strSQL & ", LAST_MOD_EMP_NUM = " & v_lTRISUser
        l_strSQL = l_strSQL & ", LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP"
        l_strSQL = l_strSQL & " WHERE EMP_NUM = " & !KEY_EMP_NUM
        l_strSQL = l_strSQL & " AND CHARGE_DATE = " & dbFormat(!KEY_CHARGE_DATE)
        l_strSQL = l_strSQL & " AND LAST_MOD_TIMESTAMP = '" & !LAST_MOD_TIMESTAMP & "'"
    End With
   
    If l_strSQL <> "" Then
       cn.Execute l_strSQL, l_lngRecsAffected
       If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
            Err.Raise trErrAppCommitConflict
       End If
    End If
    
    UpdateSchedChgPay = True
Exit Function

eh:
    GetObjectContext.SetAbort
    UpdateSchedChgPay = False
    
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = _
                Join( _
                    Array( _
                        "1", _
                        CStr(v_rsSchedChgPayRequest!KEY_EMP_NUM), _
                        Format(v_rsSchedChgPayRequest!KEY_CHARGE_DATE, "mm/dd/yyyy")) _
                    , ";")
            TrsRaiseError m_strCMODULENAME, "UpdateSchedChgPay", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateSchedChgPay", , True, , DescArgsFrom(l_strSQL), True
    End Select
    
End Function


Private Function OpenConnection() As ADODB.Connection
    Dim cm As New util_DBConnectMgr.clsuDBConnectMgr
    Dim cn As ADODB.Connection

On Error GoTo OpenConnection_EH

    Set cn = New ADODB.Connection

    With cn
        .CursorLocation = adUseClient       'Must use client to support RecordCount
        .ConnectionString = cm.GetConnectionString(m_strConnectStringReg)
                
        .Open
    End With

    Set OpenConnection = cn
    Set cm = Nothing
    Set cn = Nothing

Exit Function
OpenConnection_EH:
    Set cm = Nothing
    Set cn = Nothing

    TrsRaiseError m_strCMODULENAME, "OpenConnection"
End Function


Private Function dbFormat(ByVal v_varData As Variant) As String
    Dim l_strFormated As String, fieldType As Integer
    
    On Error GoTo eh
    
    fieldType = VarType(v_varData)
    Select Case fieldType
        Case vbString
            l_strFormated = "'" & Replace(v_varData, "'", "''") & "'"
        Case vbInteger, vbLong
            l_strFormated = CStr(v_varData)
        Case vbDate
            l_strFormated = "'" & Format(v_varData, "mm/dd/yyyy") & "'"
        Case Else
            l_strFormated = Format(v_varData)
    End Select
    dbFormat = l_strFormated
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "dbFormat", , , , DescArgsFrom(CStr(v_varData)), True
End Function




'Public Function DeleteSchedChgPmts(ByVal v_adrsSchedChgPmts) As Boolean
'    On Error GoTo DeleteSchedChgPmts_EH
'
'    Dim l_strSQL As String
'    Dim l_objCN As ADODB.Connection
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'
'    'Get Connection
'    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
'    Set l_objCN = New ADODB.Connection
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'    Set l_objCM = Nothing
'
'    DeleteSchedChgPmts = True
'
'    v_adrsSchedChgPmts.MoveFirst
'    With l_objCN
'        'Open Connection
'        .Open
'        Do While Not v_adrsSchedChgPmts.EOF
'            'Format SQL Statement
'            l_strSQL = "DELETE " & _
'                            " FROM methdba.ttr_sched_chg_pmts " & _
'                            " WHERE emp_num = " & CLng(v_adrsSchedChgPmts!EMP_NUM) & _
'                            " AND date = '" & Format(v_adrsSchedChgPmts!SCHED_DATE, "mm/dd/yyyy") & "'" & _
'                            " AND start_time = '" & v_adrsSchedChgPmts!START_TIME & "'"
'
'            If Not .Execute(l_strSQL) Then
'                DeleteSchedChgPmts = False
'                GoTo DeleteSchedChgPmts_EH
'            End If
'            v_adrsSchedChgPmts.MoveNext
'        Loop
'    End With
'
'    Set l_objCN = Nothing
'    GetObjectContext.SetComplete
'
'    Exit Function
'
'DeleteSchedChgPmts_EH:
'    GetObjectContext.SetAbort
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "DeleteSchedChgPmts", , , , DescArgsFrom(l_strSQL), True
'
'End Function
'
'
'Public Function DeleteSchedChgPmtsForRange _
'    (ByVal v_aryEmployees As Variant, _
'    ByVal v_dtStartDate As Date, _
'    ByVal v_dtEndDate As Date) As Boolean
'
'    On Error GoTo DeleteSchedChgPmtsForRange_EH
'
'    Dim l_strSQL As String
'    Dim l_objCN As ADODB.Connection
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'
'    'Get Connection
'    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
'    Set l_objCN = New ADODB.Connection
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'    Set l_objCM = Nothing
'
'    DeleteSchedChgPmtsForRange = True
'
'    With l_objCN
'        'Open Connection
'        .Open
'        'Format SQL Statement
'        l_strSQL = "DELETE " & _
'                        " FROM methdba.ttr_sched_chg_pmts " & _
'                        " WHERE " & BuildInClause("emp_num", v_aryEmployees, False) & _
'                        " AND charge_date BETWEEN " & _
'                        "'" & Format(v_dtStartDate, "mm/dd/yyyy") & "'  AND '" & Format(v_dtEndDate, "mm/dd/yyyy") & "' "
'        If Not .Execute(l_strSQL) Then
'            DeleteSchedChgPmtsForRange = False
'            GoTo DeleteSchedChgPmtsForRange_EH
'        End If
'    End With
'
'    Set l_objCN = Nothing
'    GetObjectContext.SetComplete
'
'    Exit Function
'
'DeleteSchedChgPmtsForRange_EH:
'    GetObjectContext.SetAbort
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "DeleteSchedChgPmtsForRange", , , , DescArgsFrom(l_strSQL), True
'
'End Function
'Public Function InsertSchedChgPmts _
'    (ByVal v_adrsSchedChgPmts) As Boolean
'
'    On Error GoTo InsertSchedChgPmts_EH
''*****************************************************************************************
''*  Schedule Change Payments ADO
''*****************************************************************************************
''* EMP_NUM
''* SCHEDULE_DATE
''* START_TIME
''* EXCEPTION_CD
''* CHARGE_DATE
''* PAYMENT_TYPE
''* ORIG_SCHED_CODE
''* LAST_MOD_EMP
''*****************************************************************************************
'    Dim l_strSQL As String
'    Dim l_objCN As ADODB.Connection
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'
'    'Get Connection
'    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
'    Set l_objCN = New ADODB.Connection
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'    Set l_objCM = Nothing
'
'    InsertSchedChgPmts = True
'
'    With v_adrsSchedChgPmts
'        'Open Connection
'        l_objCN.Open
'        'Format SQL Statement
'        v_adrsSchedChgPmts.MoveFirst
'        Do While Not v_adrsSchedChgPmts.EOF
'            l_strSQL = "INSERT INTO methdba.ttr_sched_chg_pmts " & _
'                            "(emp_num, date, start_time, end_time, exception_cd, " & _
'                            "charge_date, payment_type, orig_sched_cd, last_mod_emp_num) " & _
'                            "VALUES( " & v_adrsSchedChgPmts!EMP_NUM & "," & _
'                            "'" & Format(!SCHEDULE_DATE, "mm/dd/yyyy") & "' ," & _
'                            "'" & !START_TIME & "' ," & _
'                            "'" & !EXCEPTION_CD & "', " & _
'                            "'" & Format(!CHARGE_DATE, "mm/dd/yyyy") & "'," & _
'                            "'" & !PAYMENT_TYPE & "', " & _
'                            "'" & !ORIG_SCHED_CODE & "'," & _
'                            !LAST_MOD_EMP & ")"
'            If Not l_objCN.Execute(l_strSQL) Then
'                InsertSchedChgPmts = False
'                GoTo InsertSchedChgPmts_EH
'            End If
'        Loop
'    End With
'
'    Set l_objCN = Nothing
'    GetObjectContext.SetComplete
'
'    Exit Function
'
'InsertSchedChgPmts_EH:
'    GetObjectContext.SetAbort
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "InsertSchedChgPmts", , , , DescArgsFrom(l_strSQL), True
'
'End Function


