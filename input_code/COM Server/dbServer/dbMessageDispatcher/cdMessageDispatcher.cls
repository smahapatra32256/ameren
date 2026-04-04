VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdMessageDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdMessageDispatcher"


'****************************************************************
' Start Public EMPRV Procedures
'****************************************************************

Public Sub AddEMPRVTransaction( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_strDeptNum As String, _
                    ByVal v_strExcCd As String, _
                    ByVal v_datDate As Date, _
                    ByVal v_varHours As Variant, _
                    ByVal v_strTranCd As String, _
                    ByVal v_lngLocCd As Long)
'WHITE200
    Dim l_varTrans(0) As Variant
On Error GoTo eh
    
    l_varTrans(0) = Array( _
            v_lngEmpNum, v_strDeptNum, v_strExcCd, v_datDate, _
            v_varHours, v_strTranCd, v_lngLocCd)
    InsertEMPRVTransactions (l_varTrans)
    
    GetObjectContext.SetComplete
    Exit Sub

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "AddEMPRVTransaction", , , , , True
End Sub
Public Sub AddEMPRVTransactions(ByVal v_varTranRecs As Variant)
    
    Dim l_strSQL As String
    Dim i As Integer

On Error GoTo eh

    If Not IsArray(v_varTranRecs) Then
        Err.Raise vbObjectError + 1001      'Invalid argument - not array
    Else
        If ArraySize(v_varTranRecs) < 1 Then
            Err.Raise vbObjectError + 1002  'Invalid argument - empty array
        End If
    End If
    InsertEMPRVTransactions v_varTranRecs
    
    GetObjectContext.SetComplete
    Exit Sub

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "AddEMPRVTransactions", , , , , True
End Sub
Public Sub AddEMPRVJobTasks(ByVal v_rsJobTaskEntries As ADODB.Recordset)

    Dim l_objCN As New ADODB.Connection
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    If v_rsJobTaskEntries.BOF And v_rsJobTaskEntries.EOF Then
        GetObjectContext.SetComplete
        Exit Sub
    End If
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("EMPRV")
    'l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    
    l_objCN.Open

    If Not v_rsJobTaskEntries.EOF Then
        v_rsJobTaskEntries.MoveFirst
    End If
    
    Do Until v_rsJobTaskEntries.EOF
        InsertEMPRVJobTasks v_rsJobTaskEntries, l_objCN
        v_rsJobTaskEntries.MoveNext
    Loop

    l_objCN.Close
    
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    
Exit Sub

eh:
    GetObjectContext.SetAbort
    Set l_objCN = Nothing
    TrsRaiseError m_strCMODULENAME, "AddEMPRVJobTaskTransactions", , , , , True
End Sub
Private Sub InsertEMPRVJobTasks(v_rsJobTaskEntries As ADODB.Recordset, cn As ADODB.Connection)

    Dim ignoreFields As Variant
    Dim strSQL As String
    Dim errArgs As New Collection
    Dim addDefaultFields() As Variant
    
On Error GoTo InsertEMPRVJobTasks_EH
    
    'strSQL = "INSERT INTO METHDBA.TTR_EMPRV_HRS(" & _
        "EMP_NUM, DATE_WORKED, CHARGE_DATE, START_TIME, END_TIME, WG_NUM, JOB_NBR " & _
        ", JOBTASK_NBR, JOBTASK_STEP_NBR, EMPRV_DATABASE_ID, EXCEPTION_CD " & _
        ", DISTRIBUTION_ID, HOURS, OT_IND, CHG_IND)" & _
        " VALUES(" & _
            CStr(v_rsJobTaskEntries.Fields("EMP_NUM")) & ", " & _
            DateAsEnclosedString(v_rsJobTaskEntries.Fields("DATE_WORKED")) & ", " & _
            DateAsEnclosedString(v_rsJobTaskEntries.Fields("CHARGE_DATE")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("START_TIME")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("END_TIME")) & ", " & _
            CStr(v_rsJobTaskEntries.Fields("WG_NUM")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("JOB_NBR")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("JOBTASK_NBR")) & ", " & _
            CStr(v_rsJobTaskEntries.Fields("JOBTASK_STEP_NBR")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("EMPRV_DATABASE_ID")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("EXCEPTION_CD")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("DISTRIBUTION_ID")) & ", " & _
            CStr(v_rsJobTaskEntries.Fields("HOURS")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("OT_IND")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("CHG_IND")) & ")"
            
    strSQL = "insert into dbo.tpotrishours" & _
             "(emplnbr,chrgdt,sttm,workeddt,endtm,wrkgrpnbr,fk_jobtaskjobnbr," & _
             "fk_jobtasktasknbr,fk_jobstepseqnbr,fk_emprvdbcode,excptcd,distrid," & _
             "rptdhrs,otind,chgind,fk_costcatcode)" & _
             "values(" & _
                CStr(v_rsJobTaskEntries.Fields("EMP_NUM")) & "," & _
                DateAsEnclosedString(v_rsJobTaskEntries.Fields("CHARGE_DATE")) & ", " & _
                SinqleQuotedString(v_rsJobTaskEntries.Fields("START_TIME")) & ", " & _
                DateAsEnclosedString(v_rsJobTaskEntries.Fields("DATE_WORKED")) & ", " & _
                SinqleQuotedString(v_rsJobTaskEntries.Fields("END_TIME")) & ", " & _
                CStr(v_rsJobTaskEntries.Fields("WG_NUM")) & ", " & _
                SinqleQuotedString(v_rsJobTaskEntries.Fields("JOB_NBR")) & ", " & _
                SinqleQuotedString(v_rsJobTaskEntries.Fields("JOBTASK_NBR")) & ", " & _
                CStr(v_rsJobTaskEntries.Fields("JOBTASK_STEP_NBR")) & ", " & _
                SinqleQuotedString(v_rsJobTaskEntries.Fields("EMPRV_DATABASE_ID")) & ", " & _
                SinqleQuotedString(v_rsJobTaskEntries.Fields("EXCEPTION_CD")) & ", " & _
                SinqleQuotedString(v_rsJobTaskEntries.Fields("DISTRIBUTION_ID")) & ", " & _
                CStr(v_rsJobTaskEntries.Fields("HOURS")) & ", " & _
                SinqleQuotedString(v_rsJobTaskEntries.Fields("OT_IND")) & ", " & _
                SinqleQuotedString(v_rsJobTaskEntries.Fields("CHG_IND")) & ", " & _
                SinqleQuotedString(v_rsJobTaskEntries.Fields("DEPT_NUM")) & ")"

    If strSQL <> "" Then
        cn.Execute strSQL
    End If
    
Exit Sub
InsertEMPRVJobTasks_EH:

    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "InsertEMPRVJobTasks", , True, , errArgs, True

End Sub

Private Function OpenConnection() As ADODB.Connection
  
    Dim cm As New util_DBConnectMgr.clsuDBConnectMgr
    Dim cn As ADODB.Connection
    
On Error GoTo OpenConnection_EH
    
    Set cn = New ADODB.Connection
    
    With cn
        .CursorLocation = adUseClient       'Must use client to support RecordCount
        .ConnectionString = cm.GetConnectionString()
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


'****************************************************************
' End Public EMPRV Procedures
'****************************************************************



'****************************************************************
' Private Procedures
'****************************************************************

Private Sub InsertEMPRVTransactions(ByVal v_varTranRecs As Variant)
'v_varTranRecs must be a varian array of variant arrays where each
' arrays contains v_lngEmpNum, v_strDeptNum, v_strExcCd, v_datDate,
' v_varHours, v_strTranCd, v_lngLocCd
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_intLowBound As Integer
    Dim i As Integer

On Error GoTo eh

    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("EMPRV")
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    For i = LBound(v_varTranRecs) To UBound(v_varTranRecs)
        l_intLowBound = LBound(v_varTranRecs(i))
        'l_strSQL = InsertEMPRVTranSQL( _
                    v_varTranRecs(i)(l_intLowBound), _
                    v_varTranRecs(i)(l_intLowBound + 1), _
                    v_varTranRecs(i)(l_intLowBound + 2), _
                    v_varTranRecs(i)(l_intLowBound + 3), _
                    v_varTranRecs(i)(l_intLowBound + 4), _
                    v_varTranRecs(i)(l_intLowBound + 5), _
                    v_varTranRecs(i)(l_intLowBound + 6))
                    
        l_strSQL = "insert into dbo.tpo_tris_exceptions" & _
                 "(postind,empnbr,deptcode,excpcode,excpdt,excphrs,trancode,loccode, " & _
                    " savedt, postdt) " & _
                 "values(" & _
                    "'0', " & _
                    CStr(v_varTranRecs(i)(l_intLowBound)) & "," & _
                    SinqleQuotedString(v_varTranRecs(i)(l_intLowBound + 1)) & "," & _
                    SinqleQuotedString(v_varTranRecs(i)(l_intLowBound + 2)) & "," & _
                    DateAsEnclosedString(v_varTranRecs(i)(l_intLowBound + 3)) & "," & _
                    CStr(v_varTranRecs(i)(l_intLowBound + 4)) & "," & _
                    SinqleQuotedString(v_varTranRecs(i)(l_intLowBound + 5)) & "," & _
                    CStr(v_varTranRecs(i)(l_intLowBound + 6)) & "," & _
                    " getdate(), null)"
                    
                    'DateAsEnclosedString (v_rsJobTaskEntries.Fields("CHARGE_DATE")) & ", " & _
                    'SinqleQuotedString(v_rsJobTaskEntries.Fields("START_TIME")) & ", " & _
                    'DateAsEnclosedString(v_rsJobTaskEntries.Fields("DATE_WORKED")) & ", " & _
                    'SinqleQuotedString(v_rsJobTaskEntries.Fields("END_TIME")) & ", " & _
                    'CStr(v_rsJobTaskEntries.Fields("WG_NUM")) & ", " & _
                    'SinqleQuotedString(v_rsJobTaskEntries.Fields("JOB_NBR")) & ", " & _
                    'SinqleQuotedString(v_rsJobTaskEntries.Fields("JOBTASK_NBR")) & ", " & _
                    'CStr(v_rsJobTaskEntries.Fields("JOBTASK_STEP_NBR")) & ", " & _
                    'SinqleQuotedString(v_rsJobTaskEntries.Fields("EMPRV_DATABASE_ID")) & ", " & _
                    'SinqleQuotedString(v_rsJobTaskEntries.Fields("EXCEPTION_CD")) & ", " & _
                    'SinqleQuotedString(v_rsJobTaskEntries.Fields("DISTRIBUTION_ID")) & ", " & _
                    'CStr(v_rsJobTaskEntries.Fields("HOURS")) & ", " & _
                    'SinqleQuotedString(v_rsJobTaskEntries.Fields("OT_IND")) & ", " & _
                    'SinqleQuotedString(v_rsJobTaskEntries.Fields("CHG_IND")) & ", " & _
                    'SinqleQuotedString(v_rsJobTaskEntries.Fields("DEPT_NUM")) & ")"

        l_objCN.Execute l_strSQL
    Next

    l_objCN.Close
    Set l_objCN = Nothing

    Exit Sub

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    'TrsRaiseError m_strCMODULENAME, "InsertEMPRVTransactions", , , , DescArgsFrom(l_strSQL), True
End Sub

Private Function InsertEMPRVTranSQL( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_strDeptNum As String, _
                    ByVal v_strExcCd As String, _
                    ByVal v_datDate As Date, _
                    ByVal v_varHours As Variant, _
                    ByVal v_strTranCd As String, _
                    ByVal v_lngLocCd As Long) As String
    Dim l_strSQL As String
    
On Error GoTo eh

    l_strSQL = "INSERT INTO METHDBA.TTR_EMPRV_EXC(" & _
        "EMP_NUM, DEPT_NUM, EXCEPTION_CD, DATE, HOURS, TRAN_CODE, LOC_CD)" & _
        " VALUES(" & _
            CStr(v_lngEmpNum) & ", " & _
            SinqleQuotedString(v_strDeptNum) & ", " & _
            SinqleQuotedString(v_strExcCd) & ", " & _
            DateAsEnclosedString(v_datDate) & ", " & _
            CStr(v_varHours) & ", " & _
            SinqleQuotedString(v_strTranCd) & ", " & _
            CStr(v_lngLocCd) & ")"
            
    InsertEMPRVTranSQL = l_strSQL
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "InsertEMPRVTranSQL", , , , DescArgsFrom(l_strSQL)
End Function


Public Sub AddMAXIMOJobTasks(ByVal v_rsJobTaskEntries As ADODB.Recordset)

    Dim l_objCN As New ADODB.Connection
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    If v_rsJobTaskEntries.BOF And v_rsJobTaskEntries.EOF Then
        GetObjectContext.SetComplete
        Exit Sub
    End If
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    
    l_objCN.Open

    If Not v_rsJobTaskEntries.EOF Then
        v_rsJobTaskEntries.MoveFirst
    End If
    
    Do Until v_rsJobTaskEntries.EOF
        InsertMAXIMOJobTasks v_rsJobTaskEntries, l_objCN
        v_rsJobTaskEntries.MoveNext
    Loop

    l_objCN.Close
    
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    
Exit Sub

eh:
    GetObjectContext.SetAbort
    Set l_objCN = Nothing
    TrsRaiseError m_strCMODULENAME, "AddMAXIMOJobTaskTransactions", , , , , True
End Sub
Private Sub InsertMAXIMOJobTasks(v_rsJobTaskEntries As ADODB.Recordset, cn As ADODB.Connection)

    Dim ignoreFields As Variant
    Dim strSQL As String
    Dim errArgs As New Collection
    Dim addDefaultFields() As Variant
    
On Error GoTo InsertMAXIMOJobTasks_EH
    
    strSQL = "INSERT INTO METHDBA.TTR_ASPIRE_MAXIMO_TIME_CHANGES(" & _
        "EMP_NUM, DATE_WORKED, START_TIME, END_TIME, JOB_NBR " & _
        ", JOBTASK_NBR, EMPRV_DATABASE_ID, EXCEPTION_CD, DEPT_NUM " & _
        ", HOURS, RECORD_STATUS, SRC_ROW_ID)" & _
        " VALUES(" & _
            CStr(v_rsJobTaskEntries.Fields("EMP_NUM")) & ", " & _
            DateAsEnclosedString(v_rsJobTaskEntries.Fields("DATE_WORKED")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("START_TIME")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("END_TIME")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("JOB_NBR")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("JOBTASK_NBR")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("EMPRV_DATABASE_ID")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("EXCEPTION_CD")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("DEPT_NUM")) & ", " & _
            CStr(v_rsJobTaskEntries.Fields("HOURS")) & ", " & _
            SinqleQuotedString(v_rsJobTaskEntries.Fields("CHG_IND")) & ", " & _
            CStr(v_rsJobTaskEntries.Fields("SRC_ROW_ID")) & " ) "
   

    If strSQL <> "" Then
        cn.Execute strSQL
    End If
    
Exit Sub
InsertMAXIMOJobTasks_EH:

    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "InsertMAXIMOJobTasks", , True, , errArgs, True

End Sub


