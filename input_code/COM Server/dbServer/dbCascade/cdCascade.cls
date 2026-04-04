VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdCascade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cdCascade"


'****************************************************************
' Public Procedures
'****************************************************************

Public Function GetScheduledExceptionShifts( _
                    ByVal v_varEmployees As Variant, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date) _
                As ADODB.Recordset
' v_varEmployees should be either a long or an array of longs
'
    Dim l_strSQL As String
    Dim l_strEmpPredicate As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString()
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_varEmployees) Then
        l_strEmpPredicate = BuildInClause("EMP_NUM", v_varEmployees, False)
    Else
        l_strEmpPredicate = "EMP_NUM = " & v_varEmployees
    End If
    
    l_strSQL = "SELECT DISTINCT EMP_NUM, CHARGE_DATE, SHIFT_NUM, SCHED_CD, SCHED_SRC_CD" & _
        " FROM METHDBA.TTR_EMP_SCHED_EXC" & _
        " WHERE " & l_strEmpPredicate & _
        " AND CHARGE_DATE BETWEEN " & DateAsEnclosedString(v_datStartDate) & _
        " AND " & DateAsEnclosedString(v_datEndDate) & _
        " ORDER BY EMP_NUM, CHARGE_DATE, SHIFT_NUM"
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetScheduledExceptionShifts = l_objRS
    Set l_objRS = Nothing

    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduledExceptionShifts", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetAvailableSchedules( _
                    ByVal v_varSchedCodes As Variant) _
                As ADODB.Recordset
'
' v_varSchedCodes should be either a long or an array of longs
'
    Dim l_dbSchedule As dbSchedule.cdSchedule
    Dim l_colSchedCodes As Collection

On Error GoTo eh

    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    
    If IsArray(v_varSchedCodes) Then
        Set l_colSchedCodes = _
            CollectionFromArray(v_varSchedCodes, LBound(v_varSchedCodes))
    Else
        Set l_colSchedCodes = New Collection
        l_colSchedCodes.Add v_varSchedCodes
    End If
    
    Set GetAvailableSchedules = l_dbSchedule.GetAvailSchedInfo(l_colSchedCodes)

    GetObjectContext.SetComplete
    Set l_dbSchedule = Nothing
    Set l_colSchedCodes = Nothing
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_dbSchedule = Nothing
    Set l_colSchedCodes = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAvailableSchedules", , , , , True
End Function


Public Function GetActualTime( _
                    ByVal v_varEmployees As Variant, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date) _
                As ADODB.Recordset
' v_varEmployees should be either a long or an array of longs
'
    Dim l_dbTimecard As dbTimeCard.cdTimeCard
    Dim l_varEmpNums As Variant

On Error GoTo eh

    If IsArray(v_varEmployees) Then
        l_varEmpNums = v_varEmployees
    Else
        l_varEmpNums = Array(v_varEmployees)
    End If
    
    Set l_dbTimecard = CreateObject("dbTimeCard.cdTimeCard")
    Set GetActualTime = l_dbTimecard.GetActualTimeCascadeExists( _
                            l_varEmpNums, v_datStartDate, v_datEndDate)

    GetObjectContext.SetComplete
    Set l_dbTimecard = Nothing
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_dbTimecard = Nothing
    TrsRaiseError m_strCMODULENAME, "GetActualTime", , , , , True
End Function


Public Function GetScheduledExceptions( _
                    ByVal v_varEmployees As Variant, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date) _
                As ADODB.Recordset
'
' 10/10/01 RSG Review the resulting sql for compatibility with sql from dbSchedule,
'  If similar enough, or minor changes, just call dbSchedule
'
' v_varEmployees should be either a long or an array of longs
'
    Dim l_strSQL As String
    Dim l_strEmpPredicate As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString()
    Set l_objCM = Nothing
    
    l_objCN.Open

    If IsArray(v_varEmployees) Then
        l_strEmpPredicate = BuildInClause("A.EMP_NUM", v_varEmployees, False)
    Else
        l_strEmpPredicate = "A.EMP_NUM = " & v_varEmployees
    End If

    l_strSQL = "SELECT NE2.EMP_NUM , NE2.CHARGE_DATE, NE2.SHIFT_NUM, NE2.HOURS_OFFSET_NUM,"
    l_strSQL = l_strSQL & " NE2.HOURS, NE2.EXCEPTION_CD, NE2.OVER_EXC_CD, NE2.FMLA_IND,"
    l_strSQL = l_strSQL & " NE2.CASE_NUM , NE2.WG_NUM, NE2.JOB_NUM, NE2.RELIEF_COMP,"
    l_strSQL = l_strSQL & " NE2.COMP_CD, NE2.SHIFT_COMP, NE2.SCHED_CD, NE2.SCHED_SRC_CD,"
    l_strSQL = l_strSQL & " NE2.CREATED_EMP_NUM, NE2.CREATED_TIMESTAMP, NE2.LAST_MOD_EMP_NUM, NE2.LAST_MOD_TIMESTAMP,"
    l_strSQL = l_strSQL & " NE2.EMP_NAME, NE2.B_DEPT_EFF_DATE , NE2.DEPT_NUM, NE2.INTERFACE_TO_CD,"
    l_strSQL = l_strSQL & " NE2.EXCEPTION_TYPE , NE2.EXCEPTION_DESC,"
    l_strSQL = l_strSQL & " CHAR(LEFT(VALUE(E.EG_SUB_CD, ' '),1),1) AS GROUP_IND,"
    l_strSQL = l_strSQL & " CHAR(VALUE(F.EXCEPTION_DESC,' ')) AS OVER_EXC_DESC"
    l_strSQL = l_strSQL & " FROM"
    l_strSQL = l_strSQL & " (SELECT NE1.EMP_NUM , NE1.CHARGE_DATE, NE1.SHIFT_NUM, NE1.HOURS_OFFSET_NUM,"
    l_strSQL = l_strSQL & " NE1.HOURS, NE1.EXCEPTION_CD, NE1.OVER_EXC_CD, NE1.FMLA_IND,"
    l_strSQL = l_strSQL & " NE1.CASE_NUM , NE1.WG_NUM, NE1.JOB_NUM, NE1.RELIEF_COMP,"
    l_strSQL = l_strSQL & " NE1.COMP_CD, NE1.SHIFT_COMP, NE1.SCHED_CD, NE1.SCHED_SRC_CD,"
    l_strSQL = l_strSQL & " NE1.CREATED_EMP_NUM, NE1.CREATED_TIMESTAMP, NE1.LAST_MOD_EMP_NUM, NE1.LAST_MOD_TIMESTAMP,"
    l_strSQL = l_strSQL & " NE1.EMP_NAME, NE1.B_DEPT_EFF_DATE , NE1.DEPT_NUM, D.INTERFACE_TO_CD,"
    l_strSQL = l_strSQL & " C.EXCEPTION_TYPE , C.EXCEPTION_DESC"
    l_strSQL = l_strSQL & " FROM"
    l_strSQL = l_strSQL & " (SELECT A.EMP_NUM , A.CHARGE_DATE, A.SHIFT_NUM, A.HOURS_OFFSET_NUM, A.HOURS,"
    l_strSQL = l_strSQL & " A.EXCEPTION_CD, A.OVER_EXC_CD, A.FMLA_IND, A.CASE_NUM, A.WG_NUM, A.JOB_NUM,"
    l_strSQL = l_strSQL & " A.RELIEF_COMP, A.COMP_CD, A.SHIFT_COMP, A.SCHED_CD, A.SCHED_SRC_CD,"
    l_strSQL = l_strSQL & " A.CREATED_EMP_NUM,"
    l_strSQL = l_strSQL & " CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP,"
    l_strSQL = l_strSQL & " A.LAST_MOD_EMP_NUM,"
    l_strSQL = l_strSQL & " CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP,"
    l_strSQL = l_strSQL & " B.EMP_NAME,"
    l_strSQL = l_strSQL & " VALUE(NULLIF(B.DEPT_EFF_DATE,'0001-01-01'),DATE('0100-01-01')) AS B_DEPT_EFF_DATE,"
    l_strSQL = l_strSQL & " CASE WHEN A.CHARGE_DATE >="
    l_strSQL = l_strSQL & " VALUE(NULLIF(B.DEPT_EFF_DATE,'0001-01-01'),DATE('0100-01-01'))"
    l_strSQL = l_strSQL & " THEN B.DEPT_NUM"
    l_strSQL = l_strSQL & " WHEN B.PREV_DEPT_NUM = '' THEN B.DEPT_NUM"
    l_strSQL = l_strSQL & " ELSE B.PREV_DEPT_NUM"
    l_strSQL = l_strSQL & " END AS DEPT_NUM"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EMP_SCHED_EXC A,"
    l_strSQL = l_strSQL & " METHDBA.TTR_EMPLOYEE B"
    l_strSQL = l_strSQL & " WHERE " & l_strEmpPredicate
    l_strSQL = l_strSQL & " AND (CHARGE_DATE BETWEEN " & DateAsEnclosedString(v_datStartDate)
    l_strSQL = l_strSQL & " AND " & DateAsEnclosedString(v_datEndDate) & ")"
    l_strSQL = l_strSQL & " AND B.EMP_NUM = A.EMP_NUM"
    l_strSQL = l_strSQL & ") NE1"
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_DEPT D"
    l_strSQL = l_strSQL & " ON D.DEPT_NUM = NE1.DEPT_NUM"
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_EXCEPTION_CDS C"
    l_strSQL = l_strSQL & " ON C.EXCEPTION_CD = NE1.EXCEPTION_CD"
    l_strSQL = l_strSQL & ") NE2"
    l_strSQL = l_strSQL & " LEFT JOIN"
    l_strSQL = l_strSQL & "  (SELECT EG_SUB_CD, EXCEPTION_CD"
    l_strSQL = l_strSQL & "   FROM METHDBA.TTR_EXC_GROUP"
    l_strSQL = l_strSQL & "   WHERE EG_CD = 1) E"
    l_strSQL = l_strSQL & " ON NE2.EXCEPTION_CD = E.EXCEPTION_CD"
    l_strSQL = l_strSQL & " LEFT JOIN METHDBA.TTR_EXCEPTION_CDS F"
    l_strSQL = l_strSQL & " ON NE2.OVER_EXC_CD = F.EXCEPTION_CD"


    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetScheduledExceptions = l_objRS
    Set l_objRS = Nothing

    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduledExceptions", , , , DescArgsFrom(l_strSQL), True
End Function


Public Function UpdateCascadeExceptions( _
                    ByVal v_rsExcUpdates As ADODB.Recordset, _
                    ByVal v_lngEmpNum As Long) As Variant
' Added this new method to enable user employee number to be passed for audit.
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_lngRecsAffected As Long
    Dim l_cexRecordset As utilDataServiceLib.CExRecordset
    Dim l_objSQLBuilder As utilDataServiceLib.cuSQLBuilder
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_lngShiftNum As Long
    Dim l_strExcCd As String
    Dim l_strAuditSQL As String
    Dim primKeys As Variant
    Dim ignoreFields As Variant
    
    
On Error GoTo eh
'    Bring the deletes back into view
    FilterPendingChanges v_rsExcUpdates
    If Not RecordsExist(v_rsExcUpdates) Then
        UpdateCascadeExceptions = Empty
        Exit Function
    Else
        v_rsExcUpdates.MoveFirst
    End If
    

    ignoreFields = Array("LAST_MOD_TIMESTAMP", "CREATED_TIMESTAMP", "EMP_NAME", "B_DEPT_EFF_DATE", "DEPT_NUM", "INTERFACE_TO_CD", "EXCEPTION_TYPE", "EXCEPTION_DESC", "GROUP_IND", "OVER_EXC_DESC")


    Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString()
    Set l_objCM = Nothing
    
    l_objCN.Open

    Set l_cexRecordset = BuildEmpSchedExcProps
    Set l_objSQLBuilder = New cuSQLBuilder
    
    Do Until v_rsExcUpdates.EOF
        l_lngEmpNum = AscertainKeyValue(v_rsExcUpdates, "EMP_NUM")
        l_datChargeDate = AscertainKeyValue(v_rsExcUpdates, "CHARGE_DATE")
        l_lngShiftNum = AscertainKeyValue(v_rsExcUpdates, "SHIFT_NUM")
        l_strExcCd = AscertainKeyValue(v_rsExcUpdates, "EXCEPTION_CD")
        
        l_strAuditSQL = ""
        Select Case v_rsExcUpdates.Status
           Case adRecModified
               l_strSQL = l_objSQLBuilder.UpdateStatement(v_rsExcUpdates, l_cexRecordset)
               If l_strSQL <> "" And ExceptionCodeChanged(v_rsExcUpdates) Then
                    l_strAuditSQL = BuildExceptionAuditSQL(v_rsExcUpdates, v_lngEmpNum)
               End If
           Case adRecNew
            l_strSQL = InsertStatement(v_rsExcUpdates, "METHDBA.TTR_EMP_SCHED_EXC", ignoreFields)
               'l_strSQL = l_objSQLBuilder.InsertStatement(v_rsExcUpdates, l_cexRecordset)
           Case adRecDeleted
               l_strSQL = l_objSQLBuilder.DeleteStatement(v_rsExcUpdates, l_cexRecordset)
               If l_strSQL <> "" Then
                    l_strAuditSQL = BuildExceptionAuditSQL(v_rsExcUpdates, v_lngEmpNum)
               End If
            Case Else
               l_strSQL = ""
        End Select
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCascadeCommitConflict
            End If
            If l_strAuditSQL <> "" Then
                l_objCN.Execute l_strAuditSQL
            End If
        End If
        v_rsExcUpdates.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    Set l_cexRecordset = Nothing
    Set l_objSQLBuilder = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_cexRecordset = Nothing
    Set l_objSQLBuilder = Nothing
    
    Select Case Err.Number
        Case trErrAppCascadeCommitConflict
            Err.Source = _
                Join( _
                    Array( _
                        CStr(TrType_CA), _
                        CStr(l_lngEmpNum), _
                        Format(l_datChargeDate, "mm/dd/yyyy"), _
                        CStr(l_lngShiftNum), _
                        l_strExcCd) _
                    , ";")
            TrsRaiseError m_strCMODULENAME, "UpdateCascadeExceptions", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateCascadeExceptions", , True, , DescArgsFrom(l_strSQL), True
    End Select
    
End Function
Public Function GetCascadeExceptionUpdates( _
                    ByVal v_rsExcUpdates As ADODB.Recordset, _
                    ByVal v_lngEmpNum As Long) As Variant
' Answer a variant array of strings.

On Error GoTo eh

    GetCascadeExceptionUpdates = CascadeExceptionSQL(v_rsExcUpdates, v_lngEmpNum)
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetCascadeExceptionUpdates", , , , , True
End Function


Public Function SetEmployeeScheduleModified( _
                    ByVal v_varEmployees As Variant, _
                    Optional ByVal v_lngScheduleSource As Long = 0, _
                    Optional ByVal v_strTimeStamp As String = "CURRENT") _
                As Boolean
' v_varEmployees should be either a long or an array of longs
'
    Dim l_strSQL As String
    Dim l_lngEmpArray As Variant
    Dim l_objCN As ADODB.Connection
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strTimestamp As String
    Dim i As Integer
    Dim l_lngRecsAffected As Long

On Error GoTo eh
    
    If v_strTimeStamp = "CURRENT" Then
        l_strTimestamp = "CURRENT TIMESTAMP"
    Else
        If IsDB2Timestamp(v_strTimeStamp) Then
            l_strTimestamp = EncloseWithSingleQuote(v_strTimeStamp)
        Else
            l_strTimestamp = "CURRENT TIMESTAMP"
        End If
    End If

    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString()
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_varEmployees) Then
        l_lngEmpArray = v_varEmployees
    Else
        l_lngEmpArray = Array(v_varEmployees)
    End If
    
    For i = LBound(l_lngEmpArray) To UBound(l_lngEmpArray)
        l_strSQL = "UPDATE METHDBA.TTR_EMP_MOD" & _
            " SET LAST_MOD_SCHED_TM = " & l_strTimestamp & _
            " WHERE EMP_NUM = " & l_lngEmpArray(i)
        l_objCN.Execute l_strSQL, l_lngRecsAffected
        If l_lngRecsAffected = 0 Or l_lngRecsAffected = -1 Then
            l_strSQL = "INSERT INTO METHDBA.TTR_EMP_MOD" & _
                " (EMP_NUM, LAST_MOD_SCHED_TM)" & _
                " VALUES(" & CStr(l_lngEmpArray(i)) & ", " & l_strTimestamp & ")"
            l_objCN.Execute l_strSQL
        End If
    Next
        
    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    TrsRaiseError m_strCMODULENAME, "SetEmployeeScheduleModified", , , , DescArgsFrom(l_strSQL), True
End Function


Public Function GetEmployeeScheduleModified( _
                    ByVal v_varEmployees As Variant) _
                As ADODB.Recordset
' v_varEmployees should be either a long or an array of longs
'
    Dim l_strSQL As String
    Dim l_strEmpPredicate As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString()
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_varEmployees) Then
        l_strEmpPredicate = BuildInClause("EMP_NUM", v_varEmployees, False)
    Else
        l_strEmpPredicate = "EMP_NUM = " & v_varEmployees
    End If
    
    l_strSQL = "SELECT EMP_NUM, CHAR(LAST_MOD_SCHED_TM) AS LAST_MOD_SCHED_TM" & _
        " FROM METHDBA.TTR_EMP_MOD" & _
        " WHERE " & l_strEmpPredicate & _
        " ORDER BY EMP_NUM"
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetEmployeeScheduleModified = l_objRS
    Set l_objRS = Nothing

    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeScheduleModified", , , , DescArgsFrom(l_strSQL), True
End Function



'****************************************************************
' Private Procedures
'****************************************************************

Private Function CascadeExceptionSQL( _
                    ByVal v_rsExcUpdates As ADODB.Recordset, _
                    ByVal v_lngEmpNum As Long) As Variant
' Answer a variant array of strings or empty.
    Dim l_strSQL As String
    Dim l_cexRecordset As utilDataServiceLib.CExRecordset
    Dim l_objSQLBuilder As utilDataServiceLib.cuSQLBuilder
    Dim l_colSQL As Collection
    Dim l_strAuditSQL As String

On Error GoTo eh
'    Bring the deletes back into view
    FilterPendingChanges v_rsExcUpdates
    If Not RecordsExist(v_rsExcUpdates) Then
        CascadeExceptionSQL = Empty
        Exit Function
    Else
        v_rsExcUpdates.MoveFirst
    End If
    
    Set l_cexRecordset = BuildEmpSchedExcProps
    Set l_colSQL = New Collection

    Set l_objSQLBuilder = New cuSQLBuilder
    Do Until v_rsExcUpdates.EOF
        l_strAuditSQL = ""
        Select Case v_rsExcUpdates.Status
           Case adRecModified
               l_strSQL = l_objSQLBuilder.UpdateStatement(v_rsExcUpdates, l_cexRecordset)
               If l_strSQL <> "" And ExceptionCodeChanged(v_rsExcUpdates) Then
                    l_strAuditSQL = BuildExceptionAuditSQL(v_rsExcUpdates, v_lngEmpNum)
               End If
           Case adRecNew
               l_strSQL = l_objSQLBuilder.InsertStatement(v_rsExcUpdates, l_cexRecordset)
           Case adRecDeleted
               l_strSQL = l_objSQLBuilder.DeleteStatement(v_rsExcUpdates, l_cexRecordset)
               If l_strSQL <> "" Then
                    l_strAuditSQL = BuildExceptionAuditSQL(v_rsExcUpdates, v_lngEmpNum)
               End If
            Case Else
               l_strSQL = ""
        End Select
        If Trim(l_strSQL) <> "" Then
            l_colSQL.Add l_strSQL
        End If
        If Trim(l_strAuditSQL) <> "" Then
            l_colSQL.Add l_strAuditSQL
        End If
        v_rsExcUpdates.MoveNext
    Loop

    CascadeExceptionSQL = ArrayFromCollection(l_colSQL)
    Set l_colSQL = Nothing
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    Exit Function

eh:
    Set l_colSQL = Nothing
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    TrsRaiseError m_strCMODULENAME, "CascadeExceptionSQL", , , , DescArgsFrom(l_strSQL), True
End Function


Private Function BuildEmpSchedExcProps() As utilDataServiceLib.CExRecordset
' Added LAST_MOD_TIMESTAMP to primary key, but really just concurrency key
' used for optimistic locking.
    Dim l_cexRecordset As utilDataServiceLib.CExRecordset

On Error GoTo eh
    Set l_cexRecordset = New CExRecordset
    l_cexRecordset.SetFullTableName "METHDBA.TTR_EMP_SCHED_EXC"
    l_cexRecordset.AddFields "PRIMARYKEY", True, Array("EMP_NUM", "CHARGE_DATE", "SHIFT_NUM", "HOURS_OFFSET_NUM", "CREATED_TIMESTAMP", "LAST_MOD_TIMESTAMP")
    l_cexRecordset.AddField "LAST_MOD_TIMESTAMP", Array("SETCURRENTTIMESTAMP"), Array(1)
    l_cexRecordset.AddField "CREATED_TIMESTAMP", Array("SETCURRENTTIMESTAMP"), Array(3)

    Set BuildEmpSchedExcProps = l_cexRecordset

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildEmpSchedExcProps"
End Function


Private Function AscertainKeyValue( _
                    ByVal v_rsDBRequest As ADODB.Recordset, _
                    ByVal v_FieldName As String) As Variant

On Error GoTo eh

    If v_rsDBRequest.Status = adRecDeleted Or v_rsDBRequest.Status = adRecModified Then
        AscertainKeyValue = v_rsDBRequest.Fields(v_FieldName).OriginalValue
    ElseIf v_rsDBRequest.Status = adRecNew Then
        AscertainKeyValue = v_rsDBRequest.Fields(v_FieldName).Value
    Else
        AscertainKeyValue = 0
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "AscertainKeyValue"
End Function


Private Function ExceptionCodeChanged( _
                    ByVal v_rsException As ADODB.Recordset) As Boolean

On Error GoTo eh

    ExceptionCodeChanged = _
        v_rsException.Fields("EXCEPTION_CD").Value <> _
        v_rsException.Fields("EXCEPTION_CD").OriginalValue

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "AscertainKeyValue"
End Function


Private Function BuildExceptionAuditSQL( _
                    ByVal v_rsException As ADODB.Recordset, _
                    ByVal v_lngUserNum As Long _
                ) As String
    
    Dim l_strSQL As String
    Dim l_objField As ADODB.Field
    Dim l_intI As Integer
    Dim l_strFieldValue As String
    Dim l_intScale As Integer
    Dim l_intPrecision As Integer
    Dim l_intCount As Integer
    Dim l_strFormat As String
    
    Dim l_strColumns As String
    
On Error GoTo eh
    
'    If v_rsException.Fields.Count = 0 Then
'        BuildAuditSQL = ""
'        Exit Function
'    End If
    
    l_strColumns = _
        "(EMP_NUM, CHARGE_DATE, SHIFT_NUM, HOURS_OFFSET_NUM, HOURS, EXCEPTION_CD, OVER_EXC_CD," & _
        " FMLA_IND, CASE_NUM, WG_NUM, JOB_NUM, RELIEF_COMP, COMP_CD, SHIFT_COMP, SCHED_CD, SCHED_SRC_CD," & _
        " CREATED_EMP_NUM, LAST_MOD_EMP_NUM, CREATED_TIMESTAMP, LAST_MOD_TIMESTAMP," & _
        " AUD_EMP_NUM, AUD_TIMESTAMP, AUD_SUB_SYSTEM_IND)"

    l_strSQL = "INSERT INTO METHDBA.TTR_EMP_SCH_EX_AUD "
    l_strSQL = l_strSQL & l_strColumns
    l_strSQL = l_strSQL & " VALUES ("
    l_strSQL = l_strSQL & CStr(v_rsException.Fields("EMP_NUM").OriginalValue) & ", "
    l_strSQL = l_strSQL & DateAsString(v_rsException.Fields("CHARGE_DATE").OriginalValue, "'", True) & ", "
    l_strSQL = l_strSQL & CStr(v_rsException.Fields("SHIFT_NUM").OriginalValue) & ", "
    l_strSQL = l_strSQL & FormattedDecimal( _
                            v_rsException.Fields("HOURS_OFFSET_NUM").OriginalValue, _
                            v_rsException.Fields("HOURS_OFFSET_NUM").Precision, _
                            v_rsException.Fields("HOURS_OFFSET_NUM").NumericScale) & ", "
    l_strSQL = l_strSQL & FormattedDecimal( _
                            v_rsException.Fields("HOURS").OriginalValue, _
                            v_rsException.Fields("HOURS").Precision, _
                            v_rsException.Fields("HOURS").NumericScale) & ", "
    l_strSQL = l_strSQL & SinqleQuotedString(v_rsException.Fields("EXCEPTION_CD").OriginalValue) & ", "
    l_strSQL = l_strSQL & SinqleQuotedString(v_rsException.Fields("OVER_EXC_CD").OriginalValue) & ", "
    l_strSQL = l_strSQL & SinqleQuotedString(v_rsException.Fields("FMLA_IND").OriginalValue) & ", "
    l_strSQL = l_strSQL & CStr(v_rsException.Fields("CASE_NUM").OriginalValue) & ", "
    l_strSQL = l_strSQL & CStr(v_rsException.Fields("WG_NUM").OriginalValue) & ", "
    l_strSQL = l_strSQL & SinqleQuotedString(v_rsException.Fields("JOB_NUM").OriginalValue) & ", "
    l_strSQL = l_strSQL & SinqleQuotedString(v_rsException.Fields("RELIEF_COMP").OriginalValue) & ", "
    l_strSQL = l_strSQL & SinqleQuotedString(v_rsException.Fields("COMP_CD").OriginalValue) & ", "
    l_strSQL = l_strSQL & SinqleQuotedString(v_rsException.Fields("SHIFT_COMP").OriginalValue) & ", "
    l_strSQL = l_strSQL & CStr(v_rsException.Fields("SCHED_CD").OriginalValue) & ", "
    l_strSQL = l_strSQL & CStr(v_rsException.Fields("SCHED_SRC_CD").OriginalValue) & ", "
    l_strSQL = l_strSQL & CStr(v_rsException.Fields("CREATED_EMP_NUM").OriginalValue) & ", "
    l_strSQL = l_strSQL & CStr(v_rsException.Fields("LAST_MOD_EMP_NUM").OriginalValue) & ", "
    l_strSQL = l_strSQL & SinqleQuotedString(v_rsException.Fields("CREATED_TIMESTAMP").OriginalValue) & ", "
    l_strSQL = l_strSQL & SinqleQuotedString(v_rsException.Fields("LAST_MOD_TIMESTAMP").OriginalValue) & ", "
    l_strSQL = l_strSQL & CStr(v_lngUserNum) & ", "     'AUD_EMP_NUM
    l_strSQL = l_strSQL & "CURRENT TIMESTAMP, "         'AUD_TIMESTAMP
    l_strSQL = l_strSQL & "'CS')"                        'AUD_SUB_SYSTEM_IND - CASCADE
      
    BuildExceptionAuditSQL = l_strSQL
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "BuildExceptionAuditSQL"
End Function

