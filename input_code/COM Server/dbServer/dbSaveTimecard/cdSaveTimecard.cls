VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdSaveTimecard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdSaveTimecard"

Public Function SaveTimeCard( _
    ByVal prTimeCard As ADODB.Recordset _
    ) As Boolean

    Dim mbRetTimeEntry As Boolean
    
    Dim cn As ADODB.Connection

    On Error GoTo eh
    
    If prTimeCard.EOF Then
        SaveTimeCard = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    Set cn = OpenConnection
    
    prTimeCard.MoveFirst
    
    Do Until prTimeCard.EOF
        If IsNew(prTimeCard) Then
            InsertTimecardEntry prTimeCard, cn
        ElseIf IsModified(prTimeCard) Then
            UpdateTimecardEntry prTimeCard, cn
        ElseIf IsDeleted(prTimeCard) Then
            DeleteTimecardEntry prTimeCard, cn
        End If
        prTimeCard.MoveNext
    Loop
    
    prTimeCard.MoveFirst
    cn.Close
    Set cn = Nothing
    GetObjectContext.SetComplete
    
    SaveTimeCard = True
    
    Exit Function
    
eh:

    GetObjectContext.SetAbort
    Set cn = Nothing
    
    TrsRaiseError m_strCMODULENAME, "SaveTimeCard", , True, , , True

End Function




Private Function OpenConnection() As ADODB.Connection
  
    Dim cm As New util_DBConnectMgr.clsuDBConnectMgr
    Dim cn As ADODB.Connection
    
    On Error GoTo OpenConnection_Eh
    
    Set cn = New ADODB.Connection
    With cn
        .CursorLocation = adUseClient       'Must use client to support RecordCount
        .ConnectionString = cm.GetConnectionString
        .Open
    End With
    
    Set OpenConnection = cn
    Set cm = Nothing
    Set cn = Nothing
    Exit Function
    
OpenConnection_Eh:
    Set cm = Nothing
    Set cn = Nothing
    
    TrsRaiseError m_strCMODULENAME, "OpenConnection"
 
End Function

Private Function IsNew(prTimeCard As ADODB.Recordset) As Boolean
    
    If prTimeCard!RECORD_STATUS = "I" Then
        IsNew = True
    End If
    
End Function

Private Function IsModified(prTimeCard As ADODB.Recordset) As Boolean
    
    If prTimeCard!RECORD_STATUS = "C" Then
        IsModified = True
    End If
    
End Function

Private Function IsDeleted(prTimeCard As ADODB.Recordset) As Boolean
    
    If prTimeCard!RECORD_STATUS = "D" Then
        IsDeleted = True
    End If
    
End Function

Private Sub InsertTimecardEntry( _
            prTimeCard As ADODB.Recordset, cn As ADODB.Connection)
    Dim primKeys As Variant
    Dim ignoreFields As Variant
    Dim l_strSQL As String
    
    On Error GoTo eh
    
    If Trim(prTimeCard!TIME_STAMP) = "" Then
        ignoreFields = Array("TIME_STAMP", "INDEX_TIME", "RECORD_STATUS", "SCHED_TIME_IND", "OT_ADJUST_DESC", "ROW_ID")
    Else
        ignoreFields = Array("INDEX_TIME", "RECORD_STATUS", "SCHED_TIME_IND", "OT_ADJUST_DESC", "ROW_ID")
    End If
    l_strSQL = InsertStatement(prTimeCard, "METHDBA.TTR_ACTUAL_TIME", ignoreFields)
    If l_strSQL <> "" Then
    
'test only
        #If TRACENEON Then
            Dim strDummySelect As String
            strDummySelect = "SELECT " _
                & prTimeCard!EMP_NUM & " as EMP_NUM" _
                & ", '" & prTimeCard!CHARGE_DATE & "' as CHG_DAT" _
                & ", '" & prTimeCard!START_TIME & "' as ST_TIM" _
                & " FROM SYSIBM.SYSDUMMY1"
            cn.Execute strDummySelect
            TraceNeonMessage prTimeCard!EMP_NUM, "INSTAT", strDummySelect
        #End If
    
        cn.Execute l_strSQL
    End If
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InsertTimecardEntry", , True, , DescArgsFrom(l_strSQL), True

End Sub

Private Sub UpdateTimecardEntry( _
                prTimeCard As ADODB.Recordset, cn As ADODB.Connection)
    Dim primKeys As Variant
    Dim ignoreFields As Variant
    Dim l_strSQL As String
    
    On Error GoTo eh
    
    prTimeCard!INTERFACE_IND = ""
'    primKeys = Array("EMP_NUM", "CHARGE_DATE", "START_TIME", "INDEX_TIME")
    primKeys = Array("ROW_ID")
    If Trim(prTimeCard!TIME_STAMP) = "" Then
        ignoreFields = Array("TIME_STAMP", "INDEX_TIME", "RECORD_STATUS", "SCHED_TIME_IND", "OT_ADJUST_DESC")
    Else
        ignoreFields = Array("INDEX_TIME", "RECORD_STATUS", "SCHED_TIME_IND", "OT_ADJUST_DESC")
    End If
    l_strSQL = UpdateStatement(prTimeCard, "METHDBA.TTR_ACTUAL_TIME", primKeys, ignoreFields)
'    guard against record marked as modified but not attributes changed
    If l_strSQL <> "" Then
       cn.Execute l_strSQL
    End If
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "UpdateTimecardEntry", , True, , DescArgsFrom(l_strSQL), True

End Sub

Private Sub DeleteTimecardEntry( _
                prTimeCard As ADODB.Recordset, cn As ADODB.Connection)
    Dim primKeys As Variant
    Dim ignoreFields As Variant
    Dim l_strSQL As String
    
    On Error GoTo eh
    
    primKeys = Array("EMP_NUM", "CHARGE_DATE", "START_TIME", "INDEX_TIME")
    l_strSQL = DeleteStatement(prTimeCard, "METHDBA.TTR_ACTUAL_TIME", primKeys)
    If l_strSQL <> "" Then
        cn.Execute l_strSQL
    End If
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "DeleteTimecardEntry", , True, , DescArgsFrom(l_strSQL), True

End Sub
'***XXX***
Public Function SaveOtherPay(ByVal v_rsOtherPay As ADODB.Recordset) As Boolean

    
    Dim cn As ADODB.Connection

On Error GoTo SaveOtherPay_EH
    
    If v_rsOtherPay.BOF And v_rsOtherPay.EOF Then
        SaveOtherPay = True
        GetObjectContext.SetComplete
        
        Exit Function
    End If
    
    
    Set cn = OpenConnection
    
'  Per Prem - The client will insure that duplicates are not sent to Server
'  Per Prem - The client may reuse a deleted entry Start_Time perform deletes first
    
'  Perform Deletes
    v_rsOtherPay.MoveFirst
    Do Until v_rsOtherPay.EOF
        If IsDeleted(v_rsOtherPay) Then
            DeleteOtherPayEntry v_rsOtherPay, cn
        End If
        v_rsOtherPay.MoveNext
    Loop
    
'  Perform Changes
    v_rsOtherPay.MoveFirst
    Do Until v_rsOtherPay.EOF
        If IsModified(v_rsOtherPay) Then
            UpdateOtherPayEntry v_rsOtherPay, cn
        End If
        v_rsOtherPay.MoveNext
    Loop
    
'  Perform Inserts
    v_rsOtherPay.MoveFirst
    Do Until v_rsOtherPay.EOF
        If IsNew(v_rsOtherPay) Then
            InsertOtherPayEntry v_rsOtherPay, cn
        End If
        v_rsOtherPay.MoveNext
    Loop

    v_rsOtherPay.MoveFirst
    cn.Close
    
    Set cn = Nothing
    GetObjectContext.SetComplete
    
    SaveOtherPay = True
    
    
Exit Function
SaveOtherPay_EH:

    GetObjectContext.SetAbort
    Set cn = Nothing
    
    TrsRaiseError m_strCMODULENAME, "SaveOtherPay", , True, , , True

End Function

'***XXX***
Private Sub InsertOtherPayEntry(rsOtherPay As ADODB.Recordset, cn As ADODB.Connection)

    Dim ignoreFields As Variant
    Dim l_strSQL As String
    Dim addDefaultFields() As Variant
    
On Error GoTo InsertOtherPayEntry_EH
    
    ReDim addDefaultFields(0, 1)
    addDefaultFields(0, 0) = "FAMILY_LEAVE_IND"
    addDefaultFields(0, 1) = "'N'"
    ignoreFields = Array("TIME_STAMP", "INDEX_TIME", "RECORD_STATUS")
    l_strSQL = InsertStatement(rsOtherPay, "METHDBA.TTR_ACTUAL_TIME", ignoreFields, addDefaultFields)
    
    If l_strSQL <> "" Then
       cn.Execute l_strSQL
    End If
    
Exit Sub
InsertOtherPayEntry_EH:

    TrsRaiseError m_strCMODULENAME, "InsertTimecardEntry", , True, , DescArgsFrom(l_strSQL), True

End Sub

'***XXX***
Private Sub UpdateOtherPayEntry(rsOtherPay As ADODB.Recordset, cn As ADODB.Connection)

    Dim primKeys As Variant
    Dim ignoreFields As Variant
    Dim l_strSQL As String
    
    On Error GoTo UpdateOtherPayEntry_EH
    
'    primKeys = Array("EMP_NUM", "CHARGE_DATE", "DATE_WORKED", "INDEX_TIME")
    primKeys = Array("ROW_ID")
    ignoreFields = Array("TIME_STAMP", "INDEX_TIME", "RECORD_STATUS")
    l_strSQL = UpdateStatement(rsOtherPay, "METHDBA.TTR_ACTUAL_TIME", primKeys, ignoreFields)
    
    If l_strSQL <> "" Then
       cn.Execute l_strSQL
    End If
    
Exit Sub
UpdateOtherPayEntry_EH:
    
    TrsRaiseError m_strCMODULENAME, "UpdateOtherPayEntry", , True, , DescArgsFrom(l_strSQL), True

End Sub

'***XXX***
Private Sub DeleteOtherPayEntry(rsOtherPay As ADODB.Recordset, cn As ADODB.Connection)

    Dim primKeys As Variant
    Dim l_strSQL As String
    
    On Error GoTo DeleteOtherPayEntry_EH
    
    primKeys = Array("EMP_NUM", "CHARGE_DATE", "DATE_WORKED", "INDEX_TIME")
    l_strSQL = DeleteStatement(rsOtherPay, "METHDBA.TTR_ACTUAL_TIME", primKeys)
    
    If l_strSQL <> "" Then
       cn.Execute l_strSQL
    End If
    
Exit Sub
DeleteOtherPayEntry_EH:
    
    TrsRaiseError m_strCMODULENAME, "DeleteOtherPayEntry", , True, , DescArgsFrom(l_strSQL), True

End Sub

Public Function UpdateOTBankChargeDateOnTimeEntries( _
            ByVal v_rsUpdateRequest As ADODB.Recordset _
            ) As Boolean


    Dim cn As ADODB.Connection
    Dim l_lngEmployee As Long
    Dim l_datChargeDate As Date
    Dim l_strStartTime As String
    Dim l_ExceptionCD As String
    Dim l_OverExcCD As String
    Dim l_datOTBankChargeDate As Date
    Dim l_strSQL As String

    On Error GoTo eh



    If v_rsUpdateRequest.EOF And v_rsUpdateRequest.BOF Then
        UpdateOTBankChargeDateOnTimeEntries = True
        GetObjectContext.SetComplete
        Exit Function
    End If

    Set cn = OpenConnection
    v_rsUpdateRequest.MoveFirst
    While Not v_rsUpdateRequest.EOF
        l_strSQL = GetOTBankChargeDateSQL( _
                    v_rsUpdateRequest!EMP_NUM, _
                    v_rsUpdateRequest!CHARGE_DATE, _
                    v_rsUpdateRequest!START_TIME, _
                    v_rsUpdateRequest!EXCEPTION_CD, _
                    v_rsUpdateRequest!OVER_EXC_CD, _
                    v_rsUpdateRequest!BANK_CHARGE_DATE _
                    )
        cn.Execute l_strSQL
        v_rsUpdateRequest.MoveNext
    Wend

    GetObjectContext.SetComplete
    Set cn = Nothing

    Exit Function

eh:

    GetObjectContext.SetAbort
    Set cn = Nothing


    TrsRaiseError m_strCMODULENAME, "UpdateOTBankChargeDateOnTimeEntries", , True, , DescArgsFrom(l_strSQL), True



End Function

Private Function GetOTBankChargeDateSQL( _
            ByVal v_lngEmployee As Long, _
            ByVal v_datChargeDate As Date, _
            ByVal v_strStartTime As String, _
            ByVal v_ExceptionCD As String, _
            ByVal v_OverExcCD As String, _
            ByVal v_datOTBankChargeDate As Date _
            ) As String

    Dim l_strSQL As String

    On Error GoTo eh


    l_strSQL = "Update methdba.ttr_actual_time " _
        & " Set OT_BANK_CHRG_DT = " & DateAsEnclosedString(v_datOTBankChargeDate) _
        & " Where emp_num = " & str(v_lngEmployee) _
        & "   and charge_date = " & DateAsEnclosedString(v_datChargeDate) _
        & "   and start_time = " & EncloseWithSingleQuote(v_strStartTime) _
        & "   and exception_cd = " & EncloseWithSingleQuote(v_ExceptionCD) _
        & "   and over_exc_cd = " & EncloseWithSingleQuote(v_OverExcCD)

    GetOTBankChargeDateSQL = l_strSQL

    Exit Function

eh:

    TrsRaiseError m_strCMODULENAME, "GetOTBankChargeDateSQL", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function UpdateApprovalInd( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_aryWorkGroupNum As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            ByVal v_lngEmployeeId As Long, _
            ByVal v_bolApprovedAllFlagInd As Boolean _
            ) As Boolean
' This is invoked to update actual time as
' as approved in time collection by default. If v_bolApprovedAllFlagInd is
' set to false, update as unapproved in time collection
' WHITE173/WHITE174
    
    Dim cn As ADODB.Connection
    Dim l_strSQL As String
    Dim l_strApprovedAllFlagInd As String
    Dim l_strApprovedDate As String
    Dim l_lngApprovedByEmployee As Long
    Dim l_lngStart As Long

    On Error GoTo eh
    
    If v_lngEmployeeId < 1 Then
        Err.Raise trErrdbInvalidZerolngEmpNum
    Else
    End If
    
    l_lngStart = LBound(v_aryEmployees)
    If Val(v_aryEmployees(l_lngStart)) < 1 Then
        Err.Raise trErrdbGetEmptyArrEmployees
    Else
    End If
    
    l_lngStart = LBound(v_aryWorkGroupNum)
    If Val(v_aryWorkGroupNum(l_lngStart)) < 1 Then
        Err.Raise trErrdbGetEmptyArrWorkGroups
    Else
    End If
    
    
    Set cn = OpenConnection

    If v_bolApprovedAllFlagInd = False Then
        l_strApprovedAllFlagInd = "N"
        l_strApprovedDate = "0001-01-01"
        l_lngApprovedByEmployee = 0
    Else
        l_strApprovedAllFlagInd = "Y"           ' set to other pay when input parm v_bolOtherPayAsApproved is set to true
        l_strApprovedDate = Format(Date, "mm/dd/yyyy")
        l_lngApprovedByEmployee = v_lngEmployeeId
    End If
    
    l_strSQL = "Update methdba.ttr_actual_time " _
        & " SET APPR_IND = " & EncloseWithSingleQuote(l_strApprovedAllFlagInd) & "," _
        & " APPR_DATE = " & "'" & l_strApprovedDate & "'" & "," _
        & " APPR_BY_EMP_NUM = " & l_lngApprovedByEmployee _
        & " WHERE " & BuildInClause("EMP_NUM", v_aryEmployees, False) & " " _
        & " AND " & BuildInClause("WG_NUM", v_aryWorkGroupNum, False) & " " _
        & " AND CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "mm/dd/yyyy") & "'" _
        & " AND '" & Format(v_datEndDate, "mm/dd/yyyy") & "'" _
        & " AND NOT APPR_IND = " & EncloseWithSingleQuote(l_strApprovedAllFlagInd)
    
    cn.Execute l_strSQL
 
    Set cn = Nothing
    GetObjectContext.SetComplete
    UpdateApprovalInd = True
    
    Exit Function

eh:

    GetObjectContext.SetAbort
    Set cn = Nothing
    
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "UpdateApprovalInd", trErrdbInvalidZerolngEmpNum, , , , True
        Case trErrdbGetEmptyArrEmployees
            TrsRaiseError m_strCMODULENAME, "UpdateApprovalInd", trErrdbGetEmptyArrEmployees, , , , True
        Case trErrdbGetEmptyArrWorkGroups
            TrsRaiseError m_strCMODULENAME, "UpdateApprovalInd", trErrdbGetEmptyArrWorkGroups, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateApprovalInd", , , , DescArgsFrom(l_strSQL), True
    End Select
    
End Function


Public Function DeleteActualTime( _
    ByVal v_rsActualTime As ADODB.Recordset _
    ) As Boolean
    
'This method will only delete time entries.
'Each record in the request is a delete of a time entry.
'If you want to delete all time entries on a date, all you have to do
' is perform a SQL of the TimeEntries and pass it to this method.
'Since we always inform the user of what we have done,
'this should not be a problem.

    Dim cn As ADODB.Connection
    Dim l_strSQL As String

    On Error GoTo eh
    
    If Not RecordsExist(v_rsActualTime) Then
        DeleteActualTime = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    
    Set cn = OpenConnection
    v_rsActualTime.MoveFirst
    While Not v_rsActualTime.EOF
        DeleteTimecardEntry v_rsActualTime, cn
        v_rsActualTime.MoveNext
    Wend

    DeleteActualTime = True
    Exit Function
    
eh:

    TrsRaiseError m_strCMODULENAME, "DeleteActualTime", , , , , True

    
End Function

Public Function RemoveExceptionsFromSchedule( _
    ByVal v_varEmpChgDates As Variant, _
    Optional v_strUserId As String = "77777" _
    ) As Boolean
' This is invoked when any time is entered in time collection.
' v_varEmpChgDates should be a 2-dim array with n-number of
' distinct employee number/charge date pairs.
' WHITE117
    
    Dim cn As ADODB.Connection
    Dim l_strSQL As String
    Dim l_strDeleteSQL As String
    Dim i As Integer

    On Error GoTo eh
    
    Set cn = OpenConnection
    
    For i = LBound(v_varEmpChgDates, 1) To UBound(v_varEmpChgDates, 1)
        l_strSQL = "Insert into  METHDBA.TTR_EMP_SCH_EX_AUD " & _
            "select EMP_NUM, CHARGE_DATE, SHIFT_NUM, HOURS_OFFSET_NUM, HOURS " & _
            ", EXCEPTION_CD, OVER_EXC_CD, FMLA_IND, CASE_NUM, WG_NUM " & _
            ",JOB_NUM, RELIEF_COMP, COMP_CD, SHIFT_COMP, SCHED_CD " & _
            ",SCHED_SRC_CD, CREATED_EMP_NUM, LAST_MOD_EMP_NUM " & _
            ",CREATED_TIMESTAMP,  LAST_MOD_TIMESTAMP " & _
            "," & v_strUserId & ", current timestamp, 'TS' " & _
        "From METHDBA.TTR_EMP_SCHED_EXC " & _
        "Where emp_num = " & CStr(v_varEmpChgDates(i, 0)) & _
        "  and charge_date =  " & DateAsEnclosedString(v_varEmpChgDates(i, 1))
        
        l_strDeleteSQL = "Delete From METHDBA.TTR_EMP_SCHED_EXC " & _
            "Where emp_num = " & CStr(v_varEmpChgDates(i, 0)) & _
            "  and charge_date =  " & DateAsEnclosedString(v_varEmpChgDates(i, 1))
        cn.Execute l_strSQL
        cn.Execute l_strDeleteSQL
    Next
    
    Set cn = Nothing
    GetObjectContext.SetComplete
    RemoveExceptionsFromSchedule = True
    
    Exit Function

eh:

    GetObjectContext.SetAbort
    Set cn = Nothing
    
    TrsRaiseError m_strCMODULENAME, "RemoveExceptionsFromSchedule", , , , DescArgsFrom(l_strSQL), True
    
End Function

Public Function GetScheduleTimeStamp( _
    ByVal v_lngEmpNum As Long _
    ) As String

    Dim l_strSQL As String
    Dim l_rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    On Error GoTo eh
    
    
    l_strSQL = "select char(LAST_MOD_SCHED_TM) as LAST_MOD_SCHED_TM " & _
              " from methdba.TTR_EMP_MOD " & _
              " where emp_num = " & str(v_lngEmpNum)
              
    Set cn = OpenConnection
    Set l_rs = New ADODB.Recordset
    l_rs.Open l_strSQL, cn
    
    If l_rs.BOF And l_rs.EOF Then
        GetScheduleTimeStamp = ""
    Else
        GetScheduleTimeStamp = l_rs!LAST_MOD_SCHED_TM
    End If
    Set l_rs = Nothing
    cn.Close
        
    Exit Function

eh:
    Set l_rs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduleTimeStamp", , , , DescArgsFrom(l_strSQL), True
    
    
    
End Function

