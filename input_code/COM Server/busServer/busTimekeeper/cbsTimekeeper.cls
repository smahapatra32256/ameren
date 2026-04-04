VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsTimekeeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'this component does not raise errors
Private Enum TrTimekeeperErrorEnum
    trErrValidationError = 30110 + vbObjectError
    trErrAppendEntryError
End Enum

Private Enum TrsPrivateTimecardErrorEnum
    trLockedBypayroll = 30103 + vbObjectError
    trProcessedBypayroll = 30104 + vbObjectError
End Enum

Dim m_rsEmployeeData As ADODB.Recordset
Dim m_rsCollPeriodData As ADODB.Recordset
'for older clients (6.0.0) that are not sending this column to the server
Private m_blnPreArrangedOTIndMissing As Boolean
Private m_blnIsPOETActive As Boolean
Private m_blnIsOTTCEffective As Boolean

Private Const mstrCMODULENAME = "cbsTimekeeper"

Private Property Get PreArrangedOTIndMissing() As Boolean
    PreArrangedOTIndMissing = m_blnPreArrangedOTIndMissing
End Property

Private Property Let PreArrangedOTIndMissing(blnPreArrangedOTIndMissing As Boolean)
    m_blnPreArrangedOTIndMissing = blnPreArrangedOTIndMissing
End Property

Private Sub AppendEntryError( _
            ByRef r_rsErrors As ADODB.Recordset, _
            ByRef r_rsTimeentryRequest As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDescription As String, _
            Optional ByVal v_strErrorSource As String)

    Dim l_varEmp_Num As Long
    Dim l_varChargeDate As Date
    Dim l_varDateWorked As Date
    Dim l_varStartTime As String
    Dim l_varEndTime As String
    Dim l_varHours As Variant
    Dim l_varExceptionCode As String
    Dim l_varOverExcCode As String
    Dim l_varErrorSource As String
    
    On Error GoTo eh
    
    l_varErrorSource = IIf(v_strErrorSource = "", mstrCMODULENAME, v_strErrorSource)
    
    If r_rsTimeentryRequest.BOF Or r_rsTimeentryRequest.EOF Then
        AppendError _
            r_rsErrors, 0, Date, Date, "00:00", "00:00", 0, "", "", _
            v_lngErrorNumber, v_strErrorType, _
            v_strErrorDescription, l_varErrorSource
        Exit Sub
    End If

    If IsNull(r_rsTimeentryRequest!EMP_NUM) Then
        l_varEmp_Num = r_rsTimeentryRequest!KEY_EMP_NUM
    Else
        l_varEmp_Num = r_rsTimeentryRequest!EMP_NUM
    End If
    
    If IsNull(r_rsTimeentryRequest!CHARGE_DATE) Then
        l_varChargeDate = r_rsTimeentryRequest!KEY_CHARGE_DATE
    Else
        l_varChargeDate = r_rsTimeentryRequest!CHARGE_DATE
    End If
        
    If IsNull(r_rsTimeentryRequest!DATE_WORKED) Then
        l_varDateWorked = Date
    Else
        l_varDateWorked = r_rsTimeentryRequest!DATE_WORKED
    End If
    
    If IsNull(r_rsTimeentryRequest!START_TIME) Then
        l_varStartTime = ""
    Else
        l_varStartTime = r_rsTimeentryRequest!START_TIME
    End If
        
    If IsNull(r_rsTimeentryRequest!END_TIME) Then
        l_varEndTime = ""
    Else
        l_varEndTime = r_rsTimeentryRequest!END_TIME
    End If
    
    If IsNull(r_rsTimeentryRequest!Hours) Then
        l_varHours = CDec(0)
    Else
        l_varHours = r_rsTimeentryRequest!Hours
    End If
    
    If IsNull(r_rsTimeentryRequest!EXCEPTION_CD) Then
        l_varExceptionCode = ""
    Else
        l_varExceptionCode = r_rsTimeentryRequest!EXCEPTION_CD
    End If
    
    If IsNull(r_rsTimeentryRequest!OVER_EXC_CD) Then
        l_varOverExcCode = ""
    Else
        l_varOverExcCode = r_rsTimeentryRequest!OVER_EXC_CD
    End If
    
    AppendError _
        r_rsErrors, _
        l_varEmp_Num, _
        l_varChargeDate, _
        l_varDateWorked, _
        l_varStartTime, _
        l_varEndTime, _
        l_varHours, _
        l_varExceptionCode, _
        l_varOverExcCode, _
        v_lngErrorNumber, _
        v_strErrorType, _
        v_strErrorDescription, _
        l_varErrorSource

    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "AppendEntryError", _
        trErrbusAppendEntryError, True, , , True
End Sub

Private Sub AppendEntryErrorOtherPay( _
            ByRef r_rsErrors As ADODB.Recordset, _
            ByRef r_rsRequest As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDescription As String, _
            Optional ByVal v_strErrorSource As String)

    Dim l_varEmp_Num As Long
    Dim l_varChargeDate As Date
    Dim l_varDateWorked As Date
    Dim l_varStartTime As String
    Dim l_varEndTime As String
    Dim l_varHours As Variant
    Dim l_varExceptionCode As String
    Dim l_varOverExcCode As String
    Dim l_varErrorSource As String
    
    On Error GoTo AppendEntryErrorOtherPay_EH
    
    l_varErrorSource = IIf(v_strErrorSource = "", mstrCMODULENAME, v_strErrorSource)
    
    If r_rsRequest.BOF Or r_rsRequest.EOF Then
        AppendError _
            r_rsErrors, 0, Date, Date, "00:00", "00:00", 0, "", "", _
            v_lngErrorNumber, v_strErrorType, _
            v_strErrorDescription, l_varErrorSource
        Exit Sub
    End If

    If IsNull(r_rsRequest!EMP_NUM) Then
        l_varEmp_Num = 0
    Else
        l_varEmp_Num = r_rsRequest!EMP_NUM
    End If
    
    If IsNull(r_rsRequest!CHARGE_DATE) Then
        l_varChargeDate = Date
    Else
        l_varChargeDate = r_rsRequest!CHARGE_DATE
    End If
        
    l_varDateWorked = Date
    
    If IsNull(r_rsRequest!START_TIME) Then
        l_varStartTime = ""
    Else
        l_varStartTime = r_rsRequest!START_TIME
    End If
        
    l_varEndTime = ""
    
    If IsNull(r_rsRequest!Hours) Then
        l_varHours = CDec(0)
    Else
        l_varHours = r_rsRequest!Hours
    End If
    
    If IsNull(r_rsRequest!EXCEPTION_CD) Then
        l_varExceptionCode = ""
    Else
        l_varExceptionCode = r_rsRequest!EXCEPTION_CD
    End If
    
    l_varOverExcCode = ""
    
    AppendError _
        r_rsErrors, _
        l_varEmp_Num, _
        l_varChargeDate, _
        l_varDateWorked, _
        l_varStartTime, _
        l_varEndTime, _
        l_varHours, _
        l_varExceptionCode, _
        l_varOverExcCode, _
        v_lngErrorNumber, _
        v_strErrorType, _
        v_strErrorDescription, _
        l_varErrorSource

    Exit Sub
    
AppendEntryErrorOtherPay_EH:
    TrsRaiseError mstrCMODULENAME, "AppendEntryErrorOtherPay", _
        trErrbusAppendEntryError, True, , , True
End Sub

Public Function ProcessApproveWGTime( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_aryWorkgroups As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            ByVal v_strApprovalInd As String, _
            ByVal v_lngTRISEmployeeNum As Long, _
            ByRef r_strReturnMessage As String _
        ) As Boolean
    Dim lobjContext As ObjectContext
    Dim lobjSaveTimecard As dbSaveTimecard.cdSaveTimecard
    Dim lblnApprovalIndicator As Boolean
    Dim lblnResult As Boolean
        
    Set lobjContext = GetObjectContext
    If UCase(Left(LTrim(v_strApprovalInd), 1)) = "Y" Then
        lblnApprovalIndicator = True
    Else
        lblnApprovalIndicator = False
    End If
    Set lobjSaveTimecard = _
        CreateObject("dbSaveTimecard.cdSaveTimecard")
    lblnResult = _
            lobjSaveTimecard.UpdateApprovalInd( _
                v_aryEmployees, _
                v_aryWorkgroups, _
                v_datStartDate, _
                v_datEndDate, _
                v_lngTRISEmployeeNum, _
                lblnApprovalIndicator)
    If lblnResult Then
        lobjContext.SetComplete
        r_strReturnMessage = ""
        ProcessApproveWGTime = True
    Else
        lobjContext.SetAbort
        r_strReturnMessage = "Update approval by workgroup failed"
        ProcessApproveWGTime = False
    End If
    Exit Function
        
eh:
    lobjContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "ProcessApproveWGTime", , , , , True
        
End Function

Public Function ProcessTimeEntries( _
        ByVal v_rsTimeEntriesRequest As ADODB.Recordset, _
        ByVal v_rsOtherPay As ADODB.Recordset, _
        ByVal v_rsECDrawTime As ADODB.Recordset, _
        ByVal v_startDate As Date, _
        ByVal v_endDate As Date, _
        ByVal l_rsEmployees As ADODB.Recordset, _
        ByVal l_rsEmpWorkgroups As ADODB.Recordset, _
        ByVal v_rsSchedules As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset, _
        ByVal v_UserId As String) As Boolean
 
TrsRaiseError mstrCMODULENAME, "ProcessTimeEntries", trErrObsoleteMethod, , , , True

'
' Per discussion with Greg B. and Prem C. 5/2/00
' Considering the clients unit of work may be divided when
'  processed, error communication with the client will not be
'  handled by raising errors.
'
' Answer True when all time entries requests have been processed.
'  If any time entry request fails return False.
' Warnings will not have any effect on the return value.
'
' This component will trap all errors and return the error
'  information in the error recordset.  This component will
'  not re-raise errors to the client.
'
'    Dim l_objTimecard As busTimecard.cbsTimecard
'    Dim l_objCtx As ObjectContext
'    Dim l_blnResult As Boolean
'    Dim l_objErrors As ADODB.Recordset
'    Dim l_colEmployees As cuSet
''    Dim l_rsEmployees As ADODB.Recordset
''    Dim l_rsEmpWorkgroups As ADODB.Recordset
'    Dim l_rsTimecardEntries As ADODB.Recordset
'    Dim l_rsOtherpayEntries As ADODB.Recordset
'    Dim l_rsScheduleChgPmts As ADODB.Recordset
'    Dim l_strEmpNum As Variant
'    Dim l_rsAccrualBankUpdates As ADODB.Recordset
'    Dim l_rsOvertimeBankUpdates As ADODB.Recordset
'    Dim l_colEMPRVInterface As Collection
'    Dim l_rsEMPRVJobTaskEntries As ADODB.Recordset
'    Dim l_objDistributionMgr As busDistributionMgr.cbsDistributionMgr
'    Dim l_blnDistributionValidationSuccessful As Boolean
'    Dim l_blnOtherPayInput As Boolean
'    Dim l_aryEmployees As Variant
'    Dim l_rsEmployeeData As ADODB.Recordset
'    Dim l_rsCollPeriodData As ADODB.Recordset
'    Dim l_lngEmpNum As Long
'
'    On Error GoTo eh
'
'    'THIS CODE HAS BEEN MOVED TO PROCESSTIMEENTRIES2
'
'xit:
'    If Not l_objErrors Is Nothing Then
'        If l_objErrors.RecordCount = 0 Then
'            Set r_rsErrors = Nothing
'        Else
'            Set r_rsErrors = l_objErrors
'        End If
'    Else
'        Set r_rsErrors = Nothing
'    End If
'
'    Set l_objTimecard = Nothing
'
'    #If TRTRACE Then
'        TraceMessage "", "End ProcessTimeEntries"
'    #End If
'
'    #If TRACENEON Then
'        TraceNeonMessage CStr(lngTraceEmpNum), "PROCTE", "EXIT"
'    #End If
'    Exit Function
'
'eh:
'    l_objCtx.SetAbort
'    Set r_rsErrors = Nothing
'    TrsRaiseError mstrCMODULENAME, "ProcessTimeEntries", , , , , True
   
End Function

'//still used by TRIS 6.6 clients
Public Function ProcessTimeEntries2( _
        ByVal v_rsTimeEntriesRequest As ADODB.Recordset, _
        ByVal v_rsOtherPay As ADODB.Recordset, _
        ByVal v_rsECDrawTime As ADODB.Recordset, _
        ByVal v_rsJobTaskSteps As ADODB.Recordset, _
        ByVal v_startDate As Date, _
        ByVal v_endDate As Date, _
        ByVal l_rsEmployees As ADODB.Recordset, _
        ByVal l_rsEmpWorkgroups As ADODB.Recordset, _
        ByVal v_rsSchedules As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset, _
        ByVal v_UserId As String) As Boolean
        
ProcessTimeEntries2 = ProcessTimeEntries3(v_rsTimeEntriesRequest, v_rsOtherPay, v_rsJobTaskSteps, _
                                        v_startDate, v_endDate, l_rsEmployees, l_rsEmpWorkgroups, v_rsSchedules, r_rsErrors, v_UserId)

End Function

Public Function ProcessTimeEntries3( _
        ByVal v_rsTimeEntriesRequest As ADODB.Recordset, _
        ByVal v_rsOtherPay As ADODB.Recordset, _
        ByVal v_rsJobTaskSteps As ADODB.Recordset, _
        ByVal v_startDate As Date, _
        ByVal v_endDate As Date, _
        ByVal l_rsEmployees As ADODB.Recordset, _
        ByVal l_rsEmpWorkgroups As ADODB.Recordset, _
        ByVal v_rsSchedules As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset, _
        ByVal v_UserId As String) As Boolean
'
' Answer True when all time entries requests have been processed.
'  If any time entry request fails return False.
' Warnings will not have any effect on the return value.
'
' This component will trap all errors and return the error
'  information in the error recordset.  This component will
'  not re-raise errors to the client.
'
    Dim l_objTimecard As busTimecard.cbsTimecard
    Dim l_objCtx As ObjectContext
    Dim l_blnResult As Boolean
    Dim l_objErrors As ADODB.Recordset
    Dim l_colEmployees As cuSet
'    Dim l_rsEmployees As ADODB.Recordset
'    Dim l_rsEmpWorkgroups As ADODB.Recordset
    Dim l_rsTimecardEntries As ADODB.Recordset
    Dim l_rsOtherpayEntries As ADODB.Recordset
    Dim l_rsScheduleChgPmts As ADODB.Recordset
    Dim l_strEmpNum As Variant
    Dim l_rsAccrualBankUpdates As ADODB.Recordset
    Dim l_rsOvertimeBankUpdates As ADODB.Recordset
    Dim l_colEMPRVInterface As Collection
    Dim l_rsEMPRVJobTaskEntries As ADODB.Recordset
    Dim l_rsMAXIMOJobTaskEntries As ADODB.Recordset
    Dim l_objDistributionMgr As busDistributionMgr.cbsDistributionMgr
    Dim l_blnDistributionValidationSuccessful As Boolean
    Dim l_blnOtherPayInput As Boolean
    Dim l_aryEmployees As Variant
    Dim l_rsEmployeeData As ADODB.Recordset
    Dim l_rsCollPeriodData As ADODB.Recordset
    Dim l_lngEmpNum As Long
    Dim jimb As Integer

    On Error GoTo eh
    
    #If TRACENEON Then
        Dim lngTraceEmpNum As Long
        lngTraceEmpNum = v_rsTimeEntriesRequest!KEY_EMP_NUM
        TraceNeonMessage CStr(lngTraceEmpNum), "PROCTE", "BEGIN"
    #End If
    
    #If TRTRACE Then
        TraceMessage "", "Begin ProcessTimeEntries3"
    #End If
    
    Set l_objCtx = GetObjectContext
    Set l_objErrors = NewErrorset
    
    #If TRTRACE Then
        TraceMessage "", "Begin CheckForInput"
    #End If
    If Not CheckForInput(v_rsTimeEntriesRequest, v_rsOtherPay, l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries3 = False
        GoTo xit
    End If
    #If TRTRACE Then
        TraceMessage "", "End CheckForInput"
    #End If
    
    #If TRTRACE Then
        TraceMessage "", "Begin Validate"
    #End If
    
    'for older clients (6.0.0) that are not sending this column to the server
    PreArrangedOTIndMissing = Not FieldExists("PREARRANGED_OT_IND", v_rsTimeEntriesRequest)
    
    If Not ValidateInput(v_rsTimeEntriesRequest, l_objErrors) Or _
        Not ValidateInputOtherPay(v_rsOtherPay, l_objErrors) Or _
        Not ValidateEMPRVJobTaskTime(v_rsTimeEntriesRequest, v_rsJobTaskSteps, l_objErrors) Then
            l_objCtx.SetAbort
            ProcessTimeEntries3 = False
            GoTo xit
    End If
    
    #If TRTRACE Then
        TraceMessage "", "End Validate"
    #End If
    
    Set l_colEmployees = GetEmployeeList(v_rsTimeEntriesRequest, v_rsOtherPay)
    #If TRTRACE Then
        TraceMessage "", "End Validate"
    #End If

    #If TRTRACE Then
        TraceMessage "", "Begin GetEmployeeState"
    #End If
    
    l_blnOtherPayInput = RecordsExist(v_rsOtherPay)
    GetEmployeeState l_colEmployees, v_startDate, v_endDate, l_blnOtherPayInput, l_rsEmployees, l_rsEmpWorkgroups, l_rsTimecardEntries, l_rsOtherpayEntries, v_rsSchedules, l_rsScheduleChgPmts
    #If TRTRACE Then
        TraceMessage "", "End GetEmployeeState"
    #End If
    
    #If TRTRACE Then
        TraceMessage "", "Begin Validate Distribution IDs"
    #End If
    
    Set l_objDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    l_blnDistributionValidationSuccessful = _
            l_objDistributionMgr.ValidateDistributionIdsFromTimeEntries2( _
                v_rsTimeEntriesRequest, _
                v_rsOtherPay, _
                l_rsEmployees, _
                l_rsEmpWorkgroups, _
                v_startDate, _
                v_endDate, _
                l_objErrors, _
                v_rsJobTaskSteps)
    
    '''''l_blnDistributionValidationSuccessful = True
    
    If Not l_blnDistributionValidationSuccessful Then
        Set r_rsErrors = l_objErrors
        ProcessTimeEntries3 = False
        Exit Function
    End If


    #If TRTRACE Then
        TraceMessage "", "End Validate Distribution IDs"
    #End If
    
    #If TRTRACE Then
        TraceMessage "", "Begin Get LockStatus State"
    #End If
    l_aryEmployees = ArrayFromCollection(l_colEmployees.Contents)
    Set m_rsEmployeeData = GetLockStatusData(l_aryEmployees)
    Set m_rsCollPeriodData = GetCurrentCollPeriods()
    
    If Not ValidLockStatus(v_rsTimeEntriesRequest, "T", l_objErrors) Or _
        Not ValidLockStatus(v_rsOtherPay, "O", l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries3 = False
        GoTo xit
    End If

    
    #If TRTRACE Then
        TraceMessage "", "End Get LockStatus State"
    #End If

    '//for EMPRV move 8 char projects into feeder ref
    UpdateEMPRVFeederRef v_rsTimeEntriesRequest, v_rsJobTaskSteps

    For Each l_strEmpNum In l_colEmployees
        #If TRTRACE Then
            TraceMessage "", "Begin RecordTimeEntry for: " & l_strEmpNum
        #End If
        PrepareRequest CStr(l_strEmpNum), l_rsEmployees, l_rsEmpWorkgroups, l_rsTimecardEntries, l_rsOtherpayEntries, v_rsSchedules, l_rsScheduleChgPmts, v_rsTimeEntriesRequest, v_rsOtherPay
        Set l_objTimecard = CreateObject("busTimecard.cbsTimecard")
        
    #If TRACENEON Then
        TraceNeonMessage l_strEmpNum, "PROCTE", "BEFORE RecordTimeEntry"
    #End If
        If TypeName(l_strEmpNum) = "Long" Then
            l_lngEmpNum = l_strEmpNum
        Else
            l_lngEmpNum = CLng(l_strEmpNum)
        End If
        
        jimb = 1
                
        ProcessTimeEntries3 = _
            l_objTimecard.RecordTimeEntry4( _
                v_rsTimeEntriesRequest, _
                v_rsOtherPay, _
                l_rsEmployees, _
                l_rsEmpWorkgroups, _
                l_rsTimecardEntries, _
                l_rsOtherpayEntries, _
                v_rsSchedules, _
                l_rsScheduleChgPmts, _
                v_startDate, _
                v_endDate, _
                l_rsAccrualBankUpdates, _
                l_rsOvertimeBankUpdates, _
                l_colEMPRVInterface, _
                l_rsEMPRVJobTaskEntries, _
                l_rsMAXIMOJobTaskEntries, _
                l_objErrors, _
                v_UserId)
        If Not ProcessTimeEntries3 Then
            l_objCtx.SetAbort
            GoTo xit
        End If
    Next
    
    #If TRACENEON Then
        TraceNeonMessage CStr(lngTraceEmpNum), "PROCTE", "AFTER RecordTimeEntry, ProcessTimeEntries returned: " & CStr(ProcessTimeEntries)
    #End If
    
    #If TRTRACE Then
        TraceMessage "", "Begin UpdateAccrualBanks"
    #End If
    
    #If TRACENEON Then
        TraceNeonMessage CStr(lngTraceEmpNum), "PROCTE", "BEFORE UpdateAccrualBanks"
    #End If
    
    UpdateAccrualBanks l_rsAccrualBankUpdates, Nothing, v_UserId
    #If TRTRACE Then
        TraceMessage "", "End UpdateAccrualBanks"
    #End If
    
    
    #If TRTRACE Then
        TraceMessage "", "Begin UpdateOvertimeBanks"
    #End If
    
    #If TRACENEON Then
        TraceNeonMessage CStr(lngTraceEmpNum), "PROCTE", "BEFORE UpdateOvertimeBanks"
    #End If
    
    UpdateOvertimeBanks l_rsOvertimeBankUpdates, l_rsEmployees, Nothing, v_UserId
    #If TRTRACE Then
        TraceMessage "", "End UpdateOvertimeBanks"
    #End If
    
       
    #If TRTRACE Then
        TraceMessage "", "Begin EMPRV Interface"
    #End If
    
    If Not l_colEMPRVInterface Is Nothing Then
    If l_colEMPRVInterface.Count > 0 Then
        #If TRACENEON Then
            TraceNeonMessage CStr(lngTraceEmpNum), "PROCTE", "BEFORE Intiating EMPVR Interface"
        #End If
        SaveEMPRV l_colEMPRVInterface
    End If
    End If
    #If TRTRACE Then
        TraceMessage "", "End EMPRV Interface"
    #End If
    
    'Call to SaveJobTaskEMPRV Interface here
    #If TRTRACE Then
        TraceMessage "", "Begin EMPRV JobTasks Interface"
    #End If
    
    If Not l_rsEMPRVJobTaskEntries Is Nothing Then
        If l_rsEMPRVJobTaskEntries.RecordCount > 0 Then
            If Not SaveEMPRVJobTasks(l_rsEMPRVJobTaskEntries, l_objErrors) Then
                l_objCtx.SetAbort
                ProcessTimeEntries3 = False
                GoTo xit
            End If
        End If
    End If
    
    #If TRTRACE Then
        TraceMessage "", "End EMPRV JobTasks Interface"
    #End If
        
        
    'Call to SaveMAXIMOTaskEMPRV Interface here
    #If TRTRACE Then
        TraceMessage "", "Begin MAXIMO JobTasks Interface"
    #End If
    
    If Not l_rsMAXIMOJobTaskEntries Is Nothing Then
        If l_rsMAXIMOJobTaskEntries.RecordCount > 0 Then
            If Not SaveMAXIMOJobTasks(l_rsMAXIMOJobTaskEntries, l_objErrors) Then
                l_objCtx.SetAbort
                ProcessTimeEntries3 = False
                GoTo xit
            End If
        End If
    End If
    
    #If TRTRACE Then
        TraceMessage "", "End MAXIMO JobTasks Interface"
    #End If
                        
    #If TRACENEON Then
        TraceNeonMessage CStr(lngTraceEmpNum), "PROCTE", "BEFORE SETCOMPLETE"
    #End If
       
    #If TRTRACE Then
        TraceMessage "", "Begin SetComplete"
    #End If
    l_objCtx.SetComplete
    
    #If TRACENEON Then
        TraceNeonMessage CStr(lngTraceEmpNum), "PROCTE", "AFTER SETCOMPLETE"
    #End If
    
    #If TRTRACE Then
        TraceMessage "", "Begin SetComplete"
    #End If
   
xit:
    If Not l_objErrors Is Nothing Then
        If l_objErrors.RecordCount = 0 Then
            Set r_rsErrors = Nothing
        Else
            Set r_rsErrors = l_objErrors
        End If
    Else
        Set r_rsErrors = Nothing
    End If
    
    Set l_objTimecard = Nothing
    
    #If TRTRACE Then
        TraceMessage "", "End ProcessTimeEntries3"
    #End If
    
    #If TRACENEON Then
        TraceNeonMessage CStr(lngTraceEmpNum), "PROCTE", "EXIT"
    #End If
    Exit Function
       
eh:
    l_objCtx.SetAbort
    Set r_rsErrors = Nothing
    TrsRaiseError mstrCMODULENAME, "ProcessTimeEntries3", , , , , True
   
End Function

Private Sub UpdateEMPRVFeederRef(ByVal v_rsTimeEntries As ADODB.Recordset, _
                                ByVal v_rsJobTaskSteps As ADODB.Recordset) '//with 8 char project

On Error GoTo eh
    Dim l_sProjectNumber As String, l_sJobNbr As String, l_sJobTaskNbr As String, l_sDBId As String
    Dim l_strDistID As String, l_strDistType As String, l_sFilter As String
    
    If RecordsExist(v_rsTimeEntries) Then
    Else
        Exit Sub
    End If
    TimeEntriesFilterAndSort v_rsTimeEntries, False
    If v_rsTimeEntries.EOF Then
        v_rsTimeEntries.Filter = adFilterNone
        If RecordsExist(v_rsTimeEntries) Then
            v_rsTimeEntries.MoveFirst
        End If

        Exit Sub
    End If
    
    v_rsTimeEntries.MoveFirst
    Do While Not v_rsTimeEntries.EOF
        l_strDistID = GetDistributionID(v_rsTimeEntries)
        l_strDistType = GetDistributionType(v_rsTimeEntries)

        l_sProjectNumber = ""
        '//try to get 8 char project id's from client job task steps (from picklists)
        If l_strDistType = "J" Then
            If FieldExists("VALIDATION_TYPE", v_rsJobTaskSteps) Then
                l_sJobNbr = v_rsTimeEntries!JOB_NBR
                l_sJobTaskNbr = v_rsTimeEntries!JOBTASK_NBR
                l_sDBId = v_rsTimeEntries!EMPRV_DATABASE_ID
                
                l_sFilter = "JOB_NBR = '" & Trim(l_sJobNbr) & "'" & _
                            " AND JOBTASK_NBR = '" & Trim(l_sJobTaskNbr) & "'" & _
                            " AND EMPRV_DATABASE_ID = '" & Trim(l_sDBId) & "'" & _
                            " AND VALIDATION_TYPE = 'P'"
                v_rsJobTaskSteps.Filter = l_sFilter
                If Not v_rsJobTaskSteps.BOF And Not v_rsJobTaskSteps.EOF Then '//RecordsExist in timekeeper removes filter!!!
                    v_rsJobTaskSteps.MoveFirst
                    l_sProjectNumber = v_rsJobTaskSteps!PROJECT_NUMBER
                    
                    If Len(Trim(l_sProjectNumber)) > 5 Then
                        If Trim(GetFeederRef39(l_strDistID)) = "" Then
                            SetFeederRef39 l_strDistID, PadRight(l_sProjectNumber, 10)
                            v_rsTimeEntries!DISTRIBUTION_ID = l_strDistID
                            v_rsTimeEntries.Update
                        End If
                    End If
                End If
            End If
        End If
               
        v_rsTimeEntries.MoveNext
    Loop
    
    v_rsTimeEntries.Filter = adFilterNone
    If RecordsExist(v_rsTimeEntries) Then
        v_rsTimeEntries.MoveFirst
    End If
    
    v_rsJobTaskSteps.Filter = adFilterNone
    If RecordsExist(v_rsJobTaskSteps) Then
        v_rsJobTaskSteps.MoveFirst
    End If
    Exit Sub
eh:
    TrsRaiseError mstrCMODULENAME, "UpdateEMPRVFeederRef"
End Sub

'//copied from busDistributionMgr
Private Sub TimeEntriesFilterAndSort( _
            ByRef r_rsTimeEntries As ADODB.Recordset, _
            Optional v_blnOtherPay As Boolean = False)
                       
On Error GoTo Sort_EH:

                       
    If v_blnOtherPay Then
        r_rsTimeEntries.Filter = MDAC25FilterFix("RECORD_STATUS <> 'D' AND EXCEPTION_CD = 'TTA'")
    Else
        r_rsTimeEntries.Filter = "RECORD_STATUS <> 'D'"
    End If
    If r_rsTimeEntries.EOF Then
        Exit Sub
    End If
    r_rsTimeEntries.Sort = "KEY_EMP_NUM, KEY_CHARGE_DATE, KEY_START_TIME, END_TIME"
             
    Exit Sub
    
Sort_EH:
    TrsRaiseError mstrCMODULENAME, "TimeEntriesFilterAndSort"
            
End Sub

Private Function GetDistributionID( _
            ByVal v_rsTimeEntries As ADODB.Recordset _
            ) As String
                   
    Dim l_strDistID As String
    
    On Error GoTo GetDistributionID_EH
    
    If IsEmpty(v_rsTimeEntries!DISTRIBUTION_ID) Or IsNull(v_rsTimeEntries!DISTRIBUTION_ID) Then
        l_strDistID = ""
    Else
        l_strDistID = v_rsTimeEntries!DISTRIBUTION_ID
    End If
    
    GetDistributionID = l_strDistID
    Exit Function
    
GetDistributionID_EH:
    TrsRaiseError mstrCMODULENAME, "GetDistributionID"
    
End Function

Private Function GetDistributionType( _
            ByVal v_rsTimeEntries As ADODB.Recordset _
            ) As String
            
    Dim l_strDistID As String
    Dim l_strDistType As String
    
    On Error GoTo GetDistributionType_EH
    
    l_strDistID = GetDistributionID(v_rsTimeEntries)
    
    If IsEmpty(v_rsTimeEntries!DISTRIBUTION_TYPE) Or IsNull(v_rsTimeEntries!DISTRIBUTION_TYPE) Then
        If Trim(v_rsTimeEntries!DISTRIBUTION_TYPE) <> "J" Then
            l_strDistType = DetermineDistributionType(l_strDistID)
            v_rsTimeEntries!DISTRIBUTION_TYPE = l_strDistType
        End If
    Else
        If Trim(v_rsTimeEntries!DISTRIBUTION_TYPE = "") Then
            l_strDistType = DetermineDistributionType(l_strDistID)
            v_rsTimeEntries!DISTRIBUTION_TYPE = l_strDistType
        Else
            l_strDistType = v_rsTimeEntries!DISTRIBUTION_TYPE
        End If
    End If
    
    GetDistributionType = l_strDistType
    
    Exit Function
    
GetDistributionType_EH:
    TrsRaiseError mstrCMODULENAME, "GetDistributionType"
      
        
End Function

Private Function DetermineDistributionType( _
            ByVal v_strDistributionId As Variant _
            ) As String
            
On Error GoTo eh
            
    If Not IsArray(v_strDistributionId) Then
        If Trim(v_strDistributionId) = "" Then
            DetermineDistributionType = " "
            Exit Function
        End If
        If Mid$(v_strDistributionId, 6, 2) = "WR" And Trim(Mid$(v_strDistributionId, 5, 1)) = "" Then
            DetermineDistributionType = "W"
            Exit Function
        End If
        If IsServiceRequest(v_strDistributionId) Then
            DetermineDistributionType = "S"
            Exit Function
        End If
        ' There is no true way to tell if it is full accounting
        ' or the description of the exception code
        ' I am assuming that all exception code descriptions
        ' have at least 4 characters
        If IsNumeric(Mid$(v_strDistributionId, 4, 3)) Or Trim(Mid$(v_strDistributionId, 4, 3)) = "" Then
            DetermineDistributionType = "A"
        Else
            DetermineDistributionType = " "
        End If
    Else
        If Trim(v_strDistributionId(1)) = "WR" Then
            DetermineDistributionType = "W"
            Exit Function
        End If
        If IsServiceRequest(v_strDistributionId) Then
            DetermineDistributionType = "S"
            Exit Function
        End If
        If IsNumeric(v_strDistributionId(2)) Or Trim(v_strDistributionId(2)) = "" Then
            DetermineDistributionType = "A"
        Else
            DetermineDistributionType = " "
        End If
    End If
    
Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "DetermineDistributionType"
            
End Function

Private Function IsServiceRequest(ByVal v_varDistributionId As Variant) As Boolean

' A Distribution ID is a service request
' if only the following 3 fields are entered
' FMC, Project, Activity
' The Feeder_ref (Memo) field is ignored
' The Project field must be entered to be a service request

    Dim l_strTempHold As Variant
    Dim l_intNumOfFields As Integer
    
On Error GoTo IsServiceRequest_EH
    
    Dim l_strDistributionId(11) As String
    Dim l_strEmptyFields As String
    
    If Not IsArray(v_varDistributionId) Then
        l_strDistributionId(0) = Left$(v_varDistributionId, 1) '(U)Utility
        l_strDistributionId(1) = Mid$(v_varDistributionId, 2, 2) '(BD) Business Division
        l_strDistributionId(2) = Mid$(v_varDistributionId, 4, 3) '(MAJ)FERC Major
        l_strDistributionId(3) = Mid$(v_varDistributionId, 7, 3) '(MIN)FERC Minor
        l_strDistributionId(4) = Mid$(v_varDistributionId, 10, 3) '(FMC) Fund Mgt. Center
        l_strDistributionId(5) = Mid$(v_varDistributionId, 13, 3) '(RMC) Resource Mgt. Center
        l_strDistributionId(6) = Mid$(v_varDistributionId, 16, 1) '(T) Trans Type
        l_strDistributionId(7) = Mid$(v_varDistributionId, 17, 5) '(PROJT) Project
        l_strDistributionId(8) = Mid$(v_varDistributionId, 22, 2) '(PR) Product
        l_strDistributionId(9) = Mid$(v_varDistributionId, 24, 4) '(ACTV) Activity
        l_strDistributionId(10) = Mid$(v_varDistributionId, 28, 2) '(RT) Resource Type
        l_strDistributionId(11) = Mid$(v_varDistributionId, 30, 10) '(FEEDER_REF) Memo
    Else
        l_strDistributionId(0) = v_varDistributionId(0) '(U)Utility
        l_strDistributionId(1) = v_varDistributionId(1) '(BD)  Business Division
        l_strDistributionId(2) = v_varDistributionId(2) '(MAJ)FERC Major
        l_strDistributionId(3) = v_varDistributionId(3) '(MIN)FERC Minor
        l_strDistributionId(4) = v_varDistributionId(4) '(FMC) Fund Mgt. Center
        l_strDistributionId(5) = v_varDistributionId(5) '(RMC) Resource Mgt. Center
        l_strDistributionId(6) = v_varDistributionId(6) '(T) Trans Type
        l_strDistributionId(7) = v_varDistributionId(7) '(PROJT) Project
        l_strDistributionId(8) = v_varDistributionId(8) '(PR) Product
        l_strDistributionId(9) = v_varDistributionId(9) '(ACTV) Activity
        l_strDistributionId(10) = v_varDistributionId(10) '(RT) Resource Type
        l_strDistributionId(11) = v_varDistributionId(11) '(FEEDER_REF) Memo
    End If
    
    l_strEmptyFields = l_strDistributionId(0) & _
                       l_strDistributionId(1) & _
                       l_strDistributionId(2) & _
                       l_strDistributionId(3) & _
                       l_strDistributionId(5) & _
                       l_strDistributionId(6) & _
                       l_strDistributionId(8) & _
                       l_strDistributionId(10)
    If Trim(l_strEmptyFields) <> "" Then
        IsServiceRequest = False
        Exit Function
    End If
    If Trim(l_strDistributionId(7)) <> "" Then
'Please not that FMC and Activity may be missing
        IsServiceRequest = True
    Else
        IsServiceRequest = False
    End If
    Exit Function
IsServiceRequest_EH:
    TrsRaiseError mstrCMODULENAME, "IsServiceRequest", , , , , True
    
End Function

'//end copy from busDistributionMgr

Private Sub UpdateAccrualBanks( _
                ByRef r_rsAccrualBankUpdates As ADODB.Recordset, _
                ByRef r_rsErrors As ADODB.Recordset, _
                ByVal v_strEmployeeNumber As String)
    Dim lobjAccrualMgr As busAccrualBankMgr.cbsAccrualBankMgr

On Error GoTo eh
    If RecordsExist(r_rsAccrualBankUpdates) Then
        Set lobjAccrualMgr = _
            CreateObject("busAccrualBankMgr.cbsAccrualBankMgr")
        lobjAccrualMgr.UpdateBanksForTimeEntry r_rsAccrualBankUpdates, r_rsErrors, v_strEmployeeNumber
    End If
    Exit Sub

eh:
    Set lobjAccrualMgr = Nothing
    TrsRaiseError mstrCMODULENAME, "UpdateAccrualBanks"

End Sub

Private Sub UpdateOvertimeBanks( _
                ByRef r_rsOvertimeBankUpdates As ADODB.Recordset, _
                ByRef r_rsEmployees As ADODB.Recordset, _
                ByRef r_rsErrors As ADODB.Recordset, _
                ByVal v_strEmployeeNumber As String)
    Dim lobjOTMgr As busOvertimeEqualizer.cbsOTEqualizer

On Error GoTo eh
    If RecordsExist(r_rsOvertimeBankUpdates) Then
        Set lobjOTMgr = _
            CreateObject("busOvertimeEqualizer.cbsOTEqualizer")
        lobjOTMgr.UpdateOvertimeBanksForTimeEntry r_rsOvertimeBankUpdates, r_rsEmployees, r_rsErrors, v_strEmployeeNumber
    End If
    Exit Sub

eh:
    Set lobjOTMgr = Nothing
    TrsRaiseError mstrCMODULENAME, "UpdateOvertimeBanks"

End Sub

Private Sub PrepareRequest( _
                l_strEmpNum As String, _
                l_rsEmployees As ADODB.Recordset, _
                l_rsEmpWorkgroups As ADODB.Recordset, _
                l_rsTimecardEntries As ADODB.Recordset, _
                l_rsOtherpayEntries As ADODB.Recordset, _
                v_rsSchedules As ADODB.Recordset, _
                v_rsScheduleChgPmts As ADODB.Recordset, _
                v_rsTimeEntriesRequest As ADODB.Recordset, _
                v_rsOtherPay As ADODB.Recordset)

On Error GoTo eh
    If Not l_rsEmployees Is Nothing Then
        l_rsEmployees.Filter = adFilterNone
    End If
    If RecordsExist(l_rsEmployees) Then
        l_rsEmployees.Filter = "EMP_NUM = " & l_strEmpNum
    End If
    
    If Not l_rsEmpWorkgroups Is Nothing Then
        l_rsEmpWorkgroups.Filter = adFilterNone
    End If
    If RecordsExist(l_rsEmpWorkgroups) Then
        l_rsEmpWorkgroups.Filter = "EMP_NUM = " & l_strEmpNum
    End If
    
    If Not l_rsTimecardEntries Is Nothing Then
        l_rsTimecardEntries.Filter = adFilterNone
    End If
    If RecordsExist(l_rsTimecardEntries) Then
        l_rsTimecardEntries.Filter = "EMP_NUM = " & l_strEmpNum
        l_rsTimecardEntries.Sort = "CHARGE_DATE ASC"
    End If
    
    If Not l_rsOtherpayEntries Is Nothing Then
        l_rsOtherpayEntries.Filter = adFilterNone
    End If
    If RecordsExist(l_rsOtherpayEntries) Then
        l_rsOtherpayEntries.Filter = "EMP_NUM = " & l_strEmpNum
        l_rsOtherpayEntries.Sort = "CHARGE_DATE ASC"
    End If
    
    If Not v_rsSchedules Is Nothing Then
        v_rsSchedules.Filter = adFilterNone
    End If
    If RecordsExist(v_rsSchedules) Then
        v_rsSchedules.Filter = "EMP_NUM = " & l_strEmpNum
        v_rsSchedules.Sort = "CHARGE_DATE ASC"
    End If
    
    If Not v_rsScheduleChgPmts Is Nothing Then
        v_rsScheduleChgPmts.Filter = adFilterNone
    End If
    If RecordsExist(v_rsScheduleChgPmts) Then
        v_rsScheduleChgPmts.Filter = "EMP_NUM = " & l_strEmpNum
        v_rsScheduleChgPmts.Sort = "CHARGE_DATE ASC"
    End If
    
    If Not v_rsTimeEntriesRequest Is Nothing Then
        v_rsTimeEntriesRequest.Filter = adFilterNone
    End If
    If RecordsExist(v_rsTimeEntriesRequest) Then
        v_rsTimeEntriesRequest.Filter = "KEY_EMP_NUM = " & l_strEmpNum
        v_rsTimeEntriesRequest.Sort = "KEY_CHARGE_DATE ASC, EXCEPTION_CD DESC"
    End If
    
    If Not v_rsOtherPay Is Nothing Then
        v_rsOtherPay.Filter = adFilterNone
    End If
    If RecordsExist(v_rsOtherPay) Then
        v_rsOtherPay.Filter = "KEY_EMP_NUM = " & l_strEmpNum
        v_rsOtherPay.Sort = "KEY_CHARGE_DATE ASC"
    End If
    
'    If Not v_rsECDrawTime Is Nothing Then
'        v_rsECDrawTime.Filter = adFilterNone
'    End If
'    If RecordsExist(v_rsECDrawTime) Then
'        v_rsECDrawTime.Filter = "KEY_EMP_NUM = " & l_strEmpNum
'        v_rsECDrawTime.Sort = "WEEK_ENDING_DATE ASC"
'    End If
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "PrepareRequest"

End Sub

Private Function GetEmployeeList( _
            r_rsTimeEntriesRequest As ADODB.Recordset, _
            r_rsOtherPay As ADODB.Recordset) _
                As cuSet
    Dim l_objEmpList As cuSet

On Error GoTo eh
    Set l_objEmpList = New cuSet
    If RecordsExist(r_rsTimeEntriesRequest) Then
        r_rsTimeEntriesRequest.MoveFirst
        Do Until r_rsTimeEntriesRequest.EOF
            If Not IsNull(r_rsTimeEntriesRequest!KEY_EMP_NUM) Then
                l_objEmpList.Add r_rsTimeEntriesRequest.Fields("KEY_EMP_NUM").Value
            End If
            r_rsTimeEntriesRequest.MoveNext
        Loop
        r_rsTimeEntriesRequest.MoveFirst
    End If
    
    If RecordsExist(r_rsOtherPay) Then
        r_rsOtherPay.MoveFirst
        Do Until r_rsOtherPay.EOF
            If Not IsNull(r_rsOtherPay!KEY_EMP_NUM) Then
                l_objEmpList.Add r_rsOtherPay.Fields("KEY_EMP_NUM").Value
            End If
            r_rsOtherPay.MoveNext
        Loop
        r_rsOtherPay.MoveFirst
    End If
    
'    If RecordsExist(v_rsECDrawTime) Then
'        v_rsECDrawTime.MoveFirst
'        Do Until v_rsECDrawTime.EOF
'            If Not IsNull(v_rsECDrawTime!KEY_EMP_NUM) Then
'                l_objEmpList.Add v_rsECDrawTime.Fields("KEY_EMP_NUM").Value
'            End If
'            v_rsECDrawTime.MoveNext
'        Loop
'        v_rsECDrawTime.MoveFirst
'    End If
    
    Set GetEmployeeList = l_objEmpList
    Exit Function
    
eh:
    Set l_objEmpList = Nothing
    TrsRaiseError mstrCMODULENAME, "GetEmployeeList"

End Function

Private Function GetEmployeeState( _
        ByRef l_colEmployees As cuSet, _
        ByVal startDt As Date, _
        ByVal endDt As Date, _
        ByVal v_blnOtherPayReq As Boolean, _
        ByRef l_rsEmployees As ADODB.Recordset, _
        ByRef l_rsEmpWorkgroups As ADODB.Recordset, _
        ByRef l_rsTimecardEntries As ADODB.Recordset, _
        ByRef l_rsOtherpayEntries As ADODB.Recordset, _
        ByRef l_rsSchedules As ADODB.Recordset, _
        ByRef l_rsScheduleChangePayments As ADODB.Recordset)
    Dim l_aryEmployees As Variant

On Error GoTo eh

    l_aryEmployees = ArrayFromCollection(l_colEmployees.Contents)
    
    If l_rsEmployees Is Nothing Then
        Set l_rsEmployees = _
            GetEmployees(l_aryEmployees, startDt, endDt)
    End If
    If l_rsEmpWorkgroups Is Nothing Then
        Set l_rsEmpWorkgroups = _
            GetEmpWorkgroups(l_aryEmployees, startDt, endDt)
    End If
    Set l_rsTimecardEntries = _
        GetTimecardEntries(l_aryEmployees, startDt, endDt)
    If CheckForScheduleChangePayments(l_rsSchedules) Then
        Set l_rsScheduleChangePayments = _
            GetScheduleChangePayments(l_aryEmployees, startDt, endDt)
    End If
    If v_blnOtherPayReq Then
        Set l_rsOtherpayEntries = _
            GetOtherPayEntries(l_aryEmployees, startDt, endDt)
    End If

    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "GetEmployeeState"

End Function

Private Function CheckForScheduleChangePayments( _
            l_rsSchedules As ADODB.Recordset) As Boolean
' If schedules are not provided, must check for schedule change
' payments.  If schedules are provided, check for schedule change
' payements if any schedule indicates schedule change payment.

On Error GoTo eh

    If Not RecordsExist(l_rsSchedules) Then
        CheckForScheduleChangePayments = True
        Exit Function
    End If

    l_rsSchedules.Filter = "SCHED_CHANGE_PMT = 'Y'"
    If l_rsSchedules.EOF Then
        CheckForScheduleChangePayments = False
    Else
        CheckForScheduleChangePayments = True
    End If
    l_rsSchedules.Filter = ""
    l_rsSchedules.MoveFirst
    
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "CheckForScheduleChangePayments"

End Function

Private Function GetTimecardEntries( _
            ByVal v_aryEmployees As Variant, _
            ByVal startDt As Date, _
            ByVal endDt As Date) As ADODB.Recordset
    Dim ocdTimeCard As dbTimeCard.cdTimeCard

On Error GoTo Restore_eh
    
    Set ocdTimeCard = _
        CreateObject("dbTimeCard.cdTimeCard")
    Set GetTimecardEntries = _
        ocdTimeCard.Get_TimeCard(v_aryEmployees, startDt, endDt)
    Set ocdTimeCard = Nothing
    Exit Function
    
Restore_eh:
    Set ocdTimeCard = Nothing
    TrsRaiseError mstrCMODULENAME, "GetTimecardEntries"
    
End Function


Private Function GetOtherPayEntries( _
            ByVal v_aryEmployees As Variant, _
            ByVal startDt As Date, _
            ByVal endDt As Date) As ADODB.Recordset
    Dim ocdTimeCard As dbTimeCard.cdTimeCard

On Error GoTo Restore_eh
    
    Set ocdTimeCard = _
        CreateObject("dbTimeCard.cdTimeCard")
    Set GetOtherPayEntries = _
        ocdTimeCard.Get_OtherPay(v_aryEmployees, startDt, endDt)
    Set ocdTimeCard = Nothing
    Exit Function
    
Restore_eh:
    Set ocdTimeCard = Nothing
    TrsRaiseError mstrCMODULENAME, "GetOtherPayEntries"
End Function


Private Function GetScheduleChangePayments( _
            ByVal v_aryEmployees As Variant, _
            ByVal startDt As Date, _
            ByVal endDt As Date) As ADODB.Recordset
    Dim ocdSchedule As dbSchedule.cdSchedule
    
On Error GoTo eh
    
    Set ocdSchedule = _
        CreateObject("dbSchedule.cdSchedule")
    Set GetScheduleChangePayments = _
        ocdSchedule.GetScheduleChangePayments(v_aryEmployees, startDt, endDt)
    Set ocdSchedule = Nothing
    Exit Function
    
eh:
    Set ocdSchedule = Nothing
    TrsRaiseError mstrCMODULENAME, "GetScheduleChangePayments"
    
End Function

Private Function GetEmployees( _
            ByVal v_aryEmployees As Variant, _
            ByVal startDt As Date, _
            ByVal endDt As Date) As ADODB.Recordset
    Dim ocdEmployee As dbEmployee.cdEmployee

On Error GoTo Restore_eh
    
    Set ocdEmployee = _
        CreateObject("dbEmployee.cdEmployee")
    Set GetEmployees = _
        ocdEmployee.GetEmployee(v_aryEmployees)
'        ocdEmployee.GetEmployee2(v_aryEmployees, startDt, endDt)
    Set ocdEmployee = Nothing
    Exit Function
    
Restore_eh:
    Set ocdEmployee = Nothing
    TrsRaiseError mstrCMODULENAME, "GetEmployees"
    
End Function

Private Function GetEmpWorkgroups( _
            ByVal v_aryEmployees As Variant, _
            ByVal startDt As Date, _
            ByVal endDt As Date) As ADODB.Recordset
    Dim ocdWorkGroup As dbWorkGroup.cdWorkGroup

On Error GoTo Restore_eh
    
    Set ocdWorkGroup = _
        CreateObject("dbWorkgroup.cdWorkgroup")
    Set GetEmpWorkgroups = _
        ocdWorkGroup.GetEmployeeWorkGroups(v_aryEmployees, startDt, endDt, True)
    Set ocdWorkGroup = Nothing
    Exit Function
    
Restore_eh:
    Set ocdWorkGroup = Nothing
    TrsRaiseError mstrCMODULENAME, "GetEmpWorkgroups"
    
End Function

Private Function RecordsExist( _
                r_rsAnyRecordset As ADODB.Recordset) As Boolean

On Error GoTo eh

    If Not r_rsAnyRecordset Is Nothing Then
        r_rsAnyRecordset.Filter = ""
        If Not (r_rsAnyRecordset.BOF And r_rsAnyRecordset.EOF) Then
            RecordsExist = True
        End If
    End If
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "RecordsExist"

End Function

Private Function CheckForInput( _
        r_rsTimeEntriesRequest As ADODB.Recordset, _
        r_rsOtherpayRequest As ADODB.Recordset, _
        r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_blnEmptyTimeRequest As Boolean
    Dim l_blnEmptyOtherRequest As Boolean
'    Dim l_blnEmptyECRequest As Boolean
    
On Error GoTo eh
    If Not r_rsTimeEntriesRequest Is Nothing Then
        If r_rsTimeEntriesRequest.BOF And _
            r_rsTimeEntriesRequest.EOF Then
            l_blnEmptyTimeRequest = True
        End If
    Else
        l_blnEmptyTimeRequest = True
    End If

    If Not r_rsOtherpayRequest Is Nothing Then
        If r_rsOtherpayRequest.BOF And _
            r_rsOtherpayRequest.EOF Then
            l_blnEmptyOtherRequest = True
        End If
    Else
        l_blnEmptyOtherRequest = True
    End If
    
'    If Not r_rsECDrawTime Is Nothing Then
'        If r_rsECDrawTime.BOF And _
'            r_rsECDrawTime.EOF Then
'            l_blnEmptyECRequest = True
'        End If
'    Else
'        l_blnEmptyECRequest = True
'    End If
    
'    If l_blnEmptyTimeRequest And l_blnEmptyOtherRequest _
'            And l_blnEmptyECRequest Then
    If l_blnEmptyTimeRequest And l_blnEmptyOtherRequest Then
        AppendError _
            r_rsErrors, 0, Date, Date, "00:00", "00:00", 0, "", "", _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "Time entry and other pay requests are empty", mstrCMODULENAME
        Exit Function
    End If
    CheckForInput = True
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "CheckForInput"
    
End Function

Private Function ValidateInput( _
        ByRef r_rsTimeEntriesRequest As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_blnValidationError As Boolean

On Error GoTo eh
    If r_rsTimeEntriesRequest Is Nothing Then
        ValidateInput = True
        Exit Function
    End If

    If r_rsTimeEntriesRequest.BOF And _
        r_rsTimeEntriesRequest.EOF Then
'        AppendError _
'            r_rsErrors, 0, Date, Date, "00:00", "00:00", 0, "", "", _
'            trErrbusValidationError, mstrCBUSINESSERR, _
'            "Timekeeper request is empty", mstrCMODULENAME
        ValidateInput = True
        Exit Function
    End If

    r_rsTimeEntriesRequest.MoveFirst
    
'    temp while not partitioning uow
'    Dim l_lngFirstEmpNum As Long
'    l_lngFirstEmpNum = r_rsTimeEntriesRequest!EMP_NUM
    
    Do Until r_rsTimeEntriesRequest.EOF
        Select Case r_rsTimeEntriesRequest!RECORD_STATUS
            Case "I"
                If Not ValidateNewTimeEntry(r_rsTimeEntriesRequest, r_rsErrors) Then
                    l_blnValidationError = True
                End If
            Case "C"
                If Not ValidateChangeTimeEntry(r_rsTimeEntriesRequest, r_rsErrors) Then
                    l_blnValidationError = True
                End If
            Case "D"
                If Not ValidateDeleteTimeEntry(r_rsTimeEntriesRequest, r_rsErrors) Then
                    l_blnValidationError = True
                End If
            Case Else
                AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Record status must be I, C, or D"
                l_blnValidationError = True
        End Select
        r_rsTimeEntriesRequest.MoveNext
    Loop
    
    ValidateInput = Not l_blnValidationError
    r_rsTimeEntriesRequest.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ValidateInput"

End Function

Private Function GetMaxChargeDate() As Date
    Dim l_obj As busEnvironment.cbsEnvironment
On Error GoTo eh

    Set l_obj = CreateObject("busEnvironment.cbsEnvironment")

    GetMaxChargeDate = l_obj.GetMaxChargeDate()
    Set l_obj = Nothing
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError mstrCMODULENAME, "GetMaxChargeDate"
End Function

Private Function GetMaxChargeDateOffset() As Long
    Dim l_obj As busEnvironment.cbsEnvironment
On Error GoTo eh

    Set l_obj = CreateObject("busEnvironment.cbsEnvironment")

    GetMaxChargeDateOffset = l_obj.GetMaxChargeDateOffset()
    Set l_obj = Nothing
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError mstrCMODULENAME, "GetMaxChargeDateOffset"
End Function

Private Function GetMaxChargeDateError() As String
    Dim NumberOfDays As Long
On Error GoTo eh
    
    NumberOfDays = GetMaxChargeDateOffset
    If NumberOfDays = 1 Then
        GetMaxChargeDateError = "Charge Date is more than 1 day in future"
    Else
        GetMaxChargeDateError = "Charge Date is more than " & CStr(NumberOfDays) & " days in future (" & CStr(GetMaxChargeDate()) & ")"
    End If
    
    Exit Function
eh:
    TrsRaiseError mstrCMODULENAME, "GetMaxChargeDateError"
End Function

Private Function ValidateNewTimeEntry( _
        ByRef r_rsTimeEntriesRequest As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset) As Boolean
' answer true when the request is valid
    Dim l_strExcCd As String
    Dim l_strOverExcCd As String
    Dim l_blnValidationError As Boolean
'    Dim l_intDateDiff As Integer
    Dim l_lngDateDiff As Long
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_decHours As Variant

On Error GoTo eh

'    temp while not partitioning uow, only allow 1 employee per uow
' 7/5/00 commented out for now until uow decided.
'    If Not l_lngFirstEmpNum = r_rsTimeEntriesRequest!EMP_NUM Then
'        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Multiple employees not allowed"
'        l_blnValidationError = True
'    End If
    
    If Not IsNumeric(r_rsTimeEntriesRequest!EMP_NUM) Or _
            r_rsTimeEntriesRequest!EMP_NUM = 0 Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Employee Id is invalid"
        l_blnValidationError = True
    End If

' 9/13/2001 RSG
' Reject uninitialized dates and dates set to vbNull
    If Not IsDateValid(r_rsTimeEntriesRequest!CHARGE_DATE.Value) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Charge Date is not a valid date. New time entry"
        l_blnValidationError = True
    Else
        If r_rsTimeEntriesRequest!CHARGE_DATE > GetMaxChargeDate() Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, GetMaxChargeDateError()
            l_blnValidationError = True
        End If
    End If

' 9/13/2001 RSG
' Reject uninitialized dates and dates set to vbNull
    If Not IsDateValid(r_rsTimeEntriesRequest!DATE_WORKED.Value) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Date worked is not a valid date"
        l_blnValidationError = True
    End If

' 9/13/2001 RSG
' Charge_Date may only be upto 1 day before or after Date_Worked.
    If IsDateValid(r_rsTimeEntriesRequest!CHARGE_DATE.Value) And _
        IsDateValid(r_rsTimeEntriesRequest!DATE_WORKED.Value) Then
        l_lngDateDiff = Abs(DateDiff("d", r_rsTimeEntriesRequest!CHARGE_DATE.Value, r_rsTimeEntriesRequest!DATE_WORKED.Value))
        If l_lngDateDiff > 1 Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Charge date/Date worked difference more than 1 day"
            l_blnValidationError = True
        End If
    End If

    If Not IsTimeValid(r_rsTimeEntriesRequest!START_TIME) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start time is not valid"
        l_blnValidationError = True
    End If

    If Not IsTimeValid(r_rsTimeEntriesRequest!END_TIME) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "End time is not valid"
        l_blnValidationError = True
    End If

    If Not IsNumeric(r_rsTimeEntriesRequest!Hours) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not numeric"
        l_blnValidationError = True
    Else
        If Not (r_rsTimeEntriesRequest!Hours >= 0 And _
                r_rsTimeEntriesRequest!Hours <= 24) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not a valid range"
            l_blnValidationError = True
        End If
    End If
    
    If Not l_blnValidationError Then
        l_strStartTime = Trim(r_rsTimeEntriesRequest!START_TIME)
        l_strEndTime = Trim(r_rsTimeEntriesRequest!END_TIME)
        If l_strStartTime = l_strEndTime Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start time can not equal end time"
            l_blnValidationError = True
        End If
        If l_strStartTime > l_strEndTime Then
            If l_strEndTime = "00:00" Then
                r_rsTimeEntriesRequest!END_TIME = "24:00"
            Else
                AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start time is greater than end time. Time may not be split properly over midnight."
                l_blnValidationError = True
            End If
        End If
    End If
    If Not l_blnValidationError Then
        l_strStartTime = Trim(r_rsTimeEntriesRequest!START_TIME)
        l_strEndTime = Trim(r_rsTimeEntriesRequest!END_TIME)
        l_decHours = CDec(Left(l_strEndTime, 2)) + CDec(Right(l_strEndTime, 2)) / 60 _
                    - CDec(Left(l_strStartTime, 2)) - CDec(Right(l_strStartTime, 2)) / 60
        r_rsTimeEntriesRequest!Hours = l_decHours
    End If

    ' Schedule Indicator
    If Not IsYorN(r_rsTimeEntriesRequest!SCHED_TIME_IND) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Schedule Indicator must be Y or N"
        l_blnValidationError = True
    End If
    
    If Not ValidateOptionalFields(r_rsTimeEntriesRequest, r_rsErrors) Then
        l_blnValidationError = True
    End If
    
    If Not r_rsTimeEntriesRequest!KEY_EMP_NUM = _
            r_rsTimeEntriesRequest!EMP_NUM Then
        l_blnValidationError = True
    End If

    If Not r_rsTimeEntriesRequest!KEY_CHARGE_DATE = _
            r_rsTimeEntriesRequest!CHARGE_DATE Then
        l_blnValidationError = True
    End If

    If Not r_rsTimeEntriesRequest!KEY_START_TIME = _
            r_rsTimeEntriesRequest!START_TIME Then
        l_blnValidationError = True
    End If
    
    ValidateNewTimeEntry = Not l_blnValidationError
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ValidateNewTimeEntry"

End Function

Private Function ValidateChangeTimeEntry( _
        ByRef r_rsTimeEntriesRequest As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset) As Boolean
' answer true when the request is valid
    Dim l_blnValidationError As Boolean
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_decHours As Variant

On Error GoTo eh
' Ensure key components are provided, if not stop here
    If Not ValidateKeyComponents(r_rsTimeEntriesRequest, r_rsErrors) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCAPPLICATIONERR, "One or more key components are invalid"
        ValidateChangeTimeEntry = False
        Exit Function
    End If

    If Not IsNull(r_rsTimeEntriesRequest!EMP_NUM) Then
        If Not r_rsTimeEntriesRequest!EMP_NUM = _
                    r_rsTimeEntriesRequest!KEY_EMP_NUM Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Employee Id is not valid"
            l_blnValidationError = True
        End If
    End If

' 9/13/2001 RSG
' Reject uninitialized dates and dates set to vbNull
    If Not IsNull(r_rsTimeEntriesRequest!CHARGE_DATE) Then
        If Not IsDateValid(r_rsTimeEntriesRequest!CHARGE_DATE.Value) Or _
                Not r_rsTimeEntriesRequest!CHARGE_DATE = _
                    r_rsTimeEntriesRequest!KEY_CHARGE_DATE Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Charge Date is not a valid date.  Change time entry."
            l_blnValidationError = True
        End If
    End If
    If IsDateValid(r_rsTimeEntriesRequest!KEY_CHARGE_DATE.Value) Then
        If r_rsTimeEntriesRequest!KEY_CHARGE_DATE > GetMaxChargeDate() Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, GetMaxChargeDateError()
            l_blnValidationError = True
        End If
    End If

' 9/13/2001 RSG
' Reject uninitialized dates and dates set to vbNull
    If Not IsNull(r_rsTimeEntriesRequest!DATE_WORKED) Then
        If Not IsDateValid(r_rsTimeEntriesRequest!DATE_WORKED.Value) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Date worked is not a valid date"
            l_blnValidationError = True
        End If
    End If

    If Not IsNull(r_rsTimeEntriesRequest!START_TIME) Then
        If Not IsTimeValid(r_rsTimeEntriesRequest!START_TIME) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start time is not valid"
            l_blnValidationError = True
        End If
    End If

    If Not IsNull(r_rsTimeEntriesRequest!END_TIME) Then
        If Not IsTimeValid(r_rsTimeEntriesRequest!END_TIME) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "End time is not valid"
            l_blnValidationError = True
        End If
    End If

    If Not IsNull(r_rsTimeEntriesRequest!Hours) Then
        If Not IsNumeric(r_rsTimeEntriesRequest!Hours) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not numeric"
            l_blnValidationError = True
        Else
            If Not (r_rsTimeEntriesRequest!Hours >= 0 And _
                    r_rsTimeEntriesRequest!Hours <= 24) Then
                AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not a valid range"
                l_blnValidationError = True
            End If
        End If
    End If
    
    If Not l_blnValidationError Then
        l_strStartTime = Trim(r_rsTimeEntriesRequest!START_TIME)
        l_strEndTime = Trim(r_rsTimeEntriesRequest!END_TIME)
        If l_strStartTime = l_strEndTime Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start time can not equal end time"
            l_blnValidationError = True
        End If
        If l_strStartTime > l_strEndTime Then
            If l_strEndTime = "00:00" Then
                r_rsTimeEntriesRequest!END_TIME = "24:00"
            Else
                AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start time is greater than end time. Time may not be split properly over midnight."
                l_blnValidationError = True
            End If
        End If
    End If
    If Not l_blnValidationError Then
        l_strStartTime = Trim(r_rsTimeEntriesRequest!START_TIME)
        l_strEndTime = Trim(r_rsTimeEntriesRequest!END_TIME)
        l_decHours = CDec(Left(l_strEndTime, 2)) + CDec(Right(l_strEndTime, 2)) / 60 _
                    - CDec(Left(l_strStartTime, 2)) - CDec(Right(l_strStartTime, 2)) / 60
        r_rsTimeEntriesRequest!Hours = l_decHours
    End If

    ' Schedule Indicator
    If Not IsYorN(r_rsTimeEntriesRequest!SCHED_TIME_IND) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Schedule Indicator must be Y or N"
        l_blnValidationError = True
    End If
    
    If Not ValidateOptionalFields(r_rsTimeEntriesRequest, r_rsErrors) Then
        l_blnValidationError = True
    End If
    
    ValidateChangeTimeEntry = Not l_blnValidationError
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ValidateChangeTimeEntry"

End Function

Private Function ValidateOptionalFields( _
        ByRef r_rsTimeEntriesRequest As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset) As Boolean
' answer true when the request is valid
    Dim l_strExcCd As String
    Dim l_strOverExcCd As String
    Dim l_strDistType As String
    Dim l_strDistributionId As String
    Dim l_strfeeder As String
    Dim l_len As Integer
    Dim l_blnValidationError As Boolean

On Error GoTo eh
    '  Exception Code
'    This code will be used when the application can handle exceptions
    If Not IsNull(r_rsTimeEntriesRequest!EXCEPTION_CD) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!EXCEPTION_CD) And r_rsTimeEntriesRequest!EXCEPTION_CD <> "" Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception code is not valid"
        End If
    End If
    
' temporary
' 7/18/00 removed the filter allowing all exception codes
' so that accrual bank manager could be tested.
'    If Not IsNull(r_rsTimeEntriesRequest!EXCEPTION_CD) Then
'        l_strExcCd = Trim(r_rsTimeEntriesRequest!EXCEPTION_CD)
'        If Not l_strExcCd = "" And Not l_strExcCd = "CS" Then
'            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception Code is not valid"
'            l_blnValidationError = True
'        End If
'    End If
    
    '  OverrideException Code
    If Not IsNull(r_rsTimeEntriesRequest!OVER_EXC_CD) And Trim(r_rsTimeEntriesRequest!OVER_EXC_CD) <> "" Then
       If Not IsAlphaNumeric(r_rsTimeEntriesRequest!OVER_EXC_CD) And r_rsTimeEntriesRequest!OVER_EXC_CD <> "" Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Override Exception code is not valid"
        l_blnValidationError = True
       End If
       ''jimbo
        If Not IsNull(r_rsTimeEntriesRequest!DISTRIBUTION_TYPE) Then
            l_strDistType = Trim(r_rsTimeEntriesRequest!DISTRIBUTION_TYPE)
        Else
            l_strDistType = ""
        End If
        If Not IsNull(r_rsTimeEntriesRequest!DISTRIBUTION_ID) Then
            l_strDistributionId = r_rsTimeEntriesRequest!DISTRIBUTION_ID
        Else
            l_strDistributionId = ""
        End If
        If l_strDistType = "M" Or l_strDistType = "J" Then
            l_len = Len(l_strDistributionId)
            l_strfeeder = Trim(Mid(l_strDistributionId, 30, 10))
            l_strOverExcCd = Trim(r_rsTimeEntriesRequest!OVER_EXC_CD)
            If l_strfeeder = l_strOverExcCd Then
                r_rsTimeEntriesRequest!DISTRIBUTION_TYPE = ""
                r_rsTimeEntriesRequest!JOB_NBR = ""
                r_rsTimeEntriesRequest!JOBTASK_NBR = ""
                r_rsTimeEntriesRequest!JOBTASK_STEP_NBR = 0
                r_rsTimeEntriesRequest!EMPRV_DATABASE_ID = ""
            End If
        End If
    End If
       
    '  Work Group
    If Not IsNull(r_rsTimeEntriesRequest!WG_NUM) Then
        If Not IsNumeric(r_rsTimeEntriesRequest!WG_NUM) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Workgroup is not numeric"
            l_blnValidationError = True
        End If
    End If
        
    '  Distribution Department
    If Not IsNull(r_rsTimeEntriesRequest!DISTR_DEPT_NUM) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!DISTR_DEPT_NUM) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Department is invalid or missing"
            l_blnValidationError = True
        End If
    End If
       
    '  Job Number
    If Not IsNull(r_rsTimeEntriesRequest!JOB_NUM) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!JOB_NUM) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Job Number is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    '  OT Reason Code
    If Not IsNull(r_rsTimeEntriesRequest!OT_REASON_CD) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!OT_REASON_CD) And Trim(r_rsTimeEntriesRequest!OT_REASON_CD) <> "" Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "OT Reason Code is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    '  Pre-Arranged OT Indicator
    If Not PreArrangedOTIndMissing Then
        If Not IsNull(r_rsTimeEntriesRequest!PREARRANGED_OT_IND) Then
            If Not IsYorN(r_rsTimeEntriesRequest!PREARRANGED_OT_IND) Then
            'TODO: PreArrangedOTIndicator should not be empty string
                AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Pre-Arranged OT indicator is not valid."
                l_blnValidationError = True
            End If
        End If
    End If
    
    '  OT Opportunity Description
    If Not IsNull(r_rsTimeEntriesRequest!OT_OPP_DESC) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!OT_OPP_DESC) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "OT Opportunity Description is not valid"
            l_blnValidationError = True
        End If
    End If

    '  Approval Indicator
    If Not IsNull(r_rsTimeEntriesRequest!APPR_IND) Then
        If Not IsYorN(r_rsTimeEntriesRequest!APPR_IND) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Approval indicator is not valid"
            l_blnValidationError = True
        End If
    End If

    '  Component Shift Code
    If Not IsNull(r_rsTimeEntriesRequest!COMP_SHIFT_CD) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!COMP_SHIFT_CD) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Component Shift Code is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    '  Component Code
    If Not IsNull(r_rsTimeEntriesRequest!COMP_CD) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!COMP_CD) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Component Code is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    '  Relief Code
    If Not IsNull(r_rsTimeEntriesRequest!REL_CD) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!REL_CD) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Relief Code is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    ' Family Leave Indicator
    If Not IsNull(r_rsTimeEntriesRequest!FAMILY_LEAVE_IND) Then
        If Not IsAlpha(r_rsTimeEntriesRequest!FAMILY_LEAVE_IND) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Family Leave Indicator is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    '  Schedule Change Exception Code
    If Not IsNull(r_rsTimeEntriesRequest!SCHED_CHG_EXC_CD) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!SCHED_CHG_EXC_CD) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Schedule Change Exception Code is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    '  Distribution Id
    Dim l_blnDistTypeRequired As Boolean
    l_blnDistTypeRequired = False
    If Not IsNull(r_rsTimeEntriesRequest!DISTRIBUTION_ID) Then
        l_blnDistTypeRequired = True
    End If
    If Not IsNull(r_rsTimeEntriesRequest!DISTRIBUTION_ID) And _
            Not IsDistributionIDDescription(r_rsTimeEntriesRequest!DISTRIBUTION_TYPE) Then
        If Not (m_blnIsPOETActive And m_blnIsOTTCEffective) Then
            If Not IsAlphaNumeric(r_rsTimeEntriesRequest!DISTRIBUTION_ID) Then
                AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Id is invalid"
                l_blnValidationError = True
            End If
        Else '//POET
            Dim FeederRef As String
            FeederRef = GetFeederRef39(r_rsTimeEntriesRequest!DISTRIBUTION_ID)
            If Not IsAlphaNumeric(FeederRef) Then
                AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Feeder Ref is invalid (alpha-numeric only)"
                l_blnValidationError = True
            End If
        End If
    End If
    
    '  Distribution type and Distribution Id should both be provided
    '  or neither provided.
    
    '  Distribution type
    '//Dim l_strDistType As String
    If Not IsNull(r_rsTimeEntriesRequest!DISTRIBUTION_TYPE) Then
        l_strDistType = Trim(r_rsTimeEntriesRequest!DISTRIBUTION_TYPE)
        If Not l_strDistType = "" And Not l_strDistType = "A" _
                And Not l_strDistType = "W" And Not l_strDistType = "S" _
                And Not l_strDistType = "J" And Not l_strDistType = "M" _
                And Not l_strDistType = "O" Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution type is invalid"
            l_blnValidationError = True
        End If
    Else
        If l_blnDistTypeRequired Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution type is required when distribution id is provided"
            l_blnValidationError = True
        End If
    End If
    
    ' Timestamp
    If Not IsNull(r_rsTimeEntriesRequest!TIME_STAMP) Then
        If Not IsDB2Timestamp(r_rsTimeEntriesRequest!TIME_STAMP) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Timestamp is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    'Job Nbr
    If Not IsNull(r_rsTimeEntriesRequest!JOB_NBR) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!JOB_NBR) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "EMPRV Job Nbr is invalid or missing"
            l_blnValidationError = True
        End If
    End If
        
    'JobTask Nbr
    If Not IsNull(r_rsTimeEntriesRequest!JOBTASK_NBR) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!JOBTASK_NBR) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "EMPRV JobTask Nbr is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    'JobTask Step Nbr
    If Not IsNull(r_rsTimeEntriesRequest!JOBTASK_STEP_NBR) Then
        If Not IsNumeric(r_rsTimeEntriesRequest!JOBTASK_STEP_NBR) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "EMPRV JobTask Step Nbr is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    'Database ID
    If Not IsNull(r_rsTimeEntriesRequest!EMPRV_DATABASE_ID) Then
        If Not IsAlphaNumeric(r_rsTimeEntriesRequest!EMPRV_DATABASE_ID) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "EMPRV Database ID is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    ValidateOptionalFields = Not l_blnValidationError
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ValidateOptionalFields"

End Function

Private Function ValidateDeleteTimeEntry( _
        ByRef r_rsTimeEntriesRequest As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo eh
    If Trim(r_rsTimeEntriesRequest!SCHED_TIME_IND) = "Y" Then
        ValidateDeleteTimeEntry = _
            ValidateScheduledDeleteTimeEntry( _
                r_rsTimeEntriesRequest, _
                r_rsErrors)
    Else
        ValidateDeleteTimeEntry = _
            ValidateActualDeleteTimeEntry( _
                r_rsTimeEntriesRequest, _
                r_rsErrors)
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ValidateDeleteTimeEntry"

End Function

Private Function ValidateScheduledDeleteTimeEntry( _
        ByRef r_rsTimeEntriesRequest As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset) As Boolean
' Answer true when the request is valid.
    Dim l_blnValidationError As Boolean

On Error GoTo eh
' Ensure required components are provided
    If Not IsNumeric(r_rsTimeEntriesRequest!EMP_NUM) Or _
            r_rsTimeEntriesRequest!EMP_NUM = 0 Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Employee Id is required"
        l_blnValidationError = True
    End If

' 9/13/2001 RSG
' Reject uninitialized dates and dates set to vbNull
    If Not IsDateValid(r_rsTimeEntriesRequest!CHARGE_DATE.Value) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Charge date is required.  Scheduled delete time entry"
        l_blnValidationError = True
    Else
        If r_rsTimeEntriesRequest!CHARGE_DATE > GetMaxChargeDate() Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, GetMaxChargeDateError()
            l_blnValidationError = True
        End If
    End If

    If Not IsNumeric(r_rsTimeEntriesRequest!Hours) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not numeric"
        l_blnValidationError = True
    Else
        If Not (r_rsTimeEntriesRequest!Hours >= 0 And _
                r_rsTimeEntriesRequest!Hours <= 24) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not a valid range"
            l_blnValidationError = True
        End If
    End If

    If Not IsAlphaNumeric(r_rsTimeEntriesRequest!EXCEPTION_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception Code is required"
        l_blnValidationError = True
    End If

'    If Not IsAlpha(r_rsTimeEntriesRequest!FAMILY_LEAVE_IND) Then
'        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Family Leave Indicator is invalid or missing"
'        l_blnValidationError = True
'    End If

' Ensure required key components are provided
    If Not IsNumeric(r_rsTimeEntriesRequest!KEY_EMP_NUM) Or _
            r_rsTimeEntriesRequest!KEY_EMP_NUM = 0 Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Key Employee Id is required"
        l_blnValidationError = True
    End If

' 9/13/2001 RSG
' Reject uninitialized dates and dates set to vbNull
    If Not IsDateValid(r_rsTimeEntriesRequest!KEY_CHARGE_DATE.Value) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Key Charge date is required. Scheduled delete time entry"
        l_blnValidationError = True
    End If

' Ensure all other fields are Null
    If Not IsNull(r_rsTimeEntriesRequest!DATE_WORKED) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Date worked is not null"
        l_blnValidationError = True
    End If

    If Not IsNull(r_rsTimeEntriesRequest!START_TIME) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start time is not null"
        l_blnValidationError = True
    End If

    If Not IsNull(r_rsTimeEntriesRequest!END_TIME) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "End time is not null"
        l_blnValidationError = True
    End If
'
'    If Not IsNull(r_rsTimeEntriesRequest!Hours) Then
'        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not null"
'        l_blnValidationError = True
'    End If
    
'    '  Exception Code
'    If Not IsNull(r_rsTimeEntriesRequest!EXCEPTION_CD) Then
'        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception Code is not null"
'        l_blnValidationError = True
'    End If
    
    '  OverrideException Code
    If Not IsNull(r_rsTimeEntriesRequest!OVER_EXC_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Override Exception code is not null"
        l_blnValidationError = True
    End If
       
    '  Work Group
    If Not IsNull(r_rsTimeEntriesRequest!WG_NUM) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Workgroup is not null"
        l_blnValidationError = True
    End If
        
    '  Distribution Department
    If Not IsNull(r_rsTimeEntriesRequest!DISTR_DEPT_NUM) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Department is not null"
        l_blnValidationError = True
    End If
       
    '  Job Number
    If Not IsNull(r_rsTimeEntriesRequest!JOB_NUM) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Job Number is not null"
        l_blnValidationError = True
    End If
    
    '  OT Reason Code
    If Not IsNull(r_rsTimeEntriesRequest!OT_REASON_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "OT Reason Code is not null"
        l_blnValidationError = True
    End If
    
    '  Pre-Arranged OT Indicator
    If Not PreArrangedOTIndMissing Then
        If Not IsNull(r_rsTimeEntriesRequest!PREARRANGED_OT_IND) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Pre-Arranged OT indicator is not null"
            l_blnValidationError = True
        End If
    End If

    '  OT Opportunity Description
    If Not IsNull(r_rsTimeEntriesRequest!OT_OPP_DESC) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "OT Opportunity Description is not null"
        l_blnValidationError = True
    End If
        
    '  Approval Indicator
    If Not IsNull(r_rsTimeEntriesRequest!APPR_IND) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Approval indicator is not null"
        l_blnValidationError = True
    End If
        
    '  Component Shift Code
    If Not IsNull(r_rsTimeEntriesRequest!COMP_SHIFT_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Component Shift Code is not null"
        l_blnValidationError = True
    End If
    
    '  Component Code
    If Not IsNull(r_rsTimeEntriesRequest!COMP_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Component Code is not null"
        l_blnValidationError = True
    End If
    
    '  Relief Code
    If Not IsNull(r_rsTimeEntriesRequest!REL_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Relief Code is not null"
        l_blnValidationError = True
    End If
    
'    ' Family Leave Indicator
'    If Not IsNull(r_rsTimeEntriesRequest!FAMILY_LEAVE_IND) Then
'        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Family Leave Indicator is not null"
'        l_blnValidationError = True
'    End If
    
    '  Schedule Change Exception Code
    If Not IsNull(r_rsTimeEntriesRequest!SCHED_CHG_EXC_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Schedule Change Exception Code is not null"
        l_blnValidationError = True
    End If
    
    '  Distribution Id
    If Not IsNull(r_rsTimeEntriesRequest!DISTRIBUTION_ID) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Id is not null"
        l_blnValidationError = True
    End If
    
    '  Distribution type
    If Not IsNull(r_rsTimeEntriesRequest!DISTRIBUTION_TYPE) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution type is not null"
        l_blnValidationError = True
    End If
    
    '  Job Nbr
    If Not IsNull(r_rsTimeEntriesRequest!JOB_NBR) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Job Nbr is not null"
        l_blnValidationError = True
    End If
    
    '  JobTask Nbr
    If Not IsNull(r_rsTimeEntriesRequest!JOBTASK_NBR) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "JobTask Nbr is not null"
        l_blnValidationError = True
    End If

    '  JobTask Step Nbr
    If Not IsNull(r_rsTimeEntriesRequest!JOBTASK_STEP_NBR) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "JobTask Step Nbr is not null"
        l_blnValidationError = True
    End If
    
    '  Database ID
    If Not IsNull(r_rsTimeEntriesRequest!EMPRV_DATABASE_ID) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Database ID is not null"
        l_blnValidationError = True
    End If
    
    ValidateScheduledDeleteTimeEntry = Not l_blnValidationError
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ValidateScheduledDeleteTimeEntry"

End Function


Private Function ValidateActualDeleteTimeEntry( _
        ByRef r_rsTimeEntriesRequest As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset) As Boolean
' Answer true when the request is valid.
' A delete request for existing time entry should only have the
' key fields.  Most other fields should be Null.
    Dim l_blnValidationError As Boolean

On Error GoTo eh
' Ensure key components are provided, if not stop here
    If Not ValidateKeyComponents(r_rsTimeEntriesRequest, r_rsErrors) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCAPPLICATIONERR, "One or more key components are invalid"
        ValidateActualDeleteTimeEntry = False
        Exit Function
    End If
    
    If IsDateValid(r_rsTimeEntriesRequest!KEY_CHARGE_DATE.Value) Then
        If r_rsTimeEntriesRequest!KEY_CHARGE_DATE > GetMaxChargeDate() Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, GetMaxChargeDateError()
            l_blnValidationError = True
        End If
    End If

    ' Schedule Indicator
    If Not IsYorN(r_rsTimeEntriesRequest!SCHED_TIME_IND) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Schedule Indicator must be Y or N"
        l_blnValidationError = True
    End If

' Ensure all other fields are Null
    If Not IsNull(r_rsTimeEntriesRequest!EMP_NUM) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Employee Id is not null"
        l_blnValidationError = True
    End If

    If Not IsNull(r_rsTimeEntriesRequest!CHARGE_DATE) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Charge Date is not null"
        l_blnValidationError = True
    End If

    If Not IsNull(r_rsTimeEntriesRequest!DATE_WORKED) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Date worked is not null"
        l_blnValidationError = True
    End If

    If Not IsNull(r_rsTimeEntriesRequest!START_TIME) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start time is not null"
        l_blnValidationError = True
    End If

    If Not IsNull(r_rsTimeEntriesRequest!END_TIME) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "End time is not null"
        l_blnValidationError = True
    End If

    If Not IsNull(r_rsTimeEntriesRequest!Hours) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not null"
        l_blnValidationError = True
    End If
    
    '  Exception Code
    If Not IsNull(r_rsTimeEntriesRequest!EXCEPTION_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception Code is not null"
        l_blnValidationError = True
    End If
    
    '  OverrideException Code
    If Not IsNull(r_rsTimeEntriesRequest!OVER_EXC_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Override Exception code is not null"
        l_blnValidationError = True
    End If
       
    '  Work Group
    If Not IsNull(r_rsTimeEntriesRequest!WG_NUM) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Workgroup is not null"
        l_blnValidationError = True
    End If
        
    '  Distribution Department
    If Not IsNull(r_rsTimeEntriesRequest!DISTR_DEPT_NUM) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Department is not null"
        l_blnValidationError = True
    End If
       
    '  Job Number
    If Not IsNull(r_rsTimeEntriesRequest!JOB_NUM) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Job Number is not null"
        l_blnValidationError = True
    End If
    
    '  OT Reason Code
    If Not IsNull(r_rsTimeEntriesRequest!OT_REASON_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "OT Reason Code is not null"
        l_blnValidationError = True
    End If
    
    '  Pre-Arranged OT Indicator
    If Not PreArrangedOTIndMissing Then
        If Not IsNull(r_rsTimeEntriesRequest!PREARRANGED_OT_IND) Then
            AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Pre-Arranged OT indicator is not null"
            l_blnValidationError = True
        End If
    End If
        
    '  OT Opportunity Description
    If Not IsNull(r_rsTimeEntriesRequest!OT_OPP_DESC) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "OT Opportunity Description is not null"
        l_blnValidationError = True
    End If
        
    '  Approval Indicator
    If Not IsNull(r_rsTimeEntriesRequest!APPR_IND) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Approval indicator is not null"
        l_blnValidationError = True
    End If
        
    '  Component Shift Code
    If Not IsNull(r_rsTimeEntriesRequest!COMP_SHIFT_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Component Shift Code is not null"
        l_blnValidationError = True
    End If
    
    '  Component Code
    If Not IsNull(r_rsTimeEntriesRequest!COMP_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Component Code is not null"
        l_blnValidationError = True
    End If
    
    '  Relief Code
    If Not IsNull(r_rsTimeEntriesRequest!REL_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Relief Code is not null"
        l_blnValidationError = True
    End If
    
    ' Family Leave Indicator
    If Not IsNull(r_rsTimeEntriesRequest!FAMILY_LEAVE_IND) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Family Leave Indicator is not null"
        l_blnValidationError = True
    End If
    
    '  Schedule Change Exception Code
    If Not IsNull(r_rsTimeEntriesRequest!SCHED_CHG_EXC_CD) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Schedule Change Exception Code is not null"
        l_blnValidationError = True
    End If
    
    '  Distribution Id
    If Not IsNull(r_rsTimeEntriesRequest!DISTRIBUTION_ID) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Id is not null"
        l_blnValidationError = True
    End If
    
    '  Distribution type
    If Not IsNull(r_rsTimeEntriesRequest!DISTRIBUTION_TYPE) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution type is not null"
        l_blnValidationError = True
    End If
    
    '  Job Nbr
    If Not IsNull(r_rsTimeEntriesRequest!JOB_NBR) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Job Nbr is not null"
        l_blnValidationError = True
    End If
    
    '  JobTask Nbr
    If Not IsNull(r_rsTimeEntriesRequest!JOBTASK_NBR) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "JobTask Nbr is not null"
        l_blnValidationError = True
    End If

    '  JobTask Step Nbr
    If Not IsNull(r_rsTimeEntriesRequest!JOBTASK_STEP_NBR) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "JobTask Step Nbr is not null"
        l_blnValidationError = True
    End If

    '  Database ID
    If Not IsNull(r_rsTimeEntriesRequest!EMPRV_DATABASE_ID) Then
        AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Database ID is not null"
        l_blnValidationError = True
    End If

    ValidateActualDeleteTimeEntry = Not l_blnValidationError
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ValidateActualDeleteTimeEntry"

End Function

Private Function ValidateKeyComponents( _
            r_rsTimeEntriesRequest As ADODB.Recordset, _
            r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_blnValidationError As Boolean

On Error GoTo eh
    If Not IsNumeric(r_rsTimeEntriesRequest!KEY_EMP_NUM) Or _
            r_rsTimeEntriesRequest!KEY_EMP_NUM = 0 Then
        l_blnValidationError = True
    End If

' 9/13/2001 RSG
' Reject uninitialized dates and dates set to vbNull
    If Not IsDateValid(r_rsTimeEntriesRequest!KEY_CHARGE_DATE.Value) Then
        l_blnValidationError = True
    End If

    If Not IsTimeValid(r_rsTimeEntriesRequest!KEY_START_TIME) Then
        l_blnValidationError = True
    End If

    If Not IsDB2Timestamp(r_rsTimeEntriesRequest!KEY_INDEX_TIME) Then
        l_blnValidationError = True
    End If
    
    ValidateKeyComponents = Not l_blnValidationError
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "ValidateKeyComponents"

End Function

'Private Function ValidateECDrawTime( _
'                    ByRef v_rsECDrawTime As ADODB.Recordset, _
'                    ByRef r_rsErrors As ADODB.Recordset _
'                    ) As Boolean
'    Dim l_blnValidationError As Boolean
'
'On Error GoTo eh
'    If v_rsECDrawTime Is Nothing Then
'        ValidateECDrawTime = True
'        Exit Function
'    End If
'
'    If v_rsECDrawTime.BOF And _
'            v_rsECDrawTime.EOF Then
'        ValidateECDrawTime = True
'        Exit Function
'    End If
'
'    v_rsECDrawTime.MoveFirst
'    Do Until v_rsECDrawTime.EOF
'        If Not IsNull(v_rsECDrawTime!EMP_NUM) Then
'            If v_rsECDrawTime!EMP_NUM < 1 Then
'                AppendECDrawError r_rsErrors, "ECDraw - Invalid employee number"
'                l_blnValidationError = True
'            End If
'        Else
'            AppendECDrawError r_rsErrors, "ECDraw - Invalid employee number; NULL"
'            l_blnValidationError = True
'        End If
'
'        If IsNull(v_rsECDrawTime!WEEK_ENDING_DATE) Then
'            AppendECDrawError r_rsErrors, "ECDraw - Invalid week ending date; NULL"
'            l_blnValidationError = True
'        End If
'
'        If IsNull(v_rsECDrawTime!DRWG_HOURS) Then
'            AppendECDrawError r_rsErrors, "ECDraw - Invalid draw hours; NULL"
'            l_blnValidationError = True
'        End If
'
'        If IsNull(v_rsECDrawTime!ACCT_NUM) Then
'            AppendECDrawError r_rsErrors, "ECDraw - Invalid account number; NULL"
'            l_blnValidationError = True
'        ElseIf Trim(v_rsECDrawTime!ACCT_NUM) = "" Then
'            AppendECDrawError r_rsErrors, "ECDraw - Invalid account number; empty string"
'            l_blnValidationError = True
'        End If
'
''        If IsNull(v_rsECDrawTime!WORK_NUM) Then
''            AppendECDrawError r_rsErrors, "ECDraw - Invalid work number; NULL"
''            l_blnValidationError = True
''        ElseIf Trim(v_rsECDrawTime!WORK_NUM) = "" Then
''            AppendECDrawError r_rsErrors, "ECDraw - Invalid work number; empty string"
''            l_blnValidationError = True
''        End If
''
''        If IsNull(v_rsECDrawTime!CLASS_NUM) Then
''            AppendECDrawError r_rsErrors, "ECDraw - Invalid class number; NULL"
''            l_blnValidationError = True
''        ElseIf Trim(v_rsECDrawTime!CLASS_NUM) = "" Then
''            AppendECDrawError r_rsErrors, "ECDraw - Invalid class number; empty string"
''            l_blnValidationError = True
''        End If
'
'        If Not IsNull(v_rsECDrawTime!RECORD_STATUS) Then
'            If Trim(v_rsECDrawTime!RECORD_STATUS) = "D" Then
'                If IsNull(v_rsECDrawTime!TIME_STAMP) Then
'                    AppendECDrawError r_rsErrors, "ECDraw - Invalid timestamp; NULL"
'                    l_blnValidationError = True
'                ElseIf Not IsDB2Timestamp(v_rsECDrawTime!TIME_STAMP) Then
'                    AppendECDrawError r_rsErrors, "ECDraw - Invalid timestamp; invalid format"
'                    l_blnValidationError = True
'                End If
'            Else
'                If IsNull(v_rsECDrawTime!WORK_NUM) Then
'                    AppendECDrawError r_rsErrors, "ECDraw - Invalid work number; NULL"
'                    l_blnValidationError = True
'                ElseIf Trim(v_rsECDrawTime!WORK_NUM) = "" Then
'                    AppendECDrawError r_rsErrors, "ECDraw - Invalid work number; empty string"
'                    l_blnValidationError = True
'                End If
'                If IsNull(v_rsECDrawTime!CLASS_NUM) Then
'                    AppendECDrawError r_rsErrors, "ECDraw - Invalid class number; NULL"
'                    l_blnValidationError = True
'                ElseIf Trim(v_rsECDrawTime!CLASS_NUM) = "" Then
'                    AppendECDrawError r_rsErrors, "ECDraw - Invalid class number; empty string"
'                    l_blnValidationError = True
'                End If
'            End If
'        Else
'            AppendECDrawError r_rsErrors, "ECDraw - Invalid record status; NULL"
'            l_blnValidationError = True
'        End If
'
'        v_rsECDrawTime.MoveNext
'    Loop
'
'    v_rsECDrawTime.MoveFirst
'    ValidateECDrawTime = Not l_blnValidationError
'    Exit Function
'
'eh:
'    TrsRaiseError mstrCMODULENAME, "ValidateECDrawTime"
'
'End Function
Private Function ValidateEMPRVJobTaskTime( _
                    ByVal v_rsTimeEntries As ADODB.Recordset, _
                    ByVal v_rsJobTaskSteps As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                    ) As Boolean
    
    Dim l_blnValidationError As Boolean
    Dim l_blnValidStep As Boolean
    Dim l_strFilter As String
    Dim l_strFilter2 As String
    Dim l_strFilter3 As String
    Dim l_objbusCodeTables As busCodeTables.cbsCodeTables
    Dim l_adrsJobTaskExceptions As ADODB.Recordset
    
On Error GoTo eh

    If v_rsTimeEntries Is Nothing Then
        ValidateEMPRVJobTaskTime = True
        Exit Function
    End If

    If v_rsTimeEntries.BOF And _
            v_rsTimeEntries.EOF Then
        ValidateEMPRVJobTaskTime = True
        Exit Function
    End If
   
    l_strFilter2 = "RECORD_STATUS <> 'D' AND JOB_NBR <> ''"
    v_rsTimeEntries.Filter = l_strFilter2
    If Not v_rsTimeEntries.BOF And Not v_rsTimeEntries.EOF Then
        v_rsTimeEntries.MoveFirst
        Set l_objbusCodeTables = CreateObject("busCodeTables.cbsCodeTables")
        Set l_adrsJobTaskExceptions = l_objbusCodeTables.GetJobTaskExceptions
        Set l_objbusCodeTables = Nothing
        Do Until v_rsTimeEntries.EOF
            If Not IsNumeric(v_rsTimeEntries!JOBTASK_STEP_NBR) Then
                AppendEntryError r_rsErrors, v_rsTimeEntries, trErrbusValidationError, mstrCBUSINESSERR, "EMPRV JobTask - Invalid step number"
                l_blnValidationError = True
            End If
            
            If Trim(v_rsTimeEntries!JOB_NBR) <> "" Or _
               Trim(v_rsTimeEntries!JOBTASK_NBR) <> "" Or _
               Trim(v_rsTimeEntries!EMPRV_DATABASE_ID) <> "" Then
               
               '//older clients may not have new column
               If FieldExists("VALIDATION_TYPE", v_rsJobTaskSteps) Then
                    l_strFilter = "JOB_NBR = '" & Trim(v_rsTimeEntries!JOB_NBR) & "'" & _
                                " AND JOBTASK_NBR = '" & Trim(v_rsTimeEntries!JOBTASK_NBR) & "'" & _
                                " AND EMPRV_DATABASE_ID = '" & Trim(v_rsTimeEntries!EMPRV_DATABASE_ID) & "'" & _
                                " AND VALIDATION_TYPE = 'S'"
                Else
                    l_strFilter = "JOB_NBR = '" & Trim(v_rsTimeEntries!JOB_NBR) & "'" & _
                                " AND JOBTASK_NBR = '" & Trim(v_rsTimeEntries!JOBTASK_NBR) & "'" & _
                                " AND EMPRV_DATABASE_ID = '" & Trim(v_rsTimeEntries!EMPRV_DATABASE_ID) & "'"
                End If
                            
                v_rsJobTaskSteps.Filter = l_strFilter
                If Not v_rsJobTaskSteps.BOF And Not v_rsJobTaskSteps.EOF Then '//RecordsExist in timekeeper removes filter!!!
                    If v_rsTimeEntries!JOBTASK_STEP_NBR = 0 Then
                        AppendEntryError r_rsErrors, v_rsTimeEntries, trErrbusValidationError, mstrCBUSINESSERR, "EMPRV JobTask - Must select a step number"
                        l_blnValidationError = True
                    End If
                End If
            End If
            v_rsJobTaskSteps.Filter = adFilterNone
            
            If Trim(v_rsTimeEntries!EXCEPTION_CD) <> "" Then
                l_strFilter3 = "EXCEPTION_CD = '" & Trim(v_rsTimeEntries!EXCEPTION_CD) & "'"
                l_adrsJobTaskExceptions.Filter = l_strFilter3
                If l_adrsJobTaskExceptions.BOF And l_adrsJobTaskExceptions.EOF Then
                    AppendEntryError r_rsErrors, v_rsTimeEntries, trErrbusValidationError, mstrCBUSINESSERR, "This exception code is not valid for an EMPRV JobTask"
                    l_blnValidationError = True
                End If
                l_adrsJobTaskExceptions.Filter = adFilterNone
            End If
                
            v_rsTimeEntries.MoveNext
        Loop
    Else
        v_rsTimeEntries.Filter = adFilterNone
        v_rsTimeEntries.MoveFirst
        ValidateEMPRVJobTaskTime = True
        Exit Function
    End If

    v_rsTimeEntries.MoveFirst
    ValidateEMPRVJobTaskTime = Not l_blnValidationError
    Exit Function

eh:
    Set l_objbusCodeTables = Nothing
    TrsRaiseError mstrCMODULENAME, "ValidateEMPRVJobTaskTime"

End Function

'Private Sub AppendECDrawError( _
'                ByRef r_rsErrors As ADODB.Recordset, _
'                ByVal v_strErrorMsg As String)
'
'    AppendError r_rsErrors, 0, Date, Date, "", "", 0, "", "", 40100 + vbObjectError, mstrCBUSINESSERR, v_strErrorMsg, mstrCMODULENAME
'
'End Sub

Private Function ValidateInputOtherPay( _
                    ByRef r_rsOtherpayRequest As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                    ) As Boolean
    
    Dim l_blnValidationError As Boolean
    
On Error GoTo ValidateInputOtherPay_EH

    ' If there are no records then bypass Other Pay
    If Not r_rsOtherpayRequest Is Nothing Then
        If r_rsOtherpayRequest.BOF And r_rsOtherpayRequest.EOF Then
            ValidateInputOtherPay = True
            Exit Function
        End If
    Else
        ValidateInputOtherPay = True
        Exit Function
    End If
    
    r_rsOtherpayRequest.MoveFirst
    
    ' Each type of request requires different validations
    Do Until r_rsOtherpayRequest.EOF
        Select Case UCase(r_rsOtherpayRequest!RECORD_STATUS)
            Case "I"
                If Not ValidateInputNewOtherPay(r_rsOtherpayRequest, r_rsErrors) Then
                    l_blnValidationError = True
                End If
            Case "C"
                If Not ValidateInputUpdateOtherPay(r_rsOtherpayRequest, r_rsErrors) Then
                    l_blnValidationError = True
                End If
            Case "D"
                If Not ValidateInputDeleteOtherPay(r_rsOtherpayRequest, r_rsErrors) Then
                    l_blnValidationError = True
                End If
            Case Else
                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Record status must be I, C, or D"
                l_blnValidationError = True
        End Select
        r_rsOtherpayRequest.MoveNext
    Loop

    ValidateInputOtherPay = Not l_blnValidationError
    r_rsOtherpayRequest.MoveFirst
    Exit Function
    
ValidateInputOtherPay_EH:

    TrsRaiseError mstrCMODULENAME, "ValidateInputOtherPay"

End Function

Private Function ValidateInputOtherPayOptionalFields( _
                ByRef r_rsOtherpayRequest As ADODB.Recordset, _
                ByRef r_rsErrors As ADODB.Recordset _
                ) As Boolean
                
    Dim l_blnValidationError As Boolean
    Dim l_blnDistTypeReqd As Boolean
    Dim l_strDistType As String
    
On Error GoTo ValidateOtherPayOptionalFields_EH
        
        ' Job Number
        If Not IsNull(r_rsOtherpayRequest!JOB_NUM) Then
            If Not IsAlphaNumeric(r_rsOtherpayRequest!JOB_NUM) Then
                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Job Number is not numeric."
                l_blnValidationError = True
            End If
        End If
        
        '  WorkGroup
        If Not IsNull(r_rsOtherpayRequest!WG_NUM) Then
            If Not IsNumeric(r_rsOtherpayRequest!WG_NUM) Then
                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Workgroup is not numeric."
                l_blnValidationError = True
            End If
        End If
        
'        ' Department Number
'        If Not IsNull(r_rsOtherPayRequest!DEPT_NUM) Then
'            If Not IsAlphaNumeric(r_rsOtherPayRequest!DEPT_NUM) Then
'                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherPayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Department Number is not valid."
'                l_blnValidationError = True
'            End If
'        End If

        ' Start Location
        If Not IsNull(r_rsOtherpayRequest!START_LOC_CD) Then
            If Not IsNumeric(r_rsOtherpayRequest!START_LOC_CD) Then
                    AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start Location Code is not numeric"
                    l_blnValidationError = True
            End If
        End If
                
        ' End Location
        If Not IsNull(r_rsOtherpayRequest!END_LOC_CD) Then
            If Not IsNumeric(r_rsOtherpayRequest!END_LOC_CD) Then
                    AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "End Location Code is not numeric"
                    l_blnValidationError = True
            End If
        End If
        
'        ' Location Code
'        If Not IsNull(r_rsOtherPayRequest!LOC_CD) Then
'            If Not IsNumeric(r_rsOtherPayRequest!LOC_CD) Then
'                    AppendEntryErrorOtherPay r_rsErrors, r_rsOtherPayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Location Code is not numeric"
'                    l_blnValidationError = True
'            End If
'        End If
        
        ' Component Shift Code
        If Not IsNull(r_rsOtherpayRequest!COMP_SHIFT_CD) Then
            If Not IsAlphaNumeric(r_rsOtherpayRequest!COMP_SHIFT_CD) Then
                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Component Shift Code is not valid or missing"
                l_blnValidationError = True
            End If
        End If
        
        ' Distribution Id
        l_blnDistTypeReqd = False
        If Not IsNull(r_rsOtherpayRequest!DISTRIBUTION_ID) Then
            l_blnDistTypeReqd = True                                                                        ' We need a Distribution Type
            If Not (m_blnIsPOETActive And m_blnIsOTTCEffective) Then
                If Not IsAlphaNumeric(r_rsOtherpayRequest!DISTRIBUTION_ID) Then
                    AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Id is not valid or missing"
                    l_blnValidationError = True
                End If
            Else '//POET
                Dim FeederRef As String
                FeederRef = GetFeederRef39(r_rsOtherpayRequest!DISTRIBUTION_ID)
                If Not IsAlphaNumeric(FeederRef) Then
                    AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Feeder Ref is invalid (alpha-numeric only)"
                    l_blnValidationError = True
                End If
            End If
        End If
        
        ' Distribution Type
        If Not IsNull(r_rsOtherpayRequest!DISTRIBUTION_TYPE) Then
            l_strDistType = Trim(r_rsOtherpayRequest!DISTRIBUTION_TYPE)
            If Not l_strDistType = "" And Not l_strDistType = "A" And Not l_strDistType = "W" And Not l_strDistType = "S" Then
                    AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Type is not valid"
                    l_blnValidationError = True
            End If
        Else
            If l_blnDistTypeReqd Then
                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Type is required with Distribution Id."
                l_blnValidationError = True
            End If
        End If
        
        ' Distribution Number
        If Not IsNull(r_rsOtherpayRequest!DISTR_DEPT_NUM) Then
            If Not IsAlphaNumeric(r_rsOtherpayRequest!DISTR_DEPT_NUM) Then
                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Department is not valid or missing."
                l_blnValidationError = True
            End If
        End If
        
        ' Approval Indicator
        If Not IsNull(r_rsOtherpayRequest!APPR_IND) Then
            If Not IsYorN(r_rsOtherpayRequest!APPR_IND) Then
                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Approval Indicator is not valid or missing."
                l_blnValidationError = True
            End If
        End If
               
    ValidateInputOtherPayOptionalFields = Not l_blnValidationError
    Exit Function
    
ValidateOtherPayOptionalFields_EH:
    TrsRaiseError mstrCMODULENAME, "ValidateInputOtherPayOptionalFields"

End Function

Private Function ValidateInputNewOtherPay( _
                    ByRef r_rsOtherpayRequest As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                    ) As Boolean
    Dim l_strExcCd As String
    Dim l_blnValidationError As Boolean
    Dim l_lngFirstEmpNum As Long

On Error GoTo ValidateInputNewOtherPay_EH
    
    ' KEY Employee Number
    If Not r_rsOtherpayRequest!KEY_EMP_NUM = r_rsOtherpayRequest!EMP_NUM Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Client Key application error."
        l_blnValidationError = True
    End If
    
    ' KEY Charge Date
    If Not r_rsOtherpayRequest!KEY_CHARGE_DATE = r_rsOtherpayRequest!CHARGE_DATE Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Client Key application error."
        l_blnValidationError = True
    End If
    
    ' KEY Start Time
    If Not r_rsOtherpayRequest!KEY_START_TIME = r_rsOtherpayRequest!START_TIME Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Client Key application error."
        l_blnValidationError = True
    End If
        
    ' Employee Number
    If Not IsNumeric(r_rsOtherpayRequest!EMP_NUM) Or r_rsOtherpayRequest!EMP_NUM = 0 Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Employee Id is invalid"
            l_blnValidationError = True
    End If
    
    '  Charge Date
' 9/13/2001 RSG
' Reject uninitialized dates and dates set to vbNull
    If Not IsDateValid(r_rsOtherpayRequest!CHARGE_DATE.Value) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Charge Date is not a valid date. New other pay."
        l_blnValidationError = True
    End If
    
    '  Date Worked
' 9/13/2001 RSG
' Reject uninitialized dates and dates set to vbNull
    If Not IsDateValid(r_rsOtherpayRequest!DATE_WORKED.Value) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Date Worked is not a valid date"
        l_blnValidationError = True
    End If

    '  Start Time
    If InStr(r_rsOtherpayRequest!START_TIME, "99") = 0 Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start time is not valid Other Pay Start Time."
        l_blnValidationError = True
    End If
    
    ' End Time
    If InStr(r_rsOtherpayRequest!END_TIME, "99") = 0 Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "End time is not valid Other Pay End Time."
        l_blnValidationError = True
    End If
                
    '  Exception Code
    If Not IsNull(r_rsOtherpayRequest!EXCEPTION_CD) Then
        If Not IsAlphaNumeric(r_rsOtherpayRequest!EXCEPTION_CD) And r_rsOtherpayRequest!EXCEPTION_CD <> "" Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception code is not valid"
            l_blnValidationError = True
        End If
    Else
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception code is not valid; Null."
        l_blnValidationError = True
    End If
    
    ' Hours
    If Not IsNull(r_rsOtherpayRequest!Hours) Then
        If Not IsNumeric(r_rsOtherpayRequest!Hours) Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not numeric"
            l_blnValidationError = True
        Else
            If Not (r_rsOtherpayRequest!Hours >= 0 And r_rsOtherpayRequest!Hours <= 24) Then
                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not a valid range"
                l_blnValidationError = True
            End If
        End If
    End If
    
    ' Amount
    If Not IsNull(r_rsOtherpayRequest!AMOUNT) Then
        If Not IsNumeric(r_rsOtherpayRequest!AMOUNT) Then
                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Amount is not numeric"
                l_blnValidationError = True
        End If
    End If

    ' Optional Fields
    If Not ValidateInputOtherPayOptionalFields(r_rsOtherpayRequest, r_rsErrors) Then
        l_blnValidationError = True
    End If
    
    ValidateInputNewOtherPay = Not l_blnValidationError
    Exit Function
    
ValidateInputNewOtherPay_EH:
    TrsRaiseError mstrCMODULENAME, "ValidateInputNewOtherPay"

End Function

Private Function ValidateInputDeleteOtherPay( _
                    ByRef r_rsOtherpayRequest As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                    ) As Boolean

    Dim l_blnValidationError As Boolean

On Error GoTo ValidateInputDeleteOtherPay_EH

    ' Key fields must be present
    If Not IsNumeric(r_rsOtherpayRequest!KEY_EMP_NUM) Or r_rsOtherpayRequest!KEY_EMP_NUM = 0 Then
        l_blnValidationError = True
    End If

    If Not IsDate(r_rsOtherpayRequest!KEY_CHARGE_DATE) Then
        l_blnValidationError = True
    End If

    If InStr(r_rsOtherpayRequest!KEY_START_TIME, "99") = 0 Then
        l_blnValidationError = True
    End If

    If Not IsDB2Timestamp(r_rsOtherpayRequest!KEY_INDEX_TIME) Then
        l_blnValidationError = True
    End If

    ' All other fields must be Null
    ' Employee Number
    If Not IsNull(r_rsOtherpayRequest!EMP_NUM) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Employee Number is not null."
        l_blnValidationError = True
    End If
    
    ' Charge Date
    If Not IsNull(r_rsOtherpayRequest!CHARGE_DATE) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Charge Date is not null."
        l_blnValidationError = True
    End If

    ' Date Worked
    If Not IsNull(r_rsOtherpayRequest!DATE_WORKED) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Date Worked is not null."
        l_blnValidationError = True
    End If
    
    ' Start Time
    If Not IsNull(r_rsOtherpayRequest!START_TIME) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start Time is not null."
        l_blnValidationError = True
    End If
    
    ' End Time
    If Not IsNull(r_rsOtherpayRequest!END_TIME) Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "End time is not null."
            l_blnValidationError = True
    End If
    
    ' Job Number
    If Not IsNull(r_rsOtherpayRequest!JOB_NUM) Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Job Number is not null."
            l_blnValidationError = True
    End If
    
    '  WorkGroup
    If Not IsNull(r_rsOtherpayRequest!WG_NUM) Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Workgroup is not null."
            l_blnValidationError = True
    End If
    
'    ' Department Number
'    If Not IsNull(r_rsOtherPayRequest!DEPT_NUM) Then
'            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherPayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Department Number is not null."
'            l_blnValidationError = True
'    End If

    ' Hours
    If Not IsNull(r_rsOtherpayRequest!Hours) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours are not null."
        l_blnValidationError = True
    End If

    ' Amount
    If Not IsNull(r_rsOtherpayRequest!AMOUNT) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Amount is not null."
        l_blnValidationError = True
    End If

    ' Start Location
    If Not IsNull(r_rsOtherpayRequest!START_LOC_CD) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start Location Code is not null"
        l_blnValidationError = True
    End If
            
    ' End Location
    If Not IsNull(r_rsOtherpayRequest!END_LOC_CD) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "End Location Code is not null"
        l_blnValidationError = True
    End If
    
'    ' Location Code
'    If Not IsNull(r_rsOtherPayRequest!LOC_CD) Then
'        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherPayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Location Code is not null"
'        l_blnValidationError = True
'    End If
    
    ' Component Shift Code
    If Not IsNull(r_rsOtherpayRequest!COMP_SHIFT_CD) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Component Shift Code is not null."
        l_blnValidationError = True
    End If
    
    ' Distribution Id
    If Not IsNull(r_rsOtherpayRequest!DISTRIBUTION_ID) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Id is not null."
        l_blnValidationError = True
    End If
    
    ' Distribution Type
    If Not IsNull(r_rsOtherpayRequest!DISTRIBUTION_TYPE) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Type is not null."
        l_blnValidationError = True
    End If
    
    ' Distribution Number
    If Not IsNull(r_rsOtherpayRequest!DISTR_DEPT_NUM) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Distribution Department is not null."
        l_blnValidationError = True
    End If
    
    ' Exception Code
    If Not IsNull(r_rsOtherpayRequest!EXCEPTION_CD) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception Code is not null."
        l_blnValidationError = True
    End If
    
    ' Approval Indicator
    If Not IsNull(r_rsOtherpayRequest!APPR_IND) Then
        AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Approval Indicator is not null."
        l_blnValidationError = True
    End If

    ValidateInputDeleteOtherPay = Not l_blnValidationError
    Exit Function
    
ValidateInputDeleteOtherPay_EH:
    TrsRaiseError mstrCMODULENAME, "ValidateInputDeleteOtherPay"

End Function

Private Function ValidateInputUpdateOtherPay( _
                    ByRef r_rsOtherpayRequest As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset) As Boolean

    Dim l_blnValidationError As Boolean

On Error GoTo ValidateInputUpdateOtherPay_EH

    '  All Key Components must be present
    If Not IsNumeric(r_rsOtherpayRequest!KEY_EMP_NUM) Or r_rsOtherpayRequest!KEY_EMP_NUM = 0 Then
        l_blnValidationError = True
    End If

' 9/13/2001 RSG
' Reject uninitialized dates and dates set to vbNull
    If Not IsDateValid(r_rsOtherpayRequest!KEY_CHARGE_DATE.Value) Then
        l_blnValidationError = True
    End If

    If InStr(r_rsOtherpayRequest!KEY_START_TIME, "99") = 0 Then
        l_blnValidationError = True
    End If

    If Not IsDB2Timestamp(r_rsOtherpayRequest!KEY_INDEX_TIME) Then
        l_blnValidationError = True
    End If

    ' Validate Required Components
    '  Employee Number
    If Not IsNull(r_rsOtherpayRequest!EMP_NUM) Then
        If Not r_rsOtherpayRequest!EMP_NUM = r_rsOtherpayRequest!KEY_EMP_NUM Then
            AppendEntryError r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Employee Id is not valid"
            l_blnValidationError = True
        End If
    End If

    '  Charge Date
    If Not IsNull(r_rsOtherpayRequest!CHARGE_DATE) Then
        If Not r_rsOtherpayRequest!CHARGE_DATE = r_rsOtherpayRequest!KEY_CHARGE_DATE Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Charge Date is not a valid date. Change other pay"
            l_blnValidationError = True
        End If
    End If

    ' Date Worked
' 9/13/2001 RSG
' Reject uninitialized dates and dates set to vbNull
    If Not IsNull(r_rsOtherpayRequest!DATE_WORKED) Then
        If Not IsDateValid(r_rsOtherpayRequest!DATE_WORKED.Value) Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Date worked is not a valid date"
            l_blnValidationError = True
        End If
    End If
    
    ' Start Time
    If Not IsNull(r_rsOtherpayRequest!START_TIME) Then
        If InStr(r_rsOtherpayRequest!START_TIME, "99") = 0 Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Start time is not valid"
            l_blnValidationError = True
        End If
    End If
    
    '  End Time
    If Not IsNull(r_rsOtherpayRequest!END_TIME) Then
        If InStr(r_rsOtherpayRequest!START_TIME, "99") = 0 Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "End time is not valid"
            l_blnValidationError = True
        End If
    End If

    ' Hours
    If Not IsNull(r_rsOtherpayRequest!Hours) Then
        If Not IsNumeric(r_rsOtherpayRequest!Hours) Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not numeric"
            l_blnValidationError = True
        Else
            If Not (r_rsOtherpayRequest!Hours >= 0 And _
                    r_rsOtherpayRequest!Hours <= 24) Then
                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Hours is not a valid range"
                l_blnValidationError = True
            End If
        End If
    End If
    
    ' Amount
    If Not IsNull(r_rsOtherpayRequest!AMOUNT) Then
        If Not IsNumeric(r_rsOtherpayRequest!AMOUNT) Then
                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Amount is not numeric"
                l_blnValidationError = True
        End If
    End If
    
    If Not ValidateInputOtherPayOptionalFields(r_rsOtherpayRequest, r_rsErrors) Then
        l_blnValidationError = True
    End If
    
    ' Exception Code
    If Not IsNull(r_rsOtherpayRequest!EXCEPTION_CD) Then
        If Not IsAlphaNumeric(r_rsOtherpayRequest!EXCEPTION_CD) And r_rsOtherpayRequest!EXCEPTION_CD <> "" Then
            AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception Code is not valid"
            l_blnValidationError = True
        End If
    End If
    
    ValidateInputUpdateOtherPay = Not l_blnValidationError
    Exit Function

ValidateInputUpdateOtherPay_EH:
    TrsRaiseError mstrCMODULENAME, "ValidateInputUpdateOtherPay"

End Function

Private Function IsDistributionIDDescription(distType As Variant) As Boolean
    If Not IsNull(distType) Then
        If Trim(CStr(distType)) = "" Then
            IsDistributionIDDescription = True
            Exit Function
        End If
    End If
    IsDistributionIDDescription = False

End Function

Private Function IsDateValid(aDate As Variant) As Boolean
'Reject aDate when unitialized(vbEmpty) or set to vbNull.
'Guard against Null.

On Error GoTo eh
    If IsNull(aDate) Then
        IsDateValid = False
    Else
        IsDateValid = _
            IsDate(aDate) And aDate <> vbEmpty And aDate <> vbNull
    End If
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsDateValid"
End Function

'Recordset definintion we are expecting

'.Fields.Append "EMP_NUM", adInteger, , l_lngCFIELDATTR
'.Fields.Append "CHARGE_DATE", adDate, , l_lngCFIELDATTR
'.Fields.Append "DATE_WORKED", adDate, , l_lngCFIELDATTR
'.Fields.Append "DISTRIBUTION_ID", adChar, 29, l_lngCFIELDATTR
'.Fields.Append "DISTRIBUTION_TYPE", adChar, 1, l_lngCFIELDATTR
'.Fields.Append "HOURS", adDecimal, , l_lngCFIELDATTR
'    .Fields.Item("HOURS").Precision = 3
'    .Fields.Item("HOURS").NumericScale = 1
'.Fields.Append "START_TIME", adChar, 5, l_lngCFIELDATTR
'.Fields.Append "END_TIME", adChar, 5, l_lngCFIELDATTR
'.Fields.Append "EXCEPTION_CD", adChar, 3, l_lngCFIELDATTR
'.Fields.Append "OVER_EXC_CD", adChar, 3, l_lngCFIELDATTR
'.Fields.Append "JOB_NUM", adChar, 6, l_lngCFIELDATTR
'.Fields.Append "OT_REASON_CD", adChar, 3, l_lngCFIELDATTR
'.Fields.Append "OT_OPP_DESC", adChar, 30, l_lngCFIELDATTR
'.Fields.Append "WG_NUM", adInteger, , l_lngCFIELDATTR
'.Fields.Append "CASE_NUM", adInteger, , l_lngCFIELDATTR
'.Fields.Append "FAMILY_LEAVE_IND", adChar, 1, l_lngCFIELDATTR
'.Fields.Append "COMP_SHIFT_CD", adChar, 3, l_lngCFIELDATTR
'.Fields.Append "COMP_CD", adChar, 3, l_lngCFIELDATTR
'.Fields.Append "REL_CD", adChar, 3, l_lngCFIELDATTR
'.Fields.Append "SCHED_CHG_EXC_CD", adChar, 3, l_lngCFIELDATTR
'.Fields.Append "DISTR_DEPT_NUM", adChar, 3, l_lngCFIELDATTR
'.Fields.Append "APPR_IND", adChar, 1, l_lngCFIELDATTR
'.Fields.Append "RECORD_STATUS", adChar, 1, adFldUpdatable
'.Fields.Append "SCHED_TIME_IND", adChar, 1, adFldUpdatable
'.Fields.Append "TIME_STAMP", adChar, 26, adFldUpdatable
''KEY INFORMATION
'.Fields.Append "KEY_EMP_NUM", adInteger, , l_lngCFIELDATTR
'.Fields.Append "KEY_CHARGE_DATE", adDate, , l_lngCFIELDATTR
'.Fields.Append "KEY_START_TIME", adChar, 5, l_lngCFIELDATTR
'.Fields.Append "KEY_INDEX_TIME", adChar, 26, l_lngCFIELDATTR


Private Sub SaveEMPRV( _
    ByVal v_colEMPRVTrans As Collection)

    Dim l_varEMPRVTrans As Variant
    Dim l_obj As dbMessageDispatcher.cdMessageDispatcher
    
    On Error GoTo eh

    l_varEMPRVTrans = ArrayFromCollection(v_colEMPRVTrans)
    Set l_obj = _
        CreateObject("dbMessageDispatcher.cdMessageDispatcher")
    l_obj.AddEMPRVTransactions l_varEMPRVTrans
    Set l_obj = Nothing
    Exit Sub
    
    
eh:
    Set l_obj = Nothing
    TrsRaiseError mstrCMODULENAME, "SaveEMPRV"

End Sub
Private Function SaveEMPRVJobTasks(v_rsEMPRVJobTasks As ADODB.Recordset, _
        r_rsErrors As ADODB.Recordset) As Boolean

    Dim l_obj As dbMessageDispatcher.cdMessageDispatcher
    
    On Error GoTo eh

    Set l_obj = _
        CreateObject("dbMessageDispatcher.cdMessageDispatcher")
    l_obj.AddEMPRVJobTasks v_rsEMPRVJobTasks
    Set l_obj = Nothing
    SaveEMPRVJobTasks = True
    Exit Function
    
eh:
    Set l_obj = Nothing
    SaveEMPRVJobTasks = False
    AppendError r_rsErrors, 0, Date, Date, "00:00", "00:00", 0, "", "", _
        trErrdbSaveEMPRVJobTasks, mstrCBUSINESSERR, _
        "EMPRV is unavailable therefore EMPRV Job Task Charges cannot be saved", _
        mstrCMODULENAME
    
End Function

Private Function GetLockStatusData( _
    ByVal v_aryEmployees As Variant _
    ) As ADODB.Recordset

    Dim l_obj As dbEmployee.cdEmployee

    On Error GoTo eh

    Set l_obj = CreateObject("dbEmployee.cdEmployee")
    Set GetLockStatusData = l_obj.GetEmployeeExcSchedInfo(v_aryEmployees)
    Set l_obj = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError mstrCMODULENAME, "GetLockStatusData"


End Function


Private Function GetCurrentCollPeriods() As ADODB.Recordset

    Dim l_obj As dbCollectionPeriod.cdCollectionPeriod


    On Error GoTo eh

    Set l_obj = CreateObject("dbCollectionPeriod.cdCollectionPeriod")

    Set GetCurrentCollPeriods = l_obj.GetCurrentCollectionPeriods()

    Set l_obj = Nothing

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetCurrentCollPeriods"

End Function


        
Private Function GetLockStatus( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As String

    Dim l_obj As busEmployee.cbsEmployee
    
    On Error GoTo eh

    Set l_obj = CreateObject("busEmployee.cbsEmployee")

    GetLockStatus = l_obj.GetLockStatus( _
        v_lngEmpNum, v_datDate, m_rsEmployeeData, m_rsCollPeriodData)

    Set l_obj = Nothing

    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "GetLockStatus"
    
End Function

Private Function AppendLockStatusBRV( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datChargeDate As Date, _
            ByVal v_strLockStatus As String, _
            ByVal v_strEntrySource As String, _
            ByVal v_rsRequest As ADODB.Recordset, _
            ByRef r_rsErrors As ADODB.Recordset _
            ) As Boolean
            
    Dim l_strResDescription As String
    Dim l_bolError As Boolean
    Dim l_strDept As String
    Dim l_lngErrorNum As Long
    
    On Error GoTo AppendLockStatusBRV_EH
    
    Select Case v_strLockStatus
        Case "Y"
            l_strResDescription = _
                TrsGetErrDescription(LongAppErr(trProcessedBypayroll), DescArgsFrom(CStr(v_datChargeDate), CStr(v_lngEmpNum)))
            l_lngErrorNum = trProcessedBypayroll
        Case "L"
            l_strDept = GetEmployeeDept(v_lngEmpNum, v_datChargeDate)
            l_strResDescription = _
                    TrsGetErrDescription(LongAppErr(trLockedBypayroll), DescArgsFrom(l_strDept, CStr(v_lngEmpNum)))
            l_lngErrorNum = trLockedBypayroll
    End Select
    
    Select Case v_strEntrySource
        Case "T" 'TimeEntries
            AppendEntryError r_rsErrors, v_rsRequest, l_lngErrorNum, mstrCBUSINESSERR, l_strResDescription, mstrCMODULENAME
        
        Case "O" 'Other Pay Entries
            AppendEntryErrorOtherPay r_rsErrors, v_rsRequest, l_lngErrorNum, mstrCBUSINESSERR, l_strResDescription, mstrCMODULENAME
        
    End Select

    AppendLockStatusBRV = l_bolError
    Exit Function
    
AppendLockStatusBRV_EH:
    TrsRaiseError mstrCMODULENAME, "AppendLockStatusBRV"

End Function

Private Function GetEmployeeDept( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date _
    ) As String
    
    Dim l_obj As busEmployee.cbsEmployee
    Dim l_strDept As String
    Dim l_strPrevDept As String
    Dim l_datDeptEffDate As Date
    
    On Error GoTo eh
    
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    If Not mADOExtensions.RecordsExist(m_rsEmployeeData) Then
        GetEmployeeDept = ""
        m_rsEmployeeData.Filter = ""
        Exit Function
    End If
    l_strDept = m_rsEmployeeData!DEPT_NUM
    l_strPrevDept = m_rsEmployeeData!PREV_DEPT_NUM
    l_datDeptEffDate = m_rsEmployeeData!DEPT_EFF_DATE
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    GetEmployeeDept = _
        l_obj.EmployeeDepartment(l_datDeptEffDate, l_strDept, _
            l_strPrevDept, v_datChargeDate)

    Set l_obj = Nothing
    
    m_rsEmployeeData.Filter = ""
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError mstrCMODULENAME, "GetEmployeeDept"

End Function

Private Function ValidLockStatus( _
    ByVal v_rsRequest As ADODB.Recordset, _
    ByVal v_strEntrySource As String, _
    ByRef r_objErrors As ADODB.Recordset _
    ) As Boolean
    
    Dim l_blnValidLockStatus As Boolean
    Dim l_strLockStatus As String
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    
    On Error GoTo eh
    
    If Not RecordsExist(v_rsRequest) Then
        ValidLockStatus = True
        Exit Function
    End If
    
    l_blnValidLockStatus = True
    v_rsRequest.MoveFirst
    
    While Not v_rsRequest.EOF
        l_lngEmpNum = v_rsRequest!KEY_EMP_NUM
        l_datChargeDate = v_rsRequest!KEY_CHARGE_DATE
        l_strLockStatus = GetLockStatus(l_lngEmpNum, l_datChargeDate)
        If l_strLockStatus <> "N" Then
            l_blnValidLockStatus = False
            AppendLockStatusBRV l_lngEmpNum, l_datChargeDate, _
                l_strLockStatus, v_strEntrySource, v_rsRequest, r_objErrors
        End If
        v_rsRequest.MoveNext
    Wend
    
    ValidLockStatus = l_blnValidLockStatus
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ValidLockStatus"
    
End Function
           
Private Function SaveMAXIMOJobTasks(v_rsMAXIMOJobTasks As ADODB.Recordset, _
        r_rsErrors As ADODB.Recordset) As Boolean

    Dim l_obj As dbMessageDispatcher.cdMessageDispatcher
    
    On Error GoTo eh

    Set l_obj = _
        CreateObject("dbMessageDispatcher.cdMessageDispatcher")
    l_obj.AddMAXIMOJobTasks v_rsMAXIMOJobTasks
    Set l_obj = Nothing
    SaveMAXIMOJobTasks = True
    Exit Function
    
eh:
    Set l_obj = Nothing
    SaveMAXIMOJobTasks = False

End Function
           
Private Sub Class_Initialize()
    m_blnIsPOETActive = IsPOETActive()
    m_blnIsOTTCEffective = IsOTTCEffective()
End Sub

'//copied
Private Function IsPOETActive() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim UseOTTC As String
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsPOETActive = False
    
    UseOTTC = UCase(dbEnv.GetSysParmsValue(50, 1))
    
    If UseOTTC = "Y" Or UseOTTC = "TRUE" Or UseOTTC = "1" Then
        IsPOETActive = True
    End If
        
    Set dbEnv = Nothing

    Exit Function

eh:
    Set dbEnv = Nothing
   
    TrsRaiseError mstrCMODULENAME, "IsPOETActive", , , , DescArgsFrom(UseOTTC), True
End Function

'//actual date to start using POET distributions after cutover
Private Function IsOTTCEffective() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim sEffDate As String
    Dim EffDateTime
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsOTTCEffective = False
    
    sEffDate = dbEnv.GetSysParmsValue(53, 1)
    
    EffDateTime = DateTimeFromDB2Timestamp(sEffDate)
    
    IsOTTCEffective = Now > EffDateTime
        
    Set dbEnv = Nothing

    Exit Function
eh:
    Set dbEnv = Nothing
   
    TrsRaiseError mstrCMODULENAME, "IsOTTCEffective", , , , DescArgsFrom(sEffDate), True
End Function
