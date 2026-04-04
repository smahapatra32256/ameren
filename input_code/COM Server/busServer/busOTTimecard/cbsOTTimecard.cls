VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsOTTimecard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum TrsPrivateTimecardErrorEnum
    trLockedBypayroll = 30103 + vbObjectError
    trProcessedBypayroll = 30104 + vbObjectError
End Enum
'
Dim m_rsEmployeeData As ADODB.Recordset
Dim m_rsCollPeriodData As ADODB.Recordset
Dim m_rsEmployees As ADODB.Recordset

Private Const mstrCMODULENAME = "cbsOTTimecard"


Public Function ProcessTimeEntries( _
        ByVal v_rsTimeEntriesRequest As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset, _
        ByVal v_UserId As String) As Boolean
' Answer True when all time entries requests have been processed.
'  If any time entry request fails return False.

    Dim l_objTimecard As busTimecard.cbsTimecard
    Dim l_objCtx As ObjectContext
    Dim l_blnResult As Boolean
    Dim l_objErrors As ADODB.Recordset
    Dim l_colEmployees As Collection
    Dim l_rsEmployees As ADODB.Recordset
    Dim l_rsEmpWorkgroups As ADODB.Recordset
    Dim l_rsSchedules As ADODB.Recordset
    Dim l_rsTimecardEntries As ADODB.Recordset
    Dim l_rsAccrualBankUpdates As ADODB.Recordset
    Dim l_rsOvertimeBankUpdates As ADODB.Recordset
    Dim l_colEMPRVInterface As Collection
    Dim l_strEmpNum As Variant
'    Dim l_objDistributionMgr As busDistributionMgr.cbsDistributionMgr
    Dim l_blnDistributionValidationSuccessful As Boolean
    Dim l_aryEmployees As Variant
    Dim l_rsEmployeeData As ADODB.Recordset
    Dim l_rsCollPeriodData As ADODB.Recordset
    Dim l_lngEmpNum As Long
    Dim l_lowDate As Date
    Dim l_highDate As Date
    Dim l_startDate As Date
    Dim l_endDate As Date
    Dim l_rsTimecardRequest As ADODB.Recordset


On Error GoTo eh
    
    Set l_objCtx = GetObjectContext
    Set l_objErrors = NewErrorset
    
    If Not CheckForInput(v_rsTimeEntriesRequest, Nothing, l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries = False
        GoTo xit
    End If
    
    If Not ValidateInput(v_rsTimeEntriesRequest, l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries = False
        GoTo xit
    End If
    
    Set l_colEmployees = GetEmployeeList(v_rsTimeEntriesRequest, Nothing)
    l_aryEmployees = ArrayFromCollection(l_colEmployees)

    l_lowDate = FindLowDate(v_rsTimeEntriesRequest)
    l_highDate = FindHighDate(v_rsTimeEntriesRequest)
    l_startDate = WorkWeekBeginDateFor(l_lowDate)
    l_endDate = WorkWeekEndDateFor(l_highDate)
    
    GetEmployeeState l_aryEmployees, l_startDate, l_endDate, False, _
        l_rsEmployees, l_rsEmpWorkgroups, l_rsTimecardEntries, Nothing, l_rsSchedules
    Set m_rsEmployees = l_rsEmployees
    
    CopyTimecardRequest v_rsTimeEntriesRequest, l_rsTimecardRequest
    
    If Not RecordsExist(l_rsTimecardRequest) Then
        l_objCtx.SetComplete
        ProcessTimeEntries = True
        GoTo xit
    End If
    
'    Set l_objDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
'    l_blnDistributionValidationSuccessful = _
'            l_objDistributionMgr.ValidateDistributionIdsFromTimeEntries( _
'                l_rsTimecardRequest, _
'                Nothing, _
'                l_rsEmployees, _
'                l_rsEmpWorkgroups, _
'                l_startDate, _
'                l_endDate, _
'                l_objErrors)
'
'    If Not l_blnDistributionValidationSuccessful Then
'        Set r_rsErrors = l_objErrors
'        ProcessTimeEntries = False
'        Exit Function
'    End If


    Set m_rsEmployeeData = GetLockStatusData(l_aryEmployees)
    Set m_rsCollPeriodData = GetCurrentCollPeriods()

    If Not ValidLockStatus(v_rsTimeEntriesRequest, "T", l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries = False
        GoTo xit
    End If

    For Each l_strEmpNum In l_colEmployees
        PrepareRequest CStr(l_strEmpNum), l_rsEmployees, l_rsEmpWorkgroups, _
            l_rsTimecardEntries, Nothing, l_rsSchedules, v_rsTimeEntriesRequest, Nothing
        Set l_objTimecard = CreateObject("busTimecard.cbsTimecard")
        
        l_lngEmpNum = CLng(l_strEmpNum)
        
        ProcessTimeEntries = _
            l_objTimecard.RecordEqualizedOvertime( _
                l_rsTimecardRequest, _
                l_rsEmployees, _
                l_rsEmpWorkgroups, _
                l_rsTimecardEntries, _
                l_rsSchedules, _
                l_startDate, _
                l_endDate, _
                l_objErrors, _
                v_UserId)
        If Not ProcessTimeEntries Then
            l_objCtx.SetAbort
            GoTo xit
        End If
    Next
    
'
'    UpdateAccrualBanks l_rsAccrualBankUpdates, Nothing, v_UserId
'
'    UpdateOvertimeBanks l_rsOvertimeBankUpdates, l_rsEmployees, Nothing, v_UserId
'
'    If Not l_colEMPRVInterface Is Nothing Then
'    If l_colEMPRVInterface.Count > 0 Then
'        SaveEMPRV l_colEMPRVInterface
'    End If

    l_objCtx.SetComplete

 
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
    Exit Function
       
eh:
    l_objCtx.SetAbort
    Set r_rsErrors = Nothing
    TrsRaiseError mstrCMODULENAME, "ProcessTimeEntries", , , , , True
   
End Function


Public Function ProcessTimeEntries3( _
        ByVal v_rsTimecardRequest As ADODB.Recordset, _
        ByVal v_rsOtherPayRequest As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset, _
        ByVal v_UserId As String) As Boolean
' Answer True when all time entries requests have been processed.
'  If any time entry request fails return False.

    Dim l_objTimecard As busTimecard.cbsTimecard
    Dim l_objCtx As ObjectContext
    Dim l_blnResult As Boolean
    Dim l_objErrors As ADODB.Recordset
    Dim l_colEmployees As Collection
    Dim l_rsEmployees As ADODB.Recordset
    Dim l_rsEmpWorkgroups As ADODB.Recordset
    Dim l_rsSchedules As ADODB.Recordset
    Dim l_rsTimecardEntries As ADODB.Recordset
    Dim l_rsOtherPayEntries As ADODB.Recordset
    Dim l_rsAccrualBankUpdates As ADODB.Recordset
    Dim l_rsOvertimeBankUpdates As ADODB.Recordset
    Dim l_colEMPRVInterface As Collection
    Dim l_blnOtherPayInput As Boolean
    Dim l_strEmpNum As Variant
    Dim l_objDistributionMgr As busDistributionMgr.cbsDistributionMgr
    Dim l_blnDistributionValidationSuccessful As Boolean
    Dim l_aryEmployees As Variant
    Dim l_rsEmployeeData As ADODB.Recordset
    Dim l_rsCollPeriodData As ADODB.Recordset
    Dim l_lngEmpNum As Long
    Dim l_lowDate As Date
    Dim l_highDate As Date
    Dim l_startDate As Date
    Dim l_endDate As Date
    Dim l_rsTimecardRequest As ADODB.Recordset


On Error GoTo eh
    
    Set l_objCtx = GetObjectContext
    Set l_objErrors = NewErrorset
    
    If Not CheckForInput(v_rsTimecardRequest, v_rsOtherPayRequest, l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries3 = False
        GoTo xit
    End If
    
    If Not ValidateInput(v_rsTimecardRequest, l_objErrors) Or _
            Not ValidateInputOtherPay(v_rsOtherPayRequest, l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries3 = False
        GoTo xit
    End If
    
    Set l_colEmployees = GetEmployeeList(v_rsTimecardRequest, v_rsOtherPayRequest)
    l_aryEmployees = ArrayFromCollection(l_colEmployees)

    l_lowDate = FindLowDate(v_rsTimecardRequest)
    l_highDate = FindHighDate(v_rsTimecardRequest)
    l_startDate = WorkWeekBeginDateFor(l_lowDate)
    l_endDate = WorkWeekEndDateFor(l_highDate)
    
    l_blnOtherPayInput = RecordsExist(v_rsOtherPayRequest)
    GetEmployeeState l_aryEmployees, l_startDate, l_endDate, l_blnOtherPayInput, _
        l_rsEmployees, l_rsEmpWorkgroups, l_rsTimecardEntries, _
        l_rsOtherPayEntries, l_rsSchedules
    Set m_rsEmployees = l_rsEmployees
    
    CopyTimecardRequest v_rsTimecardRequest, l_rsTimecardRequest
    
    If Not RecordsExist(l_rsTimecardRequest) Then
        l_objCtx.SetComplete
        ProcessTimeEntries3 = True
        GoTo xit
    Else
        l_rsTimecardRequest.MoveFirst
    End If
    
    If RecordsExist(v_rsOtherPayRequest) Then
        v_rsOtherPayRequest.MoveFirst
    End If
    Set l_objDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    l_blnDistributionValidationSuccessful = _
            l_objDistributionMgr.ValidateDistributionIdsFromTimeEntries( _
                l_rsTimecardRequest, _
                v_rsOtherPayRequest, _
                l_rsEmployees, _
                l_rsEmpWorkgroups, _
                l_startDate, _
                l_endDate, _
                l_objErrors)
    If Not l_blnDistributionValidationSuccessful Then
        ProcessTimeEntries3 = False
        l_objCtx.SetAbort
        GoTo xit
    Else
        l_rsTimecardRequest.Filter = adFilterNone
        v_rsOtherPayRequest.Filter = adFilterNone
    End If

    Set m_rsEmployeeData = GetLockStatusData(l_aryEmployees)
    Set m_rsCollPeriodData = GetCurrentCollPeriods()

    If Not ValidLockStatus(v_rsTimecardRequest, "T", l_objErrors) Or _
            Not ValidLockStatus(v_rsOtherPayRequest, "O", l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries3 = False
        GoTo xit
    End If

    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long

    For Each l_strEmpNum In l_colEmployees

        PrepareRequest CStr(l_strEmpNum), l_rsEmployees, l_rsEmpWorkgroups, _
            l_rsTimecardEntries, l_rsOtherPayEntries, l_rsSchedules, _
            l_rsTimecardRequest, v_rsOtherPayRequest
        Set l_objTimecard = CreateObject("busTimecard.cbsTimecard")
        
        l_lngEmpNum = CLng(l_strEmpNum)
        
        'Bypass RecordEqualizedOvertime2 when both request recordsets are empty.
        'Currently, this can occur since refused timecharges are sent to OTTimecard
        'and the empNum is included in array, but no timecard requests exist.
        
        If RecordsExist(l_rsTimecardRequest) Or _
                RecordsExist(v_rsOtherPayRequest) Then
                
            #If TRTRACE Then
                stTime = GetTickCount()
            #End If
                
            ProcessTimeEntries3 = _
                l_objTimecard.RecordEqualizedOvertime2( _
                    l_rsTimecardRequest, _
                    v_rsOtherPayRequest, _
                    l_rsEmployees, _
                    l_rsEmpWorkgroups, _
                    l_rsTimecardEntries, _
                    l_rsOtherPayEntries, _
                    l_rsSchedules, _
                    l_startDate, _
                    l_endDate, _
                    l_objErrors, _
                    v_UserId)
                    
            #If TRTRACE Then
                ndTime = GetTickCount()
                dfTime = ndTime - stTime
                TraceAction App.ThreadID, "3", "ProcessTimeEntries3", Format(Now, "mm/dd/yyyy hh:mm:ss"), _
                    CStr(dfTime / 1000)
            #End If
                    
            If Not ProcessTimeEntries3 Then
                l_objCtx.SetAbort
                GoTo xit
            End If
        End If
    Next
    
'
'    UpdateAccrualBanks l_rsAccrualBankUpdates, Nothing, v_UserId
'
'    UpdateOvertimeBanks l_rsOvertimeBankUpdates, l_rsEmployees, Nothing, v_UserId
'
'    If Not l_colEMPRVInterface Is Nothing Then
'    If l_colEMPRVInterface.Count > 0 Then
'        SaveEMPRV l_colEMPRVInterface
'    End If

    l_objCtx.SetComplete

 
xit:
    AppendTimecardErrors l_objErrors, r_rsErrors
    Set l_objTimecard = Nothing
    Exit Function
 
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
'    Exit Function
       
eh:
    l_objCtx.SetAbort
    Set r_rsErrors = Nothing
    TrsRaiseError mstrCMODULENAME, "ProcessTimeEntries3", , , , , True
   
End Function



Public Function ProcessTimeEntries2( _
                    ByVal v_rsTimeChargeRequests As ADODB.Recordset, _
                    ByVal v_blnGenDeletes As Boolean, _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByVal v_UserId As String) As Boolean
' Answer True when all time entries requests have been processed.
'  If any time entry request fails return False.
'
' The following arguments are required when v_blnGenDeletes is True:
'   v_lngOTOppNum, v_aryEmployees, v_datStartDate, v_datEndDate
'

    Dim l_objTimecard As busTimecard.cbsTimecard
    Dim l_objCtx As ObjectContext
    Dim l_blnResult As Boolean
    Dim l_objErrors As ADODB.Recordset
    Dim l_colEmployees As Collection
    Dim l_rsEmployees As ADODB.Recordset
    Dim l_rsEmpWorkgroups As ADODB.Recordset
    Dim l_rsSchedules As ADODB.Recordset
    Dim l_rsTimecardEntries As ADODB.Recordset
    Dim l_rsAccrualBankUpdates As ADODB.Recordset
    Dim l_rsOvertimeBankUpdates As ADODB.Recordset
    Dim l_colEMPRVInterface As Collection
    Dim l_strEmpNum As String
'    Dim l_objDistributionMgr As busDistributionMgr.cbsDistributionMgr
    Dim l_blnDistributionValidationSuccessful As Boolean
    Dim l_aryEmployees As Variant
    Dim l_rsEmployeeData As ADODB.Recordset
    Dim l_rsCollPeriodData As ADODB.Recordset
'    Dim l_lngEmpNum As Long
    Dim l_lowDate As Date
    Dim l_highDate As Date
    Dim l_startDate As Date
    Dim l_endDate As Date
    Dim l_rsTimecardDeleteRequest As ADODB.Recordset
    Dim i As Integer

On Error GoTo eh
    
    Set l_objCtx = GetObjectContext
    Set l_objErrors = NewErrorset
    
    If Not CheckForInput(v_rsTimeChargeRequests, Nothing, l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries2 = False
        GoTo xit
    End If
    
    If Not ValidateInput(v_rsTimeChargeRequests, l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries2 = False
        GoTo xit
    End If
    
    'RSG - I am not sure union is necessary?  Will v_aryEmployees be full list?
    Set l_colEmployees = GetEmployeeList(v_rsTimeChargeRequests, Nothing)
    If v_blnGenDeletes Then
        Set l_colEmployees = UnionEmployeeList(l_colEmployees, v_aryEmployees)
    End If
    l_aryEmployees = ArrayFromCollection(l_colEmployees)

    l_lowDate = FindLowDate(v_rsTimeChargeRequests)
    l_highDate = FindHighDate(v_rsTimeChargeRequests)
    If v_blnGenDeletes Then
        l_lowDate = MinDate(l_lowDate, v_datStartDate)
        l_highDate = MaxDate(l_highDate, v_datEndDate)
    End If
    l_startDate = WorkWeekBeginDateFor(l_lowDate)
    l_endDate = WorkWeekEndDateFor(l_highDate)
    
    GetEmployeeState l_aryEmployees, l_startDate, l_endDate, False, _
        l_rsEmployees, l_rsEmpWorkgroups, l_rsTimecardEntries, Nothing, l_rsSchedules
    Set m_rsEmployees = l_rsEmployees
    
    If v_blnGenDeletes Then
        Set l_rsTimecardDeleteRequest = _
            GenerateDeleteRequests(v_lngOTOppNum, v_aryEmployees, l_rsTimecardEntries)
        AppendDeleteRequests v_lngOTOppNum, l_rsTimecardDeleteRequest, v_rsTimeChargeRequests
    End If

    
'    Set l_objDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
'    l_blnDistributionValidationSuccessful = _
'            l_objDistributionMgr.ValidateDistributionIdsFromTimeEntries( _
'                l_rsTimecardDeleteRequest, _
'                Nothing, _
'                l_rsEmployees, _
'                l_rsEmpWorkgroups, _
'                l_startDate, _
'                l_endDate, _
'                l_objErrors)
'
'    If Not l_blnDistributionValidationSuccessful Then
'        Set r_rsErrors = l_objErrors
'        ProcessTimeEntries2 = False
'        Exit Function
'    End If


    Set m_rsEmployeeData = GetLockStatusData(l_aryEmployees)
    Set m_rsCollPeriodData = GetCurrentCollPeriods()

    If Not ValidLockStatus(v_rsTimeChargeRequests, "T", l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries2 = False
        GoTo xit
    End If

    For i = LBound(l_aryEmployees) To UBound(l_aryEmployees)
        l_strEmpNum = CStr(l_aryEmployees(i))
        PrepareRequest CStr(l_strEmpNum), l_rsEmployees, l_rsEmpWorkgroups, _
            l_rsTimecardEntries, Nothing, l_rsSchedules, v_rsTimeChargeRequests, Nothing
        
        If RecordsExist(v_rsTimeChargeRequests) Then
            Set l_objTimecard = CreateObject("busTimecard.cbsTimecard")
            ProcessTimeEntries2 = _
                l_objTimecard.RecordEqualizedOvertime( _
                    v_rsTimeChargeRequests, _
                    l_rsEmployees, _
                    l_rsEmpWorkgroups, _
                    l_rsTimecardEntries, _
                    l_rsSchedules, _
                    l_startDate, _
                    l_endDate, _
                    l_objErrors, _
                    v_UserId)
            If Not ProcessTimeEntries2 Then
                l_objCtx.SetAbort
                GoTo xit
            End If
        End If
    Next
    
'
'    UpdateAccrualBanks l_rsAccrualBankUpdates, Nothing, v_UserId
'
'    UpdateOvertimeBanks l_rsOvertimeBankUpdates, l_rsEmployees, Nothing, v_UserId
'
'    If Not l_colEMPRVInterface Is Nothing Then
'    If l_colEMPRVInterface.Count > 0 Then
'        SaveEMPRV l_colEMPRVInterface
'    End If

    l_objCtx.SetComplete

 
xit:
    AppendTimecardErrors l_objErrors, r_rsErrors
    Set l_objTimecard = Nothing
    Exit Function
       
eh:
    l_objCtx.SetAbort
    Set r_rsErrors = Nothing
    TrsRaiseError mstrCMODULENAME, "ProcessTimeEntries2", , , , , True
   
End Function


Public Function ProcessTimeEntries4( _
                    ByVal v_rsTimeChargeRequests As ADODB.Recordset, _
                    ByVal v_rsOtherPayRequests As ADODB.Recordset, _
                    ByVal v_blnGenDeletes As Boolean, _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByVal v_UserId As String) As Boolean
' Answer True when all time entries requests have been processed.
'  If any time entry request fails return False.
'
' The following arguments are required when v_blnGenDeletes is True:
'   v_lngOTOppNum, v_aryEmployees, v_datStartDate, v_datEndDate
'

    Dim l_objTimecard As busTimecard.cbsTimecard
    Dim l_objCtx As ObjectContext
    Dim l_blnResult As Boolean
    Dim l_objErrors As ADODB.Recordset
    Dim l_colEmployees As Collection
    Dim l_rsEmployees As ADODB.Recordset
    Dim l_rsEmpWorkgroups As ADODB.Recordset
    Dim l_rsSchedules As ADODB.Recordset
    Dim l_rsTimecardEntries As ADODB.Recordset
    Dim l_rsOtherPayEntries As ADODB.Recordset
    Dim l_rsAccrualBankUpdates As ADODB.Recordset
    Dim l_rsOvertimeBankUpdates As ADODB.Recordset
    Dim l_colEMPRVInterface As Collection
    Dim l_strEmpNum As String
    Dim l_objDistributionMgr As busDistributionMgr.cbsDistributionMgr
    Dim l_blnDistributionValidationSuccessful As Boolean
    Dim l_aryEmployees As Variant
    Dim l_rsEmployeeData As ADODB.Recordset
    Dim l_rsCollPeriodData As ADODB.Recordset
'    Dim l_lngEmpNum As Long
    Dim l_lowDate As Date
    Dim l_highDate As Date
    Dim l_startDate As Date
    Dim l_endDate As Date
    Dim l_rsTimecardDeleteRequest As ADODB.Recordset
    Dim l_rsOtherPayDeleteRequest As ADODB.Recordset
    Dim i As Integer

On Error GoTo eh
    
    Set l_objCtx = GetObjectContext
    Set l_objErrors = NewErrorset
    
    If Not CheckForInput(v_rsTimeChargeRequests, v_rsOtherPayRequests, l_objErrors) _
                And Not v_blnGenDeletes Then
        l_objCtx.SetAbort
        ProcessTimeEntries4 = False
        GoTo xit
    End If
    
    If Not ValidateInput(v_rsTimeChargeRequests, l_objErrors) Or _
            Not ValidateInputOtherPay(v_rsOtherPayRequests, l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries4 = False
        GoTo xit
    End If
    
' Union list for deletes
    Set l_colEmployees = GetEmployeeList(v_rsTimeChargeRequests, v_rsOtherPayRequests)
    If v_blnGenDeletes Then
        Set l_colEmployees = UnionEmployeeList(l_colEmployees, v_aryEmployees)
    End If
    l_aryEmployees = ArrayFromCollection(l_colEmployees)

    l_lowDate = FindLowDate(v_rsTimeChargeRequests)
    l_highDate = FindHighDate(v_rsTimeChargeRequests)
    If v_blnGenDeletes Then
        l_lowDate = MinDate(l_lowDate, v_datStartDate)
        l_highDate = MaxDate(l_highDate, v_datEndDate)
    End If
    l_startDate = WorkWeekBeginDateFor(l_lowDate)
    l_endDate = WorkWeekEndDateFor(l_highDate)
    
    GetEmployeeState l_aryEmployees, l_startDate, l_endDate, True, _
        l_rsEmployees, l_rsEmpWorkgroups, l_rsTimecardEntries, l_rsOtherPayEntries, l_rsSchedules
    Set m_rsEmployees = l_rsEmployees
    
    If v_blnGenDeletes Then
        Set l_rsTimecardDeleteRequest = _
            GenerateDeleteRequests(v_lngOTOppNum, v_aryEmployees, l_rsTimecardEntries)
        AppendDeleteRequests v_lngOTOppNum, l_rsTimecardDeleteRequest, v_rsTimeChargeRequests
        Set l_rsOtherPayDeleteRequest = _
            GenerateOtherPayDeleteRequests(v_lngOTOppNum, v_aryEmployees, l_rsOtherPayEntries)
        AppendOtherPayDeleteRequests v_lngOTOppNum, l_rsOtherPayDeleteRequest, v_rsOtherPayRequests
    End If
    
    If RecordsExist(v_rsTimeChargeRequests) Then
        v_rsTimeChargeRequests.MoveFirst
    End If
    If RecordsExist(v_rsOtherPayRequests) Then
        v_rsOtherPayRequests.MoveFirst
    End If
    Set l_objDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    l_blnDistributionValidationSuccessful = _
            l_objDistributionMgr.ValidateDistributionIdsFromTimeEntries( _
                v_rsTimeChargeRequests, _
                v_rsOtherPayRequests, _
                l_rsEmployees, _
                l_rsEmpWorkgroups, _
                l_startDate, _
                l_endDate, _
                l_objErrors)
    If Not l_blnDistributionValidationSuccessful Then
        ProcessTimeEntries4 = False
        l_objCtx.SetAbort
        GoTo xit
    Else
        v_rsTimeChargeRequests.Filter = adFilterNone
        v_rsOtherPayRequests.Filter = adFilterNone
    End If


    Set m_rsEmployeeData = GetLockStatusData(l_aryEmployees)
    Set m_rsCollPeriodData = GetCurrentCollPeriods()

    If Not ValidLockStatus(v_rsTimeChargeRequests, "T", l_objErrors) Or _
            Not ValidLockStatus(v_rsOtherPayRequests, "O", l_objErrors) Then
        l_objCtx.SetAbort
        ProcessTimeEntries4 = False
        GoTo xit
    End If

    For i = LBound(l_aryEmployees) To UBound(l_aryEmployees)
        l_strEmpNum = CStr(l_aryEmployees(i))
        PrepareRequest CStr(l_strEmpNum), l_rsEmployees, l_rsEmpWorkgroups, _
            l_rsTimecardEntries, l_rsOtherPayEntries, l_rsSchedules, _
            v_rsTimeChargeRequests, v_rsOtherPayRequests
        
        If RecordsExist(v_rsTimeChargeRequests) Then
            Set l_objTimecard = CreateObject("busTimecard.cbsTimecard")
            ProcessTimeEntries4 = _
                l_objTimecard.RecordEqualizedOvertime2( _
                    v_rsTimeChargeRequests, _
                    v_rsOtherPayRequests, _
                    l_rsEmployees, _
                    l_rsEmpWorkgroups, _
                    l_rsTimecardEntries, _
                    l_rsOtherPayEntries, _
                    l_rsSchedules, _
                    l_startDate, _
                    l_endDate, _
                    l_objErrors, _
                    v_UserId)
                    
            If Not ProcessTimeEntries4 Then
                l_objCtx.SetAbort
                GoTo xit
            End If
        Else
            ProcessTimeEntries4 = True
        End If
    Next

    l_objCtx.SetComplete

 
xit:
    AppendTimecardErrors l_objErrors, r_rsErrors
    Set l_objTimecard = Nothing
    Exit Function
       
eh:
    l_objCtx.SetAbort
    Set r_rsErrors = Nothing
    TrsRaiseError mstrCMODULENAME, "ProcessTimeEntries4", , , , , True
   
End Function



Public Function DeleteTimeEntries( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByVal v_UserId As String) As Boolean
' Answer True when all time entries requests have been processed.
'  If any time entry request fails return False.

    Dim l_objTimecard As busTimecard.cbsTimecard
    Dim l_objCtx As ObjectContext
    Dim l_blnResult As Boolean
    Dim l_objErrors As ADODB.Recordset
'    Dim l_colEmployees As Collection
    Dim l_rsEmployees As ADODB.Recordset
    Dim l_rsEmpWorkgroups As ADODB.Recordset
    Dim l_rsSchedules As ADODB.Recordset
    Dim l_rsTimecardEntries As ADODB.Recordset
    Dim l_rsOtherPayEntries As ADODB.Recordset
    Dim l_rsAccrualBankUpdates As ADODB.Recordset
    Dim l_rsOvertimeBankUpdates As ADODB.Recordset
    Dim l_colEMPRVInterface As Collection
    Dim l_strEmpNum As String
    Dim l_blnDistributionValidationSuccessful As Boolean
'    Dim l_aryEmployees As Variant
    Dim l_rsEmployeeData As ADODB.Recordset
    Dim l_rsCollPeriodData As ADODB.Recordset
'    Dim l_lngEmpNum As Long
    Dim l_lowDate As Date
    Dim l_highDate As Date
    Dim l_startDate As Date
    Dim l_endDate As Date
    Dim l_rsTimecardRequest As ADODB.Recordset
    Dim l_rsOtherPayRequest As ADODB.Recordset
    Dim i As Integer


On Error GoTo eh
    
    Set l_objCtx = GetObjectContext
    Set l_objErrors = NewErrorset
    
'    If Not CheckForInput(v_rsTimeEntriesRequest, l_objErrors) Then
'        l_objCtx.SetAbort
'        DeleteTimeEntries = False
'        GoTo xit
'    End If
'
'    If Not ValidateInput(v_rsTimeEntriesRequest, l_objErrors) Then
'        l_objCtx.SetAbort
'        DeleteTimeEntries = False
'        GoTo xit
'    End If
    
'    Set l_colEmployees = CollectionFromArray(v_aryEmployees, 0)

'    l_lowDate = FindLowDate(v_rsTimeEntriesRequest)
'    l_highDate = FindHighDate(v_rsTimeEntriesRequest)
    l_startDate = WorkWeekBeginDateFor(v_datStartDate)
    l_endDate = WorkWeekEndDateFor(v_datEndDate)
    
    GetEmployeeState v_aryEmployees, l_startDate, l_endDate, True, _
        l_rsEmployees, l_rsEmpWorkgroups, l_rsTimecardEntries, l_rsOtherPayEntries, l_rsSchedules
    Set m_rsEmployees = l_rsEmployees
    
    CreateDeleteTimecardRequest v_lngOTOppNum, l_rsTimecardEntries, l_rsTimecardRequest
    CreateDeleteOtherPayRequest v_lngOTOppNum, l_rsOtherPayEntries, l_rsOtherPayRequest
    
    If Not RecordsExist(l_rsTimecardRequest) Then
        l_objCtx.SetComplete
        DeleteTimeEntries = True
        GoTo xit
    End If
    
    Set m_rsEmployeeData = GetLockStatusData(v_aryEmployees)
    Set m_rsCollPeriodData = GetCurrentCollPeriods()

    If Not ValidLockStatus(l_rsTimecardRequest, "T", l_objErrors) Or _
            Not ValidLockStatus(l_rsOtherPayRequest, "O", l_objErrors) Then
        l_objCtx.SetAbort
        DeleteTimeEntries = False
        GoTo xit
    End If

    For i = 0 To UBound(v_aryEmployees)
        l_strEmpNum = CStr(v_aryEmployees(i))
        PrepareRequest l_strEmpNum, l_rsEmployees, l_rsEmpWorkgroups, _
            l_rsTimecardEntries, l_rsOtherPayEntries, l_rsSchedules, _
            l_rsTimecardRequest, l_rsOtherPayRequest
                        
        If RecordsExist(l_rsTimecardRequest) Or _
                RecordsExist(l_rsOtherPayRequest) Then
            Set l_objTimecard = CreateObject("busTimecard.cbsTimecard")
            DeleteTimeEntries = _
                l_objTimecard.RecordEqualizedOvertime2( _
                    l_rsTimecardRequest, _
                    l_rsOtherPayRequest, _
                    l_rsEmployees, _
                    l_rsEmpWorkgroups, _
                    l_rsTimecardEntries, _
                    l_rsOtherPayEntries, _
                    l_rsSchedules, _
                    l_startDate, _
                    l_endDate, _
                    l_objErrors, _
                    v_UserId)
            If Not DeleteTimeEntries Then
                l_objCtx.SetAbort
                GoTo xit
            End If
        End If
    Next

    l_objCtx.SetComplete

 
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
    Exit Function
       
eh:
    l_objCtx.SetAbort
    Set r_rsErrors = Nothing
    TrsRaiseError mstrCMODULENAME, "DeleteTimeEntries", , , , , True
   
End Function


Private Function CheckForInput( _
                    ByRef r_rsTimeEntriesRequest As ADODB.Recordset, _
                    ByRef r_rsOtherPayRequest As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
    
On Error GoTo eh
'    If Not RecordsExist(r_rsTimeEntriesRequest) Then
'        AppendError _
'            r_rsErrors, 0, Date, Date, "00:00", "00:00", 0, "", "", _
'            trErrbusValidationError, mstrCBUSINESSERR, _
'            "Time entry requests are empty", mstrCMODULENAME
'        Exit Function
'    End If
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
    If Not RecordsExist(r_rsTimeEntriesRequest) Then
        ValidateInput = True
        Exit Function
    End If

    r_rsTimeEntriesRequest.MoveFirst
    
'RSG - needs work
    
'    Do Until r_rsTimeEntriesRequest.EOF
'        Select Case r_rsTimeEntriesRequest!RECORD_STATUS
'            Case "I"
'                If Not ValidateNewTimeEntry(r_rsTimeEntriesRequest, r_rsErrors) Then
'                    l_blnValidationError = True
'                End If
'            Case "C"
'                If Not ValidateChangeTimeEntry(r_rsTimeEntriesRequest, r_rsErrors) Then
'                    l_blnValidationError = True
'                End If
'            Case "D"
'                If Not ValidateDeleteTimeEntry(r_rsTimeEntriesRequest, r_rsErrors) Then
'                    l_blnValidationError = True
'                End If
'            Case Else
'                AppendEntryError r_rsErrors, r_rsTimeEntriesRequest, trErrbusValidationError, mstrCBUSINESSERR, "Record status must be I, C, or D"
'                l_blnValidationError = True
'        End Select
'        r_rsTimeEntriesRequest.MoveNext
'    Loop
    
    'test only
    l_blnValidationError = False
    'test only
    
    ValidateInput = Not l_blnValidationError
    r_rsTimeEntriesRequest.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ValidateInput"

End Function

Private Function ValidateInputOtherPay( _
                    ByRef r_rsOtherPayRequest As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                    ) As Boolean
    
    Dim l_blnValidationError As Boolean
    
On Error GoTo eh

'RSG - needs work

    ' If there are no records then bypass Other Pay
    If Not RecordsExist(r_rsOtherPayRequest) Then
        ValidateInputOtherPay = True
        Exit Function
    End If
    
    r_rsOtherPayRequest.MoveFirst
    
    ' Each type of request requires different validations
'    Do Until r_rsOtherpayRequest.EOF
'        Select Case UCase(r_rsOtherpayRequest!RECORD_STATUS)
'            Case "I"
'                If Not ValidateInputNewOtherPay(r_rsOtherpayRequest, r_rsErrors) Then
'                    l_blnValidationError = True
'                End If
'            Case "D"
'                If Not ValidateInputDeleteOtherPay(r_rsOtherpayRequest, r_rsErrors) Then
'                    l_blnValidationError = True
'                End If
'            Case Else
'                AppendEntryErrorOtherPay r_rsErrors, r_rsOtherpayRequest, trErrbusValidationError, mstrCBUSINESSERR, "Record status must be I, C, or D"
'                l_blnValidationError = True
'        End Select
'        r_rsOtherpayRequest.MoveNext
'    Loop

    'test only
    l_blnValidationError = False
    'test only

    ValidateInputOtherPay = Not l_blnValidationError
    r_rsOtherPayRequest.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ValidateInputOtherPay"

End Function


Private Function GetEmployeeList( _
                    ByRef r_rsTimeEntriesRequest As ADODB.Recordset, _
                    ByRef r_rsOtherPay As ADODB.Recordset) As Collection
    Dim l_objEmpList As cuSet

On Error GoTo eh
    Set l_objEmpList = New cuSet
    If RecordsExist(r_rsTimeEntriesRequest) Then
        r_rsTimeEntriesRequest.MoveFirst
        Do Until r_rsTimeEntriesRequest.EOF
'            If Not IsNull(r_rsTimeEntriesRequest!KEY_EMP_NUM) Then
'                l_objEmpList.Add r_rsTimeEntriesRequest.Fields("KEY_EMP_NUM").Value
'            ElseIf IsNull(r_rsTimeEntriesRequest!EMP_NUM) Then
'                l_objEmpList.Add r_rsTimeEntriesRequest.Fields("EMP_NUM").Value
'            End If
            l_objEmpList.Add r_rsTimeEntriesRequest.Fields("EMP_NUM").Value

            r_rsTimeEntriesRequest.MoveNext
        Loop
        r_rsTimeEntriesRequest.MoveFirst
    End If
    
    If RecordsExist(r_rsOtherPay) Then
        r_rsOtherPay.MoveFirst
        Do Until r_rsOtherPay.EOF
            If Not IsNull(r_rsOtherPay!KEY_EMP_NUM) Then
                l_objEmpList.Add r_rsOtherPay.Fields("EMP_NUM").Value
            End If
            r_rsOtherPay.MoveNext
        Loop
        r_rsOtherPay.MoveFirst
    End If
    
    Set GetEmployeeList = l_objEmpList.Contents
    Exit Function
    
eh:
    Set l_objEmpList = Nothing
    TrsRaiseError mstrCMODULENAME, "GetEmployeeList"

End Function


Private Function UnionEmployeeList( _
                    ByVal v_colSource As Collection, _
                    ByRef r_aryEmployees As Variant) As Collection
    Dim l_objEmpList As cuSet
    Dim i As Integer

On Error GoTo eh

    Set l_objEmpList = New cuSet
    For i = 1 To v_colSource.Count
        l_objEmpList.Add v_colSource.Item(i)
    Next
    
    If ArraySize(r_aryEmployees) > 0 Then
        For i = LBound(r_aryEmployees) To UBound(r_aryEmployees)
            l_objEmpList.Add r_aryEmployees(i)
        Next
    End If
    
    Set UnionEmployeeList = l_objEmpList.Contents
    Exit Function
    
eh:
    Set l_objEmpList = Nothing
    TrsRaiseError mstrCMODULENAME, "UnionEmployeeList"

End Function


Private Function GetEmployeeState( _
                    ByVal l_aryEmployees As Variant, _
                    ByVal startDt As Date, _
                    ByVal endDt As Date, _
                    ByVal v_blnOtherPayReq As Boolean, _
                    ByRef l_rsEmployees As ADODB.Recordset, _
                    ByRef l_rsEmpWorkgroups As ADODB.Recordset, _
                    ByRef l_rsTimecardEntries As ADODB.Recordset, _
                    ByRef l_rsOtherPayEntries As ADODB.Recordset, _
                    ByRef l_rsSchedules As ADODB.Recordset)
'    Dim l_aryEmployees As Variant

    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long

On Error GoTo eh

    #If TRTRACE Then
        stTime = GetTickCount()
    #End If

'    l_aryEmployees = ArrayFromCollection(l_colEmployees)
    
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

    Set l_rsSchedules = _
        GetSchedules(l_aryEmployees, startDt, endDt)

    If v_blnOtherPayReq Then
        Set l_rsOtherPayEntries = _
            GetOtherPayEntries(l_aryEmployees, startDt, endDt)
    End If

    #If TRTRACE Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        TraceAction App.ThreadID, "3", "GetEmployeeState", Format(Now, "mm/dd/yyyy hh:mm:ss"), _
            CStr(dfTime / 1000)
    #End If

    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "GetEmployeeState"

End Function


Private Function GetEmployees( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal startDt As Date, _
                    ByVal endDt As Date) _
                As ADODB.Recordset
    Dim ocdEmployee As dbEmployee.cdEmployee

On Error GoTo eh
    
    Set ocdEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployees = ocdEmployee.GetEmployee(v_aryEmployees)
    Set ocdEmployee = Nothing
    Exit Function
    
eh:
    Set ocdEmployee = Nothing
    TrsRaiseError mstrCMODULENAME, "GetEmployees"
    
End Function


Private Function GetEmpWorkgroups( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal startDt As Date, _
                    ByVal endDt As Date) _
                As ADODB.Recordset
    Dim ocdWorkGroup As dbWorkGroup.cdWorkGroup

On Error GoTo eh
    
    Set ocdWorkGroup = CreateObject("dbWorkgroup.cdWorkgroup")
    Set GetEmpWorkgroups = _
        ocdWorkGroup.GetEmployeeWorkGroups(v_aryEmployees, startDt, endDt, True)
    Set ocdWorkGroup = Nothing
    Exit Function
    
eh:
    Set ocdWorkGroup = Nothing
    TrsRaiseError mstrCMODULENAME, "GetEmpWorkgroups"
    
End Function


Private Function GetTimecardEntries( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal startDt As Date, _
                    ByVal endDt As Date) _
                As ADODB.Recordset
    Dim ocdTimeCard As dbTimeCard.cdTimeCard

On Error GoTo eh
    
    Set ocdTimeCard = CreateObject("dbTimeCard.cdTimeCard")
    Set GetTimecardEntries = _
        ocdTimeCard.Get_TimeCard(v_aryEmployees, startDt, endDt)
    Set ocdTimeCard = Nothing
    Exit Function
    
eh:
    Set ocdTimeCard = Nothing
    TrsRaiseError mstrCMODULENAME, "GetTimecardEntries"
    
End Function


Private Function GetOtherPayEntries( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal startDt As Date, _
                    ByVal endDt As Date) As ADODB.Recordset
    Dim ocdTimeCard As dbTimeCard.cdTimeCard

On Error GoTo eh
    
    Set ocdTimeCard = _
        CreateObject("dbTimeCard.cdTimeCard")
    Set GetOtherPayEntries = _
        ocdTimeCard.Get_OtherPay(v_aryEmployees, startDt, endDt)
    Set ocdTimeCard = Nothing
    Exit Function
    
eh:
    Set ocdTimeCard = Nothing
    TrsRaiseError mstrCMODULENAME, "GetOtherPayEntries"
    
End Function


Private Function GetSchedules( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal startDt As Date, _
                    ByVal endDt As Date) _
                As ADODB.Recordset
    Dim obsSchedule As busSchedule.cbsSchedule

On Error GoTo eh
        
    Set obsSchedule = CreateObject("busSchedule.cbsSchedule")
    Set GetSchedules = _
        obsSchedule.GetAdjustedBaseSchedule(v_aryEmployees, startDt, endDt, Nothing, , trAdjBaseSchedWithExceptions)
    Set obsSchedule = Nothing
    Exit Function
    
eh:
    Set obsSchedule = Nothing
    TrsRaiseError mstrCMODULENAME, "GetSchedules"
    
End Function


Private Function FindLowDate( _
                    ByVal v_rsTimeentryRequest As ADODB.Recordset) As Date
                    
    Dim l_datDate As Date

On Error GoTo eh
    
    l_datDate = #12/31/9999#
    If Not RecordsExist(v_rsTimeentryRequest) Then
        GoTo xit
    End If
    
    v_rsTimeentryRequest.MoveFirst
    Do Until v_rsTimeentryRequest.EOF
        l_datDate = MinDate(l_datDate, v_rsTimeentryRequest!CHARGE_DATE)
        v_rsTimeentryRequest.MoveNext
    Loop
    
    v_rsTimeentryRequest.MoveFirst
    
xit:
    FindLowDate = l_datDate
    Exit Function
    
eh:
    
    TrsRaiseError mstrCMODULENAME, "FindLowDate"
        
End Function


Private Function FindHighDate( _
                    ByVal v_rsTimeentryRequest As ADODB.Recordset) As Date
                    
    Dim l_datDate As Date

On Error GoTo eh
    
    l_datDate = #1/1/1901#
    If Not RecordsExist(v_rsTimeentryRequest) Then
        GoTo xit
    End If
    
    v_rsTimeentryRequest.MoveFirst
    Do Until v_rsTimeentryRequest.EOF
        l_datDate = MaxDate(l_datDate, v_rsTimeentryRequest!CHARGE_DATE)
        v_rsTimeentryRequest.MoveNext
    Loop
    
    v_rsTimeentryRequest.MoveFirst
    
xit:
    FindHighDate = l_datDate
    Exit Function
    
eh:
    
    TrsRaiseError mstrCMODULENAME, "FindHighDate"
        
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
        l_lngEmpNum = v_rsRequest!EMP_NUM
        l_datChargeDate = v_rsRequest!CHARGE_DATE
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

'    If IsNull(r_rsTimeentryRequest!EMP_NUM) Then
'        l_varEmp_Num = r_rsTimeentryRequest!KEY_EMP_NUM
'    Else
        l_varEmp_Num = r_rsTimeentryRequest!EMP_NUM
'    End If
    
'    If IsNull(r_rsTimeentryRequest!CHARGE_DATE) Then
'        l_varChargeDate = r_rsTimeentryRequest!KEY_CHARGE_DATE
'    Else
        l_varChargeDate = r_rsTimeentryRequest!CHARGE_DATE
'    End If
        
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


Private Sub CopyTimecardRequest( _
                ByRef r_rsOvertimeEntries As ADODB.Recordset, _
                ByRef r_rsTimecardEntries As ADODB.Recordset)
'Filter out non paid overtime charges.

    Dim l_objEmployee As busEmployee.cbsEmployee
    Dim l_objCodeTables As dbCodeTables.cdCodeTables
    Dim l_strUnionCd As String
    Dim l_strExecConInd As String
    Dim l_strEarningsCode As String

On Error GoTo eh

    Set l_objEmployee = CreateObject("busEmployee.cbsEmployee")
    Set l_objCodeTables = CreateObject("dbCodeTables.cdCodeTables")

    Set r_rsTimecardEntries = CreateTimecardRequest
    r_rsOvertimeEntries.MoveFirst
    Do Until r_rsOvertimeEntries.EOF
        With r_rsOvertimeEntries
            m_rsEmployees.Filter = "EMP_NUM = " & !EMP_NUM
            l_strUnionCd = _
                l_objEmployee.EmployeeUnion( _
                    m_rsEmployees!JOB_EFF_DATE, _
                    m_rsEmployees!UNION_CD, _
                    m_rsEmployees!PRIOR_UNION_CD, _
                    !CHARGE_DATE)
            l_strExecConInd = _
                l_objCodeTables.UnionExecutiveContractIndicator(l_strUnionCd)
            l_strEarningsCode = _
                l_objCodeTables.ExceptionEarningsCode( _
                    !EXCEPTION_CD, _
                    l_strExecConInd)
            If Trim(l_strEarningsCode) <> "99" Or Trim(!OVER_EXC_CD) = "OE" Or Trim(!EXCEPTION_CD) = "ONP" Then
                r_rsTimecardEntries.AddNew
                r_rsTimecardEntries!EMP_NUM = !EMP_NUM
                r_rsTimecardEntries!OT_OPP_NUM = !OT_OPP_NUM
                r_rsTimecardEntries!CHARGE_DATE = !CHARGE_DATE
                r_rsTimecardEntries!DATE_WORKED = !DATE_WORKED
                r_rsTimecardEntries!START_TIME = !START_TIME
                r_rsTimecardEntries!END_TIME = !END_TIME
                r_rsTimecardEntries!Hours = !Hours
                r_rsTimecardEntries!EXCEPTION_CD = !EXCEPTION_CD
                r_rsTimecardEntries!OVER_EXC_CD = !OVER_EXC_CD
                r_rsTimecardEntries!JOB_NUM = !JOB_NUM
                r_rsTimecardEntries!COMP_SHIFT_CD = !COMP_SHIFT_CD
                r_rsTimecardEntries!OT_REASON_CD = !OT_REASON_CD
                r_rsTimecardEntries!WG_NUM = !WG_NUM
                r_rsTimecardEntries!DISTRIBUTION_ID = !DISTRIBUTION_ID
                r_rsTimecardEntries!DISTRIBUTION_TYPE = !DISTRIBUTION_TYPE
                r_rsTimecardEntries!JOB_NBR = !JOB_NBR
                r_rsTimecardEntries!JOBTASK_NBR = !JOBTASK_NBR
                r_rsTimecardEntries!JOBTASK_STEP_NBR = !JOBTASK_STEP_NBR
                r_rsTimecardEntries!EMPRV_DATABASE_ID = !EMPRV_DATABASE_ID
                If !TRAVEL_PAY_HRS > 0 Then
                    r_rsTimecardEntries!OT_TRAVEL_PAY_IND = "Y"
                End If
                r_rsTimecardEntries!KEY_EMP_NUM = !EMP_NUM
                r_rsTimecardEntries!KEY_CHARGE_DATE = !CHARGE_DATE
                r_rsTimecardEntries!KEY_START_TIME = !START_TIME
                r_rsTimecardEntries!KEY_INDEX_TIME = ""
                r_rsTimecardEntries!TIME_STAMP = !LAST_MOD_TIMESTAMP
                r_rsTimecardEntries!RECORD_STATUS = "I"
            End If
            .MoveNext
        End With
    Loop

    m_rsEmployees.Filter = adFilterNone
    Set l_objEmployee = Nothing
    Set l_objCodeTables = Nothing
    Exit Sub

eh:
    Set l_objEmployee = Nothing
    Set l_objCodeTables = Nothing
    TrsRaiseError mstrCMODULENAME, "CopyTimecardRequest"

End Sub


Private Sub CreateDeleteTimecardRequest( _
                ByVal v_lngOTOppNum As Long, _
                ByVal v_rsTimecardEntries As ADODB.Recordset, _
                ByRef r_rsTimecardRequest As ADODB.Recordset)

On Error GoTo eh

    If Not RecordsExist(v_rsTimecardEntries) Then
        Exit Sub
    End If

    Set r_rsTimecardRequest = CreateTimecardRequest
    v_rsTimecardEntries.MoveFirst
    Do Until v_rsTimecardEntries.EOF
        With v_rsTimecardEntries
            If v_rsTimecardEntries!OT_OPP_NUM = v_lngOTOppNum Then
                r_rsTimecardRequest.AddNew
                r_rsTimecardRequest!EMP_NUM = !EMP_NUM
                r_rsTimecardRequest!OT_OPP_NUM = !OT_OPP_NUM
                r_rsTimecardRequest!CHARGE_DATE = !CHARGE_DATE
                r_rsTimecardRequest!DATE_WORKED = !DATE_WORKED
                r_rsTimecardRequest!START_TIME = !START_TIME
                r_rsTimecardRequest!END_TIME = !END_TIME
                r_rsTimecardRequest!Hours = !Hours
                r_rsTimecardRequest!EXCEPTION_CD = !EXCEPTION_CD
                r_rsTimecardRequest!OVER_EXC_CD = !OVER_EXC_CD
                r_rsTimecardRequest!JOB_NUM = !JOB_NUM
                r_rsTimecardRequest!COMP_SHIFT_CD = !COMP_SHIFT_CD
                r_rsTimecardRequest!OT_REASON_CD = !OT_REASON_CD
                r_rsTimecardRequest!WG_NUM = !WG_NUM
                r_rsTimecardRequest!DISTRIBUTION_ID = !DISTRIBUTION_ID
                r_rsTimecardRequest!DISTRIBUTION_TYPE = !DISTRIBUTION_TYPE
                r_rsTimecardRequest!JOB_NBR = !JOB_NBR
                r_rsTimecardRequest!JOBTASK_NBR = !JOBTASK_NBR
                r_rsTimecardRequest!JOBTASK_STEP_NBR = !JOBTASK_STEP_NBR
                r_rsTimecardRequest!EMPRV_DATABASE_ID = !EMPRV_DATABASE_ID
                If !TRAVEL_PAY_HRS > 0 Then
                    r_rsTimecardRequest!OT_TRAVEL_PAY_IND = "Y"
                End If
                r_rsTimecardRequest!KEY_EMP_NUM = !EMP_NUM
                r_rsTimecardRequest!KEY_CHARGE_DATE = !CHARGE_DATE
                r_rsTimecardRequest!KEY_START_TIME = !START_TIME
                r_rsTimecardRequest!KEY_INDEX_TIME = !INDEX_TIME
'                r_rsTimecardRequest!TIME_STAMP = !LAST_MOD_TIMESTAMP
                r_rsTimecardRequest!RECORD_STATUS = "D"
            End If
            .MoveNext
        End With
    Loop

    Exit Sub

eh:
    TrsRaiseError mstrCMODULENAME, "CreateDeleteTimecardRequest"

End Sub


Private Sub CreateDeleteOtherPayRequest( _
                ByVal v_lngOTOppNum As Long, _
                ByVal v_rsOtherPayEntries As ADODB.Recordset, _
                ByRef r_rsOtherPayRequest As ADODB.Recordset)

On Error GoTo eh

    If Not RecordsExist(v_rsOtherPayEntries) Then
        Exit Sub
    End If

    Set r_rsOtherPayRequest = CreateOtherPayRequest
    v_rsOtherPayEntries.MoveFirst
    Do Until v_rsOtherPayEntries.EOF
        With v_rsOtherPayEntries
            If v_rsOtherPayEntries!OT_OPP_NUM = v_lngOTOppNum Then
                r_rsOtherPayRequest.AddNew
                r_rsOtherPayRequest!EMP_NUM = !EMP_NUM
                r_rsOtherPayRequest!OT_OPP_NUM = !OT_OPP_NUM
                r_rsOtherPayRequest!CHARGE_DATE = !CHARGE_DATE
                r_rsOtherPayRequest!DATE_WORKED = !DATE_WORKED
                r_rsOtherPayRequest!START_TIME = !START_TIME
                r_rsOtherPayRequest!END_TIME = !END_TIME
                r_rsOtherPayRequest!EXCEPTION_CD = !EXCEPTION_CD
'                r_rsOtherPayRequest!OVER_EXC_CD = !OVER_EXC_CD
                r_rsOtherPayRequest!START_LOC_CD = !START_LOC_CD
                r_rsOtherPayRequest!END_LOC_CD = !END_LOC_CD

                r_rsOtherPayRequest!KEY_EMP_NUM = !EMP_NUM
                r_rsOtherPayRequest!KEY_CHARGE_DATE = !CHARGE_DATE
                r_rsOtherPayRequest!KEY_START_TIME = !START_TIME
                r_rsOtherPayRequest!KEY_INDEX_TIME = !INDEX_TIME
                r_rsOtherPayRequest!RECORD_STATUS = "D"
            End If
            .MoveNext
        End With
    Loop

    Exit Sub

eh:
    TrsRaiseError mstrCMODULENAME, "CreateDeleteOtherPayRequest"

End Sub


Private Function GenerateDeleteRequests( _
                ByVal v_lngOTOppNum As Long, _
                ByVal v_aryEmployees As Variant, _
                ByVal v_rsTimecardEntries As ADODB.Recordset) As ADODB.Recordset

    Dim l_rsTimecardRequest As ADODB.Recordset
    Dim i As Integer

On Error GoTo eh

    Set l_rsTimecardRequest = CreateTimecardRequest
    If Not RecordsExist(v_rsTimecardEntries) Or _
            ArraySize(v_aryEmployees) < 1 Then
        GoTo xit
    End If
    
    For i = LBound(v_aryEmployees) To UBound(v_aryEmployees)
        v_rsTimecardEntries.Filter = "EMP_NUM = " & v_aryEmployees(i)
        Do Until v_rsTimecardEntries.EOF
            If v_rsTimecardEntries!OT_OPP_NUM = v_lngOTOppNum Then
            With v_rsTimecardEntries
                l_rsTimecardRequest.AddNew
                l_rsTimecardRequest!EMP_NUM = !EMP_NUM
                l_rsTimecardRequest!OT_OPP_NUM = !OT_OPP_NUM
                l_rsTimecardRequest!CHARGE_DATE = !CHARGE_DATE
                l_rsTimecardRequest!DATE_WORKED = !DATE_WORKED
                l_rsTimecardRequest!START_TIME = !START_TIME
                l_rsTimecardRequest!END_TIME = !END_TIME
                l_rsTimecardRequest!Hours = !Hours
                l_rsTimecardRequest!EXCEPTION_CD = !EXCEPTION_CD
                l_rsTimecardRequest!OVER_EXC_CD = !OVER_EXC_CD
                l_rsTimecardRequest!JOB_NUM = !JOB_NUM
                l_rsTimecardRequest!COMP_SHIFT_CD = !COMP_SHIFT_CD
                l_rsTimecardRequest!OT_REASON_CD = !OT_REASON_CD
                l_rsTimecardRequest!WG_NUM = !WG_NUM
                l_rsTimecardRequest!DISTRIBUTION_ID = !DISTRIBUTION_ID
                l_rsTimecardRequest!DISTRIBUTION_TYPE = !DISTRIBUTION_TYPE
                l_rsTimecardRequest!JOB_NBR = !JOB_NBR
                l_rsTimecardRequest!JOBTASK_NBR = !JOBTASK_NBR
                l_rsTimecardRequest!JOBTASK_STEP_NBR = !JOBTASK_STEP_NBR
                l_rsTimecardRequest!EMPRV_DATABASE_ID = !EMPRV_DATABASE_ID
                If !TRAVEL_PAY_HRS > 0 Then
                    l_rsTimecardRequest!OT_TRAVEL_PAY_IND = "Y"
                End If
                l_rsTimecardRequest!KEY_EMP_NUM = !EMP_NUM
                l_rsTimecardRequest!KEY_CHARGE_DATE = !CHARGE_DATE
                l_rsTimecardRequest!KEY_START_TIME = !START_TIME
                l_rsTimecardRequest!KEY_INDEX_TIME = !INDEX_TIME
                l_rsTimecardRequest!RECORD_STATUS = "D"
            End With
            End If
            v_rsTimecardEntries.MoveNext
        Loop
    Next

xit:
    v_rsTimecardEntries.Filter = adFilterNone
    Set GenerateDeleteRequests = l_rsTimecardRequest
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GenerateDeleteRequests"

End Function


Private Function GenerateOtherPayDeleteRequests( _
                ByVal v_lngOTOppNum As Long, _
                ByVal v_aryEmployees As Variant, _
                ByVal v_rsOtherPayEntries As ADODB.Recordset) As ADODB.Recordset

    Dim l_rsGenOtherPayRequest As ADODB.Recordset
    Dim i As Integer

On Error GoTo eh

    Set l_rsGenOtherPayRequest = CreateOtherPayRequest
    If Not RecordsExist(v_rsOtherPayEntries) Or _
            ArraySize(v_aryEmployees) < 1 Then
        GoTo xit
    End If
    
    For i = LBound(v_aryEmployees) To UBound(v_aryEmployees)
        v_rsOtherPayEntries.Filter = "EMP_NUM = " & v_aryEmployees(i)
        Do Until v_rsOtherPayEntries.EOF
            If v_rsOtherPayEntries!OT_OPP_NUM = v_lngOTOppNum Then
            With v_rsOtherPayEntries
                l_rsGenOtherPayRequest.AddNew
                l_rsGenOtherPayRequest!EMP_NUM = !EMP_NUM
                l_rsGenOtherPayRequest!OT_OPP_NUM = !OT_OPP_NUM
                l_rsGenOtherPayRequest!CHARGE_DATE = !CHARGE_DATE
                l_rsGenOtherPayRequest!DATE_WORKED = !DATE_WORKED
                l_rsGenOtherPayRequest!START_TIME = !START_TIME
                l_rsGenOtherPayRequest!EXCEPTION_CD = !EXCEPTION_CD
                l_rsGenOtherPayRequest!WG_NUM = !WG_NUM

                l_rsGenOtherPayRequest!KEY_EMP_NUM = !EMP_NUM
                l_rsGenOtherPayRequest!KEY_CHARGE_DATE = !CHARGE_DATE
                l_rsGenOtherPayRequest!KEY_START_TIME = !START_TIME
                l_rsGenOtherPayRequest!KEY_INDEX_TIME = !INDEX_TIME
                l_rsGenOtherPayRequest!RECORD_STATUS = "D"
            End With
            End If
            v_rsOtherPayEntries.MoveNext
        Loop
    Next

xit:
    v_rsOtherPayEntries.Filter = adFilterNone
    Set GenerateOtherPayDeleteRequests = l_rsGenOtherPayRequest
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GenerateOtherPayDeleteRequests"

End Function


Private Sub AppendDeleteRequests( _
                ByVal v_lngOTOppNum As Long, _
                ByVal v_rsGenDeletes As ADODB.Recordset, _
                ByVal v_rsTimecardRequest As ADODB.Recordset)
' Do not append duplicates to v_rsTarget.

On Error GoTo eh

    v_rsTimecardRequest.Filter = "RECORD_STATUS = 'D'"
    If RecordsExist(v_rsGenDeletes) Then
        v_rsGenDeletes.MoveFirst
    End If
    
    Do Until v_rsGenDeletes.EOF
        If Not TimecardRequestIncludes(v_rsGenDeletes!KEY_EMP_NUM, _
                v_rsGenDeletes!KEY_CHARGE_DATE, v_rsGenDeletes!KEY_START_TIME, v_rsTimecardRequest) Then
            CopyDeleteTimecardRequest v_rsGenDeletes, v_rsTimecardRequest
        End If
        v_rsGenDeletes.MoveNext
    Loop
    
    Exit Sub

eh:
    TrsRaiseError mstrCMODULENAME, "AppendDeleteRequests"
    
End Sub


Private Sub AppendOtherPayDeleteRequests( _
                ByVal v_lngOTOppNum As Long, _
                ByVal v_rsGenDeletes As ADODB.Recordset, _
                ByVal v_rsOtherPayRequest As ADODB.Recordset)
' Do not append duplicates to v_rsTarget.

On Error GoTo eh

    v_rsOtherPayRequest.Filter = "RECORD_STATUS = 'D'"
    If RecordsExist(v_rsGenDeletes) Then
        v_rsGenDeletes.MoveFirst
    End If
    
    Do Until v_rsGenDeletes.EOF
        If Not TimecardRequestIncludes(v_rsGenDeletes!KEY_EMP_NUM, _
                v_rsGenDeletes!KEY_CHARGE_DATE, v_rsGenDeletes!KEY_START_TIME, v_rsOtherPayRequest) Then
            CopyDeleteOtherPayRequest v_rsGenDeletes, v_rsOtherPayRequest
        End If
        v_rsGenDeletes.MoveNext
    Loop
    
    Exit Sub

eh:
    TrsRaiseError mstrCMODULENAME, "AppendOtherPayDeleteRequests"
    
End Sub


Private Function TimecardRequestIncludes( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_strStartTime As String, _
                    ByRef r_rsTimecardRequest As ADODB.Recordset) As Boolean

On Error GoTo eh

    FindTimecardRequest v_lngEmpNum, _
                v_datChargeDate, v_strStartTime, r_rsTimecardRequest
                
    TimecardRequestIncludes = Not r_rsTimecardRequest.EOF

    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "TimecardRequestIncludes"
    
End Function


Private Sub FindTimecardRequest( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date, _
                ByVal v_strStartTime As String, _
                ByRef r_rsTimecardRequest As ADODB.Recordset)

On Error GoTo eh

    If Not RecordsExist(r_rsTimecardRequest) Then
        Exit Sub
    End If
    
    r_rsTimecardRequest.MoveFirst
    Do Until r_rsTimecardRequest.EOF
        If r_rsTimecardRequest!KEY_EMP_NUM = v_lngEmpNum And _
            r_rsTimecardRequest!KEY_CHARGE_DATE = v_datChargeDate And _
            r_rsTimecardRequest!KEY_START_TIME = v_strStartTime Then
            Exit Do
        Else
            r_rsTimecardRequest.MoveNext
        End If
    Loop

    Exit Sub


eh:
    TrsRaiseError mstrCMODULENAME, "FindTimecardRequest"
                
End Sub


Private Sub AppendTimecardErrors( _
                ByVal v_rsTimecardErrors As ADODB.Recordset, _
                ByRef r_rsOvertimeErrors As ADODB.Recordset)

On Error GoTo eh

    If Not RecordsExist(v_rsTimecardErrors) Then
        GoTo xit
    End If
    
    If r_rsOvertimeErrors Is Nothing Then
        Set r_rsOvertimeErrors = CreateNewBusinessErrorSet
    End If

    v_rsTimecardErrors.MoveFirst
    Do Until v_rsTimecardErrors.EOF
        r_rsOvertimeErrors.AddNew
        r_rsOvertimeErrors!ERR_NUMBER = v_rsTimecardErrors!ERR_NUMBER
        r_rsOvertimeErrors!ERR_TYPE = v_rsTimecardErrors!ERR_TYPE
        r_rsOvertimeErrors!ERR_DESCRIPTION = v_rsTimecardErrors!ERR_DESCRIPTION
        r_rsOvertimeErrors!ERR_SOURCE = v_rsTimecardErrors!ERR_SOURCE
        r_rsOvertimeErrors!RECORD_TYPE = TrType_OT_TC
        r_rsOvertimeErrors!FIELD_NAMES = "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD"
        r_rsOvertimeErrors!KEY_FIELD1 = CStr(v_rsTimecardErrors!EMP_NUM)
        r_rsOvertimeErrors!KEY_FIELD2 = Format(v_rsTimecardErrors!CHARGE_DATE, "mm/dd/yyyy")
        r_rsOvertimeErrors!KEY_FIELD3 = Format(v_rsTimecardErrors!DATE_WORKED, "mm/dd/yyyy")
        r_rsOvertimeErrors!KEY_FIELD4 = v_rsTimecardErrors!START_TIME
        r_rsOvertimeErrors!KEY_FIELD5 = CStr(v_rsTimecardErrors!ELAPSED_HOURS)
        r_rsOvertimeErrors!KEY_FIELD6 = v_rsTimecardErrors!EXCEPTION_CD
        v_rsTimecardErrors.MoveNext
    Loop
    
    If RecordsExist(r_rsOvertimeErrors) Then
        r_rsOvertimeErrors.MoveFirst
    End If

xit:
    Exit Sub

eh:
    TrsRaiseError mstrCMODULENAME, "AppendTimecardErrors"
    
End Sub


Private Sub CopyDeleteTimecardRequest( _
                ByVal v_rsDeleteRequest As ADODB.Recordset, _
                ByRef r_rsTimecardRequest As ADODB.Recordset)

On Error GoTo eh
    
    With v_rsDeleteRequest
        r_rsTimecardRequest.AddNew
        r_rsTimecardRequest!EMP_NUM = !EMP_NUM
        r_rsTimecardRequest!OT_OPP_NUM = !OT_OPP_NUM
        r_rsTimecardRequest!CHARGE_DATE = !CHARGE_DATE
        r_rsTimecardRequest!DATE_WORKED = !DATE_WORKED
        r_rsTimecardRequest!START_TIME = !START_TIME
        r_rsTimecardRequest!END_TIME = !END_TIME
        r_rsTimecardRequest!Hours = !Hours
        r_rsTimecardRequest!EXCEPTION_CD = !EXCEPTION_CD
        r_rsTimecardRequest!OVER_EXC_CD = !OVER_EXC_CD
        r_rsTimecardRequest!JOB_NUM = !JOB_NUM
        r_rsTimecardRequest!COMP_SHIFT_CD = !COMP_SHIFT_CD
        r_rsTimecardRequest!OT_REASON_CD = !OT_REASON_CD
        r_rsTimecardRequest!WG_NUM = !WG_NUM
        r_rsTimecardRequest!DISTRIBUTION_ID = !DISTRIBUTION_ID
        r_rsTimecardRequest!DISTRIBUTION_TYPE = !DISTRIBUTION_TYPE
        r_rsTimecardRequest!JOB_NBR = !JOB_NBR
        r_rsTimecardRequest!JOBTASK_NBR = !JOBTASK_NBR
        r_rsTimecardRequest!JOBTASK_STEP_NBR = !JOBTASK_STEP_NBR
        r_rsTimecardRequest!EMPRV_DATABASE_ID = !EMPRV_DATABASE_ID
        r_rsTimecardRequest!OT_TRAVEL_PAY_IND = !OT_TRAVEL_PAY_IND
        r_rsTimecardRequest!KEY_EMP_NUM = !EMP_NUM
        r_rsTimecardRequest!KEY_CHARGE_DATE = !CHARGE_DATE
        r_rsTimecardRequest!KEY_START_TIME = !START_TIME
        r_rsTimecardRequest!KEY_INDEX_TIME = !KEY_INDEX_TIME
'        r_rsTimecardRequest!TIME_STAMP = !LAST_MOD_TIMESTAMP
        r_rsTimecardRequest!RECORD_STATUS = !RECORD_STATUS
    End With

    Exit Sub

eh:
    TrsRaiseError mstrCMODULENAME, "CopyDeleteTimecardRequest"

End Sub


Private Sub CopyDeleteOtherPayRequest( _
                ByVal v_rsDeleteRequest As ADODB.Recordset, _
                ByRef r_rsOtherPayRequest As ADODB.Recordset)

On Error GoTo eh
    
    With v_rsDeleteRequest
        r_rsOtherPayRequest.AddNew
        r_rsOtherPayRequest!EMP_NUM = !EMP_NUM
        r_rsOtherPayRequest!OT_OPP_NUM = !OT_OPP_NUM
        r_rsOtherPayRequest!CHARGE_DATE = !CHARGE_DATE
        r_rsOtherPayRequest!DATE_WORKED = !DATE_WORKED
        r_rsOtherPayRequest!START_TIME = !START_TIME
        r_rsOtherPayRequest!END_TIME = !END_TIME
        r_rsOtherPayRequest!AMOUNT = !AMOUNT
        r_rsOtherPayRequest!EXCEPTION_CD = !EXCEPTION_CD
        r_rsOtherPayRequest!START_LOC_CD = !START_LOC_CD
        r_rsOtherPayRequest!END_LOC_CD = !END_LOC_CD
        r_rsOtherPayRequest!WG_NUM = !WG_NUM
        
        r_rsOtherPayRequest!KEY_EMP_NUM = !EMP_NUM
        r_rsOtherPayRequest!KEY_CHARGE_DATE = !CHARGE_DATE
        r_rsOtherPayRequest!KEY_START_TIME = !START_TIME
        r_rsOtherPayRequest!KEY_INDEX_TIME = !KEY_INDEX_TIME
'        r_rsOtherPayRequest!TIME_STAMP = !LAST_MOD_TIMESTAMP
        r_rsOtherPayRequest!RECORD_STATUS = !RECORD_STATUS
    End With


    Exit Sub

eh:
    TrsRaiseError mstrCMODULENAME, "CopyDeleteTimecardRequest"

End Sub


Private Function CreateTimecardRequest() As ADODB.Recordset

On Error GoTo eh

    Const l_lngCFIELDATTR = adFldUpdatable Or adFldIsNullable
    Dim l_adrsSaveTime As ADODB.Recordset

    Set l_adrsSaveTime = New ADODB.Recordset
    With l_adrsSaveTime
        .Fields.Append "EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CHARGE_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "DATE_WORKED", adDate, , l_lngCFIELDATTR
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, l_lngCFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "HOURS", adDecimal, , l_lngCFIELDATTR
            .Fields.Item("HOURS").Precision = 3
            .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "START_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "END_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "OVER_EXC_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, l_lngCFIELDATTR
        .Fields.Append "OT_REASON_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "OT_OPP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "OT_OPP_DESC", adVarChar, 60, l_lngCFIELDATTR
        .Fields.Append "OT_TRAVEL_PAY_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "OT_ADJUST_DESC", adChar, 30, l_lngCFIELDATTR
        .Fields.Append "WG_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CASE_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "FAMILY_LEAVE_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "COMP_SHIFT_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "COMP_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "REL_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "SCHED_CHG_EXC_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "DISTR_DEPT_NUM", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "APPR_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "TIME_STAMP", adChar, 26, l_lngCFIELDATTR
        .Fields.Append "JOB_NBR", adVarChar, 8, l_lngCFIELDATTR
        .Fields.Append "JOBTASK_NBR", adVarChar, 4, l_lngCFIELDATTR
        .Fields.Append "JOBTASK_STEP_NBR", adInteger, 1, l_lngCFIELDATTR
        .Fields.Append "EMPRV_DATABASE_ID", adVarChar, 8, l_lngCFIELDATTR

        'Non-nullable information
        '   RECORD_STATUS CAN BE
        '   D - DELETE
        '   I - INSERT
        '   C - CHANGE
        '   SCHED_TIME_IND CAN BE
        '   Y - Yes Scheduled Time
        '   N - No Actual Time
        .Fields.Append "RECORD_STATUS", adChar, 1, adFldUpdatable
        .Fields.Append "SCHED_TIME_IND", adChar, 1, adFldUpdatable

        'KEY INFORMATION
        .Fields.Append "KEY_EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "KEY_CHARGE_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "KEY_START_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "KEY_INDEX_TIME", adChar, 26, l_lngCFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic

        .Open
    End With

    Set CreateTimecardRequest = l_adrsSaveTime
    Set l_adrsSaveTime = Nothing

    Exit Function

eh:
    Set l_adrsSaveTime = Nothing
    Set CreateTimecardRequest = Nothing

    TrsRaiseError mstrCMODULENAME, "CreateTimecardRequest"

End Function


Private Function CreateOtherPayRequest() As ADODB.Recordset

    Const l_lngCFIELDATTR = adFldUpdatable Or adFldIsNullable
    Dim l_adrsSaveOtherPay As ADODB.Recordset

On Error GoTo eh

    Set l_adrsSaveOtherPay = New ADODB.Recordset

    With l_adrsSaveOtherPay
        .Fields.Append "EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CHARGE_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "DATE_WORKED", adDate, , l_lngCFIELDATTR
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, l_lngCFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "HOURS", adDecimal, , l_lngCFIELDATTR
            .Fields.Item("HOURS").Precision = 3
            .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "AMOUNT", adDecimal, , l_lngCFIELDATTR
            .Fields.Item("AMOUNT").Precision = 9
            .Fields.Item("AMOUNT").NumericScale = 2
        .Fields.Append "START_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "END_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "START_LOC_CD", adInteger, , l_lngCFIELDATTR
        .Fields.Append "END_LOC_CD", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CHEM_MTL_TSTR_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "HALF_TRIP_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, l_lngCFIELDATTR
        .Fields.Append "WG_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "COMP_SHIFT_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "DISTR_DEPT_NUM", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "APPR_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "OT_OPP_NUM", adInteger, , l_lngCFIELDATTR

        .Fields.Append "LOC_CD", adInteger, , l_lngCFIELDATTR
        .Fields.Append "DEPT_NUM", adChar, 3, l_lngCFIELDATTR
        
        .Fields.Append "RECORD_STATUS", adChar, 1, adFldUpdatable
        .Fields.Append "SCHED_TIME_IND", adChar, 1, adFldUpdatable

        'KEY INFORMATION
        .Fields.Append "KEY_EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "KEY_CHARGE_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "KEY_START_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "KEY_INDEX_TIME", adChar, 26, l_lngCFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic

        .Open
    End With

    Set CreateOtherPayRequest = l_adrsSaveOtherPay
    Set l_adrsSaveOtherPay = Nothing

    Exit Function

eh:
    Set l_adrsSaveOtherPay = Nothing
    Set CreateOtherPayRequest = Nothing

    TrsRaiseError mstrCMODULENAME, "CreateOtherPayRequest"

End Function


Private Sub PrepareRequest( _
                l_strEmpNum As String, _
                l_rsEmployees As ADODB.Recordset, _
                l_rsEmpWorkgroups As ADODB.Recordset, _
                l_rsTimecardEntries As ADODB.Recordset, _
                l_rsOtherPayEntries As ADODB.Recordset, _
                v_rsSchedules As ADODB.Recordset, _
                v_rsTimecardRequest As ADODB.Recordset, _
                v_rsOtherPayRequest As ADODB.Recordset)

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
    
    If Not l_rsOtherPayEntries Is Nothing Then
        l_rsOtherPayEntries.Filter = adFilterNone
    End If
    If RecordsExist(l_rsOtherPayEntries) Then
        l_rsOtherPayEntries.Filter = "EMP_NUM = " & l_strEmpNum
        l_rsOtherPayEntries.Sort = "CHARGE_DATE ASC"
    End If
    
    If Not v_rsSchedules Is Nothing Then
        v_rsSchedules.Filter = adFilterNone
    End If
    If RecordsExist(v_rsSchedules) Then
        v_rsSchedules.Filter = "EMP_NUM = " & l_strEmpNum
        v_rsSchedules.Sort = "CHARGE_DATE ASC"
    End If
    
    If Not v_rsTimecardRequest Is Nothing Then
        v_rsTimecardRequest.Filter = adFilterNone
    End If
    If RecordsExist(v_rsTimecardRequest) Then
'        v_rsTimecardRequest.Filter = "KEY_EMP_NUM = " & l_strEmpNum
'        v_rsTimecardRequest.Sort = "KEY_CHARGE_DATE ASC, EXCEPTION_CD DESC"
        v_rsTimecardRequest.Filter = "EMP_NUM = " & l_strEmpNum
        v_rsTimecardRequest.Sort = "CHARGE_DATE ASC, EXCEPTION_CD DESC"
    End If
    
    If Not v_rsOtherPayRequest Is Nothing Then
        v_rsOtherPayRequest.Filter = adFilterNone
    End If
    If RecordsExist(v_rsOtherPayRequest) Then
        v_rsOtherPayRequest.Filter = "KEY_EMP_NUM = " & l_strEmpNum
        v_rsOtherPayRequest.Sort = "KEY_CHARGE_DATE ASC, EXCEPTION_CD DESC"
'        v_rsOtherPayRequest.Filter = "EMP_NUM = " & l_strEmpNum
'        v_rsOtherPayRequest.Sort = "CHARGE_DATE ASC, EXCEPTION_CD DESC"
    End If
    
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "PrepareRequest"

End Sub


