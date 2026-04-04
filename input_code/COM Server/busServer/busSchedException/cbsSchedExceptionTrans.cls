VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsSchedExcTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cbsSchedExcTrans"
Private m_lngUserNum As Long
Private m_strUserId As String
Private m_rsErrors As ADODB.Recordset
Private m_rsSchedules As ADODB.Recordset
Private m_rsExceptions As ADODB.Recordset
Private m_rsActualTime As ADODB.Recordset
Private m_colEmployees As Collection
Private m_aryEmployees As Variant
Private m_rsDeleteActualTime As ADODB.Recordset
Private m_rsEmployeeData As ADODB.Recordset
Private m_rsCollPeriods As ADODB.Recordset
Private m_rsAccrualBankReq As ADODB.Recordset
Private m_rsEmployeeWorkGroups As ADODB.Recordset
Private m_colEMPRVTrans As Collection
Private m_rsSchedChgPmts As ADODB.Recordset
Private m_rsOriginalRequest As ADODB.Recordset
Private m_rsHolidays As ADODB.Recordset
Private m_rsRequest As ADODB.Recordset

Private Enum EMPRVTranEnum
    etEMPNUM = 0
    etDEPTNUM
    etEXCCD
    etCHGDATE
    etHOURS
    etTRANCD
    etLOCCD
End Enum

'Use for validation
Private mlngErrorNumber As Long
Private mstrErrorDescription As String

Public Function SaveExceptions( _
    ByVal v_rsRequest As ADODB.Recordset, _
    ByRef r_rsBRVs As ADODB.Recordset, _
    ByVal v_strLastModEmpNum As String _
    ) As Boolean

    Dim l_objCtx As ObjectContext
    Dim l_blnError As Boolean

    On Error GoTo eh

    Set l_objCtx = GetObjectContext
    If Not RecordsExist(v_rsRequest) Then
        SaveExceptions = True
        GetObjectContext.SetComplete
        GoTo Xit
    End If

    m_strUserId = v_strLastModEmpNum
    m_lngUserNum = ParseUserID(v_strLastModEmpNum)
    Set m_rsErrors = CreateNewBusinessErrorSet()

    GetState v_rsRequest

    If Not PreValidate(v_rsRequest) Then
        SaveExceptions = False
        GetObjectContext.SetAbort
        GoTo Xit
    End If

    ModifyRequest v_rsRequest

    If Not ProcessExceptions(v_rsRequest) Then
        SaveExceptions = False
        GetObjectContext.SetAbort
        GoTo Xit
    End If
    
    If Not PostValidate() Then
        SaveExceptions = False
        GetObjectContext.SetAbort
        GoTo Xit
    End If
    
    If Not DeleteActualTime() Then
        SaveExceptions = False
        GetObjectContext.SetAbort
        GoTo Xit
    End If

    If Not UpdateAccrualBanks() Then
        SaveExceptions = False
        GetObjectContext.SetAbort
        GoTo Xit
    End If
        
    If Not SaveExceptionsDB() Then
        SaveExceptions = False
        GetObjectContext.SetAbort
        GoTo Xit
    End If

    If m_colEMPRVTrans.Count > 0 Then
         SaveEMPRV
    End If
    
   SaveExceptions = True

    GetObjectContext.SetComplete
Xit:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsBRVs = Nothing
        Else
            Set r_rsBRVs = m_rsErrors
        End If
    Else
        Set r_rsBRVs = Nothing
    End If
    SetGlobalObjectsToNothing
    
    Exit Function
    
eh:
    SetGlobalObjectsToNothing
    Set r_rsBRVs = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "SaveExceptions", , , , , True
    
End Function

Private Function BuildGABRequest( _
    ByVal v_objRS As ADODB.Recordset _
    ) As ADODB.Recordset

    Dim l_rsReq As ADODB.Recordset
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date

    On Error GoTo eh

    Set l_rsReq = CreateGABRequestRecordset()

    v_objRS.Filter = "REC_STATUS <> 'D'"
    If Not RecordsExist(v_objRS) Then
        v_objRS.Filter = ""
        Set BuildGABRequest = Nothing
        Exit Function
    End If

    v_objRS.Sort = "EMP_NUM, CHARGE_DATE"
    v_objRS.MoveFirst
    l_lngEmpNum = 0
    l_datChargeDate = CDate("1980-01-01")


    While Not v_objRS.EOF
        If v_objRS!EMP_NUM = l_lngEmpNum Then
            If v_objRS!CHARGE_DATE = l_datChargeDate Then
            Else
                l_rsReq.AddNew
                l_rsReq!EMP_NUM = v_objRS!EMP_NUM
                l_rsReq!START_DATE = v_objRS!CHARGE_DATE
                l_rsReq!END_DATE = v_objRS!CHARGE_DATE
            End If
        Else
            l_rsReq.AddNew
            l_rsReq!EMP_NUM = v_objRS!EMP_NUM
            l_rsReq!START_DATE = v_objRS!CHARGE_DATE
            l_rsReq!END_DATE = v_objRS!CHARGE_DATE
        End If
        l_lngEmpNum = v_objRS!EMP_NUM
        l_datChargeDate = v_objRS!CHARGE_DATE
        v_objRS.MoveNext
    Wend

    v_objRS.Filter = ""
    Set BuildGABRequest = l_rsReq
    Set l_rsReq = Nothing
    
    Exit Function
    
eh:
    Set l_rsReq = Nothing
    TrsRaiseError m_strCMODULENAME, "BuildGABRequest"


End Function

Private Function CreateGABRequestRecordset()

    Dim l_rs As ADODB.Recordset

    On Error GoTo eh

    Set l_rs = New ADODB.Recordset

    With l_rs
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "START_DATE", adDate, , adFldUpdatable
      .Fields.Append "END_DATE", adDate, , adFldUpdatable
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With

    Set CreateGABRequestRecordset = l_rs
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CreateGABRequestRecordset"


End Function

Private Function GetSchedOnExceptionDays( _
    ByVal v_rsReq As ADODB.Recordset _
    ) As ADODB.Recordset

    Dim l_obj As busSchedule.cbsSchedule
    
    On Error GoTo eh

    Set l_obj = CreateObject("busSchedule.cbsSchedule")

    Set GetSchedOnExceptionDays = l_obj.GABEmployeeSchedOnSpecificDates( _
        v_rsReq, Nothing, False, _
        trBaseSchedNoDetails, trAllScheduleLevels)

    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetSchedOnExceptionDays"

End Function

Private Function CreateAllTimeCasesReqRS() As ADODB.Recordset
    Dim l_rsRequest As ADODB.Recordset

    On Error GoTo eh

    Set l_rsRequest = New ADODB.Recordset
    With l_rsRequest
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "CASE_GROUP", adChar, 3, adFldUpdatable
      .Fields.Append "CASE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "REC_STATUS", adChar, 1, adFldUpdatable

      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockBatchOptimistic
      .Open
    End With

    Set CreateAllTimeCasesReqRS = l_rsRequest
    Set l_rsRequest = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CreateAllTimeCasesReqRS"

End Function


Private Sub GetState(ByVal v_rsReq As ADODB.Recordset)

    Dim l_rsGABReq As ADODB.Recordset
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_colEMPRVEmployees As Collection
    Dim l_aryEMPRVEmployees As Variant

    On Error GoTo eh


    Set m_rsOriginalRequest = CopyRecordset(v_rsReq)
' Get all exception from 30 days in past to 12/31/9999.
' Build gab request for each new or changed
' exception in the client request.
    Set m_colEmployees = GetEmployeesInRequest(v_rsReq)
    m_aryEmployees = ArrayFromCollection(m_colEmployees)
    
    Set m_rsExceptions = GetScheduledExceptions(m_aryEmployees, DateAdd("d", -30, Date), #12/31/9999#)

    Set l_rsGABReq = BuildGABRequest(v_rsReq)
    If RecordsExist(m_rsExceptions) Then
        m_rsExceptions.Sort = "CHARGE_DATE"
        m_rsExceptions.MoveFirst
        l_datStart = CDate(m_rsExceptions!CHARGE_DATE)
        m_rsExceptions.MoveLast
        l_datEnd = CDate(m_rsExceptions!CHARGE_DATE)
    Else
        l_datStart = CDate("9999-12-31")
        l_datEnd = CDate("1000-01-01")
    End If
    v_rsReq.Sort = "CHARGE_DATE"
    v_rsReq.MoveFirst
    If CDate(v_rsReq!CHARGE_DATE) < l_datStart Then
        l_datStart = v_rsReq!CHARGE_DATE
    End If
    v_rsReq.MoveLast
    If CDate(v_rsReq!CHARGE_DATE) > l_datEnd Then
        l_datEnd = v_rsReq!CHARGE_DATE
    End If
    
    Set m_rsSchedules = GetSchedOnExceptionDays(l_rsGABReq)
    Set m_rsActualTime = GetActualTime(m_aryEmployees, l_datStart, l_datEnd)
    Set m_rsCollPeriods = GetCurrentCollPeriods()
    Set m_rsSchedChgPmts = GetScheduleChangePayments(m_aryEmployees, l_datStart, l_datEnd)
    Set m_rsEmployeeData = GetEmployeeData(m_aryEmployees)
    Set l_colEMPRVEmployees = DetermineEMPRVEmployees()
    If l_colEMPRVEmployees.Count > 0 Then
        l_aryEMPRVEmployees = ArrayFromCollection(l_colEMPRVEmployees)
        Set m_rsEmployeeWorkGroups = GetEmployeeWorkGroups(l_aryEMPRVEmployees, l_datStart, l_datEnd)
    Else
        Set m_rsEmployeeWorkGroups = Nothing
    End If
    Set m_colEMPRVTrans = New Collection
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetState"


End Sub

Private Function GetEmployeesInRequest( _
    ByVal v_objRS As ADODB.Recordset _
    ) As Collection

    Dim l_colEmployees As Collection
    Dim l_strEmployee As String
    Dim l_aryEmployees As Variant

    On Error GoTo eh

    Set l_colEmployees = New Collection

    If Not RecordsExist(v_objRS) Then
        Set GetEmployeesInRequest = l_colEmployees
        Exit Function
    End If

    v_objRS.Sort = "EMP_NUM"
    v_objRS.MoveFirst
    l_strEmployee = ""
    While Not v_objRS.EOF
        If Trim(str(v_objRS!EMP_NUM)) <> l_strEmployee Then
            l_strEmployee = Trim(str(v_objRS!EMP_NUM))
            If Not IncludesString(l_strEmployee, l_colEmployees) Then
                l_colEmployees.Add l_strEmployee
            End If
        End If
        v_objRS.MoveNext
    Wend
    v_objRS.Sort = ""
    v_objRS.MoveFirst

    Set GetEmployeesInRequest = l_colEmployees

    Exit Function

eh:

  TrsRaiseError m_strCMODULENAME, "GetEmployeesInRequest"

End Function

Private Function GetScheduledExceptions( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStart As Date, _
    ByVal l_datEnd As Date _
    ) As ADODB.Recordset

'    ByVal l_datLastPayEnd As Date,
'    ByVal l_datLastDayOfOpenPP As Date

   Dim l_obj As dbSchedException.cdSchedException

    On Error GoTo eh

   Set l_obj = CreateObject("dbSchedException.cdSchedException")

   Set GetScheduledExceptions = l_obj.GetScheduledExceptions( _
    v_aryEmployees, v_datStart, l_datEnd)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetScheduledExceptions"


End Function

Private Function GetActualTime( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date _
    ) As ADODB.Recordset

    Dim l_obj As dbTimeCard.cdTimeCard

    On Error GoTo eh

    Set l_obj = CreateObject("dbTimeCard.cdTimeCard")

    Set GetActualTime = l_obj.GetExceptionSchedTimeEntries( _
        v_aryEmployees, v_datStart, v_datEnd)

    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTime"


End Function


Private Function ProcessExceptions( _
    ByVal v_rsRequest As ADODB.Recordset _
    ) As Boolean

    Dim l_blnOk As Boolean
    Dim l_blnScheduleExc As Boolean

    On Error GoTo eh

    Set m_rsDeleteActualTime = CreateActualTimeDeleteRS()
    Set m_rsAccrualBankReq = CreateAccrualBankMgrRecordset()
    l_blnOk = True
    v_rsRequest.Sort = "REC_STATUS Desc, EMP_NUM, CHARGE_DATE"
    v_rsRequest.MoveFirst
    While Not v_rsRequest.EOF
        l_blnScheduleExc = True
        If Not CanActualTimeBeDeleted(v_rsRequest) Then
            l_blnOk = False
            l_blnScheduleExc = False
        End If
        If l_blnScheduleExc Then
            If Not ProcessScheduledException(v_rsRequest) Then
                l_blnOk = False
            End If
        End If

        v_rsRequest.MoveNext
    Wend

    ProcessExceptions = l_blnOk
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ProcessExceptions"

End Function

Private Function CanActualTimeBeDeleted( _
    ByVal v_rsRequest As ADODB.Recordset _
    ) As Boolean
    
    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strFilter As String
    Dim l_blnProcessed As Boolean
    Dim l_blnApproved As Boolean
    Dim l_strProcessedInd As String
    Dim l_strDeleteActualTime As String
    Dim l_strApprAuthInd As String
    
    On Error GoTo eh
    
    If IsNull(v_rsRequest!DELETE_ACTUAL_TIME) Or _
        IsEmpty(v_rsRequest!DELETE_ACTUAL_TIME) Then
        l_strDeleteActualTime = "N"
    Else
        l_strDeleteActualTime = v_rsRequest!DELETE_ACTUAL_TIME
        If l_strDeleteActualTime <> "Y" Then
            l_strDeleteActualTime = "N"
        End If
    End If
    
    If IsNull(v_rsRequest!APPR_AUTH_IND) Or _
        IsEmpty(v_rsRequest!APPR_AUTH_IND) Then
        l_strApprAuthInd = "N"
    Else
        l_strApprAuthInd = v_rsRequest!APPR_AUTH_IND
        If l_strApprAuthInd <> "Y" Then
            l_strApprAuthInd = "N"
        End If
    End If
    
    
    
    l_lngEmpNum = v_rsRequest!EMP_NUM
    l_datChargeDate = v_rsRequest!CHARGE_DATE
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
        " AND CHARGE_DATE = " & str(l_datChargeDate)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsActualTime.Filter = l_strFilter
    
    If Not RecordsExist(m_rsActualTime) Then
        CanActualTimeBeDeleted = True
        Exit Function
    End If
    
    l_blnProcessed = False
    l_blnApproved = False
    m_rsActualTime.MoveFirst
    While Not m_rsActualTime.EOF _
        And Not l_blnProcessed _
        And Not l_blnApproved
        l_strProcessedInd = Trim(m_rsActualTime!PROCESSED_IND)
        If l_strProcessedInd = "" Or l_strProcessedInd = "N" Then
        Else
            l_blnProcessed = True
        End If
        If Trim(m_rsActualTime!APPR_IND) <> "N" Then
            l_blnApproved = True
        End If
        m_rsActualTime.MoveNext
    Wend
    
    If l_blnProcessed Then
        CanActualTimeBeDeleted = False
        AppendEntryError v_rsRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception can not be processed.  Can not delete time that has already been processed by payroll."
        Exit Function
    End If
    
    If l_strDeleteActualTime <> "Y" Then
        CanActualTimeBeDeleted = False
        AppendEntryError v_rsRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception can not be processed.  Actual Time exists. If you wish to delete the actual time, please check the delete actual time box and save again."
        Exit Function
    End If
    
    If l_blnApproved Then
        If l_strApprAuthInd = "Y" Then
            CanActualTimeBeDeleted = True
            AddActualTimeDelete l_lngEmpNum, l_datChargeDate
        Else
            CanActualTimeBeDeleted = False
            AppendEntryError v_rsRequest, trErrbusValidationError, mstrCBUSINESSERR, "Exception can not be processed.  Actual Time can not be deleted becuase it has been approved and you do not have the authority to delete this approved time."
        End If
    Else
        CanActualTimeBeDeleted = True
        AddActualTimeDelete l_lngEmpNum, l_datChargeDate
    End If

    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CanActualTimeBeDeleted"
    
    
End Function
    
    
Private Function GetScheduleChangePayments( _
                    ByVal v_aryEmpNums As Variant, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date) _
                As ADODB.Recordset

    Dim l_obj As dbSchedChgPay.cdSchedChgPay
        
On Error GoTo eh

    Set l_obj = CreateObject("dbSchedChgPay.cdSchedChgPay")
    Set GetScheduleChangePayments = _
        l_obj.GetSchedChgPay(v_aryEmpNums, v_datStartDate, v_datEndDate)
    Set l_obj = Nothing
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetScheduleChangePayments"
End Function
    
    
Private Function GetCurrentCollPeriods() As ADODB.Recordset

    Dim l_obj As dbCollectionPeriod.cdCollectionPeriod
    
    
    On Error GoTo eh

    Set l_obj = CreateObject("dbCollectionPeriod.cdCollectionPeriod")

    Set GetCurrentCollPeriods = l_obj.GetCurrentCollectionPeriods()

    Set l_obj = Nothing

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetCurrentCollPeriods"
    
End Function

Private Function GetEmployeeData( _
    ByVal v_aryEmps As Variant _
    ) As ADODB.Recordset

    Dim l_obj As dbEmployee.cdEmployee
    
    
    On Error GoTo eh

    Set l_obj = CreateObject("dbEmployee.cdEmployee")

    Set GetEmployeeData = l_obj.GetEmployeeExcSchedInfo(v_aryEmps)

    Set l_obj = Nothing

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployeeData"
    
End Function


Private Function GetLockStatus( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As String

    Dim l_obj As busEmployee.cbsEmployee
    
    On Error GoTo eh

    Set l_obj = CreateObject("busEmployee.cbsEmployee")

    GetLockStatus = l_obj.GetLockStatus( _
        v_lngEmpNum, v_datDate, m_rsEmployeeData, m_rsCollPeriods)

    Set l_obj = Nothing

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetLockStatus"
    
End Function

Private Sub SetGlobalObjectsToNothing()


    Set m_rsSchedules = Nothing
    Set m_rsExceptions = Nothing
    Set m_colEmployees = Nothing
    Set m_rsDeleteActualTime = Nothing
    Set m_rsErrors = Nothing
    Set m_rsEmployeeData = Nothing
    Set m_rsActualTime = Nothing
    Set m_rsCollPeriods = Nothing
    Set m_rsAccrualBankReq = Nothing
    Set m_rsEmployeeWorkGroups = Nothing
    Set m_colEMPRVTrans = Nothing
    Set m_rsSchedChgPmts = Nothing


End Sub

Private Function PreValidate( _
    ByVal v_rsRequest As ADODB.Recordset _
    ) As Boolean
    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_rsLockStatus As String
    Dim l_blnValid As Boolean
    Dim l_obj As dbCodeTables.cdCodeTables
    Dim StartDate As Date
    Dim EndDate As Date
    
    On Error GoTo eh
    
    If Not RecordsExist(v_rsRequest) Then
        PreValidate = True
        Exit Function
    End If
    v_rsRequest.Sort = "CHARGE_DATE"
    v_rsRequest.MoveFirst
    StartDate = v_rsRequest!CHARGE_DATE
    v_rsRequest.MoveLast
    EndDate = v_rsRequest!CHARGE_DATE
    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    Set m_rsHolidays = l_obj.GetHolidaysBetween(StartDate, EndDate)
    Set m_rsRequest = CopyRecordset(v_rsRequest)
    
    l_blnValid = True
    v_rsRequest.Sort = "EMP_NUM, CHARGE_DATE"
    v_rsRequest.MoveFirst
    While Not v_rsRequest.EOF
        l_lngEmpNum = v_rsRequest!EMP_NUM
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_rsLockStatus = GetLockStatus(l_lngEmpNum, l_datChargeDate)
        If l_rsLockStatus <> "N" Then
            l_blnValid = False
            AppendLockStatusBRV l_lngEmpNum, l_datChargeDate, l_rsLockStatus
        End If
        If ValidRequest(v_rsRequest) Then
            If Not ValidFieldValues(v_rsRequest) Then
                l_blnValid = False
            End If
        Else
            l_blnValid = False
        End If
        
        v_rsRequest.MoveNext
    Wend
    
    PreValidate = l_blnValid
    
    
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "PreValidate"
    
    
End Function


Private Function AppendLockStatusBRV( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datChargeDate As Date, _
            ByVal v_strLockStatus As String _
            ) As Boolean
            
    Dim l_bolError As Boolean
    Dim lngKeyEmpNum As Long
    Dim l_datKeyChargeDate As Date
    Dim l_strErrDesc As String
    
    On Error GoTo AppendLockStatusBRV_EH
    
    lngKeyEmpNum = v_lngEmpNum
    l_datKeyChargeDate = v_datChargeDate
    
    Select Case v_strLockStatus
        Case "Y"
            l_strErrDesc = "The pay period is closed."
        Case "L"
            l_strErrDesc = "The department is locked by payroll."
    End Select
    
    l_bolError = AppendBusinessError(m_rsErrors, _
            trErrbusValidationError, _
            mstrCBUSINESSERR, _
            l_strErrDesc, _
            m_strCMODULENAME, _
            90, , , "EMP_NUM;CHARGE_DATE", _
            str(lngKeyEmpNum), _
            str(l_datKeyChargeDate))

    AppendLockStatusBRV = l_bolError
    Exit Function
    
AppendLockStatusBRV_EH:
    TrsRaiseError m_strCMODULENAME, "AppendLockStatusBRV"

End Function


Private Function ValidRequest( _
        ByRef r_rsReq As ADODB.Recordset _
        ) As Boolean
' answer true when the request is valid
    Dim l_strExcCd As String
    Dim l_strOverExcCd As String
    Dim l_blnValidationError As Boolean
    Dim l_lngDateDiff As Long
    Dim l_strFullShiftInd As String
    Dim l_strStartTime As String
    Dim l_strEndTime As String

On Error GoTo eh
    
    If Not IsValidRecStatus(r_rsReq) Then
        l_blnValidationError = True
    Else
        If r_rsReq!REC_STATUS = "D" Then
            ValidRequest = True
            Exit Function
        End If
    End If
    
    If IsNull(r_rsReq!START_TIME) Or IsEmpty(r_rsReq!START_TIME) Then
        l_strStartTime = ""
        r_rsReq!START_TIME = ""
    Else
        l_strStartTime = Trim(r_rsReq!START_TIME)
    End If
    
    If IsNull(r_rsReq!END_TIME) Or IsEmpty(r_rsReq!END_TIME) Then
        l_strEndTime = ""
        r_rsReq!END_TIME = ""
    Else
        l_strEndTime = Trim(r_rsReq!END_TIME)
    End If
    
    If IsNull(r_rsReq!FULL_SHIFT_EXC_IND) Or IsEmpty(r_rsReq!FULL_SHIFT_EXC_IND) Then
        If l_strStartTime = "" And l_strEndTime = "" Then
            l_strFullShiftInd = "Y"
        Else
            l_strFullShiftInd = "N"
        End If
    Else
        If l_strStartTime = "" And l_strEndTime = "" Then
            l_strFullShiftInd = "Y"
        Else
            If Trim(r_rsReq!FULL_SHIFT_EXC_IND) <> "Y" Then
                l_strFullShiftInd = "N"
            Else
                l_strFullShiftInd = "Y"
            End If
        End If
    End If
    r_rsReq!FULL_SHIFT_EXC_IND = l_strFullShiftInd
    r_rsReq!START_TIME = l_strStartTime
    r_rsReq!END_TIME = l_strEndTime

    If Not IsNumeric(r_rsReq!EMP_NUM) Or _
            r_rsReq!EMP_NUM = 0 Then
        AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Employee Id is invalid"
        l_blnValidationError = True
    End If

' Reject uninitialized dates and dates set to vbNull
    If Not IsDateValid(r_rsReq!CHARGE_DATE.Value) Then
        AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Charge Date is not a valid date"
        l_blnValidationError = True
    End If

    If l_strStartTime <> "" Then
        If Not IsTimeValid(r_rsReq!START_TIME) Then
            AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Start time is not valid"
            l_blnValidationError = True
        End If
    End If

    If l_strEndTime <> "" Then
        If Not IsTimeValid(r_rsReq!END_TIME) Then
            AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "End time is not valid"
            l_blnValidationError = True
        End If
    End If
    
    If Not ValidateOptionalRequestFields(r_rsReq) Then
        l_blnValidationError = True
    End If
    
    ValidRequest = Not l_blnValidationError
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidRequest"

End Function

Private Function ValidateOptionalRequestFields( _
        ByVal r_rsReq As ADODB.Recordset _
        ) As Boolean
' answer true when the request is valid
    Dim l_strExcCd As String
    Dim l_strOverExcCd As String
    Dim l_blnValidationError As Boolean

On Error GoTo eh
    '  Exception Code
'    This code will be used when the application can handle exceptions
    If Not IsNull(r_rsReq!EXCEPTION_CD) Then
        If Not IsAlphaNumeric(r_rsReq!EXCEPTION_CD) And r_rsReq!EXCEPTION_CD <> "" Then
            AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Exception code is not valid"
        End If
    End If
    
    '  OverrideException Code
    If Not IsNull(r_rsReq!OVER_EXC_CD) Then
       If Not IsAlphaNumeric(r_rsReq!OVER_EXC_CD) And r_rsReq!OVER_EXC_CD <> "" Then
        AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Override Exception code is not valid"
        l_blnValidationError = True
       End If
    End If
       
    '  Work Group
    If Not IsNull(r_rsReq!WG_NUM) Then
        If Not IsNumeric(r_rsReq!WG_NUM) Then
            AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Workgroup is not numeric"
            l_blnValidationError = True
        End If
    End If
       
    '  Job Number
    If Not IsNull(r_rsReq!JOB_NUM) Then
        If Not IsAlphaNumeric(r_rsReq!JOB_NUM) Then
            AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Job Number is invalid or missing"
            l_blnValidationError = True
        End If
    End If

    '  Component Shift Code
    If Not IsNull(r_rsReq!SHIFT_COMP) Then
        If Not IsAlphaNumeric(r_rsReq!SHIFT_COMP) Then
            AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Component Shift Code is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    '  Component Code
    If Not IsNull(r_rsReq!COMP_CD) Then
        If Not IsAlphaNumeric(r_rsReq!COMP_CD) Then
            AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Component Code is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    '  Relief Code
    If Not IsNull(r_rsReq!RELIEF_COMP) Then
        If Not IsAlphaNumeric(r_rsReq!RELIEF_COMP) Then
            AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Relief Code is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    ' Family Leave Indicator
    If Not IsNull(r_rsReq!FMLA_IND) Then
        If Not IsAlphaNumeric(r_rsReq!FMLA_IND) Then
            AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Family Leave Indicator is invalid or missing"
            l_blnValidationError = True
        End If
    End If
    
    ValidateOptionalRequestFields = Not l_blnValidationError
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateOptionalRequestFields"

End Function

            
Private Sub AppendEntryError( _
            ByVal v_rsReq As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDescription As String, _
            Optional ByVal v_strErrorSource As String)
            
            
            
    Dim l_bolError As Boolean
    Dim lngKeyEmpNum As Long
    Dim l_datKeyChargeDate As Date
    Dim lngKeyShiftNum As Long
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_strExcCd As String
    Dim l_strOverExcCd As String

    On Error GoTo AppendEntryError_EH
    
    lngKeyEmpNum = v_rsReq!EMP_NUM
    l_datKeyChargeDate = v_rsReq!CHARGE_DATE
    lngKeyShiftNum = v_rsReq!SHIFT_NUM
    
    If IsNull(v_rsReq!START_TIME) Or IsEmpty(v_rsReq!START_TIME) Then
        l_strStartTime = ""
    Else
        l_strStartTime = v_rsReq!START_TIME
    End If
    If IsNull(v_rsReq!END_TIME) Or IsEmpty(v_rsReq!END_TIME) Then
        l_strEndTime = ""
    Else
        l_strEndTime = v_rsReq!END_TIME
    End If
    If IsNull(v_rsReq!EXCEPTION_CD) Or IsEmpty(v_rsReq!EXCEPTION_CD) Then
        l_strExcCd = ""
    Else
        l_strExcCd = v_rsReq!EXCEPTION_CD
    End If
    
    If IsNull(v_rsReq!OVER_EXC_CD) Or IsEmpty(v_rsReq!OVER_EXC_CD) Then
        l_strOverExcCd = ""
    Else
        l_strOverExcCd = v_rsReq!OVER_EXC_CD
    End If
    
    
    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDescription, _
            m_strCMODULENAME, _
            90, , , "EMP_NUM;CHARGE_DATE;SHIFT_NUM,START_TIME,END_TIME,EXCEPTION_CD, OVER_EXC_CD", _
            str(lngKeyEmpNum), _
            str(l_datKeyChargeDate), str(lngKeyShiftNum), _
            l_strStartTime, l_strEndTime, l_strExcCd, l_strOverExcCd)

    Exit Sub
    
AppendEntryError_EH:
    TrsRaiseError m_strCMODULENAME, "AppendEntryError"

End Sub


Private Function ValidFieldValues( _
        ByRef r_rsReq As ADODB.Recordset _
        ) As Boolean
' answer true when the request is valid
    Dim l_strExcCd As String
    Dim l_strOverExcCd As String
    Dim l_blnIsValid As Boolean
    Dim l_lngDateDiff As Long
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strUnionCode As String
    Dim l_strJob As String
    Dim l_strValue As String
    Dim l_lngValue As Long
    Dim l_blnIsExective As Boolean
    Dim l_strBenefitsInd As String
    Dim l_strBaseJob As String

On Error GoTo eh

    If r_rsReq!REC_STATUS = "D" Then
        ValidFieldValues = True
        Exit Function
    End If

    l_blnIsValid = True
    l_lngEmpNum = r_rsReq!EMP_NUM
    l_datChargeDate = r_rsReq!CHARGE_DATE
    l_strUnionCode = GetEffectiveUnion(l_lngEmpNum, l_datChargeDate)
    l_strBaseJob = GetEffectiveJob(l_lngEmpNum, l_datChargeDate)
    l_blnIsExective = IsExecutive(l_lngEmpNum, l_datChargeDate)
    
    If IsNull(r_rsReq!EXCEPTION_CD) Or IsEmpty(r_rsReq!EXCEPTION_CD) Then
        l_strExcCd = ""
    Else
        l_strExcCd = Trim(r_rsReq!EXCEPTION_CD)
    End If
    If l_strExcCd = "" Then
        l_strExcCd = "ST"
        r_rsReq!EXCEPTION_CD = "ST"
    End If
    If IsNull(r_rsReq!OVER_EXC_CD) Or IsEmpty(r_rsReq!OVER_EXC_CD) Then
        l_strOverExcCd = ""
        r_rsReq!OVER_EXC_CD = ""
    Else
        l_strOverExcCd = Trim(r_rsReq!OVER_EXC_CD)
    End If
    If IsNull(r_rsReq!JOB_NUM) Or IsEmpty(r_rsReq!JOB_NUM) Then
        l_strJob = ""
    Else
        l_strJob = Trim(r_rsReq!JOB_NUM)
    End If
    If l_strJob <> "" Then
        If l_strBaseJob = l_strJob Then
' Callaway needs to pay vacation at the base job number
' an employee is upgrade so often, they just override it
'            l_strJob = ""
'            r_rsReq!JOB_NUM = ""
            If l_strOverExcCd = "JC" Then
                l_strOverExcCd = ""
                r_rsReq!OVER_EXC_CD = ""
            End If
            AppendWarning r_rsReq, trWarDiscardJobUpgrade, "Job number select is the same as the employee's base job (" & l_strBaseJob & ")."
        Else
            If l_strOverExcCd = "" Then
                r_rsReq!OVER_EXC_CD = "JC"
                l_strOverExcCd = "JC"
            End If
        End If
    End If
    
    If IsNull(r_rsReq!RELIEF_COMP) Or IsEmpty(r_rsReq!RELIEF_COMP) Then
        l_strValue = ""
        r_rsReq!RELIEF_COMP = ""
    Else
        l_strValue = Trim(r_rsReq!RELIEF_COMP)
    End If
    If Not IsValidVacationReliefCode(l_strUnionCode, l_datChargeDate, l_strValue) Then
        AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Relief component code is not valid"
        l_blnIsValid = False
    End If
    
    If IsNull(r_rsReq!COMP_CD) Or IsEmpty(r_rsReq!COMP_CD) Then
        l_strValue = ""
        r_rsReq!COMP_CD = ""
    Else
        l_strValue = Trim(r_rsReq!COMP_CD)
    End If
    If Not IsValidComponentCode(l_lngEmpNum, l_strUnionCode, l_datChargeDate, l_strValue, l_strExcCd) Then
        AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Component code is not valid"
        l_blnIsValid = False
    End If
    
    If IsNull(r_rsReq!SHIFT_COMP) Or IsEmpty(r_rsReq!SHIFT_COMP) Then
        l_strValue = ""
        r_rsReq!SHIFT_COMP = ""
    Else
        l_strValue = Trim(r_rsReq!SHIFT_COMP)
    End If
    If Not IsValidShiftComponentCode(r_rsReq, l_lngEmpNum, l_strUnionCode, l_datChargeDate, l_strValue, l_strExcCd) Then
        AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, mstrErrorDescription
        l_blnIsValid = False
    End If
    
    If Not IsValidJobNumber(l_strJob, l_strUnionCode, l_blnIsExective, l_strExcCd) Then
        AppendEntryError r_rsReq, mlngErrorNumber, mstrCBUSINESSERR, mstrErrorDescription
        l_blnIsValid = False
    End If

    r_rsReq!CASE_NUM = 0
    r_rsReq.Update
    l_strBenefitsInd = EmployeeBenefitPackage(l_lngEmpNum, l_datChargeDate)

    If Not IsValidExceptionCode(l_strExcCd, l_strUnionCode, l_strBenefitsInd) Then
        AppendEntryError r_rsReq, mlngErrorNumber, mstrCBUSINESSERR, mstrErrorDescription
        l_blnIsValid = False
    End If

    If Not IsValidOverrideExceptionCode(l_strOverExcCd, l_strExcCd, l_strJob, r_rsReq) Then
        AppendEntryError r_rsReq, mlngErrorNumber, mstrCBUSINESSERR, mstrErrorDescription
        l_blnIsValid = False
    End If
    
    If IsNull(r_rsReq!FMLA_IND) Or IsEmpty(r_rsReq!FMLA_IND) Then
        l_strValue = "N"
        r_rsReq!FMLA_IND = "N"
    Else
        l_strValue = Trim(r_rsReq!FMLA_IND)
        If l_strValue = "" Then
            l_strValue = "N"
            r_rsReq!FMLA_IND = "N"
        End If
    End If
    If Not IsValidFamilyLeaveIndicator(l_strValue, l_strExcCd) Then
        AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Family leave indicator is not valid"
        l_blnIsValid = False
    End If
    
    If IsNull(r_rsReq!WG_NUM) Or IsEmpty(r_rsReq!WG_NUM) Then
        l_lngValue = 0
        r_rsReq!WG_NUM = 0
    Else
        l_lngValue = Trim(r_rsReq!WG_NUM)
    End If
    If Not IsValidWorkGroup(l_lngValue, l_datChargeDate) Then
        AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Workgroup Number is not valid"
        l_blnIsValid = False
    End If
    
    If Not IsValidEntryDate(l_lngEmpNum, l_datChargeDate) Then
        AppendEntryError r_rsReq, trErrbusValidationError, mstrCBUSINESSERR, "Exception charge date is before employee's start date."
        l_blnIsValid = False
    End If
    
    If Not IsValidHolidayBanked(l_datChargeDate, l_strExcCd, l_strUnionCode) Then
        AppendEntryError r_rsReq, mlngErrorNumber, mstrCBUSINESSERR, mstrErrorDescription
        l_blnIsValid = False
    End If
    
    If Not IsValidHolidayDate(l_lngEmpNum, l_datChargeDate, l_strExcCd, l_strUnionCode, r_rsReq) Then
        AppendEntryError r_rsReq, mlngErrorNumber, mstrCBUSINESSERR, mstrErrorDescription
        l_blnIsValid = False
    End If
    
    ValidFieldValues = l_blnIsValid
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidFieldValues"

End Function


Private Function IsValidShiftComponentCode( _
        ByRef r_rsReq As ADODB.Recordset, _
        ByVal v_lngEmpNum As Long, _
        ByVal v_strUnionCode As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCode As String, _
        ByVal v_strExceptionCode _
        ) As Boolean

    
On Error GoTo eh

'Changes to IsEligibleForShiftComponent will cause changes to IsValidShiftComponentCode

'JIMB 05/06/2003
'Added Logic for CILCO.
'Change code to drop down logic to avoid nasty nested if statements
    Dim str As String
    
On Error GoTo eh

    If v_strCompCode = "" Then
        IsValidShiftComponentCode = True
        Exit Function
    End If
        
    If IsEligibleForShiftComponent(v_lngEmpNum, v_datEffDate, v_strExceptionCode) Then
        'Is the shift component valid for the union in the ttr_comp_amount table
        If IsShiftCompCode(v_strUnionCode, v_datEffDate, v_strCompCode) Then
            IsValidShiftComponentCode = True
        Else
            AppendWarning r_rsReq, trWarShiftCompRemoved, "Shift component code (" & v_strCompCode & ") is not valid for union (" & v_strUnionCode & "). Shift component was removed"
            r_rsReq!SHIFT_COMP = ""
            IsValidShiftComponentCode = True
        End If
        Exit Function
    End If

'AmerenCIPS Contract and AmerenCILCO local 8 employees are the only employees
'that get a shift component removed for certain exceptions
'The exception is ineligible for a shift component
    
    If IsAmerenCILCOLocal8Employee(v_lngEmpNum, v_datEffDate) And _
            Not IsExceptionIsEligibleForShiftComponent(v_strExceptionCode) Then
        AppendWarning r_rsReq, trWarShiftCompRemoved, "Exception (" & v_strExceptionCode & ") is not eligible for shift component (" & v_strCompCode & "). Shift component was removed"
        r_rsReq!SHIFT_COMP = ""
        IsValidShiftComponentCode = True
        Exit Function
    End If
    

'Must be a AmerenCIPS contract employee with non work time

    AppendWarning r_rsReq, trWarShiftCompRemoved, "Non-work time exception (" & v_strExceptionCode & ") is not eligible for shift component (" & v_strCompCode & "). Shift component was removed"
    r_rsReq!SHIFT_COMP = ""
    IsValidShiftComponentCode = True

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsValidShiftComponentCode"

End Function


Private Function IsShiftCompCode( _
        ByVal v_strUnionCode As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCode As String) As Boolean

On Error GoTo eh

    Dim l_obj As busCodeTables.cbsCodeTables

    Set l_obj = CreateObject("busCodeTables.cbsCodeTables")

    IsShiftCompCode = _
        l_obj.IsShiftCompCode( _
                v_strUnionCode, v_datEffDate, v_strCompCode)
    Set l_obj = Nothing
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsShiftCompCode"

End Function

Private Function IsEligibleForShiftComponent( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date, _
    ByVal v_strExcCd As String _
    ) As Boolean
' BLUE013, WHITE052
    On Error GoTo eh
    
    Dim l_blnIsExecutive As Boolean
    Dim l_blnIsContract As Boolean
    
    l_blnIsExecutive = IsExecutive(v_lngEmpNum, v_datDate)
    l_blnIsContract = Not l_blnIsExecutive
    
    If ((l_blnIsExecutive And _
            IsAmerenCIPS(v_lngEmpNum, v_datDate)) Or _
        (l_blnIsExecutive And _
            IsAmerenCIPSBenefitParticipant(v_lngEmpNum, v_datDate))) And _
        IsNonWorkTime(v_strExcCd) Then
        IsEligibleForShiftComponent = False
    Else
        If IsAmerenCILCOLocal8Employee(v_lngEmpNum, v_datDate) And _
            Not IsExceptionIsEligibleForShiftComponent(v_strExcCd) Then
            IsEligibleForShiftComponent = False
        Else
            IsEligibleForShiftComponent = True
        End If
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsEligibleForShiftComponent"
    
End Function

Private Function IsNonWorkTime( _
    ByVal v_strExceptionCode As String) As Boolean
    
    IsNonWorkTime = _
        IsStringInList(v_strExceptionCode, _
            Array("V", "E", "EF", "EB", "EC", "PB", "REL", "FH", "DF", _
            "PPT", "H", "J", "AD", "ED", "EDP", "AE", "JIE", "HR", "VH"))
End Function


Private Function IsExecutive( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As Boolean
    
    Dim l_datJobEffDate As Date
    Dim l_strUnionCd As String
    Dim l_strPriorUnionCd As String
    Dim l_strExcConInd As String
    Dim l_strPriorExcConInd As String
    Dim l_obj As busEmployee.cbsEmployee
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")

    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    
    l_datJobEffDate = m_rsEmployeeData!JOB_EFF_DATE
    l_strUnionCd = m_rsEmployeeData!UNION_CD
    l_strPriorUnionCd = m_rsEmployeeData!PRIOR_UNION_CD
    l_strExcConInd = m_rsEmployeeData!EXEC_CONTRACT_IND
    l_strPriorExcConInd = m_rsEmployeeData!PRIOR_EXEC_CON_IND
    
    IsExecutive = l_obj.IsExecutive( _
        l_datJobEffDate, l_strUnionCd, l_strPriorUnionCd, _
        l_strExcConInd, l_strPriorExcConInd, v_datDate)
        
    Set l_obj = Nothing

    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsExecutive"
    
End Function

Private Function GetEffectiveUnion( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date _
    ) As String
    
    
    Dim l_datJobEffDate As Date
    Dim l_strUnionCd As String
    Dim l_strPriorUnionCd As String
    Dim l_obj As busEmployee.cbsEmployee
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    
    l_datJobEffDate = m_rsEmployeeData!JOB_EFF_DATE
    l_strUnionCd = m_rsEmployeeData!UNION_CD
    l_strPriorUnionCd = m_rsEmployeeData!PRIOR_UNION_CD
    
    GetEffectiveUnion = l_obj.EmployeeUnion( _
        l_datJobEffDate, l_strUnionCd, l_strPriorUnionCd, v_datChargeDate)
    
    Set l_obj = Nothing
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetEffectiveUnion"
    
    
End Function

'//expose this method only for unit testing
Public Function IsAmerenCIPSTest( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As Boolean
On Error GoTo eh
        
    Set m_rsEmployeeData = GetEmployeeData(v_lngEmpNum)
    
    IsAmerenCIPSTest = IsAmerenCIPS(v_lngEmpNum, v_datDate)
    Set m_rsEmployeeData = Nothing
    
    Exit Function
eh:
    Set m_rsEmployeeData = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPSTest"
End Function

Private Function IsAmerenCIPS( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As Boolean
    
    Dim l_datDeptEffDate As Date
    Dim l_strDept As String
    Dim l_strPrevDept As String
    Dim l_strCompany As String
    Dim l_strPriorCompany  As String
    Dim l_strUnionCode As String
    Dim l_strPayrollGroupCd As String
    Dim l_obj As busEmployee.cbsEmployee
    
    On Error GoTo eh
    
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    
    l_datDeptEffDate = m_rsEmployeeData!DEPT_EFF_DATE
    l_strDept = m_rsEmployeeData!DEPT_NUM
    l_strPrevDept = m_rsEmployeeData!PREV_DEPT_NUM
    l_strCompany = m_rsEmployeeData!CMPY_CD
    l_strPriorCompany = m_rsEmployeeData!PREV_CMPY_CD
    l_strUnionCode = m_rsEmployeeData!UNION_CD
    l_strPayrollGroupCd = m_rsEmployeeData!PAYROLL_GROUP_CD
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    
    IsAmerenCIPS = l_obj.IsAmerenCIPSorCILCO(dbRules.TRSRuleCode.trIsAmerenCIPS, _
        l_datDeptEffDate, l_strDept, l_strPrevDept, _
        l_strCompany, l_strPriorCompany, v_datDate, l_strUnionCode, l_strPayrollGroupCd)
        
    Set l_obj = Nothing
    
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPS"
End Function


Private Function IsAmerenCIPSBenefitParticipant( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As Boolean
    
    Dim l_strBenPackage As String
    
    On Error GoTo eh
    
    l_strBenPackage = EmployeeBenefitPackage(v_lngEmpNum, v_datDate)
    
    IsAmerenCIPSBenefitParticipant = l_strBenPackage = "CIPS"

    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPSBenefitParticipant"
    
End Function

Private Function IsValidVacationReliefCode( _
    ByVal v_strUnionCode As String, _
    ByVal v_datEffDate As Date, _
    ByVal v_strCompCode As String _
    ) As Boolean
'BLUE022, WHITE071
    Dim str As String
    
On Error GoTo eh
    
    If v_strCompCode = "" Then
        IsValidVacationReliefCode = True
    Else
        IsValidVacationReliefCode = _
            IsVacationReliefCode( _
                v_strUnionCode, v_datEffDate, v_strCompCode)
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidVacationReliefCode"

End Function

Private Function IsVacationReliefCode( _
        ByVal v_strUnionCode As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCode As String) As Boolean

On Error GoTo eh
    
    Dim l_obj As busCodeTables.cbsCodeTables

    Set l_obj = CreateObject("busCodeTables.cbsCodeTables")

    IsVacationReliefCode = _
        l_obj.IsVacationReliefCode( _
                v_strUnionCode, v_datEffDate, v_strCompCode)
                
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsVacationReliefCode"

End Function

Private Function IsValidComponentCode( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_strUnionCode As String, _
            ByVal v_datEffDate As Date, _
            ByVal v_strCompCode, _
            ByVal v_strExcCd As String _
            ) As Boolean
            
'WHITE073, WHITE052
    Dim str As String

    On Error GoTo eh
    If v_strCompCode = "" Then
        IsValidComponentCode = True
    Else
        If IsEligibleForShiftComponent(v_lngEmpNum, v_datEffDate, v_strExcCd) Then
            IsValidComponentCode = _
                IsComponentCode(v_strUnionCode, v_datEffDate, v_strCompCode)
        Else
            IsValidComponentCode = False
        End If
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidComponentCode"

End Function

Private Function IsComponentCode( _
        ByVal v_strUnionCode As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCode As String) As Boolean

On Error GoTo eh
    Dim l_obj As busCodeTables.cbsCodeTables

    Set l_obj = CreateObject("busCodeTables.cbsCodeTables")
    
    IsComponentCode = _
        l_obj.IsComponentCode( _
                v_strUnionCode, v_datEffDate, v_strCompCode)
                
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsComponentCode"

End Function

Private Sub AppendWarning( _
            ByVal v_rsReq As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            Optional ByVal v_strErrorDescription As String, _
            Optional ByRef vcolDescriptionArgs As Collection)
    Dim l_strErrDesc As String
    
    If v_strErrorDescription = "" Then
        l_strErrDesc = TrsGetErrDescription(LongAppErr(v_lngErrorNumber), vcolDescriptionArgs)
    Else
        l_strErrDesc = v_strErrorDescription
    End If
    AppendEntryError v_rsReq, v_lngErrorNumber, mstrCBUSINESSWAR, l_strErrDesc

End Sub

Private Sub SetError( _
            ByVal vlngErrorNumber As Long, _
            Optional ByVal vstrErrorDescription As String)
' vlngErrorNumber should be vbObjectError + appErrorNumber

    mlngErrorNumber = vlngErrorNumber
    mstrErrorDescription = vstrErrorDescription
    
End Sub

Private Function IsValidJobNumber( _
    ByVal v_strJob As String, _
    ByVal v_strUnion As String, _
    ByVal v_blnIsExective As Boolean, _
    ByVal v_strExcCd As String _
    ) As Boolean

' BLUE021
' RED151
On Error GoTo eh

    If v_strJob = "" Then
        IsValidJobNumber = True
        Exit Function
    End If

    If Not JobNumberIsValid(v_strJob) Then
        SetError trErrbusValidationError, "Job number is not valid"
        IsValidJobNumber = False
        Exit Function
    End If
    
    If Not IsJobUpgradeValid(v_strJob, v_strUnion, v_blnIsExective) Then
        SetError trErrbusValidationError, "Job upgrade is not valid for the employee's union"
        IsValidJobNumber = False
        Exit Function
    End If
    
    If Not IsEligibleForJobUpgrade(v_strUnion, v_strExcCd) Then
        SetError trErrbusValidationError, "Job upgrade is not valid for the exception code"
        IsValidJobNumber = False
        Exit Function
    End If
    
    
    IsValidJobNumber = True
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsValidJobNumber"
    
End Function


Private Function IsJobUpgradeValid( _
    ByVal v_strJobNum As String, _
    ByVal v_strEmpsUnion As String, _
    ByVal v_blnIsExective As Boolean _
    ) As Boolean
' RED151
' v_strJobNum is the job number changed from the base job.
'
    Dim l_strUnionCd As String
    Dim l_blnIsValid As Boolean
    
On Error GoTo eh
    l_strUnionCd = Left(v_strJobNum, 2)
    If v_blnIsExective Then
        If v_strEmpsUnion = l_strUnionCd Then
            IsJobUpgradeValid = True
        Else
            IsJobUpgradeValid = False
        End If
    Else
        If ExecutiveContractIndicatorForUnion(l_strUnionCd) = _
           ExecutiveContractIndicatorForUnion(v_strEmpsUnion) Then
            IsJobUpgradeValid = True
        Else
            IsJobUpgradeValid = False
        End If
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsJobUpgradeValid"
    
End Function

Private Function JobNumberIsValid( _
    ByVal v_strJob As String _
    ) As Boolean
    
 On Error GoTo eh
    Dim l_obj As busCodeTables.cbsCodeTables

    Set l_obj = CreateObject("busCodeTables.cbsCodeTables")
    
    JobNumberIsValid = _
        l_obj.JobNumberIsValid(v_strJob)
                
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "JobNumberIsValid"
    
    
End Function

Private Function ExecutiveContractIndicatorForUnion( _
                    ByVal v_strUnionCd As String) As String
                    
On Error GoTo eh
    Dim l_obj As dbCodeTables.cdCodeTables

    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    
    ExecutiveContractIndicatorForUnion = _
        l_obj.UnionExecutiveContractIndicator(v_strUnionCd)
        
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExecutiveContractIndicatorForUnion"
End Function

'TODO: obsolete IsValidCaseNumber
'Private Function IsValidCaseNumber( _
'    ByVal v_lngCaseNum As Long, _
'    ByVal v_strExceptionCode As String _
'    ) As Boolean
'
'On Error GoTo eh
'
'    IsValidCaseNumber = True
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "IsValidCaseNumber"
'
'End Function

'TODO: obsolete IsCaseNumberRequired
'Private Function IsCaseNumberRequired( _
'    ByVal v_strExceptionCode As String _
'    ) As Boolean
'
'    On Error GoTo eh
'
'    IsCaseNumberRequired = False
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "IsCaseNumberRequired"
'
'End Function

Private Function CaseGroupFor( _
            ByVal v_strExcCd As String) As String
' BLUE082
On Error GoTo eh
    Dim l_obj As dbCodeTables.cdCodeTables

    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    
    CaseGroupFor = _
        l_obj.ExceptionCaseGroup(v_strExcCd)
        
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CaseGroupFor"
    
End Function

Private Function IsValidExceptionCode( _
    ByVal v_strExceptionCode As String, _
    ByVal v_strEmployeeUnion As String, _
    ByVal v_strBenefitsInd As String _
    ) As Boolean

'BLUE094, RED038

    Dim l_strErrorMsg As String
    Dim l_blnIsValid As Boolean
    Dim l_strExecConInd As String
    Dim l_strBenefitsInd As String

On Error GoTo eh
    
    If v_strExceptionCode = "" Then
        SetError trErrbusValidationError, "Invalid Exception Code. Exception code was not entered"
        IsValidExceptionCode = False
        Exit Function
    End If
       
    l_strExecConInd = ExecutiveContractIndicatorForUnion(v_strEmployeeUnion)
    
    If CanExceptionBeScheduled(v_strExceptionCode) Then
        l_blnIsValid = _
            IsValidExceptionCodeFor( _
                v_strExceptionCode, _
                l_strExecConInd, _
                v_strBenefitsInd)
    Else
            SetError trErrbusValidationError, "Invalid Exception Code. Exception code can not be scheduled"
            IsValidExceptionCode = False
            Exit Function
    End If
                    
    If l_blnIsValid Then
        ' Continue
    Else
        If l_strExecConInd = "E" Then
            l_strErrorMsg = "executive"
        Else
            l_strErrorMsg = "contract"
        End If
'        SetError trErrbusValidationError, "Exception Code is not valid for " & l_strErrorMsg & " employee."
        SetError trErrbusValidationError, "Exception Code is not valid."
        l_blnIsValid = False
    End If
              
    IsValidExceptionCode = l_blnIsValid

    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsValidExceptionCode"

End Function


Private Function CanExceptionBeScheduled( _
            ByVal v_strExcCd As String _
            ) As Boolean
' BLUE082
On Error GoTo eh
    Dim l_obj As dbCodeTables.cdCodeTables
    Dim l_strSchedInd As String

    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    
    l_strSchedInd = _
        l_obj.ExceptionSchedInd(v_strExcCd)
            
    If l_strSchedInd = "Y" Then
        CanExceptionBeScheduled = True
    Else
        CanExceptionBeScheduled = False
    End If
    
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CanExceptionBeScheduled"
    
End Function


Private Function IsValidExceptionCodeFor( _
            ByVal v_strExcCd As String, _
            ByVal v_strExecConInd As String, _
            ByVal v_strBenefitsInd As String _
            ) As Boolean
' BLUE094, WHITE043
            
On Error GoTo eh
    Dim l_obj As busCodeTables.cbsCodeTables


    Set l_obj = CreateObject("busCodeTables.cbsCodeTables")

    IsValidExceptionCodeFor = _
        l_obj.ExceptionCodeIsValidFor(v_strExcCd, v_strExecConInd, v_strBenefitsInd)
        
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidExceptionCodeFor"
        
End Function


Private Function IsValidOverrideExceptionCode( _
    ByVal v_strOverExcCd As String, _
    ByVal v_strExceptionCode As String, _
    ByVal v_strJob As String, _
    ByRef r_rsReq As ADODB.Recordset _
    ) As Boolean

On Error GoTo eh

    Dim l_blnIsValid As Boolean
    Dim l_strExcType As String
    
'''''    l_strExcType = GetExceptionType(v_strExceptionCode)
'''''    If l_strExcType <> "P" And v_strOverExcCd = "" Then
'''''        SetError trErrbusValidationError, "Must enter an override with this (" & v_strExceptionCode & ") exception code"
'''''        IsValidOverrideExceptionCode = False
'''''        Exit Function
'''''    End If
    

    If v_strOverExcCd <> "" Then
        l_blnIsValid = _
            IsOverrideExceptionCode(v_strOverExcCd)
        If Not l_blnIsValid Then
            IsValidOverrideExceptionCode = False
            SetError trErrbusValidationError, "Override exception code is not valid"
            Exit Function
        End If
    Else
        IsValidOverrideExceptionCode = True
    End If
    
    If v_strOverExcCd = "JC" Then
        If v_strJob = "" Then
            r_rsReq!OVER_EXC_CD = ""
            AppendWarning r_rsReq, vbObjectError + 20000, "Removed JC override code because, the job upgrade(JC) is not valid without a job change"
            IsValidOverrideExceptionCode = True
            Exit Function
        End If
    End If
    
    IsValidOverrideExceptionCode = True
    
    

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsValidOverrideExceptionCode"

End Function

Private Function IsOverrideExceptionCode( _
            ByVal v_strExcCd As String) As Boolean
    Dim l_strExcType As String
            
On Error GoTo eh
    Dim l_obj As busCodeTables.cbsCodeTables

    Set l_obj = CreateObject("busCodeTables.cbsCodeTables")
    
    IsOverrideExceptionCode = _
        l_obj.ExceptionIsOverride(v_strExcCd)
                
    Set l_obj = Nothing

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsOverrideExceptionCode"
        
End Function

Private Function IsValidFamilyLeaveIndicator( _
    ByVal v_strFMLA As String, _
    ByVal v_strExcCd As String _
    ) As Boolean
' WHITE067
On Error GoTo eh
    If Not IsYorN(v_strFMLA) Then
        IsValidFamilyLeaveIndicator = False
        Exit Function
    End If

    If v_strFMLA = "Y" Then
        IsValidFamilyLeaveIndicator = IsEligibleForFamilyLeave(v_strExcCd)
    Else
        IsValidFamilyLeaveIndicator = True
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsValidFamilyLeaveIndicator"

End Function

Private Function IsEligibleForFamilyLeave( _
            ByVal v_strExcCd As String) As Boolean
'WHITE066
On Error GoTo eh

    Dim l_obj As busCodeTables.cbsCodeTables

    Set l_obj = CreateObject("busCodeTables.cbsCodeTables")
    
    IsEligibleForFamilyLeave = _
        l_obj.ExceptionIsEligibleForFamilyLeave(v_strExcCd)
                
    Set l_obj = Nothing

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsEligibleForFamilyLeave"
    
End Function


Private Function IsValidWorkGroup( _
            ByVal v_lngWorkGroupNum As Long, _
            ByVal v_datEffDate As Date) As Boolean
' WHITE088
    Dim lobjWGMgr As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh
    
    If v_lngWorkGroupNum = 0 Then
        IsValidWorkGroup = True
        Exit Function
    End If
    
    Set lobjWGMgr = CreateObject("dbWorkGroup.cdWorkGroup")
    IsValidWorkGroup = lobjWGMgr.IsValidWorkGroup(v_lngWorkGroupNum, v_datEffDate)
    
    Set lobjWGMgr = Nothing
    Exit Function
    
eh:
    Set lobjWGMgr = Nothing
    TrsRaiseError m_strCMODULENAME, "IsValidWorkGroup"

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
    TrsRaiseError m_strCMODULENAME, "IsDateValid"
End Function


Private Function CreateActualTimeDeleteRS() As ADODB.Recordset
    Dim l_rsRequest As ADODB.Recordset

    On Error GoTo eh

    Set l_rsRequest = New ADODB.Recordset
    With l_rsRequest
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "HOURS", adDecimal, adFldUpdatable
      .Fields.Item("HOURS").Precision = 3
      .Fields.Item("HOURS").NumericScale = 1
      .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "CASE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "FMLA_IND", adChar, 1, adFldUpdatable
      .Fields.Append "INDEX_TIME", adChar, 26, adFldUpdatable
      .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockBatchOptimistic
      .Open
    End With

    Set CreateActualTimeDeleteRS = l_rsRequest
    Set l_rsRequest = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CreateActualTimeDeleteRS"

End Function


Private Sub AddActualTimeDelete( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date)


    Dim l_strFilter As String

    On Error GoTo eh
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
        " AND CHARGE_DATE = " & str(v_datChargeDate)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsDeleteActualTime.Filter = l_strFilter
    
    If RecordsExist(m_rsDeleteActualTime) Then
        m_rsDeleteActualTime.Filter = ""
        Exit Sub
    End If
    
    
    m_rsActualTime.MoveFirst
    While Not m_rsActualTime.EOF
        m_rsDeleteActualTime.AddNew
        m_rsDeleteActualTime!EMP_NUM = m_rsActualTime!EMP_NUM
        m_rsDeleteActualTime!CHARGE_DATE = m_rsActualTime!CHARGE_DATE
        m_rsDeleteActualTime!Hours = CDec(m_rsActualTime!Hours)
        m_rsDeleteActualTime!START_TIME = m_rsActualTime!START_TIME
        m_rsDeleteActualTime!END_TIME = m_rsActualTime!END_TIME
        m_rsDeleteActualTime!EXCEPTION_CD = m_rsActualTime!EXCEPTION_CD
        m_rsDeleteActualTime!OVER_EXC_CD = m_rsActualTime!OVER_EXC_CD
        m_rsDeleteActualTime!CASE_NUM = 0
        m_rsDeleteActualTime!INDEX_TIME = m_rsActualTime!INDEX_TIME
        m_rsDeleteActualTime!FMLA_IND = m_rsActualTime!FAMILY_LEAVE_IND
        m_rsDeleteActualTime!WG_NUM = m_rsActualTime!WG_NUM
        m_rsDeleteActualTime.Update
        m_rsActualTime.MoveNext
    Wend
    
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "AddActualTimeDelete"
    
End Sub

Private Function GetExceptionType( _
            ByVal v_strExcCd As String _
            ) As String

On Error GoTo eh
    Dim l_obj As dbCodeTables.cdCodeTables


    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    
    GetExceptionType = _
        l_obj.ExceptionType(v_strExcCd)
    
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetExceptionType"
    
End Function


Private Sub ModifyRequest( _
    ByRef r_rsRequest As ADODB.Recordset)
    
'This routine will change the rec status of all deletes
' from a "D" to a "Z".  This is so all deletes are processed first
'It will also turn a change request into a delete and an insert

    Dim l_strTimeStamp As String
    Dim l_lngShiftNum As Long
    Dim l_strStart As String
    Dim l_strEnd As String
    Dim l_rsChanges As ADODB.Recordset


    On Error GoTo eh

    If Not RecordsExist(r_rsRequest) Then
        Exit Sub
    End If
    
    r_rsRequest.Filter = "REC_STATUS = 'D'"
    If RecordsExist(r_rsRequest) Then
        r_rsRequest.MoveFirst
        While Not r_rsRequest.EOF
            r_rsRequest!REC_STATUS = "Z"
            r_rsRequest.Update
            r_rsRequest.MoveNext
        Wend
    End If
    

    
'Change Updates to a Delete and an insert
    r_rsRequest.Filter = "REC_STATUS = 'C'"
    If RecordsExist(r_rsRequest) Then
        Set l_rsChanges = CopyRecordset(r_rsRequest)
        l_rsChanges.MoveFirst
        While Not l_rsChanges.EOF
            l_strTimeStamp = Trim(l_rsChanges!CREATED_TIMESTAMP)
            m_rsExceptions.Filter = "CREATED_TIMESTAMP = '" & l_strTimeStamp & "'"
            r_rsRequest.Filter = "CREATED_TIMESTAMP = '" & l_strTimeStamp & "'"
            r_rsRequest!REC_STATUS = "I"
            r_rsRequest.Update
            AddDeleteRequest r_rsRequest
            l_rsChanges.MoveNext
        Wend
    End If

    r_rsRequest.Filter = "REC_STATUS = 'I'"
    If RecordsExist(r_rsRequest) Then
        r_rsRequest.MoveFirst
        While Not r_rsRequest.EOF
            If IsNull(r_rsRequest!SHIFT_NUM) Or IsEmpty(r_rsRequest!SHIFT_NUM) Then
                DetermineShift r_rsRequest
            Else
                l_lngShiftNum = r_rsRequest!SHIFT_NUM
                If IsNull(r_rsRequest!START_TIME) Or _
                    IsEmpty(r_rsRequest!START_TIME) Then
                    l_strStart = ""
                Else
                    l_strStart = Trim(r_rsRequest!START_TIME)
                End If
                If IsNull(r_rsRequest!END_TIME) Or _
                    IsEmpty(r_rsRequest!END_TIME) Then
                    l_strEnd = ""
                Else
                    l_strEnd = Trim(r_rsRequest!END_TIME)
                End If
                If l_lngShiftNum <= 0 Or l_lngShiftNum > 2 _
                    Or l_strStart <> "" Or l_strEnd <> "" Then
                    DetermineShift r_rsRequest
                Else
                    FillInMissingTime r_rsRequest
                End If
            End If
            DayOffProcess r_rsRequest
            r_rsRequest.Update
            r_rsRequest.MoveNext
        Wend
    End If
    r_rsRequest.Filter = ""
    r_rsRequest.MoveFirst
   
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "ModifyRequest"
    
    
End Sub

Private Function ProcessScheduledException( _
    ByVal v_rsRequest As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strRecStatus As String
    Dim l_blnOk As Boolean
    
    On Error GoTo eh
    
    l_strRecStatus = v_rsRequest!REC_STATUS
    
    Select Case l_strRecStatus
        Case "Z" 'Delete
            l_blnOk = DeleteScheduledException(v_rsRequest)
        Case "I" 'Insert
            l_blnOk = InsertScheduledException(v_rsRequest)
        Case Else
        
    End Select
    
    ProcessScheduledException = l_blnOk
    
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ProcessScheduledException"
    
End Function


Private Function DeleteScheduledException( _
    ByVal v_rsRequest As ADODB.Recordset _
    ) As Boolean
    
    'JIMB 01/08/2004
    'Logic was added to display warnings when the VCO or VEF exceptions were deleted
    'This is for the Local 148 (OA) vacation in lieu of sick leave rules
    'In the main logic, a change request is converted to an insert and delete request
    'The BRV meesages need to know if the original request was a change request
    ' Scenarios from Exception Scheduling
    '   1. VC0 –  User deleted a VCO exception
    '   2. VCO – User changed VCO exception to another exception
    '   3. VCO – User change the charge date on a VCO exception and it remained a VCO exception
    '   4. VEF – User deleted a VEF exception
    '   5. VEF – User changed VEF to VE.  Also changed the charge date.
    '   6. VEF – User changed VEF to VE, but did not change the charge date.
    '   7. VEF – User changed VEF to another exception code that was not VE or VEF
    '   8. VEF – User changed the charge date on a VEF exception and it remained a VEF exception

    ' Scenarios 5 and 6 will not display a warning.
    'VEF is an exception that marks a frozen vacation week
    'The employee must take vacation that week or may use them for
    ' vacation in lieu of illness (VE).

    
    Dim l_strTimeStamp As String
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_lngWGNum As Long
    Dim l_blnEMPRV As Boolean
    Dim l_intLocCd As Integer
    Dim l_blnDelete As Boolean
    Dim l_strExcCd As String
    Dim l_strWarningDesc As String
    Dim l_strReqExcCd As String
    
    On Error GoTo eh
    
    l_strTimeStamp = v_rsRequest!CREATED_TIMESTAMP
    
    m_rsExceptions.Filter = "CREATED_TIMESTAMP = '" & l_strTimeStamp & "'"
    If Not RecordsExist(m_rsExceptions) Then
        AppendEntryError v_rsRequest, _
            trErrbusInvalidDeleteEntryNotFound, mstrCBUSINESSERR, _
            "Invalid delete request.  Exception to be deleted does not exist."
        l_blnDelete = False
    Else
        AppendAccrualBankUpdate "D", "S", m_rsAccrualBankReq, m_rsExceptions
        l_lngEmpNum = m_rsExceptions!EMP_NUM
        l_datChargeDate = m_rsExceptions!CHARGE_DATE
        l_blnEMPRV = IsEmployeeEMPRV(l_lngEmpNum, l_datChargeDate)
        If l_blnEMPRV Then
            l_lngWGNum = m_rsExceptions!WG_NUM
            l_intLocCd = EmployeeWorkgroupLocation(m_rsEmployeeWorkGroups, _
                l_lngWGNum, l_lngEmpNum, l_datChargeDate)
            m_colEMPRVTrans.Add BuildEMPRVTransaction(m_rsExceptions, l_intLocCd, "D")
        End If
'        m_rsExceptions.Delete
        l_blnDelete = True
    End If
    
    DeleteScheduledException = l_blnDelete
    If Not l_blnDelete Then
        Exit Function
    End If
    
    l_strExcCd = m_rsExceptions!EXCEPTION_CD
    
    m_rsOriginalRequest.Filter = "CREATED_TIMESTAMP = '" & l_strTimeStamp & "'"
    If IsNull(m_rsOriginalRequest!EXCEPTION_CD) Or _
      IsEmpty(m_rsOriginalRequest!EXCEPTION_CD) Then
        l_strReqExcCd = ""
    Else
        l_strReqExcCd = Trim(m_rsOriginalRequest!EXCEPTION_CD)
    End If
    Select Case l_strExcCd
        Case "VCO"
            l_strWarningDesc = ""
            m_rsOriginalRequest.Filter = "CREATED_TIMESTAMP = '" & l_strTimeStamp & "'"
            If m_rsOriginalRequest!REC_STATUS = "D" Then
        '1
'                l_strWarningDesc = "VCO (Vacation Carry Over) was deleted.  Please verify that all VCO exceptions are scheduled."
                l_strWarningDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                    & " is part of a partial week of vacation that has been carried over from the previous year. " _
                    & " This partial week of vacation cannot be further split and must be scheduled in a block."
            Else
                If l_strReqExcCd <> "VCO" Then
        '2
'                    l_strWarningDesc = "VCO (Vacation Carry Over) was changed to " & l_strReqExcCd & "." _
                        & "  Please verify that all VCO exceptions are scheduled."
                l_strWarningDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                    & " is part of a partial week of vacation that has been carried over from the previous year. " _
                    & " This partial week of vacation cannot be further split and must be scheduled in a block."
                Else
                    If m_rsOriginalRequest!CHARGE_DATE <> _
                      m_rsExceptions!CHARGE_DATE Then
        '3
'                        l_strWarningDesc = "VCO (Vacation Carry Over) was moved from " & _
                            str(m_rsExceptions!CHARGE_DATE) & " to " & str(m_rsOriginalRequest!CHARGE_DATE) & "." _
                        & "  Please verify that all VCO exceptions are scheduled in the same week."
                        l_strWarningDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                            & " is part of a partial week of vacation that has been carried over from the previous year. " _
                            & " This partial week of vacation cannot be further split and must be scheduled in a block."
                    
                    End If
                End If
            End If
        Case "VEF"
            l_strWarningDesc = ""
            m_rsOriginalRequest.Filter = "CREATED_TIMESTAMP = '" & l_strTimeStamp & "'"
            If m_rsOriginalRequest!REC_STATUS = "D" Then
        '4
'                l_strWarningDesc = "VEF (Vacation for Illness Frozen) was deleted. Please verify that all VEF exceptions are scheduled."
                l_strWarningDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                    & " is part of a third week of vacation that has been split for vacation in lieu of sick leave. " _
                    & " The remaining hours of vacation in this week cannot be moved except to be taken as vacation in lieu of sick leave up to this date."
            Else
                Select Case l_strReqExcCd
                    Case "VE"
        '5, 6       'No warning needed
                    Case "VEF"
                        If m_rsOriginalRequest!CHARGE_DATE <> _
                          m_rsExceptions!CHARGE_DATE Then
        '8
'                            l_strWarningDesc = "VEF (Vacation for Illness Frozen) was moved from " & _
                                str(m_rsExceptions!CHARGE_DATE) & " to " & str(m_rsOriginalRequest!CHARGE_DATE) & "." _
                                & "  Please verify that all VEF exceptions are scheduled in the same week."
                            l_strWarningDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                                & " is part of a third week of vacation that has been split for vacation in lieu of sick leave. " _
                                & " The remaining hours of vacation in this week cannot be moved except to be taken as vacation in lieu of sick leave up to this date."
                        End If
                    Case Else
        '7
'                        l_strWarningDesc = "VEF (Vacation for Illness Frozen) was changed to " & l_strReqExcCd & "." _
                             & "  Please verify that all VEF exceptions are scheduled."
                        l_strWarningDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                            & " is part of a third week of vacation that has been split for vacation in lieu of sick leave. " _
                            & " The remaining hours of vacation in this week cannot be moved except to be taken as vacation in lieu of sick leave up to this date."
                End Select
            End If
        Case Else
            l_strWarningDesc = ""
    End Select
    
    If l_strWarningDesc <> "" Then
        AppendWarning m_rsOriginalRequest, vbObjectError + 20000, l_strWarningDesc
    End If

    m_rsExceptions.Delete
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DeleteScheduledException"
   
    
End Function


Private Function InsertScheduledException( _
    ByVal v_rsRequest As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_strExcCd As String
    Dim l_decShiftHours As Variant
    Dim l_strShiftStartTime As String
    Dim l_strShiftEndTime As String
    Dim l_datShiftStartDate As Date
    Dim l_datShiftEndDate As Date
    Dim l_strFullShiftInd As String
    Dim l_strFilter As String
    Dim l_datShiftStartDateTime As Date
    Dim l_datShiftEndDateTime As Date
    Dim l_datExcStartDateTime As Date
    Dim l_datExcEndDateTime As Date
    Dim l_decExceptionMins As Variant
    Dim l_datStartDateTime As Date
    Dim l_datEndDateTime As Date
    Dim l_decHours As Variant
    Dim l_decOffSetHours As Variant
    Dim l_intEndHour As Integer
    Dim l_intHoursIncr As Integer
    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strUnionCode As String
    
    On Error GoTo eh
    
    l_strExcCd = Trim(v_rsRequest!EXCEPTION_CD)
    l_strFullShiftInd = v_rsRequest!FULL_SHIFT_EXC_IND
    l_lngEmpNum = v_rsRequest!EMP_NUM
    l_datChargeDate = v_rsRequest!CHARGE_DATE
    l_strStartTime = v_rsRequest!START_TIME
    l_strEndTime = v_rsRequest!END_TIME
    
    l_strFilter = FilterSchedule(v_rsRequest)
    
    If Not RecordsExist(m_rsSchedules) Then
        InsertScheduledException = False
        AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSERR, _
            "Can not find schedule for exception."
        Exit Function
    End If
    
    l_decShiftHours = CDec(m_rsSchedules!SHIFT_HOURS)
    
 'On off days, set the full shift indicator when the start and end times are
 ' the default values.
 'This will allow the user to change HR to HB and HB to HR, without removing
 'the start and end times.
    If l_decShiftHours = 0 And l_strStartTime = "08:00" And _
        l_strEndTime = "16:00" Then
        l_strFullShiftInd = "Y"
    End If
    
    
    
    m_rsExceptions.Filter = l_strFilter
    
'Full Shift Exceptions
    If l_strFullShiftInd = "Y" Then
        If RecordsExist(m_rsExceptions) Then
            AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSERR, _
                "Can not schedule a full shift exception, because an exception already exists."
            InsertScheduledException = False
            Exit Function
        Else
            If l_decShiftHours = 0 Then
                If CanBeScheduledOnOffDay(l_strExcCd) Then
                    If Not FlexPolicyApplies(l_lngEmpNum, l_datChargeDate) Then
                        AddException v_rsRequest, l_strExcCd, CDec(0), 8
                        InsertScheduledException = True
                    Else
                        If l_strExcCd = "HR" Then
                            l_strUnionCode = GetEffectiveUnion(l_lngEmpNum, l_datChargeDate)
                            If IsValidHolidayBanked(l_datChargeDate, "HB", l_strUnionCode) Then
                                AddException v_rsRequest, "HB", CDec(0), 8
                                AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSWAR, _
                                    "Can not schedule HR it was changed to HB"
                            Else
                                    AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSWAR, _
                                    "Can not schedule HR exception because of Flex Policy."
                            End If
                        Else
                            AddException v_rsRequest, l_strExcCd, CDec(0), 8
                        End If
                        InsertScheduledException = True
                    End If
                Else
                    If l_strExcCd = "H" Then
                        If Not FlexPolicyApplies(l_lngEmpNum, l_datChargeDate) Then
                            AddException v_rsRequest, "HR", CDec(0), 8
                            AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSWAR, _
                                "The H (Holiday) exception was changed to HR (Holiday on day of Rest) because the employee is schedule off on the charge date."
                            InsertScheduledException = True
                        Else
                            l_strUnionCode = GetEffectiveUnion(l_lngEmpNum, l_datChargeDate)
                            If IsValidHolidayDate(l_lngEmpNum, l_datChargeDate, l_strExcCd, l_strUnionCode, v_rsRequest) Then
                                AddException v_rsRequest, "HB", CDec(0), 8
                                AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSWAR, _
                                    "The H (Holiday) exception was changed to HB (Holiday Banked) because the employee is schedule off on the Holiday and is not allowed HR."
                                InsertScheduledException = True
                            Else
                                AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSWAR, _
                                    "Can not schedule exception on an off day."
                                InsertScheduledException = True
                            End If
                        End If
                    Else
                        AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSWAR, _
                            "Can not schedule exception on an off day."
                        InsertScheduledException = True
                        Exit Function
                    End If
                End If
            Else
                If CanBeScheduledOnWorkDay(l_strExcCd) Then
                    AddException v_rsRequest, l_strExcCd, CDec(0), l_decShiftHours
                    CheckScheduleChangePayments l_strExcCd, l_lngEmpNum, l_datChargeDate
                    InsertScheduledException = True
                Else
                    AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSWAR, _
                        "Can not schedule exception on an working shift."
                    InsertScheduledException = True
                    Exit Function
                End If
            End If
        End If
        Exit Function
    End If
    

'Partial Shift Exceptions

    l_intHoursIncr = PartialHourMinutes(v_rsRequest!EMP_NUM)

    If l_decShiftHours = 0 Then
        AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSWAR, _
            "Can not schedule partial exception on an off day."
        InsertScheduledException = True
        Exit Function
    Else
        If CanBeScheduledOnWorkDay(l_strExcCd) Then
            If l_strExcCd = "PBL" Then
            AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSERR, _
                "Can not schedule PBL as a partial shift."
            InsertScheduledException = False
            Exit Function
            End If
        Else
            AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSWAR, _
                "Can not schedule exception on an working shift."
            InsertScheduledException = True
            Exit Function
        End If
    End If
    
'    l_strStartTime = v_rsRequest!START_TIME
'    l_strEndTime = v_rsRequest!END_TIME
    l_strExcCd = v_rsRequest!EXCEPTION_CD
    l_strShiftStartTime = m_rsSchedules!SHIFT_START_TIME
    l_datShiftStartDate = m_rsSchedules!SHIFT_START_DATE
    
    l_datShiftStartDateTime = DateTimeFrom(l_datShiftStartDate, l_strShiftStartTime)
    l_datExcStartDateTime = DateTimeFrom(l_datShiftStartDate, l_strStartTime)
    If DateDiff("s", l_datExcStartDateTime, l_datShiftStartDateTime) > 0 Then
        l_datExcStartDateTime = DateAdd("d", 1, l_datExcStartDateTime)
    End If
    
    If l_strEndTime = "00:00" Then
        If l_strStartTime <> "00:00" Then
            l_strEndTime = "24:00"
        End If
    End If
    l_intEndHour = CInt(Left(l_strEndTime, 2))
    
    If l_strEndTime < l_strStartTime Then
        l_intEndHour = l_intEndHour + 24
    End If
    l_decExceptionMins = CDec(l_intEndHour) * 60 + CDec(Right(l_strEndTime, 2)) _
        - CDec(Left(l_strStartTime, 2)) * 60 - CDec(Right(l_strStartTime, 2))
    l_datExcEndDateTime = DateAdd("n", l_decExceptionMins, l_datExcStartDateTime)
    
    l_strShiftEndTime = m_rsSchedules!SHIFT_END_TIME
    l_datShiftEndDate = m_rsSchedules!SHIFT_END_DATE
    l_datShiftEndDateTime = DateTimeFrom(l_datShiftEndDate, l_strShiftEndTime)

    If Not TimePeriodsIntersect2( _
        l_datShiftStartDateTime, l_datShiftEndDateTime, _
        l_datExcStartDateTime, l_datExcEndDateTime) Then
        AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSERR, _
                "Can not schedule the exception, because it does not fall within the selected shift."
            InsertScheduledException = False
        Exit Function
    End If
    
    'Does the new exception overlapp existing exceptions
    m_rsExceptions.Sort = "HOURS_OFFSET_NUM"
    While Not m_rsExceptions.EOF
        DetermineExceptionDateTimes l_datStartDateTime, l_datEndDateTime, _
            m_rsSchedules, m_rsExceptions
        'Need to check if they overrlapp
        If TimePeriodsIntersect2( _
            l_datStartDateTime, l_datEndDateTime, _
            l_datExcStartDateTime, l_datExcEndDateTime) Then
            AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSERR, _
                    "Can not schedule the exception because it overlaps an existing entry."
            InsertScheduledException = False
            Exit Function
        End If
        m_rsExceptions.MoveNext
    Wend

    DetermineExceptionHours l_decHours, l_decOffSetHours, _
        l_datExcStartDateTime, l_datExcEndDateTime, m_rsSchedules
        
    If DateDiff("s", l_datShiftEndDateTime, l_datExcEndDateTime) > 0 Then
        AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSERR, _
        "Can not schedule the exception because it ends after the shift ends."
        InsertScheduledException = False
    ElseIf DateDiff("s", l_datExcStartDateTime, l_datExcEndDateTime) = 0 Then
        AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSERR, _
        "Can not schedule the exception because the start time equals the end time(" & Format(l_datExcStartDateTime, "hh:mm") & ")"
        InsertScheduledException = False
    ElseIf l_decHours = 0 Then
        AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSERR, _
        "Can not schedule the exception because the hours equal zero after adjusted for lunch."
        InsertScheduledException = False
    ElseIf Not IsHoursIncrementValid(l_decHours, l_intHoursIncr) Then
        AppendEntryError v_rsRequest, trErrbusShiftInvalidStartTime, mstrCBUSINESSERR, _
        "Can not schedule the exception because the hours(" & CStr(l_decHours) _
            & ") is not an increment of " & CStr(l_intHoursIncr) & " minutes."
        InsertScheduledException = False
    Else
        InsertScheduledException = True
        AddException v_rsRequest, l_strExcCd, l_decOffSetHours, l_decHours
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "InsertScheduledException"
End Function

Private Sub DetermineExceptionDateTimes( _
    ByRef r_datExcStartDateTime As Date, _
    ByRef r_datExcEndDateTime As Date, _
    ByVal v_rsAdjustedBaseSched As ADODB.Recordset, _
    ByVal v_rsExceptionSched As ADODB.Recordset)
    
    
    Dim l_decExceptionHours As Variant
    Dim l_decExceptionOffSetHours As Variant
    Dim l_decShiftHours As Variant
    Dim l_strShiftStartTime As String
    Dim l_strShiftEndTime As String
    Dim l_datShiftStartDate As Date
    Dim l_datShiftStartDateTime As Date
    Dim l_datShiftEndDate As Date
    Dim l_intExceptionHours As Integer
    Dim l_intExceptionMinutes As Integer
    Dim l_datExceptionStartDateTime As Date
    Dim l_datExceptionEndDateTime As Date
    Dim l_strLunchStartTime As String
    Dim l_strLunchEndTime As String
    Dim l_decHoursOfLunch As Variant
    Dim l_decHoursBeforeLunch As Variant
    Dim l_datLunchStartDate As Date
    Dim l_datLunchStartDateTime As Date
    Dim l_datLunchEndDate As Date
    Dim l_datLunchEndDateTime As Date
    Dim l_decExceptionPhysicalHours As Variant
    

    On Error GoTo eh

    l_datShiftStartDate = v_rsAdjustedBaseSched!SHIFT_START_DATE
    l_strShiftStartTime = Trim(v_rsAdjustedBaseSched!SHIFT_START_TIME)
    l_strLunchStartTime = Trim(v_rsAdjustedBaseSched!SHIFT_LUNCH_START_TIME)
    l_strLunchEndTime = Trim(v_rsAdjustedBaseSched!SHIFT_LUNCH_END_TIME)
    If l_strLunchStartTime = "" Or l_strLunchEndTime = "" Then
        l_strLunchStartTime = l_strShiftStartTime
        l_strLunchEndTime = l_strShiftStartTime
    End If
    l_decExceptionOffSetHours = CDec(v_rsExceptionSched!HOURS_OFFSET_NUM)
    l_decExceptionHours = v_rsExceptionSched!Hours
    If l_strShiftStartTime > l_strLunchStartTime Then
        l_datLunchStartDate = DateAdd("d", 1, l_datShiftStartDate)
    Else
        l_datLunchStartDate = l_datShiftStartDate
    End If
'There are currently no lunches that cross midnight (2002-04-12)
    If l_strLunchStartTime > l_strLunchEndTime Then
        l_datLunchEndDate = DateAdd("d", 1, l_datLunchStartDate)
    Else
        l_datLunchEndDate = l_datLunchStartDate
    End If
    l_datShiftStartDateTime = DateTimeFrom(l_datShiftStartDate, l_strShiftStartTime)
    l_datLunchStartDateTime = DateTimeFrom(l_datLunchStartDate, l_strLunchStartTime)
    l_decHoursBeforeLunch = CDec(DateDiff("n", l_datShiftStartDateTime, l_datLunchStartDateTime) / 60)
    l_datLunchEndDateTime = DateTimeFrom(l_datLunchEndDate, l_strLunchEndTime)
    l_decHoursOfLunch = CDec(DateDiff("n", l_datLunchStartDateTime, l_datLunchEndDateTime) / 60)
    
    
    If l_decExceptionOffSetHours >= l_decHoursBeforeLunch Then
    'Exception starts after lunch
        l_datExceptionStartDateTime = DateAdd("n", ((l_decExceptionOffSetHours + l_decHoursOfLunch) * 60), l_datShiftStartDateTime)
        l_datExceptionEndDateTime = DateAdd("n", l_decExceptionHours * 60, l_datExceptionStartDateTime)
    Else
    'Exception starts before lunch
        l_datExceptionStartDateTime = DateAdd("n", (l_decExceptionOffSetHours * 60), l_datShiftStartDateTime)
        'Does exception end before lunch
        If (l_decExceptionHours + l_decExceptionOffSetHours) <= l_decHoursBeforeLunch Then
        'Yes have the correct hours to calcualate end time
            l_decExceptionPhysicalHours = l_decExceptionHours
        Else
            'NO, must add lunch time to calculate end time
            l_decExceptionPhysicalHours = l_decHoursOfLunch + l_decExceptionHours
        End If
        l_datExceptionEndDateTime = _
                DateAdd("n", l_decExceptionPhysicalHours * 60, _
                          l_datExceptionStartDateTime)
    End If

    r_datExcStartDateTime = l_datExceptionStartDateTime
    r_datExcEndDateTime = l_datExceptionEndDateTime
    
    
    Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "DetermineExceptionDateTimes"
    
    
End Sub


Private Sub AddException( _
    ByVal v_rsRequest As ADODB.Recordset, _
    ByVal v_strExcCd As String, _
    ByVal v_decOffset As Variant, _
    ByVal v_decHours As Variant)
    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_lngWGNum As Long
    Dim l_blnEMPRV As Boolean
    Dim l_intLocCd As Integer
    
    On Error GoTo eh
    l_lngEmpNum = v_rsRequest!EMP_NUM
    l_datChargeDate = v_rsRequest!CHARGE_DATE
    m_rsExceptions.AddNew
    m_rsExceptions!EMP_NUM = l_lngEmpNum
    m_rsExceptions!CHARGE_DATE = l_datChargeDate
    m_rsExceptions!SHIFT_NUM = v_rsRequest!SHIFT_NUM
    m_rsExceptions!HOURS_OFFSET_NUM = v_decOffset
    m_rsExceptions!Hours = v_decHours
    m_rsExceptions!EXCEPTION_CD = v_strExcCd
    m_rsExceptions!SCHED_CD = m_rsSchedules!SCHED_CD
    m_rsExceptions!SCHED_SRC_CD = m_rsSchedules!SCHED_SRC_CD
    m_rsExceptions!CREATED_EMP_NUM = m_lngUserNum
    m_rsExceptions!LAST_MOD_EMP_NUM = m_lngUserNum
    
    If IsNull(v_rsRequest!OVER_EXC_CD) Or IsEmpty(v_rsRequest!OVER_EXC_CD) Then
        m_rsExceptions!OVER_EXC_CD = ""
    Else
        m_rsExceptions!OVER_EXC_CD = Trim(v_rsRequest!OVER_EXC_CD)
    End If
    
    If IsNull(v_rsRequest!FMLA_IND) Or IsEmpty(v_rsRequest!FMLA_IND) Then
        m_rsExceptions!FMLA_IND = "N"
    Else
        m_rsExceptions!FMLA_IND = UCase(Trim(v_rsRequest!FMLA_IND))
    End If

    m_rsExceptions!CASE_NUM = 0
    If IsNull(v_rsRequest!WG_NUM) Or IsEmpty(v_rsRequest!WG_NUM) Then
        l_lngWGNum = 0
    Else
        l_lngWGNum = v_rsRequest!WG_NUM
    End If
    m_rsExceptions!WG_NUM = l_lngWGNum
    
    If IsNull(v_rsRequest!JOB_NUM) Or IsEmpty(v_rsRequest!JOB_NUM) Then
        m_rsExceptions!JOB_NUM = ""
    Else
        m_rsExceptions!JOB_NUM = UCase(Trim(v_rsRequest!JOB_NUM))
    End If
    
    If IsNull(v_rsRequest!RELIEF_COMP) Or IsEmpty(v_rsRequest!RELIEF_COMP) Then
        m_rsExceptions!RELIEF_COMP = ""
    Else
        m_rsExceptions!RELIEF_COMP = UCase(Trim(v_rsRequest!RELIEF_COMP))
    End If
    
    If IsNull(v_rsRequest!COMP_CD) Or IsEmpty(v_rsRequest!COMP_CD) Then
        m_rsExceptions!COMP_CD = ""
    Else
        m_rsExceptions!COMP_CD = UCase(Trim(v_rsRequest!COMP_CD))
    End If
    
    If IsNull(v_rsRequest!SHIFT_COMP) Or IsEmpty(v_rsRequest!SHIFT_COMP) Then
        m_rsExceptions!SHIFT_COMP = ""
    Else
        m_rsExceptions!SHIFT_COMP = UCase(Trim(v_rsRequest!SHIFT_COMP))
    End If
    
    m_rsExceptions.Update
    
    AppendAccrualBankUpdate "I", "S", m_rsAccrualBankReq, m_rsExceptions
    
    l_blnEMPRV = IsEmployeeEMPRV(l_lngEmpNum, l_datChargeDate)
    If l_blnEMPRV Then
        l_intLocCd = EmployeeWorkgroupLocation(m_rsEmployeeWorkGroups, _
            l_lngWGNum, l_lngEmpNum, l_datChargeDate)
        m_colEMPRVTrans.Add BuildEMPRVTransaction(m_rsExceptions, l_intLocCd, "I")
    End If
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AddException"
End Sub
    
    
    
Private Sub DetermineExceptionHours( _
    ByRef r_decExcHours As Variant, _
    ByRef r_decOffsetHours As Variant, _
    ByVal v_datExcStartDateTime As Date, _
    ByVal v_datExcEndDateTime As Date, _
    ByVal v_rsAdjustedBaseSched As ADODB.Recordset)
    
    
    
    Dim l_decExceptionHours As Variant
    Dim l_decExceptionOffSetHours As Variant
    Dim l_decShiftHours As Variant
    Dim l_strShiftStartTime As String
    Dim l_strShiftEndTime As String
    Dim l_datShiftStartDate As Date
    Dim l_datShiftStartDateTime As Date
    Dim l_datShiftEndDate As Date
    Dim l_strLunchStartTime As String
    Dim l_strLunchEndTime As String
    Dim l_decHoursOfLunch As Variant
    Dim l_decHoursBeforeLunch As Variant
    Dim l_datLunchStartDate As Date
    Dim l_datLunchStartDateTime As Date
    Dim l_datLunchEndDate As Date
    Dim l_datLunchEndDateTime As Date
    

    On Error GoTo eh

    l_decShiftHours = CDec(v_rsAdjustedBaseSched!SHIFT_HOURS)
    l_datShiftStartDate = v_rsAdjustedBaseSched!SHIFT_START_DATE
    l_strShiftStartTime = Trim(v_rsAdjustedBaseSched!SHIFT_START_TIME)
    l_strLunchStartTime = Trim(v_rsAdjustedBaseSched!SHIFT_LUNCH_START_TIME)
    l_strLunchEndTime = Trim(v_rsAdjustedBaseSched!SHIFT_LUNCH_END_TIME)
    If l_strLunchStartTime = "" Or l_strLunchEndTime = "" Then
        l_strLunchStartTime = l_strShiftStartTime
        l_strLunchEndTime = l_strShiftStartTime
    End If
'''''    l_decExceptionOffSetHours = CDec(v_rsExceptionSched!HOURS_OFFSET_NUM)
'''''    l_decExceptionHours = v_rsExceptionSched!Hours
    If l_strShiftStartTime > l_strLunchStartTime Then
        l_datLunchStartDate = DateAdd("d", 1, l_datShiftStartDate)
    Else
        l_datLunchStartDate = l_datShiftStartDate
    End If
'There are currently no lunches that cross midnight (2002-04-12)
    If l_strLunchStartTime > l_strLunchEndTime Then
        l_datLunchEndDate = DateAdd("d", 1, l_datLunchStartDate)
    Else
        l_datLunchEndDate = l_datLunchStartDate
    End If
    l_datShiftStartDateTime = DateTimeFrom(l_datShiftStartDate, l_strShiftStartTime)
    l_datLunchStartDateTime = DateTimeFrom(l_datLunchStartDate, l_strLunchStartTime)
    l_decHoursBeforeLunch = CDec(DateDiff("n", l_datShiftStartDateTime, l_datLunchStartDateTime) / 60)
    l_datLunchEndDateTime = DateTimeFrom(l_datLunchEndDate, l_strLunchEndTime)
    l_decHoursOfLunch = CDec(DateDiff("n", l_datLunchStartDateTime, l_datLunchEndDateTime) / 60)
    l_decExceptionOffSetHours = CDec(DateDiff("n", l_datShiftStartDateTime, v_datExcStartDateTime) / 60)
    l_decExceptionHours = CDec(DateDiff("n", v_datExcStartDateTime, v_datExcEndDateTime) / 60)

    r_decOffsetHours = l_decExceptionOffSetHours
    If l_decExceptionOffSetHours >= l_decHoursBeforeLunch Then
    'Exception starts after lunch
        r_decExcHours = l_decExceptionHours
        If l_decExceptionOffSetHours = l_decHoursBeforeLunch Then
            r_decExcHours = l_decExceptionHours - l_decHoursOfLunch
        Else
            If l_decExceptionOffSetHours >= _
                (l_decHoursBeforeLunch + l_decHoursOfLunch) Then
                r_decOffsetHours = l_decExceptionOffSetHours - l_decHoursOfLunch
            Else
                r_decExcHours = l_decExceptionHours - (l_decHoursBeforeLunch + l_decHoursOfLunch - l_decExceptionOffSetHours)
                r_decOffsetHours = l_decHoursBeforeLunch
'                r_decExcHours = l_decShiftHours - l_decHoursBeforeLunch
            End If
        End If
        Exit Sub
    End If
       
    'Exception starts before lunch
    'Does exception end before lunch
    If (l_decExceptionHours + l_decExceptionOffSetHours) <= l_decHoursBeforeLunch Then
    'Yes have the correct hours to calcualate end time
        r_decExcHours = l_decExceptionHours
        Exit Sub
    End If
    
    'NO, must subtract lunch time
    'Does Exception End after Lunch
    If (l_decExceptionHours + l_decExceptionOffSetHours) _
        >= (l_decHoursOfLunch + l_decHoursBeforeLunch) Then
        r_decExcHours = l_decExceptionHours - l_decHoursOfLunch
    Else
        r_decExcHours = l_decHoursBeforeLunch - l_decExceptionOffSetHours
    End If

    
    Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "DetermineExceptionHours"
    
    
End Sub

Private Function CanBeScheduledOnOffDay( _
    ByVal v_strExcCd As String _
    ) As Boolean
    
    On Error GoTo eh
    
    Select Case Trim(v_strExcCd)
        Case "HR", "HB"
            CanBeScheduledOnOffDay = True
        Case Else
            CanBeScheduledOnOffDay = False
    End Select
    Exit Function
eh:

    TrsRaiseError m_strCMODULENAME, "CanBeScheduledOnOffDay"
    
End Function

Private Function SaveExceptionsDB() As Boolean

    Dim l_obj As dbSchedException.cdSchedExceptionTrans

    On Error GoTo eh

    m_rsExceptions.Filter = ""
    Set l_obj = CreateObject("dbSchedException.cdSchedExceptionTrans")

    SaveExceptionsDB = l_obj.SaveExceptions(m_rsExceptions, m_lngUserNum)
    
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "SaveExceptionsDB"
End Function

Private Sub AppendAccrualBankUpdate( _
            ByVal v_strRecStatus As String, _
            ByVal v_strRecSource As String, _
            ByRef r_rsAccBankUpdates As ADODB.Recordset, _
            ByRef r_rsExceptions As ADODB.Recordset)
    
On Error GoTo eh
   
    With r_rsAccBankUpdates
        .AddNew
        !EMP_NUM = r_rsExceptions!EMP_NUM
        !CHARGE_DATE = r_rsExceptions!CHARGE_DATE
        !EXCEPTION_CD = r_rsExceptions!EXCEPTION_CD
        !OVER_EXC_CD = r_rsExceptions!OVER_EXC_CD
        !Hours = r_rsExceptions!Hours
        !BANK_CODE = ExceptionBankCode(r_rsExceptions!EXCEPTION_CD)
        !FMLA_IND = r_rsExceptions!FMLA_IND
        !REC_STATUS = v_strRecStatus
        !CASE_NUM = 0
        !CASE_GROUP = ""
        If v_strRecSource = "A" Then
            !ACTUAL_TIME_IND = "Y"
        Else
            !ACTUAL_TIME_IND = "N"
        End If
        .Update
    End With

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendAccrualBankUpdate"
End Sub

Private Function CreateAccrualBankMgrRecordset() As ADODB.Recordset
    Dim l_rsRequest As ADODB.Recordset

On Error GoTo eh

    Set l_rsRequest = New ADODB.Recordset
    With l_rsRequest
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("HOURS").Precision = 3
      .Fields.Item("HOURS").NumericScale = 1
      .Fields.Append "BANK_CODE", adChar, 2, adFldUpdatable
      .Fields.Append "FMLA_IND", adChar, 1, adFldUpdatable
      .Fields.Append "REC_STATUS", adChar, 1, adFldUpdatable
      .Fields.Append "CASE_GROUP", adChar, 4, adFldUpdatable
      .Fields.Append "CASE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "ACTUAL_TIME_IND", adChar, 1, adFldUpdatable

      .CursorLocation = adUseClientBatch
      .CursorType = adOpenStatic
      .LockType = adLockBatchOptimistic
      .Open
    End With
    Set CreateAccrualBankMgrRecordset = l_rsRequest
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CreateAccrualBankMgrRecordset"
End Function


Private Function IsEmployeeEMPRV( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As Boolean
    
    
    Dim l_datDeptEffDate As Date
    Dim l_strDept As String
    Dim l_strPrevDept As String
    Dim l_strInterfaceCD As String
    Dim l_strPrevInterfaceCD As String
    Dim l_obj As busEmployee.cbsEmployee
    
On Error GoTo eh
    
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    l_datDeptEffDate = m_rsEmployeeData!DEPT_EFF_DATE
    l_strDept = m_rsEmployeeData!DEPT_NUM
    l_strPrevDept = m_rsEmployeeData!PREV_DEPT_NUM
    l_strInterfaceCD = Trim(m_rsEmployeeData!INTERFACE_TO_CD)
    l_strPrevInterfaceCD = Trim(m_rsEmployeeData!INTERFACE_TO_CD)
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")

    IsEmployeeEMPRV = l_obj.IsEmployeeEMPRV( _
        l_datDeptEffDate, l_strDept, l_strPrevDept, l_strInterfaceCD, l_strPrevInterfaceCD, v_datDate)

    Set l_obj = Nothing


    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsEmployeeEMPRV"
End Function


Private Function EmployeeWorkgroupLocation( _
                ByRef r_rsEmpWorkgroups As ADODB.Recordset, _
                ByVal v_lngWgNum As Long, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date) As Long
'If WG_NUM is > 0 on exception, look up LOC_CD for it and answer it.
' Otherwise determine the primary workgroup on the charge date and look up LOC_CD.
' cdWorkgroup caches workgroup data in SPM, as requested.  May need to do some
' fast initialization of the SPM cache earlier in ProcessExceptions to speed things up here.
    Dim l_cdWorkgroup As dbWorkGroup.cdWorkGroup

On Error GoTo eh

    If v_lngWgNum > 0 Then
        Set l_cdWorkgroup = CreateObject("dbWorkGroup.cdWorkGroup")
        EmployeeWorkgroupLocation = _
            CLng(l_cdWorkgroup.LocationCode(v_lngWgNum))
        Set l_cdWorkgroup = Nothing
    Else
        PositionOnWorkgroup r_rsEmpWorkgroups, v_lngEmpNum, v_datChargeDate
        If RecordsExist(r_rsEmpWorkgroups) Then
            EmployeeWorkgroupLocation = r_rsEmpWorkgroups!LOC_CD
        Else
            EmployeeWorkgroupLocation = 0
        End If
    End If

    Exit Function

eh:
    Set l_cdWorkgroup = Nothing
    TrsRaiseError m_strCMODULENAME, "EmployeeWorkgroupLocation"
End Function

Private Function GetEmployeeWorkGroups( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStart As Date, _
    ByVal l_datEnd As Date _
    ) As ADODB.Recordset
    
    Dim l_obj As dbWorkGroup.cdWorkGroup
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeeWorkGroups = _
        l_obj.GetEmployeeWorkGroups( _
            v_aryEmployees, v_datStart, l_datEnd, True)
    Set l_obj = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeWorkGroups"
End Function

Private Function DetermineEMPRVEmployees() As Collection

    Dim l_colEmployees As Collection
    Dim l_strEmployee As String
    Dim l_strInterfaceCD As String
    Dim l_strPrevInterfaceCD As String

On Error GoTo eh

    Set l_colEmployees = New Collection

    If Not RecordsExist(m_rsEmployeeData) Then
        Set DetermineEMPRVEmployees = l_colEmployees
        Exit Function
    End If

    m_rsEmployeeData.Sort = "EMP_NUM"
    m_rsEmployeeData.MoveFirst
    While Not m_rsEmployeeData.EOF
        l_strEmployee = Trim(str(m_rsEmployeeData!EMP_NUM))
        l_strInterfaceCD = Trim(m_rsEmployeeData!INTERFACE_TO_CD)
        l_strPrevInterfaceCD = Trim(m_rsEmployeeData!INTERFACE_TO_CD)
        If l_strInterfaceCD = "E" Or l_strPrevInterfaceCD = "E" Then
                l_colEmployees.Add l_strEmployee
        End If
        m_rsEmployeeData.MoveNext
    Wend
    m_rsEmployeeData.Sort = ""
    m_rsEmployeeData.MoveFirst
    
    Set DetermineEMPRVEmployees = l_colEmployees
    Set l_colEmployees = Nothing
    Exit Function

eh:
    Set l_colEmployees = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineEMPRVEmployees"
End Function


Private Function BuildEMPRVTransaction( _
    ByRef r_rsExceptions As ADODB.Recordset, _
    ByVal v_intLocationCd As Integer, _
    ByVal v_strTransType As String _
    ) As Variant
    
    
    Dim l_varEMPRVTran(0 To 6) As Variant

    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strDept As String
    

On Error GoTo eh

    l_lngEmpNum = r_rsExceptions!EMP_NUM
    l_datChargeDate = r_rsExceptions!CHARGE_DATE
    l_strDept = GetDepartment(l_lngEmpNum, l_datChargeDate)
    l_varEMPRVTran(etEMPNUM) = r_rsExceptions!EMP_NUM
        l_varEMPRVTran(etDEPTNUM) = l_strDept
    If Trim(r_rsExceptions!OVER_EXC_CD) <> "" Then
        l_varEMPRVTran(etEXCCD) = Trim(r_rsExceptions!OVER_EXC_CD)
    Else
        l_varEMPRVTran(etEXCCD) = Trim(r_rsExceptions!EXCEPTION_CD)
    End If
    l_varEMPRVTran(etCHGDATE) = r_rsExceptions!CHARGE_DATE
    l_varEMPRVTran(etHOURS) = r_rsExceptions!Hours
    l_varEMPRVTran(etLOCCD) = v_intLocationCd
    l_varEMPRVTran(etTRANCD) = v_strTransType

    BuildEMPRVTransaction = l_varEMPRVTran
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "BuildEMPRVTransaction"
End Function


Private Sub SaveEMPRV()

    Dim l_varEMPRVTrans As Variant
    Dim l_obj As dbMessageDispatcher.cdMessageDispatcher
    
    On Error GoTo eh

    l_varEMPRVTrans = ArrayFromCollection(m_colEMPRVTrans)
    Set l_obj = _
        CreateObject("dbMessageDispatcher.cdMessageDispatcher")
    l_obj.AddEMPRVTransactions l_varEMPRVTrans
    Set l_obj = Nothing
    Exit Sub
    
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveEMPRV"

End Sub

Private Function GetDepartment( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As String

    Dim l_datDeptEffDate As Date
    Dim l_strDept As String
    Dim l_strPrevDept As String
    Dim l_obj As busEmployee.cbsEmployee
    
    On Error GoTo eh
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    
    l_datDeptEffDate = m_rsEmployeeData!DEPT_EFF_DATE
    l_strDept = m_rsEmployeeData!DEPT_NUM
    l_strPrevDept = m_rsEmployeeData!PREV_DEPT_NUM

    Set l_obj = CreateObject("busEmployee.cbsEmployee")

    GetDepartment = l_obj.EmployeeDepartment( _
        l_datDeptEffDate, l_strDept, l_strPrevDept, v_datDate)

    Set l_obj = Nothing

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetDepartment"
    
End Function


Private Sub PositionOnWorkgroup( _
                ByRef r_rsEmpWorkgroups As ADODB.Recordset, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date)
    Dim l_strFilter As String
    
On Error GoTo eh
    l_strFilter = "EMP_NUM = " & v_lngEmpNum & " AND " & _
        "EFF_START_DATE <= " & _
        "#" & Format(v_datChargeDate, "MM/DD/YYYY") & "#" & _
        " AND EFF_END_DATE >= " & _
        "#" & Format(v_datChargeDate, "MM/DD/YYYY") & "#"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsEmpWorkgroups.Filter = l_strFilter
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "PositionOnWorkgroup"
End Sub

Private Function ExceptionBankCode( _
            ByVal v_strExcCd As String) As String
    Dim l_cdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_cdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    ExceptionBankCode = _
        l_cdCodeTables.ExceptionBankCode(Trim(v_strExcCd))
    Set l_cdCodeTables = Nothing
    Exit Function

eh:
    Set l_cdCodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionBankCode"
End Function

Private Function UpdateAccrualBanks() As Boolean

    Dim l_obj As busAccrualBankMgr.cbsAccrualBankMgr
    
    On Error GoTo eh
    
    m_rsAccrualBankReq.Filter = ""
    If Not RecordsExist(m_rsAccrualBankReq) Then
        UpdateAccrualBanks = True
        Exit Function
    End If
    
    Set l_obj = CreateObject("busAccrualBankMgr.cbsAccrualBankMgr")
    
    UpdateAccrualBanks = _
        l_obj.UpdateBanksForExceptionSched( _
            m_rsAccrualBankReq, Nothing, m_lngUserNum)
    
    Set l_obj = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "UpdateAccrualBanks"

End Function

Private Function IsValidRecStatus( _
    ByRef r_rsReq As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strRecStatus As String
    Dim l_strTimeStamp As String

On Error GoTo eh

    If IsNull(r_rsReq!REC_STATUS) Or IsEmpty(r_rsReq!REC_STATUS) Then
        l_strRecStatus = ""
    Else
        l_strRecStatus = r_rsReq!REC_STATUS
    End If
    
    If IsNull(r_rsReq!CREATED_TIMESTAMP) Or IsEmpty(r_rsReq!CREATED_TIMESTAMP) Then
        l_strTimeStamp = ""
    Else
        l_strTimeStamp = Trim(r_rsReq!CREATED_TIMESTAMP)
    End If
    
    m_rsExceptions.Filter = "CREATED_TIMESTAMP = '" & l_strTimeStamp & "'"
    
    Select Case l_strRecStatus
        Case "D"
            If Not RecordsExist(m_rsExceptions) Then
                AppendEntryError r_rsReq, _
                    trErrbusInvalidDeleteEntryNotFound, mstrCBUSINESSERR, _
                    "Invalid delete request.  Exception to be deleted does not exist."
                IsValidRecStatus = False
            Else
                IsValidRecStatus = True
            End If
        
        Case "C"
            If Not RecordsExist(m_rsExceptions) Then
                AppendEntryError r_rsReq, _
                    trErrbusInvalidOPChangeEntryNotFound, mstrCBUSINESSERR, _
                    "Invalid change request.  Exception to be changed does not exist."
                IsValidRecStatus = False
            Else
                IsValidRecStatus = True
            End If
        
        Case "I"
            IsValidRecStatus = True
        Case Else
            AppendEntryError r_rsReq, _
                trErrbusInvalidRecordStatus, mstrCAPPLICATIONERR, _
                    "Invalid request.  Please retry or contact support."
            IsValidRecStatus = False
        
    End Select
      
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsValidRecStatus"

End Function

Private Sub AddDeleteRequest( _
    ByRef r_rsRequest As ADODB.Recordset)
    
    On Error GoTo eh
    
    r_rsRequest.AddNew
    r_rsRequest!EMP_NUM = m_rsExceptions!EMP_NUM
    r_rsRequest!CHARGE_DATE = m_rsExceptions!CHARGE_DATE
    r_rsRequest!SHIFT_NUM = m_rsExceptions!SHIFT_NUM
    r_rsRequest!EXCEPTION_CD = m_rsExceptions!EXCEPTION_CD
    r_rsRequest!OVER_EXC_CD = m_rsExceptions!OVER_EXC_CD
    r_rsRequest!CREATED_TIMESTAMP = m_rsExceptions!CREATED_TIMESTAMP
    r_rsRequest!REC_STATUS = "Z"
    r_rsRequest.Update

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AddDeleteRequest"
    
    
End Sub

Private Function DeleteActualTime() As Boolean

    Dim l_obj As dbSaveTimecard.cdSaveTimecard
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_blnEMPRV As Boolean
    Dim l_intLocCd As Integer
    Dim l_lngWGNum As Long
    
    
    On Error GoTo eh
    
    m_rsDeleteActualTime.Filter = ""
    If Not RecordsExist(m_rsDeleteActualTime) Then
        DeleteActualTime = True
        Exit Function
    End If
    
    Set l_obj = CreateObject("dbSaveTimecard.cdSaveTimecard")
    
    DeleteActualTime = _
        l_obj.DeleteActualTime(m_rsDeleteActualTime)
    Set l_obj = Nothing
        
    m_rsDeleteActualTime.MoveFirst
    While Not m_rsDeleteActualTime.EOF
        AppendAccrualBankUpdate "D", "A", m_rsAccrualBankReq, m_rsDeleteActualTime
        AppendActualTimeWarning
        l_lngEmpNum = m_rsDeleteActualTime!EMP_NUM
        l_datChargeDate = m_rsDeleteActualTime!CHARGE_DATE
        l_blnEMPRV = IsEmployeeEMPRV(l_lngEmpNum, l_datChargeDate)
        If l_blnEMPRV Then
            l_lngWGNum = m_rsDeleteActualTime!WG_NUM
            l_intLocCd = EmployeeWorkgroupLocation(m_rsEmployeeWorkGroups, _
                l_lngWGNum, l_lngEmpNum, l_datChargeDate)
            m_colEMPRVTrans.Add BuildEMPRVTransaction(m_rsDeleteActualTime, l_intLocCd, "D")
        End If
        m_rsDeleteActualTime.MoveNext
    Wend
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "DeleteActualTime"

End Function



Private Sub AppendActualTimeWarning()

    Dim l_strMsg As String
    Dim l_bolError As Boolean
    Dim lngKeyEmpNum As Long
    Dim l_datKeyChargeDate As Date
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_strExcCd As String
    Dim l_strOverExcCd As String


    On Error GoTo eh

    l_strMsg = "Actual time has been deleted."
    
    lngKeyEmpNum = m_rsDeleteActualTime!EMP_NUM
    l_datKeyChargeDate = m_rsDeleteActualTime!CHARGE_DATE
    l_strStartTime = m_rsDeleteActualTime!START_TIME
    l_strEndTime = m_rsDeleteActualTime!END_TIME
    l_strExcCd = m_rsDeleteActualTime!EXCEPTION_CD
    l_strOverExcCd = m_rsDeleteActualTime!OVER_EXC_CD
    
    l_bolError = AppendBusinessError(m_rsErrors, _
            20000, _
            mstrCBUSINESSWAR, _
            l_strMsg, _
            m_strCMODULENAME, _
            91, , , "EMP_NUM;CHARGE_DATE;SHIFT_NUM,START_TIME,END_TIME,EXCEPTION_CD, OVER_EXC_CD", _
            str(lngKeyEmpNum), _
            str(l_datKeyChargeDate), str(0), _
            l_strStartTime, l_strEndTime, l_strExcCd, l_strOverExcCd)

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AppendActualTimeWarning"

End Sub


Private Sub DetermineShift( _
    ByRef r_rsRequest As ADODB.Recordset)
    
    Dim l_datChargeDate As Date
    Dim l_lngEmpNum As Long
    Dim l_strFullShiftInd As String
    Dim l_strFilter As String
    Dim l_blnShiftFound As Boolean
    Dim l_strExcStartTime As String
    Dim l_strExcEndTime As String
    
    
    On Error GoTo eh
    l_lngEmpNum = r_rsRequest!EMP_NUM
    l_datChargeDate = r_rsRequest!CHARGE_DATE
    If IsNull(r_rsRequest!START_TIME) Or _
        IsEmpty(r_rsRequest!START_TIME) Then
        r_rsRequest!START_TIME = ""
    End If
    If IsNull(r_rsRequest!END_TIME) Or _
        IsEmpty(r_rsRequest!END_TIME) Then
        r_rsRequest!END_TIME = ""
    End If
    l_strExcStartTime = Trim(r_rsRequest!START_TIME)
    l_strExcEndTime = Trim(r_rsRequest!END_TIME)
    
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
        " AND CHARGE_DATE = " & str(l_datChargeDate) & _
        " AND SHIFT_HOURS > 0 "
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsSchedules.Filter = l_strFilter
    If Not RecordsExist(m_rsSchedules) Then
'If there are no working shifts it will default to shift 1
'An error will display in InsertException if there is no shift 1
        l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
        " AND CHARGE_DATE = " & str(l_datChargeDate)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        m_rsSchedules.Filter = l_strFilter
        If Not RecordsExist(m_rsSchedules) Then
            r_rsRequest!SHIFT_NUM = 1
        Else
            m_rsSchedules.Sort = "SHIFT_NUM"
            m_rsSchedules.MoveFirst
            r_rsRequest!SHIFT_NUM = m_rsSchedules!SHIFT_NUM
            SetExceptionTimes l_strExcStartTime, l_strExcEndTime, r_rsRequest
        End If
        Exit Sub
    End If
    
    m_rsSchedules.Sort = "SHIFT_NUM"
    m_rsSchedules.MoveFirst
    If m_rsSchedules.RecordCount = 1 Then
        r_rsRequest!SHIFT_NUM = m_rsSchedules!SHIFT_NUM
        SetExceptionTimes l_strExcStartTime, l_strExcEndTime, r_rsRequest
        Exit Sub
    End If
    l_strFullShiftInd = r_rsRequest!FULL_SHIFT_EXC_IND
    If l_strFullShiftInd = "Y" Then
        If l_strExcStartTime = "" And l_strExcEndTime = "" Then
' Full shift exception without a designated shift
' and there is more than 1 shift on a day
            r_rsRequest!SHIFT_NUM = 1
            SetExceptionTimes l_strExcStartTime, l_strExcEndTime, r_rsRequest
            Exit Sub
        End If
    End If
    
    l_blnShiftFound = FindShift(r_rsRequest, 0)
    If l_blnShiftFound Then
        r_rsRequest!SHIFT_NUM = m_rsSchedules!SHIFT_NUM
        SetExceptionTimes l_strExcStartTime, l_strExcEndTime, r_rsRequest
        Exit Sub
    End If
    l_blnShiftFound = FindShift(r_rsRequest, 1)
    If l_blnShiftFound Then
        r_rsRequest!SHIFT_NUM = m_rsSchedules!SHIFT_NUM
        SetExceptionTimes l_strExcStartTime, l_strExcEndTime, r_rsRequest
        Exit Sub
    End If
    l_blnShiftFound = FindShift(r_rsRequest, -1)
    If l_blnShiftFound Then
        r_rsRequest!SHIFT_NUM = m_rsSchedules!SHIFT_NUM
        SetExceptionTimes l_strExcStartTime, l_strExcEndTime, r_rsRequest
    Else
'Shift not Found let it error out later
        r_rsRequest!SHIFT_NUM = 1
    End If

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineShift"

    
End Sub

Private Function FindShift( _
    ByRef r_rsRequest As ADODB.Recordset, _
    ByVal v_intChargeDateOffSet As Integer _
    ) As Boolean
    
    Dim l_strShiftStartTime As String
    Dim l_strShiftEndTime As String
    Dim l_datShiftStartDate As Date
    Dim l_datShiftEndDate As Date
    Dim l_datShiftStartDateTime As Date
    Dim l_datShiftEndDateTime As Date
    Dim l_datExcStartDateTime As Date
    Dim l_datExcEndDateTime As Date
    Dim l_blnShiftFound As Boolean
    Dim l_strExcStartTime As String
    Dim l_strExcEndTime As String
    Dim l_datExcStartDate As Date
    Dim l_datExcEndDate As Date
    Dim l_datChargeDate As Date
    
    On Error GoTo eh
    
    l_datChargeDate = r_rsRequest!CHARGE_DATE
    l_strExcStartTime = Trim(r_rsRequest!START_TIME)
    l_strExcEndTime = Trim(r_rsRequest!END_TIME)
    
 'Please note that if Start and End Time are both empty
 ' it is a full shift exception and this method will not be called
 
    If l_strExcStartTime = "" Then
        l_strExcStartTime = l_strExcEndTime
    End If
    
    If l_strExcEndTime = "" Then
        l_strExcEndTime = l_strExcStartTime
    End If
    
    If l_strExcStartTime > l_strExcEndTime Then
'Crosses Midnight
        l_datExcStartDate = DateAdd("d", v_intChargeDateOffSet - 1, l_datChargeDate)
        l_datExcEndDate = DateAdd("d", 1, l_datExcStartDate)
    Else
        l_datExcStartDate = DateAdd("d", v_intChargeDateOffSet, l_datChargeDate)
        l_datExcEndDate = DateAdd("d", v_intChargeDateOffSet, l_datChargeDate)
    End If
'Find Shift
    l_blnShiftFound = False
    l_datExcStartDateTime = DateTimeFrom(l_datExcStartDate, l_strExcStartTime)
    l_datExcEndDateTime = DateTimeFrom(l_datExcEndDate, l_strExcEndTime)
    m_rsSchedules.MoveFirst
    While Not m_rsSchedules.EOF And Not l_blnShiftFound
        l_strShiftStartTime = m_rsSchedules!SHIFT_START_TIME
        l_datShiftStartDate = m_rsSchedules!SHIFT_START_DATE
        l_strShiftEndTime = m_rsSchedules!SHIFT_END_TIME
        l_datShiftEndDate = m_rsSchedules!SHIFT_END_DATE
        l_datShiftStartDateTime = DateTimeFrom(l_datShiftStartDate, l_strShiftStartTime)
        l_datShiftEndDateTime = DateTimeFrom(l_datShiftEndDate, l_strShiftEndTime)
        If TimePeriodsIntersect2( _
            l_datShiftStartDateTime, l_datShiftEndDateTime, _
            l_datExcStartDateTime, l_datExcEndDateTime) Then
            l_blnShiftFound = True
        Else
            m_rsSchedules.MoveNext
        End If
    Wend
    
    FindShift = l_blnShiftFound
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "FindShift"
End Function

Private Function CanBeScheduledOnWorkDay( _
    ByVal v_strExcCd As String _
    ) As Boolean
    
    On Error GoTo eh
    
    Select Case Trim(v_strExcCd)
        Case "HR"
            CanBeScheduledOnWorkDay = False
        Case Else
            CanBeScheduledOnWorkDay = True
    End Select
    Exit Function
eh:

    TrsRaiseError m_strCMODULENAME, "CanBeScheduledOnWorkDay"
    
End Function

Private Function GetEffectiveJob( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date _
    ) As String
    
    
    Dim l_datJobEffDate As Date
    Dim l_strJobNum As String
    Dim l_strPriorJobNum As String
    Dim l_datPriorJobEffDate As Date
    Dim l_obj As busEmployee.cbsEmployee
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    
    l_datJobEffDate = m_rsEmployeeData!JOB_EFF_DATE
    l_strJobNum = m_rsEmployeeData!JOB_NUM
    l_strPriorJobNum = m_rsEmployeeData!PRIOR_JOB_NUM
    l_datPriorJobEffDate = m_rsEmployeeData!PRIOR_JOB_EFF_DT
    
    GetEffectiveJob = l_obj.EmployeeJobNum( _
        l_datJobEffDate, l_strJobNum, l_strPriorJobNum, l_datPriorJobEffDate, v_datChargeDate)
    
    Set l_obj = Nothing
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetEffectiveJob"
    
    
End Function

Private Function IsEligibleForJobUpgrade( _
    ByVal v_strUnionCode As String, _
    ByVal v_strExcCode As String _
    ) As Boolean
    
'BLUE021
    Dim l_strCaseGroup As String
    
    On Error GoTo eh
    
    If v_strUnionCode = trUnion148 Then
        l_strCaseGroup = CaseGroupFor(v_strExcCode)
        If l_strCaseGroup = trCaseGroupIllness Or _
            l_strCaseGroup = trCaseGroupSuspension Or _
            ExceptionBankTypeFor(v_strExcCode) = trBankTypeVacation Then
            IsEligibleForJobUpgrade = False
        Else
            IsEligibleForJobUpgrade = True
        End If
    Else
        IsEligibleForJobUpgrade = True
    End If
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsEligibleForJobUpgrade"
    
End Function


Private Function ExceptionBankTypeFor( _
            ByVal v_strExcCd As String) As String
' Answer the first character of bank code.
' BLUE021
    Dim l_obj As dbCodeTables.cdCodeTables

On Error GoTo eh

    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")

    ExceptionBankTypeFor = _
        l_obj.ExceptionBankType(v_strExcCd)
    Set l_obj = Nothing
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionBankTypeFor"
End Function



Private Sub SetExceptionTimes( _
    ByVal v_strExcStart As String, _
    ByVal v_strExcEnd As String, _
    ByRef r_rsRequest As ADODB.Recordset)
    
    On Error GoTo eh
    
    If v_strExcStart = "" Then
        r_rsRequest!START_TIME = m_rsSchedules!SHIFT_START_TIME
    End If
    
    If v_strExcEnd = "" Then
        r_rsRequest!END_TIME = m_rsSchedules!SHIFT_END_TIME
    End If
    
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "SetExceptionTimes"
    
End Sub

Private Function FilterSchedule( _
    ByVal v_rsRequest As ADODB.Recordset _
    ) As String
    
    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_lngShiftNum As Long
    Dim l_strFilter As String
    
    
    On Error GoTo eh
    
    l_lngEmpNum = v_rsRequest!EMP_NUM
    l_datChargeDate = v_rsRequest!CHARGE_DATE
    
    l_lngShiftNum = v_rsRequest!SHIFT_NUM
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
        " AND CHARGE_DATE = " & str(l_datChargeDate) & _
        " AND SHIFT_NUM = " & str(l_lngShiftNum)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsSchedules.Filter = l_strFilter
    
    FilterSchedule = l_strFilter
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FilterSchedule"
    
End Function


Private Sub FillInMissingTime( _
    ByRef r_rsRequest As ADODB.Recordset)
    
    Dim l_strExcStartTime As String
    Dim l_strExcEndTime As String
    
    On Error GoTo eh
    
    FilterSchedule r_rsRequest
    If Not RecordsExist(m_rsSchedules) Then
        Exit Sub
    'let it error out later with BRV
    End If
    
    If IsNull(r_rsRequest!START_TIME) Or _
        IsEmpty(r_rsRequest!START_TIME) Then
        r_rsRequest!START_TIME = ""
    End If
    If IsNull(r_rsRequest!END_TIME) Or _
        IsEmpty(r_rsRequest!END_TIME) Then
        r_rsRequest!END_TIME = ""
    End If
    l_strExcStartTime = Trim(r_rsRequest!START_TIME)
    l_strExcEndTime = Trim(r_rsRequest!END_TIME)
    
    SetExceptionTimes l_strExcStartTime, l_strExcEndTime, r_rsRequest
                
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "FillInMissingTime"
    
End Sub


Private Function CheckScheduleChangePayments( _
                    ByVal v_strExcCd As String, _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date _
                ) As Boolean
' Check for any schedule change payments on the charge date of the exception.
' If a payment is found, just warn the user so they can take
' appropriate action, e.g. move the payment to the next scheduled work day.
' Answer TRUE if a payment exists, otherwise FALSE.
    Dim l_blnExists As Boolean
    Dim l_strFilter As String
                    
On Error GoTo eh
                    
    l_blnExists = False
    If Not RecordsExist(m_rsSchedChgPmts) Then
        CheckScheduleChangePayments = False
        Exit Function
    End If
    
    l_strFilter = "EMP_NUM = " & CStr(v_lngEmpNum) & _
        " AND CHARGE_DATE = #" & Format(v_datChargeDate, "mm/dd/yyyy") & "#"
    m_rsSchedChgPmts.Filter = MDAC25FilterFix(l_strFilter)

    If RecordsExist(m_rsSchedChgPmts) Then
        If Not IsScheduleChangePaymentPayable(v_strExcCd) Then
            l_blnExists = True
            AppendWarningScheduleChangePayment _
                    v_lngEmpNum, _
                    m_rsSchedChgPmts!EMP_NAME, _
                    v_datChargeDate, _
                    v_strExcCd
            Exit Function
        End If
    End If
        
    m_rsSchedChgPmts.Filter = ""
    CheckScheduleChangePayments = l_blnExists
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CheckScheduleChangePayments"
End Function


Private Function IsScheduleChangePaymentPayable( _
                    ByVal v_strExcCd As String) As Boolean

    Dim l_obj As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    IsScheduleChangePaymentPayable = l_obj.ExceptionScheduleChgInd(v_strExcCd) = "Y"
    
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsScheduleChangePaymentPayable"
End Function


Private Function AppendWarningScheduleChangePayment( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_strEmpName As String, _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_strExcCd As String)

    Dim l_strWarDesc As String
    Dim l_colDescArgs As Collection

On Error GoTo eh

    Set l_colDescArgs = New Collection
    l_colDescArgs.Add CStr(v_lngEmpNum)
    l_colDescArgs.Add Trim(v_strEmpName)
    l_colDescArgs.Add Format(v_datChargeDate, "mm/dd/yyyy")
    l_colDescArgs.Add v_strExcCd
    l_strWarDesc = TrsGetErrDescription(LongAppErr(trWarSchedChgPmtOffExc), l_colDescArgs)
    Set l_colDescArgs = Nothing
    AppendBusinessError _
        m_rsErrors, _
        LongAppErr(trWarSchedChgPmtOffExc), _
        mstrCBUSINESSWAR, _
        l_strWarDesc, _
        m_strCMODULENAME, _
        TrType_EXS, _
        , , _
        "EMP_NUM;CHARGE_DATE", _
        CStr(v_lngEmpNum), _
        Format(v_datChargeDate, "mm/dd/yyyy")

    Exit Function

eh:
    Set l_colDescArgs = Nothing
    TrsRaiseError m_strCMODULENAME, "AppendWarningScheduleChangePayment"
End Function


Private Function PartialHourMinutes(ByVal v_lngEmpNum As Long) As Integer
    
On Error GoTo eh
    
    m_rsEmployeeData.Filter = "EMP_NUM = " & CStr(v_lngEmpNum)
    PartialHourMinutes = m_rsEmployeeData!TIME_INCR
    m_rsEmployeeData.Filter = adFilterNone
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "PartialHourMinutes"
End Function


Private Function IsHoursIncrementValid( _
                    ByVal v_decHours As Variant, ByVal v_intHoursIncr As Integer) As Boolean
    Dim l_lngRemainder As Long
    
On Error GoTo eh
    
    If v_intHoursIncr = 0 Then
        IsHoursIncrementValid = False
        Exit Function
    End If
    
    l_lngRemainder = HoursAsMinutes(v_decHours) Mod v_intHoursIncr
    If l_lngRemainder = 0 Then
        IsHoursIncrementValid = True
    Else
        IsHoursIncrementValid = False
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsHoursIncrementValid"
End Function


Private Function HoursAsMinutes(ByVal v_decHours As Variant) As Integer

On Error GoTo eh
    
    HoursAsMinutes = v_decHours * 60
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "HoursAsMinutes"
End Function

Private Sub DayOffProcess( _
    ByRef r_rsRequest As ADODB.Recordset)
    
'On Days off, the exception will be assigned to shift 1 always.
'This code will set shift 2 to shift 1, if the user selected shift 2

    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strFilter As String
    
    
On Error GoTo eh


    l_lngEmpNum = r_rsRequest!EMP_NUM
    l_datChargeDate = r_rsRequest!CHARGE_DATE
    
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
        " AND CHARGE_DATE = " & str(l_datChargeDate) & _
        " AND SHIFT_HOURS > 0 "
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsSchedules.Filter = l_strFilter
    If Not RecordsExist(m_rsSchedules) Then
'If there are no working shifts it will default to shift 1
'An error will display in InsertException if there is no shift 1
        r_rsRequest!SHIFT_NUM = 1
    End If
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "DayOffProcess"
    
End Sub
Private Function IsAmerenCILCOLocal8Employee( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As Boolean
    
        
    Dim l_datJobEffDate As Date
    Dim l_strUnionCd As String
    Dim l_strPriorUnionCd As String
    Dim l_obj As busEmployee.cbsEmployee
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")

    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    
    l_datJobEffDate = m_rsEmployeeData!JOB_EFF_DATE
    l_strUnionCd = m_rsEmployeeData!UNION_CD
    l_strPriorUnionCd = m_rsEmployeeData!PRIOR_UNION_CD
    
    IsAmerenCILCOLocal8Employee = l_obj.IsAmerenCILCOLocal8Employee( _
        l_datJobEffDate, l_strUnionCd, l_strPriorUnionCd, v_datDate)
    

    Set l_obj = Nothing

    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCILCOLocal8Employee"
    
End Function



Private Function IsExceptionIsEligibleForShiftComponent( _
    ByVal v_strExcCd As String) As Boolean

'RULE0020
    On Error GoTo eh

    IsExceptionIsEligibleForShiftComponent = Not _
        DoesRuleApply(20, trExceptionCode, v_strExcCd)
        
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsExceptionIsEligibleForShiftComponent"
        
End Function


Private Function DoesRuleApply( _
    ByVal v_lngRuleNum As Long, _
    ByVal v_lngBusType As TRSRulesDataTypes, _
    ByVal v_strValue As String _
    ) As Boolean
       
    Dim l_obj As dbRules.cdRules
       
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    DoesRuleApply = l_obj.DoesRuleApply(v_lngRuleNum, v_lngBusType, v_strValue)
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "DoesRuleApply"
    
End Function

Private Function EmployeeBenefitPackage( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As String
    
        
    Dim l_datJobEffDate As Date
    Dim l_strUnionCd As String
    Dim l_strPriorUnionCd As String
    Dim l_strExcConInd As String
    Dim l_strPriorExcConInd As String
    Dim l_obj As busEmployee.cbsEmployee
    Dim l_strMgmtBenPackage As String
    Dim l_strUnionBenPackage As String
    Dim l_strPriorUnionBenPackage As String
    Dim l_strBenPackage As String
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")

    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    
    l_datJobEffDate = m_rsEmployeeData!JOB_EFF_DATE
    l_strUnionCd = m_rsEmployeeData!UNION_CD
    l_strPriorUnionCd = m_rsEmployeeData!PRIOR_UNION_CD
    l_strExcConInd = m_rsEmployeeData!EXEC_CONTRACT_IND
    l_strPriorExcConInd = m_rsEmployeeData!PRIOR_EXEC_CON_IND
    l_strMgmtBenPackage = m_rsEmployeeData!MGMT_BEN_PKG_IND
    l_strUnionBenPackage = m_rsEmployeeData!UNION_BEN_PKG_IND
    l_strPriorUnionBenPackage = m_rsEmployeeData!PRIOR_UNION_BN_PKG
    
    EmployeeBenefitPackage = l_obj.EmployeeBenefitPackage( _
        l_datJobEffDate, l_strMgmtBenPackage, l_strUnionCd, l_strPriorUnionCd, _
        l_strExcConInd, l_strPriorExcConInd, _
        l_strUnionBenPackage, l_strPriorUnionBenPackage, v_datDate)
    
    Set l_obj = Nothing

    
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "EmployeeBenefitPackage"
    
End Function

Private Function IsValidHolidayBanked( _
    ByVal v_datChargeDate As Date, _
    ByVal v_strExceptionCode As String, _
    ByVal v_strEmployeeUnion As String _
    ) As Boolean

'BLUE094, RED038

    Dim l_strErrorMsg As String
    Dim l_blnIsValid As Boolean
    Dim l_strExecConInd As String
    Dim l_strBenefitsInd As String

On Error GoTo eh
    
    If Trim(v_strExceptionCode) <> "HB" Then
        IsValidHolidayBanked = True
        Exit Function
    End If
    
    If m_rsHolidays Is Nothing Then
        SetError trErrbusValidationError, "Holiday Banked can only be scheduled on a Holiday"
        IsValidHolidayBanked = False
        Exit Function
    Else
        m_rsHolidays.Filter = "HOLIDAY_DATE = " & str(v_datChargeDate) & " and UNION_CD = '" & v_strEmployeeUnion & "' "
        If m_rsHolidays.BOF And m_rsHolidays.EOF Then
            SetError trErrbusValidationError, "Holiday Banked can only be scheduled on a Holiday"
            IsValidHolidayBanked = False
        Else
            IsValidHolidayBanked = True
        End If
    End If
       
    m_rsHolidays.Filter = ""
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsValidHolidayBanked"

End Function


Private Function IsValidHolidayDate( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date, _
    ByVal v_strExceptionCode As String, _
    ByVal v_strEmployeeUnion As String, _
    ByVal v_rsReq As ADODB.Recordset _
    ) As Boolean

On Error GoTo eh

    Dim l_blnIsValid As Boolean
    Dim l_blnIsDateAHoliday As Boolean
    Dim l_decHours As Variant
    Dim l_blnFlexEmp As Boolean

    If Trim(v_strExceptionCode) <> "H" Then
        IsValidHolidayDate = True
        Exit Function
    End If

    If Trim(v_rsReq!REC_STATUS) = "D" Then
        IsValidHolidayDate = True
        Exit Function
    End If
    
    If FlexPolicyApplies(v_lngEmpNum, v_datChargeDate) Then
        l_blnFlexEmp = True
    Else
        l_blnFlexEmp = False
    End If

    l_blnIsValid = True
    l_blnIsDateAHoliday = True
    If m_rsHolidays Is Nothing Then
        l_blnIsDateAHoliday = False
    Else
        m_rsHolidays.Filter = "HOLIDAY_DATE = " & str(v_datChargeDate) & " and UNION_CD = '" & v_strEmployeeUnion & "' "
        If m_rsHolidays.BOF And m_rsHolidays.EOF Then
            l_blnIsDateAHoliday = False
        End If
    End If
    If Not l_blnIsDateAHoliday Then
        If l_blnFlexEmp Then
            SetError trErrbusValidationError, "Date " & v_datChargeDate & " is not on the Holiday Calendar."
            l_blnIsValid = False
        Else
            AppendWarning v_rsReq, vbObjectError + 20000, "Date " & v_datChargeDate & " is not on the Holiday Calendar."
            l_blnIsValid = True
        End If
    End If
    
    If Not m_rsHolidays Is Nothing Then
        m_rsHolidays.Filter = ""
    End If

    IsValidHolidayDate = l_blnIsValid


    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsValidHolidayDate"

End Function


Private Function IsValidHolidayHours( _
        ByVal v_lngEmpNum As Long, _
        ByVal v_rsExceptions As ADODB.Recordset _
    ) As Boolean

On Error GoTo eh

    Dim l_blnIsValid As Boolean
    Dim l_datChargeDate As Date
    Dim l_datChargeDateHold As Date
    Dim l_decHours As Variant
    
    v_rsExceptions.Filter = "EMP_NUM = " & v_lngEmpNum & " and EXCEPTION_CD = 'H' "
    If v_rsExceptions.EOF And v_rsExceptions.BOF Then
        IsValidHolidayHours = True
        v_rsExceptions.Filter = ""
        Exit Function
    End If
    
    l_blnIsValid = True
    
    v_rsExceptions.Sort = "CHARGE_DATE"
    v_rsExceptions.MoveFirst
    l_decHours = 0
    l_datChargeDateHold = v_rsExceptions!CHARGE_DATE
    While Not v_rsExceptions.EOF
        l_datChargeDate = v_rsExceptions!CHARGE_DATE
        If l_datChargeDate = l_datChargeDateHold Then
            l_decHours = l_decHours + v_rsExceptions!Hours
        Else
            If l_decHours > 8 Then
                If LimitEightHoursOfHoliday(v_lngEmpNum, l_datChargeDateHold) Then
                    l_blnIsValid = False
                    SetError trErrbusValidationError, "Can't Schedule more than 8 hours of Holiday"
                    AppendExcessHolidayEntryError v_lngEmpNum, l_datChargeDateHold
                End If
            End If
            l_decHours = v_rsExceptions!Hours
            l_datChargeDateHold = l_datChargeDate
        End If
        v_rsExceptions.MoveNext
    Wend

    
    If l_decHours > 8 Then
        If LimitEightHoursOfHoliday(v_lngEmpNum, l_datChargeDateHold) Then
            l_blnIsValid = False
            SetError trErrbusValidationError, "Can't Schedule more than 8 hours of Holiday"
            AppendExcessHolidayEntryError v_lngEmpNum, l_datChargeDateHold
        End If
    End If

    
    IsValidHolidayHours = l_blnIsValid
    
    

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsValidHoliday"

End Function





Private Function PostValidate() As Boolean

    Dim l_blnIsValid As Boolean
    Dim l_strEmployee As String
    Dim l_colEmployees As Collection
    Dim l_intI As Integer
    Dim l_intNumOfEmps As Integer
    Dim l_lngEmpNum As Long
    Dim l_rsExceptions As ADODB.Recordset

    On Error GoTo eh

    Set l_colEmployees = New Collection
    Set l_rsExceptions = CopyRecordset(m_rsExceptions)
    If l_rsExceptions.BOF And l_rsExceptions.EOF Then
        PostValidate = True
        Exit Function
    End If
    l_rsExceptions.Filter = ""
    l_rsExceptions.MoveFirst
    While Not l_rsExceptions.EOF
        Select Case l_rsExceptions.Status
            Case adRecModified
                l_strEmployee = Trim(str(l_rsExceptions!EMP_NUM))
            Case adRecNew
                l_strEmployee = Trim(str(l_rsExceptions!EMP_NUM))
            Case adRecDeleted
                l_strEmployee = ""
        End Select
        If l_strEmployee <> "" Then
            If Not IncludesString(l_strEmployee, l_colEmployees) Then
                l_colEmployees.Add l_strEmployee
            End If
        End If
        l_rsExceptions.MoveNext
    Wend
    
    If l_colEmployees.Count = 0 Then
        PostValidate = True
        Exit Function
    End If
    
    l_blnIsValid = True
    l_intNumOfEmps = l_colEmployees.Count
    For l_intI = 1 To l_intNumOfEmps
        l_lngEmpNum = CLng(l_colEmployees.Item(l_intI))
        If Not IsValidHolidayHours(l_lngEmpNum, l_rsExceptions) Then
            l_blnIsValid = False
        End If
    Next l_intI

    PostValidate = l_blnIsValid
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PostValidate"
End Function


            
Private Sub AppendExcessHolidayEntryError( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datChargeDate As Date)
            
            
            
    Dim l_bolError As Boolean
    Dim lngKeyEmpNum As Long
    Dim l_datKeyChargeDate As Date
    Dim lngKeyShiftNum As Long
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_strExcCd As String
    Dim l_strOverExcCd As String

    On Error GoTo AppendExcessHolidayEntryError_EH
    
    lngKeyEmpNum = v_lngEmpNum
    l_datKeyChargeDate = v_datChargeDate
    lngKeyShiftNum = 1
    
    l_strStartTime = ""
    l_strEndTime = ""
    l_strExcCd = "H"
    l_strOverExcCd = ""

    l_bolError = AppendBusinessError(m_rsErrors, _
            mlngErrorNumber, _
            mstrCBUSINESSERR, _
            mstrErrorDescription, _
            m_strCMODULENAME, _
            90, , , "EMP_NUM;CHARGE_DATE;SHIFT_NUM,START_TIME,END_TIME,EXCEPTION_CD, OVER_EXC_CD", _
            str(lngKeyEmpNum), _
            str(l_datKeyChargeDate), str(lngKeyShiftNum), _
            l_strStartTime, l_strEndTime, l_strExcCd, l_strOverExcCd)

    Exit Sub
    
AppendExcessHolidayEntryError_EH:
    TrsRaiseError m_strCMODULENAME, "AppendExcessHolidayEntryError"

End Sub

Private Function LimitEightHoursOfHoliday( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As Boolean

    Dim l_intLocationCode As Integer
    Dim l_strDepartmentCode As String
    Dim l_strOTExclusion As String
    
'RULE0020
    On Error GoTo eh
    
    If Not IsExecutive(v_lngEmpNum, v_datDate) Then
        LimitEightHoursOfHoliday = False
        Exit Function
    End If
    
    l_intLocationCode = GetLocation(v_lngEmpNum, v_datDate)
    If DoesRuleApply(600, trLocationCode, Trim(str(l_intLocationCode))) Then
        LimitEightHoursOfHoliday = False
        Exit Function
    End If
    
    l_strDepartmentCode = GetDepartment(v_lngEmpNum, v_datDate)
    l_strOTExclusion = GetOTExclusion(v_lngEmpNum)
    
        
    If DoesRuleApply(601, trDepartmentNumber, Trim(l_strDepartmentCode)) And l_strOTExclusion = "0000" Then
        LimitEightHoursOfHoliday = False
        Exit Function
    End If
    
    If DoesRuleApply(602, trDepartmentNumber, Trim(l_strDepartmentCode)) And l_strOTExclusion = "0000" Then
        LimitEightHoursOfHoliday = False
        Exit Function
    End If
    
    If DoesRuleApply(600, trDepartmentNumber, Trim(l_strDepartmentCode)) Then
        LimitEightHoursOfHoliday = False
        Exit Function
    End If
        
    LimitEightHoursOfHoliday = True
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "LimitEightHoursOfHoliday"
        
End Function

Private Function GetLocation( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As Integer

    
    On Error GoTo eh
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    
    GetLocation = m_rsEmployeeData!LOC_CD


    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetLocation"
    
End Function

Private Function FlexPolicyApplies( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date _
    ) As Boolean
    
    FlexPolicyApplies = LimitEightHoursOfHoliday(v_lngEmpNum, v_datDate)
End Function

Private Function GetOTExclusion(ByVal v_lngEmpNum As Long) As String

    On Error GoTo eh
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    
    GetOTExclusion = m_rsEmployeeData!AR_OT_EXCLUSN

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetOTExclusion"
    
End Function

Private Function IsValidEntryDate( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date _
    ) As Boolean
   
    Dim l_datEmployDate As Date
    
 On Error GoTo eh
 
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    l_datEmployDate = m_rsEmployeeData!EMPLOY_DATE
    If l_datEmployDate > v_datChargeDate Then
        IsValidEntryDate = False
    Else
        IsValidEntryDate = True
    End If
 
     Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidEntryDate"
    
    End Function




