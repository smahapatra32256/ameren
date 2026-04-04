VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cboTimecardEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'******************************************************
' Steps to add a persistent property
' 1. Add private module/instance variables.
' 2. Add Get/Let property procedures.
' 3. Set value in cboTimecardEntry.Restore.
' 4. Store new property in save recordset; cbsTimecard.AddNewTimeEntry and cbsTimecard.UpdateTimeEntry
' 5. Copy new property in SplitSimple.
' 6. If default or derived value for new timecard entries, add appropriate code to cboTimecard.BuildMissingComponents.
' 7. Do the same as above in cboTimecard.UpdateTimecardEntry to handle updates.
'******************************************************

Private Const mstrCMODULENAME = "cboTimecardEntry"
Private mobjTimecard As cboTimecard

Private mlngObjectStatus As Long

'Instance variables
Private mlngEmployeeNumber As Long
Private mdatChargeDate As Date
Private mdatDateWorked As Date
Private mstrDateWorkedDayName As String
Private mlngWorkgroupNumber As Long
Private mstrStartTime As String
Private mstrEndTime As String
Private mstrIndexTime As String
Private mstrGroupTimestamp As String
Private mstrDepartmentNumber As String
Private mvarHours As Variant
Private mintLocationCode As Integer
Private mstrUnionCode As String
Private mstrExceptionCode As String
Private mstrOverrideExceptionCode As String
Private mstrOvertimeReasonCode As String
Private mstrPreArrangedOTInd As String
Private mdatOvertimeBankChargeDate As Date
Private mvarTravelPayHours As Variant
Private mstrEarningsCode As String
Private mstrTradeExceptionCode As String
Private mstrReliefCode As String
Private mstrComponentCode As String
Private mvarComponentAmount As Variant
Private mstrApprovalIndicator As String
Private mlngApprovedBy As Long
Private mdatApprovedDate As Date
Private mstrFamilyLeaveIndicator As String
Private mstrJobNumber As String
Private mstrScheduleChangeExceptionCode As String
Private mstrScheduleChangeEarningsCode As String
Private mlngCaseNumber As Long
Private mstrShiftComponentCode As String
Private mvarShiftComponentAmount As Variant
Private mlngOvertimeOpportunityNumber As Long
Private mstrOvertimeOpportunityDescription As String
Private mstrDistributionId As String
Private mstrDistributionType As String
Private mstrJobNbr As String
Private mstrJobTaskNbr As String
Private mintJobTaskStepNbr As Integer
Private mstrEMPRVDatabaseID As String
Private mstrDistributionDepartmentNumber As String
Private mlngScheduleCode As Long
Private mlngScheduleSourceCode As Long

' mstrSchedTimeInd is a transient attribute that only has
' meaning for an update request.
Private mstrSchedTimeInd As String
Private mstrOTAdjustDesc As String

'Use for validation
Private mlngErrorNumber As Long
Private mstrErrorDescription As String
'for older clients (6.0.0) that are not sending this column to the server
Private m_blnPreArrangedOTIndMissing As Boolean

Private Const l_intSplitNone As Integer = 0
Private Const l_intSplitStart As Integer = 1
Private Const l_intSplitEnd As Integer = 2
Private Const l_intSplitBoth As Integer = 3

Private Const l_strCFIXEDACCOUNTTYPE As String = "F"
Private Const l_strCEXCEPTIONACCOUNTTYPE As String = "E"
Private Const l_strCACCOUNTNOTEXISTS As String = ""
Private Const l_strCACCOUNTNOTAPPLY As String = "N/A"

'Error set is for validation and warnings only.
Public Function ErrorSet() As ADODB.Recordset
    Set ErrorSet = Timecard.ErrorSet
End Function

'Private Sub RaiseBusinessEntryError( _
'            ByVal v_lngErrorNumber As Long, _
'            Optional ByVal v_strErrorDescription As String, _
'            Optional ByRef vcolDescriptionArgs As Collection)
'    Dim l_strErrDesc As String
'
'    If v_strErrorDescription = "" Then
'        l_strErrDesc = TrsGetErrDescription(LongAppErr(v_lngErrorNumber), vcolDescriptionArgs)
'    Else
'        l_strErrDesc = v_strErrorDescription
'    End If
'    AppendEntryError v_lngErrorNumber, mstrCBUSINESSERR, l_strErrDesc
'    Err.Raise v_lngErrorNumber, mstrCMODULENAME, l_strErrDesc
'
'End Sub

Private Sub SetError( _
            ByVal vlngErrorNumber As Long, _
            Optional ByVal vstrErrorDescription As String)
' vlngErrorNumber should be vbObjectError + appErrorNumber

    mlngErrorNumber = vlngErrorNumber
    mstrErrorDescription = vstrErrorDescription
    
End Sub

Private Sub ResetError()

    SetError 0, ""
    
End Sub

Private Sub AppendEntryError( _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDescription As String, _
            Optional ByVal v_strErrorSource As String)
    Dim l_varErrorSource As String
    
    On Error GoTo eh
    
    l_varErrorSource = IIf(v_strErrorSource = "", mstrCMODULENAME, v_strErrorSource)
    
    AppendError _
        ErrorSet, _
        EmployeeNumber, _
        ChargeDate, _
        DateWorked, _
        startTime, _
        EndTime, _
        Hours, _
        ExceptionCode, _
        OverrideExceptionCode, _
        v_lngErrorNumber, _
        v_strErrorType, _
        v_strErrorDescription, _
        l_varErrorSource

    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "AppendEntryError", _
        trErrbusAppendEntryError, True, , , True
End Sub

Public Sub AppendWarning( _
            ByVal v_lngErrorNumber As Long, _
            Optional ByVal v_strErrorDescription As String, _
            Optional ByRef vcolDescriptionArgs As Collection)
    Dim l_strErrDesc As String
    
    If v_strErrorDescription = "" Then
        l_strErrDesc = TrsGetErrDescription(LongAppErr(v_lngErrorNumber), vcolDescriptionArgs)
    Else
        l_strErrDesc = v_strErrorDescription
    End If
    AppendEntryError v_lngErrorNumber, mstrCBUSINESSWAR, l_strErrDesc

End Sub

Private Function GetDepartmentManager() As dbDepartment.cdDepartment
    Set GetDepartmentManager = Timecard.GetBSTimecard.GetDepartmentManager
End Function

Private Function GetBusDepartmentManager() As busDepartment.cbsDepartment
    
    Set GetBusDepartmentManager = _
        CreateObject("busDepartment.cbsDepartment")
    
End Function

Private Function GetJobManager() As dbCodeTables.cdCodeTables
    Set GetJobManager = Timecard.GetBSTimecard.GetJobManager
End Function

Public Function GetEffectiveUnionCode() As String
    GetEffectiveUnionCode = Timecard.Employee.UnionCode(ChargeDate)
End Function

'REMOVE DATAMANAGER
Public Function Restore(rs As ADODB.Recordset)

On Error GoTo eh
    mlngEmployeeNumber = rs!EMP_NUM
    mdatChargeDate = rs!CHARGE_DATE
    mdatDateWorked = rs!DATE_WORKED
    mstrDateWorkedDayName = Trim(rs!DAY_NAME)
    mlngWorkgroupNumber = rs!WG_NUM
    mstrStartTime = RTrim(rs!START_TIME)
    mstrEndTime = RTrim(rs!END_TIME)
    mstrIndexTime = RTrim(rs!INDEX_TIME)
    mstrGroupTimestamp = RTrim(rs!TIME_STAMP)
    mstrDepartmentNumber = RTrim(rs!DEPT_NUM)
    mvarHours = rs!Hours
    mintLocationCode = rs!LOC_CD
    mstrUnionCode = RTrim(rs!UNION_CD)
    mstrExceptionCode = RTrim(rs!EXCEPTION_CD)
    mstrOverrideExceptionCode = RTrim(rs!OVER_EXC_CD)
    mstrOvertimeReasonCode = RTrim(rs!OT_REASON_CD)
    
    'for older clients (6.0.0) that are not sending this column to the server
    If Not PreArrangedOTIndMissing Then
        mstrPreArrangedOTInd = RTrim(rs!PREARRANGED_OT_IND)
    Else
        mstrPreArrangedOTInd = "N"
    End If
    
    mdatOvertimeBankChargeDate = rs!OT_BANK_CHRG_DT
    mvarTravelPayHours = rs!TRAVEL_PAY_HRS
    mstrEarningsCode = RTrim(rs!EARNINGS_CD)
    mstrTradeExceptionCode = RTrim(rs!TRADE_EXC_CD)
    mstrReliefCode = RTrim(rs!REL_CD)
    mstrComponentCode = RTrim(rs!COMP_CD)
    mvarComponentAmount = rs!COMP_AMT
    mstrApprovalIndicator = RTrim(rs!APPR_IND)
    mlngApprovedBy = rs!APPR_BY_EMP_NUM
    mdatApprovedDate = rs!APPR_DATE
    mstrFamilyLeaveIndicator = RTrim(rs!FAMILY_LEAVE_IND)
    mstrJobNumber = RTrim(rs!JOB_NUM)
    mstrScheduleChangeExceptionCode = RTrim(rs!SCHED_CHG_EXC_CD)
    mstrScheduleChangeEarningsCode = RTrim(rs!SCHED_CHG_EARN_CD)
    mlngCaseNumber = rs!CASE_NUM
    mstrShiftComponentCode = RTrim(rs!COMP_SHIFT_CD)
    mvarShiftComponentAmount = rs!COMP_SHIFT_CD_AMT
    mlngOvertimeOpportunityNumber = rs!OT_OPP_NUM
    mstrOvertimeOpportunityDescription = RTrim(rs!OT_OPP_DESC)
    'DO NOT TRIM DISTRIBUTION_ID WITH LEADING SPACES
    mstrDistributionId = RTrim(rs!DISTRIBUTION_ID)
    mstrDistributionType = RTrim(rs!DISTRIBUTION_TYPE)
    mstrJobNbr = RTrim(rs!JOB_NBR)
    mstrJobTaskNbr = RTrim(rs!JOBTASK_NBR)
    mintJobTaskStepNbr = rs!JOBTASK_STEP_NBR
    mstrEMPRVDatabaseID = RTrim(rs!EMPRV_DATABASE_ID)
    mstrDistributionDepartmentNumber = RTrim(rs!DISTR_DEPT_NUM)
    mlngScheduleCode = rs!SCHED_CD
    mlngScheduleSourceCode = rs!SCHED_SRC_CD
    
    Me.SetRestored
    Exit Function
     
eh:
    TrsRaiseError mstrCMODULENAME, "Restore"
End Function

Public Property Get ScheduleTimeIndicator() As String
    ScheduleTimeIndicator = mstrSchedTimeInd
End Property

Public Property Let ScheduleTimeIndicator(schedTimeInd As String)
    mstrSchedTimeInd = schedTimeInd
End Property

Public Property Get OTAdjustDesc() As String
    OTAdjustDesc = mstrOTAdjustDesc
End Property

Public Property Let OTAdjustDesc(v_strOTAdjustDesc As String)
    mstrOTAdjustDesc = v_strOTAdjustDesc
End Property

Public Property Get EmployeeNumber() As Long
    EmployeeNumber = mlngEmployeeNumber
End Property

Public Property Let EmployeeNumber(EmpNum As Long)
    If Not mlngEmployeeNumber = EmpNum Then
        mlngEmployeeNumber = EmpNum
        SetModified
    End If
End Property

Public Property Get ChargeDate() As Date
    ChargeDate = mdatChargeDate
End Property

Public Property Let ChargeDate(chgDate As Date)
    If Not mdatChargeDate = chgDate Then
        mdatChargeDate = chgDate
        SetModified
    End If
End Property

Public Property Get DateWorked() As Date
    DateWorked = mdatDateWorked
End Property

Public Property Let DateWorked(wrkDate As Date)
    If Not mdatDateWorked = wrkDate Then
        mdatDateWorked = wrkDate
        DateWorkedDayName = DayNameOfDate(wrkDate)
        SetModified
    End If
End Property

Public Property Get DateWorkedDayName() As String
    DateWorkedDayName = mstrDateWorkedDayName
End Property

Public Property Let DateWorkedDayName(wrkDateName As String)
    If Not Trim(mstrDateWorkedDayName) = Trim(wrkDateName) Then
        mstrDateWorkedDayName = Trim(wrkDateName)
        SetModified
    End If
End Property

Public Property Get WorkgroupNumber() As Long
    WorkgroupNumber = mlngWorkgroupNumber
End Property

Public Property Let WorkgroupNumber(wgNum As Long)
    If Not mlngWorkgroupNumber = wgNum Then
        mlngWorkgroupNumber = wgNum
        SetModified
    End If
End Property

Public Property Get startTime() As String
    startTime = mstrStartTime
End Property

Public Function StartTimeAsDate() As Date
' see comment on EndTimeAsDate
    StartTimeAsDate = DateTimeFrom(DateWorked, startTime)
End Function

Public Property Let startTime(stTime As String)
    If Not Trim(mstrStartTime) = Trim(stTime) Then
        mstrStartTime = Trim(stTime)
        SetModified
    End If
End Property

Public Property Get IndexTime() As String
    IndexTime = mstrIndexTime
End Property

Public Property Let IndexTime(ndxTime As String)
    If Not Trim(mstrIndexTime) = Trim(ndxTime) Then
        mstrIndexTime = Trim(ndxTime)
        SetModified
    End If
End Property

Public Property Get GroupTimestamp() As String
    GroupTimestamp = mstrGroupTimestamp
End Property

Public Property Let GroupTimestamp(v_strTimestamp As String)
    If Not Trim(mstrGroupTimestamp) = Trim(v_strTimestamp) Then
        mstrGroupTimestamp = Trim(v_strTimestamp)
        SetModified
    End If
End Property

Public Property Get EndTime() As String
    EndTime = mstrEndTime
End Property

Public Property Let EndTime(ndTime As String)
    If Not Trim(mstrEndTime) = Trim(ndTime) Then
        mstrEndTime = Trim(ndTime)
        SetModified
    End If
End Property

Public Function EndTimeAsDate() As Date
' timecard entries do not span midnight, so the
' date worked is the physical date for both start and
' end time.
    If EndTime = "00:00" Then
        EndTimeAsDate = DateTimeFrom(DateWorked, "24:00")
    Else
        EndTimeAsDate = DateTimeFrom(DateWorked, EndTime)
    End If
End Function

Public Property Get DepartmentNumber() As String
    DepartmentNumber = mstrDepartmentNumber
End Property

Public Property Let DepartmentNumber(deptNum As String)
    If Not Trim(mstrDepartmentNumber) = Trim(deptNum) Then
        mstrDepartmentNumber = Trim(deptNum)
        SetModified
    End If
End Property

Public Property Get Hours() As Variant
'DECIMAL(3,1)
    Hours = mvarHours
End Property

Public Property Let Hours(hrs As Variant)
'DECIMAL(3,1)
    If Not mvarHours = hrs Then
        mvarHours = hrs
        SetModified
    End If
End Property

Public Property Get LocationCode() As Integer
    LocationCode = mintLocationCode
End Property

Public Property Let LocationCode(ByVal v_intLocCd As Integer)
    If Not mintLocationCode = v_intLocCd Then
        mintLocationCode = v_intLocCd
        SetModified
    End If
End Property

Public Property Get UnionCode() As String
    UnionCode = mstrUnionCode
End Property

Public Property Let UnionCode(ByVal v_strUnionCd As String)
    If Not Trim(mstrUnionCode) = Trim(v_strUnionCd) Then
        mstrUnionCode = Trim(v_strUnionCd)
        SetModified
    End If
End Property

Public Property Get ExceptionCode() As String
'trim code here to temp correct test for overlapping time
' failing to accept equal exc codes. rsg 6/16/00
    ExceptionCode = Trim(mstrExceptionCode)
End Property

Public Property Let ExceptionCode(excCd As String)
    If Not Trim(mstrExceptionCode) = Trim(excCd) Then
        mstrExceptionCode = Trim(excCd)
        SetModified
    End If
End Property

Public Property Get OverrideExceptionCode() As String
    OverrideExceptionCode = mstrOverrideExceptionCode
End Property

Public Property Let OverrideExceptionCode(overExcCd As String)
    If Not Trim(mstrOverrideExceptionCode) = Trim(overExcCd) Then
        mstrOverrideExceptionCode = Trim(overExcCd)
        SetModified
    End If
End Property

Public Property Get OvertimeReasonCode() As String
    OvertimeReasonCode = mstrOvertimeReasonCode
End Property

Public Property Let OvertimeReasonCode(otRsnCd As String)
    If Not Trim(mstrOvertimeReasonCode) = Trim(otRsnCd) Then
        mstrOvertimeReasonCode = Trim(otRsnCd)
        SetModified
    End If
End Property

Public Property Get PreArrangedOTIndicator() As String
    PreArrangedOTIndicator = mstrPreArrangedOTInd
End Property

Public Property Let PreArrangedOTIndicator(ByVal v_strPreArrangedOTIndicator As String)
    If Not Trim(UCase(mstrPreArrangedOTInd)) = Trim(UCase(v_strPreArrangedOTIndicator)) Then
        mstrPreArrangedOTInd = Trim(UCase(v_strPreArrangedOTIndicator))
        SetModified
    End If
End Property

Public Property Get OvertimeBankChargeDate() As Date
    OvertimeBankChargeDate = mdatOvertimeBankChargeDate
End Property

Public Property Let OvertimeBankChargeDate(otBankChgDate As Date)
    If Not mdatOvertimeBankChargeDate = otBankChgDate Then
        mdatOvertimeBankChargeDate = otBankChgDate
        SetModified
    End If
End Property

Public Property Get TravelPayHours() As Variant
' Decimal 3,1
    TravelPayHours = mvarTravelPayHours
End Property

Public Property Let TravelPayHours(travelPayHrs As Variant)
' Decimal 3,1
    If Not mvarTravelPayHours = travelPayHrs Then
        mvarTravelPayHours = CDec(travelPayHrs)
        SetModified
    End If
End Property

Public Property Get EarningsCode() As String
    EarningsCode = mstrEarningsCode
End Property

Public Property Let EarningsCode(ByVal v_strEarnCd As String)
    If Not Trim(mstrEarningsCode) = Trim(v_strEarnCd) Then
        mstrEarningsCode = Trim(v_strEarnCd)
        SetModified
    End If
End Property

Public Property Get TradeExceptionCode() As String
    TradeExceptionCode = mstrTradeExceptionCode
End Property

Public Property Let TradeExceptionCode(tradeExcCd As String)
    If Not Trim(mstrTradeExceptionCode) = Trim(tradeExcCd) Then
        mstrTradeExceptionCode = Trim(tradeExcCd)
        SetModified
    End If
End Property

Public Property Get ReliefCode() As String
    ReliefCode = mstrReliefCode
End Property

Public Property Let ReliefCode(ByVal v_strReliefCode As String)
    If Not Trim(mstrReliefCode) = Trim(v_strReliefCode) Then
        mstrReliefCode = Trim(v_strReliefCode)
        SetModified
    End If
End Property

Public Property Get ComponentCode() As String
    ComponentCode = mstrComponentCode
End Property

Public Property Let ComponentCode(ByVal v_strCompCode As String)
    If Not Trim(mstrComponentCode) = Trim(v_strCompCode) Then
        mstrComponentCode = Trim(v_strCompCode)
        SetModified
    End If
End Property

Public Property Get ComponentAmount() As Variant
'DECIMAL(5,4)
    ComponentAmount = mvarComponentAmount
End Property

Public Property Let ComponentAmount(compAmt As Variant)
'DECIMAL(5,4)
    If Not mvarComponentAmount = compAmt Then
        mvarComponentAmount = compAmt
        SetModified
    End If
End Property

Public Sub SetApprovalIndicator( _
                ByVal v_strApprInd As String, _
                ByVal v_strApprovedBy As String)
    
    ApprovalIndicator = v_strApprInd
    If ApprovalIndicator = "Y" Then
        ApprovedBy = CStr(v_strApprovedBy)
        ApprovedDate = Date
    Else
        ApprovedBy = 0
        ApprovedDate = CDate(CDEFAULTDATE)
    End If

End Sub

Public Property Get ApprovalIndicator() As String
    ApprovalIndicator = mstrApprovalIndicator
End Property

Public Property Let ApprovalIndicator(ByVal v_strApprInd As String)
    If Not Trim(UCase(mstrApprovalIndicator)) = Trim(UCase(v_strApprInd)) Then
        mstrApprovalIndicator = Trim(UCase(v_strApprInd))
        SetModified
    End If
End Property

Public Property Get ApprovedBy() As Long
    ApprovedBy = mlngApprovedBy
End Property

Public Property Let ApprovedBy(ByVal v_lngApprovedBy As Long)
    If Not mlngApprovedBy = v_lngApprovedBy Then
        mlngApprovedBy = v_lngApprovedBy
        SetModified
    End If
End Property

Public Property Get ApprovedDate() As Date
    ApprovedDate = mdatApprovedDate
End Property

Public Property Let ApprovedDate(ByVal v_datApprovedDate As Date)
    If Not mdatApprovedDate = v_datApprovedDate Then
        mdatApprovedDate = v_datApprovedDate
        SetModified
    End If
End Property

Public Property Get FamilyLeaveIndicator() As String
    FamilyLeaveIndicator = mstrFamilyLeaveIndicator
End Property

Public Property Let FamilyLeaveIndicator(ByVal v_strFamLvInd As String)
    If Not Trim(mstrFamilyLeaveIndicator) = Trim(v_strFamLvInd) Then
        mstrFamilyLeaveIndicator = Trim(v_strFamLvInd)
        SetModified
    End If
End Property

Public Property Get JobNumber() As String
    JobNumber = mstrJobNumber
End Property

Public Property Let JobNumber(ByVal v_strJobNum As String)
    If Not Trim(mstrJobNumber) = Trim(v_strJobNum) Then
        mstrJobNumber = Trim(v_strJobNum)
        SetModified
    End If
End Property

Public Property Get ScheduleChangeExceptionCode() As String
    ScheduleChangeExceptionCode = mstrScheduleChangeExceptionCode
End Property

Public Property Let ScheduleChangeExceptionCode(schedChgExcCd As String)
    If Not Trim(mstrScheduleChangeExceptionCode) = Trim(schedChgExcCd) Then
        mstrScheduleChangeExceptionCode = Trim(schedChgExcCd)
        SetModified
    End If
End Property

Private Function IsExceptionEligibleForScheduleChange() As Boolean
' Is a schedule change payment allowed for this entries
' exception code.
' WHITE062, WHITE047
    IsExceptionEligibleForScheduleChange = _
        IsEligibleForScheduleChangePayment(ExceptionCode)
End Function

Public Property Get ScheduleChangeEarningsCode() As String
    ScheduleChangeEarningsCode = mstrScheduleChangeEarningsCode
End Property

Public Property Let ScheduleChangeEarningsCode(schedChgEarnCd As String)
    If Not Trim(mstrScheduleChangeEarningsCode) = Trim(schedChgEarnCd) Then
        mstrScheduleChangeEarningsCode = Trim(schedChgEarnCd)
        SetModified
    End If
End Property

Public Property Get CaseNumber() As Long
    CaseNumber = 0
End Property

Public Property Let CaseNumber(ByVal v_lngCaseNum As Long)
    mlngCaseNumber = 0
End Property

Public Property Get ShiftComponentCode() As String
    ShiftComponentCode = mstrShiftComponentCode
End Property

Public Property Let ShiftComponentCode(compShiftCd As String)
    If Not Trim(mstrShiftComponentCode) = Trim(compShiftCd) Then
        mstrShiftComponentCode = Trim(compShiftCd)
        SetModified
    End If
End Property

Public Property Get ShiftComponentAmount() As Variant
'DECIMAL(5,4)
    ShiftComponentAmount = mvarShiftComponentAmount
End Property

Public Property Let ShiftComponentAmount(compShiftCdAmt As Variant)
'DECIMAL(5,4)
    If Not mvarShiftComponentAmount = compShiftCdAmt Then
        mvarShiftComponentAmount = compShiftCdAmt
        SetModified
    End If
End Property

Public Property Get OvertimeOpportunityNumber() As Long
    OvertimeOpportunityNumber = mlngOvertimeOpportunityNumber
End Property

Public Property Let OvertimeOpportunityNumber(otOppNum As Long)
    If Not mlngOvertimeOpportunityNumber = otOppNum Then
        mlngOvertimeOpportunityNumber = otOppNum
        SetModified
    End If
End Property

Public Property Get OvertimeOpportunityDescription() As String
    OvertimeOpportunityDescription = mstrOvertimeOpportunityDescription
End Property

Public Property Let OvertimeOpportunityDescription(otOppDesc As String)
    If Not Trim(mstrOvertimeOpportunityDescription) = Trim(otOppDesc) Then
        mstrOvertimeOpportunityDescription = Trim(otOppDesc)
        SetModified
    End If
End Property

Public Property Get DistributionId() As String
    DistributionId = mstrDistributionId
End Property

Public Property Let DistributionId(distId As String)
' RTRIM ONLY - MAY HAVE LEADING SPACES
    If Not RTrim(mstrDistributionId) = RTrim(distId) Then
        mstrDistributionId = RTrim(distId)
        SetModified
    End If
End Property

Public Property Get DistributionType() As String
    DistributionType = mstrDistributionType
End Property

Public Property Let DistributionType(distType As String)
    If Not Trim(mstrDistributionType) = Trim(distType) Then
        mstrDistributionType = Trim(distType)
        SetModified
    End If
End Property
Public Property Get JobNbr() As String
    JobNbr = mstrJobNbr
End Property

Public Property Let JobNbr(jobNo As String)
    If Not Trim(mstrJobNbr) = Trim(jobNo) Then
        mstrJobNbr = Trim(jobNo)
        SetModified
    End If
End Property
Public Property Get JobTaskNbr() As String
    JobTaskNbr = mstrJobTaskNbr
End Property

Public Property Let JobTaskNbr(jobTask As String)
    If Not Trim(mstrJobTaskNbr) = Trim(jobTask) Then
        mstrJobTaskNbr = Trim(jobTask)
        SetModified
    End If
End Property
Public Property Get JobTaskStepNbr() As Integer
    JobTaskStepNbr = mintJobTaskStepNbr
End Property

Public Property Let JobTaskStepNbr(jobStep As Integer)
    If Not Trim(mintJobTaskStepNbr) = jobStep Then
        mintJobTaskStepNbr = jobStep
        SetModified
    End If
End Property
Public Property Get EMPRVDatabaseID() As String
    EMPRVDatabaseID = mstrEMPRVDatabaseID
End Property

Public Property Let EMPRVDatabaseID(dbID As String)
    If Not Trim(mstrEMPRVDatabaseID) = Trim(dbID) Then
        mstrEMPRVDatabaseID = Trim(dbID)
        SetModified
    End If
End Property

Public Property Get DistributionDepartmentNumber() As String
    DistributionDepartmentNumber = mstrDistributionDepartmentNumber
End Property

Public Property Let DistributionDepartmentNumber(distDeptNum As String)
    If Not Trim(mstrDistributionDepartmentNumber) = Trim(distDeptNum) Then
        mstrDistributionDepartmentNumber = Trim(distDeptNum)
        SetModified
    End If
End Property

Public Property Get ScheduleCode() As Long
    ScheduleCode = mlngScheduleCode
End Property

Public Property Let ScheduleCode(schedCd As Long)
    If Not mlngScheduleCode = schedCd Then
        mlngScheduleCode = schedCd
        SetModified
    End If
End Property

Public Property Get ScheduleSourceCode() As Long
    ScheduleSourceCode = mlngScheduleSourceCode
End Property

Public Property Let ScheduleSourceCode(schedSrcCd As Long)
    If Not mlngScheduleSourceCode = schedSrcCd Then
        mlngScheduleSourceCode = schedSrcCd
        SetModified
    End If
End Property

Public Property Get PreArrangedOTIndMissing() As Boolean
    PreArrangedOTIndMissing = m_blnPreArrangedOTIndMissing
End Property

Public Property Let PreArrangedOTIndMissing(blnPreArrangedOTIndMissing As Boolean)
    m_blnPreArrangedOTIndMissing = blnPreArrangedOTIndMissing
End Property

Public Sub Delete()
    SetDeleted
End Sub

Public Sub Release()
'Donot enable an error handler since this is called from an error handler.
    Set mobjTimecard = Nothing
    GetObjectContext.SetComplete
End Sub

Public Function IsRestored() As Boolean
'Has this object been restored from the database and unchanged?
    IsRestored = CBool(mlngObjectStatus And trObjectRestored)
End Function

Public Function IsNew() As Boolean
'Is this object new and not yet added to database?
    IsNew = CBool(mlngObjectStatus And trObjectNew)
End Function

Public Function IsDeleted() As Boolean
'Has this object been marked for deletion from the database?
    IsDeleted = CBool(mlngObjectStatus And trObjectDeleted)
End Function

Public Function IsModified() As Boolean
'Has this object been modified?
'Only answer true when the object has been restored from database.
'Dont answer True when the object is new.
    IsModified = CBool(mlngObjectStatus And trObjectModified) _
        And Not Me.IsNew
End Function

Public Sub SetRestored()
    mlngObjectStatus = mlngObjectStatus Or trObjectRestored
End Sub

Public Sub SetNew()
    mlngObjectStatus = mlngObjectStatus Or trObjectNew
End Sub

Public Sub SetModified()
    mlngObjectStatus = mlngObjectStatus Or trObjectModified
End Sub

Public Sub SetDeleted()
    mlngObjectStatus = mlngObjectStatus Or trObjectDeleted
End Sub

Public Function Split(ByVal v_strSplitTime As String, ByRef r_colEntries As Collection) As cboTimecardEntry
    Dim l_intPosition As Integer
    Dim l_objSplitEntry As cboTimecardEntry
    
    On Error GoTo eh
    
    Set l_objSplitEntry = SplitSimple(v_strSplitTime)
    l_intPosition = CollectionPosition(SafeRef(Me), r_colEntries)
    r_colEntries.Add l_objSplitEntry, , , l_intPosition
    Set Split = l_objSplitEntry
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "Split"
    
End Function

'All MTS bo_ objects must be a safe reference
Public Property Set Timecard(tc As cboTimecard)
    Set mobjTimecard = tc
End Property

Public Property Get Timecard() As cboTimecard
    Set Timecard = mobjTimecard
End Property

Public Function Employee() As cboEmployee
    Set Employee = Timecard.Employee
End Function

Public Function Schedule() As cboSchedule
    Set Schedule = Timecard.Schedule
End Function

Public Function IsRegularScheduledHours() As Boolean

    On Error GoTo eh

    If Me.ExceptionCode = "LU" Or Me.ExceptionCode = "LP" Then
        IsRegularScheduledHours = False
    Else
        IsRegularScheduledHours = _
            Timecard.Schedule.IsAnyPeriodWithinSchedule( _
                ChargeDate, StartTimeAsDate, EndTimeAsDate)
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsRegularScheduledHours"
    
End Function

Public Sub SetDefaultChargedWorkgroupNumber()
    Dim l_lngWgNum As Long

On Error GoTo eh
    l_lngWgNum = DetermineChargedWorkgroupNumber
    UpdateChargedWorkgroupNumber l_lngWgNum

    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "SetDefaultChargedWorkgroupNumber"
    
End Sub

Private Function DetermineChargedWorkgroupNumber() As Long

On Error GoTo eh
    
    DetermineChargedWorkgroupNumber = _
        Timecard.Employee.GetPrimaryWorkgroupNumber(ChargeDate)

    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineChargedWorkgroupNumber"
    
End Function


Public Sub UpdateChargedWorkgroupNumber(ByVal v_lngWgNum As Long)
    Dim l_lngWgNum As Long
    Dim l_blnDefDistDeptNum As Boolean
    

On Error GoTo eh
    
    If v_lngWgNum = WorkgroupNumber Then
        Exit Sub
    End If

    l_lngWgNum = v_lngWgNum
    If l_lngWgNum = 0 Then
        l_lngWgNum = DetermineChargedWorkgroupNumber
        l_blnDefDistDeptNum = True
    End If
    
' ensure we have a valid workgroup before continue
    If l_lngWgNum = 0 Then
        Err.Raise trErrEntryUpdateWorkgroup
    End If
    
    WorkgroupNumber = l_lngWgNum
    If l_blnDefDistDeptNum Then
        DetermineDistributionDepartmentNumber
    End If
    DetermineLocationCode
    If UnionCode = trUnion148 Then
        If ComponentCode = "" Then
            If IsPPMProjectEffective Then
                UpdateComponentCode (trInterplantCode)
            End If
        ElseIf ComponentCode = trInterplantCode Then
            If Not IsPPMProjectEffective Then
                UpdateComponentCode ("")
            End If
        End If
    End If
    Exit Sub
    
eh:
    Select Case Err.Number
        Case trErrEntryUpdateWorkgroup
            Err.Source = FormatEntryBRV
            TrsRaiseError mstrCMODULENAME, "UpdateChargedWorkgroupNumber", Err.Number
        Case Else
            TrsRaiseError mstrCMODULENAME, "UpdateChargedWorkgroupNumber"
    End Select

End Sub

Public Sub UpdateJobNumber(ByVal v_strJobNum As String)
' Dependent on exception code
    Dim l_strJobNum As String
    Dim lcolErrDescArgs As New Collection
    
On Error GoTo eh
        
    If Trim(v_strJobNum) = Trim(JobNumber) Then
        Exit Sub
    End If
    
    l_strJobNum = Trim(v_strJobNum)
    If l_strJobNum = "" Then
        l_strJobNum = DetermineJobNumber
    End If
        
    If l_strJobNum = "" Then
        Err.Raise trErrbusService, , "Could not determine job number"
    ElseIf Not GetJobManager.JobNumberIsValid(l_strJobNum, ChargeDate) Then
        lcolErrDescArgs.Add l_strJobNum
        lcolErrDescArgs.Add Format(ChargeDate, "mm/dd/yyyy")
        Err.Raise trErrEntryUpdateJob
    End If
        
    JobNumber = l_strJobNum
    
    If JobNumber = Employee.JobNumber(ChargeDate) Then
        RevertJobUpgrade
    Else
        ApplyJobUpgrade
    End If
    Exit Sub
    
eh:
    Select Case Err.Number
        Case trErrEntryUpdateJob
            Err.Source = FormatEntryBRV
            TrsRaiseError mstrCMODULENAME, "UpdateJobNumber", Err.Number, , lcolErrDescArgs
        Case Else
            TrsRaiseError mstrCMODULENAME, "UpdateJobNumber"
    End Select

End Sub

Public Sub UpdateExceptionCode(ByVal v_strExcCd As String)
'
' RED038
    Dim l_strExcCd As String
    Dim l_blnIsDateAHoliday As Boolean
    
On Error GoTo eh
    
    If Trim(v_strExcCd) = Trim(ExceptionCode) Then
        Exit Sub
    End If
    
    If IsCanvassNonCanvass And Not Me.Timecard.IsCallerOvertime Then
        Err.Raise trErrExcCdChgOTCanvass
    End If
    
    l_strExcCd = Trim(v_strExcCd)
    If l_strExcCd = "" Then
        l_strExcCd = DetermineExceptionCodePrivate
    End If
    
' Defer this validation until IsValid since determining if callout
' standby is dependent on having the override exception.
'    l_blnIsDateAHoliday = Timecard.IsDateAHoliday(ChargeDate)
'    If IsRegularScheduledHours And Not l_blnIsDateAHoliday Then
'        If Not IsExceptionValidInsideOfSchedule(l_strExcCd) Then
'            Err.Raise trErrExcCdNotValidInSched
'        End If
'    Else
'        If Not IsExceptionValidOutsideOfSchedule(l_strExcCd) Then
'            Err.Raise trErrExcCdNotValidOutSched
'        End If
'    End If
       
    If IsExceptionOvertimeRelated(ExceptionCode) Then
        If Not IsExceptionOvertimeRelated(l_strExcCd) Then
            If OvertimeOpportunityNumber >= trLegacyMinDefaultOTOppNum Then
                OvertimeOpportunityNumber = 0
                OvertimeReasonCode = ""
                DetermineOvertimeBankChargeDate (l_strExcCd)
            Else
                Err.Raise trErrExcCdChgOTCanvass
            End If
        End If
    Else
        If IsExceptionOvertimeRelated(l_strExcCd) Then
            If Not Me.Timecard.IsCallerOvertime Then
                OvertimeOpportunityNumber = trDefaultOTOppNum
                If OvertimeReasonCode = "" Then
                    OvertimeReasonCode = _
                        DefaultOvertimeReasonCode(Employee.DepartmentNumber(ChargeDate))
                End If
            End If
            DetermineOvertimeBankChargeDate (l_strExcCd)
        End If
    End If

    CaseNumber = 0

    If IsExceptionOvertimeRelated(l_strExcCd) Or _
            IsExceptionUnpaid(l_strExcCd, Employee.ExecContractIndicator(ChargeDate)) Then
        If Not l_blnIsDateAHoliday Then
            ReliefCode = ""
        End If
    End If
    
    ExceptionCode = l_strExcCd
    
'pts 345 to prevent HR on a schedule work day
    If IsRegularScheduledHours And IsHolidayOnRest(v_strExcCd) Then
       Err.Raise trErrExcCdNotValidInSched
    Else
    End If
    
    DetermineEarningsCode

    Exit Sub
    
eh:
    Select Case Err.Number
        Case trErrExcCdNotValidInSched, trErrExcCdNotValidOutSched, trErrExcCdChgOTCanvass
            Err.Source = FormatEntryBRV
            TrsRaiseError mstrCMODULENAME, "UpdateExceptionCode", Err.Number
        Case Else
            TrsRaiseError mstrCMODULENAME, "UpdateExceptionCode"
    End Select

End Sub

Public Function IsCanvassNonCanvass() As Boolean
   
On Error GoTo eh
   If OvertimeOpportunityNumber > 0 And _
            OvertimeOpportunityNumber < trLegacyMinDefaultOTOppNum Then
        IsCanvassNonCanvass = True
    End If
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsCanvassNonCanvass"

End Function

Public Sub UpdateScheduleChangeExceptionCode(ByVal v_strSchedChgExcCd As String)
' Dependent on UpdateExceptionCode
' WHITE062, WHITE063
    Dim l_strEarnCd As String
    Dim l_strExecConInd As String
    Dim lcolErrDescArgs As New Collection
     
On Error GoTo eh
    
    If Trim(v_strSchedChgExcCd) = Trim(ScheduleChangeExceptionCode) Then
        Exit Sub
    End If
    
    ScheduleChangeExceptionCode = Trim(v_strSchedChgExcCd)
    
    If Trim(v_strSchedChgExcCd) = "" Then
        l_strEarnCd = EarningsCodeForExceptionCode(v_strSchedChgExcCd, _
            Timecard.Employee.ExecContractIndicator(ChargeDate))
        ScheduleChangeEarningsCode = l_strEarnCd
        DetermineEarningsCode
    Else
        l_strExecConInd = _
            Timecard.Employee.ExecContractIndicator(ChargeDate)
        If Not IsValidScheduleChangeCode( _
                v_strSchedChgExcCd, _
                l_strExecConInd) Then
            lcolErrDescArgs.Add v_strSchedChgExcCd
            Select Case l_strExecConInd
                Case "E"
                    lcolErrDescArgs.Add "Executive"
                Case "C"
                    lcolErrDescArgs.Add "Contract"
                Case Else
                    lcolErrDescArgs.Add "ExecContractInd: " & l_strExecConInd
            End Select
            Err.Raise trErrSchedChgCdNotValid
        Else
            ' this may override the exception earnings code
            DetermineScheduleChangeEarningsCode
        End If
    End If
    Exit Sub
    
eh:
    Select Case Err.Number
        Case trErrSchedChgCdNotValid
            Err.Source = FormatEntryBRV
            TrsRaiseError mstrCMODULENAME, "UpdateScheduleChangeExceptionCode", Err.Number, , lcolErrDescArgs
        Case Else
            TrsRaiseError mstrCMODULENAME, "UpdateScheduleChangeExceptionCode"
    End Select
    
End Sub

Public Sub UpdateOverrideExceptionCode(ByVal v_strOverExcCd As String)
' Dependent on exception code
On Error GoTo eh
    
    If v_strOverExcCd = OverrideExceptionCode Then
        Exit Sub
    End If
    
    OverrideExceptionCode = v_strOverExcCd
    
'    DetermineDefaultAccounting - BLUE016 Iteration 3
    
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "UpdateOverrideExceptionCode"
    
End Sub


Public Sub UpdateComponentCode(ByVal v_strCompCd As String)
On Error GoTo eh
    
    If Trim(v_strCompCd) = Trim(ComponentCode) Then
        Exit Sub
    End If
    
    ComponentCode = v_strCompCd
    DetermineComponentAmount
    
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "UpdateComponentCode"
    
End Sub

Public Sub UpdateDistributionDepartementNumber(ByVal v_strDistDeptNum As String)
    Dim l_strDistDeptNum As String

On Error GoTo eh
    
    If Trim(v_strDistDeptNum) = Trim(DistributionDepartmentNumber) Then
        Exit Sub
    End If
    
    l_strDistDeptNum = Trim(v_strDistDeptNum)
    If l_strDistDeptNum = "" Then
        l_strDistDeptNum = DetermineDistributionDepartmentNumberPrivate
    End If
    
    DistributionDepartmentNumber = l_strDistDeptNum
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "UpdateDistributionDepartementNumber"
    
End Sub

Public Sub DetermineLocationCode()
'
' Dependent on DetermineChargedWorkgroupNumber
'
' Set the location code when necessary.

    Dim l_intLocCd As Integer
    
    On Error GoTo eh
    
    l_intLocCd = _
        Timecard.Employee.GetWorkgroupLocationCode(WorkgroupNumber)
        
' if workgroup is not a primary or secondary, just get the location
    If l_intLocCd = 0 Then
        l_intLocCd = GetLocationCode(WorkgroupNumber)
    End If
            
    If l_intLocCd <> LocationCode Then
        LocationCode = l_intLocCd
    End If
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineLocationCode"
    
End Sub

Private Function GetLocationCode(ByVal v_lngWgNum As Long) As Integer
    Dim lobjWGMgr As dbWorkgroup.cdWorkGroup
    
On Error GoTo eh
    
    Set lobjWGMgr = CreateObject("dbWorkgroup.cdWorkGroup")
    GetLocationCode = lobjWGMgr.LocationCode(v_lngWgNum)
    
    Set lobjWGMgr = Nothing
    Exit Function
    
eh:
    Set lobjWGMgr = Nothing
    TrsRaiseError mstrCMODULENAME, "GetLocationCode"

End Function


Public Sub DetermineDistributionDepartmentNumber()
'
'Dependent on DetermineChargedWorkgroupNumber
'
'BLUE018
    On Error GoTo eh
    
    If Not IsDOJMWorkgroupNumber(WorkgroupNumber) And _
        DepartmentNumber <> "086" And _
        (Not GetBusDepartmentManager.IsDOJMDepartmentNumber(DepartmentNumber)) Then
        DistributionDepartmentNumber = DepartmentNumber
    Else
        DistributionDepartmentNumber = WorkgroupDepartmentNumber
    End If
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineDistributionDepartmentNumber"
    
End Sub

Private Function DetermineDistributionDepartmentNumberPrivate() As String
' 9/9/00 rsg
' This is temporary name used with update until merged with new time.
'
'Dependent on WorkgroupNumber and DepartmentNumber
'
'BLUE018
On Error GoTo eh
    
    If Not IsDOJMWorkgroupNumber(WorkgroupNumber) And _
        DepartmentNumber <> "086" And _
        (Not GetBusDepartmentManager.IsDOJMDepartmentNumber(DepartmentNumber)) Then
        DetermineDistributionDepartmentNumberPrivate = DepartmentNumber
    Else
        DetermineDistributionDepartmentNumberPrivate = WorkgroupDepartmentNumber
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineDistributionDepartmentNumberPrivate"
    
End Function

' REMOVED 3/14/2002
'Public Sub DetermineDistributionId()


Public Sub DetermineDistributionDescription()
' Override distributionId with exception description
' Dependent on DetermineDistributionDepartmentNumber
'  and DetermineChargedWorkgroupNumber
'  and DetermineExceptionCode
'
    Dim l_strDistDeptNum As String
    Dim l_strExcepDesc As String
    
    On Error GoTo eh

    If Trim(DistributionType) = "" Then
        DistributionId = GetExceptionDescription()
    End If
    
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineDistributionDescription"

End Sub

' REMOVED 3/14/2002
'This function is now useless in Timcard
'busDistributionMgr determines default accounting prior
' to busTimeCard being called
'Private Function DetermineDefaultAccountNumber() As String


Private Function DetermineDistributionType(ByVal v_strDistributionId) As String
    
    Dim l_strUtility As String
    Dim l_strDivision As String
    Dim l_strMajor As String
    Dim l_strMinor As String
    Dim l_strNEC As String
    Dim l_strBudget As String
    Dim l_strTran As String
    Dim l_strWorkOrder As String
    Dim l_strDistrict As String
    Dim l_strWorkRequest As String
    Dim l_strIdentifier As String
    Dim l_strType As String
    
On Error GoTo eh

'I HAVE COMMENTED OUT THE CODE IN THIS PROCEDURE SINCE THIS PROCEDURE IS UNUSED AND THERE IS
'NOT A WAY TO DETERMINE IF THE DISTRIBUTION IS AN EMPRV JOB TASK - L. BERTKE

    'l_strUtility = Trim(Mid(v_strDistributionId, 1, 1))
    'l_strDivision = Trim(Mid(v_strDistributionId, 2, 2))
    'l_strMajor = Trim(Mid(v_strDistributionId, 4, 3))
    'l_strMinor = Trim(Mid(v_strDistributionId, 7, 3))
    'l_strNEC = Trim(Mid(v_strDistributionId, 10, 3))
    'l_strBudget = Trim(Mid(v_strDistributionId, 13, 3))
    'l_strTran = Trim(Mid(v_strDistributionId, 16, 1))
    'l_strWorkOrder = Trim(Mid(v_strDistributionId, 17, 5))
    
    'l_strDistrict = Trim(Mid(v_strDistributionId, 1, 4))
    'l_strIdentifier = Trim(Mid(v_strDistributionId, 5, 3))
    'l_strWorkRequest = Trim(Mid(v_strDistributionId, 8, 6))
    
    'If ((l_strIdentifier = "WR") And (l_strDistrict <> "") And (l_strWorkRequest <> "")) Or _
        ((l_strDistrict <> "") And (l_strWorkRequest <> "") And (Trim(Mid(v_strDistributionId, 14, 15))) = "") Then
            'DetermineDistributionType = "W"
    'Else
        'If (l_strUtility = "" And l_strDivision = "" And l_strMajor = "" And l_strMinor = "" And _
            l_strNEC = "" And l_strBudget = "" And l_strTran = "") And (l_strWorkOrder <> "") Then
                'DetermineDistributionType = "S"
        'Else
            ' Utility must be 1 (electric), 2 (gas), 3 (water), 4 (steam)
            'If (l_strUtility <> "" And l_strDivision <> "" And l_strMajor <> "") Then
                'DetermineDistributionType = "A"
            'Else
                DetermineDistributionType = ""
            'End If
        'End If
    'End If
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "DetermineDistributionType"
                
End Function

Private Function GetExceptionAccountNumber() As String
    Dim l_strAcctNum As String

    On Error GoTo eh
    
    If Trim(OverrideExceptionCode) <> "" Then
        l_strAcctNum = OverrideExceptionAccountNumber
    Else
        l_strAcctNum = ExceptionAccountNumber
    End If
    GetExceptionAccountNumber = l_strAcctNum
'New requirement, just need to know if it exists
'  to determine if the exception description is placed in
' the distribution id field
'
'    If l_strAcctNum = l_strCACCOUNTNOTEXISTS Or _
'        l_strAcctNum = l_strCACCOUNTNOTAPPLY Then
'        GetExceptionAccountNumber = l_strAcctNum
'    Else
'        GetExceptionAccountNumber = FormatAccount( _
'            l_strAcctNum, GetDepartmentManager.GetAccountDivision(v_strDistDeptNum))
'    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "GetExceptionAccountNumber"

End Function

Private Function GetExceptionDescription() As String
    
    Dim l_strExcpDesc As String

    On Error GoTo eh
    
    If Trim(OverrideExceptionCode) <> "" Then
        l_strExcpDesc = ExceptionDescriptionFor(OverrideExceptionCode)
    Else
        l_strExcpDesc = ExceptionDescriptionFor(ExceptionCode)
    End If
    GetExceptionDescription = l_strExcpDesc
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "GetExceptionDescription"

End Function

Private Function OverrideExceptionAccountNumber() As String
    OverrideExceptionAccountNumber = _
        ExceptionAccountNumberFor(OverrideExceptionCode)
End Function

Private Function ExceptionAccountNumber() As String
    ExceptionAccountNumber = _
        ExceptionAccountNumberFor(ExceptionCode)
End Function

Private Function FormatAccount(ByVal v_strAcctNum As String, _
            ByVal v_strAcctDivNum As String) As String
'Assuming division is 2 char
    On Error GoTo eh
    
    FormatAccount = Left(v_strAcctNum, 1) & v_strAcctDivNum & _
            Mid(v_strAcctNum, 4)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "FormatAccount"
    
End Function

Private Function WorkgroupDepartmentNumber() As String
'
' Dependent on DetermineChargedWorkgroupNumber
'
' BLUE056
    Dim lobjWGMgr As dbWorkgroup.cdWorkGroup
    
On Error GoTo eh
    
    Set lobjWGMgr = CreateObject("dbWorkgroup.cdWorkGroup")
    WorkgroupDepartmentNumber = lobjWGMgr.DepartmentNumber(WorkgroupNumber)
    
    Set lobjWGMgr = Nothing
    Exit Function
    
eh:
    Set lobjWGMgr = Nothing
    TrsRaiseError mstrCMODULENAME, "WorkgroupDepartmentNumber"

End Function

Public Sub DetermineExceptionCode()
'
' Dependent on DetermineDistributionDepartmentNumber
'
' BLUE030, RED057
    'l_blnIsDateAHoliday is used to determine if a holiday was reported
    'on that day
'    Dim l_blnIsDateAHoliday As Boolean
    On Error GoTo eh
    
'    l_blnIsDateAHoliday = Timecard.IsDateAHoliday(ChargeDate)
    If IsRegularScheduledHours And Not Timecard.DoesTimeEntryOverlapWithHolidayOrVacation(SafeRef(Me)) Then
        ExceptionCode = DetermineStraightTimeExceptionCode
    Else
        ExceptionCode = _
            DefaultOvertimeRate(Employee.UnionCode(ChargeDate))
        If OvertimeReasonCode = "" Then
            OvertimeReasonCode = _
                DefaultOvertimeReasonCode(Employee.DepartmentNumber(ChargeDate))
        End If
        OvertimeOpportunityNumber = trDefaultOTOppNum
        ' apply exceptions
'        If Employee.IsExecutive(ChargeDate) Then
'            If Employee.IsAmerenLeadership(ChargeDate) Then
'                ExceptionCode = "ONP"
'            ElseIf Me.Employee.IsExecutiveShiftWorker(ChargeDate) And Me.Employee.OTType = "Q" Then
'                If Timecard.IsDateAHoliday(ChargeDate) Then
'                    ExceptionCode = "OT6"
'                End If
'            End If
'        End If
        If ExceptionCode = "ONP" Then
            OvertimeOpportunityNumber = 0
            OvertimeReasonCode = ""
        Else
'            If Trim(OvertimeReasonCode) = "" Then
'                OvertimeReasonCode = _
'                    DefaultOvertimeReasonCode(Employee.DepartmentNumber(ChargeDate))
'            End If
        End If
    End If
    
    DetermineOvertimeBankChargeDate (ExceptionCode)
    DetermineEarningsCode
    
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineExceptionCode"
    
End Sub

Private Function DetermineExceptionCodePrivate() As String
' 9/9/00 rsg
' This is temporary name used with update until merged with new time.
'
' Dependent on DetermineDistributionDepartmentNumber
'
' BLUE030, RED057
    Dim l_strExcCd As String

    On Error GoTo eh
    
    l_strExcCd = "ST"
    If IsRegularScheduledHours Then
        If ScheduledSundayPremiumApplies Then
            If IsCallowayMaintenanceDepartment(DistributionDepartmentNumber) Then
                l_strExcCd = "SSC"
            ElseIf IsSelfTradedShift And UnionCode = trUnion148 Then
                If Not ScheduledBeforeTrade Then
                    l_strExcCd = "ST"
                Else
                    l_strExcCd = "SUS"
                End If
            Else
                l_strExcCd = "SUS"
            End If
        ElseIf ScheduledSaturdayPremiumApplies Then
            l_strExcCd = "SAS"
        Else
            l_strExcCd = "ST"
        End If
    Else
        l_strExcCd = _
            DefaultOvertimeRate(Employee.UnionCode(ChargeDate))
        ' apply exceptions
        If Employee.IsExecutive(ChargeDate) Then
            If Employee.IsAmerenLeadership(ChargeDate) Then
                l_strExcCd = "ONP"
            ElseIf Me.Employee.IsExecutiveShiftWorker(ChargeDate) And Me.Employee.OTType = "Q" Then
                If Timecard.IsDateAHoliday(ChargeDate) Then
                    l_strExcCd = "OT6"
                End If
            End If
        End If
    End If
    
    DetermineExceptionCodePrivate = l_strExcCd
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineExceptionCodePrivate"
    
End Function

'
'Private Function IsHoliday() As Boolean
'' cache the holidays later
'    Dim l_cdCodeTable As dbCodeTables.cdCodeTables
'    Dim l_rsHolidays As ADODB.Recordset
'    Dim l_objCtx As ObjectContext
'    Dim l_datStart As Date
'    Dim l_datEnd As Date
'
'On Error GoTo eh
'    l_datStart = DateAdd("yyyy", -1, ChargeDate)
'    l_datEnd = DateAdd("yyyy", 1, ChargeDate)
'
'    Set l_objCtx = GetObjectContext
'    Set l_cdCodeTable = CreateObject("dbCodeTables.cdCodeTables")
'
'    Set l_rsHolidays = _
'        l_cdCodeTable.GetHolidaysBetween(l_datStart, l_datEnd)
'    If Not (l_rsHolidays.BOF And l_rsHolidays.EOF) Then
'        l_rsHolidays.Filter = _
'            "UNION_CD = " & "'" & Employee.UnionCode(ChargeDate) & _
'            "'" & " AND HOLIDAY_DATE = " & "#" & _
'            Format(ChargeDate, "mm/dd/yyyy") & "#"
'        IsHoliday = l_rsHolidays.RecordCount > 0
'    Else
'        IsHoliday = False
'    End If
'    Exit Function
'
'eh:
'    TrsRaiseError mstrCMODULENAME, "IsHoliday"
'
'End Function

Private Function DefaultOvertimeRate(ByVal v_strUnionCode As String) As String
'RED061, GOLD001
    Dim l_strUnionCode As String
    
    l_strUnionCode = Trim(UCase(v_strUnionCode))
    
    If Timecard.Employee.IsExecutive(ChargeDate) Then
        DefaultOvertimeRate = DefaultExecutiveOvertimeRate(l_strUnionCode)
    Else
        Select Case l_strUnionCode
            Case "BA", "CA", "DA", "EA", "GA", "HA", "KA", "MA", "NA", _
                    "PA", "QA", "SA", "VA", "WA", _
                    "BW", "NE", "NT", "ND"
                DefaultOvertimeRate = "OT2"
            Case "FA", "IA", "JA", "TA", "JB", "TB"
                DefaultOvertimeRate = "OT4"
            Case "LA", "OA"
                DefaultOvertimeRate = "OT5"
            Case Else
                DefaultOvertimeRate = "OT2"
        End Select
    End If

End Function

Private Function DefaultOvertimeReasonCode(ByVal v_strDeptNum As String) As String
'GOLD004
On Error GoTo eh
    DefaultOvertimeReasonCode = _
        GetDepartmentManager.GetOvertimeReasonCode(v_strDeptNum)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "DefaultOvertimeReasonCode"
End Function

Public Sub DetermineEarningsCode()
'
' Dependent on DetermineExceptionCode
'
' Set the earnings code when necessary.
' BLUE096
    Dim l_strEarnCd As String
    
On Error GoTo eh
    
    l_strEarnCd = _
        EarningsCodeForExceptionCode(ExceptionCode, _
            Timecard.Employee.ExecContractIndicator(ChargeDate))
            
    If l_strEarnCd <> EarningsCode Then
        EarningsCode = l_strEarnCd
    End If
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineEarningsCode"

End Sub


Public Sub DetermineOvertimeBankChargeDate(ByVal v_strExcCd As String)
'
' Dependent on DetermineExceptionCode
'
' Set the overtime bank charge date, when necessary.
' RED060
    Dim ldatOTBankChgDt As Date
    Dim lrsOTBankChgDtInfo As ADODB.Recordset
    Dim lobjOTEqualizer As busOvertimeEqualizer.cbsOTEqualizer
    
On Error GoTo eh
    
    If IsExceptionOvertimeBankRelated(v_strExcCd) Then
        Set lrsOTBankChgDtInfo = _
            Employee.GetOTBankChargeInfoByChargeDate( _
                EmployeeNumber, ChargeDate, ChargeDate)
        Set lobjOTEqualizer = CreateObject("busOvertimeEqualizer.cbsOTEqualizer")
        ldatOTBankChgDt = _
            lobjOTEqualizer.GetBankChargeDate(lrsOTBankChgDtInfo)
        Set lobjOTEqualizer = Nothing
        Set lrsOTBankChgDtInfo = Nothing
    Else
        ldatOTBankChgDt = CDate(CDEFAULTDATE)
    End If
    
    If ldatOTBankChgDt <> OvertimeBankChargeDate Then
        OvertimeBankChargeDate = ldatOTBankChgDt
    End If
    Exit Sub
    
eh:
    If Not lobjOTEqualizer Is Nothing Then
        Set lobjOTEqualizer = Nothing
    End If
    If Not lrsOTBankChgDtInfo Is Nothing Then
        Set lrsOTBankChgDtInfo = Nothing
    End If
    TrsRaiseError mstrCMODULENAME, "DetermineOvertimeBankChargeDate"
    
End Sub


Public Sub DetermineTradeExceptionCode()
'
' Dependent on DetermineExceptionCode
' BLUE029
On Error GoTo eh

    TradeExceptionCode = _
        Timecard.Schedule.TradeExceptionCode( _
            ChargeDate, StartTimeAsDate, EndTimeAsDate)
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineTradeExceptionCode"
        
End Sub

Public Sub DetermineVacationReliefCode()
' BLUE064
    Dim l_strRelCd As String
    Dim l_colErrDescArgs As Collection

On Error GoTo eh
    
    l_strRelCd = _
        Timecard.Schedule.ReliefCode( _
            ChargeDate, StartTimeAsDate, EndTimeAsDate)
            
    If l_strRelCd <> "" Then
        If Not IsEligibleForVacationRelief(Employee.UnionCode(ChargeDate)) Then
            Set l_colErrDescArgs = New Collection
            l_colErrDescArgs.Add l_strRelCd
            l_colErrDescArgs.Add ExceptionCode
            AppendWarning trWarDiscardReliefComp, , l_colErrDescArgs
            Set l_colErrDescArgs = Nothing
            l_strRelCd = ""
        End If
    End If
    
    ReliefCode = l_strRelCd
    Exit Sub
    
eh:
    Set l_colErrDescArgs = Nothing
    TrsRaiseError mstrCMODULENAME, "DetermineVacationReliefCode"
End Sub

Public Sub DetermineOvertimeReasonCode()
'
' Dependent on DetermineExceptionCode
'
On Error GoTo eh

    If IsExceptionOvertimeRelated(ExceptionCode) Then
        OvertimeReasonCode = _
            DefaultOvertimeReasonCode(Employee.DepartmentNumber(ChargeDate))
    End If
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineOvertimeReasonCode"
    
End Sub

Public Sub DetermineComponentCode()
'
' Dependent on DetermineChargedWorkgroup
' BLUE085
On Error GoTo eh

    If IsPPMProjectEffective Then
        If IsJobEligibleForPPMComponent Then
            ComponentCode = "IC"
            DetermineComponentAmount
        End If
    End If
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineComponentCode"
    
End Sub

Private Function IsPPMProjectEffective() As Boolean
' WHITE092
' 9/11/00 rsg changed implementation to use ChargeDate instead of DateWorked.
    Dim l_varPeriod As Variant

On Error GoTo eh

    l_varPeriod = Timecard.Employee.GetPPMProjectPeriod(WorkgroupNumber)
    If Not IsEmpty(l_varPeriod) Then
        If IsValidPeriod(l_varPeriod) Then
            If ChargeDate >= l_varPeriod(0) And _
                ChargeDate <= l_varPeriod(1) Then
                IsPPMProjectEffective = True
                Exit Function
            End If
        End If
    End If
    IsPPMProjectEffective = False
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsPPMProjectEffective"
    
End Function

Private Function IsValidPeriod(r_varPeriod As Variant) As Boolean
' Is r_varPeriod a 1 based array with 2 date elements

On Error GoTo eh
    
    If IsArray(r_varPeriod) Then
        If LBound(r_varPeriod) = 0 And _
            UBound(r_varPeriod) = 1 Then
            If IsDate(r_varPeriod(0)) And _
                IsDate(r_varPeriod(1)) Then
                If CDate(r_varPeriod(0)) <= CDate(r_varPeriod(0)) Then
                    IsValidPeriod = True
                    Exit Function
                End If
            End If
        End If
    End If
    IsValidPeriod = False
    Exit Function
        
eh:
    TrsRaiseError mstrCMODULENAME, "IsValidPeriod"
    
End Function

Public Sub DetermineShiftComponentCode()
' Dependent on DetermineExceptionCode
' BLUE063
    If IsEligibleForShiftComponent Then
        ShiftComponentCode = _
            Timecard.Schedule.ShiftComponentOnJobNum( _
                ChargeDate, _
                StartTimeAsDate, _
                EndTimeAsDate, _
                JobNumber)
            
        DetermineShiftComponentAmount
    End If
End Sub

Public Sub DetermineShiftComponentAmount()
'
    Dim l_varAmt As Variant
    Dim l_strShiftComp As String
    Dim lcolErrDescArgs As New Collection
    
On Error GoTo eh
    
    l_strShiftComp = ShiftComponentCode
    If ShiftComponentCode = "" Then
        ShiftComponentAmount = CDec(0)
        Exit Sub
    End If

    l_varAmt = ShiftComponentAmountFor(UnionCode, ChargeDate, l_strShiftComp)
    If Not IsNull(l_varAmt) Then
        ShiftComponentAmount = CDec(l_varAmt)
    Else
        ShiftComponentAmount = CDec(0)
        ShiftComponentCode = ""
        lcolErrDescArgs.Add l_strShiftComp
        lcolErrDescArgs.Add ChargeDate
        AppendWarning trWarShiftCompAmtNotFound, , lcolErrDescArgs
        Set lcolErrDescArgs = Nothing
    End If
    Exit Sub
        
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineShiftComponentAmount"
    
End Sub

Public Sub DetermineComponentAmount()
'
' Dependent on DetermineComponentCode and UnionCode
    Dim l_varAmt As Variant
    Dim lcolErrArgs As New Collection
    
On Error GoTo eh
    If ComponentCode = "" Then
        ComponentAmount = CDec(0)
        Exit Sub
    End If

    l_varAmt = ComponentAmountFor(UnionCode, ChargeDate, ComponentCode)
    If Not IsNull(l_varAmt) Then
        ComponentAmount = CDec(l_varAmt)
    Else
        lcolErrArgs.Add UnionCode
        lcolErrArgs.Add ChargeDate
        lcolErrArgs.Add ComponentCode
        Err.Raise trErrbusCompAmountNotFound
    End If
    Exit Sub
    
eh:
    Select Case Err.Number
        Case trErrbusCompAmountNotFound
            Err.Source = FormatEntryBRV
            TrsRaiseError mstrCMODULENAME, "DetermineComponentAmount", Err.Number, , lcolErrArgs
        Case Else
            TrsRaiseError mstrCMODULENAME, "DetermineComponentAmount"
    End Select
    
End Sub

Private Function DetermineJobNumber() As String
    Dim l_strSchedJobNum As String
    Dim l_strBaseJobNum As String
    Dim l_colErrArgs As New Collection
' BLUE081

On Error GoTo eh

    l_strBaseJobNum = Me.Timecard.Employee.JobNumber(ChargeDate)
    
    'debug
    l_colErrArgs.Add "Base job number: " & l_strBaseJobNum
    
    l_strSchedJobNum = Me.Timecard.Schedule.ScheduledJobUpgrade(ChargeDate, StartTimeAsDate, EndTimeAsDate)
    
    'debug
    l_colErrArgs.Add "Scheduled job number: " & l_strSchedJobNum
    
    If l_strSchedJobNum <> "" Then
        DetermineJobNumber = l_strSchedJobNum
    Else
        DetermineJobNumber = l_strBaseJobNum
    End If
    Set l_colErrArgs = Nothing
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineJobNumber", , , , l_colErrArgs
    
End Function


Public Sub SetDefaultJobNumber()
    Dim l_strJobNum As String
    Dim l_colErrDescArgs As New Collection
' BLUE081

On Error GoTo eh

    l_strJobNum = DetermineJobNumber
    ' job upgrade
    If l_strJobNum <> Employee.JobNumber(ChargeDate) Then
        If IsEligibleForJobUpgrade(Employee.UnionCode(ChargeDate)) Then
            UpdateJobNumber l_strJobNum
        Else
            'warning
            l_colErrDescArgs.Add l_strJobNum
            l_colErrDescArgs.Add ExceptionCode
            AppendWarning trWarDiscardJobUpgrade, , l_colErrDescArgs
            UpdateJobNumber Employee.JobNumber(ChargeDate)
        End If
    Else
        UpdateJobNumber l_strJobNum
    End If
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "SetDefaultJobNumber"
    
End Sub

Public Sub DetermineJobUpgrade()
' Dependent on DetermineJobNumber
' BLUE020
    Dim l_strBaseJobNum As String
    
    On Error GoTo eh
    
    l_strBaseJobNum = Employee.JobNumber(ChargeDate)
    If l_strBaseJobNum <> JobNumber Then
        ApplyJobUpgrade
    End If
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineJobUpgrade"

End Sub

Public Sub DeterminePayrollDepartmentNumber()

    On Error GoTo eh
    
    DepartmentNumber = _
        Timecard.Employee.DepartmentNumber(ChargeDate)
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DeterminePayrollDepartmentNumber"
        
End Sub

Public Sub DetermineUnionCode()
'
' Dependent on DetermineDepartmentNumber
'
    UnionCode = GetEffectiveUnionCode
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineUnionCode"
    
End Sub

Public Sub DetermineScheduleChangeExceptionCode()

    On Error GoTo eh
    
    ScheduleChangeExceptionCode = ScheduleChangePaymentScheduled
    
    DetermineScheduleChangeEarningsCode
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineScheduleChangeExceptionCode"
    
End Sub

Public Sub DetermineScheduleChangeEarningsCode()
' Dependent on schedule change exception code and
'   must follow DetermineEarningsCode.
' BLUE087, WHITE063
    Dim l_strSchedChgEarnCd As String
    
    On Error GoTo eh
    
    l_strSchedChgEarnCd = _
        EarningsCodeForExceptionCode(ScheduleChangeExceptionCode, _
            Timecard.Employee.ExecContractIndicator(ChargeDate))
            
    If l_strSchedChgEarnCd <> ScheduleChangeEarningsCode Then
        ScheduleChangeEarningsCode = l_strSchedChgEarnCd
    End If
'
' Override the earnings code when a schedule change earnings exists
'
    If l_strSchedChgEarnCd <> "" Then
        EarningsCode = ScheduleChangeEarningsCode
    Else
        If Trim(ScheduleChangeExceptionCode) <> "" Then
            AppendWarning trWarShiftCompRemoved, "Schedule Change Payment (" & ScheduleChangeExceptionCode & ") is not valid for union (" & Employee.UnionCode(ChargeDate) & "). Schedule Change Payment was removed"
            ScheduleChangeExceptionCode = ""
        End If
    End If
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineScheduleChangeEarningsCode"
    
End Sub

Public Sub DetermineScheduleCode()
    Dim l_aryShiftDetails As Variant

    l_aryShiftDetails = _
        Timecard.Schedule.ShiftDetailsOn( _
                ChargeDate, _
                StartTimeAsDate, _
                EndTimeAsDate)
    Me.ScheduleCode = l_aryShiftDetails(0)
    Me.ScheduleSourceCode = l_aryShiftDetails(1)

    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineScheduleCode"
End Sub

'Private Sub ApplyJobNumber(ByVal v_strJobNum As String, _
'        ByVal v_blnJobUpgrade As Boolean)
'
'    JobNumber = v_strJobNum
'
'    If v_blnJobUpgrade Then
'        If Trim(OverrideExceptionCode) = "" Then
'            OverrideExceptionCode = "JC"
'        Else
'            AppendEntryError vbObjectError + 30000, mstrCBUSINESSWAR, "Warning: Override already exists - JC not applied"
'        End If
'    Else
'        If Trim(OverrideExceptionCode) = "JC" Then
'            OverrideExceptionCode = ""
'        End If
'    End If
'End Sub

Private Sub ApplyJobUpgrade()
    Dim l_strSchedJob As String
    Dim l_colErrDescArgs As New Collection
        
On Error GoTo eh
    If Not IsEligibleForJobUpgrade(Employee.UnionCode(ChargeDate)) Then
        l_strSchedJob = Trim(Me.Timecard.Schedule.ScheduledJobUpgrade( _
            ChargeDate, StartTimeAsDate, EndTimeAsDate))
        If l_strSchedJob <> "" And JobNumber = l_strSchedJob Then
            JobNumber = Employee.JobNumber(ChargeDate)
            RevertJobUpgrade
            l_colErrDescArgs.Add l_strSchedJob
            l_colErrDescArgs.Add ExceptionCode
            AppendWarning trWarDiscardJobUpgrade, , l_colErrDescArgs
            Set l_colErrDescArgs = Nothing
            Exit Sub
        Else
            Err.Raise trErrbusNotEligibleJobUpgrade
        End If
    ElseIf Not IsJobUpgradeValid(JobNumber) Then
        Err.Raise trErrbusInvalidJobUpgrade
    End If
    
    If OverrideExceptionCode = "" Then
        OverrideExceptionCode = trOverExcJobChange
    End If
    Exit Sub
    
eh:
    Set l_colErrDescArgs = Nothing
    Select Case Err.Number
        Case trErrbusInvalidJobUpgrade, trErrbusNotEligibleJobUpgrade
            Err.Source = FormatEntryBRV
            TrsRaiseError mstrCMODULENAME, "ApplyJobUpgrade", Err.Number
        Case Else
            TrsRaiseError mstrCMODULENAME, "ApplyJobUpgrade"
    End Select
End Sub

Private Sub RevertJobUpgrade()

    If OverrideExceptionCode = trOverExcJobChange Then
        OverrideExceptionCode = ""
    End If

End Sub

Private Function ScheduledSaturdayPremiumApplies() As Boolean
    
    On Error GoTo eh
    
    If (IsScheduledSaturday Or IsExceptionToPremiumPayRules) _
        And IsEntitledToSaturdayPremium Then
    'continue
    Else
        ScheduledSaturdayPremiumApplies = False
        Exit Function
    End If
    
    If IsExceptionToPremiumPayRules Then
        ScheduledSaturdayPremiumApplies = Weekday(DateWorked) = 7
    Else
        ScheduledSaturdayPremiumApplies = IsScheduledSaturday
    End If

    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ScheduledSaturdayPremiumApplies"
        
End Function

Private Function ScheduledSundayPremiumApplies() As Boolean

    On Error GoTo eh
    Dim l_blnSundayPremiumApplies As Boolean
    
    If IsScheduledSunday Or IsExceptionToPremiumPayRules Then
    'continue
    Else
        ScheduledSundayPremiumApplies = False
        Exit Function
    End If
    
    If IsEntitledToSundayPremium Then
    'continue
    Else
        ScheduledSundayPremiumApplies = False
        Exit Function
    End If
    
    If IsEligibleCIPSJobForSundayPremium Then
    'continue
    Else
        ScheduledSundayPremiumApplies = False
        Exit Function
    End If
    
    If IsExceptionToPremiumPayRules Then
        ScheduledSundayPremiumApplies = Weekday(DateWorked) = 1
    Else
        ScheduledSundayPremiumApplies = IsScheduledSunday
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ScheduledSundayPremiumApplies"
        
End Function

Private Function IsEntitledToSaturdayPremium() As Boolean
    
    On Error GoTo eh
    
    IsEntitledToSaturdayPremium = _
        Me.Timecard.Employee.IsEntitledToSaturdayPremium
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsEntitledToSaturdayPremium"
        
End Function

Private Function IsEntitledToSundayPremium() As Boolean
    
    On Error GoTo eh
    
    IsEntitledToSundayPremium = _
        Me.Timecard.Employee.IsEntitledToSundayPremium
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsEntitledToSundayPremium"
        
End Function

Public Function IsAllTimeWithinRegularScheduledHours() As Boolean
    
    On Error GoTo eh
    
    IsAllTimeWithinRegularScheduledHours = _
        Not IsOutsideRegularScheduledHours
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsAllTimeWithinRegularScheduledHours"
        
End Function

Public Function IsAllTimeOutsideRegularScheduledHours() As Boolean
        
    On Error GoTo eh
    
    IsAllTimeOutsideRegularScheduledHours = _
        Not IsRegularScheduledHours
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsAllTimeOutsideRegularScheduledHours"
                
End Function


Public Function IsStraightTime() As Boolean
    
    On Error GoTo eh
    
    Select Case ExceptionCode
        Case "ST", "STP", "SAS", "SUS", "SSC"
            IsStraightTime = True
        Case Else
            IsStraightTime = False
    End Select
    
'    IsStraightTime = _
'        IsStringInList(ExceptionCode, _
'            Array("ST", "STP", "SAS", "SUS", "SSC"))
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsStraightTime"
            
End Function

Public Function IsReliefTime() As Boolean
    
On Error GoTo eh
    Select Case ExceptionCode
        Case "RW1", "RW2", "RW3"
            IsReliefTime = True
        Case Else
            IsReliefTime = False
    End Select
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsReliefTime"
                    
End Function

Public Function IsOvertime() As Boolean
    On Error GoTo eh
    
    IsOvertime = IsExceptionOvertimeRelated(ExceptionCode)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsOvertime"
    
End Function

Public Function IsVacation() As Boolean
    
    On Error GoTo eh
    
    IsVacation = (ExceptionCode = "V")
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsVacation"
    
End Function

Public Function IsExceptionTime() As Boolean
' BLUE075
    On Error GoTo eh
    
    If IsStraightTime Or IsOvertime Then
        IsExceptionTime = False
    Else
        IsExceptionTime = True
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsVacation"
End Function

Public Function IsNonWorkTime() As Boolean
' BLUE062,WHITE052
' 8/4/00 rsg - there is a difference between the list in each rule
' above.  using list in WHITE052 as it matches the code.
    IsNonWorkTime = _
        IsStringInList(ExceptionCode, _
            Array("V", "E", "EF", "EB", "EC", "PB", "REL", "FH", "DF", _
            "PPT", "H", "J", "AD", "ED", "EDP", "AE", "JIE", "HR", "VH"))
End Function

Public Function IsEligibleForComponent() As Boolean
' WHITE052
' Currently, the rule is the same as shift component.
    IsEligibleForComponent = IsEligibleForShiftComponent
End Function

Public Function IsEligibleForShiftComponent() As Boolean

'Changes to IsEligibleForShiftComponent will cause changes to IsValidShiftComponentCode
' BLUE013, WHITE052
'AmerenCILCO Local 8 employees do not get a shift component for illness
'AmerenIP employees do not get shift component for time not worked
'Employees that are a CIPS benefits participant do not get shift component for time not worked
'The CIPS exception is Newton with time reported to the following 2 exception codes
'       Injury on day of straight time (A)
'       Rest Period (RP)

' Callaway Security Does not get shift components for Turnover time (NRC,NRT)
' and security drills.  These are indetified by override code (rule 213)


    Dim l_blnCILCOLocal8Employee As Boolean
    Dim l_blnCILCOEmployee As Boolean
    Dim l_blnCallawaySecurity As Boolean

    On Error GoTo eh
    
    l_blnCallawaySecurity = Me.Employee.IsCallawaySecurity(ChargeDate)
    If l_blnCallawaySecurity Then
        If Not CallawaySecurityShiftComponentEligible(OverrideExceptionCode) Then
            IsEligibleForShiftComponent = False
            Exit Function
        End If
    End If
    
    If IsWorkingException(ExceptionCode) Then
        IsEligibleForShiftComponent = True
        Exit Function
    End If
    
    If Timecard.Employee.IsExecutive(ChargeDate) Or l_blnCallawaySecurity Then
        IsEligibleForShiftComponent = False
        Exit Function
    End If
       
    l_blnCILCOLocal8Employee = Me.Employee.IsAmerenCILCOLocal8Employee(ChargeDate)
    If l_blnCILCOLocal8Employee Then
        IsEligibleForShiftComponent = IsExceptionIsEligibleForShiftComponent
        Exit Function
    End If
    
    If Timecard.Employee.IsAmerenUEBenefitParticipant(ChargeDate) Then
        IsEligibleForShiftComponent = True
    Else
        If Not IsWorkingException(ExceptionCode) Then
            If IsExceptionToRemoveShiftComponentRule() Then
                IsEligibleForShiftComponent = True
            Else
                IsEligibleForShiftComponent = False
            End If
        End If
    End If

    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsEligibleForShiftComponent"
    
End Function

Public Function IsEligibleForJobUpgrade(ByVal v_strUnionCode As String) As Boolean
'BLUE021
    Dim l_strCaseGroup As String
    
    If v_strUnionCode = trUnion148 Then
        l_strCaseGroup = CaseGroupFor(ExceptionCode)
        If (l_strCaseGroup = trCaseGroupIllness Or _
            l_strCaseGroup = trCaseGroupSuspension Or _
            l_strCaseGroup = trCaseGroupInjured Or _
            ExceptionBankTypeFor(ExceptionCode) = trBankTypeVacation) _
            And Not ExcludeFromJobUpgradeRule(ExceptionCode) Then
            IsEligibleForJobUpgrade = False
        Else
            IsEligibleForJobUpgrade = True
        End If
    Else
        IsEligibleForJobUpgrade = True
    End If

End Function

Public Function IsEligibleForVacationRelief(ByVal v_strUnionCode As String) As Boolean
'BLUE021
    IsEligibleForVacationRelief = IsEligibleForJobUpgrade(v_strUnionCode) _
        And ExceptionCode <> "HR"
End Function

Public Function IsEligibleForTravelPay(ByVal v_strUnionCode As String) As Boolean
'RED044
    
    IsEligibleForTravelPay = v_strUnionCode = trUnion148

End Function

Public Function IsCoffeen() As Boolean
    
    On Error GoTo eh
    
    IsCoffeen = _
        Timecard.Employee.IsCoffeen(ChargeDate)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsCoffeen"
        
End Function


Public Function IsSelfTradedShift() As Boolean
    
    On Error GoTo eh
    
    IsSelfTradedShift = _
        Timecard.Schedule.IsSelfTradedShift( _
            ChargeDate, _
            StartTimeAsDate, _
            EndTimeAsDate)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsSelfTradedShift"
        
End Function

Public Function ScheduledBeforeTrade() As Boolean
' Answer true when a schedule shift is found for this
' charge date checking demand level and below.
' Schedule request should always answer a record,
' just check the first record since the request is
' for a specific charge date.
    Dim l_bsSchedule As busSchedule.cbsSchedule
    Dim l_rsSchedule As ADODB.Recordset
    Dim l_rsScheduleWOExc As ADODB.Recordset

    On Error GoTo eh
    
    Set l_bsSchedule = CreateObject("busSchedule.cbsSchedule")
    Set l_rsSchedule = _
        l_bsSchedule.GetAdjustedBaseSchedule( _
            EmployeeNumber, ChargeDate, ChargeDate, l_rsScheduleWOExc, False, trAdjustedBaseSched, trDemandLineBaseSchedLevel)
            
    If Not (l_rsSchedule.BOF And l_rsSchedule.EOF) Then
        l_rsSchedule.MoveFirst
        ScheduledBeforeTrade = l_rsSchedule!SHIFT_HOURS > 0
    Else
        ScheduledBeforeTrade = False
    End If
    Set l_rsScheduleWOExc = Nothing
    Exit Function
    
eh:
    Set l_rsScheduleWOExc = Nothing
    TrsRaiseError mstrCMODULENAME, "ScheduledBeforeTrade"

End Function

Public Function IsTradedWorkShift() As Boolean
    IsTradedWorkShift = _
        Timecard.Schedule.IsTradedWorkShift(ChargeDate, StartTimeAsDate, EndTimeAsDate)
End Function

Public Function IsScheduledSunday() As Boolean
    IsScheduledSunday = _
        Timecard.Schedule.IsScheduledFirstDayOff(ChargeDate)
End Function

Public Function IsScheduledSaturday() As Boolean
    IsScheduledSaturday = _
        Timecard.Schedule.IsScheduledSecondDayOff(ChargeDate)
End Function

Public Function IsOutsideRegularScheduledHours() As Boolean
    IsOutsideRegularScheduledHours = _
        Timecard.Schedule.IsAnyPeriodOutsideSchedule(ChargeDate, StartTimeAsDate, EndTimeAsDate)
End Function

Public Function HasScheduleChangePayment() As Boolean
'07/03/2002 -- JIMB
'This is dead code this method on cboTimecardEntry is not called
'The method it calls on cboSchedule is not called from any other place
    HasScheduleChangePayment = _
        Timecard.Schedule.HasScheduleChangePayment(ChargeDate, DateWorked, startTime, EndTime)
End Function

Public Function IsScheduleChangePaymentSplitRequired() As Boolean
'A split is required when a schedule change payment exists where
' the time periods intersect and the time periods are not equal
' or the time entry period is a subset of the schedule change payment.
    Dim l_varSCP As Variant
    Dim l_strSCPStart As String
    Dim l_strSCPEnd As String

'07/03/2002 -- JIMB
'This methdod is dead code
       
    On Error GoTo eh
    
    l_varSCP = Timecard.Schedule.FirstScheduleChangePayment(ChargeDate, DateWorked, startTime, EndTime)
    If IsEmpty(l_varSCP) Then
        IsScheduleChangePaymentSplitRequired = False
        Exit Function
    End If
    
    l_strSCPStart = l_varSCP(4)
    l_strSCPEnd = l_varSCP(5)
    
    IsScheduleChangePaymentSplitRequired = SplitRule(l_strSCPStart, l_strSCPEnd)
    
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsScheduleChangePaymentSplitRequired"
    
End Function

Public Function ScheduleChangePaymentScheduled() As String
' BLUE086
    
    On Error GoTo eh
    
    ScheduleChangePaymentScheduled = _
        Timecard.Schedule.ScheduleChangePaymentScheduled( _
            ChargeDate, DateWorked, startTime, EndTime)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ScheduleChangePaymentScheduled"

End Function


Public Function SplitComplex( _
        ByVal v_strSplitStart As String, _
        ByVal v_strSplitEnd As String, _
        ByRef r_colEntries As Collection, _
        ByRef r_NextTimeEntry As cboTimecardEntry _
        ) As cboTimecardEntry
'This functionality is based on the user entering 1 time entry.
'The system may break this up into multiple time entries.
'This function will split a time entry, which may have
' already been split by the system

'This operation will split a time entry into upto 3 time entries
' depending on how the split request intersects with the time entry.
' e.g.      time entry          split
'            start    end      start    end         No. splits
'  1         07:00   15:00     08:00    11:00       2
'  2         07:00   15:00     07:00    11:00       1
'  3         07:00   15:00     11:00    15:00       1
'Answer the time entry that is associated with the split request.
' e.g. in split request 8-11 above answer the resulting time entry
' that starts at 8 and ends at 11.  Answer Nothing when no split.

'JIMB
'Added code to simplify multiple splits on a time entry
'Given the 2nd example above, the function returns 7 - 11.
'The time from 11 - 15 may also need split.
'The parameter NextTimeEntry will pass back 11 - 15.
'The time from 7 - 11 is the new values of original of the 7 - 15 entry
'Given the 3rd example above, the function returns 11 - 15.
'The time from 7 - 11 is the new values of original of the 7 - 15 entry
'The parameter NextTimeEntry will pass back nothing
'Given the 1st example above, the function returns 8 - 11.
'The time from 11 - 15 may also need split.
'The parameter NextTimeEntry will pass back 11 - 15.
'The time from 7 - 8 is what is left of the original 7 - 15 entry.
'The time from 7-8 may also need split, but if you split
'chronologically it is not needed to be passed back


    Dim newEntry As cboTimecardEntry

    On Error GoTo eh
    
    Select Case SplitRule(v_strSplitStart, v_strSplitEnd)
    Case l_intSplitNone
        Set SplitComplex = Nothing
        Set r_NextTimeEntry = Nothing
    Case l_intSplitStart
        Set SplitComplex = Split(v_strSplitStart, r_colEntries)
        Set r_NextTimeEntry = Nothing
    Case l_intSplitEnd
        Set r_NextTimeEntry = Split(v_strSplitEnd, r_colEntries)
        Set SplitComplex = SafeRef(Me)
    Case l_intSplitBoth
        Set newEntry = Split(v_strSplitStart, r_colEntries)
        Set r_NextTimeEntry = newEntry.Split(v_strSplitEnd, r_colEntries)
        Set SplitComplex = newEntry
    End Select
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "SplitComplex"

End Function

Private Function SplitRule( _
        ByVal v_strSplitStart As String, _
        ByVal v_strSplitEnd As String) As Integer
    On Error GoTo eh
    
    If v_strSplitStart = startTime And _
        v_strSplitEnd = EndTime Then
        SplitRule = l_intSplitNone
        Exit Function
    End If
    
    If Not TimePeriodStringsIntersect( _
            v_strSplitStart, v_strSplitEnd, _
            startTime, EndTime) Then
        SplitRule = l_intSplitNone
        Exit Function
    End If
            
    If v_strSplitStart <= startTime Then
        If IsTimeBetween(v_strSplitEnd, startTime, EndTime, False) Then
            SplitRule = l_intSplitEnd
            Exit Function
        Else
            SplitRule = l_intSplitNone
            Exit Function
        End If
    End If
    
    If v_strSplitStart > startTime Then
        If v_strSplitEnd >= EndTime Then
            SplitRule = l_intSplitStart
            Exit Function
        ElseIf IsTimeBetween(v_strSplitEnd, startTime, EndTime, False) Then
            SplitRule = l_intSplitBoth
            Exit Function
        Else
            SplitRule = l_intSplitNone
        End If
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "SplitRule"
End Function

Private Function SplitSimple(ByVal v_strSplitTime As String) _
                    As cboTimecardEntry
'Assuming time entries are already split on midnight.

    Dim newEntry As cboTimecardEntry
    Dim lcolErrArgs As New Collection

    On Error GoTo eh

' Ensure splitTime is a string in 24 hour format, e.g. 08:00
    If Not IsTimeValid(v_strSplitTime) Then
        Err.Raise trErrbusService, , "Split time is not valid"
    End If

'Ensure splitTime is between start and end time, not inclusive
    If Not IsTimeBetween(v_strSplitTime, startTime, EndTime, False) Then
        lcolErrArgs.Add "SplitTime: " & v_strSplitTime
        lcolErrArgs.Add "StartTime: " & startTime
        lcolErrArgs.Add "EndTime: " & EndTime
        Err.Raise trErrbusService, , "Split time is not inclusive"
    End If

    Set newEntry = CreateObject("busTimecard.cboTimecardEntry")
    newEntry.SetNew

'the split entry should be next to the original entry
    Timecard.AddTimecardEntry newEntry

    With newEntry
        .EmployeeNumber = Me.EmployeeNumber
        .ChargeDate = Me.ChargeDate
        .DateWorked = Me.DateWorked
        .startTime = v_strSplitTime
        .EndTime = Me.EndTime
        .DepartmentNumber = Me.DepartmentNumber
        .ExceptionCode = Me.ExceptionCode
        .DistributionId = Me.DistributionId
        .DistributionType = Me.DistributionType
        .JobNbr = Me.JobNbr
        .JobTaskNbr = Me.JobTaskNbr
        .JobTaskStepNbr = Me.JobTaskStepNbr
        .EMPRVDatabaseID = Me.EMPRVDatabaseID
        .OverrideExceptionCode = Me.OverrideExceptionCode
        .DepartmentNumber = Me.DepartmentNumber
        .DistributionDepartmentNumber = Me.DistributionDepartmentNumber
        .WorkgroupNumber = Me.WorkgroupNumber
        .JobNumber = Me.JobNumber
        .ShiftComponentCode = Me.ShiftComponentCode
        .FamilyLeaveIndicator = Me.FamilyLeaveIndicator
        .ScheduleChangeExceptionCode = Me.ScheduleChangeExceptionCode
        .ComponentCode = Me.ComponentCode
        .ComponentAmount = Me.ComponentAmount
        .OvertimeReasonCode = Me.OvertimeReasonCode
        .PreArrangedOTIndicator = Me.PreArrangedOTIndicator
        .CaseNumber = Me.CaseNumber
        .ReliefCode = Me.ReliefCode
        .OvertimeOpportunityNumber = Me.OvertimeOpportunityNumber
        .OvertimeOpportunityDescription = Me.OvertimeOpportunityDescription
        .ApprovedBy = Me.ApprovedBy
        .ApprovedDate = Me.ApprovedDate
        .ApprovalIndicator = Me.ApprovalIndicator
        .ScheduleTimeIndicator = Me.ScheduleTimeIndicator
        .GroupTimestamp = Me.GroupTimestamp
        .ScheduleCode = Me.ScheduleCode
        .ScheduleSourceCode = Me.ScheduleSourceCode
    End With

    newEntry.Elapse

    Me.EndTime = v_strSplitTime
    Me.Elapse

    Set SplitSimple = newEntry

    Exit Function

eh:
    Select Case Err.Number
        Case trErrbusService
            TrsRaiseError mstrCMODULENAME, "SplitSimple", , , , lcolErrArgs, True
        Case Else
            TrsRaiseError mstrCMODULENAME, "SplitSimple"
    End Select

End Function


Public Sub Elapse()

On Error GoTo eh
    Hours = DateDiff("n", StartTimeAsDate, EndTimeAsDate) / 60
    Exit Sub
eh:
    TrsRaiseError mstrCMODULENAME, "Elapse"
End Sub


Private Function IncludesTime(aTime As String) As Boolean
    IncludesTime = IsTimeBetweenStartEnd(aTime)
End Function

Private Function IsTimeBetweenStartEnd(aTime As String) As Boolean
    IsTimeBetweenStartEnd = IsTimeBetween(aTime, startTime, EndTime, False)
End Function

Private Function HoursAsMinutes() As Integer

    On Error GoTo eh
    
    HoursAsMinutes = Hours * 60
    Exit Function
eh:
    TrsRaiseError mstrCMODULENAME, "HoursAsMinutes"
End Function

Public Function IsValid() As Boolean
' Check each attribute for validity, noting errors and warnings.
    Dim l_strReturnMessage As String
    Dim l_colErrDescArgs As Collection
    Dim l_strErrMessage As String
    
    On Error GoTo eh
    
    ResetError
    IsValid = True
    If Not IsHoursIncrementValid Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Hours increment is not valid"
        IsValid = False
    End If

'4/9/01 RSG untested, will include in next release
'    If Not IsHoursValid Then
'        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "LU hours exceeded scheduled lunch period."
'        IsValid = False
'    End If
'
    If Not IsValidVacationReliefCode(UnionCode, ChargeDate, ReliefCode) Then
        If Not IsVacationReliefCode(UnionCode, ChargeDate, ReliefCode) Then
            l_strErrMessage = "Relief (" & ReliefCode & ") code " & _
                "is not valid for employee's union (" & UnionCode & ")."
        Else
            l_strErrMessage = ""
        End If
        Set l_colErrDescArgs = New Collection
        l_colErrDescArgs.Add ReliefCode
        l_colErrDescArgs.Add ExceptionCode
        AppendWarning trWarDiscardReliefComp, l_strErrMessage, l_colErrDescArgs
        Set l_colErrDescArgs = Nothing
        ReliefCode = ""
'        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Relief component code is not valid"
'        IsValid = False
    End If
    
    If Not IsValidComponentCode(UnionCode, ChargeDate, ComponentCode) Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Component code is not valid"
        IsValid = False
    End If
    
    If Not IsValidShiftComponentCode(UnionCode, ChargeDate, ShiftComponentCode) Then
        AppendEntryError mlngErrorNumber, mstrCBUSINESSERR, mstrErrorDescription
        IsValid = False
    Else
        If IsBusinessWarning(mlngErrorNumber) Then
            AppendEntryError mlngErrorNumber, mstrCBUSINESSWAR, mstrErrorDescription
        End If
    End If
    
    If Not IsValidJobNumber() Then
        AppendEntryError mlngErrorNumber, mstrCBUSINESSERR, mstrErrorDescription
        IsValid = False
    End If
    
    If Not IsValidScheduleChangeExceptionCode Then
        AppendEntryError mlngErrorNumber, mstrCBUSINESSERR, mstrErrorDescription
        IsValid = False
    End If
    
    If Not IsValidTravelPayHours Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Travel pay hours are not valid"
        IsValid = False
    End If
    
    If Not IsValidApprovalIndicator Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Approval indicator is not valid"
        IsValid = False
    End If
    
    If Not IsValidApprovedBy Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Approved by is not valid"
        IsValid = False
    End If
        
    If Not IsValidApprovedDate Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Approved date is not valid"
        IsValid = False
    End If
    
    If Not IsValidExceptionCode Then
'        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Exception Code is not valid"
        IsValid = False
    End If
    
    If Not IsValidPremiumTime Then
        IsValid = False
    End If
    
    If Not IsValidOverrideExceptionCode Then
        AppendEntryError mlngErrorNumber, mstrCBUSINESSERR, mstrErrorDescription
        IsValid = False
    End If
    
    If Not IsValidNote Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "This code requires a note to be entered.  Please enter one."
        IsValid = False
    End If
    
    If Not IsValidOvertimeReasonCode Then
        If Trim(OvertimeReasonCode) <> "" Then
            AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Overtime Reason Code is not valid.  Please choose a valid one from the list."
        Else
            AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Overtime Reason Code is missing.  Please enter one. "
        End If
        IsValid = False
    End If
    
    If Not IsValidOvertimeOpportunityDescription Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "A justification for the COP hours is not provided.  Please enter a justification in the notes' section. "
        IsValid = False
    End If
       
    If Not IsValidPreArrangedOTIndicator Then
        IsValid = False
    End If
    
    If Not IsValidFamilyLeaveIndicator Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Family leave indicator is not valid"
        IsValid = False
    End If
    
    If Not IsValidDepartment(DepartmentNumber) Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Department number is not valid"
        IsValid = False
    End If
    
    If Not IsValidDistributionDepartment(DistributionDepartmentNumber, l_strErrMessage) Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, l_strErrMessage ' "Distribution department number is not valid"
        IsValid = False
    End If
    
    If Not IsValidEntryDate Then
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Time Entry exists before hire date"
        IsValid = False
    End If

    ResetError
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsValid"

End Function

Private Function IsValidWorkGroup( _
            ByVal v_lngWorkGroupNum As Long, _
            ByVal v_datEffDate As Date) As Boolean
' WHITE088
    Dim lobjWGMgr As dbWorkgroup.cdWorkGroup
    
On Error GoTo eh
    
    Set lobjWGMgr = CreateObject("dbWorkGroup.cdWorkGroup")
    IsValidWorkGroup = lobjWGMgr.IsValidWorkGroup(v_lngWorkGroupNum, v_datEffDate)
    
    Set lobjWGMgr = Nothing
    Exit Function
    
eh:
    Set lobjWGMgr = Nothing
    TrsRaiseError mstrCMODULENAME, "IsValidWorkGroup"

End Function

Private Function IsValidDepartment(ByVal v_lngDistDeptNum As String) As Boolean
' WHITE085
    On Error GoTo eh
    
    IsValidDepartment = GetDepartmentManager.IsDepartmentValid(v_lngDistDeptNum)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsValidDepartment"

End Function

Private Function IsValidDistributionDepartment( _
                    ByVal v_lngDistDeptNum As String, _
                    ByRef r_strErrorMessage As String) As Boolean
' WHITE085

    Dim l_blnDepartmentExists As Boolean
    Dim l_strPayrollDept As String
    Dim l_strPayrollDeptCMPY As String
    Dim l_strDistrDeptCMPY As String
    Dim l_dbEffectiveDateRules As dbRules.cdEffectiveDateRules
    
    On Error GoTo eh
    
    r_strErrorMessage = ""
    
    l_blnDepartmentExists = IsValidDepartment(v_lngDistDeptNum)
    If l_blnDepartmentExists Then
        l_strPayrollDept = mobjTimecard.Employee.DepartmentNumber(ChargeDate)
        l_strPayrollDeptCMPY = GetDepartmentManager.GetCompanyCode(l_strPayrollDept)
        l_strDistrDeptCMPY = GetDepartmentManager.GetCompanyCode(v_lngDistDeptNum)
        
        If l_strPayrollDeptCMPY = l_strDistrDeptCMPY Then
            IsValidDistributionDepartment = True
        Else
            '//TODO: Not needed after AIC implementation
            Set l_dbEffectiveDateRules = CreateObject("dbRules.cdEffectiveDateRules")
            If l_dbEffectiveDateRules.DoesEffectiveDateRuleApply(dbRules.TRSRuleCode.trAICDeptCutOverEffDate, trCompanyCode, _
                                                                l_strPayrollDeptCMPY, ChargeDate) Then
                IsValidDistributionDepartment = True
            Else
                IsValidDistributionDepartment = False
            
                r_strErrorMessage = "Employee's payroll department's ({0}) company ({1}) does not match the distribution department's ({2}) company ({3})"
                r_strErrorMessage = StringFormat(r_strErrorMessage, l_strPayrollDept, l_strPayrollDeptCMPY, v_lngDistDeptNum, l_strDistrDeptCMPY)
            End If
            Set l_dbEffectiveDateRules = Nothing
        End If
    Else
        IsValidDistributionDepartment = False
        r_strErrorMessage = "Distribution department number is not valid"
    End If
    
    Exit Function
eh:
    Set l_dbEffectiveDateRules = Nothing
    TrsRaiseError mstrCMODULENAME, "IsValidDistributionDepartment"
End Function


Private Function IsValidExceptionCode() As Boolean
' SSO is a system exception that indicates overtime
' should be posted.
'BLUE094, RED038

'Release 07/25/2001 -- JFB  Made changes to allow overtime
'                            within schedule if it overlaps
'                            a vacation time entry
'Release 08/13/2001 -- JFB  Made changes to allow straight time
'                           and holiday time on the same day
'                           Change Holiday to be just like vacation
'                           Allows overtime to overlapp holiday time

    Dim l_strErrorMsg As String

On Error GoTo eh
    
    If IsTimeCollectionExceptionCode(ExceptionCode) Then
        'continue
    Else
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Invalid Exception Code. " & _
            "Exception code can not be used in Time Collection. " & _
            "If you need to use this (" & ExceptionCode & _
            ") please contact payroll."
        IsValidExceptionCode = False
        Exit Function
    End If
    
    Select Case ExceptionCode
        Case "SSO"
            AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Invalid Exception Code. Exception code SSO can not be used"
            IsValidExceptionCode = False
'        Case "HB", "VH"
'            IsValidExceptionCode = IsHoliday
        Case Else
            IsValidExceptionCode = _
                IsValidExceptionCodeFor( _
                    ExceptionCode, _
                    Employee.ExecContractIndicator(ChargeDate), _
                    Employee.EmployeeBenefitPackage)
    End Select
    If IsValidExceptionCode Then
        ' Continue
    Else
        If Employee.ExecContractIndicator(ChargeDate) = "E" Then
            l_strErrorMsg = "executive"
        Else
            l_strErrorMsg = "contract"
        End If
        AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Exception Code is not valid for " & l_strErrorMsg & " employee."
        IsValidExceptionCode = False
        Exit Function
    End If
    
    If IsValidExceptionCode And ExceptionCode = "NAH" Then
        If Not DoesRuleApply(190, trEmployeeNumber, CStr(EmployeeNumber)) Then
            l_strErrorMsg = "NAH Exception Code is not valid for this employee. This is an unpaid exception code, so its use is limited. If you are required to use this code have your TRIS administrator contact the TRIS team."
            AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, l_strErrorMsg
            IsValidExceptionCode = False
            Exit Function
        End If
    End If
              
'Is Exception Outside of Schedule Valid
    If Not IsRegularScheduledHours Or _
            IsScheduledHolidayOrVacation Then           'outside of schedule
        If IsExceptionValidOutsideOfSchedule(ExceptionCode) Then
            If Timecard.Employee.IsExecutive(ChargeDate) And _
                (IsExceptionOvertimeRelated(ExceptionCode) Or _
                    ExceptionCode = "ONP") Then
                IsValidExceptionCode = IsValidExecutiveOvertimeExceptionCode()
            Else
                IsValidExceptionCode = True
            End If
        Else
            AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Exception Code is not valid outside of schedule."
            IsValidExceptionCode = False
        End If
        Exit Function
    End If
            
 'Is Exception Inside of Schedule Valid
    If Timecard.Employee.IsLunchFloater(ChargeDate) Then
        'Allows Overtime anywhere within schedule
        IsValidExceptionCode = True
        Exit Function
    End If

    If Not IsExceptionValidInsideOfSchedule(ExceptionCode) Then
        ' Overtime with callout standby (CS) is valid inside.
        If IsOverrideThatAllowsOverlap(OverrideExceptionCode) And _
            IsExceptionOvertimeRelated(ExceptionCode) Then
            ' acceptable
            IsValidExceptionCode = True
        Else
            'Overtime is allowed inside of schedule
            ' if it overlapps Holiday or Vacation time
            If Timecard.DoesTimeEntryOverlapWithHolidayOrVacation(SafeRef(Me)) Then
                IsValidExceptionCode = True
            Else
                Select Case ExceptionCode
                    Case "ONP"
                        l_strErrorMsg = "Invalid Exception Code. Overtime Not Paid (ONP) is not valid inside of schedule."
                    Case "OST"
                        l_strErrorMsg = "Invalid Exception Code. Straight time outside of schedule (OST) is not valid inside of schedule."
                    Case Else
                        l_strErrorMsg = "Invalid Exception Code. Overtime is not valid inside of schedule."
                End Select
                AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, l_strErrorMsg
                IsValidExceptionCode = False
            End If
        End If
    End If
    
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidExceptionCode"

End Function


Private Function IsScheduledHolidayOrVacation() As Boolean

On Error GoTo eh

    IsScheduledHolidayOrVacation = _
        Me.Timecard.Schedule.IsScheduledHolidayOrVacation(Me.ChargeDate, Me.StartTimeAsDate, Me.EndTimeAsDate)

    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsScheduledHolidayOrVacation"

End Function


Private Function IsValidOverrideExceptionCode() As Boolean

On Error GoTo eh

    If OverrideExceptionCode <> "" Then
        IsValidOverrideExceptionCode = _
            IsOverrideExceptionCode(OverrideExceptionCode)
        If Not IsValidOverrideExceptionCode Then
            SetError trErrbusTimecardentryValidationError, "Override exception code is not valid"
'            AppendEntryError vbObjectError + 30000, mstrCBUSINESSERR, "Override exception code is not valid"
            Exit Function
        End If
    Else
        IsValidOverrideExceptionCode = True
    End If
    
    If OverrideExceptionCode = "JC" Then
        If JobNumber = Employee.JobNumber(ChargeDate) Then
            OverrideExceptionCode = ""
            AppendWarning vbObjectError + 30000, "Removed JC override code because, the job upgrade(JC) is not valid without a job change"
'            SetError trErrbusTimecardentryValidationError, "Job upgrade(JC) is not valid without a job change"
'            AppendEntryError vbObjectError + 30000, mstrCBUSINESSERR, "Job upgrade(JC) is not valid without a job change"
            IsValidOverrideExceptionCode = True
            Exit Function
        End If
    End If
    
    If DoesRuleApply(176, trExceptionCode, OverrideExceptionCode) Then
    'If OverrideExceptionCode = "TA" Then
        If Trim(OvertimeOpportunityDescription) = "" Then '//-----Notes column in TimeCollection-----//
            SetError trErrbusTimecardentryValidationError, "Notes/Comments are required when using Override exception code " & OverrideExceptionCode & "."
            IsValidOverrideExceptionCode = False
            Exit Function
        End If
    End If

    
'    AppendError ErrorSet, _
'    Timecard.Employee.EmployeeNumber, _
'    l_datChargeDate, _
'    l_datChargeDate, _
'    "", _
'    "", _
'    CStr(l_decReportedHours), _
'    "", _
'    "", _
'    trWarTimecardHoursLessSched, _
'    mstrCBUSINESSWAR, _
'    lstrErrDesc, _
'    m_strCMODULENAME

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidOverrideExceptionCode"

End Function


Private Function IsValidOvertimeReasonCode() As Boolean
'RED039, BLUE090
On Error GoTo eh

    If IsExceptionOvertimeRelated(ExceptionCode) And Not IsRefusalExceptionCode Then
        If Trim(OvertimeReasonCode) <> "" Then
            IsValidOvertimeReasonCode = IsOvertimeReasonCode(OvertimeReasonCode)
        Else
            IsValidOvertimeReasonCode = False
        End If
    Else
        If Trim(OvertimeReasonCode) = "" Then
            IsValidOvertimeReasonCode = True
        Else
            OvertimeReasonCode = ""
'            AppendWarning vbObjectError + 30000, "Removed overtime reason code because time is not overtime."
            IsValidOvertimeReasonCode = True
        End If
    End If

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidOvertimeReasonCode"

End Function

Private Function IsValidPreArrangedOTIndicator() As Boolean
On Error GoTo eh
    If Not IsNull(PreArrangedOTIndicator) Then
        If Not IsYorN(PreArrangedOTIndicator) Then
            AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Pre-Arranged OT indicator is not valid."
            IsValidPreArrangedOTIndicator = False
            Exit Function
        End If
    
        If PreArrangedOTIndicator = "Y" Then
            If IsExceptionOvertimeRelated(ExceptionCode) Then
                IsValidPreArrangedOTIndicator = True
            Else
                AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "Pre-Arranged OT indicator can only be used with overtime exceptions codes."
                IsValidPreArrangedOTIndicator = False
            End If
        Else
            IsValidPreArrangedOTIndicator = True
        End If
    Else
        IsValidPreArrangedOTIndicator = True
    End If
    
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidPreArrangedOTIndicator"
End Function

Private Function IsValidTravelPayHours() As Boolean
On Error GoTo eh

    If TravelPayHours <> 0 Then
        IsValidTravelPayHours = _
            IsEligibleForTravelPay(UnionCode) And _
                IsExceptionOvertimeRelated(ExceptionCode)
    Else
        IsValidTravelPayHours = True
    End If

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidTravelPayHours"

End Function

Private Function IsValidJobNumber() As Boolean
' BLUE021
' RED151
On Error GoTo eh

    If Not Timecard.GetJobManager.JobNumberIsValid(JobNumber) Then
        SetError trErrbusTimecardentryValidationError, "Job number is not valid"
        IsValidJobNumber = False
        Exit Function
    End If
' 4/2/01 RSG removed edit to avoid problem with new EFFDT.
'' Determine if job is effective
'    If Not GetJobManager.JobNumberIsValid(JobNumber, ChargeDate) Then
'        SetError trErrbusTimecardentryValidationError, "Job number is not in effect"
'        IsValidJobNumber = False
'        Exit Function
'    End If
    
    If JobNumber <> Employee.JobNumber(ChargeDate) Then
        If Not IsEligibleForJobUpgrade(Employee.UnionCode(ChargeDate)) Then
        SetError trErrbusTimecardentryValidationError, "Not eligible for job upgrade"
            IsValidJobNumber = False
            Exit Function
        ElseIf Not IsJobUpgradeValid(JobNumber) Then
            SetError trErrbusTimecardentryValidationError, "Job upgrade is not valid for employees union"
            IsValidJobNumber = False
            Exit Function
        End If
    End If
    IsValidJobNumber = True
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidJobNumber"
    
End Function

Private Function IsJobUpgradeValid(ByVal v_strJobNum As String) As Boolean
' RED151
' v_strJobNum is the job number changed from the base job.
'
    Dim l_strUnionCd As String
    
On Error GoTo eh
    l_strUnionCd = GetJobManager.JobUnionCode(v_strJobNum)
    If Employee.IsExecutive(ChargeDate) Then
        If Employee.UnionCode(ChargeDate) = l_strUnionCd Then
            IsJobUpgradeValid = True
        Else
            IsJobUpgradeValid = False
        End If
    Else
        If ExecutiveContractIndicatorForUnion(l_strUnionCd) = _
                Employee.ExecContractIndicator(ChargeDate) Then
            IsJobUpgradeValid = True
        Else
            IsJobUpgradeValid = False
        End If
    End If
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsJobUpgradeValid"
    
End Function

Private Function IsValidScheduleChangeExceptionCode() As Boolean

On Error GoTo eh
    If ScheduleChangeExceptionCode <> "" Then
        If IsValidScheduleChangeCode( _
                ScheduleChangeExceptionCode, _
                Employee.ExecContractIndicator) Then
            IsValidScheduleChangeExceptionCode = True
        Else
            SetError trErrbusTimecardentryValidationError, "Schedule change code is not valid"
'            AppendEntryError vbObjectError + 30000, mstrCBUSINESSERR, "Schedule change code is not valid"
            IsValidScheduleChangeExceptionCode = False
        End If
        If IsExceptionEligibleForScheduleChange Then
            IsValidScheduleChangeExceptionCode = True
        Else
            AppendWarning trWarBaseMessage, "Schedule change code was removed because it is not valid with exception code: " & ExceptionCode
            ScheduleChangeExceptionCode = ""
            ScheduleChangeEarningsCode = ""
            DetermineEarningsCode
'            AppendEntryError vbObjectError + 30000, mstrCBUSINESSERR, "Schedule change code is not valid with exception code: " & ExceptionCode
            IsValidScheduleChangeExceptionCode = True
        End If
    Else
        IsValidScheduleChangeExceptionCode = True
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsValidScheduleChangeExceptionCode"

End Function

Private Function IsHoursIncrementValid() As Boolean
' Validate the hours are correct increment
' BLUE045, BLUE046
    Dim l_lngRemainder As Long
    
    On Error GoTo eh
    
    l_lngRemainder = HoursAsMinutes Mod Employee.PartialHourMinutes(ChargeDate)
    If l_lngRemainder = 0 Then
        IsHoursIncrementValid = True
    Else
        IsHoursIncrementValid = False
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsHoursIncrementValid"
    
End Function

'Private Function IsHoursValid() As Boolean
'    Dim lobjSchedShift As cboScheduledShift
''4/9/01 RSG untested, will include in next release
'On Error GoTo eh
'    If ExceptionCode = "LU" Then
'        Set lobjSchedShift = Timecard.Schedule.ShiftScheduledDuring( _
'                    ChargeDate, StartTimeAsDate, EndTimeAsDate)
'        If lobjSchedShift Is Nothing Then
'            IsHoursValid = False
'        Else
'            If HoursAsMinutes > _
'                lobjSchedShift.LunchPeriodAsMinutes Then
'                IsHoursValid = False
'            Else
'                IsHoursValid = True
'            End If
'        End If
'    Else
'        IsHoursValid = True
'    End If
'    Set lobjSchedShift = Nothing
'    Exit Function
'
'eh:
'    Set lobjSchedShift = Nothing
'    TrsRaiseError mstrCMODULENAME, "IsHoursValid"
'
'End Function

Private Function IsValidVacationReliefCode(ByVal v_strUnionCode As String, ByVal v_datEffDate As Date, ByVal v_strCompCode As String) As Boolean
'BLUE022, WHITE071
    Dim str As String
    
On Error GoTo eh
    
    If v_strCompCode = "" Then
        IsValidVacationReliefCode = True
    Else
        If IsEligibleForVacationRelief(v_strUnionCode) Then
            IsValidVacationReliefCode = IsVacationReliefCode(v_strUnionCode, v_datEffDate, v_strCompCode)
        Else
            IsValidVacationReliefCode = False
        End If
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsValidVacationReliefCode"

End Function

Private Function IsValidComponentCode( _
            ByVal v_strUnionCode As String, _
            ByVal v_datEffDate As Date, _
            ByVal v_strCompCode As String) As Boolean
'WHITE073, WHITE052
    Dim str As String

    On Error GoTo eh
    If v_strCompCode = "" Then
        IsValidComponentCode = True
    Else
        If IsEligibleForComponent Then
            IsValidComponentCode = _
                IsComponentCode(v_strUnionCode, v_datEffDate, v_strCompCode)
        Else
            IsValidComponentCode = False
        End If
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsValidComponentCode"

End Function

Private Function IsValidShiftComponentCode( _
        ByVal v_strUnionCode As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCode As String) As Boolean
        
'Changes to IsEligibleForShiftComponent will cause changes to IsValidShiftComponentCode

'WHITE078, WHITE052
'JimB 09/27/2001
'Change to display approriate warning and clear shift component

'JIMB 05/06/2003
'Added Logic for CILCO.
'Change code to drop down logic to avoid nasty nested if statements
    Dim str As String
    Dim l_blnCILCOLocal8Employee As Boolean
    Dim l_blnCILCOEmployee As Boolean
    Dim l_blnCallawaySecurity As Boolean
    
On Error GoTo eh

    If v_strCompCode = "" Then
        IsValidShiftComponentCode = True
        Exit Function
    End If
    
    If ExceptionCode = "ONP" Then
        ShiftComponentCode = ""
        ShiftComponentAmount = 0
        IsValidShiftComponentCode = True
        Exit Function
    End If
        
    If IsEligibleForShiftComponent Then
        'Is the shift component valid for the union in the ttr_comp_amount table
        If IsShiftCompCode(v_strUnionCode, v_datEffDate, v_strCompCode) Then
            IsValidShiftComponentCode = True
        Else
            AppendWarning trWarShiftCompRemoved, "Shift component code (" & v_strCompCode & ") is not valid for union (" & v_strUnionCode & "). Shift component was removed"
            ShiftComponentCode = ""
            ShiftComponentAmount = 0
            IsValidShiftComponentCode = True
        End If
        Exit Function
    End If

'AmerenCIPS Contract and AmerenCILCO local 8 employees are the only employees
'that get a shift component removed for certain exceptions
'The exception is ineligible for a shift component
'JIMB 07/11/2003
'AmerenCILCO employees, except local 8, only get a shift component for time worked.
'Local 8 gets it for all time (Vacation being one of them) reported, except
'certain exceptions (such as Illness)
    
    l_blnCILCOEmployee = Timecard.Employee.IsAmerenCILCO(ChargeDate)
    l_blnCILCOLocal8Employee = Timecard.Employee.IsAmerenCILCOLocal8Employee(ChargeDate)
    If l_blnCILCOEmployee Then
        If l_blnCILCOLocal8Employee Then
            If Not IsExceptionIsEligibleForShiftComponent Then
                AppendWarning trWarShiftCompRemoved, "Exception (" & ExceptionCode & ") is not eligible for shift component (" & v_strCompCode & "). Shift component was removed"
                ShiftComponentCode = ""
                ShiftComponentAmount = 0
            End If
            IsValidShiftComponentCode = True
            Exit Function
        End If
        AppendWarning trWarShiftCompRemoved, "The non working Exception (" & ExceptionCode & ") is not eligible for shift component (" & v_strCompCode & "). Shift component was removed"
        ShiftComponentCode = ""
        ShiftComponentAmount = 0
        IsValidShiftComponentCode = True
        Exit Function
    End If

    l_blnCallawaySecurity = Timecard.Employee.IsCallawaySecurity(ChargeDate)
    If l_blnCallawaySecurity Then
        If Not CallawaySecurityShiftComponentEligible(OverrideExceptionCode) Then
            AppendWarning trWarShiftCompRemoved, "Shift component code (" & v_strCompCode & ") is not valid for override code (" & OverrideExceptionCode & "). Shift component was removed"
            ShiftComponentCode = ""
            ShiftComponentAmount = 0
            IsValidShiftComponentCode = True
            Exit Function
        End If
    End If
    
'Must be a AmerenCIPS contract employee with non work time

    AppendWarning trWarShiftCompRemoved, "Non-work time exception (" & ExceptionCode & ") is not eligible for shift component (" & v_strCompCode & "). Shift component was removed"
    ShiftComponentCode = ""
    ShiftComponentAmount = 0
    IsValidShiftComponentCode = True

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidShiftComponentCode"

End Function

'TODO: obsolete IsValidCaseNumber
'Private Function IsValidCaseNumber() As Boolean
'On Error GoTo eh
'
'    IsValidCaseNumber = True
'
'    Exit Function
'
'eh:
'    TrsRaiseError mstrCMODULENAME, "IsValidCaseNumber"
'
'End Function

'TODO: obsolete IsCaseNumberRequired
'Private Function IsCaseNumberRequired() As Boolean
'
'    On Error GoTo eh
'
'    IsCaseNumberRequired = False
'
'    Exit Function
'
'eh:
'    TrsRaiseError mstrCMODULENAME, "IsCaseNumberRequired"
'
'End Function

Private Function IsValidFamilyLeaveIndicator() As Boolean
' WHITE067
On Error GoTo eh
    If Not IsYorN(FamilyLeaveIndicator) Then
        IsValidFamilyLeaveIndicator = False
        Exit Function
    End If

    If FamilyLeaveIndicator = "Y" Then
        IsValidFamilyLeaveIndicator = IsEligibleForFamilyLeave(ExceptionCode)
    Else
        IsValidFamilyLeaveIndicator = True
    End If
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidFamilyLeaveIndicator"

End Function

Private Function IsValidApprovalIndicator() As Boolean

    On Error GoTo eh

    Select Case Trim(Me.ApprovalIndicator)
        Case "Y", "N"
            IsValidApprovalIndicator = True
        Case Else
            IsValidApprovalIndicator = False
    End Select
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidApprovalIndicator"

End Function

Private Function IsValidApprovedBy() As Boolean

    On Error GoTo eh

    If Trim(Me.ApprovalIndicator) = "Y" Then
       If Me.ApprovedBy > 0 Then
          IsValidApprovedBy = True
       End If
    ElseIf Trim(Me.ApprovalIndicator) = "N" Then
        If Me.ApprovedBy = 0 Then
            IsValidApprovedBy = True
        End If
    End If
    
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidApprovedBy"

End Function

Private Function IsValidApprovedDate() As Boolean

    On Error GoTo eh

    IsValidApprovedDate = IsDate(Me.ApprovedDate)
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidApprovedDate"

End Function


Private Function IsJobEligibleForPPMComponent( _
            ) As Boolean
           
 On Error GoTo eh
 
    Select Case JobNumber
        Case "OA0746", "OA0747", "OA0748", "OA0749", _
             "OA0750", "OA0751", "OA0752", "OA0753", _
             "OA1702", "OA1703", "OA1704", _
             "OA1705", "OA1716", "OA1723", _
             "OA1724", "OA1726", "OA1728", _
             "OA1730", "OA1732", "OA1734", _
             "OA1735", "OA1736", "OA1738", _
             "OA1875", "OA1884", "OA1885", _
             "OA2707", "OA2771", "OA2875", _
             "OA3100", "OA3101", "OA3103", _
             "OA3235", "OA3250", "OA3251", _
             "OA3252", "OA3253", "OA3255", _
             "OA3256", "OA3257", "OA3258", _
             "OA3259", "OA3260", "OA3261", _
             "OA3262", "OA3263", "0A3264", _
             "OA3265", "OA3320", "OA3350", _
             "OA3351", "OA3352", "OA3353", _
             "OA3354", "OA3355", "OA3356", _
             "OA3357", "OA3358", "OA3359", _
             "OA3360", "OA3361", "OA3362", _
             "OA3947", "OA3950", "OA3953"
            IsJobEligibleForPPMComponent = True
        Case Else
            IsJobEligibleForPPMComponent = False
    End Select
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsJobEligibleForPPMComponent"

End Function

Private Function FormatEntryBRV() As String

    FormatEntryBRV = _
        FormatBRV(TrType_TC, EmployeeNumber, ChargeDate, DateWorked, startTime, EndTime, Hours, ExceptionCode, OverrideExceptionCode)

End Function

'PTS 345 check for Holiday on Rest
Private Function IsHolidayOnRest( _
            ByVal v_strExcCd As String) As Boolean

On Error GoTo eh
    If Trim(v_strExcCd) = "HR" Then
        IsHolidayOnRest = True
    End If
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsHolidayOnRest"

End Function

Public Function DetermineStraightTimeExceptionCode() As String

    'To avoid a big nested if statement drop down logic was applied
    'Only 1 exception code applies.

    If ScheduledSaturdayPremiumApplies Then
        DetermineStraightTimeExceptionCode = "SAS"
        Exit Function
        'No need to continue found exception code
    End If

    If ScheduledSundayPremiumApplies Then
        'continue
    Else
        DetermineStraightTimeExceptionCode = "ST"
        Exit Function
    End If
    
    If IsCallowayMaintenanceDepartment(DistributionDepartmentNumber) Then
        DetermineStraightTimeExceptionCode = "SSC"
        Exit Function
    End If
    
    If mobjTimecard.Employee.IsAmerenCILCO(ChargeDate) Or mobjTimecard.Employee.IsAmerenCILCODispatch(ChargeDate) Then
        DetermineStraightTimeExceptionCode = "SU1"
        Exit Function
    End If
    
    If mobjTimecard.Employee.IsAmerenIPBenefitParticipant(ChargeDate) Then
        DetermineStraightTimeExceptionCode = "SU3"
        Exit Function
    End If

    If UnionCode = trUnion148 And IsSelfTradedShift Then
        If Not ScheduledBeforeTrade Then
            DetermineStraightTimeExceptionCode = "ST"
        Else
            DetermineStraightTimeExceptionCode = "SUS"
        End If
    Else
        DetermineStraightTimeExceptionCode = "SUS"
    End If

End Function
Private Function IsEligibleCIPSJobForSundayPremium( _
            ) As Boolean
           
    Dim l_strJobNumber As String
     
    
 On Error GoTo eh
 
    l_strJobNumber = Employee.JobNumber(ChargeDate)
    Select Case l_strJobNumber
        Case "YA6646", "AG6646"
            'Chemists at Springfield office are not eligible for SUS
            ' Also they do not work Sundays either
            If Employee.DepartmentNumber(ChargeDate) = "534" Then
                IsEligibleCIPSJobForSundayPremium = False
            Else
                IsEligibleCIPSJobForSundayPremium = True
            End If
        Case Else
            IsEligibleCIPSJobForSundayPremium = Employee.IsJobEligibleForSundayPremium(ChargeDate, JobNumber)
    End Select
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsEligibleCIPSJobForSundayPremium"

End Function

Private Function IsExceptionToPremiumPayRules()

' All of Ameren uses charge date in determining if SUS or SAS time should be applied.
' There is one exception, and those employees are at Callaway.
' Those employees use date worked in determining if SUS or SAS time should be applied.
'    Those employees are
'            JA (1455 -- IBEW - OFFICE WORKERS )
'        in the departments
'           154   ADMN -SVCS & RECDS
'           156   MATERIALS MNGMNT

    Dim l_bln As Boolean
    Dim l_strDept
    
    On Error GoTo eh
    
    l_strDept = Employee.DepartmentNumber(ChargeDate)
    
'    If l_strDept = "154" Or l_strDept = "156" Then
'    Else
'        IsExceptionToPremiumPayRules = False
'        Exit Function
'    End If

    Select Case l_strDept
        Case "154", "156", "007"
            If Trim(Employee.UnionCode(ChargeDate)) = "JA" Or Trim(Employee.UnionCode(ChargeDate)) = "JB" Then
                IsExceptionToPremiumPayRules = True
            Else
                IsExceptionToPremiumPayRules = False
            End If
        Case Else
            IsExceptionToPremiumPayRules = False
    End Select

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsExceptionToPremiumPayRules"
 
End Function

Private Function IsValidPremiumTime() As Boolean

    On Error GoTo eh

    IsValidPremiumTime = True
    Select Case ExceptionCode
        Case "SU3"
            
            If Not IsScheduledSunday Then
                AppendEntryError trWarTimecardValidation, mstrCBUSINESSWAR, "Please verify SU3 Premium time.  The charge date should be a Sunday."
                IsValidPremiumTime = True
            End If
                
            If Not mobjTimecard.Employee.IsAmerenIPBenefitParticipant(ChargeDate) Then
                AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "SU3 Premium time not allowed, must be an IP employee."
                IsValidPremiumTime = False
            End If
            
        Case "SU1"
            
            If Not IsScheduledSunday Then
                AppendEntryError trWarTimecardValidation, mstrCBUSINESSWAR, "Please verify SU1 Premium time.  The charge date should be a Sunday."
                IsValidPremiumTime = True
            End If
                
            If Not (mobjTimecard.Employee.IsAmerenCILCO(ChargeDate) Or mobjTimecard.Employee.IsAmerenCILCODispatch(ChargeDate)) Then
                AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "SU1 Premium time not allowed, must be a CILCO employee."
                IsValidPremiumTime = False
            End If
            
        Case "SU2"
            
            If Not IsScheduledSunday Then
                AppendEntryError trWarTimecardValidation, mstrCBUSINESSWAR, "Please verify SU2 Premium time.  The charge date should be a Sunday."
                IsValidPremiumTime = True
            End If
                
            If Not mobjTimecard.Employee.IsAmerenCILCO(ChargeDate) Then
                AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "SU2 Premium time not allowed, must be a CILCO employee."
                IsValidPremiumTime = False
            End If
    
        Case "SUS"
 'Allows SUS when the charge date is Sunday,
 ' unless the employee has a certain CIPS job
 'Some employees can get SUS time if the
 ' date worked is Sunday and the charge date is a
 ' Saturday or Monday
            If IsEligibleCIPSJobForSundayPremium Then
            Else
                'CIPS Job can not get SUS time
                AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSWAR, "SUS Premium time not allowed, so the system removed it.  CIPS job can not get premium time"
                ExceptionCode = "ST"
                IsValidPremiumTime = True
                Exit Function
            End If
            
            'SUS is valid if 1 of the following is true
            '1. The charge date is a Sunday
            '2. Date worked is a Sunday and the Employee gets SUS based on date worked
            If IsExceptionToPremiumPayRules Then
                If IsScheduledSunday Or Weekday(DateWorked) = 1 Then
                Else
'                    AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "SUS Premium time not allowed.  The charge date must be a Sunday or the date worked must be a Sunday."
'                    IsValidPremiumTime = False
                    AppendEntryError trWarTimecardValidation, mstrCBUSINESSWAR, "Please verify SUS Premium time.  The charge date should be a Sunday or the date worked should be a Sunday."
                    IsValidPremiumTime = True
                End If
            Else
                If Not IsScheduledSunday Then
'                    AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "SUS Premium time not allowed.  The charge date must be a Sunday."
'                    IsValidPremiumTime = False
                    If WorkgroupNumber = 16769 Then
                'Remove the warning for the Gas Dispatchers in Decatur
                    Else
                    AppendEntryError trWarTimecardValidation, mstrCBUSINESSWAR, "Please verify SUS Premium time.  The charge date should be a Sunday."
                    IsValidPremiumTime = True
                    End If
                End If
            End If
                
'Comment code below used the same code that defaulted
' SUS time
'            If Not ScheduledSundayPremiumApplies Then
'                AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "SUS Premium time not allowed.  Please check date and/or employee."
'                IsValidPremiumTime = False
'            End If
            
        Case "SAS"
'Only employees that entilted to Saturday Premium
' can get SAS time.
' It can not be forced in
'Hopefully we do not have a problem with this

'10/30/2001 JIMB
'It looks like we did so I added more code

            If Not IsEntitledToSaturdayPremium Then
                AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "SAS Premium time not allowed. The Employee is not entitled to SAS time."
                IsValidPremiumTime = False
                Exit Function
            End If
            
            'SAS is valid if 1 of the following is true
            '1. The charge date is a Saturday
            '2. Date worked is a Saturday and the Employee gets SAS based on date worked
            If IsExceptionToPremiumPayRules Then
                If IsScheduledSaturday Or Weekday(DateWorked) = 7 Then
                Else
'                    AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "SAS Premium time not allowed.  The charge date must be a Saturday or the date worked must be a Saturday."
'                    IsValidPremiumTime = False
                    AppendEntryError trWarTimecardValidation, mstrCBUSINESSWAR, "Please verify SAS Premium time.  The charge date should be a Saturday or the date worked should be a Saturday."
                    IsValidPremiumTime = True
                End If
            Else
                If Not IsScheduledSaturday Then
'                    AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "SAS Premium time not allowed.  The charge date must be a Saturday."
'                    IsValidPremiumTime = False
                    AppendEntryError trWarTimecardValidation, mstrCBUSINESSWAR, "Please verify SAS Premium time.  The charge date should be a Saturday."
                    IsValidPremiumTime = True
                End If
            End If
            
        Case "SSC"
'Callaway maintenace employees are the only employees
' that can get SSC time
' The charge date must be a Sunday
            If Not IsScheduledSunday Then
                AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "SSC Premium time not allowed.  The charge date must be a Sunday."
                IsValidPremiumTime = False
            Else
                If Not IsCallowayMaintenanceDepartment(DistributionDepartmentNumber) Then
                    AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, "SSC Premium time not allowed.  The employee must be in the Callaway Maintenance Department."
                    IsValidPremiumTime = False
                End If
            End If
                
    End Select
    
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidPremiumTime"

End Function


Private Function IsExceptionIsEligibleForShiftComponent() As Boolean

'RULE0020
    On Error GoTo eh

    IsExceptionIsEligibleForShiftComponent = Not _
        DoesRuleApply(20, trExceptionCode, ExceptionCode)
        
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsExceptionIsEligibleForShiftComponent"
        
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
    TrsRaiseError mstrCMODULENAME, "DoesRuleApply"
    
End Function

Private Function ExcludeFromJobUpgradeRule( _
    ByVal v_strExcCd As String _
    ) As Boolean
    
    On Error GoTo eh
    
    Select Case Trim(v_strExcCd)
        Case "A"
            ExcludeFromJobUpgradeRule = True
        Case Else
            ExcludeFromJobUpgradeRule = False
    End Select
    
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ExcludeFromJobUpgradeRule"
    
End Function


Private Function IsValidExecutiveOvertimeExceptionCode() As Boolean

    If ChargeDate < CDate("2004-02-16") Then
        IsValidExecutiveOvertimeExceptionCode = True
        Exit Function
    End If
    
    If IsRefusalExceptionCode Then
        IsValidExecutiveOvertimeExceptionCode = True
        Exit Function
    End If
    
    If Timecard.Employee.IsNonExemptExecutive(ChargeDate) Then
        If ExceptionCode = "ONP" Or DoesRuleApply(134, trExceptionCode, ExceptionCode) Then
            AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, _
                    "Overtime exception code is invalid for a non-exempt employee." & _
                    "  Non-exempt employees should report overtime to either " & _
                    "the OT1, OT2, or OT5 exception code."
            IsValidExecutiveOvertimeExceptionCode = False
        Else
            IsValidExecutiveOvertimeExceptionCode = True
        End If
    Else
        If ExceptionCode = "ONP" Then
            'Valid always
            IsValidExecutiveOvertimeExceptionCode = True
        Else
            If DoesRuleApply(134, trExceptionCode, ExceptionCode) Then
                'Not valid for employees that can not get paid for OT
                If Timecard.Employee.OTExclusn(ChargeDate) = "9999" Then
                    AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, _
                        "Overtime exception code is invalid for an Ameren Leadership Team employee." & _
                        "  Ameren Leadership Team employees should report overtime to the ONP exception code"
                        IsValidExecutiveOvertimeExceptionCode = False
                Else
                    IsValidExecutiveOvertimeExceptionCode = True
                End If
            Else
                'Only ONP and COP are valid for Exempt Executives
                If IsExceptionOvertimeRelated(ExceptionCode) Then
                    AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, _
                        "Overtime exception code is invalid for an exempt employee." & _
                        "  Exempt employees should report overtime to either ONP or COP exception code"
                    IsValidExecutiveOvertimeExceptionCode = False
                Else
                    IsValidExecutiveOvertimeExceptionCode = True
                End If
            End If
        End If
    End If
        
'        Select Case ExceptionCode
'            Case "ONP", "COP", "GI1"
'                AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, _
'                        "Overtime exception code is invalid for a non-exempt employee." & _
'                        "  Non-exempt employees should report overtime to either " & _
'                        "the OT1, OT2, or OT5 exception code."
'                IsValidExecutiveOvertimeExceptionCode = False
'            Case Else
'                IsValidExecutiveOvertimeExceptionCode = True
'        End Select
'    Else
'        Select Case ExceptionCode
'            Case "ONP"
'                'Valid Always
'                IsValidExecutiveOvertimeExceptionCode = True
'            Case "COP", "GI1"
'                'Not Valid for employees that can not get paid for OT
'                If Timecard.Employee.OTExclusn(ChargeDate) = "9999" Then
'                    AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, _
'                        "Overtime exception code is invalid for an Ameren Leadership Team employee." & _
'                        "  Ameren Leadership Team employees should report overtime to the ONP exception code"
'                        IsValidExecutiveOvertimeExceptionCode = False
'                Else
'                    IsValidExecutiveOvertimeExceptionCode = True
'                End If
'            Case Else
'                'Only ONP and COP are valid for Exempt Executives
'                If IsExceptionOvertimeRelated(ExceptionCode) Then
'                    AppendEntryError trErrbusTimecardentryValidationError, mstrCBUSINESSERR, _
'                        "Overtime exception code is invalid for an exempt employee." & _
'                        "  Exempt employees should report overtime to either ONP or COP exception code"
'                    IsValidExecutiveOvertimeExceptionCode = False
'                Else
'                    IsValidExecutiveOvertimeExceptionCode = True
'                End If
'        End Select
'    End If
    
        Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsValidExecutiveOvertimeExceptionCode"
    
            
    

End Function


Private Function DefaultExecutiveOvertimeRate( _
    ByVal v_strUnionCode As String) As String


On Error GoTo eh
    
    If ChargeDate < CDate("2004-02-16") Then
        Select Case v_strUnionCode
            Case "AH", "XA", "ZA", "AL", "AN", "AP", "AV"
                DefaultExecutiveOvertimeRate = "OTE"
            Case "AG", "YA"
                DefaultExecutiveOvertimeRate = "OT1"
            Case Else
                DefaultExecutiveOvertimeRate = "ONP"
        End Select
        If Employee.IsAmerenLeadership(ChargeDate) Then
                DefaultExecutiveOvertimeRate = "ONP"
        End If
    Else
        If Timecard.Employee.IsNonExemptExecutive(ChargeDate) Then
            DefaultExecutiveOvertimeRate = "OT2"
        Else
            DefaultExecutiveOvertimeRate = "ONP"
        End If
    End If
    
    Exit Function
        
eh:
    TrsRaiseError mstrCMODULENAME, "DefaultExecutiveOvertimeRate"
    
            
            

End Function

Private Function IsExceptionToRemoveShiftComponentRule()

    Dim l_strExcCd As String
    
    On Error GoTo eh
    
    If Me.Employee.LocationCode = 121 Then
        l_strExcCd = Trim(ExceptionCode)
        Select Case l_strExcCd
            Case "A"
                IsExceptionToRemoveShiftComponentRule = True
        End Select
    Else
        IsExceptionToRemoveShiftComponentRule = False
    End If
    
    
    

    Exit Function
        
eh:
    TrsRaiseError mstrCMODULENAME, "IsExceptionToRemoveShiftComponentRule"
    
End Function

Private Function IsRefusalExceptionCode() As Boolean

    Select Case ExceptionCode
        Case "REF", "RF1", "RF2", "RF3", "RF4", "RF5"
            IsRefusalExceptionCode = True
        Case Else
            IsRefusalExceptionCode = False
    End Select

    Exit Function
        
eh:
    TrsRaiseError mstrCMODULENAME, "IsRefusalExceptionCode"
    
End Function


Private Function IsValidOvertimeOpportunityDescription() As Boolean
'Callaway Managment employees must enter this with COP time
On Error GoTo eh

'if ot is not entered in Time Collection
'OT is comming for canvassed or non canvassed OT
    If Not OvertimeOpportunityNumber = trDefaultOTOppNum Then
        IsValidOvertimeOpportunityDescription = True
        Exit Function
    End If

    If Not Timecard.Employee.IsExecutive(ChargeDate) Then
        IsValidOvertimeOpportunityDescription = True
        Exit Function
    End If
    
    If Timecard.Employee.IsNonExemptExecutive(ChargeDate) Then
        IsValidOvertimeOpportunityDescription = True
        Exit Function
    End If
    
' All Exempt Management employees must enter notes for overtime
'''''    If Not DoesLocationRequireNotes Then
'''''        IsValidOvertimeOpportunityDescription = True
'''''        Exit Function
'''''    End If

'Do not show the missing notes BRV for ALT members
'They can only enter ONP
    If Timecard.Employee.OTExclusn(ChargeDate) = "9999" Then
        IsValidOvertimeOpportunityDescription = True
        Exit Function
    End If
    
    If IsExceptionOvertimeRelated(ExceptionCode) Then
        If Trim(OvertimeOpportunityDescription) <> "" Then
            IsValidOvertimeOpportunityDescription = True
        Else

            IsValidOvertimeOpportunityDescription = False
        End If
    Else
        IsValidOvertimeOpportunityDescription = True
    End If

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidOvertimeOpportunityDescription"

End Function
Private Function DoesLocationRequireNotes() As Boolean

    On Error GoTo eh

    DoesLocationRequireNotes = _
        DoesRuleApply(240, trLocationCode, Timecard.Employee.LocationCode)
        
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "DoesLocationRequireNotes"
        
End Function

Private Function IsValidNote() As Boolean
'All Employees must enter a note for IL time
On Error GoTo eh


    If Trim(OvertimeOpportunityDescription) <> "" Then
        IsValidNote = True
        Exit Function
    End If

    If IsNoteRequired Then
        IsValidNote = False
    Else
        IsValidNote = True
    End If

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidNote"

End Function

Private Function IsNoteRequired() As Boolean

    On Error GoTo eh

    IsNoteRequired = _
        DoesRuleApply(259, trExceptionCode, ExceptionCode) Or DoesRuleApply(259, trExceptionCode, OverrideExceptionCode)
        
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsNoteRequired"
        
End Function

Private Function IsValidEntryDate() As Boolean

On Error GoTo eh


    If Employee.EmploymentDate > ChargeDate Then
        IsValidEntryDate = False
    Else
'        If Employee.Status = "T" And Employee.StatusEffDate < ChargeDate Then
'            IsValidEntryDate = False
'        Else
            IsValidEntryDate = True
'        End If
    End If

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsValidEntryDate"

End Function

