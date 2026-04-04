VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cboOtherPayEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrCMODULENAME = "cboOtherPayEntry"
Private Const mstrCACCOUNTNOTEXISTS As String = ""
Private Const mstrCACCOUNTNOTAPPLY As String = "N/A"

'Use for validation
Private mlngErrorNumber As Long
Private mstrErrorDescription As String

Private mobjOtherPay As cboOtherPay

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
Private mstrDepartmentNumber As String
Private mvarHours As Variant
Private mvarAmount As Variant
Private mintLocationCode As Integer
Private mintStartLocationCode As Integer
Private mintEndLocationCode As Integer
Private mstrUnionCode As String
Private mstrExceptionCode As String
Private mstrEarningsCode As String
Private mstrShiftComponentCode As String
Private mvarShiftComponentAmount As Variant
Private mstrApprovalIndicator As String
Private mlngApprovedBy As Long
Private mdatApprovedDate As Date
Private mstrDistributionId As String
Private mstrDistributionType As String
Private mstrDistributionDepartmentNumber As String
Private mstrJobNumber As String
Private mlngLastModEmpNum As Long
Private mstrChemTesterIndicator As String
Private mstrHalfTripIndicator As String
Private mlngOvertimeOpportunityNumber As Long


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
    mstrDepartmentNumber = RTrim(rs!DEPT_NUM)
    mvarHours = rs!Hours
    mvarAmount = rs!AMOUNT
    mintLocationCode = rs!LOC_CD
    mintStartLocationCode = rs!START_LOC_CD
    mintEndLocationCode = rs!END_LOC_CD
    mstrUnionCode = RTrim(rs!UNION_CD)
    mstrExceptionCode = RTrim(rs!EXCEPTION_CD)
    mstrEarningsCode = RTrim(rs!EARNINGS_CD)
    mstrShiftComponentCode = RTrim(rs!COMP_SHIFT_CD)
    mvarShiftComponentAmount = rs!COMP_SHIFT_CD_AMT
    mstrApprovalIndicator = RTrim(rs!APPR_IND)
    mlngApprovedBy = rs!APPR_BY_EMP_NUM
    mdatApprovedDate = rs!APPR_DATE
    'DO NOT TRIM DISTRIBUTION_ID WITH LEADING SPACES
    mstrDistributionId = RTrim(rs!DISTRIBUTION_ID)
    mstrDistributionType = RTrim(rs!DISTRIBUTION_TYPE)
    mstrDistributionDepartmentNumber = RTrim(rs!DISTR_DEPT_NUM)
    mstrJobNumber = RTrim(rs!JOB_NUM)
    mlngLastModEmpNum = rs!LAST_MOD_EMP_NUM
    mstrChemTesterIndicator = RTrim(rs!CHEM_MTL_TSTR_IND)
    mstrHalfTripIndicator = RTrim(rs!HALF_TRIP_IND)
    mlngOvertimeOpportunityNumber = rs!OT_OPP_NUM
        
    Me.SetRestored
    Exit Function
     
eh:
    TrsRaiseError mstrCMODULENAME, "Restore"
End Function


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
        DayName = DayNameOfDate(wrkDate)
        SetModified
    End If
End Property

Public Property Get DayName() As String
    DayName = mstrDateWorkedDayName
End Property

Public Property Let DayName(wrkDateName As String)
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

Public Property Let startTime(stTime As String)
    If Not Trim(mstrStartTime) = Trim(stTime) Then
        mstrStartTime = Trim(stTime)
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

Public Property Get IndexTime() As String
    IndexTime = mstrIndexTime
End Property

Public Property Let IndexTime(ndxTime As String)
    If Not Trim(mstrIndexTime) = Trim(ndxTime) Then
        mstrIndexTime = Trim(ndxTime)
        SetModified
    End If
End Property

Public Property Get DepartmentNumber() As String
    DepartmentNumber = mstrDepartmentNumber
End Property

Public Property Let DepartmentNumber(deptNum As String)
    If Not Trim(mstrDepartmentNumber) = Trim(deptNum) Then
        mstrDepartmentNumber = Trim(deptNum)
        SetModified
    End If
End Property

Public Property Get hrs() As Variant
'DECIMAL(3,1)
    hrs = mvarHours
End Property

Public Property Let hrs(vHrs As Variant)
'DECIMAL(3,1)
    If Not mvarHours = vHrs Then
        mvarHours = vHrs
        SetModified
    End If
End Property

Public Property Get Amt() As Variant
    Amt = mvarAmount
End Property

Public Property Let Amt(vAmnt As Variant)
    If Not mvarAmount = vAmnt Then
        mvarAmount = vAmnt
        SetModified
    End If
End Property

Public Property Get Location() As Integer
    Location = mintLocationCode
End Property

Public Property Let Location(ByVal v_intLocCd As Integer)
    If Not mintLocationCode = v_intLocCd Then
        mintLocationCode = v_intLocCd
        SetModified
    End If
End Property

Public Property Get StartLocation() As Integer
    StartLocation = mintStartLocationCode
End Property

Public Property Let StartLocation(ByVal v_intLocCd As Integer)
    If Not mintStartLocationCode = v_intLocCd Then
        mintStartLocationCode = v_intLocCd
        SetModified
    End If
End Property

Public Property Get EndLocation() As Integer
    EndLocation = mintEndLocationCode
End Property

Public Property Let EndLocation(ByVal v_intLocCd As Integer)
    If Not mintEndLocationCode = v_intLocCd Then
        mintEndLocationCode = v_intLocCd
        SetModified
    End If
End Property

Public Property Get Union() As String
    Union = mstrUnionCode
End Property

Public Property Let Union(ByVal v_strUnionCd As String)
    If Not Trim(mstrUnionCode) = Trim(v_strUnionCd) Then
        mstrUnionCode = Trim(v_strUnionCd)
        SetModified
    End If
End Property

Public Property Get ExceptionCode() As String
    ExceptionCode = mstrExceptionCode
End Property

Public Property Let ExceptionCode(excCd As String)
    If Not Trim(mstrExceptionCode) = Trim(excCd) Then
        mstrExceptionCode = Trim(excCd)
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

Public Property Get ApprovalInd() As String
    ApprovalInd = mstrApprovalIndicator
End Property

Public Property Let ApprovalInd(ByVal v_strApprInd As String)
    If Not Trim(UCase(mstrApprovalIndicator)) = Trim(UCase(v_strApprInd)) Then
        mstrApprovalIndicator = Trim(UCase(v_strApprInd))
        SetModified
    End If
End Property

Public Property Get ApproveByEmpNum() As Long
    ApproveByEmpNum = mlngApprovedBy
End Property

Public Property Let ApproveByEmpNum(ByVal v_lngApprovedBy As Long)
    If Not mlngApprovedBy = v_lngApprovedBy Then
        mlngApprovedBy = v_lngApprovedBy
        SetModified
    End If
End Property

Public Property Get ApprovalDate() As Date
    ApprovalDate = mdatApprovedDate
End Property

Public Property Let ApprovalDate(ByVal v_datApprovedDate As Date)
    If Not mdatApprovedDate = v_datApprovedDate Then
        mdatApprovedDate = v_datApprovedDate
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

Public Property Get DistributionDeptNum() As String
    DistributionDeptNum = mstrDistributionDepartmentNumber
End Property

Public Property Let DistributionDeptNum(distDeptNum As String)
    If Not Trim(mstrDistributionDepartmentNumber) = Trim(distDeptNum) Then
        mstrDistributionDepartmentNumber = Trim(distDeptNum)
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

Public Property Get LastModEmployeeNumber() As Long
    LastModEmployeeNumber = mlngLastModEmpNum
End Property

Public Property Let LastModEmployeeNumber(EmpNum As Long)
    If Not mlngLastModEmpNum = EmpNum Then
        mlngLastModEmpNum = EmpNum
        SetModified
    End If
End Property

Public Property Get ChemTesterInd() As String
    ChemTesterInd = mstrChemTesterIndicator
End Property

Public Property Let ChemTesterInd(ByVal v_strChemTestInd As String)
    If Not Trim(UCase(mstrChemTesterIndicator)) = Trim(UCase(v_strChemTestInd)) Then
        mstrChemTesterIndicator = Trim(UCase(v_strChemTestInd))
        SetModified
    End If
End Property

Public Property Get HalfTripInd() As String
    HalfTripInd = mstrHalfTripIndicator
End Property

Public Property Let HalfTripInd(ByVal v_strHalfTripInd As String)
    If Not Trim(UCase(mstrHalfTripIndicator)) = Trim(UCase(v_strHalfTripInd)) Then
        mstrHalfTripIndicator = Trim(UCase(v_strHalfTripInd))
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


Private Sub AppendEntryErrorOtherPay( _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDescription As String, _
            Optional ByVal v_strErrorSource As String)

    Dim l_varErrorSource As String
    Dim l_strSpaces As String
    
On Error GoTo AppendEntryErrorOtherPay_EH
    
    l_strSpaces = ""
    l_varErrorSource = IIf(v_strErrorSource = "", mstrCMODULENAME, v_strErrorSource)
        
    AppendError _
        ErrorSet, _
        EmployeeNumber, _
        ChargeDate, _
        DateWorked, _
        startTime, _
        EndTime, _
        hrs, _
        ExceptionCode, _
        l_strSpaces, _
        v_lngErrorNumber, _
        v_strErrorType, _
        v_strErrorDescription, _
        l_varErrorSource

Exit Sub
AppendEntryErrorOtherPay_EH:
    
    TrsRaiseError mstrCMODULENAME, "AppendEntryErrorOtherPay"

End Sub

Public Sub DetermineChargedWorkgroupNumber()

    Dim l_lngWgNum As Long
    Dim lcolErrArgs As New Collection

On Error GoTo DetermineChargedWorkgroupNumber_EH
    
    l_lngWgNum = OtherPay.Employee.GetPrimaryWorkgroupNumber(ChargeDate)
    
    If Not l_lngWgNum = 0 Then
        WorkgroupNumber = l_lngWgNum
        DetermineLocationCode
    Else
        lcolErrArgs.Add EmployeeNumber
        lcolErrArgs.Add ChargeDate
        Err.Raise trErrbusEmpWorkgroupNotFound
    End If
    Exit Sub
    
DetermineChargedWorkgroupNumber_EH:
    Select Case Err.Number
        Case trErrbusEmpWorkgroupNotFound
            Err.Source = FormatEntryBRV
            TrsRaiseError mstrCMODULENAME, "DetermineChargedWorkgroupNumber", Err.Number, , lcolErrArgs
        Case Else
            TrsRaiseError mstrCMODULENAME, "DetermineChargedWorkgroupNumber"
    End Select

End Sub

Public Sub UpdateChargedWorkgroupNumber(ByVal v_lngWgNum As Long)
    Dim l_lngWgNum As Long
    Dim l_blnDefDistDeptNum As Boolean

On Error GoTo eh
    
    If v_lngWgNum = WorkgroupNumber Then
        Exit Sub
    End If

    l_lngWgNum = v_lngWgNum
    If l_lngWgNum = 0 Then
        l_lngWgNum = OtherPay.Employee.GetPrimaryWorkgroupNumber(ChargeDate)
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
    Dim l_strJobNum As String

On Error GoTo eh
    
    If Trim(v_strJobNum) = Trim(JobNumber) Then
        Exit Sub
    End If

    l_strJobNum = Trim(v_strJobNum)
    If l_strJobNum = "" Then
        l_strJobNum = OtherPay.Employee.JobNumber(ChargeDate)
    End If
    
' ensure we have a valid job number before continue
    If l_strJobNum = "" Then
        Err.Raise trErrbusService, , "Job number is not provided on employee record"
    End If
   
    JobNumber = l_strJobNum
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "UpdateJobNumber"
    
End Sub


Public Sub UpdateDistDeptNumber(ByVal v_strDistDeptNum As String)
    Dim l_strDistDeptNum As String

On Error GoTo eh
    
    If Trim(v_strDistDeptNum) = Trim(DistributionDeptNum) Then
        Exit Sub
    End If

    l_strDistDeptNum = Trim(v_strDistDeptNum)
    If l_strDistDeptNum = "" Then
        l_strDistDeptNum = ReallyDetermineDistributionDepartmentNumber
    End If
    
    DistributionDeptNum = l_strDistDeptNum
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "UpdateJobNumber"
    
End Sub


Public Sub DetermineDistributionDepartmentNumber()
'
'Dependent on DetermineChargedWorkgroupNumber
'
On Error GoTo DetermineDistributionDepartmentNumber_EH
    
    DistributionDeptNum = ReallyDetermineDistributionDepartmentNumber

Exit Sub
DetermineDistributionDepartmentNumber_EH:

    TrsRaiseError mstrCMODULENAME, "DetermineDistributionDepartmentNumber"

End Sub

Private Function ReallyDetermineDistributionDepartmentNumber() As String
'
'Dependent on DetermineChargedWorkgroupNumber
'
On Error GoTo DetermineDistributionDepartmentNumber_EH
    
    If Not IsDOJMWorkgroupNumber(WorkgroupNumber) And _
        DepartmentNumber <> "086" And _
        (Not GetBusDepartment.IsDOJMDepartmentNumber(DepartmentNumber)) Then
            ReallyDetermineDistributionDepartmentNumber = DepartmentNumber
    Else
            ReallyDetermineDistributionDepartmentNumber = WorkgroupDepartmentNumber
    End If
    
Exit Function
DetermineDistributionDepartmentNumber_EH:

    TrsRaiseError mstrCMODULENAME, "ReallyDetermineDistributionDepartmentNumber"

End Function


Public Sub DetermineJobNumber()
' Per Prem and Greg, Job Upgrade is not  an issue with Other Pay.
'  If no job number is sent by the client get the Employee's Job Number
On Error GoTo DetermineJobNumber_EH

        JobNumber = OtherPay.Employee.JobNumber(ChargeDate)

Exit Sub
DetermineJobNumber_EH:
    
    TrsRaiseError mstrCMODULENAME, "DetermineJobNumber"

End Sub

Public Sub DetermineLocationCode()
'
' Dependent on DetermineChargedWorkgroupNumber
'
' Set the location code when necessary.
    
    Dim l_intLocCd As Integer
    
On Error GoTo DetermineLocationCode_EH

    l_intLocCd = OtherPay.Employee.GetWorkgroupLocationCode(WorkgroupNumber)

' if workgroup is not a primary or secondary, just get the location
    If l_intLocCd = 0 Then
        l_intLocCd = GetLocationCode(WorkgroupNumber)
    End If
    
    If l_intLocCd <> Me.Location Then
        Location = l_intLocCd
    End If

Exit Sub
DetermineLocationCode_EH:

    TrsRaiseError mstrCMODULENAME, "DetermineLocationCode"

End Sub

Public Sub DeterminePayrollDepartmentNumber()

On Error GoTo DeterminePayrollDepartmentNumber_EH

    DepartmentNumber = Me.Employee.DepartmentNumber(ChargeDate)
    
Exit Sub
DeterminePayrollDepartmentNumber_EH:

    TrsRaiseError mstrCMODULENAME, "DeterminePayrollDepartmentNumber"

End Sub

Public Sub DetermineShiftComponentCode()
' BLUE063
'  Start Date and End Time are not valid arguments for the determination of Shift Component On
'  There will be no determination of the Shift Component, the client will pass or not pass this code
''    If IsEligibleForShiftComponent Then
''        ShiftComponentCode = _
''            OtherPay.Schedule.ShiftComponentOn(ChargeDate, StartTimeAsDate, EndTimeAsDate, IsCoffeen)
''
''        DetermineShiftComponentAmount
''    End If

    Dim l_objShift As cboScheduledShift
    Dim l_strShiftComponentOn As String
    Dim l_colShifts As Collection
    
    Set l_colShifts = OtherPay.Schedule.ScheduledShiftsOn(ChargeDate)
    If l_colShifts.Count = 0 Then
        l_strShiftComponentOn = ""
    Else
        Set l_objShift = l_colShifts.Item(1)
        l_strShiftComponentOn = l_objShift.ShiftComponentCode
    End If
    If Me.Employee.IsAmerenCIPS(ChargeDate) Then
        l_strShiftComponentOn = ""
    End If
    
    ShiftComponentCode = l_strShiftComponentOn
    
End Sub

Public Sub DetermineShiftComponentAmount()

    Dim l_varAmt As Variant
    Dim lcolErrArgs As New Collection
    
On Error GoTo DetermineShiftComponentAmount_EH
    
    If ShiftComponentCode = "" Then
        ShiftComponentAmount = CDec(0)
        Exit Sub
    End If

    l_varAmt = ShiftComponentAmountFor(Union, ChargeDate, ShiftComponentCode)
    
    If Not IsNull(l_varAmt) Then
        ShiftComponentAmount = CDec(l_varAmt)
    Else
        ShiftComponentAmount = 0
    End If
    
Exit Sub
DetermineShiftComponentAmount_EH:
    Select Case Err.Number
        Case trErrbusShiftCompAmountNotFound
            Err.Source = FormatEntryBRV
            TrsRaiseError mstrCMODULENAME, "DetermineShiftComponentAmount", Err.Number, , lcolErrArgs
        Case Else
            TrsRaiseError mstrCMODULENAME, "DetermineShiftComponentAmount"
    End Select
    
End Sub

Public Sub DetermineUnionCode()
'
' Dependent on DetermineDepartmentNumber
'
On Error GoTo DetermineUnionCode_EH

    Union = GetEffectiveUnionCode
    
Exit Sub
DetermineUnionCode_EH:

    TrsRaiseError mstrCMODULENAME, "DetermineUnionCode"

End Sub

Public Sub DetermineEarningsCode()

' Set the earnings code when necessary.
' BLUE096
    Dim l_strEarnCd As String
    
On Error GoTo DetermineEarningsCode_EH
    
    l_strEarnCd = EarningsCodeForExceptionCode( _
                            ExceptionCode, _
                            OtherPay.Employee.ExecContractIndicator(ChargeDate))
            
    If l_strEarnCd <> EarningsCode Then
        EarningsCode = l_strEarnCd
    End If

Exit Sub
DetermineEarningsCode_EH:

    TrsRaiseError mstrCMODULENAME, "DetermineEarningsCode"

End Sub

Public Function DetermineDayName(ByVal v_datDate As Date) As String
    
    DetermineDayName = DayNameOfDate(v_datDate)

End Function
'Public Function IsChemTester() As Boolean
''RED106 , 107
'On Error GoTo IsChemTester_EH
'
'    If Union = "OA" Then
'        Select Case JobNumber
'            Case "OA2871", "OA2878", "OA2879", "OA2881", "OA2882"
'                IsChemTester = True
'        End Select
'    End If
'
'Exit Function
'IsChemTester_EH:
'
'    TrsRaiseError mstrCMODULENAME, "IsChemTester"
'
'End Function

Public Function IsEligibleForShiftComponent() As Boolean
    
    Dim l_blnCILCOLocal8Employee As Boolean
On Error GoTo IsEligibleForShiftComponent_EH

    If (Me.Employee.IsExecutive(ChargeDate) _
        And Me.Employee.IsAmerenCIPS(ChargeDate)) _
        Or (Me.Employee.IsContract(ChargeDate) _
        And Me.Employee.IsAmerenCIPSBenefitParticipant(ChargeDate)) _
        And IsNonWorkTime(ExceptionCode) Then
            IsEligibleForShiftComponent = False
    Else
        l_blnCILCOLocal8Employee = Me.Employee.IsAmerenCILCOLocal8Employee(ChargeDate)
        If l_blnCILCOLocal8Employee And _
            Not IsExceptionIsEligibleForShiftComponent Then
            IsEligibleForShiftComponent = False
        Else
            IsEligibleForShiftComponent = True
        End If
    End If

Exit Function
IsEligibleForShiftComponent_EH:

    TrsRaiseError mstrCMODULENAME, "IsEligibleForShiftComponent"

End Function

Public Function IsNonWorkTime(ByVal v_strExceptionCode As String) As Boolean
    
    Select Case v_strExceptionCode
        Case "B", "E", "EF", "EC", "PB", "REL", "FH", "DF", _
            "PPT", "H", "J", "AD", "ED", "EDP", "AE", "JIE", "HR"
            IsNonWorkTime = True
    End Select
            
End Function

Private Function WorkgroupDepartmentNumber() As String
' Dependent on DetermineChargedWorkgroupNumber
    Dim dbWorkgroup As dbWorkgroup.cdWorkGroup

    On Error GoTo WorkGroupDepartmentNumber_EH
    Set dbWorkgroup = CreateObject("dbWorkgroup.cdWorkgroup")
    WorkgroupDepartmentNumber = dbWorkgroup.GetDepartmentNumber(WorkgroupNumber)
    Set dbWorkgroup = Nothing

Exit Function
WorkGroupDepartmentNumber_EH:
    Set dbWorkgroup = Nothing
    TrsRaiseError mstrCMODULENAME, "WorkGroupDepartmentNumber"
    
End Function

'All MTS bo_ objects must be a safe reference
Public Property Set OtherPay(objOtherPay As cboOtherPay)
    Set mobjOtherPay = objOtherPay
End Property

Public Property Get OtherPay() As cboOtherPay
    Set OtherPay = mobjOtherPay
End Property

Public Function Employee() As cboEmployee
    Set Employee = OtherPay.Employee
End Function

Public Function ErrorSet() As ADODB.Recordset
    Set ErrorSet = OtherPay.ErrorSet
End Function

Private Function GetDepartmentManager() As dbDepartment.cdDepartment

    Set GetDepartmentManager = OtherPay.GetBSTimecard.GetDepartmentManager
    
End Function

Private Function GetBusDepartment() As busDepartment.cbsDepartment

    Set GetBusDepartment = CreateObject("busDepartment.cbsDepartment")
    
End Function


Public Function IsValid() As Boolean

' Validate the Fields
On Error GoTo IsValid_EH

    ResetError
    IsValid = True
'  Exception Code    Other Pay is dependant on Exception Code
    If Not IsValidExceptionCode(ExceptionCode) Then
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Exception Code is not valid"
        IsValid = False
    Else
        ' Based on Exception Code are all components entered
        If Not IsValidEntry() Then
            IsValid = False
        End If
    End If
    
'  WorkGroup Number
    If Not IsValidWorkGroup(WorkgroupNumber, ChargeDate) Then
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Workgroup Number is not valid"
        IsValid = False
    End If

'  Department Number
    If Not IsValidDepartment(DepartmentNumber) Then
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Department Number is not valid"
        IsValid = False
    End If

'  Distribution Department Number
    If Not IsValidDepartment(DistributionDeptNum) Then
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Distribution Department Number is not valid"
        IsValid = False
    End If

'  Shift Component Code/Shift Component Amt
    If Trim(ShiftComponentCode) <> "" Then
        If Not IsValidShiftComponentCode(Union, ChargeDate, ShiftComponentCode) Then
            AppendEntryErrorOtherPay trWarShiftCompRemoved, mstrCBUSINESSWAR, "Shift component code (" & ShiftComponentCode & ") is not valid for union (" & Union & "). Shift component was removed"
            ShiftComponentCode = ""
            ShiftComponentAmount = 0
        End If
    End If

'  Earnings Code
    If Not IsValidEarningsCode() Then
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Earnings Code is not valid"
        IsValid = False
    End If

'  Approval Indicator
    If Not IsValidApprovalIndicator() Then
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Approval Indicator is not valid"
        IsValid = False
    End If

'  Approved By
    If Not IsValidApprovedBy() Then
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Approved By is not valid"
        IsValid = False
    End If

'  Approved Date
    If Not IsValidApprovedDate() Then
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Approved Date is not valid"
        IsValid = False
    End If

'  Day Name
    If Not IsValidDayName() Then
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Day Name is not valid"
        IsValid = False
    End If

' Job Number
' needs work  (Upgrades not considered)
    If Not IsValidJobNumber() Then
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Job Number is not valid"
        IsValid = False
    End If
    ResetError
    
    Exit Function

IsValid_EH:
    TrsRaiseError mstrCMODULENAME, "IsValid"

End Function

Private Function IsValidEntry() As Boolean
'  Based on the Exception code, is this a valid entry

    Dim l_blnValidationError As Boolean
    Dim l_strCompanyCode As String
    Dim l_strReturnMessage As String

On Error GoTo IsValidEntry_EH

    Select Case UCase(ExceptionCode)
        Case "TTA"
            ' Check Amount
            If Not IsValidAmount Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Invalid Amount for Other Pay"
                l_blnValidationError = True
                Exit Function
            End If
            
            ' Check Start/End Location
            If Not IsValidStartEndLocation Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Invalid Start and End Location for Other Pay"
                l_blnValidationError = True
                Exit Function
            End If
        
            ' Check Hours  (If not zero set to zero and warn if not already)
            If Not IsValidHoursEntry Then
                l_blnValidationError = True
                Exit Function
            End If
'Begin Prem PTS 2091
        Case "JSR", "MAT", "MPT", "MAN" _
            , "ULA", "SB3", "SB4", "JSS", "JSO" _
            , "BBF", "BPG", "DTR", "TRV"
'End Prem PTS 2091
            ' Check Amount
            If Not IsValidAmount Then
                l_blnValidationError = True
                Exit Function
            End If

            ' Check Start/End Location (Set to zero and warn if not already)
            If Not IsValidStartEndLocationEntry Then
                l_blnValidationError = True
                Exit Function
            End If
            ' Check Hours (Set to zero and warn if not already)
            If Not IsValidHoursEntry Then
                l_blnValidationError = True
                Exit Function
            End If
    
        Case "DM2", "DM4", "DM5", "MM2", "MM4", "MM5", "MM1"
            ' Check Hours
            If Not IsValidHours Then
                l_blnValidationError = True
                Exit Function
            End If
        
            ' Check Amount (Set to zero and warn if not already)
            If Not IsValidAmountEntry Then
                l_blnValidationError = True
                Exit Function
            End If
            
            ' Check Start/End Location (Set to zero and warn if not already)
            If Not IsValidStartEndLocationEntry Then
                l_blnValidationError = True
                Exit Function
            End If
            
        Case Else
            If Not IsValidAmount Then
                l_blnValidationError = True
                Exit Function
            End If
            
            If IsEmpty(Amt) Or Amt = 0 Then
                If IsEmpty(hrs) Or hrs = 0 Then
                    AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount or Hours must be entered for Exception Code."
                    l_blnValidationError = True
                    Exit Function
                Else
                    If Not IsValidHours Then
                        l_blnValidationError = True
                        Exit Function
                    End If
                End If
            Else
                If IsEmpty(hrs) Or hrs = 0 Then
                Else
                    AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount and Hours can not both be entered for Exception Code."
                    l_blnValidationError = True
                    Exit Function
                End If
            End If
        
    End Select
    
    IsValidEntry = Not l_blnValidationError
    Exit Function

IsValidEntry_EH:
    TrsRaiseError mstrCMODULENAME, "IsValidEntry"

End Function

Private Function IsValidAmount() As Boolean
' RED102 , 109
'  Amounts are required for certain Other Pay Exception Codes, there are valid ranges for these codes

On Error GoTo IsValidAmount_EH
    
    Dim l_strResponse As String
    Dim l_blnValidationError As Boolean
    
    '  Is Amount Valid
    Select Case UCase(ExceptionCode)
'Begin Prem PTS 2091
        Case "TRV"
            If Not (Amt >= -2000 And Amt <= 2000) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount is not in valid range for Exception Code."
                l_blnValidationError = True
            End If
            If (IsEmpty(Amt) Or Amt = 0) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Exception Code."
                l_blnValidationError = True
            End If
'End Prem PTS 2091

        Case "JSR"
            If Not (Amt >= -100 And Amt <= 100) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount is not in valid range for Exception Code."
                l_blnValidationError = True
            End If
            If IsEmpty(Amt) Or Amt = 0 Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Exception Code."
                l_blnValidationError = True
            End If
         Case "MPT"
            If Not (Amt >= -50 And Amt <= 50) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount is not in valid range for Exception Code."
                l_blnValidationError = True
            End If
            If IsEmpty(Amt) Or Amt = 0 Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Exception Code."
                l_blnValidationError = True
            End If
        Case "MAN", "MAT"
            If Not (Amt >= -75 And Amt <= 75) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount is not in valid range for Exception Code."
                l_blnValidationError = True
            End If
            If IsEmpty(Amt) Or Amt = 0 Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Exception Code."
                l_blnValidationError = True
            End If
        Case "ULA"
            If Not (Amt >= -8 And Amt <= 8) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount is not in valid range for Exception Code."
                l_blnValidationError = True
            End If
            If IsEmpty(Amt) Or Amt = 0 Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Exception Code."
                l_blnValidationError = True
            End If
        Case "TTA"
            ' I need more here - the amount should be valid based on TTA Rules
            ' 08/07/2000 Per Prem and Greg - only check if between -50 and 50, client will calculate
            If Not (Amt >= -50 And Amt <= 50) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount is not in valid range for Exception Code."
                l_blnValidationError = True
            End If
            If IsEmpty(Amt) Or Amt = 0 Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Exception Code."
                l_blnValidationError = True
            End If
            
        Case "SB3"
            If Not (Amt >= -50 And Amt <= 50) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount is not in valid range for Exception Code."
                l_blnValidationError = True
            End If
            If IsEmpty(Amt) Or Amt = 0 Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Exception Code."
                l_blnValidationError = True
            End If
            
        Case "SB4"
            If Not (Amt >= -75 And Amt <= 75) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount is not in valid range for Exception Code."
                l_blnValidationError = True
            End If
            If IsEmpty(Amt) Or Amt = 0 Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Exception Code."
                l_blnValidationError = True
            End If
            
        Case "JSS"
            If Not (Amt >= -72 And Amt <= 72) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount is not in valid range for Exception Code."
                l_blnValidationError = True
            End If
            If IsEmpty(Amt) Or Amt = 0 Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Exception Code."
                l_blnValidationError = True
            End If
        Case "JSO"
            If Not (Amt >= -150 And Amt <= 150) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount is not in valid range for Exception Code."
                l_blnValidationError = True
            End If
            If IsEmpty(Amt) Or Amt = 0 Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Exception Code."
                l_blnValidationError = True
            End If
        Case "BBF", "BPG"
            If Not (Amt >= -1000 And Amt <= 1000) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount is not in valid range for Exception Code."
                l_blnValidationError = True
            End If
            If IsEmpty(Amt) Or Amt = 0 Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Exception Code."
                l_blnValidationError = True
            End If
        Case "DTR"
            If Not (Amt >= -20 And Amt <= 20) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Diesel Truck Reimbursement can not be more than $20.  Create another entry if a higher reimbursement is needed."
                l_blnValidationError = True
            End If
            If IsEmpty(Amt) Or Amt = 0 Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount must be entered for Diesel Truck Reimbursement Exception Code."
                l_blnValidationError = True
            End If
        Case Else
            If Not (Amt >= -100 And Amt <= 100) Then
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Exception Code is outside the default amount range of -100 to +100.  Please contact payroll if you need to enter an amount outside that range."
                l_blnValidationError = True
            End If
    End Select

    IsValidAmount = Not l_blnValidationError
        
Exit Function
IsValidAmount_EH:
    
    TrsRaiseError mstrCMODULENAME, "IsValidAmount"

End Function

Private Function IsValidApprovalIndicator() As Boolean

On Error GoTo IsValidApprovalIndicator_EH

    Select Case Trim(UCase(ApprovalInd))
        Case "Y", "N"
            IsValidApprovalIndicator = True
        Case Else
            IsValidApprovalIndicator = False
    End Select

Exit Function
IsValidApprovalIndicator_EH:
    
    TrsRaiseError mstrCMODULENAME, "IsValidApprovalIndicator"

End Function
Private Function IsValidApprovedBy() As Boolean

    On Error GoTo IsValidApprovedBy_EH

    If Trim(ApprovalInd) = "Y" Then
       If Me.ApproveByEmpNum > 0 Then
          IsValidApprovedBy = True
       End If
    ElseIf Trim(ApprovalInd) = "N" Then
        If ApproveByEmpNum = 0 Then
            IsValidApprovedBy = True
        End If
    End If
    
Exit Function
IsValidApprovedBy_EH:

    TrsRaiseError mstrCMODULENAME, "IsValidApprovedBy"

End Function

Private Function IsValidApprovedDate() As Boolean

On Error GoTo IsValidApprovedDate_EH

    IsValidApprovedDate = IsDate(ApprovalDate)
    
Exit Function
IsValidApprovedDate_EH:

    TrsRaiseError mstrCMODULENAME, "IsValidApprovedDate"

End Function


Private Function IsValidDayName() As Boolean

On Error GoTo IsValidDayName_EH

    Select Case Trim(UCase(DayName))
        Case "SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"
            IsValidDayName = True
        Case Else
            IsValidDayName = False
    End Select

Exit Function
IsValidDayName_EH:
    TrsRaiseError mstrCMODULENAME, "IsValidDayName"

End Function

Private Function IsValidHours() As Boolean
'RED110
'RED103

    Dim l_blnValidationError As Boolean
    
On Error GoTo IsValidHours_EH
'  Hours are required for certain Exception Codes, they must be between 0 and 24

    '  Is Hours Valid
    If Not (hrs > 0 And hrs <= 24) Then
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Hours is not valid."
        l_blnValidationError = True
    End If
        
    IsValidHours = Not l_blnValidationError
    
Exit Function
IsValidHours_EH:

    TrsRaiseError mstrCMODULENAME, "IsValidHours"

End Function

Private Function IsValidEndTime() As Boolean

On Error GoTo IsValidEndTime_EH

    If startTime = EndTime Then
        IsValidEndTime = True
    Else
        IsValidEndTime = False
    End If

Exit Function
IsValidEndTime_EH:

        TrsRaiseError mstrCMODULENAME, "IsValidEndTime"

End Function


Private Function IsValidWorkGroup( _
                    ByVal v_lngWorkGroupNum As Long, _
                    ByVal v_datEffDate As Date _
                        ) As Boolean
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

Private Function IsValidUnionCode() As Boolean
    '  Per Greg B. Job Number is validated by the client and there are caching and rule issues to be solved
    
On Error GoTo IsValidUnionCode_EH

    IsValidUnionCode = True

Exit Function
IsValidUnionCode_EH:

    TrsRaiseError mstrCMODULENAME, "IsValidUnionCode"
    
End Function

Private Function IsValidJobNumber() As Boolean
    Dim lobjJobMgr As dbCodeTables.cdCodeTables

On Error GoTo IsValidJobNumber_EH

    If Trim(JobNumber) = "" Then
        IsValidJobNumber = True
        Exit Function
    End If
    
' Determine if job is effective
    Set lobjJobMgr = CreateObject("dbCodeTables.cdCodeTables")
' 4/2/01 RSG Removed edit for effective date problem.
'    IsValidJobNumber = _
'        lobjJobMgr.JobNumberIsValid(JobNumber, ChargeDate)
    IsValidJobNumber = _
        lobjJobMgr.JobNumberIsValid(JobNumber)
    
    Set lobjJobMgr = Nothing
    Exit Function

IsValidJobNumber_EH:
    Set lobjJobMgr = Nothing
    TrsRaiseError mstrCMODULENAME, "IsValidJobNumber"

End Function

Private Function IsValidDepartment(ByVal v_lngDistDeptNum As String) As Boolean
' WHITE085

    On Error GoTo IsValidDepartment_EH
    
    IsValidDepartment = GetDepartmentManager.IsDepartmentValid(v_lngDistDeptNum)
    Exit Function
    
IsValidDepartment_EH:

    TrsRaiseError mstrCMODULENAME, "IsValidDepartment"

End Function

Private Function IsValidExceptionCode(ByVal v_strExceptionCode As String) As Boolean
'BLUE094

On Error GoTo IsValidExceptionCode_EH
    
    If Trim(UCase(ExceptionCode = "SSO")) Then
        IsValidExceptionCode = False
    Else
        IsValidExceptionCode = IsOtherPayExceptionCode(v_strExceptionCode)
    End If

Exit Function
IsValidExceptionCode_EH:

    TrsRaiseError mstrCMODULENAME, "IsValidExceptionCode"

End Function

Private Function IsValidShiftComponentCode(ByVal v_strUnionCode As String, _
                        ByVal v_datEffDate As Date, _
                        ByVal v_strCompCode As String _
                        ) As Boolean
    'WHITE078
'  Validate only if passed by the client.
    Dim str As String
    
On Error GoTo IsValidShiftComponentCode_EH

    If v_strCompCode = "" Then
        IsValidShiftComponentCode = True
    Else
        IsValidShiftComponentCode = IsShiftCompCode(UCase(v_strUnionCode), v_datEffDate, UCase(v_strCompCode))
    End If
    
Exit Function
IsValidShiftComponentCode_EH:

    TrsRaiseError mstrCMODULENAME, "IsValidShiftComponentCode"

End Function

Private Function IsValidEarningsCode() As Boolean
'  work on
On Error GoTo IsValidEarningsCode_EH

    IsValidEarningsCode = True
    
Exit Function
IsValidEarningsCode_EH:

    TrsRaiseError mstrCMODULENAME, "IsValidEarningsCode"
    
End Function

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

Public Function IsCoffeen() As Boolean
    
On Error GoTo IsCoffeen_EH
    
    IsCoffeen = OtherPay.Employee.IsCoffeen(ChargeDate)

Exit Function
IsCoffeen_EH:
    
    TrsRaiseError mstrCMODULENAME, "IsCoffeen"
        
End Function

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

''***OP2***
Public Function IsChemTester() As String
'RED106 , 107
On Error GoTo IsChemTester_EH

    If ChemTesterInd = "Y" Then
        IsChemTester = "Y"
    Else
        IsChemTester = "N"
    End If

Exit Function
IsChemTester_EH:

    TrsRaiseError mstrCMODULENAME, "IsChemTester"

End Function

'***OP2***
Public Function DetermineTTAAmount() As Boolean

    Dim l_strIsChemTester As String
    Dim l_strIsHalfTrip As String
    Dim l_intStartLoc As Integer
    Dim l_intEndLoc As Integer
    Dim l_strDistId As String * 39
    Dim l_varAmt As Variant
    
    
On Error GoTo DetermineTTAAmount_EH
    
            l_intStartLoc = StartLocation
            l_intEndLoc = EndLocation
            l_strIsChemTester = IsChemTester
            l_strIsHalfTrip = IsHalfTrip
            
    If (l_intStartLoc > 0) And (l_intEndLoc > 0) Then
        ' Determine Amount and Distribution Id
            If DetermineDistIdandAmount( _
                l_strIsChemTester, _
                l_intStartLoc, l_intEndLoc, _
                l_strIsHalfTrip, ChargeDate, _
                l_strDistId, l_varAmt) Then
                                                    
                Amt = l_varAmt
' JimB This Module is no longer determining Distribution ID
' It is being done in busDistributionMgr
' The code change here is to simply comment out one line of code
' The Amount will still be determined as before
'                DistributionId = l_strDistId
                DistributionType = "A"
                
                 '  Half Trip
                If l_strIsHalfTrip = "Y" Then
                    Amt = Amt / 2
                End If
               
            Else
                ' Error, Amount and Distribution Id could not be found
                DetermineTTAAmount = False
                AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount was not able to be determined for Exception Code."
            End If
    Else
        ' Determine Distribution Id and determine Amount provided by client
        If l_strIsHalfTrip = "Y" Then
            Amt = Amt / 2
        End If

        ' Set Start and End Locations equal to zero
        StartLocation = 0
        EndLocation = 0
        
        ' Make sure Distribution Id is not null or spaces
        ' Make sure Distribution Type is A
        If DistributionType <> "A" Then
            ' Error Distribution Type must be A
            DetermineTTAAmount = False
            AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Distribution Type was invalid for Exception Code."
        End If
        
        If DistributionId = "" Then
            ' Distribution Id is required
            DetermineTTAAmount = False
            AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Distribution Id is required for Exception Code."
        End If
    
    End If
    
    DetermineTTAAmount = True
    
Exit Function
DetermineTTAAmount_EH:

    DetermineTTAAmount = False
    TrsRaiseError mstrCMODULENAME, "DetermineTTAAmount"

End Function

'***OP2***
Public Function IsHalfTrip() As String
'RED106 , 107
On Error GoTo IsHalfTrip_EH

    If HalfTripInd = "Y" Then
        IsHalfTrip = "Y"
    Else
        IsHalfTrip = "N"
    End If
    
Exit Function
IsHalfTrip_EH:
    
    TrsRaiseError mstrCMODULENAME, "IsHalfTrip"

End Function

Public Function GetEffectiveUnionCode() As String
    GetEffectiveUnionCode = OtherPay.Employee.UnionCode(ChargeDate)
End Function

Public Sub Delete()
    SetDeleted
End Sub

Public Sub Release()
'Donot enable an error handler since this is called from an error handler.
    Set mobjOtherPay = Nothing
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

Public Sub SetApprovalIndicator( _
                ByVal v_strApprInd As String, _
                ByVal v_strApprovedBy As String)
    
    ApprovalInd = v_strApprInd
    If ApprovalInd = "Y" Then
        ApproveByEmpNum = CStr(v_strApprovedBy)
        ApprovalDate = Date
    Else
        ApproveByEmpNum = 0
        ApprovalDate = CDate(CDEFAULTDATE)
    End If

End Sub

Private Function IsValidHoursEntry() As Boolean
'  Hours entered were not valid for Exception Code
    If hrs <> 0 Then
        IsValidHoursEntry = False
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Hours are not allowed for this Exception Code."
    Else
        IsValidHoursEntry = True
    End If
End Function

Private Function IsValidStartEndLocationEntry() As Boolean
'  Start and End Locations were not valid for Exception Code, will be set to 0 and user will be warned
    If StartLocation <> 0 Or EndLocation <> 0 Then
        IsValidStartEndLocationEntry = False
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Start/End Locations entered were not required for Exception Code."
    Else
        IsValidStartEndLocationEntry = True
    End If
End Function

Private Function IsValidStartEndLocation() As Boolean
' 08/28/2000 Per Prem, a TTA can have a start and end location of 0
    Dim lobjLocMgr As dbCodeTables.cdCodeTables
    Dim lstrStLoc As String
    Dim lstrNdLoc As String
    
    lstrStLoc = StartLocation
    lstrNdLoc = EndLocation
    If lstrStLoc = 0 And lstrNdLoc = 0 Then
        IsValidStartEndLocation = True
    Else
        Set lobjLocMgr = CreateObject("dbCodeTables.cdCodeTables")
        IsValidStartEndLocation = _
            (lobjLocMgr.LocationCodeIsValid(lstrStLoc) And _
            lobjLocMgr.LocationCodeIsValid(lstrNdLoc))
        Set lobjLocMgr = Nothing
    End If
    
End Function

Private Function IsValidAmountEntry() As Boolean
'  The Amount entered was not valid for the Exception Code, set to zero and warn user
    If Amt <> 0 Then
        IsValidAmountEntry = False
        AppendEntryErrorOtherPay trErrbusOtherPayEntryValidationError, mstrCBUSINESSERR, "Amount entered was not required for Exception Code."
    Else
        IsValidAmountEntry = True
    End If
End Function

Private Function FormatEntryBRV() As String

    FormatEntryBRV = _
        FormatBRV(TrType_TC, EmployeeNumber, ChargeDate, DateWorked, startTime, EndTime, hrs, ExceptionCode, "")

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

