VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cboEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Company code is determined by getting the effective
'   department and its associated company code.
' Union code and prior union code are derived from the first
'   two characters of job number and prior job number.
' ExecutiveContract indicator is determined by getting the
'   effective union and its associated execContract indicator.
' UnionBenefitPackage indicator is determined by getting the
'   effective union and its associated UnionBenefitPackage indicator.
'
Private mlngEmployeeNumber As Long
Private mstrEmployeeName As String
Private mintLocationCode As Integer
Private mstrShiftInd As String
Private mstrEmployeeStatus As String
Private mdatStatusEffDate As Date
Private mstrAccountType As String
Private mstrPayrollGroupCode As String
Private mstrInterfaceToCode As String
Private mstrAccountDivisionNum As String

Private mstrDepartmentNumber As String
Private mstrPreviousDepartmentNumber As String
Private mdatDepartmentEffDate As Date
Private mstrJobNumber As String
Private mdatJobEffDate As Date
Private mstrPriorJobNumber As String
Private mdatPriorJobEffDate As Date

Private mstrUnionBenPkg As String
Private mstrPriorUnionBenPkg As String
Private mstrManagementBenPkg As String
Private mstrUnionCode As String
Private mstrPriorUnionCode As String
Private mstrExecContractInd As String
Private mstrPriorExecContractInd As String
Private mstrCompanyCode As String
Private mstrPreviousCompanyCode As String
Private mstrOTType As String
Private mstrPriorOTType As String
Private mstrOTExclusn As String
Private mstrPriorOTExclusn As String
Private mstrOTExclusnEffDt As Date
Private mstrPriorOTExclusnEffDt As Date
Private mdatEmploymentDate As Date

Private mrsWorkgroups As ADODB.Recordset
'Preserve the start and end dates passed on restore
'request for lazy materialization.
Private mdatStartDate As Date
Private mdatEndDate As Date


Private Const m_strCMODULENAME As String = "cboEmployee"
Const mstrDefaultDate As String = "01/01/0100"

Public Property Get EmployeeNumber() As Long
    EmployeeNumber = mlngEmployeeNumber
End Property

Public Property Get EmployeeName() As String
    EmployeeName = mstrEmployeeName
End Property

Public Property Get EmployeeBenefitPackage(Optional _
        ByVal v_strChgDate As Date = mstrDefaultDate) As String
' WHITE039
    Dim l_datChgDate As Date

    On Error GoTo eh
    
    l_datChgDate = EffectiveDate(v_strChgDate)
    If IsExecutive(l_datChgDate) Then
        EmployeeBenefitPackage = mstrManagementBenPkg
    ElseIf UseCurrentJobAttribute(l_datChgDate) Then
        EmployeeBenefitPackage = mstrUnionBenPkg
    Else
        EmployeeBenefitPackage = mstrPriorUnionBenPkg
    End If
    
    If IsExecutive(l_datChgDate) Then
        EmployeeBenefitPackage = "UE"
    End If
    
    If IsAmerenCILCOLocal8Employee(l_datChgDate) Then
        EmployeeBenefitPackage = "UE"
    End If
    
    Exit Property
    
eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeBenefitPackage", , , , , True
End Property

Public Property Get LocationCode() As Integer
    LocationCode = mintLocationCode
End Property

Public Property Get JobEffectiveDate() As Date
    JobEffectiveDate = mdatJobEffDate
End Property

Public Property Get PriorJobNumber() As String
    PriorJobNumber = mstrPriorJobNumber
End Property

Public Property Get PriorJobEffDate() As Date
    PriorJobEffDate = mdatPriorJobEffDate
End Property

Public Property Get PreviousDepartmentNumber() As String
    PreviousDepartmentNumber = mstrPreviousDepartmentNumber
End Property

Public Property Get DepartmentEffDate() As Date
    DepartmentEffDate = mdatDepartmentEffDate
End Property

Public Property Get ShiftIndicator() As String
' WHITE038
' Y means a shift working executive
    ShiftIndicator = mstrShiftInd
End Property

Public Property Get Status() As String
    Status = mstrEmployeeStatus
End Property

Public Property Get StatusEffDate() As Date
    StatusEffDate = mdatStatusEffDate
End Property

Public Property Get EmploymentDate() As Date
    EmploymentDate = mdatEmploymentDate
End Property


Public Property Get AccountType() As String
    AccountType = mstrAccountType
End Property

Public Property Get PayrollGroupCode() As String
' RED053
    PayrollGroupCode = mstrPayrollGroupCode
End Property

Public Property Get InterfaceToCode() As String
    InterfaceToCode = mstrInterfaceToCode
End Property

Public Property Get CompanyCode(Optional _
        ByVal v_strChargeDate As Date = mstrDefaultDate) As String
' CompanyCode and PreviousCompanyCode are attributes of the
' effective department.

    On Error GoTo eh
    
    If EffectiveDate(v_strChargeDate) >= mdatDepartmentEffDate Then
        CompanyCode = mstrCompanyCode
    Else
        CompanyCode = mstrPreviousCompanyCode
    End If
    
    Exit Property

eh:
    TrsRaiseError m_strCMODULENAME, "CompanyCode", , , , , True

End Property


Public Property Get AccountDivisionNum() As String
    AccountDivisionNum = mstrAccountDivisionNum
End Property

Public Sub RestoreWithState( _
                ByRef r_rsEmployee As ADODB.Recordset, _
                ByRef r_rsEmpWorkgroups As ADODB.Recordset)
'Currently assuming r_rsEmployee positioned on a record and
' r_rsEmpWorkgroups contains at least 1 record.
'Must ensure nothing changes the filter on these recordsets
' or passed out of process.
    On Error GoTo Restore_eh
    
    SetState r_rsEmployee
    Set mrsWorkgroups = r_rsEmpWorkgroups
    
    Exit Sub
    
Restore_eh:

    Dim lcolErrArgs As Collection
    
    Set lcolErrArgs = New Collection
    If Not r_rsEmployee Is Nothing Then
        lcolErrArgs.Add "r_rsEmployee.Filter=" & r_rsEmployee.Filter
    Else
        lcolErrArgs.Add "The r_rsEmployee recordset = nothing"
    End If
    If Not r_rsEmpWorkgroups Is Nothing Then
        lcolErrArgs.Add "r_rsEmpWorkgroups.Filter=" & r_rsEmpWorkgroups.Filter
    Else
        lcolErrArgs.Add "The r_rsEmpWorkgroups recordset = nothing"
    End If
    Set r_rsEmployee = Nothing
    Set r_rsEmpWorkgroups = Nothing

    TrsRaiseError m_strCMODULENAME, "RestoreWithState", , , , lcolErrArgs, True
    
End Sub


Public Sub Restore( _
            ByVal v_lngEmpId As Long, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_blnDeferLoad As Boolean = True)
    
    On Error GoTo Restore_eh
    
    mdatStartDate = v_datStartDate
    mdatEndDate = v_datEndDate
    
    RestoreBasic v_lngEmpId
    
    If Not v_blnDeferLoad Then
        RestoreWorkgroups
    End If
    
    Exit Sub
    
Restore_eh:

    TrsRaiseError m_strCMODULENAME, "Restore", , , , , True
    
End Sub

Private Sub RestoreBasic(ByVal v_lngEmpId As Long)
    Dim dbEmployee As dbEmployee.cdEmployee
    Dim rsEmployee As ADODB.Recordset
    
    
    On Error GoTo eh
    
    
    Set dbEmployee = CreateObject("dbEmployee.cdEmployee")
    Set rsEmployee = dbEmployee.GetEmployee(v_lngEmpId)
    If rsEmployee.BOF And rsEmployee.EOF Then
        Err.Raise trErrboEmpNotFound ', , "Employee not found.."
    Else
        SetState rsEmployee
    End If
    Set rsEmployee = Nothing
    Set dbEmployee = Nothing
    
    Exit Sub
    
eh:
    Set rsEmployee = Nothing
    Set dbEmployee = Nothing
    Select Case Err.Number
       Case trErrboEmpNotFound
            TrsRaiseError m_strCMODULENAME, "RestoreBasic", trErrboEmpNotFound, , , , True
       Case Else
            TrsRaiseError m_strCMODULENAME, "RestoreBasic", , , , , True
    End Select
    
End Sub

Private Sub RestoreWorkgroups()
' mrsWorkGroups will contain primary workgroups only that the employee
' was assigned to between the start and end date.
    Dim dbWorkGroup As dbWorkGroup.cdWorkGroup
    Dim rsWorkGroups As ADODB.Recordset
    
    
    On Error GoTo RestoreWorkgroups_eh
    
    
    Set dbWorkGroup = CreateObject("dbWorkgroup.cdWorkgroup")
    Set rsWorkGroups = dbWorkGroup.GetEmployeeWorkGroups( _
                EmployeeNumber, mdatStartDate, mdatEndDate, True)
    If rsWorkGroups.BOF And rsWorkGroups.EOF Then
        Err.Raise trErrboWGNotFound  'must have at least 1 workgroup - Error occurred attempting to retrieve from Workgroup Table Information database '10011 replaced by 30083 + vbObjectError
    Else
        Set mrsWorkgroups = rsWorkGroups
    End If
    Set dbWorkGroup = Nothing
    Set rsWorkGroups = Nothing
       
    Exit Sub
    
RestoreWorkgroups_eh:
    Set dbWorkGroup = Nothing
    Set rsWorkGroups = Nothing
    Dim l_colDescArgs As Collection
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(EmployeeNumber)
    l_colDescArgs.Add str(mdatStartDate)
    l_colDescArgs.Add str(mdatEndDate)

    Select Case Err.Number
       Case trErrboWGNotFound
            TrsRaiseError m_strCMODULENAME, "RestoreWorkgroups", trErrboWGNotFound, , , l_colDescArgs, True
       Case Else
            TrsRaiseError m_strCMODULENAME, "RestoreWorkgroups", , , , , True
    End Select
    
End Sub

Private Sub SetState(rs As ADODB.Recordset)

    On Error GoTo eh

    mlngEmployeeNumber = rs!EMP_NUM
    mstrEmployeeName = Trim(rs!EMP_NAME)
    mstrUnionCode = Trim(rs!UNION_CD)
    mstrDepartmentNumber = Trim(rs!DEPT_NUM)
    mstrManagementBenPkg = Trim(rs!MGMT_BEN_PKG_IND)
    mstrUnionBenPkg = Trim(rs!UNION_BEN_PKG_IND)
    mintLocationCode = rs!LOC_CD
    mstrJobNumber = Trim(rs!JOB_NUM)
    mdatJobEffDate = rs!JOB_EFF_DATE
    mstrPriorJobNumber = Trim(rs!PRIOR_JOB_NUM)
    mdatPriorJobEffDate = rs!PRIOR_JOB_EFF_DT
    mstrPreviousDepartmentNumber = Trim(rs!PREV_DEPT_NUM)
    mdatDepartmentEffDate = rs!DEPT_EFF_DATE
    mstrShiftInd = Trim(rs!SHIFT_IND)
    mstrExecContractInd = Trim(rs!EXEC_CONTRACT_IND)
    mstrEmployeeStatus = Trim(rs!EMP_STATUS)
    mdatStatusEffDate = rs!STATUS_EFF_DATE
    mdatEmploymentDate = rs!EMPLOY_DATE
    mstrAccountType = Trim(rs!ACCT_TYPE)
    mstrPayrollGroupCode = Trim(rs!PAYROLL_GROUP_CD)
    mstrInterfaceToCode = Trim(rs!INTERFACE_TO_CD)
    mstrCompanyCode = Trim(rs!CMPY_CD)
    mstrAccountDivisionNum = Trim(rs!ACCT_DIV_NUM)
    mstrPriorUnionCode = Trim(rs!PRIOR_UNION_CD)
    mstrPriorExecContractInd = Trim(rs!PRIOR_EXEC_CON_IND)
    mstrPriorUnionBenPkg = Trim(rs!PRIOR_UNION_BN_PKG)
    mstrPreviousCompanyCode = Trim(rs!PREV_CMPY_CD)
    mstrOTType = Trim(rs!OT_TYPE)
    mstrPriorOTType = Trim(rs!PRIOR_OT_TYPE)
    mstrOTExclusn = Trim(rs!AR_OT_EXCLUSN)
    mstrPriorOTExclusn = Trim(rs!PR_AR_OT_EXCLUSN)
    mstrOTExclusnEffDt = rs!EFFDT
    mstrPriorOTExclusnEffDt = rs!PRIOR_EFFDT
      
    Exit Sub

eh:

    TrsRaiseError m_strCMODULENAME, "SetState", , , , , True

End Sub

Public Sub Release()

    Set mrsWorkgroups = Nothing

End Sub

Public Function PartialHourMinutes(ByVal v_datChargeDate As Date) As Integer
' Answer the minutes that partial hours may be entered.
    Dim l_strDepartmentNumber As String
    
    On Error GoTo eh
    
    l_strDepartmentNumber = DepartmentNumber(v_datChargeDate)
    If DoesRuleApply(166, trDepartmentNumber, l_strDepartmentNumber) Then
        PartialHourMinutes = 3  'Callaway may be going to 3 minute increments
    Else
        If IsContract(v_datChargeDate) Then
            PartialHourMinutes = 6
        Else                            'must be executive
            If DoesRuleApply(165, trDepartmentNumber, l_strDepartmentNumber) Then
                    PartialHourMinutes = 6
            Else
                    PartialHourMinutes = 30
            End If
        End If
    End If
    
    Exit Function

eh:

    TrsRaiseError m_strCMODULENAME, "PartialHourMinutes", , , , , True

End Function

Public Function GetWorkgroups() As ADODB.Recordset
    If mrsWorkgroups Is Nothing Then
        RestoreWorkgroups
    End If
    Set GetWorkgroups = mrsWorkgroups
End Function

Public Function GetPrimaryWorkgroupNumber(ByVal v_strChargeDate As Date) As Long
'  Determine the Employees primary workgroup number based on Charge Date
'  Answer zero if not found.
    Dim l_rsWg As ADODB.Recordset

On Error GoTo eh

    Set l_rsWg = GetWorkgroups
    l_rsWg.MoveFirst
    Do While Not l_rsWg.EOF
        If l_rsWg!PRIMARY_IND = "Y" And _
            v_strChargeDate >= l_rsWg!EFF_START_DATE _
             And v_strChargeDate <= l_rsWg!EFF_END_DATE Then
            GetPrimaryWorkgroupNumber = l_rsWg!WG_NUM
            Exit Function
        End If
        l_rsWg.MoveNext
    Loop

    GetPrimaryWorkgroupNumber = 0
    Set l_rsWg = Nothing
    Exit Function

eh:
    Set l_rsWg = Nothing

    TrsRaiseError m_strCMODULENAME, "GetPrimaryWorkgroupNumber", , , , , True

End Function

Public Function GetProjectId(ByVal v_lngWgNum As Long) As Long
' BLUE084
    Dim lobjWgMgr As dbWorkGroup.cdWorkGroup

On Error GoTo eh

    Set lobjWgMgr = CreateObject("dbWorkGroup.cdWorkGroup")
    GetProjectId = lobjWgMgr.ProjectID(v_lngWgNum)
    
    Set lobjWgMgr = Nothing
    Exit Function
eh:
    Set lobjWgMgr = Nothing
    TrsRaiseError m_strCMODULENAME, "GetProjectId", , True, , , True

End Function

Public Function GetPPMProjectPeriod(ByVal v_lngWgNum As Long) As Variant
' If v_lngWgNum is a PPM workgroup(i.e. project_id <> 0),
' answer an array with 2 elements, first project start date and
' second project end date.  If not found, answer empty.
' BLUE083
    Dim l_varPeriod As Variant
    Dim lobjWgMgr As dbWorkGroup.cdWorkGroup
    Dim lobjWgRecord As Variant

On Error GoTo eh

    Set lobjWgMgr = CreateObject("dbWorkGroup.cdWorkGroup")
    If lobjWgMgr.ProjectID(v_lngWgNum) <> 0 Then
        l_varPeriod = _
            Array(lobjWgMgr.ProjectStartDate(v_lngWgNum), _
                lobjWgMgr.ProjectEndDate(v_lngWgNum))
    Else
        l_varPeriod = Empty
    End If
    GetPPMProjectPeriod = l_varPeriod
    Set lobjWgMgr = Nothing

    Exit Function
eh:
    Set lobjWgMgr = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPPMProjectPeriod", , , , , True

End Function

Public Property Get DepartmentNumber(Optional _
        ByVal v_strChargeDate As Date = mstrDefaultDate) As String
' RED054
    On Error GoTo eh
    
    If EffectiveDate(v_strChargeDate) >= mdatDepartmentEffDate Then
        DepartmentNumber = mstrDepartmentNumber
    Else
        DepartmentNumber = PreviousDepartmentNumber
    End If
    
    Exit Property

eh:
    TrsRaiseError m_strCMODULENAME, "DepartmentNumber", , , , , True

End Property

Public Property Get JobNumber(Optional _
        ByVal v_strChargeDate As Date = mstrDefaultDate) As String
' WHITE082
    Dim l_datChgDate As Date
    Dim l_strJobNum As String
    
    On Error GoTo eh

    l_datChgDate = EffectiveDate(v_strChargeDate)

    If l_datChgDate >= JobEffectiveDate Then
        l_strJobNum = mstrJobNumber
    ElseIf l_datChgDate >= PriorJobEffDate Then
        l_strJobNum = PriorJobNumber
        If l_strJobNum = "" Then
            l_strJobNum = mstrJobNumber
        End If
    Else
        l_strJobNum = mstrJobNumber
    End If
    
    JobNumber = l_strJobNum
    
    Exit Property

eh:
    TrsRaiseError m_strCMODULENAME, "JobNumber", , , , , True

End Property

Public Property Get OTType(Optional _
        ByVal v_strChargeDate As Date = mstrDefaultDate) As String
' WHITE082
' The OT Type is 100% directly related to the Job
' Uses same logic to determine effective OT Type as
' the logic to determine effective job number
' Q indicates a quarterly executive
' H indicates an executive that gets paid for all ot
' N indicates executive does not get paid for any ot
' " " blank indicates non-executive
    Dim l_datChgDate As Date
    Dim l_strOTType As String
    
    On Error GoTo eh

    l_datChgDate = EffectiveDate(v_strChargeDate)

    If l_datChgDate >= JobEffectiveDate Then
        l_strOTType = mstrOTType
    ElseIf l_datChgDate >= PriorJobEffDate Then
        l_strOTType = mstrPriorOTType
        If l_strOTType = "" Then
            l_strOTType = mstrOTType
        End If
    Else
        l_strOTType = mstrOTType
    End If
    
    OTType = l_strOTType
    
    Exit Property

eh:
    TrsRaiseError m_strCMODULENAME, "OTType", , , , , True

End Property
Public Function GetWorkgroupLocationCode(ByVal v_lngWgNum As Long) As Integer
    Dim lobjWgMgr As dbWorkGroup.cdWorkGroup

On Error GoTo eh

    Set lobjWgMgr = CreateObject("dbWorkGroup.cdWorkGroup")
    GetWorkgroupLocationCode = lobjWgMgr.LocationCode(v_lngWgNum)

    Set lobjWgMgr = Nothing
    Exit Function
eh:
    Set lobjWgMgr = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkgroupLocationCode", , True, , , True

End Function

Public Property Get UnionCode(Optional _
        ByVal v_strChargeDate As Date = mstrDefaultDate) As String
' BLUE067, RED152, RED153
    On Error GoTo eh
    
    If UseCurrentJobAttribute(EffectiveDate(v_strChargeDate)) Then
        UnionCode = mstrUnionCode
    Else
        UnionCode = mstrPriorUnionCode
    End If
    
    Exit Property

eh:
    TrsRaiseError m_strCMODULENAME, "UnionCode", , , , , True

End Property

Public Property Get ExecContractIndicator(Optional _
        ByVal v_strChargeDate As Date = mstrDefaultDate) As String
' WHITE037, WHITE084
    On Error GoTo eh
    
    If UseCurrentJobAttribute(EffectiveDate(v_strChargeDate)) Then
        ExecContractIndicator = mstrExecContractInd
    Else
        ExecContractIndicator = mstrPriorExecContractInd
    End If
    Exit Property
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExecContractIndicator", , , , , True

End Property

Private Function UseCurrentJobAttribute( _
            ByVal v_datChgDate As Date) As Boolean
            
    If v_datChgDate >= JobEffectiveDate Or _
            PriorJobNumber = "" Then
        UseCurrentJobAttribute = True
    End If

End Function

Public Function IsEntitledToSaturdayPremium() As Boolean
' BLUE066
    IsEntitledToSaturdayPremium = _
        IsStringInList(UnionCode, Array("JA", "TA", "FA", "JB", "TB"))

End Function

Public Function IsEntitledToSundayPremium() As Boolean
' BLUE068
    Dim l_varUnionList As Variant
'    l_varUnionList = Array("XA", "ZA", "IA", "JA", "TA", "FA", _
'                        "PA", "VA", "SA", "EA", "DA", "HA", _
'                        "QA", "KA", "OA", "NA", "MA", "LA")
'The list above is incomplete, did not include new GENCO unions
'There are a select few unions that do not get SUS time

    l_varUnionList = Array("AA", "AC", "AD", "BA", "CA", "RA", "WA", "IP")
    If IsStringInList(UnionCode, l_varUnionList) Then
        IsEntitledToSundayPremium = False
    Else
        IsEntitledToSundayPremium = True
    End If
    
    

End Function

Public Function IsUnion148(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' WHITE053
    On Error GoTo eh
    
    IsUnion148 = _
        UnionCode(EffectiveDate(v_datChgDate)) = trUnion148
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsUnion148", , , , , True
End Function

Public Function IsExecutive(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' BLUE011, BLUE059
    On Error GoTo eh
    
    IsExecutive = _
        ExecContractIndicator(EffectiveDate(v_datChgDate)) = "E"
        
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsExecutive", , , , , True
End Function

Public Function IsContract(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' BLUE061, RED043
    On Error GoTo eh
    
    IsContract = Not IsExecutive(EffectiveDate(v_datChgDate))
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsContract", , , , , True
End Function


Public Function IsLunchFloater(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
    
On Error GoTo eh
    
    If IsDOJM(v_datChgDate) Then
        IsLunchFloater = True
    ElseIf IsSubstationOrTransmission(v_datChgDate) Then
        IsLunchFloater = True
    Else
        IsLunchFloater = False
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsLunchFloater", , , , , True
End Function


Public Function IsDOJM(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' cache this later
    Dim l_strDeptNum As String
    Dim l_objDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh
    l_strDeptNum = DepartmentNumber(EffectiveDate(v_datChgDate))
    Set l_objDepartment = _
        CreateObject("busDepartment.cbsDepartment")
    '//IsDOJM = l_objDepartment.IsDOJMDepartmentNumber(l_strDeptNum)
    IsDOJM = DoesRuleApply(1001, trDepartmentNumber, l_strDeptNum)
    Set l_objDepartment = Nothing
    Exit Function
    
eh:
    Set l_objDepartment = Nothing
    TrsRaiseError m_strCMODULENAME, "IsDOJM", , , , , True
End Function

Public Function IsSubstationOrTransmission(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
    Dim l_strDeptNum As String
    Dim l_strUnionCd As String
    
On Error GoTo eh
    
    l_strDeptNum = DepartmentNumber(EffectiveDate(v_datChgDate))
    l_strUnionCd = UnionCode(EffectiveDate(v_datChgDate))
    IsSubstationOrTransmission = _
        IsSubstationOrTransmissionEmployee(l_strUnionCd, l_strDeptNum)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsSubstationOrTransmission", , , , , True
End Function

Public Function IsSubstationOrTransmissionEmployee( _
            ByVal v_strUnionCode As String, _
            ByVal v_strDeptNum As String) As Boolean
'WHITE003

On Error GoTo eh
    Dim rc As Boolean
    
    rc = False
    rc = DoesRuleApply(dbRules.TRSRuleCode.trIsSubstationOrTransmissionEmployee, trDepartmentNumber, v_strDeptNum)
    
    If Not rc Then
        rc = DoesRuleApply(dbRules.TRSRuleCode.trIsSubstationOrTransmissionEmployee, trDeptNumAndUnionCode, v_strDeptNum & ":" & PadRight(v_strUnionCode, 3))
    End If

    IsSubstationOrTransmissionEmployee = rc
    
'    If ((Trim(v_strUnionCode) = "IA") And _
'        (v_strDeptNum = "033" Or _
'         v_strDeptNum = "041" Or _
'         v_strDeptNum = "048" Or _
'         v_strDeptNum = "056" Or _
'         v_strDeptNum = "060" Or _
'         v_strDeptNum = "410" Or _
'         v_strDeptNum = "175" Or _
'         v_strDeptNum = "176" Or _
'         v_strDeptNum = "177" Or _
'         v_strDeptNum = "178")) Or _
'       (v_strDeptNum = "333") Or _
'       (v_strDeptNum = "360") Then
'        IsSubstationOrTransmissionEmployee = True
'    Else
'        IsSubstationOrTransmissionEmployee = False
'    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsSubstationOrTransmissionEmployee", , , , , True
End Function

Public Function IsExecutiveShiftWorker(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' WHITE038
    On Error GoTo eh
    
    IsExecutiveShiftWorker = _
        IsExecutive(v_datChgDate) And ShiftIndicator = "Y"
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsExecutiveShiftWorker", , , , , True
End Function

Public Function IsQuarterlyExecutive(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
    On Error GoTo eh
    
    IsQuarterlyExecutive = OTType(v_datChgDate) = "Q"
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsQuarterlyExecutive", , , , , True
End Function

Public Function IsAmerenCIPS(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean

' RED154
'RULE0011
On Error GoTo eh

    IsAmerenCIPS = IsAmerenCIPSorCILCO(dbRules.TRSRuleCode.trIsAmerenCIPS, v_datChgDate)

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPS", , , , , True
End Function

'//for unit testing
Public Function IsAmerenCIPSorCILCOTest(ByVal v_lngRuleNum As Long, _
                                        ByVal v_lngEmpId As Long, _
                                        ByVal v_datChgDate As Date) As Boolean

On Error GoTo eh

    Me.Restore v_lngEmpId, v_datChgDate, v_datChgDate
    IsAmerenCIPSorCILCOTest = IsAmerenCIPSorCILCO(v_lngRuleNum, v_datChgDate)

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPSorCILCOTest", , , , , True
End Function

Public Function IsAmerenUE(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean

    On Error GoTo eh
    
    IsAmerenUE = CompanyCode(EffectiveDate(v_datChgDate)) = "UEC"
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenUE", , , , , True
End Function

Public Function IsCoffeen(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' Currently, this is used to determine shift component
' which has a special rule for coffeen maintenance and
' operations.  However, this is not inclusive of all
' coffeen.  Location code should be used to determine that.
    On Error GoTo eh
    
    IsCoffeen = IsCoffeenMaintenance(v_datChgDate) Or _
                    IsCoffeenOperations(v_datChgDate)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsCoffeen", , , , , True

End Function

Public Function IsCoffeenMaintenance(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
    ' RED155
    'RULE0009
    
    Dim l_obj As busDepartment.cbsDepartment
    Dim l_strDept As String

    On Error GoTo eh
    
    l_strDept = DepartmentNumber(EffectiveDate(v_datChgDate))
        
    Set l_obj = CreateObject("busDepartment.cbsDepartment")
    If l_obj.IsCoffeenMaintenance(l_strDept) Then 'RULE0009
        IsCoffeenMaintenance = True
    Else
        IsCoffeenMaintenance = False
    End If
    
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsCoffeenMaintenance", , , , , True

End Function

Public Function IsAmerenLeadership(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' Always an executive employee, big dog.
    Dim l_strDeptNum As String
    
    On Error GoTo eh
    
    l_strDeptNum = DepartmentNumber(EffectiveDate(v_datChgDate))
    Select Case l_strDeptNum
        Case "545"
            IsAmerenLeadership = True
        Case Else
            IsAmerenLeadership = False
    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenLeadership", , , , , True

End Function

Public Function IsCoffeenOperations(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' RED156
'RULE0010

    Dim l_obj As busDepartment.cbsDepartment
    Dim l_strDept As String
    
    On Error GoTo eh
    
    l_strDept = DepartmentNumber(EffectiveDate(v_datChgDate))
        
    Set l_obj = CreateObject("busDepartment.cbsDepartment")
    If l_obj.IsCoffeenOperations(l_strDept) Then 'RULE0010
        IsCoffeenOperations = True
    Else
        IsCoffeenOperations = False
    End If
    
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsCoffeenOperations", , , , , True

End Function

Public Function IsAmerenCIPSBenefitParticipant(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' RED067
    On Error GoTo eh
       
    IsAmerenCIPSBenefitParticipant = _
        EmployeeBenefitPackage(EffectiveDate(v_datChgDate)) = "CIPS"
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPSBenefitParticipant", , , , , True
End Function

Private Function EffectiveDate(ByVal v_datChgDate As Date) As Date
    
    If v_datChgDate = mstrDefaultDate Then
        EffectiveDate = Date
    Else
        EffectiveDate = v_datChgDate
    End If

End Function

'''''
'''''
'''''Public Function GetTimeIncrementByChargeDate( _
'''''        ByVal v_lngEmpNum As Long, _
'''''        ByVal v_datStartDate As Date, _
'''''        ByVal v_datEndDate As Date, _
'''''        Optional ByRef r_bytProperties As Variant, _
'''''        Optional ByVal v_intRequestType As Integer = 1) As ADODB.Recordset
'''''' Answer a recordset and optional properties depending on the type
'''''' of request.  Trying to make this a dual mode procedure with
'''''' capability to operate in stateless and stateful mode.
'''''' Properties are answered as a byte array which are the contents of
'''''' a property bag.  Pass as byte array for marshalling.
'''''' 1 (default) - answer employee DEPT_NUM, PREV_DEPT_NUM, DEPT_EFF_DATE
'''''    Dim l_rsTimeIncrByChargeDate As ADODB.Recordset
'''''    Dim l_lngNumberofDays As Long
'''''    Dim l_intIndex As Integer
'''''    Dim l_datChargeDate As Date
'''''    Dim l_intAllowTimeIncr As Integer
'''''    Dim l_pbProperties As PropertyBag
'''''
'''''    On Error GoTo eh
'''''
'''''    If Not IsRestored Then
'''''        Restore v_lngEmpNum, v_datStartDate, v_datEndDate
'''''    End If
'''''
'''''    Select Case v_intRequestType
'''''        Case 1
'''''            Set l_pbProperties = New PropertyBag
'''''            With l_pbProperties
'''''                .WriteProperty "DEPT_NUM", DepartmentNumber
'''''                .WriteProperty "PREV_DEPT_NUM", PreviousDepartmentNumber
'''''                .WriteProperty "DEPT_EFF_DATE", DepartmentEffDate
'''''                r_bytProperties = .Contents
'''''            End With
'''''    End Select
'''''    'added new field UNION_CD needed for cbsEmployee to determine subtation or transmission 'WHITE003
'''''    Set l_rsTimeIncrByChargeDate = New ADODB.Recordset
'''''    With l_rsTimeIncrByChargeDate
'''''      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
'''''      .Fields.Append "ALLOW_TIME_INCR", adInteger, , adFldUpdatable
'''''      .Fields.Append "UNION_CD", adChar, 3, adFldUpdatable
'''''      .Open
'''''    End With
'''''
'''''    l_lngNumberofDays = DateDiff("d", v_datStartDate, v_datEndDate)
'''''    For l_intIndex = 0 To l_lngNumberofDays
'''''        l_datChargeDate = DateAdd("d", l_intIndex, v_datStartDate)
'''''        l_intAllowTimeIncr = PartialHourMinutes(l_datChargeDate)
'''''        l_rsTimeIncrByChargeDate.AddNew
'''''        l_rsTimeIncrByChargeDate!CHARGE_DATE = l_datChargeDate
'''''        l_rsTimeIncrByChargeDate!ALLOW_TIME_INCR = l_intAllowTimeIncr
'''''        l_rsTimeIncrByChargeDate!UNION_CD = UnionCode(l_datChargeDate)
'''''        l_rsTimeIncrByChargeDate.Update
'''''    Next
'''''
'''''    l_rsTimeIncrByChargeDate.MoveFirst
'''''    Set GetTimeIncrementByChargeDate = l_rsTimeIncrByChargeDate
'''''    Set l_rsTimeIncrByChargeDate = Nothing
'''''    Exit Function
'''''
'''''eh:
'''''    Set l_rsTimeIncrByChargeDate = Nothing
'''''    TrsRaiseError m_strCMODULENAME, "GetTimeIncrementByChargeDate", , , , , True
'''''
'''''End Function

Private Function IsRestored() As Boolean

    IsRestored = mlngEmployeeNumber <> 0

End Function

Private Sub Class_Initialize()
    mdatStartDate = Date
    mdatEndDate = Date
End Sub



Public Function GetTimeIncrementByChargeDate( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_rsEmployeeData As ADODB.Recordset _
        ) As ADODB.Recordset
        
' This method is purely stateless
' The employee data is already retrieved prior to this method
' being called so it is passed.
' The code is in this module because all employee rules are
' in this object.  Thus all code in one place, but interface
' is different

    Dim l_rsTimeIncrByChargeDate As ADODB.Recordset
    Dim l_lngNumberofDays As Long
    Dim l_intIndex As Integer
    Dim l_datChargeDate As Date
    Dim l_intAllowTimeIncr As Integer
    Dim l_strEmpNum As String
    
    
    On Error GoTo eh
    
    
    Set l_rsTimeIncrByChargeDate = New ADODB.Recordset
    With l_rsTimeIncrByChargeDate
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "ALLOW_TIME_INCR", adInteger, , adFldUpdatable
      .Fields.Append "UNION_CD", adChar, 3, adFldUpdatable
      .Open
    End With
    
    l_lngNumberofDays = DateDiff("d", v_datStartDate, v_datEndDate)
    
    If RecordsExist(v_rsEmployeeData) Then
        v_rsEmployeeData.MoveFirst
        While Not v_rsEmployeeData.EOF
            SetState v_rsEmployeeData
            For l_intIndex = 0 To l_lngNumberofDays
                l_datChargeDate = DateAdd("d", l_intIndex, v_datStartDate)
                l_intAllowTimeIncr = PartialHourMinutes(l_datChargeDate)
                l_rsTimeIncrByChargeDate.AddNew
                l_rsTimeIncrByChargeDate!EMP_NUM = EmployeeNumber
                l_rsTimeIncrByChargeDate!CHARGE_DATE = l_datChargeDate
                l_rsTimeIncrByChargeDate!ALLOW_TIME_INCR = l_intAllowTimeIncr
                l_rsTimeIncrByChargeDate!UNION_CD = UnionCode(l_datChargeDate)
                l_rsTimeIncrByChargeDate.Update
            Next l_intIndex
            v_rsEmployeeData.MoveNext
        Wend
    End If
    
    l_rsTimeIncrByChargeDate.Sort = "EMP_NUM, CHARGE_DATE"
    l_rsTimeIncrByChargeDate.MoveFirst
    Set GetTimeIncrementByChargeDate = l_rsTimeIncrByChargeDate
    Set l_rsTimeIncrByChargeDate = Nothing
    Exit Function
    
eh:
    Set l_rsTimeIncrByChargeDate = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimeIncrementByChargeDate", , , , , True

End Function

Public Function GetOTBankChargeInfoByChargeDate( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_rsEmployeeData As ADODB.Recordset = Nothing _
        ) As ADODB.Recordset

    Dim l_rsOTInfoByChargeDate As ADODB.Recordset
    Dim l_lngNumberofDays As Long
    Dim l_intIndex As Integer
    Dim l_datChargeDate As Date
    Dim l_strEffectiveUnionCode As String
    Dim l_strEffectiveExecContractInd As String
    Dim l_strEffectiveOTType As String
    Dim l_strShiftInd As String
    Dim l_strPayrollGroupCd As String
    Dim l_lngStartEmployee As Long
    Dim l_lngEndEmployee As Long
    Dim l_lngI As Long
    Dim l_rsEmployees As ADODB.Recordset
    Dim l_strDeptPayrollGroup(50, 1) As String
    Dim l_intDeptPayrollGroupCount As Integer
    Dim l_strDept As String
    Dim l_intPos As Integer
    Dim l_aryEmployees As Variant
    
    
    On Error GoTo eh
    
    If v_rsEmployeeData Is Nothing Then
        Set l_rsEmployees = GetEmployees(v_aryEmployees)
    Else
        Set l_rsEmployees = v_rsEmployeeData
    End If
    Set l_rsOTInfoByChargeDate = New ADODB.Recordset
    With l_rsOTInfoByChargeDate
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "UNION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "OT_TYPE", adChar, 1, adFldUpdatable
      .Fields.Append "EXEC_CONTRACT_IND", adChar, 1, adFldUpdatable
      .Fields.Append "SHIFT_IND", adChar, 1, adFldUpdatable
      .Fields.Append "PAYROLL_GROUP_CD", adChar, 1, adFldUpdatable
      .Fields.Append "DEPT_NUM", adChar, 3, adFldUpdatable
      .Fields.Append "LOC_CD", adInteger, , adFldUpdatable
' 11/17/00 rsg let this default
'      .CursorLocation = adUseClient
'      .CursorType = adOpenDynamic
'      .LockType = adLockOptimistic
      .Open
    End With
    

' Please Note that Payroll Group Code is not effective
' date driven.  Jim B. chose not to address it at this time.
' It could be by using the department effective
' date.  The previous department payroll group code
' will need to be retrieved.
' This can be done by
    ' Changing method GetDepartment in the dbDepartment object
    '            to retrieve the Payroll_group_cd
    ' Changing the following in dbEmployee Object
    '   -- Change SQL in method GetEmployeeBasic to have a
    '            placeholder for prior_pay_grp_cd
    '   -- Change method SetPriorDepartmentAttributes to
    '            set the payroll group code for the
    '            prior department
    

    If IsArray(v_aryEmployees) Then
        l_aryEmployees = v_aryEmployees
    Else
        l_aryEmployees = Array(CLng(v_aryEmployees))
    End If
    l_lngStartEmployee = LBound(l_aryEmployees)
    l_lngEndEmployee = UBound(l_aryEmployees)
    l_lngNumberofDays = DateDiff("d", v_datStartDate, v_datEndDate)
    
    For l_lngI = l_lngStartEmployee To l_lngEndEmployee
        l_rsEmployees.Filter = "EMP_NUM = " & l_aryEmployees(l_lngI)
        SetState l_rsEmployees
        l_strShiftInd = ShiftIndicator()
        l_strPayrollGroupCd = PayrollGroupCode()
        For l_intIndex = 0 To l_lngNumberofDays
            l_datChargeDate = DateAdd("d", l_intIndex, v_datStartDate)
            l_strEffectiveUnionCode = UnionCode(l_datChargeDate)
            l_strEffectiveExecContractInd = ExecContractIndicator(l_datChargeDate)
            l_strEffectiveOTType = OTType(l_datChargeDate)
            l_rsOTInfoByChargeDate.AddNew
            l_rsOTInfoByChargeDate!EMP_NUM = l_aryEmployees(l_lngI)
            l_rsOTInfoByChargeDate!CHARGE_DATE = l_datChargeDate
            l_rsOTInfoByChargeDate!UNION_CD = l_strEffectiveUnionCode
            l_rsOTInfoByChargeDate!OT_TYPE = l_strEffectiveOTType
            l_rsOTInfoByChargeDate!EXEC_CONTRACT_IND = l_strEffectiveExecContractInd
            l_rsOTInfoByChargeDate!SHIFT_IND = l_strShiftInd
            l_rsOTInfoByChargeDate!PAYROLL_GROUP_CD = l_strPayrollGroupCd
            l_rsOTInfoByChargeDate!DEPT_NUM = DepartmentNumber(l_datChargeDate)
            l_rsOTInfoByChargeDate!LOC_CD = LocationCode
            l_rsOTInfoByChargeDate.Update
        Next l_intIndex
    Next l_lngI
    
'Determine Payroll Group for each Department
    l_rsOTInfoByChargeDate.Sort = "EMP_NUM, CHARGE_DATE"
    If Not l_rsOTInfoByChargeDate.EOF And Not l_rsOTInfoByChargeDate.BOF Then
        l_rsOTInfoByChargeDate.MoveFirst
    End If
    l_intDeptPayrollGroupCount = 0
    While Not l_rsOTInfoByChargeDate.EOF
        l_strDept = l_rsOTInfoByChargeDate!DEPT_NUM
        If IsStringInList2(l_strDept, l_strDeptPayrollGroup, l_intDeptPayrollGroupCount, l_intPos) Then
            l_strPayrollGroupCd = l_strDeptPayrollGroup(l_intPos, 1)
        Else
            l_strPayrollGroupCd = GetPayrollGroupForDept(l_strDept)
            l_strDeptPayrollGroup(l_intDeptPayrollGroupCount, 0) = l_strDept
            l_strDeptPayrollGroup(l_intDeptPayrollGroupCount, 1) = l_strPayrollGroupCd
            l_intDeptPayrollGroupCount = l_intDeptPayrollGroupCount + 1
        End If
        l_rsOTInfoByChargeDate!PAYROLL_GROUP_CD = l_strPayrollGroupCd
        l_rsOTInfoByChargeDate.MoveNext
    Wend

    l_rsOTInfoByChargeDate.MoveFirst
    Set GetOTBankChargeInfoByChargeDate = l_rsOTInfoByChargeDate

    Set l_rsOTInfoByChargeDate = Nothing
    Exit Function
eh:
    Set l_rsOTInfoByChargeDate = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTBankChargeInfoByChargeDate", , , , , True

End Function

Private Function GetEmployees( _
                ByVal v_aryEmployees As Variant _
                ) As ADODB.Recordset
                
    Dim dbEmployee As dbEmployee.cdEmployee
    Dim rsEmployee As ADODB.Recordset
    
    
    On Error GoTo eh
    
    
    Set dbEmployee = CreateObject("dbEmployee.cdEmployee")
    Set rsEmployee = dbEmployee.GetEmployee(v_aryEmployees)
    If rsEmployee.BOF And rsEmployee.EOF Then
        Err.Raise trErrboEmpNotFound ', , "Employee not found..."
    End If
    Set GetEmployees = rsEmployee
    Set dbEmployee = Nothing
    Set rsEmployee = Nothing
    
    Exit Function
    
eh:
    Set dbEmployee = Nothing
    Set rsEmployee = Nothing
    Select Case Err.Number
       Case trErrboEmpNotFound
            TrsRaiseError m_strCMODULENAME, "GetEmployees", trErrboEmpNotFound, , , , True
       Case Else
            TrsRaiseError m_strCMODULENAME, "GetEmployees", , , , , True
    End Select
    
End Function




Public Function GetValidateDistIdInfoByChargeDate( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_rsEmployeeData As ADODB.Recordset _
        ) As ADODB.Recordset

    Dim l_rsValidateInfoByChargeDate As ADODB.Recordset
    Dim l_lngNumberofDays As Long
    Dim l_intIndex As Integer
    Dim l_datChargeDate As Date
    Dim l_strEffectiveDepartment As String
    Dim l_lngStartEmployee As Long
    Dim l_lngEndEmployee As Long
    Dim l_lngI As Long
    Dim l_rsEmployees As ADODB.Recordset
'    Dim l_obj As dbDepartment.cdDepartment
    Dim l_strDepartments As Variant
    Dim l_dbEffectiveDateRules As dbRules.cdEffectiveDateRules
    Dim l_dbRules As dbRules.cdRules
    Dim l_strAICOldDepartment As String
    
    On Error GoTo eh
    
'    Set l_obj = CreateObject("dbDepartment.cdDepartment")
    If v_rsEmployeeData Is Nothing Then
        Set l_rsEmployees = GetEmployees(v_aryEmployees)
    Else
        Set l_rsEmployees = v_rsEmployeeData
    End If
    Set l_rsValidateInfoByChargeDate = New ADODB.Recordset
    With l_rsValidateInfoByChargeDate
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "DEPT_NUM", adChar, 3, adFldUpdatable
      .Fields.Append "CMPY_CD", adChar, 3, adFldUpdatable
      .Fields.Append "ACCT_TYPE", adChar, 1, adFldUpdatable
'      .Fields.Append "INTERFACE_TO_CD", adChar, 1, adFldUpdatable
      .Fields.Append "ACCT_DIV_NUM", adChar, 3, adFldUpdatable
      .Open
    End With
    
    If l_rsEmployees.EOF And l_rsEmployees.BOF Then
        Set GetValidateDistIdInfoByChargeDate = l_rsValidateInfoByChargeDate
        Exit Function
    End If
    
    Set l_dbEffectiveDateRules = CreateObject("dbRules.cdEffectiveDateRules")
    Set l_dbRules = CreateObject("dbRules.cdRules")

    l_lngNumberofDays = DateDiff("d", v_datStartDate, v_datEndDate)
    l_rsEmployees.MoveFirst
    While Not l_rsEmployees.EOF
        SetState l_rsEmployees
        For l_intIndex = 0 To l_lngNumberofDays
            l_datChargeDate = DateAdd("d", l_intIndex, v_datStartDate)
            l_strEffectiveDepartment = DepartmentNumber(l_datChargeDate)
            l_rsValidateInfoByChargeDate.AddNew
            l_rsValidateInfoByChargeDate!EMP_NUM = l_rsEmployees!EMP_NUM
            l_rsValidateInfoByChargeDate!CHARGE_DATE = l_datChargeDate
            
            '//for AIC Department cut-over only, since AIC departments will be converted prior to 1/1 use the old departments
            '//TODO: Not needed after AIC implementation
            If l_dbEffectiveDateRules.DoesEffectiveDateRuleApply(dbRules.TRSRuleCode.trAICDeptCutOverEffDate, trCompanyCode, _
                                                                CompanyCode(l_datChargeDate), l_datChargeDate) Then
                l_strAICOldDepartment = l_dbRules.GetRuleReverseLookup(dbRules.TRSRuleCode.trAICDeptMapping, trDepartmentNumber, l_strEffectiveDepartment)
                If Trim(l_strAICOldDepartment) <> "" Then '//may not have been brought over from PeopleSoft yet
                    l_strEffectiveDepartment = l_strAICOldDepartment
                End If
            End If
            
            l_rsValidateInfoByChargeDate!DEPT_NUM = l_strEffectiveDepartment
            
'            l_rsValidateInfoByChargeDate!CMPY_CD = l_obj.GetCompanyCode(l_strEffectiveDepartment)
'            l_rsValidateInfoByChargeDate!ACCT_TYPE = l_obj.GetAccountType(l_strEffectiveDepartment)
''            l_rsValidateInfoByChargeDate!INTERFACE_TO_CD = l_obj.InterfaceCode(l_strEffectiveDepartment)
'            l_rsValidateInfoByChargeDate!ACCT_DIV_NUM = l_obj.GetAccountDivision(l_strEffectiveDepartment)
            l_rsValidateInfoByChargeDate.Update
        Next l_intIndex
        l_rsEmployees.MoveNext
    Wend
    
'Get Department Information
'Being smart as to only hit the cache once per department

    l_strDepartments = GetValidateDistIdInfoDistinctDepartments(l_rsValidateInfoByChargeDate)
    GetValidateDistIdInfoDepartmentInfo l_rsValidateInfoByChargeDate, l_strDepartments
    

    
    l_rsValidateInfoByChargeDate.Sort = "EMP_NUM, CHARGE_DATE"
    l_rsValidateInfoByChargeDate.MoveFirst
    Set GetValidateDistIdInfoByChargeDate = l_rsValidateInfoByChargeDate
    Set l_rsValidateInfoByChargeDate = Nothing
    Set l_rsEmployees = Nothing
    Set l_dbEffectiveDateRules = Nothing
    Set l_dbRules = Nothing
    
    Exit Function
eh:
    Set l_rsValidateInfoByChargeDate = Nothing
    Set l_rsEmployees = Nothing
    Set l_dbEffectiveDateRules = Nothing
    Set l_dbRules = Nothing
    TrsRaiseError m_strCMODULENAME, "GetValidateDistIdInfoByChargeDate", , , , , True

End Function

Private Function GetValidateDistIdInfoDistinctDepartments( _
            ByVal v_rsValidateInfoByChargeDate As ADODB.Recordset _
            ) As Variant

    Dim l_strDepartments() As String
    Dim l_intNumDepts As Integer
    Dim l_strDept As String
    
    v_rsValidateInfoByChargeDate.Sort = "DEPT_NUM"
    v_rsValidateInfoByChargeDate.MoveFirst
    ReDim l_strDepartments(0) As String
    l_strDepartments(0) = v_rsValidateInfoByChargeDate!DEPT_NUM
    l_intNumDepts = 1
    v_rsValidateInfoByChargeDate.MoveNext
    While Not v_rsValidateInfoByChargeDate.EOF
        l_strDept = v_rsValidateInfoByChargeDate!DEPT_NUM
        If IsStringInList(l_strDept, l_strDepartments) Then
        Else
            ReDim Preserve l_strDepartments(l_intNumDepts) As String
            l_strDepartments(l_intNumDepts) = l_strDept
            l_intNumDepts = l_intNumDepts + 1
        End If
        v_rsValidateInfoByChargeDate.MoveNext
    Wend

    GetValidateDistIdInfoDistinctDepartments = l_strDepartments

End Function


Private Sub GetValidateDistIdInfoDepartmentInfo( _
                ByRef r_rsValidateInfoByChargeDate As ADODB.Recordset, _
                ByVal v_strDepartments As Variant)
                
    Dim l_lngStartDept As Long
    Dim l_lngEndDept As Long
    Dim l_lngI As Long
    Dim l_strDept As String
    Dim l_obj As dbDepartment.cdDepartment
    Dim l_strCompanyCode As String
    Dim l_strAcctType As String
    Dim l_strAcctDivNum As String
    
    Set l_obj = CreateObject("dbDepartment.cdDepartment")
    l_lngStartDept = LBound(v_strDepartments)
    l_lngEndDept = UBound(v_strDepartments)
    For l_lngI = l_lngStartDept To l_lngEndDept
        l_strDept = v_strDepartments(l_lngI)
        l_strCompanyCode = l_obj.GetCompanyCode(l_strDept)
        l_strAcctType = l_obj.GetAccountType(l_strDept)
        l_strAcctDivNum = l_obj.GetAccountDivision(l_strDept)
        
        r_rsValidateInfoByChargeDate.Filter = "DEPT_NUM = '" & l_strDept & "'"
        While Not r_rsValidateInfoByChargeDate.EOF
            r_rsValidateInfoByChargeDate!CMPY_CD = l_strCompanyCode
            r_rsValidateInfoByChargeDate!ACCT_TYPE = l_strAcctType
            r_rsValidateInfoByChargeDate!ACCT_DIV_NUM = l_strAcctDivNum
            r_rsValidateInfoByChargeDate.MoveNext
        Wend
    
    Next l_lngI
    r_rsValidateInfoByChargeDate.Filter = ""

    Set l_obj = Nothing

End Sub

Private Function GetPayrollGroupForDept( _
            ByVal v_strDept As String _
            ) As String
            
    On Error GoTo eh
            
    Dim l_obj As dbDepartment.cdDepartment
 
    Set l_obj = CreateObject("dbDepartment.cdDepartment")
    
    GetPayrollGroupForDept = l_obj.GetPayrollGroup(v_strDept)
    
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPayrollGroupForDept"
            
            
End Function

Public Function IsAmerenCILCO(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean

'RULE0012
On Error GoTo eh

    IsAmerenCILCO = IsAmerenCIPSorCILCO(dbRules.TRSRuleCode.trIsAmerenCILCO, v_datChgDate)

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCILCO", , , , , True
End Function

Private Function IsAmerenCIPSorCILCO(ByVal v_lngRuleNum As Long, _
                            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
    Dim l_obj As busDepartment.cbsDepartment
    Dim l_strCompany As String
    Dim l_strDept As String
    Dim l_strUnionCode As String
    Dim l_datChargeDate As Date
    Dim l_strPayrollGroupCd As String
On Error GoTo eh

    Set l_obj = CreateObject("busDepartment.cbsDepartment")
    l_datChargeDate = EffectiveDate(v_datChgDate)
    l_strDept = DepartmentNumber(l_datChargeDate)
    l_strCompany = CompanyCode(l_datChargeDate)
    l_strUnionCode = UnionCode(l_datChargeDate)
    l_strPayrollGroupCd = PayrollGroupCode()

    IsAmerenCIPSorCILCO = l_obj.IsAmerenCIPSorCILCO(v_lngRuleNum, l_strDept, l_strCompany, l_strUnionCode, l_datChargeDate, l_strPayrollGroupCd)

    Set l_obj = Nothing
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPSorCILCO", , , , , True
End Function

Public Function IsEdwardsOperations(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean

' CILCO
'RULE0013
'RULE0014
    
    Dim l_obj As busDepartment.cbsDepartment
    Dim l_strDept As String

    On Error GoTo eh
    
    l_strDept = DepartmentNumber(EffectiveDate(v_datChgDate))
        
    Set l_obj = CreateObject("busDepartment.cbsDepartment")
    If l_obj.IsEdwardsOperations(l_strDept) Then 'RULE0013
        Set l_obj = Nothing
        IsEdwardsOperations = True
        Exit Function
    End If
    Set l_obj = Nothing
    
    IsEdwardsOperations = DoesRuleApply(14, trEmployeeNumber, EmployeeNumber)
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsEdwardsOperations", , , , , True

End Function

Public Function IsAmerenCILCODispatch(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean

' CILCO
'RULE0015
'RULE0016
    
    Dim l_obj As busDepartment.cbsDepartment
    Dim l_strDept As String

    On Error GoTo eh
    
    l_strDept = DepartmentNumber(EffectiveDate(v_datChgDate))
        
    Set l_obj = CreateObject("busDepartment.cbsDepartment")
    If l_obj.IsAmerenCILCODispatch(l_strDept) Then 'RULE0015
        Set l_obj = Nothing
        IsAmerenCILCODispatch = True
        Exit Function
    End If
    Set l_obj = Nothing
    
    IsAmerenCILCODispatch = DoesRuleApply(16, trEmployeeNumber, EmployeeNumber)
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAmerenCILCODispatch", , , , , True

End Function

Public Function IsAmerenCILCOLocal8Employee( _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As String
            
'RULE0019

    Dim l_datChgDate As Date
    Dim l_strUnion As String

On Error GoTo eh

    l_strUnion = UnionCode(EffectiveDate(v_datChgDate))
    
    IsAmerenCILCOLocal8Employee = DoesRuleApply(19, trUnionCode, l_strUnion)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCILCOLocal8Employee", , , , , True
        
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

Public Property Get OTExclusn(Optional _
        ByVal v_strChargeDate As Date = mstrDefaultDate) As String

' The OT Exclusion is 100% directly related to the Job
' Uses same logic to determine effective OT Exclusion as
' the logic to determine effective job number
' This will be used instead of the OTType
' 0000 indicates employee gets paid for all overtime
' 9998 indicates an executive that gets paid for all ot
' 9999 indicates executive does not get paid for any ot

    Dim l_datChgDate As Date
    Dim l_strOTExclusn As String
    
    On Error GoTo eh

    l_datChgDate = EffectiveDate(v_strChargeDate)

    If l_datChgDate >= JobEffectiveDate Then
        l_strOTExclusn = mstrOTExclusn
    ElseIf l_datChgDate >= PriorJobEffDate Then
        l_strOTExclusn = mstrPriorOTExclusn
        If l_strOTExclusn = "" Then
            l_strOTExclusn = mstrOTExclusn
        End If
    Else
        l_strOTExclusn = mstrOTExclusn
    End If
    
    OTExclusn = l_strOTExclusn
    
    Exit Property

eh:
    TrsRaiseError m_strCMODULENAME, "OTExclusn", , , , , True

End Property

Public Function IsNonExemptExecutive(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' BLUE011, BLUE059
    On Error GoTo eh
    
    IsNonExemptExecutive = _
        OTExclusn(EffectiveDate(v_datChgDate)) = "0000"
        
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsNonExemptExecutive", , , , , True
End Function

Public Function IsAmerenUEBenefitParticipant(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' RED067
    On Error GoTo eh
       
    IsAmerenUEBenefitParticipant = _
        EmployeeBenefitPackage(EffectiveDate(v_datChgDate)) = "UE"
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenUEBenefitParticipant", , , , , True
End Function

Public Function IsAmerenIPBenefitParticipant(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' RED067
    On Error GoTo eh
       
    IsAmerenIPBenefitParticipant = _
        EmployeeBenefitPackage(EffectiveDate(v_datChgDate)) = "IPC"
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenIPBenefitParticipant", , , , , True
End Function

Public Function IsEligibleJobForSundayPremium(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean

' Some jobs do not get Sunday Premium Time (SUS) because of agreements
'This method should no longer be used, but will call the new method
'Delete when compatability is finally broken

'    Dim l_strJobNum As String
    
    On Error GoTo eh

'    l_strJobNum = JobNumber(EffectiveDate(v_datChgDate))

    IsEligibleJobForSundayPremium = IsJobEligibleForSundayPremium(v_datChgDate)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsEligibleJobForSundayPremium", , , , , True

End Function

Public Function IsCoffeen12HourAgreement(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' Employees under the 12 hour agreement
'   get the shift component C12 for all time worked between 19:00 and 07:00
'   do not get Sunday Premium Time (SUS)
'This method should no longer be used, but will call the new method
'Delete when compatability is finally broken


'    Dim l_strJobNum As String
    
    On Error GoTo eh
    
'    l_strJobNum = JobNumber(EffectiveDate(v_datChgDate))

    IsCoffeen12HourAgreement = IsCoffeen12HourAgreementJob(v_datChgDate)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsCoffeen12HourAgreement", , , , , True

End Function

Public Function IsHutsonville12HourAgreement(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' Employees under the 12 hour agreement
' This employees will work their straight time shift from 05:00 to 17:00
'   get the shift component O for all Straight time worked between 05:00 and 07:00
'   get the shift component P for all Straight time worked between 15:00 and 17:00
'   do not get Sunday Premium Time (SUS)
'This method should no longer be used, but will call the new method
'Delete when compatability is finally broken


'    Dim l_strJobNum As String
    
    On Error GoTo eh
    
'    l_strJobNum = JobNumber(EffectiveDate(v_datChgDate))

    IsHutsonville12HourAgreement = IsHutsonville12HourAgreementJob(v_datChgDate)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsHutsonville12HourAgreement", , , , , True

End Function

Public Function IsHutsonvilleLocal148(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' Employees under the 12 hour agreement
' These employees will work their straight time shift from 05:00 to 17:00
'   get the shift component O for all Straight time worked between 05:00 and 07:00
'   get the shift component P for all Straight time worked between 15:00 and 17:00
'   do not get Sunday Premium Time (SUS)

    Dim l_strUnionCd As String
    
    On Error GoTo eh
    
    l_strUnionCd = UnionCode(EffectiveDate(v_datChgDate))
        
    If l_strUnionCd = "MA" Then
        IsHutsonvilleLocal148 = True
    Else
        IsHutsonvilleLocal148 = False
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsHutsonvilleLocal148", , , , , True

End Function

Public Function IsCILCOLocal8MaterialHandler(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' These employees will work their straight time shift from 07:00 to 15:00
'   get the shift component O for all Straight time worked between 23:00 and 07:00
'   get the shift component P for all Straight time worked between 15:00 and 23:00


    Dim l_strJobNum As String
    
    On Error GoTo eh
    
    l_strJobNum = JobNumber(EffectiveDate(v_datChgDate))

    IsCILCOLocal8MaterialHandler = DoesRuleApply(36, trJobNumber, l_strJobNum)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsCILCOLocal8MaterialHandler", , , , , True

End Function


Public Function IsJobEligibleForSundayPremium(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate, _
        Optional ByVal v_strJob As String = "") As Boolean

' Some jobs do not get Sunday Premium Time (SUS) because of agreements

    Dim l_strJobNum As String
    
    On Error GoTo eh
    
    If Trim(v_strJob) = "" Then
        l_strJobNum = JobNumber(EffectiveDate(v_datChgDate))
    Else
        l_strJobNum = v_strJob
    End If

    IsJobEligibleForSundayPremium = Not DoesRuleApply(154, trJobNumber, l_strJobNum)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsEligibleJobForSundayPremium", , , , , True

End Function

Public Function IsCoffeen12HourAgreementJob(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate, _
        Optional ByVal v_strJob As String = "") As Boolean
' Employees under the 12 hour agreement
'   get the shift component C12 for all time worked between 19:00 and 07:00
'   do not get Sunday Premium Time (SUS)

    Dim l_strJobNum As String
    
    On Error GoTo eh
    
    If Trim(v_strJob) = "" Then
        l_strJobNum = JobNumber(EffectiveDate(v_datChgDate))
    Else
        l_strJobNum = v_strJob
    End If

    IsCoffeen12HourAgreementJob = DoesRuleApply(155, trJobNumber, l_strJobNum)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsCoffeen12HourAgreementJob", , , , , True

End Function

Public Function IsHutsonville12HourAgreementJob(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate, _
        Optional ByVal v_strJob As String = "") As Boolean
' Employees under the 12 hour agreement
' This employees will work their straight time shift from 05:00 to 17:00
'   get the shift component O for all Straight time worked between 05:00 and 07:00
'   get the shift component P for all Straight time worked between 15:00 and 17:00
'   do not get Sunday Premium Time (SUS)

    Dim l_strJobNum As String
    
    On Error GoTo eh
    
    If Trim(v_strJob) = "" Then
        l_strJobNum = JobNumber(EffectiveDate(v_datChgDate))
    Else
        l_strJobNum = v_strJob
    End If

    IsHutsonville12HourAgreementJob = DoesRuleApply(160, trJobNumber, l_strJobNum)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsHutsonville12HourAgreementJob", , , , , True

End Function

Public Function IsHutsonvilleMaint(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
' Employees in the Hutsonville Maintenance department
' These employees will work their straight time shift from 07:00 to 15:30
'   get NO shift component for all Overtime time worked between 07:00 and 15:30
'   get the shift component P for all Overtime time worked between 15:30 and 23:00
'   get the shift component O for all Overtime time worked between 23:00 and 07:00

    Dim l_strDept As String
    
    On Error GoTo eh
    
    l_strDept = DepartmentNumber(EffectiveDate(v_datChgDate))
        
    IsHutsonvilleMaint = DoesRuleApply(161, trDepartmentNumber, l_strDept)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsHutsonvilleMaint", , , , , True

End Function


Public Function IsHutsonvilleMaintJob(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate, _
        Optional ByVal v_strJob As String = "") As Boolean
' Employees in the Hutsonville Maintenance department
' These employees will work their straight time shift from 07:00 to 15:30
'   get NO shift component for all Overtime time worked between 07:00 and 15:30
'   get the shift component P for all Overtime time worked between 15:30 and 23:00
'   get the shift component O for all Overtime time worked between 23:00 and 07:00
'Some Operations employees follow the rules of Maintenance


    Dim l_strJobNum As String
    
    On Error GoTo eh
    
    If Trim(v_strJob) = "" Then
        l_strJobNum = JobNumber(EffectiveDate(v_datChgDate))
    Else
        l_strJobNum = v_strJob
    End If
        
    IsHutsonvilleMaintJob = DoesRuleApply(162, trJobNumber, l_strJobNum)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsHutsonvilleMaintJob", , , , , True

End Function

Public Function IsCallawaySecurity(Optional _
        ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean

' Callaway Security Personal need to be indentified for the system assigns the shift component based on time worked
' Called from cboSchedule in busTimecard
    
    Dim l_obj As busDepartment.cbsDepartment
    Dim l_strDept As String

    On Error GoTo eh
    
    l_strDept = DepartmentNumber(EffectiveDate(v_datChgDate))
    
    IsCallawaySecurity = DoesRuleApply(212, trDepartmentNumber, l_strDept)
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsCallawaySecurity", , , , , True

End Function

