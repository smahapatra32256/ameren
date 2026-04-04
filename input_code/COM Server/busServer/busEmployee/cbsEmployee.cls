VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cbsEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'COMMENTS: OCT, 2005 - CHANGE CREATEINSTANCE
'TO CREATEOBJECT -- PREM

Private Const m_strCMODULENAME As String = "cbsEmployee"
Private Const mstrDefaultDate As String = "01/01/0100"
Private Const mstrNotAvailable As String = "NA"
Private Const mstrNotAvailable2 As String = "Z#"

Public Function GetDepartment(ByVal v_lngEmpNum As Long, ByVal v_datDate As Date) As String
    
'Jimb 04/21/2003
'This seems to be dead code
'I can find any other code that calls it
'The dbEmployee code for this is dead also
    
    Dim l_objEmployee As dbEmployee.cdEmployee
    
    On Error GoTo GetDepartment_EH
    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    GetDepartment = l_objEmployee.GetDepartment(v_lngEmpNum, v_datDate)

    Set l_objEmployee = Nothing
    Exit Function
    
GetDepartment_EH:
    Set l_objEmployee = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDepartment", , , , , True

End Function
Public Function GetWorkReasons(ByVal v_lngEmpNum As Long) As ADODB.Recordset
TrsRaiseError m_strCMODULENAME, "GetWorkReasons", trErrObsoleteMethod, , , , True
'    Dim l_objEmployee As dbEmployee.cdEmployee
'
'    On Error GoTo GetWorkReasons_EH
'
'    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
'    Set GetWorkReasons = l_objEmployee.GetWorkReasons(v_lngEmpNum)
'
'    Set l_objEmployee = Nothing
'Exit Function
'
'GetWorkReasons_EH:
'    Set l_objEmployee = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetWorkReasons", , , , , True

End Function
Public Function IsValidEmployee(ByVal v_lngEmpNum As Long) As Boolean
'  Is the employee a valid employee
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo IsValidEmployee_EH

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    IsValidEmployee = l_objEmployee.IsValidEmployee(v_lngEmpNum)
    Set l_objEmployee = Nothing
Exit Function

IsValidEmployee_EH:
    IsValidEmployee = False
    Set l_objEmployee = Nothing
    
    TrsRaiseError m_strCMODULENAME, "IsValidEmployee", , , , , True

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
'       (v_strDeptNum = "360") Or _
'       (v_strDeptNum = "18U") Or _
'       (v_strDeptNum = "18W") Then
'        IsSubstationOrTransmissionEmployee = True
'    Else
'        IsSubstationOrTransmissionEmployee = False
'    End If
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsSubstationOrTransmissionEmployee", , , , , True
End Function

Private Function IsEmployeeDepartmentDOJM(ByVal v_strDeptNum As String, ByVal v_objrsDepartment As ADODB.Recordset) As Boolean
'  BLUE037
'  IsEmployeeDepartmentDOJM
    Dim lcolErrDescArgs As New Collection
    
    On Error GoTo IsEmployeeDepartmentDOJM_EH

'    v_objrsDepartment.Filter = "DEPT_NUM = '" & v_strDeptNum & "' "
'    If v_objrsDepartment.BOF And v_objrsDepartment.EOF Then
'        lcolErrDescArgs.Add v_strDeptNum
'        Err.Raise trErrbusDeptNotInCollection
'    End If
    '//If v_objrsDepartment!INTERFACE_TO_CD = "D" Then
    If DoesRuleApply(1001, trDepartmentNumber, v_strDeptNum) Then
        IsEmployeeDepartmentDOJM = True
    Else
        IsEmployeeDepartmentDOJM = False
    End If
    
'    v_objrsDepartment.Filter = ""
    Exit Function

IsEmployeeDepartmentDOJM_EH:
    Select Case Err.Number
        Case trErrbusDeptNotInCollection
            TrsRaiseError m_strCMODULENAME, "IsEmployeeDepartmentDOJM", Err.Number, , lcolErrDescArgs
        Case Else
            TrsRaiseError m_strCMODULENAME, "IsEmployeeDepartmentDOJM"
    End Select
    
End Function
Public Function GetEmployeesInDepartment( _
            ByVal v_aryDepartments As Variant, _
            Optional v_bolActiveEmployeesOnly As Boolean = False, _
            Optional v_datStartDate As Date = "0100-01-01", _
            Optional v_datEndDate As Date = "0100-01-01" _
            ) As ADODB.Recordset
            
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo GetEmployeesInDepartment_EH
    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeesInDepartment = l_objEmployee.GetEmployeesInDepartment( _
                v_aryDepartments _
                , v_bolActiveEmployeesOnly _
                , v_datStartDate _
                , v_datEndDate)

    Set l_objEmployee = Nothing
Exit Function
GetEmployeesInDepartment_EH:

    Set l_objEmployee = Nothing

    TrsRaiseError m_strCMODULENAME, "GetEmployeesInDepartment", , , , , True
            
End Function
Public Function GetPrimaryWorkGroups( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional bln_PrimaryOnly As Boolean = False _
            ) As ADODB.Recordset
    
    Dim l_objWorkGroup As dbWorkGroup.cdWorkGroup

    On Error GoTo eh

    Set l_objWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")    'MTS standard
    Set GetPrimaryWorkGroups = l_objWorkGroup.GetEmployeeWorkGroups(v_aryEmployees, v_datStartDate, v_datEndDate, bln_PrimaryOnly)
    Set l_objWorkGroup = Nothing

Exit Function

eh:
    Set GetPrimaryWorkGroups = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetPrimaryWorkGroups", , , , , True
    
End Function
Public Function GetAllEmployeeInfo( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            ByVal v_blnPrimaryOnly As Boolean, _
            ByRef r_rsEmployee As ADODB.Recordset, _
            ByRef r_rsPrimaryWorkGroups As ADODB.Recordset, _
            ByRef r_rsTimeEntryLockStatus As ADODB.Recordset, _
            ByRef r_rsCaseNumbers As ADODB.Recordset, _
            ByRef r_rsEmpBanks As ADODB.Recordset, _
            ByRef r_rsEmpOTBanks As ADODB.Recordset, _
            ByRef r_rsQuarterlyBanks As ADODB.Recordset _
            ) As Boolean
            
    Dim l_bolGetEmpBanks As Boolean
    Dim l_datToday As Date
    Dim l_aryEmployees As Variant

On Error GoTo eh
    
    If IsArray(v_aryEmployees) Then
        l_aryEmployees = v_aryEmployees
    Else
        l_aryEmployees = Array(str(CLng(v_aryEmployees)))
    End If
    l_datToday = Date
    Set r_rsEmployee = GetEmployee(l_aryEmployees)
    Set r_rsPrimaryWorkGroups = GetPrimaryWorkGroups(l_aryEmployees, v_datStartDate, v_datEndDate, v_blnPrimaryOnly)
    Set r_rsTimeEntryLockStatus = GetTimeEntryLockStatus(v_datStartDate, v_datEndDate, r_rsEmployee)
    'Set r_rsCaseNumbers = GetCaseGroupsAndNumbers(l_aryEmployees, v_datStartDate, v_datEndDate)
    Set r_rsCaseNumbers = Nothing '//this is intentional as it is not used anymore
    
    l_bolGetEmpBanks = GetEmployeeBanks(l_aryEmployees, l_datToday, r_rsEmpBanks, r_rsEmpOTBanks, Nothing)
    Set r_rsQuarterlyBanks = Nothing '//this is intentional as it is not used anymore

    GetAllEmployeeInfo = True
    Exit Function
    
eh:
    Dim l_colErrDesc As Collection
    Dim l_lngLow As Long
    Dim l_lngHigh As Long
    Dim l_lngI As Long
    Dim l_lngErrNum As Long
    
    l_lngErrNum = Err.Number
    
    Set l_colErrDesc = New Collection
    l_colErrDesc.Add "v_datStartDate =" & v_datStartDate
    l_colErrDesc.Add "v_datEndDate =" & v_datEndDate
    l_colErrDesc.Add "Begin Employees in Request"
    If IsArray(v_aryEmployees) Then
        l_lngLow = LBound(v_aryEmployees)
        l_lngHigh = UBound(v_aryEmployees)
        For l_lngI = l_lngLow To l_lngHigh
            l_colErrDesc.Add "Emp Num =" & v_aryEmployees(l_lngI)
        Next l_lngI
    Else
        l_colErrDesc.Add "Emp Num =" & v_aryEmployees
    End If
    l_colErrDesc.Add "End Employees in Request"

    Set r_rsEmployee = Nothing
    Set r_rsPrimaryWorkGroups = Nothing
    Set r_rsTimeEntryLockStatus = Nothing
    Set r_rsCaseNumbers = Nothing
    Set r_rsEmpBanks = Nothing
    Set r_rsEmpOTBanks = Nothing
    Set r_rsQuarterlyBanks = Nothing
    GetAllEmployeeInfo = False
    TrsRaiseError m_strCMODULENAME, "GetAllEmployeeInfo", l_lngErrNum, , , l_colErrDesc, True

End Function
Public Function GetEmployee( _
        ByVal v_aryEmployees As Variant _
        ) As ADODB.Recordset

    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo eh

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployee = l_objEmployee.GetEmployee(v_aryEmployees)
    Set l_objEmployee = Nothing
    
Exit Function
eh:
    Set GetEmployee = Nothing

    TrsRaiseError m_strCMODULENAME, "GetEmployee", , , , , True

End Function
Private Function GetTimeEntryLockStatus( _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            ByVal v_rsEmployeeData As ADODB.Recordset _
            ) As ADODB.Recordset
            
            
'JimB 04/15/2003
'This method was re-written for readablity and to remove reference to boCommonLib

            

'  Loop through the dates between the start and end date and
'  determine the department and lock status.
' BLUE040
' BLUE041
' BLUE037

    Dim l_strDeptNum As String
    Dim l_strPrevDeptNum As String
    Dim l_strEnumDeptNum As String
    Dim l_datDeptEffDate As Date
    Dim l_strPayrollGroupCode As String
    Dim l_strPrevPayrollGroupCode As String
    Dim l_strPayProcInd As String
    Dim l_strLockStatus As String
    Dim l_strDepartmentLockStatus As String
    Dim l_datDate As Date
    Dim l_objdbDepartment As dbDepartment.cdDepartment   'BLUE037
    Dim l_rsGetDepartments As ADODB.Recordset            'BLUE037
    Dim l_colDept As New Collection                      'BLUE037
    Dim l_bolIsSubstationOrTransmission As Boolean       'BLUE037
    Dim l_bolIsDOJM As Boolean                           'BLUE037
    Dim l_strUnionCd As String                           'WHITE003
    Dim l_objdbCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    Dim l_rsTimeEntryLockStatus As ADODB.Recordset
    Dim l_intAllowTimeIncr As Integer
    Dim l_rsCollPeriods As ADODB.Recordset

    Dim l_lngEmpNum As Long
    Dim l_lngI As Long
    Dim l_colPayrollGroupCodes As Collection
    Dim l_strDeptPayrollGroupCode As String
    Dim l_strDeptLockStatus As String
    Dim l_strCollPeriodLockStatus As String
    Dim l_strFilter As String
    Dim l_strPriorUnionCd As String
    Dim l_strEffUnionCd As String
    Dim l_datJobEffDate As Date
    Dim l_blnIsExecutive As Boolean
    Dim l_strExecContractInd As String
    Dim l_strPriorExecContractInd As String
    Dim l_strCompany As String
    Dim l_strPriorCompany As String
    Dim l_strEffCompany As String


On Error GoTo eh

    Set l_objdbCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    Set l_objdbDepartment = CreateObject("dbDepartment.cdDepartment")
    Set l_rsTimeEntryLockStatus = New ADODB.Recordset

    '  Create an empty recordset
    With l_rsTimeEntryLockStatus
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        .Fields.Append "DEPT_NUM", adChar, 3, adFldUpdatable
        .Fields.Append "LOCK_STATUS", adChar, 1, adFldUpdatable
        .Fields.Append "ALLOW_TIME_INCR", adInteger, , adFldUpdatable
        .Fields.Append "SUB_OR_TRANS_FLAG", adBoolean, , adFldUpdatable
        .Fields.Append "DOJM_FLAG", adBoolean, adFldUpdatable
        .Open
    End With
    If Not RecordsExist(v_rsEmployeeData) Then
        Set GetTimeEntryLockStatus = l_rsTimeEntryLockStatus
        Exit Function
    Else
        v_rsEmployeeData.MoveFirst
    End If

        ' Get Lock Status Recordset
    Set l_colPayrollGroupCodes = New Collection
    Set l_rsCollPeriods = l_objdbCollectionPeriod.GetCollectionPeriods(v_datStartDate, v_datEndDate, l_colPayrollGroupCodes)
    Set l_objdbCollectionPeriod = Nothing
    Set l_colDept = GetDistinctDepartments(v_rsEmployeeData)
    Set l_rsGetDepartments = l_objdbDepartment.GetDepartments(l_colDept) 'BLUE037
    
    
    While Not v_rsEmployeeData.EOF
        l_lngEmpNum = v_rsEmployeeData!EMP_NUM
        l_strDeptNum = Trim(v_rsEmployeeData!DEPT_NUM)
        l_strPrevDeptNum = Trim(v_rsEmployeeData!PREV_DEPT_NUM)
        l_datDeptEffDate = v_rsEmployeeData!DEPT_EFF_DATE
        l_strUnionCd = Trim(v_rsEmployeeData!UNION_CD)
        l_strPriorUnionCd = Trim(v_rsEmployeeData!PRIOR_UNION_CD)
        l_datJobEffDate = v_rsEmployeeData!JOB_EFF_DATE
        l_strExecContractInd = v_rsEmployeeData!EXEC_CONTRACT_IND
        l_strPriorExecContractInd = v_rsEmployeeData!PRIOR_EXEC_CON_IND
        l_strCompany = v_rsEmployeeData!CMPY_CD
        l_strPriorCompany = v_rsEmployeeData!PREV_CMPY_CD
    
        l_datDate = v_datStartDate
        Do Until l_datDate > v_datEndDate
            l_strEffUnionCd = EmployeeUnion( _
                l_datJobEffDate, l_strUnionCd, l_strPriorUnionCd, l_datDate)
            l_blnIsExecutive = IsExecutive(l_datJobEffDate, l_strUnionCd, l_strPriorUnionCd, _
                l_strExecContractInd, l_strPriorExecContractInd, l_datDate)
            l_strEnumDeptNum = EffectiveDepartment( _
                l_strDeptNum, l_datDeptEffDate, l_strPrevDeptNum, l_datDate)
            l_strEffCompany = CompanyCode(l_datDate, l_strDeptNum, l_strPrevDeptNum, l_datDeptEffDate, _
                l_strCompany, l_strPriorCompany)
            l_intAllowTimeIncr = PartialHourMinutes(l_strEnumDeptNum, l_blnIsExecutive, l_strEffCompany)
            l_bolIsSubstationOrTransmission = IsSubstationOrTransmissionEmployee(l_strEffUnionCd, l_strEnumDeptNum)
            l_rsGetDepartments.Filter = "DEPT_NUM = '" & l_strEnumDeptNum & "'"
            If RecordsExist(l_rsGetDepartments) Then
                l_bolIsDOJM = IsEmployeeDepartmentDOJM(l_strEnumDeptNum, l_rsGetDepartments)
                l_strDeptPayrollGroupCode = l_rsGetDepartments!PAYROLL_GROUP_CD
                l_strDeptLockStatus = l_rsGetDepartments!LOCK_STATUS
                l_strFilter = "PAYROLL_GROUP_CD = '" & l_strDeptPayrollGroupCode & "'" _
                 & "  AND PAY_PER_START_DATE <= '" & Format(l_datDate, "MM/DD/YYYY") & "' " _
                 & "  AND PAY_PER_END_DATE >= '" & Format(l_datDate, "MM/DD/YYYY") & "' "
                l_strFilter = MDAC25FilterFix(l_strFilter)
                l_rsCollPeriods.Filter = l_strFilter
                    'raise an error here if there are no collection periods
                If Not RecordsExist(l_rsCollPeriods) Then
                    Err.Raise trErrbusNoCollectionPeriods
                End If
                l_strCollPeriodLockStatus = l_rsCollPeriods!PAY_PROCESSED_IND
            Else
                l_bolIsDOJM = False
                l_strDeptPayrollGroupCode = "N"
                l_strDeptLockStatus = "N"
                l_strCollPeriodLockStatus = "N"
            End If
            l_strLockStatus = DetermineLockStatus(l_strCollPeriodLockStatus, l_strDeptLockStatus)
            With l_rsTimeEntryLockStatus
                .AddNew
                .Fields("EMP_NUM") = l_lngEmpNum
                .Fields("CHARGE_DATE") = l_datDate
                .Fields("DEPT_NUM") = l_strEnumDeptNum
                .Fields("LOCK_STATUS") = l_strLockStatus
                .Fields("ALLOW_TIME_INCR") = l_intAllowTimeIncr
                .Fields("SUB_OR_TRANS_FLAG") = l_bolIsSubstationOrTransmission
                .Fields("DOJM_FLAG") = l_bolIsDOJM
                .Update
            End With
            l_datDate = DateAdd("d", 1, l_datDate)
        Loop
        v_rsEmployeeData.MoveNext
    Wend
    v_rsEmployeeData.MoveFirst

    If RecordsExist(l_rsTimeEntryLockStatus) Then
        l_rsTimeEntryLockStatus.MoveFirst
    End If

    Set GetTimeEntryLockStatus = l_rsTimeEntryLockStatus

    Set l_rsTimeEntryLockStatus = Nothing
    Set l_rsCollPeriods = Nothing
    Set l_objdbDepartment = Nothing
    Set l_rsGetDepartments = Nothing
    Exit Function

eh:
    Set l_rsCollPeriods = Nothing
    Set l_rsTimeEntryLockStatus = Nothing
    Set l_objdbDepartment = Nothing
    Set l_rsGetDepartments = Nothing
    
    Dim l_colErrDesc As Collection
    Dim l_lngErrNum As Long
    
    l_lngErrNum = Err.Number
    
    Set l_colErrDesc = New Collection
    
'JIMB HERE

    l_colErrDesc.Add "Begin GetTimeEntryLockStatus Log"
    l_colErrDesc.Add "Look at the payroll group"
    l_colErrDesc.Add "If it equals N, contact payroll "
    l_colErrDesc.Add " and they will handle it in PeopleSoft."
    l_colErrDesc.Add "v_rsEmployeeData.Filter=" & v_rsEmployeeData.Filter
    If RecordsExist(v_rsEmployeeData) Then
        l_colErrDesc.Add "v_rsEmployeeData.Filter=" & v_rsEmployeeData.Filter
    End If
    l_colErrDesc.Add "l_lngEmpNum =" & str(l_lngEmpNum)
    l_colErrDesc.Add "l_strDeptNum =" & l_strDeptNum
    l_colErrDesc.Add "l_strPrevDeptNum =" & l_strPrevDeptNum
    If IsNull(l_datDate) Or IsEmpty(l_datDate) Then
    l_colErrDesc.Add "l_datDate=NOTHING"
    Else
        If IsDate(l_datDeptEffDate) Then
            l_colErrDesc.Add "l_datDate=" & l_datDate
        Else
            l_colErrDesc.Add "l_datDate=NOTHING"
        End If
    End If
    If IsNull(l_datDeptEffDate) Or IsEmpty(l_datDeptEffDate) Then
    l_colErrDesc.Add "l_datDeptEffDate=NOTHING"
    Else
        If IsDate(l_datDeptEffDate) Then
            l_colErrDesc.Add "l_datDeptEffDate=" & l_datDeptEffDate
        Else
            l_colErrDesc.Add "l_datDeptEffDate=NOTHING"
        End If
    End If
    l_colErrDesc.Add "l_strDeptPayrollGroupCode=" & l_strDeptPayrollGroupCode
    
    l_colErrDesc.Add "End GetTimeEntryLockStatus Log"
    
    TrsRaiseError m_strCMODULENAME, "GetTimeEntryLockStatus", l_lngErrNum, , , l_colErrDesc
    

End Function
Private Function GetDistinctDepartments( _
            ByVal v_rsEmployeeData As ADODB.Recordset _
            ) As Collection

    Dim l_colDepts As Collection
    Dim l_lngCount, l_lngI As Long

    Set l_colDepts = New Collection

    v_rsEmployeeData.Filter = ""
    v_rsEmployeeData.MoveFirst
    While Not v_rsEmployeeData.EOF
        AddtoDistinctCollection l_colDepts, v_rsEmployeeData!DEPT_NUM
        If Trim(v_rsEmployeeData!PREV_DEPT_NUM) <> "" Then
           AddtoDistinctCollection l_colDepts, v_rsEmployeeData!PREV_DEPT_NUM
        End If
        v_rsEmployeeData.MoveNext
    Wend
    v_rsEmployeeData.MoveFirst

    Set GetDistinctDepartments = l_colDepts

End Function
Private Sub AddtoDistinctCollection( _
                ByRef r_colData As Collection, _
                ByVal v_strData As String)
                
    Dim l_lngCount As Long
    Dim l_lngI As Long
    Dim l_strCompare As String
    Dim l_bolInList As Boolean
                
    l_bolInList = False
    l_lngCount = r_colData.Count
    l_lngI = 0
    While l_lngI < l_lngCount And Not l_bolInList
        l_lngI = l_lngI + 1
        If v_strData = r_colData.Item(l_lngI) Then
            l_bolInList = True
        End If
    Wend
    
    If Not l_bolInList Then
        r_colData.Add v_strData
    End If
    
                
End Sub
Public Function GetEmployeeBanks( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datCurrDate As Date, _
                    ByRef r_rsBanks As ADODB.Recordset, _
                    ByRef r_rsOTBanks As ADODB.Recordset, _
                    ByRef r_rsQuarterlyOTBanks As ADODB.Recordset _
                    ) As Boolean
                                                        
'  Get the Employee Banks (Vacation, FMLA, Holiday, Compensatory) and Quarterly Overtime Banks
'  Passing in the Current Year results in getting the Prior Year and Next Year data as well
                                                        
    Dim l_objdbAccrualBankMgr As dbAccrualBankMgr.cdAccrualBankMgr
    Dim l_objdbOTEqualizer As dbOvertimeEqualizer.cdOvertimeEqualizer
    
On Error GoTo eh
    
    Set l_objdbAccrualBankMgr = CreateObject("dbAccrualBankMgr.cdAccrualBankMgr")
    Set r_rsBanks = l_objdbAccrualBankMgr.GetEmployeeBanks(v_aryEmployees, v_datCurrDate)
    Set l_objdbAccrualBankMgr = Nothing
    
    Set l_objdbOTEqualizer = CreateObject("dbOvertimeEqualizer.cdOvertimeEqualizer")
'    Set r_rsQuarterlyOTBanks = l_objdbOTEqualizer.GetEmployeeQuarterlyOTBanks(v_aryEmployees, v_datCurrDate)
    Set r_rsQuarterlyOTBanks = Nothing '//this is intentional as it is not used anymore
    Set r_rsOTBanks = l_objdbOTEqualizer.GetEmployeeOvertimeBanks(v_aryEmployees)
    Set l_objdbOTEqualizer = Nothing
                                                        
    GetEmployeeBanks = True
    Exit Function

eh:
    Set r_rsBanks = Nothing
    Set r_rsOTBanks = Nothing
    Set r_rsQuarterlyOTBanks = Nothing
    Set l_objdbAccrualBankMgr = Nothing
    GetEmployeeBanks = False
    TrsRaiseError m_strCMODULENAME, "GetEmployeeBanks", , , , , True

End Function
Public Function GetAllEmployees( _
            Optional v_bolActiveEmployeesOnly As Boolean = True _
            ) As ADODB.Recordset
            
    Dim l_objEmployee As dbEmployee.cdEmployee
    
    On Error GoTo GetAllEmployees_EH
    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetAllEmployees = l_objEmployee.GetAllEmployees(v_bolActiveEmployeesOnly)
    
    Set l_objEmployee = Nothing
Exit Function
GetAllEmployees_EH:

    Set l_objEmployee = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllEmployees", , , , , True
            
End Function
Public Function GetBasicEmployeeInformation(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo GetBasicEmployeeInformation_EH

    'Set l_objEmployee = New dbEmployee.cdEmployee
    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetBasicEmployeeInformation = l_objEmployee.GetBasicEmployeeInformation(v_lngEmpNum)

    Set l_objEmployee = Nothing
    Exit Function
    
GetBasicEmployeeInformation_EH:
    Set GetBasicEmployeeInformation = Nothing
    TrsRaiseError m_strCMODULENAME, "GetBasicEmployeeInformation", , , , , True

End Function
Public Function GetEmployeePrimaryWorkGroup(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo GetEmployeePrimaryWorkGroup_EH

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeePrimaryWorkGroup = l_objEmployee.GetEmployeePrimaryWorkGroup(v_lngEmpNum)

    Set l_objEmployee = Nothing
    Exit Function
    
GetEmployeePrimaryWorkGroup_EH:
    Set GetEmployeePrimaryWorkGroup = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeePrimaryWorkGroup", , , , , True

End Function
Public Function GetEmployeeWorkGroups(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo GetEmployeeWorkGroups_EH

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeeWorkGroups = l_objEmployee.GetEmployeeWorkGroups(v_lngEmpNum)

    Set l_objEmployee = Nothing
    Exit Function
    
GetEmployeeWorkGroups_EH:
    Set GetEmployeeWorkGroups = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeWorkGroups", , , , , True

End Function
Public Function GetEmployeeReliefGroups(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo GetEmployeeReliefGroups_EH

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeeReliefGroups = l_objEmployee.GetEmployeeReliefGroups(v_lngEmpNum)

    Set l_objEmployee = Nothing
    Exit Function
    
GetEmployeeReliefGroups_EH:
    Set GetEmployeeReliefGroups = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeReliefGroups", , , , , True

End Function
Public Function GetEmployeeQualifications(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo GetEmployeeQualifications_EH

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeeQualifications = l_objEmployee.GetEmployeeQualifications(v_lngEmpNum)

    Set l_objEmployee = Nothing
    Exit Function
    
GetEmployeeQualifications_EH:
    Set GetEmployeeQualifications = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeQualifications", , , , , True

End Function
Public Function GetEmployeeSkills(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo GetEmployeeSkills_EH

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeeSkills = l_objEmployee.GetEmployeeSkills(v_lngEmpNum)

    Set l_objEmployee = Nothing
    Exit Function
    
GetEmployeeSkills_EH:
    Set GetEmployeeSkills = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeSkills", , , , , True

End Function
Public Function EmployeeBenefitPackage( _
            ByVal v_datJobEffDate As Date, _
            ByVal v_strManagementBenPkg As String, _
            ByVal v_strUnionCode As String, _
            ByVal v_strPriorUnionCode As String, _
            Optional ByVal v_strExecContractInd As String = mstrNotAvailable, _
            Optional ByVal v_strPriorExecContractInd As String = mstrNotAvailable, _
            Optional ByVal v_strUnionBenPkg As String = mstrNotAvailable, _
            Optional ByVal v_strPriorUnionBenPkg As String = mstrNotAvailable, _
            Optional ByVal v_strChgDate As Date = mstrDefaultDate) As String
' WHITE039
' v_datJobEffDate is the union effective date
    Dim l_datChgDate As Date
    Dim l_blnIsExecutive As Boolean

On Error GoTo eh

    l_datChgDate = EffectiveDate(v_strChgDate)
    l_blnIsExecutive = IsExecutivePrivate(l_datChgDate, v_strExecContractInd, v_strPriorExecContractInd, _
            v_datJobEffDate, v_strUnionCode, v_strPriorUnionCode)
    If l_blnIsExecutive Then
        EmployeeBenefitPackage = v_strManagementBenPkg
    ElseIf UseCurrentUnionAttribute(l_datChgDate, v_datJobEffDate, v_strPriorUnionCode) Then
        EmployeeBenefitPackage = UnionBenefitPackage(v_strUnionCode, v_strUnionBenPkg)
    Else
        EmployeeBenefitPackage = UnionBenefitPackage(v_strPriorUnionCode, v_strPriorUnionBenPkg)
    End If
    
    If l_blnIsExecutive And v_strManagementBenPkg = "CIPS" Then
        EmployeeBenefitPackage = "UE"
    End If
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeBenefitPackage", , , , , True
End Function
Private Function UnionBenefitPackage( _
            ByVal v_strUnionCode As String, _
            ByVal v_strUnionBenPkg As String) As String
    Dim l_cdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    
    If v_strUnionBenPkg = mstrNotAvailable Then
        Set l_cdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
        UnionBenefitPackage = _
            l_cdCodeTables.UnionBenefitPackageIndicator(v_strUnionCode)
        Set l_cdCodeTables = Nothing
    Else
        UnionBenefitPackage = v_strUnionBenPkg
    End If
    Exit Function
    
eh:
    Set l_cdCodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "UnionBenefitPackage"
End Function
Private Function IsExecutivePrivate( _
            ByVal v_datChgDate As Date, _
            ByVal v_strExecContractInd As String, _
            ByVal v_strPriorExecContractInd As String, _
            ByVal v_datUnionEffDate As Date, _
            ByVal v_strUnionCode As String, _
            ByVal v_strPriorUnionCode As String) As Boolean
    
On Error GoTo eh
    
    IsExecutivePrivate = _
        ExecContractIndicator(v_datChgDate, v_strExecContractInd, v_strPriorExecContractInd, _
                v_datUnionEffDate, v_strUnionCode, v_strPriorUnionCode) = "E"
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsExecutivePrivate"
End Function
Private Function ExecContractIndicator( _
            ByVal v_datChgDate As Date, _
            ByVal v_strExecContractInd As String, _
            ByVal v_strPriorExecContractInd As String, _
            ByVal v_datUnionEffDate As Date, _
            ByVal v_strUnionCode As String, _
            ByVal v_strPriorUnionCode As String) As String

On Error GoTo eh
        
    If UseCurrentUnionAttribute(v_datChgDate, v_datUnionEffDate, v_strPriorUnionCode) Then
        ExecContractIndicator = GetExecContractInd(v_strUnionCode, v_strExecContractInd)
    Else
        ExecContractIndicator = GetExecContractInd(v_strPriorUnionCode, v_strPriorExecContractInd)
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ExecContractIndicator"
End Function
Private Function GetExecContractInd( _
            ByVal v_strUnionCode As String, _
            ByVal v_strExecContractInd As String) As String
    Dim l_cdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    
    If v_strExecContractInd = mstrNotAvailable Then
        Set l_cdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
        GetExecContractInd = _
            l_cdCodeTables.UnionExecutiveContractIndicator(v_strUnionCode)
        Set l_cdCodeTables = Nothing
    Else
        GetExecContractInd = v_strExecContractInd
    End If
    Exit Function
    
eh:
    Set l_cdCodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetExecContractInd"
End Function
Private Function UseCurrentUnionAttribute( _
            ByVal v_datChgDate As Date, _
            ByVal v_datUnionEffDate As Date, _
            ByVal v_strPriorUnionCode As String) As Boolean

On Error GoTo eh
    If v_datChgDate >= v_datUnionEffDate Or _
            Trim(v_strPriorUnionCode) = "" Then
        UseCurrentUnionAttribute = True
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "UseCurrentUnionAttribute"
End Function
Private Function EffectiveDate(ByVal v_datChgDate As Date) As Date
'    If v_datChgDate is the default date, answer current date;
'    otherwise answer v_datChgDate
On Error GoTo eh
    If v_datChgDate = mstrDefaultDate Then
        EffectiveDate = Date
    Else
        EffectiveDate = v_datChgDate
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "EffectiveDate"
End Function
Public Function GetLockStatus( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datChargeDate As Date, _
            ByVal v_rsEmpLockData As ADODB.Recordset, _
            ByVal v_rsCollPeriods As ADODB.Recordset _
            ) As String
            
    Dim l_strLockStatus As String
    Dim l_strDept As String
    Dim l_strPayGroup As String
    Dim l_strDeptLockStatus As String
    Dim l_strPayProcessedInd As String
            
    On Error GoTo eh
            
    GetEffectiveDeptInfo v_lngEmpNum, v_datChargeDate, v_rsEmpLockData, _
        l_strDept, l_strPayGroup, l_strDeptLockStatus
    l_strPayProcessedInd = GetPayGroupProcessedInd(l_strPayGroup, v_datChargeDate, v_rsCollPeriods)
    l_strLockStatus = DetermineLockStatus(l_strPayProcessedInd, l_strDeptLockStatus)
    
    GetLockStatus = l_strLockStatus
    v_rsEmpLockData.Filter = ""
    v_rsCollPeriods.Filter = ""

    Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "GetLockStatus", , , , , True
            
End Function
Private Function DetermineLockStatus( _
    ByVal v_strPayProcessedInd As String, _
    ByVal v_strDeptLockStatus As String _
    ) As String
    
    Dim l_strLockStatus As String
    
    On Error GoTo eh
    
    Select Case v_strPayProcessedInd
        Case "I"                                        ' In Process
                l_strLockStatus = "Y"
        Case "L", "R"                              ' In Final Review, Reviewing
                '  Get the Department Lock Status
                Select Case v_strDeptLockStatus
                    Case "Y"
                        l_strLockStatus = "L"       ' Department Locked
                    Case "N"
                        l_strLockStatus = "N"       ' Department Not Locked
                End Select
        
        Case "N"
            l_strLockStatus = "N"                   ' Not Processed
        Case "Y"
            l_strLockStatus = "Y"                   ' Processed
    End Select
            
    DetermineLockStatus = l_strLockStatus
    
    Exit Function
    
eh:
   
    TrsRaiseError m_strCMODULENAME, "DetermineLockStatus"
    
End Function
Private Function GetPayGroupProcessedInd( _
    ByVal v_strPayGroup As String, _
    ByVal v_datChargeDate As Date, _
    ByVal v_rsCollPeriods As ADODB.Recordset _
    ) As String
    
    Dim l_strFilter As String
    Dim l_strPayProcessInd As String
    Dim l_datFirstStart As Date
    Dim l_datLastEnd As Date
    Dim l_blnPeriodFound As Boolean
    
    'v_rsCollPeriods has all the pay groups from the ttr_coll_period table
    'it does not have collection period information from closed payrolls
    'It only has collection information for the open payrolls
    'This is to limit the data that needs to be retrieved
    'All pay periods prior to the first open pay period are all closed.
    'All pay periods after the first open pay period are all open
    
    On Error GoTo eh
    
    l_strFilter = "PAYROLL_GROUP_CD = '" & v_strPayGroup & "'"
    v_rsCollPeriods.Filter = l_strFilter
    If Not RecordsExist(v_rsCollPeriods) Then
    'Should never happen -- This probably should be an error
        GetPayGroupProcessedInd = "N"
        Exit Function
    End If
    v_rsCollPeriods.Sort = "PAY_PER_START_DATE"
    v_rsCollPeriods.MoveFirst
    l_datFirstStart = v_rsCollPeriods!PAY_PER_START_DATE
    v_rsCollPeriods.MoveLast
    l_datLastEnd = v_rsCollPeriods!PAY_PER_END_DATE
    
    If v_datChargeDate < l_datFirstStart Then
        GetPayGroupProcessedInd = "Y"
        Exit Function
    End If
    
    If v_datChargeDate > l_datLastEnd Then
        GetPayGroupProcessedInd = "N"
        Exit Function
    End If
    
    v_rsCollPeriods.MoveFirst
    l_blnPeriodFound = False
    
    While Not v_rsCollPeriods.EOF And Not l_blnPeriodFound
        If v_datChargeDate >= v_rsCollPeriods!PAY_PER_START_DATE _
            And v_datChargeDate <= v_rsCollPeriods!PAY_PER_END_DATE Then
            l_blnPeriodFound = True
        Else
            v_rsCollPeriods.MoveNext
        End If
    Wend
    
    If Not v_rsCollPeriods.EOF Then
        GetPayGroupProcessedInd = v_rsCollPeriods!PAY_PROCESSED_IND
    Else
    'Should never happen -- This probably should be an error
        GetPayGroupProcessedInd = "N"
    End If
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetPayGroupProcessedInd"
    
End Function
Private Sub GetEffectiveDeptInfo( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date, _
    ByVal v_rsEmpLockData As ADODB.Recordset, _
    ByRef r_strDept As String, _
    ByRef r_strPayGroup As String, _
    ByRef r_strDeptLockStatus As String)
    
    Dim l_strDept As String
    Dim l_strPrevDept As String
    Dim l_datDeptEffDate As Date
    Dim l_strEffDept As String

    
    On Error GoTo eh
    
    v_rsEmpLockData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    l_strDept = v_rsEmpLockData!DEPT_NUM
    l_strPrevDept = v_rsEmpLockData!PREV_DEPT_NUM
    l_datDeptEffDate = v_rsEmpLockData!DEPT_EFF_DATE
    
    l_strEffDept = EffectiveDepartment( _
        l_strDept, l_datDeptEffDate, l_strPrevDept, v_datChargeDate)
    
    If l_strEffDept = l_strDept Then
        r_strPayGroup = v_rsEmpLockData!PAYROLL_GROUP_CD
        r_strDeptLockStatus = v_rsEmpLockData!LOCK_STATUS
    Else
        r_strPayGroup = v_rsEmpLockData!PREV_PAY_GROUP
        r_strDeptLockStatus = v_rsEmpLockData!PREV_LOCK_STATUS
    End If
    r_strDept = l_strEffDept
    
    
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "GetEffectiveDeptInfo"
    
End Sub
Private Function EffectiveDepartment( _
    ByVal v_strDept As String, _
    ByVal v_datDeptEffDate As Date, _
    ByVal v_strPrevDept As String, _
    Optional ByVal v_datDate As Date = mstrDefaultDate _
    ) As String
' RED054
    On Error GoTo eh

    If EffectiveDate(v_datDate) >= v_datDeptEffDate Then
        EffectiveDepartment = v_strDept
    Else
        EffectiveDepartment = v_strPrevDept
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "EffectiveDepartment"

End Function
Public Function IsExecutive( _
            ByVal v_datJobEffDate As Date, _
            ByVal v_strUnionCode As String, _
            ByVal v_strPriorUnionCode As String, _
            Optional ByVal v_strExecContractInd As String = mstrNotAvailable, _
            Optional ByVal v_strPriorExecContractInd As String = mstrNotAvailable, _
            Optional ByVal v_strChgDate As Date = mstrDefaultDate) As String
            
    Dim l_datChgDate As Date

On Error GoTo eh

    l_datChgDate = EffectiveDate(v_strChgDate)
    IsExecutive = IsExecutivePrivate( _
        l_datChgDate, v_strExecContractInd, v_strPriorExecContractInd, _
        v_datJobEffDate, v_strUnionCode, v_strPriorUnionCode)
        
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsExecutive", , , , , True
        
End Function
Public Function EmployeeCompany( _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_strCompanyCode As String = mstrNotAvailable, _
            Optional ByVal v_strPriorCompanyCode As String = mstrNotAvailable, _
            Optional ByVal v_strChgDate As Date = mstrDefaultDate) As String
            
     Dim l_datChgDate As Date

On Error GoTo eh

    l_datChgDate = EffectiveDate(v_strChgDate)
    
    EmployeeCompany = CompanyCode( _
        l_datChgDate, v_strDeptNum, v_strPriorDeptNum, _
        v_datDeptEffDate, v_strCompanyCode, v_strPriorCompanyCode)
    
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeCompany", , , , , True
            
End Function
Private Function GetCompanyCode( _
            ByVal v_strDept As String, _
            ByVal v_strCompanyCode As String) As String
            
    Dim l_obj As dbDepartment.cdDepartment
    
On Error GoTo eh
    
    If v_strCompanyCode = mstrNotAvailable Then
        Set l_obj = CreateObject("dbDepartment.cdDepartment")
        'Set l_obj = New dbDepartment.cdDepartment
        GetCompanyCode = _
            l_obj.GetCompanyCode(v_strDept)
        Set l_obj = Nothing
    Else
        GetCompanyCode = v_strCompanyCode
    End If
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCompanyCode"
End Function
Private Function CompanyCode( _
            ByVal v_datChgDate As Date, _
            ByVal v_strDept As String, _
            ByVal v_strPriorDept As String, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strCompanyCode As String, _
            ByVal v_strPriorCompanyCode As String) As String

    Dim l_blnUseCurrent As Boolean
    
On Error GoTo eh
        
    If EffectiveDate(v_datChgDate) >= v_datDeptEffDate Then
        l_blnUseCurrent = True
    Else
        If Trim(v_strPriorDept) = "" Then
            l_blnUseCurrent = True
        Else
            l_blnUseCurrent = False
        End If
    End If
        
    If l_blnUseCurrent Then
        CompanyCode = GetCompanyCode(v_strDept, v_strCompanyCode)
    Else
        CompanyCode = GetCompanyCode(v_strPriorDept, v_strPriorCompanyCode)
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "CompanyCode"
End Function

Public Function IsAmerenCIPS( _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_strCompanyCode As String = mstrNotAvailable, _
            Optional ByVal v_strPriorCompanyCode As String = mstrNotAvailable, _
            Optional ByVal v_strChgDate As Date = mstrDefaultDate) As String

TrsRaiseError m_strCMODULENAME, "IsAmerenCIPS", trErrObsoleteMethod, , , , True

''RULE0011
'
'     Dim l_obj As busDepartment.cbsDepartment
'     Dim l_strCompany As String
'     Dim l_strDept As String
'
'
'On Error GoTo eh
'
'    Set l_obj = CreateObject("busDepartment.cbsDepartment")
'
'    l_strDept = EmployeeDepartment( _
'    v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_strChgDate)
'
'    l_strCompany = EmployeeCompany( _
'        v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, _
'        v_strCompanyCode, v_strPriorCompanyCode, v_strChgDate)
'
'    IsAmerenCIPS = l_obj.IsAmerenCIPS(l_strDept, l_strCompany)
'
'    Set l_obj = Nothing
    
    Exit Function

eh:
'    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPS", , , , , True
            
End Function

Public Function IsAmerenCIPSorCILCO( _
            ByVal v_lngRuleNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_strCompanyCode As String = mstrNotAvailable, _
            Optional ByVal v_strPriorCompanyCode As String = mstrNotAvailable, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate, _
            Optional ByVal v_strUnionCode As String = mstrNotAvailable2, _
            Optional ByVal v_strPayrollGroupCd As String = mstrNotAvailable2) As Boolean
            
     Dim l_obj As busDepartment.cbsDepartment
     Dim l_strCompany As String
     Dim l_strDept As String
On Error GoTo eh
   
    Set l_obj = CreateObject("busDepartment.cbsDepartment")
    
    l_strDept = EmployeeDepartment(v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
        
    l_strCompany = EmployeeCompany(v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, _
                                    v_strCompanyCode, v_strPriorCompanyCode, v_datChgDate)
         
    IsAmerenCIPSorCILCO = l_obj.IsAmerenCIPSorCILCO(v_lngRuleNum, l_strDept, l_strCompany, v_strUnionCode, v_datChgDate, v_strPayrollGroupCd)
    
    Set l_obj = Nothing
    
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPSorCILCO", , , , , True
End Function

Public Function EmployeeUnion( _
            ByVal v_datJobEffDate As Date, _
            ByVal v_strUnionCode As String, _
            ByVal v_strPriorUnionCode As String, _
            Optional ByVal v_strChgDate As Date = mstrDefaultDate) As String
            
    Dim l_datChgDate As Date

On Error GoTo eh

    l_datChgDate = EffectiveDate(v_strChgDate)
    If UseCurrentUnionAttribute(l_datChgDate, v_datJobEffDate, v_strPriorUnionCode) Then
        EmployeeUnion = v_strUnionCode
    Else
        EmployeeUnion = v_strPriorUnionCode
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeUnion", , , , , True
        
End Function
Public Function EmployeeDepartment( _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_strChgDate As Date = mstrDefaultDate _
            ) As String
            
     Dim l_datChgDate As Date

On Error GoTo eh

    EmployeeDepartment = _
        EffectiveDepartment( _
        v_strDeptNum, v_datDeptEffDate, v_strPriorDeptNum, v_strChgDate)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeDepartment", , , , , True
            
End Function
Public Function IsEmployeeEMPRV( _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_strInterfaceCode As String = mstrNotAvailable, _
            Optional ByVal v_strPriorInterfaceCode As String = mstrNotAvailable, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate _
            ) As String
            
    Dim l_datChgDate As Date
    Dim l_blnUseCurrent As Boolean
    Dim l_strInterfaceCode As String
    
On Error GoTo eh
        
    If EffectiveDate(v_datChgDate) >= v_datDeptEffDate Then
        l_blnUseCurrent = True
    Else
        If Trim(v_strPriorDeptNum) = "" Then
            l_blnUseCurrent = True
        Else
            l_blnUseCurrent = False
        End If
    End If
        
    If l_blnUseCurrent Then
        l_strInterfaceCode = GetInterfaceCode(v_strDeptNum, v_strInterfaceCode)
    Else
        l_strInterfaceCode = GetInterfaceCode(v_strPriorDeptNum, v_strPriorInterfaceCode)
    End If
    
    If l_strInterfaceCode = "E" Then
        IsEmployeeEMPRV = True
    Else
        IsEmployeeEMPRV = False
    End If
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsEmployeeEMPRV", , , , , True
            
End Function
Private Function GetInterfaceCode( _
            ByVal v_strDept As String, _
            ByVal v_strInterfaceCode As String) As String
            
    Dim l_obj As dbDepartment.cdDepartment
    
On Error GoTo eh
    
    If v_strInterfaceCode = mstrNotAvailable Then
        Set l_obj = CreateObject("dbDepartment.cdDepartment")
        GetInterfaceCode = _
            l_obj.GetSystemInterfaceCode(v_strDept)
        Set l_obj = Nothing
    Else
        GetInterfaceCode = v_strInterfaceCode
    End If
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetInterfaceCode"
End Function
Public Function EmployeeJobNum( _
            ByVal v_datJobEffDate As Date, _
            ByVal v_strJobNum As String, _
            ByVal v_strPriorJobNum As String, _
            Optional v_datPriorJobEffDate As Date = mstrDefaultDate, _
            Optional ByVal v_strChgDate As Date = mstrDefaultDate) As String
            
    Dim l_datChgDate As Date
    Dim l_colError As Collection
On Error GoTo eh

    l_datChgDate = EffectiveDate(v_strChgDate)
    If UseCurrentJobAttribute(l_datChgDate, v_datJobEffDate, v_strPriorJobNum, v_datPriorJobEffDate) Then
        EmployeeJobNum = v_strJobNum
    Else
        EmployeeJobNum = v_strPriorJobNum
    End If

    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeJobNum", , , , , True
        
End Function
Private Function UseCurrentJobAttribute( _
            ByVal v_datChgDate As Date, _
            ByVal v_datJobEffDate As Date, _
            ByVal v_strPriorJobNum As String, _
            ByVal v_datPriorJobEffDate _
            ) As Boolean

On Error GoTo eh

    If v_datChgDate >= v_datJobEffDate Then
        UseCurrentJobAttribute = True
    ElseIf v_datChgDate >= v_datPriorJobEffDate Then
        If Trim(v_strPriorJobNum) = "" Then
            UseCurrentJobAttribute = True
        End If
    Else
       UseCurrentJobAttribute = True
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "UseCurrentJobAttribute"
End Function
'Rule WHITE046
'Get summary case groups and case numbers with first and last date used (start date)
Public Function GetCaseGroupsAndNumbers( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date _
        ) As ADODB.Recordset
TrsRaiseError m_strCMODULENAME, "GetCaseGroupsAndNumbers", trErrObsoleteMethod, , , , True
'    Dim l_objEmployee As dbEmployee.cdEmployee
'    Dim l_datStartDate As Date
'    Dim l_datEndDate As Date
'
'    On Error GoTo eh
'
'    l_datStartDate = DateSerial(Year(v_datStartDate), 1, 1)
'    l_datEndDate = DateSerial(Year(v_datEndDate), 12, 31)
''//-----don't know why this was added in version 71
''//-----discovered the project was modified while doing TRIS 5.0 build on 3/15/04, so a release label was added
''//-----SchedExceptions call this method with and end date of #12/31/9999# which fails
''    l_datStartDate = DateAdd("yyyy", -1, l_datStartDate)
''    l_datEndDate = DateAdd("yyyy", 1, l_datEndDate)
'
'    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
'
'    Set GetCaseGroupsAndNumbers = l_objEmployee.GetCaseGroupsAndNumbers( _
'        v_aryEmployees, l_datStartDate, l_datEndDate)
'    Set l_objEmployee = Nothing
'
'    Exit Function
'
'eh:
'    Set l_objEmployee = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetCaseGroupsAndNumbers", , , , , True

End Function
Public Function PartialHourMinutes( _
        ByVal v_strDept As String, _
        ByVal v_blnIsExecutive As Boolean, _
        Optional ByVal v_strCompany As String = mstrNotAvailable) As Integer
    
' Answer the minutes that partial hours may be entered.
    Dim l_strDepartmentNumber As String
    
    On Error GoTo eh
    
    If DoesRuleApply(166, trDepartmentNumber, v_strDept) Then
        PartialHourMinutes = 3  'Callaway may be going to 3 minute increments
    Else
        If Not v_blnIsExecutive Then
            PartialHourMinutes = 6
        Else                            'must be executive
            If DoesRuleApply(165, trDepartmentNumber, v_strDept) Then
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

Public Function IsAmerenCILCO( _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_strCompanyCode As String = mstrNotAvailable, _
            Optional ByVal v_strPriorCompanyCode As String = mstrNotAvailable, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
            
TrsRaiseError m_strCMODULENAME, "IsAmerenCIPS", trErrObsoleteMethod, , , , True

''RULE0012
'
'     Dim l_obj As busDepartment.cbsDepartment
'     Dim l_strCompany As String
'     Dim l_strDept As String
'
'
'On Error GoTo eh
'
'    Set l_obj = CreateObject("busDepartment.cbsDepartment")
'
'    l_strDept = EmployeeDepartment( _
'    v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
'
'    l_strCompany = EmployeeCompany( _
'        v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, _
'        v_strCompanyCode, v_strPriorCompanyCode, v_datChgDate)
'
'    IsAmerenCILCO = l_obj.IsAmerenCILCO(l_strDept, l_strCompany)
'
'    Set l_obj = Nothing
    
    Exit Function

eh:
'    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAmerenCILCO", , , , , True
    
    
End Function
Public Function IsAmerenCILCODispatch( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
            
'RULE0015
'RULE0016
            
     Dim l_obj As busDepartment.cbsDepartment
     Dim l_strDept As String
     
On Error GoTo eh
    
    Set l_obj = CreateObject("busDepartment.cbsDepartment")
         
    l_strDept = EmployeeDepartment( _
        v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
    
    If l_obj.IsAmerenCILCODispatch(l_strDept) Then 'RULE0015
        IsAmerenCILCODispatch = True
    Else
        IsAmerenCILCODispatch = DoesRuleApply(16, trEmployeeNumber, v_lngEmpNum)
    End If
    
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAmerenCILCODispatch", , , , , True
    
            
End Function
Public Function IsEdwardsOperations( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate _
            ) As Boolean
' CILCO
'RULE0013
'RULE0014
    
    Dim l_obj As busDepartment.cbsDepartment
    Dim l_strDept As String

    On Error GoTo eh
    
    l_strDept = EmployeeDepartment( _
        v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
        

    Set l_obj = CreateObject("busDepartment.cbsDepartment")
    If l_obj.IsEdwardsOperations(l_strDept) Then 'RULE0013
        IsEdwardsOperations = True
    Else
        IsEdwardsOperations = DoesRuleApply(14, trEmployeeNumber, v_lngEmpNum)
    End If
    
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsEdwardsOperations", , , , , True

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
Public Function IsCoffeenOperations( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate _
            ) As Boolean

'RULE0010

    
    Dim l_obj As busDepartment.cbsDepartment
    Dim l_strDept As String

    On Error GoTo eh
    
    l_strDept = EmployeeDepartment( _
        v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
        
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
Public Function IsCoffeenMaintenance( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate _
            ) As Boolean

'RULE0009

    
    Dim l_obj As busDepartment.cbsDepartment
    Dim l_strDept As String

    On Error GoTo eh
    
    l_strDept = EmployeeDepartment( _
        v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
        
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
    TrsRaiseError m_strCMODULENAME, "IsCoffeenOperations", , , , , True

End Function
Public Function IsAmerenCILCOLocal8Employee( _
            ByVal v_datJobEffDate As Date, _
            ByVal v_strUnionCode As String, _
            ByVal v_strPriorUnionCode As String, _
            Optional ByVal v_strChgDate As Date = mstrDefaultDate) As String
'RULE0019
            
    Dim l_datChgDate As Date
    Dim l_strUnion As String

On Error GoTo eh

    l_datChgDate = EffectiveDate(v_strChgDate)
    If UseCurrentUnionAttribute(l_datChgDate, v_datJobEffDate, v_strPriorUnionCode) Then
        l_strUnion = v_strUnionCode
    Else
        l_strUnion = v_strPriorUnionCode
    End If
    
    IsAmerenCILCOLocal8Employee = DoesRuleApply(19, trUnionCode, l_strUnion)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCILCOLocal8Employee", , , , , True
        
End Function
'PREM BEGIN TRIS CONVERSION PHASE 3 (PHASE 1 .NET)
'The following function retrieves Qualifications
' and Special Skills assigned to an employee or employees
Public Function GetQualsAndSkillsForEmployees( _
                    ByVal vadrsEmployees As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo GetQualsAndSkillsForEmployees_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee
Dim llIndex As Long, lsWhereClause As String

    If (mADOExtensions.RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                For llIndex = 1 To .RecordCount
                    lsWhereClause = lsWhereClause & _
                                    CStr(!EMP_NUM) & ","
                    .MoveNext
                Next llIndex

                If (Trim(lsWhereClause) <> "") Then
                    If (Right(lsWhereClause, 1) = ",") Then
                        lsWhereClause = Left(lsWhereClause, Len(lsWhereClause) - 1)
                    End If
                End If
            End If
        End With
    Else
        Set GetQualsAndSkillsForEmployees = Nothing
        Exit Function
    End If

    If (Trim(lsWhereClause) <> "") Then
        'Construct
        Set lobjdbEmployee = CreateObject( _
                                "dbEmployee.cdEmployee")
        'Set lobjdbEmployee = New dbEmployee.cdEmployee
        'Get and set for return
        Set GetQualsAndSkillsForEmployees = _
                lobjdbEmployee.GetQualsAndSkillsForEmployees( _
                    lsWhereClause)
    End If
    'Destroy
    Set lobjdbEmployee = Nothing
    Exit Function
GetQualsAndSkillsForEmployees_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetQualsAndSkillsForEmployees" _
        , , , , _
        , True

End Function
Public Function GetQualsAndSkillsForEmployees2( _
                    ByVal vadrsEmployees As ADODB.Recordset _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetQualsAndSkillsForEmployees2_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee
Dim llIndex As Long, lsWhereClause As String

    If (mADOExtensions.RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                For llIndex = 1 To .RecordCount
                    lsWhereClause = lsWhereClause & _
                                    CStr(!EMP_NUM) & ","
                    .MoveNext
                Next llIndex

                If (Trim(lsWhereClause) <> "") Then
                    If (Right(lsWhereClause, 1) = ",") Then
                        lsWhereClause = Left(lsWhereClause, Len(lsWhereClause) - 1)
                    End If
                End If
            End If
        End With
    Else
        Set GetQualsAndSkillsForEmployees2 = Nothing
        Exit Function
    End If

    If (Trim(lsWhereClause) <> "") Then
        'Construct
        Set lobjdbEmployee = CreateObject( _
                                "dbEmployee.cdEmployee")
        'Set lobjdbEmployee = New dbEmployee.cdEmployee
        'Get and set for return
        Set GetQualsAndSkillsForEmployees2 = _
                lobjdbEmployee.GetQualsAndSkillsForEmployees2( _
                    lsWhereClause _
                    , vdtStartDate _
                    , vdtEndDate)
    End If
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetQualsAndSkillsForEmployees2_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetQualsAndSkillsForEmployees2" _
        , , , , _
        , True

End Function
Public Function GetExpiringQAndSByLocation( _
                    ByVal viLocation As Integer _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetExpiringQAndSByLocation_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetExpiringQAndSByLocation = _
            lobjdbEmployee.GetExpiringQAndSByLocation( _
                viLocation _
                , vdtStartDate _
                , vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetExpiringQAndSByLocation_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetExpiringQAndSByLocation" _
        , , , , _
        , True

End Function
Public Function GetExpiringQAndSByDept( _
                    ByVal vsDepartment As String _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetExpiringQAndSByDept_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetExpiringQAndSByDept = _
            lobjdbEmployee.GetExpiringQAndSByDept( _
                vsDepartment _
                , vdtStartDate _
                , vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetExpiringQAndSByDept_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetExpiringQAndSByDept" _
        , , , , _
        , True

End Function
Public Function GetExpiringQAndSByWG( _
                    ByVal vlWorkgroup As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetExpiringQAndSByWG_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetExpiringQAndSByWG = _
            lobjdbEmployee.GetExpiringQAndSByWG( _
                vlWorkgroup _
                , vdtStartDate _
                , vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetExpiringQAndSByWG_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetExpiringQAndSByWG" _
        , , , , _
        , True

End Function
'The following function retrieves seniority information
'for the given employee
Public Function GetEmployeeSeniority( _
                    ByVal vlEmployeeNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeSeniority_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Set lobjdbEmployee = New dbEmployee.cdEmployee
    'Construct
    Set lobjdbEmployee = CreateObject( _
                        "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetEmployeeSeniority = _
            lobjdbEmployee.GetEmployeeSeniority( _
                vlEmployeeNumber)
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeSeniority_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeSeniority" _
        , , , , _
        , True

End Function
'The following function retrieves frozen seniority information
'for the given employee
Public Function GetEmployeeFrozenSeniority( _
                    ByVal vlEmployeeNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeFrozenSeniority_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Set lobjdbEmployee = New dbEmployee.cdEmployee
    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetEmployeeFrozenSeniority = _
            lobjdbEmployee.GetEmployeeFrozenSeniority( _
                vlEmployeeNumber)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeFrozenSeniority_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeFrozenSeniority" _
        , , , , _
        , True

End Function
'The following subroutine gets both seniority and frozen seniority information
Public Sub GetAllEmployeeSeniorityInformation( _
                    ByVal vlEmployeeNumber As Long _
                    , ByRef vbIsAGencoEmployee As Boolean _
                    , ByRef radrsSeniority As ADODB.Recordset _
                    , ByRef radrsFrozenSeniority As ADODB.Recordset)
On Error GoTo GetAllEmployeeSeniorityInformation_EH
Dim ladrsEmployeeInformation As ADODB.Recordset

    'Get basic employee information
    Set ladrsEmployeeInformation = GetBasicEmployeeInformation(vlEmployeeNumber)
    'check to see if the information exists
    If (mADOExtensions.RecordsExist(ladrsEmployeeInformation)) Then
        With ladrsEmployeeInformation
            'Check to see if the Employee's Company Code follows
            'AmerenCIPS rules (CIP, GEN, IHC, AED, IMS)
            'the Rule Code is 11, the Business Type is Company
            'the value is returned by Employee Company
            'vbIsAGencoEmployee = False
            vbIsAGencoEmployee = IsAGencoEmployee(EmployeeCompany(!DEPT_EFF_DATE, !DEPT_NUM, !PREV_DEPT_NUM), _
                                                EmployeeUnion(!JOB_EFF_DATE, Trim(!UNION_CD), Trim(!PRIOR_UNION_CD)), _
                                                !PAYROLL_GROUP_CD, _
                                                Date)

'            vbIsAGencoEmployee = DoesRuleApply(11 _
'                                    , trCompanyCode _
'                                    , EmployeeCompany(!DEPT_EFF_DATE _
'                                        , !DEPT_NUM _
'                                        , !PREV_DEPT_NUM))
        End With

        Set ladrsEmployeeInformation = Nothing
    End If
    Set ladrsEmployeeInformation = Nothing

    Set radrsSeniority = GetEmployeeSeniority( _
                            vlEmployeeNumber)
    Set radrsFrozenSeniority = GetEmployeeFrozenSeniority( _
                                    vlEmployeeNumber)
    'Destroy
    Exit Sub
GetAllEmployeeSeniorityInformation_EH:
    Set ladrsEmployeeInformation = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetAllEmployeeSeniorityInformation" _
        , , , , _
        , True

End Sub

'//made public for unit testing
Public Function IsAGencoEmployee(ByVal v_strCmpyCd As String, ByVal v_strUnionCd As String, ByVal v_strPayPeriodCd As String, ByVal v_datChargeDate As Date) As Boolean
    Dim lbIsAGencoEmployee As Boolean
    Dim oRules As dbRules.cdEffectiveDateRules
    On Error GoTo eh
 
    Set oRules = CreateObject("dbRules.cdEffectiveDateRules")
    
    If oRules.DoesEffectiveDateRuleApply(dbRules.TRSRuleCode.trAICEffectiveDates, trPayrollGroupCode, v_strPayPeriodCd, v_datChargeDate) Then
        lbIsAGencoEmployee = DoesRuleApply(dbRules.TRSRuleCode.trIsAmerenCIPS _
                                        , trUnionCode _
                                        , v_strUnionCd)
    Else
        lbIsAGencoEmployee = DoesRuleApply(dbRules.TRSRuleCode.trIsAmerenCIPS _
                                        , trCompanyCode _
                                        , v_strCmpyCd)
    End If
    
    IsAGencoEmployee = lbIsAGencoEmployee
    
    Set oRules = Nothing
    Exit Function
eh:
    Set oRules = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAGencoEmployee", , , , , True
End Function

'The following function retrieves seniority information
'for the given department
Public Function GetEmployeeSeniorityForDept( _
                    ByVal vsDepartment As String _
                    , ByVal vbActiveEmployeesOnly As Boolean _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeSeniorityForDept_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Set lobjdbEmployee = New dbEmployee.cdEmployee
    'Construct
    Set lobjdbEmployee = CreateObject( _
                        "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetEmployeeSeniorityForDept = _
            lobjdbEmployee.GetEmployeeSeniorityForDept( _
                vsDepartment, vbActiveEmployeesOnly)
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeSeniorityForDept_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeSeniorityForDept" _
        , , , , _
        , True

End Function
'The following function retrieves frozen seniority information
'for the given department
Public Function GetEmployeeFrznSnrtyForDept( _
                    ByVal vsDepartment As String _
                    , ByVal vbActiveEmployeesOnly As Boolean _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeFrznSnrtyForDept_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Set lobjdbEmployee = New dbEmployee.cdEmployee
    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetEmployeeFrznSnrtyForDept = _
            lobjdbEmployee.GetEmployeeFrznSnrtyForDept( _
                vsDepartment, vbActiveEmployeesOnly)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeFrznSnrtyForDept_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeFrznSnrtyForDept" _
        , , , , _
        , True

End Function

'TODO: What if the department's union code is blank ?
'The following subroutine gets both seniority and frozen seniority information
Public Sub GetAllEmployeeSeniorityForDept( _
                    ByVal vsDepartment As String _
                    , ByVal vbActiveEmployeesOnly As Boolean _
                    , ByRef vbIsAGencoDepartment As Boolean _
                    , ByRef radrsSeniority As ADODB.Recordset _
                    , ByRef radrsFrozenSeniority As ADODB.Recordset)
On Error GoTo eh
    Dim lobjdbDepartment As dbDepartment.cdDepartment
    Dim lsCompany As String
    Dim lsUnionCode As String
    Dim lsPayrollGroupCd As String

    lsCompany = ""
    Set lobjdbDepartment = CreateObject("dbDepartment.cdDepartment")
    lsCompany = lobjdbDepartment.GetCompanyCode(vsDepartment)
    lsUnionCode = lobjdbDepartment.GetUnionCode(vsDepartment)
    lsPayrollGroupCd = lobjdbDepartment.GetPayrollGroup(vsDepartment)
    Set lobjdbDepartment = Nothing

    vbIsAGencoDepartment = IsAGencoEmployee(lsCompany, _
                                        lsUnionCode, _
                                        lsPayrollGroupCd, _
                                        Date)

'    If (lsCompany <> "") Then
        'Check to see if the Employee's Company Code follows
        'AmerenCIPS rules (CIP, GEN, IHC, AED, IMS)
        'the Rule Code is 11, the Business Type is Company
        'the value is returned by Employee Company
        'vbIsAGencoEmployee = False
        
'        vbIsAGencoDepartment = DoesRuleApply(11 _
'                                , trCompanyCode _
'                                , lsCompany)
'    End If

    Set radrsSeniority = GetEmployeeSeniorityForDept(vsDepartment, vbActiveEmployeesOnly)

    Set radrsFrozenSeniority = GetEmployeeFrznSnrtyForDept(vsDepartment, vbActiveEmployeesOnly)
    'Destroy
    Exit Sub
eh:
    Set lobjdbDepartment = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetAllEmployeeSeniorityForDept" _
        , , , , _
        , True
End Sub

'The following function retrieves seniority information
'for the given workgroup
Public Function GetEmployeeSeniorityForWG( _
                    ByVal vlWorkgroup As Long _
                    , ByVal vbActiveEmployeesOnly As Boolean _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeSeniorityForWG_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Set lobjdbEmployee = New dbEmployee.cdEmployee
    'Construct
    Set lobjdbEmployee = CreateObject( _
                        "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetEmployeeSeniorityForWG = _
            lobjdbEmployee.GetEmployeeSeniorityForWG( _
                vlWorkgroup, vbActiveEmployeesOnly)
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeSeniorityForWG_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeSeniorityForWG" _
        , , , , _
        , True

End Function
'The following function retrieves frozen seniority information
'for the given workgroup
Public Function GetEmployeeFrznSnrtyForWG( _
                    ByVal vlWorkgroup As Long _
                    , ByVal vbActiveEmployeesOnly As Boolean _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeFrznSnrtyForWG_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Set lobjdbEmployee = New dbEmployee.cdEmployee
    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetEmployeeFrznSnrtyForWG = _
            lobjdbEmployee.GetEmployeeFrznSnrtyForWG( _
                vlWorkgroup, vbActiveEmployeesOnly)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeFrznSnrtyForWG_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeFrznSnrtyForWG" _
        , , , , _
        , True

End Function

'//the vbIsAGencoWorkgroup should be defined ByRef
'//TODO: Make this obsolete after all clients have been updated to use method 2
Public Sub GetAllEmployeeSeniorityForWG( _
                    ByVal vlWorkgroup As Long _
                    , ByVal vbActiveEmployeesOnly As Boolean _
                    , ByVal vbIsAGencoWorkgroup As Boolean _
                    , ByRef radrsSeniority As ADODB.Recordset _
                    , ByRef radrsFrozenSeniority As ADODB.Recordset)
On Error GoTo eh

    GetAllEmployeeSeniorityForWG2 vlWorkgroup, vbActiveEmployeesOnly, vbIsAGencoWorkgroup, radrsSeniority, radrsFrozenSeniority
    'Destroy
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "GetAllEmployeeSeniorityForWG", , , , , True
End Sub

'TODO: What if the department's union code is blank ?
'The following subroutine gets both seniority and frozen seniority information
Public Sub GetAllEmployeeSeniorityForWG2( _
                    ByVal vlWorkgroup As Long _
                    , ByVal vbActiveEmployeesOnly As Boolean _
                    , ByRef vbIsAGencoWorkgroup As Boolean _
                    , ByRef radrsSeniority As ADODB.Recordset _
                    , ByRef radrsFrozenSeniority As ADODB.Recordset)
On Error GoTo eh
    Dim lobjdbWorkgroup As dbWorkGroup.cdWorkGroup
    Dim lobjdbDepartment As dbDepartment.cdDepartment
    Dim lsCompany As String, lsDepartment As String
    Dim lsUnionCode As String
    Dim lsPayrollGroupCd As String
    lsDepartment = ""
    lsCompany = ""

    Set lobjdbWorkgroup = CreateObject("dbWorkGroup.cdWorkGroup")
    lsDepartment = lobjdbWorkgroup.GetDepartmentNumber(vlWorkgroup)
    Set lobjdbWorkgroup = Nothing

    If (lsDepartment <> "") Then
        Set lobjdbDepartment = CreateObject("dbDepartment.cdDepartment")
        lsCompany = lobjdbDepartment.GetCompanyCode(lsDepartment)
        lsUnionCode = lobjdbDepartment.GetUnionCode(lsDepartment)
        lsPayrollGroupCd = lobjdbDepartment.GetPayrollGroup(lsDepartment)
        Set lobjdbDepartment = Nothing

        'Check to see if the Employee's Company Code follows
        'AmerenCIPS rules (CIP, GEN, IHC, AED, IMS)
        'the Rule Code is 11, the Business Type is Company
        'the value is returned by Employee Company
        'vbIsAGencoEmployee = False
        vbIsAGencoWorkgroup = IsAGencoEmployee(lsCompany, _
                                        lsUnionCode, _
                                        lsPayrollGroupCd, _
                                        Date)

'        If (lsCompany <> "") Then
'            vbIsAGencoWorkgroup = DoesRuleApply(11 _
'                                , trCompanyCode _
'                                , lsCompany)
'        End If
    End If

    Set radrsSeniority = GetEmployeeSeniorityForWG(vlWorkgroup, vbActiveEmployeesOnly)

    Set radrsFrozenSeniority = GetEmployeeFrznSnrtyForWG(vlWorkgroup, vbActiveEmployeesOnly)
    'Destroy
    Exit Sub
eh:
    Set lobjdbDepartment = Nothing
    Set lobjdbWorkgroup = Nothing

    TrsRaiseError m_strCMODULENAME, "GetAllEmployeeSeniorityForWG2", , , , , True
End Sub

'The following Function retrieves Overtime Canvass Jobs for Employees
Public Function GetEmployeeCanvassJob( _
                    ByVal vadrsEmployees As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeCanvassJob_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee
Dim llIndex As Long, lsWhereClause As String

    If (mADOExtensions.RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                For llIndex = 1 To .RecordCount
                    lsWhereClause = lsWhereClause & _
                                    CStr(!EMP_NUM) & ","
                    .MoveNext
                Next llIndex

                If (Trim(lsWhereClause) <> "") Then
                    If (Right(lsWhereClause, 1) = ",") Then
                        lsWhereClause = Left(lsWhereClause _
                                            , Len(lsWhereClause) - 1)
                    End If
                End If
            End If
        End With
    Else
        Set GetEmployeeCanvassJob = Nothing
        Exit Function
    End If

    If (Trim(lsWhereClause) <> "") Then
        'Construct
        Set lobjdbEmployee = CreateObject( _
                                "dbEmployee.cdEmployee")
        'Get and set for return
        Set GetEmployeeCanvassJob = _
                lobjdbEmployee.GetEmployeeCanvassJob( _
                    lsWhereClause)
    End If
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeCanvassJob_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeCanvassJob" _
        , , , , _
        , True

End Function
Public Function GetEmployeeCanvassJob2( _
                    ByVal vadrsEmployees As ADODB.Recordset _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeCanvassJob2_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee
Dim llIndex As Long, lsWhereClause As String

    If (mADOExtensions.RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                For llIndex = 1 To .RecordCount
                    lsWhereClause = lsWhereClause & _
                                    CStr(!EMP_NUM) & ","
                    .MoveNext
                Next llIndex

                If (Trim(lsWhereClause) <> "") Then
                    If (Right(lsWhereClause, 1) = ",") Then
                        lsWhereClause = Left(lsWhereClause _
                                            , Len(lsWhereClause) - 1)
                    End If
                End If
            End If
        End With
    Else
        Set GetEmployeeCanvassJob2 = Nothing
        Exit Function
    End If

    If (Trim(lsWhereClause) <> "") Then
        'Construct
        Set lobjdbEmployee = CreateObject( _
                                "dbEmployee.cdEmployee")
        'Get and set for return
        Set GetEmployeeCanvassJob2 = _
                lobjdbEmployee.GetEmployeeCanvassJob2( _
                    lsWhereClause _
                    , vdtStartDate _
                    , vdtEndDate)
    End If
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeCanvassJob2_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeCanvassJob2" _
        , , , , _
        , True

End Function
'This Function will retrieve Employee Overtime Ineligibilities
Public Function GetEmployeeIneligibilities( _
                    ByVal vadrsEmployees As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeIneligibilities_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee
Dim llIndex As Long, lsWhereClause As String

    If (mADOExtensions.RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                For llIndex = 1 To .RecordCount
                    lsWhereClause = lsWhereClause & _
                                    CStr(!EMP_NUM) & ","
                    .MoveNext
                Next llIndex

                If (Trim(lsWhereClause) <> "") Then
                    If (Right(lsWhereClause, 1) = ",") Then
                        lsWhereClause = Left(lsWhereClause _
                                            , Len(lsWhereClause) - 1)
                    End If
                End If
            End If
        End With
    Else
        Set GetEmployeeIneligibilities = Nothing
        Exit Function
    End If

    If (Trim(lsWhereClause) <> "") Then
        'Construct
        'Construct
        Set lobjdbEmployee = CreateObject( _
                                "dbEmployee.cdEmployee")
        'Get and set for return
        Set GetEmployeeIneligibilities = _
                lobjdbEmployee.GetEmployeeIneligibilities( _
                    lsWhereClause)
    End If
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeIneligibilities_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeIneligibilities" _
        , , , , _
        , True

End Function
Public Function GetEmployeeIneligibilities2( _
                    ByVal vadrsEmployees As ADODB.Recordset _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeIneligibilities2_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee
Dim llIndex As Long, lsWhereClause As String

    If (mADOExtensions.RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                For llIndex = 1 To .RecordCount
                    lsWhereClause = lsWhereClause & _
                                    CStr(!EMP_NUM) & ","
                    .MoveNext
                Next llIndex

                If (Trim(lsWhereClause) <> "") Then
                    If (Right(lsWhereClause, 1) = ",") Then
                        lsWhereClause = Left(lsWhereClause _
                                            , Len(lsWhereClause) - 1)
                    End If
                End If
            End If
        End With
    Else
        Set GetEmployeeIneligibilities2 = Nothing
        Exit Function
    End If

    If (Trim(lsWhereClause) <> "") Then
        'Construct
        'Construct
        Set lobjdbEmployee = CreateObject( _
                                "dbEmployee.cdEmployee")
        'Get and set for return
        Set GetEmployeeIneligibilities2 = _
                lobjdbEmployee.GetEmployeeIneligibilities2( _
                    lsWhereClause _
                    , vdtStartDate _
                    , vdtEndDate)
    End If
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeIneligibilities2_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeIneligibilities2" _
        , , , , _
        , True

End Function
'This Function will retrieve Newton Plant Overtime Volunteer Information
Public Function GetVolunteerInformation() As ADODB.Recordset
On Error GoTo GetVolunteerInformation_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee
Dim ladrsVlntrInfo As ADODB.Recordset
Dim ladrsVlntrAction As ADODB.Recordset

Dim llRecordIndex As Long
'NOTE: I am not sure if the following is okay. i.e. Construct
'the object call multiple methods in the object and then
'destroy it.
    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    'Get and set for return
    Set ladrsVlntrInfo = _
            lobjdbEmployee.GetVolunteerInformation()

    Set ladrsVlntrAction = _
            lobjdbEmployee.GetVolunteerAction()
    'Destroy
    Set lobjdbEmployee = Nothing

    If (mADOExtensions.RecordsExist(ladrsVlntrAction) And _
        mADOExtensions.RecordsExist(ladrsVlntrInfo)) Then
        With ladrsVlntrAction
            .MoveFirst
            For llRecordIndex = 1 To .RecordCount
                ladrsVlntrInfo.MoveFirst

                ladrsVlntrInfo.Find "EMP_NUM = " & !EMP_NUM
                If Not (ladrsVlntrInfo.EOF) Then
                    Select Case (UCase(Trim(!ACTN_CD)))
                        Case "YES"
                            ladrsVlntrInfo!VLNTR_STATUS = "YES"
                        Case "NO"
                            ladrsVlntrInfo!VLNTR_STATUS = "NO"
                    End Select
                    ladrsVlntrInfo!STUS_EFF_DT = !STUS_EFF_DT
                    ladrsVlntrInfo!TIME_STAMP = !TIME_STAMP
                    ladrsVlntrInfo.Update
                End If

                .MoveNext
            Next llRecordIndex
        End With

        ladrsVlntrInfo.MoveFirst
    End If

    'RETURN
    Set GetVolunteerInformation = ladrsVlntrInfo

    'DESTROY
    Set ladrsVlntrAction = Nothing
    Set ladrsVlntrInfo = Nothing

    Exit Function
GetVolunteerInformation_EH:
    Set lobjdbEmployee = Nothing

    Set ladrsVlntrAction = Nothing
    Set ladrsVlntrInfo = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetVolunteerInformation" _
        , , , , _
        , True

End Function
'PROJECT BANKS MAINTENANCE
Public Function GetEmployeesInProject(ByVal vlProjectID As Long _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999# _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeesInProject_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    If (vdtStartDate = #1/1/100#) Then
        vdtStartDate = Date
    End If
    If (vdtStartDate = #12/31/9999#) Then
        vdtEndDate = Date
    End If

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetEmployeesInProject = _
            lobjdbEmployee.GetEmployeesInProject(vlProjectID _
                , vdtStartDate _
                , vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeesInProject_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeesInProject" _
        , , , , _
        , True

End Function
Public Function GetEmployeesForProjectByWorkgroup(ByVal vlProjectID As Long _
                    , ByVal vlWorkgroupNumber As Long _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999# _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeesForProjectByWorkgroup_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    If (vdtStartDate = #1/1/100#) Then
        vdtStartDate = Date
    End If
    If (vdtStartDate = #12/31/9999#) Then
        vdtEndDate = Date
    End If

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetEmployeesForProjectByWorkgroup = _
            lobjdbEmployee.GetEmployeesForProjectByWorkgroup(vlProjectID _
                , vlWorkgroupNumber _
                , vdtStartDate _
                , vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeesForProjectByWorkgroup_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeesForProjectByWorkgroup" _
        , , , , _
        , True

End Function
Public Function GetEmployeesForProjectByLocation(ByVal vlProjectID As Long _
                    , ByVal vlLocationCode As Long _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999# _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeesForProjectByLocation_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    If (vdtStartDate = #1/1/100#) Then
        vdtStartDate = Date
    End If
    If (vdtStartDate = #12/31/9999#) Then
        vdtEndDate = Date
    End If

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetEmployeesForProjectByLocation = _
            lobjdbEmployee.GetEmployeesForProjectByLocation(vlProjectID _
                , vlLocationCode _
                , vdtStartDate _
                , vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeesForProjectByLocation_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeesForProjectByLocation" _
        , , , , _
        , True

End Function
Public Function GetEmployeesForProjectByCraft(ByVal vlProjectID As Long _
                    , ByVal vsCraftNumber As String _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999# _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeesForProjectByCraft_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    If (vdtStartDate = #1/1/100#) Then
        vdtStartDate = Date
    End If
    If (vdtStartDate = #12/31/9999#) Then
        vdtEndDate = Date
    End If

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    'Get and set for return
    Set GetEmployeesForProjectByCraft = _
            lobjdbEmployee.GetEmployeesForProjectByCraft(vlProjectID _
                , vsCraftNumber _
                , vdtStartDate _
                , vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeesForProjectByCraft_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeesForProjectByCraft" _
        , , , , _
        , True

End Function
Public Function GetEmployeeECWorkReasons(ByVal vbActiveEmployeesOnly As Boolean _
                    ) As ADODB.Recordset
TrsRaiseError m_strCMODULENAME, "GetEmployeeECWorkReasons", trErrObsoleteMethod, , , , True
'On Error GoTo GetEmployeeECWorkReasons_EH
''Declare
'Dim lobjdbEmployee As dbEmployee.cdEmployee
'Dim ldtStatusEffectiveFromDate As Date
'
'    If (vbActiveEmployeesOnly) Then
'        ldtStatusEffectiveFromDate = _
'                CDate(Format(DateAdd("m", -6, Date) _
'                    , "mm/dd/yyyy"))
'    Else
'        ldtStatusEffectiveFromDate = #1/1/1900#
'    End If
'
'    'Construct
'    Set lobjdbEmployee = CreateObject( _
'                            "dbEmployee.cdEmployee")
'    'Get and set for return
'    Set GetEmployeeECWorkReasons = _
'            lobjdbEmployee.GetEmployeeECWorkReasons(vbActiveEmployeesOnly _
'                            , ldtStatusEffectiveFromDate)
'    'Destroy
'    Set lobjdbEmployee = Nothing
'
'    Exit Function
'GetEmployeeECWorkReasons_EH:
'    Set lobjdbEmployee = Nothing
'
'    TrsRaiseError m_strCMODULENAME _
'        , "GetEmployeesForProjectByCraft" _
'        , , , , _
'        , True
End Function
Public Function GetEmployeeSickLeaveInformation(ByVal vlEmployeeNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeSickLeaveInformation_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    Set GetEmployeeSickLeaveInformation = lobjdbEmployee.GetEmployeeSickLeaveInformation( _
                vlEmployeeNumber)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeSickLeaveInformation_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeSickLeaveInformation" _
        , , , , _
        , True
End Function
Public Function GetEmployeeSickLeaveInformation2(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date) As ADODB.Recordset
On Error GoTo GetEmployeeSickLeaveInformation2_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    Set GetEmployeeSickLeaveInformation2 = lobjdbEmployee.GetEmployeeSickLeaveInformation2( _
                vlEmployeeNumber _
                , vdtStartDate _
                , vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeSickLeaveInformation2_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeSickLeaveInformation2" _
        , , , , _
        , True
End Function
Public Sub GetEmployeeAllowedTimeInformation(ByVal vlEmployeeNumber As Long _
                    , ByRef radrsAllowedTimeDetails As ADODB.Recordset _
                    , ByRef radrsAllowedTimeTotals As ADODB.Recordset _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999#)
On Error GoTo GetEmployeeAllowedTimeInformation_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    Set radrsAllowedTimeDetails = lobjdbEmployee.GetEmployeeAllowedTimeDetails( _
                vlEmployeeNumber, vdtStartDate, vdtEndDate)
    Set radrsAllowedTimeTotals = lobjdbEmployee.GetEmployeeAllowedTimeTotals( _
                vlEmployeeNumber, vdtStartDate, vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Sub
GetEmployeeAllowedTimeInformation_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeAllowedTimeInformation" _
        , , , , _
        , True
End Sub
Public Sub GetEmployeeCertificationInformation(ByVal vlEmployeeNumber As Long _
                    , ByRef radrsSummary As ADODB.Recordset _
                    , ByRef radrsDetails As ADODB.Recordset _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999#)
On Error GoTo GetEmployeeCertificationInformation_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    Set radrsSummary = lobjdbEmployee.GetEmployeeCertificationSummary( _
                vlEmployeeNumber, vdtStartDate, vdtEndDate)
    Set radrsDetails = lobjdbEmployee.GetEmployeeCertificationDetails( _
                vlEmployeeNumber, vdtStartDate, vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Sub
GetEmployeeCertificationInformation_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeCertificationInformation" _
        , , , , _
        , True
End Sub
Public Function GetEmployeeExcusedTimeInformation(ByVal vlEmployeeNumber As Long _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999# _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeExcusedTimeInformation_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    Set GetEmployeeExcusedTimeInformation = lobjdbEmployee.GetEmployeeExcusedTimeInformation( _
                vlEmployeeNumber, vdtStartDate, vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeExcusedTimeInformation_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeExcusedTimeInformation" _
        , , , , _
        , True
End Function
'PREM END TRIS CONVERSION PHASE 3 (PHASE 1 .NET)
'Begin Prem TRIS 5.7
Public Function GetEmployeePSContactInformation(ByVal vlEmployee As Long _
                ) As ADODB.Recordset
On Error GoTo GetEmployeePSContactInformation_EH
Dim lobjdbEmployee As dbEmployee.cdEmployee

    Set GetEmployeePSContactInformation = Nothing

    Set lobjdbEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeePSContactInformation = _
            lobjdbEmployee.GetEmployeePSContactInformation(vlEmployee)
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeePSContactInformation_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeePSContactInformation" _
        , , , , _
        , True

End Function
Public Sub GetEmployeeCertificationInformation2(ByVal vlEmployeeNumber As Long _
                    , ByRef radrsSummaryByYear As ADODB.Recordset _
                    , ByRef radrsSummaryByDate As ADODB.Recordset _
                    , ByRef radrsDetails As ADODB.Recordset _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999#)
On Error GoTo GetEmployeeCertificationInformation2_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    Set radrsSummaryByYear = lobjdbEmployee.GetEmployeeCertificationByYear( _
                vlEmployeeNumber, vdtStartDate, vdtEndDate)
    Set radrsSummaryByDate = lobjdbEmployee.GetEmployeeCertificationByDate( _
                vlEmployeeNumber, vdtStartDate, vdtEndDate)
    Set radrsDetails = lobjdbEmployee.GetEmployeeCertificationDetails( _
                vlEmployeeNumber, vdtStartDate, vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Sub
GetEmployeeCertificationInformation2_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeCertificationInformation2" _
        , , , , _
        , True
End Sub
'//for TRIS 6.6 Clients
Public Sub GetEmployeeExcusedTimeInformation2(ByVal vlEmployeeNumber As Long _
                    , ByRef radrsSummary As ADODB.Recordset _
                    , ByRef radrsDetails As ADODB.Recordset _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999#)
On Error GoTo GetEmployeeExcusedTimeInformation2_EH
    'Declare
    Dim lobjdbEmployee As dbEmployee.cdEmployee
    Dim radrsTotal As ADODB.Recordset

    'Construct
    Set lobjdbEmployee = CreateObject("dbEmployee.cdEmployee")
    Set radrsSummary = lobjdbEmployee.GetEmployeeExcusedTimeSummary(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    Set radrsDetails = lobjdbEmployee.GetEmployeeExcusedTimeDetails(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    Set radrsTotal = lobjdbEmployee.GetEmployeeExcusedTimeSummaryTotal(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    
    '//expecting 1 row from SummaryTotal
    
    If Not radrsSummary Is Nothing Then
        If RecordsExist(radrsTotal) Then
            radrsTotal.MoveFirst
            
            radrsSummary.AddNew
            radrsSummary.Fields("CHARGE_DATE").Value = "12/31/9999"
            radrsSummary.Fields("RPTD_HRS").Value = radrsTotal.Fields("RPTD_HRS").Value
            radrsSummary.Fields("EXCUSED_HRS").Value = radrsTotal.Fields("EXCUSED_HRS").Value
            radrsSummary.Fields("UNEXCUSED_HRS").Value = radrsTotal.Fields("UNEXCUSED_HRS").Value
            radrsSummary.Update
        End If
    End If
    
    'Destroy
    Set lobjdbEmployee = Nothing
    Set radrsTotal = Nothing
    
    Exit Sub
GetEmployeeExcusedTimeInformation2_EH:
    Set lobjdbEmployee = Nothing
    Set radrsTotal = Nothing

    TrsRaiseError m_strCMODULENAME, "GetEmployeeExcusedTimeInformation2", , , , , True
End Sub
'//for TRIS 6.7+ clients
Public Sub GetEmployeeExcusedTimeInformation3(ByVal vlEmployeeNumber As Long _
                    , ByRef radrsSummary As ADODB.Recordset _
                    , ByRef radrsDetails As ADODB.Recordset _
                    , ByRef radrsTotal As ADODB.Recordset _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999#)
On Error GoTo eh
    'Declare
    Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Construct
    Set lobjdbEmployee = CreateObject("dbEmployee.cdEmployee")
    Set radrsSummary = lobjdbEmployee.GetEmployeeExcusedTimeSummary(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    Set radrsDetails = lobjdbEmployee.GetEmployeeExcusedTimeDetails(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    Set radrsTotal = lobjdbEmployee.GetEmployeeExcusedTimeSummaryTotal(vlEmployeeNumber, vdtStartDate, vdtEndDate)
        
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Sub
eh:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME, "GetEmployeeExcusedTimeInformation3", , , , , True
End Sub

Public Function GetEmployeeFMLADetails(ByVal vlEmployeeNumber As Long _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999# _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeFMLADetails_EH
'Declare
Dim lobjdbEmployee As dbEmployee.cdEmployee

    'Construct
    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployee")
    Set GetEmployeeFMLADetails = lobjdbEmployee.GetEmployeeFMLADetails( _
                vlEmployeeNumber, vdtStartDate, vdtEndDate)
    'Destroy
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeeFMLADetails_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeeFMLADetails" _
        , , , , _
        , True
End Function
'''Public Function GetPSPhoneBook() As ADODB.Recordset
'''On Error GoTo GetPSPhoneBook_EH
''''Declare
'''Dim lsSQL As String, lsConnectionString As String
'''Dim lobjConnectionMgr As util_DBConnectMgr.clsuDBConnectMgr
'''
'''    Set GetPSPhoneBook = Nothing
'''
'''    lsSQL = "SELECT  INTEGER(J.EMPLID) AS EMP_NUM, N.NAME AS EMP_NAME, P.PHONE AS EMP_PHONE" & _
'''        "       , E.CONTACT_NAME, E.PRIMARY_CONTACT, E.PHONE AS CNTCT_PHONE" & _
'''        " FROM  PS_JOB J" & _
'''        "       INNER JOIN PS_NAMES N" & _
'''        "       ON J.EMPL_RCD = 0" & _
'''        "       AND J.EMPL_STATUS IN ('A', 'L', 'P')" & _
'''        "       AND J.DEPTID NOT IN ('CONS', 'CONSCAL')" & _
'''        "       AND J.EFFDT = (SELECT MAX(J2.EFFDT) FROM PS_JOB J2" & _
'''        "                       WHERE   J2.EMPLID = J.EMPLID" & _
'''        "                       AND     J2.EMPL_RCD = J.EMPL_RCD" & _
'''        "                       AND     J2.EFFDT <= SYSDATE)" & _
'''        "       AND J.EFFSEQ = (SELECT MAX(J3.EFFSEQ) FROM PS_JOB J3" & _
'''        "                       WHERE   J3.EMPLID = J.EMPLID" & _
'''        "                       AND     J3.EMPL_RCD = J.EMPL_RCD" & _
'''        "                       AND     J3.EFFDT = J.EFFDT)" & _
'''        "       AND J.EMPLID = N.EMPLID" & _
'''        "       AND N.NAME_TYPE = 'PRI'" & _
'''        "       AND N.EFFDT <= SYSDATE" & _
'''        "       INNER JOIN PS_EMERGENCY_CNTCT E" & _
'''        "       ON J.EMPLID = E.EMPLID" & _
'''        "       INNER JOIN PS_PERSONAL_PHONE P" & _
'''        "       ON J.EMPLID = P.EMPLID"
'''
'''    Set lobjConnectionMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
'''    lsConnectionString = lobjConnectionMgr.ConnectionString("PeopleSoft")
'''    Set lobjConnectionMgr = Nothing
'''
'''    Set GetPSPhoneBook = New ADODB.Recordset
'''    GetPSPhoneBook.Open lsSQL, lsConnectionString _
'''        , adOpenStatic _
'''        , adLockReadOnly
'''    Set GetPSPhoneBook.ActiveConnection = Nothing
'''
'''    If Not (GetPSPhoneBook Is Nothing) Then
'''        If Not (GetPSPhoneBook.EOF And _
'''            GetPSPhoneBook.BOF) Then
'''            GetPSPhoneBook.MoveFirst
'''        End If
'''    End If
'''
'''    Exit Function
'''GetPSPhoneBook_EH:
'''    Set lobjConnectionMgr = Nothing
'''
'''    'Call Standard Module Function
'''    'Raise the appropriate error message depending on the error number
'''    Select Case (Err.Number)
'''            Case trErrutOpenConnection
'''                TrsRaiseError m_strCMODULENAME _
'''                    , "GetPSPhoneBook"
'''            Case trErrdb_ExecuteSQL_F
'''                TrsRaiseError m_strCMODULENAME _
'''                    , "GetPSPhoneBook" _
'''                    , trErrdb_ExecuteSQL_F _
'''                    , , _
'''                    , DescArgsFrom(lsSQL) _
'''                    , True
'''            Case Else
'''                TrsRaiseError m_strCMODULENAME _
'''                    , "GetPSPhoneBook" _
'''                    , trErrdbGetExcusedTime _
'''                    , , _
'''                    , DescArgsFrom(lsSQL) _
'''                    , True
'''    End Select
'''
'''End Function
'End Prem TRIS 5.7

Public Function IsCallawaySecurity( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
                       
     Dim l_obj As busDepartment.cbsDepartment
     Dim l_strDept As String
     
On Error GoTo eh
    
    Set l_obj = CreateObject("busDepartment.cbsDepartment")
         
    l_strDept = EmployeeDepartment( _
        v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
    
    IsCallawaySecurity = DoesRuleApply(212, trDepartmentNumber, l_strDept)
    
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsCallawaySecurity", , , , , True
    
            
End Function



