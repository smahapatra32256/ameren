VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsEmployeeProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function GetDepartment ( _
	ByVal v_lngEmpNum As Long,  _
	ByVal v_datDate As Date _
	)  As  String
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    GetDepartment = l_objcbsEmployee.GetDepartment(v_lngEmpNum, v_datDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetDepartment", , , , , True

End Function


Public Function GetWorkReasons ( _
	ByVal v_lngEmpNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetWorkReasons = l_objcbsEmployee.GetWorkReasons(v_lngEmpNum)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetWorkReasons", , , , , True

End Function


Public Function IsValidEmployee ( _
	ByVal v_lngEmpNum As Long _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsValidEmployee = l_objcbsEmployee.IsValidEmployee(v_lngEmpNum)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsValidEmployee", , , , , True

End Function


Public Function IsSubstationOrTransmissionEmployee ( _
	ByVal v_strUnionCode As String,  _
	ByVal v_strDeptNum As String _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsSubstationOrTransmissionEmployee = l_objcbsEmployee.IsSubstationOrTransmissionEmployee(v_strUnionCode, v_strDeptNum)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsSubstationOrTransmissionEmployee", , , , , True

End Function


Public Function GetEmployeesInDepartment ( _
	ByVal v_aryDepartments As Variant,  _
	Optional ByRef v_bolActiveEmployeesOnly As Boolean = "False",  _
	Optional ByRef v_datStartDate As Date = "01/01/100",  _
	Optional ByRef v_datEndDate As Date = "01/01/100" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeesInDepartment = l_objcbsEmployee.GetEmployeesInDepartment(v_aryDepartments, v_bolActiveEmployeesOnly, v_datStartDate, v_datEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeesInDepartment", , , , , True

End Function


Public Function GetPrimaryWorkGroups ( _
	ByVal v_aryEmployees As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByRef bln_PrimaryOnly As Boolean = "False" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetPrimaryWorkGroups = l_objcbsEmployee.GetPrimaryWorkGroups(v_aryEmployees, v_datStartDate, v_datEndDate, bln_PrimaryOnly)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetPrimaryWorkGroups", , , , , True

End Function


Public Function GetAllEmployeeInfo ( _
	ByVal v_aryEmployees As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	ByVal v_blnPrimaryOnly As Boolean,  _
	ByRef r_rsEmployee As ADODB.Recordset,  _
	ByRef r_rsPrimaryWorkGroups As ADODB.Recordset,  _
	ByRef r_rsTimeEntryLockStatus As ADODB.Recordset,  _
	ByRef r_rsCaseNumbers As ADODB.Recordset,  _
	ByRef r_rsEmpBanks As ADODB.Recordset,  _
	ByRef r_rsEmpOTBanks As ADODB.Recordset,  _
	ByRef r_rsQuarterlyBanks As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    GetAllEmployeeInfo = l_objcbsEmployee.GetAllEmployeeInfo(v_aryEmployees, v_datStartDate, v_datEndDate, v_blnPrimaryOnly, r_rsEmployee, r_rsPrimaryWorkGroups, r_rsTimeEntryLockStatus, r_rsCaseNumbers, r_rsEmpBanks, r_rsEmpOTBanks, r_rsQuarterlyBanks)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetAllEmployeeInfo", , , , , True

End Function


Public Function GetEmployee ( _
	ByVal v_aryEmployees As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployee = l_objcbsEmployee.GetEmployee(v_aryEmployees)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployee", , , , , True

End Function


Public Function GetEmployeeBanks ( _
	ByVal v_aryEmployees As Variant,  _
	ByVal v_datCurrDate As Date,  _
	ByRef r_rsBanks As ADODB.Recordset,  _
	ByRef r_rsOTBanks As ADODB.Recordset,  _
	ByRef r_rsQuarterlyOTBanks As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    GetEmployeeBanks = l_objcbsEmployee.GetEmployeeBanks(v_aryEmployees, v_datCurrDate, r_rsBanks, r_rsOTBanks, r_rsQuarterlyOTBanks)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeBanks", , , , , True

End Function


Public Function GetAllEmployees ( _
	Optional ByRef v_bolActiveEmployeesOnly As Boolean = "True" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetAllEmployees = l_objcbsEmployee.GetAllEmployees(v_bolActiveEmployeesOnly)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetAllEmployees", , , , , True

End Function


Public Function GetBasicEmployeeInformation ( _
	ByVal v_lngEmpNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetBasicEmployeeInformation = l_objcbsEmployee.GetBasicEmployeeInformation(v_lngEmpNum)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetBasicEmployeeInformation", , , , , True

End Function


Public Function GetEmployeePrimaryWorkGroup ( _
	ByVal v_lngEmpNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeePrimaryWorkGroup = l_objcbsEmployee.GetEmployeePrimaryWorkGroup(v_lngEmpNum)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeePrimaryWorkGroup", , , , , True

End Function


Public Function GetEmployeeWorkGroups ( _
	ByVal v_lngEmpNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeWorkGroups = l_objcbsEmployee.GetEmployeeWorkGroups(v_lngEmpNum)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeWorkGroups", , , , , True

End Function


Public Function GetEmployeeReliefGroups ( _
	ByVal v_lngEmpNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeReliefGroups = l_objcbsEmployee.GetEmployeeReliefGroups(v_lngEmpNum)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeReliefGroups", , , , , True

End Function


Public Function GetEmployeeQualifications ( _
	ByVal v_lngEmpNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeQualifications = l_objcbsEmployee.GetEmployeeQualifications(v_lngEmpNum)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeQualifications", , , , , True

End Function


Public Function GetEmployeeSkills ( _
	ByVal v_lngEmpNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeSkills = l_objcbsEmployee.GetEmployeeSkills(v_lngEmpNum)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeSkills", , , , , True

End Function


Public Function EmployeeBenefitPackage ( _
	ByVal v_datJobEffDate As Date,  _
	ByVal v_strManagementBenPkg As String,  _
	ByVal v_strUnionCode As String,  _
	ByVal v_strPriorUnionCode As String,  _
	Optional ByVal v_strExecContractInd As String = "NA",  _
	Optional ByVal v_strPriorExecContractInd As String = "NA",  _
	Optional ByVal v_strUnionBenPkg As String = "NA",  _
	Optional ByVal v_strPriorUnionBenPkg As String = "NA",  _
	Optional ByVal v_strChgDate As Date = "01/01/100" _
	)  As  String
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    EmployeeBenefitPackage = l_objcbsEmployee.EmployeeBenefitPackage(v_datJobEffDate, v_strManagementBenPkg, v_strUnionCode, v_strPriorUnionCode, v_strExecContractInd, v_strPriorExecContractInd, v_strUnionBenPkg, v_strPriorUnionBenPkg, v_strChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "EmployeeBenefitPackage", , , , , True

End Function


Public Function GetLockStatus ( _
	ByVal v_lngEmpNum As Long,  _
	ByVal v_datChargeDate As Date,  _
	ByVal v_rsEmpLockData As ADODB.Recordset,  _
	ByVal v_rsCollPeriods As ADODB.Recordset _
	)  As  String
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    GetLockStatus = l_objcbsEmployee.GetLockStatus(v_lngEmpNum, v_datChargeDate, v_rsEmpLockData, v_rsCollPeriods)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetLockStatus", , , , , True

End Function


Public Function IsExecutive ( _
	ByVal v_datJobEffDate As Date,  _
	ByVal v_strUnionCode As String,  _
	ByVal v_strPriorUnionCode As String,  _
	Optional ByVal v_strExecContractInd As String = "NA",  _
	Optional ByVal v_strPriorExecContractInd As String = "NA",  _
	Optional ByVal v_strChgDate As Date = "01/01/100" _
	)  As  String
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsExecutive = l_objcbsEmployee.IsExecutive(v_datJobEffDate, v_strUnionCode, v_strPriorUnionCode, v_strExecContractInd, v_strPriorExecContractInd, v_strChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsExecutive", , , , , True

End Function


Public Function EmployeeCompany ( _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_strCompanyCode As String = "NA",  _
	Optional ByVal v_strPriorCompanyCode As String = "NA",  _
	Optional ByVal v_strChgDate As Date = "01/01/100" _
	)  As  String
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    EmployeeCompany = l_objcbsEmployee.EmployeeCompany(v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_strCompanyCode, v_strPriorCompanyCode, v_strChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "EmployeeCompany", , , , , True

End Function


Public Function IsAmerenCIPS ( _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_strCompanyCode As String = "NA",  _
	Optional ByVal v_strPriorCompanyCode As String = "NA",  _
	Optional ByVal v_strChgDate As Date = "01/01/100" _
	)  As  String
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsAmerenCIPS = l_objcbsEmployee.IsAmerenCIPS(v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_strCompanyCode, v_strPriorCompanyCode, v_strChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsAmerenCIPS", , , , , True

End Function


Public Function IsAmerenCIPSorCILCO ( _
	ByVal v_lngRuleNum As Long,  _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_strCompanyCode As String = "NA",  _
	Optional ByVal v_strPriorCompanyCode As String = "NA",  _
	Optional ByVal v_datChgDate As Date = "01/01/100",  _
	Optional ByVal v_strUnionCode As String = "Z#",  _
	Optional ByVal v_strPayrollGroupCd As String = "Z#" _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsAmerenCIPSorCILCO = l_objcbsEmployee.IsAmerenCIPSorCILCO(v_lngRuleNum, v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_strCompanyCode, v_strPriorCompanyCode, v_datChgDate, v_strUnionCode, v_strPayrollGroupCd)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsAmerenCIPSorCILCO", , , , , True

End Function


Public Function EmployeeUnion ( _
	ByVal v_datJobEffDate As Date,  _
	ByVal v_strUnionCode As String,  _
	ByVal v_strPriorUnionCode As String,  _
	Optional ByVal v_strChgDate As Date = "01/01/100" _
	)  As  String
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    EmployeeUnion = l_objcbsEmployee.EmployeeUnion(v_datJobEffDate, v_strUnionCode, v_strPriorUnionCode, v_strChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "EmployeeUnion", , , , , True

End Function


Public Function EmployeeDepartment ( _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_strChgDate As Date = "01/01/100" _
	)  As  String
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    EmployeeDepartment = l_objcbsEmployee.EmployeeDepartment(v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_strChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "EmployeeDepartment", , , , , True

End Function


Public Function IsEmployeeEMPRV ( _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_strInterfaceCode As String = "NA",  _
	Optional ByVal v_strPriorInterfaceCode As String = "NA",  _
	Optional ByVal v_datChgDate As Date = "01/01/100" _
	)  As  String
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsEmployeeEMPRV = l_objcbsEmployee.IsEmployeeEMPRV(v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_strInterfaceCode, v_strPriorInterfaceCode, v_datChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsEmployeeEMPRV", , , , , True

End Function


Public Function EmployeeJobNum ( _
	ByVal v_datJobEffDate As Date,  _
	ByVal v_strJobNum As String,  _
	ByVal v_strPriorJobNum As String,  _
	Optional ByRef v_datPriorJobEffDate As Date = "01/01/100",  _
	Optional ByVal v_strChgDate As Date = "01/01/100" _
	)  As  String
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    EmployeeJobNum = l_objcbsEmployee.EmployeeJobNum(v_datJobEffDate, v_strJobNum, v_strPriorJobNum, v_datPriorJobEffDate, v_strChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "EmployeeJobNum", , , , , True

End Function


Public Function GetCaseGroupsAndNumbers ( _
	ByVal v_aryEmployees As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetCaseGroupsAndNumbers = l_objcbsEmployee.GetCaseGroupsAndNumbers(v_aryEmployees, v_datStartDate, v_datEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetCaseGroupsAndNumbers", , , , , True

End Function


Public Function PartialHourMinutes ( _
	ByVal v_strDept As String,  _
	ByVal v_blnIsExecutive As Boolean,  _
	Optional ByVal v_strCompany As String = "NA" _
	)  As  Integer
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    PartialHourMinutes = l_objcbsEmployee.PartialHourMinutes(v_strDept, v_blnIsExecutive, v_strCompany)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "PartialHourMinutes", , , , , True

End Function


Public Function IsAmerenCILCO ( _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_strCompanyCode As String = "NA",  _
	Optional ByVal v_strPriorCompanyCode As String = "NA",  _
	Optional ByVal v_datChgDate As Date = "01/01/100" _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsAmerenCILCO = l_objcbsEmployee.IsAmerenCILCO(v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_strCompanyCode, v_strPriorCompanyCode, v_datChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsAmerenCILCO", , , , , True

End Function


Public Function IsAmerenCILCODispatch ( _
	ByVal v_lngEmpNum As Long,  _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_datChgDate As Date = "01/01/100" _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsAmerenCILCODispatch = l_objcbsEmployee.IsAmerenCILCODispatch(v_lngEmpNum, v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsAmerenCILCODispatch", , , , , True

End Function


Public Function IsEdwardsOperations ( _
	ByVal v_lngEmpNum As Long,  _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_datChgDate As Date = "01/01/100" _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsEdwardsOperations = l_objcbsEmployee.IsEdwardsOperations(v_lngEmpNum, v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsEdwardsOperations", , , , , True

End Function


Public Function IsCoffeenOperations ( _
	ByVal v_lngEmpNum As Long,  _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_datChgDate As Date = "01/01/100" _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsCoffeenOperations = l_objcbsEmployee.IsCoffeenOperations(v_lngEmpNum, v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsCoffeenOperations", , , , , True

End Function


Public Function IsCoffeenMaintenance ( _
	ByVal v_lngEmpNum As Long,  _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_datChgDate As Date = "01/01/100" _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsCoffeenMaintenance = l_objcbsEmployee.IsCoffeenMaintenance(v_lngEmpNum, v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsCoffeenMaintenance", , , , , True

End Function


Public Function IsAmerenCILCOLocal8Employee ( _
	ByVal v_datJobEffDate As Date,  _
	ByVal v_strUnionCode As String,  _
	ByVal v_strPriorUnionCode As String,  _
	Optional ByVal v_strChgDate As Date = "01/01/100" _
	)  As  String
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsAmerenCILCOLocal8Employee = l_objcbsEmployee.IsAmerenCILCOLocal8Employee(v_datJobEffDate, v_strUnionCode, v_strPriorUnionCode, v_strChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsAmerenCILCOLocal8Employee", , , , , True

End Function


Public Function GetQualsAndSkillsForEmployees ( _
	ByVal vadrsEmployees As ADODB.Recordset _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetQualsAndSkillsForEmployees = l_objcbsEmployee.GetQualsAndSkillsForEmployees(vadrsEmployees)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetQualsAndSkillsForEmployees", , , , , True

End Function


Public Function GetQualsAndSkillsForEmployees2 ( _
	ByVal vadrsEmployees As ADODB.Recordset,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetQualsAndSkillsForEmployees2 = l_objcbsEmployee.GetQualsAndSkillsForEmployees2(vadrsEmployees, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetQualsAndSkillsForEmployees2", , , , , True

End Function


Public Function GetExpiringQAndSByLocation ( _
	ByVal viLocation As Integer,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetExpiringQAndSByLocation = l_objcbsEmployee.GetExpiringQAndSByLocation(viLocation, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetExpiringQAndSByLocation", , , , , True

End Function


Public Function GetExpiringQAndSByDept ( _
	ByVal vsDepartment As String,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetExpiringQAndSByDept = l_objcbsEmployee.GetExpiringQAndSByDept(vsDepartment, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetExpiringQAndSByDept", , , , , True

End Function


Public Function GetExpiringQAndSByWG ( _
	ByVal vlWorkgroup As Long,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetExpiringQAndSByWG = l_objcbsEmployee.GetExpiringQAndSByWG(vlWorkgroup, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetExpiringQAndSByWG", , , , , True

End Function


Public Function GetEmployeeSeniority ( _
	ByVal vlEmployeeNumber As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeSeniority = l_objcbsEmployee.GetEmployeeSeniority(vlEmployeeNumber)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeSeniority", , , , , True

End Function


Public Function GetEmployeeFrozenSeniority ( _
	ByVal vlEmployeeNumber As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeFrozenSeniority = l_objcbsEmployee.GetEmployeeFrozenSeniority(vlEmployeeNumber)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeFrozenSeniority", , , , , True

End Function


Public Sub GetAllEmployeeSeniorityInformation ( _
	ByVal vlEmployeeNumber As Long,  _
	ByRef vbIsAGencoEmployee As Boolean,  _
	ByRef radrsSeniority As ADODB.Recordset,  _
	ByRef radrsFrozenSeniority As ADODB.Recordset _
	)  
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    l_objcbsEmployee.GetAllEmployeeSeniorityInformation vlEmployeeNumber, vbIsAGencoEmployee, radrsSeniority, radrsFrozenSeniority
    
    Set l_objcbsEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetAllEmployeeSeniorityInformation", , , , , True

End Sub


Public Function IsAGencoEmployee ( _
	ByVal v_strCmpyCd As String,  _
	ByVal v_strUnionCd As String,  _
	ByVal v_strPayPeriodCd As String,  _
	ByVal v_datChargeDate As Date _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsAGencoEmployee = l_objcbsEmployee.IsAGencoEmployee(v_strCmpyCd, v_strUnionCd, v_strPayPeriodCd, v_datChargeDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsAGencoEmployee", , , , , True

End Function


Public Function GetEmployeeSeniorityForDept ( _
	ByVal vsDepartment As String,  _
	ByVal vbActiveEmployeesOnly As Boolean _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeSeniorityForDept = l_objcbsEmployee.GetEmployeeSeniorityForDept(vsDepartment, vbActiveEmployeesOnly)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeSeniorityForDept", , , , , True

End Function


Public Function GetEmployeeFrznSnrtyForDept ( _
	ByVal vsDepartment As String,  _
	ByVal vbActiveEmployeesOnly As Boolean _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeFrznSnrtyForDept = l_objcbsEmployee.GetEmployeeFrznSnrtyForDept(vsDepartment, vbActiveEmployeesOnly)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeFrznSnrtyForDept", , , , , True

End Function


Public Sub GetAllEmployeeSeniorityForDept ( _
	ByVal vsDepartment As String,  _
	ByVal vbActiveEmployeesOnly As Boolean,  _
	ByRef vbIsAGencoDepartment As Boolean,  _
	ByRef radrsSeniority As ADODB.Recordset,  _
	ByRef radrsFrozenSeniority As ADODB.Recordset _
	)  
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    l_objcbsEmployee.GetAllEmployeeSeniorityForDept vsDepartment, vbActiveEmployeesOnly, vbIsAGencoDepartment, radrsSeniority, radrsFrozenSeniority
    
    Set l_objcbsEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetAllEmployeeSeniorityForDept", , , , , True

End Sub


Public Function GetEmployeeSeniorityForWG ( _
	ByVal vlWorkgroup As Long,  _
	ByVal vbActiveEmployeesOnly As Boolean _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeSeniorityForWG = l_objcbsEmployee.GetEmployeeSeniorityForWG(vlWorkgroup, vbActiveEmployeesOnly)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeSeniorityForWG", , , , , True

End Function


Public Function GetEmployeeFrznSnrtyForWG ( _
	ByVal vlWorkgroup As Long,  _
	ByVal vbActiveEmployeesOnly As Boolean _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeFrznSnrtyForWG = l_objcbsEmployee.GetEmployeeFrznSnrtyForWG(vlWorkgroup, vbActiveEmployeesOnly)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeFrznSnrtyForWG", , , , , True

End Function


Public Sub GetAllEmployeeSeniorityForWG ( _
	ByVal vlWorkgroup As Long,  _
	ByVal vbActiveEmployeesOnly As Boolean,  _
	ByVal vbIsAGencoWorkgroup As Boolean,  _
	ByRef radrsSeniority As ADODB.Recordset,  _
	ByRef radrsFrozenSeniority As ADODB.Recordset _
	)  
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    l_objcbsEmployee.GetAllEmployeeSeniorityForWG vlWorkgroup, vbActiveEmployeesOnly, vbIsAGencoWorkgroup, radrsSeniority, radrsFrozenSeniority
    
    Set l_objcbsEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetAllEmployeeSeniorityForWG", , , , , True

End Sub


Public Sub GetAllEmployeeSeniorityForWG2 ( _
	ByVal vlWorkgroup As Long,  _
	ByVal vbActiveEmployeesOnly As Boolean,  _
	ByRef vbIsAGencoWorkgroup As Boolean,  _
	ByRef radrsSeniority As ADODB.Recordset,  _
	ByRef radrsFrozenSeniority As ADODB.Recordset _
	)  
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    l_objcbsEmployee.GetAllEmployeeSeniorityForWG2 vlWorkgroup, vbActiveEmployeesOnly, vbIsAGencoWorkgroup, radrsSeniority, radrsFrozenSeniority
    
    Set l_objcbsEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetAllEmployeeSeniorityForWG2", , , , , True

End Sub


Public Function GetEmployeeCanvassJob ( _
	ByVal vadrsEmployees As ADODB.Recordset _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeCanvassJob = l_objcbsEmployee.GetEmployeeCanvassJob(vadrsEmployees)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeCanvassJob", , , , , True

End Function


Public Function GetEmployeeCanvassJob2 ( _
	ByVal vadrsEmployees As ADODB.Recordset,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeCanvassJob2 = l_objcbsEmployee.GetEmployeeCanvassJob2(vadrsEmployees, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeCanvassJob2", , , , , True

End Function


Public Function GetEmployeeIneligibilities ( _
	ByVal vadrsEmployees As ADODB.Recordset _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeIneligibilities = l_objcbsEmployee.GetEmployeeIneligibilities(vadrsEmployees)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeIneligibilities", , , , , True

End Function


Public Function GetEmployeeIneligibilities2 ( _
	ByVal vadrsEmployees As ADODB.Recordset,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeIneligibilities2 = l_objcbsEmployee.GetEmployeeIneligibilities2(vadrsEmployees, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeIneligibilities2", , , , , True

End Function


Public Function GetVolunteerInformation ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetVolunteerInformation = l_objcbsEmployee.GetVolunteerInformation()
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetVolunteerInformation", , , , , True

End Function


Public Function GetEmployeesInProject ( _
	ByVal vlProjectID As Long,  _
	Optional ByVal vdtStartDate As Date = "01/01/100",  _
	Optional ByVal vdtEndDate As Date = "12/31/9999" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeesInProject = l_objcbsEmployee.GetEmployeesInProject(vlProjectID, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeesInProject", , , , , True

End Function


Public Function GetEmployeesForProjectByWorkgroup ( _
	ByVal vlProjectID As Long,  _
	ByVal vlWorkgroupNumber As Long,  _
	Optional ByVal vdtStartDate As Date = "01/01/100",  _
	Optional ByVal vdtEndDate As Date = "12/31/9999" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeesForProjectByWorkgroup = l_objcbsEmployee.GetEmployeesForProjectByWorkgroup(vlProjectID, vlWorkgroupNumber, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeesForProjectByWorkgroup", , , , , True

End Function


Public Function GetEmployeesForProjectByLocation ( _
	ByVal vlProjectID As Long,  _
	ByVal vlLocationCode As Long,  _
	Optional ByVal vdtStartDate As Date = "01/01/100",  _
	Optional ByVal vdtEndDate As Date = "12/31/9999" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeesForProjectByLocation = l_objcbsEmployee.GetEmployeesForProjectByLocation(vlProjectID, vlLocationCode, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeesForProjectByLocation", , , , , True

End Function


Public Function GetEmployeesForProjectByCraft ( _
	ByVal vlProjectID As Long,  _
	ByVal vsCraftNumber As String,  _
	Optional ByVal vdtStartDate As Date = "01/01/100",  _
	Optional ByVal vdtEndDate As Date = "12/31/9999" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeesForProjectByCraft = l_objcbsEmployee.GetEmployeesForProjectByCraft(vlProjectID, vsCraftNumber, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeesForProjectByCraft", , , , , True

End Function


Public Function GetEmployeeECWorkReasons ( _
	ByVal vbActiveEmployeesOnly As Boolean _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeECWorkReasons = l_objcbsEmployee.GetEmployeeECWorkReasons(vbActiveEmployeesOnly)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeECWorkReasons", , , , , True

End Function


Public Function GetEmployeeSickLeaveInformation ( _
	ByVal vlEmployeeNumber As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeSickLeaveInformation = l_objcbsEmployee.GetEmployeeSickLeaveInformation(vlEmployeeNumber)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeSickLeaveInformation", , , , , True

End Function


Public Function GetEmployeeSickLeaveInformation2 ( _
	ByVal vlEmployeeNumber As Long,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeSickLeaveInformation2 = l_objcbsEmployee.GetEmployeeSickLeaveInformation2(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeSickLeaveInformation2", , , , , True

End Function


Public Sub GetEmployeeAllowedTimeInformation ( _
	ByVal vlEmployeeNumber As Long,  _
	ByRef radrsAllowedTimeDetails As ADODB.Recordset,  _
	ByRef radrsAllowedTimeTotals As ADODB.Recordset,  _
	Optional ByVal vdtStartDate As Date = "01/01/100",  _
	Optional ByVal vdtEndDate As Date = "12/31/9999" _
	)  
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    l_objcbsEmployee.GetEmployeeAllowedTimeInformation vlEmployeeNumber, radrsAllowedTimeDetails, radrsAllowedTimeTotals, vdtStartDate, vdtEndDate
    
    Set l_objcbsEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeAllowedTimeInformation", , , , , True

End Sub


Public Sub GetEmployeeCertificationInformation ( _
	ByVal vlEmployeeNumber As Long,  _
	ByRef radrsSummary As ADODB.Recordset,  _
	ByRef radrsDetails As ADODB.Recordset,  _
	Optional ByVal vdtStartDate As Date = "01/01/100",  _
	Optional ByVal vdtEndDate As Date = "12/31/9999" _
	)  
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    l_objcbsEmployee.GetEmployeeCertificationInformation vlEmployeeNumber, radrsSummary, radrsDetails, vdtStartDate, vdtEndDate
    
    Set l_objcbsEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeCertificationInformation", , , , , True

End Sub


Public Function GetEmployeeExcusedTimeInformation ( _
	ByVal vlEmployeeNumber As Long,  _
	Optional ByVal vdtStartDate As Date = "01/01/100",  _
	Optional ByVal vdtEndDate As Date = "12/31/9999" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeExcusedTimeInformation = l_objcbsEmployee.GetEmployeeExcusedTimeInformation(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeExcusedTimeInformation", , , , , True

End Function


Public Function GetEmployeePSContactInformation ( _
	ByVal vlEmployee As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeePSContactInformation = l_objcbsEmployee.GetEmployeePSContactInformation(vlEmployee)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeePSContactInformation", , , , , True

End Function


Public Sub GetEmployeeCertificationInformation2 ( _
	ByVal vlEmployeeNumber As Long,  _
	ByRef radrsSummaryByYear As ADODB.Recordset,  _
	ByRef radrsSummaryByDate As ADODB.Recordset,  _
	ByRef radrsDetails As ADODB.Recordset,  _
	Optional ByVal vdtStartDate As Date = "01/01/100",  _
	Optional ByVal vdtEndDate As Date = "12/31/9999" _
	)  
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    l_objcbsEmployee.GetEmployeeCertificationInformation2 vlEmployeeNumber, radrsSummaryByYear, radrsSummaryByDate, radrsDetails, vdtStartDate, vdtEndDate
    
    Set l_objcbsEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeCertificationInformation2", , , , , True

End Sub


Public Sub GetEmployeeExcusedTimeInformation2 ( _
	ByVal vlEmployeeNumber As Long,  _
	ByRef radrsSummary As ADODB.Recordset,  _
	ByRef radrsDetails As ADODB.Recordset,  _
	Optional ByVal vdtStartDate As Date = "01/01/100",  _
	Optional ByVal vdtEndDate As Date = "12/31/9999" _
	)  
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    l_objcbsEmployee.GetEmployeeExcusedTimeInformation2 vlEmployeeNumber, radrsSummary, radrsDetails, vdtStartDate, vdtEndDate
    
    Set l_objcbsEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeExcusedTimeInformation2", , , , , True

End Sub


Public Sub GetEmployeeExcusedTimeInformation3 ( _
	ByVal vlEmployeeNumber As Long,  _
	ByRef radrsSummary As ADODB.Recordset,  _
	ByRef radrsDetails As ADODB.Recordset,  _
	ByRef radrsTotal As ADODB.Recordset,  _
	Optional ByVal vdtStartDate As Date = "01/01/100",  _
	Optional ByVal vdtEndDate As Date = "12/31/9999" _
	)  
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    l_objcbsEmployee.GetEmployeeExcusedTimeInformation3 vlEmployeeNumber, radrsSummary, radrsDetails, radrsTotal, vdtStartDate, vdtEndDate
    
    Set l_objcbsEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeExcusedTimeInformation3", , , , , True

End Sub


Public Function GetEmployeeFMLADetails ( _
	ByVal vlEmployeeNumber As Long,  _
	Optional ByVal vdtStartDate As Date = "01/01/100",  _
	Optional ByVal vdtEndDate As Date = "12/31/9999" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    Set GetEmployeeFMLADetails = l_objcbsEmployee.GetEmployeeFMLADetails(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "GetEmployeeFMLADetails", , , , , True

End Function


Public Function IsCallawaySecurity ( _
	ByVal v_lngEmpNum As Long,  _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_datChgDate As Date = "01/01/100" _
	)  As  Boolean
  
    Dim l_objcbsEmployee As busEmployee.cbsEmployee
    
On Error GoTo eh

    Set l_objcbsEmployee = CreateObject("busEmployee.cbsEmployee")
    
    IsCallawaySecurity = l_objcbsEmployee.IsCallawaySecurity(v_lngEmpNum, v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_datChgDate)
    
    Set l_objcbsEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployee = Nothing

    TrsRaiseError "busEmployee.cbsEmployee", "IsCallawaySecurity", , , , , True

End Function