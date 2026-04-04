VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cboEmployeeProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Sub RestoreWithState ( _
	ByRef r_rsEmployee As ADODB.Recordset,  _
	ByRef r_rsEmpWorkgroups As ADODB.Recordset _
	)  
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    l_objcboEmployee.RestoreWithState r_rsEmployee, r_rsEmpWorkgroups
    
    Set l_objcboEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "RestoreWithState", , , , , True

End Sub


Public Sub Restore ( _
	ByVal v_lngEmpId As Long,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByVal v_blnDeferLoad As Boolean = "True" _
	)  
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    l_objcboEmployee.Restore v_lngEmpId, v_datStartDate, v_datEndDate, v_blnDeferLoad
    
    Set l_objcboEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "Restore", , , , , True

End Sub


Public Sub Release ( _
	)  
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    l_objcboEmployee.Release 
    
    Set l_objcboEmployee = Nothing
    
    Exit Sub
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "Release", , , , , True

End Sub


Public Function PartialHourMinutes ( _
	ByVal v_datChargeDate As Date _
	)  As  Integer
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    PartialHourMinutes = l_objcboEmployee.PartialHourMinutes(v_datChargeDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "PartialHourMinutes", , , , , True

End Function


Public Function GetWorkgroups ( _
	)  As  ADODB.Recordset
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    Set GetWorkgroups = l_objcboEmployee.GetWorkgroups()
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "GetWorkgroups", , , , , True

End Function


Public Function GetPrimaryWorkgroupNumber ( _
	ByVal v_strChargeDate As Date _
	)  As  Long
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    GetPrimaryWorkgroupNumber = l_objcboEmployee.GetPrimaryWorkgroupNumber(v_strChargeDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "GetPrimaryWorkgroupNumber", , , , , True

End Function


Public Function GetProjectId ( _
	ByVal v_lngWgNum As Long _
	)  As  Long
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    GetProjectId = l_objcboEmployee.GetProjectId(v_lngWgNum)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "GetProjectId", , , , , True

End Function


Public Function GetPPMProjectPeriod ( _
	ByVal v_lngWgNum As Long _
	)  As  Variant
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    GetPPMProjectPeriod = l_objcboEmployee.GetPPMProjectPeriod(v_lngWgNum)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "GetPPMProjectPeriod", , , , , True

End Function


Public Function GetWorkgroupLocationCode ( _
	ByVal v_lngWgNum As Long _
	)  As  Integer
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    GetWorkgroupLocationCode = l_objcboEmployee.GetWorkgroupLocationCode(v_lngWgNum)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "GetWorkgroupLocationCode", , , , , True

End Function


Public Function IsEntitledToSaturdayPremium ( _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsEntitledToSaturdayPremium = l_objcboEmployee.IsEntitledToSaturdayPremium()
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsEntitledToSaturdayPremium", , , , , True

End Function


Public Function IsEntitledToSundayPremium ( _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsEntitledToSundayPremium = l_objcboEmployee.IsEntitledToSundayPremium()
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsEntitledToSundayPremium", , , , , True

End Function


Public Function IsUnion148 ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsUnion148 = l_objcboEmployee.IsUnion148(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsUnion148", , , , , True

End Function


Public Function IsExecutive ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsExecutive = l_objcboEmployee.IsExecutive(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsExecutive", , , , , True

End Function


Public Function IsContract ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsContract = l_objcboEmployee.IsContract(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsContract", , , , , True

End Function


Public Function IsLunchFloater ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsLunchFloater = l_objcboEmployee.IsLunchFloater(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsLunchFloater", , , , , True

End Function


Public Function IsDOJM ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsDOJM = l_objcboEmployee.IsDOJM(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsDOJM", , , , , True

End Function


Public Function IsSubstationOrTransmission ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsSubstationOrTransmission = l_objcboEmployee.IsSubstationOrTransmission(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsSubstationOrTransmission", , , , , True

End Function


Public Function IsSubstationOrTransmissionEmployee ( _
	ByVal v_strUnionCode As String,  _
	ByVal v_strDeptNum As String _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsSubstationOrTransmissionEmployee = l_objcboEmployee.IsSubstationOrTransmissionEmployee(v_strUnionCode, v_strDeptNum)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsSubstationOrTransmissionEmployee", , , , , True

End Function


Public Function IsExecutiveShiftWorker ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsExecutiveShiftWorker = l_objcboEmployee.IsExecutiveShiftWorker(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsExecutiveShiftWorker", , , , , True

End Function


Public Function IsQuarterlyExecutive ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsQuarterlyExecutive = l_objcboEmployee.IsQuarterlyExecutive(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsQuarterlyExecutive", , , , , True

End Function


Public Function IsAmerenCIPS ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsAmerenCIPS = l_objcboEmployee.IsAmerenCIPS(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsAmerenCIPS", , , , , True

End Function


Public Function IsAmerenCIPSorCILCOTest ( _
	ByVal v_lngRuleNum As Long,  _
	ByVal v_lngEmpId As Long,  _
	ByVal v_datChgDate As Date _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsAmerenCIPSorCILCOTest = l_objcboEmployee.IsAmerenCIPSorCILCOTest(v_lngRuleNum, v_lngEmpId, v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsAmerenCIPSorCILCOTest", , , , , True

End Function


Public Function IsAmerenUE ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsAmerenUE = l_objcboEmployee.IsAmerenUE(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsAmerenUE", , , , , True

End Function


Public Function IsCoffeen ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsCoffeen = l_objcboEmployee.IsCoffeen(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsCoffeen", , , , , True

End Function


Public Function IsCoffeenMaintenance ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsCoffeenMaintenance = l_objcboEmployee.IsCoffeenMaintenance(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsCoffeenMaintenance", , , , , True

End Function


Public Function IsAmerenLeadership ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsAmerenLeadership = l_objcboEmployee.IsAmerenLeadership(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsAmerenLeadership", , , , , True

End Function


Public Function IsCoffeenOperations ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsCoffeenOperations = l_objcboEmployee.IsCoffeenOperations(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsCoffeenOperations", , , , , True

End Function


Public Function IsAmerenCIPSBenefitParticipant ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsAmerenCIPSBenefitParticipant = l_objcboEmployee.IsAmerenCIPSBenefitParticipant(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsAmerenCIPSBenefitParticipant", , , , , True

End Function


Public Function GetTimeIncrementByChargeDate ( _
	ByVal v_aryEmployees As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	ByVal v_rsEmployeeData As ADODB.Recordset _
	)  As  ADODB.Recordset
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    Set GetTimeIncrementByChargeDate = l_objcboEmployee.GetTimeIncrementByChargeDate(v_aryEmployees, v_datStartDate, v_datEndDate, v_rsEmployeeData)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "GetTimeIncrementByChargeDate", , , , , True

End Function


Public Function GetOTBankChargeInfoByChargeDate ( _
	ByVal v_aryEmployees As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByVal v_rsEmployeeData As ADODB.Recordset _
	)  As  ADODB.Recordset
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    Set GetOTBankChargeInfoByChargeDate = l_objcboEmployee.GetOTBankChargeInfoByChargeDate(v_aryEmployees, v_datStartDate, v_datEndDate, v_rsEmployeeData)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "GetOTBankChargeInfoByChargeDate", , , , , True

End Function


Public Function GetValidateDistIdInfoByChargeDate ( _
	ByVal v_aryEmployees As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	ByVal v_rsEmployeeData As ADODB.Recordset _
	)  As  ADODB.Recordset
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    Set GetValidateDistIdInfoByChargeDate = l_objcboEmployee.GetValidateDistIdInfoByChargeDate(v_aryEmployees, v_datStartDate, v_datEndDate, v_rsEmployeeData)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "GetValidateDistIdInfoByChargeDate", , , , , True

End Function


Public Function IsAmerenCILCO ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsAmerenCILCO = l_objcboEmployee.IsAmerenCILCO(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsAmerenCILCO", , , , , True

End Function


Public Function IsEdwardsOperations ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsEdwardsOperations = l_objcboEmployee.IsEdwardsOperations(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsEdwardsOperations", , , , , True

End Function


Public Function IsAmerenCILCODispatch ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsAmerenCILCODispatch = l_objcboEmployee.IsAmerenCILCODispatch(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsAmerenCILCODispatch", , , , , True

End Function


Public Function IsAmerenCILCOLocal8Employee ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  String
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsAmerenCILCOLocal8Employee = l_objcboEmployee.IsAmerenCILCOLocal8Employee(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsAmerenCILCOLocal8Employee", , , , , True

End Function


Public Function IsNonExemptExecutive ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsNonExemptExecutive = l_objcboEmployee.IsNonExemptExecutive(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsNonExemptExecutive", , , , , True

End Function


Public Function IsAmerenUEBenefitParticipant ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsAmerenUEBenefitParticipant = l_objcboEmployee.IsAmerenUEBenefitParticipant(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsAmerenUEBenefitParticipant", , , , , True

End Function


Public Function IsAmerenIPBenefitParticipant ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsAmerenIPBenefitParticipant = l_objcboEmployee.IsAmerenIPBenefitParticipant(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsAmerenIPBenefitParticipant", , , , , True

End Function


Public Function IsEligibleJobForSundayPremium ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsEligibleJobForSundayPremium = l_objcboEmployee.IsEligibleJobForSundayPremium(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsEligibleJobForSundayPremium", , , , , True

End Function


Public Function IsCoffeen12HourAgreement ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsCoffeen12HourAgreement = l_objcboEmployee.IsCoffeen12HourAgreement(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsCoffeen12HourAgreement", , , , , True

End Function


Public Function IsHutsonville12HourAgreement ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsHutsonville12HourAgreement = l_objcboEmployee.IsHutsonville12HourAgreement(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsHutsonville12HourAgreement", , , , , True

End Function


Public Function IsHutsonvilleLocal148 ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsHutsonvilleLocal148 = l_objcboEmployee.IsHutsonvilleLocal148(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsHutsonvilleLocal148", , , , , True

End Function


Public Function IsCILCOLocal8MaterialHandler ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsCILCOLocal8MaterialHandler = l_objcboEmployee.IsCILCOLocal8MaterialHandler(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsCILCOLocal8MaterialHandler", , , , , True

End Function


Public Function IsJobEligibleForSundayPremium ( _
	Optional ByVal v_datChgDate As Date = "1/1/100",  _
	Optional ByVal v_strJob As String = "" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsJobEligibleForSundayPremium = l_objcboEmployee.IsJobEligibleForSundayPremium(v_datChgDate, v_strJob)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsJobEligibleForSundayPremium", , , , , True

End Function


Public Function IsCoffeen12HourAgreementJob ( _
	Optional ByVal v_datChgDate As Date = "1/1/100",  _
	Optional ByVal v_strJob As String = "" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsCoffeen12HourAgreementJob = l_objcboEmployee.IsCoffeen12HourAgreementJob(v_datChgDate, v_strJob)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsCoffeen12HourAgreementJob", , , , , True

End Function


Public Function IsHutsonville12HourAgreementJob ( _
	Optional ByVal v_datChgDate As Date = "1/1/100",  _
	Optional ByVal v_strJob As String = "" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsHutsonville12HourAgreementJob = l_objcboEmployee.IsHutsonville12HourAgreementJob(v_datChgDate, v_strJob)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsHutsonville12HourAgreementJob", , , , , True

End Function


Public Function IsHutsonvilleMaint ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsHutsonvilleMaint = l_objcboEmployee.IsHutsonvilleMaint(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsHutsonvilleMaint", , , , , True

End Function


Public Function IsHutsonvilleMaintJob ( _
	Optional ByVal v_datChgDate As Date = "1/1/100",  _
	Optional ByVal v_strJob As String = "" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsHutsonvilleMaintJob = l_objcboEmployee.IsHutsonvilleMaintJob(v_datChgDate, v_strJob)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsHutsonvilleMaintJob", , , , , True

End Function


Public Function IsCallawaySecurity ( _
	Optional ByVal v_datChgDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcboEmployee As boCommonLib.cboEmployee
    
On Error GoTo eh

    Set l_objcboEmployee = CreateObject("boCommonLib.cboEmployee")
    
    IsCallawaySecurity = l_objcboEmployee.IsCallawaySecurity(v_datChgDate)
    
    Set l_objcboEmployee = Nothing
    
    Exit Function
    
eh:
    Set l_objcboEmployee = Nothing

    TrsRaiseError "boCommonLib.cboEmployee", "IsCallawaySecurity", , , , , True

End Function