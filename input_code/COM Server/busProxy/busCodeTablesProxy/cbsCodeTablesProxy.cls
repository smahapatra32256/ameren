VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsCodeTablesProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Sub GetAllSupportInformation ( _
	ByVal v_aryRequest As Variant,  _
	ByRef r_rsAllExceptions As ADODB.Recordset,  _
	ByRef r_rsAllComponents As ADODB.Recordset,  _
	ByRef r_rsAllHolidays As ADODB.Recordset,  _
	ByRef r_rsAllJobs As ADODB.Recordset,  _
	ByRef r_rsAllOvertimeReasons As ADODB.Recordset,  _
	ByRef r_rsAllLocations As ADODB.Recordset,  _
	ByRef r_rsAllUnions As ADODB.Recordset,  _
	ByRef r_rsAllDepartments As ADODB.Recordset,  _
	ByRef r_rsAllTravelTripLocations As ADODB.Recordset,  _
	ByRef r_rsAllTravelTripAllowances As ADODB.Recordset,  _
	ByRef r_rsAllActivityList As ADODB.Recordset,  _
	ByRef r_rsAllProjectInfo As ADODB.Recordset _
	)  
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    l_objcbsCodeTables.GetAllSupportInformation v_aryRequest, r_rsAllExceptions, r_rsAllComponents, r_rsAllHolidays, r_rsAllJobs, r_rsAllOvertimeReasons, r_rsAllLocations, r_rsAllUnions, r_rsAllDepartments, r_rsAllTravelTripLocations, r_rsAllTravelTripAllowances, r_rsAllActivityList, r_rsAllProjectInfo
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllSupportInformation", , , , , True

End Sub


Public Function GetAllJobs ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllJobs = l_objcbsCodeTables.GetAllJobs()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllJobs", , , , , True

End Function


Public Function GetAllLocations ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllLocations = l_objcbsCodeTables.GetAllLocations()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllLocations", , , , , True

End Function


Public Function GetAllUnions ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllUnions = l_objcbsCodeTables.GetAllUnions()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllUnions", , , , , True

End Function


Public Function GetAllOvertimeReasons ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllOvertimeReasons = l_objcbsCodeTables.GetAllOvertimeReasons()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllOvertimeReasons", , , , , True

End Function


Public Function GetEffectiveUnionComponents ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetEffectiveUnionComponents = l_objcbsCodeTables.GetEffectiveUnionComponents()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetEffectiveUnionComponents", , , , , True

End Function


Public Function GetUnionComponents ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetUnionComponents = l_objcbsCodeTables.GetUnionComponents()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetUnionComponents", , , , , True

End Function


Public Function GetAllExceptions ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllExceptions = l_objcbsCodeTables.GetAllExceptions()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllExceptions", , , , , True

End Function


Public Function GetAllHolidays ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllHolidays = l_objcbsCodeTables.GetAllHolidays()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllHolidays", , , , , True

End Function


Public Function GetAllEarningCodes ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllEarningCodes = l_objcbsCodeTables.GetAllEarningCodes()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllEarningCodes", , , , , True

End Function


Public Function GetTravelAllowances ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetTravelAllowances = l_objcbsCodeTables.GetTravelAllowances()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetTravelAllowances", , , , , True

End Function


Public Function GetTravelTripLocations ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetTravelTripLocations = l_objcbsCodeTables.GetTravelTripLocations()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetTravelTripLocations", , , , , True

End Function


Public Function GetUniqueComponentCodes ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetUniqueComponentCodes = l_objcbsCodeTables.GetUniqueComponentCodes()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetUniqueComponentCodes", , , , , True

End Function


Public Function GetAllDepartments ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllDepartments = l_objcbsCodeTables.GetAllDepartments()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllDepartments", , , , , True

End Function


Public Function IsVacationReliefCode ( _
	ByVal v_strUnionCd As String,  _
	ByVal v_datEffDate As Date,  _
	ByVal v_strCompCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    IsVacationReliefCode = l_objcbsCodeTables.IsVacationReliefCode(v_strUnionCd, v_datEffDate, v_strCompCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "IsVacationReliefCode", , , , , True

End Function


Public Function IsShiftCompCode ( _
	ByVal v_strUnionCd As String,  _
	ByVal v_datEffDate As Date,  _
	ByVal v_strCompCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    IsShiftCompCode = l_objcbsCodeTables.IsShiftCompCode(v_strUnionCd, v_datEffDate, v_strCompCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "IsShiftCompCode", , , , , True

End Function


Public Function IsComponentCode ( _
	ByVal v_strUnionCd As String,  _
	ByVal v_datEffDate As Date,  _
	ByVal v_strCompCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    IsComponentCode = l_objcbsCodeTables.IsComponentCode(v_strUnionCd, v_datEffDate, v_strCompCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "IsComponentCode", , , , , True

End Function


Public Function ExceptionCodeIsValid ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionCodeIsValid = l_objcbsCodeTables.ExceptionCodeIsValid(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionCodeIsValid", , , , , True

End Function


Public Function ExceptionIsCaseRelated ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsCaseRelated = l_objcbsCodeTables.ExceptionIsCaseRelated(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsCaseRelated", , , , , True

End Function


Public Function ExceptionIsValidInsideOfSchedule ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsValidInsideOfSchedule = l_objcbsCodeTables.ExceptionIsValidInsideOfSchedule(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsValidInsideOfSchedule", , , , , True

End Function


Public Function ExceptionIsValidOutsideOfSchedule ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsValidOutsideOfSchedule = l_objcbsCodeTables.ExceptionIsValidOutsideOfSchedule(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsValidOutsideOfSchedule", , , , , True

End Function


Public Function ExceptionIsOvertimeRelated ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsOvertimeRelated = l_objcbsCodeTables.ExceptionIsOvertimeRelated(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsOvertimeRelated", , , , , True

End Function


Public Function ExceptionIsOvertimeBankRelated ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsOvertimeBankRelated = l_objcbsCodeTables.ExceptionIsOvertimeBankRelated(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsOvertimeBankRelated", , , , , True

End Function


Public Function ExceptionIsUnpaid ( _
	ByVal v_strExcCd As String,  _
	ByVal v_strExecConInd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsUnpaid = l_objcbsCodeTables.ExceptionIsUnpaid(v_strExcCd, v_strExecConInd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsUnpaid", , , , , True

End Function


Public Function ExceptionIsOtherPay ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsOtherPay = l_objcbsCodeTables.ExceptionIsOtherPay(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsOtherPay", , , , , True

End Function


Public Function ExceptionIsECNonProductive ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsECNonProductive = l_objcbsCodeTables.ExceptionIsECNonProductive(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsECNonProductive", , , , , True

End Function


Public Function ExceptionIsOverride ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsOverride = l_objcbsCodeTables.ExceptionIsOverride(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsOverride", , , , , True

End Function


Public Function ExceptionIsEligibleForScheduleChangePayment ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsEligibleForScheduleChangePayment = l_objcbsCodeTables.ExceptionIsEligibleForScheduleChangePayment(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsEligibleForScheduleChangePayment", , , , , True

End Function


Public Function ExceptionIsEligibleForFamilyLeave ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsEligibleForFamilyLeave = l_objcbsCodeTables.ExceptionIsEligibleForFamilyLeave(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsEligibleForFamilyLeave", , , , , True

End Function


Public Function ExceptionIsValidScheduleChangeCode ( _
	ByVal v_strExcCd As String,  _
	ByVal v_strExecConInd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionIsValidScheduleChangeCode = l_objcbsCodeTables.ExceptionIsValidScheduleChangeCode(v_strExcCd, v_strExecConInd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionIsValidScheduleChangeCode", , , , , True

End Function


Public Function ExceptionScheduleChangeExists ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionScheduleChangeExists = l_objcbsCodeTables.ExceptionScheduleChangeExists(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionScheduleChangeExists", , , , , True

End Function


Public Function ExceptionCodeIsWorking ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionCodeIsWorking = l_objcbsCodeTables.ExceptionCodeIsWorking(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionCodeIsWorking", , , , , True

End Function


Public Function IsVacationSymbolException ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    IsVacationSymbolException = l_objcbsCodeTables.IsVacationSymbolException(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "IsVacationSymbolException", , , , , True

End Function


Public Function IsHolidaySymbolException ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    IsHolidaySymbolException = l_objcbsCodeTables.IsHolidaySymbolException(v_strExcCd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "IsHolidaySymbolException", , , , , True

End Function


Public Function GetAvailableScheduleDetails ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAvailableScheduleDetails = l_objcbsCodeTables.GetAvailableScheduleDetails()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAvailableScheduleDetails", , , , , True

End Function


Public Function GetAvailableScheduleHeader ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAvailableScheduleHeader = l_objcbsCodeTables.GetAvailableScheduleHeader()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAvailableScheduleHeader", , , , , True

End Function


Public Function GetAllActivityList ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllActivityList = l_objcbsCodeTables.GetAllActivityList()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllActivityList", , , , , True

End Function


Public Function ValidateScheduleCode ( _
	ByVal v_lScheduleCode As Long _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ValidateScheduleCode = l_objcbsCodeTables.ValidateScheduleCode(v_lScheduleCode)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ValidateScheduleCode", , , , , True

End Function


Public Function JobNumberIsValid ( _
	ByVal v_strJobNum As String,  _
	Optional ByVal v_datEffDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    JobNumberIsValid = l_objcbsCodeTables.JobNumberIsValid(v_strJobNum, v_datEffDate)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "JobNumberIsValid", , , , , True

End Function


Public Function GetProjectInfo ( _
	Optional ByVal v_intProjectId As Integer,  _
	Optional ByVal v_intPPMInd As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetProjectInfo = l_objcbsCodeTables.GetProjectInfo(v_intProjectId, v_intPPMInd)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetProjectInfo", , , , , True

End Function


Public Function GetAllProjects ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllProjects = l_objcbsCodeTables.GetAllProjects()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllProjects", , , , , True

End Function


Public Function GetLocationsAssociatedToDepartment ( _
	ByVal v_varDepts As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetLocationsAssociatedToDepartment = l_objcbsCodeTables.GetLocationsAssociatedToDepartment(v_varDepts)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetLocationsAssociatedToDepartment", , , , , True

End Function


Public Function GetAllScheduleCodes ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllScheduleCodes = l_objcbsCodeTables.GetAllScheduleCodes()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllScheduleCodes", , , , , True

End Function


Public Function GetRotationSchedHeader ( _
	Optional ByVal v_intRotationCode As Integer _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetRotationSchedHeader = l_objcbsCodeTables.GetRotationSchedHeader(v_intRotationCode)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetRotationSchedHeader", , , , , True

End Function


Public Function GetRotationSchedDetail ( _
	Optional ByVal v_intRotationCode As Integer,  _
	Optional ByVal v_intRotationWeek As Integer _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetRotationSchedDetail = l_objcbsCodeTables.GetRotationSchedDetail(v_intRotationCode, v_intRotationWeek)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetRotationSchedDetail", , , , , True

End Function


Public Function GetExceptionGroupForColoring ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetExceptionGroupForColoring = l_objcbsCodeTables.GetExceptionGroupForColoring()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetExceptionGroupForColoring", , , , , True

End Function


Public Function ExceptionCodeIsValidFor ( _
	ByVal v_strExcCd As String,  _
	ByVal v_strExecConInd As String,  _
	Optional ByVal v_strBenefitsIND As String = "" _
	)  As  Boolean
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    ExceptionCodeIsValidFor = l_objcbsCodeTables.ExceptionCodeIsValidFor(v_strExcCd, v_strExecConInd, v_strBenefitsIND)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "ExceptionCodeIsValidFor", , , , , True

End Function


Public Function GetAllCrafts ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllCrafts = l_objcbsCodeTables.GetAllCrafts()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllCrafts", , , , , True

End Function


Public Function GetAllQualifications ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllQualifications = l_objcbsCodeTables.GetAllQualifications()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllQualifications", , , , , True

End Function


Public Function GetAllSpecialSkills ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllSpecialSkills = l_objcbsCodeTables.GetAllSpecialSkills()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllSpecialSkills", , , , , True

End Function


Public Function GetAllOvertimeResponses ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllOvertimeResponses = l_objcbsCodeTables.GetAllOvertimeResponses()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllOvertimeResponses", , , , , True

End Function


Public Function GetAllPromotionalLines ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllPromotionalLines = l_objcbsCodeTables.GetAllPromotionalLines()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllPromotionalLines", , , , , True

End Function


Public Function GetAllCompanies ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllCompanies = l_objcbsCodeTables.GetAllCompanies()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllCompanies", , , , , True

End Function


Public Function GetAllOvertimeBankHeaders ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllOvertimeBankHeaders = l_objcbsCodeTables.GetAllOvertimeBankHeaders()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllOvertimeBankHeaders", , , , , True

End Function


Public Function GetAllBudgetNumbers ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetAllBudgetNumbers = l_objcbsCodeTables.GetAllBudgetNumbers()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetAllBudgetNumbers", , , , , True

End Function


Public Function GetSeniorityReportSortOrder ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetSeniorityReportSortOrder = l_objcbsCodeTables.GetSeniorityReportSortOrder()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetSeniorityReportSortOrder", , , , , True

End Function


Public Function GetPreOppReportSortOrder ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetPreOppReportSortOrder = l_objcbsCodeTables.GetPreOppReportSortOrder()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetPreOppReportSortOrder", , , , , True

End Function


Public Function GetOTEqualReportSortOrder ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetOTEqualReportSortOrder = l_objcbsCodeTables.GetOTEqualReportSortOrder()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetOTEqualReportSortOrder", , , , , True

End Function


Public Function GetVacSumReportSortOrder ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetVacSumReportSortOrder = l_objcbsCodeTables.GetVacSumReportSortOrder()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetVacSumReportSortOrder", , , , , True

End Function


Public Function GetHolidays ( _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date,  _
	Optional ByVal vsUnionCode As String = "" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetHolidays = l_objcbsCodeTables.GetHolidays(vdtStartDate, vdtEndDate, vsUnionCode)
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetHolidays", , , , , True

End Function


Public Function GetJobTaskExceptions ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCodeTables As busCodeTables.cbsCodeTables
    
On Error GoTo eh

    Set l_objcbsCodeTables = CreateObject("busCodeTables.cbsCodeTables")
    
    Set GetJobTaskExceptions = l_objcbsCodeTables.GetJobTaskExceptions()
    
    Set l_objcbsCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTables = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTables", "GetJobTaskExceptions", , , , , True

End Function