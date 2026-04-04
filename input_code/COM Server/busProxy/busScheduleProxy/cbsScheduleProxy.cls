VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsScheduleProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//
Public Enum TRSReturnOption
    trBaseSchedNoDetails = 1
    trAdjustedBaseSched = 2
    trAdjBaseSchedWithExceptions = 3
    trAdjBaseSchedWithExceptionsAndOvertime = 4
End Enum

Public Enum TRSSchedSourceStartLevel
    trAllScheduleLevels = 50
    trDemandLineTrades = 50
    trDemandLineBaseSchedLevel = 40
    trEmployeeTradesLevel = 30
    trNonDemandLevel = 20
    trWorkGroupBaseScheduleLevel = 10
End Enum

Public Enum TRSLineReturnOption
    trLineBaseSchedNoDetails = 1
    trAdjustedLineBaseSched = 2
End Enum

Public Enum TRSLineSchedSourceStartLevel
    trLineAllLevels = 50
    trLineTrades = 50
    trLineBaseSched = 40
End Enum

Public Enum TRSSchedLineAssignmentSource
    trRelief = 20
    trDemandScheduling = 10
End Enum

Public Enum TRSScheduledExceptionsOption
    trAllExceptions = 1
    trOnlyPendingExceptions = 2
End Enum

Public Enum TRSScheduleTypes
    trScheduleChange = 10
    trTradeDoubleShift = 20
    trPartialShiftException = 110
    trFullShiftException = 115
    trOvertimeException = 120
End Enum

Public Enum TRSScheduleSources
    trWorkGroupBaseSchedule = 10
    trNonDemandSchedule = 20
    trDemandSchedule = 30
    trScheduleChanged = 41
    trTradedDouble = 42
End Enum

Public Enum TRSbusScheduleErrorEnum
    trErrInvalidRequestOnStartDate = -2147201502
    trErrInvalidRequestOption = -2147201501
    trErrEmployeeDoesNotExist = -2147191461
    trErrScheduleCodeDoesNotExist = -2147191460
    trErrEmpExceptionDoesNotMatchSchedule = -2147191459
End Enum





Public Function GetAdjustedBaseSchedule ( _
	ByVal v_aryEmployees As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	ByRef r_rsScheduleWithoutExceptions As ADODB.Recordset,  _
	Optional ByVal v_blnReturnScheduleWithoutExceptions As Boolean = "False",  _
	Optional ByVal v_ReturnOption As TRSReturnOption = "2",  _
	Optional ByVal v_ScheduleSourceStartLevel As TRSSchedSourceStartLevel = "50",  _
	Optional ByVal v_SchedLineSource As TRSSchedLineAssignmentSource,  _
	Optional ByVal v_ScheduleOption As TRSScheduledExceptionsOption = "2" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSchedule As busSchedule.cbsSchedule
    
On Error GoTo eh

    Set l_objcbsSchedule = CreateObject("busSchedule.cbsSchedule")
    
    Set GetAdjustedBaseSchedule = l_objcbsSchedule.GetAdjustedBaseSchedule(v_aryEmployees, v_datStartDate, v_datEndDate, r_rsScheduleWithoutExceptions, v_blnReturnScheduleWithoutExceptions, v_ReturnOption, v_ScheduleSourceStartLevel, v_SchedLineSource, v_ScheduleOption)
    
    Set l_objcbsSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedule = Nothing

    TrsRaiseError "busSchedule.cbsSchedule", "GetAdjustedBaseSchedule", , , , , True

End Function


Public Function GetEffectiveWorkGroupBaseSchedule ( _
	ByVal v_lngWorkGroup As Long,  _
	ByVal v_datDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSchedule As busSchedule.cbsSchedule
    
On Error GoTo eh

    Set l_objcbsSchedule = CreateObject("busSchedule.cbsSchedule")
    
    Set GetEffectiveWorkGroupBaseSchedule = l_objcbsSchedule.GetEffectiveWorkGroupBaseSchedule(v_lngWorkGroup, v_datDate)
    
    Set l_objcbsSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedule = Nothing

    TrsRaiseError "busSchedule.cbsSchedule", "GetEffectiveWorkGroupBaseSchedule", , , , , True

End Function


Public Function GetSchedulesInRotation ( _
	ByVal v_intRotationCode As Integer,  _
	ByVal v_datDate As Date,  _
	ByVal v_datEffectiveDate As Date,  _
	ByVal v_intStartWeek As Integer _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSchedule As busSchedule.cbsSchedule
    
On Error GoTo eh

    Set l_objcbsSchedule = CreateObject("busSchedule.cbsSchedule")
    
    Set GetSchedulesInRotation = l_objcbsSchedule.GetSchedulesInRotation(v_intRotationCode, v_datDate, v_datEffectiveDate, v_intStartWeek)
    
    Set l_objcbsSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedule = Nothing

    TrsRaiseError "busSchedule.cbsSchedule", "GetSchedulesInRotation", , , , , True

End Function


Public Function GetAdjustedBaseWorkGroupSchedules ( _
	ByVal v_aryWorkgroups As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByVal v_ReturnOption As TRSReturnOption = "2",  _
	Optional ByVal v_ScheduleSourceStartLevel As TRSSchedSourceStartLevel = "50" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSchedule As busSchedule.cbsSchedule
    
On Error GoTo eh

    Set l_objcbsSchedule = CreateObject("busSchedule.cbsSchedule")
    
    Set GetAdjustedBaseWorkGroupSchedules = l_objcbsSchedule.GetAdjustedBaseWorkGroupSchedules(v_aryWorkgroups, v_datStartDate, v_datEndDate, v_ReturnOption, v_ScheduleSourceStartLevel)
    
    Set l_objcbsSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedule = Nothing

    TrsRaiseError "busSchedule.cbsSchedule", "GetAdjustedBaseWorkGroupSchedules", , , , , True

End Function


Public Function GABEmployeeSchedOnSpecificDates ( _
	ByVal v_rsEmployeesAndDates As ADODB.Recordset,  _
	ByRef r_rsScheduleWithoutExceptions As ADODB.Recordset,  _
	Optional ByVal v_blnReturnScheduleWithoutExceptions As Boolean = "False",  _
	Optional ByVal v_ReturnOption As TRSReturnOption = "2",  _
	Optional ByVal v_ScheduleSourceStartLevel As TRSSchedSourceStartLevel = "50" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSchedule As busSchedule.cbsSchedule
    
On Error GoTo eh

    Set l_objcbsSchedule = CreateObject("busSchedule.cbsSchedule")
    
    Set GABEmployeeSchedOnSpecificDates = l_objcbsSchedule.GABEmployeeSchedOnSpecificDates(v_rsEmployeesAndDates, r_rsScheduleWithoutExceptions, v_blnReturnScheduleWithoutExceptions, v_ReturnOption, v_ScheduleSourceStartLevel)
    
    Set l_objcbsSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedule = Nothing

    TrsRaiseError "busSchedule.cbsSchedule", "GABEmployeeSchedOnSpecificDates", , , , , True

End Function


Public Function GetPendingAdjustedBaseSchedule ( _
	ByVal v_rsEmployeesAndDates As ADODB.Recordset,  _
	ByRef r_rsScheduleWithoutExceptions As ADODB.Recordset,  _
	Optional ByVal v_blnReturnScheduleWithoutExceptions As Boolean = "False",  _
	Optional ByVal v_ReturnOption As TRSReturnOption = "2",  _
	Optional ByVal v_ScheduleSourceStartLevel As TRSSchedSourceStartLevel = "50",  _
	Optional ByVal v_rsWGAssignmentSPA As ADODB.Recordset,  _
	Optional ByVal v_rsWGBaseSchedChanges As ADODB.Recordset,  _
	Optional ByVal v_rsWGLineSchedChanges As ADODB.Recordset,  _
	Optional ByVal v_rsWGAssignChanges As ADODB.Recordset,  _
	Optional ByVal v_rsWGAssignOtherChanges As ADODB.Recordset,  _
	Optional ByVal v_rsEmpLineAssignChanges As ADODB.Recordset,  _
	Optional ByVal v_rsPersonalChanges As ADODB.Recordset,  _
	Optional ByVal v_rsLineTradeChanges As ADODB.Recordset,  _
	Optional ByVal v_rsEmpTradeChanges As ADODB.Recordset _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSchedule As busSchedule.cbsSchedule
    
On Error GoTo eh

    Set l_objcbsSchedule = CreateObject("busSchedule.cbsSchedule")
    
    Set GetPendingAdjustedBaseSchedule = l_objcbsSchedule.GetPendingAdjustedBaseSchedule(v_rsEmployeesAndDates, r_rsScheduleWithoutExceptions, v_blnReturnScheduleWithoutExceptions, v_ReturnOption, v_ScheduleSourceStartLevel, v_rsWGAssignmentSPA, v_rsWGBaseSchedChanges, v_rsWGLineSchedChanges, v_rsWGAssignChanges, v_rsWGAssignOtherChanges, v_rsEmpLineAssignChanges, v_rsPersonalChanges, v_rsLineTradeChanges, v_rsEmpTradeChanges)
    
    Set l_objcbsSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedule = Nothing

    TrsRaiseError "busSchedule.cbsSchedule", "GetPendingAdjustedBaseSchedule", , , , , True

End Function


Public Function GetRotationScheduleCode ( _
	ByVal v_datChargeDate As Date,  _
	ByVal v_intRotationCode As Integer,  _
	ByVal v_datEffDate As Date,  _
	ByVal v_intStartPosition As Integer,  _
	ByRef r_intRecordPosition As Integer _
	)  As  Integer
  
    Dim l_objcbsSchedule As busSchedule.cbsSchedule
    
On Error GoTo eh

    Set l_objcbsSchedule = CreateObject("busSchedule.cbsSchedule")
    
    GetRotationScheduleCode = l_objcbsSchedule.GetRotationScheduleCode(v_datChargeDate, v_intRotationCode, v_datEffDate, v_intStartPosition, r_intRecordPosition)
    
    Set l_objcbsSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedule = Nothing

    TrsRaiseError "busSchedule.cbsSchedule", "GetRotationScheduleCode", , , , , True

End Function


Public Function GetAdjustedBaseDepartmentSchedules ( _
	ByVal v_aryDepartments As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByVal v_ReturnOption As TRSReturnOption = "2",  _
	Optional ByVal v_ScheduleSourceStartLevel As TRSSchedSourceStartLevel = "50" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSchedule As busSchedule.cbsSchedule
    
On Error GoTo eh

    Set l_objcbsSchedule = CreateObject("busSchedule.cbsSchedule")
    
    Set GetAdjustedBaseDepartmentSchedules = l_objcbsSchedule.GetAdjustedBaseDepartmentSchedules(v_aryDepartments, v_datStartDate, v_datEndDate, v_ReturnOption, v_ScheduleSourceStartLevel)
    
    Set l_objcbsSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedule = Nothing

    TrsRaiseError "busSchedule.cbsSchedule", "GetAdjustedBaseDepartmentSchedules", , , , , True

End Function


Public Function KeepsShiftComponentTest ( _
	ByVal v_lngEmpNum As Long,  _
	ByVal v_datChargeDate As Date _
	)  As  Boolean
  
    Dim l_objcbsSchedule As busSchedule.cbsSchedule
    
On Error GoTo eh

    Set l_objcbsSchedule = CreateObject("busSchedule.cbsSchedule")
    
    KeepsShiftComponentTest = l_objcbsSchedule.KeepsShiftComponentTest(v_lngEmpNum, v_datChargeDate)
    
    Set l_objcbsSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedule = Nothing

    TrsRaiseError "busSchedule.cbsSchedule", "KeepsShiftComponentTest", , , , , True

End Function


Public Function GetScheduleForOTCanvassing ( _
	ByVal v_aryEmployees As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSchedule As busSchedule.cbsSchedule
    
On Error GoTo eh

    Set l_objcbsSchedule = CreateObject("busSchedule.cbsSchedule")
    
    Set GetScheduleForOTCanvassing = l_objcbsSchedule.GetScheduleForOTCanvassing(v_aryEmployees, v_datStartDate, v_datEndDate)
    
    Set l_objcbsSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedule = Nothing

    TrsRaiseError "busSchedule.cbsSchedule", "GetScheduleForOTCanvassing", , , , , True

End Function