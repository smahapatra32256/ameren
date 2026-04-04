VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdCodeTablesProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//
Public Enum TrsComponentType
    trReliefComponent = 1
    trShiftComponent = 2
    trOtherComponent = 3
End Enum





Public Sub GetAllSupportInformation ( _
	ByRef r_rsAllExceptions As ADODB.Recordset,  _
	ByRef r_rsAllComponents As ADODB.Recordset,  _
	ByRef r_rsAllHolidays As ADODB.Recordset,  _
	ByRef r_rsAllJobs As ADODB.Recordset,  _
	ByRef r_rsAllOvertimeReasons As ADODB.Recordset,  _
	ByRef r_rsAllLocations As ADODB.Recordset,  _
	ByRef r_rsAllUnions As ADODB.Recordset,  _
	ByRef r_rsAllDepartments As ADODB.Recordset,  _
	ByRef r_rsAllActivityList As ADODB.Recordset _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.GetAllSupportInformation r_rsAllExceptions, r_rsAllComponents, r_rsAllHolidays, r_rsAllJobs, r_rsAllOvertimeReasons, r_rsAllLocations, r_rsAllUnions, r_rsAllDepartments, r_rsAllActivityList
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllSupportInformation", , , , , True

End Sub


Public Function GetAllDepartments ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllDepartments = l_objcdCodeTables.GetAllDepartments()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllDepartments", , , , , True

End Function


Public Function GetAllJobs ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllJobs = l_objcdCodeTables.GetAllJobs()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllJobs", , , , , True

End Function


Public Function GetJob ( _
	ByVal v_strJobNum As String _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetJob = l_objcdCodeTables.GetJob(v_strJobNum)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetJob", , , , , True

End Function


Public Function GetUnions ( _
	ByRef r_colCriteria As Collection _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetUnions = l_objcdCodeTables.GetUnions(r_colCriteria)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetUnions", , , , , True

End Function


Public Function GetAllUnions ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllUnions = l_objcdCodeTables.GetAllUnions()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllUnions", , , , , True

End Function


Public Function GetAllEarningCodes ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllEarningCodes = l_objcdCodeTables.GetAllEarningCodes()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllEarningCodes", , , , , True

End Function


Public Function GetTravelAllowances ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetTravelAllowances = l_objcdCodeTables.GetTravelAllowances()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetTravelAllowances", , , , , True

End Function


Public Function GetTravelTripLocations ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetTravelTripLocations = l_objcdCodeTables.GetTravelTripLocations()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetTravelTripLocations", , , , , True

End Function


Public Function GetOvertimeExceptions ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetOvertimeExceptions = l_objcdCodeTables.GetOvertimeExceptions()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetOvertimeExceptions", , , , , True

End Function


Public Function GetHolidaysBetween ( _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByVal v_colUnionArgs As Collection _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetHolidaysBetween = l_objcdCodeTables.GetHolidaysBetween(v_datStartDate, v_datEndDate, v_colUnionArgs)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetHolidaysBetween", , , , , True

End Function


Public Function GetAllHolidays ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllHolidays = l_objcdCodeTables.GetAllHolidays()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllHolidays", , , , , True

End Function


Public Function GetExceptionRecords ( _
	Optional ByVal v_varExcCodes As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetExceptionRecords = l_objcdCodeTables.GetExceptionRecords(v_varExcCodes)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetExceptionRecords", , , , , True

End Function


Public Function GetUnionComponentRecords ( _
	Optional ByVal v_varKeys As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetUnionComponentRecords = l_objcdCodeTables.GetUnionComponentRecords(v_varKeys)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetUnionComponentRecords", , , , , True

End Function


Public Function GetAllExceptions ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllExceptions = l_objcdCodeTables.GetAllExceptions()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllExceptions", , , , , True

End Function


Public Function GetAllLocations ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllLocations = l_objcdCodeTables.GetAllLocations()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllLocations", , , , , True

End Function


Public Function GetAllOvertimeReasons ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllOvertimeReasons = l_objcdCodeTables.GetAllOvertimeReasons()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllOvertimeReasons", , , , , True

End Function


Public Function GetAllOvertimeResponses ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllOvertimeResponses = l_objcdCodeTables.GetAllOvertimeResponses()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllOvertimeResponses", , , , , True

End Function


Public Function LocationRecord ( _
	ByVal v_strLocCd As String _
	)  As  Variant
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    LocationRecord = l_objcdCodeTables.LocationRecord(v_strLocCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "LocationRecord", , , , , True

End Function


Public Function LocationCodeIsValid ( _
	ByVal v_strLocCd As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    LocationCodeIsValid = l_objcdCodeTables.LocationCodeIsValid(v_strLocCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "LocationCodeIsValid", , , , , True

End Function


Public Function LocationDescription ( _
	ByVal v_strLocCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    LocationDescription = l_objcdCodeTables.LocationDescription(v_strLocCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "LocationDescription", , , , , True

End Function


Public Function LocationPlantCode ( _
	ByVal v_strLocCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    LocationPlantCode = l_objcdCodeTables.LocationPlantCode(v_strLocCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "LocationPlantCode", , , , , True

End Function


Public Function JobRecord ( _
	ByVal v_strJobNum As String _
	)  As  Variant
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    JobRecord = l_objcdCodeTables.JobRecord(v_strJobNum)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "JobRecord", , , , , True

End Function


Public Function JobNumberIsValid ( _
	ByVal v_strJobNum As String,  _
	Optional ByVal v_datEffDate As Date = "1/1/100" _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    JobNumberIsValid = l_objcdCodeTables.JobNumberIsValid(v_strJobNum, v_datEffDate)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "JobNumberIsValid", , , , , True

End Function


Public Function JobRecordExists ( _
	ByVal v_strJobNum As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    JobRecordExists = l_objcdCodeTables.JobRecordExists(v_strJobNum)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "JobRecordExists", , , , , True

End Function


Public Function JobUnionCode ( _
	ByVal v_strJobNum As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    JobUnionCode = l_objcdCodeTables.JobUnionCode(v_strJobNum)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "JobUnionCode", , , , , True

End Function


Public Function JobOvertimeType ( _
	ByVal v_strJobNum As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    JobOvertimeType = l_objcdCodeTables.JobOvertimeType(v_strJobNum)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "JobOvertimeType", , , , , True

End Function


Public Function JobDescription ( _
	ByVal v_strJobNum As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    JobDescription = l_objcdCodeTables.JobDescription(v_strJobNum)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "JobDescription", , , , , True

End Function


Public Function JobCraftNumber ( _
	ByVal v_strJobNum As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    JobCraftNumber = l_objcdCodeTables.JobCraftNumber(v_strJobNum)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "JobCraftNumber", , , , , True

End Function


Public Function JobEffectiveDate ( _
	ByVal v_strJobNum As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    JobEffectiveDate = l_objcdCodeTables.JobEffectiveDate(v_strJobNum)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "JobEffectiveDate", , , , , True

End Function


Public Function JobEffectiveStatus ( _
	ByVal v_strJobNum As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    JobEffectiveStatus = l_objcdCodeTables.JobEffectiveStatus(v_strJobNum)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "JobEffectiveStatus", , , , , True

End Function


Public Function ExceptionRecord ( _
	ByVal v_strExcCd As String _
	)  As  Variant
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionRecord = l_objcdCodeTables.ExceptionRecord(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionRecord", , , , , True

End Function


Public Function ExceptionType ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionType = l_objcdCodeTables.ExceptionType(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionType", , , , , True

End Function


Public Function ExceptionCaseGroup ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionCaseGroup = l_objcdCodeTables.ExceptionCaseGroup(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionCaseGroup", , , , , True

End Function


Public Function ExceptionDescription ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionDescription = l_objcdCodeTables.ExceptionDescription(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionDescription", , , , , True

End Function


Public Function ExceptionAccountNumber ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionAccountNumber = l_objcdCodeTables.ExceptionAccountNumber(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionAccountNumber", , , , , True

End Function


Public Function ExceptionBankCode ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionBankCode = l_objcdCodeTables.ExceptionBankCode(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionBankCode", , , , , True

End Function


Public Function ExceptionBankType ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionBankType = l_objcdCodeTables.ExceptionBankType(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionBankType", , , , , True

End Function


Public Function ExceptionEarningsCode ( _
	ByVal v_strExcCd As String,  _
	ByVal v_strExecConInd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionEarningsCode = l_objcdCodeTables.ExceptionEarningsCode(v_strExcCd, v_strExecConInd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionEarningsCode", , , , , True

End Function


Public Function ExceptionExecutiveEarningsCode ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionExecutiveEarningsCode = l_objcdCodeTables.ExceptionExecutiveEarningsCode(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionExecutiveEarningsCode", , , , , True

End Function


Public Function ExceptionContractEarningsCode ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionContractEarningsCode = l_objcdCodeTables.ExceptionContractEarningsCode(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionContractEarningsCode", , , , , True

End Function


Public Function ExceptionOtherPayIndicator ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionOtherPayIndicator = l_objcdCodeTables.ExceptionOtherPayIndicator(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionOtherPayIndicator", , , , , True

End Function


Public Function ExceptionScheduleChgInd ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionScheduleChgInd = l_objcdCodeTables.ExceptionScheduleChgInd(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionScheduleChgInd", , , , , True

End Function


Public Function ExceptionFamilyLeaveInd ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionFamilyLeaveInd = l_objcdCodeTables.ExceptionFamilyLeaveInd(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionFamilyLeaveInd", , , , , True

End Function


Public Function ExceptionCodeIsValid ( _
	ByVal v_strExcCd As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionCodeIsValid = l_objcdCodeTables.ExceptionCodeIsValid(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionCodeIsValid", , , , , True

End Function


Public Function ExceptionCodeIsValidFor ( _
	ByVal v_strExcCd As String,  _
	ByVal v_strExecConInd As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionCodeIsValidFor = l_objcdCodeTables.ExceptionCodeIsValidFor(v_strExcCd, v_strExecConInd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionCodeIsValidFor", , , , , True

End Function


Public Function ExceptionWorkRestInd ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionWorkRestInd = l_objcdCodeTables.ExceptionWorkRestInd(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionWorkRestInd", , , , , True

End Function


Public Function UnionRecord ( _
	ByVal v_strUnionCd As String _
	)  As  Variant
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    UnionRecord = l_objcdCodeTables.UnionRecord(v_strUnionCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "UnionRecord", , , , , True

End Function


Public Function UnionDescription ( _
	ByVal v_strUnionCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    UnionDescription = l_objcdCodeTables.UnionDescription(v_strUnionCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "UnionDescription", , , , , True

End Function


Public Function UnionExecutiveContractIndicator ( _
	ByVal v_strUnionCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    UnionExecutiveContractIndicator = l_objcdCodeTables.UnionExecutiveContractIndicator(v_strUnionCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "UnionExecutiveContractIndicator", , , , , True

End Function


Public Function UnionBenefitPackageIndicator ( _
	ByVal v_strUnionCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    UnionBenefitPackageIndicator = l_objcdCodeTables.UnionBenefitPackageIndicator(v_strUnionCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "UnionBenefitPackageIndicator", , , , , True

End Function


Public Function UnionLocal ( _
	ByVal v_strUnionCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    UnionLocal = l_objcdCodeTables.UnionLocal(v_strUnionCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "UnionLocal", , , , , True

End Function


Public Function UnionCodeIsValid ( _
	ByVal v_strUnionCd As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    UnionCodeIsValid = l_objcdCodeTables.UnionCodeIsValid(v_strUnionCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "UnionCodeIsValid", , , , , True

End Function


Public Function UnionComponentVacationReliefAmount ( _
	ByVal v_strUnionCd As String,  _
	ByVal v_datEffDate As Date,  _
	ByVal v_strCompCd As String _
	)  As  Variant
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    UnionComponentVacationReliefAmount = l_objcdCodeTables.UnionComponentVacationReliefAmount(v_strUnionCd, v_datEffDate, v_strCompCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "UnionComponentVacationReliefAmount", , , , , True

End Function


Public Function UnionComponentShiftAmount ( _
	ByVal v_strUnionCd As String,  _
	ByVal v_datEffDate As Date,  _
	ByVal v_strCompCd As String _
	)  As  Variant
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    UnionComponentShiftAmount = l_objcdCodeTables.UnionComponentShiftAmount(v_strUnionCd, v_datEffDate, v_strCompCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "UnionComponentShiftAmount", , , , , True

End Function


Public Function UnionComponentAmount ( _
	ByVal v_strUnionCd As String,  _
	ByVal v_datEffDate As Date,  _
	ByVal v_strCompCd As String _
	)  As  Variant
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    UnionComponentAmount = l_objcdCodeTables.UnionComponentAmount(v_strUnionCd, v_datEffDate, v_strCompCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "UnionComponentAmount", , , , , True

End Function


Public Function OvertimeReasonIsValid ( _
	ByVal v_strOTReasonCd As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    OvertimeReasonIsValid = l_objcdCodeTables.OvertimeReasonIsValid(v_strOTReasonCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "OvertimeReasonIsValid", , , , , True

End Function


Public Function OvertimeReasonDescription ( _
	ByVal v_strOTReasonCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    OvertimeReasonDescription = l_objcdCodeTables.OvertimeReasonDescription(v_strOTReasonCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "OvertimeReasonDescription", , , , , True

End Function


Public Function OvertimeResponseIsValid ( _
	ByVal v_strOTResponse As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    OvertimeResponseIsValid = l_objcdCodeTables.OvertimeResponseIsValid(v_strOTResponse)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "OvertimeResponseIsValid", , , , , True

End Function


Public Function OvertimeResponseDescription ( _
	ByVal v_strOTResponse As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    OvertimeResponseDescription = l_objcdCodeTables.OvertimeResponseDescription(v_strOTResponse)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "OvertimeResponseDescription", , , , , True

End Function


Public Function OvertimeResponseWorkIndicator ( _
	ByVal v_strOTResponse As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    OvertimeResponseWorkIndicator = l_objcdCodeTables.OvertimeResponseWorkIndicator(v_strOTResponse)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "OvertimeResponseWorkIndicator", , , , , True

End Function


Public Function OvertimeResponseChargeIndicator ( _
	ByVal v_strOTResponse As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    OvertimeResponseChargeIndicator = l_objcdCodeTables.OvertimeResponseChargeIndicator(v_strOTResponse)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "OvertimeResponseChargeIndicator", , , , , True

End Function


Public Function OvertimeExceptionFactor ( _
	ByVal v_strExceptionCode As String _
	)  As  Variant
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    OvertimeExceptionFactor = l_objcdCodeTables.OvertimeExceptionFactor(v_strExceptionCode)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "OvertimeExceptionFactor", , , , , True

End Function


Public Function RotationCodeIsValid ( _
	ByVal v_strRotCd As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    RotationCodeIsValid = l_objcdCodeTables.RotationCodeIsValid(v_strRotCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "RotationCodeIsValid", , , , , True

End Function


Public Function RotationWeeksIsValid ( _
	ByVal v_strRotCd As String,  _
	ByVal v_lngRot_Weeks As Long _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    RotationWeeksIsValid = l_objcdCodeTables.RotationWeeksIsValid(v_strRotCd, v_lngRot_Weeks)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "RotationWeeksIsValid", , , , , True

End Function


Public Function RotationWeeks ( _
	ByVal v_strRotCd As String _
	)  As  Long
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    RotationWeeks = l_objcdCodeTables.RotationWeeks(v_strRotCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "RotationWeeks", , , , , True

End Function


Public Function ScheduleCodeIsValid ( _
	ByVal v_strSchedCd As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ScheduleCodeIsValid = l_objcdCodeTables.ScheduleCodeIsValid(v_strSchedCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ScheduleCodeIsValid", , , , , True

End Function


Public Function ComponentCodeIsValid ( _
	ByVal v_strCompCd As String,  _
	ByVal v_lngCompType As TrsComponentType _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ComponentCodeIsValid = l_objcdCodeTables.ComponentCodeIsValid(v_strCompCd, v_lngCompType)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ComponentCodeIsValid", , , , , True

End Function


Public Function EmployeeName ( _
	ByVal v_lngEmpNum As Long _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    EmployeeName = l_objcdCodeTables.EmployeeName(v_lngEmpNum)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "EmployeeName", , , , , True

End Function


Public Function GetGroupRecords ( _
	ByVal v_strGroupName As String,  _
	ByVal v_varKeys As Variant,  _
	ByRef r_varKeyArray As Variant,  _
	ByRef r_varColumnArray As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetGroupRecords = l_objcdCodeTables.GetGroupRecords(v_strGroupName, v_varKeys, r_varKeyArray, r_varColumnArray)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetGroupRecords", , , , , True

End Function


Public Sub InvalidateAvailScheduleCache ( _
	ByVal v_strSchedCd As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateAvailScheduleCache v_strSchedCd
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateAvailScheduleCache", , , , , True

End Sub


Public Sub InvalidateRotationCache ( _
	ByVal v_strRotCd As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateRotationCache v_strRotCd
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateRotationCache", , , , , True

End Sub


Public Sub InvalidateExceptionCache ( _
	ByVal v_strExcCd As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateExceptionCache v_strExcCd
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateExceptionCache", , , , , True

End Sub


Public Sub InvalidateOTExceptionCache ( _
	ByVal v_strOTExcCd As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateOTExceptionCache v_strOTExcCd
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateOTExceptionCache", , , , , True

End Sub


Public Sub InvalidateOTReasonCache ( _
	ByVal v_strOTRsnCd As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateOTReasonCache v_strOTRsnCd
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateOTReasonCache", , , , , True

End Sub


Public Sub InvalidateOTResponseCache ( _
	ByVal v_strOTRspCd As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateOTResponseCache v_strOTRspCd
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateOTResponseCache", , , , , True

End Sub


Public Sub InvalidateJobCache ( _
	ByVal v_strJobNum As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateJobCache v_strJobNum
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateJobCache", , , , , True

End Sub


Public Sub InvalidateHSCutoffDateCache ( _
	ByVal v_strHSCutoffKey As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateHSCutoffDateCache v_strHSCutoffKey
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateHSCutoffDateCache", , , , , True

End Sub


Public Sub InvalidateLocationCache ( _
	ByVal v_strLocCd As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateLocationCache v_strLocCd
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateLocationCache", , , , , True

End Sub


Public Sub InvalidateUnionCache ( _
	ByVal v_strUnionCd As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateUnionCache v_strUnionCd
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateUnionCache", , , , , True

End Sub


Public Sub InvalidateUnionComponentCache ( _
	ByVal v_strUnionCd As String,  _
	ByVal v_strCompType As String,  _
	ByVal v_strCompCd As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateUnionComponentCache v_strUnionCd, v_strCompType, v_strCompCd
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateUnionComponentCache", , , , , True

End Sub


Public Sub InvalidateComponentCache ( _
	ByVal v_strCompCd As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateComponentCache v_strCompCd
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateComponentCache", , , , , True

End Sub


Public Sub InvalidateExceptionGroupsCache ( _
	ByVal v_intEgCd As String,  _
	ByVal v_strExcCd As String,  _
	ByVal v_strEgSubCd As String _
	)  
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_objcdCodeTables.InvalidateExceptionGroupsCache v_intEgCd, v_strExcCd, v_strEgSubCd
    
    Set l_objcdCodeTables = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "InvalidateExceptionGroupsCache", , , , , True

End Sub


Public Function GetEffectiveUnionComponents ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetEffectiveUnionComponents = l_objcdCodeTables.GetEffectiveUnionComponents()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetEffectiveUnionComponents", , , , , True

End Function


Public Function GetUnionComponents ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetUnionComponents = l_objcdCodeTables.GetUnionComponents()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetUnionComponents", , , , , True

End Function


Public Function GetActualTimeHSCutoffDate ( _
	ByVal v_strHSKey As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    GetActualTimeHSCutoffDate = l_objcdCodeTables.GetActualTimeHSCutoffDate(v_strHSKey)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetActualTimeHSCutoffDate", , , , , True

End Function


Public Function ExceptionSymbolCd ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionSymbolCd = l_objcdCodeTables.ExceptionSymbolCd(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionSymbolCd", , , , , True

End Function


Public Function GetAvailableScheduleDetails ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAvailableScheduleDetails = l_objcdCodeTables.GetAvailableScheduleDetails()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAvailableScheduleDetails", , , , , True

End Function


Public Function GetAvailableScheduleHeader ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAvailableScheduleHeader = l_objcdCodeTables.GetAvailableScheduleHeader()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAvailableScheduleHeader", , , , , True

End Function


Public Function GetAllActivityList ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllActivityList = l_objcdCodeTables.GetAllActivityList()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllActivityList", , , , , True

End Function


Public Function ValidateScheduleCode ( _
	ByVal v_lScheduleCode As Long _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ValidateScheduleCode = l_objcdCodeTables.ValidateScheduleCode(v_lScheduleCode)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ValidateScheduleCode", , , , , True

End Function


Public Function GetProjectInfo ( _
	Optional ByVal v_intProjectId As Integer,  _
	Optional ByVal v_intPPMInd As String _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetProjectInfo = l_objcdCodeTables.GetProjectInfo(v_intProjectId, v_intPPMInd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetProjectInfo", , , , , True

End Function


Public Function GetLocationsAssociatedToDepartment ( _
	ByRef v_colDeptNums As Collection _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetLocationsAssociatedToDepartment = l_objcdCodeTables.GetLocationsAssociatedToDepartment(v_colDeptNums)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetLocationsAssociatedToDepartment", , , , , True

End Function


Public Function GetDivisionAndFunctionAssociatedToDepartmentAndLocation ( _
	ByVal v_strDeptNum As String,  _
	ByVal v_intLocCd As Integer,  _
	ByRef r_strDivisionNum As String,  _
	ByRef r_strFunctionNum As String _
	)  As  Variant
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    GetDivisionAndFunctionAssociatedToDepartmentAndLocation = l_objcdCodeTables.GetDivisionAndFunctionAssociatedToDepartmentAndLocation(v_strDeptNum, v_intLocCd, r_strDivisionNum, r_strFunctionNum)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetDivisionAndFunctionAssociatedToDepartmentAndLocation", , , , , True

End Function


Public Function GetAllProjects ( _
	ByVal v_datEffective As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllProjects = l_objcdCodeTables.GetAllProjects(v_datEffective)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllProjects", , , , , True

End Function


Public Function GetAllScheduleCodes ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllScheduleCodes = l_objcdCodeTables.GetAllScheduleCodes()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllScheduleCodes", , , , , True

End Function


Public Function GetExceptionGroupForColoring ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetExceptionGroupForColoring = l_objcdCodeTables.GetExceptionGroupForColoring()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetExceptionGroupForColoring", , , , , True

End Function


Public Function ExceptionSchedInd ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionSchedInd = l_objcdCodeTables.ExceptionSchedInd(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionSchedInd", , , , , True

End Function


Public Function IsExceptionInGroup ( _
	ByVal v_lngGroupCode As Long,  _
	ByVal v_strExceptionCode As String,  _
	ByVal v_strSubGroupCode As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsExceptionInGroup = l_objcdCodeTables.IsExceptionInGroup(v_lngGroupCode, v_strExceptionCode, v_strSubGroupCode)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsExceptionInGroup", , , , , True

End Function


Public Function GetUniqueComponentCodes ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetUniqueComponentCodes = l_objcdCodeTables.GetUniqueComponentCodes()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetUniqueComponentCodes", , , , , True

End Function


Public Function GetAllCrafts ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllCrafts = l_objcdCodeTables.GetAllCrafts()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllCrafts", , , , , True

End Function


Public Function GetAllQualifications ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllQualifications = l_objcdCodeTables.GetAllQualifications()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllQualifications", , , , , True

End Function


Public Function GetAllSpecialSkills ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllSpecialSkills = l_objcdCodeTables.GetAllSpecialSkills()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllSpecialSkills", , , , , True

End Function


Public Function GetAllPromotionalLines ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllPromotionalLines = l_objcdCodeTables.GetAllPromotionalLines()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllPromotionalLines", , , , , True

End Function


Public Function GetAllCompanies ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllCompanies = l_objcdCodeTables.GetAllCompanies()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllCompanies", , , , , True

End Function


Public Function IsAValidQualificationCode ( _
	ByVal vsQualificationCode As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAValidQualificationCode = l_objcdCodeTables.IsAValidQualificationCode(vsQualificationCode)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAValidQualificationCode", , , , , True

End Function


Public Function IsAValidSpecialSkillCode ( _
	ByVal vsSpecialSkillCode As String,  _
	ByVal vsQualificationCode As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAValidSpecialSkillCode = l_objcdCodeTables.IsAValidSpecialSkillCode(vsSpecialSkillCode, vsQualificationCode)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAValidSpecialSkillCode", , , , , True

End Function


Public Function IsAValidSpecialSkillCode2 ( _
	ByVal vsSpecialSkillCode As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAValidSpecialSkillCode2 = l_objcdCodeTables.IsAValidSpecialSkillCode2(vsSpecialSkillCode)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAValidSpecialSkillCode2", , , , , True

End Function


Public Function IsAValidDepartment ( _
	ByVal vsDepartmentNumber As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAValidDepartment = l_objcdCodeTables.IsAValidDepartment(vsDepartmentNumber)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAValidDepartment", , , , , True

End Function


Public Function IsAValidCraft ( _
	ByVal vsCraftNumber As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAValidCraft = l_objcdCodeTables.IsAValidCraft(vsCraftNumber)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAValidCraft", , , , , True

End Function


Public Function IsAValidCompany ( _
	ByVal vsCompanyCode As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAValidCompany = l_objcdCodeTables.IsAValidCompany(vsCompanyCode)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAValidCompany", , , , , True

End Function


Public Function IsAValidIndustrialRelationsArea ( _
	ByVal vsIRArea As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAValidIndustrialRelationsArea = l_objcdCodeTables.IsAValidIndustrialRelationsArea(vsIRArea)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAValidIndustrialRelationsArea", , , , , True

End Function


Public Function IsAValidIndustrialRelationsDivision ( _
	ByVal vsIRDivision As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAValidIndustrialRelationsDivision = l_objcdCodeTables.IsAValidIndustrialRelationsDivision(vsIRDivision)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAValidIndustrialRelationsDivision", , , , , True

End Function


Public Function GetAllOvertimeBankHeaders ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllOvertimeBankHeaders = l_objcdCodeTables.GetAllOvertimeBankHeaders()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllOvertimeBankHeaders", , , , , True

End Function


Public Function DoesBankExist ( _
	ByVal viOTBankCode As Integer _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    DoesBankExist = l_objcdCodeTables.DoesBankExist(viOTBankCode)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "DoesBankExist", , , , , True

End Function


Public Function GetAllBudgetNumbers ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllBudgetNumbers = l_objcdCodeTables.GetAllBudgetNumbers()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetAllBudgetNumbers", , , , , True

End Function


Public Function IsAValidBudget ( _
	ByVal vsBudgetNumber As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAValidBudget = l_objcdCodeTables.IsAValidBudget(vsBudgetNumber)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAValidBudget", , , , , True

End Function


Public Function IsAValidOpportunity ( _
	ByVal vlOTOpportunity As Long _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAValidOpportunity = l_objcdCodeTables.IsAValidOpportunity(vlOTOpportunity)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAValidOpportunity", , , , , True

End Function


Public Function IsAnAssignedQualification ( _
	ByVal vsQualification As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAnAssignedQualification = l_objcdCodeTables.IsAnAssignedQualification(vsQualification)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAnAssignedQualification", , , , , True

End Function


Public Function IsAnAssignedSkill ( _
	ByVal vsSpecialSkill As String,  _
	ByVal vsQualification As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsAnAssignedSkill = l_objcdCodeTables.IsAnAssignedSkill(vsSpecialSkill, vsQualification)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsAnAssignedSkill", , , , , True

End Function


Public Function IsCraftUsedinTRIS ( _
	ByVal vsCraftNumber As String,  _
	Optional ByVal vsJobNumber As String = "" _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsCraftUsedinTRIS = l_objcdCodeTables.IsCraftUsedinTRIS(vsCraftNumber, vsJobNumber)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsCraftUsedinTRIS", , , , , True

End Function


Public Function IsOTReasonUsedinTRIS ( _
	ByVal vsOTReason As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsOTReasonUsedinTRIS = l_objcdCodeTables.IsOTReasonUsedinTRIS(vsOTReason)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsOTReasonUsedinTRIS", , , , , True

End Function


Public Function IsOTResponseUsedinTRIS ( _
	ByVal vsOTResponse As String _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsOTResponseUsedinTRIS = l_objcdCodeTables.IsOTResponseUsedinTRIS(vsOTResponse)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsOTResponseUsedinTRIS", , , , , True

End Function


Public Function IsOTBankUsedinTRIS ( _
	ByVal viOTBankCode As Integer _
	)  As  Boolean
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    IsOTBankUsedinTRIS = l_objcdCodeTables.IsOTBankUsedinTRIS(viOTBankCode)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "IsOTBankUsedinTRIS", , , , , True

End Function


Public Function GetJobTaskExceptions ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetJobTaskExceptions = l_objcdCodeTables.GetJobTaskExceptions()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetJobTaskExceptions", , , , , True

End Function


Public Function GetNRCRuleEffectiveDate ( _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    GetNRCRuleEffectiveDate = l_objcdCodeTables.GetNRCRuleEffectiveDate()
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "GetNRCRuleEffectiveDate", , , , , True

End Function


Public Function ExceptionAccountNumberAIC ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionAccountNumberAIC = l_objcdCodeTables.ExceptionAccountNumberAIC(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionAccountNumberAIC", , , , , True

End Function


Public Function ExceptionAccountNumberAMS ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionAccountNumberAMS = l_objcdCodeTables.ExceptionAccountNumberAMS(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionAccountNumberAMS", , , , , True

End Function


Public Function ExceptionAccountNumberUEC ( _
	ByVal v_strExcCd As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionAccountNumberUEC = l_objcdCodeTables.ExceptionAccountNumberUEC(v_strExcCd)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionAccountNumberUEC", , , , , True

End Function


Public Function ExceptionAccountNumberByCompany ( _
	ByVal v_strExcCd As String,  _
	ByVal v_strCompany As String _
	)  As  String
  
    Dim l_objcdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_objcdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    ExceptionAccountNumberByCompany = l_objcdCodeTables.ExceptionAccountNumberByCompany(v_strExcCd, v_strCompany)
    
    Set l_objcdCodeTables = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTables = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTables", "ExceptionAccountNumberByCompany", , , , , True

End Function