VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdWorkGrpProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function GetWorkGroup ( _
	ByVal v_lngEmpNum As Long,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWorkGroup = l_objcdWorkGroup.GetWorkGroup(v_lngEmpNum, v_datStartDate, v_datEndDate)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWorkGroup", , , , , True

End Function


Public Function GetAccounting ( _
	ByVal v_lngWgNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetAccounting = l_objcdWorkGroup.GetAccounting(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetAccounting", , , , , True

End Function


Public Function GetDepartmentNumber ( _
	ByVal v_lngWgNum As Long _
	)  As  String
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    GetDepartmentNumber = l_objcdWorkGroup.GetDepartmentNumber(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetDepartmentNumber", , , , , True

End Function


Public Function GetCrewNumber ( _
	ByVal v_lngWgNum As Long _
	)  As  String
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    GetCrewNumber = l_objcdWorkGroup.GetCrewNumber(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetCrewNumber", , , , , True

End Function


Public Function GetDOJMDistrictAndHeadquarters ( _
	ByVal v_strCrewNum As String _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetDOJMDistrictAndHeadquarters = l_objcdWorkGroup.GetDOJMDistrictAndHeadquarters(v_strCrewNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetDOJMDistrictAndHeadquarters", , , , , True

End Function


Public Function GetEmployeeWorkGroups ( _
	ByVal v_aryEmployees As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByRef bln_PrimaryOnly As Boolean = "False" _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeeWorkGroups = l_objcdWorkGroup.GetEmployeeWorkGroups(v_aryEmployees, v_datStartDate, v_datEndDate, bln_PrimaryOnly)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetEmployeeWorkGroups", , , , , True

End Function


Public Function GetWorkgroupRecord ( _
	ByVal v_lngWgNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWorkgroupRecord = l_objcdWorkGroup.GetWorkgroupRecord(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWorkgroupRecord", , , , , True

End Function


Public Function GetWorkgroupRecords ( _
	ByVal v_lngWgNums As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWorkgroupRecords = l_objcdWorkGroup.GetWorkgroupRecords(v_lngWgNums)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWorkgroupRecords", , , , , True

End Function


Public Function EffectiveStartDate ( _
	ByVal v_lngWgNum As Long _
	)  As  Variant
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    EffectiveStartDate = l_objcdWorkGroup.EffectiveStartDate(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "EffectiveStartDate", , , , , True

End Function


Public Function EffectiveEndDate ( _
	ByVal v_lngWgNum As Long _
	)  As  Variant
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    EffectiveEndDate = l_objcdWorkGroup.EffectiveEndDate(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "EffectiveEndDate", , , , , True

End Function


Public Function ProjectStartDate ( _
	ByVal v_lngWgNum As Long _
	)  As  Variant
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    ProjectStartDate = l_objcdWorkGroup.ProjectStartDate(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "ProjectStartDate", , , , , True

End Function


Public Function ProjectEndDate ( _
	ByVal v_lngWgNum As Long _
	)  As  Variant
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    ProjectEndDate = l_objcdWorkGroup.ProjectEndDate(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "ProjectEndDate", , , , , True

End Function


Public Function ProjectID ( _
	ByVal v_lngWgNum As Long _
	)  As  Long
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    ProjectID = l_objcdWorkGroup.ProjectID(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "ProjectID", , , , , True

End Function


Public Function Supervisor ( _
	ByVal v_lngWgNum As Long _
	)  As  Long
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Supervisor = l_objcdWorkGroup.Supervisor(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "Supervisor", , , , , True

End Function


Public Function DemandIndicator ( _
	ByVal v_lngWgNum As Long _
	)  As  String
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    DemandIndicator = l_objcdWorkGroup.DemandIndicator(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "DemandIndicator", , , , , True

End Function


Public Function WorkgroupDescription ( _
	ByVal v_lngWgNum As Long _
	)  As  String
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    WorkgroupDescription = l_objcdWorkGroup.WorkgroupDescription(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "WorkgroupDescription", , , , , True

End Function


Public Function CrewNumber ( _
	ByVal v_lngWgNum As Long _
	)  As  String
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    CrewNumber = l_objcdWorkGroup.CrewNumber(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "CrewNumber", , , , , True

End Function


Public Function DOJMDistrict ( _
	ByVal v_lngWgNum As Long _
	)  As  String
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    DOJMDistrict = l_objcdWorkGroup.DOJMDistrict(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "DOJMDistrict", , , , , True

End Function


Public Function DOJMHeadquarters ( _
	ByVal v_lngWgNum As Long _
	)  As  String
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    DOJMHeadquarters = l_objcdWorkGroup.DOJMHeadquarters(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "DOJMHeadquarters", , , , , True

End Function


Public Function DepartmentNumber ( _
	ByVal v_lngWgNum As Long _
	)  As  String
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    DepartmentNumber = l_objcdWorkGroup.DepartmentNumber(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "DepartmentNumber", , , , , True

End Function


Public Function LocationCode ( _
	ByVal v_lngWgNum As Long _
	)  As  Long
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    LocationCode = l_objcdWorkGroup.LocationCode(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "LocationCode", , , , , True

End Function


Public Function IsValidWorkgroup ( _
	ByVal v_lngWgNum As Long,  _
	ByVal v_datEffDate As Date _
	)  As  Boolean
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    IsValidWorkgroup = l_objcdWorkGroup.IsValidWorkgroup(v_lngWgNum, v_datEffDate)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "IsValidWorkgroup", , , , , True

End Function


Public Function Record ( _
	ByVal v_lngWgNum As Long _
	)  As  Variant
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Record = l_objcdWorkGroup.Record(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "Record", , , , , True

End Function


Public Function Item ( _
	ByVal v_strFieldName As String,  _
	ByVal v_varRecord As Variant _
	)  As  Variant
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Item = l_objcdWorkGroup.Item(v_strFieldName, v_varRecord)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "Item", , , , , True

End Function


Public Function GetGroupRecords ( _
	ByVal v_strGroupName As String,  _
	ByVal v_varKeys As Variant,  _
	ByRef r_varKeyArray As Variant,  _
	ByRef r_varColumnArray As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetGroupRecords = l_objcdWorkGroup.GetGroupRecords(v_strGroupName, v_varKeys, r_varKeyArray, r_varColumnArray)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetGroupRecords", , , , , True

End Function


Public Sub InvalidateWorkGroupCache ( _
	ByVal v_lngWgNum As Long _
	)  
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    l_objcdWorkGroup.InvalidateWorkGroupCache v_lngWgNum
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "InvalidateWorkGroupCache", , , , , True

End Sub


Public Function GetActivities ( _
	ByVal v_lngWgNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetActivities = l_objcdWorkGroup.GetActivities(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetActivities", , , , , True

End Function


Public Function GetSteps ( _
	ByVal v_sJobTask As String,  _
	ByVal v_sDBID As String _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetSteps = l_objcdWorkGroup.GetSteps(v_sJobTask, v_sDBID)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetSteps", , , , , True

End Function


Public Function GetEmployeesInWorkGroup ( _
	ByVal v_aryWorkGroups As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByVal v_bolActiveEmployeesOnly As Boolean = "False" _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeesInWorkGroup = l_objcdWorkGroup.GetEmployeesInWorkGroup(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetEmployeesInWorkGroup", , , , , True

End Function


Public Function GetEmployeesInWorkGroupTimeCharged ( _
	ByVal v_aryWorkGroups As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByVal v_bolActiveEmployeesOnly As Boolean = "False" _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeesInWorkGroupTimeCharged = l_objcdWorkGroup.GetEmployeesInWorkGroupTimeCharged(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetEmployeesInWorkGroupTimeCharged", , , , , True

End Function


Public Function GetEmployeesInDepartmentTimeCharged ( _
	ByVal v_aryDepartments As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByVal v_bolActiveEmployeesOnly As Boolean = "False" _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeesInDepartmentTimeCharged = l_objcdWorkGroup.GetEmployeesInDepartmentTimeCharged(v_aryDepartments, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetEmployeesInDepartmentTimeCharged", , , , , True

End Function


Public Function GetEmployeesAssignedToWorkGroup ( _
	ByVal v_aryWorkGroups As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByVal v_bolActiveEmployeesOnly As Boolean = "False" _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeesAssignedToWorkGroup = l_objcdWorkGroup.GetEmployeesAssignedToWorkGroup(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetEmployeesAssignedToWorkGroup", , , , , True

End Function


Public Function GetEmployeeAssignmentsInWorkGroup ( _
	ByVal v_lngWorkGroup As Long,  _
	ByVal v_datDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeeAssignmentsInWorkGroup = l_objcdWorkGroup.GetEmployeeAssignmentsInWorkGroup(v_lngWorkGroup, v_datDate)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetEmployeeAssignmentsInWorkGroup", , , , , True

End Function


Public Function GetEmployeesInWorkGroupViewTime ( _
	ByVal v_aryWorkGroups As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByVal v_bolActiveEmployeesOnly As Boolean = "False" _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeesInWorkGroupViewTime = l_objcdWorkGroup.GetEmployeesInWorkGroupViewTime(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetEmployeesInWorkGroupViewTime", , , , , True

End Function


Public Function GetWorkGroupAssignments ( _
	ByVal v_lngWorkGroup As Long,  _
	ByVal v_datDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWorkGroupAssignments = l_objcdWorkGroup.GetWorkGroupAssignments(v_lngWorkGroup, v_datDate)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWorkGroupAssignments", , , , , True

End Function


Public Function GetEmployeeWorkGroupAssignments ( _
	ByVal v_aryEmployees As Variant,  _
	ByVal v_datDate As Date,  _
	ByVal v_lngWorkGroup As Long,  _
	Optional ByRef v_blnPrimaryOnly As Boolean = "True" _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeeWorkGroupAssignments = l_objcdWorkGroup.GetEmployeeWorkGroupAssignments(v_aryEmployees, v_datDate, v_lngWorkGroup, v_blnPrimaryOnly)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetEmployeeWorkGroupAssignments", , , , , True

End Function


Public Function GetWGDemandLineAssigmments ( _
	ByVal v_lngWgNum As Long,  _
	ByVal v_datDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWGDemandLineAssigmments = l_objcdWorkGroup.GetWGDemandLineAssigmments(v_lngWgNum, v_datDate)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWGDemandLineAssigmments", , , , , True

End Function


Public Function GetWorkGroupAssociations ( _
	ByVal v_lngWgNums As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWorkGroupAssociations = l_objcdWorkGroup.GetWorkGroupAssociations(v_lngWgNums)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWorkGroupAssociations", , , , , True

End Function


Public Function GetWorkgroupAttributes ( _
	ByVal v_aryWgNums As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWorkgroupAttributes = l_objcdWorkGroup.GetWorkgroupAttributes(v_aryWgNums)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWorkgroupAttributes", , , , , True

End Function


Public Function GetWorkGroupActivities ( _
	ByVal v_lngWgNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWorkGroupActivities = l_objcdWorkGroup.GetWorkGroupActivities(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWorkGroupActivities", , , , , True

End Function


Public Function GetWorkGroupDistribution ( _
	ByVal v_lngWgNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWorkGroupDistribution = l_objcdWorkGroup.GetWorkGroupDistribution(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWorkGroupDistribution", , , , , True

End Function


Public Function GetWorkgroupForUpdate ( _
	ByVal v_aryWgNums As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWorkgroupForUpdate = l_objcdWorkGroup.GetWorkgroupForUpdate(v_aryWgNums)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWorkgroupForUpdate", , , , , True

End Function


Public Function GetDemandLineBaseSchedules ( _
	ByVal v_aryWgNums As Variant,  _
	ByVal v_datDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetDemandLineBaseSchedules = l_objcdWorkGroup.GetDemandLineBaseSchedules(v_aryWgNums, v_datDate)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetDemandLineBaseSchedules", , , , , True

End Function


Public Function GetBaseSchedules ( _
	ByVal v_aryWgNums As Variant,  _
	ByVal v_datDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetBaseSchedules = l_objcdWorkGroup.GetBaseSchedules(v_aryWgNums, v_datDate)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetBaseSchedules", , , , , True

End Function


Public Function GetDemandLineSchedules ( _
	ByVal v_lngWorkGroup As Long,  _
	ByVal v_datMinEffDate As Date,  _
	ByVal v_datMaxEffDate As Date,  _
	ByVal v_varLineNums As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetDemandLineSchedules = l_objcdWorkGroup.GetDemandLineSchedules(v_lngWorkGroup, v_datMinEffDate, v_datMaxEffDate, v_varLineNums)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetDemandLineSchedules", , , , , True

End Function


Public Function GetWorkGroupDistributionForUpdate ( _
	ByVal v_lngWgNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWorkGroupDistributionForUpdate = l_objcdWorkGroup.GetWorkGroupDistributionForUpdate(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWorkGroupDistributionForUpdate", , , , , True

End Function


Public Function GetWorkGroupActivitiesForUpdate ( _
	ByVal v_lngWgNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetWorkGroupActivitiesForUpdate = l_objcdWorkGroup.GetWorkGroupActivitiesForUpdate(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetWorkGroupActivitiesForUpdate", , , , , True

End Function


Public Function GetGLActivities ( _
	ByVal v_strActivities As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetGLActivities = l_objcdWorkGroup.GetGLActivities(v_strActivities)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetGLActivities", , , , , True

End Function


Public Function GetEmployeesForWGSchedRpt ( _
	ByVal vlWGNum As Long,  _
	ByVal vblnUnmannedLines As Boolean,  _
	ByVal vblnDemandLine As Boolean,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeesForWGSchedRpt = l_objcdWorkGroup.GetEmployeesForWGSchedRpt(vlWGNum, vblnUnmannedLines, vblnDemandLine, vdtStartDate, vdtEndDate)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetEmployeesForWGSchedRpt", , , , , True

End Function


Public Function GetEmployeesForDailyRosterRpt ( _
	ByVal vlWGNum As Long,  _
	ByVal vdtRequestDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeesForDailyRosterRpt = l_objcdWorkGroup.GetEmployeesForDailyRosterRpt(vlWGNum, vdtRequestDate)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetEmployeesForDailyRosterRpt", , , , , True

End Function


Public Function GetEmployeesForDailyRosterRpt2 ( _
	ByVal vlWGNum As Long,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeesForDailyRosterRpt2 = l_objcdWorkGroup.GetEmployeesForDailyRosterRpt2(vlWGNum, vdtStartDate, vdtEndDate)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetEmployeesForDailyRosterRpt2", , , , , True

End Function


Public Function GetActiveWorkgroups ( _
	ByVal vlngEmpNum As Long,  _
	ByVal vdtDate As Date _
	)  As  Long
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    GetActiveWorkgroups = l_objcdWorkGroup.GetActiveWorkgroups(vlngEmpNum, vdtDate)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "GetActiveWorkgroups", , , , , True

End Function


Public Function IsJobTaskWorkgroup ( _
	ByVal v_lngWgNum As Long _
	)  As  Boolean
  
    Dim l_objcdWorkGroup As dbWorkGroup.cdWorkGroup
    
On Error GoTo eh

    Set l_objcdWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    
    IsJobTaskWorkgroup = l_objcdWorkGroup.IsJobTaskWorkgroup(v_lngWgNum)
    
    Set l_objcdWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroup = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroup", "IsJobTaskWorkgroup", , , , , True

End Function