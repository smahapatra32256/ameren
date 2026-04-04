VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsWorkGrpProxy"
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
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetWorkGroup = l_objcbsWorkGroup.GetWorkGroup(v_lngEmpNum, v_datStartDate, v_datEndDate)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetWorkGroup", , , , , True

End Function


Public Function GetAccounting ( _
	ByVal v_lngWgNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetAccounting = l_objcbsWorkGroup.GetAccounting(v_lngWgNum)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetAccounting", , , , , True

End Function


Public Function GetDOJMDistrictAndHeadquarters ( _
	ByVal v_lngWGNumber As Long,  _
	ByRef r_strDOJMDistrict As Variant,  _
	ByRef r_strDOJMHeadQtr As Variant _
	)  As  Variant
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    GetDOJMDistrictAndHeadquarters = l_objcbsWorkGroup.GetDOJMDistrictAndHeadquarters(v_lngWGNumber, r_strDOJMDistrict, r_strDOJMHeadQtr)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetDOJMDistrictAndHeadquarters", , , , , True

End Function


Public Function GetEmployeesInWorkGroup ( _
	ByVal v_aryWorkGroups As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByRef v_bolActiveEmployeesOnly As Boolean = "False" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetEmployeesInWorkGroup = l_objcbsWorkGroup.GetEmployeesInWorkGroup(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetEmployeesInWorkGroup", , , , , True

End Function


Public Function GetActivities ( _
	ByVal v_lngWgNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetActivities = l_objcbsWorkGroup.GetActivities(v_lngWgNum)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetActivities", , , , , True

End Function


Public Function GetSteps ( _
	ByVal v_sJobTask As String,  _
	ByVal v_sDBID As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetSteps = l_objcbsWorkGroup.GetSteps(v_sJobTask, v_sDBID)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetSteps", , , , , True

End Function


Public Function GetTransferWorkGroup ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetTransferWorkGroup = l_objcbsWorkGroup.GetTransferWorkGroup()
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetTransferWorkGroup", , , , , True

End Function


Public Function GetEmployeeAssignmentsInWorkGroup ( _
	ByVal v_lngWorkGroup As Long,  _
	ByVal v_datDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetEmployeeAssignmentsInWorkGroup = l_objcbsWorkGroup.GetEmployeeAssignmentsInWorkGroup(v_lngWorkGroup, v_datDate)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetEmployeeAssignmentsInWorkGroup", , , , , True

End Function


Public Function GetWorkGroupDistribution ( _
	ByVal v_lngWgNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetWorkGroupDistribution = l_objcbsWorkGroup.GetWorkGroupDistribution(v_lngWgNum)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetWorkGroupDistribution", , , , , True

End Function


Public Function GetWorkGroupActivities ( _
	ByVal v_lngWgNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetWorkGroupActivities = l_objcbsWorkGroup.GetWorkGroupActivities(v_lngWgNum)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetWorkGroupActivities", , , , , True

End Function


Public Function GetWorkGroupAttributes ( _
	ByVal v_lngWorkGroups As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetWorkGroupAttributes = l_objcbsWorkGroup.GetWorkGroupAttributes(v_lngWorkGroups)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetWorkGroupAttributes", , , , , True

End Function


Public Function GetEmployeesAssignedToWorkGroup ( _
	ByVal v_aryWorkGroups As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByRef v_bolActiveEmployeesOnly As Boolean = "False" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetEmployeesAssignedToWorkGroup = l_objcbsWorkGroup.GetEmployeesAssignedToWorkGroup(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetEmployeesAssignedToWorkGroup", , , , , True

End Function


Public Function GetEmployeesInWorkGroupViewTime ( _
	ByVal v_aryWorkGroups As Variant,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	Optional ByRef v_bolActiveEmployeesOnly As Boolean = "False" _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetEmployeesInWorkGroupViewTime = l_objcbsWorkGroup.GetEmployeesInWorkGroupViewTime(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetEmployeesInWorkGroupViewTime", , , , , True

End Function


Public Function GetWGSchedule ( _
	ByVal vlngWGNum As Long,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date,  _
	ByVal vblnUnmannedLines As Boolean _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetWGSchedule = l_objcbsWorkGroup.GetWGSchedule(vlngWGNum, vdtStartDate, vdtEndDate, vblnUnmannedLines)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetWGSchedule", , , , , True

End Function


Public Function GetDailyRosterSchedule ( _
	ByVal vlngWGNum As Long,  _
	ByVal vdtRequestDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetDailyRosterSchedule = l_objcbsWorkGroup.GetDailyRosterSchedule(vlngWGNum, vdtRequestDate)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetDailyRosterSchedule", , , , , True

End Function


Public Function GetDailyRosterSchedule2 ( _
	ByVal vlngWGNum As Long,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetDailyRosterSchedule2 = l_objcbsWorkGroup.GetDailyRosterSchedule2(vlngWGNum, vdtStartDate, vdtEndDate)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetDailyRosterSchedule2", , , , , True

End Function


Public Function GetDailyRosterSchedule3 ( _
	ByVal vlngWGNum As Long,  _
	ByVal varrEmployees As Variant,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    Set GetDailyRosterSchedule3 = l_objcbsWorkGroup.GetDailyRosterSchedule3(vlngWGNum, varrEmployees, vdtStartDate, vdtEndDate)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetDailyRosterSchedule3", , , , , True

End Function


Public Function GetJobDesc ( _
	ByVal vlngJobNum As Variant _
	)  As  String
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    GetJobDesc = l_objcbsWorkGroup.GetJobDesc(vlngJobNum)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "GetJobDesc", , , , , True

End Function


Public Function IsJobTaskWorkgroup ( _
	ByVal v_lngWgNum As Long _
	)  As  Boolean
  
    Dim l_objcbsWorkGroup As busWorkGroup.cbsWorkGroup
    
On Error GoTo eh

    Set l_objcbsWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    
    IsJobTaskWorkgroup = l_objcbsWorkGroup.IsJobTaskWorkgroup(v_lngWgNum)
    
    Set l_objcbsWorkGroup = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroup = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroup", "IsJobTaskWorkgroup", , , , , True

End Function