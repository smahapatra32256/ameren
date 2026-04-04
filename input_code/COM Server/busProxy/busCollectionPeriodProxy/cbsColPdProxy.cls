VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsColPdProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function GetCollectionPeriodStatusByEmployee ( _
	ByVal v_lngEmpNum As Long,  _
	Optional ByVal v_datEffectiveDate As Date = "12:00:00 AM",  _
	Optional ByVal v_intMaxNoOfPayPeriods As Integer _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCollectionPeriod As busCollectionPeriod.cbsCollectionPeriod
    
On Error GoTo eh

    Set l_objcbsCollectionPeriod = CreateObject("busCollectionPeriod.cbsCollectionPeriod")
    
    Set GetCollectionPeriodStatusByEmployee = l_objcbsCollectionPeriod.GetCollectionPeriodStatusByEmployee(v_lngEmpNum, v_datEffectiveDate, v_intMaxNoOfPayPeriods)
    
    Set l_objcbsCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCollectionPeriod = Nothing

    TrsRaiseError "busCollectionPeriod.cbsCollectionPeriod", "GetCollectionPeriodStatusByEmployee", , , , , True

End Function


Public Function GetPayProcIndInRange ( _
	ByVal v_strPayrollGroupCode As String,  _
	ByVal v_datChargeDate As Date _
	)  As  String
  
    Dim l_objcbsCollectionPeriod As busCollectionPeriod.cbsCollectionPeriod
    
On Error GoTo eh

    Set l_objcbsCollectionPeriod = CreateObject("busCollectionPeriod.cbsCollectionPeriod")
    
    GetPayProcIndInRange = l_objcbsCollectionPeriod.GetPayProcIndInRange(v_strPayrollGroupCode, v_datChargeDate)
    
    Set l_objcbsCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCollectionPeriod = Nothing

    TrsRaiseError "busCollectionPeriod.cbsCollectionPeriod", "GetPayProcIndInRange", , , , , True

End Function


Public Function GetCollectionPeriodLockStatus ( _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPrevDeptNum As String,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCollectionPeriod As busCollectionPeriod.cbsCollectionPeriod
    
On Error GoTo eh

    Set l_objcbsCollectionPeriod = CreateObject("busCollectionPeriod.cbsCollectionPeriod")
    
    Set GetCollectionPeriodLockStatus = l_objcbsCollectionPeriod.GetCollectionPeriodLockStatus(v_strDeptNum, v_strPrevDeptNum, v_datStartDate, v_datEndDate)
    
    Set l_objcbsCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCollectionPeriod = Nothing

    TrsRaiseError "busCollectionPeriod.cbsCollectionPeriod", "GetCollectionPeriodLockStatus", , , , , True

End Function


Public Function GetCollectionPeriodsByDepartment ( _
	ByVal v_strDeptNum As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCollectionPeriod As busCollectionPeriod.cbsCollectionPeriod
    
On Error GoTo eh

    Set l_objcbsCollectionPeriod = CreateObject("busCollectionPeriod.cbsCollectionPeriod")
    
    Set GetCollectionPeriodsByDepartment = l_objcbsCollectionPeriod.GetCollectionPeriodsByDepartment(v_strDeptNum)
    
    Set l_objcbsCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCollectionPeriod = Nothing

    TrsRaiseError "busCollectionPeriod.cbsCollectionPeriod", "GetCollectionPeriodsByDepartment", , , , , True

End Function


Public Function GetCollectionPeriodsByWorkGroup ( _
	ByVal v_lngWorkGroup As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsCollectionPeriod As busCollectionPeriod.cbsCollectionPeriod
    
On Error GoTo eh

    Set l_objcbsCollectionPeriod = CreateObject("busCollectionPeriod.cbsCollectionPeriod")
    
    Set GetCollectionPeriodsByWorkGroup = l_objcbsCollectionPeriod.GetCollectionPeriodsByWorkGroup(v_lngWorkGroup)
    
    Set l_objcbsCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCollectionPeriod = Nothing

    TrsRaiseError "busCollectionPeriod.cbsCollectionPeriod", "GetCollectionPeriodsByWorkGroup", , , , , True

End Function


Public Function GetLatestProcessedPayPeriodEndDate ( _
	ByVal v_strPayrollGroupCode As String _
	)  As  Date
  
    Dim l_objcbsCollectionPeriod As busCollectionPeriod.cbsCollectionPeriod
    
On Error GoTo eh

    Set l_objcbsCollectionPeriod = CreateObject("busCollectionPeriod.cbsCollectionPeriod")
    
    GetLatestProcessedPayPeriodEndDate = l_objcbsCollectionPeriod.GetLatestProcessedPayPeriodEndDate(v_strPayrollGroupCode)
    
    Set l_objcbsCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCollectionPeriod = Nothing

    TrsRaiseError "busCollectionPeriod.cbsCollectionPeriod", "GetLatestProcessedPayPeriodEndDate", , , , , True

End Function