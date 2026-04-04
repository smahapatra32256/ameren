VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdColPdProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function GetCollectionPeriodStatusByPayGroup ( _
	ByVal v_datDate As Date,  _
	ByVal v_strPayrollCode As String _
	)  As  ADODB.Recordset
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    Set GetCollectionPeriodStatusByPayGroup = l_objcdCollectionPeriod.GetCollectionPeriodStatusByPayGroup(v_datDate, v_strPayrollCode)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "GetCollectionPeriodStatusByPayGroup", , , , , True

End Function


Public Function GetCollPerPayProcInd ( _
	ByVal v_datDate As Date,  _
	ByVal v_strPayrollCode As String _
	)  As  String
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    GetCollPerPayProcInd = l_objcdCollectionPeriod.GetCollPerPayProcInd(v_datDate, v_strPayrollCode)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "GetCollPerPayProcInd", , , , , True

End Function


Public Function GetPayProcIndInRange ( _
	ByVal v_strPayrollGroupCode As String,  _
	ByVal v_datChargeDate As Date _
	)  As  String
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    GetPayProcIndInRange = l_objcdCollectionPeriod.GetPayProcIndInRange(v_strPayrollGroupCode, v_datChargeDate)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "GetPayProcIndInRange", , , , , True

End Function


Public Function GetCollectionPeriodLockStatus ( _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPrevDeptNum As String,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    Set GetCollectionPeriodLockStatus = l_objcdCollectionPeriod.GetCollectionPeriodLockStatus(v_strDeptNum, v_strPrevDeptNum, v_datStartDate, v_datEndDate)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "GetCollectionPeriodLockStatus", , , , , True

End Function


Public Function GetCollectionPeriods ( _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	ByVal v_strPayrollCodes As Collection _
	)  As  ADODB.Recordset
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    Set GetCollectionPeriods = l_objcdCollectionPeriod.GetCollectionPeriods(v_datStartDate, v_datEndDate, v_strPayrollCodes)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "GetCollectionPeriods", , , , , True

End Function


Public Function GetCollectionPeriodsByDepartment ( _
	ByVal v_strDeptNum As String,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    Set GetCollectionPeriodsByDepartment = l_objcdCollectionPeriod.GetCollectionPeriodsByDepartment(v_strDeptNum, v_datStartDate, v_datEndDate)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "GetCollectionPeriodsByDepartment", , , , , True

End Function


Public Function GetCollectionPeriodsByWorkGroup ( _
	ByVal v_lngWorkGroup As Long,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    Set GetCollectionPeriodsByWorkGroup = l_objcdCollectionPeriod.GetCollectionPeriodsByWorkGroup(v_lngWorkGroup, v_datStartDate, v_datEndDate)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "GetCollectionPeriodsByWorkGroup", , , , , True

End Function


Public Function GetLatestProcessedPayPeriodEndDate ( _
	ByVal v_strPayrollGroupCode As String _
	)  As  Date
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    GetLatestProcessedPayPeriodEndDate = l_objcdCollectionPeriod.GetLatestProcessedPayPeriodEndDate(v_strPayrollGroupCode)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "GetLatestProcessedPayPeriodEndDate", , , , , True

End Function


Public Function PayPeriodStartDate ( _
	ByVal v_strPayrollGrpCd As String,  _
	ByVal v_datChgDate As Date _
	)  As  Variant
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    PayPeriodStartDate = l_objcdCollectionPeriod.PayPeriodStartDate(v_strPayrollGrpCd, v_datChgDate)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "PayPeriodStartDate", , , , , True

End Function


Public Function PayPeriodEndDate ( _
	ByVal v_strPayrollGrpCd As String,  _
	ByVal v_datChgDate As Date _
	)  As  Variant
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    PayPeriodEndDate = l_objcdCollectionPeriod.PayPeriodEndDate(v_strPayrollGrpCd, v_datChgDate)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "PayPeriodEndDate", , , , , True

End Function


Public Function PayProcessedIndicator ( _
	ByVal v_strPayrollGrpCd As String,  _
	ByVal v_datChgDate As Date _
	)  As  Variant
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    PayProcessedIndicator = l_objcdCollectionPeriod.PayProcessedIndicator(v_strPayrollGrpCd, v_datChgDate)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "PayProcessedIndicator", , , , , True

End Function


Public Function GetGroupRecords ( _
	ByVal v_strGroupName As String,  _
	ByVal v_varKeys As Variant,  _
	ByRef r_varKeyArray As Variant,  _
	ByRef r_varColumnArray As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    Set GetGroupRecords = l_objcdCollectionPeriod.GetGroupRecords(v_strGroupName, v_varKeys, r_varKeyArray, r_varColumnArray)
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "GetGroupRecords", , , , , True

End Function


Public Sub InvalidateCollectionPeriodCache ( _
	ByVal v_strPayrollGrpCd As String _
	)  
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    l_objcdCollectionPeriod.InvalidateCollectionPeriodCache v_strPayrollGrpCd
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "InvalidateCollectionPeriodCache", , , , , True

End Sub


Public Function GetCurrentCollectionPeriods ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
    
On Error GoTo eh

    Set l_objcdCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    
    Set GetCurrentCollectionPeriods = l_objcdCollectionPeriod.GetCurrentCollectionPeriods()
    
    Set l_objcdCollectionPeriod = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCollectionPeriod = Nothing

    TrsRaiseError "dbCollectionPeriod.cdCollectionPeriod", "GetCurrentCollectionPeriods", , , , , True

End Function