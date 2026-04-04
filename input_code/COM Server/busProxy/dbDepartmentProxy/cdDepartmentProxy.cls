VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdDepartmentProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function GetDepartment ( _
	ByVal v_strDeptNum As String _
	)  As  ADODB.Recordset
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    Set GetDepartment = l_objcdDepartment.GetDepartment(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetDepartment", , , , , True

End Function


Public Function GetDepartments ( _
	ByRef v_colDeptNums As Collection _
	)  As  ADODB.Recordset
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    Set GetDepartments = l_objcdDepartment.GetDepartments(v_colDeptNums)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetDepartments", , , , , True

End Function


Public Function GetAllDepartments ( _
	)  As  ADODB.Recordset
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    Set GetAllDepartments = l_objcdDepartment.GetAllDepartments()
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetAllDepartments", , , , , True

End Function


Public Function GetDepartmentLockInfo ( _
	ByVal v_lngEmpNum As Long,  _
	ByVal v_datChargeDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    Set GetDepartmentLockInfo = l_objcdDepartment.GetDepartmentLockInfo(v_lngEmpNum, v_datChargeDate)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetDepartmentLockInfo", , , , , True

End Function


Public Function GetInterfaceInd ( _
	ByVal v_strDeptNum As String _
	)  As  ADODB.Recordset
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    Set GetInterfaceInd = l_objcdDepartment.GetInterfaceInd(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetInterfaceInd", , , , , True

End Function


Public Function PayrollGroupCodeFor ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    PayrollGroupCodeFor = l_objcdDepartment.PayrollGroupCodeFor(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "PayrollGroupCodeFor", , , , , True

End Function


Public Function IsDepartmentValid ( _
	ByVal v_strDeptNum As String _
	)  As  Boolean
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    IsDepartmentValid = l_objcdDepartment.IsDepartmentValid(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "IsDepartmentValid", , , , , True

End Function


Public Function GetDepartmentName ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetDepartmentName = l_objcdDepartment.GetDepartmentName(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetDepartmentName", , , , , True

End Function


Public Function GetSystemInterfaceCode ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetSystemInterfaceCode = l_objcdDepartment.GetSystemInterfaceCode(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetSystemInterfaceCode", , , , , True

End Function


Public Function GetCompanyCode ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetCompanyCode = l_objcdDepartment.GetCompanyCode(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetCompanyCode", , , , , True

End Function


Public Function GetPayrollGroup ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetPayrollGroup = l_objcdDepartment.GetPayrollGroup(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetPayrollGroup", , , , , True

End Function


Public Function GetUnionCode ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetUnionCode = l_objcdDepartment.GetUnionCode(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetUnionCode", , , , , True

End Function


Public Function GetNonProdLoadInd ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetNonProdLoadInd = l_objcdDepartment.GetNonProdLoadInd(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetNonProdLoadInd", , , , , True

End Function


Public Function GetAccountType ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetAccountType = l_objcdDepartment.GetAccountType(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetAccountType", , , , , True

End Function


Public Function GetAccountNumber ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetAccountNumber = l_objcdDepartment.GetAccountNumber(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetAccountNumber", , , , , True

End Function


Public Function GetAccountDivision ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetAccountDivision = l_objcdDepartment.GetAccountDivision(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetAccountDivision", , , , , True

End Function


Public Function GetOvertimeReasonCode ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetOvertimeReasonCode = l_objcdDepartment.GetOvertimeReasonCode(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetOvertimeReasonCode", , , , , True

End Function


Public Function GetECWorkReasonInd ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetECWorkReasonInd = l_objcdDepartment.GetECWorkReasonInd(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetECWorkReasonInd", , , , , True

End Function


Public Function GetECDrawDetailsInd ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetECDrawDetailsInd = l_objcdDepartment.GetECDrawDetailsInd(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetECDrawDetailsInd", , , , , True

End Function


Public Function GetWorkReasonInd ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetWorkReasonInd = l_objcdDepartment.GetWorkReasonInd(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetWorkReasonInd", , , , , True

End Function


Public Function GetDepartmentLockStatus ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetDepartmentLockStatus = l_objcdDepartment.GetDepartmentLockStatus(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetDepartmentLockStatus", , , , , True

End Function


Public Function GetBudgetNumber ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    GetBudgetNumber = l_objcdDepartment.GetBudgetNumber(v_strDeptNum)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetBudgetNumber", , , , , True

End Function


Public Function GetGroupRecords ( _
	ByVal v_strGroupName As String,  _
	ByVal v_varKeys As Variant,  _
	ByRef r_varKeyArray As Variant,  _
	ByRef r_varColumnArray As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    Set GetGroupRecords = l_objcdDepartment.GetGroupRecords(v_strGroupName, v_varKeys, r_varKeyArray, r_varColumnArray)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetGroupRecords", , , , , True

End Function


Public Sub InvalidateDepartmentCache ( _
	ByVal v_strDeptNum As String _
	)  
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    l_objcdDepartment.InvalidateDepartmentCache v_strDeptNum
    
    Set l_objcdDepartment = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "InvalidateDepartmentCache", , , , , True

End Sub


Public Function GetEmpsinDept ( _
	ByVal v_strDeptNum As String,  _
	ByVal v_dtChargeDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    Set GetEmpsinDept = l_objcdDepartment.GetEmpsinDept(v_strDeptNum, v_dtChargeDate)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetEmpsinDept", , , , , True

End Function


Public Function GetEmpsinDept2 ( _
	ByVal v_strDeptNum As String,  _
	ByVal vdtStartDate As Date,  _
	ByVal vdtEndDate As Date _
	)  As  ADODB.Recordset
  
    Dim l_objcdDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh

    Set l_objcdDepartment = CreateObject("dbDepartment.cdDepartment")
    
    Set GetEmpsinDept2 = l_objcdDepartment.GetEmpsinDept2(v_strDeptNum, vdtStartDate, vdtEndDate)
    
    Set l_objcdDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcdDepartment = Nothing

    TrsRaiseError "dbDepartment.cdDepartment", "GetEmpsinDept2", , , , , True

End Function