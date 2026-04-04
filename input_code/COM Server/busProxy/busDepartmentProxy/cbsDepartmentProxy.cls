VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsDepartmentProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function GetDepartment ( _
	ByVal v_strDeptNum As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    Set GetDepartment = l_objcbsDepartment.GetDepartment(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "GetDepartment", , , , , True

End Function


Public Function GetAllDepartments ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    Set GetAllDepartments = l_objcbsDepartment.GetAllDepartments()
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "GetAllDepartments", , , , , True

End Function


Public Function IsDepartmentValid ( _
	ByVal v_strDeptNum As String _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsDepartmentValid = l_objcbsDepartment.IsDepartmentValid(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsDepartmentValid", , , , , True

End Function


Public Function GetPayrollGroup ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    GetPayrollGroup = l_objcbsDepartment.GetPayrollGroup(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "GetPayrollGroup", , , , , True

End Function


Public Function GetDepartmentLockStatus ( _
	ByVal v_strDeptNum As String _
	)  As  String
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    GetDepartmentLockStatus = l_objcbsDepartment.GetDepartmentLockStatus(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "GetDepartmentLockStatus", , , , , True

End Function


Public Function GetAccountType ( _
	ByVal v_strDeptNum As Variant _
	)  As  String
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    GetAccountType = l_objcbsDepartment.GetAccountType(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "GetAccountType", , , , , True

End Function


Public Function GetSystemInterfaceCode ( _
	ByVal v_strDeptNum As Variant _
	)  As  String
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    GetSystemInterfaceCode = l_objcbsDepartment.GetSystemInterfaceCode(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "GetSystemInterfaceCode", , , , , True

End Function


Public Function IsDOJMDepartmentNumber ( _
	ByVal v_strDeptNum As Variant _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsDOJMDepartmentNumber = l_objcbsDepartment.IsDOJMDepartmentNumber(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsDOJMDepartmentNumber", , , , , True

End Function


Public Function IsEMPRVDepartmentNumber ( _
	ByVal v_strDeptNum As Variant _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsEMPRVDepartmentNumber = l_objcbsDepartment.IsEMPRVDepartmentNumber(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsEMPRVDepartmentNumber", , , , , True

End Function


Public Function IsJobTaskDepartmentNumber ( _
	ByVal v_strDeptNum As Variant _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsJobTaskDepartmentNumber = l_objcbsDepartment.IsJobTaskDepartmentNumber(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsJobTaskDepartmentNumber", , , , , True

End Function


Public Function IsDepartmentException ( _
	ByVal v_strDeptNum As Variant _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsDepartmentException = l_objcbsDepartment.IsDepartmentException(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsDepartmentException", , , , , True

End Function


Public Function IsDepartmentVariable ( _
	ByVal v_strDeptNum As Variant _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsDepartmentVariable = l_objcbsDepartment.IsDepartmentVariable(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsDepartmentVariable", , , , , True

End Function


Public Function IsDepartmentFixed ( _
	ByVal v_strDeptNum As Variant _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsDepartmentFixed = l_objcbsDepartment.IsDepartmentFixed(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsDepartmentFixed", , , , , True

End Function


Public Function IsECDepartmentNumber ( _
	ByVal v_strDeptNum As Variant _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsECDepartmentNumber = l_objcbsDepartment.IsECDepartmentNumber(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsECDepartmentNumber", , , , , True

End Function


Public Function IsWorkReasonIndicatorRequired ( _
	ByVal v_strDeptNum As String _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsWorkReasonIndicatorRequired = l_objcbsDepartment.IsWorkReasonIndicatorRequired(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsWorkReasonIndicatorRequired", , , , , True

End Function


Public Function IsCoffeenMaintenance ( _
	ByVal v_strDeptNum As String _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsCoffeenMaintenance = l_objcbsDepartment.IsCoffeenMaintenance(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsCoffeenMaintenance", , , , , True

End Function


Public Function IsCoffeenOperations ( _
	ByVal v_strDeptNum As String _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsCoffeenOperations = l_objcbsDepartment.IsCoffeenOperations(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsCoffeenOperations", , , , , True

End Function


Public Function IsCoffeen ( _
	ByVal v_strDeptNum As String _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsCoffeen = l_objcbsDepartment.IsCoffeen(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsCoffeen", , , , , True

End Function


Public Function IsEdwardsOperations ( _
	ByVal v_strDeptNum As String _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsEdwardsOperations = l_objcbsDepartment.IsEdwardsOperations(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsEdwardsOperations", , , , , True

End Function


Public Function IsAmerenCIPS ( _
	ByVal v_strDeptNum As String,  _
	Optional ByVal v_strCompanyCode As String = "NA" _
	)  As  String
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsAmerenCIPS = l_objcbsDepartment.IsAmerenCIPS(v_strDeptNum, v_strCompanyCode)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsAmerenCIPS", , , , , True

End Function


Public Function IsAmerenCILCO ( _
	ByVal v_strDeptNum As String,  _
	Optional ByVal v_strCompanyCode As String = "NA" _
	)  As  String
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsAmerenCILCO = l_objcbsDepartment.IsAmerenCILCO(v_strDeptNum, v_strCompanyCode)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsAmerenCILCO", , , , , True

End Function


Public Function IsAmerenCIPSorCILCO ( _
	ByVal v_lngRuleNum As Long,  _
	ByVal v_strDeptNum As String,  _
	Optional ByVal v_strCompanyCode As String = "NA",  _
	Optional ByVal v_strUnionCode As String = "Z#",  _
	Optional ByVal v_datChgDate As Date = "1/1/100",  _
	Optional ByVal v_strPayrollGroupCd As String = "Z#" _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsAmerenCIPSorCILCO = l_objcbsDepartment.IsAmerenCIPSorCILCO(v_lngRuleNum, v_strDeptNum, v_strCompanyCode, v_strUnionCode, v_datChgDate, v_strPayrollGroupCd)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsAmerenCIPSorCILCO", , , , , True

End Function


Public Function IsAmerenCILCODispatch ( _
	ByVal v_strDeptNum As String _
	)  As  Boolean
  
    Dim l_objcbsDepartment As busDepartment.cbsDepartment
    
On Error GoTo eh

    Set l_objcbsDepartment = CreateObject("busDepartment.cbsDepartment")
    
    IsAmerenCILCODispatch = l_objcbsDepartment.IsAmerenCILCODispatch(v_strDeptNum)
    
    Set l_objcbsDepartment = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDepartment = Nothing

    TrsRaiseError "busDepartment.cbsDepartment", "IsAmerenCILCODispatch", , , , , True

End Function