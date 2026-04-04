VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cbsDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsDepartment"
Private Const mstrDefaultDate As String = "01/01/0100"
Private Const mstrNotAvailable As String = "NA" 'keep for compatability
Private Const mstrNotAvailable2 As String = "Z#" 'zee pound or zee sharp?


Public Function GetDepartment( _
            ByVal v_strDeptNum As Long) As ADODB.Recordset
    
    Dim l_objWkGp As dbDepartment.cdDepartment
    
    Dim l_objRS As Recordset
    
    On Error GoTo GetDepartment_EH
    
    
    Set l_objWkGp = CreateObject("dbDepartment.cdDepartment")

    Set l_objRS = l_objWkGp.GetDepartment(v_strDeptNum)
    Set GetDepartment = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    Exit Function
    
GetDepartment_EH:
    Set l_objWkGp = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDepartment", , , , , True

End Function

Public Function GetAllDepartments() As ADODB.Recordset

    Dim l_objDepartment As dbDepartment.cdDepartment
    

    On Error GoTo GetAllDepartments_EH

    
    Set l_objDepartment = CreateObject("dbDepartment.cdDepartment")
    
    Set GetAllDepartments = l_objDepartment.GetAllDepartments
    Set l_objDepartment = Nothing
    Exit Function

GetAllDepartments_EH:
    Set l_objDepartment = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllDepartments", , , , , True

End Function

Public Function IsDepartmentValid( _
            ByVal v_strDeptNum As String) As Boolean
    Dim l_objDepartment As dbDepartment.cdDepartment
    
    
On Error GoTo eh
    
    Set l_objDepartment = _
        CreateObject("dbDepartment.cdDepartment")
    IsDepartmentValid = l_objDepartment.IsDepartmentValid(v_strDeptNum)
    Set l_objDepartment = Nothing
    Exit Function

eh:
    Set l_objDepartment = Nothing
    TrsRaiseError m_strCMODULENAME, "IsDepartmentValid", , , , , True
End Function

Public Function GetPayrollGroup( _
            ByVal v_strDeptNum As String) As String
    Dim l_objDepartment As dbDepartment.cdDepartment
    
    
On Error GoTo eh
    
    Set l_objDepartment = CreateObject("dbDepartment.cdDepartment")
    GetPayrollGroup = l_objDepartment.GetPayrollGroup(v_strDeptNum)
    Set l_objDepartment = Nothing
    Exit Function
    
eh:
    Set l_objDepartment = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPayrollGroup", , , , , True
End Function

Public Function GetDepartmentLockStatus( _
            ByVal v_strDeptNum As String) As String
    Dim l_objDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh
    
    Set l_objDepartment = CreateObject("dbDepartment.cdDepartment")
    GetDepartmentLockStatus = l_objDepartment.GetDepartmentLockStatus(v_strDeptNum)
    Set l_objDepartment = Nothing
    Exit Function

eh:
    Set l_objDepartment = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDepartmentLockStatus", , , , , True
    
End Function

Public Function GetAccountType(ByVal v_strDeptNum) As String
    Dim l_objDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh
    Set l_objDepartment = _
        CreateObject("dbDepartment.cdDepartment")
    GetAccountType = l_objDepartment.GetAccountType(v_strDeptNum)
    Set l_objDepartment = Nothing
    Exit Function

eh:
    Set l_objDepartment = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAccountType", , , , , True
    
End Function

Public Function GetSystemInterfaceCode(ByVal v_strDeptNum) As String
    Dim l_objDepartment As dbDepartment.cdDepartment
    
On Error GoTo eh
    Set l_objDepartment = _
        CreateObject("dbDepartment.cdDepartment")
    GetSystemInterfaceCode = _
        l_objDepartment.GetSystemInterfaceCode(v_strDeptNum)
    Set l_objDepartment = Nothing
    Exit Function

eh:
    Set l_objDepartment = Nothing
    TrsRaiseError m_strCMODULENAME, "GetSystemInterfaceCode", , , , , True

End Function


Public Function IsDOJMDepartmentNumber(ByVal v_strDeptNum) As Boolean
' BLUE037
    Dim l_objDepartment As dbDepartment.cdDepartment
    
    
On Error GoTo eh
    
    Set l_objDepartment = _
        CreateObject("dbDepartment.cdDepartment")
    IsDOJMDepartmentNumber = _
        (l_objDepartment.GetSystemInterfaceCode(v_strDeptNum) = "D")
    Set l_objDepartment = Nothing
    Exit Function

eh:
    Set l_objDepartment = Nothing
    TrsRaiseError m_strCMODULENAME, "IsDOJMDepartmentNumber", , , , , True

End Function

Public Function IsEMPRVDepartmentNumber(ByVal v_strDeptNum) As Boolean
'WHITE199
    Dim l_objDepartment As dbDepartment.cdDepartment

On Error GoTo eh
    
    Set l_objDepartment = _
        CreateObject("dbDepartment.cdDepartment")
    IsEMPRVDepartmentNumber = _
        (l_objDepartment.GetSystemInterfaceCode(v_strDeptNum) = "E")
    Set l_objDepartment = Nothing
    Exit Function

eh:
    Set l_objDepartment = Nothing
    TrsRaiseError m_strCMODULENAME, "IsEMPRVDepartmentNumber", , , , , True

End Function
Public Function IsJobTaskDepartmentNumber(ByVal v_strDeptNum) As Boolean
'WHITE199
    Dim l_objDepartment As dbDepartment.cdDepartment
    Dim l_rsInterfaceInd As ADODB.Recordset

On Error GoTo eh
    
    Set l_objDepartment = _
        CreateObject("dbDepartment.cdDepartment")
    Set l_rsInterfaceInd = l_objDepartment.GetInterfaceInd(v_strDeptNum)
    
    If Not l_rsInterfaceInd Is Nothing Then
        l_rsInterfaceInd.MoveFirst
        If Not (l_rsInterfaceInd.BOF And l_rsInterfaceInd.EOF) Then
            If l_rsInterfaceInd.Fields("INTERFACE_IND") = "E" Then
                IsJobTaskDepartmentNumber = True
            Else
                IsJobTaskDepartmentNumber = False
            End If
        Else
            IsJobTaskDepartmentNumber = False
        End If
    Else
        IsJobTaskDepartmentNumber = False
    End If

    Set l_objDepartment = Nothing
    Exit Function

eh:
    Set l_objDepartment = Nothing
    Set l_rsInterfaceInd = Nothing
    TrsRaiseError m_strCMODULENAME, "IsJobTaskDepartmentNumber", , , , , True

End Function

Public Function IsDepartmentException(ByVal v_strDeptNum) As Boolean

On Error GoTo eh

    IsDepartmentException = _
        (GetAccountType(v_strDeptNum) = "E")
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsDepartmentException", , , , , True
    
End Function

Public Function IsDepartmentVariable(ByVal v_strDeptNum) As Boolean

On Error GoTo eh

    IsDepartmentVariable = _
        (GetAccountType(v_strDeptNum) = "V")
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsDepartmentVariable", , , , , True
    
End Function


Public Function IsDepartmentFixed(ByVal v_strDeptNum) As Boolean

On Error GoTo eh

    IsDepartmentFixed = _
        (GetAccountType(v_strDeptNum) = "F")
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsDepartmentFixed", , , , , True
    
End Function

Public Function IsECDepartmentNumber(ByVal v_strDeptNum) As Boolean
TrsRaiseError m_strCMODULENAME, "IsECDepartmentNumber", trErrObsoleteMethod, , , , True
'On Error GoTo eh
'    IsECDepartmentNumber = _
'        (GetSystemInterfaceCode(v_strDeptNum) = "A")
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "IsECDepartmentNumber", , , , , True
'
End Function

Public Function IsWorkReasonIndicatorRequired( _
            ByVal v_strDeptNum As String) As Boolean
TrsRaiseError m_strCMODULENAME, "IsWorkReasonIndicatorRequired", trErrObsoleteMethod, , , , True
'    Dim l_objDepartment As dbDepartment.cdDepartment
'
'On Error GoTo eh
'
'    Set l_objDepartment = _
'        CreateObject("dbDepartment.cdDepartment")
'    IsWorkReasonIndicatorRequired = _
'        (l_objDepartment.GetECWorkReasonInd(v_strDeptNum) = "Y")
'    Set l_objDepartment = Nothing
'    Exit Function
'
'eh:
'    Set l_objDepartment = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetECWorkReasonInd", , , , , True

End Function

Public Function IsCoffeenMaintenance( _
        ByVal v_strDeptNum As String) As Boolean
' RED155
'RULE0009
    
    On Error GoTo eh
    
    IsCoffeenMaintenance = DoesRuleApply(9, trDepartmentNumber, v_strDeptNum)

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsCoffeenMaintenance", , , , , True

End Function
    
Public Function IsCoffeenOperations( _
        ByVal v_strDeptNum As String) As Boolean
' RED156
'RULE0010

    On Error GoTo eh
    
    IsCoffeenOperations = DoesRuleApply(10, trDepartmentNumber, v_strDeptNum)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsCoffeenOperations", , , , , True

End Function

Public Function IsCoffeen( _
        ByVal v_strDeptNum As String) As Boolean
' Currently, this is used to determine shift component
' which has a special rule for coffeen maintenance and
' operations.  However, this is not inclusive of all
' coffeen.  Location code should be used to determine that.
    On Error GoTo eh
    
    IsCoffeen = IsCoffeenMaintenance(v_strDeptNum) Or _
                    IsCoffeenOperations(v_strDeptNum)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsCoffeen", , , , , True

End Function

Public Function IsEdwardsOperations( _
        ByVal v_strDeptNum As String) As Boolean
' CILCO
'RULE0013

    On Error GoTo eh
    
    IsEdwardsOperations = DoesRuleApply(13, trDepartmentNumber, v_strDeptNum)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsEdwardsOperations", , , , , True

End Function

Public Function IsAmerenCIPS( _
            ByVal v_strDeptNum As String, _
            Optional ByVal v_strCompanyCode As String = mstrNotAvailable _
            ) As String

TrsRaiseError m_strCMODULENAME, "IsAmerenCIPS", trErrObsoleteMethod, , , , True
''RULE0011
'
'     Dim l_strCompany As String
'
'On Error GoTo eh
'
'
'    l_strCompany = GetCompanyCode(v_strDeptNum, v_strCompanyCode)
'    IsAmerenCIPS = DoesRuleApply(11, trCompanyCode, l_strCompany)

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPS", , , , , True
End Function

Private Function GetCompanyCode( _
            ByVal v_strDept As String, _
            ByVal v_strCompanyCode As String) As String
            
    Dim l_obj As dbDepartment.cdDepartment
    
On Error GoTo eh
    
    If v_strCompanyCode = mstrNotAvailable Then
        Set l_obj = CreateObject("dbDepartment.cdDepartment")
        GetCompanyCode = _
            l_obj.GetCompanyCode(v_strDept)
        Set l_obj = Nothing
    Else
        GetCompanyCode = v_strCompanyCode
    End If
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCompanyCode"
End Function

Private Function GetUnionCode( _
            ByVal v_strDept As String, _
            ByVal v_strUnionCode As String) As String
            
    Dim l_obj As dbDepartment.cdDepartment
On Error GoTo eh
    
    If v_strUnionCode = mstrNotAvailable2 Then
        Set l_obj = CreateObject("dbDepartment.cdDepartment")
        GetUnionCode = l_obj.GetUnionCode(v_strDept)
        Set l_obj = Nothing
    Else
        GetUnionCode = v_strUnionCode
    End If
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetUnionCode"
End Function

Public Function IsAmerenCILCO( _
            ByVal v_strDeptNum As String, _
            Optional ByVal v_strCompanyCode As String = mstrNotAvailable _
        ) As String
        
TrsRaiseError m_strCMODULENAME, "IsAmerenCILCO", trErrObsoleteMethod, , , , True
''CILCO
''RULE0012
'
'     Dim l_strCompany As String
'
'On Error GoTo eh
'
'    l_strCompany = GetCompanyCode(v_strDeptNum, v_strCompanyCode)
'    IsAmerenCILCO = DoesRuleApply(12, trCompanyCode, l_strCompany)
        
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCILCO", , , , , True
End Function

'//refactored to pass in rule number rather than to copy and paste
Public Function IsAmerenCIPSorCILCO( _
            ByVal v_lngRuleNum As Long, _
            ByVal v_strDeptNum As String, _
            Optional ByVal v_strCompanyCode As String = mstrNotAvailable, _
            Optional ByVal v_strUnionCode As String = mstrNotAvailable2, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate, _
            Optional ByVal v_strPayrollGroupCd As String = mstrNotAvailable2 _
        ) As Boolean
'CILCO
'RULE0012
            
    Dim l_strCompany As String
    Dim l_strPayrollGroupCd As String
On Error GoTo eh
    
    If v_strPayrollGroupCd = mstrNotAvailable2 Then
        l_strPayrollGroupCd = GetPayrollGroup(v_strDeptNum)
    Else
        l_strPayrollGroupCd = v_strPayrollGroupCd
    End If
    
    If DoesEffectiveDateRuleApply(dbRules.TRSRuleCode.trAICEffectiveDates, trPayrollGroupCode, l_strPayrollGroupCd, v_datChgDate) Then
        IsAmerenCIPSorCILCO = IsAmerenCIPSorCILCOByUnion(v_lngRuleNum, v_strDeptNum, v_strUnionCode, v_datChgDate)
    Else
        IsAmerenCIPSorCILCO = IsAmerenCIPSorCILCOByCompany(v_lngRuleNum, v_strDeptNum, v_strCompanyCode, v_datChgDate)
    End If
            
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPSorCILCO", , , , , True
End Function

Private Function IsAmerenCIPSorCILCOByCompany(ByVal v_lngRuleNum As Long, _
                                        ByVal v_strDeptNum As String, _
                                Optional ByVal v_strCompanyCode As String = mstrNotAvailable, _
                                Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
    Dim l_strCompany As String
On Error GoTo eh
    l_strCompany = GetCompanyCode(v_strDeptNum, v_strCompanyCode)
    
    IsAmerenCIPSorCILCOByCompany = DoesRuleApply(v_lngRuleNum, trCompanyCode, l_strCompany)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPSorCILCOByCompany", , , , , True
End Function

Private Function IsAmerenCIPSorCILCOByUnion(ByVal v_lngRuleNum As Long, _
                                            ByVal v_strDeptNum As String, _
                                    Optional ByVal v_strUnionCode As String = mstrNotAvailable2, _
                                    Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
    Dim l_strUnion As String
On Error GoTo eh
    l_strUnion = GetUnionCode(v_strDeptNum, v_strUnionCode)

    IsAmerenCIPSorCILCOByUnion = DoesRuleApply(v_lngRuleNum, trUnionCode, l_strUnion)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPSorCILCOByUnion", , , , , True
End Function

Public Function IsAmerenCILCODispatch( _
        ByVal v_strDeptNum As String) As Boolean
' CILCO
'RULE0015

    On Error GoTo eh
    
    IsAmerenCILCODispatch = DoesRuleApply(15, trDepartmentNumber, v_strDeptNum)

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsAmerenCILCODispatch", , , , , True

End Function

Private Function DoesRuleApply( _
    ByVal v_lngRuleNum As Long, _
    ByVal v_lngBusType As TRSRulesDataTypes, _
    ByVal v_strValue As String _
    ) As Boolean
       
    Dim l_obj As dbRules.cdRules
       
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    DoesRuleApply = l_obj.DoesRuleApply(v_lngRuleNum, v_lngBusType, v_strValue)
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "DoesRuleApply"
    
End Function

Private Function DoesEffectiveDateRuleApply( _
    ByVal v_lngRuleNum As Long, _
    ByVal v_lngBusType As TRSRulesDataTypes, _
    ByVal v_strValue As String, _
    ByVal v_datDate As Date) As Boolean
       
    Dim l_obj As dbRules.cdEffectiveDateRules
On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdEffectiveDateRules")
    DoesEffectiveDateRuleApply = l_obj.DoesEffectiveDateRuleApply(v_lngRuleNum, v_lngBusType, v_strValue, v_datDate)
    Set l_obj = Nothing
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "DoesEffectiveDateRuleApply"
End Function

