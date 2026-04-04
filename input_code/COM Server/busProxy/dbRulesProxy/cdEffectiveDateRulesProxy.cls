VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdEffectiveDateRulesProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//
'//Public Enum TRSRulesDataTypes
'//    trExceptionCode = 1
'//    trEmployeeNumber = 2
'//    trBenPkgInd = 3
'//    trDepartmentNumber = 4
'//    trCompanyCode = 5
'//    trUnionCode = 6
'//    trJobNumber = 7
'//    trLocationCode = 8
'//    trPayrollGroupCode = 23
'//    trDeptNumAndUnionCode = 75
'//End Enum
'//
'//Public Enum TRSRuleCode
'//    trIsAmerenCIPS = 11
'//    trIsAmerenCILCO = 12
'//    trAICDeptMapping = 203
'//    trAICEffectiveDates = 206
'//    trIsSubstationOrTransmissionEmployee = 207
'//    trIsCIPSForSickLeaveBanks = 209
'//    trAICDeptCutOverEffDate = 211
'//End Enum
'//
'//




Public Function DoesEffectiveDateRuleApply ( _
	ByVal v_lngRule As Long,  _
	ByVal v_lngDataType As TRSRulesDataTypes,  _
	ByVal v_strValue As String,  _
	ByVal v_datDate As Date _
	)  As  Boolean
  
    Dim l_objcdEffectiveDateRules As dbRules.cdEffectiveDateRules
    
On Error GoTo eh

    Set l_objcdEffectiveDateRules = CreateObject("dbRules.cdEffectiveDateRules")
    
    DoesEffectiveDateRuleApply = l_objcdEffectiveDateRules.DoesEffectiveDateRuleApply(v_lngRule, v_lngDataType, v_strValue, v_datDate)
    
    Set l_objcdEffectiveDateRules = Nothing
    
    Exit Function
    
eh:
    Set l_objcdEffectiveDateRules = Nothing

    TrsRaiseError "dbRules.cdEffectiveDateRules", "DoesEffectiveDateRuleApply", , , , , True

End Function


Public Function GetGroupRecords ( _
	ByVal v_strGroupName As String,  _
	ByVal v_varKeys As Variant,  _
	ByRef r_varKeyArray As Variant,  _
	ByRef r_varColumnArray As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdEffectiveDateRules As dbRules.cdEffectiveDateRules
    
On Error GoTo eh

    Set l_objcdEffectiveDateRules = CreateObject("dbRules.cdEffectiveDateRules")
    
    Set GetGroupRecords = l_objcdEffectiveDateRules.GetGroupRecords(v_strGroupName, v_varKeys, r_varKeyArray, r_varColumnArray)
    
    Set l_objcdEffectiveDateRules = Nothing
    
    Exit Function
    
eh:
    Set l_objcdEffectiveDateRules = Nothing

    TrsRaiseError "dbRules.cdEffectiveDateRules", "GetGroupRecords", , , , , True

End Function


Public Sub InvalidateRulesCache ( _
	ByVal v_lngRule As Long,  _
	ByVal v_lngBusTypeCd As Long,  _
	ByVal v_strRgValueCd As String _
	)  
  
    Dim l_objcdEffectiveDateRules As dbRules.cdEffectiveDateRules
    
On Error GoTo eh

    Set l_objcdEffectiveDateRules = CreateObject("dbRules.cdEffectiveDateRules")
    
    l_objcdEffectiveDateRules.InvalidateRulesCache v_lngRule, v_lngBusTypeCd, v_strRgValueCd
    
    Set l_objcdEffectiveDateRules = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdEffectiveDateRules = Nothing

    TrsRaiseError "dbRules.cdEffectiveDateRules", "InvalidateRulesCache", , , , , True

End Sub


Public Function GetEffectiveDateRule ( _
	ByVal v_lngRule As Long,  _
	Optional ByVal v_strSubRule As String = "" _
	)  As  ADODB.Recordset
  
    Dim l_objcdEffectiveDateRules As dbRules.cdEffectiveDateRules
    
On Error GoTo eh

    Set l_objcdEffectiveDateRules = CreateObject("dbRules.cdEffectiveDateRules")
    
    Set GetEffectiveDateRule = l_objcdEffectiveDateRules.GetEffectiveDateRule(v_lngRule, v_strSubRule)
    
    Set l_objcdEffectiveDateRules = Nothing
    
    Exit Function
    
eh:
    Set l_objcdEffectiveDateRules = Nothing

    TrsRaiseError "dbRules.cdEffectiveDateRules", "GetEffectiveDateRule", , , , , True

End Function


Public Function GetEffectiveDateRules ( _
	ByVal v_varRules As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdEffectiveDateRules As dbRules.cdEffectiveDateRules
    
On Error GoTo eh

    Set l_objcdEffectiveDateRules = CreateObject("dbRules.cdEffectiveDateRules")
    
    Set GetEffectiveDateRules = l_objcdEffectiveDateRules.GetEffectiveDateRules(v_varRules)
    
    Set l_objcdEffectiveDateRules = Nothing
    
    Exit Function
    
eh:
    Set l_objcdEffectiveDateRules = Nothing

    TrsRaiseError "dbRules.cdEffectiveDateRules", "GetEffectiveDateRules", , , , , True

End Function