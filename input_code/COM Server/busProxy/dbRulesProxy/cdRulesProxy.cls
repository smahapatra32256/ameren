VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdRulesProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//
Public Enum TRSRulesDataTypes
    trExceptionCode = 1
    trEmployeeNumber = 2
    trBenPkgInd = 3
    trDepartmentNumber = 4
    trCompanyCode = 5
    trUnionCode = 6
    trJobNumber = 7
    trLocationCode = 8
    trPayrollGroupCode = 23
    trDeptNumAndUnionCode = 75
End Enum

Public Enum TRSRuleCode
    trIsAmerenCIPS = 11
    trIsAmerenCILCO = 12
    trAICDeptMapping = 203
    trAICEffectiveDates = 206
    trIsSubstationOrTransmissionEmployee = 207
    trIsCIPSForSickLeaveBanks = 209
    trAICDeptCutOverEffDate = 211
End Enum





Public Function GetRule ( _
	ByVal v_lngRule As Long,  _
	Optional ByVal v_strSubRule As String = "" _
	)  As  ADODB.Recordset
  
    Dim l_objcdRules As dbRules.cdRules
    
On Error GoTo eh

    Set l_objcdRules = CreateObject("dbRules.cdRules")
    
    Set GetRule = l_objcdRules.GetRule(v_lngRule, v_strSubRule)
    
    Set l_objcdRules = Nothing
    
    Exit Function
    
eh:
    Set l_objcdRules = Nothing

    TrsRaiseError "dbRules.cdRules", "GetRule", , , , , True

End Function


Public Function GetRules ( _
	ByVal v_varRules As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdRules As dbRules.cdRules
    
On Error GoTo eh

    Set l_objcdRules = CreateObject("dbRules.cdRules")
    
    Set GetRules = l_objcdRules.GetRules(v_varRules)
    
    Set l_objcdRules = Nothing
    
    Exit Function
    
eh:
    Set l_objcdRules = Nothing

    TrsRaiseError "dbRules.cdRules", "GetRules", , , , , True

End Function


Public Function DoesRuleApply ( _
	ByVal v_lngRule As Long,  _
	ByVal v_lngDataType As TRSRulesDataTypes,  _
	ByVal v_strValue As String _
	)  As  Boolean
  
    Dim l_objcdRules As dbRules.cdRules
    
On Error GoTo eh

    Set l_objcdRules = CreateObject("dbRules.cdRules")
    
    DoesRuleApply = l_objcdRules.DoesRuleApply(v_lngRule, v_lngDataType, v_strValue)
    
    Set l_objcdRules = Nothing
    
    Exit Function
    
eh:
    Set l_objcdRules = Nothing

    TrsRaiseError "dbRules.cdRules", "DoesRuleApply", , , , , True

End Function


Public Function GetRuleLookup ( _
	ByVal v_lngRule As Long,  _
	ByVal v_lngDataType As TRSRulesDataTypes,  _
	ByVal v_strValuePart1 As String _
	)  As  String
  
    Dim l_objcdRules As dbRules.cdRules
    
On Error GoTo eh

    Set l_objcdRules = CreateObject("dbRules.cdRules")
    
    GetRuleLookup = l_objcdRules.GetRuleLookup(v_lngRule, v_lngDataType, v_strValuePart1)
    
    Set l_objcdRules = Nothing
    
    Exit Function
    
eh:
    Set l_objcdRules = Nothing

    TrsRaiseError "dbRules.cdRules", "GetRuleLookup", , , , , True

End Function


Public Function GetRuleReverseLookup ( _
	ByVal v_lngRule As Long,  _
	ByVal v_lngDataType As TRSRulesDataTypes,  _
	ByVal v_strValuePart2 As String _
	)  As  String
  
    Dim l_objcdRules As dbRules.cdRules
    
On Error GoTo eh

    Set l_objcdRules = CreateObject("dbRules.cdRules")
    
    GetRuleReverseLookup = l_objcdRules.GetRuleReverseLookup(v_lngRule, v_lngDataType, v_strValuePart2)
    
    Set l_objcdRules = Nothing
    
    Exit Function
    
eh:
    Set l_objcdRules = Nothing

    TrsRaiseError "dbRules.cdRules", "GetRuleReverseLookup", , , , , True

End Function


Public Function GetGroupRecords ( _
	ByVal v_strGroupName As String,  _
	ByVal v_varKeys As Variant,  _
	ByRef r_varKeyArray As Variant,  _
	ByRef r_varColumnArray As Variant _
	)  As  ADODB.Recordset
  
    Dim l_objcdRules As dbRules.cdRules
    
On Error GoTo eh

    Set l_objcdRules = CreateObject("dbRules.cdRules")
    
    Set GetGroupRecords = l_objcdRules.GetGroupRecords(v_strGroupName, v_varKeys, r_varKeyArray, r_varColumnArray)
    
    Set l_objcdRules = Nothing
    
    Exit Function
    
eh:
    Set l_objcdRules = Nothing

    TrsRaiseError "dbRules.cdRules", "GetGroupRecords", , , , , True

End Function


Public Sub InvalidateRulesCache ( _
	ByVal v_lngRule As Long,  _
	ByVal v_lngBusTypeCd As Long,  _
	ByVal v_strRgValueCd As String _
	)  
  
    Dim l_objcdRules As dbRules.cdRules
    
On Error GoTo eh

    Set l_objcdRules = CreateObject("dbRules.cdRules")
    
    l_objcdRules.InvalidateRulesCache v_lngRule, v_lngBusTypeCd, v_strRgValueCd
    
    Set l_objcdRules = Nothing
    
    Exit Sub
    
eh:
    Set l_objcdRules = Nothing

    TrsRaiseError "dbRules.cdRules", "InvalidateRulesCache", , , , , True

End Sub


Public Function GetEffectiveDateRule ( _
	ByVal v_lngRule As Long _
	)  As  ADODB.Recordset
  
    Dim l_objcdRules As dbRules.cdRules
    
On Error GoTo eh

    Set l_objcdRules = CreateObject("dbRules.cdRules")
    
    Set GetEffectiveDateRule = l_objcdRules.GetEffectiveDateRule(v_lngRule)
    
    Set l_objcdRules = Nothing
    
    Exit Function
    
eh:
    Set l_objcdRules = Nothing

    TrsRaiseError "dbRules.cdRules", "GetEffectiveDateRule", , , , , True

End Function