VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsRules"

'//this was orginally defined in dbRules.cdRules, but need it for the client
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
End Enum

Public Function GetRule( _
    ByVal v_lngRule As Long, _
    Optional ByVal v_strSubRule As String = "" _
    ) As ADODB.Recordset
    
    Dim l_obj As dbRules.cdRules

    On Error GoTo eh

    Set l_obj = CreateObject("dbRules.cdRules")
    Set GetRule = l_obj.GetRule(v_lngRule, v_strSubRule)
    Set l_obj = Nothing
    
Exit Function
eh:
    Set l_obj = Nothing
    Set GetRule = Nothing

    TrsRaiseError m_strCMODULENAME, "GetRule", , , , , True
End Function

'//for the client to cache multiple rules in a single recordset
Public Function GetRules(ByVal v_varRules As Variant) As ADODB.Recordset
    Dim l_obj As dbRules.cdRules

    On Error GoTo eh

    Set l_obj = CreateObject("dbRules.cdRules")
    Set GetRules = l_obj.GetRules(v_varRules)
    Set l_obj = Nothing
Exit Function
eh:
    Set l_obj = Nothing
    Set GetRules = Nothing

    TrsRaiseError m_strCMODULENAME, "GetRules", , , , , True
End Function

Public Function GetEffectiveDateRule(ByVal v_lngRule As Long, Optional ByVal v_strSubRule As String = "") As ADODB.Recordset
    Dim l_obj As dbRules.cdEffectiveDateRules

    On Error GoTo eh

    Set l_obj = CreateObject("dbRules.cdEffectiveDateRules")
    Set GetEffectiveDateRule = l_obj.GetEffectiveDateRule(v_lngRule, v_strSubRule)
    Set l_obj = Nothing
    
Exit Function
eh:
    Set l_obj = Nothing
    Set GetEffectiveDateRule = Nothing

    TrsRaiseError m_strCMODULENAME, "GetEffectiveDateRule", , , , , True
End Function

Public Function GetEffectiveDateRules(ByVal v_varRules As Variant) As ADODB.Recordset
    Dim l_obj As dbRules.cdEffectiveDateRules

    On Error GoTo eh

    Set l_obj = CreateObject("dbRules.cdEffectiveDateRules")
    Set GetEffectiveDateRules = l_obj.GetEffectiveDateRules(v_varRules)
    Set l_obj = Nothing
    
Exit Function
eh:
    Set l_obj = Nothing
    Set GetEffectiveDateRules = Nothing

    TrsRaiseError m_strCMODULENAME, "GetEffectiveDateRules", , , , , True
End Function
