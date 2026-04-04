Attribute VB_Name = "mbusTimecard"
Option Explicit

Public Enum TrsObjectStatusEnum
    trObjectRestored = 8
    trObjectNew = 1
    trObjectModified = 2
    trObjectDeleted = 4
End Enum

Private Const mstrCMODULENAME As String = "mbusTimecard"


Public Function ExecutiveContractIndicatorForUnion( _
                    ByVal v_strUnionCd As String) As String
                    
On Error GoTo eh

    ExecutiveContractIndicatorForUnion = _
        GetUnionManager.UnionExecutiveContractIndicator(v_strUnionCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ExecutiveContractIndicatorForUnion"
End Function

Public Function EarningsCodeForExceptionCode( _
            ByVal v_strExcCd As String, _
            ByVal v_strExecConInd As String) As String

On Error GoTo eh
    EarningsCodeForExceptionCode = _
        GetExceptionCodeManager.ExceptionEarningsCode( _
                v_strExcCd, v_strExecConInd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "EarningsCodeForExceptionCode", Err.Number
    
End Function

Public Function ExceptionBankCodeFor( _
            ByVal v_strExcCd As String) As String
' BLUE076
On Error GoTo eh

    ExceptionBankCodeFor = _
        GetExceptionCodeManager.ExceptionBankCode(v_strExcCd)
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "ExceptionBankCodeFor", Err.Number
End Function

Public Function IsExceptionOvertimeRelated( _
            ByVal v_strExcCd As String) As Boolean
' WHITE056, WHITE070
On Error GoTo eh

    IsExceptionOvertimeRelated = _
        GetBusExceptionCodeManager.ExceptionIsOvertimeRelated(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsExceptionOvertimeRelated", Err.Number
End Function

Public Function IsExceptionOvertimeBankRelated( _
            ByVal v_strExcCd As String) As Boolean
' BLUE075, BLUE076, WHITE056
On Error GoTo eh

    IsExceptionOvertimeBankRelated = _
        GetBusExceptionCodeManager.ExceptionIsOvertimeBankRelated(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsExceptionOvertimeBankRelated", Err.Number
End Function


Public Function IsExceptionUnpaid( _
            ByVal v_strExcCd As String, _
            ByVal v_strExecConInd As String) As Boolean
            
    IsExceptionUnpaid = _
        GetBusExceptionCodeManager.ExceptionIsUnpaid(v_strExcCd, v_strExecConInd)
            
End Function

Public Function IsExceptionCaseRelated( _
            ByVal v_strExcCd As String) As Boolean

On Error GoTo eh

    IsExceptionCaseRelated = _
        GetBusExceptionCodeManager.ExceptionIsCaseRelated(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsExceptionCaseRelated", Err.Number
End Function

Public Function IsExceptionValidInsideOfSchedule( _
            ByVal v_strExcCd As String) As Boolean
' GOLD003, BLUE090
On Error GoTo eh

    IsExceptionValidInsideOfSchedule = _
        GetBusExceptionCodeManager.ExceptionIsValidInsideOfSchedule(v_strExcCd) _
        And v_strExcCd <> "ONP"
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsExceptionValidInsideOfSchedule", Err.Number
End Function

Public Function IsExceptionValidOutsideOfSchedule( _
            ByVal v_strExcCd As String) As Boolean
' GOLD002, BLUE090
On Error GoTo eh

    IsExceptionValidOutsideOfSchedule = _
        GetBusExceptionCodeManager.ExceptionIsValidOutsideOfSchedule(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsExceptionValidOutsideOfSchedule", Err.Number
End Function

Public Function ExceptionBankTypeFor( _
            ByVal v_strExcCd As String) As String
' Answer the first character of bank code.
' BLUE021
On Error GoTo eh
    ExceptionBankTypeFor = _
        GetExceptionCodeManager.ExceptionBankType(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ExceptionBankTypeFor", Err.Number
End Function

Public Function IsExceptionCode( _
            ByVal v_strExcCd As String) As Boolean
' BLUE094, WHITE043
            
On Error GoTo eh
    IsExceptionCode = _
        GetExceptionCodeManager.ExceptionCodeIsValid(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsExceptionCode", Err.Number
        
End Function


Public Function IsValidExceptionCodeFor( _
            ByVal v_strExcCd As String, _
            ByVal v_strExecConInd As String, _
            ByVal v_strBenefitsInd) As Boolean
' BLUE094, WHITE043
            
On Error GoTo eh

    IsValidExceptionCodeFor = _
        GetBusExceptionCodeManager.ExceptionCodeIsValidFor(v_strExcCd, v_strExecConInd, v_strBenefitsInd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsValidExceptionCodeFor", Err.Number
        
End Function


Public Function IsOtherPayExceptionCode( _
            ByVal v_strExcCd As String) As Boolean
' BLUE094
            
On Error GoTo eh

    IsOtherPayExceptionCode = _
        GetBusExceptionCodeManager.ExceptionIsOtherPay(v_strExcCd)
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsOtherPayExceptionCode", Err.Number
        
End Function

Public Function IsOverrideExceptionCode( _
            ByVal v_strExcCd As String) As Boolean
    Dim l_strExcType As String
            
On Error GoTo eh

    IsOverrideExceptionCode = _
        GetBusExceptionCodeManager.ExceptionIsOverride(v_strExcCd)
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsOverrideExceptionCode", Err.Number
        
End Function


Public Function IsEligibleForScheduleChangePayment( _
            ByVal v_strExcCd As String) As Boolean
            
On Error GoTo eh

    IsEligibleForScheduleChangePayment = _
        GetBusExceptionCodeManager.ExceptionIsEligibleForScheduleChangePayment(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsEligibleForScheduleChangePayment", Err.Number
    
End Function

Public Function IsEligibleForFamilyLeave( _
            ByVal v_strExcCd As String) As Boolean
'WHITE066
On Error GoTo eh
    IsEligibleForFamilyLeave = _
        GetBusExceptionCodeManager.ExceptionIsEligibleForFamilyLeave(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsEligibleForFamilyLeave", Err.Number
    
End Function

Public Function IsValidScheduleChangeCode( _
                    ByVal v_strExcCd As String, _
                    ByVal v_strExecConInd As String) As Boolean
' WHITE062
On Error GoTo eh

    IsValidScheduleChangeCode = _
        GetBusExceptionCodeManager.ExceptionIsValidScheduleChangeCode(v_strExcCd, v_strExecConInd)
    Exit Function
  
eh:
    TrsRaiseError mstrCMODULENAME, "IsValidScheduleChangeCode", Err.Number
End Function

Public Function ScheduleChangeExceptionCodeExists( _
            ByVal v_strExcCd As String) As Boolean
            
On Error GoTo eh

    ScheduleChangeExceptionCodeExists = _
        GetBusExceptionCodeManager.ExceptionScheduleChangeExists(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ScheduleChangeExceptionCodeExists", Err.Number
    
End Function

Public Function CaseGroupFor( _
            ByVal v_strExcCd As String) As String
' BLUE082
On Error GoTo eh

    CaseGroupFor = _
        GetExceptionCodeManager.ExceptionCaseGroup(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "CaseGroupFor", Err.Number
    
End Function

Public Function ExceptionDescriptionFor( _
            ByVal v_strExcCd As String) As String
' BLUE079
On Error GoTo eh

    ExceptionDescriptionFor = _
        GetExceptionCodeManager.ExceptionDescription(v_strExcCd)
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "ExceptionDescriptionFor", Err.Number
End Function

Public Function ExceptionAccountNumberFor( _
            ByVal v_strExcCd As String) As String
'BLUE074, BLUE077
On Error GoTo eh
    ExceptionAccountNumberFor = _
        GetExceptionCodeManager.ExceptionAccountNumber(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ExceptionAccountNumberFor", Err.Number
End Function

Public Function IsDOJMWorkgroupNumber(ByVal v_lngWgNum As Long) As Boolean
'    This should be on cboWorkgroup but I have not defined
'    that class yet.
'
'Needswork
On Error GoTo eh
    IsDOJMWorkgroupNumber = _
        v_lngWgNum >= 100 And v_lngWgNum <= 1999
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsDOJMWorkgroupNumber", Err.Number
End Function

Public Function IsCallowayMaintenanceDepartment(ByVal v_strDeptNum) As Boolean
'
'Needswork
On Error GoTo eh
    IsCallowayMaintenanceDepartment = _
        v_strDeptNum = "152" Or v_strDeptNum = "158" Or v_strDeptNum = "170"
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsCallowayMaintenanceDepartment", Err.Number
        
End Function

Private Function GetExceptionCodeManager() As dbCodeTables.cdCodeTables
'Try using CreateInstance here, if doesnt work out move it to cbsTimecard
    Set GetExceptionCodeManager = _
        CreateObject("dbCodeTables.cdCodeTables")

End Function

Private Function GetBusExceptionCodeManager() As busCodeTables.cbsCodeTables
'Try using CreateInstance here, if doesnt work out move it to cbsTimecard
    Set GetBusExceptionCodeManager = _
        CreateObject("busCodeTables.cbsCodeTables")
End Function

Private Function GetUnionManager() As dbCodeTables.cdCodeTables
'Try using CreateInstance here, if doesnt work out move it to cbsTimecard
    Set GetUnionManager = _
        CreateObject("dbCodeTables.cdCodeTables")
End Function

Public Function IsOvertimeReasonCode(ByVal v_strOTRsnCd As String) As Boolean
    Dim l_objReasonMgr As dbCodeTables.cdCodeTables
                      
On Error GoTo eh
    
    Set l_objReasonMgr = _
        CreateObject("dbCodeTables.cdCodeTables")
    IsOvertimeReasonCode = l_objReasonMgr.OvertimeReasonIsValid(v_strOTRsnCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsOvertimeReasonCode"

End Function

' Release 07/25/2001 -- JFB   Added 2 methods to remove
'                               hard coding

Public Function IsAHolidayException( _
            ByVal v_strExcCd As String _
            ) As Boolean
    
    On Error GoTo eh

    If Trim(v_strExcCd) = "" Then
        IsAHolidayException = False
    Else
        IsAHolidayException = GetBusExceptionCodeManager.IsHolidaySymbolException(v_strExcCd)
    End If

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsAHolidayException"
            
End Function

Public Function IsAVacationException( _
            ByVal v_strExcCd As String _
            ) As Boolean
    
    On Error GoTo eh
    
    If Trim(v_strExcCd) = "" Then
        IsAVacationException = False
    Else
        IsAVacationException = GetBusExceptionCodeManager.IsVacationSymbolException(v_strExcCd)
    End If
    
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsAVacationException"
            
End Function


' cuContainer wrappers
Public Function CopyObjects(colOriginal As Collection) As Collection
    Dim l_ContainerLib As utilTRISCommonLib.cuContainer

    Set l_ContainerLib = New cuContainer
    Set CopyObjects = _
        l_ContainerLib.CopyObjects(colOriginal)
    Set l_ContainerLib = Nothing
End Function

Public Function IsIncluded( _
            objIn As Object, _
            col As Collection) As Boolean
    Dim l_ContainerLib As utilTRISCommonLib.cuContainer

    Set l_ContainerLib = New cuContainer
    IsIncluded = _
        l_ContainerLib.IsIncluded(objIn, col)
    Set l_ContainerLib = Nothing
End Function

Public Function CollectionPosition( _
        anObject As Object, _
        aCol As Collection) As Integer
    Dim l_ContainerLib As utilTRISCommonLib.cuContainer

    Set l_ContainerLib = New cuContainer
    CollectionPosition = _
        l_ContainerLib.CollectionPosition(anObject, aCol)
    Set l_ContainerLib = Nothing
End Function

Public Function IsStringInList( _
        ByVal v_strValue As String, _
        ByVal v_strList As Variant) As Boolean
    Dim l_ContainerLib As utilTRISCommonLib.cuContainer

    Set l_ContainerLib = New cuContainer
    IsStringInList = _
        l_ContainerLib.IsStringInList(v_strValue, v_strList)
    Set l_ContainerLib = Nothing
End Function

Public Function IncludesString( _
        ByVal v_strValue As String, _
        ByVal v_colList As Collection) As Boolean
    Dim l_ContainerLib As utilTRISCommonLib.cuContainer

    Set l_ContainerLib = New cuContainer
    IncludesString = _
        l_ContainerLib.IncludesString(v_strValue, v_colList)
    Set l_ContainerLib = Nothing
End Function

Public Function IsWorkingException( _
            ByVal v_strExcCd As String _
            ) As Boolean
' BLUE094, WHITE043
            
On Error GoTo eh

    IsWorkingException = _
        GetBusExceptionCodeManager.ExceptionCodeIsWorking(v_strExcCd)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsWorkingException", Err.Number
        
End Function

Public Function IsOverrideThatAllowsOverlap( _
    ByVal v_strOverExc As String) As Boolean

    On Error GoTo eh

    If CanOverrideCodeOverlap(v_strOverExc) Or IsSpecialOverrideCode(v_strOverExc) Then
        IsOverrideThatAllowsOverlap = True
    Else
        IsOverrideThatAllowsOverlap = False
    End If

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsOverrideThatAllowsOverlap"

End Function

Public Function IsTimeCollectionExceptionCode( _
            ByVal v_strExcCd As String) As Boolean

    Dim l_strTimeType As String
    Dim l_strExcType As String
    
On Error GoTo eh
    l_strTimeType = _
        Trim(GetExceptionCodeManager.ExceptionOtherPayIndicator(v_strExcCd))
    l_strExcType = Trim(GetExceptionCodeManager.ExceptionType(v_strExcCd))
    If l_strTimeType = "T" Or l_strTimeType = "B" Then
        Select Case l_strExcType
            Case "S", "M", "P"
                IsTimeCollectionExceptionCode = True
            Case Else
                IsTimeCollectionExceptionCode = False
        End Select
    Else
        IsTimeCollectionExceptionCode = False
    End If
    
    
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ExceptionAccountNumberFor", Err.Number
End Function



Public Function CallawaySecurityShiftComponentEligible(ByVal v_strOverExcCd As String) As Boolean

    Dim l_obj As dbRules.cdRules
    Dim l_rs As ADODB.Recordset
    Dim l_strOverExcCd As String
    
    On Error GoTo eh
    
    l_strOverExcCd = Trim(v_strOverExcCd)
    If l_strOverExcCd = "" Then
        CallawaySecurityShiftComponentEligible = True
        Exit Function
    End If
    
    Set l_obj = CreateObject("dbRules.cdRules")
    If l_obj.DoesRuleApply(213, trExceptionCode, l_strOverExcCd) Then
        CallawaySecurityShiftComponentEligible = False
    Else
        CallawaySecurityShiftComponentEligible = True
    End If
    
    Set l_obj = Nothing

    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError mstrCMODULENAME, "CallawaySecurityShiftComponentEligible"
End Function

Public Function GetCallawayNightShiftComponent( _
    ByVal v_blnExemptEmployee As Boolean _
    ) As String
       
    Dim l_obj As dbRules.cdRules
    Dim l_rs As ADODB.Recordset
    Dim l_strShiftComp As String
    
    On Error GoTo eh
    
    l_strShiftComp = ""
    Set l_obj = CreateObject("dbRules.cdRules")
    If v_blnExemptEmployee Then
        l_strShiftComp = l_obj.GetRuleLookup(214, 77, "Y")
    Else
        l_strShiftComp = l_obj.GetRuleLookup(214, 77, "N")
    End If
    
    Set l_obj = Nothing

    
    GetCallawayNightShiftComponent = l_strShiftComp
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError mstrCMODULENAME, "GetCallawayNightShiftComponent"
    
End Function

Public Function IsSpecialOverrideCode( _
    ByVal v_strOverExc As String) As Boolean

    Dim l_obj As dbRules.cdRules

    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    IsSpecialOverrideCode = l_obj.DoesRuleApply(280, trExceptionCode, v_strOverExc)


    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsSpecialOverrideCode"

End Function

Public Function CanOverrideCodeOverlap( _
    ByVal v_strOverExc As String) As Boolean

    Dim l_obj As dbRules.cdRules

    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    CanOverrideCodeOverlap = l_obj.DoesRuleApply(281, trExceptionCode, v_strOverExc)


    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "CanOverrideCodeOverlap"

End Function
