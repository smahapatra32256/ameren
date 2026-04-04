VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cbsCodeTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsCodeTables"
Private Const CDEFAULTDATE As String = "01/01/0100"

Public Sub GetAllSupportInformation( _
        ByVal v_aryRequest As Variant, _
        ByRef r_rsAllExceptions As ADODB.Recordset, _
        ByRef r_rsAllComponents As ADODB.Recordset, _
        ByRef r_rsAllHolidays As ADODB.Recordset, _
        ByRef r_rsAllJobs As ADODB.Recordset, _
        ByRef r_rsAllOvertimeReasons As ADODB.Recordset, _
        ByRef r_rsAllLocations As ADODB.Recordset, _
        ByRef r_rsAllUnions As ADODB.Recordset, _
        ByRef r_rsAllDepartments As ADODB.Recordset, _
        ByRef r_rsAllTravelTripLocations As ADODB.Recordset, _
        ByRef r_rsAllTravelTripAllowances As ADODB.Recordset, _
        ByRef r_rsAllActivityList As ADODB.Recordset, _
        ByRef r_rsAllProjectInfo As ADODB.Recordset)

' v_aryRequest should be 1 dim long array with values
' associated with TRCodeTablesToReturn enum.  Return the recordsets
' defined on v_aryRequest.  Order of the request is not important.
' May be zero or 1 based lower bound.
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    If Not IsArray(v_aryRequest) Then
        Exit Sub
    End If

    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    'Set l_CodeTables = New dbCodeTables.cdCodeTables

    If RequestIncludes(v_aryRequest, trCTAllExceptions) Then
        Set r_rsAllExceptions = l_CodeTables.GetAllExceptions
    End If
    If RequestIncludes(v_aryRequest, trCTAllComponents) Then
'        Set r_rsAllComponents = l_CodeTables.GetEffectiveUnionComponents 'replaced per pts388
        Set r_rsAllComponents = l_CodeTables.GetUniqueComponentCodes 'replaced per pts 870 l_CodeTables.GetUnionComponents
    End If
    If RequestIncludes(v_aryRequest, trCTAllHolidays) Then
        Set r_rsAllHolidays = l_CodeTables.GetAllHolidays
    End If
    If RequestIncludes(v_aryRequest, trCTAllJobs) Then
        Set r_rsAllJobs = l_CodeTables.GetAllJobs
    End If
    If RequestIncludes(v_aryRequest, trCTAllOTReasons) Then
        Set r_rsAllOvertimeReasons = l_CodeTables.GetAllOvertimeReasons
    End If
    If RequestIncludes(v_aryRequest, trCTAllLocations) Then
        Set r_rsAllLocations = l_CodeTables.GetAllLocations
    End If
    If RequestIncludes(v_aryRequest, trCTAllUnions) Then
        Set r_rsAllUnions = l_CodeTables.GetAllUnions
    End If
    If RequestIncludes(v_aryRequest, trCTAllDepartments) Then
        Set r_rsAllDepartments = l_CodeTables.GetAllDepartments
    End If
    If RequestIncludes(v_aryRequest, trCTAllTravelTripLocations) Then
        Set r_rsAllTravelTripLocations = l_CodeTables.GetTravelTripLocations
    End If
    If RequestIncludes(v_aryRequest, trCTAllTravelTripAllowances) Then
        Set r_rsAllTravelTripAllowances = l_CodeTables.GetTravelAllowances
    End If
    If RequestIncludes(v_aryRequest, trCTAllActivityList) Then
        Set r_rsAllActivityList = l_CodeTables.GetAllActivityList
    End If
    If RequestIncludes(v_aryRequest, trCTAllProjectInfo) Then
        Set r_rsAllProjectInfo = l_CodeTables.GetProjectInfo
    End If

    Exit Sub

eh:
     Set r_rsAllExceptions = Nothing
     Set r_rsAllComponents = Nothing
     Set r_rsAllHolidays = Nothing
     Set r_rsAllJobs = Nothing
     Set r_rsAllOvertimeReasons = Nothing
     Set r_rsAllLocations = Nothing
     Set r_rsAllUnions = Nothing
     Set r_rsAllDepartments = Nothing
     Set r_rsAllTravelTripLocations = Nothing
     Set r_rsAllTravelTripAllowances = Nothing
     Set r_rsAllActivityList = Nothing

    TrsRaiseError m_strCMODULENAME, "GetAllSupportInformation", , , , , True

End Sub

Private Function RequestIncludes( _
                ByVal v_aryRequest As Variant, _
                ByVal v_lngRequestNumber As TRCodeTablesToReturn) As Boolean
' Already checked for an array.
    Dim i As Integer

On Error GoTo eh
    For i = LBound(v_aryRequest) To UBound(v_aryRequest)
        If v_aryRequest(i) = v_lngRequestNumber Then
            RequestIncludes = True
            Exit Function
        End If
    Next
    Exit Function
                
eh:
    TrsRaiseError m_strCMODULENAME, "RequestIncludes", , True, , , True
                        
End Function

Public Function GetAllJobs() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo GetAllJobs_EH
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    'Set l_CodeTables = New dbCodeTables.cdCodeTables
    Set GetAllJobs = l_CodeTables.GetAllJobs
    Set l_CodeTables = Nothing
    Exit Function

GetAllJobs_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllJobs", , , , , True

End Function


Public Function GetAllLocations() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo GetAllLocations_EH
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetAllLocations = l_CodeTables.GetAllLocations
    Set l_CodeTables = Nothing
    Exit Function

GetAllLocations_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllLocations", , , , , True

End Function

Public Function GetAllUnions() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo GetAllUnions_EH
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetAllUnions = l_CodeTables.GetAllUnions
    Set l_CodeTables = Nothing
    Exit Function

GetAllUnions_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllUnions", , , , , True

End Function

Public Function GetAllOvertimeReasons() As ADODB.Recordset
On Error GoTo GetAllOvertimeReasons_EH
    Dim l_CodeTables As dbCodeTables.cdCodeTables

    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    'Set l_CodeTables = New dbCodeTables.cdCodeTables
    Set GetAllOvertimeReasons = l_CodeTables.GetAllOvertimeReasons

    Set l_CodeTables = Nothing
    Exit Function

GetAllOvertimeReasons_EH:
    Set l_CodeTables = Nothing

    TrsRaiseError m_strCMODULENAME, "GetAllOvertimeReasons", , , , , True

End Function

Public Function GetEffectiveUnionComponents() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo GetEffectiveUnionComponents_EH
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetEffectiveUnionComponents = l_CodeTables.GetEffectiveUnionComponents
    Set l_CodeTables = Nothing
    Exit Function

GetEffectiveUnionComponents_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEffectiveUnionComponents", , , , , True

End Function

Public Function GetUnionComponents() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
    
On Error GoTo GetUnionComponents_EH

    
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    Set GetUnionComponents = l_CodeTables.GetUnionComponents
    Set l_CodeTables = Nothing
    
    Exit Function

GetUnionComponents_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetUnionComponents", , , , , True

End Function

Public Function GetAllExceptions() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo GetAllExceptions_EH
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetAllExceptions = l_CodeTables.GetAllExceptions
    Set l_CodeTables = Nothing
    Exit Function

GetAllExceptions_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllExceptions", , , , , True

End Function

Public Function GetAllHolidays() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo GetAllHolidays_EH
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetAllHolidays = l_CodeTables.GetAllHolidays
    Set l_CodeTables = Nothing
    Exit Function

GetAllHolidays_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllHolidays", , , , , True

End Function

Public Function GetAllEarningCodes() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo GetAllEarningCodes_EH
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetAllEarningCodes = l_CodeTables.GetAllEarningCodes
    Set l_CodeTables = Nothing
    Exit Function

GetAllEarningCodes_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllEarningCodes", , , , , True

End Function


Public Function GetTravelAllowances() As ADODB.Recordset

    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
    
    On Error GoTo GetTravelAllowances_EH
    
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    Set GetTravelAllowances = l_CodeTables.GetTravelAllowances
    Set l_CodeTables = Nothing
    Exit Function

GetTravelAllowances_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTravelAllowances", , , , , True

End Function

Public Function GetTravelTripLocations() As ADODB.Recordset
' REED104
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
    
    On Error GoTo GetTravelTripLocations_EH
    
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    Set GetTravelTripLocations = l_CodeTables.GetTravelTripLocations
    Set l_CodeTables = Nothing
    Exit Function

GetTravelTripLocations_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTravelTripLocations", , , , , True

End Function

'added for multi-threading
Public Function GetUniqueComponentCodes() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo eh
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetUniqueComponentCodes = l_CodeTables.GetUniqueComponentCodes
    Set l_CodeTables = Nothing
    Exit Function

eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetUniqueComponentCodes", , , , , True

End Function

Public Function GetAllDepartments() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo eh
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetAllDepartments = l_CodeTables.GetAllDepartments
    Set l_CodeTables = Nothing
    Exit Function

eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllDepartments", , , , , True

End Function

'*********************************************************
' Begin Union component services
'*********************************************************
Public Function IsVacationReliefCode( _
        ByVal v_strUnionCd As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCd As String) As Boolean
'  Is this a valid vacation relief code for union and date
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    IsVacationReliefCode = _
        Not IsNull(l_CodeTables.UnionComponentVacationReliefAmount( _
                        v_strUnionCd, _
                        v_datEffDate, _
                        v_strCompCd))
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "IsVacationReliefCode", , , , , True
End Function

Public Function IsShiftCompCode( _
        ByVal v_strUnionCd As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCd As String) As Boolean
'  Is this a valid shift component for union and date
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    IsShiftCompCode = _
        Not IsNull(l_CodeTables.UnionComponentShiftAmount( _
                        v_strUnionCd, _
                        v_datEffDate, _
                        v_strCompCd))
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "IsShiftCompCode", , , , , True
End Function

Public Function IsComponentCode( _
        ByVal v_strUnionCd As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCd As String) As Boolean
'  Is this a valid component code for union and date
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    IsComponentCode = _
        Not IsNull(l_CodeTables.UnionComponentAmount( _
                        v_strUnionCd, _
                        v_datEffDate, _
                        v_strCompCd))
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "IsComponentCode", , , , , True
End Function
'*********************************************************
' End Union component services
'*********************************************************

'*********************************************************
' Begin Exception code services
'*********************************************************
Public Function ExceptionCodeIsValid( _
            ByVal v_strExcCd As String) As Boolean
' BLUE094, WHITE043
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    ExceptionCodeIsValid = _
        l_CodeTables.ExceptionCodeIsValid(v_strExcCd)
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionCodeIsValid", , , , , True
End Function

Public Function ExceptionIsCaseRelated( _
            ByVal v_strExcCd As String) As Boolean
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    ExceptionIsCaseRelated = _
        Not (l_CodeTables.ExceptionCaseGroup(v_strExcCd) = "")
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionIsCaseRelated", , , , , True
End Function

Public Function ExceptionIsValidInsideOfSchedule( _
            ByVal v_strExcCd As String) As Boolean
' GOLD003, BLUE090
On Error GoTo eh
    ExceptionIsValidInsideOfSchedule = _
        Not ExceptionIsOvertimeRelated(v_strExcCd) And Trim(v_strExcCd) <> "OST"
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionIsValidInsideOfSchedule", , , , , True
End Function

Public Function ExceptionIsValidOutsideOfSchedule( _
            ByVal v_strExcCd As String) As Boolean
' GOLD002, BLUE090
' 8/3/01 RSG Added "LP"
' 10/09/09 Removed hardcoding to add "NAH"
On Error GoTo eh

    If DoesRuleApply(188, trExceptionCode, v_strExcCd) Then
       ExceptionIsValidOutsideOfSchedule = True
    Else
       ExceptionIsValidOutsideOfSchedule = ExceptionIsOvertimeRelated(v_strExcCd)
    End If
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionIsValidOutsideOfSchedule", , , , , True
End Function

Public Function ExceptionIsOvertimeRelated( _
            ByVal v_strExcCd As String) As Boolean
' WHITE056, WHITE070
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    If l_CodeTables.ExceptionBankCode(v_strExcCd) = "O" Then
        ExceptionIsOvertimeRelated = True
    Else
        ExceptionIsOvertimeRelated = False
    End If
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionIsOvertimeRelated", , , , , True
End Function

Public Function ExceptionIsOvertimeBankRelated( _
            ByVal v_strExcCd As String) As Boolean
' BLUE075, BLUE076, WHITE056
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    ExceptionIsOvertimeBankRelated = _
        l_CodeTables.ExceptionBankCode(v_strExcCd) = "O"
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionIsOvertimeBankRelated", , , , , True
End Function

Public Function ExceptionIsUnpaid( _
            ByVal v_strExcCd As String, _
            ByVal v_strExecConInd As String) As Boolean
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    ExceptionIsUnpaid = _
        l_CodeTables.ExceptionEarningsCode( _
                        v_strExcCd, v_strExecConInd) = "99"
    Set l_CodeTables = Nothing
    Exit Function
            
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionIsUnpaid", , , , , True
End Function

Public Function ExceptionIsOtherPay( _
            ByVal v_strExcCd As String) As Boolean
' BLUE094
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    Dim l_strExcType As String
    Dim l_strOtherPayInd As String
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    l_strExcType = l_CodeTables.ExceptionType(v_strExcCd)
    l_strOtherPayInd = l_CodeTables.ExceptionOtherPayIndicator(v_strExcCd)
    ExceptionIsOtherPay = _
        (l_strExcType = "S" Or l_strExcType = "P" _
            Or l_strExcType = "M") And _
        (l_strOtherPayInd = "O" Or l_strOtherPayInd = "B")
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionIsOtherPay", , , , , True
End Function

Public Function ExceptionIsECNonProductive( _
            ByVal v_strExcCd As String) As Boolean
' BLUE094
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    Dim l_strExcType As String
    Dim l_strOtherPayInd As String
    Dim l_strSchedChgInd As String
    Dim l_strAcctNum As String
            
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    l_strExcType = l_CodeTables.ExceptionType(v_strExcCd)
    l_strOtherPayInd = l_CodeTables.ExceptionOtherPayIndicator(v_strExcCd)
    l_strSchedChgInd = l_CodeTables.ExceptionScheduleChgInd(v_strExcCd)
    l_strAcctNum = l_CodeTables.ExceptionAccountNumber(v_strExcCd)
    ExceptionIsECNonProductive = _
        (l_strExcType = "S" Or l_strExcType = "P" _
            Or l_strExcType = "M") And _
        (l_strOtherPayInd = "T" Or l_strOtherPayInd = "B") And _
        (l_strSchedChgInd <> "S") And _
        (l_strAcctNum <> "")
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionIsECNonProductive", , , , , True
End Function

Public Function ExceptionIsOverride( _
            ByVal v_strExcCd As String) As Boolean
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    Dim l_strExcType As String
            
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    l_strExcType = l_CodeTables.ExceptionType(v_strExcCd)
    If l_strExcType = "P" Or l_strExcType = "M" Then
        If l_CodeTables.ExceptionExecutiveEarningsCode(v_strExcCd) = "" And _
            l_CodeTables.ExceptionContractEarningsCode(v_strExcCd) = "" Then
            ExceptionIsOverride = True
        End If
    End If
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionIsOverride", , , , , True
End Function

Public Function ExceptionIsEligibleForScheduleChangePayment( _
            ByVal v_strExcCd As String) As Boolean
    Dim l_CodeTables As dbCodeTables.cdCodeTables
            
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    ExceptionIsEligibleForScheduleChangePayment = _
        (l_CodeTables.ExceptionScheduleChgInd(v_strExcCd) = "Y")
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionIsEligibleForScheduleChangePayment", , , , , True
End Function

Public Function ExceptionIsEligibleForFamilyLeave( _
            ByVal v_strExcCd As String) As Boolean
'WHITE066
    Dim l_CodeTables As dbCodeTables.cdCodeTables
            
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    ExceptionIsEligibleForFamilyLeave = _
        (l_CodeTables.ExceptionFamilyLeaveInd(v_strExcCd) = "Y")
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
   Set l_CodeTables = Nothing
   TrsRaiseError m_strCMODULENAME, "ExceptionIsEligibleForFamilyLeave", , , , , True
End Function

Public Function ExceptionIsValidScheduleChangeCode( _
                    ByVal v_strExcCd As String, _
                    ByVal v_strExecConInd As String) As Boolean
' WHITE062
    Dim l_CodeTables As dbCodeTables.cdCodeTables
            
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    If l_CodeTables.ExceptionScheduleChgInd(v_strExcCd) <> "S" Then
        Exit Function
    End If
    If Trim(v_strExecConInd) = "E" Then
        ExceptionIsValidScheduleChangeCode = _
            l_CodeTables.ExceptionExecutiveEarningsCode(v_strExcCd) <> ""
    ElseIf Trim(v_strExecConInd) = "C" Then
        ExceptionIsValidScheduleChangeCode = _
            l_CodeTables.ExceptionContractEarningsCode(v_strExcCd) <> ""
    Else
        ExceptionIsValidScheduleChangeCode = False
    End If
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionIsValidScheduleChangeCode", , , , , True
End Function

Public Function ExceptionScheduleChangeExists( _
            ByVal v_strExcCd As String) As Boolean
    Dim l_CodeTables As dbCodeTables.cdCodeTables
            
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    ExceptionScheduleChangeExists = _
        l_CodeTables.ExceptionScheduleChgInd(v_strExcCd) = "S"
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionScheduleChangeExists", , , , , True
End Function

Public Function ExceptionCodeIsWorking( _
            ByVal v_strExcCd As String) As Boolean
' BLUE094, WHITE043
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    ExceptionCodeIsWorking = _
        l_CodeTables.ExceptionWorkRestInd(v_strExcCd) = "W"
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionCodeIsWorking", , , , , True
End Function

'*********************************************************
' End Exception code services
'*********************************************************

Public Function IsVacationSymbolException( _
            ByVal v_strExcCd As String) As Boolean
' There are 3 vacation execptions treated the same
' They are 'V', 'VH', 'VBY'  each one has the same rules
' Instead of hardcoding each the symbol code (symbol_cd) of 'V'
'  indicates all follow the vacation rules

    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    IsVacationSymbolException = _
        l_CodeTables.ExceptionSymbolCd(v_strExcCd) = "V"
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "IsVacationSymbolException", , , , , True
End Function

Public Function IsHolidaySymbolException( _
            ByVal v_strExcCd As String) As Boolean
' There are 8 Holiday exceptions treated the same
' They are 'H','HR','HD','HF','HRV','HA', 'HNT', 'HE' each one has the same rules
' Instead of hardcoding each the symbol code (symbol_cd) of 'V'
'  indicates all follow the vacation rules

    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    IsHolidaySymbolException = _
        l_CodeTables.ExceptionSymbolCd(v_strExcCd) = "H"
    Set l_CodeTables = Nothing
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "IsHolidaySymbolException", , , , , True
End Function

Public Function GetAvailableScheduleDetails() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo GetAvailableScheduleDetails_EH
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetAvailableScheduleDetails = l_CodeTables.GetAvailableScheduleDetails
    Set l_CodeTables = Nothing
    Exit Function

GetAvailableScheduleDetails_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAvailableScheduleDetails", , , , , True

End Function

Public Function GetAvailableScheduleHeader() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo GetAvailableScheduleHeader_EH
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetAvailableScheduleHeader = l_CodeTables.GetAvailableScheduleHeader
    Set l_CodeTables = Nothing
    Exit Function

GetAvailableScheduleHeader_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAvailableScheduleHeader", , , , , True

End Function


Public Function GetAllActivityList() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo GetAllActivityList_EH
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetAllActivityList = l_CodeTables.GetAllActivityList
    Set l_CodeTables = Nothing
    Exit Function

GetAllActivityList_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllActivityList", , , , , True

End Function
Public Function ValidateScheduleCode(ByVal v_lScheduleCode As Long) As Boolean
    'NOTE:  If the calling routine has already retrieved the list of avail schedules (GetAllAvailSchedDetails)
    '  the schedule code may be validated against the GetAllAvailSchedDetails recordset
    
    Dim l_objScheduleDetail As dbCodeTables.cdCodeTables
    
    Set l_objScheduleDetail = CreateObject("dbCodeTables.cdCodeTables")
    
    ValidateScheduleCode = l_objScheduleDetail.ValidateScheduleCode(v_lScheduleCode)
    
    Set l_objScheduleDetail = Nothing
    
    Exit Function
    
ValidateScheduleCode_EH:
    If Not l_objScheduleDetail Is Nothing Then
        Set l_objScheduleDetail = Nothing
    End If
    TrsRaiseError m_strCMODULENAME, "ValidateScheduleCode", , , , , True

End Function
Public Function JobNumberIsValid( _
            ByVal v_strJobNum As String, _
            Optional ByVal v_datEffDate As Date = CDEFAULTDATE) As Boolean

  
    Dim l_objJobNumber As dbCodeTables.cdCodeTables
    
    Set l_objJobNumber = CreateObject("dbCodeTables.cdCodeTables")
    
    JobNumberIsValid = l_objJobNumber.JobNumberIsValid(v_strJobNum, v_datEffDate)
    
    Set l_objJobNumber = Nothing
    
    Exit Function
    
JobNumberIsValid_EH:
    If Not l_objJobNumber Is Nothing Then
        Set l_objJobNumber = Nothing
    End If
    TrsRaiseError m_strCMODULENAME, "JobNumberIsValid", , , , , True

End Function


Public Function GetProjectInfo _
                (Optional ByVal v_intProjectId As Integer, _
                 Optional ByVal v_intPPMInd As String _
                ) As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                 'MTS standard
    
    On Error GoTo GetProjectInfo_EH
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetProjectInfo = l_CodeTables.GetProjectInfo(v_intProjectId, v_intPPMInd)
    Set l_CodeTables = Nothing
    Exit Function

GetProjectInfo_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetProjectInfo", , , , , True

End Function

Public Function GetAllProjects() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
    Dim l_dat1MonthAgo As Date
                                                                 
    
    On Error GoTo GetAllProjects_EH
    l_dat1MonthAgo = DateAdd("m", -1, Date)
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetAllProjects = l_CodeTables.GetAllProjects(l_dat1MonthAgo)
    Set l_CodeTables = Nothing
    Exit Function

GetAllProjects_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllProjects", , , , , True

End Function


Public Function GetLocationsAssociatedToDepartment _
                (ByVal v_varDepts As Variant _
                ) As ADODB.Recordset
                
'v_varDepts is an array of departments stored as strings
'It could be just a string
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables 'MTS standard
    Dim l_colDepts As Collection
                     
    On Error GoTo GetLocationsAssociatedToDepartment_EH
    
    Set l_colDepts = New Collection
    If IsArray(v_varDepts) Then
        Set l_colDepts = CollectionFromArray(v_varDepts, 0)
    Else
        l_colDepts.Add v_varDepts
    End If
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    Set GetLocationsAssociatedToDepartment = l_CodeTables.GetLocationsAssociatedToDepartment(l_colDepts)
    Set l_CodeTables = Nothing
    Exit Function

GetLocationsAssociatedToDepartment_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetLocationsAssociatedToDepartment", , , , , True

End Function

Public Function GetAllScheduleCodes() As ADODB.Recordset
    Dim l_objCodeTables As dbCodeTables.cdCodeTables
    
    On Error GoTo GetAllScheduleCodes_eh

    Set l_objCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetAllScheduleCodes = l_objCodeTables.GetAllScheduleCodes
    
    Set l_objCodeTables = Nothing
    
    Exit Function

GetAllScheduleCodes_eh:
    Set l_objCodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllScheduleCodes", , , , , True

End Function


Public Function GetRotationSchedHeader _
                (Optional ByVal v_intRotationCode As Integer _
                ) As ADODB.Recordset

On Error GoTo GetRotationSchedHeader_EH
    
    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation

    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")

    Set GetRotationSchedHeader = l_dbSchedRotation.GetRotationSchedHeader(v_intRotationCode)

    Set l_dbSchedRotation = Nothing
    
    Exit Function

GetRotationSchedHeader_EH:
    Set l_dbSchedRotation = Nothing
    TrsRaiseError m_strCMODULENAME, "GetRotationSchedHeader"

End Function
'public call to get rotation detail
Public Function GetRotationSchedDetail _
                (Optional ByVal v_intRotationCode As Integer, _
                 Optional ByVal v_intRotationWeek As Integer _
                ) As ADODB.Recordset

On Error GoTo GetRotationSchedDetail_EH

    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation

    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")

    Set GetRotationSchedDetail = l_dbSchedRotation.GetRotationSchedDetail(v_intRotationCode, v_intRotationWeek)

    Set l_dbSchedRotation = Nothing
    
Exit Function

GetRotationSchedDetail_EH:
    Set l_dbSchedRotation = Nothing
    TrsRaiseError m_strCMODULENAME, "GetRotationSchedDetail"

End Function
'Added 5/09/2002 per Prem's request
Public Function GetExceptionGroupForColoring() As ADODB.Recordset
    
    On Error GoTo GetExceptionGroupForColoring_EH
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables                          'MTS standard
                                                                     'MTS standard
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")    'MTS standard
    'Set l_CodeTables = New dbCodeTables.cdCodeTables
    Set GetExceptionGroupForColoring = l_CodeTables.GetExceptionGroupForColoring
    Set l_CodeTables = Nothing
    Exit Function

GetExceptionGroupForColoring_EH:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "GetExceptionGroupForColoring", , , , , True

End Function

Public Function ExceptionCodeIsValidFor( _
            ByVal v_strExcCd As String, _
            ByVal v_strExecConInd As String, _
            Optional ByVal v_strBenefitsIND As String = "" _
            ) As Boolean
            
'This method is called from busTimeCard and busSchedException
'JimB 04/21/2003
'This method was suffixed with a 2 and had a duplicate method without the suffix.
'The only difference in the 2 methods was a check to see if it was valid
' on benefits Indicator
'busCascade calls the dbCodeTables directly

    Dim l_CodeTables As dbCodeTables.cdCodeTables
    Dim l_blnIsValid As Boolean

On Error GoTo eh
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    l_blnIsValid = _
        l_CodeTables.ExceptionCodeIsValidFor(v_strExcCd, v_strExecConInd)
    
    If l_blnIsValid Then
'If the exception can only be entered as other pay, it is invalid
        If l_CodeTables.ExceptionOtherPayIndicator(v_strExcCd) = "O" Then
            l_blnIsValid = False
        End If
    End If
    
    If l_blnIsValid Then
    'If Benefits Indicator was not passed do not do the last validation, it is valid
        If Trim(v_strBenefitsIND) = "" Then
            ExceptionCodeIsValidFor = True
            Exit Function
        End If
    End If
    
    If l_blnIsValid Then
        If v_strBenefitsIND = "CIPS" And _
            v_strExecConInd = "C" Then
            l_blnIsValid = Not l_CodeTables.IsExceptionInGroup(3, v_strExcCd, "")
        Else
            l_blnIsValid = Not l_CodeTables.IsExceptionInGroup(4, v_strExcCd, "")
        End If
    End If
    
    Set l_CodeTables = Nothing
    ExceptionCodeIsValidFor = l_blnIsValid
    Exit Function
    
eh:
    Set l_CodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionCodeIsValidFor"
End Function
'PREM BEGIN TRIS CONVERSION PHASE 3 (PHASE 1 .NET)
'The following functions retrieve Crafts, Qualifications
', Special Skills, and OT Responses information for Code
'Table Maintenance
Public Function GetAllCrafts() As ADODB.Recordset
On Error GoTo GetAllCrafts_EH
'Declare
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables

    'Construct
    Set lobjdbCodeTables = CreateObject( _
                            "dbCodeTables.cdCodeTables")
    'Get and set for return
    Set GetAllCrafts = lobjdbCodeTables.GetAllCrafts
    'Destroy
    Set lobjdbCodeTables = Nothing

    Exit Function
GetAllCrafts_EH:
    Set lobjdbCodeTables = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetAllCrafts" _
        , , , , _
        , True

End Function
Public Function GetAllQualifications() As ADODB.Recordset
On Error GoTo GetAllQualifications_EH
'Declare
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables

    'Construct
    Set lobjdbCodeTables = CreateObject( _
                            "dbCodeTables.cdCodeTables")
    'Get and set for return
    Set GetAllQualifications = lobjdbCodeTables.GetAllQualifications
    'Destroy
    Set lobjdbCodeTables = Nothing

    Exit Function
GetAllQualifications_EH:
    Set lobjdbCodeTables = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetAllQualifications" _
        , , , , _
        , True

End Function
Public Function GetAllSpecialSkills() As ADODB.Recordset
On Error GoTo GetAllSpecialSkills_EH
'Declare
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables

    'Construct
    Set lobjdbCodeTables = CreateObject( _
                            "dbCodeTables.cdCodeTables")
    'Get and set for return
    Set GetAllSpecialSkills = lobjdbCodeTables.GetAllSpecialSkills
    'Destroy
    Set lobjdbCodeTables = Nothing

    Exit Function
GetAllSpecialSkills_EH:
    Set lobjdbCodeTables = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetAllSpecialSkills" _
        , , , , _
        , True

End Function
Public Function GetAllOvertimeResponses() As ADODB.Recordset
On Error GoTo GetAllOvertimeResponses_EH
'Declare
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables
Dim ladrs As ADODB.Recordset, lsFilterString As String
Dim llRecordIndex As Long

    'Construct
    Set lobjdbCodeTables = CreateObject( _
                            "dbCodeTables.cdCodeTables")
    'Get and set for return
    Set GetAllOvertimeResponses = lobjdbCodeTables.GetAllOvertimeResponses
    'Destroy
    Set lobjdbCodeTables = Nothing

    Exit Function
GetAllOvertimeResponses_EH:
    Set lobjdbCodeTables = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetAllOvertimeResponses" _
        , , , , _
        , True

End Function
Public Function GetAllPromotionalLines() As ADODB.Recordset
On Error GoTo GetAllPromotionalLines_EH

'Declare
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables

    'Construct
    Set lobjdbCodeTables = CreateObject( _
                            "dbCodeTables.cdCodeTables")
    'Get and set for return
    Set GetAllPromotionalLines = lobjdbCodeTables.GetAllPromotionalLines
    'Destroy
    Set lobjdbCodeTables = Nothing

    Exit Function
GetAllPromotionalLines_EH:
    Set lobjdbCodeTables = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetAllPromotionalLines" _
        , , , , _
        , True

End Function
Public Function GetAllCompanies() As ADODB.Recordset
On Error GoTo GetAllCompanies_EH
'Declare
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables

    'Construct
    Set lobjdbCodeTables = CreateObject( _
                            "dbCodeTables.cdCodeTables")
    'Get and set for return
    Set GetAllCompanies = lobjdbCodeTables.GetAllCompanies
    'Destroy
    Set lobjdbCodeTables = Nothing

    Exit Function
GetAllCompanies_EH:
    Set lobjdbCodeTables = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetAllCompanies" _
        , , , , _
        , True

End Function
Public Function GetAllOvertimeBankHeaders() As ADODB.Recordset
On Error GoTo GetAllOvertimeBankHeaders_EH
'Declare
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables

    'Construct
    Set lobjdbCodeTables = CreateObject( _
                            "dbCodeTables.cdCodeTables")
    'Get and set for return
    Set GetAllOvertimeBankHeaders = lobjdbCodeTables.GetAllOvertimeBankHeaders
    'Destroy
    Set lobjdbCodeTables = Nothing

    Exit Function
GetAllOvertimeBankHeaders_EH:
    Set lobjdbCodeTables = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetAllOvertimeBankHeaders" _
        , , , , _
        , True

End Function
Public Function GetAllBudgetNumbers() As ADODB.Recordset
On Error GoTo GetAllBudgetNumbers_EH
'Declare
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables

    'Construct
    Set lobjdbCodeTables = CreateObject( _
                            "dbCodeTables.cdCodeTables")
    'Get and set for return
    Set GetAllBudgetNumbers = lobjdbCodeTables.GetAllBudgetNumbers
    'Destroy
    Set lobjdbCodeTables = Nothing

    Exit Function
GetAllBudgetNumbers_EH:
    Set lobjdbCodeTables = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetAllBudgetNumbers" _
        , , , , _
        , True

End Function
Public Function GetSeniorityReportSortOrder() As ADODB.Recordset
On Error GoTo GetSeniorityReportSortOrder_EH
Dim ladrs As ADODB.Recordset

    Set GetSeniorityReportSortOrder = Nothing

    Set ladrs = DefineSortOrderRS
    With ladrs
        .AddNew
        !SO_NUM = 1
        !SO_DESC = "Department"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 2
        !SO_DESC = "Employee Number"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 3
        !SO_DESC = "Job Class"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 4
        !SO_DESC = "Location"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 5
        !SO_DESC = "Workgroup"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 6
        !SO_DESC = "Union Code"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 7
        !SO_DESC = "Corporate Seniority"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 8
        !SO_DESC = "Company Seniority"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 9
        !SO_DESC = "Union Seniority"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 10
        !SO_DESC = "Plant Seniority"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 11
        !SO_DESC = "Department/Shop Seniority"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 12
        !SO_DESC = "Job Seniority"
        !SO_SYS_DESC = ""
        .Update
        .AddNew
        !SO_NUM = 13
        !SO_DESC = "Promotional Seniority"
        !SO_SYS_DESC = ""
        .Update
    End With

    Set GetSeniorityReportSortOrder = ladrs

    Exit Function
GetSeniorityReportSortOrder_EH:
    TrsRaiseError m_strCMODULENAME _
        , "GetSeniorityReportSortOrder" _
        , , , , _
        , True

End Function
Public Function GetPreOppReportSortOrder() As ADODB.Recordset
On Error GoTo GetPreOppReportSortOrder_EH
Dim ladrs As ADODB.Recordset

    Set GetPreOppReportSortOrder = Nothing

    Set ladrs = DefineSortOrderRS
    With ladrs
        .AddNew
        !SO_NUM = 1
        !SO_DESC = "Total OT Hours Low"
        !SO_SYS_DESC = "BY LOWEST TOT_OT_HRS"
        .Update
        .AddNew
        !SO_NUM = 2
        !SO_DESC = "Total OT Hours High"
        !SO_SYS_DESC = "BY HIGHEST TOT_OT_HRS"
        .Update
        .AddNew
        !SO_NUM = 3
        !SO_DESC = "Promotional/Plant Seniority Low"
        !SO_SYS_DESC = "BY LOWEST LOC_SENIORITY"
        .Update
        .AddNew
        !SO_NUM = 4
        !SO_DESC = "Promotional/Plant Seniority High"
        !SO_SYS_DESC = "BY HIGHEST LOC_SENIORITY"
        .Update
        .AddNew
        !SO_NUM = 5
        !SO_DESC = "Union Seniority Low"
        !SO_SYS_DESC = "BY LOWEST UNION_SENIORITY"
        .Update
        .AddNew
        !SO_NUM = 6
        !SO_DESC = "Union Seniority High"
        !SO_SYS_DESC = "BY HIGHEST UNION_SENIORITY"
        .Update
        .AddNew
        !SO_NUM = 7
        !SO_DESC = "UE Seniority Low"
        !SO_SYS_DESC = "BY LOWEST SYSTEM_SENIORITY"
        .Update
        .AddNew
        !SO_NUM = 8
        !SO_DESC = "UE Seniority High"
        !SO_SYS_DESC = "BY HIGHEST SYSTEM_SENIORITY"
        .Update
        .AddNew
        !SO_NUM = 9
        !SO_DESC = "Department Seniority Low"
        !SO_SYS_DESC = "BY HIGHEST DEPT_EFF_DATE"
        .Update
        .AddNew
        !SO_NUM = 10
        !SO_DESC = "Department Seniority High"
        !SO_SYS_DESC = "BY LOWEST DEPT_EFF_DATE"
        .Update
        .AddNew
        !SO_NUM = 11
        !SO_DESC = "Employee Number Low"
        !SO_SYS_DESC = "BY LOWEST EMP_NUM"
        .Update
        .AddNew
        !SO_NUM = 12
        !SO_DESC = "Employee Number High"
        !SO_SYS_DESC = "BY HIGHEST EMP_NUM"
        .Update
        .AddNew
        !SO_NUM = 13
        !SO_DESC = "Job Seniority Low"
        !SO_SYS_DESC = "BY HIGHEST JOB_EFF_DATE"
        .Update
        .AddNew
        !SO_NUM = 14
        !SO_DESC = "Job Seniority High"
        !SO_SYS_DESC = "BY LOWEST JOB_EFF_DATE"
        .Update
        .AddNew
        !SO_NUM = 15
        !SO_DESC = "Callaway Seniority Low"
        !SO_SYS_DESC = "BY LOWEST CALLAWAY_PLT_SNRTY"
        .Update
        .AddNew
        !SO_NUM = 16
        !SO_DESC = "Callaway Seniority High"
        !SO_SYS_DESC = "BY HIGHEST CALLAWAY_PLT_SNRTY"
        .Update
        .AddNew
        !SO_NUM = 17
        !SO_DESC = "Age Low"
        !SO_SYS_DESC = "BY HIGHEST DATE_OF_BIRTH"
        .Update
        .AddNew
        !SO_NUM = 18
        !SO_DESC = "Age High"
        !SO_SYS_DESC = "BY LOWEST DATE_OF_BIRTH"
        .Update
        .AddNew
        !SO_NUM = 19
        !SO_DESC = "Job Classification Low"
        !SO_SYS_DESC = "BY LOWEST JOB_NUM"
        .Update
        .AddNew
        !SO_NUM = 20
        !SO_DESC = "Job Classification High"
        !SO_SYS_DESC = "BY HIGHEST JOB_NUM"
        .Update
        .AddNew
        !SO_NUM = 21
        !SO_DESC = "Workgroup Number Low"
        !SO_SYS_DESC = "BY LOWEST WG_NUM"
        .Update
        .AddNew
        !SO_NUM = 22
        !SO_DESC = "Workgroup Number High"
        !SO_SYS_DESC = "BY HIGHEST WG_NUM"
        .Update
    End With

    Set GetPreOppReportSortOrder = ladrs

    Exit Function
GetPreOppReportSortOrder_EH:
    TrsRaiseError m_strCMODULENAME _
        , "GetPreOppReportSortOrder" _
        , , , , _
        , True

End Function
Public Function GetOTEqualReportSortOrder() As ADODB.Recordset
On Error GoTo GetOTEqualReportSortOrder_EH
Dim ladrs As ADODB.Recordset

    Set GetOTEqualReportSortOrder = Nothing

    Set ladrs = DefineSortOrderRS
    With ladrs
        .AddNew
        !SO_NUM = 1
        !SO_DESC = "Total OT Hours Low"
        !SO_SYS_DESC = "BY LOWEST TOT_OT_HRS"
        .Update
        .AddNew
        !SO_NUM = 2
        !SO_DESC = "Total OT Hours High"
        !SO_SYS_DESC = "BY HIGHEST TOT_OT_HRS"
        .Update
        .AddNew
        !SO_NUM = 3
        !SO_DESC = "Plant Seniority Low"
        !SO_SYS_DESC = "BY LOWEST PLANT_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 4
        !SO_DESC = "Plant Seniority High"
        !SO_SYS_DESC = "BY HIGHEST PLANT_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 5
        !SO_DESC = "Union Seniority Low"
        !SO_SYS_DESC = "BY LOWEST UNION_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 6
        !SO_DESC = "Union Seniority High"
        !SO_SYS_DESC = "BY HIGHEST UNION_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 7
        !SO_DESC = "Corporate Seniority Low"
        !SO_SYS_DESC = "BY LOWEST CORP_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 8
        !SO_DESC = "Corporate Seniority High"
        !SO_SYS_DESC = "BY HIGHEST CORP_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 9
        !SO_DESC = "Department Seniority Low"
        !SO_SYS_DESC = "BY HIGHEST DEPT_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 10
        !SO_DESC = "Department Seniority High"
        !SO_SYS_DESC = "BY LOWEST DEPT_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 11
        !SO_DESC = "Employee Number Low"
        !SO_SYS_DESC = "BY LOWEST EMP_NUM"
        .Update
        .AddNew
        !SO_NUM = 12
        !SO_DESC = "Employee Number High"
        !SO_SYS_DESC = "BY HIGHEST EMP_NUM"
        .Update
        .AddNew
        !SO_NUM = 13
        !SO_DESC = "Job Seniority Low"
        !SO_SYS_DESC = "BY HIGHEST JOB_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 14
        !SO_DESC = "Job Seniority High"
        !SO_SYS_DESC = "BY LOWEST JOB_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 15
        !SO_DESC = "Line of Promotion Seniority Low"
        !SO_SYS_DESC = "BY LOWEST LOP_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 16
        !SO_DESC = "Line of Promotion Seniority High"
        !SO_SYS_DESC = "BY HIGHEST LOP_SNRTY_QTY"
        .Update
        .AddNew
        !SO_NUM = 17
        !SO_DESC = "Job Classification"
        !SO_SYS_DESC = "BY LOWEST JOB_NUM"
        .Update
    End With

    Set GetOTEqualReportSortOrder = ladrs

    Exit Function
GetOTEqualReportSortOrder_EH:
    TrsRaiseError m_strCMODULENAME _
        , "GetOTEqualReportSortOrder" _
        , , , , _
        , True

End Function
Public Function GetVacSumReportSortOrder() As ADODB.Recordset
On Error GoTo GetVacSumReportSortOrder_EH
Dim ladrs As ADODB.Recordset

    Set GetVacSumReportSortOrder = Nothing

    Set ladrs = DefineSortOrderRS
    With ladrs
        .AddNew
        !SO_NUM = 1
        !SO_DESC = "Employee Number"
        !SO_SYS_DESC = "BY EMP_NUM NOPRINT"
        .Update
        .AddNew
        !SO_NUM = 2
        !SO_DESC = "Employee Name"
        !SO_SYS_DESC = "BY EMP_NAME NOPRINT"
        .Update
        .AddNew
        !SO_NUM = 3
        !SO_DESC = "Workgroup"
        !SO_SYS_DESC = "BY WG_NUM NOPRINT"
        .Update
        .AddNew
        !SO_NUM = 4
        !SO_DESC = "Department"
        !SO_SYS_DESC = "BY DEPT_NUM NOPRINT"
        .Update
        .AddNew
        !SO_NUM = 5
        !SO_DESC = "Job Classification"
        !SO_SYS_DESC = "BY JOB_NUM NOPRINT"
        .Update
        .AddNew
        !SO_NUM = 6
        !SO_DESC = "Page Break"
        !SO_SYS_DESC = "PAGE-BREAK"
        .Update
        .AddNew
        !SO_NUM = 7
        !SO_DESC = "Corporate Seniority"
        !SO_SYS_DESC = "BY HIGHEST CS NOPRINT"
        .Update
        .AddNew
        !SO_NUM = 8
        !SO_DESC = "Union Seniority"
        !SO_SYS_DESC = "BY HIGHEST US NOPRINT"
        .Update
        .AddNew
        !SO_NUM = 9
        !SO_DESC = "Plant Seniority"
        !SO_SYS_DESC = "BY HIGHEST LS NOPRINT"
        .Update
        .AddNew
        !SO_NUM = 10
        !SO_DESC = "LOP Seniority"
        !SO_SYS_DESC = "BY HIGHEST PS NOPRINT"
        .Update
    End With

    Set GetVacSumReportSortOrder = ladrs

    Exit Function
GetVacSumReportSortOrder_EH:
    TrsRaiseError m_strCMODULENAME _
        , "GetVacSumReportSortOrder" _
        , , , , _
        , True

End Function
Private Function DefineSortOrderRS() As ADODB.Recordset
On Error GoTo DefineSortOrderRS_EH
Dim ladrs As ADODB.Recordset

    Set ladrs = New ADODB.Recordset
    With ladrs
        .Fields.Append "SO_NUM", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "SO_DESC", adVarChar, 60, adFldIsNullable + adFldUpdatable
        .Fields.Append "SO_SYS_DESC", adVarChar, 120, adFldIsNullable + adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic

        .Open
    End With

    Set DefineSortOrderRS = ladrs
    Exit Function
DefineSortOrderRS_EH:
    TrsRaiseError m_strCMODULENAME _
        , "DefineSortOrderRS"

End Function
'PREM END TRIS CONVERSION PHASE 3 (PHASE 1 .NET)
'BEGIN PREM TRIS 5.5
Public Function GetHolidays(ByVal vdtStartDate As Date _
        , ByVal vdtEndDate As Date _
        , Optional ByVal vsUnionCode As String = "" _
        ) As ADODB.Recordset
On Error GoTo GetHolidays_EH
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables
Dim lcollUnion As Collection

    If (vsUnionCode <> "") Then
        Set lcollUnion = New Collection
        lcollUnion.Add vsUnionCode
    End If

    Set lobjdbCodeTables = CreateObject( _
                            "dbCodeTables.cdCodeTables")
    Set GetHolidays = lobjdbCodeTables.GetHolidaysBetween( _
                        vdtStartDate _
                        , vdtEndDate _
                        , lcollUnion)
    Set lobjdbCodeTables = Nothing
    Set lcollUnion = Nothing

    Exit Function
GetHolidays_EH:
    Set lobjdbCodeTables = Nothing
    Set lcollUnion = Nothing
    TrsRaiseError m_strCMODULENAME _
        , "GetHolidays" _
        , , , , _
        , True

End Function
'END PREM TRIS 5.5
Public Function GetJobTaskExceptions() As ADODB.Recordset
    
    Dim l_CodeTables As dbCodeTables.cdCodeTables
    
    On Error GoTo GetJobtaskexceptions_EH
                                                                    
    Set l_CodeTables = CreateObject("dbCodeTables.cdCodeTables")
    Set GetJobTaskExceptions = l_CodeTables.GetJobTaskExceptions
    Set l_CodeTables = Nothing
    Exit Function

GetJobtaskexceptions_EH:
    Set l_CodeTables = Nothing

    TrsRaiseError m_strCMODULENAME, "GetJobTaskExceptions", , , , , True

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
    TrsRaiseError m_strCMODULENAME, "DoesRuleApply", , , , , True
    
End Function

