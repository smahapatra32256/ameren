VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const m_strCMODULENAME As String = "cbsEmployee"
Private Const mstrDefaultDate As String = "01/01/0100"
'Private Const mstrNotAvailable As String = "NA"

Public Function GetViewEmployeeInformation( _
            ByVal v_lngEmpNum As Long _
            , ByRef r_adrsBasicInformation As ADODB.Recordset _
            , ByRef r_adrsPrimaryWG As ADODB.Recordset _
            , ByRef r_adrsEmployeeWorkGroups As ADODB.Recordset _
            , ByRef r_adrsEmployeeReliefGroups As ADODB.Recordset _
            , ByRef r_adrsEmployeeBanks As ADODB.Recordset _
            , ByRef r_adrsEmployeeOTBanks As ADODB.Recordset _
            , ByRef r_adrsEmployeeOTQTRBanks As ADODB.Recordset _
            , ByRef r_adrsEmployeeQuals As ADODB.Recordset _
            , ByRef r_adrsEmployeeSpecSkills As ADODB.Recordset _
            ) As Boolean
            
TrsRaiseError m_strCMODULENAME, "GetViewEmployeeInformation", trErrObsoleteMethod, , , , True
'    Dim l_aryEmployees(0) As String
'
'    Dim l_bolGetEmpBanks As Boolean
'    Dim l_datToday As Date
'    Dim l_rsSched As ADODB.Recordset
'
'On Error GoTo GetViewEmployeeInformation_EH
'
'    l_datToday = Date
'    Set l_rsSched = GetSchedule(v_lngEmpNum, l_datToday, l_datToday)
'    Set r_adrsBasicInformation = GetBasicEmployeeInformation(v_lngEmpNum)
'    FillScheduleData r_adrsBasicInformation, l_rsSched
'    Set r_adrsPrimaryWG = GetEmployeePrimaryWorkGroup(v_lngEmpNum)
'    Set r_adrsEmployeeWorkGroups = GetEmployeeWorkGroups(v_lngEmpNum)
'    Set r_adrsEmployeeReliefGroups = GetEmployeeReliefGroups(v_lngEmpNum)
'    Set r_adrsEmployeeQuals = GetEmployeeQualifications(v_lngEmpNum)
'    Set r_adrsEmployeeSpecSkills = GetEmployeeSkills(v_lngEmpNum)
'    l_aryEmployees(0) = v_lngEmpNum
'    l_bolGetEmpBanks = GetEmployeeBanks(l_aryEmployees, l_datToday, r_adrsEmployeeBanks, r_adrsEmployeeOTBanks, r_adrsEmployeeOTQTRBanks)
'
'    GetViewEmployeeInformation = True
'
'Exit Function
'
'GetViewEmployeeInformation_EH:
'    Set r_adrsBasicInformation = Nothing
'    Set r_adrsPrimaryWG = Nothing
'    Set r_adrsEmployeeWorkGroups = Nothing
'    Set r_adrsEmployeeReliefGroups = Nothing
'    Set r_adrsEmployeeBanks = Nothing
'    Set r_adrsEmployeeOTBanks = Nothing
'    Set r_adrsEmployeeOTQTRBanks = Nothing
'    Set r_adrsEmployeeQuals = Nothing
'    Set r_adrsEmployeeSpecSkills = Nothing
'    GetViewEmployeeInformation = False
'    TrsRaiseError m_strCMODULENAME, "GetViewEmployeeInformation", , , , , True

End Function

Private Function GetSchedule( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date _
    ) As ADODB.Recordset

    Dim l_obj As busSchedule.cbsSchedule
    Dim l_rsDummy As ADODB.Recordset

    On Error GoTo eh
    Set l_obj = CreateObject("busSchedule.cbsSchedule")
    Set GetSchedule = l_obj.GetAdjustedBaseSchedule( _
        v_lngEmpNum, v_datStart, v_datEnd, l_rsDummy)

    Set l_obj = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetSchedule"

End Function


Private Sub FillScheduleData( _
    ByRef r_rsEmpInfo As ADODB.Recordset, _
    ByVal v_rsSched As ADODB.Recordset)

    Dim l_intRotCode As Integer
    Dim l_intSchedCode As Integer
    Dim l_strRotDesc As String
    Dim l_strSchedDesc As String

    On Error GoTo eh

    If Not RecordsExist(r_rsEmpInfo) Then
        Exit Sub
    End If

    v_rsSched.MoveFirst
    l_intSchedCode = v_rsSched!SCHED_CD
    r_rsEmpInfo!SCHED_CD = l_intSchedCode
    r_rsEmpInfo!SCHED_DESC = GetSchedCodeDesc(l_intSchedCode)
    r_rsEmpInfo!SCHED_SRC_CD = v_rsSched!SCHED_SRC_CD
    l_intRotCode = v_rsSched!ROT_CD
    r_rsEmpInfo!ROT_CD = l_intRotCode
    r_rsEmpInfo!ROT_WEEK = v_rsSched!ROT_WEEK
    r_rsEmpInfo!ROT_DESC = GetRotCodeDesc(l_intRotCode)

    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "FillScheduleData"

End Sub
Private Function GetSchedCodeDesc( _
    ByVal v_intCode As Integer _
    ) As String

    Dim l_obj As dbSchedCodeMaint.cdSchedCodeMaint
    Dim l_rs As ADODB.Recordset

    On Error GoTo eh

    If v_intCode <= 0 _
        And v_intCode <> -9 _
        And v_intCode <> -2 Then
        GetSchedCodeDesc = ""
        Exit Function
    End If

    Set l_obj = CreateObject("dbSchedCodeMaint.cdSchedCodeMaint")
    Set l_rs = l_obj.GetScheduleDetail(v_intCode)
    If RecordsExist(l_rs) Then
        GetSchedCodeDesc = l_rs!SCHED_DESC
    Else
        GetSchedCodeDesc = ""
    End If
    Set l_obj = Nothing
    Set l_rs = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    Set l_rs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetSchedCodeDesc"
End Function


Private Function GetRotCodeDesc( _
    ByVal v_intCode As Integer _
    ) As String

    Dim l_obj As dbSchedRotation.cdSchedRotation
    Dim l_rs As ADODB.Recordset

    On Error GoTo eh

    If v_intCode <= 0 Then
        GetRotCodeDesc = ""
        Exit Function
    End If

    Set l_obj = CreateObject("dbSchedRotation.cdSchedRotation")
    Set l_rs = l_obj.GetRotationSchedHeader(v_intCode)
    If RecordsExist(l_rs) Then
        GetRotCodeDesc = l_rs!ROT_DESC
    Else
        GetRotCodeDesc = ""
    End If
    Set l_obj = Nothing
    Set l_rs = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    Set l_rs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetRotCodeDesc"
End Function


Private Function GetEmployeePrimaryWorkGroup(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo eh

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeePrimaryWorkGroup = l_objEmployee.GetEmployeePrimaryWorkGroup(v_lngEmpNum)

    Set l_objEmployee = Nothing
    Exit Function
    
eh:
    Set GetEmployeePrimaryWorkGroup = Nothing
    Set l_objEmployee = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeePrimaryWorkGroup", , , , , True

End Function

Private Function GetEmployeeWorkGroups(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo GetEmployeeWorkGroups_EH

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeeWorkGroups = l_objEmployee.GetEmployeeWorkGroups(v_lngEmpNum)

    Set l_objEmployee = Nothing
    Exit Function
    
GetEmployeeWorkGroups_EH:
    Set GetEmployeeWorkGroups = Nothing
    Set l_objEmployee = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeWorkGroups", , , , , True

End Function

Private Function GetEmployeeReliefGroups(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo GetEmployeeReliefGroups_EH

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeeReliefGroups = l_objEmployee.GetEmployeeReliefGroups(v_lngEmpNum)

    Set l_objEmployee = Nothing
    Exit Function
    
GetEmployeeReliefGroups_EH:
    Set GetEmployeeReliefGroups = Nothing
    Set l_objEmployee = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeReliefGroups", , , , , True

End Function

Private Function GetEmployeeQualifications(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo eh

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeeQualifications = l_objEmployee.GetEmployeeQualifications(v_lngEmpNum)
    Set l_objEmployee = Nothing
    Exit Function
    
eh:
    Set GetEmployeeQualifications = Nothing
    Set l_objEmployee = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeQualifications", , , , , True

End Function

Private Function GetEmployeeSkills(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    
    Dim l_objEmployee As dbEmployee.cdEmployee

    On Error GoTo eh

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeeSkills = l_objEmployee.GetEmployeeSkills(v_lngEmpNum)
    Set l_objEmployee = Nothing
    Exit Function
    
eh:
    Set GetEmployeeSkills = Nothing
    Set l_objEmployee = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeSkills", , , , , True

End Function

Private Function GetEmployeeBanks( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datCurrDate As Date, _
                    ByRef r_rsBanks As ADODB.Recordset, _
                    ByRef r_rsOTBanks As ADODB.Recordset, _
                    ByRef r_rsQuarterlyOTBanks As ADODB.Recordset _
                    ) As Boolean
                                                        
'  Get the Employee Banks (Vacation, FMLA, Holiday, Compensatory) and Quarterly Overtime Banks
'  Passing in the Current Year results in getting the Prior Year and Next Year data as well
                                                        
    Dim l_objdbAccrualBankMgr As dbAccrualBankMgr.cdAccrualBankMgr
    Dim l_objdbOTEqualizer As dbOvertimeEqualizer.cdOvertimeEqualizer
    
On Error GoTo eh
    
    Set l_objdbAccrualBankMgr = CreateObject("dbAccrualBankMgr.cdAccrualBankMgr")
    Set r_rsBanks = l_objdbAccrualBankMgr.GetEmployeeBanks(v_aryEmployees, v_datCurrDate)
    Set l_objdbAccrualBankMgr = Nothing
    
    Set l_objdbOTEqualizer = CreateObject("dbOvertimeEqualizer.cdOvertimeEqualizer")
'    Set r_rsQuarterlyOTBanks = l_objdbOTEqualizer.GetEmployeeQuarterlyOTBanks(v_aryEmployees, v_datCurrDate)
    Set r_rsQuarterlyOTBanks = Nothing '//this is intentional as it is not used anymore
    Set r_rsOTBanks = l_objdbOTEqualizer.GetEmployeeOvertimeBanks(v_aryEmployees)
    Set l_objdbOTEqualizer = Nothing
                                                        
    GetEmployeeBanks = True
    Exit Function

eh:
    Set r_rsBanks = Nothing
    Set r_rsOTBanks = Nothing
    Set r_rsQuarterlyOTBanks = Nothing
    Set l_objdbAccrualBankMgr = Nothing
    GetEmployeeBanks = False
    TrsRaiseError m_strCMODULENAME, "GetEmployeeBanks", , , , , True

End Function

Private Function GetBasicEmployeeInformation(ByVal v_lngEmpNum As Long) As ADODB.Recordset
    Dim l_objEmployee As dbEmployee.cdEmployee
    Dim l_objbusEmployee As busEmployee.cbsEmployee
    Dim l_sEmployeeBenefitPackage As String
    On Error GoTo eh

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetBasicEmployeeInformation = l_objEmployee.GetBasicEmployeeInformation(v_lngEmpNum)
    Set l_objEmployee = Nothing
    
    If RecordsExist(GetBasicEmployeeInformation) Then
        GetBasicEmployeeInformation.MoveFirst
        Set l_objbusEmployee = CreateObject("busEmployee.cbsEmployee")
        Do While Not GetBasicEmployeeInformation.EOF
            l_sEmployeeBenefitPackage = l_objbusEmployee.EmployeeBenefitPackage(GetBasicEmployeeInformation!JOB_EFF_DATE, GetBasicEmployeeInformation!MGMT_BEN_PKG_IND, _
                                                            GetBasicEmployeeInformation!UNION_CD, GetBasicEmployeeInformation!PRIOR_UNION_CD, _
                                                            GetBasicEmployeeInformation!EXEC_CONTRACT_IND)
            
            GetBasicEmployeeInformation!EMP_BEN_PKG_IND = l_sEmployeeBenefitPackage
            GetBasicEmployeeInformation!CIPS_IND = IIf(l_objbusEmployee.IsAmerenCIPSorCILCO(dbRules.TRSRuleCode.trIsAmerenCIPS, GetBasicEmployeeInformation!DEPT_EFF_DATE, GetBasicEmployeeInformation!DEPT_NUM, GetBasicEmployeeInformation!PREV_DEPT_NUM, , , Date), "1", "0")
            GetBasicEmployeeInformation!CILCO_IND = IIf(l_objbusEmployee.IsAmerenCIPSorCILCO(dbRules.TRSRuleCode.trIsAmerenCILCO, GetBasicEmployeeInformation!DEPT_EFF_DATE, GetBasicEmployeeInformation!DEPT_NUM, GetBasicEmployeeInformation!PREV_DEPT_NUM, , , Date), "1", "0")
            GetBasicEmployeeInformation.Update
            
            GetBasicEmployeeInformation.MoveNext
        Loop
        GetBasicEmployeeInformation.MoveFirst
        Set l_objbusEmployee = Nothing
    End If
        
    Exit Function
eh:
    Set GetBasicEmployeeInformation = Nothing
    Set l_objEmployee = Nothing
    Set l_objbusEmployee = Nothing
    TrsRaiseError m_strCMODULENAME, "GetBasicEmployeeInformation", , , , , True
End Function

'Begin Prem TRIS 5.7
Public Sub GetViewEmployeeInformation2(ByVal vlEmployee As Long _
            , ByRef radrsBasicInformation As ADODB.Recordset _
            , ByRef radrsPrimaryWG As ADODB.Recordset _
            , ByRef radrsEmployeeWorkGroups As ADODB.Recordset _
            , ByRef radrsEmployeeReliefGroups As ADODB.Recordset _
            , ByRef radrsEmployeeBanks As ADODB.Recordset _
            , ByRef radrsEmployeeOTBanks As ADODB.Recordset _
            , ByRef radrsEmployeeOTQTRBanks As ADODB.Recordset _
            , ByRef radrsEmployeeQuals As ADODB.Recordset _
            , ByRef radrsEmployeeSpecSkills As ADODB.Recordset _
            , ByRef radrsContact As ADODB.Recordset)
On Error GoTo GetViewEmployeeInformation2_EH
Dim lARREmployees(0) As String, lbGetEmpBanks As Boolean
Dim ladrsSchedule As ADODB.Recordset

    Set ladrsSchedule = GetSchedule(vlEmployee, Date, Date)
    Set radrsBasicInformation = GetBasicEmployeeInformation(vlEmployee)

    FillScheduleData radrsBasicInformation, ladrsSchedule
    Set ladrsSchedule = Nothing

    Set radrsPrimaryWG = GetEmployeePrimaryWorkGroup(vlEmployee)
    Set radrsEmployeeWorkGroups = GetEmployeeWorkGroups(vlEmployee)
    Set radrsEmployeeReliefGroups = GetEmployeeReliefGroups(vlEmployee)
    Set radrsEmployeeQuals = GetEmployeeQualifications(vlEmployee)
    Set radrsEmployeeSpecSkills = GetEmployeeSkills(vlEmployee)

    lARREmployees(0) = CStr(vlEmployee)
    lbGetEmpBanks = GetEmployeeBanks(lARREmployees _
                        , Date _
                        , radrsEmployeeBanks _
                        , radrsEmployeeOTBanks _
                        , Nothing)
    Set radrsEmployeeOTQTRBanks = Nothing '//this is intentional as it is not used anymore
    
    Set radrsContact = GetEmployeePSContact(vlEmployee)

    Exit Sub
GetViewEmployeeInformation2_EH:
    Set ladrsSchedule = Nothing

    Set radrsBasicInformation = Nothing
    Set radrsPrimaryWG = Nothing
    Set radrsEmployeeWorkGroups = Nothing
    Set radrsEmployeeReliefGroups = Nothing
    Set radrsEmployeeBanks = Nothing
    Set radrsEmployeeOTBanks = Nothing
    Set radrsEmployeeOTQTRBanks = Nothing
    Set radrsEmployeeQuals = Nothing
    Set radrsEmployeeSpecSkills = Nothing
    Set radrsContact = Nothing

    TrsRaiseError m_strCMODULENAME, "GetViewEmployeeInformation", , , , , True
End Sub
Private Function GetEmployeePSContact(ByVal vlEmployee As Long _
        ) As ADODB.Recordset
On Error GoTo GetEmployeePSContact_EH
Dim lobjdbEmployee As dbEmployee.cdEmployee

    Set GetEmployeePSContact = Nothing

    Set lobjdbEmployee = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployeePSContact = _
            lobjdbEmployee.GetEmployeePSContactInformation(vlEmployee)
    Set lobjdbEmployee = Nothing

    Exit Function
GetEmployeePSContact_EH:
    Set lobjdbEmployee = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployeePSContact" _
        , , , , _
        , True
End Function
'End Prem TRIS 5.7
