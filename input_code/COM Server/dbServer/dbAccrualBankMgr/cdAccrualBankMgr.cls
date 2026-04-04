VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdAccrualBankMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Private Const m_strCMODULENAME As String = "cdAccrualBankMgr"
Private Const m_DBSCRIPTSROOT As String = "C:\TRIS SOLUTION\dbScripts\TRIS_db2"
Private Const m_DBSCRIPTSPATH As String = "\dbAccrualBankMgr\cdAccrualBankMgr"
Private m_ResourceMgr As utilADOHelper.cuResourceMgr

Public Function UpdateEmployeeBanks( _
                    ByVal v_rsEmployeeBanks As ADODB.Recordset, _
                    ByVal v_rsAuditTrail As ADODB.Recordset, _
                    ByVal v_lngLastModEmpNum As Long _
                    ) As Boolean
    TrsRaiseError m_strCMODULENAME, "This method has been moved to cdAccrualBankMgrTrans.", , , , , True
End Function

Public Function GetEmployeeQuarterlyOTBanks(ByVal v_lngEmpNum As Long, _
                                                                        ByVal v_datStartDate As Date _
                                                                        ) As ADODB.Recordset
    
TrsRaiseError m_strCMODULENAME, "GetEmployeeQuarterlyOTBanks", trErrObsoleteMethod, , , , True
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'    Dim l_lngCurrentYear As Long
'    Dim l_lngPriorYear As Long
'    Dim l_lngNextYear As Long
'    Dim l_strInClause As String
'
'    On Error GoTo GetEmployeeQuarterlyOTBanks_EH
'
'
'    Set l_objCN = New ADODB.Connection
'    Set l_objRS = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
'
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.CursorLocation = adUseClient
'    l_objCN.Open
'
'    l_lngCurrentYear = Year(v_datStartDate)
'    l_lngPriorYear = l_lngCurrentYear - 1
'    l_lngNextYear = l_lngCurrentYear + 1
'    l_strInClause = str(l_lngCurrentYear) & "," & str(l_lngPriorYear) & "," & str(l_lngNextYear)
'
'    l_strSQL = "SELECT EMP_NUM, EFF_YEAR, QTR_1_HRS, QTR_2_HRS, QTR_3_HRS, QTR_4_HRS "
'    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_EXEC_HRS "
'    l_strSQL = l_strSQL & "WHERE EMP_NUM = " & v_lngEmpNum & " "
'    l_strSQL = l_strSQL & "AND EFF_YEAR IN (" & l_strInClause & ") "
'    l_strSQL = l_strSQL & "ORDER BY EFF_YEAR DESC"
'
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'    Set l_objCN = Nothing
'
'    Set GetEmployeeQuarterlyOTBanks = l_objRS
'
'    Set l_objRS = Nothing
'
'Exit Function
'GetEmployeeQuarterlyOTBanks_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetEmployeeQuarterlyOTBanks", , , , DescArgsFrom(l_strSQL), True

End Function

Private Function AreFamilyCareBanksActive() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim UseNewBanks As String
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    AreFamilyCareBanksActive = False
    
    UseNewBanks = UCase(dbEnv.GetSysParmsValue(51, 1))
    
    If UseNewBanks = "Y" Or UseNewBanks = "TRUE" Or UseNewBanks = "1" Then
        AreFamilyCareBanksActive = True
    End If
        
    Set dbEnv = Nothing

    Exit Function

eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "AreFamilyCareBanksActive", , , , DescArgsFrom(UseNewBanks), True
End Function

Public Function GetEmployeeBanks( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date) As ADODB.Recordset
    
    
'    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_lngCurrentYear As Long
    Dim l_lngPriorYear As Long
    Dim l_lngNextYear As Long
    Dim l_strInClause As String
    Dim l_strEmployeeSelection As String
    Dim oADOHelper As utilADOHelper.cuADOHelper

    On Error GoTo eh
    
    
'    Set l_objCN = New ADODB.Connection
'    Set l_objRS = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
'
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.CursorLocation = adUseClient
'    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    
    l_lngCurrentYear = Year(v_datStartDate)
    l_lngPriorYear = l_lngCurrentYear - 1
    l_lngNextYear = l_lngCurrentYear + 1
    l_strInClause = str(l_lngCurrentYear) & "," & str(l_lngPriorYear) & "," & str(l_lngNextYear)
    
    Set oADOHelper = New utilADOHelper.cuADOHelper

    If AreFamilyCareBanksActive Then
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetEmployeeBanks.sql")
    Else
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetEmployeeBanks_FamilyCareBanksNOTActive.sql")
    End If
    
    l_strSQL = Replace(l_strSQL, "{EffYearInClause}", l_strInClause)
    l_strSQL = Replace(l_strSQL, "{EmployeeSelection}", l_strEmployeeSelection)
    
    Set l_objRS = oADOHelper.ExecuteSelect(l_strSQL, adOpenStatic, adLockBatchOptimistic)
       
'    l_strSQL = "select A.EMP_NUM, EFF_YEAR, VAC_HRS_ALLOTTED, VAC_HRS_USED, VAC_HRS_SCHED , " _
'        & " decimal((VAC_HRS_ALLOTTED - VAC_HRS_USED - VAC_HRS_SCHED), 7,1) AS VAC_HRS_AVAIL, " _
'        & " COMP_STR_TIME_HRS, COMP_TAKEN_SCHED, " _
'        & " decimal((COMP_STR_TIME_HRS - COMP_TAKEN_SCHED),7,1) AS COMP_HRS_AVAIL, " _
'        & " HOL_HRS_ACCRUED, HOL_TAKEN_SCHED, " _
'        & " decimal((HOL_HRS_ACCRUED - HOL_TAKEN_SCHED),7,1) AS HOL_HRS_AVAIL, " _
'        & " FMLA_HRS_ALLOWED, FMLA_HRS_USED , FMLA_HRS_SCHED, " _
'        & " decimal((FMLA_HRS_ALLOWED - FMLA_HRS_USED - FMLA_HRS_SCHED),7,1) AS FMLA_HRS_AVAIL , " _
'        & " B.BIG_BANK_FULL_HRS, B.BIG_BANK_HALF_HRS, B.BIG_BANK_DELAY_HRS, " _
'        & " VALUE(NULLIF(B.BIG_BANK_INT_DATE,'0001-01-01'),Date('0100-01-01')) AS BIG_BANK_INT_DATE , " _
'        & " VALUE(NULLIF(B.BIG_BANK_RST_DATE,'0001-01-01'),Date('0100-01-01')) AS BIG_BANK_RST_DATE , " _
'        & " B.LIT_BANK_TOT_HRS, B.LIT_BANK_DELAY_HRS, " _
'        & " VALUE(NULLIF(B.LIT_BANK_INT_DATE,'0001-01-01'),Date('0100-01-01')) AS LIT_BANK_INT_DATE , " _
'        & " VALUE(NULLIF(B.LIT_BANK_RST_DATE,'0001-01-01'),Date('0100-01-01')) AS LIT_BANK_RST_DATE,  " _
'        & " FIRST_HR_NO, BASIC_HR_NO, ACCUM_HR_NO, EXTRA_FIRST_HR_NO , " _
'        & " ALWD_NON_AH_NO, ACCUM_OFLOW_HR_NO  " _
'        & "From methdba.ttr_emp_banks A, methdba.ttr_emp_sl_banks B " _
'        & "where " & l_strEmployeeSelection _
'        & "  and B.emp_num = A.emp_num " _
'        & " and eff_year in (" & l_strInClause & ") " _
'        & " order by A.EMP_NUM, eff_year"
        
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing

    Set GetEmployeeBanks = l_objRS
    
    Set l_objRS = Nothing
    Set oADOHelper = Nothing

    Exit Function
eh:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
    Set l_objRS = Nothing
    Set oADOHelper = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeeBanks", , , , DescArgsFrom(l_strSQL), True
End Function


Public Function GetEmployeeVacationBankAudit( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    l_strSQL = "SELECT EMP_NUM,CHARGE_DATE, EFF_YEAR, BUCKET_IND_CD, " _
        & " SUM(HOURS) as HOURS " _
        & " FROM METHDBA.TTR_EMP_BANK_AUDIT " _
        & " WHERE " & BuildInClause("EMP_NUM", v_aryEmployees, False) _
        & " AND BANK_CD = 'V'" _
        & " AND CHARGE_DATE >= " & DateAsEnclosedString(v_datStartDate) _
        & " GROUP BY EMP_NUM, CHARGE_DATE, EFF_YEAR, BUCKET_IND_CD " _
        & " HAVING SUM(HOURS) > 0 " _
        & " ORDER BY EMP_NUM, CHARGE_DATE, EFF_YEAR, BUCKET_IND_CD"
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeVacationBankAudit = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing

    TrsRaiseError m_strCMODULENAME, "GetEmployeeVacationBankAudit", , , , DescArgsFrom(l_strSQL), True
End Function


Public Function GetEmployeeHolidayBankAudit( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    
    l_strSQL = "SELECT EMP_NUM,CHARGE_DATE, EFF_YEAR, BUCKET_IND_CD, " _
        & " SUM(HOURS) as HOURS " _
        & " FROM METHDBA.TTR_EMP_BANK_AUDIT " _
        & " WHERE " & BuildInClause("EMP_NUM", v_aryEmployees, False) _
        & " AND BANK_CD = 'H'" _
        & " AND BUCKET_IND_CD = 'A'" _
        & " AND CHARGE_DATE >= " & DateAsEnclosedString(v_datStartDate) _
        & " GROUP BY EMP_NUM, CHARGE_DATE, EFF_YEAR, BUCKET_IND_CD " _
        & " HAVING SUM(HOURS) > 0 " _
        & " ORDER BY EMP_NUM, CHARGE_DATE, EFF_YEAR, BUCKET_IND_CD"
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeHolidayBankAudit = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing

    TrsRaiseError m_strCMODULENAME, "GetEmployeeHolidayBankAudit", , , , DescArgsFrom(l_strSQL), True
End Function

Private Sub Class_Initialize()
    Set m_ResourceMgr = New utilADOHelper.cuResourceMgr
End Sub

Private Sub Class_Terminate()
    Set m_ResourceMgr = Nothing
End Sub
