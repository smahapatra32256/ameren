VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdCodeTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

' Static data constants
Private Const CRECORDNOTFOUND = "RECORDNOTFOUND"
Private Const CSPINVALID = "SPINVALID"
Private Const CDEFAULTDATE As String = "01/01/0100"

' Location code constants
Private Const CLC_GROUPNAME = "LOCATIONCACHE"
' index for record dynamic array
'Private Const CLC_LOC_CD = 0
'Private Const CLC_LOC_DESC = 1
'Private Const CLC_PLANT_CD = 2
'Private Const CLC_LAST_SLOT = 2
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const CLC_LOC_CD = 1
Private Const CLC_LOC_DESC = 2
Private Const CLC_PLANT_CD = 3
Private Const CLC_LAST_SLOT = 3

' Job number constants
Private Const CJB_GROUPNAME = "JOBCACHE"
' index for record dynamic array
'Private Const CJB_JOB_NUM = 0
'Private Const CJB_JOB_DESC = 1
'Private Const CJB_CRAFT_NUM = 2
'Private Const CJB_OT_TYPE = 3
'Private Const CJB_EFF_DT = 4
'Private Const CJB_EFF_STATUS = 5
'Private Const CJB_UNION_CD = 6
'Private Const CJB_LAST_SLOT = 6
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const CJB_JOB_NUM = 1
Private Const CJB_JOB_DESC = 2
Private Const CJB_CRAFT_NUM = 3
Private Const CJB_OT_TYPE = 4
Private Const CJB_EFF_DT = 5
Private Const CJB_EFF_STATUS = 6
Private Const CJB_UNION_CD = 7
Private Const CJB_LAST_SLOT = 7

' Exception code constants
Private Const CEX_GROUPNAME = "EXCEPTIONCACHE"
' index for record dynamic array
'Private Const CEX_EXCEPTION_CD = 0
'Private Const CEX_EXCEPTION_DESC = 1
'Private Const CEX_EXCEPTION_TYPE = 2
'Private Const CEX_CONTRACT_EARN_CD = 3
'Private Const CEX_EXEC_EARN_CD = 4
'Private Const CEX_SYMBOL_TIME_IND = 5
'Private Const CEX_EXCEPTION_BANK_CD = 6
'Private Const CEX_INELIGIBILITY_IND = 7
'Private Const CEX_OTHERPAY_IND = 8
'Private Const CEX_CASE_GROUP = 9
'Private Const CEX_SCHED_IND = 10
'Private Const CEX_FMLA_IND = 11
'Private Const CEX_WORK_REST_IND = 12
'Private Const CEX_SCHED_CHG_IND = 13
'Private Const CEX_ACCT_NUM = 14
'Private Const CEX_EXCEPTION_SYMBOL_CD = 15
'Private Const CEX_LAST_SLOT = 15
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const CEX_EXCEPTION_CD = 1
Private Const CEX_EXCEPTION_DESC = 2
Private Const CEX_EXCEPTION_TYPE = 3
Private Const CEX_CONTRACT_EARN_CD = 4
Private Const CEX_EXEC_EARN_CD = 5
Private Const CEX_SYMBOL_TIME_IND = 6
Private Const CEX_EXCEPTION_BANK_CD = 7
Private Const CEX_INELIGIBILITY_IND = 8
Private Const CEX_OTHERPAY_IND = 9
Private Const CEX_CASE_GROUP = 10
Private Const CEX_SCHED_IND = 11
Private Const CEX_FMLA_IND = 12
Private Const CEX_WORK_REST_IND = 13
Private Const CEX_SCHED_CHG_IND = 14
Private Const CEX_ACCT_NUM = 15
Private Const CEX_EXCEPTION_SYMBOL_CD = 16
Private Const CEX_AIC_ACCT_NUM = 17
Private Const CEX_AMS_ACCT_NUM = 18
Private Const CEX_UEC_ACCT_NUM = 19
Private Const CEX_LAST_SLOT = 19

' Union code constants
Private Const CUN_GROUPNAME = "UNIONCACHE"
' index for record dynamic array
'Private Const CUN_UNION_CD = 0
'Private Const CUN_UNION_LOCAL = 1
'Private Const CUN_UNION_DESC = 2
'Private Const CUN_EXEC_CONTRACT_IND = 3
'Private Const CUN_UNION_BEN_PKG_IND = 4
'Private Const CUN_LAST_SLOT = 4
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const CUN_UNION_CD = 1
Private Const CUN_UNION_LOCAL = 2
Private Const CUN_UNION_DESC = 3
Private Const CUN_EXEC_CONTRACT_IND = 4
Private Const CUN_UNION_BEN_PKG_IND = 5
Private Const CUN_LAST_SLOT = 5

' Union component constants
Private Const CUC_GROUPNAME = "EFFUNIONCOMPCACHE"
Private Const CUC_OTHER_COMP_TYPE = "O"
Private Const CUC_RELIEF_COMP_TYPE = "R"
Private Const CUC_SHIFT_COMP_TYPE = "S"
' SharedProperty value(CUC_ROWVALUES) is a 2 dim array with 2 fields,
' EFF_DATE and COMP_AMT.
' index for record dynamic array
Private Const CUC_COMP_KEY = 0
Private Const CUC_ROWVALUES = 1
Private Const CUC_LAST_SLOT = 1

Private Const CUC_EFF_DATE = 0
Private Const CUC_COMP_AMT = 1

' Overtime reason code constants
Private Const COR_GROUPNAME = "OTREASONCACHE"
' index for record dynamic array
'Private Const COR_OT_REASON_CD = 0
'Private Const COR_OT_REASON_DESC = 1
'Private Const COR_LAST_SLOT = 1
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const COR_OT_REASON_CD = 1
Private Const COR_OT_REASON_DESC = 2
Private Const COR_LAST_SLOT = 2

' Overtime response code constants
Private Const COP_GROUPNAME = "OTRESPONSECACHE"
' index for record dynamic array
'Private Const COP_OT_RESP_CD = 0
'Private Const COP_RESPONSE_DESC = 1
'Private Const COP_OT_CHRG_IND = 2
'Private Const COP_OT_WORKED_IND = 3
'Private Const COP_LAST_SLOT = 3
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const COP_OT_RESP_CD = 1
Private Const COP_RESPONSE_DESC = 2
Private Const COP_OT_CHRG_IND = 3
Private Const COP_OT_WORKED_IND = 4
Private Const COP_LAST_SLOT = 4

' Overtime exception code constants
Private Const COE_GROUPNAME = "OTEXCEPTIONCACHE"
' index for record dynamic array
'Private Const COE_EXCEPTION_CD = 0
'Private Const COE_EXCEPTION_DESC = 1
'Private Const COE_FACTOR = 2
'Private Const COE_EXEC_CONTRACT_IND = 3
'Private Const COE_EARNINGS_CD = 4
'Private Const COE_LAST_SLOT = 4
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const COE_EXCEPTION_CD = 1
Private Const COE_EXCEPTION_DESC = 2
Private Const COE_FACTOR = 3
Private Const COE_EXEC_CONTRACT_IND = 4
Private Const COE_EARNINGS_CD = 5
Private Const COE_LAST_SLOT = 5

' Actual Time History CutOff date constants
Private Const CHS_GROUPNAME = "HSCUTOFFDATECACHE"
' index for record dynamic array
'Private Const CHS_CUTOFF_CD = 0
'Private Const CHS_CUTOFF_DT = 1
'Private Const CHS_LAST_SLOT = 1
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const CHS_CUTOFF_CD = 1
Private Const CHS_CUTOFF_DT = 2
Private Const CHS_LAST_SLOT = 2

' Available schedule code constants
Private Const CSC_GROUPNAME = "AVAILSCHEDCACHE"
' index for record dynamic array
'Private Const CSC_SCHED_CD = 0
'Private Const CSC_LAST_SLOT = 0
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const CSC_SCHED_CD = 1
Private Const CSC_LAST_SLOT = 1

' Rotation code constants
Private Const CRC_GROUPNAME = "ROTATIONCACHE"
' index for record dynamic array
'Private Const CRC_ROT_CD = 0
'Private Const CRC_ROT_WEEKS = 1
'Private Const CRC_LAST_SLOT = 1
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const CRC_ROT_CD = 1
Private Const CRC_ROT_WEEKS = 2
Private Const CRC_LAST_SLOT = 2

' Component code constants
Private Const CCC_GROUPNAME = "COMPCODECACHE"
' index for record dynamic array
'Private Const CCC_COMP_CD = 0
'Private Const CCC_COMP_TYPE = 1
'Private Const CCC_LAST_SLOT = 1
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const CCC_COMP_CD = 1
Private Const CCC_COMP_TYPE = 2
Private Const CCC_LAST_SLOT = 2

' Exception Group constants
Private Const CEG_GROUPNAME = "EXCEPTIONGROUP"
Private Const CEG_KEY = 0
Private Const CEG_SEQNUM = 1
Private Const CEG_LAST_SLOT = 1

'Enable trace info by providing conditional constant
' TRDEBUG = -1 here or project properties dialog.
' Disable with TRDEBUG = 0.
'#Const TRDEBUG = -1

Public Enum TrsComponentType
    trReliefComponent = 1
    trShiftComponent
    trOtherComponent
End Enum

Private Const m_strCMODULENAME As String = "cdCodeTables"
Private Const m_DBSCRIPTSROOT As String = "C:\TRIS SOLUTION\dbScripts\TRIS_db2"
Private Const m_DBSCRIPTSPATH As String = "\dbCodeTables\cdCodeTables"
Private m_ResourceMgr As utilADOHelper.cuResourceMgr

'*********************************************************
' Begin Public procedures
'*********************************************************
Public Sub GetAllSupportInformation( _
        ByRef r_rsAllExceptions As ADODB.Recordset, _
        ByRef r_rsAllComponents As ADODB.Recordset, _
        ByRef r_rsAllHolidays As ADODB.Recordset, _
        ByRef r_rsAllJobs As ADODB.Recordset, _
        ByRef r_rsAllOvertimeReasons As ADODB.Recordset, _
        ByRef r_rsAllLocations As ADODB.Recordset, _
        ByRef r_rsAllUnions As ADODB.Recordset, _
        ByRef r_rsAllDepartments As ADODB.Recordset, _
        ByRef r_rsAllActivityList As ADODB.Recordset)
        
On Error GoTo eh
    
    Set r_rsAllExceptions = GetAllExceptions
    Set r_rsAllComponents = GetUniqueComponentCodes 'PTS 870 replaced GetUnionComponents
    Set r_rsAllHolidays = GetAllHolidays
    Set r_rsAllJobs = GetAllJobs
    Set r_rsAllOvertimeReasons = GetAllOvertimeReasons
    Set r_rsAllLocations = GetAllLocations
    Set r_rsAllUnions = GetAllUnions
    Set r_rsAllDepartments = GetAllDepartments
    Set r_rsAllActivityList = GetAllActivityList
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "GetAllSupportInformation", , , , , True
End Sub

Public Function GetAllDepartments() As ADODB.Recordset
    Dim l_dbDepartment As dbDepartment.cdDepartment

On Error GoTo eh

    Set l_dbDepartment = _
        CreateObject("dbDepartment.cdDepartment")
    'Set l_dbDepartment = New dbDepartment.cdDepartment
    Set GetAllDepartments = l_dbDepartment.GetAllDepartments
    Set l_dbDepartment = Nothing
    Exit Function

eh:
    Set l_dbDepartment = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllDepartments", , , , , True
End Function

Public Function GetAllJobs() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetAllJobs_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
'PREM BEGIN TRIS CONVERSION PHASE 3 (.NET PHASE 1)
    l_strSQL = "SELECT  VARCHAR(LTRIM(RTRIM(JOB_NUM)), 6) AS JOB_NUM, JOB_DESC" & _
            "           , VARCHAR(LTRIM(RTRIM(CRAFT_NUM)), 3) AS CRAFT_NUM, OT_TYPE" & _
            "           , COALESCE(EFFDT, '01/01/0100') AS EFFDT" & _
            "           , EFF_STATUS, VARCHAR(LTRIM(RTRIM(UNION_CD)), 3) AS UNION_CD" & _
            "           , LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP" & _
            " FROM      METHDBA.TTR_JOBS" & _
            " ORDER BY JOB_NUM "
'PREM END TRIS CONVERSION PHASE 3 (.NET PHASE 1)
'//TNV
    '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
'    l_strSQL = "SELECT CAST(LTRIM(RTRIM(JOB_NUM)) AS VARCHAR(6)) AS JOB_NUM, JOB_DESC, CRAFT_NUM, OT_TYPE, " _
'        & "VALUE(NULLIF(EFFDT,'0001-01-01'),DATE('0100-01-01')) as EFFDT, " _
'        & "EFF_STATUS, CAST(LTRIM(RTRIM(UNION_CD)) AS VARCHAR(3)) AS UNION_CD " _
'        & "FROM METHDBA.TTR_JOBS " _
'        & "ORDER BY JOB_NUM "
'//TNV

'    l_strSQL = "SELECT JOB_NUM, JOB_DESC, CRAFT_NUM, OT_TYPE, " _
'        & "VALUE(NULLIF(EFFDT,'0001-01-01'),DATE('0100-01-01')) as EFFDT, " _
'        & "EFF_STATUS, UNION_CD " _
'        & "FROM METHDBA.TTR_JOBS " _
'        & "ORDER BY JOB_NUM "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetAllJobs = l_objRS

    Set l_objRS = Nothing
    Exit Function

GetAllJobs_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllJobs", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetJob(ByVal v_strJobNum As String) As ADODB.Recordset
    
On Error GoTo eh
    Set GetJob = GetJobBasic(v_strJobNum)
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetJob", , , , , True
End Function

Private Function GetJobBasic(ByVal v_strJobNum As String) As ADODB.Recordset
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
    l_objCN.Open

    l_strSQL = "SELECT JOB_NUM, JOB_DESC, CRAFT_NUM, OT_TYPE, " _
        & "VALUE(NULLIF(EFFDT,'0001-01-01'),DATE('0100-01-01')) as EFFDT, " _
        & "EFF_STATUS, UNION_CD " _
        & "FROM METHDBA.TTR_JOBS " _
        & "WHERE JOB_NUM = '" & Trim(v_strJobNum) & "' "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetJobBasic = l_objRS
    Set l_objRS = Nothing
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetJobBasic", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetUnions(r_colCriteria As Collection) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open
    l_strSQL = "SELECT UNION_CD, UNION_LOCAL, UNION_DESC, " _
             & "EXEC_CONTRACT_IND, UNION_BEN_PKG_IND " _
             & "FROM METHDBA.TTR_UNIONS " _
             & "WHERE UNION_CD IN " _
             & AsDelimitedString(r_colCriteria) _
             & "ORDER BY UNION_CD "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetUnions = l_objRS
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetUnions", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetAllUnions() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open

'//TNV
    '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
    l_strSQL = "SELECT CAST(LTRIM(RTRIM(UNION_CD)) AS VARCHAR(3)) AS UNION_CD, UNION_LOCAL, UNION_DESC, " _
             & "EXEC_CONTRACT_IND, UNION_BEN_PKG_IND " _
             & "FROM METHDBA.TTR_UNIONS " _
             & "ORDER BY UNION_CD "
'//TNV
    
'    l_strSQL = "SELECT UNION_CD, UNION_LOCAL, UNION_DESC, " _
'             & "EXEC_CONTRACT_IND, UNION_BEN_PKG_IND " _
'             & "FROM METHDBA.TTR_UNIONS " _
'             & "ORDER BY UNION_CD "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetAllUnions = l_objRS
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllUnions", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetAllEarningCodes() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetAllEarningCodes_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "select EARNINGS_CD, EARNINGS_CD_DESC, EXEC_CONTRACT_IND, " _
             & " FACTOR, PAY_DISTR_TYPE " _
             & " from METHDBA.TTR_EARN_CODES " _
             & " order by EXEC_CONTRACT_IND, EARNINGS_CD "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetAllEarningCodes = Nothing
    Else
        Set GetAllEarningCodes = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetAllEarningCodes_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllEarningCodes", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetTravelAllowances() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetTravelAllowances_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    l_objCN.Open
    l_strSQL = "SELECT USER_TYPE, EFF_YEAR, START_LOC_CD, END_LOC_CD, " _
             & " ALLOWANCE_AMT, DISTRIBUTION_ID, LAST_MOD_EMP_NUM " _
             & " FROM METHDBA.TTR_TRAV_ALLOWANCE "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetTravelAllowances = Nothing
    Else
        Set GetTravelAllowances = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetTravelAllowances_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTravelAllowances", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetTravelTripLocations() As ADODB.Recordset
' RED104
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetTravelTripLocations_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT DISTINCT B.LOC_CD, B.LOC_DESC "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_TRAV_ALLOWANCE A, METHDBA.TTR_LOCATIONS B "
    l_strSQL = l_strSQL & "WHERE (B.LOC_CD = A.START_LOC_CD "
    l_strSQL = l_strSQL & "OR B.LOC_CD = A.END_LOC_CD) "
    l_strSQL = l_strSQL & "ORDER BY B.LOC_DESC "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    If l_objRS.BOF And l_objRS.EOF Then
        Set GetTravelTripLocations = Nothing
    Else
        Set GetTravelTripLocations = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetTravelTripLocations_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTravelTripLocations", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetOvertimeExceptions() As ADODB.Recordset
'RED051 and RED062
'This function returns overtime exception codes for
'contract and executive employees
'The exec_contract_ind ('C' or 'E') can be filter
'to get a list specfic to contract or executive employees
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetOvertimeExceptions_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    l_objCN.Open
    l_strSQL = "Select exception_cd , EXCEPTION_DESC, " & _
                    " factor, EXEC_CONTRACT_IND, " & _
                    " EXEC_EARN_CD As EARNINGS_CD, 'N' as REFUSED_IND " & _
                "from methdba.ttr_exception_cds, " & _
                "     methdba.TTR_EARN_CODES " & _
                "Where EARNINGS_CD = EXEC_EARN_CD " & _
                "  and EXEC_CONTRACT_IND = 'E' " & _
                "  and otherpay_ind in ('T','B') " & _
                "  and EXCEPTION_BANK_CD = 'O' " & _
                "  and EARNINGS_CD <> '99' " & _
                "  and EXCEPTION_CD not in " & _
                " (select RG_VALUE_CD " & _
                "   from METHDBA.TTR_RG_GROUP " & _
                "   where RG_CD = 6)  "
    l_strSQL = l_strSQL & " UNION  " & _
                "Select exception_cd , EXCEPTION_DESC, " & _
                    " factor, EXEC_CONTRACT_IND, " & _
                    " CONTRACT_EARN_CD As EARNINGS_CD, 'N' as REFUSED_IND " & _
                "from methdba.ttr_exception_cds, " & _
                "     methdba.TTR_EARN_CODES " & _
                "Where EARNINGS_CD = CONTRACT_EARN_CD " & _
                "  and EXEC_CONTRACT_IND = 'C' " & _
                "  and otherpay_ind in ('T','B') " & _
                "  and EXCEPTION_BANK_CD = 'O' " & _
                 "  and EARNINGS_CD <> '99' "
    l_strSQL = l_strSQL & "UNION " & _
                "Select b.exception_cd , EXCEPTION_DESC, " & _
                "DECIMAL(factor/100,5,4) as factor " & _
                ",'C','99', LEFT(RG_VALUE_CD,1) as REFUSED_IND " & _
                "from methdba.ttr_exception_cds a, " & _
                    " methdba.ttr_rg_bank_exc_cd b " & _
                "Where b.rg_cd = 107 " & _
                "  and a.exception_cd = b.exception_cd "
    l_strSQL = l_strSQL & " ORDER BY 4, 1  "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    If l_objRS.BOF And l_objRS.EOF Then
        Set GetOvertimeExceptions = Nothing
    Else
        Set GetOvertimeExceptions = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetOvertimeExceptions_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeExceptions", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetHolidaysBetween( _
        ByVal v_datStartDate As Date _
        , ByVal v_datEndDate As Date _
        , Optional ByVal v_colUnionArgs As Collection _
        ) As ADODB.Recordset
On Error GoTo GetHolidaysBetween_EH
' WHITE048
Dim l_objCN As ADODB.Connection
Dim l_objRS As ADODB.Recordset
Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
Dim l_strSQL As String, l_blnUnionsProvided As Boolean

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset

    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing

    l_objCN.Open

    If Not v_colUnionArgs Is Nothing Then
        If v_colUnionArgs.Count > 0 Then
            l_blnUnionsProvided = True
        End If
    End If

    If Not l_blnUnionsProvided Then
        l_strSQL = "SELECT  UNION_CD, HOLIDAY_DATE" & _
            "       , HOLIDAY_DESC" & _
            " FROM  METHDBA.TTR_HOLIDAYS" & _
            " WHERE HOLIDAY_DATE BETWEEN" & _
            " '" & Format(v_datStartDate, "mm/dd/yyyy") & "'" & _
            " AND " & _
            " '" & Format(v_datEndDate, "mm/dd/yyyy") & "'" & _
            " ORDER BY UNION_CD, HOLIDAY_DATE "
    Else
        l_strSQL = "SELECT  UNION_CD, HOLIDAY_DATE" & _
            "       , HOLIDAY_DESC" & _
            " FROM METHDBA.TTR_HOLIDAYS" & _
            " WHERE HOLIDAY_DATE BETWEEN" & _
            " '" & Format(v_datStartDate, "mm/dd/yyyy") & "'" & _
            " AND" & _
            " '" & Format(v_datEndDate, "mm/dd/yyyy") & "'" & _
            " AND UNION_CD IN" & AsDelimitedString(v_colUnionArgs) & _
            " ORDER BY UNION_CD, HOLIDAY_DATE"
    End If

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    l_objRS.ActiveConnection = Nothing

    l_objCN.Close
    Set l_objCN = Nothing

    If l_objRS.BOF And l_objRS.EOF Then
        Set GetHolidaysBetween = Nothing
    Else
        Set GetHolidaysBetween = l_objRS
    End If

    Set l_objRS = Nothing
    Exit Function
GetHolidaysBetween_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing

    TrsRaiseError m_strCMODULENAME, "GetHolidaysBetween", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetAllHolidays() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetAllHolidays_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    l_objCN.Open

'//TNV
    '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
    l_strSQL = "select HOLIDAY_DATE, HOLIDAY_DESC, CAST(LTRIM(RTRIM(UNION_CD)) AS VARCHAR(3)) AS UNION_CD " _
             & " from METHDBA.TTR_HOLIDAYS " _
             & " where HOLIDAY_DATE >= " & "'" & Format(Now, "mm/dd/yyyy") & "'" _
             & " order by HOLIDAY_DATE, UNION_CD "
             
'//TNV
'    l_strSQL = "select HOLIDAY_DATE, HOLIDAY_DESC, UNION_CD " _
'             & " from METHDBA.TTR_HOLIDAYS " _
'             & " where HOLIDAY_DATE >= " & "'" & Format(Now, "mm/dd/yyyy") & "'" _
'             & " order by HOLIDAY_DATE, UNION_CD "

    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    If l_objRS.BOF And l_objRS.EOF Then
        Set GetAllHolidays = Nothing
    Else
        Set GetAllHolidays = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetAllHolidays_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllHolidays", , , , DescArgsFrom(l_strSQL), True
End Function

'//-------------------------------------------------------//
'//-----Retrieve methods for use with utilRecordCache-----//
'//-------------------------------------------------------//
Public Function GetExceptionRecords(Optional ByVal v_varExcCodes As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String

On Error GoTo eh
    
'    If IsMissing(v_varExcCodes) Or IsEmpty(v_varExcCodes) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varExcCodes) Then
'        l_strWhereClause = " WHERE " & BuildInClause("EXC_CD", v_varExcCodes, True)
'    Else
'        l_strWhereClause = " WHERE EXCEPTION_CD = " & "'" & v_varExcCodes & "'"
'    End If

    l_strWhereClause = BuildWhereClause("WHERE", "EXCEPTION_CD", v_varExcCodes, True)
    
'   '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
'    l_strSQL = "select CAST(LTRIM(RTRIM(EXCEPTION_CD)) AS VARCHAR(3)) AS EXCEPTION_CD, " _
'             & " EXCEPTION_DESC, EXCEPTION_TYPE, " _
'             & " CONTRACT_EARN_CD, EXEC_EARN_CD, SYMBOL_TIME_IND, " _
'             & " EXCEPTION_BANK_CD, INELIGIBILITY_IND, OTHERPAY_IND, " _
'             & " CASE_GROUP, SCHED_IND, 'N' as FMLA_IND, " _
'             & " WORK_REST_IND, SCHED_CHG_IND, ACCT_NUM, SYMBOL_CD " _
'             & " , AIC_ACCT_NUM, AMS_ACCT_NUM, UEC_ACCT_NUM " _
'             & " from METHDBA.TTR_EXCEPTION_CDS " _
'             & l_strWhereClause

    If IsPOETActive And IsOTTCEffective Then
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetExceptionRecordsOTTC.sql")
    Else
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetExceptionRecords.sql")
    End If
    
    l_strSQL = Replace(l_strSQL, "{WhereClause}", l_strWhereClause)
    '//Set l_objRS = oADOHelper.ExecuteSelect(l_strSQL, adOpenStatic, adLockBatchOptimistic)

    Set l_objRS = GetCacheRecords(l_strSQL)

    Set GetExceptionRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetExceptionRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'//copied
Private Function IsPOETActive() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim UseOTTC As String
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsPOETActive = False
    
    UseOTTC = UCase(dbEnv.GetSysParmsValue(50, 1))
    
    If UseOTTC = "Y" Or UseOTTC = "TRUE" Or UseOTTC = "1" Then
        IsPOETActive = True
    End If
        
    Set dbEnv = Nothing

    Exit Function

eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsPOETActive", , , , DescArgsFrom(UseOTTC), True
End Function

'//actual date to start using POET distributions after cutover
Private Function IsOTTCEffective() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim sEffDate As String
    Dim EffDateTime
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsOTTCEffective = False
    
    sEffDate = dbEnv.GetSysParmsValue(53, 1)
    
    EffDateTime = DateTimeFromDB2Timestamp(sEffDate)
    
    IsOTTCEffective = Now > EffDateTime
        
    Set dbEnv = Nothing

    Exit Function
eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsOTTCEffective", , , , DescArgsFrom(sEffDate), True
End Function

'//-----previously GetEffectiveUnionComponents
Public Function GetUnionComponentRecords(Optional ByVal v_varKeys As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    Dim l_colDescArgs As Collection
On Error GoTo eh
    
    If IsMissing(v_varKeys) Or IsEmpty(v_varKeys) Then
        l_strWhereClause = ""
    ElseIf IsArray(v_varKeys) Then
        If UBound(v_varKeys) - LBound(v_varKeys) = 2 Then
            l_strWhereClause = BuildWhereClauseForMultiKey(Array("AND", "AND", "AND"), _
                                                    Array("UNION_CD", "A.COMP_TYPE", "A.COMP_CD"), _
                                                    v_varKeys, _
                                                    Array(True, True, True))
        Else
            Err.Raise trErrdbGetCacheSQLKey
        End If
    Else
        Err.Raise trErrdbGetCacheSQLKey
    End If
    
    '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
    l_strSQL = "SELECT DISTINCT CAST(LTRIM(RTRIM(A.UNION_CD)) AS VARCHAR(3)) AS UNION_CD, " _
             & " CAST(LTRIM(RTRIM(A.COMP_CD)) AS VARCHAR(3)) AS COMP_CD," _
             & " B.COMP_DESC, A.EFF_DATE, A.COMP_TYPE, A.COMP_AMT " _
             & "FROM METHDBA.TTR_UNION_COMP_AMT A, METHDBA.TTR_COMP_CD B " _
             & "WHERE A.COMP_CD = B.COMP_CD " _
             & "AND A.COMP_TYPE = B.COMP_TYPE " _
             & "AND B.COMP_CD not in " _
             & " (select RG_VALUE_CD " _
             & "   from METHDBA.TTR_RG_GROUP " _
             & "   where RG_CD = 85) " _
             & l_strWhereClause _
             & "ORDER BY UNION_CD, COMP_TYPE, COMP_CD, A.EFF_DATE DESC"

    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetUnionComponentRecords = l_objRS

    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    Select Case Err.Number
    Case trErrdbGetCacheSQLKey
        Set l_colDescArgs = New Collection
    
        AddToErrorCollection l_colDescArgs, v_varKeys, "v_varKeys"
        TrsRaiseError m_strCMODULENAME, "GetUnionComponentRecords", Err.Number, , l_colDescArgs, , True
    Case Else
    TrsRaiseError m_strCMODULENAME, "GetUnionComponentRecords", , , , DescArgsFrom(l_strSQL), True
    End Select
    

End Function

'//-----previously GetJobBasic
Private Function GetJobRecords(Optional ByVal v_varJobNum As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    
On Error GoTo eh
    
'    If IsMissing(v_varJobNum) Or IsEmpty(v_varJobNum) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varJobNum) Then
'        l_strWhereClause = " WHERE " & BuildInClause("JOB_NUM", v_varJobNum, True)
'    Else
'        l_strWhereClause = " WHERE JOB_NUM = '" & Trim(v_varJobNum) & "' "
'    End If
    
    l_strWhereClause = BuildWhereClause("WHERE", "JOB_NUM", v_varJobNum, True)
    
    l_strSQL = "SELECT JOB_NUM, JOB_DESC, CRAFT_NUM, OT_TYPE, " _
        & "VALUE(NULLIF(EFFDT,'0001-01-01'),DATE('0100-01-01')) as EFFDT, " _
        & "EFF_STATUS, UNION_CD " _
        & "FROM METHDBA.TTR_JOBS " _
        & l_strWhereClause
    
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetJobRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function

eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetJobRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'//-----previously from GetExceptionGroupCacheSQL
Private Function GetExceptionGroupsRecords(Optional ByVal v_varKeys As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    Dim l_colDescArgs As Collection
On Error GoTo eh
    
    If IsMissing(v_varKeys) Or IsEmpty(v_varKeys) Then
        l_strWhereClause = ""
    ElseIf IsArray(v_varKeys) Then
        If UBound(v_varKeys) - LBound(v_varKeys) = 2 Then
            l_strWhereClause = BuildWhereClauseForMultiKey(Array("WHERE", "AND", "AND"), _
                                                    Array("EG_CD", "EXCEPTION_CD", "EG_SUB_CD"), _
                                                    v_varKeys, _
                                                    Array(False, True, True))
        Else
            Err.Raise trErrdbGetCacheSQLKey
        End If
    Else
        Err.Raise trErrdbGetCacheSQLKey
    End If
    
    l_strSQL = "SELECT EG_CD,EG_SUB_CD,EXCEPTION_CD,EXC_SEQ_NUM " _
         & " FROM METHDBA.TTR_EXC_GROUP "
    l_strSQL = l_strSQL & l_strWhereClause
    l_strSQL = l_strSQL & " ORDER BY EG_CD,EG_SUB_CD,EXCEPTION_CD "
    
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetExceptionGroupsRecords = l_objRS
        
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    
    Select Case Err.Number
    Case trErrdbGetCacheSQLKey
        Set l_colDescArgs = New Collection
    
        AddToErrorCollection l_colDescArgs, v_varKeys, "v_varKeys"
        TrsRaiseError m_strCMODULENAME, "GetExceptionGroupsRecords", Err.Number, , l_colDescArgs, , True
    Case Else
        TrsRaiseError m_strCMODULENAME, "GetExceptionGroupsRecords", , , , DescArgsFrom(l_strSQL), True
    End Select
    
End Function

'same as GetAllLocations
Private Function GetLocationRecords(Optional ByVal v_varLocCd As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    
On Error GoTo eh
    
'    If IsMissing(v_varLocCd) Or IsEmpty(v_varLocCd) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varLocCd) Then
'        l_strWhereClause = " WHERE " & BuildInClause("LOC_CD", v_varLocCd, False)
'    Else
'        l_strWhereClause = " WHERE LOC_CD = " & Trim(v_varLocCd)
'    End If
    
    l_strWhereClause = BuildWhereClause("WHERE", "LOC_CD", v_varLocCd, False)
    
    l_strSQL = "select LOC_CD, LOC_DESC, PLANT_CD " _
             & "from METHDBA.TTR_LOCATIONS " _
            & l_strWhereClause
    
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetLocationRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function

eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetLocationRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'same as GetAllUnions
Private Function GetUnionRecords(Optional ByVal v_varUnionCd As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    
On Error GoTo eh
    
'    If IsMissing(v_varUnionCd) Or IsEmpty(v_varUnionCd) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varUnionCd) Then
'        l_strWhereClause = " WHERE " & BuildInClause("UNION_CD", v_varUnionCd, True)
'    Else
'        l_strWhereClause = " WHERE UNION_CD = '" & Trim(v_varUnionCd) & "' "
'    End If
    
    l_strWhereClause = BuildWhereClause("WHERE", "UNION_CD", v_varUnionCd, True)
    
    l_strSQL = "SELECT CAST(LTRIM(RTRIM(UNION_CD)) AS VARCHAR(3)) AS UNION_CD, UNION_LOCAL, UNION_DESC, " _
                & "EXEC_CONTRACT_IND, UNION_BEN_PKG_IND " _
                & " FROM METHDBA.TTR_UNIONS " _
                & l_strWhereClause _
                & " ORDER BY UNION_CD "
    
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetUnionRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetUnionRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'same as GetAllOvertimeReasons
'not utilizing GetOTReasonCacheSQL
Private Function GetOvertimeReasonRecords(Optional ByVal v_varOTReasonCd As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    
On Error GoTo eh
    
'    If IsMissing(v_varOTReasonCd) Or IsEmpty(v_varOTReasonCd) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varOTReasonCd) Then
'        l_strWhereClause = " WHERE " & BuildInClause("OT_REASON_CD", v_varOTReasonCd, True)
'    Else
'        l_strWhereClause = " WHERE OT_REASON_CD = '" & Trim(v_varOTReasonCd) & "'"
'    End If
    
    l_strWhereClause = BuildWhereClause("WHERE", "OT_REASON_CD", v_varOTReasonCd, True)
    
    l_strSQL = "select CAST(LTRIM(RTRIM(OT_REASON_CD)) AS VARCHAR(3)) AS OT_REASON_CD, OT_REASON_DESC " _
                & "from METHDBA.TTR_OT_REASONS " _
                & l_strWhereClause & _
                " ORDER BY OT_REASON_CD"
                
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetOvertimeReasonRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeReasonRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'same as GetAllOvertimeResponses
'not utilizing GetOTResponseCacheSQL
Private Function GetOvertimeResponseRecords(Optional ByVal v_varOTRespCd As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    
On Error GoTo eh
    
'    If IsMissing(v_varOTRespCd) Or IsEmpty(v_varOTRespCd) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varOTRespCd) Then
'        l_strWhereClause = " WHERE " & BuildInClause("OT_RESP_CD", v_varOTRespCd, True)
'    Else
'        l_strWhereClause = " WHERE OT_RESP_CD = '" & Trim(v_varOTRespCd) & "'"
'    End If
    
    l_strWhereClause = BuildWhereClause("WHERE", "OT_RESP_CD", v_varOTRespCd, True)
'BEGIN PREM TRIS CONVERSION PHASE 3
    l_strSQL = "SELECT  OT_RESP_CD , RESPONSE_DESC " & _
                "       , OT_CHRG_IND, OT_WORKED_IND " & _
                "       , OT_CONTACTED_IND " & _
                " FROM  METHDBA.TTR_OT_RESPONSES " _
                & l_strWhereClause & _
                " ORDER BY OT_RESP_CD"
'END PREM TRIS CONVERSION PHASE 3
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetOvertimeResponseRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeResponseRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'same as GetOvertimeExceptions
'previously GetOTExceptionCacheSQL (not utilized for optimization)
Private Function GetOvertimeExceptionRecords(Optional ByVal v_varExcCd As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    
On Error GoTo eh
    
'    If IsMissing(v_varExcCd) Or IsEmpty(v_varExcCd) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varExcCd) Then
'        l_strWhereClause = " AND " & BuildInClause("EXCEPTION_CD", v_varExcCd, True)
'    Else
'        l_strWhereClause = " AND EXCEPTION_CD = '" & Trim(v_varExcCd) & "'"
'    End If
    
    l_strWhereClause = BuildWhereClause("AND", "a.EXCEPTION_CD", v_varExcCd, True)
    
    l_strSQL = "Select exception_cd , EXCEPTION_DESC, " & _
                    " factor, EXEC_CONTRACT_IND, " & _
                    " EXEC_EARN_CD As EARNINGS_CD " & _
                "from methdba.ttr_exception_cds a, " & _
                "     methdba.TTR_EARN_CODES " & _
                "Where EARNINGS_CD = EXEC_EARN_CD " & _
                "  and EXEC_CONTRACT_IND = 'E' " & _
                "  and otherpay_ind in ('T','B') " & _
                "  and EXCEPTION_BANK_CD = 'O' " & _
                "  and EARNINGS_CD <> '99' " & _
                l_strWhereClause
    l_strSQL = l_strSQL & " UNION  " & _
                "Select exception_cd , EXCEPTION_DESC, " & _
                    " factor, EXEC_CONTRACT_IND, " & _
                    " CONTRACT_EARN_CD As EARNINGS_CD " & _
                "from methdba.ttr_exception_cds a, " & _
                "     methdba.TTR_EARN_CODES " & _
                "Where EARNINGS_CD = CONTRACT_EARN_CD " & _
                "  and EXEC_CONTRACT_IND = 'C' " & _
                "  and otherpay_ind in ('T','B') " & _
                "  and EXCEPTION_BANK_CD = 'O' " & _
                "  and EARNINGS_CD <> '99' " & _
                l_strWhereClause
    l_strSQL = l_strSQL & "UNION " & _
                "Select b.exception_cd , EXCEPTION_DESC, " & _
                "DECIMAL(factor/100,5,4) as factor " & _
                ",'C','99' " & _
                "from methdba.ttr_exception_cds a, " & _
                    " methdba.ttr_rg_bank_exc_cd b " & _
                "Where b.rg_cd = 107 " & _
                "  and a.exception_cd = b.exception_cd " & _
                l_strWhereClause
    l_strSQL = l_strSQL & " ORDER BY 4, 1  "
    
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetOvertimeExceptionRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeExceptionRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'previously GetHistoryCutoffDate
Private Function GetHSCutOffDateRecords(Optional ByVal v_varCutoffKey As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    
On Error GoTo eh
    
'    If IsMissing(v_varCutoffKey) Or IsEmpty(v_varCutoffKey) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varCutoffKey) Then
'        l_strWhereClause = " AND " & BuildInClause("HS_CUTOFF_KEY", v_varCutoffKey, True)
'    Else
'        l_strWhereClause = " AND HS_CUTOFF_KEY = '" & Trim(v_varCutoffKey) & "'"
'    End If
    
    l_strWhereClause = BuildWhereClause("WHERE", "HS_CUTOFF_KEY", v_varCutoffKey, True)
    
    l_strSQL = "SELECT *" _
                & " From" _
                & " (" _
    & " SELECT '" & CKEYCUTOFFDATE & "' AS HS_CUTOFF_KEY, PARM_VALUE AS HS_CUTOFF_DATE1, " _
                & " Cast('0100-01-01' as Date) AS HS_CUTOFF_DATE" _
                & " FROM METHDBA.TTR_SYSPARMS " _
                & " WHERE PARM_NUM = 18 AND PARM_SUFFIX = 1" _
                & " ) NTE0" _
                & l_strWhereClause
 
    Set l_objRS = GetCacheRecords(l_strSQL, adLockBatchOptimistic) '//-----so that this can be updated later
       
    '//-----Type Cast the char column so that the cache will return a date in the proper format
    If RecordsExist(l_objRS) Then
        l_objRS.MoveFirst
        Do While Not l_objRS.EOF
            If IsDate(Trim(l_objRS!HS_CUTOFF_DATE1)) Then
                l_objRS!HS_CUTOFF_DATE = CDate(Trim(l_objRS!HS_CUTOFF_DATE1))
                l_objRS.Update
            End If
            l_objRS.MoveNext
        Loop
    End If
    
    Set GetHSCutOffDateRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetHSCutOffDateRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'previously GetAvailScheduleCacheSQL
Private Function GetScheduleRecords(Optional ByVal v_varSchedCd As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    
On Error GoTo eh
    
'    If IsMissing(v_varSchedCd) Or IsEmpty(v_varSchedCd) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varSchedCd) Then
'        l_strWhereClause = " WHERE " & BuildInClause("SCHED_CD", v_varSchedCd, False)
'    Else
'        l_strWhereClause = " WHERE SCHED_CD = " & Trim(v_varSchedCd)
'    End If
    
    l_strWhereClause = BuildWhereClause("WHERE", "SCHED_CD", v_varSchedCd, False)
    
    l_strSQL = "SELECT SCHED_CD" & _
               " FROM METHDBA.TTR_AVAIL_SCHED" & _
                l_strWhereClause & _
                " ORDER BY SCHED_CD"

    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetScheduleRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduleRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'previously GetRotationCacheSQL
Private Function GetRotationRecords(Optional ByVal v_varRotCd As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    
On Error GoTo eh
    
'    If IsMissing(v_varRotCd) Or IsEmpty(v_varRotCd) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varRotCd) Then
'        l_strWhereClause = " WHERE " & BuildInClause("ROT_CD", v_varRotCd, False)
'    Else
'        l_strWhereClause = " WHERE ROT_CD = " & Trim(v_varRotCd)
'    End If
    
    l_strWhereClause = BuildWhereClause("WHERE", "ROT_CD", v_varRotCd, False)
    
    l_strSQL = "SELECT ROT_CD, MAX(ROT_WEEK) AS ROT_WEEKS" & _
                " FROM METHDBA.TTR_ROTATION_SCHED" & _
                l_strWhereClause & _
                " GROUP BY ROT_CD"
    
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetRotationRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetRotationRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'previously GetComponentCodeCacheSQL
Private Function GetComponentRecords(Optional ByVal v_varCompCd As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    
On Error GoTo eh
    
'    If IsMissing(v_varCompCd) Or IsEmpty(v_varCompCd) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varCompCd) Then
'        l_strWhereClause = " WHERE " & BuildInClause("COMP_CD", v_varCompCd, True)
'    Else
'        l_strWhereClause = " WHERE COMP_CD = '" & Trim(v_varCompCd) & "'"
'    End If
    
    l_strWhereClause = BuildWhereClause("WHERE", "COMP_CD", v_varCompCd, True)
    
    l_strSQL = "SELECT COMP_CD, COMP_TYPE " & _
                " FROM METHDBA.TTR_COMP_CD " & _
                l_strWhereClause
    If Trim(l_strWhereClause) = "" Then
        l_strSQL = l_strSQL _
             & " WHERE COMP_CD not in "
    Else
        l_strSQL = l_strSQL _
             & " AND COMP_CD not in "
    End If
    l_strSQL = l_strSQL _
             & " (select RG_VALUE_CD " _
             & "   from METHDBA.TTR_RG_GROUP " _
             & "   where RG_CD = 85) " _
             & " ORDER BY COMP_CD"
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetComponentRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetComponentRecords", , , , DescArgsFrom(l_strSQL), True
End Function

Private Function BuildWhereClause(ByVal v_strSQLKeyWord As String, _
                                    ByVal v_strColumn As String, _
                                    ByVal v_varColumnValues As Variant, _
                                    ByVal v_blnDelimit As Boolean) As String
    Dim l_strWhereClause As String
    Dim l_colErrArgs As Collection
On Error GoTo eh
    
    If IsMissing(v_varColumnValues) Or IsEmpty(v_varColumnValues) Then
        l_strWhereClause = ""
    ElseIf IsArray(v_varColumnValues) Then
        l_strWhereClause = " " & v_strSQLKeyWord & " " & BuildInClause(v_strColumn, v_varColumnValues, v_blnDelimit)
    Else
        l_strWhereClause = " " & v_strSQLKeyWord & " " & v_strColumn & " = " & _
            IIf(v_blnDelimit, StringEnclosedBy(Trim(CStr(v_varColumnValues)), "'"), Trim(CStr(v_varColumnValues)))
    End If

    BuildWhereClause = l_strWhereClause
    Exit Function
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strSQLKeyWord = " & v_strSQLKeyWord
    l_colErrArgs.Add "v_strColumn = " & v_strColumn
    If IsMissing(v_varColumnValues) Or IsEmpty(v_varColumnValues) Then
        l_colErrArgs.Add "v_varColumnValues = empty"
    ElseIf IsArray(v_varColumnValues) Then
        l_colErrArgs.Add "v_varColumnValues() = (" & Join(v_varColumnValues, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varColumnValues = " & v_varColumnValues
    End If
    l_colErrArgs.Add "v_blnDelimit = " & v_blnDelimit
    
    TrsRaiseError m_strCMODULENAME, "BuildWhereClause", , , , l_colErrArgs, True
End Function

Private Function BuildWhereClauseForMultiKey(ByVal v_strSQLKeyWord As Variant, _
                                    ByVal v_strColumn As Variant, _
                                    ByVal v_varColumnValues As Variant, _
                                    ByVal v_blnDelimit As Variant) As Variant
    Dim l_strWhereClause As String
    Dim l_colErrArgs As Collection
    Dim nKey As Long
On Error GoTo eh
    
    If IsMissing(v_varColumnValues) Or IsEmpty(v_varColumnValues) Then
        l_strWhereClause = ""
    Else
        If UBound(v_strSQLKeyWord) = UBound(v_strSQLKeyWord) And _
            UBound(v_strSQLKeyWord) = UBound(v_strColumn) And _
            UBound(v_strSQLKeyWord) = UBound(v_varColumnValues) And _
            UBound(v_strSQLKeyWord) = UBound(v_blnDelimit) Then
            l_strWhereClause = ""
            For nKey = LBound(v_strSQLKeyWord) To UBound(v_strSQLKeyWord)
                l_strWhereClause = l_strWhereClause & BuildWhereClause(v_strSQLKeyWord(nKey), v_strColumn(nKey), v_varColumnValues(nKey), v_blnDelimit(nKey)) & " "
            Next nKey
        Else
        End If
    End If
    
    BuildWhereClauseForMultiKey = l_strWhereClause
    Exit Function
eh:
    Set l_colErrArgs = New Collection
    
    AddToErrorCollection l_colErrArgs, v_strSQLKeyWord, "v_strSQLKeyWord"
    AddToErrorCollection l_colErrArgs, v_strColumn, "v_strColumn"
    AddToErrorCollection l_colErrArgs, v_varColumnValues, "v_varColumnValues"
    AddToErrorCollection l_colErrArgs, v_blnDelimit, "v_blnDelimit"
         
    TrsRaiseError m_strCMODULENAME, "BuildWhereClauseForMultiKey", , , , l_colErrArgs, True
End Function

Private Function AddToErrorCollection(ByRef r_colErrArgs As Collection, _
                                        ByRef v_varArgs As Variant, _
                                        ByVal v_strArgName As String)
    If IsMissing(v_varArgs) Or IsEmpty(v_varArgs) Then
        r_colErrArgs.Add v_strArgName & " = empty"
    ElseIf IsArray(v_varArgs) Then
        r_colErrArgs.Add v_strArgName & "() = (" & Join(v_varArgs, ", ") & ")"
    Else
        r_colErrArgs.Add v_strArgName & " = " & v_varArgs
    End If
End Function

Public Function GetAllExceptions() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetAllExceptions_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

   '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
'    l_strSQL = "select CAST(LTRIM(RTRIM(EXCEPTION_CD)) AS VARCHAR(3)) AS EXCEPTION_CD, " _
'             & " EXCEPTION_DESC, EXCEPTION_TYPE, " _
'             & " CONTRACT_EARN_CD, EXEC_EARN_CD, SYMBOL_TIME_IND, " _
'             & " EXCEPTION_BANK_CD, INELIGIBILITY_IND, OTHERPAY_IND, " _
'             & " CASE_GROUP, SCHED_IND, 'N' as FMLA_IND, " _
'             & " WORK_REST_IND, SCHED_CHG_IND, ACCT_NUM, SYMBOL_CD " _
'             & " , AIC_ACCT_NUM, AMS_ACCT_NUM, UEC_ACCT_NUM " _
'             & " from METHDBA.TTR_EXCEPTION_CDS " _
'             & " where EXCEPTION_CD not in " _
'             & " (select RG_VALUE_CD " _
'             & "   from METHDBA.TTR_RG_GROUP " _
'             & "   where RG_CD = 6)"

'    l_strSQL = "select EXCEPTION_CD, EXCEPTION_DESC, EXCEPTION_TYPE, " _
'             & " CONTRACT_EARN_CD, EXEC_EARN_CD, SYMBOL_TIME_IND, " _
'             & " EXCEPTION_BANK_CD, INELIGIBILITY_IND, OTHERPAY_IND, " _
'             & " CASE_GROUP, SCHED_IND, FMLA_IND, " _
'             & " WORK_REST_IND, SCHED_CHG_IND, ACCT_NUM " _
'             & " from METHDBA.TTR_EXCEPTION_CDS "
    
    If IsPOETActive And IsOTTCEffective Then
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetAllExceptionsOTTC.sql")
    Else
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetAllExceptions.sql")
    End If
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    If l_objRS.BOF And l_objRS.EOF Then
        Set GetAllExceptions = Nothing
    Else
        Set GetAllExceptions = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetAllExceptions_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllExceptions", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetAllLocations() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetAllLocations_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "select LOC_CD, LOC_DESC, PLANT_CD " _
             & "from METHDBA.TTR_LOCATIONS " _
             & "ORDER BY LOC_CD "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetAllLocations = Nothing
    Else
        Set GetAllLocations = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetAllLocations_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllLocations", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetAllOvertimeReasons() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetAllOvertimeReasons_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

'BEGIN PREM TRIS CONVERSION PHASE 3 (.NET PHASE 1)
'ADDED TIME_STAMP AND LAST_MOD_EMP_NUM TO THE SQL STATEMENT
    l_strSQL = "SELECT  VARCHAR(LTRIM(RTRIM(OT_REASON_CD)), 3) AS OT_REASON_CD" & _
                "       , OT_REASON_DESC, LAST_MOD_EMP_NUM" & _
                "       , CHAR(TIME_STAMP) AS TIME_STAMP" & _
                " FROM  METHDBA.TTR_OT_REASONS" & _
                " ORDER BY OT_REASON_CD"
'END PREM TRIS CONVERSION PHASE 3 (.NET PHASE 1)

'//TNV
    '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
'    l_strSQL = "select CAST(LTRIM(RTRIM(OT_REASON_CD)) AS VARCHAR(3)) AS OT_REASON_CD, OT_REASON_DESC " _
             & "from METHDBA.TTR_OT_REASONS "
'//TNV

'    l_strSQL = "select OT_REASON_CD, OT_REASON_DESC " _
'             & "from METHDBA.TTR_OT_REASONS "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetAllOvertimeReasons = Nothing
    Else
        Set GetAllOvertimeReasons = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetAllOvertimeReasons_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllOvertimeReasons", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetAllOvertimeResponses() As ADODB.Recordset
    
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
    l_objCN.Open

'BEGIN PREM TRIS CONVERSION PHASE 3 (.NET PHASE 1)
'ADDED OT_CONTACTED_IND TIME_STAMP AND LAST_MOD_EMP_NUM
'TO THE SQL STATEMENT
    l_strSQL = "SELECT  VARCHAR(LTRIM(RTRIM(OT_RESP_CD)), 5) AS OT_RESP_CD" & _
                "       , RESPONSE_DESC, OT_WORKED_IND, CASE(OT_CHRG_IND)" & _
                "       WHEN 'C' THEN 'Y'" & _
                "       WHEN 'N' THEN 'N'" & _
                "       ELSE ' ' END AS OT_CHRG_IND" & _
                "       , OT_CONTACTED_IND" & _
                "       , LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP" & _
                " FROM  METHDBA.TTR_OT_RESPONSES" & _
                " ORDER BY OT_RESP_CD"
'END PREM TRIS CONVERSION PHASE 3 (.NET PHASE 1)

    'l_strSQL = "select OT_RESP_CD , RESPONSE_DESC, OT_CHRG_IND, OT_WORKED_IND " & _
               " From methdba.ttr_ot_responses "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    
    Set GetAllOvertimeResponses = l_objRS
    Set l_objRS = Nothing
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllOvertimeResponses", , , , DescArgsFrom(l_strSQL), True
End Function


'*********************************************************
' Location getters
'*********************************************************
Public Function LocationRecord( _
            ByVal v_strLocCd As String) As Variant
' Public/Friend method.  This should only be called by a friend
' that has knowledge of array structure; i.e. cbsCodeTables.
' Answer an array or empty.
' Array will be a 2 dim zero based array with the field name as the first
' element and the value as the second element.  This will allow
' friend clients to discover the structure instead of hard coding it.
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim lvarRecordArray(1, CLC_LAST_SLOT) As Variant
    Dim i As Integer

On Error GoTo eh
    
    l_varRecord = GetRecord(Array(v_strLocCd), CLC_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    lvarRecordArray(0, CLC_LOC_CD) = "LOC_CD"
    lvarRecordArray(1, CLC_LOC_CD) = l_varRecord(CLC_LOC_CD)
    lvarRecordArray(0, CLC_LOC_DESC) = "LOC_DESC"
    lvarRecordArray(1, CLC_LOC_DESC) = l_varRecord(CLC_LOC_DESC)
    lvarRecordArray(0, CLC_PLANT_CD) = "PLANT_CD"
    lvarRecordArray(1, CLC_PLANT_CD) = l_varRecord(CLC_PLANT_CD)
    LocationRecord = lvarRecordArray
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "LocationRecord"
End Function

Public Function LocationCodeIsValid( _
            ByVal v_strLocCd As String) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    l_varRecord = GetRecord(Array(v_strLocCd), CLC_GROUPNAME, lblnExists)
    LocationCodeIsValid = Not IsEmpty(l_varRecord)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "LocationCodeIsValid"
End Function

Public Function LocationDescription( _
            ByVal v_strLocCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(Array(v_strLocCd), CLC_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    LocationDescription = l_varRecord(CLC_LOC_DESC)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "LocationDescription"
End Function

Public Function LocationPlantCode( _
            ByVal v_strLocCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
' utilSPMCache returns Empty when record is not found, or is
' logically removed.
            
On Error GoTo eh
    l_varRecord = GetRecord(Array(v_strLocCd), CLC_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    LocationPlantCode = l_varRecord(CLC_PLANT_CD)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "LocationPlantCode"
End Function

'*********************************************************
' Begin Job getters
'*********************************************************
Public Function JobRecord( _
            ByVal v_strJobNum As String) As Variant
' Public/Friend method.  This should only be called by a friend
' that has knowledge of array structure; i.e. cbsCodeTables.
' Answer an array or empty.
' Array will be a 2 dim zero based array with the field name as the first
' element and the value as the second element.  This will allow
' friend clients to discover the structure instead of hard coding it.
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim lvarRecordArray(1, CJB_LAST_SLOT) As Variant
    Dim i As Integer
            
On Error GoTo eh
    l_varRecord = GetRecord(v_strJobNum, CJB_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    lvarRecordArray(0, CJB_JOB_NUM) = "JOB_NUM"
    lvarRecordArray(1, CJB_JOB_NUM) = l_varRecord(CJB_JOB_NUM)
    lvarRecordArray(0, CJB_JOB_DESC) = "JOB_DESC"
    lvarRecordArray(1, CJB_JOB_DESC) = l_varRecord(CJB_JOB_DESC)
    lvarRecordArray(0, CJB_CRAFT_NUM) = "CRAFT_NUM"
    lvarRecordArray(1, CJB_CRAFT_NUM) = l_varRecord(CJB_CRAFT_NUM)
    lvarRecordArray(0, CJB_OT_TYPE) = "OT_TYPE"
    lvarRecordArray(1, CJB_OT_TYPE) = l_varRecord(CJB_OT_TYPE)
    lvarRecordArray(0, CJB_EFF_DT) = "EFFDT"
    lvarRecordArray(1, CJB_EFF_DT) = l_varRecord(CJB_EFF_DT)
    lvarRecordArray(0, CJB_EFF_STATUS) = "EFF_STATUS"
    lvarRecordArray(1, CJB_EFF_STATUS) = l_varRecord(CJB_EFF_STATUS)
    lvarRecordArray(0, CJB_UNION_CD) = "UNION_CD"
    lvarRecordArray(1, CJB_UNION_CD) = l_varRecord(CJB_UNION_CD)
    JobRecord = lvarRecordArray
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "JobRecord"
End Function

Public Function JobNumberIsValid( _
            ByVal v_strJobNum As String, _
            Optional ByVal v_datEffDate As Date = CDEFAULTDATE) As Boolean
'When v_datEffDate is default date, use current date.
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim l_strEffStatus As String
    Dim l_datEffDate As Date
    Dim l_datEffDateIn As Date

On Error GoTo eh
    l_varRecord = GetRecord(v_strJobNum, CJB_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    l_strEffStatus = Trim(l_varRecord(CJB_EFF_STATUS))
    If IsDate(l_varRecord(CJB_EFF_DT)) Then
        l_datEffDate = l_varRecord(CJB_EFF_DT)
    Else
        l_datEffDate = CDate(CDEFAULTDATE)
    End If
    l_datEffDateIn = EffectiveDate(v_datEffDate)
    If l_strEffStatus = "A" Then
        JobNumberIsValid = l_datEffDateIn >= l_datEffDate
    ElseIf l_strEffStatus = "I" Then
        JobNumberIsValid = l_datEffDateIn < l_datEffDate
    Else
        JobNumberIsValid = False
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "JobNumberIsValid"
End Function

Public Function JobRecordExists( _
            ByVal v_strJobNum As String) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strJobNum, CJB_GROUPNAME, lblnExists)
    JobRecordExists = Not IsEmpty(l_varRecord)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "JobRecordExists"
End Function

Public Function JobUnionCode( _
            ByVal v_strJobNum As String) As String
' RED152
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strJobNum, CJB_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    JobUnionCode = l_varRecord(CJB_UNION_CD)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "JobUnionCode"
End Function

Public Function JobOvertimeType( _
            ByVal v_strJobNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strJobNum, CJB_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    JobOvertimeType = l_varRecord(CJB_OT_TYPE)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "JobOvertimeType"
End Function

Public Function JobDescription( _
            ByVal v_strJobNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strJobNum, CJB_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    JobDescription = l_varRecord(CJB_JOB_DESC)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "JobDescription"
End Function

Public Function JobCraftNumber( _
            ByVal v_strJobNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strJobNum, CJB_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    JobCraftNumber = l_varRecord(CJB_CRAFT_NUM)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "JobCraftNumber"
                        
End Function

Public Function JobEffectiveDate( _
            ByVal v_strJobNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strJobNum, CJB_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    JobEffectiveDate = l_varRecord(CJB_EFF_DT)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "JobEffectiveDate"
End Function

Public Function JobEffectiveStatus( _
            ByVal v_strJobNum As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strJobNum, CJB_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    JobEffectiveStatus = l_varRecord(CJB_EFF_STATUS)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "JobEffectiveStatus"
End Function

'*********************************************************
' End Job getters
'*********************************************************

'*********************************************************
' Begin Exception getters
'*********************************************************
Public Function ExceptionRecord( _
            ByVal v_strExcCd As String) As Variant
' Public/Friend method.  This should only be called by a friend
' that has knowledge of array structure; i.e. cbsCodeTables.
' Answer an array or empty.
' Array will be a 2 dim zero based array with the field name as the first
' element and the value as the second element.  This will allow
' friend clients to discover the structure instead of hard coding it.
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim lvarRecordArray(1, CEX_LAST_SLOT) As Variant
    Dim i As Integer
            
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    lvarRecordArray(0, CEX_EXCEPTION_CD) = "EXCEPTION_CD"
    lvarRecordArray(1, CEX_EXCEPTION_CD) = l_varRecord(CEX_EXCEPTION_CD)
    lvarRecordArray(0, CEX_EXCEPTION_DESC) = "EXCEPTION_DESC"
    lvarRecordArray(1, CEX_EXCEPTION_DESC) = l_varRecord(CEX_EXCEPTION_DESC)
    lvarRecordArray(0, CEX_CONTRACT_EARN_CD) = "CONTRACT_EARN_CD"
    lvarRecordArray(1, CEX_CONTRACT_EARN_CD) = l_varRecord(CEX_CONTRACT_EARN_CD)
    lvarRecordArray(0, CEX_EXEC_EARN_CD) = "EXEC_EARN_CD"
    lvarRecordArray(1, CEX_EXEC_EARN_CD) = l_varRecord(CEX_EXEC_EARN_CD)
    lvarRecordArray(0, CEX_SYMBOL_TIME_IND) = "SYMBOL_TIME_IND"
    lvarRecordArray(1, CEX_SYMBOL_TIME_IND) = l_varRecord(CEX_SYMBOL_TIME_IND)
    lvarRecordArray(0, CEX_EXCEPTION_BANK_CD) = "EXCEPTION_BANK_CD"
    lvarRecordArray(1, CEX_EXCEPTION_BANK_CD) = l_varRecord(CEX_EXCEPTION_BANK_CD)
    lvarRecordArray(0, CEX_INELIGIBILITY_IND) = "INELIGIBILITY_IND"
    lvarRecordArray(1, CEX_INELIGIBILITY_IND) = l_varRecord(CEX_INELIGIBILITY_IND)
    lvarRecordArray(0, CEX_OTHERPAY_IND) = "OTHERPAY_IND"
    lvarRecordArray(1, CEX_OTHERPAY_IND) = l_varRecord(CEX_OTHERPAY_IND)
    lvarRecordArray(0, CEX_CASE_GROUP) = "CASE_GROUP"
    lvarRecordArray(1, CEX_CASE_GROUP) = l_varRecord(CEX_CASE_GROUP)
    lvarRecordArray(0, CEX_SCHED_IND) = "SCHED_IND"
    lvarRecordArray(1, CEX_SCHED_IND) = l_varRecord(CEX_SCHED_IND)
    lvarRecordArray(0, CEX_FMLA_IND) = "FMLA_IND"
    lvarRecordArray(1, CEX_FMLA_IND) = l_varRecord(CEX_FMLA_IND)
    lvarRecordArray(0, CEX_WORK_REST_IND) = "WORK_REST_IND"
    lvarRecordArray(1, CEX_WORK_REST_IND) = l_varRecord(CEX_WORK_REST_IND)
    lvarRecordArray(0, CEX_SCHED_CHG_IND) = "SCHED_CHG_IND"
    lvarRecordArray(1, CEX_SCHED_CHG_IND) = l_varRecord(CEX_SCHED_CHG_IND)
    lvarRecordArray(0, CEX_ACCT_NUM) = "ACCT_NUM"
    lvarRecordArray(1, CEX_ACCT_NUM) = l_varRecord(CEX_ACCT_NUM)
    lvarRecordArray(0, CEX_EXCEPTION_SYMBOL_CD) = "SYMBOL_CD"
    lvarRecordArray(1, CEX_EXCEPTION_SYMBOL_CD) = l_varRecord(CEX_EXCEPTION_SYMBOL_CD)
    lvarRecordArray(0, CEX_ACCT_NUM) = "AIC_ACCT_NUM"
    lvarRecordArray(1, CEX_ACCT_NUM) = l_varRecord(CEX_AIC_ACCT_NUM)
    lvarRecordArray(0, CEX_ACCT_NUM) = "AMS_ACCT_NUM"
    lvarRecordArray(1, CEX_ACCT_NUM) = l_varRecord(CEX_AMS_ACCT_NUM)
    lvarRecordArray(0, CEX_ACCT_NUM) = "UEC_ACCT_NUM"
    lvarRecordArray(1, CEX_ACCT_NUM) = l_varRecord(CEX_UEC_ACCT_NUM)
    ExceptionRecord = lvarRecordArray
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionRecord"
End Function

Public Function ExceptionType( _
            ByVal v_strExcCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionType = l_varRecord(CEX_EXCEPTION_TYPE)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionType"
End Function

Public Function ExceptionCaseGroup( _
            ByVal v_strExcCd As String) As String
' BLUE082
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionCaseGroup = l_varRecord(CEX_CASE_GROUP)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionCaseGroup"
End Function

Public Function ExceptionDescription( _
            ByVal v_strExcCd As String) As String
' BLUE079
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionDescription = l_varRecord(CEX_EXCEPTION_DESC)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionDescription"
End Function

Public Function ExceptionAccountNumber( _
            ByVal v_strExcCd As String) As String
'BLUE074, BLUE077
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionAccountNumber = l_varRecord(CEX_ACCT_NUM)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionAccountNumber"
End Function

Public Function ExceptionBankCode( _
            ByVal v_strExcCd As String) As String
'BLUE076
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionBankCode = l_varRecord(CEX_EXCEPTION_BANK_CD)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionBankCode"
End Function

Public Function ExceptionBankType( _
            ByVal v_strExcCd As String) As String
' Answer the first character of bank code.
' BLUE021
On Error GoTo eh
    ExceptionBankType = Left(ExceptionBankCode(v_strExcCd), 1)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionBankType"
End Function

Public Function ExceptionEarningsCode( _
            ByVal v_strExcCd As String, _
            ByVal v_strExecConInd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim l_strEarnCd As String
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    If Trim(v_strExecConInd) = "E" Then
        l_strEarnCd = l_varRecord(CEX_EXEC_EARN_CD)
    ElseIf Trim(v_strExecConInd) = "C" Then
        l_strEarnCd = l_varRecord(CEX_CONTRACT_EARN_CD)
    Else
        l_strEarnCd = ""
    End If
    ExceptionEarningsCode = l_strEarnCd
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionEarningsCode"
End Function

Public Function ExceptionExecutiveEarningsCode( _
            ByVal v_strExcCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionExecutiveEarningsCode = l_varRecord(CEX_EXEC_EARN_CD)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionExecutiveEarningsCode"
End Function

Public Function ExceptionContractEarningsCode( _
            ByVal v_strExcCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionContractEarningsCode = l_varRecord(CEX_CONTRACT_EARN_CD)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionContractEarningsCode"
End Function

Public Function ExceptionOtherPayIndicator( _
            ByVal v_strExcCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionOtherPayIndicator = l_varRecord(CEX_OTHERPAY_IND)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionOtherPayIndicator"
End Function

Public Function ExceptionScheduleChgInd( _
            ByVal v_strExcCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionScheduleChgInd = l_varRecord(CEX_SCHED_CHG_IND)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionScheduleChgInd"
End Function

Public Function ExceptionFamilyLeaveInd( _
            ByVal v_strExcCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionFamilyLeaveInd = l_varRecord(CEX_FMLA_IND)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionFamilyLeaveInd"
End Function

Public Function ExceptionCodeIsValid( _
            ByVal v_strExcCd As String) As Boolean
' BLUE094, WHITE043
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionCodeIsValid = Not IsEmpty(l_varRecord)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionCodeIsValid"
End Function

Public Function ExceptionCodeIsValidFor( _
            ByVal v_strExcCd As String, _
            ByVal v_strExecConInd As String) As Boolean

On Error GoTo eh
    ExceptionCodeIsValidFor = _
        (Trim(ExceptionEarningsCode(v_strExcCd, v_strExecConInd)) <> "")
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionCodeIsValidFor"
End Function

Public Function ExceptionWorkRestInd( _
            ByVal v_strExcCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionWorkRestInd = l_varRecord(CEX_WORK_REST_IND)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionWorkRestInd"
End Function
'*********************************************************
' End Exception getters
'*********************************************************

'*********************************************************
' Begin Union getters
'*********************************************************
Public Function UnionRecord( _
            ByVal v_strUnionCd As String) As Variant
' Public/Friend method.  This should only be called by a friend
' that has knowledge of array structure; i.e. cbsCodeTables.
' Answer an array or empty.
' Array will be a 2 dim zero based array with the field name as the first
' element and the value as the second element.  This will allow
' friend clients to discover the structure instead of hard coding it.
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim lvarRecordArray(1, CUN_LAST_SLOT) As Variant
    Dim i As Integer
            
On Error GoTo eh
    l_varRecord = GetRecord(v_strUnionCd, CUN_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    lvarRecordArray(0, CUN_UNION_CD) = "UNION_CD"
    lvarRecordArray(1, CUN_UNION_CD) = l_varRecord(CUN_UNION_CD)
    lvarRecordArray(0, CUN_UNION_LOCAL) = "UNION_LOCAL"
    lvarRecordArray(1, CUN_UNION_LOCAL) = l_varRecord(CUN_UNION_LOCAL)
    lvarRecordArray(0, CUN_UNION_DESC) = "UNION_DESC"
    lvarRecordArray(1, CUN_UNION_DESC) = l_varRecord(CUN_UNION_DESC)
    lvarRecordArray(0, CUN_EXEC_CONTRACT_IND) = "EXEC_CONTRACT_IND"
    lvarRecordArray(1, CUN_EXEC_CONTRACT_IND) = l_varRecord(CUN_EXEC_CONTRACT_IND)
    lvarRecordArray(0, CUN_UNION_BEN_PKG_IND) = "UNION_BEN_PKG_IND"
    lvarRecordArray(1, CUN_UNION_BEN_PKG_IND) = l_varRecord(CUN_UNION_BEN_PKG_IND)
    UnionRecord = lvarRecordArray
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "UnionRecord"
End Function

Public Function UnionDescription( _
            ByVal v_strUnionCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    l_varRecord = GetRecord(v_strUnionCd, CUN_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    UnionDescription = l_varRecord(CUN_UNION_DESC)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "UnionDescription"
End Function

Public Function UnionExecutiveContractIndicator( _
            ByVal v_strUnionCd As String) As String
' WHITE084
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    l_varRecord = GetRecord(v_strUnionCd, CUN_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    UnionExecutiveContractIndicator = _
        l_varRecord(CUN_EXEC_CONTRACT_IND)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "UnionExecutiveContractIndicator"
End Function

Public Function UnionBenefitPackageIndicator( _
            ByVal v_strUnionCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    l_varRecord = GetRecord(v_strUnionCd, CUN_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    UnionBenefitPackageIndicator = _
        l_varRecord(CUN_UNION_BEN_PKG_IND)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "UnionBenefitPackageIndicator"
End Function

Public Function UnionLocal( _
            ByVal v_strUnionCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    l_varRecord = GetRecord(v_strUnionCd, CUN_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    UnionLocal = l_varRecord(CUN_UNION_LOCAL)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "UnionLocal"
End Function

Public Function UnionCodeIsValid( _
            ByVal v_strUnionCd As String) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    l_varRecord = GetRecord(v_strUnionCd, CUN_GROUPNAME, lblnExists)
    UnionCodeIsValid = Not IsEmpty(l_varRecord)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "UnionCodeIsValid"
End Function

'*********************************************************
' End Union getters
'*********************************************************


'*********************************************************
' Begin Union component getters
'*********************************************************

Public Function UnionComponentVacationReliefAmount( _
        ByVal v_strUnionCd As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCd As String) As Variant
    
On Error GoTo eh
    UnionComponentVacationReliefAmount = _
        GetComponentAmountOn( _
                        v_strUnionCd, _
                        CUC_RELIEF_COMP_TYPE, _
                        v_strCompCd, _
                        v_datEffDate)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "UnionComponentVacationReliefAmount"
End Function

Public Function UnionComponentShiftAmount( _
        ByVal v_strUnionCd As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCd As String) As Variant
    
On Error GoTo eh
    UnionComponentShiftAmount = _
        GetComponentAmountOn( _
                        v_strUnionCd, _
                        CUC_SHIFT_COMP_TYPE, _
                        v_strCompCd, _
                        v_datEffDate)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "UnionComponentShiftAmount"
End Function

Public Function UnionComponentAmount( _
        ByVal v_strUnionCd As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCd As String) As Variant
    
On Error GoTo eh
    UnionComponentAmount = _
        GetComponentAmountOn( _
                        v_strUnionCd, _
                        CUC_OTHER_COMP_TYPE, _
                        v_strCompCd, _
                        v_datEffDate)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "UnionComponentAmount"
End Function

'*********************************************************
' End Union component getters
'*********************************************************

'*********************************************************
' Begin Overtime Reason getters
'*********************************************************
Public Function OvertimeReasonIsValid( _
            ByVal v_strOTReasonCd As String) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    l_varRecord = GetRecord(v_strOTReasonCd, COR_GROUPNAME, lblnExists)
    OvertimeReasonIsValid = Not IsEmpty(l_varRecord)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OvertimeReasonIsValid"
End Function

Public Function OvertimeReasonDescription( _
            ByVal v_strOTReasonCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strOTReasonCd, COR_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    OvertimeReasonDescription = l_varRecord(COR_OT_REASON_DESC)
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "OvertimeReasonDescription"
End Function

'*********************************************************
' End Overtime Reason getters
'*********************************************************

'*********************************************************
' Begin Overtime Response getters
'*********************************************************
Public Function OvertimeResponseIsValid( _
            ByVal v_strOTResponse As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strOTResponse, COP_GROUPNAME, lblnExists)
    OvertimeResponseIsValid = Not IsEmpty(l_varRecord)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OvertimeResponseIsValid"
End Function

Public Function OvertimeResponseDescription( _
            ByVal v_strOTResponse As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strOTResponse, COP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    OvertimeResponseDescription = l_varRecord(COP_RESPONSE_DESC)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OvertimeResponseDescription"
End Function

Public Function OvertimeResponseWorkIndicator( _
            ByVal v_strOTResponse As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strOTResponse, COP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    OvertimeResponseWorkIndicator = l_varRecord(COP_OT_WORKED_IND)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OvertimeResponseWorkIndicator"
End Function

Public Function OvertimeResponseChargeIndicator( _
            ByVal v_strOTResponse As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strOTResponse, COP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    OvertimeResponseChargeIndicator = l_varRecord(COP_OT_CHRG_IND)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OvertimeResponseChargeIndicator"
End Function
'*********************************************************
' End Overtime Response getters
'*********************************************************

'*********************************************************
' Begin Overtime Exception getters
'*********************************************************
Public Function OvertimeExceptionFactor( _
            ByVal v_strExceptionCode As String) As Variant
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExceptionCode, COE_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    OvertimeExceptionFactor = CDec(l_varRecord(COE_FACTOR))
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OvertimeExceptionFactor"
End Function

'*********************************************************
' End Overtime Exception getters
'*********************************************************

'*********************************************************
' Begin Rotation getters
'*********************************************************
Public Function RotationCodeIsValid( _
            ByVal v_strRotCd As String) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean

On Error GoTo eh

    l_varRecord = GetRecord(Array(v_strRotCd), CRC_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    RotationCodeIsValid = Not IsEmpty(l_varRecord)
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "RotationCodeIsValid"
End Function


Public Function RotationWeeksIsValid( _
            ByVal v_strRotCd As String, _
            ByVal v_lngRot_Weeks As Long) As Boolean
' Zero is not a valid rotation week for a valid rotation code.
' Rows exist on TTR_ROTATION and TTR_ROTATION_SCHED with ROT_CD and ROT_WEEK
' zero values.  These exist to support referential integrity.
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
On Error GoTo eh

    l_varRecord = GetRecord(Array(v_strRotCd), CRC_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    RotationWeeksIsValid = _
        v_lngRot_Weeks > 0 And _
        v_lngRot_Weeks <= l_varRecord(CRC_ROT_WEEKS)
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "RotationWeeksIsValid"
End Function


Public Function RotationWeeks( _
            ByVal v_strRotCd As String) As Long
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean

On Error GoTo eh

    l_varRecord = GetRecord(Array(v_strRotCd), CRC_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    RotationWeeks = l_varRecord(CRC_ROT_WEEKS)
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "RotationWeeks"
End Function

'*********************************************************
' End Rotation getters
'*********************************************************

'*********************************************************
' Begin Available schedule getters
'*********************************************************

Public Function ScheduleCodeIsValid( _
            ByVal v_strSchedCd As String) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
   
On Error GoTo eh
    l_varRecord = GetRecord(Array(v_strSchedCd), CSC_GROUPNAME, lblnExists)
    ScheduleCodeIsValid = Not IsEmpty(l_varRecord)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ScheduleCodeIsValid"
End Function

''//-----Each field in this compound key needs to match the corresponding column defined size in the cached recordset-----//
'Private Function IntegerKey( _
'    ByVal v_lngKey As Long, _
'    ByVal v_eDataType As DataTypeEnum _
'    ) As String
'
'    Dim lstrInteger As String
'    Dim l_lngLength As Long
'
'On Error GoTo eh
'    Select Case v_eDataType
'    Case adInteger
'        l_lngLength = 11 '//extra for sign
'    Case adSmallInt
'        l_lngLength = 6 '//extra for sign
'    Case Else
'        l_lngLength = 6 '//extra for sign
'    End Select
'
'    lstrInteger = String(l_lngLength, "0") '//-----PadLeft with zeros
'    lstrInteger = lstrInteger & Abs(Format(v_lngKey))
'    IntegerKey = IIf(v_lngKey < 0, "-", "0") & Right(lstrInteger, l_lngLength - 1)
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "IntegerKey"
'End Function

'*********************************************************
' End Available schedule getters
'*********************************************************

'*********************************************************
' Begin Component code getters
'*********************************************************
Public Function ComponentCodeIsValid( _
            ByVal v_strCompCd As String, _
            ByVal v_lngCompType As TrsComponentType) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim l_strCompType As String
    
On Error GoTo eh
    Select Case v_lngCompType
        Case trReliefComponent
            l_strCompType = "R"
        Case trShiftComponent
            l_strCompType = "S"
        Case trOtherComponent
            l_strCompType = "O"
    End Select
    l_varRecord = GetRecord(v_strCompCd, CCC_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ComponentCodeIsValid = Trim(l_varRecord(CCC_COMP_TYPE)) = l_strCompType
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "ComponentCodeIsValid"
End Function

'*********************************************************
' End Component code getters
'*********************************************************

'*********************************************************
' Begin Employee getters
'*********************************************************

Public Function EmployeeName( _
            ByVal v_lngEmpNum As Long) As String
'    Dim l_varRecord As Variant
'    Dim lblnExists As Boolean
'
'On Error GoTo eh
'    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
'    If IsEmpty(l_varRecord) Then
'        Exit Function
'    End If
'    ExceptionAccountNumber = l_varRecord(CEX_ACCT_NUM)
'    Exit Function
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo eh
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString()
    Set l_objCM = Nothing
    l_objCN.Open

    l_strSQL = "SELECT EMP_NAME" _
        & " FROM METHDBA.TTR_EMPLOYEE" _
        & " WHERE EMP_NUM = " & v_lngEmpNum
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    
    If RecordsExist(l_objRS) Then
        EmployeeName = Trim(l_objRS!EMP_NAME)
    End If
    Set l_objRS = Nothing

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeName"
End Function

'*********************************************************
' Begin Employee getters
'*********************************************************


'*********************************************************
' Begin Public Cache management methods
'*********************************************************

Public Function GetGroupRecords( _
                    ByVal v_strGroupName As String, _
                    ByVal v_varKeys As Variant, _
                    ByRef r_varKeyArray As Variant, _
                    ByRef r_varColumnArray As Variant) As ADODB.Recordset
    Dim l_rsRecords As ADODB.Recordset
    Dim l_colErrArgs As Collection
On Error GoTo eh

    Select Case v_strGroupName
        Case CLC_GROUPNAME
            Set l_rsRecords = GetLocationRecords(v_varKeys)
            r_varKeyArray = Array("LOC_CD")
            r_varColumnArray = Empty
        Case CEX_GROUPNAME
            Set l_rsRecords = GetExceptionRecords(v_varKeys)
            r_varKeyArray = Array("EXCEPTION_CD") '//-----Columns for building Key '//ExceptionGroupKeys
            r_varColumnArray = Empty '//-----cache all columns in recordset
        Case CUN_GROUPNAME
            Set l_rsRecords = GetUnionRecords(v_varKeys)
            r_varKeyArray = Array("UNION_CD")
            r_varColumnArray = Empty
        Case CUC_GROUPNAME
            Set l_rsRecords = GetUnionComponentRecords(v_varKeys)
            r_varKeyArray = Array("UNION_CD", "COMP_TYPE", "COMP_CD")
            r_varColumnArray = Array("EFF_DATE", "COMP_AMT")
        Case CJB_GROUPNAME
            Set l_rsRecords = GetJobRecords(v_varKeys)
            r_varKeyArray = Array("JOB_NUM")
            r_varColumnArray = Empty
        Case COR_GROUPNAME
            Set l_rsRecords = GetOvertimeReasonRecords(v_varKeys)
            r_varKeyArray = Array("OT_REASON_CD")
            r_varColumnArray = Empty
        Case COP_GROUPNAME
            Set l_rsRecords = GetOvertimeResponseRecords(v_varKeys)
            r_varKeyArray = Array("OT_RESP_CD")
            r_varColumnArray = Empty
        Case COE_GROUPNAME
            Set l_rsRecords = GetOvertimeExceptionRecords(v_varKeys)
            r_varKeyArray = Array("EXCEPTION_CD")
            r_varColumnArray = Empty
        Case CHS_GROUPNAME
            Set l_rsRecords = GetHSCutOffDateRecords(v_varKeys)
            r_varKeyArray = Array("HS_CUTOFF_KEY")
            r_varColumnArray = Array("HS_CUTOFF_KEY", "HS_CUTOFF_DATE")
        Case CSC_GROUPNAME
            Set l_rsRecords = GetScheduleRecords(v_varKeys)
            r_varKeyArray = Array("SCHED_CD")
            r_varColumnArray = Empty
        Case CRC_GROUPNAME
            Set l_rsRecords = GetRotationRecords(v_varKeys)
            r_varKeyArray = Array("ROT_CD")
            r_varColumnArray = Empty
        Case CCC_GROUPNAME
            Set l_rsRecords = GetComponentRecords(v_varKeys)
            r_varKeyArray = Array("COMP_CD")
            r_varColumnArray = Empty
        Case CEG_GROUPNAME
            Set l_rsRecords = GetExceptionGroupsRecords(v_varKeys)
            r_varKeyArray = Array("EG_CD", "EXCEPTION_CD", "EG_SUB_CD")
            r_varColumnArray = Array("EXC_SEQ_NUM")
    End Select
    Set GetGroupRecords = l_rsRecords

    Set l_rsRecords = Nothing
    Exit Function
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    l_colErrArgs.Add "v_varKeys = " & v_varKeys
    l_colErrArgs.Add "l_rsRecords.Source = " & l_rsRecords.Source
    
    Set l_rsRecords = Nothing
    TrsRaiseError m_strCMODULENAME, "GetGroupRecords", , , , l_colErrArgs, True
End Function



Public Sub InvalidateAvailScheduleCache(ByVal v_strSchedCd As String)
    Dim lstrKey As String
On Error GoTo eh
'//    lstrKey = IntegerKey(v_strSchedCd, adSmallInt)

    InvalidateCache Array(v_strSchedCd), CSC_GROUPNAME
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateAvailScheduleCache"
End Sub

Public Sub InvalidateRotationCache(ByVal v_strRotCd As String)
    Dim lstrKey As String
On Error GoTo eh
'//    lstrKey = IntegerKey(v_strRotCd, adSmallInt)

    InvalidateCache Array(v_strRotCd), CRC_GROUPNAME
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateRotationCache"
End Sub

Public Sub InvalidateExceptionCache(ByVal v_strExcCd As String)

On Error GoTo eh

    InvalidateCache v_strExcCd, CEX_GROUPNAME
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateExceptionCache"
End Sub

Public Sub InvalidateOTExceptionCache(ByVal v_strOTExcCd As String)

On Error GoTo eh

    InvalidateCache v_strOTExcCd, COE_GROUPNAME
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateOTExceptionCache"
End Sub

Public Sub InvalidateOTReasonCache(ByVal v_strOTRsnCd As String)

On Error GoTo eh

    InvalidateCache v_strOTRsnCd, COR_GROUPNAME
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateOTReasonCache"
End Sub

Public Sub InvalidateOTResponseCache(ByVal v_strOTRspCd As String)

On Error GoTo eh

    InvalidateCache v_strOTRspCd, COP_GROUPNAME
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateOTResponseCache"
End Sub

Public Sub InvalidateJobCache(ByVal v_strJobNum As String)

On Error GoTo eh

    InvalidateCache v_strJobNum, CJB_GROUPNAME
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateJobCache"
End Sub

Public Sub InvalidateHSCutoffDateCache(ByVal v_strHSCutoffKey As String)

On Error GoTo eh

    InvalidateCache v_strHSCutoffKey, CHS_GROUPNAME
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateHSCutoffDateCache"
End Sub

'//-----
Public Sub InvalidateLocationCache(ByVal v_strLocCd As String)

On Error GoTo eh

    '//-----location code is a small int in the db
    '//-----wrap it with array in order to use the meta data which will format it correctly "1"->000001
    InvalidateCache Array(v_strLocCd), CLC_GROUPNAME
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateLocationCache"
End Sub

Public Sub InvalidateUnionCache(ByVal v_strUnionCd As String)

On Error GoTo eh

    InvalidateCache v_strUnionCd, CUN_GROUPNAME
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateUnionCache"
End Sub

Public Sub InvalidateUnionComponentCache(ByVal v_strUnionCd As String, _
                                        ByVal v_strCompType As String, _
                                        ByVal v_strCompCd As String)
On Error GoTo eh

    InvalidateCache Array(v_strUnionCd, v_strCompType, v_strCompCd), CUC_GROUPNAME
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateUnionComponentCache"
End Sub

Public Sub InvalidateComponentCache(ByVal v_strCompCd As String)

On Error GoTo eh

    InvalidateCache v_strCompCd, CCC_GROUPNAME
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateComponentCache"
End Sub

Public Sub InvalidateExceptionGroupsCache(ByVal v_intEgCd As String, _
                                        ByVal v_strExcCd As String, _
                                        ByVal v_strEgSubCd As String)

On Error GoTo eh

    InvalidateCache Array(v_intEgCd, v_strExcCd, v_strEgSubCd), CEG_GROUPNAME
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateExceptionGroupsCache"
End Sub

'*********************************************************
' End Public Cache management methods
'*********************************************************




'*********************************************************
' Begin Private procedures
'*********************************************************
Private Function GetComponentAmountOn( _
            ByVal v_strUnionCd As String, _
            ByVal v_strCompType As String, _
            ByVal v_strCompCd As String, _
            ByVal v_datEffDate As Date) As Variant
' Return the effective amount on v_datEffDate or NULL.
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim lstrKey As String
            
On Error GoTo eh
    '//lstrKey = UnionComponentKey(v_strUnionCd, v_strCompType, v_strCompCd)
    l_varRecord = GetRecord(Array(v_strUnionCd, v_strCompType, v_strCompCd), CUC_GROUPNAME, lblnExists)
    GetComponentAmountOn = _
        FindComponentAmountIn(v_datEffDate, l_varRecord)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetComponentAmountOn"
End Function

Private Function FindComponentAmountIn( _
            ByVal v_datEffDate As Date, _
            r_varRecord As Variant) As Variant
' Return the effective amount on v_datEffDate or NULL.
' l_varRowValues is a 2 dim array sorted in EFF_DATE descending order
    Dim i As Integer
    Dim l_varEffCompAmt As Variant
    Dim l_varRowValues As Variant

On Error GoTo eh
    If IsEmpty(r_varRecord) Then
        FindComponentAmountIn = Null
        Exit Function
    End If
    
    l_varRowValues = r_varRecord(CUC_ROWVALUES)
    For i = LBound(l_varRowValues, 2) To UBound(l_varRowValues, 2)
        If v_datEffDate >= l_varRowValues(CUC_EFF_DATE, i) Then
            l_varEffCompAmt = l_varRowValues(CUC_COMP_AMT, i)
            Exit For
        End If
    Next

    If IsEmpty(l_varEffCompAmt) Then
        l_varEffCompAmt = Null
    End If
    FindComponentAmountIn = l_varEffCompAmt
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "FindComponentAmountIn"
End Function

Private Function EffectiveDate(ByVal v_datChgDate As Date) As Date
    If v_datChgDate = CDEFAULTDATE Then
        EffectiveDate = Date
    Else
        EffectiveDate = v_datChgDate
    End If
End Function


'*********************************************************
' Begin Private Caching methods
'*********************************************************

Private Function GetRecord( _
            ByVal v_varKey As Variant, _
            ByVal v_strGroupName As String, _
            ByRef r_blnExists As Boolean) As Variant
' Answer an variant array, representing a record, or empty.
' This method should not invoke InitializeGroup unless the
' group is setup for bulk initialization.  If setup for lazy
' initialization it should only get the single record and add it
' to the cache.  If the record is not found, using lazy, it should
' add a value to the cache to denote invalid so that it does not
' attempt to retrieve it from the database again.  Currently, a string
' constant is added for record not found and the spmCache component
' will answer that it exists, but the property value is a
' RECORDNOTFOUND string.
'
    Dim recCache As utilRecordCache.cuRecordCache
    Dim l_colErrArgs As Collection
'    Dim l_varRecord As Variant
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim lblnGroupExists As Boolean
           
On Error GoTo eh

'    Select Case Trim(v_strGroupName)
'    Case CEX_GROUPNAME, CUC_GROUPNAME, CJB_GROUPNAME, CEG_GROUPNAME

'//--------------------------------------------------------------------------------------------------------------//
'//-----http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cossdk/htm/imtxas_2bs6.asp
'//-----Do not use the SafeRef method in COM+. In MTS, SafeRef was used by an object to obtain
'//-----a reference to itself that could be passed outside its context wrapper.
'//-----With contexts and interception fully integrated into COM+, this is no longer necessary.
'//-----
'//-----[C/C++]
'//-----Note   The C++ function SafeRef should be used instead of this method.
'//--------------------------------------------------------------------------------------------------------------//

        Set recCache = CreateObject("utilRecordCache.cuRecordCache")
    
        If IsArray(v_varKey) Then
            GetRecord = recCache.GetRecord2(v_varKey, v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                r_blnExists, Me)
        Else
            GetRecord = recCache.GetRecord(CStr(v_varKey), v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                r_blnExists, Me)
        End If
        
        Set recCache = Nothing
'    Case Else
'        Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'        l_varRecord = _
'            spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
'        If Not lblnGroupExists Then
'            If IsLazyInitialized(v_strGroupName) Then
'                LoadRecord v_strKey, v_strGroupName
'                l_varRecord = _
'                    spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
'            Else
'                InitializeGroup (v_strGroupName)
'                l_varRecord = _
'                    spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
'            End If
'        Else
'            If Not r_blnExists Then
'                If IsLazyInitialized(v_strGroupName) Then
'                    LoadRecord v_strKey, v_strGroupName
'                    l_varRecord = _
'                        spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
'                End If
'            Else    'it exists and INVALID - reload it
'                    '***There must be a Load method for every cache that supports Invalidate
'                If Not IsArray(l_varRecord) Then
'                    If l_varRecord = CSPINVALID And _
'                            GroupSupportsLoadRecord(v_strGroupName) Then
'                        LoadRecord v_strKey, v_strGroupName
'                        l_varRecord = _
'                            spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
'                    End If
'                End If
'            End If
'        End If
'        ' test for array first since 99.9% of time it should be one,
'        ' otherwise answer empty.
'        If IsArray(l_varRecord) Then
'            GetRecord = l_varRecord
'        Else
'            GetRecord = Empty
'        End If
'        Set spmWrapper = Nothing
'    End Select
    
    Exit Function
eh:
    Set recCache = Nothing
'    Set spmWrapper = Nothing

    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & CStr(v_varKey)
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    
    TrsRaiseError m_strCMODULENAME, "GetRecord", , , , l_colErrArgs, True
End Function

'//-----code has been factored into utilRecordCache-----//
Private Sub InvalidateCache( _
                ByVal v_varKey As Variant, _
                ByVal v_strGroupName As String)

    Dim recCache As utilRecordCache.cuRecordCache
    Dim l_colErrArgs As Collection
    Dim l_blnExists As Boolean
On Error GoTo eh

    Set recCache = CreateObject("utilRecordCache.cuRecordCache")
    
    If IsArray(v_varKey) Then '//-----uses metadata to format key, may need to retrieve data to get metadata
        recCache.InvalidateCache2 v_varKey, v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                l_blnExists, Me
    Else
        recCache.InvalidateCache CStr(v_varKey), v_strGroupName
    End If
        
    Set recCache = Nothing

    Exit Sub
eh:
    Set recCache = Nothing
    
    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & CStr(v_varKey)
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    
    TrsRaiseError m_strCMODULENAME, "InvalidateCache", , , , l_colErrArgs, True
End Sub
                
''//-----this could be factored into utilRecordCache-----//
'Private Sub InvalidateCache( _
'                ByVal v_strKey As String, _
'                ByVal v_strGroupName As String)
'' RSG 11/07/02 - Invalidate all servers in cluster.
'    Dim spmProxy As utilSPMCacheProxy.cuSPMCacheProxy
'    Dim l_strKey As String
'    Dim l_strThisComputerName As String
'    Dim l_varAllServers As Variant
'    Dim i As Integer
'
'On Error GoTo eh
'
'    l_strThisComputerName = Trim(UCase(ThisComputerName))
'    l_varAllServers = GetAllServerNames(l_strThisComputerName)
'    For i = LBound(l_varAllServers) To UBound(l_varAllServers)
'        If l_varAllServers(i) = l_strThisComputerName Then
'            Set spmProxy = CreateObject("utilSPMCacheProxy.cuSPMCacheProxy")
'        Else
'            Set spmProxy = CreateObject("utilSPMCacheProxy.cuSPMCacheProxy", l_varAllServers(i))
'        End If
'        l_strKey = Trim(v_strKey)
'        spmProxy.Invalidate l_strKey, v_strGroupName
'        Set spmProxy = Nothing
'    Next
'    Exit Sub
'
'eh:
'    Set spmProxy = Nothing
'    TrsRaiseError m_strCMODULENAME, "InvalidateCache"
'End Sub
'
''//-----this could be factored into utilRecordCache or some other util object for reusablity-----//
'Private Function GetAllServerNames(ByVal v_strComputerName As String) As Variant
'' Answer an array.
'' Get all available servers in the same cluster as v_strComputerName,
'' including v_strComputerName.  If GetServersInMyCluster answers empty then
'' v_strComputerName is a developer workstation, just add it to the array.
'    Dim oPing As utilTRISPingRouter.cuTRISPingRouter
'    Dim l_varAllServers As Variant
'
'On Error GoTo eh
'
'    Set oPing = CreateObject("utilTRISPingRouter.cuTRISPingRouter")
'    l_varAllServers = oPing.GetServersInMyCluster(v_strComputerName)
'    If IsEmpty(l_varAllServers) Then
'        ReDim l_varAllServers(0)
'        l_varAllServers(0) = v_strComputerName
'    End If
'    GetAllServerNames = l_varAllServers
'    Set oPing = Nothing
'    Exit Function
'
'eh:
'    Set oPing = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetAllServerNames"
'End Function
'
'
'Private Function InitializeGroup( _
'            ByVal v_strGroupName As String) As Variant
'
'On Error GoTo eh
'    Select Case v_strGroupName
''        Case CLC_GROUPNAME
''            InitializeLocationGroup
''//        Case CEX_GROUPNAME
''//            InitializeExceptionGroup
''        Case CUN_GROUPNAME
''            InitializeUnionGroup
''//        Case CUC_GROUPNAME
''//            InitializeUnionComponentGroup
''        Case COR_GROUPNAME
''            InitializeOvertimeReasonGroup
''        Case COP_GROUPNAME
''            InitializeOvertimeResponseGroup
''        Case COE_GROUPNAME
''            InitializeOvertimeExceptionGroup
''        Case CHS_GROUPNAME
''            InitializeHSCutOffDateGroup
''        Case CSC_GROUPNAME
''            InitializeAvailableScheduleGroup
''        Case CRC_GROUPNAME
''            InitializeRotationGroup
''        Case CCC_GROUPNAME
''            InitializeComponentCodeGroup
''//        Case CEG_GROUPNAME
''//            InitializeExceptionGroupsGroup
'    End Select
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "InitializeGroup"
'End Function

'Private Function InitializeUnionComponentGroup() As Boolean
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    Dim lvarRecord(CUC_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'    Dim dc As utilTRISCommonLib.cuSet
'    Dim lvarKey As Variant
'    Dim lstrFilter As String
'    Dim laryFields As Variant
'    Dim laryValues As Variant
'    Dim lstrKey As String * 7
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'    Set rs = GetEffectiveUnionComponents
'    If rs.BOF And rs.EOF Then
'        InitializeUnionComponentGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
''get distinct list then filter and sort each, with getrows
'    rs.MoveFirst
'    Set dc = New cuSet
'    Do Until rs.EOF
'        lstrKey = UnionComponentKey(rs!UNION_CD, rs!COMP_TYPE, rs!COMP_CD)
'        dc.Add lstrKey
'        rs.MoveNext
'    Loop
'    rs.MoveFirst
'
'    llngRecordCount = dc.Count
'    ReDim lvarRecordArray(1, llngRecordCount - 1)
'
'    laryFields = Array("EFF_DATE", "COMP_AMT")
'    For Each lvarKey In dc
'        lstrFilter = _
'            "UNION_CD = '" & Mid(lvarKey, 1, 3) & "' AND " & _
'            "COMP_TYPE = '" & Mid(lvarKey, 4, 1) & "' AND " & _
'            "COMP_CD = '" & Mid(lvarKey, 5, 3) & "'"
'        lstrFilter = MDAC25FilterFix(lstrFilter)
'        rs.Filter = lstrFilter
'        rs.Sort = "EFF_DATE DESC"
'        laryValues = rs.GetRows(, , laryFields)
'        lvarRecord(CUC_COMP_KEY) = lvarKey
'        lvarRecord(CUC_ROWVALUES) = laryValues
'        lvarRecordArray(0, i) = lvarRecord(CUC_COMP_KEY)
'        lvarRecordArray(1, i) = lvarRecord
'        i = i + 1
'    Next
'
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, CUC_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build union component group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeUnionComponentGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeUnionComponentGroup"
'End Function
'
''//-----Each field in this compound key needs to match the corresponding column defined size in the cached recordset-----//
'Private Function UnionComponentKey( _
'                    ByVal v_strUnionCd As String, _
'                    ByVal v_strCompType As String, _
'                    ByVal v_strCompCd As String) As String
'    Dim lstrResult As String * 7
'
'On Error GoTo eh
'    lstrResult = Space(7)
'    Mid(lstrResult, 1, 3) = v_strUnionCd
'    Mid(lstrResult, 4, 1) = v_strCompType
'    Mid(lstrResult, 5, 3) = v_strCompCd
'    UnionComponentKey = lstrResult
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "UnionComponentKey"
'End Function
'
'
'Private Function InitializeUnionGroup() As Boolean
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    Dim lvarRecord(CUN_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'    Set rs = GetAllUnions
'    If rs.BOF And rs.EOF Then
'        InitializeUnionGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    llngRecordCount = rs.RecordCount
'    ReDim lvarRecordArray(1, llngRecordCount - 1)
'    rs.MoveFirst
'    Do Until rs.EOF
'        lvarRecord(CUN_UNION_CD) = Trim(rs!UNION_CD)
'        lvarRecord(CUN_UNION_LOCAL) = Trim(rs!UNION_LOCAL)
'        lvarRecord(CUN_UNION_DESC) = Trim(rs!UNION_DESC)
'        lvarRecord(CUN_EXEC_CONTRACT_IND) = Trim(rs!EXEC_CONTRACT_IND)
'        lvarRecord(CUN_UNION_BEN_PKG_IND) = Trim(rs!UNION_BEN_PKG_IND)
'        lvarRecordArray(0, i) = lvarRecord(CUN_UNION_CD)
'        lvarRecordArray(1, i) = lvarRecord
'        rs.MoveNext
'        i = i + 1
'    Loop
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, CUN_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build union code group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeUnionGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeUnionGroup"
'End Function


'Private Function InitializeLocationGroup() As Boolean
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    Dim lvarRecord(CLC_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'    Set rs = GetAllLocations
'    If rs.BOF And rs.EOF Then
'        InitializeLocationGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    llngRecordCount = rs.RecordCount
'    ReDim lvarRecordArray(1, llngRecordCount - 1)
'    rs.MoveFirst
'    Do Until rs.EOF
'        lvarRecord(CLC_LOC_CD) = Trim(rs!LOC_CD)
'        lvarRecord(CLC_LOC_DESC) = Trim(rs!LOC_DESC)
'        lvarRecord(CLC_PLANT_CD) = Trim(rs!PLANT_CD)
'        lvarRecordArray(0, i) = lvarRecord(CLC_LOC_CD)
'        lvarRecordArray(1, i) = lvarRecord
'        rs.MoveNext
'        i = i + 1
'    Loop
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, CLC_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build location code group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeLocationGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeLocationGroup"
'End Function
'
'
'Private Function InitializeOvertimeReasonGroup() As Boolean
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    Dim lvarRecord(COR_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'    Set rs = GetAllOvertimeReasons
'    If rs.BOF And rs.EOF Then
'        InitializeOvertimeReasonGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    llngRecordCount = rs.RecordCount
'    ReDim lvarRecordArray(1, llngRecordCount - 1)
'    rs.MoveFirst
'    Do Until rs.EOF
'        lvarRecord(COR_OT_REASON_CD) = Trim(rs!OT_REASON_CD)
'        lvarRecord(COR_OT_REASON_DESC) = Trim(rs!OT_REASON_DESC)
'        lvarRecordArray(0, i) = lvarRecord(COR_OT_REASON_CD)
'        lvarRecordArray(1, i) = lvarRecord
'        rs.MoveNext
'        i = i + 1
'    Loop
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, COR_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build overtime reason group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeOvertimeReasonGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeOvertimeReasonGroup"
'End Function


'Private Function InitializeOvertimeResponseGroup() As Boolean
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    Dim lvarRecord(COP_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'    Set rs = GetAllOvertimeResponses
'    If rs.BOF And rs.EOF Then
'        InitializeOvertimeResponseGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    llngRecordCount = rs.RecordCount
'    ReDim lvarRecordArray(1, llngRecordCount - 1)
'    rs.MoveFirst
'    Do Until rs.EOF
'        lvarRecord(COP_OT_RESP_CD) = Trim(rs!OT_RESP_CD)
'        lvarRecord(COP_RESPONSE_DESC) = Trim(rs!RESPONSE_DESC)
'        lvarRecord(COP_OT_CHRG_IND) = Trim(rs!OT_CHRG_IND)
'        lvarRecord(COP_OT_WORKED_IND) = Trim(rs!OT_WORKED_IND)
'        lvarRecordArray(0, i) = lvarRecord(COP_OT_RESP_CD)
'        lvarRecordArray(1, i) = lvarRecord
'        rs.MoveNext
'        i = i + 1
'    Loop
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, COP_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build overtime response group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeOvertimeResponseGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeOvertimeResponseGroup"
'End Function

'
'Private Function InitializeOvertimeExceptionGroup() As Boolean
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    Dim lvarRecord(COE_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'    Set rs = GetOvertimeExceptions
'    If rs.BOF And rs.EOF Then
'        InitializeOvertimeExceptionGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    llngRecordCount = rs.RecordCount
'    ReDim lvarRecordArray(1, llngRecordCount - 1)
'    rs.MoveFirst
'    Do Until rs.EOF
'        lvarRecord(COE_EXCEPTION_CD) = Trim(rs!EXCEPTION_CD)
'        lvarRecord(COE_EXCEPTION_DESC) = Trim(rs!EXCEPTION_DESC)
'        lvarRecord(COE_FACTOR) = Trim(CStr(rs!Factor))
'        lvarRecord(COE_EXEC_CONTRACT_IND) = Trim(rs!EXEC_CONTRACT_IND)
'        lvarRecord(COE_EARNINGS_CD) = Trim(rs!EARNINGS_CD)
'        lvarRecordArray(0, i) = lvarRecord(COE_EXCEPTION_CD)
'        lvarRecordArray(1, i) = lvarRecord
'        rs.MoveNext
'        i = i + 1
'    Loop
'
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, COE_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build overtime exception group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeOvertimeExceptionGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeOvertimeExceptionGroup"
'End Function


'Private Function ExceptionGroupKeys() As Variant
'' This name must match recordset name exactly
'
'On Error GoTo eh
'    ExceptionGroupKeys = Array("EXCEPTION_CD")
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "ExceptionGroupKeys"
'End Function


'Private Function InitializeExceptionGroup() As Boolean
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    'group specific
'    Dim lvarRecord(CEX_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'    'group specific
'    Set rs = GetAllExceptions
'    If rs.BOF And rs.EOF Then
'        InitializeExceptionGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    llngRecordCount = rs.RecordCount
'    ReDim lvarRecordArray(1, llngRecordCount - 1)
'    rs.MoveFirst
'    'group specific
'    Do Until rs.EOF
'        lvarRecord(CEX_EXCEPTION_CD) = Trim(rs!EXCEPTION_CD)
'        lvarRecord(CEX_EXCEPTION_DESC) = Trim(rs!EXCEPTION_DESC)
'        lvarRecord(CEX_EXCEPTION_TYPE) = UCase(Trim(rs!EXCEPTION_TYPE))
'        lvarRecord(CEX_CONTRACT_EARN_CD) = Trim(rs!CONTRACT_EARN_CD)
'        lvarRecord(CEX_EXEC_EARN_CD) = Trim(rs!EXEC_EARN_CD)
'        lvarRecord(CEX_SYMBOL_TIME_IND) = Trim(rs!SYMBOL_TIME_IND)
'        lvarRecord(CEX_EXCEPTION_BANK_CD) = Trim(rs!EXCEPTION_BANK_CD)
'        lvarRecord(CEX_INELIGIBILITY_IND) = Trim(rs!INELIGIBILITY_IND)
'        lvarRecord(CEX_OTHERPAY_IND) = UCase(Trim(rs!OTHERPAY_IND))
'        lvarRecord(CEX_CASE_GROUP) = Trim(rs!CASE_GROUP)
'        lvarRecord(CEX_SCHED_IND) = Trim(rs!SCHED_IND)
'        lvarRecord(CEX_FMLA_IND) = Trim(rs!FMLA_IND)
'        lvarRecord(CEX_WORK_REST_IND) = Trim(rs!WORK_REST_IND)
'        lvarRecord(CEX_SCHED_CHG_IND) = Trim(rs!SCHED_CHG_IND)
'        lvarRecord(CEX_ACCT_NUM) = rs!ACCT_NUM
'        lvarRecord(CEX_EXCEPTION_SYMBOL_CD) = Trim(rs!SYMBOL_CD)
'        lvarRecordArray(0, i) = lvarRecord(CEX_EXCEPTION_CD)
'        lvarRecordArray(1, i) = lvarRecord
'        rs.MoveNext
'        i = i + 1
'    Loop
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    'group specific
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, CEX_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build exception code group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeExceptionGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeExceptionGroup"
'End Function


'Private Function InitializeAvailableScheduleGroup() As Boolean
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    Dim lvarRecord(CSC_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'    Dim l_strSQL As String
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    l_strSQL = GetAvailScheduleCacheSQL()
'    Set rs = GetCacheRecords(l_strSQL)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Get available schedule records in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    If Not RecordsExist(rs) Then
'        InitializeAvailableScheduleGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    llngRecordCount = rs.RecordCount
'    ReDim lvarRecordArray(1, llngRecordCount - 1)
'    rs.MoveFirst
'    Do Until rs.EOF
'        lvarRecord(CSC_SCHED_CD) = CStr(rs!SCHED_CD)
'        lvarRecordArray(0, i) = lvarRecord(CSC_SCHED_CD)
'        lvarRecordArray(1, i) = lvarRecord
'        rs.MoveNext
'        i = i + 1
'    Loop
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, CSC_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build available schedule code group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeAvailableScheduleGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeAvailableScheduleGroup"
'End Function


'Private Function InitializeRotationGroup() As Boolean
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    Dim lvarRecord(CRC_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'    Dim l_strSQL As String
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    l_strSQL = GetRotationCacheSQL()
'    Set rs = GetCacheRecords(l_strSQL)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Get rotation week records in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    If Not RecordsExist(rs) Then
'        InitializeRotationGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    llngRecordCount = rs.RecordCount
'    ReDim lvarRecordArray(1, llngRecordCount - 1)
'    rs.MoveFirst
'    Do Until rs.EOF
'        lvarRecord(CRC_ROT_CD) = CStr(rs!ROT_CD)
'        lvarRecord(CRC_ROT_WEEKS) = CStr(rs!ROT_WEEKS)
'        lvarRecordArray(0, i) = lvarRecord(CRC_ROT_CD)
'        lvarRecordArray(1, i) = lvarRecord
'        rs.MoveNext
'        i = i + 1
'    Loop
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, CRC_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build rotation week group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeRotationGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeRotationGroup"
'End Function


'Private Function InitializeComponentCodeGroup() As Boolean
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    Dim lvarRecord(CCC_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'    Dim l_strSQL As String
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    l_strSQL = GetComponentCodeCacheSQL()
'    Set rs = GetCacheRecords(l_strSQL)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Get component code records in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    If Not RecordsExist(rs) Then
'        InitializeComponentCodeGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    llngRecordCount = rs.RecordCount
'    ReDim lvarRecordArray(1, llngRecordCount - 1)
'    rs.MoveFirst
'    Do Until rs.EOF
'        lvarRecord(CCC_COMP_CD) = Trim(rs!COMP_CD)
'        lvarRecord(CCC_COMP_TYPE) = Trim(rs!COMP_TYPE)
'        lvarRecordArray(0, i) = lvarRecord(CCC_COMP_CD)
'        lvarRecordArray(1, i) = lvarRecord
'        rs.MoveNext
'        i = i + 1
'    Loop
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, CCC_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build component code group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeComponentCodeGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeComponentCodeGroup"
'End Function


'//-----used by RecordCache-----//

Private Function IsLazyInitialized(ByVal v_strGroupName As String) As Boolean
' If group is not lazy initialized, then it is bulk initialized.
On Error GoTo eh
    Select Case Trim(v_strGroupName)
    Case CJB_GROUPNAME
        IsLazyInitialized = True
    Case Else
        IsLazyInitialized = False
    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsLazyInitialized"
End Function


Private Function GroupSupportsLoadRecord(ByVal v_strGroupName As String) As Boolean
'All group caches that support Invalidate must use parameterized SQL in corresponding GetxxxRecords (GetGroupRecords)
On Error GoTo eh
    Select Case Trim(v_strGroupName)
    Case CJB_GROUPNAME, CSC_GROUPNAME, CRC_GROUPNAME, CEX_GROUPNAME, _
            COE_GROUPNAME, COR_GROUPNAME, COP_GROUPNAME, CHS_GROUPNAME, _
            CLC_GROUPNAME, CUN_GROUPNAME, CUC_GROUPNAME, CCC_GROUPNAME, CEG_GROUPNAME
        GroupSupportsLoadRecord = True
    Case Else
        GroupSupportsLoadRecord = False
    End Select
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GroupSupportsLoadRecord"
End Function


Private Function OneRowForEachKey(ByVal v_strGroupName As String) As Boolean
'//-----normally each key/row is cached into a 1-dimensional array that represents each column-----//
'//-----union components will have multiple rows per key and will be cached into a 2-dimensional array-----//
On Error GoTo eh
    Select Case Trim(v_strGroupName)
    Case CUC_GROUPNAME
        OneRowForEachKey = False
    Case Else
        OneRowForEachKey = True
    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OneRowForEachKey"
End Function


'Private Function LoadRecord( _
'            ByVal v_strKey As String, _
'            ByVal v_strGroupName As String) As Variant
'
'On Error GoTo eh
'    Select Case Trim(v_strGroupName)
''//        Case CJB_GROUPNAME
''//            LoadJobRecord (v_strKey)
''        Case CSC_GROUPNAME
''            LoadScheduleRecord (v_strKey)
''        Case CRC_GROUPNAME
''            LoadRotationRecord (v_strKey)
''//        Case CEX_GROUPNAME
''//            LoadExceptionRecord (v_strKey)
''        Case COE_GROUPNAME
''            LoadOTExceptionRecord (v_strKey)
''        Case COR_GROUPNAME
''            LoadOTReasonRecord (v_strKey)
''        Case COP_GROUPNAME
''            LoadOTResponseRecord (v_strKey)
'    End Select
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "LoadRecord"
'End Function
'

'Private Sub LoadJobRecord( _
'                ByVal v_strJobNum As String)
'    Dim rs As ADODB.Recordset
'    Dim lvarRecord(CJB_LAST_SLOT) As Variant
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim llngRetVal As Long
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    Set rs = GetJobBasic(v_strJobNum)
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    If rs.BOF And rs.EOF Then
'        llngRetVal = _
'            spmWrapper.Add(v_strJobNum, CRECORDNOTFOUND, CJB_GROUPNAME)
'    Else
'        lvarRecord(CJB_JOB_NUM) = Trim(rs!JOB_NUM)
'        lvarRecord(CJB_JOB_DESC) = Trim(rs!JOB_DESC)
'        lvarRecord(CJB_CRAFT_NUM) = Trim(rs!CRAFT_NUM)
'        lvarRecord(CJB_OT_TYPE) = Trim(rs!OT_TYPE)
'        lvarRecord(CJB_EFF_DT) = rs!EFFDT
'        lvarRecord(CJB_EFF_STATUS) = Trim(rs!EFF_STATUS)
'        lvarRecord(CJB_UNION_CD) = Trim(rs!UNION_CD)
'        llngRetVal = _
'            spmWrapper.Add(v_strJobNum, lvarRecord, CJB_GROUPNAME)
'    End If
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Load job record in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    Exit Sub
'eh:
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "LoadJobRecord"
'End Sub
'
'
'Private Sub LoadScheduleRecord( _
'                ByVal v_strSchedCode As String)
'    Dim rs As ADODB.Recordset
'    Dim l_strSchedCode As String
'    Dim lvarRecord(CSC_LAST_SLOT) As Variant
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim llngRetVal As Long
'    Dim l_strSQL As String
'
'On Error GoTo eh
'    l_strSchedCode = Trim(v_strSchedCode)
'    l_strSQL = GetAvailScheduleCacheSQL(l_strSchedCode)
'    Set rs = GetCacheRecords(l_strSQL)
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    If Not RecordsExist(rs) Then
'        llngRetVal = _
'            spmWrapper.Remove(l_strSchedCode, CSC_GROUPNAME)
'    Else
'        lvarRecord(CSC_SCHED_CD) = l_strSchedCode
'        llngRetVal = _
'            spmWrapper.Add(l_strSchedCode, lvarRecord, CSC_GROUPNAME)
'    End If
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    Exit Sub
'
'eh:
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "LoadScheduleRecord"
'End Sub

'
'Private Sub LoadRotationRecord( _
'                ByVal v_strRotCd As String)
'    Dim l_strRotCd As String
'    Dim rs As ADODB.Recordset
'    Dim lvarRecord(CRC_LAST_SLOT) As Variant
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim llngRetVal As Long
'    Dim l_strSQL As String
'
'On Error GoTo eh
'    l_strRotCd = Trim(v_strRotCd)
'    l_strSQL = GetRotationCacheSQL(l_strRotCd)
'    Set rs = GetCacheRecords(l_strSQL)
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    If Not RecordsExist(rs) Then
'        llngRetVal = _
'            spmWrapper.Remove(l_strRotCd, CRC_GROUPNAME)
'    Else
'        lvarRecord(CRC_ROT_CD) = rs!ROT_CD
'        lvarRecord(CRC_ROT_WEEKS) = rs!ROT_WEEKS
'        llngRetVal = _
'            spmWrapper.Add(l_strRotCd, lvarRecord, CRC_GROUPNAME)
'    End If
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    Exit Sub
'
'eh:
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "LoadRotationRecord"
'End Sub


'Private Sub LoadExceptionRecord( _
'                ByVal v_strExcCd As String)
'    Dim l_strExcCd As String
'    Dim rs As ADODB.Recordset
'    Dim lvarRecord(CEX_LAST_SLOT) As Variant
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim llngRetVal As Long
'    Dim l_strSQL As String
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    l_strExcCd = Trim(v_strExcCd)
'    l_strSQL = GetExceptionCacheSQL(l_strExcCd)
'    Set rs = GetCacheRecords(l_strSQL)
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    If Not RecordsExist(rs) Then
'        llngRetVal = _
'            spmWrapper.Remove(l_strExcCd, CEX_GROUPNAME)
'    Else
'        lvarRecord(CEX_EXCEPTION_CD) = Trim(rs!EXCEPTION_CD)
'        lvarRecord(CEX_EXCEPTION_DESC) = Trim(rs!EXCEPTION_DESC)
'        lvarRecord(CEX_EXCEPTION_TYPE) = UCase(Trim(rs!EXCEPTION_TYPE))
'        lvarRecord(CEX_CONTRACT_EARN_CD) = Trim(rs!CONTRACT_EARN_CD)
'        lvarRecord(CEX_EXEC_EARN_CD) = Trim(rs!EXEC_EARN_CD)
'        lvarRecord(CEX_SYMBOL_TIME_IND) = Trim(rs!SYMBOL_TIME_IND)
'        lvarRecord(CEX_EXCEPTION_BANK_CD) = Trim(rs!EXCEPTION_BANK_CD)
'        lvarRecord(CEX_INELIGIBILITY_IND) = Trim(rs!INELIGIBILITY_IND)
'        lvarRecord(CEX_OTHERPAY_IND) = UCase(Trim(rs!OTHERPAY_IND))
'        lvarRecord(CEX_CASE_GROUP) = Trim(rs!CASE_GROUP)
'        lvarRecord(CEX_SCHED_IND) = Trim(rs!SCHED_IND)
'        lvarRecord(CEX_FMLA_IND) = Trim(rs!FMLA_IND)
'        lvarRecord(CEX_WORK_REST_IND) = Trim(rs!WORK_REST_IND)
'        lvarRecord(CEX_SCHED_CHG_IND) = Trim(rs!SCHED_CHG_IND)
'        lvarRecord(CEX_ACCT_NUM) = rs!ACCT_NUM
'        lvarRecord(CEX_EXCEPTION_SYMBOL_CD) = Trim(rs!SYMBOL_CD)
'        llngRetVal = _
'            spmWrapper.Add(l_strExcCd, lvarRecord, CEX_GROUPNAME)
'    End If
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "LoadExceptionRecord in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    Exit Sub
'
'eh:
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "LoadExceptionRecord"
'End Sub


'Private Sub LoadOTReasonRecord( _
'                ByVal v_strOTRsnCd As String)
'    Dim l_strOTRsnCd As String
'    Dim rs As ADODB.Recordset
'    Dim lvarRecord(COR_LAST_SLOT) As Variant
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim llngRetVal As Long
'    Dim l_strSQL As String
'
'On Error GoTo eh
'    l_strOTRsnCd = Trim(v_strOTRsnCd)
'    l_strSQL = GetOTReasonCacheSQL(l_strOTRsnCd)
'    Set rs = GetCacheRecords(l_strSQL)
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    If Not RecordsExist(rs) Then
'        llngRetVal = _
'            spmWrapper.Remove(l_strOTRsnCd, COR_GROUPNAME)
'    Else
'        lvarRecord(COR_OT_REASON_CD) = Trim(rs!OT_REASON_CD)
'        lvarRecord(COR_OT_REASON_DESC) = Trim(rs!OT_REASON_DESC)
'        llngRetVal = _
'            spmWrapper.Add(l_strOTRsnCd, lvarRecord, COR_GROUPNAME)
'    End If
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    Exit Sub
'
'eh:
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "LoadOTReasonRecord"
'End Sub


'Private Sub LoadOTResponseRecord( _
'                ByVal v_strOTRspCd As String)
'    Dim l_strOTRspCd As String
'    Dim rs As ADODB.Recordset
'    Dim lvarRecord(COP_LAST_SLOT) As Variant
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim llngRetVal As Long
'    Dim l_strSQL As String
'
'On Error GoTo eh
'    l_strOTRspCd = Trim(v_strOTRspCd)
'    l_strSQL = GetOTResponseCacheSQL(l_strOTRspCd)
'    Set rs = GetCacheRecords(l_strSQL)
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    If Not RecordsExist(rs) Then
'        llngRetVal = _
'            spmWrapper.Remove(l_strOTRspCd, COP_GROUPNAME)
'    Else
'        lvarRecord(COP_OT_RESP_CD) = Trim(rs!OT_RESP_CD)
'        lvarRecord(COP_RESPONSE_DESC) = Trim(rs!RESPONSE_DESC)
'        lvarRecord(COP_OT_CHRG_IND) = Trim(rs!OT_CHRG_IND)
'        lvarRecord(COP_OT_WORKED_IND) = Trim(rs!OT_WORKED_IND)
'        llngRetVal = _
'            spmWrapper.Add(l_strOTRspCd, lvarRecord, COP_GROUPNAME)
'    End If
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    Exit Sub
'
'eh:
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "LoadOTResponseRecord"
'End Sub


'Private Sub LoadOTExceptionRecord( _
'                ByVal v_strOTExcCd As String)
'    Dim l_strOTExcCd As String
'    Dim rs As ADODB.Recordset
'    Dim lvarRecord(COE_LAST_SLOT) As Variant
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim llngRetVal As Long
'    Dim l_strSQL As String
'
'On Error GoTo eh
'    l_strOTExcCd = Trim(v_strOTExcCd)
'    l_strSQL = GetOTExceptionCacheSQL(l_strOTExcCd)
'    Set rs = GetCacheRecords(l_strSQL)
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    If Not RecordsExist(rs) Then
'        llngRetVal = _
'            spmWrapper.Remove(l_strOTExcCd, COE_GROUPNAME)
'    Else
'        lvarRecord(COE_EXCEPTION_CD) = Trim(rs!EXCEPTION_CD)
'        lvarRecord(COE_EXCEPTION_DESC) = Trim(rs!EXCEPTION_DESC)
'        lvarRecord(COE_FACTOR) = Trim(CStr(rs!Factor))
'        lvarRecord(COE_EXEC_CONTRACT_IND) = Trim(rs!EXEC_CONTRACT_IND)
'        lvarRecord(COE_EARNINGS_CD) = Trim(rs!EARNINGS_CD)
'        llngRetVal = _
'            spmWrapper.Add(l_strOTExcCd, lvarRecord, COE_GROUPNAME)
'    End If
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    Exit Sub
'
'eh:
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "LoadOTExceptionRecord"
'End Sub


Private Function GetCacheRecords(ByVal v_strSQL As String, Optional eLockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long
    
On Error GoTo eh
    
    #If TRDEBUG Then
        stTime = GetTickCount()
    #End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open v_strSQL, l_objCN, adOpenStatic, eLockType
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    
    If l_objRS.BOF And l_objRS.EOF Then
        Set l_objRS = Nothing
    End If
    
    Set GetCacheRecords = l_objRS
    
    Set l_objRS = Nothing
    Set l_objCN = Nothing
    Set l_objCM = Nothing
    
    #If TRDEBUG Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        Debug.Print "GetCacheRecords: " & CStr(dfTime / 1000)
    #End If
    
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCacheRecords", , , , DescArgsFrom(v_strSQL), True
End Function


'Private Function GetAvailScheduleCacheSQL( _
'                    Optional ByVal v_strSchedCd As String) As String
'
'    Dim l_strSQL As String
'    Dim l_strSchedCd As String
'
'On Error GoTo eh
'
'    l_strSchedCd = Trim(v_strSchedCd)
'    l_strSQL = "SELECT SCHED_CD" & _
'               " FROM METHDBA.TTR_AVAIL_SCHED"
'    If l_strSchedCd = "" Then
'        l_strSQL = l_strSQL & " ORDER BY SCHED_CD"
'    Else
'        l_strSQL = l_strSQL & " WHERE SCHED_CD = " & l_strSchedCd
'    End If
'    GetAvailScheduleCacheSQL = l_strSQL
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetAvailScheduleCacheSQL"
'End Function

'
'Private Function GetRotationCacheSQL( _
'                    Optional ByVal v_strRotCd As String) As String
'
'    Dim l_strSQL As String
'    Dim l_strRotCd As String
'
'On Error GoTo eh
'
'    l_strRotCd = Trim(v_strRotCd)
'    l_strSQL = "SELECT ROT_CD, MAX(ROT_WEEK) AS ROT_WEEKS" & _
'               " FROM METHDBA.TTR_ROTATION_SCHED"
'
'    If l_strRotCd = "" Then
'        l_strSQL = l_strSQL & " GROUP BY ROT_CD"
'    Else
'        l_strSQL = l_strSQL & " WHERE ROT_CD = " & l_strRotCd _
'                            & " GROUP BY ROT_CD"
'    End If
'    GetRotationCacheSQL = l_strSQL
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetRotationCacheSQL"
'End Function

'//-----not used any more by cache-----//
'Private Function GetExceptionCacheSQL( _
'                    Optional ByVal v_strExcCd As String) As String
'
'    Dim l_strSQL As String
'    Dim l_strExcCd As String
'
'On Error GoTo eh
'    l_strExcCd = Trim(v_strExcCd)
'    l_strSQL = "SELECT EXCEPTION_CD, " _
'         & " EXCEPTION_DESC, EXCEPTION_TYPE, " _
'         & " CONTRACT_EARN_CD, EXEC_EARN_CD, SYMBOL_TIME_IND, " _
'         & " EXCEPTION_BANK_CD, INELIGIBILITY_IND, OTHERPAY_IND, " _
'         & " CASE_GROUP, SCHED_IND, 'N' as FMLA_IND, " _
'         & " WORK_REST_IND, SCHED_CHG_IND, ACCT_NUM, SYMBOL_CD " _
'         & " FROM METHDBA.TTR_EXCEPTION_CDS "
'
'    If l_strExcCd = "" Then
'        l_strSQL = l_strSQL & " ORDER BY EXCEPTION_CD"
'    Else
'        l_strSQL = l_strSQL & " WHERE EXCEPTION_CD = " & "'" & l_strExcCd & "'"
'    End If
'    GetExceptionCacheSQL = l_strSQL
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetExceptionCacheSQL"
'End Function

'
'Private Function GetOTReasonCacheSQL( _
'                    Optional ByVal v_strOTRsnCd As String) As String
'
'    Dim l_strSQL As String
'    Dim l_strOTRsnCd As String
'
'On Error GoTo eh
'    l_strOTRsnCd = Trim(v_strOTRsnCd)
'    l_strSQL = "SELECT OT_REASON_CD, OT_REASON_DESC" _
'             & " FROM METHDBA.TTR_OT_REASONS"
'    If l_strOTRsnCd = "" Then
'        l_strSQL = l_strSQL & " ORDER BY OT_REASON_CD"
'    Else
'        l_strSQL = l_strSQL & " WHERE OT_REASON_CD = " & "'" & l_strOTRsnCd & "'"
'    End If
'    GetOTReasonCacheSQL = l_strSQL
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetOTReasonCacheSQL"
'End Function

'
'Private Function GetOTResponseCacheSQL( _
'                    Optional ByVal v_strOTRspCd As String) As String
'
'    Dim l_strSQL As String
'    Dim l_strOTRspCd As String
'
'On Error GoTo eh
'    l_strOTRspCd = Trim(v_strOTRspCd)
'    l_strSQL = "SELECT OT_RESP_CD, RESPONSE_DESC, OT_CHRG_IND, OT_WORKED_IND " & _
'               " FROM METHDBA.TTR_OT_RESPONSES"
'    If l_strOTRspCd = "" Then
'        l_strSQL = l_strSQL & " ORDER BY OT_RESP_CD"
'    Else
'        l_strSQL = l_strSQL & " WHERE OT_RESP_CD = " & "'" & l_strOTRspCd & "'"
'    End If
'    GetOTResponseCacheSQL = l_strSQL
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetOTResponseCacheSQL"
'End Function

'
'Private Function GetOTExceptionCacheSQL( _
'                    ByVal v_strOTExcCd As String) As String
'' Only use this for a single record.
'    Dim l_strSQL As String
'    Dim l_strOTExcCd As String
'
'On Error GoTo eh
'    l_strOTExcCd = Trim(v_strOTExcCd)
'    Select Case l_strOTExcCd
'        Case "ON1"
'            l_strSQL = "SELECT EXCEPTION_CD, EXCEPTION_DESC, " & _
'                    " 1.0 AS FACTOR,  'C' AS EXEC_CONTRACT_IND, '99' AS EARNINGS_CD" & _
'                " FROM METHDBA.TTR_EXCEPTION_CDS" & _
'                " WHERE EXCEPTION_CD = 'ON1'"
'        Case "ON2"
'            l_strSQL = "SELECT EXCEPTION_CD, EXCEPTION_DESC, " & _
'                    " 2.0 AS FACTOR,  'C' AS EXEC_CONTRACT_IND, '99' AS EARNINGS_CD" & _
'                " FROM METHDBA.TTR_EXCEPTION_CDS" & _
'                " WHERE EXCEPTION_CD = 'ON2'"
'        Case "ON7"
'            l_strSQL = "SELECT EXCEPTION_CD, EXCEPTION_DESC, " & _
'                    " 1.75 AS FACTOR,  'C' AS EXEC_CONTRACT_IND, '99' AS EARNINGS_CD" & _
'                " FROM METHDBA.TTR_EXCEPTION_CDS" & _
'                " WHERE EXCEPTION_CD = 'ON7'"
'        Case Else
'            l_strSQL = "SELECT EXCEPTION_CD, EXCEPTION_DESC," & _
'                    " FACTOR, EXEC_CONTRACT_IND," & _
'                    " EXEC_EARN_CD As EARNINGS_CD" & _
'                " FROM METHDBA.TTR_EXCEPTION_CDS," & _
'                "      METHDBA.TTR_EARN_CODES" & _
'                " WHERE EARNINGS_CD = EXEC_EARN_CD" & _
'                " AND EXCEPTION_CD = " & "'" & l_strOTExcCd & "'"
'    End Select
'
'    GetOTExceptionCacheSQL = l_strSQL
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetOTExceptionCacheSQL"
'End Function


'Private Function GetComponentCodeCacheSQL() As String
'
'    Dim l_strSQL As String
'
'On Error GoTo eh
'    l_strSQL = "SELECT COMP_CD, COMP_TYPE " & _
'               " FROM METHDBA.TTR_COMP_CD" & _
'               " ORDER BY COMP_CD"
'    GetComponentCodeCacheSQL = l_strSQL
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetComponentCodeCacheSQL"
'End Function

'*********************************************************
' End Private Caching methods
'*********************************************************



'//-----not used any more by cache-----//
Public Function GetEffectiveUnionComponents() As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetEffectiveUnionComponents_EH


    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")

    Set l_objCM = Nothing

    l_objCN.Open

    '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
    l_strSQL = "SELECT DISTINCT CAST(LTRIM(RTRIM(A.UNION_CD)) AS VARCHAR(3)) AS UNION_CD, " _
             & " CAST(LTRIM(RTRIM(A.COMP_CD)) AS VARCHAR(3)) AS COMP_CD," _
             & " B.COMP_DESC, A.EFF_DATE, A.COMP_TYPE, A.COMP_AMT " _
             & "FROM METHDBA.TTR_UNION_COMP_AMT A, METHDBA.TTR_COMP_CD B " _
             & "WHERE A.COMP_CD = B.COMP_CD " _
             & "AND A.COMP_TYPE = B.COMP_TYPE " _
             & "AND B.COMP_CD not in " _
             & " (select RG_VALUE_CD " _
             & "   from METHDBA.TTR_RG_GROUP " _
             & "   where RG_CD = 85) " _
             & "ORDER BY UNION_CD, COMP_CD, A.EFF_DATE "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    If l_objRS.BOF And l_objRS.EOF Then
        Set GetEffectiveUnionComponents = Nothing
    Else
        Set GetEffectiveUnionComponents = l_objRS
    End If

    Set l_objRS = Nothing

    Exit Function

GetEffectiveUnionComponents_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEffectiveUnionComponents", , , , , True

End Function

Public Function GetUnionComponents() As ADODB.Recordset

    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetUnionComponents_EH

    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")

    Set l_objCM = Nothing

    l_objCN.Open

'//TNV
    '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
     l_strSQL = "SELECT DISTINCT CAST(LTRIM(RTRIM(A.UNION_CD)) AS VARCHAR(3)) AS UNION_CD," _
             & "CAST(LTRIM(RTRIM(A.COMP_CD)) AS VARCHAR(3)) AS COMP_CD," _
             & "B.COMP_DESC, A.COMP_TYPE " _
             & "FROM METHDBA.TTR_UNION_COMP_AMT A, METHDBA.TTR_COMP_CD B " _
             & "Where A.COMP_CD = B.COMP_CD " _
             & "AND A.COMP_TYPE = B.COMP_TYPE " _
             & "AND B.COMP_CD not in " _
             & " (select RG_VALUE_CD " _
             & "   from METHDBA.TTR_RG_GROUP " _
             & "   where RG_CD = 85) " _
             & "ORDER BY UNION_CD, COMP_CD "
'//TNV

'    l_strSQL = "SELECT DISTINCT A.UNION_CD, A.COMP_CD, B.COMP_DESC, A.COMP_TYPE " _
'             & "FROM METHDBA.TTR_UNION_COMP_AMT A, METHDBA.TTR_COMP_CD B " _
'             & "Where A.COMP_CD = B.COMP_CD " _
'             & "AND A.COMP_TYPE = B.COMP_TYPE " _
'             & "ORDER BY A.UNION_CD, A.COMP_CD "

'  Spoke with Prem and Greg.  Both felt that the self join served no purpose in the implementation of this operation
'    l_strSQL = "SELECT DISTINCT A.UNION_CD, A.COMP_CD, B.COMP_DESC, A.COMP_TYPE " _
'             & "FROM METHDBA.TTR_UNION_COMP_AMT A, METHDBA.TTR_COMP_CD B " _
'             & "Where A.COMP_CD = B.COMP_CD " _
'             & "AND A.COMP_TYPE = B.COMP_TYPE " _
'             & "AND A.EFF_DATE IN (SELECT MAX(C.EFF_DATE) " _
'             & "FROM METHDBA.TTR_UNION_COMP_AMT C " _
'             & "Where C.COMP_CD = A.COMP_CD " _
'             & "AND C.COMP_TYPE = A.COMP_TYPE " _
'             & "AND C.UNION_CD = A.UNION_CD) " _
'             & "ORDER BY A.UNION_CD, A.COMP_CD "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    If l_objRS.BOF And l_objRS.EOF Then
        Set GetUnionComponents = Nothing
    Else
        Set GetUnionComponents = l_objRS
    End If

    Set l_objRS = Nothing

    Exit Function

GetUnionComponents_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetUnionComponents", , , , , True

End Function


Public Function GetActualTimeHSCutoffDate( _
            ByVal v_strHSKey As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strHSKey, CHS_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    GetActualTimeHSCutoffDate = l_varRecord(CHS_CUTOFF_DT)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTimeHSCutOffDate"
End Function


'Private Function InitializeHSCutOffDateGroup() As Boolean
'
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    Dim lvarRecord(CHS_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'    Dim l_strHSCutOffDate As String
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'    l_strHSCutOffDate = GetHistoryCutoffDate
'    If l_strHSCutOffDate = "" Then
'        InitializeHSCutOffDateGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    ReDim lvarRecordArray(1, CHS_LAST_SLOT)
'
'    lvarRecord(CHS_CUTOFF_CD) = CKEYCUTOFFDATE
'    lvarRecord(CHS_CUTOFF_DT) = CDate(Trim(l_strHSCutOffDate))
'    lvarRecordArray(0, 0) = lvarRecord(CHS_CUTOFF_CD)
'    lvarRecordArray(1, 0) = lvarRecord
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, CHS_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build History CutOff date group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeHSCutOffDateGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeHSCutOffDateGroup"
'End Function
    

''Get TTR_ACTUAL_TIME_HS cutoff Date
'Private Function GetHistoryCutoffDate() As String
'
'    Dim l_strSQL As String
'    Dim l_objCN As ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'
'    On Error GoTo GetHistoryCutoffDate_EH
'    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
'    Set l_objCN = New ADODB.Connection
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open
'
'    l_strSQL = "SELECT PARM_VALUE " _
'             & "FROM METHDBA.TTR_SYSPARMS " _
'             & "WHERE PARM_NUM = 18 AND PARM_SUFFIX = 1"
'
'    Set l_objRS = New ADODB.Recordset
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'
'    If l_objRS.BOF And l_objRS.EOF Then
'        GetHistoryCutoffDate = ""
'    Else
'        GetHistoryCutoffDate = l_objRS.Fields(0)
'    End If
'
'    Set l_objRS = Nothing
'    Exit Function
'
'GetHistoryCutoffDate_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetHistoryCutoffDate", , , , DescArgsFrom(l_strSQL), True
'
'End Function



Public Function ExceptionSymbolCd( _
            ByVal v_strExcCd As String) As String
' The symbol code is used to link exceptions together that
' have the same behaivor (rules) associated to them
' For example, there a few exceptions for Holiday (H,HR,HD,...)
'   they should all have the same rules applied to them.
'  The symbol_cd for Holiday is 'H'
'  The symbol_cd for Vacation is 'V'
' There are others, but at this time all is needed is
' Holiday and Vacation


    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionSymbolCd = l_varRecord(CEX_EXCEPTION_SYMBOL_CD)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionSymbolCd"
End Function


Public Function GetAvailableScheduleDetails() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long
    
On Error GoTo GetAvailableScheduleDetails_EH
    
    #If TRDEBUG Then
        stTime = GetTickCount()
    #End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT  SCHED_CD, SCHED_DESC, SCHED_SHIFT_CD "
    l_strSQL = l_strSQL & _
                ", SUN_ACT_DAY, SUN_SHIFT_CD, SUN_START_TIME" & _
                ", SUN_END_TIME, SUN_LUNCH_START, SUN_LUNCH_END" & _
                ", SUN_STR_HRS, SUN_OT_START_TIME, SUN_OT_END_TIME" & _
                ", SUN_OT_HOURS, SUN_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", MON_ACT_DAY, MON_SHIFT_CD, MON_START_TIME" & _
                ", MON_END_TIME, MON_LUNCH_START, MON_LUNCH_END" & _
                ", MON_STR_HRS, MON_OT_START_TIME, MON_OT_END_TIME" & _
                ", MON_OT_HOURS, MON_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", TUE_ACT_DAY, TUE_SHIFT_CD, TUE_START_TIME" & _
                ", TUE_END_TIME, TUE_LUNCH_START, TUE_LUNCH_END" & _
                ", TUE_STR_HRS, TUE_OT_START_TIME, TUE_OT_END_TIME" & _
                ", TUE_OT_HOURS, TUE_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", WED_ACT_DAY, WED_SHIFT_CD, WED_START_TIME" & _
                ", WED_END_TIME, WED_LUNCH_START, WED_LUNCH_END" & _
                ", WED_STR_HRS, WED_OT_START_TIME, WED_OT_END_TIME" & _
                ", WED_OT_HOURS, WED_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", THU_ACT_DAY, THU_SHIFT_CD, THU_START_TIME" & _
                ", THU_END_TIME, THU_LUNCH_START, THU_LUNCH_END" & _
                ", THU_STR_HRS, THU_OT_START_TIME, THU_OT_END_TIME" & _
                ", THU_OT_HOURS, THU_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", FRI_ACT_DAY, FRI_SHIFT_CD, FRI_START_TIME" & _
                ", FRI_END_TIME, FRI_LUNCH_START, FRI_LUNCH_END" & _
                ", FRI_STR_HRS, FRI_OT_START_TIME, FRI_OT_END_TIME" & _
                ", FRI_OT_HOURS, FRI_COMP_SHIFT_CD"
    l_strSQL = l_strSQL & _
                ", SAT_ACT_DAY, SAT_SHIFT_CD, SAT_START_TIME" & _
                ", SAT_END_TIME, SAT_LUNCH_START, SAT_LUNCH_END" & _
                ", SAT_STR_HRS, SAT_OT_START_TIME, SAT_OT_END_TIME" & _
                ", SAT_OT_HOURS, SAT_COMP_SHIFT_CD, LAST_MOD_EMP_NUM,Char(TIME_STAMP) as TIME_STAMP "
    l_strSQL = l_strSQL & _
                " FROM       METHDBA.TTR_AVAIL_SCHED" & _
                " WHERE      SCHED_CD > 0 " & _
                " ORDER BY   SCHED_CD"


    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetAvailableScheduleDetails = Nothing
    Else
        Set GetAvailableScheduleDetails = l_objRS
    End If
    Set l_objRS = Nothing
    
    #If TRDEBUG Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        Debug.Print "GetAvailableScheduleDetails: " & CStr(dfTime / 1000)
    #End If
    
    Exit Function

GetAvailableScheduleDetails_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAvailableScheduleDetails", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetAvailableScheduleHeader() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetAvailableScheduleHeader_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT SCHED_CD, SCHED_DESC, SUN_SHIFT_CD" & _
                ", MON_SHIFT_CD, TUE_SHIFT_CD, WED_SHIFT_CD" & _
                ", THU_SHIFT_CD, FRI_SHIFT_CD, SAT_SHIFT_CD" & _
                ", 'L' AS RECORD_STATUS, LAST_MOD_EMP_NUM, Char(TIME_STAMP) as TIME_STAMP" & _
                " FROM       METHDBA.TTR_AVAIL_SCHED" & _
                " WHERE      SCHED_CD > 0 " & _
                " ORDER BY   SCHED_CD"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetAvailableScheduleHeader = Nothing
    Else
        Set GetAvailableScheduleHeader = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetAvailableScheduleHeader_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAvailableScheduleHeader", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetAllActivityList() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetAllActivityList_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT ACTIVITY_CD, ACTIVITY_DESC " & _
                " FROM       METHDBA.TTR_ACTIVITY"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetAllActivityList = Nothing
    Else
        Set GetAllActivityList = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetAllActivityList_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllActivityList", , , , DescArgsFrom(l_strSQL), True
End Function


Public Function ValidateScheduleCode(ByVal v_lScheduleCode As Long) As Boolean

On Error GoTo eh

    ValidateScheduleCode = ScheduleCodeIsValid(v_lScheduleCode)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateScheduleCode"
End Function

Public Function GetProjectInfo _
                (Optional ByVal v_intProjectId As Integer, _
                 Optional ByVal v_intPPMInd As String _
                ) As ADODB.Recordset


    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetProjectInfo_EH
    
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = RetrieveProjectInfo(v_intProjectId, v_intPPMInd)

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetProjectInfo = Nothing
    Else
        Set GetProjectInfo = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetProjectInfo_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetProjectInfo", , , , DescArgsFrom(l_strSQL), True
    
End Function

Private Function RetrieveProjectInfo( _
        Optional ByVal v_intProjectId As Integer, _
        Optional ByVal v_intPPMInd As String) As String

    Dim strSQL As String
    Dim errArgs As New Collection
    Dim l_objRS As ADODB.Recordset
    
    On Error GoTo RetrieveProjectInfo_EH
    
 
    If v_intProjectId = 0 Then
        If v_intPPMInd = "" Then
            strSQL = "SELECT PROJECT_ID, PROJECT_DESC, TEMPLATE_NUM, EFF_START_DATE, " & _
                " EFF_END_DATE, PROJECT_LOC_CD, PPM_IND, TRAVEL_PAY_IND, " & _
                " FUNCTION_NUM, LAST_MOD_EMP_NUM , char(TIME_STAMP) AS TIME_STAMP, ORD_SET_ID, " & _
                "  COUNT_TRAINING, COUNT_VACATION " & _
                " FROM       METHDBA.TTR_PROJECTS"
        Else
            strSQL = "SELECT PROJECT_ID, PROJECT_DESC, TEMPLATE_NUM, EFF_START_DATE, " & _
                " EFF_END_DATE, PROJECT_LOC_CD, PPM_IND, TRAVEL_PAY_IND, " & _
                " FUNCTION_NUM, LAST_MOD_EMP_NUM , CHAR(TIME_STAMP) AS TIME_STAMP, ORD_SET_ID, " & _
                "  COUNT_TRAINING, COUNT_VACATION " & _
                " FROM       METHDBA.TTR_PROJECTS" & _
                " WHERE PPM_IND = '" & UCase(Trim(v_intPPMInd)) & "'"
        End If
    Else
        If v_intPPMInd = "" Then
            strSQL = "SELECT PROJECT_ID, PROJECT_DESC, TEMPLATE_NUM, EFF_START_DATE, " & _
                " EFF_END_DATE, PROJECT_LOC_CD, PPM_IND, TRAVEL_PAY_IND, " & _
                " FUNCTION_NUM, LAST_MOD_EMP_NUM , CHAR(TIME_STAMP) AS TIME_STAMP, ORD_SET_ID, " & _
                "  COUNT_TRAINING, COUNT_VACATION " & _
                " FROM       METHDBA.TTR_PROJECTS" & _
                " WHERE PROJECT_ID = " & v_intProjectId
        Else
            strSQL = "SELECT PROJECT_ID, PROJECT_DESC, TEMPLATE_NUM, EFF_START_DATE, " & _
                " EFF_END_DATE, PROJECT_LOC_CD, PPM_IND, TRAVEL_PAY_IND, " & _
                " FUNCTION_NUM, LAST_MOD_EMP_NUM , CHAR(TIME_STAMP) AS TIME_STAMP, ORD_SET_ID, " & _
                "  COUNT_TRAINING, COUNT_VACATION " & _
                " FROM       METHDBA.TTR_PROJECTS" & _
                " WHERE PROJECT_ID = " & v_intProjectId & _
                " AND PPM_IND = '" & UCase(Trim(v_intPPMInd)) & "'"
        End If
    End If
    
    RetrieveProjectInfo = strSQL

Exit Function
RetrieveProjectInfo_EH:
    
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "RetrieveProjectInfo", , True, , errArgs, True

End Function


Public Function GetLocationsAssociatedToDepartment( _
        ByRef v_colDeptNums As Collection) As ADODB.Recordset

    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
   
    l_strSQL = _
        "SELECT A.DEPT_NUM, A.LOC_CD, B.LOC_DESC,FUNCTION_NUM, DIVISION_NUM " _
        & "FROM METHDBA.TTR_ASSOCIATIONS A " _
        & ",  METHDBA.TTR_LOCATIONS B " _
        & "WHERE A.WG_NUM = 0 " _
        & " AND A.DEPT_NUM IN " _
        & AsDelimitedString(v_colDeptNums) _
        & " AND A.LOC_CD = B.LOC_CD "
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetLocationsAssociatedToDepartment = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetLocationsAssociatedToDepartment", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetDivisionAndFunctionAssociatedToDepartmentAndLocation( _
                                            ByVal v_strDeptNum As String, _
                                            ByVal v_intLocCd As Integer, _
                                            ByRef r_strDivisionNum As String, _
                                            ByRef r_strFunctionNum As String)
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
   
    l_strSQL = _
                "SELECT DIVISION_NUM, FUNCTION_NUM" & _
                " FROM METHDBA.TTR_ASSOCIATIONS" & _
                " WHERE DEPT_NUM = '" & v_strDeptNum & "'" & _
                " AND LOC_CD = " & v_intLocCd & _
                " AND WG_NUM = 0"
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    
    If RecordsExist(l_objRS) Then
        r_strDivisionNum = l_objRS.Fields("DIVISION_NUM").Value
        r_strFunctionNum = l_objRS.Fields("FUNCTION_NUM").Value
    Else
        r_strDivisionNum = ""
        r_strFunctionNum = ""
    End If
       
    Set l_objRS = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDivisionAndFunctionAssociatedToDepartmentAndLocation", , , , DescArgsFrom(l_strSQL), True
End Function


Public Function GetAllProjects( _
            ByVal v_datEffective As Date _
            ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open

    l_strSQL = "SELECT PROJECT_ID, PROJECT_DESC " _
             & "FROM METHDBA.TTR_PROJECTS " _
             & "WHERE PPM_IND = 'Y' " _
             & "AND EFF_END_DATE >= '" & Format(v_datEffective, "mm/dd/yyyy") & "' " _
             & "ORDER BY PROJECT_ID "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetAllProjects = l_objRS
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllProjects", , , , DescArgsFrom(l_strSQL), True
End Function


Public Function GetAllScheduleCodes() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    
On Error GoTo GetAllScheduleCodes_eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open

    '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
    l_strSQL = "SELECT sched_cd, sched_desc " & _
                    " FROM methdba.ttr_avail_sched  " & _
                    " ORDER BY sched_cd "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetAllScheduleCodes = l_objRS
    Exit Function

GetAllScheduleCodes_eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllScheduleCodes", , , , DescArgsFrom(l_strSQL), True
End Function


'Added 5/09/2002 per Prem's request
Public Function GetExceptionGroupForColoring() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetExceptionGroupForColoring_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT EG_SUB_CD, EXCEPTION_CD " & _
                "       , EXC_SEQ_NUM" & _
                " FROM  METHDBA.TTR_EXC_GROUP" & _
                " WHERE EG_CD = 2" & _
                " ORDER BY  EG_SUB_CD"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetExceptionGroupForColoring = Nothing
    Else
        Set GetExceptionGroupForColoring = l_objRS
    End If
    Set l_objRS = Nothing
    Exit Function

GetExceptionGroupForColoring_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetExceptionGroupForColoring", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function ExceptionSchedInd( _
            ByVal v_strExcCd As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionSchedInd = l_varRecord(CEX_SCHED_IND)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionSchedInd"
End Function


Public Function IsExceptionInGroup( _
    ByVal v_lngGroupCode As Long, _
    ByVal v_strExceptionCode As String, _
    ByVal v_strSubGroupCode As String _
    ) As Boolean

    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
'//    Dim lstrKey As String
            
On Error GoTo eh
'//    lstrKey = ExceptionGroupKey(v_lngGroupCode, v_strExceptionCode, v_strSubGroupCode)
    l_varRecord = GetRecord(Array(v_lngGroupCode, v_strExceptionCode, v_strSubGroupCode), CEG_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        IsExceptionInGroup = False
    Else
        IsExceptionInGroup = True
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsExceptionInGroup"
End Function

''//-----Each field in this compound key needs to match the corresponding column defined size in the cached recordset-----//
'Private Function ExceptionGroupKey( _
'    ByVal v_lngGroupCode As Long, _
'    ByVal v_strExceptionCode As String, _
'    ByVal v_strSubGroupCode As String _
'    ) As String
'
'
'    Dim lstrResult As String * 17
'    Dim l_str As String
'
'On Error GoTo eh
'    lstrResult = Space(17)
'    l_str = Trim(str(v_lngGroupCode))
'    Mid(lstrResult, 1, 11) = String(11 - Len(l_str), "0") & l_str
'    Mid(lstrResult, 12, 3) = v_strExceptionCode
'    Mid(lstrResult, 15, 3) = v_strSubGroupCode
'
''    lstrResult = Space(10)
''    l_str = Trim(str(v_lngGroupCode))
''    Mid(lstrResult, 1, 4) = String(4 - Len(l_str), "0") & l_str
''    Mid(lstrResult, 5, 3) = v_strExceptionCode
''    Mid(lstrResult, 8, 3) = v_strSubGroupCode
'
'    ExceptionGroupKey = lstrResult
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "ExceptionGroupKey"
'End Function
'
'Private Function InitializeExceptionGroupsGroup() As Boolean
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim rs As ADODB.Recordset
'    Dim lvarRecordArray() As Variant
'    Dim lvarRecord(CEG_LAST_SLOT) As Variant
'    Dim llngRetVal As Long
'    Dim llngRecordCount As Long
'    Dim i As Integer
'    Dim lstrKey As String * 12
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'    Dim l_strSQL As String
'
'On Error GoTo eh
'    l_strSQL = GetExceptionGroupCacheSQL()
'    Set rs = GetCacheRecords(l_strSQL)
'    If rs.BOF And rs.EOF Then
'        InitializeExceptionGroupsGroup = False
'        Exit Function
'    End If
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    llngRecordCount = rs.RecordCount
'    ReDim lvarRecordArray(1, llngRecordCount - 1)
'    rs.MoveFirst
'    Do Until rs.EOF
'        lstrKey = ExceptionGroupKey(rs!EG_CD, rs!EXCEPTION_CD, rs!EG_SUB_CD)
'        lvarRecord(CEG_KEY) = lstrKey
'        lvarRecord(CEG_SEQNUM) = rs!EXC_SEQ_NUM
'        lvarRecordArray(0, i) = lstrKey
'        lvarRecordArray(1, i) = lvarRecord
'        i = i + 1
'        rs.MoveNext
'    Loop
'    rs.MoveFirst
'
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    llngRetVal = spmWrapper.AddAll(lvarRecordArray, CEG_GROUPNAME)
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Build exception groups group in seconds: " & CStr(dfTime / 1000)
'    #End If
'
'    InitializeExceptionGroupsGroup = (llngRetVal = 0)
'    Set spmWrapper = Nothing
'    Exit Function
'
'eh:
'    Set spmWrapper = Nothing
'    TrsRaiseError m_strCMODULENAME, "InitializeExceptionGroupsGroup"
'End Function

'Private Function GetExceptionGroupCacheSQL( _
'  ) As String
'
'    Dim l_strSQL As String
'
'On Error GoTo eh
'
'    l_strSQL = "SELECT EG_CD,EG_SUB_CD,EXCEPTION_CD,EXC_SEQ_NUM " _
'         & " FROM METHDBA.TTR_EXC_GROUP "
'
'    l_strSQL = l_strSQL & " ORDER BY EG_CD,EG_SUB_CD,EXCEPTION_CD "
'
'    GetExceptionGroupCacheSQL = l_strSQL
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetExceptionGroupCacheSQL"
'End Function

Public Function GetUniqueComponentCodes() As ADODB.Recordset

    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetUniqueComponentCodes_EH

    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")

    Set l_objCM = Nothing

    l_objCN.Open

'//TNV
    '//-----return a varchar(3) for use by drop down search mechanism('ST ' -> 'ST')-----//
     l_strSQL = "SELECT DISTINCT " _
             & "CAST(LTRIM(RTRIM(A.COMP_CD)) AS VARCHAR(3)) AS COMP_CD," _
             & "B.COMP_DESC, A.COMP_TYPE " _
             & "FROM METHDBA.TTR_UNION_COMP_AMT A, METHDBA.TTR_COMP_CD B " _
             & "Where A.COMP_CD = B.COMP_CD " _
             & "AND A.COMP_TYPE = B.COMP_TYPE " _
             & "AND B.COMP_CD not in " _
             & " (select RG_VALUE_CD " _
             & "   from METHDBA.TTR_RG_GROUP " _
             & "   where RG_CD = 85) " _
             & "ORDER BY COMP_CD "
'//TNV


    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    If l_objRS.BOF And l_objRS.EOF Then
        Set GetUniqueComponentCodes = Nothing
    Else
        Set GetUniqueComponentCodes = l_objRS
    End If

    Set l_objRS = Nothing

    Exit Function

GetUniqueComponentCodes_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetUniqueComponentCodes", , , , , True

End Function
'PREM BEGIN TRIS CONVERSION PHASE 3 (PHASE 1 .NET)
'The following functions retrieve Crafts, Qualifications
', Special Skills, and OT Responses information for Code
'Table Maintenance
Public Function GetAllCrafts() As ADODB.Recordset
On Error GoTo GetAllCrafts_EH
'Define
Dim lsSQL As String

    'SQL statement
    lsSQL = "SELECT    VARCHAR(LTRIM(RTRIM(CRAFT_NUM)), 3) AS CRAFT_NUM" & _
            "           , CRAFT_DESC, LAST_MOD_EMP_NUM" & _
            "           , CHAR(TIME_STAMP) AS TIME_STAMP, CRAFT_GROUP" & _
            "           , CASE(JOURN_APPR)" & _
            "           WHEN 'J' THEN 'JOURNEY'" & _
            "           WHEN 'A' THEN 'APPRENTICE'" & _
            "           ELSE ' ' END AS JOURN_APPR" & _
            " FROM      METHDBA.TTR_CRAFTS" & _
            " ORDER BY  CRAFT_NUM" & _
            " FOR FETCH ONLY WITH UR;"
    'Call ExecuteSQL and retrieve information
    Set GetAllCrafts = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    Exit Function
GetAllCrafts_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllCrafts"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllCrafts" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllCrafts" _
                    , trErrdbGetAllCrafts _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetAllQualifications() As ADODB.Recordset
On Error GoTo GetAllQualifications_EH
'Define
Dim lsSQL As String

    'SQL statement
    lsSQL = "SELECT    VARCHAR(LTRIM(RTRIM(QUAL_CD)), 3) AS QUAL_CD" & _
            "           , QUAL_DESC, LAST_MOD_EMP_NUM" & _
            "           , CHAR(TIME_STAMP) AS TIME_STAMP" & _
            " FROM      METHDBA.TTR_QUALIFICATIONS" & _
            " ORDER BY  QUAL_CD" & _
            " FOR FETCH ONLY WITH UR;"
    'Call ExecuteSQL and retrieve information
    Set GetAllQualifications = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    Exit Function
GetAllQualifications_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllQualifications"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllQualifications" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllQualifications" _
                    , trErrdbGetAllQuals _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetAllSpecialSkills() As ADODB.Recordset
On Error GoTo GetAllSpecialSkills_EH
'Define
Dim lsSQL As String

    'SQL statement
    lsSQL = " SELECT    VARCHAR(LTRIM(RTRIM(SPEC_SKILL_CD)), 3) AS SPEC_SKILL_CD" & _
            "           , SPEC_SKILL_DESC, VARCHAR(LTRIM(RTRIM(QUAL_CD)), 3) AS QUAL_CD" & _
            "           , LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP" & _
            " FROM      METHDBA.TTR_SPEC_SKILLS" & _
            " ORDER BY  SPEC_SKILL_CD" & _
            " FOR FETCH ONLY WITH UR;"

    'Call ExecuteSQL and retrieve information
    Set GetAllSpecialSkills = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    Exit Function
GetAllSpecialSkills_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllSpecialSkills"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllSpecialSkills" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllSpecialSkills" _
                    , trErrdbGetAllSkills _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetAllPromotionalLines() As ADODB.Recordset
On Error GoTo GetAllPromotionalLines_EH
'Define
Dim lsSQL As String

    'SQL statement
    lsSQL = "SELECT    DEPT_NUM, HR_LOC" & _
            "           , JOB_NUM, IR_AREA_CD" & _
            "           , IR_DIV_CD, LOP_CURR_LEV_IND" & _
            "           , LOP_NEXT_LEV_IND, LAST_MOD_EMP_NUM" & _
            "           , CHAR(TIME_STAMP) AS TIME_STAMP" & _
            " FROM      METHDBA.TTR_PROM_LINES" & _
            " ORDER BY  DEPT_NUM" & _
            " FOR FETCH ONLY WITH UR;"

    'Call ExecuteSQL and retrieve information
    Set GetAllPromotionalLines = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)
    Exit Function
GetAllPromotionalLines_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllPromotionalLines"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllPromotionalLines" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllPromotionalLines" _
                    , trErrdbGetAllPromotionalLines _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetAllCompanies() As ADODB.Recordset
On Error GoTo GetAllCompanies_EH
'Define
Dim lsSQL As String

    'SQL statement
    lsSQL = " SELECT    CMPY_CD, CMPY_DESC_TXT, CMPY_ORG_CD" & _
            "           , LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP" & _
            " FROM      METHDBA.TTR_COMPANY" & _
            " ORDER BY  CMPY_CD" & _
            " FOR FETCH ONLY WITH UR;"
    'Call ExecuteSQL and retrieve information
    Set GetAllCompanies = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)
    Exit Function
GetAllCompanies_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllCompanies"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllCompanies" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllCompanies" _
                    , trErrdbGetCompany _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'This function returns true if the provided Qualification Code
'exists in the database
Public Function IsAValidQualificationCode( _
                ByVal vsQualificationCode As String _
                ) As Boolean
On Error GoTo IsAValidQualificationCode_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    lsSQL = " SELECT   QUAL_CD" & _
            " FROM      METHDBA.TTR_QUALIFICATIONS" & _
            " WHERE     QUAL_CD = ?" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsQualificationCode
        .Update
    End With

'    Set ladrs = ExecuteSQL(lsSQL _
'            , adOpenForwardOnly _
'            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenForwardOnly _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    IsAValidQualificationCode = mADOExtensions.RecordsExist(ladrs)

    Set ladrs = Nothing
   
    Exit Function
IsAValidQualificationCode_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing
    
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidQualificationCode"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidQualificationCode" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidQualificationCode" _
                    , trErrdbGetAllQuals _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select
End Function
'This function returns true if the provided Special Skill Code
'exists in the database
Public Function IsAValidSpecialSkillCode( _
                ByVal vsSpecialSkillCode As String _
                , ByVal vsQualificationCode As String _
                ) As Boolean
On Error GoTo IsAValidSpecialSkillCode_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    lsSQL = " SELECT   SPEC_SKILL_CD" & _
            " FROM      METHDBA.TTR_SPEC_SKILLS" & _
            " WHERE     SPEC_SKILL_CD = ?" & _
            " AND       QUAL_CD = ?" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsSpecialSkillCode
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsQualificationCode
        .Update
    End With

'    Set ladrs = ExecuteSQL(lsSQL _
'            , adOpenForwardOnly _
'            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenForwardOnly _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    IsAValidSpecialSkillCode = mADOExtensions.RecordsExist(ladrs)

    Set ladrs = Nothing
   
    Exit Function
IsAValidSpecialSkillCode_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidSpecialSkillCode"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidSpecialSkillCode" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidSpecialSkillCode" _
                    , trErrdbGetAllSkills _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAValidSpecialSkillCode2(ByVal vsSpecialSkillCode As String _
                ) As Boolean
On Error GoTo IsAValidSpecialSkillCode2_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    lsSQL = " SELECT   SPEC_SKILL_CD" & _
            " FROM      METHDBA.TTR_SPEC_SKILLS" & _
            " WHERE     SPEC_SKILL_CD = ?" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsSpecialSkillCode
        .Update
    End With

'    Set ladrs = ExecuteSQL(lsSQL _
'            , adOpenForwardOnly _
'            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenForwardOnly _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    IsAValidSpecialSkillCode2 = mADOExtensions.RecordsExist(ladrs)

    Set ladrs = Nothing

    Exit Function
IsAValidSpecialSkillCode2_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidSpecialSkillCode2"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidSpecialSkillCode2" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidSpecialSkillCode2" _
                    , trErrdbGetAllSkills _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAValidDepartment( _
                    ByVal vsDepartmentNumber As String _
                    ) As Boolean
On Error GoTo IsAValidDepartment_EH
Dim lobjdbDepartment As dbDepartment.cdDepartment

    Set lobjdbDepartment = CreateObject( _
                            "dbDepartment.cdDepartment")
    'Set lobjdbDepartment = New dbDepartment.cdDepartment
    IsAValidDepartment = lobjdbDepartment.IsDepartmentValid( _
                            vsDepartmentNumber)
    Set lobjdbDepartment = Nothing

    Exit Function
IsAValidDepartment_EH:
    Set lobjdbDepartment = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "IsAValidDepartment"

End Function
Public Function IsAValidCraft( _
                    ByVal vsCraftNumber As String _
                    ) As Boolean
On Error GoTo IsAValidCraft_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    lsSQL = " SELECT   CRAFT_NUM" & _
            " FROM      METHDBA.TTR_CRAFTS" & _
            " WHERE     CRAFT_NUM = ?" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsCraftNumber
        .Update
    End With

'    Set ladrs = ExecuteSQL(lsSQL _
'            , adOpenForwardOnly _
'            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenForwardOnly _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    IsAValidCraft = mADOExtensions.RecordsExist(ladrs)

    Set ladrs = Nothing

    Exit Function
IsAValidCraft_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidCraft"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidCraft" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidCraft" _
                    , trErrdbGetAllCrafts _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAValidCompany( _
                    ByVal vsCompanyCode As String _
                    ) As Boolean
On Error GoTo IsAValidCompany_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    lsSQL = " SELECT   CMPY_CD" & _
            " FROM      METHDBA.TTR_COMPANY" & _
            " WHERE     CMPY_CD = ?" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsCompanyCode
        .Update
    End With

'    Set ladrs = ExecuteSQL(lsSQL _
'            , adOpenForwardOnly _
'            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenForwardOnly _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    IsAValidCompany = mADOExtensions.RecordsExist(ladrs)

    Set ladrs = Nothing

    Exit Function
IsAValidCompany_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidCompany"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidCompany" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidCompany" _
                    , trErrdbGetCompany _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAValidIndustrialRelationsArea( _
                    ByVal vsIRArea As String _
                    ) As Boolean
On Error GoTo IsAValidIndustrialRelationsArea_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    lsSQL = " SELECT   IR_AREA_CD" & _
            " FROM      METHDBA.TTR_PROM_LINES" & _
            " WHERE     IR_AREA_CD = ?" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 2
        !PARM_VALUE = vsIRArea
        .Update
    End With

'    Set ladrs = ExecuteSQL(lsSQL _
'            , adOpenForwardOnly _
'            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenForwardOnly _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    IsAValidIndustrialRelationsArea = _
        mADOExtensions.RecordsExist(ladrs)

    Set ladrs = Nothing

    Exit Function
IsAValidIndustrialRelationsArea_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing
    
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidIndustrialRelationsArea"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidIndustrialRelationsArea" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidIndustrialRelationsArea" _
                    , trErrdbGetIndustrialRelationsInfo _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAValidIndustrialRelationsDivision( _
                    ByVal vsIRDivision As String _
                    ) As Boolean
On Error GoTo IsAValidIndustrialRelationsDivision_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    lsSQL = " SELECT   IR_DIV_CD" & _
            " FROM      METHDBA.TTR_PROM_LINES" & _
            " WHERE     IR_DIV_CD = ?" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 2
        !PARM_VALUE = vsIRDivision
        .Update
    End With

'    Set ladrs = ExecuteSQL(lsSQL _
'            , adOpenForwardOnly _
'            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenForwardOnly _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    IsAValidIndustrialRelationsDivision = _
        mADOExtensions.RecordsExist(ladrs)

    Set ladrs = Nothing

    Exit Function
IsAValidIndustrialRelationsDivision_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidIndustrialRelationsDivision"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidIndustrialRelationsDivision" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidIndustrialRelationsDivision" _
                    , trErrdbGetIndustrialRelationsInfo _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetAllOvertimeBankHeaders() As ADODB.Recordset
On Error GoTo GetAllOvertimeBankHeaders_EH
'Define
Dim lsSQL As String

    'SQL statement
    lsSQL = "SELECT    OT_BANK_CD, OT_BANK_DESC" & _
            "           , CREATED_EMP_NUM , CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
            "           , LAST_MOD_EMP_NUM, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
            " FROM      METHDBA.TTR_OT_BANK_CD" & _
            " ORDER BY  OT_BANK_CD" & _
            " FOR FETCH ONLY WITH UR;"

    'Call ExecuteSQL and retrieve information
    Set GetAllOvertimeBankHeaders = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    Exit Function
GetAllOvertimeBankHeaders_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllOvertimeBankHeaders"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllOvertimeBankHeaders" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllOvertimeBankHeaders" _
                    , trErrdbGetOTBankCode _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function DoesBankExist(ByVal viOTBankCode As Integer) As Boolean
On Error GoTo DoesBankExist_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    DoesBankExist = False

    'SQL statement
    lsSQL = "SELECT    OT_BANK_CD" & _
            " FROM      METHDBA.TTR_OT_BANK_CD" & _
            " WHERE     OT_BANK_CD = ?" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = viOTBankCode
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set ladrs = ExecuteSQL(lsSQL _
'                , adOpenForwardOnly _
'                , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
                , ladrsParms _
                , adOpenForwardOnly _
                , adLockReadOnly)
    Set ladrsParms = Nothing

    'return
    DoesBankExist = mADOExtensions.RecordsExist(ladrs)

    'Destroy recordset
    Set ladrs = Nothing

    Exit Function
DoesBankExist_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "DoesBankExist"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "DoesBankExist" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "DoesBankExist" _
                    , trErrdbGetOTBankCode _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetAllBudgetNumbers() As ADODB.Recordset
On Error GoTo GetAllBudgetNumbers_EH
'Define
Dim lsSQL As String

    'SQL statement
    lsSQL = "SELECT    DISTINCT BUDGET_NUM" & _
            " FROM      METHDBA.TTR_DEPT" & _
            " ORDER BY  BUDGET_NUM" & _
            " FOR FETCH ONLY WITH UR;"

    'Call ExecuteSQL and retrieve information
    Set GetAllBudgetNumbers = ExecuteSQL(lsSQL _
                , adOpenStatic _
                , adLockReadOnly)
    Exit Function
GetAllBudgetNumbers_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllBudgetNumbers"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllBudgetNumbers" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetAllBudgetNumbers" _
                    , trErrdbGetBudgetNumbers _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select
End Function
Public Function IsAValidBudget(ByVal vsBudgetNumber As String _
                ) As Boolean
On Error GoTo IsAValidBudget_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    IsAValidBudget = False

    'SQL statement
    lsSQL = "SELECT    BUDGET_NUM" & _
            " FROM      METHDBA.TTR_DEPT" & _
            " WHERE     BUDGET_NUM = ?" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsBudgetNumber
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set ladrs = ExecuteSQL(lsSQL _
'                , adOpenForwardOnly _
'                , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
                , ladrsParms _
                , adOpenForwardOnly _
                , adLockReadOnly)
    Set ladrsParms = Nothing

    'return recordset
    IsAValidBudget = mADOExtensions.RecordsExist(ladrs)

    'Destroy recordset
    Set ladrs = Nothing

    Exit Function
IsAValidBudget_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidBudget"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidBudget" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidBudget" _
                    , trErrdbGetBudgetNumbers _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAValidOpportunity(ByVal vlOTOpportunity As Long _
                ) As Boolean
On Error GoTo IsAValidOpportunity_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    IsAValidOpportunity = False

    'SQL statement
    lsSQL = "SELECT    OT_OPP_NUM" & _
            " FROM      METHDBA.TTR_OT_OPPORTUNITY" & _
            " WHERE     OT_OPP_NUM = ?" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOTOpportunity
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set ladrs = ExecuteSQL(lsSQL _
'                    , adOpenForwardOnly _
'                    , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenForwardOnly _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    'return
    IsAValidOpportunity = mADOExtensions.RecordsExist(ladrs)

    'Destroy recordset
    Set ladrs = Nothing

    Exit Function
IsAValidOpportunity_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidOpportunity"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidOpportunity" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsAValidOpportunity" _
                    , trErrdbOpportunityInfo _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAnAssignedQualification(ByVal vsQualification As String _
                ) As Boolean
On Error GoTo IsAnAssignedQualification_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    IsAnAssignedQualification = False

    lsSQL = "SELECT    SUM(MYFIELD) AS NUM_USED" & _
        " FROM  ("

    'SQL statement
    lsSQL = lsSQL & _
        " SELECT    COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_EMP_QUAL_SKILL" & _
        " WHERE     QUAL_CD = ?"

    lsSQL = lsSQL & " UNION ALL "

    lsSQL = lsSQL & _
        " SELECT    COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_SPEC_SKILLS" & _
        " WHERE     QUAL_CD = ?"

    lsSQL = lsSQL & ") IAAQ"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsQualification
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsQualification
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set ladrs = ExecuteSQL(lsSQL _
'                    , adOpenForwardOnly _
'                    , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenForwardOnly _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    'return
    If (mADOExtensions.RecordsExist(ladrs)) Then
        If (ladrs!NUM_USED > 0) Then
            IsAnAssignedQualification = True
        End If
    End If

    'Destroy recordset
    Set ladrs = Nothing

    Exit Function
IsAnAssignedQualification_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsAnAssignedQualification"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsAnAssignedQualification" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsAnAssignedQualification" _
                    , trErrdbGetEmpQAndS _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsAnAssignedSkill(ByVal vsSpecialSkill As String _
                , ByVal vsQualification As String _
                ) As Boolean
On Error GoTo IsAnAssignedSkill_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    IsAnAssignedSkill = False

    'SQL statement
    lsSQL = "SELECT    EMP_NUM" & _
        " FROM      METHDBA.TTR_EMP_QUAL_SKILL" & _
        " WHERE     SPEC_SKILL_CD = ?" & _
        " AND       QUAL_CD = ?" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsSpecialSkill
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsQualification
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set ladrs = ExecuteSQL(lsSQL _
'                    , adOpenForwardOnly _
'                    , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenForwardOnly _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    'return
    IsAnAssignedSkill = mADOExtensions.RecordsExist(ladrs)

    'Destroy recordset
    Set ladrs = Nothing

    Exit Function
IsAnAssignedSkill_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsAnAssignedSkill"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsAnAssignedSkill" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsAnAssignedSkill" _
                    , trErrdbGetEmpQAndS _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsCraftUsedinTRIS(ByVal vsCraftNumber As String _
                , Optional ByVal vsJobNumber As String = "" _
                ) As Boolean
On Error GoTo IsCraftUsedinTRIS_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset
'NOTE TTR_PROJ_PLANT_REL DOES NOT USE CRAFT_NUM
'ALSO TTR_PROJECT_REQ DOES NOT EXIST

    lsSQL = "SELECT    SUM(MYFIELD) AS NUM_USED" & _
        " FROM  ("

    If (Trim(vsJobNumber) <> "") Then
        lsSQL = lsSQL & _
            " SELECT   COUNT(*) AS MYFIELD" & _
            " FROM      METHDBA.TTR_CANV_MULTI" & _
            " WHERE     JOB_NUM = ?"

        lsSQL = lsSQL & " UNION ALL "
    End If

    lsSQL = lsSQL & _
        " SELECT   COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_JOBS" & _
        " WHERE     CRAFT_NUM = ?"

    lsSQL = lsSQL & " UNION ALL "

    lsSQL = lsSQL & _
        " SELECT   COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_CANV_SHEETS" & _
        " WHERE     CRAFT_NUM = ?"

    lsSQL = lsSQL & " UNION ALL "

    lsSQL = lsSQL & _
        " SELECT   COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_OT_EMP_OVRD" & _
        " WHERE     CRAFT_NUM = ?"

    lsSQL = lsSQL & ") ICUIT"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        If (Trim(vsJobNumber) <> "") Then
            .AddNew
            !PARM_ORDER = 1
            !PARM_NAME = ""
            !PARM_TYPE = adChar
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 6
            !PARM_VALUE = vsJobNumber
            .Update
        End If

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsCraftNumber
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsCraftNumber
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsCraftNumber
        .Update
    End With

'    Set ladrs = ExecuteSQL(lsSQL _
'            , adOpenForwardOnly _
'            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenForwardOnly _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    If (mADOExtensions.RecordsExist(ladrs)) Then
        If (ladrs!NUM_USED > 0) Then
            IsCraftUsedinTRIS = True
        End If
    End If

    Set ladrs = Nothing

    Exit Function
IsCraftUsedinTRIS_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsCraftUsedinTRIS"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsCraftUsedinTRIS" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsCraftUsedinTRIS" _
                    , trErrdbGetAllCrafts _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsOTReasonUsedinTRIS(ByVal vsOTReason As String _
                ) As Boolean
On Error GoTo IsOTReasonUsedinTRIS_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    lsSQL = "SELECT    SUM(MYFIELD) AS NUM_USED" & _
        " FROM  ("

    lsSQL = lsSQL & _
        " SELECT   COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_OT_OPPORTUNITY" & _
        " WHERE     OT_REASON_CD = ?"

    lsSQL = lsSQL & " UNION ALL "

    lsSQL = lsSQL & _
        " SELECT   COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_ACTUAL_TIME" & _
        " WHERE     OT_REASON_CD = ?"

    lsSQL = lsSQL & " UNION ALL "

    lsSQL = lsSQL & _
        " SELECT   COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_ADJ_TIME" & _
        " WHERE     OT_REASON_CD = ?"

    lsSQL = lsSQL & " UNION ALL "

    lsSQL = lsSQL & _
        " SELECT   COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_OT_OPP_TM_CHRG" & _
        " WHERE     OT_REASON_CD = ?"

    lsSQL = lsSQL & ") IOTRUIT"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsOTReason
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsOTReason
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsOTReason
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsOTReason
        .Update
    End With

'PREMPREM WHAT ABOUT TTR_ACTUAL_TIME_HS??

'    Set ladrs = ExecuteSQL(lsSQL _
'            , adOpenForwardOnly _
'            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenForwardOnly _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    If (mADOExtensions.RecordsExist(ladrs)) Then
        If (ladrs!NUM_USED > 0) Then
            IsOTReasonUsedinTRIS = True
        End If
    End If

    Set ladrs = Nothing

    Exit Function
IsOTReasonUsedinTRIS_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsOTReasonUsedinTRIS"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsOTReasonUsedinTRIS" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsOTReasonUsedinTRIS" _
                    , trErrdbGetAllReasons _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsOTResponseUsedinTRIS(ByVal vsOTResponse As String _
                ) As Boolean
On Error GoTo IsOTResponseUsedinTRIS_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    lsSQL = "SELECT    SUM(MYFIELD) AS NUM_USED" & _
        " FROM  ("

    lsSQL = lsSQL & _
        " SELECT    COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_CANV_SHEETS" & _
        " WHERE     CANVASS_RESP_CD = ?" & _
        " OR        FORCE_RESP_CD = ?"

    lsSQL = lsSQL & " UNION ALL "

    lsSQL = lsSQL & _
        " SELECT    COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_CANV_TMPL_RESP" & _
        " WHERE     OT_RESP_CD = ?"

    lsSQL = lsSQL & ") IOTRUIT"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 5
        !PARM_VALUE = vsOTResponse
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 5
        !PARM_VALUE = vsOTResponse
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 5
        !PARM_VALUE = vsOTResponse
        .Update
    End With

'    Set ladrs = ExecuteSQL(lsSQL _
'            , adOpenForwardOnly _
'            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenForwardOnly _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    If (mADOExtensions.RecordsExist(ladrs)) Then
        If (ladrs!NUM_USED > 0) Then
            IsOTResponseUsedinTRIS = True
        End If
    End If

    Set ladrs = Nothing

    Exit Function
IsOTResponseUsedinTRIS_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsOTResponseUsedinTRIS"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsOTResponseUsedinTRIS" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsOTResponseUsedinTRIS" _
                    , trErrdbGetAllResponses _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function IsOTBankUsedinTRIS(ByVal viOTBankCode As Integer _
                ) As Boolean
On Error GoTo IsOTBankUsedinTRIS_EH
'Define
Dim lsSQL As String, ladrs As ADODB.Recordset
Dim ladrsParms As ADODB.Recordset

    lsSQL = "SELECT    SUM(MYFIELD) AS NUM_USED" & _
        " FROM  ("

    lsSQL = lsSQL & _
        " SELECT   COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_OT_BANKS" & _
        " WHERE     OT_BANK_CD = ?"

    lsSQL = lsSQL & " UNION ALL "

    lsSQL = lsSQL & _
        " SELECT   COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_OT_GROUP" & _
        " WHERE     OT_STAND_BANK_CD = ?"

    lsSQL = lsSQL & " UNION ALL "

    lsSQL = lsSQL & _
        " SELECT   COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_OT_GROUP_BANK" & _
        " WHERE     OT_BANK_CD = ?"

    lsSQL = lsSQL & " UNION ALL "

    lsSQL = lsSQL & _
        " SELECT   COUNT(*) AS MYFIELD" & _
        " FROM      METHDBA.TTR_CANV_SHEETS" & _
        " WHERE     OT_BANK_CD = ?"

    lsSQL = lsSQL & ") IOTBUIT"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = viOTBankCode
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = viOTBankCode
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = viOTBankCode
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = viOTBankCode
        .Update
    End With

'    Set ladrs = ExecuteSQL(lsSQL _
'            , adOpenForwardOnly _
'            , adLockReadOnly)
    Set ladrs = NewExecuteSQL(lsSQL _
            , ladrsParms _
            , adOpenForwardOnly _
            , adLockReadOnly)
    Set ladrsParms = Nothing

    If (mADOExtensions.RecordsExist(ladrs)) Then
        If (ladrs!NUM_USED > 0) Then
            IsOTBankUsedinTRIS = True
        End If
    End If

    Set ladrs = Nothing

    Exit Function
IsOTBankUsedinTRIS_EH:
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "IsOTBankUsedinTRIS"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "IsOTBankUsedinTRIS" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "IsOTBankUsedinTRIS" _
                    , trErrdbGetOTBankCode _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetJobTaskExceptions() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetJobTaskExceptions_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
'    l_strSQL = "select CAST(LTRIM(RTRIM(EXCEPTION_CD)) AS VARCHAR(3)) AS EXCEPTION_CD, " _
'             & " EXCEPTION_DESC, EXCEPTION_TYPE, " _
'             & " CONTRACT_EARN_CD, EXEC_EARN_CD, SYMBOL_TIME_IND, " _
'             & " EXCEPTION_BANK_CD, INELIGIBILITY_IND, OTHERPAY_IND, " _
'             & " CASE_GROUP, SCHED_IND, 'N' as FMLA_IND, " _
'             & " WORK_REST_IND, SCHED_CHG_IND, ACCT_NUM, SYMBOL_CD " _
'             & " , AIC_ACCT_NUM, AMS_ACCT_NUM, UEC_ACCT_NUM " _
'             & " from METHDBA.TTR_EXCEPTION_CDS " _
'             & " where EXCEPTION_CD not in " _
'             & " (select RG_VALUE_CD " _
'             & "   from METHDBA.TTR_RG_GROUP " _
'             & "   where RG_CD = 6)" _
'             & " AND EXCEPTION_CD IN (SELECT RG_VALUE_CD " _
'             & "   From METHDBA.TTR_RG_GROUP " _
'             & "   WHERE RG_CD = 142)"

    If IsPOETActive And IsOTTCEffective Then
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetJobTaskExceptionsOTTC.sql")
    Else
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetJobTaskExceptions.sql")
    End If
    
    '//Set l_objRS = oADOHelper.ExecuteSelect(l_strSQL, adOpenStatic, adLockBatchOptimistic)

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetJobTaskExceptions = l_objRS

    Set l_objRS = Nothing
    Exit Function

GetJobTaskExceptions_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetJobTaskExceptions", , , , DescArgsFrom(l_strSQL), True
End Function



Public Function GetNRCRuleEffectiveDate() As String
    
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
    
    l_objCN.Open
    l_strSQL = "SELECT parm_value" _
             & " FROM methdba.ttr_sysparms" _
             & " WHERE parm_num = 43" _
             & " AND parm_suffix = 0" _

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    
    If RecordsExist(l_objRS) Then
        GetNRCRuleEffectiveDate = l_objRS!PARM_VALUE
    Else
        GetNRCRuleEffectiveDate = ""
    End If

    Set l_objRS = Nothing
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetNRCRuleEffectiveDate", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function ExceptionAccountNumberAIC( _
            ByVal v_strExcCd As String) As String
'BLUE074, BLUE077
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionAccountNumberAIC = l_varRecord(CEX_AIC_ACCT_NUM)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionAccountNumberAIC"
End Function

Public Function ExceptionAccountNumberAMS( _
            ByVal v_strExcCd As String) As String
'BLUE074, BLUE077
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionAccountNumberAMS = l_varRecord(CEX_AMS_ACCT_NUM)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionAccountNumberAMS"
End Function

Public Function ExceptionAccountNumberUEC( _
            ByVal v_strExcCd As String) As String
'BLUE074, BLUE077
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    ExceptionAccountNumberUEC = l_varRecord(CEX_UEC_ACCT_NUM)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionAccountNumberUEC"
End Function

Public Function ExceptionAccountNumberByCompany( _
            ByVal v_strExcCd As String, ByVal v_strCompany As String) As String
'BLUE074, BLUE077
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim lstrAcctNum As String
    
On Error GoTo eh
    l_varRecord = GetRecord(v_strExcCd, CEX_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    
    Select Case v_strCompany
        Case "AIC"
            lstrAcctNum = l_varRecord(CEX_AIC_ACCT_NUM)
        Case "AMS"
            lstrAcctNum = l_varRecord(CEX_AMS_ACCT_NUM)
        Case "UEC"
            lstrAcctNum = l_varRecord(CEX_UEC_ACCT_NUM)
        Case Else
            lstrAcctNum = l_varRecord(CEX_ACCT_NUM)
    
    End Select
    
    
    ExceptionAccountNumberByCompany = lstrAcctNum
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ExceptionAccountNumberByCompany"
End Function

Private Sub Class_Initialize()
    Set m_ResourceMgr = New utilADOHelper.cuResourceMgr
End Sub

Private Sub Class_Terminate()
    Set m_ResourceMgr = Nothing
End Sub
