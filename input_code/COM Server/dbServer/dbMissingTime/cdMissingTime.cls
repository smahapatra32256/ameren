VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdMissingTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Time to Pay
'Author:    Prem Chander
'Date:      January 2006
'Description: This module will retrieve
'   - contract & management employees missing time for the just ended pay period.

Option Explicit

Implements IObjectConstruct
Dim m_ConstructString As String

'#Const DEBUGGING = 0 '1 - ON; 0 - OFF

Private Const MODULENAME As String = "dbMissingTime.cdMissingTime"
Dim m_colConstructStringProperty As Collection
Dim DEBUGGING As Integer '1 - ON;0 - OFF
Dim TRACING As Integer '1 - ON;0 - OFF

Private Const m_DBSCRIPTSROOT As String = "C:\TRIS SOLUTION\dbScripts\TRIS_db2"
Private Const m_DBSCRIPTSPATH As String = "\dbMissingTime\cdMissingTime"
Private m_ResourceMgr As utilADOHelper.cuResourceMgr

Private Sub Class_Initialize()
    Set m_ResourceMgr = New utilADOHelper.cuResourceMgr
    ''test
'    DEBUGGING = 0
End Sub

Private Sub Class_Terminate()
    Set m_ResourceMgr = Nothing
End Sub

Private Sub IObjectConstruct_Construct(ByVal pCtorObj As Object)
    Dim mConstructString As IObjectConstructString
    
    Set mConstructString = pCtorObj
    
    m_ConstructString = mConstructString.ConstructString
    
    Set m_colConstructStringProperty = ParseContructString(m_ConstructString)
    
    PostConstruct
    App.LogEvent MODULENAME & " ObjectConstructString = " & m_ConstructString, vbLogEventTypeInformation
End Sub

Private Sub PostConstruct()
    If DoesPropertyExists(m_colConstructStringProperty, "DEBUGGING") Then
        If IsNumeric(m_colConstructStringProperty.Item("DEBUGGING")) Then
            DEBUGGING = CInt(m_colConstructStringProperty.Item("DEBUGGING"))
        End If
    End If
    If DoesPropertyExists(m_colConstructStringProperty, "TRACING") Then
        If IsNumeric(m_colConstructStringProperty.Item("TRACING")) Then
            TRACING = CInt(m_colConstructStringProperty.Item("TRACING"))
        End If
    End If
End Sub
    
Private Function DoesPropertyExists(ByVal colPropertyValue As Collection, ByVal Property As String) As Boolean
    Dim oValue As Variant

    On Error GoTo eh
    
    oValue = colPropertyValue.Item(Property)
    
    DoesPropertyExists = True
    Exit Function
eh:
    DoesPropertyExists = False
End Function

Private Function ParseContructString(ByVal ConstructString As String) As Collection
    Dim colConstructStringProperty As Collection
    Dim PropertyValueArray() As String
    Dim PropertyValue As Variant
    Dim PropertyValueArray2() As String

    Set colConstructStringProperty = New Collection
    PropertyValueArray = Split(ConstructString, ";")
    
    For Each PropertyValue In PropertyValueArray
        PropertyValueArray2 = Split(PropertyValue, "=")
        If UBound(PropertyValueArray2) >= 1 Then
            colConstructStringProperty.Add PropertyValueArray2(1), PropertyValueArray2(0)
        End If
    Next
       
    Set ParseContructString = colConstructStringProperty
End Function

Public Function GetMgmtMissingTimeEmployees() As ADODB.Recordset
    TrsRaiseError MODULENAME, "GetMgmtMissingTimeEmployees", trErrObsoleteMethod, , , , True
End Function

Public Function GetMgmtMissingTimeEmployees2(ByVal vdtScheduleDateTime As Date) As ADODB.Recordset
On Error GoTo eh
'Connection
Dim ladcn As ADODB.Connection, ladrsParms3 As ADODB.Recordset
'SQL String Holders
Dim lsSQL1 As String, lsSQL2 As String
Dim lsSQL3 As String, lsSQL4 As String
Dim lsSQL5 As String, lsSQL6 As String
Dim lsSQL7 As String, lsSQL8 As String
Dim lsSQL As String

    Set GetMgmtMissingTimeEmployees2 = Nothing

    lsSQL7 = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_MEMP_HRS_SCHED " & _
        "(EMP_NUM INTEGER NOT NULL, TOT_SCHED_HRS DECIMAL(4, 1) NOT NULL) " & _
        "ON COMMIT PRESERVE ROWS; "

    lsSQL8 = "INSERT INTO SESSION.GTT_MEMP_HRS_SCHED " & _
        "(EMP_NUM, TOT_SCHED_HRS) " & _
        "SELECT    A.EMP_NUM, COALESCE(SUM(B.HOURS), 0.0) AS TOT_SCHED_HRS " & _
        "FROM SESSION.GTT_MEMPS_PP A " & _
        "INNER JOIN METHDBA.TTR_EMPL_SCH B " & _
        "  ON A.EMP_NUM = B.EMP_NUM " & _
        " AND B.CHARGE_DATE BETWEEN A.PAY_PER_START_DATE " & _
        " AND A.PAY_PER_END_DATE " & _
        "GROUP BY A.EMP_NUM; "

    lsSQL1 = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_MEMPS_PP" & _
        " (EMP_NUM INTEGER NOT NULL, EMP_NAME VARCHAR(30) NOT NULL" & _
        ", CMPY_CD VARCHAR(3) NOT NULL, CMPY_DESC_TXT VARCHAR(30) NOT NULL" & _
        ", DEPT_NUM CHAR(3) NOT NULL, DEPT_NAME VARCHAR(30) NOT NULL" & _
        ", SUPERVISOR_NUM INTEGER NOT NULL, SUPERVISOR_NAME VARCHAR(30) NOT NULL" & _
        ", PAY_PER_START_DATE DATE NOT NULL, PAY_PER_END_DATE DATE NOT NULL" & _
        ", PAY_PER_TOT_HRS DECIMAL(4, 1) NOT NULL)" & _
        " ON COMMIT PRESERVE ROWS;"

    lsSQL2 = " DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_MEMP_HRS_PP" & _
        " (EMP_NUM INTEGER NOT NULL, TOT_HRS DECIMAL(4, 1) NOT NULL)" & _
        " ON COMMIT PRESERVE ROWS;"

    lsSQL3 = "INSERT INTO SESSION.GTT_MEMPS_PP" & _
        " (EMP_NUM, EMP_NAME" & _
        ", CMPY_CD, CMPY_DESC_TXT" & _
        ", DEPT_NUM, DEPT_NAME" & _
        ", SUPERVISOR_NUM, SUPERVISOR_NAME" & _
        ", PAY_PER_START_DATE, PAY_PER_END_DATE" & _
        ", PAY_PER_TOT_HRS)" & _
        " SELECT    A.EMP_NUM, A.EMP_NAME" & _
        "           , A.CMPY_CD, D.CMPY_DESC_TXT" & _
        "           , A.DEPT_NUM, C.DEPT_NAME" & _
        "           , INTEGER(E.SUPV_ID) AS SUPERVISOR_NUM, F.EMP_NAME" & _
        "           , B.PAY_PER_START_DATE, B.PAY_PER_END_DATE" & _
        "           , B.PAY_PER_TOT_HRS"
    lsSQL3 = lsSQL3 & _
        " FROM      METHDBA.TTR_EMPLOYEE A" & _
        "           INNER JOIN METHDBA.TTR_COLL_PERIODS B" & _
        "           ON CAST(? AS DATE) - 5 DAYS BETWEEN B.PAY_PER_START_DATE" & _
        "           AND B.PAY_PER_END_DATE" & _
        "           AND B.PAYROLL_GROUP_CD = 'E'" & _
        "           AND A.EMP_STATUS <> 'T'" & _
        "           AND A.STATUS_EFF_DATE <= B.PAY_PER_END_DATE" & _
        "           AND A.DEPT_EFF_DATE <= B.PAY_PER_END_DATE" & _
        "           AND A.CMPY_CD <> 'MV1'" & _
        "           AND A.DEPT_NUM NOT IN (SELECT RG_VALUE_CD FROM METHDBA.TTR_RG_GROUP WHERE RG_CD = 167 AND APPLY_IND = 1)" & _
        "           INNER JOIN METHDBA.TTR_DEPT C" & _
        "           ON C.DEPT_NUM = A.DEPT_NUM" & _
        "           AND C.DEPT_NUM > '499'" & _
        "           AND C.DEPT_NUM NOT IN (SELECT RG_VALUE_CD FROM METHDBA.TTR_RG_GROUP WHERE RG_CD = 167 AND APPLY_IND = 1)" & _
        "           AND C.CMPY_CD <> 'MV1'" & _
        "           AND C.PAYROLL_GROUP_CD = 'E'" & _
        "           AND B.PAYROLL_GROUP_CD = C.PAYROLL_GROUP_CD" & _
        "           INNER JOIN METHDBA.TTR_COMPANY D" & _
        "           ON D.CMPY_CD = A.CMPY_CD" & _
        "           AND D.CMPY_CD <> 'MV1'" & _
        "           INNER JOIN METHDBA.TAR_COMN_EMPL_TBL E" & _
        "           ON A.EMP_NUM = INTEGER(E.EMPL_ID)" & _
        "           INNER JOIN METHDBA.TTR_EMPLOYEE F" & _
        "           ON F.EMP_NUM = INTEGER(E.SUPV_ID);"

    lsSQL4 = "INSERT INTO SESSION.GTT_MEMP_HRS_PP" & _
        " (EMP_NUM, TOT_HRS)" & _
        " SELECT    A.EMP_NUM, COALESCE(SUM(B.HOURS), 0.0) AS TOT_HRS" & _
        " FROM      SESSION.GTT_MEMPS_PP A" & _
        "           INNER JOIN METHDBA.TTR_ACTUAL_TIME B" & _
        "           ON A.EMP_NUM = B.EMP_NUM" & _
        "           AND B.CHARGE_DATE BETWEEN A.PAY_PER_START_DATE" & _
        "           AND A.PAY_PER_END_DATE" & _
        "           AND B.EXCEPTION_CD <> 'ONP'" & _
        "           AND B.EXCEPTION_CD <> 'HR'" & _
        "           INNER JOIN METHDBA.TTR_EARN_CODES C" & _
        "           ON C.EARNINGS_CD = B.EARNINGS_CD" & _
        "           AND C.EXEC_CONTRACT_IND = 'E'" & _
        "           AND C.PAY_DISTR_TYPE <> 'OT'" & _
        "           INNER JOIN METHDBA.TTR_EXCEPTION_CDS D" & _
        "           ON D.EXCEPTION_CD = B.EXCEPTION_CD" & _
        "           AND D.EXCEPTION_TYPE <> 'S'" & _
        " GROUP BY A.EMP_NUM;"

    lsSQL5 = " SELECT * FROM ( " & _
        "SELECT    A.EMP_NUM, EMP_NAME" & _
        "           , DECIMAL(0.0,4,1) AS TOT_HRS" & _
        "           , DECIMAL(COALESCE(C.TOT_SCHED_HRS, 0.0),4,1) AS TOT_SCHED_HRS " & _
        "           , CMPY_CD, CMPY_DESC_TXT" & _
        "           , DEPT_NUM, DEPT_NAME" & _
        "           , SUPERVISOR_NUM, SUPERVISOR_NAME" & _
        "           , PAY_PER_START_DATE, PAY_PER_END_DATE" & _
        "           , DECIMAL(PAY_PER_TOT_HRS,4,1) AS PAY_PER_TOT_HRS, 'Y' AS SEND_IND" & _
        "           , COALESCE(SUBSTR(RG.RG_VALUE_CD, 5), RTRIM(CHAR(A.EMP_NUM))||'@ameren.com') AS NOTIFY_EMP_NUM" & _
        "           , CASE WHEN RG.RG_VALUE_CD IS NULL THEN 'Y' ELSE 'N' END AS COPY_SUPERVISOR_IND" & _
        " FROM      SESSION.GTT_MEMPS_PP A " & _
        "           LEFT OUTER JOIN SESSION.GTT_MEMP_HRS_SCHED C " & _
        "           ON A.EMP_NUM = C.EMP_NUM " & _
        "           LEFT OUTER JOIN METHDBA.TTR_RG_GROUP RG" & _
        "           ON RG.RG_CD = 171 AND RG.APPLY_IND = 1" & _
        "           AND LEFT(RG.RG_VALUE_CD, 3) = DEPT_NUM" & _
        " WHERE     A.EMP_NUM NOT IN (SELECT  EMP_NUM" & _
        "                           FROM    SESSION.GTT_MEMP_HRS_PP)" & _
        " UNION"
    lsSQL5 = lsSQL5 & _
        " SELECT    A.EMP_NUM, A.EMP_NAME" & _
        "           , DECIMAL(COALESCE(B.TOT_HRS, 0.0),4,1) AS TOT_HRS" & _
        "           , DECIMAL(COALESCE(C.TOT_SCHED_HRS, 0.0),4,1) AS TOT_SCHED_HRS " & _
        "           , A.CMPY_CD, A.CMPY_DESC_TXT" & _
        "           , A.DEPT_NUM, A.DEPT_NAME" & _
        "           , A.SUPERVISOR_NUM, A.SUPERVISOR_NAME" & _
        "           , A.PAY_PER_START_DATE, A.PAY_PER_END_DATE" & _
        "           , DECIMAL(PAY_PER_TOT_HRS,4,1) AS PAY_PER_TOT_HRS, 'Y' AS SEND_IND" & _
        "           , COALESCE(SUBSTR(RG.RG_VALUE_CD, 5), RTRIM(CHAR(A.EMP_NUM))||'@ameren.com') AS NOTIFY_EMP_NUM" & _
        "           , CASE WHEN RG.RG_VALUE_CD IS NULL THEN 'Y' ELSE 'N' END AS COPY_SUPERVISOR_IND" & _
        " FROM      SESSION.GTT_MEMPS_PP A" & _
        "           LEFT OUTER JOIN SESSION.GTT_MEMP_HRS_SCHED C " & _
        "           ON A.EMP_NUM = C.EMP_NUM " & _
        "           INNER JOIN SESSION.GTT_MEMP_HRS_PP B" & _
        "           ON A.EMP_NUM = B.EMP_NUM" & _
        "           LEFT OUTER JOIN METHDBA.TTR_RG_GROUP RG" & _
        "           ON RG.RG_CD = 171 AND RG.APPLY_IND = 1" & _
        "           AND LEFT(RG.RG_VALUE_CD, 3) = DEPT_NUM ) DET " & _
        " where TOT_SCHED_HRS > TOT_HRS   " & _
        " ORDER BY  1" & _
        " FOR FETCH ONLY WITH UR ; "

    lsSQL6 = "DROP TABLE SESSION.GTT_MEMP_HRS_PP; DROP TABLE SESSION.GTT_MEMPS_PP; DROP TABLE SESSION.GTT_MEMP_HRS_SCHED;"

    Set ladrsParms3 = CreateParameterRecordset()
    AddInputParameterRecordset ladrsParms3, 1, "", adDate, 0, MidnightDate(vdtScheduleDateTime)
    
    Set ladcn = OpenConnection

    If Not (ladcn Is Nothing) Then
        'Declare Global Temp
        lsSQL = lsSQL1
        ExecuteOTHERSQL lsSQL1, ladcn, adCmdText
        lsSQL = lsSQL2
        ExecuteOTHERSQL lsSQL2, ladcn, adCmdText

        'Select and Insert into Global Temp
        lsSQL = lsSQL3
        NewExecuteOtherSQL lsSQL3, ladrsParms3, ladcn, adCmdText
        lsSQL = lsSQL4
        ExecuteOTHERSQL lsSQL4, ladcn, adCmdText
        lsSQL = lsSQL7
        ExecuteOTHERSQL lsSQL7, ladcn, adCmdText
        lsSQL = lsSQL8
        ExecuteOTHERSQL lsSQL8, ladcn, adCmdText

        lsSQL = lsSQL5
        Set GetMgmtMissingTimeEmployees2 = _
                NewExecuteSQLWithConnection( _
                    lsSQL5 _
                    , Nothing _
                    , ladcn _
                    , adOpenStatic _
                    , adLockBatchOptimistic _
                    , adCmdText)

        lsSQL = lsSQL6
        ExecuteOTHERSQL lsSQL6, ladcn, adCmdText

        If (ladcn.State = adStateOpen) Then
            ladcn.Close
        End If

        Set ladcn = Nothing
    End If

    Exit Function
eh:
    Set ladcn = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetMgmtMissingTimeEmployees2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetMgmtMissingTimeEmployees2" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetMgmtMissingTimeEmployees2" _
                , Err.Number _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function

Public Function GetContractMissingTimeEmployees(ByVal vsPayrollGroupCode As String) As ADODB.Recordset
    TrsRaiseError MODULENAME, "GetContractMissingTimeEmployees", trErrObsoleteMethod, , , , True
End Function

Public Function GetContractMissingTimeEmployees2(ByVal vsPayrollGroupCode As String, ByVal vdtPayProcessDate As Date) As ADODB.Recordset
On Error GoTo eh
    Dim oADOHelper As utilADOHelper.cuADOHelper
    Dim lsSQL As String
    Dim lsExclusion1 As String, lsExclusion2 As String

    Set GetContractMissingTimeEmployees2 = Nothing

    lsExclusion1 = ""
    lsExclusion2 = ""
    
    If TRACING = 1 Then '//for testing to speed things up
'        lsExclusion1 = "AND D.DEPT_NUM IN ('006', '049', '069', '120', '14N', '165', '29S')"
'        lsExclusion2 = "WHERE A.DEPT_NUM IN ('006', '049', '069', '120', '14N', '165', '29S')"
        '//lsExclusion1 = "AND D.dept_num in ('30V', '31M', '30Y', '30L', '32N')"
        '//lsExclusion2 = "WHERE A.dept_num in ('30V', '31M', '30Y', '30L', '32N')"
    End If
           
    'get the SQL from SPMCache
    lsSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetContractMissingTimeEmployees2.sql")
        
    'set non-standard parameters (for testing)
    lsSQL = Replace(lsSQL, "{lsExclusion1}", lsExclusion1)
    lsSQL = Replace(lsSQL, "{lsExclusion2}", lsExclusion2)
    
    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelper
    
    oADOHelper.ADO.CreateCharacterParameter "PayGroup1", adChar, 1, vsPayrollGroupCode
    oADOHelper.ADO.CreateCharacterParameter "PayGroup2", adChar, 1, vsPayrollGroupCode
    oADOHelper.ADO.CreateCharacterParameter "PayGroup3", adChar, 1, vsPayrollGroupCode
    oADOHelper.ADO.CreateCharacterParameter "PayGroup4", adChar, 1, vsPayrollGroupCode
    oADOHelper.ADO.CreateCharacterParameter "PayGroup5", adChar, 1, vsPayrollGroupCode
    oADOHelper.ADO.CreateCharacterParameter "PayGroup6", adChar, 1, vsPayrollGroupCode
    oADOHelper.ADO.CreateCharacterParameter "PayGroup7", adChar, 1, vsPayrollGroupCode
    oADOHelper.ADO.CreateCharacterParameter "PayGroup8", adChar, 1, vsPayrollGroupCode
    oADOHelper.ADO.CreateNumericParameter "PayProcessDate1", adDate, 8, 0, vdtPayProcessDate
    oADOHelper.ADO.CreateNumericParameter "PayProcessDate2", adDate, 8, 0, vdtPayProcessDate
    oADOHelper.ADO.CreateNumericParameter "PayProcessDate3", adDate, 8, 0, vdtPayProcessDate
       
    'execute sql
    Set GetContractMissingTimeEmployees2 = oADOHelper.ExecuteSelect(lsSQL, adOpenStatic, adLockBatchOptimistic, adCmdText)
    
    Set oADOHelper = Nothing

    Exit Function
eh:
    Set oADOHelper = Nothing
    Set GetContractMissingTimeEmployees2 = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME, "GetContractMissingTimeEmployees2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME, "GetContractMissingTimeEmployees2", trErrdb_ExecuteSQL_F, , , DescArgsFrom(lsSQL), True
        Case Else
            TrsRaiseError MODULENAME, "GetContractMissingTimeEmployees2", Err.Number, , , DescArgsFrom(lsSQL), True
    End Select
End Function

Public Function GetDepartmentsAndPayPeriod(ByVal vsPayrollGroupCode As String _
                ) As ADODB.Recordset
    '//GetDepartmentsAndPayPeriod2 vsPayrollGroupCode, Date
    '//Err.Raise trErrObsoleteMethod, MODULENAME & ".GetDepartmentsAndPayPeriod", TrsGetErrDescription(LongAppErr(trErrObsoleteMethod))
    TrsRaiseError MODULENAME, "GetDepartmentsAndPayPeriod", trErrObsoleteMethod, , , , True
End Function

Public Function GetDepartmentsAndPayPeriod2(ByVal vsPayrollGroupCode As String _
                                            , ByVal vdtPayProcessDate As Date _
                ) As ADODB.Recordset
On Error GoTo eh
'//Connection
'//Dim ladcn As ADODB.Connection, ladrsParms As ADODB.Recordset
'SQL String Holders
Dim oADOHelper As utilADOHelper.cuADOHelper
Dim lsSQL As String, lsExclusion1 As String
Dim lsExclusion2 As String

    Set GetDepartmentsAndPayPeriod2 = Nothing

    lsExclusion1 = ""
    lsExclusion2 = ""

    If (UCase(Trim(vsPayrollGroupCode)) = "E" Or UCase(Trim(vsPayrollGroupCode)) = "N") Then
        lsExclusion1 = "        AND B.DEPT_NUM NOT IN (SELECT RG_VALUE_CD FROM METHDBA.TTR_RG_GROUP WHERE RG_CD = 167 AND APPLY_IND = 1)" & _
                    "       AND B.CMPY_CD <> 'MV1'"
        lsExclusion2 = "                AND D.DEPT_NUM NOT IN (SELECT RG_VALUE_CD FROM METHDBA.TTR_RG_GROUP WHERE RG_CD = 167 AND APPLY_IND = 1)" & _
                    "               AND D.CMPY_CD <> 'MV1'"
    End If

    If TRACING = 1 Then
'        lsExclusion1 = lsExclusion1 & vbCrLf & _
'                        "AND B.DEPT_NUM IN ('006', '049', '069', '120', '14N', '165', '29S')"
'        lsExclusion2 = lsExclusion2 & vbCrLf & _
'                        "AND D.DEPT_NUM IN ('006', '049', '069', '120', '14N', '165', '29S')"
    End If

    'get the SQL from SPMCache
    lsSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetDepartmentsAndPayPeriod2.sql")
        
    'set non-standard parameters (for testing)
    lsSQL = Replace(lsSQL, "{lsExclusion1}", lsExclusion1)
    lsSQL = Replace(lsSQL, "{lsExclusion2}", lsExclusion2)

    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelper
    
    oADOHelper.ADO.CreateCharacterParameter "PayGroup1", adChar, 1, vsPayrollGroupCode
    oADOHelper.ADO.CreateNumericParameter "PayProcessDate1", adDate, 8, 0, vdtPayProcessDate
    oADOHelper.ADO.CreateNumericParameter "PayProcessDate2", adDate, 8, 0, vdtPayProcessDate
    oADOHelper.ADO.CreateNumericParameter "PayProcessDate3", adDate, 8, 0, vdtPayProcessDate
    oADOHelper.ADO.CreateNumericParameter "PayProcessDate4", adDate, 8, 0, vdtPayProcessDate
    oADOHelper.ADO.CreateCharacterParameter "PayGroup2", adChar, 1, vsPayrollGroupCode
    oADOHelper.ADO.CreateCharacterParameter "PayGroup3", adChar, 1, vsPayrollGroupCode
           
    'execute sql
    Set GetDepartmentsAndPayPeriod2 = oADOHelper.ExecuteSelect(lsSQL, adOpenStatic, adLockBatchOptimistic, adCmdText)
    
    Set oADOHelper = Nothing

'    lsSQL = "SELECT A.PAYROLL_GROUP_CD" & _
'        "       , A.PAY_PER_START_DATE, A.PAY_PER_END_DATE" & _
'        "       , B.DEPT_NUM, B.DEPT_NAME" & _
'        "       , B.CMPY_CD, A.PAY_PROCESS_DATE" & _
'        " FROM  METHDBA.TTR_COLL_PERIODS A" & _
'        "       INNER JOIN METHDBA.TTR_DEPT B" & _
'        "       ON A.PAYROLL_GROUP_CD = B.PAYROLL_GROUP_CD" & _
'        "       AND A.PAYROLL_GROUP_CD = ?" & _
'        lsExclusion1 & _
'        "       AND (PAY_PROCESS_DATE = ?" & _
'        "       OR  PAY_PROCESS_DATE = CAST(? AS DATE) + 1 DAYS)" & _
'        "       AND B.TR_STATUS_EFF_DATE <= A.PAY_PER_END_DATE" & _
'        "       AND B.DEPT_NUM IN (SELECT   DISTINCT D.DEPT_NUM" & _
'        "                       FROM    METHDBA.TTR_EMPLOYEE C" & _
'        "                               INNER JOIN METHDBA.TTR_DEPT D" & _
'        "                               ON D.DEPT_NUM = C.DEPT_NUM" & _
'        "                               AND D.PAYROLL_GROUP_CD = ?" & _
'        lsExclusion2 & _
'        "                               AND ((C.EMP_STATUS = 'A'" & _
'        "                               AND C.EMPLOY_DATE <= A.PAY_PER_END_DATE)" & _
'        "                               OR  (C.EMP_STATUS = 'T'" & _
'        "                               AND C.STATUS_EFF_DATE >= A.PAY_PER_START_DATE))" & _
'        "                               AND C.DEPT_EFF_DATE <= A.PAY_PER_END_DATE)" & _
'        " ORDER BY B.DEPT_NUM, A.PAY_PER_START_DATE" & _
'        " FOR FETCH ONLY WITH UR;"

'    Set ladrsParms = CreateParameterRecordset()
'    AddInputParameterRecordset ladrsParms, 1, "", adChar, 1, vsPayrollGroupCode
'    AddInputParameterRecordset ladrsParms, 2, "", adDate, 0, vdtPayProcessDate
'    AddInputParameterRecordset ladrsParms, 3, "", adDate, 0, vdtPayProcessDate
'    AddInputParameterRecordset ladrsParms, 4, "", adChar, 1, vsPayrollGroupCode
'
'    Set GetDepartmentsAndPayPeriod2 = NewExecuteSQL(lsSQL _
'                                , ladrsParms _
'                                , adOpenStatic _
'                                , adLockReadOnly _
'                                , adCmdText)

    Exit Function
eh:
    Set oADOHelper = Nothing
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetDepartmentsAndPayPeriod"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetDepartmentsAndPayPeriod" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "GetDepartmentsAndPayPeriod" _
                , Err.Number _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function

Public Function GetPayProcessedDate(ByVal vsPayrollGroupCode As String, ByVal vdtScheduledDateTime As Date) As Date
On Error GoTo eh
'//Connection
'//Dim ladcn As ADODB.Connection, ladrsParms As ADODB.Recordset
'SQL String Holders
Dim oADOHelper As utilADOHelper.cuADOHelper
Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim ldtScheduledDate As Date
Dim ldtPayProcessDate As Date

    ldtScheduledDate = MidnightDate(vdtScheduledDateTime)
    ldtPayProcessDate = Date

    If TRACING = 1 Then
'        lsExclusion1 = lsExclusion1 & vbCrLf & _
'                        "AND B.DEPT_NUM IN ('006', '049', '069', '120', '14N', '165', '29S')"
'        lsExclusion2 = lsExclusion2 & vbCrLf & _
'                        "AND D.DEPT_NUM IN ('006', '049', '069', '120', '14N', '165', '29S')"
    End If

    'get the SQL from SPMCache
    lsSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetPayProcessedDate.sql")
        
    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelper
    
    oADOHelper.ADO.CreateCharacterParameter "PayGroup1", adChar, 1, vsPayrollGroupCode
    oADOHelper.ADO.CreateNumericParameter "ScheduledDate1", adDate, 8, 0, ldtScheduledDate
    oADOHelper.ADO.CreateNumericParameter "ScheduledDate2", adDate, 8, 0, ldtScheduledDate
    oADOHelper.ADO.CreateNumericParameter "ScheduledDate3", adDate, 8, 0, ldtScheduledDate
    oADOHelper.ADO.CreateNumericParameter "ScheduledDate4", adDate, 8, 0, ldtScheduledDate
           
    'execute sql
    Set rs = oADOHelper.ExecuteSelect(lsSQL, adOpenStatic, adLockBatchOptimistic, adCmdText)
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        ldtPayProcessDate = rs.Fields("PAY_PROCESS_DATE").Value
    End If
    
    GetPayProcessedDate = ldtPayProcessDate
    
    Set oADOHelper = Nothing
    Set rs = Nothing
    Exit Function
eh:
    Set oADOHelper = Nothing
    Set rs = Nothing
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME, "GetPayProcessedDate"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME, "GetPayProcessedDate", trErrdb_ExecuteSQL_F, , , DescArgsFrom(lsSQL), True
        Case Else
            TrsRaiseError MODULENAME, "GetPayProcessedDate", Err.Number, , , DescArgsFrom(lsSQL), True
    End Select

End Function

