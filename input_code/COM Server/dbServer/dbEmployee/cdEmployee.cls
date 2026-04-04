VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TrsGetModeEnum
    trGetBasicEmployee = 1
    trGetExtendedEmployee = 2
End Enum

Private Const m_strCMODULENAME As String = "cdEmployee"
Private Const m_DBSCRIPTSROOT As String = "C:\TRIS SOLUTION\dbScripts\TRIS_db2"
Private Const m_DBSCRIPTSPATH As String = "\dbEmployee\cdEmployee"
Private m_ResourceMgr As utilADOHelper.cuResourceMgr
Private Const m_strC1SPACESTRING As String = " "

'JimB 04/11/2003
'Made mass changes to the interfaces to allow multiple employees in the selection
'This does not affect the code of the calling module.
'The interfaces will accept 1 employee number as a long or an array of employees.
'Removed the functions that suffixed in 2.

Public Function IsValidEmployee(ByVal v_lngEmpNum As Long) As Boolean
    
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo IsValidEmployee_EH

    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT EMP_NUM From METHDBA.TTR_EMPLOYEE "
    l_strSQL = l_strSQL & "WHERE EMP_NUM = " & v_lngEmpNum
    
    Set l_objRS = New ADODB.Recordset
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    If l_objRS.BOF And l_objRS.EOF Then
        IsValidEmployee = False
    Else
        IsValidEmployee = True
    End If
    
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    Exit Function
    
IsValidEmployee_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "IsValidEmployee", , , , DescArgsFrom(l_strSQL), True

End Function
Public Function IsSecondaryWorkGroup(ByVal v_lngEmpNum As Long, _
                                    ByVal v_lngWgNum As Long, _
                                    ByVal v_dtRequestDate As Date) As Boolean
    
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo IsSecondaryWorkGroup_EH

    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT '1' From METHDBA.TTR_EMP_WG "
    l_strSQL = l_strSQL & "WHERE EMP_NUM = " & v_lngEmpNum
    l_strSQL = l_strSQL & " AND PRIMARY_IND = 'N' "
    l_strSQL = l_strSQL & "AND WG_NUM =  " & v_lngWgNum
    l_strSQL = l_strSQL & " AND '" & Format(v_dtRequestDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " BETWEEN EFF_START_DATE AND EFF_END_DATE "
    l_strSQL = l_strSQL & " UNION "
    l_strSQL = l_strSQL & "SELECT '1' FROM METHDBA.TTR_LINE_ASGNM A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_EMP_WG B "
    l_strSQL = l_strSQL & "WHERE A.EMP_NUM = " & v_lngEmpNum
    l_strSQL = l_strSQL & " AND '" & Format(v_dtRequestDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " BETWEEN A.START_DATE AND A.END_DATE "
    l_strSQL = l_strSQL & " AND B.EMP_NUM = A.EMP_NUM "
    l_strSQL = l_strSQL & " AND '" & Format(v_dtRequestDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " BETWEEN B.EFF_START_DATE AND B.EFF_END_DATE "
    l_strSQL = l_strSQL & " AND B.PRIMARY_IND = 'Y' "
    l_strSQL = l_strSQL & "AND B.WG_NUM <> A.WG_NUM"
    
    Set l_objRS = New ADODB.Recordset
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    If l_objRS.BOF And l_objRS.EOF Then
        IsSecondaryWorkGroup = False
    Else
        IsSecondaryWorkGroup = True
    End If
    
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    Exit Function
    
IsSecondaryWorkGroup_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "IsSecondaryWorkGroup", , , , DescArgsFrom(l_strSQL), True

End Function
Public Function IsSecondaryWorkGroup2(ByVal v_lngEmpNum As Long, _
                                    ByVal v_lngWgNum As Long, _
                                    ByVal v_dtStartDate As Date, _
                                    ByVal v_dtEndDate As Date) As Boolean
    

End Function
Public Function IsEmployeeInWorkGroup(ByVal v_lngEmpNum As Long, _
                                    ByVal v_lngWgNum As Long, _
                                    ByVal v_dtRequestDate) As Boolean
    
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo IsEmployeeInWorkGroup_EH

    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT '1' From METHDBA.TTR_EMP_WG "
    l_strSQL = l_strSQL & "WHERE EMP_NUM = " & v_lngEmpNum
    l_strSQL = l_strSQL & " AND WG_NUM =  " & v_lngWgNum
    l_strSQL = l_strSQL & " AND '" & Format(v_dtRequestDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " BETWEEN EFF_START_DATE AND EFF_END_DATE "
    l_strSQL = l_strSQL & " UNION "
    l_strSQL = l_strSQL & "SELECT '1' FROM METHDBA.TTR_LINE_ASGNM A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_EMP_WG B "
    l_strSQL = l_strSQL & "WHERE A.EMP_NUM = " & v_lngEmpNum
    l_strSQL = l_strSQL & " AND '" & Format(v_dtRequestDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " BETWEEN A.START_DATE AND A.END_DATE "
    l_strSQL = l_strSQL & " AND B.EMP_NUM = A.EMP_NUM "
    l_strSQL = l_strSQL & " AND '" & Format(v_dtRequestDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " BETWEEN B.EFF_START_DATE AND B.EFF_END_DATE "
    l_strSQL = l_strSQL & "AND B.WG_NUM <> A.WG_NUM"
    
    Set l_objRS = New ADODB.Recordset
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    If l_objRS.BOF And l_objRS.EOF Then
        IsEmployeeInWorkGroup = False
    Else
        IsEmployeeInWorkGroup = True
    End If
    
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    Exit Function
    
IsEmployeeInWorkGroup_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "IsEmployeeInWorkGroup", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetDepartment(ByVal v_lngEmpNum As Long, ByVal v_datDate As Date) As String
    
    'This is used by busCollectionPeriods
    'This is lousy code and should be removed.
    'Need to look at busCollectionPeriods for a better solution
    'For now, changing code to perform 1 SQL statement versus 2
    
    'Jimb 04/21/2003
    'Removed the code from busCollectionPeriods, because it was dead
    'This seems to be dead code also
    'It is only called from busEmployee from a method with the same name
    'That code seems dead also.
    'I can find any other code that calls it
    
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetDepartment_EH

    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    
    l_strSQL = "SELECT CASE " & _
                    "WHEN DEPT_EFF_DATE <=  '" & _
                        Format(v_datDate, "mm/dd/yyyy") & "' " & _
                        "THEN DEPT_NUM " & _
                    "WHEN PREV_DEPT_NUM = '' " & _
                        "THEN DEPT_NUM " & _
                    "ELSE PREV_DEPT_NUM " & _
                "END AS DEPT_NUM " & _
                "FROM METHDBA.TTR_EMPLOYEE " & _
                "WHERE EMP_NUM = " & str(v_lngEmpNum)
                    
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    If Not RecordsExist(l_objRS) Then
        GetDepartment = ""
    Else
        GetDepartment = l_objRS.Fields(0)
    End If
    
    Set l_objCN = Nothing
    
    Set l_objRS = Nothing
    Exit Function
    
GetDepartment_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDepartment", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetWorkReasons(ByVal v_aryEmployees As Variant) As ADODB.Recordset
TrsRaiseError m_strCMODULENAME, "GetWorkReasons", trErrObsoleteMethod, , , , True
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'    Dim l_strEmployeeSelection As String
'
'    On Error GoTo GetWorkReasons_EH
'
'    Set l_objCN = New ADODB.Connection
'    Set l_objRS = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
'
'    l_objCN.ConnectionString = l_objCM.GetConnectionString
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open
'
'    If IsArray(v_aryEmployees) Then
'        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
'    Else
'        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
'    End If
'
'    l_strSQL = "Select * From METHDBA.TTR_EC_WORK_REAS Where " & l_strEmployeeSelection
'
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'
'    Set GetWorkReasons = l_objRS
'
'    Set l_objRS = Nothing
'    Exit Function
'
'GetWorkReasons_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetWorkReasons", , , , DescArgsFrom(l_strSQL), True
'
End Function

Private Sub DeterminePriorUnionAttributes( _
                ByRef rsEmp As ADODB.Recordset)
'Must be positioned on a record
    Dim utilUnionMgr As dbCodeTables.cdCodeTables
    Dim l_strPriorUnionCd As String
    
    On Error GoTo eh

    l_strPriorUnionCd = Trim(rsEmp!PRIOR_UNION_CD)
    If l_strPriorUnionCd = Trim(rsEmp!UNION_CD) Then
        SetPriorUnionToCurrent rsEmp
        Exit Sub
    ElseIf l_strPriorUnionCd = "" Then
        Exit Sub
    End If
    
    Set utilUnionMgr = CreateObject("dbCodeTables.cdCodeTables")
    If utilUnionMgr.UnionCodeIsValid(l_strPriorUnionCd) Then
        rsEmp!PRIOR_EXEC_CON_IND = utilUnionMgr.UnionExecutiveContractIndicator(l_strPriorUnionCd)
        rsEmp!PRIOR_UNION_BN_PKG = utilUnionMgr.UnionBenefitPackageIndicator(l_strPriorUnionCd)
    Else
        SetPriorUnionToCurrent rsEmp
    End If
    Set utilUnionMgr = Nothing
    Exit Sub
    
eh:
    Set utilUnionMgr = Nothing
    TrsRaiseError m_strCMODULENAME, "DeterminePriorUnionAttributes"
End Sub

Private Sub SetPriorUnionToCurrent(rsEmp As ADODB.Recordset)

    If Not rsEmp.BOF And Not rsEmp.EOF Then
        rsEmp!PRIOR_EXEC_CON_IND = rsEmp!EXEC_CONTRACT_IND
        rsEmp!PRIOR_UNION_BN_PKG = rsEmp!UNION_BEN_PKG_IND
    End If
    
End Sub

Private Sub DeterminePreviousDepartmentAttributes( _
                ByRef rsEmp As ADODB.Recordset)
    Dim utilDeptMgr As dbDepartment.cdDepartment
    Dim l_strPreviousDepartmentNumber As String
    
    On Error GoTo eh

    l_strPreviousDepartmentNumber = Trim(rsEmp!PREV_DEPT_NUM)
    If l_strPreviousDepartmentNumber = Trim(rsEmp!DEPT_NUM) Then
        SetPreviousDepartmentToCurrent rsEmp
        Exit Sub
    ElseIf l_strPreviousDepartmentNumber = "" Then
        Exit Sub
    End If
    
    Set utilDeptMgr = CreateObject("dbDepartment.cdDepartment")
    If utilDeptMgr.IsDepartmentValid(l_strPreviousDepartmentNumber) Then
        rsEmp!PREV_CMPY_CD = utilDeptMgr.GetCompanyCode(l_strPreviousDepartmentNumber)
        rsEmp!PREV_INTERFC_TO_CD = utilDeptMgr.GetSystemInterfaceCode(l_strPreviousDepartmentNumber)
    Else
        SetPreviousDepartmentToCurrent rsEmp
    End If
    Set utilDeptMgr = Nothing
    Exit Sub
    
eh:
    Set utilDeptMgr = Nothing
    TrsRaiseError m_strCMODULENAME, "DeterminePreviousDepartmentAttributes"

End Sub

Private Sub SetPreviousDepartmentToCurrent( _
                ByRef rsEmp As ADODB.Recordset)

    If Not rsEmp.BOF And Not rsEmp.EOF Then
        rsEmp!PREV_CMPY_CD = rsEmp!CMPY_CD
        rsEmp!PREV_INTERFC_TO_CD = rsEmp!INTERFACE_TO_CD
    End If
    
End Sub

Private Sub DeterminePriorOTTypeAttributes( _
                ByRef rsEmp As ADODB.Recordset)
'Must be positioned on a record
    Dim utilJobMgr As dbCodeTables.cdCodeTables
    Dim l_strOTType As String
    Dim l_strPriorOTType As String
    Dim l_strJobNum As String
    Dim l_strPriorJobNum As String
    
On Error GoTo eh
    l_strJobNum = Trim(rsEmp!JOB_NUM)
    l_strPriorJobNum = Trim(rsEmp!PRIOR_JOB_NUM)
    l_strOTType = Trim(rsEmp!OT_TYPE)
    If l_strPriorJobNum = l_strJobNum Then
        rsEmp!PRIOR_OT_TYPE = l_strOTType
        Exit Sub
    ElseIf l_strPriorJobNum = "" Then
        Exit Sub
    End If

    Set utilJobMgr = CreateObject("dbCodeTables.cdCodeTables")
    If utilJobMgr.JobRecordExists(l_strPriorJobNum) Then
        rsEmp!PRIOR_OT_TYPE = utilJobMgr.JobOvertimeType(l_strPriorJobNum)
    Else
        rsEmp!PRIOR_OT_TYPE = l_strOTType
    End If
    Set utilJobMgr = Nothing
    Exit Sub
    
eh:
    Set utilJobMgr = Nothing
    TrsRaiseError m_strCMODULENAME, "DeterminePriorOTTypeAttributes"

End Sub

Public Function GetEmployee( _
        ByVal v_aryEmployees As Variant, _
        Optional ByVal v_lngMode As TrsGetModeEnum = 2 _
        ) As ADODB.Recordset
    
    Dim l_objRS As ADODB.Recordset
    
    On Error GoTo eh

    Set l_objRS = GetEmployeeBasic(v_aryEmployees)

'    If v_lngMode = 2 Then
'        If RecordsExist(l_objRS) Then
'            l_objRS.MoveFirst
'            While Not l_objRS.EOF
'                DeterminePriorUnionAttributes l_objRS
'                DeterminePreviousDepartmentAttributes l_objRS
'                DeterminePriorOTTypeAttributes l_objRS
'                l_objRS.MoveNext
'            Wend
'        End If
'    End If

    If RecordsExist(l_objRS) Then
        l_objRS.MoveFirst
    End If
    
    Set GetEmployee = l_objRS
    
    Set l_objRS = Nothing
    Exit Function

eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployee", , , , , True

End Function

Private Function GetEmployeeBasic( _
            ByVal v_aryEmployees As Variant _
            ) As ADODB.Recordset
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    l_strSQL = "SELECT " & _
                " A.EMP_NUM " & _
                ",A.EMP_NAME " & _
                ", A.SEC_LEVEL " & _
                ",A.LOC_CD " & _
                ",A.HOME_PHONE_NUM " & _
                ",A.DEPT_NUM " & _
                ",A.PREV_DEPT_NUM " & _
                ",CASE WHEN A.DEPT_EFF_DATE = '0001-01-01' " & _
                     " THEN '0100-01-01' " & _
                     " Else A.DEPT_EFF_DATE " & _
                     "END As DEPT_EFF_DATE " & _
                ",A.JOB_NUM " & _
                ",CASE WHEN A.JOB_EFF_DATE = '0001-01-01' " & _
                    " THEN '0100-01-01' " & _
                    " Else A.JOB_EFF_DATE " & _
                    "END As JOB_EFF_DATE " & _
                ",A.PRIOR_JOB_NUM " & _
                ",CASE WHEN A.PRIOR_JOB_EFF_DT = '0001-01-01' " & _
                    " THEN '0100-01-01' " & _
                    " Else A.PRIOR_JOB_EFF_DT " & _
                    "END As PRIOR_JOB_EFF_DT "
    l_strSQL = l_strSQL & _
                ",B.ACCT_TYPE " & _
                ",B.ACCT_DIV_NUM " & _
                ",B.PAYROLL_GROUP_CD " & _
                ",B.INTERFACE_TO_CD " & _
                ",COALESCE (H.INTERFACE_TO_CD,' ')  AS PREV_INTERFC_TO_CD " & _
                ",A.SHIFT_IND " & _
                ",A.MGMT_BEN_PKG_IND " & _
                ",A.EMP_STATUS " & _
                ",CASE WHEN A.STATUS_EFF_DATE = '0001-01-01' " & _
                    " THEN '0100-01-01' " & _
                    " Else A.STATUS_EFF_DATE " & _
                    "END As STATUS_EFF_DATE " & _
                ",SUBSTR(A.JOB_NUM,1,2) AS UNION_CD " & _
                ",SUBSTR(PRIOR_JOB_NUM,1,2) AS PRIOR_UNION_CD " & _
                ",C.EXEC_CONTRACT_IND " & _
                ",COALESCE (I.EXEC_CONTRACT_IND,' ') AS PRIOR_EXEC_CON_IND " & _
                ",C.UNION_BEN_PKG_IND " & _
                ",COALESCE (I.UNION_BEN_PKG_IND,'    ') AS PRIOR_UNION_BN_PKG " & _
                ",B.CMPY_CD "
    l_strSQL = l_strSQL & _
                ",COALESCE (H.CMPY_CD,'   ') AS PREV_CMPY_CD " & _
                ",E.AR_OT_EXCLUSN " & _
                ",NTE1.EFFDT " & _
                ",A.DIVISION_NUM " & _
                ",A.FUNCTION_NUM " & _
                ",' ' AS OT_TYPE " & _
                ",' ' AS PRIOR_OT_TYPE " & _
                ",COALESCE(G.AR_OT_EXCLUSN,'    ') AS PR_AR_OT_EXCLUSN " & _
                ",COALESCE(G.EFFDT,DATE('0100-01-01')) AS PRIOR_EFFDT "
    l_strSQL = l_strSQL & _
                ",A.OFFCR_FL, case when A.EMPLOY_DATE = A.ORIG_HIRE_DATE then A.EMPLOY_DATE else A.EMPLOY_DATE - 15 days end as EMPLOY_DATE "
    l_strSQL = l_strSQL & _
            "FROM METHDBA.TTR_EMPLOYEE A " & _
            "Inner Join " & _
                " METHDBA.TTR_DEPT B " & _
            "ON  B.DEPT_NUM = A.DEPT_NUM " & _
            "Inner Join " & _
                " METHDBA.TTR_UNIONS C " & _
            "ON C.UNION_CD = SUBSTR(A.JOB_NUM,1,2) " & _
            "Inner Join " & _
            " ( SELECT JOBCODE ,MAX(EFFDT) AS EFFDT " & _
                "From METHDBA.TTR_EMPLOYEE A " & _
                    "Inner Join " & _
                        " METHDBA.VJOBCODE_TBL_TR B " & _
                    "ON   B.JOBCODE = A.JOB_NUM " & _
                    "AND EFFDT <= CURRENT DATE " & _
                "WHERE " & l_strEmployeeSelection & _
                "Group By JOBCODE ) NTE1 " & _
            "ON NTE1.JOBCODE = A.JOB_NUM "
    l_strSQL = l_strSQL & _
            "INNER JOIN " & _
                " METHDBA.VJOBCODE_TBL_TR E " & _
            "ON  E.JOBCODE = NTE1.JOBCODE " & _
            "AND E.EFFDT = NTE1.EFFDT " & _
            "LEFT OUTER JOIN " & _
            " ( SELECT JOBCODE ,MAX(EFFDT) AS EFFDT " & _
                "From METHDBA.TTR_EMPLOYEE A " & _
                    "Inner Join " & _
                        " METHDBA.VJOBCODE_TBL_TR B " & _
                    "ON  B.JOBCODE = A.PRIOR_JOB_NUM " & _
                    "AND EFFDT <= CURRENT DATE " & _
                "WHERE " & l_strEmployeeSelection & _
                "Group By JOBCODE ) NTE2 " & _
            "ON NTE2.JOBCODE = A.PRIOR_JOB_NUM "
    l_strSQL = l_strSQL & _
            "LEFT OUTER JOIN " & _
                " METHDBA.VJOBCODE_TBL_TR G " & _
            "ON  G.JOBCODE = NTE2.JOBCODE " & _
            "AND G.EFFDT = NTE2.EFFDT " & _
            "LEFT OUTER JOIN " & _
                " METHDBA.TTR_DEPT H " & _
            "ON  H.DEPT_NUM = A.PREV_DEPT_NUM " & _
            "LEFT OUTER JOIN " & _
                " METHDBA.TTR_UNIONS I " & _
            "ON  I.UNION_CD = SUBSTR(A.PRIOR_JOB_NUM,1,2) "
    l_strSQL = l_strSQL & _
            "WHERE " & l_strEmployeeSelection

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeBasic = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeBasic", , , , DescArgsFrom(l_strSQL)

End Function

Public Function GetEmployeesInDepartment( _
            ByVal v_aryDepartments As Variant, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False, _
            Optional ByVal v_datStartDate As Date = "0100-01-01", _
            Optional ByVal v_datEndDate As Date = "0100-01-01" _
            ) As ADODB.Recordset
        
'    Dim l_lngStart As Long
    Dim l_datCutOffChargeDate As Date
    
On Error GoTo GetEmployeesInDepartment_EH
    
'    l_lngStart = LBound(v_aryDepartments)
'    If v_aryDepartments(l_lngStart) <> "CON" Then
'        If Val(v_aryDepartments(l_lngStart)) < 1 Then
'            Err.Raise trErrdbGetEmptyArrDept
'        Else
'        End If
'    End If
    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
   
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set GetEmployeesInDepartment = GetEmployeesInDepartmentACTData(v_aryDepartments, TrTable_Actual, v_bolActiveEmployeesOnly, v_datStartDate, v_datEndDate)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set GetEmployeesInDepartment = GetEmployeesInDepartmentACTData(v_aryDepartments, TrTable_History, v_bolActiveEmployeesOnly, v_datStartDate, v_datEndDate)
        Else
           Set GetEmployeesInDepartment = GetEmployeesInDepartmentACTData(v_aryDepartments, TrTable_ActualAndHistory, v_bolActiveEmployeesOnly, v_datStartDate, v_datEndDate) 'use both current and history data
        End If
    End If
    
    Exit Function
    
GetEmployeesInDepartment_EH:
    
    Select Case Err.Number
        Case trErrdbGetEmptyArrDept
            TrsRaiseError m_strCMODULENAME, "GetEmployeesInDepartment", trErrdbGetEmptyArrDept, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetEmployeesInDepartment", , , , , True
    End Select
    
End Function
Private Function GetEmployeesInDepartmentACTData( _
            ByVal v_aryDepartments As Variant, _
            ByVal v_TrActualTimeTable As TrActualTimeTable, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False, _
            Optional ByVal v_datStartDate As Date = "0100-01-01", _
            Optional ByVal v_datEndDate As Date = "0100-01-01" _
            ) As ADODB.Recordset
'PREM: MODIFIED THIS METHOD TO RETRIEVE EMPLOYEES
'BASED ON DEPARTMENT EFFECTIVE DATE AND
'START AND END DATES
On Error GoTo GetEmployeesInDepartmentACTData_EH
Dim lobjCM As util_DBConnectMgr.clsuDBConnectMgr
Dim ladcn As ADODB.Connection, ladrs As ADODB.Recordset
Dim lsSQL As String, lsInClausePartValueKey As String

    lsInClausePartValueKey = BuildInClause("PART_VALUE_KEY" _
                    , BuildArrayPartValueKeys(v_datStartDate, v_datEndDate) _
                        , True)

    lsSQL = "SELECT A.EMP_NUM, A.EMP_NAME" & _
        "       , A.DEPT_NUM, A.EMP_STATUS" & _
        "       , A.STATUS_EFF_DATE"

    Select Case (v_TrActualTimeTable)
        Case TrTable_Actual
          lsSQL = lsSQL & _
            "       , COALESCE(B.ACTUAL_TIME_SW, 0) AS ACTUAL_TIME_SW"
        Case TrTable_History
          lsSQL = lsSQL & _
            "       , COALESCE(C.ACTUAL_TIME_SW, 0) AS ACTUAL_TIME_SW"
        Case Else
          lsSQL = lsSQL & _
            "       , COALESCE(B.ACTUAL_TIME_SW, COALESCE(C.ACTUAL_TIME_SW,0)) AS ACTUAL_TIME_SW"
    End Select

    lsSQL = lsSQL & _
        " FROM  (SELECT EMP_NUM, EMP_NAME" & _
        "               , DEPT_NUM, EMP_STATUS, STATUS_EFF_DATE" & _
        "       FROM    METHDBA.TTR_EMPLOYEE" & _
        "       WHERE   " & BuildInClause("DEPT_NUM", v_aryDepartments, True) & _
        "       AND     DEPT_EFF_DATE <= '" & Format(v_datEndDate, "MM/dd/yyyy") & "'" & _
        "       UNION " & _
        "       SELECT  EMP_NUM, EMP_NAME" & _
        "               , PREV_DEPT_NUM AS DEPT_NUM, EMP_STATUS, STATUS_EFF_DATE" & _
        "       FROM    METHDBA.TTR_EMPLOYEE" & _
        "       WHERE   " & BuildInClause("PREV_DEPT_NUM", v_aryDepartments, True) & _
        "       AND     DEPT_EFF_DATE >= '" & Format(v_datStartDate, "MM/dd/yyyy") & "') AS A"

    If (v_TrActualTimeTable = TrTable_Actual) Or _
       (v_TrActualTimeTable = TrTable_ActualAndHistory) Then
        lsSQL = lsSQL & _
            "       LEFT OUTER JOIN" & _
            "       (SELECT DISTINCT EMP_NUM, 1 AS ACTUAL_TIME_SW" & _
            "       FROM    METHDBA.TTR_ACTUAL_TIME" & _
            "       WHERE   CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "MM/dd/yyyy") & "'" & _
            "       AND     '" & Format(v_datEndDate, "MM/dd/yyyy") & "') AS B" & _
            "       ON A.EMP_NUM = B.EMP_NUM"
    End If
    If (v_TrActualTimeTable = TrTable_History) Or _
       (v_TrActualTimeTable = TrTable_ActualAndHistory) Then
        lsSQL = lsSQL & _
            "       LEFT OUTER JOIN" & _
            "       (SELECT DISTINCT EMP_NUM, 1 AS ACTUAL_TIME_SW" & _
            "       FROM    METHDBA.TTR_ACTUAL_TIME_HS" & _
            "       WHERE   " & lsInClausePartValueKey & _
            "       AND     CHARGE_DATE BETWEEN '" & Format(v_datStartDate, "MM/dd/yyyy") & "'" & _
            "       AND     '" & Format(v_datEndDate, "MM/dd/yyyy") & "') AS C" & _
            "       ON A.EMP_NUM = C.EMP_NUM"
    End If

    If (v_bolActiveEmployeesOnly) Then
        lsSQL = lsSQL & _
            " WHERE (A.EMP_STATUS <> 'T'" & _
            " OR    (A.EMP_STATUS = 'T'" & _
            " AND   A.STATUS_EFF_DATE BETWEEN '" & Format(v_datStartDate, "MM/dd/yyyy") & "'" & _
            " AND   '" & Format(v_datEndDate, "MM/dd/yyyy") & "'))"

        If (v_TrActualTimeTable = TrTable_Actual) Or _
           (v_TrActualTimeTable = TrTable_ActualAndHistory) Then
           lsSQL = lsSQL & _
            " OR    B.ACTUAL_TIME_SW = 1 "
        End If

        If (v_TrActualTimeTable = TrTable_History) Or _
           (v_TrActualTimeTable = TrTable_ActualAndHistory) Then
           lsSQL = lsSQL & _
            " OR    C.ACTUAL_TIME_SW = 1 "
        End If
    End If

    lsSQL = lsSQL & " ORDER BY EMP_NAME" & _
            " FOR FETCH ONLY WITH UR;"

    'Connection
    Set ladcn = New ADODB.Connection

    Set lobjCM = New util_DBConnectMgr.clsuDBConnectMgr
    ladcn.ConnectionString = lobjCM.GetConnectionString("SHADOWDIRECT")
    Set lobjCM = Nothing

    ladcn.Open

    'Recordset
    Set ladrs = New ADODB.Recordset
    With ladrs
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .ActiveConnection = ladcn

        .Open lsSQL, ladcn
        .ActiveConnection = Nothing
    End With

    ladcn.Close
    Set ladcn = Nothing

    Set GetEmployeesInDepartmentACTData = ladrs

    Set ladrs = Nothing

    Exit Function
GetEmployeesInDepartmentACTData_EH:
    Set lobjCM = Nothing
    Set ladcn = Nothing
    Set ladrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeesInDepartmentACTData", , , , DescArgsFrom(lsSQL), True
    
End Function

'Get TTR_ACTUAL_TIME_HS cutoff Charge_Dt
Private Function GetActualTimeHSCutoffDate() As String
    
    Dim l_objCodeTables As dbCodeTables.cdCodeTables
    Dim l_varHSCutOffDate As Variant
    
    On Error GoTo GetActualTimeHSCutoffDate_EH
    
    Set l_objCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_varHSCutOffDate = l_objCodeTables.GetActualTimeHSCutoffDate(CKEYCUTOFFDATE)
    If IsDate(l_varHSCutOffDate) Then
        GetActualTimeHSCutoffDate = CStr(l_varHSCutOffDate)
    Else
        Err.Raise trErrdb_GetHSCutOffDateInfo
    End If
 
    Set l_objCodeTables = Nothing

    Exit Function
    
GetActualTimeHSCutoffDate_EH:
    Set l_objCodeTables = Nothing
    
    Select Case Err.Number
        Case trErrdb_GetHSCutOffDateInfo
            TrsRaiseError m_strCMODULENAME, "GetActualTimeHSCutoffDate", trErrdb_GetHSCutOffDateInfo, , , DescArgsFrom(CStr(l_varHSCutOffDate)), True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetActualTimeHSCutoffDate", , , , , True
    End Select

End Function


Public Function GetAllEmployees( _
        Optional v_bln_ActiveEmployeesOnly As Boolean = True _
            ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_datSixMonthsAgo As Date

    On Error GoTo GetAllEmployees_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "select A.EMP_NUM, A.EMP_NAME, A.UNION_CD,A.FUNCTION_NUM, "
    l_strSQL = l_strSQL & "A.SEC_LEVEL, A.LOC_CD, C.LOC_DESC,A.JOB_NUM, "
    l_strSQL = l_strSQL & "Case when A.JOB_EFF_DATE = '0001-01-01' then Date('0100-01-01') "
    l_strSQL = l_strSQL & " else A.JOB_EFF_DATE end  as JOB_EFF_DATE, "
    l_strSQL = l_strSQL & "Case when A.PRIOR_JOB_EFF_DT = '0001-01-01' then Date('0100-01-01') "
    l_strSQL = l_strSQL & " else A.PRIOR_JOB_EFF_DT  end as PRIOR_JOB_EFF_DT,  "
    l_strSQL = l_strSQL & "A.DEPT_NUM, B.DEPT_NAME,"
    l_strSQL = l_strSQL & "Case when A.DEPT_EFF_DATE = '0001-01-01' then Date('0100-01-01') "
    l_strSQL = l_strSQL & " else A.DEPT_EFF_DATE end  as DEPT_EFF_DATE, "
    l_strSQL = l_strSQL & "A.PREV_DEPT_NUM "
    l_strSQL = l_strSQL & "from METHDBA.TTR_EMPLOYEE A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_DEPT B, "
    l_strSQL = l_strSQL & "METHDBA.TTR_LOCATIONS C "
    l_strSQL = l_strSQL & "Where A.DEPT_NUM = B.DEPT_NUM "
    l_strSQL = l_strSQL & " and A.LOC_CD = C.LOC_CD"
    
    If v_bln_ActiveEmployeesOnly Then
        l_datSixMonthsAgo = DateAdd("m", -6, Date)
        l_strSQL = l_strSQL & " and (emp_status <> 'T' or "
        l_strSQL = l_strSQL & " STATUS_EFF_DATE >= '" & l_datSixMonthsAgo & "') "
    End If
    l_strSQL = l_strSQL & " order by emp_name, loc_cd"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetAllEmployees = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetAllEmployees_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllEmployees", , , , DescArgsFrom(l_strSQL)

End Function

Public Function GetBasicEmployeeInformation( _
        ByVal v_aryEmployees As Variant _
            ) As ADODB.Recordset
            
'JIMB 08/30/2002
'Field placeholders have been added to the SQL to show
' SCHED_CD, SCHED_SRC_CD, SCHED_DESC, ROT_CD, ROT_WEEK, ROT_DESC
' GetViewEmployeeInformation in busEmployee will fill the fields in
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo GetBasicEmployeeInformation_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    l_strSQL = " SELECT  A.STATUS_EFF_DATE, A.EMP_STATUS, A.PART_FULL_TIME_IND"
    l_strSQL = l_strSQL & " , A.REG_TEMP_STATUS, A.SHIFT_IND, B.OT_TYPE"
    l_strSQL = l_strSQL & " , A.WORK_PHONE_NUM, A.HOME_PHONE_NUM"
    l_strSQL = l_strSQL & " , A.SEC_LEVEL, C.SEC_LEVEL_DESC"
    l_strSQL = l_strSQL & " , A.CMPY_CD, D.CMPY_DESC_TXT"
    l_strSQL = l_strSQL & " , A.DIVISION_NUM, E.DIVISION_NAME"
    l_strSQL = l_strSQL & " , A.FUNCTION_NUM, F.FUNCTION_NAME"
    l_strSQL = l_strSQL & " , A.DEPT_NUM, G.DEPT_NAME"
    l_strSQL = l_strSQL & " , VALUE(NULLIF(A.DEPT_EFF_DATE,'0001-01-01'),DATE('0100-01-01')) AS DEPT_EFF_DATE"
    l_strSQL = l_strSQL & " , G.PAYROLL_GROUP_CD, A.PREV_DEPT_NUM"
    l_strSQL = l_strSQL & " , A.LOC_CD, H.LOC_DESC"
    l_strSQL = l_strSQL & " , A.JOB_NUM, B.JOB_DESC, A.JOB_EFF_DATE"
    l_strSQL = l_strSQL & " , A.PRIOR_JOB_NUM, A.PRIOR_JOB_EFF_DT"
    l_strSQL = l_strSQL & " , A.UNION_CD, SUBSTR(A.PRIOR_JOB_NUM,1,2) AS PRIOR_UNION_CD, I.UNION_LOCAL"
    l_strSQL = l_strSQL & " , I.UNION_DESC, I.UNION_BEN_PKG_IND"
    l_strSQL = l_strSQL & " , A.ORIG_HIRE_DATE, case when A.EMPLOY_DATE = A.ORIG_HIRE_DATE then A.EMPLOY_DATE else A.EMPLOY_DATE - 15 days end as EMPLOY_DATE "
    l_strSQL = l_strSQL & " , A.ACCPT_DT"
    l_strSQL = l_strSQL & " , VALUE(NULLIF(A.BENEFITS_DATE,'0001-01-01'),DATE('0100-01-01')) AS BENEFITS_DATE"
    l_strSQL = l_strSQL & " , A.CIPS_VAC_ELIG_DT, A.CIPS_SL_ELIG_DT"
    l_strSQL = l_strSQL & " , A.MGMT_BEN_PKG_IND, A.EXEC_CONTRACT_IND, SPACE(4) as EMP_BEN_PKG_IND, SPACE(1) as CIPS_IND, SPACE(1) as CILCO_IND"
    l_strSQL = l_strSQL & " , SMALLINT(0) as SCHED_CD, SMALLINT(0) as SCHED_SRC_CD , SPACE(30) as SCHED_DESC "
    l_strSQL = l_strSQL & " , SMALLINT(0) as ROT_CD, SMALLINT(0) as ROT_WEEK, SPACE(50) as ROT_DESC "
    l_strSQL = l_strSQL & " ,CASE WHEN VAC_ELIG_DT = '0001-01-01' THEN DATE('0100-01-01')  "
    l_strSQL = l_strSQL & "       ELSE VAC_ELIG_DT END AS VAC_ELIG_DT "
    l_strSQL = l_strSQL & " ,CASE WHEN SL_ELIG_DT = '0001-01-01' THEN DATE('0100-01-01')  "
    l_strSQL = l_strSQL & "       ELSE SL_ELIG_DT END AS SL_ELIG_DT "
    l_strSQL = l_strSQL & " FROM    METHDBA.TTR_EMPLOYEE A, METHDBA.TTR_JOBS B"
    l_strSQL = l_strSQL & "         , METHDBA.TTR_SECURITY_LEVEL C, METHDBA.TTR_COMPANY D"
    l_strSQL = l_strSQL & "         , METHDBA.TTR_DIVISION E, METHDBA.TTR_FUNCTION F"
    l_strSQL = l_strSQL & "         , METHDBA.TTR_DEPT G, METHDBA.TTR_LOCATIONS H"
    l_strSQL = l_strSQL & "         , METHDBA.TTR_UNIONS I"
    l_strSQL = l_strSQL & " WHERE  " & l_strEmployeeSelection
    l_strSQL = l_strSQL & " AND     A.JOB_NUM = B.JOB_NUM"
    l_strSQL = l_strSQL & " AND     A.SEC_LEVEL = C.SEC_LEVEL"
    l_strSQL = l_strSQL & " AND     A.CMPY_CD = D.CMPY_CD"
    l_strSQL = l_strSQL & " AND     A.DIVISION_NUM = E.DIVISION_NUM"
    l_strSQL = l_strSQL & " AND     A.FUNCTION_NUM = F.FUNCTION_NUM"
    l_strSQL = l_strSQL & " AND     A.DEPT_NUM = G.DEPT_NUM"
    l_strSQL = l_strSQL & " AND     A.LOC_CD = H.LOC_CD"
    l_strSQL = l_strSQL & " AND     A.UNION_CD = I.UNION_CD"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetBasicEmployeeInformation = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetBasicEmployeeInformation_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetBasicEmployeeInformation", , , , DescArgsFrom(l_strSQL)

End Function


Public Function GetEmployeePrimaryWorkGroup _
                (ByVal v_aryEmployees As Variant _
                ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo GetEmployeePrimaryWorkGroup_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    l_strSQL = "SELECT  A.WG_NUM, B.WG_DESC "
    l_strSQL = l_strSQL & ", A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM "
    l_strSQL = l_strSQL & ", CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP "
    l_strSQL = l_strSQL & ", CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM    METHDBA.TTR_EMP_WG A, METHDBA.TTR_WORK_GROUP B "
    l_strSQL = l_strSQL & "Where " & l_strEmployeeSelection & " "
    l_strSQL = l_strSQL & " AND     A.PRIMARY_IND = 'Y' "
    l_strSQL = l_strSQL & " AND     A.WG_NUM = B.WG_NUM "
    l_strSQL = l_strSQL & " AND  '" & Format(Date, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " BETWEEN  A.EFF_START_DATE AND A.EFF_END_DATE "
'    l_strSQL = l_strSQL & " AND  A.EFF_START_DATE <= '" & Format(Date, "mm/dd/yyyy") & "'"
'    l_strSQL = l_strSQL & " AND  A.EFF_END_DATE >= '" & Format(Date, "mm/dd/yyyy") & "'"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeePrimaryWorkGroup = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetEmployeePrimaryWorkGroup_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeePrimaryWorkGroup", , , , DescArgsFrom(l_strSQL)

End Function

Public Function GetEmployeeWorkGroups _
                (ByVal v_aryEmployees As Variant _
                ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo GetEmployeeWorkGroups_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    l_strSQL = "SELECT  A.WG_NUM, B.WG_DESC "
    l_strSQL = l_strSQL & ", A.PRIMARY_IND, A.EFF_START_DATE "
    l_strSQL = l_strSQL & ", A.EFF_END_DATE, A.TEMP_OT_BANK_IND "
    l_strSQL = l_strSQL & ", A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM "
    l_strSQL = l_strSQL & ", CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP "
    l_strSQL = l_strSQL & ", CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_EMP_WG A, METHDBA.TTR_WORK_GROUP B "
    l_strSQL = l_strSQL & "WHERE    A.WG_NUM = B.WG_NUM "
    l_strSQL = l_strSQL & " AND " & l_strEmployeeSelection & " "
    l_strSQL = l_strSQL & " ORDER BY A.EFF_START_DATE DESC"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeWorkGroups = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetEmployeeWorkGroups_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeWorkGroups", , , , DescArgsFrom(l_strSQL)

End Function

Public Function GetEmployeeReliefGroups _
                (ByVal v_aryEmployees As Variant _
                ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo GetEmployeeReliefGroups_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If


    l_strSQL = "SELECT  A.REL_GROUP_ID, B.REL_GROUP_DESC "
    l_strSQL = l_strSQL & ", B.EFF_YEAR, A.REL_TERM "
    l_strSQL = l_strSQL & ", A.REL_ASSIGN, A.REL_GROUP_POS "
    l_strSQL = l_strSQL & ", A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_REL_GRP_EMP A, METHDBA.TTR_RELIEF_GRPS B "
    l_strSQL = l_strSQL & "WHERE    A.REL_GROUP_ID = B.REL_GROUP_ID "
    l_strSQL = l_strSQL & " AND " & l_strEmployeeSelection & " "
    l_strSQL = l_strSQL & " ORDER BY B.EFF_YEAR, B.REL_GROUP_ID DESC"



    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeReliefGroups = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetEmployeeReliefGroups_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeReliefGroups", , , , DescArgsFrom(l_strSQL)

End Function

Public Function GetEmployeeQualifications _
                (ByVal v_aryEmployees As Variant _
                ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo GetEmployeeQualifications_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    l_strSQL = "SELECT  A.EMP_NUM, A.QUAL_CD, B.QUAL_DESC "
    l_strSQL = l_strSQL & ", A.EFF_START_DATE, A.EFF_END_DATE "
    l_strSQL = l_strSQL & ", A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_EMP_QUAL_SKILL A, METHDBA.TTR_QUALIFICATIONS B "
    l_strSQL = l_strSQL & "WHERE    A.QUAL_CD = B.QUAL_CD "
    l_strSQL = l_strSQL & " AND " & l_strEmployeeSelection & " "
    l_strSQL = l_strSQL & " AND A.EFF_START_DATE < current_date - 4 days "
    l_strSQL = l_strSQL & " AND A.EFF_END_DATE > current_date + 4 days "
    l_strSQL = l_strSQL & "ORDER BY A.EMP_NUM, A.QUAL_CD"



    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeQualifications = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetEmployeeQualifications_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeQualifications", , , , DescArgsFrom(l_strSQL)

End Function


Public Function GetEmployeeSkills _
                (ByVal v_aryEmployees As Variant _
                ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo GetEmployeeSkills_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If

    l_strSQL = "SELECT  A.EMP_NUM, A.SPEC_SKILL_CD, B.SPEC_SKILL_DESC "
    l_strSQL = l_strSQL & ", A.EFF_START_DATE, A.EFF_END_DATE "
    l_strSQL = l_strSQL & ", A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_EMP_QUAL_SKILL A, METHDBA.TTR_SPEC_SKILLS B "
    l_strSQL = l_strSQL & "WHERE    A.SPEC_SKILL_CD = B.SPEC_SKILL_CD "
    l_strSQL = l_strSQL & " AND " & l_strEmployeeSelection & " "
    l_strSQL = l_strSQL & " AND A.EFF_START_DATE < current_date - 4 days "
    l_strSQL = l_strSQL & " AND A.EFF_END_DATE > current_date + 4 days "
    l_strSQL = l_strSQL & "ORDER BY A.EMP_NUM, A.SPEC_SKILL_CD"



    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeSkills = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetEmployeeSkills_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeSkills", , , , DescArgsFrom(l_strSQL)

End Function


Public Function GetEmployeeBenefitsPackage( _
                    ByVal v_aryEmployees As Variant _
                ) As ADODB.Recordset
' Get all the attributes needed to determine an employees benefits_package_ind
' on a given charge_date.  See cbs.EmployeeBenefitPackage in busEmployee.
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmpPredicate As String
    
On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    
    If IsArray(v_aryEmployees) Then
        l_strEmpPredicate = BuildInClause("EMP_NUM", v_aryEmployees, False)
    Else
        l_strEmpPredicate = "EMP_NUM = " & CStr(v_aryEmployees)
    End If

    l_objCN.Open
    l_strSQL = "SELECT NE1.EMP_NUM, NE1.MGMT_BEN_PKG_IND,"
    l_strSQL = l_strSQL & " NE1.JOB_EFF_DATE, NE1.UNION_CD,"
    l_strSQL = l_strSQL & " COALESCE(LOJ1.UNION_BEN_PKG_IND, '') AS UNION_BEN_PKG_IND,"
    l_strSQL = l_strSQL & " COALESCE(LOJ1.EXEC_CONTRACT_IND, '') AS EXEC_CONTRACT_IND,"
    l_strSQL = l_strSQL & " NE1.PRIOR_UNION_CD,"
    l_strSQL = l_strSQL & " COALESCE(LOJ2.UNION_BEN_PKG_IND, '') AS PRIOR_UNION_BN_PKG,"
    l_strSQL = l_strSQL & " COALESCE(LOJ2.EXEC_CONTRACT_IND, '') AS PRIOR_EXEC_CON_IND"
    l_strSQL = l_strSQL & " FROM"
    l_strSQL = l_strSQL & " (SELECT EMP_NUM, MGMT_BEN_PKG_IND, JOB_EFF_DATE, Left(JOB_NUM,2) as UNION_CD,"
    l_strSQL = l_strSQL & " Left(PRIOR_JOB_NUM, 2) As PRIOR_UNION_CD"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EMPLOYEE"
    l_strSQL = l_strSQL & " WHERE " & l_strEmpPredicate & ") NE1"
    l_strSQL = l_strSQL & " LEFT OUTER JOIN"
    l_strSQL = l_strSQL & " METHDBA.TTR_UNIONS LOJ1"
    l_strSQL = l_strSQL & " ON LOJ1.UNION_CD = NE1.UNION_CD"
    l_strSQL = l_strSQL & " LEFT OUTER JOIN"
    l_strSQL = l_strSQL & " METHDBA.TTR_UNIONS LOJ2"
    l_strSQL = l_strSQL & " ON LOJ2.UNION_CD = NE1.PRIOR_UNION_CD"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing

    Set GetEmployeeBenefitsPackage = l_objRS
    Set l_objRS = Nothing
   
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeBenefitsPackage", , , , DescArgsFrom(l_strSQL), True
End Function


Public Function GetEmployeeEMPRV( _
                    ByVal v_aryEmployees As Variant _
                ) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_lngStart As Long
    Dim l_lngEnd As Long
    Dim l_lngIndex As Long
    
    On Error GoTo eh

    Set l_objRS = GetEmployeeEMPRVPrivate(v_aryEmployees)
    
    
    If RecordsExist(l_objRS) Then
        l_objRS.MoveFirst
        While Not l_objRS.EOF
            DeterminePreviousDepartmentAttributes l_objRS
            l_objRS.MoveNext
        Wend
    End If
    
    If RecordsExist(l_objRS) Then
        l_objRS.MoveFirst
    End If

    Set GetEmployeeEMPRV = l_objRS
    
    Set l_objRS = Nothing
    Exit Function

eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeEMPRV", , , , , True

End Function

Private Function GetEmployeeEMPRVPrivate( _
                    ByVal v_aryEmployees As Variant _
                ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If


    l_strSQL = "select A.EMP_NUM, A.LOC_CD, "
    l_strSQL = l_strSQL & "A.DEPT_NUM, A.PREV_DEPT_NUM, "
    l_strSQL = l_strSQL & _
        "VALUE(NULLIF(A.DEPT_EFF_DATE,'0001-01-01'),DATE('0100-01-01')) as DEPT_EFF_DATE, "
    l_strSQL = l_strSQL & "B.INTERFACE_TO_CD, "
    l_strSQL = l_strSQL & EncloseWithSingleQuote(String(1, m_strC1SPACESTRING))
    l_strSQL = l_strSQL & " AS PREV_INTERFC_TO_CD, "
    l_strSQL = l_strSQL & "B.CMPY_CD, "
    l_strSQL = l_strSQL & EncloseWithSingleQuote(String(3, m_strC1SPACESTRING))
    l_strSQL = l_strSQL & " AS PREV_CMPY_CD "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_EMPLOYEE A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_DEPT B "
    l_strSQL = l_strSQL & "WHERE " & l_strEmployeeSelection
    l_strSQL = l_strSQL & " AND A.DEPT_NUM = B.DEPT_NUM "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeEMPRVPrivate = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeEMPRVPrivate", , , , DescArgsFrom(l_strSQL)
End Function

Public Function GetEmployeeExcSchedInfo( _
                    ByVal v_aryEmployees As Variant _
                ) As ADODB.Recordset
' Get all the attributes needed to determine an employees benefits_package_ind
' on a given charge_date.  See cbs.EmployeeBenefitPackage in busEmployee.
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmpPredicate As String
    
On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    
    If IsArray(v_aryEmployees) Then
        l_strEmpPredicate = BuildInClause("EMP_NUM", v_aryEmployees, False)
    Else
        l_strEmpPredicate = "EMP_NUM = " & CStr(v_aryEmployees)
    End If

    l_objCN.Open
    l_strSQL = "SELECT NE1.EMP_NUM, NE1.MGMT_BEN_PKG_IND, "
    l_strSQL = l_strSQL & " NE1.JOB_EFF_DATE, NE1.UNION_CD, "
    l_strSQL = l_strSQL & " NE1.JOB_NUM, NE1.PRIOR_JOB_NUM, NE1.PRIOR_JOB_EFF_DT, "
    l_strSQL = l_strSQL & " COALESCE(LOJ1.UNION_BEN_PKG_IND, '') AS UNION_BEN_PKG_IND, "
    l_strSQL = l_strSQL & " COALESCE(LOJ1.EXEC_CONTRACT_IND, '') AS EXEC_CONTRACT_IND, "
    l_strSQL = l_strSQL & " NE1.PRIOR_UNION_CD,"
    l_strSQL = l_strSQL & " COALESCE(LOJ2.UNION_BEN_PKG_IND, '') AS PRIOR_UNION_BN_PKG, "
    l_strSQL = l_strSQL & " COALESCE(LOJ2.EXEC_CONTRACT_IND, '') AS PRIOR_EXEC_CON_IND "
    l_strSQL = l_strSQL & ",NE1.DEPT_EFF_DATE, NE1.DEPT_NUM, "
    l_strSQL = l_strSQL & "COALESCE(LOJ3.PAYROLL_GROUP_CD,' ') AS PAYROLL_GROUP_CD, "
    l_strSQL = l_strSQL & "COALESCE(LOJ3.INTERFACE_TO_CD,' ') AS INTERFACE_TO_CD, "
    l_strSQL = l_strSQL & "COALESCE(LOJ3.LOCK_STATUS,' ') AS LOCK_STATUS, "
    l_strSQL = l_strSQL & "COALESCE(LOJ3.CMPY_CD,' ') AS CMPY_CD, "
    l_strSQL = l_strSQL & "NE1.PREV_DEPT_NUM, "
    l_strSQL = l_strSQL & "COALESCE(LOJ4.PAYROLL_GROUP_CD,' ') AS PREV_PAY_GROUP, "
    l_strSQL = l_strSQL & "COALESCE(LOJ4.INTERFACE_TO_CD,' ') AS PREV_INTERFACE, "
    l_strSQL = l_strSQL & "COALESCE(LOJ4.LOCK_STATUS,' ') AS PREV_LOCK_STATUS, "
    l_strSQL = l_strSQL & "COALESCE(LOJ4.CMPY_CD,' ') AS PREV_CMPY_CD, "
    l_strSQL = l_strSQL & " case    "
    l_strSQL = l_strSQL & "        when LOJ6.RG_VALUE_CD is not null then 3 "
    l_strSQL = l_strSQL & "        when LOJ5.RG_VALUE_CD is not null then 6 "
    l_strSQL = l_strSQL & "        when LOJ1.EXEC_CONTRACT_IND = 'E' then 30    "
    l_strSQL = l_strSQL & "        when LOJ1.EXEC_CONTRACT_IND = 'C' then 6 "
    l_strSQL = l_strSQL & "        else 0   "
    l_strSQL = l_strSQL & " end as TIME_INCR,   "
    l_strSQL = l_strSQL & " case    "
    l_strSQL = l_strSQL & "    WHEN NE1.PREV_DEPT_NUM = ''    THEN  "
    l_strSQL = l_strSQL & "          case   "
    l_strSQL = l_strSQL & "             when LOJ6.RG_VALUE_CD is not null then 3    "
    l_strSQL = l_strSQL & "             when LOJ5.RG_VALUE_CD is not null then 6    "
    l_strSQL = l_strSQL & "          end    "
    l_strSQL = l_strSQL & "    when LOJ8.RG_VALUE_CD is not null then 3 "
    l_strSQL = l_strSQL & "    when LOJ7.RG_VALUE_CD is not null then 6 "
    l_strSQL = l_strSQL & "    when LOJ1.EXEC_CONTRACT_IND = 'E' then 30    "
    l_strSQL = l_strSQL & "    when LOJ1.EXEC_CONTRACT_IND = 'C' then 6 "
    l_strSQL = l_strSQL & "    else 0   "
    l_strSQL = l_strSQL & " end As PRIOR_TIME_INCR, LOC_CD, AR_OT_EXCLUSN, EMPLOY_DATE  "
    l_strSQL = l_strSQL & " FROM"
    l_strSQL = l_strSQL & " (SELECT EMP_NUM, MGMT_BEN_PKG_IND,  "
    l_strSQL = l_strSQL & " JOB_EFF_DATE, JOB_NUM, PRIOR_JOB_NUM, "
    l_strSQL = l_strSQL & " CASE "
    l_strSQL = l_strSQL & "    WHEN PRIOR_JOB_EFF_DT = DATE('0001-01-01') THEN DATE('0100-01-01') "
    l_strSQL = l_strSQL & "    WHEN PRIOR_JOB_EFF_DT IS NULL THEN DATE('0100-01-01') "
    l_strSQL = l_strSQL & "    ELSE PRIOR_JOB_EFF_DT  "
    l_strSQL = l_strSQL & " END AS PRIOR_JOB_EFF_DT, "
    l_strSQL = l_strSQL & " Left(JOB_NUM,2) as UNION_CD, Left(PRIOR_JOB_NUM, 2) As PRIOR_UNION_CD "
    l_strSQL = l_strSQL & ",DEPT_NUM, PREV_DEPT_NUM, "
    l_strSQL = l_strSQL & "CASE "
    l_strSQL = l_strSQL & "   WHEN DEPT_EFF_DATE = DATE('0001-01-01') THEN DATE('0100-01-01') "
    l_strSQL = l_strSQL & "   ELSE DEPT_EFF_DATE END AS DEPT_EFF_DATE, LOC_CD, case when EMPLOY_DATE = ORIG_HIRE_DATE then EMPLOY_DATE else EMPLOY_DATE - 15 days end as EMPLOY_DATE "
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EMPLOYEE "
    l_strSQL = l_strSQL & " WHERE " & l_strEmpPredicate & ") NE1 "
    l_strSQL = l_strSQL & " LEFT OUTER JOIN "
    l_strSQL = l_strSQL & " METHDBA.TTR_UNIONS LOJ1 "
    l_strSQL = l_strSQL & " ON LOJ1.UNION_CD = NE1.UNION_CD "
    l_strSQL = l_strSQL & " LEFT OUTER JOIN "
    l_strSQL = l_strSQL & " METHDBA.TTR_UNIONS LOJ2 "
    l_strSQL = l_strSQL & " ON LOJ2.UNION_CD = NE1.PRIOR_UNION_CD "
    l_strSQL = l_strSQL & "LEFT OUTER JOIN "
    l_strSQL = l_strSQL & " METHDBA.TTR_DEPT LOJ3 "
    l_strSQL = l_strSQL & " ON LOJ3.DEPT_NUM = NE1.DEPT_NUM "
    l_strSQL = l_strSQL & "LEFT OUTER JOIN "
    l_strSQL = l_strSQL & " METHDBA.TTR_DEPT LOJ4 "
    l_strSQL = l_strSQL & "ON LOJ4.DEPT_NUM = NE1.PREV_DEPT_NUM "
    l_strSQL = l_strSQL & "LEFT OUTER JOIN "
    l_strSQL = l_strSQL & " METHDBA.TTR_RG_GROUP LOJ5 "
    l_strSQL = l_strSQL & " ON LOJ5.RG_CD = 165 "
    l_strSQL = l_strSQL & "   AND LOJ5.RG_VALUE_CD = NE1.DEPT_NUM "
    l_strSQL = l_strSQL & "LEFT OUTER JOIN "
    l_strSQL = l_strSQL & " METHDBA.TTR_RG_GROUP LOJ6 "
    l_strSQL = l_strSQL & " ON LOJ6.RG_CD = 166 "
    l_strSQL = l_strSQL & "   AND LOJ6.RG_VALUE_CD = NE1.DEPT_NUM "
    l_strSQL = l_strSQL & "LEFT OUTER JOIN "
    l_strSQL = l_strSQL & " METHDBA.TTR_RG_GROUP LOJ7 "
    l_strSQL = l_strSQL & " ON LOJ7.RG_CD = 165 "
    l_strSQL = l_strSQL & "   AND LOJ7.RG_VALUE_CD = NE1.PREV_DEPT_NUM "
    l_strSQL = l_strSQL & "LEFT OUTER JOIN "
    l_strSQL = l_strSQL & " METHDBA.TTR_RG_GROUP LOJ8 "
    l_strSQL = l_strSQL & " ON LOJ8.RG_CD = 166 "
    l_strSQL = l_strSQL & "   AND LOJ8.RG_VALUE_CD = NE1.PREV_DEPT_NUM "
    l_strSQL = l_strSQL & " Inner Join "
    l_strSQL = l_strSQL & " ( SELECT JOBCODE ,MAX(EFFDT) AS EFFDT "
    l_strSQL = l_strSQL & "   From METHDBA.TTR_EMPLOYEE A "
    l_strSQL = l_strSQL & "   Inner Join "
    l_strSQL = l_strSQL & "        METHDBA.VJOBCODE_TBL_TR B "
    l_strSQL = l_strSQL & "    ON   B.JOBCODE = A.JOB_NUM "
    l_strSQL = l_strSQL & "    AND EFFDT <= CURRENT DATE "
    l_strSQL = l_strSQL & " WHERE " & l_strEmpPredicate & " "
    l_strSQL = l_strSQL & " Group By JOBCODE ) NTE1 "
    l_strSQL = l_strSQL & " ON NTE1.JOBCODE = NE1.JOB_NUM "
    l_strSQL = l_strSQL & " Inner Join "
    l_strSQL = l_strSQL & " METHDBA.VJOBCODE_TBL_TR E "
    l_strSQL = l_strSQL & " ON  E.JOBCODE = NTE1.JOBCODE "
    l_strSQL = l_strSQL & " AND E.EFFDT = NTE1.EFFDT "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing

    Set GetEmployeeExcSchedInfo = l_objRS
    Set l_objRS = Nothing
    
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeExcSchedInfo", , , , DescArgsFrom(l_strSQL), True
End Function
'TODO: obsolete
''BEGIN PREM TRIS CONVERSION PHASE 3
'Private Function GetCaseNumbersSQL( _
'            ByVal varyEmployees As Variant _
'            , ByVal vdtStartDate As Date _
'            , ByVal vdtEndDate As Date _
'            ) As String
'Dim lsSQL As String, lsEmpNumSQL As String
'Dim lsChargeDateSQL As String, lsCaseGroupSQL As String
'Dim lsCaseNumSQL As String, lsCaseYearSQL As String
'Dim lsHoursSQL As String, lsAdjustmentsSQL As String
'Dim lsActualTimeHSSQL As String, lsActualTimeSQL As String
'Dim lsEmpSchedExcSQL As String, lsEmployeeSelection As String
'Dim ldtHistoryEndDate As Date, ldtActualTimeStartDate As Date
'Dim llActualTimeTables As Long, lsPartValKeySel As String
'Dim lsNewAdjTableSQL As String
''Prem Case Number Bug Fix. When three years or more are involved
''the Part Value Key in SQL statement uses only the Start Date Year and
''the History End Date Year and does not account for the years in the middle.
''e.g. If Start Date Year is 2004 and end date year is 2006, then the logic determines
''the Part Value Key IN Clause as "AND     PART_VALUE_KEY IN ('4', '6'). It skips 5.
'Dim liIndex As Integer, lsTmpString As String
'
''PER USER REQUEST CASE NUMBERS WILL GREATER THAN 99 AND
'' THE START DATE WILL ALWAYS BE
'    vdtStartDate = #1/1/2004#
'
'    llActualTimeTables = DetermineActualTimeTables( _
'                            vdtStartDate _
'                            , vdtEndDate _
'                            , ldtHistoryEndDate _
'                            , ldtActualTimeStartDate)
'
'    If (IsArray(varyEmployees)) Then
'        lsEmployeeSelection = BuildInClause("EMP_NUM" _
'                                , varyEmployees _
'                                , False) & " "
'    Else
'        lsEmployeeSelection = "EMP_NUM = " & CStr(CLng(varyEmployees)) & " "
'    End If
'
'    If (Year(vdtStartDate) = Year(ldtHistoryEndDate)) Then
'        lsPartValKeySel = " = '" & Right(str(Year(vdtStartDate)), 1) & "' "
'    Else
'        'Begin Prem TRIS 5.6 Bug Fix
'        lsTmpString = ""
'        lsPartValKeySel = " IN ('"
'
'        For liIndex = Year(vdtStartDate) To Year(ldtHistoryEndDate)
'            If (liIndex = Year(vdtStartDate)) Then
'                lsTmpString = Right(CStr(liIndex), 1) & "'"
'            Else
'                lsTmpString = lsTmpString & ", '" & _
'                    Right(CStr(liIndex), 1) & "'"
'            End If
'        Next liIndex
'
'        lsPartValKeySel = " IN ('" & lsTmpString & ")"
'        'End Prem TRIS 5.6 Bug Fix
'    End If
'
'    lsSQL = "SELECT EMP_NUM, CASE_GROUP" & _
'        "           , REFERENCE_NUM, CASE_YEAR" & _
'        "           , MIN(CHARGE_DATE) AS FIRST_DATE, MAX(CHARGE_DATE) AS LAST_DATE" & _
'        " FROM      ("
'
'    lsSQL = lsSQL & _
'        "SELECT EMP_NUM, CASE_GROUP" & _
'        "       , CASE_NUM AS REFERENCE_NUM, CASE_YEAR" & _
'        "       , CHARGE_DATE" & _
'        " FROM  (SELECT    "
'
'    lsEmpNumSQL = " CASE   WHEN NE1.EMP_NUM IS NOT NULL" & _
'        "   THEN NE1.EMP_NUM" & _
'        "   WHEN NE2.EMP_NUM IS NOT NULL" & _
'        "   THEN NE2.EMP_NUM" & _
'        "   WHEN NE3.EMP_NUM IS NOT NULL" & _
'        "   THEN NE3.EMP_NUM" & _
'        "   ELSE NE4.EMP_NUM END AS EMP_NUM"
'
'    lsChargeDateSQL = " , CASE    WHEN NE1.EMP_NUM IS NOT NULL" & _
'        "   THEN NE1.DATE_WORKED" & _
'        "   WHEN NE2.EMP_NUM IS NOT NULL" & _
'        "   THEN NE2.CHARGE_DATE" & _
'        "   WHEN NE3.EMP_NUM IS NOT NULL" & _
'        "   THEN NE3.CHARGE_DATE" & _
'        "   ELSE NE4.CHARGE_DATE END AS CHARGE_DATE"
'
'    lsCaseGroupSQL = " , CASE    WHEN NE1.EMP_NUM IS NOT NULL" & _
'        "   THEN NE1.CASE_GROUP" & _
'        "   WHEN NE2.EMP_NUM IS NOT NULL" & _
'        "   THEN NE2.CASE_GROUP" & _
'        "   WHEN NE3.EMP_NUM IS NOT NULL" & _
'        "   THEN NE3.CASE_GROUP" & _
'        "   ELSE NE4.CASE_GROUP END AS CASE_GROUP"
'
'    lsCaseNumSQL = " , CASE    WHEN NE1.EMP_NUM IS NOT NULL" & _
'        "   THEN NE1.CASE_NUM" & _
'        "   WHEN NE2.EMP_NUM IS NOT NULL" & _
'        "   THEN NE2.CASE_NUM" & _
'        "   WHEN NE3.EMP_NUM IS NOT NULL" & _
'        "   THEN NE3.CASE_NUM" & _
'        "   ELSE NE4.CASE_NUM END AS CASE_NUM"
'
'    lsCaseYearSQL = " , CASE    WHEN NE1.EMP_NUM IS NOT NULL" & _
'        "   THEN NE1.CASE_YEAR" & _
'        "   WHEN NE2.EMP_NUM IS NOT NULL" & _
'        "   THEN NE2.CASE_YEAR" & _
'        "   WHEN NE3.EMP_NUM IS NOT NULL" & _
'        "   THEN NE3.CASE_YEAR" & _
'        "   ELSE NE4.CASE_YEAR END AS CASE_YEAR"
'
'    lsHoursSQL = " , COALESCE(NE1.HOURS, 0) + COALESCE(NE2.HOURS, 0)" & _
'                "  + COALESCE(NE3.HOURS, 0) + COALESCE(NE4.HOURS, 0)" & _
'                "  AS HOURS"
'
'    lsSQL = lsSQL & lsEmpNumSQL & _
'            lsChargeDateSQL & lsCaseGroupSQL & _
'            lsCaseNumSQL & lsCaseYearSQL & _
'            lsHoursSQL
'
'    lsSQL = lsSQL & _
'        " FROM  ( "
'
'    lsAdjustmentsSQL = "SELECT EMP_NUM, DATE_WORKED" & _
'        "   , CASE_GROUP, CASE_NUM" & _
'        "   , YEAR(DATE_WORKED) AS CASE_YEAR, SUM(HOURS) AS HOURS" & _
'        "   FROM    (SELECT EMP_NUM, DATE_WORKED" & _
'        "           , CASE_GROUP, CASE_NUM" & _
'        "           , YEAR(DATE_WORKED) AS CASE_YEAR" & _
'        "           , CASE WHEN ADJ_TYPE = 'A' " & _
'        "           THEN -1 * HOURS" & _
'        "           ELSE HOURS END AS HOURS" & _
'        "           FROM    METHDBA.TTR_ADJUSTMENTS A, METHDBA.TTR_EXCEPTION_CDS B" & _
'        "           WHERE   " & lsEmployeeSelection & _
'        "           AND     DATE_WORKED BETWEEN " & DateAsEnclosedString(vdtStartDate) & _
'        "           AND     " & DateAsEnclosedString(vdtEndDate) & _
'        "           AND     B.EXCEPTION_CD = A.EXCEPTION_CD"
'    lsNewAdjTableSQL = " UNION " & _
'        "           SELECT  EMP_NUM, CHARGE_DATE AS DATE_WORKED" & _
'        "                   , CASE_GROUP, CASE_NUM" & _
'        "                   , YEAR(CHARGE_DATE) AS CASE_YEAR" & _
'        "                   , HOURS" & _
'        "           FROM    METHDBA.TTR_ADJ_TIME A, METHDBA.TTR_EXCEPTION_CDS B" & _
'        "                   , METHDBA.TTR_ADJ_RQT C" & _
'        "           WHERE   " & lsEmployeeSelection & _
'        "           AND     A.ADJ_RQT_NUM = C.ADJ_RQT_NUM" & _
'        "           AND     CHARGE_DATE BETWEEN " & DateAsEnclosedString(vdtStartDate) & _
'        "           AND     " & DateAsEnclosedString(vdtEndDate) & _
'        "           AND     B.EXCEPTION_CD = A.EXCEPTION_CD" & _
'        "           ) NE1A" & _
'        " GROUP BY EMP_NUM, DATE_WORKED, CASE_GROUP, CASE_NUM"
'
'    lsSQL = lsSQL & _
'            lsAdjustmentsSQL & _
'            lsNewAdjTableSQL & _
'            ") NE1 " & _
'            " FULL OUTER JOIN ( "
'
'    lsActualTimeSQL = "SELECT  EMP_NUM, CHARGE_DATE" & _
'        "       , CASE_GROUP, CASE_NUM" & _
'        "       , YEAR(A.CHARGE_DATE) AS CASE_YEAR, SUM(HOURS) AS HOURS" & _
'        " FROM  METHDBA.TTR_ACTUAL_TIME A, METHDBA.TTR_EXCEPTION_CDS B" & _
'        " WHERE " & lsEmployeeSelection & _
'        " AND   CHARGE_DATE BETWEEN " & DateAsEnclosedString(ldtActualTimeStartDate) & _
'        " AND   " & DateAsEnclosedString(vdtEndDate) & _
'        " AND   CASE_NUM > 99" & _
'        " AND   B.EXCEPTION_CD = A.EXCEPTION_CD" & _
'        " GROUP BY EMP_NUM, CHARGE_DATE, CASE_GROUP, A.CASE_NUM"
'
'    lsSQL = lsSQL & lsActualTimeSQL & _
'        ") NE2 " & _
'        " ON  NE2.EMP_NUM = NE1.EMP_NUM" & _
'        " AND NE2.CHARGE_DATE = NE1.DATE_WORKED" & _
'        " AND NE2.CASE_GROUP = NE1.CASE_GROUP" & _
'        " AND NE2.CASE_NUM = NE1.CASE_NUM" & _
'        " FULL OUTER JOIN ("
'
'    lsActualTimeHSSQL = "SELECT  EMP_NUM, CHARGE_DATE" & _
'        "       , CASE_GROUP, CASE_NUM" & _
'        "       , YEAR(A.CHARGE_DATE) AS CASE_YEAR, SUM(HOURS) AS HOURS" & _
'        " FROM  METHDBA.TTR_ACTUAL_TIME_HS A, METHDBA.TTR_EXCEPTION_CDS B" & _
'        " WHERE " & lsEmployeeSelection & _
'        " AND   PART_VALUE_KEY " & lsPartValKeySel & _
'        " AND   CHARGE_DATE BETWEEN " & DateAsEnclosedString(vdtStartDate) & _
'        " AND   " & DateAsEnclosedString(ldtHistoryEndDate) & _
'        " AND   CASE_NUM > 99" & _
'        " AND   B.EXCEPTION_CD = A.EXCEPTION_CD" & _
'        " GROUP BY EMP_NUM, CHARGE_DATE, CASE_GROUP, A.CASE_NUM"
'
'    lsSQL = lsSQL & lsActualTimeHSSQL & _
'        ") NE3 " & _
'        " ON NE3.EMP_NUM = NE1.EMP_NUM" & _
'        " AND NE3.CHARGE_DATE = NE1.DATE_WORKED" & _
'        " AND NE3.CASE_GROUP = NE1.CASE_GROUP" & _
'        " AND NE3.CASE_NUM = NE1.CASE_NUM" & _
'        " FULL OUTER JOIN ("
'
'    lsEmpSchedExcSQL = "SELECT  EMP_NUM, CHARGE_DATE" & _
'        "       , CASE_GROUP, CASE_NUM" & _
'        "       , YEAR(A.CHARGE_DATE) AS CASE_YEAR, SUM(HOURS) AS HOURS" & _
'        " FROM  METHDBA.TTR_EMP_SCHED_EXC A, METHDBA.TTR_EXCEPTION_CDS B" & _
'        " WHERE " & lsEmployeeSelection & _
'        " AND   CHARGE_DATE BETWEEN " & DateAsEnclosedString(ldtActualTimeStartDate) & _
'        " AND   " & DateAsEnclosedString(vdtEndDate) & _
'        " AND   CASE_NUM > 99" & _
'        " AND   B.EXCEPTION_CD = A.EXCEPTION_CD" & _
'        " GROUP BY EMP_NUM, CHARGE_DATE, CASE_GROUP, A.CASE_NUM"
'
'    lsSQL = lsSQL & lsEmpSchedExcSQL & _
'        ") NE4 " & _
'        " ON  NE4.EMP_NUM = NE1.EMP_NUM" & _
'        " AND NE4.CHARGE_DATE = NE1.DATE_WORKED" & _
'        " AND NE4.CASE_GROUP = NE1.CASE_GROUP" & _
'        " AND NE4.CASE_NUM = NE1.CASE_NUM" & _
'        ") NE5 " & _
'    " WHERE CASE_NUM > 99" & _
'    " GROUP BY EMP_NUM, CASE_GROUP, CASE_NUM, CHARGE_DATE, CASE_YEAR" & _
'    " HAVING SUM(HOURS) > 0" & _
'    ") TRISRULES" & _
'    " GROUP BY EMP_NUM, CASE_GROUP, REFERENCE_NUM, CASE_YEAR"
'
'    GetCaseNumbersSQL = lsSQL
'
'End Function
''END PREM TRIS CONVERSION PHASE 3

Private Function DetermineActualTimeTables( _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date, _
                    ByRef r_datHistoryEndDate As Date, _
                    ByRef r_datActualTimeStartDate As Date _
                    ) As TrActualTimeTable
                                            
    Dim l_datCutOffChargeDate As Date

On Error GoTo eh
    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then
        DetermineActualTimeTables = TrTable_Actual
        r_datActualTimeStartDate = v_datStartDate
    ElseIf v_datEndDate <= l_datCutOffChargeDate Then
        DetermineActualTimeTables = TrTable_History
        r_datHistoryEndDate = v_datEndDate
    Else
        DetermineActualTimeTables = TrTable_ActualAndHistory
        r_datHistoryEndDate = l_datCutOffChargeDate
        r_datActualTimeStartDate = DateAdd("d", 1, l_datCutOffChargeDate)
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineActualTimeTables"
End Function


Public Function GetCaseGroupsAndNumbers( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date _
        ) As ADODB.Recordset
TrsRaiseError m_strCMODULENAME, "GetCaseGroupsAndNumbers", trErrObsoleteMethod, , , , True
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'
'    On Error GoTo eh
'
'    Set l_objCN = New ADODB.Connection
'    Set l_objRS = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
'
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open
'
'    l_strSQL = GetCaseNumbersSQL(v_aryEmployees, v_datStartDate, v_datEndDate)
'
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'
'    Set GetCaseGroupsAndNumbers = l_objRS
'
'    Set l_objRS = Nothing
'    Exit Function
'
'eh:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetCaseGroupsAndNumbers", , , , DescArgsFrom(l_strSQL), True
    
End Function
'PREM BEGIN TRIS CONVERSION PHASE 3 (PHASE 1 .NET)
'The following function retrieves Qualifications
' and Special Skills assigned to an employee or employees
Public Function GetQualsAndSkillsForEmployees(ByVal vsEmployeeWhereClause As String _
                    ) As ADODB.Recordset
On Error GoTo GetQualsAndSkillsForEmployees_EH
'SQL Statement Container
Dim lsSQL As String, ladcn As ADODB.Connection

    Set GetQualsAndSkillsForEmployees = Nothing

    lsSQL = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_EMP" & _
        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL  ) CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"
    lsSQL = lsSQL & _
        " INSERT INTO SESSION.GTT_EMP (EMP_NUM, EMP_NAME)" & _
        " SELECT A.EMP_NUM, A.EMP_NAME FROM METHDBA.TTR_EMPLOYEE A" & _
        " WHERE A.EMP_NUM IN (" & Trim(vsEmployeeWhereClause) & ")" & _
        " GROUP BY A.EMP_NUM, A.EMP_NAME;"

    Set ladcn = OpenConnection
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    lsSQL = "SELECT    A.EMP_NUM, B.EMP_NAME" & _
            "           , A.QUAL_CD, A.SPEC_SKILL_CD" & _
            "           , A.EFF_START_DATE, A.EFF_END_DATE" & _
            "           , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
            " FROM      SESSION.GTT_EMP B" & _
            "           INNER JOIN METHDBA.TTR_EMP_QUAL_SKILL A" & _
            "           ON A.EMP_NUM = B.EMP_NUM " & _
            " ORDER BY A.EMP_NUM, A.EFF_START_DATE;"
    'Call ExecuteSQL and retrieve information
'    Set GetQualsAndSkillsForEmployees = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetQualsAndSkillsForEmployees = NewExecuteSQLWithConnection(lsSQL _
                            , Nothing _
                            , ladcn _
                            , adOpenStatic _
                            , adLockReadOnly)

    lsSQL = "DROP TABLE SESSION.GTT_EMP;"
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

    Exit Function
GetQualsAndSkillsForEmployees_EH:
    If Not (ladcn Is Nothing) Then
        lsSQL = "DROP TABLE SESSION.GTT_EMP;"
        ExecuteOTHERSQL lsSQL _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetQualsAndSkillsForEmployees"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetQualsAndSkillsForEmployees" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetQualsAndSkillsForEmployees" _
                    , trErrdbGetEmpQAndS _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetQualsAndSkillsForEmployees2(ByVal vsEmployeeWhereClause As String _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999# _
                    ) As ADODB.Recordset
On Error GoTo GetQualsAndSkillsForEmployees2_EH
'SQL Statement Container
Dim lsSQL As String, ladcn As ADODB.Connection
Dim ladrsParms As ADODB.Recordset
Dim lsStartDate As String, lsEndDate As String

    Set GetQualsAndSkillsForEmployees2 = Nothing

    lsStartDate = GetValidDB2Date(vdtStartDate)
    lsEndDate = GetValidDB2Date(vdtEndDate)

    lsSQL = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_EMP" & _
        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL  ) CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"
    lsSQL = lsSQL & _
        " INSERT INTO SESSION.GTT_EMP (EMP_NUM, EMP_NAME)" & _
        " SELECT A.EMP_NUM, A.EMP_NAME FROM METHDBA.TTR_EMPLOYEE A" & _
        " WHERE A.EMP_NUM IN (" & Trim(vsEmployeeWhereClause) & ")" & _
        " GROUP BY A.EMP_NUM, A.EMP_NAME;"

    Set ladcn = OpenConnection
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    lsSQL = "SELECT    A.EMP_NUM, B.EMP_NAME" & _
        "           , A.QUAL_CD, A.SPEC_SKILL_CD" & _
        "           , A.EFF_START_DATE, A.EFF_END_DATE" & _
        "           , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
        " FROM      SESSION.GTT_EMP B" & _
        "           INNER JOIN METHDBA.TTR_EMP_QUAL_SKILL A" & _
        "           ON A.EMP_NUM = B.EMP_NUM" & _
        "           AND A.EFF_START_DATE BETWEEN ?" & _
        "           AND ?" & _
        "           AND A.EFF_END_DATE BETWEEN ?" & _
        "           AND ?" & _
        " ORDER BY A.EMP_NUM, A.EFF_START_DATE;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsStartDate)
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsEndDate)
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsStartDate)
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsEndDate)
        .Update
    End With

    Set GetQualsAndSkillsForEmployees2 = NewExecuteSQLWithConnection(lsSQL _
                            , ladrsParms _
                            , ladcn _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    lsSQL = "DROP TABLE SESSION.GTT_EMP;"
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

'        " WHERE     A.EMP_NUM IN (" & vsEmployeeWhereClause & ")" & _
'        " AND       A.EFF_START_DATE BETWEEN '" & Format(lsStartDate) & "'" & _
'        " AND       '" & Format(lsEndDate) & "'" & _
'        " AND       A.EFF_END_DATE BETWEEN '" & Format(lsStartDate) & "'" & _
'        " AND       '" & Format(lsEndDate) & "'" & _
'        " AND       A.EMP_NUM = B.EMP_NUM" & _
'        " ORDER BY A.EMP_NUM, A.EFF_START_DATE;"
    'Call ExecuteSQL and retrieve information
'    Set GetQualsAndSkillsForEmployees2 = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)

    Exit Function
GetQualsAndSkillsForEmployees2_EH:
    Set ladrsParms = Nothing

    If Not (ladcn Is Nothing) Then
        lsSQL = "DROP TABLE SESSION.GTT_EMP;"
        ExecuteOTHERSQL lsSQL _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetQualsAndSkillsForEmployees2"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetQualsAndSkillsForEmployees2" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetQualsAndSkillsForEmployees2" _
                    , trErrdbGetEmpQAndS _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetExpiringQAndSByLocation(ByVal viLocation As Integer _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetExpiringQAndSByLocation_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetExpiringQAndSByLocation = Nothing

    lsSQL = "SELECT    A.EMP_NUM, B.EMP_NAME" & _
        "       , A.EFF_END_DATE" & _
        "       , A.QUAL_CD, C.QUAL_DESC" & _
        "       , '' AS SPEC_SKILL_CD, '' AS SPEC_SKILL_DESC" & _
        " FROM  METHDBA.TTR_EMP_QUAL_SKILL A" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON B.EMP_NUM = A.EMP_NUM" & _
        "       AND B.LOC_CD = ?" & _
        "       AND B.EMP_STATUS <> 'T'" & _
        "       INNER JOIN METHDBA.TTR_QUALIFICATIONS C" & _
        "       ON C.QUAL_CD = A.QUAL_CD" & _
        " WHERE A.EFF_END_DATE BETWEEN ?" & _
        " AND   ?"

    lsSQL = lsSQL & _
        " UNION "

    lsSQL = lsSQL & _
        " SELECT    A.EMP_NUM, B.EMP_NAME" & _
        "       , A.EFF_END_DATE" & _
        "       , C.QUAL_CD, D.QUAL_DESC" & _
        "       , A.SPEC_SKILL_CD, C.SPEC_SKILL_DESC" & _
        " FROM  METHDBA.TTR_EMP_QUAL_SKILL A" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON B.EMP_NUM = A.EMP_NUM" & _
        "       AND B.LOC_CD = ?" & _
        "       AND B.EMP_STATUS <> 'T'" & _
        "       INNER JOIN METHDBA.TTR_SPEC_SKILLS C" & _
        "       ON C.SPEC_SKILL_CD = A.SPEC_SKILL_CD" & _
        "       INNER JOIN METHDBA.TTR_QUALIFICATIONS D" & _
        "       ON D.QUAL_CD = C.QUAL_CD" & _
        " WHERE A.EFF_END_DATE BETWEEN ?" & _
        " AND   ?" & _
        " ORDER BY EMP_NUM;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = viLocation
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = viLocation
        .Update

        .AddNew
        !PARM_ORDER = 5
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 6
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetExpiringQAndSByLocation = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetExpiringQAndSByLocation = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetExpiringQAndSByLocation_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetExpiringQAndSByLocation"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetExpiringQAndSByLocation" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetExpiringQAndSByLocation" _
                    , trErrdbGetEmpQAndS _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetExpiringQAndSByDept(ByVal vsDepartment As String _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetExpiringQAndSByDept_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetExpiringQAndSByDept = Nothing

    lsSQL = "SELECT    A.EMP_NUM, B.EMP_NAME" & _
        "       , A.EFF_END_DATE" & _
        "       , A.QUAL_CD, C.QUAL_DESC" & _
        "       , '' AS SPEC_SKILL_CD, '' AS SPEC_SKILL_DESC" & _
        " FROM  METHDBA.TTR_EMP_QUAL_SKILL A" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON B.EMP_NUM = A.EMP_NUM" & _
        "       AND B.DEPT_NUM = ?" & _
        "       AND B.EMP_STATUS <> 'T'" & _
        "       INNER JOIN METHDBA.TTR_QUALIFICATIONS C" & _
        "       ON C.QUAL_CD = A.QUAL_CD" & _
        " WHERE A.EFF_END_DATE BETWEEN ?" & _
        " AND   ?"

    lsSQL = lsSQL & _
        " UNION "

    lsSQL = lsSQL & _
        " SELECT    A.EMP_NUM, B.EMP_NAME" & _
        "       , A.EFF_END_DATE" & _
        "       , C.QUAL_CD, D.QUAL_DESC" & _
        "       , A.SPEC_SKILL_CD, C.SPEC_SKILL_DESC" & _
        " FROM  METHDBA.TTR_EMP_QUAL_SKILL A" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON B.EMP_NUM = A.EMP_NUM" & _
        "       AND B.DEPT_NUM = ?" & _
        "       AND B.EMP_STATUS <> 'T'" & _
        "       INNER JOIN METHDBA.TTR_SPEC_SKILLS C" & _
        "       ON C.SPEC_SKILL_CD = A.SPEC_SKILL_CD" & _
        "       INNER JOIN METHDBA.TTR_QUALIFICATIONS D" & _
        "       ON D.QUAL_CD = C.QUAL_CD" & _
        " WHERE A.EFF_END_DATE BETWEEN ?" & _
        " AND   ?" & _
        " ORDER BY EMP_NUM;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsDepartment
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsDepartment
        .Update

        .AddNew
        !PARM_ORDER = 5
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 6
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetExpiringQAndSByDept = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetExpiringQAndSByDept = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetExpiringQAndSByDept_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetExpiringQAndSByDept"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetExpiringQAndSByDept" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetExpiringQAndSByDept" _
                    , trErrdbGetEmpQAndS _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetExpiringQAndSByWG(ByVal vlWorkgroup As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetExpiringQAndSByWG_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetExpiringQAndSByWG = Nothing

    lsSQL = "SELECT    A.EMP_NUM, B.EMP_NAME" & _
        "       , A.EFF_END_DATE" & _
        "       , A.QUAL_CD, D.QUAL_DESC" & _
        "       , '' AS SPEC_SKILL_CD, '' AS SPEC_SKILL_DESC" & _
        " FROM  METHDBA.TTR_EMP_QUAL_SKILL A" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON B.EMP_NUM = A.EMP_NUM" & _
        "       AND B.EMP_STATUS <> 'T'" & _
        "       INNER JOIN METHDBA.TTR_EMP_WG C" & _
        "       ON C.EMP_NUM = B.EMP_NUM" & _
        "       AND C.WG_NUM = ?" & _
        "       AND CURRENT DATE BETWEEN C.EFF_START_DATE" & _
        "       AND C.EFF_END_DATE" & _
        "       INNER JOIN METHDBA.TTR_QUALIFICATIONS D" & _
        "       ON D.QUAL_CD = A.QUAL_CD" & _
        " WHERE A.EFF_END_DATE BETWEEN ?" & _
        " AND   ?"

    lsSQL = lsSQL & _
        " UNION "

    lsSQL = lsSQL & _
        " SELECT    A.EMP_NUM, B.EMP_NAME" & _
        "       , A.EFF_END_DATE" & _
        "       , D.QUAL_CD, E.QUAL_DESC" & _
        "       , A.SPEC_SKILL_CD, D.SPEC_SKILL_DESC" & _
        " FROM  METHDBA.TTR_EMP_QUAL_SKILL A" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
        "       ON B.EMP_NUM = A.EMP_NUM" & _
        "       AND B.EMP_STATUS <> 'T'" & _
        "       INNER JOIN METHDBA.TTR_EMP_WG C" & _
        "       ON C.EMP_NUM = B.EMP_NUM" & _
        "       AND C.WG_NUM = ?" & _
        "       AND CURRENT DATE BETWEEN C.EFF_START_DATE" & _
        "       AND C.EFF_END_DATE" & _
        "       INNER JOIN METHDBA.TTR_SPEC_SKILLS D" & _
        "       ON D.SPEC_SKILL_CD = A.SPEC_SKILL_CD" & _
        "       INNER JOIN METHDBA.TTR_QUALIFICATIONS E" & _
        "       ON E.QUAL_CD = D.QUAL_CD" & _
        " WHERE A.EFF_END_DATE BETWEEN ?" & _
        " AND   ?" & _
        " ORDER BY EMP_NUM"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlWorkgroup
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlWorkgroup
        .Update

        .AddNew
        !PARM_ORDER = 5
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 6
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetExpiringQAndSByWG = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetExpiringQAndSByWG = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetExpiringQAndSByWG_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetExpiringQAndSByWG"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetExpiringQAndSByWG" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetExpiringQAndSByWG" _
                    , trErrdbGetEmpQAndS _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'The following function retrieves seniority info. for an employee
Public Function GetEmployeeSeniority(ByVal vlEmployeeNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeSeniority_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetEmployeeSeniority = Nothing

    lsSQL = "SELECT A.EMP_NUM, D.EMP_NAME" & _
            "       , CASE (D.EMP_STATUS)" & _
            "       WHEN 'A' THEN 'ACTIVE'" & _
            "       WHEN 'L' THEN 'LTD'" & _
            "       WHEN 'I' THEN 'INACTIVE'" & _
            "       WHEN 'T' THEN 'TERMINATED'" & _
            "       ELSE ' ' END AS EMP_STATUS"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CORP_SNRTY_QTY))), 1, 3)) AS CORP_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CORP_SNRTY_QTY))), 4, 2)) CORP_DAYS" & _
            "       , CASE COALESCE(A.CORP_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.CORP_SNRTY_UP_DT, DATE('01/01/0100')) END AS CORP_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.CORP_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.CORP_SNRTY_EFF_DT, DATE('01/01/0100')) END AS CORP_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CMPY_SNRTY_QTY))), 1, 3)) AS CMPY_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CMPY_SNRTY_QTY))), 4, 2)) AS CMPY_DAYS" & _
            "       , CASE COALESCE(A.CMPY_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.CMPY_SNRTY_UP_DT, DATE('01/01/0100')) END AS CMPY_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.CMPY_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.CMPY_SNRTY_EFF_DT, DATE('01/01/0100')) END AS CMPY_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.PLANT_SNRTY_QTY))), 1, 3)) AS PLANT_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.PLANT_SNRTY_QTY))), 4, 2)) AS PLANT_DAYS" & _
            "       , CASE COALESCE(A.PLANT_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.PLANT_SNRTY_UP_DT, DATE('01/01/0100')) END AS PLANT_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.PLANT_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.PLANT_SNRTY_EFF_DT, DATE('01/01/0100')) END AS PLANT_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.UNION_SNRTY_QTY))), 1, 3)) AS UNION_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.UNION_SNRTY_QTY))), 4, 2)) AS UNION_DAYS" & _
            "       , CASE COALESCE(A.UNION_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.UNION_SNRTY_UP_DT, DATE('01/01/0100')) END AS UNION_SNRTY_UP_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.DEPT_SNRTY_QTY))), 1, 3)) AS DEPT_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.DEPT_SNRTY_QTY))), 4, 2)) AS DEPT_DAYS" & _
            "       , CASE COALESCE(A.DEPT_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.DEPT_SNRTY_UP_DT, DATE('01/01/0100')) END AS DEPT_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.DEPT_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.DEPT_SNRTY_EFF_DT, DATE('01/01/0100')) END AS DEPT_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.JOB_SNRTY_QTY))), 1, 3)) AS JOB_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.JOB_SNRTY_QTY))), 4, 2)) AS JOB_DAYS" & _
            "       , CASE COALESCE(A.JOB_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.JOB_SNRTY_UP_DT, DATE('01/01/0100')) END AS JOB_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.JOB_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.JOB_SNRTY_EFF_DT, DATE('01/01/0100')) END AS JOB_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(LOP_SNRTY_QTY))), 1, 3)) AS LOP_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(LOP_SNRTY_QTY))), 4, 2)) AS LOP_DAYS" & _
            "       , CASE COALESCE(A.LOP_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.LOP_SNRTY_UP_DT, DATE('01/01/0100')) END AS LOP_SNRTY_UP_DT"

    lsSQL = lsSQL & _
            "       , A.LTD_IND, A.MLTRY_LOA_IND" & _
            "       , A.SPCL_ATTN_FL, A.PPM_FL" & _
            "       , A.UNION_CD, A.DEPT_NUM" & _
            "       , A.JOB_NUM, A.SNRTY_PRTY_CD" & _
            "       , COALESCE(C.HR_LOC, '') AS HR_LOC" & _
            "       , COALESCE(C.IR_AREA_CD, '') AS IR_AREA_CD" & _
            "       , COALESCE(C.IR_DIV_CD, '') AS IR_DIV_CD" & _
            "       , COALESCE(C.LOP_CURR_LEV_IND, '') AS LOP_CURR_LEV_IND" & _
            "       , COALESCE(C.LOP_NEXT_LEV_IND, '') AS LOP_NEXT_LEV_IND" & _
            "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
            " FROM  METHDBA.TTR_SENIORITY A" & _
            "       LEFT OUTER JOIN METHDBA.TTR_CIPS_SNRTY_DT B" & _
            "       ON B.EMP_NUM = A.EMP_NUM" & _
            "       AND B.DEPT_NUM = A.DEPT_NUM" & _
            "       AND B.UNION_CD = A.UNION_CD" & _
            "       AND B.DEPT_NUM = A.DEPT_NUM" & _
            "       AND B.JOB_NUM = A.JOB_NUM" & _
            "       INNER JOIN METHDBA.TTR_EMPLOYEE D" & _
            "       ON D.EMP_NUM = A.EMP_NUM"
    lsSQL = lsSQL & _
            "       LEFT OUTER JOIN METHDBA.TTR_PROM_LINES C" & _
            "       ON C.DEPT_NUM = A.DEPT_NUM" & _
            "       AND C.JOB_NUM = A.JOB_NUM" & _
            "       AND C.HR_LOC = D.HR_LOC" & _
            " WHERE A.EMP_NUM = ?" & _
            " ORDER BY A.TIME_STAMP" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeSeniority = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeeSeniority = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeSeniority_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSeniority"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSeniority" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSeniority" _
                    , trErrdbGetEmpSnrty _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'This function retireves seniority for employees who belong to
'the passed workgroup.
Public Function GetEmployeeSeniorityForWG(ByVal vlWorkgroup As Long _
                    , ByVal vbActiveEmployeesOnly As Boolean _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeSeniorityForWG_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetEmployeeSeniorityForWG = Nothing

    lsSQL = "SELECT A.EMP_NUM, D.EMP_NAME" & _
            "       , CASE (D.EMP_STATUS)" & _
            "       WHEN 'A' THEN 'ACTIVE'" & _
            "       WHEN 'L' THEN 'LTD'" & _
            "       WHEN 'I' THEN 'INACTIVE'" & _
            "       WHEN 'T' THEN 'TERMINATED'" & _
            "       ELSE ' ' END AS EMP_STATUS"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CORP_SNRTY_QTY))), 1, 3)) AS CORP_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CORP_SNRTY_QTY))), 4, 2)) CORP_DAYS" & _
            "       , CASE COALESCE(A.CORP_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.CORP_SNRTY_UP_DT, DATE('01/01/0100')) END AS CORP_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.CORP_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.CORP_SNRTY_EFF_DT, DATE('01/01/0100')) END AS CORP_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CMPY_SNRTY_QTY))), 1, 3)) AS CMPY_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CMPY_SNRTY_QTY))), 4, 2)) AS CMPY_DAYS" & _
            "       , CASE COALESCE(A.CMPY_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.CMPY_SNRTY_UP_DT, DATE('01/01/0100')) END AS CMPY_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.CMPY_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.CMPY_SNRTY_EFF_DT, DATE('01/01/0100')) END AS CMPY_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.PLANT_SNRTY_QTY))), 1, 3)) AS PLANT_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.PLANT_SNRTY_QTY))), 4, 2)) AS PLANT_DAYS" & _
            "       , CASE COALESCE(A.PLANT_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.PLANT_SNRTY_UP_DT, DATE('01/01/0100')) END AS PLANT_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.PLANT_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.PLANT_SNRTY_EFF_DT, DATE('01/01/0100')) END AS PLANT_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.UNION_SNRTY_QTY))), 1, 3)) AS UNION_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.UNION_SNRTY_QTY))), 4, 2)) AS UNION_DAYS" & _
            "       , CASE COALESCE(A.UNION_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.UNION_SNRTY_UP_DT, DATE('01/01/0100')) END AS UNION_SNRTY_UP_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.DEPT_SNRTY_QTY))), 1, 3)) AS DEPT_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.DEPT_SNRTY_QTY))), 4, 2)) AS DEPT_DAYS" & _
            "       , CASE COALESCE(A.DEPT_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.DEPT_SNRTY_UP_DT, DATE('01/01/0100')) END AS DEPT_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.DEPT_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.DEPT_SNRTY_EFF_DT, DATE('01/01/0100')) END AS DEPT_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.JOB_SNRTY_QTY))), 1, 3)) AS JOB_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.JOB_SNRTY_QTY))), 4, 2)) AS JOB_DAYS" & _
            "       , CASE COALESCE(A.JOB_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.JOB_SNRTY_UP_DT, DATE('01/01/0100')) END AS JOB_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.JOB_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.JOB_SNRTY_EFF_DT, DATE('01/01/0100')) END AS JOB_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(LOP_SNRTY_QTY))), 1, 3)) AS LOP_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(LOP_SNRTY_QTY))), 4, 2)) AS LOP_DAYS" & _
            "       , CASE COALESCE(A.LOP_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.LOP_SNRTY_UP_DT, DATE('01/01/0100')) END AS LOP_SNRTY_UP_DT"

    lsSQL = lsSQL & _
            "       , A.LTD_IND, A.MLTRY_LOA_IND" & _
            "       , A.SPCL_ATTN_FL, A.PPM_FL" & _
            "       , A.UNION_CD, A.DEPT_NUM" & _
            "       , A.JOB_NUM, A.SNRTY_PRTY_CD" & _
            "       , COALESCE(C.HR_LOC, '') AS HR_LOC" & _
            "       , COALESCE(C.IR_AREA_CD, '') AS IR_AREA_CD" & _
            "       , COALESCE(C.IR_DIV_CD, '') AS IR_DIV_CD" & _
            "       , COALESCE(C.LOP_CURR_LEV_IND, '') AS LOP_CURR_LEV_IND" & _
            "       , COALESCE(C.LOP_NEXT_LEV_IND, '') AS LOP_NEXT_LEV_IND" & _
            "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP"

    lsSQL = lsSQL & _
            " FROM  METHDBA.TTR_SENIORITY A" & _
            "       INNER JOIN METHDBA.TTR_EMP_WG E" & _
            "       ON E.EMP_NUM = A.EMP_NUM" & _
            "       AND E.WG_NUM = ?" & _
            "       AND CURRENT DATE BETWEEN E.EFF_START_DATE" & _
            "       AND E.EFF_END_DATE" & _
            "       LEFT OUTER JOIN METHDBA.TTR_CIPS_SNRTY_DT B" & _
            "       ON B.EMP_NUM = A.EMP_NUM" & _
            "       AND B.DEPT_NUM = A.DEPT_NUM" & _
            "       AND B.UNION_CD = A.UNION_CD" & _
            "       AND B.DEPT_NUM = A.DEPT_NUM" & _
            "       AND B.JOB_NUM = A.JOB_NUM" & _
            "       INNER JOIN METHDBA.TTR_EMPLOYEE D" & _
            "       ON D.EMP_NUM = A.EMP_NUM"
    If (vbActiveEmployeesOnly) Then
        lsSQL = lsSQL & _
            " AND       (D.EMP_STATUS <> 'T'" & _
            " OR        D.EMP_STATUS = 'T'" & _
            " AND       D.STATUS_EFF_DATE > CURRENT DATE - 1 MONTHS)"
    End If

    lsSQL = lsSQL & _
            "       LEFT OUTER JOIN METHDBA.TTR_PROM_LINES C" & _
            "       ON C.DEPT_NUM = A.DEPT_NUM" & _
            "       AND C.JOB_NUM = A.JOB_NUM" & _
            "       AND C.HR_LOC = D.HR_LOC"

    lsSQL = lsSQL & " ORDER BY A.TIME_STAMP" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlWorkgroup
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    Set GetEmployeeSeniorityForWG = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeSeniorityForWG_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSeniorityForWG"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSeniorityForWG" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSeniorityForWG" _
                    , trErrdbGetEmpSnrty _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'This function retireves seniority for employees who belong to
'the passed department. Please note that information is not retrieved using
'department in seniority table
Public Function GetEmployeeSeniorityForDept(ByVal vsDepartment As String _
                    , ByVal vbActiveEmployeesOnly As Boolean _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeSeniorityForDept_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetEmployeeSeniorityForDept = Nothing

    lsSQL = "SELECT A.EMP_NUM, D.EMP_NAME" & _
            "       , CASE (D.EMP_STATUS)" & _
            "       WHEN 'A' THEN 'ACTIVE'" & _
            "       WHEN 'L' THEN 'LTD'" & _
            "       WHEN 'I' THEN 'INACTIVE'" & _
            "       WHEN 'T' THEN 'TERMINATED'" & _
            "       ELSE ' ' END AS EMP_STATUS"
    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CORP_SNRTY_QTY))), 1, 3)) AS CORP_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CORP_SNRTY_QTY))), 4, 2)) CORP_DAYS" & _
            "       , CASE COALESCE(A.CORP_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.CORP_SNRTY_UP_DT, DATE('01/01/0100')) END AS CORP_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.CORP_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.CORP_SNRTY_EFF_DT, DATE('01/01/0100')) END AS CORP_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CMPY_SNRTY_QTY))), 1, 3)) AS CMPY_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.CMPY_SNRTY_QTY))), 4, 2)) AS CMPY_DAYS" & _
            "       , CASE COALESCE(A.CMPY_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.CMPY_SNRTY_UP_DT, DATE('01/01/0100')) END AS CMPY_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.CMPY_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.CMPY_SNRTY_EFF_DT, DATE('01/01/0100')) END AS CMPY_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.PLANT_SNRTY_QTY))), 1, 3)) AS PLANT_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.PLANT_SNRTY_QTY))), 4, 2)) AS PLANT_DAYS" & _
            "       , CASE COALESCE(A.PLANT_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.PLANT_SNRTY_UP_DT, DATE('01/01/0100')) END AS PLANT_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.PLANT_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.PLANT_SNRTY_EFF_DT, DATE('01/01/0100')) END AS PLANT_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.UNION_SNRTY_QTY))), 1, 3)) AS UNION_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.UNION_SNRTY_QTY))), 4, 2)) AS UNION_DAYS" & _
            "       , CASE COALESCE(A.UNION_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.UNION_SNRTY_UP_DT, DATE('01/01/0100')) END AS UNION_SNRTY_UP_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.DEPT_SNRTY_QTY))), 1, 3)) AS DEPT_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.DEPT_SNRTY_QTY))), 4, 2)) AS DEPT_DAYS" & _
            "       , CASE COALESCE(A.DEPT_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.DEPT_SNRTY_UP_DT, DATE('01/01/0100')) END AS DEPT_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.DEPT_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.DEPT_SNRTY_EFF_DT, DATE('01/01/0100')) END AS DEPT_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.JOB_SNRTY_QTY))), 1, 3)) AS JOB_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.JOB_SNRTY_QTY))), 4, 2)) AS JOB_DAYS" & _
            "       , CASE COALESCE(A.JOB_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.JOB_SNRTY_UP_DT, DATE('01/01/0100')) END AS JOB_SNRTY_UP_DT" & _
            "       , CASE COALESCE(B.JOB_SNRTY_EFF_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.JOB_SNRTY_EFF_DT, DATE('01/01/0100')) END AS JOB_SNRTY_EFF_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(LOP_SNRTY_QTY))), 1, 3)) AS LOP_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(LOP_SNRTY_QTY))), 4, 2)) AS LOP_DAYS" & _
            "       , CASE COALESCE(A.LOP_SNRTY_UP_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.LOP_SNRTY_UP_DT, DATE('01/01/0100')) END AS LOP_SNRTY_UP_DT"

    lsSQL = lsSQL & _
            "       , A.LTD_IND, A.MLTRY_LOA_IND" & _
            "       , A.SPCL_ATTN_FL, A.PPM_FL" & _
            "       , A.UNION_CD, A.DEPT_NUM" & _
            "       , A.JOB_NUM, A.SNRTY_PRTY_CD" & _
            "       , COALESCE(C.HR_LOC, '') AS HR_LOC" & _
            "       , COALESCE(C.IR_AREA_CD, '') AS IR_AREA_CD" & _
            "       , COALESCE(C.IR_DIV_CD, '') AS IR_DIV_CD" & _
            "       , COALESCE(C.LOP_CURR_LEV_IND, '') AS LOP_CURR_LEV_IND" & _
            "       , COALESCE(C.LOP_NEXT_LEV_IND, '') AS LOP_NEXT_LEV_IND" & _
            "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP"

    lsSQL = lsSQL & _
            " FROM  METHDBA.TTR_SENIORITY A" & _
            "       LEFT OUTER JOIN METHDBA.TTR_CIPS_SNRTY_DT B" & _
            "       ON B.EMP_NUM = A.EMP_NUM" & _
            "       AND B.DEPT_NUM = A.DEPT_NUM" & _
            "       AND B.UNION_CD = A.UNION_CD" & _
            "       AND B.DEPT_NUM = A.DEPT_NUM" & _
            "       AND B.JOB_NUM = A.JOB_NUM" & _
            "       INNER JOIN METHDBA.TTR_EMPLOYEE D" & _
            "       ON D.EMP_NUM = A.EMP_NUM" & _
            "       AND D.DEPT_NUM = ?"
    If (vbActiveEmployeesOnly) Then
        lsSQL = lsSQL & _
            " AND   (D.EMP_STATUS <> 'T'" & _
            " OR    D.EMP_STATUS = 'T'" & _
            " AND   D.STATUS_EFF_DATE > CURRENT DATE - 1 MONTHS)"
    End If

    lsSQL = lsSQL & _
            "       LEFT OUTER JOIN METHDBA.TTR_PROM_LINES C" & _
            "       ON C.DEPT_NUM = A.DEPT_NUM" & _
            "       AND C.JOB_NUM = A.JOB_NUM" & _
            "       AND C.HR_LOC = D.HR_LOC"

    lsSQL = lsSQL & " ORDER BY A.TIME_STAMP" & _
            " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsDepartment
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    Set GetEmployeeSeniorityForDept = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeSeniorityForDept_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSeniorityForDept"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSeniorityForDept" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSeniorityForDept" _
                    , trErrdbGetEmpSnrty _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'Oh! IBM Vs. Microsoft
Private Function GetValidDB2Date(ByVal vdtADODBDate As Date _
                    ) As String
On Error GoTo GetValidDB2Date_EH

    If (vdtADODBDate = #1/1/100#) Then
        GetValidDB2Date = "01/01/0001"
    Else
        GetValidDB2Date = Format(vdtADODBDate, "mm/dd/yyyy")
    End If

    Exit Function
GetValidDB2Date_EH:
    TrsRaiseError m_strCMODULENAME _
        , "GetValidDB2Date" _
        , trErrdbSaveEmpEmpSnrty _
        , , , _
        , True

End Function
'The following function retrieves frozen seniority info. for an employee
Public Function GetEmployeeFrozenSeniority(ByVal vlEmployeeNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeFrozenSeniority_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetEmployeeFrozenSeniority = Nothing

    'Note the where clause of the following SQL statement
    'No Corporate Frozen Seniority Information is retrieved
    'for maintenance
    lsSQL = "SELECT A.EMP_NUM, B.EMP_NAME" & _
            "       , CASE (B.EMP_STATUS)" & _
            "       WHEN 'A' THEN 'ACTIVE'" & _
            "       WHEN 'L' THEN 'LTD'" & _
            "       WHEN 'I' THEN 'INACTIVE'" & _
            "       WHEN 'T' THEN 'TERMINATED'" & _
            "       ELSE '' END AS EMP_STATUS"

    lsSQL = lsSQL & _
            "       , CASE(A.SNRTY_REC_TYPE_CD)" & _
            "       WHEN 'CMP' THEN 'COMPANY'" & _
            "       WHEN 'CRP' THEN 'CORPORATE'" & _
            "       WHEN 'DPT' THEN 'DEPARTMENT'" & _
            "       WHEN 'JOB' THEN 'JOB'" & _
            "       WHEN 'LOP' THEN 'LOP'" & _
            "       WHEN 'PLT' THEN 'PLANT'" & _
            "       WHEN 'UNI' THEN 'UNION'" & _
            "       ELSE '' END AS SNRTY_REC_TYPE_CD" & _
            "       , A.CMPY_CD , A.UNION_CD" & _
            "       , A.DEPT_NUM, A.JOB_NUM" & _
            "       , A.IR_AREA_CD, A.IR_DIV_CD" & _
            "       , CASE COALESCE(A.SNRTY_FRZN_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.SNRTY_FRZN_DT, DATE('01/01/0100')) END AS SNRTY_FRZN_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.FRZN_SNRTY_REC_QTY))), 1, 3)) AS FRZN_SNRTY_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.FRZN_SNRTY_REC_QTY))), 4, 2)) FRZN_SNRTY_DAYS" & _
            "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
            " FROM  METHDBA.TTR_FRZN_SNRTY A" & _
            "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
            "       ON B.EMP_NUM = A.EMP_NUM" & _
            " WHERE A.EMP_NUM = ?" & _
            " AND   A.SNRTY_REC_TYPE_CD <> 'CRP'" & _
            " ORDER BY  A.TIME_STAMP"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeFrozenSeniority = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeeFrozenSeniority = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeFrozenSeniority_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFrozenSeniority"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFrozenSeniority" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFrozenSeniority" _
                    , trErrdbGetEmpFrznSnrty _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'The following function retrieves frozen seniority info. for
' employees belonging to the passed department
Public Function GetEmployeeFrznSnrtyForDept(ByVal vsDepartment As String _
                    , ByVal vbActiveEmployeesOnly As Boolean _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeFrznSnrtyForDept_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetEmployeeFrznSnrtyForDept = Nothing

    'Note the where clause of the following SQL statement
    'No Corporate Frozen Seniority Information is retrieved
    'for maintenance
    lsSQL = "SELECT A.EMP_NUM, B.EMP_NAME" & _
            "       , CASE (B.EMP_STATUS)" & _
            "       WHEN 'A' THEN 'ACTIVE'" & _
            "       WHEN 'L' THEN 'LTD'" & _
            "       WHEN 'I' THEN 'INACTIVE'" & _
            "       WHEN 'T' THEN 'TERMINATED'" & _
            "       ELSE '' END AS EMP_STATUS"

    lsSQL = lsSQL & _
            "       , CASE(A.SNRTY_REC_TYPE_CD)" & _
            "       WHEN 'CMP' THEN 'COMPANY'" & _
            "       WHEN 'CRP' THEN 'CORPORATE'" & _
            "       WHEN 'DPT' THEN 'DEPARTMENT'" & _
            "       WHEN 'JOB' THEN 'JOB'" & _
            "       WHEN 'LOP' THEN 'LOP'" & _
            "       WHEN 'PLT' THEN 'PLANT'" & _
            "       WHEN 'UNI' THEN 'UNION'" & _
            "       ELSE '' END AS SNRTY_REC_TYPE_CD" & _
            "       , A.CMPY_CD , A.UNION_CD" & _
            "       , A.DEPT_NUM, A.JOB_NUM" & _
            "       , A.IR_AREA_CD, A.IR_DIV_CD" & _
            "       , CASE COALESCE(A.SNRTY_FRZN_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.SNRTY_FRZN_DT, DATE('01/01/0100')) END AS SNRTY_FRZN_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.FRZN_SNRTY_REC_QTY))), 1, 3)) AS FRZN_SNRTY_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.FRZN_SNRTY_REC_QTY))), 4, 2)) FRZN_SNRTY_DAYS" & _
            "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
            " FROM  METHDBA.TTR_FRZN_SNRTY A" & _
            "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
            "       ON B.EMP_NUM = A.EMP_NUM" & _
            "       AND B.DEPT_NUM = ?" & _
            "       AND CURRENT DATE >= B.DEPT_EFF_DATE" & _
            " WHERE A.SNRTY_REC_TYPE_CD <> 'CRP'"

    If (vbActiveEmployeesOnly) Then
        lsSQL = lsSQL & _
            " AND       (B.EMP_STATUS <> 'T'" & _
            " OR        B.EMP_STATUS = 'T'" & _
            " AND       B.STATUS_EFF_DATE > CURRENT DATE - 1 MONTHS)"
    End If
    
    lsSQL = lsSQL & " ORDER BY  A.TIME_STAMP"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = vsDepartment
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    Set GetEmployeeFrznSnrtyForDept = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeFrznSnrtyForDept_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFrznSnrtyForDept"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFrznSnrtyForDept" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFrznSnrtyForDept" _
                    , trErrdbGetEmpFrznSnrty _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'The following function retrieves frozen seniority info. for
' employees belonging to the passed department
Public Function GetEmployeeFrznSnrtyForWG(ByVal vlWorkgroup As Long _
                    , ByVal vbActiveEmployeesOnly As Boolean _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeFrznSnrtyForWG_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset
    
    Set GetEmployeeFrznSnrtyForWG = Nothing

    'Note the where clause of the following SQL statement
    'No Corporate Frozen Seniority Information is retrieved
    'for maintenance
    lsSQL = "SELECT A.EMP_NUM, B.EMP_NAME" & _
            "       , CASE (B.EMP_STATUS)" & _
            "       WHEN 'A' THEN 'ACTIVE'" & _
            "       WHEN 'L' THEN 'LTD'" & _
            "       WHEN 'I' THEN 'INACTIVE'" & _
            "       WHEN 'T' THEN 'TERMINATED'" & _
            "       ELSE '' END AS EMP_STATUS"

    lsSQL = lsSQL & _
            "       , CASE(A.SNRTY_REC_TYPE_CD)" & _
            "       WHEN 'CMP' THEN 'COMPANY'" & _
            "       WHEN 'CRP' THEN 'CORPORATE'" & _
            "       WHEN 'DPT' THEN 'DEPARTMENT'" & _
            "       WHEN 'JOB' THEN 'JOB'" & _
            "       WHEN 'LOP' THEN 'LOP'" & _
            "       WHEN 'PLT' THEN 'PLANT'" & _
            "       WHEN 'UNI' THEN 'UNION'" & _
            "       ELSE '' END AS SNRTY_REC_TYPE_CD" & _
            "       , A.CMPY_CD , A.UNION_CD" & _
            "       , A.DEPT_NUM, A.JOB_NUM" & _
            "       , A.IR_AREA_CD, A.IR_DIV_CD" & _
            "       , CASE COALESCE(A.SNRTY_FRZN_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.SNRTY_FRZN_DT, DATE('01/01/0100')) END AS SNRTY_FRZN_DT"

    lsSQL = lsSQL & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.FRZN_SNRTY_REC_QTY))), 1, 3)) AS FRZN_SNRTY_MONTHS" & _
            "       , INTEGER(SUBSTR(LTRIM(RTRIM(CHAR(A.FRZN_SNRTY_REC_QTY))), 4, 2)) FRZN_SNRTY_DAYS" & _
            "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP" & _
            " FROM  METHDBA.TTR_EMP_WG C" & _
            "       LEFT OUTER JOIN METHDBA.TTR_FRZN_SNRTY A" & _
            "       ON C.EMP_NUM = A.EMP_NUM" & _
            "       AND C.WG_NUM = ?" & _
            "       AND CURRENT DATE BETWEEN C.EFF_START_DATE" & _
            "       AND C.EFF_END_DATE" & _
            "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
            "       ON B.EMP_NUM = A.EMP_NUM" & _
            " WHERE A.SNRTY_REC_TYPE_CD <> 'CRP'"

    If (vbActiveEmployeesOnly) Then
        lsSQL = lsSQL & _
            " AND       (B.EMP_STATUS <> 'T'" & _
            " OR        B.EMP_STATUS = 'T'" & _
            " AND       B.STATUS_EFF_DATE > CURRENT DATE - 1 MONTHS)"
    End If

    lsSQL = lsSQL & " ORDER BY  A.TIME_STAMP;"

   Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlWorkgroup
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    Set GetEmployeeFrznSnrtyForWG = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeFrznSnrtyForWG_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFrznSnrtyForWG"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFrznSnrtyForWG" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFrznSnrtyForWG" _
                    , trErrdbGetEmpFrznSnrty _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'The following function retrieves Employee Overtime Canvass Job Overrides
Public Function GetEmployeeCanvassJob(ByVal vsEmployeeWhereClause As String _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeCanvassJob_EH
'SQL Statement Container
Dim lsSQL As String, ladcn As ADODB.Connection

    Set GetEmployeeCanvassJob = Nothing

    lsSQL = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_EMP" & _
        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL  ) CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"
    lsSQL = lsSQL & _
        " INSERT INTO SESSION.GTT_EMP (EMP_NUM, EMP_NAME)" & _
        " SELECT A.EMP_NUM, A.EMP_NAME FROM METHDBA.TTR_EMPLOYEE A" & _
        " WHERE A.EMP_NUM IN (" & Trim(vsEmployeeWhereClause) & ")" & _
        " GROUP BY A.EMP_NUM, A.EMP_NAME;"

    Set ladcn = OpenConnection
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    lsSQL = "SELECT    A.EMP_NUM, B.EMP_NAME" & _
        "           , A.JOB_NUM, C.JOB_DESC" & _
        "           , A.CRAFT_NUM, D.CRAFT_DESC" & _
        "           , A.EFF_START_DATE, A.EFF_END_DATE" & _
        "           , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
        "           , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
        " FROM      SESSION.GTT_EMP B" & _
        "           INNER JOIN METHDBA.TTR_OT_EMP_OVRD A" & _
        "           ON A.EMP_NUM = B.EMP_NUM" & _
        "           LEFT OUTER JOIN METHDBA.TTR_JOBS C" & _
        "           ON C.JOB_NUM = A.JOB_NUM" & _
        "           LEFT OUTER JOIN METHDBA.TTR_CRAFTS D" & _
        "           ON D.CRAFT_NUM = A.CRAFT_NUM" & _
        " ORDER BY  EMP_NUM, EFF_START_DATE;"

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeCanvassJob = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeeCanvassJob = NewExecuteSQLWithConnection(lsSQL _
                            , Nothing _
                            , ladcn _
                            , adOpenStatic _
                            , adLockReadOnly)

    lsSQL = "DROP TABLE SESSION.GTT_EMP;"
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

    Exit Function
GetEmployeeCanvassJob_EH:
    If Not (ladcn Is Nothing) Then
        lsSQL = "DROP TABLE SESSION.GTT_EMP;"
        ExecuteOTHERSQL lsSQL _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCanvassJob"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCanvassJob" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCanvassJob" _
                    , trErrdbGetEmpCanvassJob _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeeCanvassJob2(ByVal vsEmployeeWhereClause As String _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999# _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeCanvassJob2_EH
'SQL Statement Container
Dim lsSQL As String, ladcn As ADODB.Connection
Dim ladrsParms As ADODB.Recordset
Dim lsStartDate As String, lsEndDate As String

    lsStartDate = GetValidDB2Date(vdtStartDate)
    lsEndDate = GetValidDB2Date(vdtEndDate)

    lsSQL = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_EMP" & _
        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL  ) CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"
    lsSQL = lsSQL & _
        " INSERT INTO SESSION.GTT_EMP (EMP_NUM, EMP_NAME)" & _
        " SELECT A.EMP_NUM, A.EMP_NAME FROM METHDBA.TTR_EMPLOYEE A" & _
        " WHERE A.EMP_NUM IN (" & Trim(vsEmployeeWhereClause) & ")" & _
        " GROUP BY A.EMP_NUM, A.EMP_NAME;"

    Set ladcn = OpenConnection
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    lsSQL = "SELECT    A.EMP_NUM, B.EMP_NAME " & _
        "           , A.JOB_NUM, C.JOB_DESC" & _
        "           , A.CRAFT_NUM, D.CRAFT_DESC" & _
        "           , A.EFF_START_DATE, A.EFF_END_DATE " & _
        "           , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP " & _
        "           , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP " & _
        " FROM      SESSION.GTT_EMP B" & _
        "           INNER JOIN METHDBA.TTR_OT_EMP_OVRD A" & _
        "           ON A.EMP_NUM = B.EMP_NUM" & _
        "           AND A.EFF_START_DATE BETWEEN ?" & _
        "           AND ?" & _
        "           AND A.EFF_END_DATE BETWEEN ?" & _
        "           AND ?" & _
        "           LEFT OUTER JOIN METHDBA.TTR_JOBS C" & _
        "           ON C.JOB_NUM = A.JOB_NUM" & _
        "           LEFT OUTER JOIN METHDBA.TTR_CRAFTS D" & _
        "           ON D.CRAFT_NUM = A.CRAFT_NUM" & _
        " ORDER BY  A.EMP_NUM, A.EFF_START_DATE;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsStartDate)
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsEndDate)
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsStartDate)
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsEndDate)
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeCanvassJob2 = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeeCanvassJob2 = NewExecuteSQLWithConnection(lsSQL _
                            , ladrsParms _
                            , ladcn _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    lsSQL = "DROP TABLE SESSION.GTT_EMP;"
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

    Exit Function
GetEmployeeCanvassJob2_EH:
    Set ladrsParms = Nothing

    If Not (ladcn Is Nothing) Then
        lsSQL = "DROP TABLE SESSION.GTT_EMP;"
        ExecuteOTHERSQL lsSQL _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCanvassJob2"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCanvassJob2" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCanvassJob2" _
                    , trErrdbGetEmpCanvassJob _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'The following function retrieves Employee Ineligibilities
Public Function GetEmployeeIneligibilities(ByVal vsEmployeeWhereClause As String _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeIneligibilities_EH
'SQL Statement Container
Dim lsSQL As String, ladcn As ADODB.Connection

    lsSQL = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_EMP" & _
        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL  ) CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"
    lsSQL = lsSQL & _
        " INSERT INTO SESSION.GTT_EMP (EMP_NUM, EMP_NAME)" & _
        " SELECT A.EMP_NUM, A.EMP_NAME FROM METHDBA.TTR_EMPLOYEE A" & _
        " WHERE A.EMP_NUM IN (" & Trim(vsEmployeeWhereClause) & ")" & _
        " GROUP BY A.EMP_NUM, A.EMP_NAME;"

    Set ladcn = OpenConnection
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    lsSQL = "SELECT    A.EMP_NUM, B.EMP_NAME, A.EXCEPTION_CD " & _
        "           , A.EFF_START_DATE, A.EFF_END_DATE " & _
        "           , A.EFF_START_TIME, A.EFF_END_TIME " & _
        "           , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP " & _
        " FROM      SESSION.GTT_EMP B" & _
        "           INNER JOIN METHDBA.TTR_INELIGIBILITY A" & _
        "           ON A.EMP_NUM = B.EMP_NUM" & _
        " ORDER BY  A.EMP_NUM, A.EFF_START_DATE, A.EFF_START_TIME;"

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeIneligibilities = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeeIneligibilities = NewExecuteSQLWithConnection(lsSQL _
                            , Nothing _
                            , ladcn _
                            , adOpenStatic _
                            , adLockReadOnly)

    lsSQL = "DROP TABLE SESSION.GTT_EMP;"
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

    Exit Function
GetEmployeeIneligibilities_EH:
    If Not (ladcn Is Nothing) Then
        lsSQL = "DROP TABLE SESSION.GTT_EMP;"
        ExecuteOTHERSQL lsSQL _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeIneligibilities"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeIneligibilities" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeIneligibilities" _
                    , trErrdbGetEmpIneligibility _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeeIneligibilities2( _
                    ByVal vsEmployeeWhereClause As String _
                    , Optional ByVal vdtStartDate As Date = #1/1/100# _
                    , Optional ByVal vdtEndDate As Date = #12/31/9999# _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeIneligibilities2_EH
'SQL Statement Container
Dim lsSQL As String, ladcn As ADODB.Connection
Dim ladrsParms As ADODB.Recordset
Dim lsStartDate As String, lsEndDate As String

    lsStartDate = GetValidDB2Date(vdtStartDate)
    lsEndDate = GetValidDB2Date(vdtEndDate)

    lsSQL = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_EMP" & _
        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, EMP_NAME CHAR(30) NOT NULL  ) CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"
    lsSQL = lsSQL & _
        " INSERT INTO SESSION.GTT_EMP (EMP_NUM, EMP_NAME)" & _
        " SELECT A.EMP_NUM, A.EMP_NAME FROM METHDBA.TTR_EMPLOYEE A" & _
        " WHERE A.EMP_NUM IN (" & Trim(vsEmployeeWhereClause) & ")" & _
        " GROUP BY A.EMP_NUM, A.EMP_NAME;"

    Set ladcn = OpenConnection
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    lsSQL = "SELECT    A.EMP_NUM, B.EMP_NAME, A.EXCEPTION_CD " & _
        "           , A.EFF_START_DATE, A.EFF_END_DATE " & _
        "           , A.EFF_START_TIME, A.EFF_END_TIME " & _
        "           , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP " & _
        " FROM      SESSION.GTT_EMP B" & _
        "           INNER JOIN METHDBA.TTR_INELIGIBILITY A" & _
        "           ON A.EMP_NUM = B.EMP_NUM" & _
        "           AND A.EFF_START_DATE BETWEEN ?" & _
        "           AND ?" & _
        "           AND A.EFF_END_DATE BETWEEN ?" & _
        "           AND ?" & _
        " ORDER BY  A.EMP_NUM, A.EFF_START_DATE, A.EFF_START_TIME;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsStartDate)
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsEndDate)
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsStartDate)
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = CDate(lsEndDate)
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeIneligibilities2 = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeeIneligibilities2 = NewExecuteSQLWithConnection(lsSQL _
                            , ladrsParms _
                            , ladcn _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    lsSQL = "DROP TABLE SESSION.GTT_EMP;"
    ExecuteOTHERSQL lsSQL _
        , ladcn _
        , adCmdText

    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

    Exit Function
GetEmployeeIneligibilities2_EH:
    Set ladrsParms = Nothing

    If Not (ladcn Is Nothing) Then
        lsSQL = "DROP TABLE SESSION.GTT_EMP;"
        ExecuteOTHERSQL lsSQL _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeIneligibilities2"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeIneligibilities2" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeIneligibilities2" _
                    , trErrdbGetEmpIneligibility _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'The following function retrieves Employee Volunteer Information
Public Function GetVolunteerInformation() As ADODB.Recordset
On Error GoTo GetVolunteerInformation_EH
'SQL Statement Container
Dim lsSQL As String

    lsSQL = " SELECT  A.EMP_NUM, A.EMP_NAME " & _
        "           , 'YES' AS VLNTR_STATUS , 'YES' AS CURR_VLNTR_STATUS " & _
        "           , '' AS STUS_EFF_DT , '' AS TIME_STAMP " & _
        " FROM      METHDBA.TTR_EMPLOYEE A, METHDBA.TTR_OT_BANKS B " & _
        " WHERE     A.EMP_NUM = B.EMP_NUM " & _
        " AND       B.OT_BANK_CD = 2 " & _
        " AND       B.TOT_OT_HRS > 0 " & _
        " AND       A.LOC_CD = 121 " & _
        " UNION " & _
        " SELECT    A.EMP_NUM, A.EMP_NAME " & _
        "           , 'NO' AS VLNTR_STATUS , 'NO' AS CURR_VLNTR_STATUS " & _
        "           , '' AS STUS_EFF_DT , '' AS TIME_STAMP " & _
        " FROM      METHDBA.TTR_EMPLOYEE A, METHDBA.TTR_OT_BANKS B " & _
        " WHERE     A.EMP_NUM = B.EMP_NUM " & _
        " AND       B.OT_BANK_CD = 2 " & _
        " AND       B.TOT_OT_HRS <= 0 " & _
        " AND       A.LOC_CD = 121 " & _
        " ORDER BY  EMP_NUM;"

    'Call ExecuteSQL and retrieve information
    Set GetVolunteerInformation = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockBatchOptimistic)

    Exit Function
GetVolunteerInformation_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetVolunteerInformation"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetVolunteerInformation" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetVolunteerInformation" _
                    , trErrdbGetVolunteer _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'The following function retrieves Employee Volunteer Action Information
'Please note that this used only by the calling bus. The TRIS Client
'does not require this information, but this information is used to prepare
'information passed to the client
Public Function GetVolunteerAction() As ADODB.Recordset
On Error GoTo GetVolunteerAction_EH
'SQL Statement Container
Dim lsSQL As String

    lsSQL = " SELECT  EMP_NUM, STUS_EFF_DT " & _
        "           , CASE(ACTN_CD) " & _
        "           WHEN 'A' THEN 'YES' " & _
        "           WHEN 'D' THEN 'NO' " & _
        "           ELSE '   ' " & _
        "           END AS ACTN_CD " & _
        "           , CHAR(TIME_STAMP) AS TIME_STAMP " & _
        " FROM      METHDBA.TTR_VOLUNTEER " & _
        " ORDER BY  EMP_NUM, STUS_EFF_DT;"

    'Call ExecuteSQL and retrieve information
    Set GetVolunteerAction = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    Exit Function
GetVolunteerAction_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetVolunteerAction"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetVolunteerAction" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetVolunteerAction" _
                    , trErrdbGetVolunteer _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function AreValidEmployees(ByVal vsEmployeeWhereClauseValues As String _
                    ) As ADODB.Recordset
On Error GoTo AreValidEmployees_EH
'SQL Statement Container
Dim lsSQL As String

    lsSQL = " SELECT  EMP_NUM " & _
        " FROM      METHDBA.TTR_EMPLOYEE " & _
        " WHERE     EMP_NUM IN (" & vsEmployeeWhereClauseValues & ")" & _
        " FOR FETCH ONLY WITH UR;"

    'Call ExecuteSQL and retrieve information
    Set AreValidEmployees = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    Exit Function
AreValidEmployees_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "AreValidEmployees"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "AreValidEmployees" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "AreValidEmployees" _
                    , , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'PROJECT BANKS MAINTENANCE
Public Function GetEmployeesInProject(ByVal vlProjectID As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeesInProject_EH
'SQL Statement Container
Dim lsSQL As String

    lsSQL = " SELECT  B.EMP_NUM, C.EMP_NAME, C.UNION_CD " & _
                    "       , D.CRAFT_NUM, C.LOC_CD " & _
                    " FROM  METHDBA.TTR_WORK_GROUP A, METHDBA.TTR_EMP_WG B " & _
                    "       , METHDBA.TTR_EMPLOYEE C, METHDBA.TTR_JOBS D " & _
                    " WHERE A.PROJECT_ID = " & CStr(vlProjectID) & _
                    " AND   A.WG_NUM = B.WG_NUM " & _
                    " AND   B.EFF_START_DATE <= '" & Format("mm/dd/yyyy", vdtStartDate) & "'" & _
                    " AND   B.EFF_END_DATE >= '" & Format("mm/dd/yyyy", vdtEndDate) & "'" & _
                    " AND   B.EMP_NUM = C.EMP_NUM " & _
                    " AND   C.EXEC_CONTRACT_IND <> 'E' " & _
                    " AND   C.JOB_NUM = D.JOB_NUM " & _
                    " ORDER BY A.EMP_NAME "

    'Call ExecuteSQL and retrieve information
    Set GetEmployeesInProject = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    Exit Function
GetEmployeesInProject_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesInProject"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesInProject" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesInProject" _
                    , , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeesForProjectByWorkgroup(ByVal vlProjectID As Long _
                    , ByVal vlWorkgroupNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset

On Error GoTo GetEmployeesForProjectByWorkgroup_EH
'SQL Statement Container
Dim lsSQL As String

    lsSQL = " SELECT  A.EMP_NUM, B.EMP_NAME, B.UNION_CD " & _
                    "       , C.CRAFT_NUM, B.LOC_CD " & _
                    " FROM  METHDBA.TTR_EMP_WG A, METHDBA.TTR_EMPLOYEE B " & _
                    "       , METHDBA.TTR_JOBS C, METHDBA.TTR_WORK_GROUP D " & _
                    " WHERE A.WG_NUM = " & CStr(vlWorkgroupNumber) & _
                    " AND   A.EFF_START_DATE <= '" & Format("mm/dd/yyyy", vdtStartDate) & "'" & _
                    " AND   A.EFF_END_DATE >= '" & Format("mm/dd/yyyy", vdtEndDate) & "'" & _
                    " AND   A.EMP_NUM = B.EMP_NUM " & _
                    " AND   B.EXEC_CONTRACT_IND <> 'E' " & _
                    " AND   D.PROJECT_ID = " & CStr(vlProjectID) & _
                    " AND   A.WG_NUM = D.WG_NUM " & _
                    " AND   B.JOB_NUM = C.JOB_NUM " & _
                    " ORDER BY A.EMP_NAME "

    'Call ExecuteSQL and retrieve information
    Set GetEmployeesForProjectByWorkgroup = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    Exit Function
GetEmployeesForProjectByWorkgroup_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForProjectByWorkgroup"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForProjectByWorkgroup" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForProjectByWorkgroup" _
                    , , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeesForProjectByLocation(ByVal vlProjectID As Long _
                    , ByVal vlLocationCode As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset

On Error GoTo GetEmployeesForProjectByLocation_EH
'SQL Statement Container
Dim lsSQL As String

    lsSQL = " SELECT  A.EMP_NUM, A.EMP_NAME, A.UNION_CD " & _
                    "       , D.CRAFT_NUM, A.LOC_CD " & _
                    " FROM  METHDBA.TTR_EMPLOYEE A, METHDBA.TTR_JOBS D " & _
                    " WHERE A.LOC_CD = " & CStr(vlLocationCode) & _
                    " AND   A.JOB_NUM = D.JOB_NUM " & _
                    " AND   A.EXEC_CONTRACT_IND <> 'E' " & _
                    " AND   A.EMP_NUM IN (SELECT B.EMP_NUM " & _
                    "           FROM    METHDBA.TTR_EMP_WG B, METHDBA.TTR_WORK_GROUP C " & _
                    "           WHERE   B.WG_NUM = C.WG_NUM " & _
                    "           AND     B.EFF_START_DATE <= '" & Format("mm/dd/yyyy", vdtStartDate) & "'" & _
                    "           AND     B.EFF_END_DATE >= '" & Format("mm/dd/yyyy", vdtEndDate) & "'" & _
                    "           AND     C.PROJECT_ID = " & CStr(vlProjectID) & ") " & _
                    " ORDER BY A.EMP_NAME "

    'Call ExecuteSQL and retrieve information
    Set GetEmployeesForProjectByLocation = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    Exit Function
GetEmployeesForProjectByLocation_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForProjectByLocation"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForProjectByLocation" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForProjectByLocation" _
                    , , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeesForProjectByCraft(ByVal vlProjectID As Long _
                    , ByVal vsCraftNumber As String _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeesForProjectByCraft_EH
'SQL Statement Container
Dim lsSQL As String

    lsSQL = " SELECT  A.EMP_NUM, A.EMP_NAME, A.UNION_CD " & _
                "       , D.CRAFT_NUM, A.LOC_CD " & _
                " FROM  METHDBA.TTR_EMPLOYEE A, METHDBA.TTR_JOBS D " & _
                " WHERE A.JOB_NUM IN (SELECT    JOB_NUM " & _
                "               FROM  METHDBA.TTR_JOBS B " & _
                "               WHERE B.CRAFT_NUM = " & vsCraftNumber & ") " & _
                " AND   A.EXEC_CONTRACT_IND <> 'E' " & _
                " AND   A.EMP_NUM IN (SELECT    B.EMP_NUM " & _
                "               FROM  METHDBA.TTR_EMP_WG B, METHDBA.TTR_WORK_GROUP C " & _
                "               WHERE B.WG_NUM = C.WG_NUM " & _
                "               AND   B.EFF_START_DATE <= '" & Format("mm/dd/yyyy", vdtStartDate) & "'" & _
                "               AND   B.EFF_END_DATE >= '" & Format("mm/dd/yyyy", vdtEndDate) & "'" & _
                "               AND   C.PROJECT_ID = " & CStr(vlProjectID) & ") " & _
                " AND    A.JOB_NUM = D.JOB_NUM " & _
                " ORDER BY A.EMP_NAME "

    'Call ExecuteSQL and retrieve information
    Set GetEmployeesForProjectByCraft = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockReadOnly)

    Exit Function
GetEmployeesForProjectByCraft_EH:
    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForProjectByCraft"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForProjectByCraft" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForProjectByCraft" _
                    , , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeeECWorkReasons(ByVal vbActiveEmployeesOnly As Boolean _
                    , ByVal vdtStatusEffectiveFromDate As Date _
                    ) As ADODB.Recordset
TrsRaiseError m_strCMODULENAME, "GetEmployeeECWorkReasons", trErrObsoleteMethod, , , , True
'On Error GoTo GetEmployeeECWorkReasons_EH
''SQL Statement Container
'Dim lsSQL As String
'
'    lsSQL = " SELECT  A.EMP_NUM, B.EMP_NAME " & _
'                "       , A.WORK_REASON " & _
'                "       , CASE (B.EMP_STATUS) " & _
'                "       WHEN 'A' THEN 'ACTIVE' " & _
'                "       WHEN 'L' THEN 'LTD' " & _
'                "       WHEN 'I' THEN 'INACTIVE' " & _
'                "       WHEN 'T' THEN 'TERMINATED' " & _
'                "       ELSE ' ' END AS EMP_STATUS " & _
'                "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP " & _
'                " FROM  METHDBA.TTR_EC_WORK_REAS A, METHDBA.TTR_EMPLOYEE B " & _
'                " WHERE A.EMP_NUM = B.EMP_NUM "
'    If (vbActiveEmployeesOnly) Then
'        lsSQL = lsSQL & _
'                " AND   (B.EMP_STATUS <> 'T' " & _
'                " OR    (B.EMP_STATUS = 'T' AND B.STATUS_EFF_DATE > CURRENT DATE - 1 MONTHS))"
'    Else
'        lsSQL = lsSQL & _
'                " AND   B.EMP_STATUS IN ('A', 'L', 'T', 'I')"
'    End If
'
'    lsSQL = lsSQL & " ORDER BY A.EMP_NUM;"
'
'    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeECWorkReasons = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
'
'    Exit Function
'GetEmployeeECWorkReasons_EH:
'    'Call Standard Module Function
'    'Raise the appropriate error message depending on the error number
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError m_strCMODULENAME _
'                    , "GetEmployeeECWorkReasons"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError m_strCMODULENAME _
'                    , "GetEmployeeECWorkReasons" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(lsSQL) _
'                    , True
'            Case Else
'                TrsRaiseError m_strCMODULENAME _
'                    , "GetEmployeeECWorkReasons" _
'                    , trErrdbGetECWorkReasons _
'                    , , _
'                    , DescArgsFrom(lsSQL) _
'                    , True
'    End Select

End Function
'Get Sick Leave Details
Public Function GetEmployeeSickLeaveInformation(ByVal vlEmployeeNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeSickLeaveInformation_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetEmployeeSickLeaveInformation = Nothing

    lsSQL = "SELECT A.OCCUR_DT, A.DATE_WORKED" & _
            "       , A.SYMBOL_CD, A.START_TIME" & _
            "       , A.END_TIME, A.HOURS" & _
            "       , A.CASE_NUM, A.BIG_BANK_FULL_HRS" & _
            "       , A.BIG_BANK_HALF_HRS, A.BIG_BANK_DELAY_HRS" & _
            "       , CASE COALESCE(A.BIG_BANK_RST_DATE, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.BIG_BANK_RST_DATE, DATE('01/01/0100')) END AS BIG_BANK_RST_DATE" & _
            "       , A.LIT_BANK_TOT_HRS, A.LIT_BANK_DELAY_HRS" & _
            "       , CASE COALESCE(A.LIT_BANK_RST_DATE, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.LIT_BANK_RST_DATE, DATE('01/01/0100')) END AS LIT_BANK_RST_DATE" & _
            "       , A.FIRST_HR_NO, A.BASIC_HR_NO" & _
            "       , A.ACCUM_HR_NO, A.ACCUM_OFLOW_HR_NO" & _
            "       , A.FMLA_IND, A.JOB_NUM, A.UNION_CD" & _
            "       , A.EXEC_CONTRACT_IND, A.COMP_CD" & _
            "       , A.BEN_PKG_IND, A.EXTRA_FIRST_HR_NO AS YR15_FH_NO"

    lsSQL = lsSQL & _
            "       , CASE COALESCE(A.EMPMT_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.EMPMT_DT, DATE('01/01/0100')) END AS EMPMT_DT" & _
            "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP " & _
            "       , CASE COALESCE(B.SL_ELIG_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.SL_ELIG_DT, DATE('01/01/0100')) END AS SL_ELIG_DT" & _
            "       , CASE COALESCE(C.BIG_BANK_INT_DATE, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(C.BIG_BANK_INT_DATE, DATE('01/01/0100')) END AS BIG_BANK_INT_DATE" & _
            "       , CASE COALESCE(C.LIT_BANK_INT_DATE, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(C.LIT_BANK_INT_DATE, DATE('01/01/0100')) END AS LIT_BANK_INT_DATE" & _
            "       , C.ALWD_NON_AH_NO "

        lsSQL = lsSQL & _
            " FROM  METHDBA.TTR_EMP_SL_DTL A" & _
            "       LEFT OUTER JOIN METHDBA.TTR_EMP_SL_BANKS C" & _
            "       ON C.EMP_NUM = A.EMP_NUM" & _
            "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
            "       ON B.EMP_NUM =A.EMP_NUM" & _
            " WHERE A.EMP_NUM = ?" & _
            " ORDER BY  A.OCCUR_DT, A.DATE_WORKED, A.START_TIME, A.TIME_STAMP;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeSickLeaveInformation = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeeSickLeaveInformation = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeSickLeaveInformation_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSickLeaveInformation"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSickLeaveInformation" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSickLeaveInformation" _
                    , trErrdbGetSickLeaveDetails _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeeSickLeaveInformation2(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeSickLeaveInformation2_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    lsSQL = "SELECT A.OCCUR_DT, A.DATE_WORKED" & _
            "       , A.SYMBOL_CD, A.START_TIME" & _
            "       , A.END_TIME, A.HOURS" & _
            "       , A.CASE_NUM, A.BIG_BANK_FULL_HRS" & _
            "       , A.BIG_BANK_HALF_HRS, A.BIG_BANK_DELAY_HRS" & _
            "       , CASE COALESCE(A.BIG_BANK_RST_DATE, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.BIG_BANK_RST_DATE, DATE('01/01/0100')) END AS BIG_BANK_RST_DATE" & _
            "       , A.LIT_BANK_TOT_HRS, A.LIT_BANK_DELAY_HRS" & _
            "       , CASE COALESCE(A.LIT_BANK_RST_DATE, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.LIT_BANK_RST_DATE, DATE('01/01/0100')) END AS LIT_BANK_RST_DATE" & _
            "       , A.FIRST_HR_NO, A.BASIC_HR_NO" & _
            "       , A.ACCUM_HR_NO, A.ACCUM_OFLOW_HR_NO" & _
            "       , A.FMLA_IND, A.JOB_NUM, A.UNION_CD" & _
            "       , A.EXEC_CONTRACT_IND, A.COMP_CD" & _
            "       , A.BEN_PKG_IND, A.EXTRA_FIRST_HR_NO AS YR15_FH_NO"

    lsSQL = lsSQL & _
            "       , CASE COALESCE(A.EMPMT_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(A.EMPMT_DT, DATE('01/01/0100')) END AS EMPMT_DT" & _
            "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP " & _
            "       , CASE COALESCE(B.SL_ELIG_DT, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(B.SL_ELIG_DT, DATE('01/01/0100')) END AS SL_ELIG_DT" & _
            "       , CASE COALESCE(C.BIG_BANK_INT_DATE, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(C.BIG_BANK_INT_DATE, DATE('01/01/0100')) END AS BIG_BANK_INT_DATE" & _
            "       , CASE COALESCE(C.LIT_BANK_INT_DATE, DATE('01/01/0100'))" & _
            "       WHEN DATE('01/01/0001') THEN DATE('01/01/0100')" & _
            "       ELSE COALESCE(C.LIT_BANK_INT_DATE, DATE('01/01/0100')) END AS LIT_BANK_INT_DATE" & _
            "       , C.ALWD_NON_AH_NO "

        lsSQL = lsSQL & _
            " FROM  METHDBA.TTR_EMP_SL_DTL A" & _
            "       LEFT OUTER JOIN METHDBA.TTR_EMP_SL_BANKS C" & _
            "       ON C.EMP_NUM = A.EMP_NUM" & _
            "       INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
            "       ON B.EMP_NUM =A.EMP_NUM" & _
            " WHERE A.EMP_NUM = ?" & _
            " AND   A.OCCUR_DT BETWEEN ?" & _
            " AND   ?" & _
            " ORDER BY  A.OCCUR_DT, A.DATE_WORKED, A.START_TIME, A.TIME_STAMP;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    Set GetEmployeeSickLeaveInformation2 = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

'    lsSQL = "SELECT A.OCCUR_DT, A.DATE_WORKED " & _
'            "       , A.SYMBOL_CD, A.START_TIME " & _
'            "       , A.END_TIME, A.HOURS " & _
'            "       , A.CASE_NUM, A.BIG_BANK_FULL_HRS " & _
'            "       , A.BIG_BANK_HALF_HRS, A.BIG_BANK_DELAY_HRS " & _
'            "       , VALUE(NULLIF(A.BIG_BANK_RST_DATE,'0001-01-01'), DATE('0100-01-01')) AS BIG_BANK_RST_DATE " & _
'            "       , A.LIT_BANK_TOT_HRS, A.LIT_BANK_DELAY_HRS " & _
'            "       , VALUE(NULLIF(A.LIT_BANK_RST_DATE,'0001-01-01'), DATE('0100-01-01')) AS LIT_BANK_RST_DATE " & _
'            "       , A.FIRST_HR_NO, A.BASIC_HR_NO " & _
'            "       , A.ACCUM_HR_NO, A.ACCUM_OFLOW_HR_NO " & _
'            "       , A.FMLA_IND, A.JOB_NUM, A.UNION_CD " & _
'            "       , A.EXEC_CONTRACT_IND, A.COMP_CD " & _
'            "       , A.BEN_PKG_IND, A.EXTRA_FIRST_HR_NO AS YR15_FH_NO " & _
'            "       , VALUE(NULLIF(A.EMPMT_DT,'0001-01-01'), DATE('0100-01-01')) AS EMPMT_DT " & _
'            "       , A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP " & _
'            "       , VALUE(NULLIF(B.SL_ELIG_DT,'0001-01-01'), DATE('0100-01-01')) AS SL_ELIG_DT " & _
'            "       , VALUE(NULLIF(C.BIG_BANK_INT_DATE,'0001-01-01'), DATE('0100-01-01')) AS BIG_BANK_INT_DATE " & _
'            "       , VALUE(NULLIF(C.LIT_BANK_INT_DATE,'0001-01-01'), DATE('0100-01-01')) AS LIT_BANK_INT_DATE " & _
'            "       , C.ALWD_NON_AH_NO "
'        lsSQL = lsSQL & _
'            " FROM  METHDBA.TTR_EMP_SL_DTL A, METHDBA.TTR_EMPLOYEE B " & _
'            "       , METHDBA.TTR_EMP_SL_BANKS C " & _
'            " WHERE A.EMP_NUM = " & CStr(vlEmployeeNumber) & _
'            " AND   A.OCCUR_DT BETWEEN '" & Format(vdtStartDate, "mm/dd/yyyy") & "'" & _
'            "       AND '" & Format(vdtEndDate, "mm/dd/yyyy") & "'" & _
'            " AND   A.EMP_NUM = B.EMP_NUM " & _
'            " AND   A.EMP_NUM = C.EMP_NUM " & _
'            " ORDER BY  A.OCCUR_DT, A.DATE_WORKED, A.START_TIME, A.TIME_STAMP"

'    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeSickLeaveInformation2 = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)

    Exit Function
GetEmployeeSickLeaveInformation2_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSickLeaveInformation2"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSickLeaveInformation2" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeSickLeaveInformation2" _
                    , trErrdbGetSickLeaveDetails _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeeAllowedTimeDetails(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeAllowedTimeDetails_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetEmployeeAllowedTimeDetails = Nothing

    Set ladrsParms = CreateParameterRS

    'From Allowed Time Details
    lsSQL = "SELECT CHARGE_DATE, SYMBOL_CD" & _
            "       , START_TIME, END_TIME" & _
            "       , HOURS, EXCEPTION_CD" & _
            "       , DATE_WORKED, JOB_NUM" & _
            "       , DEPT_NUM, FUNCTION_NUM" & _
            "       , DIVISION_NUM, CMPY_CD" & _
            "       , UNION_CD, CASE_NUM" & _
            "       , FMLA_IND, PAID_UNPD_FL" & _
            "       , EXEC_CONTRACT_IND, ADJ_IND" & _
            " FROM  METHDBA.TTR_EMP_ALW_TM_DTL" & _
            " WHERE EMP_NUM = ?" & _
            " AND   CHARGE_DATE BETWEEN ?" & _
            " AND   ?" & _
            " AND   SYMBOL_CD <> ''"

    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    lsSQL = lsSQL & " UNION ALL "

    'From Actual Time
    lsSQL = lsSQL & _
            " SELECT    A.CHARGE_DATE, B.SYMBOL_CD" & _
            "           , A.START_TIME, A.END_TIME" & _
            "           , A.HOURS, A.EXCEPTION_CD" & _
            "           , A.DATE_WORKED, A.JOB_NUM" & _
            "           , A.DEPT_NUM, '' AS FUNCTION_NUM" & _
            "           , '' AS DIVISION_NUM, '' AS CMPY_CD " & _
            "           , A.UNION_CD, A.CASE_NUM" & _
            "           , A.FAMILY_LEAVE_IND AS FMLA_IND, '' AS PAID_UNPD_FL" & _
            "           , '' AS EXEC_CONTRACT_IND, 'N' AS ADJ_IND" & _
            " FROM      METHDBA.TTR_ACTUAL_TIME A, METHDBA.TTR_EXCEPTION_CDS B" & _
            " WHERE     A.EMP_NUM = ?" & _
            " AND       A.CHARGE_DATE BETWEEN ?" & _
            " AND       ?" & _
            " AND       A.PROCESSED_IND IN ('N', '')" & _
            " AND       A.EXCEPTION_CD = B.EXCEPTION_CD" & _
            " AND       B.SYMBOL_CD <> ''"

    With ladrsParms
        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 5
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 6
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    lsSQL = lsSQL & " UNION ALL "

    'Cutoff Old Tbl Vs. New Tbl (Jan. 26, 2005)
    If (vdtStartDate < #1/27/2005#) Then
        'From Adjustments (old and new)
        lsSQL = lsSQL & _
            " SELECT    A.DATE_WORKED AS CHARGE_DATE, B.SYMBOL_CD" & _
            "           , A.START_TIME, A.END_TIME" & _
            "           , A.HOURS, A.EXCEPTION_CD" & _
            "           , A.DATE_WORKED, A.JOB_NUM" & _
            "           , A.DEPT_NUM, '' AS FUNCTION_NUM" & _
            "           , '' AS DIVISION_NUM, '' AS CMPY_CD" & _
            "           , A.UNION_CD, A.CASE_NUM" & _
            "           , A.FAMILY_LEAVE_IND AS FMLA_IND, '' AS PAID_UNPD_FL " & _
            "           , '' AS EXEC_CONTRACT_IND, 'Y' AS ADJ_IND" & _
            " FROM      METHDBA.TTR_ADJUSTMENTS A, METHDBA.TTR_EXCEPTION_CDS B" & _
            " WHERE     A.EMP_NUM = ?" & _
            " AND       A.DATE_WORKED BETWEEN ?" & _
            " AND       ?" & _
            " AND       A.PROCESSED_IND IN ('N', '')" & _
            " AND       A.EXCEPTION_CD = B.EXCEPTION_CD" & _
            " AND       B.SYMBOL_CD <> ''"

        With ladrsParms
            .AddNew
            !PARM_ORDER = 7
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vlEmployeeNumber
            .Update

            .AddNew
            !PARM_ORDER = 8
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtStartDate
            .Update

            .AddNew
            !PARM_ORDER = 9
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtEndDate
            .Update
        End With

        lsSQL = lsSQL & " UNION ALL "

        lsSQL = lsSQL & _
            " SELECT    A.CHARGE_DATE, B.SYMBOL_CD" & _
            "           , A.START_TIME, A.END_TIME" & _
            "           , A.HOURS, A.EXCEPTION_CD" & _
            "           , A.DATE_WORKED, A.JOB_NUM" & _
            "           , A.DEPT_NUM, '' AS FUNCTION_NUM" & _
            "           , '' AS DIVISION_NUM, '' AS CMPY_CD" & _
            "           , A.UNION_CD, A.CASE_NUM" & _
            "           , A.FAMILY_LEAVE_IND AS FMLA_IND, '' AS PAID_UNPD_FL" & _
            "           , '' AS EXEC_CONTRACT_IND, 'Y' AS ADJ_IND" & _
            " FROM      METHDBA.TTR_ADJ_TIME A, METHDBA.TTR_EXCEPTION_CDS B" & _
            "           , METHDBA.TTR_ADJ_RQT C" & _
            " WHERE     C.EMP_NUM = ?" & _
            " AND       A.CHARGE_DATE BETWEEN ?" & _
            " AND       ?" & _
            " AND       A.EXCEPTION_CD = B.EXCEPTION_CD" & _
            " AND       B.SYMBOL_CD <> ''" & _
            " AND       A.ADJ_RQT_NUM = C.ADJ_RQT_NUM" & _
            " AND       C.FINAL_CYCLE_NUM IS NULL"

        With ladrsParms
            .AddNew
            !PARM_ORDER = 10
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vlEmployeeNumber
            .Update

            .AddNew
            !PARM_ORDER = 11
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtStartDate
            .Update

            .AddNew
            !PARM_ORDER = 12
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtEndDate
            .Update
        End With
    Else
        lsSQL = lsSQL & _
            " SELECT    A.CHARGE_DATE, B.SYMBOL_CD" & _
            "           , A.START_TIME, A.END_TIME" & _
            "           , A.HOURS, A.EXCEPTION_CD" & _
            "           , A.DATE_WORKED, A.JOB_NUM" & _
            "           , A.DEPT_NUM, '' AS FUNCTION_NUM" & _
            "           , '' AS DIVISION_NUM, '' AS CMPY_CD" & _
            "           , A.UNION_CD, A.CASE_NUM" & _
            "           , A.FAMILY_LEAVE_IND AS FMLA_IND, '' AS PAID_UNPD_FL" & _
            "           , '' AS EXEC_CONTRACT_IND, 'Y' AS ADJ_IND" & _
            " FROM      METHDBA.TTR_ADJ_TIME A, METHDBA.TTR_EXCEPTION_CDS B" & _
            "           , METHDBA.TTR_ADJ_RQT C" & _
            " WHERE     C.EMP_NUM = ?" & _
            " AND       A.CHARGE_DATE BETWEEN ?" & _
            " AND       ?" & _
            " AND       A.EXCEPTION_CD = B.EXCEPTION_CD" & _
            " AND       B.SYMBOL_CD <> ''" & _
            " AND       A.ADJ_RQT_NUM = C.ADJ_RQT_NUM" & _
            " AND       C.FINAL_CYCLE_NUM IS NULL"

        With ladrsParms
            .AddNew
            !PARM_ORDER = 7
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vlEmployeeNumber
            .Update

            .AddNew
            !PARM_ORDER = 8
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtStartDate
            .Update

            .AddNew
            !PARM_ORDER = 9
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtEndDate
            .Update
        End With
    End If
    lsSQL = lsSQL & " ORDER BY 1, 6, 3;"

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeAllowedTimeDetails = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeeAllowedTimeDetails = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeAllowedTimeDetails_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeAllowedTimeDetails"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeAllowedTimeDetails" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeAllowedTimeDetails" _
                    , trErrdbGetAllowedTimeDetails _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeeAllowedTimeTotals(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeAllowedTimeTotals_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetEmployeeAllowedTimeTotals = Nothing

    Set ladrsParms = CreateParameterRS

    'Get Sub Total and the Grand total
    'from Allowed Time Details, Actual Time, Actual Time History
    ' and Adjustments
    lsSQL = "SELECT    SYMBOL_CD, SUM(SUB_TOT_HOURS) AS TOT_HOURS" & _
        " FROM      (SELECT SYMBOL_CD, SUM(HOURS) AS SUB_TOT_HOURS" & _
        "           FROM    METHDBA.TTR_EMP_ALW_TM_DTL" & _
        "           WHERE   EMP_NUM = ?" & _
        "           AND     CHARGE_DATE BETWEEN ?" & _
        "           AND     ?" & _
        "           AND     SYMBOL_CD <> ''" & _
        "           GROUP BY SYMBOL_CD" & _
        "           UNION ALL " & _
        "           SELECT  B.SYMBOL_CD, SUM(A.HOURS) AS SUB_TOT_HOURS" & _
        "           FROM    METHDBA.TTR_ACTUAL_TIME A, METHDBA.TTR_EXCEPTION_CDS B" & _
        "           WHERE   A.EMP_NUM = ?" & _
        "           AND     A.CHARGE_DATE BETWEEN ?" & _
        "           AND     ?" & _
        "           AND     A.EXCEPTION_CD = B.EXCEPTION_CD" & _
        "           AND     A.PROCESSED_IND IN ('N', '')" & _
        "           AND     B.SYMBOL_CD <> ''" & _
        "           GROUP BY B.SYMBOL_CD"

    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update

        .AddNew
        !PARM_ORDER = 4
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 5
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 6
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    lsSQL = lsSQL & _
        "           UNION ALL "

    'Cutoff Old Tbl Vs. New Tbl (Jan. 26, 2005)
    If (vdtStartDate < #1/27/2005#) Then
        'From Adjustments (old and new)
        lsSQL = lsSQL & _
            "           SELECT  B.SYMBOL_CD, SUM(A.HOURS) AS SUB_TOT_HOURS" & _
            "           FROM    METHDBA.TTR_ADJUSTMENTS A, METHDBA.TTR_EXCEPTION_CDS B" & _
            "           WHERE   A.EMP_NUM = ?" & _
            "           AND     A.DATE_WORKED BETWEEN ?" & _
            "           AND     ?" & _
            "           AND     A.EXCEPTION_CD = B.EXCEPTION_CD" & _
            "           AND     A.PROCESSED_IND IN ('N', '')" & _
            "           AND     B.SYMBOL_CD <> ''" & _
            "           GROUP BY B.SYMBOL_CD"

        With ladrsParms
            .AddNew
            !PARM_ORDER = 7
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vlEmployeeNumber
            .Update

            .AddNew
            !PARM_ORDER = 8
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtStartDate
            .Update

            .AddNew
            !PARM_ORDER = 9
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtEndDate
            .Update
        End With

        lsSQL = lsSQL & _
            "           UNION ALL "

        lsSQL = lsSQL & _
            "           SELECT  B.SYMBOL_CD, SUM(A.HOURS) AS SUB_TOT_HOURS" & _
            "           FROM    METHDBA.TTR_ADJ_TIME A, METHDBA.TTR_EXCEPTION_CDS B" & _
            "                   , METHDBA.TTR_ADJ_RQT C" & _
            "           WHERE   C.EMP_NUM = ?" & _
            "           AND     A.CHARGE_DATE BETWEEN ?" & _
            "           AND     ?" & _
            "           AND     A.EXCEPTION_CD = B.EXCEPTION_CD" & _
            "           AND     B.SYMBOL_CD <> '' " & _
            "           AND     A.ADJ_RQT_NUM = C.ADJ_RQT_NUM" & _
            "           AND     C.FINAL_CYCLE_NUM IS NULL" & _
            "           GROUP BY B.SYMBOL_CD " & _
            "           ) AS ONE "
    
        With ladrsParms
            .AddNew
            !PARM_ORDER = 10
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vlEmployeeNumber
            .Update

            .AddNew
            !PARM_ORDER = 11
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtStartDate
            .Update

            .AddNew
            !PARM_ORDER = 12
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtEndDate
            .Update
        End With
    Else
        lsSQL = lsSQL & _
            "           SELECT  B.SYMBOL_CD, SUM(A.HOURS) AS SUB_TOT_HOURS" & _
            "           FROM    METHDBA.TTR_ADJ_TIME A, METHDBA.TTR_EXCEPTION_CDS B" & _
            "                   , METHDBA.TTR_ADJ_RQT C" & _
            "           WHERE   C.EMP_NUM = ?" & _
            "           AND     A.CHARGE_DATE BETWEEN ?" & _
            "           AND     ?" & _
            "           AND     A.EXCEPTION_CD = B.EXCEPTION_CD" & _
            "           AND     B.SYMBOL_CD <> '' " & _
            "           AND     A.ADJ_RQT_NUM = C.ADJ_RQT_NUM" & _
            "           AND     C.FINAL_CYCLE_NUM IS NULL" & _
            "           GROUP BY B.SYMBOL_CD " & _
            "           ) AS ONE "

        With ladrsParms
            .AddNew
            !PARM_ORDER = 7
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vlEmployeeNumber
            .Update

            .AddNew
            !PARM_ORDER = 8
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtStartDate
            .Update

            .AddNew
            !PARM_ORDER = 9
            !PARM_NAME = ""
            !PARM_TYPE = adDate
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vdtEndDate
            .Update
        End With
    End If

    lsSQL = lsSQL & _
        " GROUP BY SYMBOL_CD " & _
        " ORDER BY 1;"

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeAllowedTimeTotals = ExecuteSQL(lsSQL _
'                            , adOpenStatic _
'                            , adLockReadOnly)
    Set GetEmployeeAllowedTimeTotals = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockReadOnly)
    Set ladrsParms = Nothing
    Exit Function
GetEmployeeAllowedTimeTotals_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeAllowedTimeTotals"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeAllowedTimeTotals" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeAllowedTimeTotals" _
                    , trErrdbGetAllowedTimeDetails _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeeCertificationSummary(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeCertificationSummary_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    lsSQL = "SELECT EMP_NUM, UNION_CD" & _
        "       , BANK_CD, CHARGE_YEAR" & _
        "       , ACCUM_CERT_HRS, ACCUM_NON_CERT_HRS" & _
        "       , EFF_PNLTY_WAIT_HRS, RESTR_PNLTY_HRS" & _
        " FROM  METHDBA.TTR_EMP_SL_CRT_HRS" & _
        " WHERE EMP_NUM = ?" & _
        " AND   CHARGE_YEAR BETWEEN ?" & _
        " AND   ?" & _
        " ORDER BY CHARGE_YEAR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = Year(vdtStartDate)
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = Year(vdtEndDate)
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeCertificationSummary = ExecuteSQL(lsSQL _
'                    , adOpenStatic _
'                    , adLockReadOnly)
    Set GetEmployeeCertificationSummary = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenStatic _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeCertificationSummary_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationSummary"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationSummary" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationSummary" _
                    , trErrdbGetCertSummary _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeeCertificationDetails(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeCertificationDetails_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    'Begin Prem TRIS 5.8
    'lsSQL = "SELECT    EMP_NUM, UNION_CD" & _
        "       , CHARGE_DATE, START_TIME" & _
        "       , CASE_NUM, EXCEPTION_CD" & _
        "       , CHARGE_YEAR, BANK_YEAR" & _
        "       , SOURCE_CD, RPTD_HRS" & _
        "       , ILL_HRS, INJ_HRS" & _
        "       , CERT_HRS, NON_CERT_HRS" & _
        "       , PNLTY_WAIT_HRS" & _
        " FROM  METHDBA.TTR_EMP_SL_HRS" & _
        " WHERE EMP_NUM = ?" & _
        " AND   CHARGE_DATE BETWEEN ?" & _
        " AND   ?" & _
        " ORDER BY CHARGE_DATE;"

    lsSQL = "SELECT EMP_NUM, BANK_CD" & _
            "       , START_DT, END_DT" & _
            "       , CERT_CD" & _
            "       , CREATED_EMP_NUM, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
            "       , LAST_MOD_EMP_NUM, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
            " FROM  METHDBA.TTR_EMP_SL_CERT" & _
            " WHERE EMP_NUM = ?" & _
            " AND   START_DT >= ?" & _
            " AND   END_DT <= ?" & _
            " ORDER BY START_DT"
    'End Prem TRIS 5.8

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    Set GetEmployeeCertificationDetails = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenStatic _
                    , adLockReadOnly)

    Set ladrsParms = Nothing

    Exit Function
GetEmployeeCertificationDetails_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationDetails"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationDetails" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationDetails" _
                    , trErrdbGetCertDetails _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function

'//TODO: this is probably obsolete
Public Function GetEmployeeExcusedTimeInformation(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeExcusedTimeInformation_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    lsSQL = "SELECT    *" & _
        " FROM      (SELECT     NTE0.CHARGE_DATE, NTE0.RPTD_HRS " & _
        "                       , CASE WHEN C.EMP_NUM IS NULL" & _
        "                       THEN NTE0.RPTD_HRS" & _
        "                       WHEN SUM(COALESCE(C.UNEXCUSED_HRS, 0)) <= 0" & _
        "                       THEN 0" & _
        "                       WHEN SUM(COALESCE(C.UNEXCUSED_HRS, 0)) <= NTE0.RPTD_HRS" & _
        "                       THEN SUM(COALESCE(C.UNEXCUSED_HRS, 0))" & _
        "                       ELSE NTE0.RPTD_HRS END AS UNEXCUSED_HRS" & _
        "                       , CASE WHEN C.EMP_NUM IS NULL" & _
        "                       THEN 0" & _
        "                       WHEN SUM(COALESCE(C.UNEXCUSED_HRS, 0)) <= 0" & _
        "                       THEN NTE0.RPTD_HRS" & _
        "                       WHEN SUM(COALESCE(C.UNEXCUSED_HRS, 0)) <= NTE0.RPTD_HRS" & _
        "                       THEN NTE0.RPTD_HRS - SUM(COALESCE(C.UNEXCUSED_HRS,0))" & _
        "                       ELSE 0 END AS EXCUSED_HRS"

    lsSQL = lsSQL & _
        "           FROM        (SELECT A.EMP_NUM, A.CHARGE_DATE" & _
        "                               , SUM(A.HOURS) AS RPTD_HRS" & _
        "                       FROM    METHDBA.TTR_EMP_ALW_TM_DTL A" & _
        "                               INNER JOIN METHDBA.TTR_EXCEPTION_CDS B" & _
        "                               ON B.EXCEPTION_CD = A.EXCEPTION_CD" & _
        "                               AND B.CASE_GROUP IN ('ILL','INJ')" & _
        "                       WHERE   A.EMP_NUM = ?" & _
        "                       AND     A.CHARGE_DATE BETWEEN ?" & _
        "                       AND     ?" & _
        "                       AND     A.EXEC_CONTRACT_IND = 'C'" & _
        "                       AND     A.SYMBOL_CD NOT IN ('VE', 'VED')" & _
        "                       GROUP BY A.EMP_NUM, A.CHARGE_DATE) NTE0" & _
        "                       LEFT OUTER JOIN METHDBA.TTR_EMP_SL_EXCUSED C" & _
        "                       ON  C.EMP_NUM = NTE0.EMP_NUM" & _
        "                       AND NTE0.CHARGE_DATE BETWEEN C.START_DT" & _
        "                       AND C.END_DT" & _
        "                       GROUP BY C.EMP_NUM, NTE0.CHARGE_DATE" & _
        "                               , NTE0.RPTD_HRS) NTE1"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    'Call ExecuteSQL and retrieve information
'    Set GetEmployeeExcusedTimeInformation = ExecuteSQL(lsSQL _
'                    , adOpenStatic _
'                    , adLockReadOnly)
    Set GetEmployeeExcusedTimeInformation = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenStatic _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeExcusedTimeInformation_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeExcusedTimeInformation"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeExcusedTimeInformation" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeExcusedTimeInformation" _
                    , trErrdbGetExcusedTime _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'PREM END TRIS CONVERSION PHASE 3 (PHASE 1 .NET)

'JIMB TRIS 5 enhancement

Public Function GetPeopleSoftPhoneNumber( _
            ByVal v_aryEmployees As Variant _
            ) As ADODB.Recordset
            
    
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("PeopleSoft")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = " AND " & BuildInClause("EMPLID", v_aryEmployees, True) & " "
    Else
        l_strEmployeeSelection = " AND EMPLID = '" & CStr(CLng(v_aryEmployees)) & "' "
    End If

    l_strSQL = "SELECT CAST(SUBSTR(EMPLID,1,9) AS INTEGER) as EMP_NUM, " & _
                " SUBSTR(PHONE,1,3) || SUBSTR(PHONE,5,3) || " & _
                " SUBSTR(PHONE,9,4) as HOME_PHONE_NUM " & _
            " FROM PS_PERSONAL_PHONE " & _
            " WHERE PHONE_TYPE = 'MAIN' " & l_strEmployeeSelection & _
            " ORDER BY 1"
                
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetPeopleSoftPhoneNumber = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

eh:
    If l_objCN.Errors.Count Then
        MsgBox "Error Getting a Connection to PeopleSoft", , "Hey Tim"
        Set GetPeopleSoftPhoneNumber = Nothing
        Set l_objCM = Nothing
        Set l_objCN = Nothing
        Set l_objRS = Nothing
        Exit Function
    End If
        
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPeopleSoftPhoneNumber", , , , DescArgsFrom(l_strSQL)
    
End Function
'Begin Prem TRIS 5.7
Public Function GetEmployeePSContactInformation(ByVal vlEmployee As Long _
                ) As ADODB.Recordset
On Error GoTo eh
    
    Dim datahub As Ameren_TRIS_Server_COMIntegration4.HRDataHubEmployee
    Dim rs As ADODB.Recordset
    
    Set datahub = New Ameren_TRIS_Server_COMIntegration4.HRDataHubEmployee
    
    Set rs = datahub.GetEmployeePSContactInformation_(vlEmployee)
    
    Set GetEmployeePSContactInformation = rs
    
    Set datahub = Nothing
    Set rs = Nothing

    Exit Function
eh:
    Set datahub = Nothing
    Set rs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeePSContactInformation", , , , DescArgsFrom(CStr(vlEmployee))
End Function
'End Prem TRIS 5.7
'Begin Prem TRIS 5.8
Public Function GetEmployeeCertificationByYear(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeCertificationByYear_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    lsSQL = "SELECT EMP_NUM, UNION_CD" & _
        "       , BANK_CD, CHARGE_YEAR" & _
        "       , ACCUM_CERT_HRS, ACCUM_NON_CERT_HRS" & _
        "       , EFF_PNLTY_WAIT_HRS, RESTR_PNLTY_HRS" & _
        " FROM  METHDBA.TTR_EMP_SL_CRT_HRS" & _
        " WHERE EMP_NUM = ?" & _
        " AND   CHARGE_YEAR BETWEEN ?" & _
        " AND   ?" & _
        " ORDER BY CHARGE_YEAR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = Year(vdtStartDate)
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = Year(vdtEndDate)
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    Set GetEmployeeCertificationByYear = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenStatic _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeCertificationByYear_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationByYear"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationByYear" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationByYear" _
                    , trErrdbGetCertSummary _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeeCertificationByDate(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeCertificationByDate_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    lsSQL = "SELECT  A.EMP_NUM, A.UNION_CD" & _
        "       , A.CHARGE_DATE, A.START_TIME" & _
        "       , A.CASE_NUM, A.EXCEPTION_CD" & _
        "       , A.CHARGE_YEAR, A.BANK_YEAR" & _
        "       , A.SOURCE_CD, A.RPTD_HRS" & _
        "       , A.ILL_HRS, A.INJ_HRS" & _
        "       , CASE WHEN A.EXCEPTION_CD = 'VE' OR A.EXCEPTION_CD = 'VED'" & _
        "       THEN 0" & _
        "       WHEN C.EMP_NUM IS NOT NULL" & _
        "       THEN 0" & _
        "       WHEN B.EMP_NUM  IS NULL" & _
        "       THEN A.RPTD_HRS" & _
        "       WHEN B.UNEXCUSED_HRS <= 0" & _
        "       THEN 0" & _
        "       WHEN B.UNEXCUSED_HRS <= A.RPTD_HRS" & _
        "       THEN B.UNEXCUSED_HRS" & _
        "       ELSE A.RPTD_HRS END AS UNEXCUSED_HRS"
    lsSQL = lsSQL & _
        "       , CASE WHEN A.EXCEPTION_CD = 'VE' OR A.EXCEPTION_CD = 'VED'" & _
        "       THEN A.RPTD_HRS" & _
        "       WHEN C.EMP_NUM IS NOT NULL" & _
        "       THEN A.RPTD_HRS" & _
        "       WHEN B.EMP_NUM  IS NULL" & _
        "       THEN 0" & _
        "       WHEN B.UNEXCUSED_HRS <= 0" & _
        "       THEN A.RPTD_HRS" & _
        "       WHEN B.UNEXCUSED_HRS <= A.RPTD_HRS" & _
        "       THEN A.RPTD_HRS - B.UNEXCUSED_HRS" & _
        "       ELSE 0 END AS EXCUSED_HRS" & _
        "       , CASE WHEN C.EMP_NUM IS NOT NULL" & _
        "       THEN A.RPTD_HRS" & _
        "       ELSE 0 END AS FMLA_HRS" & _
        "       , CASE WHEN D.EMP_NUM IS NOT NULL" & _
        "       THEN A.RPTD_HRS" & _
        "       ELSE 0 END AS CERT_HRS" & _
        "       , CASE WHEN D.EMP_NUM IS NULL AND C.EMP_NUM IS NULL AND A.CERT_HRS > 0" & _
        "       THEN A.CERT_HRS" & _
        "       ELSE 0 END AS OTHER_HRS" & _
        "       , A.NON_CERT_HRS, A.PNLTY_WAIT_HRS"
    lsSQL = lsSQL & _
        "       , CASE WHEN A.EXCEPTION_CD = 'VE' OR A.EXCEPTION_CD = 'VED'" & _
        "       THEN 'Y'" & _
        "       WHEN B.EMP_NUM IS NOT NULL" & _
        "       THEN 'Y'" & _
        "       WHEN C.EMP_NUM IS NOT NULL" & _
        "       THEN 'Y'" & _
        "       ELSE 'N' END AS EXCUSED_IND" & _
        "       , CASE WHEN C.EMP_NUM IS NOT NULL" & _
        "       THEN 'Y'" & _
        "       ELSE 'N' END AS FMLA_IND" & _
        "       , CASE WHEN D.EMP_NUM IS NOT NULL" & _
        "       THEN 'Y'" & _
        "       ELSE 'N' END AS CERT_IND"
    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_EMP_SL_HRS A" & _
        "       LEFT OUTER JOIN" & _
        "       METHDBA.TTR_EMP_SL_EXCUSED B" & _
        "       ON  B.EMP_NUM = A.EMP_NUM" & _
        "       AND A.CHARGE_DATE BETWEEN B.START_DT AND B.END_DT" & _
        "       LEFT OUTER JOIN" & _
        "       METHDBA.TTR_EMP_FMLA C" & _
        "       ON  C.EMP_NUM = A.EMP_NUM" & _
        "       AND A.CHARGE_DATE BETWEEN C.START_DT AND C.END_DT" & _
        "       AND C.FMLA_REASON_CD = 'EI'" & _
        "       LEFT OUTER JOIN" & _
        "       METHDBA.TTR_EMP_SL_CERT D" & _
        "       ON  D.EMP_NUM = A.EMP_NUM" & _
        "       AND A.CHARGE_DATE BETWEEN D.START_DT AND D.END_DT" & _
        " WHERE A.EMP_NUM = ?" & _
        " AND   A.CHARGE_DATE BETWEEN ?" & _
        " AND   ?" & _
        " ORDER BY A.CHARGE_DATE;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    Set GetEmployeeCertificationByDate = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenStatic _
                    , adLockReadOnly)

    Set ladrsParms = Nothing

    Exit Function
GetEmployeeCertificationByDate_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationByDate"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationByDate" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeCertificationByDate" _
                    , trErrdbGetCertSummary _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function

Public Function GetEmployeeExcusedTimeSummary(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo eh
    Dim SQL As String
    Dim oADOHelper As utilADOHelper.cuADOHelper
    Dim rs As ADODB.Recordset
    
    'get the SQL
    SQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetEmployeeExcusedTimeSummary.sql")
    
    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelper
    oADOHelper.ADO.CreateNumericParameter "EmpNum", adInteger, 4, 0, vlEmployeeNumber
    oADOHelper.ADO.CreateCharacterParameter "StartDate", adDate, 8, vdtStartDate
    oADOHelper.ADO.CreateCharacterParameter "EndDate", adDate, 8, vdtEndDate
    
    'execute sql
    Set rs = oADOHelper.ExecuteSelect(SQL, adOpenStatic, adLockBatchOptimistic)

    Set oADOHelper = Nothing
    Set GetEmployeeExcusedTimeSummary = rs
    Exit Function
eh:
    Set oADOHelper = Nothing
    Set rs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeeExcusedTimeSummary", , , , DescArgsFrom(SQL, CStr(vlEmployeeNumber), CStr(vdtStartDate), CStr(vdtEndDate)), True
End Function

Public Function GetEmployeeExcusedTimeSummaryTotal(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo eh
    Dim SQL As String
    Dim oADOHelper As utilADOHelper.cuADOHelper
    Dim rs As ADODB.Recordset
    
    'get the SQL
    SQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetEmployeeExcusedTimeSummaryTotal.sql")
    
    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelper
    oADOHelper.ADO.CreateNumericParameter "EmpNum", adInteger, 4, 0, vlEmployeeNumber
    oADOHelper.ADO.CreateCharacterParameter "StartDate", adDate, 8, vdtStartDate
    oADOHelper.ADO.CreateCharacterParameter "EndDate", adDate, 8, vdtEndDate
    
    'execute sql
    Set rs = oADOHelper.ExecuteSelect(SQL, adOpenStatic, adLockReadOnly)

    Set oADOHelper = Nothing
    Set GetEmployeeExcusedTimeSummaryTotal = rs
    Exit Function
eh:
    Set oADOHelper = Nothing
    Set rs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeeExcusedTimeSummaryTotal", , , , DescArgsFrom(SQL, CStr(vlEmployeeNumber), CStr(vdtStartDate), CStr(vdtEndDate)), True
End Function

Public Function GetEmployeeExcusedTimeDetails(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeExcusedTimeDetails_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    lsSQL = "SELECT EMP_NUM, START_DT" & _
            "       , END_DT, UNEXCUSED_HRS" & _
            "       , CREATED_EMP_NUM, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
            "       , LAST_MOD_EMP_NUM, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
            " FROM  METHDBA.TTR_EMP_SL_EXCUSED" & _
            " WHERE EMP_NUM = ?" & _
            " AND   START_DT >= ?" & _
            " AND   END_DT <= ?" & _
            " ORDER BY START_DT;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    Set GetEmployeeExcusedTimeDetails = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenStatic _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeExcusedTimeDetails_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeExcusedTimeDetails"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeExcusedTimeDetails" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeExcusedTimeDetails" _
                    , trErrdbGetExcusedTimeDetails _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function GetEmployeeFMLADetails(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeFMLADetails_EH
'SQL Statement Container
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    lsSQL = "SELECT EMP_NUM, START_DT" & _
            "       , END_DT, FMLA_HRS" & _
            "       , 'Employee Illness' AS FMLA_REASON" & _
            "       , CREATED_EMP_NUM, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
            "       , LAST_MOD_EMP_NUM, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
            " FROM  METHDBA.TTR_EMP_FMLA" & _
            " WHERE EMP_NUM = ?" & _
            " AND   START_DT >= ?" & _
            " AND   END_DT <= ?" & _
            " AND   FMLA_REASON_CD = 'EI'" & _
            " ORDER BY START_DT;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlEmployeeNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    'Call ExecuteSQL and retrieve information
    Set GetEmployeeFMLADetails = NewExecuteSQL(lsSQL _
                    , ladrsParms _
                    , adOpenStatic _
                    , adLockReadOnly)
    Set ladrsParms = Nothing

    Exit Function
GetEmployeeFMLADetails_EH:
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFMLADetails"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFMLADetails" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeeFMLADetails" _
                    , trErrdbGetFMLADetails _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
'End Prem TRIS 5.8

Private Sub Class_Initialize()
    Set m_ResourceMgr = New utilADOHelper.cuResourceMgr
End Sub

Private Sub Class_Terminate()
    Set m_ResourceMgr = Nothing
End Sub
