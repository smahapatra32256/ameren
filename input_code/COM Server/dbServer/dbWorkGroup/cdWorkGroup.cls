VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdWorkGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdWorkGroup"
Private Const m_DBSCRIPTSROOT As String = "C:\TRIS SOLUTION\dbScripts\TRIS_db2"
Private Const m_DBSCRIPTSPATH As String = "\dbWorkGroup\cdWorkGroup"
Private m_ResourceMgr As utilADOHelper.cuResourceMgr

' Static data constants
Private Const CRECORDNOTFOUND = "RECORDNOTFOUND"
Private Const CDEFAULTDATE As String = "01/01/0100"

' Overtime exception code constants
Private Const CWG_GROUPNAME = "WORKGROUPCACHE"
' index for record dynamic array
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const CWG_WG_NUM = 1
Private Const CWG_WG_DESC = 2
Private Const CWG_EFF_START_DATE = 3
Private Const CWG_EFF_END_DATE = 4
Private Const CWG_PROJECT_ID = 5
Private Const CWG_DEMAND_IND = 6
Private Const CWG_WG_SUPERVISOR = 7
Private Const CWG_CREW_NUM = 8
Private Const CWG_DEPT_NUM = 9
Private Const CWG_LOC_CD = 10
Private Const CWG_PROJ_START_DATE = 11
Private Const CWG_PROJ_END_DATE = 12
Private Const CWG_DOJM_DIST = 13
Private Const CWG_DOJM_HQ = 14
Private Const CWG_IS_DOJM = 15
'//-----Should this be changed to 15 after testing for Record ???
Private Const CWG_LAST_SLOT = 14

Private maryFieldNames As Variant

'*********************************************************
' Begin Public procedures
'*********************************************************

Public Function GetWorkGroup(ByVal v_lngEmpNum As Long, ByVal v_datStartDate As Date, ByVal v_datEndDate As Date) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetWorkGroup_EH
   
    If v_lngEmpNum < 1 Then
        Err.Raise trErrdbInvalidZerolngEmpNum
    Else
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "select WG_NUM "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_EMP_WG "
    l_strSQL = l_strSQL & "WHERE EMP_NUM = " & v_lngEmpNum
    l_strSQL = l_strSQL & " AND (('" & Format(v_datStartDate, "mm/dd/yyyy") & "' BETWEEN EFF_START_DATE AND EFF_END_DATE) "
    l_strSQL = l_strSQL & " OR   ('" & Format(v_datEndDate, "mm/dd/yyyy") & "' BETWEEN EFF_START_DATE AND EFF_END_DATE)) "
    l_strSQL = l_strSQL & " ORDER BY WG_NUM "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWorkGroup = l_objRS
    
    Set l_objRS = Nothing
    Exit Function

GetWorkGroup_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetWorkGroup", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetWorkGroup", , , , DescArgsFrom(l_strSQL), True
    End Select
   
End Function

Public Function GetAccounting(ByVal v_lngWgNum As Long) As ADODB.Recordset
    
'Similar to GetWorkGroupDistribution
'This is not used in the context of a transaction
'GetAccounting gets called often
'The seperation makes better use of the database resources
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetAccounting_EH
    
    If v_lngWgNum < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    Else
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT DISTRIBUTION_ID, DESCRIPTION "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACCT_ASSIGNMTS "
    l_strSQL = l_strSQL & "WHERE WG_NUM = " & v_lngWgNum
    l_strSQL = l_strSQL & " ORDER BY DESCRIPTION "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetAccounting = l_objRS

    Set l_objRS = Nothing
    Exit Function

GetAccounting_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetAccounting", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetAccounting", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function

Public Function GetDepartmentNumber(ByVal v_lngWgNum As Long) As String
' BLUE056
    
On Error GoTo GetDepartmentNumber_EH
    If v_lngWgNum < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    End If

    GetDepartmentNumber = DepartmentNumber(v_lngWgNum)
    If GetDepartmentNumber = "" Then
        Err.Raise trErrdbRetrieveNoAssociations
    End If
  
    Exit Function
    
GetDepartmentNumber_EH:
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum, trErrdbRetrieveNoAssociations
            TrsRaiseError m_strCMODULENAME, "GetDepartmentNumber", Err.Number, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetDepartmentNumber", , , , , True
    End Select

End Function


'''''Public Function GetEmployeeWorkGroups(ByVal v_lngEmpNum As Long, ByVal v_datStartDate As Date, ByVal v_datEndDate As Date, Optional bln_PrimaryOnly As Boolean = False) As ADODB.Recordset
'''''
'''''    Dim l_objCN As New ADODB.Connection
'''''    Dim l_objrs As ADODB.Recordset
'''''    Dim l_strSQL As String
'''''    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'''''
'''''    On Error GoTo GetEmployeeWorkGroups_EH
'''''    If v_lngEmpNum < 1 Then
'''''        Err.Raise trErrdbInvalidZerolngEmpNum
'''''    Else
'''''    End If
'''''
'''''    Set l_objCN = New ADODB.Connection
'''''    Set l_objrs = New ADODB.Recordset
'''''    Set l_objCM = New clsuDBConnectMgr
'''''
'''''    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'''''
'''''    Set l_objCM = Nothing
'''''
'''''    l_objCN.Open
'''''
''''''The first select joins on ttr_projects to pick up the eff_start_date and
'''''' eff_end_date for PPM project, but this will only yield a resulting row when the project_id
'''''' is not equal to zero.  The second select will yield a row for non-PPM workgroup,
'''''' or when the project_id = 0.
'''''
'''''    l_strSQL = "SELECT B.WG_NUM, A.WG_DESC, B.EFF_START_DATE, B.EFF_END_DATE, "
'''''    l_strSQL = l_strSQL & "A.PROJECT_ID, B.PRIMARY_IND, A.WG_SUPERVISOR, C.LOC_CD, C.DEPT_NUM, "
'''''    l_strSQL = l_strSQL & "d.eff_start_date as proj_start_date, d.eff_end_date as proj_end_date "
'''''    l_strSQL = l_strSQL & "FROM METHDBA.TTR_WORK_GROUP A, "
'''''    l_strSQL = l_strSQL & "METHDBA.TTR_EMP_WG B, "
'''''    l_strSQL = l_strSQL & "METHDBA.TTR_ASSOCIATIONS C, "
'''''    l_strSQL = l_strSQL & "METHDBA.ttr_projects d "
'''''    l_strSQL = l_strSQL & "WHERE A.WG_NUM = B.WG_NUM AND A.WG_NUM = C.WG_NUM "
'''''    l_strSQL = l_strSQL & "AND B.EMP_NUM = " & v_lngEmpNum
'''''    l_strSQL = l_strSQL & " and a.project_id = d.project_id"
'''''    l_strSQL = l_strSQL & " and d.ppm_ind = 'Y'"
'''''    l_strSQL = l_strSQL & " AND (('" & Format(v_datStartDate, "mm/dd/yyyy") & "' BETWEEN B.EFF_START_DATE AND B.EFF_END_DATE) "
'''''    l_strSQL = l_strSQL & " OR   ('" & Format(v_datEndDate, "mm/dd/yyyy") & "' BETWEEN B.EFF_START_DATE AND B.EFF_END_DATE)) "
'''''    If bln_PrimaryOnly Then
'''''        l_strSQL = l_strSQL & " AND B.PRIMARY_IND = 'Y' "
'''''    End If
'''''    l_strSQL = l_strSQL & " UNION "
'''''    l_strSQL = l_strSQL & "SELECT B.WG_NUM, A.WG_DESC, B.EFF_START_DATE, B.EFF_END_DATE, "
'''''    l_strSQL = l_strSQL & "A.PROJECT_ID, B.PRIMARY_IND, A.WG_SUPERVISOR, C.LOC_CD, C.DEPT_NUM, "
'''''    l_strSQL = l_strSQL & "date('01/01/0100') as proj_start_date, date('01/01/0100') as proj_end_date "
'''''    l_strSQL = l_strSQL & "FROM METHDBA.TTR_WORK_GROUP A, "
'''''    l_strSQL = l_strSQL & "METHDBA.TTR_EMP_WG B, "
'''''    l_strSQL = l_strSQL & "METHDBA.TTR_ASSOCIATIONS C "
'''''    l_strSQL = l_strSQL & "WHERE A.WG_NUM = B.WG_NUM AND A.WG_NUM = C.WG_NUM "
'''''    l_strSQL = l_strSQL & "AND B.EMP_NUM = " & v_lngEmpNum
'''''    l_strSQL = l_strSQL & " and A.project_id = 0"
'''''    l_strSQL = l_strSQL & " AND (('" & Format(v_datStartDate, "mm/dd/yyyy") & "' BETWEEN B.EFF_START_DATE AND B.EFF_END_DATE) "
'''''    l_strSQL = l_strSQL & " OR   ('" & Format(v_datEndDate, "mm/dd/yyyy") & "' BETWEEN B.EFF_START_DATE AND B.EFF_END_DATE)) "
'''''    If bln_PrimaryOnly Then
'''''        l_strSQL = l_strSQL & " AND B.PRIMARY_IND = 'Y' "
'''''    End If
'''''    l_strSQL = l_strSQL & " ORDER BY WG_NUM "
'''''
'''''    l_objrs.CursorLocation = adUseClient
'''''    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'''''
'''''    l_objrs.ActiveConnection = Nothing
'''''    l_objCN.Close
'''''
'''''    Set l_objCN = Nothing
'''''
'''''    Set GetEmployeeWorkGroups = l_objrs
'''''
'''''    Set l_objrs = Nothing
'''''    Exit Function
'''''
'''''GetEmployeeWorkGroups_EH:
'''''    Set l_objCM = Nothing
'''''    Set l_objCN = Nothing
'''''    Set l_objrs = Nothing
'''''
'''''    Select Case Err.Number
'''''        Case trErrdbInvalidZerolngEmpNum
'''''            TrsRaiseError m_strCMODULENAME, "GetEmployeeWorkGroups", trErrdbInvalidZerolngEmpNum, , , , True
'''''        Case Else
'''''            TrsRaiseError m_strCMODULENAME, "GetEmployeeWorkGroups", , , , DescArgsFrom(l_strSQL), True
'''''    End Select
'''''
'''''End Function


Public Function GetCrewNumber(ByVal v_lngWgNum As Long) As String
    
On Error GoTo GetCrewNumber_EH
    If v_lngWgNum < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    End If

    GetCrewNumber = CrewNumber(v_lngWgNum)
  
    Exit Function
    
GetCrewNumber_EH:
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetCrewNumber", Err.Number, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetCrewNumber", , , , , True
    End Select

End Function

Public Function GetDOJMDistrictAndHeadquarters(ByVal v_strCrewNum As String) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetDOJMDistrictAndHeadquarters_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT2")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    
    l_strSQL = "SELECT B.CDE_DIST, B.CDE_HDQTR_CRW " _
             & "FROM METHDBA.TOMCREW A, METHDBA.TOMCRWHDQTR B " _
             & "WHERE A.CDE_HDQTR_CRW = B.CDE_HDQTR_CRW " _
             & "AND A.CDE_CRW = '" & v_strCrewNum & "'"
   
   
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDOJMDistrictAndHeadquarters = l_objRS
    
    Set l_objRS = Nothing
    Exit Function

GetDOJMDistrictAndHeadquarters_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDOJMDistrictAndHeadquarters", , , , DescArgsFrom(l_strSQL), True

End Function

'Public Function IsValidWorkgroup(ByVal v_lngWorkGroupNum As Long, ByVal v_datEffDate As Date) As Boolean
''  Validate Workgroup
'
'    Dim l_strSQL As String
'    Dim l_objCN As ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'
'On Error GoTo IsValidWorkGroup_EH
'
'    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
'    Set l_objCN = New ADODB.Connection
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open
'
'    l_strSQL = "SELECT WG_NUM FROM METHDBA.TTR_WORK_GROUP "
'    l_strSQL = l_strSQL & "WHERE WG_NUM = " & v_lngWorkGroupNum & " "
'    l_strSQL = l_strSQL & "AND ('" & Format(v_datEffDate, "mm/dd/yyyy") & "' BETWEEN EFF_START_DATE AND EFF_END_DATE) "
'
'    Set l_objRS = New ADODB.Recordset
'
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'
'    If l_objRS.BOF And l_objRS.EOF Then
'        IsValidWorkgroup = False
'    Else
'        IsValidWorkgroup = True
'    End If
'
'    Set l_objRS = Nothing
'    Exit Function
'
'IsValidWorkGroup_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "IsValidWorkGroup", trErrdbDataService, , , DescArgsFrom(l_strSQL), True
'
'End Function

Public Function GetEmployeeWorkGroups( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional bln_PrimaryOnly As Boolean = False _
            ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_lngStart As Long
    Dim l_strEmployeeSelection As String
    
    On Error GoTo eh
    
    If IsArray(v_aryEmployees) Then
        l_lngStart = LBound(v_aryEmployees)
        If Val(v_aryEmployees(l_lngStart)) = 0 Then
            Err.Raise trErrdbInvalidZerolngEmpNum
        Else
            l_strEmployeeSelection = BuildInClause("B.EMP_NUM", v_aryEmployees, False)
        End If
    Else
        If Val(v_aryEmployees) = 0 Then
            Err.Raise trErrdbInvalidZerolngEmpNum
        Else
            l_strEmployeeSelection = "B.EMP_NUM = " & str(v_aryEmployees)
        End If
    End If
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
'The first select joins on ttr_projects to pick up the eff_start_date and
' eff_end_date for PPM project, but this will only yield a resulting row when the project_id
' is not equal to zero.  The second select will yield a row for non-PPM workgroup,
' or when the project_id = 0.
    
    l_strSQL = "SELECT B.EMP_NUM, B.WG_NUM, A.WG_DESC, B.EFF_START_DATE, B.EFF_END_DATE, "
    l_strSQL = l_strSQL & "A.PROJECT_ID, B.PRIMARY_IND, A.WG_SUPERVISOR, C.LOC_CD, C.DEPT_NUM, "
    l_strSQL = l_strSQL & "d.eff_start_date as proj_start_date, d.eff_end_date as proj_end_date "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_WORK_GROUP A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_EMP_WG B, "
    l_strSQL = l_strSQL & "METHDBA.TTR_ASSOCIATIONS C, "
    l_strSQL = l_strSQL & "METHDBA.ttr_projects d "
    l_strSQL = l_strSQL & "WHERE A.WG_NUM = B.WG_NUM AND A.WG_NUM = C.WG_NUM "
    l_strSQL = l_strSQL & " AND " & l_strEmployeeSelection
    l_strSQL = l_strSQL & " and a.project_id = d.project_id"
    l_strSQL = l_strSQL & " and d.ppm_ind = 'Y'"
    l_strSQL = l_strSQL & " AND B.EFF_START_DATE <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & " AND B.EFF_END_DATE >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
    If bln_PrimaryOnly Then
        l_strSQL = l_strSQL & " AND B.PRIMARY_IND = 'Y' "
    End If
    l_strSQL = l_strSQL & " UNION "
    l_strSQL = l_strSQL & "SELECT B.EMP_NUM, B.WG_NUM, A.WG_DESC, B.EFF_START_DATE, B.EFF_END_DATE, "
    l_strSQL = l_strSQL & "A.PROJECT_ID, B.PRIMARY_IND, A.WG_SUPERVISOR, C.LOC_CD, C.DEPT_NUM, "
    l_strSQL = l_strSQL & "date('01/01/0100') as proj_start_date, date('01/01/0100') as proj_end_date "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_WORK_GROUP A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_EMP_WG B, "
    l_strSQL = l_strSQL & "METHDBA.TTR_ASSOCIATIONS C "
    l_strSQL = l_strSQL & "WHERE A.WG_NUM = B.WG_NUM AND A.WG_NUM = C.WG_NUM "
    l_strSQL = l_strSQL & " AND " & l_strEmployeeSelection
    l_strSQL = l_strSQL & " and A.project_id = 0"
    l_strSQL = l_strSQL & " AND B.EFF_START_DATE <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & " AND B.EFF_END_DATE >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
    If bln_PrimaryOnly Then
        l_strSQL = l_strSQL & " AND B.PRIMARY_IND = 'Y' "
    End If
    l_strSQL = l_strSQL & " ORDER BY EMP_NUM, WG_NUM "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeWorkGroups = l_objRS
    
    Set l_objRS = Nothing
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetEmployeeWorkGroups", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetEmployeeWorkGroups2", , , , DescArgsFrom(l_strSQL), True
    End Select
    
End Function


Public Function GetWorkgroupRecord( _
            ByVal v_lngWgNum As Long) As ADODB.Recordset
' backward compatibility, remove when bincompat broken
On Error GoTo eh

    Set GetWorkgroupRecord = GetWorkgroupRecords(v_lngWgNum)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetWorkgroupRecord", , True, , , True

End Function

Public Function GetWorkgroupRecords( _
            ByVal v_lngWgNums As Variant) As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strWhere As String
    
On Error GoTo eh
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = _
        l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_lngWgNums) Then
        l_strWhere = "WHERE " & BuildInClause("A.WG_NUM", v_lngWgNums, False)
    Else
        l_strWhere = "WHERE A.WG_NUM = " & CStr(v_lngWgNums)
    End If

'The first select joins on ttr_projects to pick up the eff_start_date and
' eff_end_date for PPM project, but this will only yield a resulting row when the project_id
' is not equal to zero.  The second select will yield a row for non-PPM workgroup,
' or when the project_id = 0.
    
    l_strSQL = "SELECT A.WG_NUM, A.WG_DESC, A.EFF_START_DATE, A.EFF_END_DATE, "
    l_strSQL = l_strSQL & "A.PROJECT_ID, A.DEMAND_IND, A.WG_SUPERVISOR, "
    l_strSQL = l_strSQL & "A.CREW_NUM, B.DEPT_NUM, B.LOC_CD, "
    l_strSQL = l_strSQL & "C.EFF_START_DATE AS PROJ_START_DATE, "
    l_strSQL = l_strSQL & "C.EFF_END_DATE AS PROJ_END_DATE "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_WORK_GROUP A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_ASSOCIATIONS B, "
    l_strSQL = l_strSQL & "METHDBA.TTR_PROJECTS C "
    l_strSQL = l_strSQL & l_strWhere
    l_strSQL = l_strSQL & " AND A.WG_NUM = B.WG_NUM"
    l_strSQL = l_strSQL & " AND A.PROJECT_ID = C.PROJECT_ID"
    l_strSQL = l_strSQL & " AND C.PPM_IND = 'Y'"
    l_strSQL = l_strSQL & " UNION "
    l_strSQL = l_strSQL & "SELECT A.WG_NUM, A.WG_DESC, A.EFF_START_DATE, A.EFF_END_DATE, "
    l_strSQL = l_strSQL & "A.PROJECT_ID, A.DEMAND_IND, A.WG_SUPERVISOR, "
    l_strSQL = l_strSQL & "A.CREW_NUM, B.DEPT_NUM, B.LOC_CD, "
    l_strSQL = l_strSQL & "DATE('01/01/0100') AS PROJ_START_DATE, "
    l_strSQL = l_strSQL & "DATE('01/01/0100') AS PROJ_END_DATE "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_WORK_GROUP A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_ASSOCIATIONS B "
    l_strSQL = l_strSQL & l_strWhere
    l_strSQL = l_strSQL & " AND A.WG_NUM = B.WG_NUM"
    l_strSQL = l_strSQL & " AND A.PROJECT_ID = 0"
    If IsArray(v_lngWgNums) Then
        l_strSQL = l_strSQL & " ORDER BY WG_NUM"
    End If
        
    Set l_objRS = New ADODB.Recordset
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    
    Set GetWorkgroupRecords = l_objRS
    Set l_objRS = Nothing
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkgroupRecords", , , , DescArgsFrom(l_strSQL), True

End Function



'*********************************************************
' Begin Public getters
'*********************************************************

Public Function EffectiveStartDate( _
            ByVal v_lngWgNum As Long) As Variant
   
On Error GoTo eh
    EffectiveStartDate = GetAttribute(v_lngWgNum, CWG_EFF_START_DATE)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "EffectiveStartDate", , , , , True
End Function

Public Function EffectiveEndDate( _
            ByVal v_lngWgNum As Long) As Variant
    
On Error GoTo eh
    EffectiveEndDate = GetAttribute(v_lngWgNum, CWG_EFF_END_DATE)
    Exit Function
   
eh:
    TrsRaiseError m_strCMODULENAME, "EffectiveEndDate", , , , , True
End Function

Public Function ProjectStartDate( _
            ByVal v_lngWgNum As Long) As Variant
    
On Error GoTo eh
    ProjectStartDate = GetAttribute(v_lngWgNum, CWG_PROJ_START_DATE)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ProjectStartDate", , , , , True
                        
End Function

Public Function ProjectEndDate( _
            ByVal v_lngWgNum As Long) As Variant
    
On Error GoTo eh
    ProjectEndDate = GetAttribute(v_lngWgNum, CWG_PROJ_END_DATE)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ProjectEndDate", , , , , True
                        
End Function

Public Function ProjectID( _
            ByVal v_lngWgNum As Long) As Long
    
On Error GoTo eh
    ProjectID = GetAttribute(v_lngWgNum, CWG_PROJECT_ID, vbLong)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ProjectID", , , , , True
End Function

Public Function Supervisor( _
            ByVal v_lngWgNum As Long) As Long
    
On Error GoTo eh
    Supervisor = GetAttribute(v_lngWgNum, CWG_WG_SUPERVISOR, vbLong)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "Supervisor", , , , , True
End Function

Public Function DemandIndicator( _
            ByVal v_lngWgNum As Long) As String
    
On Error GoTo eh
    DemandIndicator = GetAttribute(v_lngWgNum, CWG_DEMAND_IND, vbString)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DemandIndicator", , , , , True
                        
End Function

Public Function WorkgroupDescription( _
            ByVal v_lngWgNum As Long) As String
   
On Error GoTo eh
    WorkgroupDescription = GetAttribute(v_lngWgNum, CWG_WG_DESC, vbString)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "WorkgroupDescription", , , , , True
                        
End Function

Public Function CrewNumber( _
            ByVal v_lngWgNum As Long) As String
    
On Error GoTo eh
    CrewNumber = GetAttribute(v_lngWgNum, CWG_CREW_NUM, vbString)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CrewNumber", , , , , True
                        
End Function

Public Function DOJMDistrict( _
            ByVal v_lngWgNum As Long) As String
    
On Error GoTo eh
    DOJMDistrict = GetAttribute(v_lngWgNum, CWG_DOJM_DIST, vbString)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DOJMDistrict", , , , , True
                        
End Function

Public Function DOJMHeadquarters( _
            ByVal v_lngWgNum As Long) As String
    
On Error GoTo eh
    DOJMHeadquarters = GetAttribute(v_lngWgNum, CWG_DOJM_HQ, vbString)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DOJMHeadquarters", , , , , True
                        
End Function

Public Function DepartmentNumber( _
            ByVal v_lngWgNum As Long) As String
    
On Error GoTo eh
    DepartmentNumber = GetAttribute(v_lngWgNum, CWG_DEPT_NUM, vbString)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DepartmentNumber", , , , , True
                        
End Function


Public Function LocationCode(ByVal v_lngWgNum As Long) As Long
    
On Error GoTo eh
    LocationCode = GetAttribute(v_lngWgNum, CWG_LOC_CD, vbLong)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "LocationCode", , , , , True
End Function

Public Function IsValidWorkgroup( _
            ByVal v_lngWgNum As Long, _
            ByVal v_datEffDate As Date) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(FormatCacheKey(v_lngWgNum), CWG_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    
    IsValidWorkgroup = _
        (v_datEffDate >= l_varRecord(CWG_EFF_START_DATE) And _
        v_datEffDate <= l_varRecord(CWG_EFF_END_DATE))
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidWorkgroup", , , , , True
End Function

Public Function Record(ByVal v_lngWgNum As Long) As Variant
' Return a 2 dim variant array where the first
' dim has the field names and the second dim has the values.
' Basically, a lightweight, readonly dictionary that is
' passed by value.  If record is not found, return Empty.
' Already dependent on the actual field names, so just set
' them for now.
    Dim l_varRecord As Variant
    Dim l_varRecordOut(1, CWG_LAST_SLOT - 1) As Variant
    Dim l_strFieldNames As Variant
    Dim l_intIndex As Integer
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(FormatCacheKey(v_lngWgNum), CWG_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    
' Set the fields names and values.
    l_strFieldNames = FieldNames
    For l_intIndex = 1 To CWG_LAST_SLOT
        l_varRecordOut(0, l_intIndex - 1) = l_strFieldNames(l_intIndex - 1)
        
        If l_intIndex = CWG_EFF_START_DATE Or _
            l_intIndex = CWG_EFF_END_DATE Or _
            l_intIndex = CWG_PROJ_START_DATE Or _
            l_intIndex = CWG_PROJ_END_DATE Then
            l_varRecordOut(1, l_intIndex - 1) = Format(l_varRecord(l_intIndex), "m/d/yyyy")
        Else
            l_varRecordOut(1, l_intIndex - 1) = l_varRecord(l_intIndex)
        End If
    Next
        
    Record = l_varRecordOut
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "Record", , , , , True
                        
End Function

Public Function Item( _
            ByVal v_strFieldName As String, _
            ByVal v_varRecord As Variant) As Variant
            
On Error GoTo eh
    Item = RecordItem(v_strFieldName, v_varRecord)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "Item", , , , , True
End Function


'*********************************************************
' Begin Private procedures
'*********************************************************

Private Function GetAttribute( _
            ByVal v_lngWgNum As Long, _
            ByVal v_strAttributeName As String, _
            Optional ByVal v_lngInitVarType As Long = vbEmpty _
            ) As Variant
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = GetRecord(FormatCacheKey(v_lngWgNum), CWG_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        Select Case v_lngInitVarType
            Case vbLong
                GetAttribute = CLng(0)
            Case vbString
                GetAttribute = ""
            Case Else
                GetAttribute = Empty
        End Select
        Exit Function
    End If
    GetAttribute = l_varRecord(v_strAttributeName)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetAttribute"
End Function

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
        Case CWG_GROUPNAME
            Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
            
            Set l_objCM = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
            If l_objCM.AreDSNsEqual Then
                Set l_rsRecords = GetWorkgroupCacheRecords(v_varKeys)
            Else '//-----retrieve the DOJM info using the 2nd dsn (ShadowDirect2)
                Set l_rsRecords = GetWorkgroupCacheRecords2(v_varKeys)
            End If
            Set l_objCM = Nothing
            
            r_varKeyArray = Array("WG_NUM")
            r_varColumnArray = Empty
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

'Public Function GetWorkgroupCacheRecordsTest(Optional ByVal v_varWGNum As Variant) As String
'    Dim l_strSQL As String
'    Dim l_strWhereClause As String
'    Dim l_blnUseDOJM As Boolean
'    Dim i As Long
'On Error GoTo eh
'
'    l_blnUseDOJM = True
'    If IsMissing(v_varWGNum) Or IsEmpty(v_varWGNum) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varWGNum) Then
'        l_strWhereClause = " WHERE " & BuildInClause("WG_NUM", v_varWGNum, False)
'        l_blnUseDOJM = False
'        For i = LBound(v_varWGNum) To UBound(v_varWGNum)
'            If IsNumeric(Trim(v_varWGNum(i))) Then
'                If IsDOJMWorkgroup(CLng(Trim(v_varWGNum(i)))) Then
'                    l_blnUseDOJM = True
'                    Exit For
'                End If
'            End If
'        Next i
'    Else
'        l_strWhereClause = " WHERE WG_NUM = " & Trim(v_varWGNum)
'        If IsNumeric(Trim(v_varWGNum)) Then
'            l_blnUseDOJM = IsDOJMWorkgroup(CLng(Trim(v_varWGNum)))
'        End If
'    End If
'
'    If l_blnUseDOJM Then
'        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetWorkgroupCacheRecordsDOJM.sql")
'    Else
'        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetWorkgroupCacheRecordsNonDOJM.sql")
'    End If
'
'    '//substitute the Where clause (which could be for multiple employees)
'    GetWorkgroupCacheRecordsTest = Replace(l_strSQL, "{0}", l_strWhereClause)
'Exit Function
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetWorkgroupCacheRecordsTest", , , , DescArgsFrom(l_strSQL), True
'End Function

Private Function GetWorkgroupCacheRecords(Optional ByVal v_varWGNum As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    Dim l_blnUseDOJM As Boolean
    Dim i As Long
On Error GoTo eh
    
    l_blnUseDOJM = True
    If IsMissing(v_varWGNum) Or IsEmpty(v_varWGNum) Then
        l_strWhereClause = ""
    ElseIf IsArray(v_varWGNum) Then
        l_strWhereClause = " WHERE " & BuildInClause("WG_NUM", v_varWGNum, False)
        l_blnUseDOJM = False
        For i = LBound(v_varWGNum) To UBound(v_varWGNum)
            If IsNumeric(Trim(v_varWGNum(i))) Then
                If IsDOJMWorkgroup(CLng(Trim(v_varWGNum(i)))) Then
                    l_blnUseDOJM = True
                    Exit For
                End If
            End If
        Next i
    Else
        l_strWhereClause = " WHERE WG_NUM = " & Trim(v_varWGNum)
        If IsNumeric(Trim(v_varWGNum)) Then
            l_blnUseDOJM = IsDOJMWorkgroup(CLng(Trim(v_varWGNum)))
        End If
    End If
            
    If l_blnUseDOJM Then
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetWorkgroupCacheRecordsDOJM.sql")
    Else
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetWorkgroupCacheRecordsNonDOJM.sql")
    End If
    
    '//substitute the Where clause (which could be for multiple employees)
    l_strSQL = Replace(l_strSQL, "{0}", l_strWhereClause)
    
'        l_strSQL = ""
'        l_strSQL = l_strSQL & "Select WG.WG_NUM, WG.WG_DESC, WG.EFF_START_DATE, WG.EFF_END_DATE,"
'        l_strSQL = l_strSQL & " WG.PROJECT_ID, WG.DEMAND_IND, WG.WG_SUPERVISOR, WG.CREW_NUM, "
'        l_strSQL = l_strSQL & " COALESCE(B.DEPT_NUM, SPACE(3)) AS DEPT_NUM, COALESCE(B.LOC_CD, 0) AS LOC_CD,"
'        l_strSQL = l_strSQL & " COALESCE(C.EFF_START_DATE, '01/01/0100') AS PROJ_START_DATE, COALESCE(C.EFF_END_DATE, '01/01/0100') AS PROJ_END_DATE,"
'        l_strSQL = l_strSQL & " COALESCE(DH1.CDE_DIST, DH2.CDE_DIST, SPACE(4)) as CDE_DIST,"
'        l_strSQL = l_strSQL & " COALESCE(DH1.CDE_HDQTR_CRW, DH2.CDE_HDQTR_CRW, Space(3)) As CDE_HDQTR_CRW,"
'        l_strSQL = l_strSQL & " IsDOJM, COALESCE(C.PPM_IND, SPACE(1)) AS PPM_IND, WG.CREW_NUM2, COALESCE(DH1.CDE_CRW, DH2.CDE_CRW) AS CDE_CRW"
'        l_strSQL = l_strSQL & " From"
'        l_strSQL = l_strSQL & " (Select WG_NUM, WG_DESC, EFF_START_DATE, EFF_END_DATE,"
'        l_strSQL = l_strSQL & "  PROJECT_ID, DEMAND_IND, WG_SUPERVISOR, CREW_NUM, "
'        l_strSQL = l_strSQL & "   case when wg_num >= 100 and wg_num < 2899 then"
'        l_strSQL = l_strSQL & "     case when LTRIM(RTRIM(CREW_NUM))<>'' then LTRIM(RTRIM(CREW_NUM))"
'        l_strSQL = l_strSQL & "         else RTRIM(char(wg_num)) end"
'        l_strSQL = l_strSQL & "   else LTRIM(RTRIM(CREW_NUM)) end as CREW_NUM2,"
'        l_strSQL = l_strSQL & "  case when wg_num >= 100 and wg_num < 2899 then '1'"
'        l_strSQL = l_strSQL & "   Else '0' end as IsDOJM"
'        l_strSQL = l_strSQL & " From METHDBA.TTR_WORK_GROUP"
'        l_strSQL = l_strSQL & l_strWhereClause & " ) WG"
'        l_strSQL = l_strSQL & " LEFT OUTER JOIN"
'        l_strSQL = l_strSQL & " (SELECT B.CDE_DIST, LEFT(B.CDE_HDQTR_CRW,3) AS CDE_HDQTR_CRW, A.CDE_CRW"
'        l_strSQL = l_strSQL & "    FROM METHDBA.TOMCREW A,"
'        l_strSQL = l_strSQL & "    METHDBA.TOMCRWHDQTR B"
'        l_strSQL = l_strSQL & "    WHERE A.CDE_HDQTR_CRW = B.CDE_HDQTR_CRW) DH1"
'        l_strSQL = l_strSQL & " on IsDOJM = '1'"
'        l_strSQL = l_strSQL & " and DH1.CDE_CRW = 'C' || RIGHT('00000' || WG.CREW_NUM2, 5)"
'        l_strSQL = l_strSQL & " LEFT OUTER JOIN"
'        l_strSQL = l_strSQL & " (SELECT B.CDE_DIST, LEFT(B.CDE_HDQTR_CRW,3) AS CDE_HDQTR_CRW, A.CDE_CRW"
'        l_strSQL = l_strSQL & "    FROM METHDBA.TOMCREW A,"
'        l_strSQL = l_strSQL & "    METHDBA.TOMCRWHDQTR B"
'        l_strSQL = l_strSQL & "    WHERE A.CDE_HDQTR_CRW = B.CDE_HDQTR_CRW) DH2"
'        l_strSQL = l_strSQL & " on IsDOJM = '1'"
'        l_strSQL = l_strSQL & " and DH2.CDE_CRW = 'U' || RIGHT('00000' || WG.CREW_NUM2, 5)"
'        l_strSQL = l_strSQL & " INNER JOIN"
'        l_strSQL = l_strSQL & "    METHDBA.TTR_ASSOCIATIONS B"
'        l_strSQL = l_strSQL & "       ON WG.WG_NUM = B.WG_NUM"
'        l_strSQL = l_strSQL & " LEFT OUTER JOIN"
'        l_strSQL = l_strSQL & "    METHDBA.TTR_PROJECTS C"
'        l_strSQL = l_strSQL & "       ON WG.PROJECT_ID = C.PROJECT_ID"

    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetWorkgroupCacheRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkgroupCacheRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'//-----for LUW
Private Function GetWorkgroupCacheRecords2(Optional ByVal v_varWGNum As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset, l_rsDOJM As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String, l_strCrewNum As String
    
On Error GoTo eh
    
    If IsMissing(v_varWGNum) Or IsEmpty(v_varWGNum) Then
        l_strWhereClause = ""
    ElseIf IsArray(v_varWGNum) Then
        l_strWhereClause = " WHERE " & BuildInClause("WG_NUM", v_varWGNum, False)
    Else
        l_strWhereClause = " WHERE WG_NUM = " & Trim(v_varWGNum)
    End If
        
    l_strSQL = ""
    l_strSQL = l_strSQL & "Select WG.WG_NUM, WG.WG_DESC, WG.EFF_START_DATE, WG.EFF_END_DATE,"
    l_strSQL = l_strSQL & " WG.PROJECT_ID, WG.DEMAND_IND, WG.WG_SUPERVISOR, WG.CREW_NUM, "
    l_strSQL = l_strSQL & " COALESCE(B.DEPT_NUM, SPACE(3)) AS DEPT_NUM, COALESCE(B.LOC_CD, 0) AS LOC_CD,"
    l_strSQL = l_strSQL & " COALESCE(C.EFF_START_DATE, '01/01/0100') AS PROJ_START_DATE, COALESCE(C.EFF_END_DATE, '01/01/0100') AS PROJ_END_DATE,"
    l_strSQL = l_strSQL & " SPACE(4) as CDE_DIST,"
    l_strSQL = l_strSQL & " Space(3) As CDE_HDQTR_CRW,"
    l_strSQL = l_strSQL & " IsDOJM, COALESCE(C.PPM_IND, SPACE(1)) AS PPM_IND, WG.CREW_NUM2, SPACE(6) AS CDE_CRW"
    l_strSQL = l_strSQL & " From"
    l_strSQL = l_strSQL & " (Select WG_NUM, WG_DESC, EFF_START_DATE, EFF_END_DATE,"
    l_strSQL = l_strSQL & "  PROJECT_ID, DEMAND_IND, WG_SUPERVISOR, CREW_NUM, "
    l_strSQL = l_strSQL & "   case when wg_num >= 100 and wg_num < 2899 then"
    l_strSQL = l_strSQL & "     case when LTRIM(RTRIM(CREW_NUM))<>'' then LTRIM(RTRIM(CREW_NUM))"
    l_strSQL = l_strSQL & "         else RTRIM(char(wg_num)) end"
    l_strSQL = l_strSQL & "   else LTRIM(RTRIM(CREW_NUM)) end as CREW_NUM2,"
    l_strSQL = l_strSQL & "  case when wg_num >= 100 and wg_num < 2899 then '1'"
    l_strSQL = l_strSQL & "   Else '0' end as IsDOJM"
    l_strSQL = l_strSQL & " From METHDBA.TTR_WORK_GROUP"
    l_strSQL = l_strSQL & l_strWhereClause & " ) WG"
'    l_strSQL = l_strSQL & " LEFT OUTER JOIN"
'    l_strSQL = l_strSQL & " (SELECT B.CDE_DIST, LEFT(B.CDE_HDQTR_CRW,3) AS CDE_HDQTR_CRW, A.CDE_CRW"
'    l_strSQL = l_strSQL & " FROM METHDBA.TOMCREW A,"
'    l_strSQL = l_strSQL & " METHDBA.TOMCRWHDQTR B"
'    l_strSQL = l_strSQL & " WHERE A.CDE_HDQTR_CRW = B.CDE_HDQTR_CRW) CH"
'    l_strSQL = l_strSQL & " on IsDOJM = '1'"
'    l_strSQL = l_strSQL & " and (CH.CDE_CRW = 'C' || RIGHT('00000' || WG.CREW_NUM2, 5)"
'    l_strSQL = l_strSQL & " OR CH.CDE_CRW = 'U' || RIGHT('00000' || WG.CREW_NUM2, 5))"
    l_strSQL = l_strSQL & " INNER JOIN"
    l_strSQL = l_strSQL & "    METHDBA.TTR_ASSOCIATIONS B"
    l_strSQL = l_strSQL & "       ON WG.WG_NUM = B.WG_NUM"
    l_strSQL = l_strSQL & " LEFT OUTER JOIN"
    l_strSQL = l_strSQL & "    METHDBA.TTR_PROJECTS C"
    l_strSQL = l_strSQL & "       ON WG.PROJECT_ID = C.PROJECT_ID"

    Set l_objRS = GetCacheRecords(l_strSQL, adLockBatchOptimistic)
    
    '//-----populate the DOJM info using the 2nd datasource (ShadowDirect2)
    If RecordsExist(l_objRS) Then
        l_objRS.MoveFirst
        Do While Not l_objRS.EOF '//-----should be only 1 row
            If l_objRS!IsDOJM = "1" Then
                l_strCrewNum = "C" & Right("00000" & Trim(l_objRS!CREW_NUM2), 5)
                Set l_rsDOJM = GetDOJMDistrictAndHeadquarters(l_strCrewNum)
                If Not RecordsExist(l_rsDOJM) Then
                    l_strCrewNum = "U" & Right("00000" & Trim(l_objRS!CREW_NUM2), 5)
                    Set l_rsDOJM = GetDOJMDistrictAndHeadquarters(l_strCrewNum)
                End If

                If RecordsExist(l_rsDOJM) Then
                    l_rsDOJM.MoveFirst
                    l_objRS!CDE_DIST = l_rsDOJM!CDE_DIST
                    l_objRS!CDE_HDQTR_CRW = Left(l_rsDOJM!CDE_HDQTR_CRW, 3)
                    l_objRS!CDE_CRW = l_strCrewNum
                    l_objRS.Update
                End If
            End If
            l_objRS.MoveNext
        Loop
    End If
    
    Set GetWorkgroupCacheRecords2 = l_objRS
    
    Set l_objRS = Nothing
    Set l_rsDOJM = Nothing
        
    Exit Function
eh:
    Set l_objRS = Nothing
    Set l_rsDOJM = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkgroupCacheRecords2", , , , DescArgsFrom(l_strSQL), True
End Function

Private Function FormatCacheKey(ByVal v_lngWgNum As Long) As String
    Dim l_lngLength As Long
    Dim l_strInteger As String
    
    l_lngLength = 11
    
    l_strInteger = String(l_lngLength, "0")  '//-----PadLeft with zeros
    l_strInteger = l_strInteger & Abs(Format(v_lngWgNum))
    FormatCacheKey = IIf(v_lngWgNum < 0, "-", "0") & Right(l_strInteger, l_lngLength - 1)
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
           
On Error GoTo eh

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
    Exit Function
eh:
    Set recCache = Nothing

    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & CStr(v_varKey)
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    
    TrsRaiseError m_strCMODULENAME, "GetRecord", , , , l_colErrArgs, True
End Function

Public Sub InvalidateWorkGroupCache(ByVal v_lngWgNum As Long)

On Error GoTo eh
   
    InvalidateCache FormatCacheKey(v_lngWgNum), CWG_GROUPNAME
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateWorkGroupCache"
End Sub

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

'//-----used by RecordCache-----//

Private Function IsLazyInitialized(ByVal v_strGroupName As String) As Boolean
' If group is not lazy initialized, then it is bulk initialized.
On Error GoTo eh
    Select Case Trim(v_strGroupName)
    Case CWG_GROUPNAME
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
    Case CWG_GROUPNAME
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
'    Select Case Trim(v_strGroupName)
'    Case CUC_GROUPNAME
'        OneRowForEachKey = False
'    Case Else
        OneRowForEachKey = True
'    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OneRowForEachKey"
End Function

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


'Private Function GetRecord( _
'            ByVal v_strKey As String, _
'            ByVal v_strGroupName As String, _
'            ByRef r_blnExists As Boolean) As Variant
'' Answer an variant array, representing a record, or empty.
'' If the record is not found, using lazy, it should
'' add a value to the cache to denote invalid so that it does not
'' attempt to retrieve it from the database again.  Currently, a string
'' constant is added for record not found and the spmCache component
'' will answer that it exists, but the property value is a
'' RECORDNOTFOUND string.  Answer Empty when RECORDNOTFOUND.
''
'    Dim l_varRecord As Variant
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim lblnExists As Boolean
'    Dim lblnGroupExists As Boolean
'
'On Error GoTo eh
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    l_varRecord = _
'        spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
'    If Not r_blnExists Then
'        LoadWorkgroupRecord v_strKey
'        l_varRecord = _
'            spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
'    End If
'    ' test for array first since 99.9% of time it should be one,
'    ' otherwise answer empty.
'    If IsArray(l_varRecord) Then
'        GetRecord = l_varRecord
'    Else
'        GetRecord = Empty
'    End If
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetRecord"
'End Function

'Private Sub LoadWorkgroupRecord( _
'                ByVal v_strWgNum As String)
'    Dim rs As ADODB.Recordset
'    Dim lvarRecord As Variant
'    Dim spmWrapper As utilSPMCache.cuSPMCache
'    Dim llngRetVal As Long
'
'On Error GoTo eh
'    lvarRecord = GetWorkgroupCacheRecord(CLng(v_strWgNum))
'    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
'    If IsEmpty(lvarRecord) Then
'        llngRetVal = _
'            spmWrapper.Add(v_strWgNum, CRECORDNOTFOUND, CWG_GROUPNAME)
'    Else
'        llngRetVal = _
'            spmWrapper.Add(v_strWgNum, lvarRecord, CWG_GROUPNAME)
'    End If
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    Exit Sub
'
'eh:
'    Set spmWrapper = Nothing
'    Set rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "LoadWorkgroupRecord"
'End Sub

Private Function FieldNames() As Variant
' Order is important here.  The field names must be in the same
' order as the index constants.
    If IsEmpty(maryFieldNames) Then
        maryFieldNames = Array( _
            "WG_NUM", "WG_DESC", "EFF_START_DATE", "EFF_END_DATE", _
            "PROJECT_ID", "DEMAND_IND", "WG_SUPERVISOR", "CREW_NUM", _
            "DEPT_NUM", "LOC_CD", "PROJ_START_DATE", "PROJ_END_DATE", _
            "DOJM_DIST", "DOJM_HQ")
    End If
    FieldNames = maryFieldNames
End Function

Private Function RecordItem( _
            ByVal v_strFieldName As String, _
            ByVal v_varRecord As Variant) As Variant
' v_varRecord must be a 2 dim array of variants where the first
' dim has the field names and the second dim has the values.
' Basically, a lightweight, readonly dictionary that is
' passed by value.
    Dim i As Integer
    Dim lstrFieldName As String
    
    If IsEmpty(v_varRecord) Then
        Exit Function
    End If
    
    lstrFieldName = Trim(v_strFieldName)
    For i = LBound(v_varRecord, 2) To UBound(v_varRecord, 2)
        If v_varRecord(LBound(v_varRecord), i) = lstrFieldName Then
            Exit For
        End If
    Next
    
    If i <= UBound(v_varRecord, 2) Then
        RecordItem = v_varRecord(UBound(v_varRecord), i)
    Else
        RecordItem = Empty
    End If

End Function

'Private Function GetWorkgroupCacheRecord( _
'            ByVal v_lngWgNum As Long) As Variant
'' Answer a record or empty.
'    Dim rs As ADODB.Recordset
'    Dim lvarRecord(CWG_LAST_SLOT) As Variant
'    Dim l_strCrewNum As String
'    Dim lcolArgs As Collection
'
'On Error GoTo eh
'    Set rs = GetWorkgroupRecords(v_lngWgNum)
'    If rs.BOF And rs.EOF Then
'        GetWorkgroupCacheRecord = Empty
'        Exit Function
'    End If
'
'    lvarRecord(CWG_WG_NUM) = CStr(rs!WG_NUM)
'    lvarRecord(CWG_WG_DESC) = Trim(rs!WG_DESC)
'    lvarRecord(CWG_EFF_START_DATE) = rs!EFF_START_DATE
'    lvarRecord(CWG_EFF_END_DATE) = rs!EFF_END_DATE
'    lvarRecord(CWG_PROJECT_ID) = Trim(rs!PROJECT_ID)
'    lvarRecord(CWG_DEMAND_IND) = Trim(rs!DEMAND_IND)
'    lvarRecord(CWG_WG_SUPERVISOR) = Trim(rs!WG_SUPERVISOR)
'    lvarRecord(CWG_CREW_NUM) = Trim(rs!CREW_NUM)
'    lvarRecord(CWG_DEPT_NUM) = Trim(rs!DEPT_NUM)
'    lvarRecord(CWG_LOC_CD) = Trim(rs!LOC_CD)
'    lvarRecord(CWG_PROJ_START_DATE) = rs!PROJ_START_DATE
'    lvarRecord(CWG_PROJ_END_DATE) = rs!PROJ_END_DATE
'    lvarRecord(CWG_DOJM_DIST) = ""
'    lvarRecord(CWG_DOJM_HQ) = ""
'    Set rs = Nothing
'
'    If Not IsDOJMWorkgroup(v_lngWgNum) Then
'        GetWorkgroupCacheRecord = lvarRecord
'        Exit Function
'    End If
'
'    l_strCrewNum = lvarRecord(CWG_CREW_NUM)
'    If Trim(l_strCrewNum) = "" Then
'        l_strCrewNum = CStr(v_lngWgNum)
'    End If
'
'    l_strCrewNum = LeftFillCrewNumber(l_strCrewNum)
'    Set rs = GetDOJMDistrictAndHeadquarters("C" & l_strCrewNum)
'    If rs.BOF And rs.EOF Then
'        Set rs = Nothing
'        Set rs = GetDOJMDistrictAndHeadquarters("U" & l_strCrewNum)
'        If rs.BOF And rs.EOF Then
'            Set rs = Nothing
'            GetWorkgroupCacheRecord = lvarRecord
'            Exit Function
'        End If
'    End If
'
'    lvarRecord(CWG_DOJM_DIST) = rs!CDE_DIST
'    lvarRecord(CWG_DOJM_HQ) = Left(rs!CDE_HDQTR_CRW, 3)
'    GetWorkgroupCacheRecord = lvarRecord
'
'    Exit Function
'
'eh:
'    Set rs = Nothing
'    Set lcolArgs = New Collection
'    lcolArgs.Add "v_lngWgNum: " & v_lngWgNum
'    TrsRaiseError m_strCMODULENAME, "GetWorkgroupCacheRecord", , , , lcolArgs
'
'End Function
'
Private Function IsDOJMWorkgroup(ByVal v_lngWg_Num As Long) As Boolean

    IsDOJMWorkgroup = (v_lngWg_Num >= 100 And v_lngWg_Num <= 2899)

End Function

'Private Function LeftFillCrewNumber( _
'            ByVal v_strCrewNum As String) As String
''v_strCrewNum should be a string upto 5 characters.
'
'    LeftFillCrewNumber = Right("00000" & v_strCrewNum, 5)
'
'End Function

Public Function GetActivities(ByVal v_lngWgNum As Long) As ADODB.Recordset
    
'Similar to GetWorkGroupActivities
'This is not used in the context of a transaction
'GetActivities gets called often
'The seperation makes better use of the database resources
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetActivities_EH
    
    If v_lngWgNum < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    Else
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT A.ACTIVITY_CD , A.ACTIVITY_DESC "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACTIVITY A "
    l_strSQL = l_strSQL & " , METHDBA.TTR_ACTIVITY_ASGNM B "
    l_strSQL = l_strSQL & "WHERE WG_NUM = " & v_lngWgNum
    l_strSQL = l_strSQL & "  and A.ACTIVITY_CD = B.ACTIVITY_CD "
    l_strSQL = l_strSQL & " ORDER BY 2 "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetActivities = l_objRS

    Set l_objRS = Nothing
    Exit Function

GetActivities_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetActivities", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetActivities", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function
Public Function GetSteps(ByVal v_sJobTask As String, ByVal v_sDBID As String) As ADODB.Recordset
    
'Similar to GetWorkGroupSteps
'This is not used in the context of a transaction
'GetSteps gets called often
'The seperation makes better use of the database resources
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objCMD As New Command
    Dim l_JobNbr As New ADODB.Parameter
    Dim l_TaskNbr As New ADODB.Parameter
    Dim l_DBCode As New ADODB.Parameter
    Dim l_ReturnValue As New ADODB.Parameter
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_rsGetSteps As New ADODB.Recordset
    
    On Error GoTo GetSteps_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objCMD = New ADODB.Command
    Set l_ReturnValue = l_objCMD.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue)
    Set l_JobNbr = l_objCMD.CreateParameter("@jobnbr", adChar, adParamInput, 8, Mid(v_sJobTask, 1, 8))
    Set l_TaskNbr = l_objCMD.CreateParameter("@tasknbr", adChar, adParamInput, 4, Mid(v_sJobTask, 9, 2))
    Set l_DBCode = l_objCMD.CreateParameter("@dbcode", adChar, adParamInput, 8, v_sDBID)
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("EMPRV")
    'l_objCN.CursorLocation = adUseClient
    'l_objCN.ConnectionString = "DSN=DB2QLUW;UID=db2admin;PWD=db2admin"
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    'l_strSQL = "SELECT A.STEP_NBR , A.STEP_DESC "
    'l_strSQL = l_strSQL & "FROM DB2ADMIN.TJOBTASKSTEP A "
    'l_strSQL = l_strSQL & "WHERE JOBTASK_NBR = '" & v_sJobTask & "'"
    'l_strSQL = l_strSQL & " ORDER BY 2 "
    
    'l_objRS.CursorLocation = adUseClient
    'l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    'l_objRS.ActiveConnection = Nothing
    
    With l_objCMD
        .Parameters.Append l_ReturnValue
        .Parameters.Append l_JobNbr
        .Parameters.Append l_TaskNbr
        .Parameters.Append l_DBCode
        .ActiveConnection = l_objCN
        .CommandType = adCmdStoredProc
        .CommandText = "dbo.po1518_sp_tris_joblist_stepno"
    End With
    
    l_objRS.Open l_objCMD
    
    If l_ReturnValue >= 0 Then
        With l_rsGetSteps
            .Fields.Append "STEP_NBR", adInteger
            .Fields.Append "STEP_DESC", adVarChar, 60
            
            .CursorLocation = adUseClient
            .LockType = adLockBatchOptimistic
            .CursorType = adOpenStatic
            .Open
            .ActiveConnection = Nothing
            
            If Not l_objRS Is Nothing Then
                If Not (l_objRS.BOF And l_objRS.EOF) Then
                    While Not l_objRS.EOF
                        .AddNew
                        !STEP_NBR = l_objRS.Fields("seqnbr").Value
                        !step_desc = Mid(l_objRS.Fields("ldesc").Value, 1, 60)
                        .Update
                        l_objRS.MoveNext
                    Wend
                'Else
                    'Set l_rsGetSteps = Nothing
                End If
            End If
        End With
    End If
    
    Set GetSteps = l_rsGetSteps
        
    l_objCN.Close
    
    Set l_objCN = Nothing
    'Set l_objRS = Nothing
    Exit Function

GetSteps_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objCMD = Nothing
    Set l_objRS = Nothing
    Set l_rsGetSteps = Nothing
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetSteps", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetSteps", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function
Public Function GetEmployeesInWorkGroup( _
            ByVal v_aryWorkGroups As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False _
            ) As ADODB.Recordset
        
    Dim l_lngStart As Long
    Dim l_datCutOffChargeDate As Date
    
On Error GoTo GetEmployeesInWorkGroup_EH
    
    l_lngStart = LBound(v_aryWorkGroups)
    If Val(v_aryWorkGroups(l_lngStart)) < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    Else
    End If
    
    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set GetEmployeesInWorkGroup = GetEmployeesInWorkGroupACTData(v_aryWorkGroups, v_datStartDate, v_datEndDate, TrTable_Actual, v_bolActiveEmployeesOnly)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set GetEmployeesInWorkGroup = GetEmployeesInWorkGroupACTData(v_aryWorkGroups, v_datStartDate, v_datEndDate, TrTable_History, v_bolActiveEmployeesOnly)
        Else
           Set GetEmployeesInWorkGroup = GetEmployeesInWorkGroupACTData(v_aryWorkGroups, v_datStartDate, v_datEndDate, TrTable_ActualAndHistory, v_bolActiveEmployeesOnly) 'use both current and history data
        End If
    End If
    
    Exit Function
    
GetEmployeesInWorkGroup_EH:
        
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroup", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroup", , , , , True
    End Select
    
End Function
Private Function GetEmployeesInWorkGroupACTData( _
            ByVal v_aryWorkGroups As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            ByVal v_TrActualTimeTable As TrActualTimeTable, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False _
            ) As ADODB.Recordset
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClauseWG As String
    Dim l_strInClausePartValueKey As String
    On Error GoTo GetEmployeesInWorkGroupACTData_EH
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
   
    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
    l_strInClauseWG = BuildInClause("WG_NUM", v_aryWorkGroups, False)
    
    l_strSQL = l_strSQL & "select distinct a.emp_num ,b.EMP_NAME ,a.WG_NUM ,b.emp_status "
    Select Case v_TrActualTimeTable
        Case TrTable_Actual
          l_strSQL = l_strSQL & " ,Coalesce(c.Actual_Time_Sw,0) as Actual_Time_Sw "
        Case TrTable_History
          l_strSQL = l_strSQL & " ,Coalesce(d.Actual_Time_Sw,0) as Actual_Time_Sw "
        Case Else
          l_strSQL = l_strSQL & " ,Coalesce(c.Actual_Time_Sw,Coalesce(d.Actual_Time_Sw,0)) as Actual_Time_Sw "
    End Select
    l_strSQL = l_strSQL & "From (select emp_num ,wg_num "
    l_strSQL = l_strSQL & "From methdba.TTR_EMP_WG "
    l_strSQL = l_strSQL & " WHERE " & l_strInClauseWG
    l_strSQL = l_strSQL & " AND EFF_START_DATE <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & " AND EFF_END_DATE >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & " ) as a "
    l_strSQL = l_strSQL & " left outer join "
    l_strSQL = l_strSQL & " (select emp_num ,emp_name ,dept_num ,emp_status "
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EMPLOYEE "
    l_strSQL = l_strSQL & " ) as b "
    l_strSQL = l_strSQL & " on a.emp_num = b.emp_num "
    If v_TrActualTimeTable = TrTable_Actual Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
       l_strSQL = l_strSQL & " left outer join "
       l_strSQL = l_strSQL & " (select emp_num ,1 as Actual_Time_Sw "
       l_strSQL = l_strSQL & " From METHDBA.ttr_actual_time "
       l_strSQL = l_strSQL & " WHERE charge_date  between '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
       l_strSQL = l_strSQL & " and '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
       l_strSQL = l_strSQL & " and emp_num in "
       l_strSQL = l_strSQL & " (select emp_num "
       l_strSQL = l_strSQL & " From methdba.TTR_EMP_WG "
       l_strSQL = l_strSQL & " WHERE " & l_strInClauseWG
       l_strSQL = l_strSQL & " AND EFF_START_DATE <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
       l_strSQL = l_strSQL & " AND EFF_END_DATE >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
       l_strSQL = l_strSQL & ") "
          l_strSQL = l_strSQL & "  ) as c on a.emp_num = c.emp_num "
    End If
    If v_TrActualTimeTable = TrTable_History Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
       l_strSQL = l_strSQL & " left outer join "
       l_strSQL = l_strSQL & " (select emp_num ,1 as Actual_Time_Sw "
       l_strSQL = l_strSQL & " From METHDBA.ttr_actual_time_hs "
       l_strSQL = l_strSQL & " WHERE " & l_strInClausePartValueKey
       l_strSQL = l_strSQL & " and charge_date  between '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
       l_strSQL = l_strSQL & " and '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
       l_strSQL = l_strSQL & " and emp_num in "
       l_strSQL = l_strSQL & " (select emp_num "
       l_strSQL = l_strSQL & " From methdba.TTR_EMP_WG "
       l_strSQL = l_strSQL & " WHERE " & l_strInClauseWG
       l_strSQL = l_strSQL & " AND EFF_START_DATE <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
       l_strSQL = l_strSQL & " AND EFF_END_DATE >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
       l_strSQL = l_strSQL & ") "
       l_strSQL = l_strSQL & "  ) as d on a.emp_num = d.emp_num "
    End If
    If v_bolActiveEmployeesOnly Then
        l_strSQL = l_strSQL & " where b.emp_status <> 'T' "
        If v_TrActualTimeTable = TrTable_Actual Or _
           v_TrActualTimeTable = TrTable_ActualAndHistory Then
           l_strSQL = l_strSQL & " or c.Actual_Time_Sw = 1 "
        End If
        If v_TrActualTimeTable = TrTable_History Or _
           v_TrActualTimeTable = TrTable_ActualAndHistory Then
           l_strSQL = l_strSQL & " or d.Actual_Time_Sw = 1 "
        End If
    End If
    l_strSQL = l_strSQL & " ORDER BY EMP_NAME "
  
     
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeesInWorkGroupACTData = l_objRS
    
    Set l_objRS = Nothing
    Exit Function

GetEmployeesInWorkGroupACTData_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroupACTData", , , , DescArgsFrom(l_strSQL), True
    
End Function

Public Function GetEmployeesInWorkGroupTimeCharged( _
            ByVal v_aryWorkGroups As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False _
            ) As ADODB.Recordset
        
    Dim l_lngStart As Long
    Dim l_datCutOffChargeDate As Date
    
On Error GoTo GetEmployeesInWorkGroupTimeCharged_EH
    
    l_lngStart = LBound(v_aryWorkGroups)
    If Val(v_aryWorkGroups(l_lngStart)) < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    Else
    End If
    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set GetEmployeesInWorkGroupTimeCharged = GetEmployeesInWorkGroupTCACTData(v_aryWorkGroups, v_datStartDate, v_datEndDate, TrTable_Actual, v_bolActiveEmployeesOnly)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set GetEmployeesInWorkGroupTimeCharged = GetEmployeesInWorkGroupTCACTData(v_aryWorkGroups, v_datStartDate, v_datEndDate, TrTable_History, v_bolActiveEmployeesOnly)
        Else
           Set GetEmployeesInWorkGroupTimeCharged = GetEmployeesInWorkGroupTCACTData(v_aryWorkGroups, v_datStartDate, v_datEndDate, TrTable_ActualAndHistory, v_bolActiveEmployeesOnly) 'use both current and history data
        End If
    End If
    
    Exit Function
    
GetEmployeesInWorkGroupTimeCharged_EH:
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroupTimeCharged", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroupTimeCharged", , , , , True
    End Select
    
End Function

Private Function GetEmployeesInWorkGroupTCACTData( _
            ByVal v_aryWorkGroups As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            ByVal v_TrActualTimeTable As TrActualTimeTable, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False _
            ) As ADODB.Recordset
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClauseWG As String
    Dim l_strInClausePartValueKey As String
    
    On Error GoTo GetEmployeesInWorkGroupTCACTData_EH
   
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
    l_strInClauseWG = BuildInClause("WG_NUM", v_aryWorkGroups, False)
    If v_TrActualTimeTable = TrTable_History Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "select distinct a.emp_num ,a.emp_name ,wg_num "
        l_strSQL = l_strSQL & "from methdba.ttr_employee a ,methdba.ttr_actual_time_hs b  "
        l_strSQL = l_strSQL & " WHERE " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & " and " & l_strInClauseWG
        l_strSQL = l_strSQL & " AND A.EMP_NUM = B.EMP_NUM "
        l_strSQL = l_strSQL & " and charge_date  >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " and charge_date  <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " UNION "
    End If
    If v_TrActualTimeTable = TrTable_Actual Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "select distinct a.emp_num ,a.emp_name ,wg_num "
        l_strSQL = l_strSQL & "from methdba.ttr_employee a ,methdba.ttr_actual_time b  "
        l_strSQL = l_strSQL & " WHERE " & l_strInClauseWG
        l_strSQL = l_strSQL & " AND A.EMP_NUM = B.EMP_NUM "
        l_strSQL = l_strSQL & " and charge_date  >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " and charge_date  <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " UNION "
    End If
    l_strSQL = l_strSQL & " select distinct a.emp_num ,a.emp_name ,wg_num "
    l_strSQL = l_strSQL & "from methdba.ttr_employee a ,METHDBA.TTR_EMP_WG b  "
    l_strSQL = l_strSQL & " WHERE " & l_strInClauseWG
    l_strSQL = l_strSQL & " AND A.EMP_NUM = B.EMP_NUM "
    l_strSQL = l_strSQL & " AND EFF_START_DATE <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & " AND EFF_END_DATE >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
    If v_bolActiveEmployeesOnly Then
        l_strSQL = l_strSQL & " AND emp_status <> 'T' "
    End If
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeesInWorkGroupTCACTData = l_objRS
    
    Set l_objRS = Nothing
    Exit Function

GetEmployeesInWorkGroupTCACTData_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroupTCACTData", , , , DescArgsFrom(l_strSQL), True
    
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

Public Function GetEmployeesInDepartmentTimeCharged( _
            ByVal v_aryDepartments As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False _
            ) As ADODB.Recordset
        
    Dim l_lngStart As Long
    Dim l_datCutOffChargeDate As Date
    
On Error GoTo GetEmployeesInDepartmentTimeCharged_EH
    
'    l_lngStart = LBound(v_aryDepartments)
'    If v_aryDepartments(l_lngStart) <> "CON" Then
'        If Val(v_aryDepartments(l_lngStart)) < 1 Then
'            Err.Raise trErrdbInvalidEmptyDept
'        Else
'        End If
'    End If
    
    If Trim(v_aryDepartments(l_lngStart)) = "" Then
        Err.Raise trErrdbInvalidEmptyDept
    End If
    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set GetEmployeesInDepartmentTimeCharged = GetEmployeesInDepartmentTCACTData(v_aryDepartments, v_datStartDate, v_datEndDate, TrTable_Actual, v_bolActiveEmployeesOnly)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set GetEmployeesInDepartmentTimeCharged = GetEmployeesInDepartmentTCACTData(v_aryDepartments, v_datStartDate, v_datEndDate, TrTable_History, v_bolActiveEmployeesOnly)
        Else
           Set GetEmployeesInDepartmentTimeCharged = GetEmployeesInDepartmentTCACTData(v_aryDepartments, v_datStartDate, v_datEndDate, TrTable_ActualAndHistory, v_bolActiveEmployeesOnly) 'use both current and history data
        End If
    End If
    
    Exit Function
    
GetEmployeesInDepartmentTimeCharged_EH:
    Select Case Err.Number
        Case trErrdbInvalidEmptyDept
            TrsRaiseError m_strCMODULENAME, "GetEmployeesInDepartmentTimeCharged", trErrdbInvalidEmptyDept, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetEmployeesInDepartmentTimeCharged", , , , , True
    End Select
    
End Function

Private Function GetEmployeesInDepartmentTCACTData( _
            ByVal v_aryDepartments As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            ByVal v_TrActualTimeTable As TrActualTimeTable, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False _
            ) As ADODB.Recordset
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClauseDEP As String
    Dim l_strInClausePrevDEP As String
    Dim l_strInClausePartValueKey As String
    
    On Error GoTo GetEmployeesInDepartmentTCACTData_EH
   
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

   
    l_strInClausePartValueKey = BuildInClause("part_value_key", BuildArrayPartValueKeys(v_datStartDate, v_datEndDate), True)
    l_strInClauseDEP = BuildInClause("b.dept_num", v_aryDepartments, True)
    l_strInClausePrevDEP = BuildInClause("b.prev_dept_num", v_aryDepartments, True)
   If v_TrActualTimeTable = TrTable_History Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "select distinct a.emp_num ,a.emp_name ,b.dept_num "
        l_strSQL = l_strSQL & "from methdba.ttr_employee a ,methdba.ttr_actual_time_hs b  "
        l_strSQL = l_strSQL & " WHERE " & l_strInClausePartValueKey
        l_strSQL = l_strSQL & " and " & l_strInClauseDEP
        l_strSQL = l_strSQL & " AND A.EMP_NUM = B.EMP_NUM "
        l_strSQL = l_strSQL & " and charge_date  >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " and charge_date  <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " UNION "
    End If
    If v_TrActualTimeTable = TrTable_Actual Or _
       v_TrActualTimeTable = TrTable_ActualAndHistory Then
        l_strSQL = l_strSQL & "select distinct a.emp_num ,a.emp_name ,b.dept_num "
        l_strSQL = l_strSQL & "from methdba.ttr_employee a ,methdba.ttr_actual_time b  "
        l_strSQL = l_strSQL & " WHERE " & l_strInClauseDEP
        l_strSQL = l_strSQL & " AND A.EMP_NUM = B.EMP_NUM "
        l_strSQL = l_strSQL & " and charge_date  >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " and charge_date  <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
        l_strSQL = l_strSQL & " UNION "
    End If
    l_strSQL = l_strSQL & " select distinct b.emp_num ,b.emp_name ,b.dept_num "
    l_strSQL = l_strSQL & "from methdba.ttr_employee b "
    l_strSQL = l_strSQL & " WHERE " & l_strInClauseDEP
    l_strSQL = l_strSQL & " AND DEPT_EFF_DATE <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
    If v_bolActiveEmployeesOnly Then
        l_strSQL = l_strSQL & " AND emp_status <> 'T' "
    End If
    l_strSQL = l_strSQL & " UNION "
    l_strSQL = l_strSQL & " select distinct b.emp_num ,b.emp_name ,b.prev_dept_num as dept_num "
    l_strSQL = l_strSQL & "from methdba.ttr_employee b "
    l_strSQL = l_strSQL & " WHERE " & l_strInClausePrevDEP
    l_strSQL = l_strSQL & " AND DEPT_EFF_DATE > '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
    If v_bolActiveEmployeesOnly Then
        l_strSQL = l_strSQL & " AND emp_status <> 'T' "
    End If

    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeesInDepartmentTCACTData = l_objRS
    
    Set l_objRS = Nothing
    Exit Function

GetEmployeesInDepartmentTCACTData_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeesInDepartmentTCACTData", , , , DescArgsFrom(l_strSQL), True
    
End Function

'select distinct a.emp_num ,a.emp_name ,b.dept_num
'from methdba.ttr_employee a ,methdba.ttr_actual_time b
'WHERE b.dept_num IN ('050','057')  AND A.EMP_NUM = B.EMP_NUM
'and charge_date  >= '03/04/2001'  and charge_date  <= '03/17/2001'
'UNION
'select distinct a.emp_num ,a.emp_name ,dept_num
'from methdba.ttr_employee a
'WHERE dept_num IN ('050','057') and DEPT_EFF_DATE <= '03/17/2001'
'union
'select distinct a.emp_num ,a.emp_name ,prev_dept_num
'from methdba.ttr_employee a
'WHERE  prev_dept_num IN ('050','057') and DEPT_EFF_DATE > '03/04/2001'



Public Function GetEmployeesAssignedToWorkGroup( _
            ByVal v_aryWorkGroups As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False _
            ) As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_lngStart As Long
    Dim l_lng1stWgNum As Long
    Dim l_strInClauseWG As String
    Dim l_dat6MonthsAgo As Date

    On Error GoTo GetEmployeesAssignedToWorkGroup_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
            
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    l_objCN.Open

'Check to see if at least 1 WorkGroup Number was passed
    If IsArray(v_aryWorkGroups) Then
        l_lngStart = LBound(v_aryWorkGroups)
        l_lng1stWgNum = Val(v_aryWorkGroups(l_lngStart))
    Else
        l_lng1stWgNum = CLng(v_aryWorkGroups)
    End If
    If l_lng1stWgNum < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    End If

'Have a work group(s) to select
    If IsArray(v_aryWorkGroups) Then
        l_strInClauseWG = BuildInClause("WG_NUM", v_aryWorkGroups, False)
    Else
        l_strInClauseWG = "WG_NUM = " & CStr(v_aryWorkGroups)
    End If

'''''    l_strSQL = "select a.emp_num ,b.emp_name,a.wg_num, b.emp_status "
'''''    l_strSQL = l_strSQL & "From methdba.TTR_EMP_WG a, methdba.TTR_EMPLOYEE b "
'''''    l_strSQL = l_strSQL & " WHERE " & l_strInClauseWG
'''''    l_strSQL = l_strSQL & " AND EFF_START_DATE <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
'''''    l_strSQL = l_strSQL & " AND EFF_END_DATE >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
'''''    l_strSQL = l_strSQL & " AND A.EMP_NUM = B.EMP_NUM "
'''''    If v_bolActiveEmployeesOnly Then
'''''        l_dat6MonthsAgo = DateAdd("m", -6, Date)
'''''        l_strSQL = l_strSQL & " AND (b.EMP_STATUS <> 'T' or status_eff_date >= '" & l_dat6MonthsAgo & "' "
'''''    End If
    l_strSQL = "select NE1.EMP_NUM, NE2.EMP_NAME, NE1.WG_NUM, NE2.EMP_STATUS "
    l_strSQL = l_strSQL & "FROM ( "
    l_strSQL = l_strSQL & "select value(y.emp_num,z.emp_num) as emp_num, "
    l_strSQL = l_strSQL & " value(y.wg_num, z.wg_num) as wg_num "
    l_strSQL = l_strSQL & " FROM ("
    l_strSQL = l_strSQL & "select distinct emp_num, wg_num from METHDBA.TTR_EMP_WG "
    l_strSQL = l_strSQL & " WHERE " & l_strInClauseWG
    l_strSQL = l_strSQL & " AND EFF_START_DATE <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & " AND EFF_END_DATE >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & " ) y FULL OUTER JOIN "
    l_strSQL = l_strSQL & "(select distinct emp_num, wg_num from METHDBA.TTR_LINE_ASGNM "
    l_strSQL = l_strSQL & " WHERE " & l_strInClauseWG
    l_strSQL = l_strSQL & " AND START_DATE <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & " AND END_DATE >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & " ) z ON y.EMP_NUM = z.EMP_NUM "
    l_strSQL = l_strSQL & " AND y.WG_NUM = z.WG_NUM) NE1 "
    l_strSQL = l_strSQL & " LEFT OUTER JOIN "
    l_strSQL = l_strSQL & " ( select emp_num, emp_name, emp_status"
    l_strSQL = l_strSQL & "   FROM METHDBA.TTR_EMPLOYEE "
    If v_bolActiveEmployeesOnly Then
        l_dat6MonthsAgo = DateAdd("m", -6, Date)
        l_strSQL = l_strSQL & " WHERE EMP_STATUS <> 'T' or status_eff_date >= '" & l_dat6MonthsAgo & "' "
    End If
    l_strSQL = l_strSQL & " ) NE2 ON NE1.EMP_NUM = NE2.EMP_NUM "
    
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetEmployeesAssignedToWorkGroup = l_objRS

    Set l_objRS = Nothing
    Exit Function
GetEmployeesAssignedToWorkGroup_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetEmployeesAssignedToWorkGroup", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetEmployeesAssignedToWorkGroup", , , , , True
    End Select


End Function

Public Function GetEmployeeAssignmentsInWorkGroup( _
            ByVal v_lngWorkGroup As Long, _
            ByVal v_datDate As Date _
            ) As ADODB.Recordset
        
'RED256
            
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    
    On Error GoTo GetEmployeesAssignedToWorkGroup_EH
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open

'Check to see if at least 1 WorkGroup Number was passed
    If v_lngWorkGroup < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    End If
    
'Have a work group(s) to select
    l_strSQL = "select primary_ind, b.emp_name,eff_start_date,eff_end_date, "
    l_strSQL = l_strSQL & "temp_ot_bank_ind, a.emp_num, "
    l_strSQL = l_strSQL & "wg_num, 'L' as RECORD_STATUS, SUBSYS_OWNER, "
    l_strSQL = l_strSQL & " CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & " CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP, "
    l_strSQL = l_strSQL & "A.LAST_MOD_EMP_NUM, A.CREATED_EMP_NUM, 'DB' as DATA_SOURCE "
    l_strSQL = l_strSQL & "From methdba.TTR_EMP_WG a, methdba.TTR_EMPLOYEE b "
    l_strSQL = l_strSQL & " WHERE WG_NUM = " & str(v_lngWorkGroup)
    l_strSQL = l_strSQL & " AND EFF_END_DATE >= '" & Format(v_datDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & " AND A.EMP_NUM = B.EMP_NUM "

         
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeAssignmentsInWorkGroup = l_objRS
    
    Set l_objRS = Nothing
    Exit Function
GetEmployeesAssignedToWorkGroup_EH:
   
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetEmployeesAssignedToWorkGroup", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetEmployeesAssignedToWorkGroup", , , , , True
    End Select
    
End Function
    
    
Public Function GetEmployeesInWorkGroupViewTime( _
            ByVal v_aryWorkGroups As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False _
            ) As ADODB.Recordset
        
    Dim l_lngStart As Long
    Dim l_datCutOffChargeDate As Date
    
On Error GoTo GetEmployeesInWorkGroupViewTime_EH
    
    l_lngStart = LBound(v_aryWorkGroups)
    If Val(v_aryWorkGroups(l_lngStart)) < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    End If
    
    l_datCutOffChargeDate = CDate(Trim(GetActualTimeHSCutoffDate))
    
    If v_datStartDate > l_datCutOffChargeDate Then    'use current data
        Set GetEmployeesInWorkGroupViewTime = GetEmployeesInWorkGroupACTData(v_aryWorkGroups, v_datStartDate, v_datEndDate, TrTable_Actual, v_bolActiveEmployeesOnly)
    Else
        If v_datEndDate <= l_datCutOffChargeDate Then 'use history data
           Set GetEmployeesInWorkGroupViewTime = GetEmployeesInWorkGroupACTData(v_aryWorkGroups, v_datStartDate, v_datEndDate, TrTable_History, v_bolActiveEmployeesOnly)
        Else
           Set GetEmployeesInWorkGroupViewTime = GetEmployeesInWorkGroupACTData(v_aryWorkGroups, v_datStartDate, v_datEndDate, TrTable_ActualAndHistory, v_bolActiveEmployeesOnly) 'use both current and history data
        End If
    End If
    
    Exit Function
    
GetEmployeesInWorkGroupViewTime_EH:
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroupViewTime", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroupViewTime", , , , , True
    End Select
    
End Function


Public Function GetWorkGroupAssignments( _
            ByVal v_lngWorkGroup As Long, _
            ByVal v_datDate As Date _
            ) As ADODB.Recordset
        
'RED256
            
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    
    On Error GoTo GetWGAssignments_EH
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
'Have a work group(s) to select
    l_strSQL = "select emp_num,wg_num, primary_ind,eff_start_date, "
    l_strSQL = l_strSQL & "eff_end_date,temp_ot_bank_ind,"
    l_strSQL = l_strSQL & "last_mod_emp_num,created_emp_num, SUBSYS_OWNER, "
    l_strSQL = l_strSQL & " CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & " CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "From methdba.TTR_EMP_WG "
    l_strSQL = l_strSQL & " WHERE WG_NUM = " & str(v_lngWorkGroup)
    l_strSQL = l_strSQL & " AND EFF_END_DATE >= '" & Format(v_datDate, "mm/dd/yyyy") & "' "

         
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWorkGroupAssignments = l_objRS
    
    Set l_objRS = Nothing
    Exit Function
    
GetWGAssignments_EH:
   
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing

    TrsRaiseError m_strCMODULENAME, "GetWorkGroupAssignments", , , , , True
  
End Function

Public Function GetEmployeeWorkGroupAssignments( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datDate As Date, _
            ByVal v_lngWorkGroup As Long, _
            Optional v_blnPrimaryOnly As Boolean = True _
            ) As ADODB.Recordset
        
'RED256
            
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetEmployeeWorkGroupAssignments_EH
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
'Have a work group(s) to select
    l_strSQL = "select emp_num,wg_num, primary_ind,eff_start_date, "
    l_strSQL = l_strSQL & "eff_end_date,temp_ot_bank_ind,"
    l_strSQL = l_strSQL & "last_mod_emp_num,created_emp_num, SUBSYS_OWNER, "
    l_strSQL = l_strSQL & " CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & " CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "From methdba.TTR_EMP_WG "
    l_strSQL = l_strSQL & " WHERE " & BuildInClause("EMP_NUM", v_aryEmployees, False)
    l_strSQL = l_strSQL & " AND WG_NUM <>  " & str(v_lngWorkGroup)
    l_strSQL = l_strSQL & " AND EFF_END_DATE >= '" & Format(v_datDate, "mm/dd/yyyy") & "' "
    If v_blnPrimaryOnly Then
        l_strSQL = l_strSQL & " and (primary_ind = 'Y' or subsys_owner = 'R')"
    End If
    l_strSQL = l_strSQL & " order by emp_num, wg_num, eff_start_date desc "

         
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeeWorkGroupAssignments = l_objRS
    
    Set l_objRS = Nothing
    Exit Function
    
GetEmployeeWorkGroupAssignments_EH:
   
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing

    TrsRaiseError m_strCMODULENAME, "GetEmployeeWorkGroupAssignments", , , , , True
  
End Function

'Public Function WeeksToSchedule( _
'            ByVal v_lngWgNum As Long) As Long
'
''This method should be deleted
'
'On Error GoTo eh
'    WeeksToSchedule = 0
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "WeeksToSchedule", , , , , True
'
'End Function

'Public Function GetWGDemandSchedules( _
'        ByVal v_lngWgNum As Long, _
'        ByVal v_datDate As Date) As ADODB.Recordset
'
'' Obsolete - Remove on the first build for TRIS4.
'
'End Function


Public Function GetWGDemandLineAssigmments( _
        ByVal v_lngWgNum As Long, _
        ByVal v_datDate As Date) As ADODB.Recordset
        
'This method gets all Demand Line Assigments for a work group
' after a certain date
'Need to verify if the validation that needs this data is still needed
'If so does the SQL return the necessary data.

    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetWGDemandLineAssigmments_EH

    If v_lngWgNum < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    Else
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")

    Set l_objCM = Nothing

    l_objCN.Open

    l_strSQL = "SELECT WG_NUM, LINE_NUM EMP_NUM, START_DATE, END_DATE  "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_LINE_ASGNM "
    l_strSQL = l_strSQL & "WHERE WG_NUM = " & v_lngWgNum
    l_strSQL = l_strSQL & "  and END_DATE > '" & Format(v_datDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & " ORDER BY 2 "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetWGDemandLineAssigmments = l_objRS

    Set l_objRS = Nothing
    Exit Function

GetWGDemandLineAssigmments_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing

    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetWGDemandLineAssigmments", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetWGDemandLineAssigmments", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function


Public Function GetWorkGroupAssociations( _
            ByVal v_lngWgNums As Variant _
            ) As ADODB.Recordset
        
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhere As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    
    On Error GoTo GetWorkGroupAssociations_EH
    
    If IsArray(v_lngWgNums) Then
        l_strWhere = "WHERE " & BuildInClause("A.WG_NUM", v_lngWgNums, False)
    Else
        l_strWhere = "WHERE A.WG_NUM = " & CStr(v_lngWgNums)
    End If
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
'Have a work group(s) to select
    l_strSQL = "select DIVISION_NUM, FUNCTION_NUM, DEPT_NUM, "
    l_strSQL = l_strSQL & " ,LOC_CD, WG_NUM, HR_LOC "
    l_strSQL = l_strSQL & "From methdba.TTR_ASSOCIATIONS "
    l_strSQL = l_strSQL & l_strWhere
    l_strSQL = l_strSQL & " and wg_num > 0 "
    l_strSQL = l_strSQL & "order by wg_num "
         
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWorkGroupAssociations = l_objRS
    
    Set l_objRS = Nothing
    Exit Function
    
GetWorkGroupAssociations_EH:
   
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing

    TrsRaiseError m_strCMODULENAME, "GetWorkGroupAssociations", , , , , True
  
End Function

Public Function GetWorkgroupAttributes( _
            ByVal v_aryWgNums As Variant) As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strWhere As String
    
On Error GoTo eh
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = _
        l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryWgNums) Then
        l_strWhere = "WHERE " & BuildInClause("A.WG_NUM", v_aryWgNums, False)
    Else
        l_strWhere = "WHERE A.WG_NUM = " & CStr(v_aryWgNums)
    End If

'The first select joins on ttr_projects to pick up the eff_start_date and
' eff_end_date for PPM project, but this will only yield a resulting row when the project_id
' is not equal to zero.  The second select will yield a row for non-PPM workgroup,
' or when the project_id = 0.
    
    l_strSQL = "SELECT A.WG_NUM, A.WG_DESC, A.EFF_START_DATE, A.EFF_END_DATE, "
    l_strSQL = l_strSQL & "A.PROJECT_ID, A.DEMAND_IND, A.WG_SUPERVISOR, "
    l_strSQL = l_strSQL & "A.CREW_NUM, B.DEPT_NUM, B.LOC_CD, "
    l_strSQL = l_strSQL & "C.EFF_START_DATE AS PROJ_START_DATE, "
    l_strSQL = l_strSQL & "C.EFF_END_DATE AS PROJ_END_DATE "
    l_strSQL = l_strSQL & ",E.EMP_NAME as WG_SUPERVISOR_NAME "
    l_strSQL = l_strSQL & ",B.HR_LOC, A.LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "'L' as RECORD_STATUS "
    l_strSQL = l_strSQL & ",CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP"
    l_strSQL = l_strSQL & ",CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP"
    l_strSQL = l_strSQL & ",CREATED_EMP_NUM AS CREATED_EMP_NUM "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_WORK_GROUP A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_ASSOCIATIONS B, "
    l_strSQL = l_strSQL & "METHDBA.TTR_PROJECTS C "
    l_strSQL = l_strSQL & ",METHDBA.TTR_EMPLOYEE E "
    l_strSQL = l_strSQL & l_strWhere
    l_strSQL = l_strSQL & " AND A.WG_NUM = B.WG_NUM"
    l_strSQL = l_strSQL & " AND A.PROJECT_ID = C.PROJECT_ID"
    l_strSQL = l_strSQL & " AND C.PPM_IND = 'Y'"
    l_strSQL = l_strSQL & " AND A.WG_SUPERVISOR = E.EMP_NUM"
    l_strSQL = l_strSQL & " UNION "
    l_strSQL = l_strSQL & "SELECT A.WG_NUM, A.WG_DESC, A.EFF_START_DATE, A.EFF_END_DATE, "
    l_strSQL = l_strSQL & "A.PROJECT_ID, A.DEMAND_IND, A.WG_SUPERVISOR, "
    l_strSQL = l_strSQL & "A.CREW_NUM, B.DEPT_NUM, B.LOC_CD, "
    l_strSQL = l_strSQL & "DATE('01/01/0100') AS PROJ_START_DATE, "
    l_strSQL = l_strSQL & "DATE('01/01/0100') AS PROJ_END_DATE "
    l_strSQL = l_strSQL & ",E.EMP_NAME as WG_SUPERVISOR_NAME "
    l_strSQL = l_strSQL & ",B.HR_LOC, A.LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & " 'L' as RECORD_STATUS "
    l_strSQL = l_strSQL & ",CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP"
    l_strSQL = l_strSQL & ",CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP"
    l_strSQL = l_strSQL & ",CREATED_EMP_NUM AS CREATED_EMP_NUM "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_WORK_GROUP A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_ASSOCIATIONS B "
    l_strSQL = l_strSQL & ",METHDBA.TTR_EMPLOYEE E "
    l_strSQL = l_strSQL & l_strWhere
    l_strSQL = l_strSQL & " AND A.WG_NUM = B.WG_NUM"
    l_strSQL = l_strSQL & " AND A.PROJECT_ID = 0"
    l_strSQL = l_strSQL & " AND A.WG_SUPERVISOR = E.EMP_NUM"
    If IsArray(v_aryWgNums) Then
        l_strSQL = l_strSQL & " ORDER BY WG_NUM"
    End If
        
    Set l_objRS = New ADODB.Recordset
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    
    Set GetWorkgroupAttributes = l_objRS
    Set l_objRS = Nothing
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkgroupAttributes", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetWorkGroupActivities(ByVal v_lngWgNum As Long) As ADODB.Recordset
    
'Similar to GetActivities
'This will be used in the context of a transaction
'GetActivities gets called often
'The seperation makes better use of the database resources

    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetWorkGroupActivities_EH
    
    If v_lngWgNum < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    Else
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT A.ACTIVITY_CD , A.ACTIVITY_DESC, "
    l_strSQL = l_strSQL & "B.WG_NUM , B.LAST_MOD_EMP_NUM, CHAR(B.TIME_STAMP) as TIME_STAMP, "
    l_strSQL = l_strSQL & "'L' as RECORD_STATUS, 'DB' as DATA_SOURCE "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACTIVITY A "
    l_strSQL = l_strSQL & " , METHDBA.TTR_ACTIVITY_ASGNM B "
    l_strSQL = l_strSQL & "WHERE WG_NUM = " & v_lngWgNum
    l_strSQL = l_strSQL & "  and A.ACTIVITY_CD = B.ACTIVITY_CD "
    l_strSQL = l_strSQL & " ORDER BY 2 "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWorkGroupActivities = l_objRS

    Set l_objRS = Nothing
    Exit Function

GetWorkGroupActivities_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetWorkGroupActivities", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetWorkGroupActivities", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function


Public Function GetWorkGroupDistribution(ByVal v_lngWgNum As Long) As ADODB.Recordset
    
'    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim oADOHelper As utilADOHelper.cuADOHelper
    On Error GoTo GetWorkGroupDistribution_EH
    
    If v_lngWgNum < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    Else
    End If
    
'    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr

'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open
    
'    l_strSQL = "SELECT DISTRIBUTION_ID, DESCRIPTION, "
'    l_strSQL = l_strSQL & "WG_NUM , LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP, "
'    l_strSQL = l_strSQL & "'L' as RECORD_STATUS, 'DB' as DATA_SOURCE "
'    l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACCT_ASSIGNMTS "
'    l_strSQL = l_strSQL & "WHERE WG_NUM = " & v_lngWgNum
'    l_strSQL = l_strSQL & " ORDER BY DESCRIPTION "
    
    If IsPOETActive Then
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetWorkGroupDistributionOTTC.sql")
    Else
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetWorkGroupDistribution.sql")
    End If
    
    Set oADOHelper = New utilADOHelper.cuADOHelper
    oADOHelper.ADO.CreateNumericParameter "WG_NUM", adInteger, 0, 0, v_lngWgNum
    Set l_objRS = oADOHelper.ExecuteSelect(l_strSQL, adOpenStatic, adLockBatchOptimistic)
    
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
    Set oADOHelper = Nothing
    Set GetWorkGroupDistribution = l_objRS

    Set l_objRS = Nothing
    Exit Function

GetWorkGroupDistribution_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
    Set l_objRS = Nothing
    Set oADOHelper = Nothing
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetWorkGroupDistribution", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetWorkGroupDistribution", , , , DescArgsFrom(l_strSQL, CStr(v_lngWgNum)), True
    End Select

End Function

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

Public Function GetWorkgroupForUpdate( _
            ByVal v_aryWgNums As Variant) As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strWhere As String
    
On Error GoTo eh
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = _
        l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryWgNums) Then
        l_strWhere = "WHERE " & BuildInClause("A.WG_NUM", v_aryWgNums, False)
    Else
        l_strWhere = "WHERE A.WG_NUM = " & CStr(v_aryWgNums)
    End If

'DEPT_NUM and LOC_CD should be on the ttr_work_group
'Jimb does not know why
    
    l_strSQL = "SELECT A.WG_NUM, A.WG_DESC, A.EFF_START_DATE, A.EFF_END_DATE, "
    l_strSQL = l_strSQL & "A.PROJECT_ID, A.DEMAND_IND, A.WG_SUPERVISOR, "
    l_strSQL = l_strSQL & "A.CREW_NUM, B.DEPT_NUM, B.LOC_CD,"
    l_strSQL = l_strSQL & "A.LAST_MOD_EMP_NUM,"
    l_strSQL = l_strSQL & "CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP,"
    l_strSQL = l_strSQL & "CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP,"
    l_strSQL = l_strSQL & "A.CREATED_EMP_NUM "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_WORK_GROUP A, "
    l_strSQL = l_strSQL & "METHDBA.TTR_ASSOCIATIONS B "
    l_strSQL = l_strSQL & l_strWhere
    l_strSQL = l_strSQL & " AND A.WG_NUM = B.WG_NUM"
    If IsArray(v_aryWgNums) Then
        l_strSQL = l_strSQL & " ORDER BY WG_NUM"
    End If
        
    Set l_objRS = New ADODB.Recordset
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    
    Set GetWorkgroupForUpdate = l_objRS
    Set l_objRS = Nothing
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkgroupAttributes", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetDemandLineBaseSchedules( _
        ByVal v_aryWgNums As Variant, _
        ByVal v_datDate As Date _
        ) As ADODB.Recordset
    
    Dim l_intIndex As Integer
    Dim l_strSQLWhereClause As String
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetDemandLineBaseSchedules_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryWgNums) Then
        l_strSQLWhereClause = "WHERE " & BuildInClause("WG_NUM", v_aryWgNums, False)
    Else
        l_strSQLWhereClause = "WHERE WG_NUM = " & CStr(v_aryWgNums)
    End If
    
    l_strSQL = "select WG_NUM, LINE_NUM, START_DATE, END_DATE, ROT_CD, ROT_WEEK, SCHED_CD  " & _
                " , JOB_NUM, REQ_LINE_IND " & _
                " , CREATED_EMP_NUM, LAST_MOD_EMP_NUM " & _
                " , Char(CREATED_TIMESTAMP) as CREATED_TIMESTAMP " & _
                " , Char(LAST_MOD_TIMESTAMP) as LAST_MOD_TIMESTAMP " & _
                " , END_DATE " & _
                "from methdba.ttr_wg_line_sched " & _
                l_strSQLWhereClause & " " & _
                " and END_DATE >= '" & Format(v_datDate, "yyyy-mm-dd") & "'  " & _
                "order by wg_num, line_num, start_date desc"

                                   
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDemandLineBaseSchedules = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetDemandLineBaseSchedules_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineBaseSchedules", , , , DescArgsFrom(l_strSQL), True
    
End Function


Public Function GetBaseSchedules( _
        ByVal v_aryWgNums As Variant, _
        ByVal v_datDate As Date _
        ) As ADODB.Recordset

    Dim l_intIndex As Integer
    Dim l_strSQLWhereClause As String
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo GetBaseSchedules_EH

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryWgNums) Then
        l_strSQLWhereClause = "WHERE " & BuildInClause("WG_NUM", v_aryWgNums, False)
    Else
        l_strSQLWhereClause = "WHERE WG_NUM = " & CStr(v_aryWgNums)
    End If
    
    l_strSQL = "select WG_NUM , EFF_DATE , ROT_CD, ROT_WEEK, SCHED_CD " & _
                " , CREATED_EMP_NUM, LAST_MOD_EMP_NUM " & _
                " , Char(CREATED_TIMESTAMP) as CREATED_TIMESTAMP " & _
                " , Char(LAST_MOD_TIMESTAMP) as LAST_MOD_TIMESTAMP " & _
                " , END_DATE " & _
                "from methdba.ttr_wg_base_sched " & _
                l_strSQLWhereClause & " " & _
                " and END_DATE >= '" & Format(v_datDate, "yyyy-mm-dd") & "' " & _
                "order by wg_num, eff_date desc"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetBaseSchedules = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetBaseSchedules_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetBaseSchedules", , , , DescArgsFrom(l_strSQL), True

    
End Function

' RSG
Public Function GetDemandLineSchedules( _
            ByVal v_lngWorkGroup As Long, _
            ByVal v_datMinEffDate As Date, _
            ByVal v_datMaxEffDate As Date, _
            ByVal v_varLineNums As Variant _
        ) As ADODB.Recordset
' Limit selection to the lines affected, minimum effective date in request,
' and maximum effective date in request.
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strMinEffDate As String
    Dim l_strMaxEffDate As String
    Dim l_strLineInClause As String
    
On Error GoTo eh
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    l_strMinEffDate = "'" & Format(v_datMinEffDate, "mm/dd/yyyy") & "'"
    l_strMaxEffDate = "'" & Format(v_datMaxEffDate, "mm/dd/yyyy") & "'"
    l_strLineInClause = BuildInClause("LINE_NUM", v_varLineNums, False)

    l_strSQL = "SELECT WG_NUM, LINE_NUM, START_DATE, " _
        & " ROT_CD, ROT_WEEK, SCHED_CD, JOB_NUM, REQ_LINE_IND, END_DATE," _
        & " CREATED_EMP_NUM, LAST_MOD_EMP_NUM, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP," _
        & " CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" _
        & " FROM METHDBA.TTR_WG_LINE_SCHED" _
        & " WHERE WG_NUM = " & str(v_lngWorkGroup) _
        & " AND " & l_strLineInClause _
        & " AND START_DATE <= " & l_strMaxEffDate _
        & " AND END_DATE >= " & l_strMinEffDate _
        & " ORDER BY WG_NUM, LINE_NUM, START_DATE"

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDemandLineSchedules = l_objRS
    Set l_objRS = Nothing
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineSchedules"
End Function

Public Function GetWorkGroupDistributionForUpdate( _
    ByVal v_lngWgNum As Long _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetWorkGroupDistributionForUpdate_EH
    
'    If v_lngWGNum < 1 Then
'        Err.Raise trErrdbInvalidZerolngWgNum
'    Else
'    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT EMP_NUM, WG_NUM, DISTRIBUTION_ID, DESCRIPTION, "
    l_strSQL = l_strSQL & " LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACCT_ASSIGNMTS "
    l_strSQL = l_strSQL & "WHERE WG_NUM = " & v_lngWgNum
    l_strSQL = l_strSQL & " ORDER BY DISTRIBUTION_ID "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWorkGroupDistributionForUpdate = l_objRS

    Set l_objRS = Nothing
    Exit Function

GetWorkGroupDistributionForUpdate_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetWorkGroupDistributionForUpdate", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetWorkGroupDistributionForUpdate", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function

Public Function GetWorkGroupActivitiesForUpdate( _
    ByVal v_lngWgNum As Long _
    ) As ADODB.Recordset
    
'Similar to GetActivities
'This will be used in the context of a transaction
'GetActivities gets called often
'The seperation makes better use of the database resources

    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetWorkGroupActivitiesForUpdate_EH
    
''    If v_lngWGNum < 1 Then
''        Err.Raise trErrdbInvalidZerolngWgNum
''    Else
''    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT  ACTIVITY_CD ,WG_NUM , LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "LVL_NUM, CHAR(TIME_STAMP) as TIME_STAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACTIVITY_ASGNM  "
    l_strSQL = l_strSQL & "WHERE WG_NUM = " & v_lngWgNum
    l_strSQL = l_strSQL & " ORDER BY 1 "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWorkGroupActivitiesForUpdate = l_objRS

    Set l_objRS = Nothing
    Exit Function

GetWorkGroupActivitiesForUpdate_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetWorkGroupActivitiesForUpdate", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetWorkGroupActivitiesForUpdate", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function


Public Function GetGLActivities( _
            ByVal v_strActivities As Variant) As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strWhere As String
    
On Error GoTo eh
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = _
        l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_strActivities) Then
        l_strWhere = "WHERE " & BuildInClause("ACTIVITY_CD", v_strActivities, True)
    Else
        l_strWhere = "WHERE ACTIVITY_CD = '" & v_strActivities & "'"
    End If
    
    l_strSQL = "SELECT ACTIVITY_CD, ACTIVITY_DESC, CHAR(TIMESTAMP) as TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_ACTIVITY  "
    l_strSQL = l_strSQL & l_strWhere
    l_strSQL = l_strSQL & " ORDER BY ACTIVITY_CD"
        
    Set l_objRS = New ADODB.Recordset
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    
    Set GetGLActivities = l_objRS
    Set l_objRS = Nothing
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetGLActivities", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetEmployeesForWGSchedRpt(ByVal vlWGNum As Long _
                    , ByVal vblnUnmannedLines As Boolean _
                    , ByVal vblnDemandLine As Boolean _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset

On Error GoTo GetEmployeesForWGSchedRpt_EH
'SQL Statement Container
Dim lsSQLStatement As String, ladrs As ADODB.Recordset

    If vblnDemandLine Then
        lsSQLStatement = " SELECT EMP_NUM, WG_NUM, PRIMARY_IND, LINE_NUM " & _
                         " , REQ_JOB_IND, JOB_NUM " & _
                         " FROM (SELECT COALESCE(NE3.EMP_NUM, NE7.EMP_NUM) AS EMP_NUM " & _
                                      " ,COALESCE(NE3.WG_NUM,NE7.WG_NUM) AS WG_NUM " & _
                                      " ,COALESCE(NE3.PRIMARY_IND,'N') AS PRIMARY_IND " & _
                                      " ,COALESCE(NE7.LINE_NUM,0) AS LINE_NUM " & _
                                      " ,COALESCE(NE7.REQ_LINE_IND,'N') AS REQ_JOB_IND " & _
                                      " ,COALESCE(NE7.JOB_NUM,'') AS JOB_NUM " & _
                                      " FROM (SELECT CASE WHEN NE1.EMP_NUM IS NOT NULL " & _
                                                         "THEN NE1.EMP_NUM ELSE NE2.EMP_NUM " & _
                                                         "END AS EMP_NUM " & _
                                                  ", CASE WHEN NE1.WG_NUM IS NOT NULL " & _
                                                         "THEN NE1.WG_NUM ELSE NE2.WG_NUM " & _
                                                         "END AS WG_NUM " & _
                                                  ", CASE WHEN NE1.PRIMARY_IND IS NOT NULL " & _
                                                         "THEN NE1.PRIMARY_IND ELSE NE2.PRIMARY_IND " & _
                                                         "END AS PRIMARY_IND " & _
                                                " FROM (SELECT EMP_NUM, WG_NUM, PRIMARY_IND " & _
                                                        " FROM METHDBA.TTR_EMP_WG " & _
                                                        " WHERE WG_NUM = " & vlWGNum & _
                                                        " AND EFF_START_DATE <= '" & Format(vdtEndDate, "mm/dd/yyyy") & "'" & _
                                                        " AND EFF_END_DATE >= '" & Format(vdtStartDate, "mm/dd/yyyy") & "'" & _
                                                        " AND PRIMARY_IND = 'Y' " & _
                                                        " GROUP BY EMP_NUM, WG_NUM, PRIMARY_IND) NE1 "
                                                        
        lsSQLStatement = lsSQLStatement & " FULL OUTER JOIN " & _
                         " (SELECT EMP_NUM, WG_NUM, 'N' AS PRIMARY_IND " & _
                         " FROM METHDBA.TTR_LINE_ASGNM " & _
                         " WHERE WG_NUM = " & vlWGNum & _
                         " AND START_DATE <= '" & Format(vdtEndDate, "mm/dd/yyyy") & "'" & _
                         " AND END_DATE >= '" & Format(vdtStartDate, "mm/dd/yyyy") & "'" & _
                         " GROUP BY EMP_NUM, WG_NUM) NE2 " & _
                         " ON NE2.EMP_NUM = NE1.EMP_NUM) NE3 " & _
                         " FULL OUTER JOIN " & _
                         " (SELECT EMP_NUM, WG_NUM, LINE_NUM, REQ_LINE_IND, JOB_NUM " & _
                            " FROM (SELECT COALESCE(NE5.EMP_NUM, 0) AS EMP_NUM " & _
                                         ", NE4.WG_NUM, NE4.LINE_NUM, NE4.REQ_LINE_IND, NE4.JOB_NUM " & _
                                    " FROM (SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE, REQ_LINE_IND, JOB_NUM " & _
                                            " FROM METHDBA.TTR_WG_LINE_SCHED " & _
                                            " WHERE WG_NUM = " & vlWGNum & _
                                            " AND START_DATE <= '" & Format(vdtEndDate, "mm/dd/yyyy") & "'" & _
                                            " AND END_DATE >= '" & Format(vdtStartDate, "mm/dd/yyyy") & "'" & _
                                            " GROUP BY WG_NUM, LINE_NUM, START_DATE, END_DATE, REQ_LINE_IND, JOB_NUM) NE4 " & _
                                            " LEFT OUTER JOIN " & _
                                            " (SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE, EMP_NUM " & _
                                              " FROM METHDBA.TTR_LINE_ASGNM " & _
                                              " WHERE WG_NUM = " & vlWGNum & _
                                              " AND START_DATE <= '" & Format(vdtEndDate, "mm/dd/yyyy") & "'" & _
                                              " AND END_DATE >= '" & Format(vdtStartDate, "mm/dd/yyyy") & "'" & _
                                              " GROUP BY WG_NUM, LINE_NUM, START_DATE, END_DATE, EMP_NUM) NE5 "

        lsSQLStatement = lsSQLStatement & " ON NE5.WG_NUM = NE4.WG_NUM " & _
                        " AND NE5.LINE_NUM = NE4.LINE_NUM " & _
                        " AND NE5.START_DATE <= NE4.END_DATE " & _
                        " AND NE5.END_DATE >= NE4.START_DATE) NE6 " & _
                        " GROUP BY EMP_NUM, WG_NUM, LINE_NUM, REQ_LINE_IND, JOB_NUM) NE7 " & _
                        " ON NE7.EMP_NUM = NE3.EMP_NUM " & _
                        " AND NE7.WG_NUM = NE3.WG_NUM) NE8 "
                        
        If vblnUnmannedLines Then
            lsSQLStatement = lsSQLStatement & _
            " GROUP BY LINE_NUM, EMP_NUM, WG_NUM, PRIMARY_IND, REQ_JOB_IND, JOB_NUM "
        Else
            lsSQLStatement = lsSQLStatement & _
            " WHERE EMP_NUM > 0 " & _
            " GROUP BY LINE_NUM, EMP_NUM, WG_NUM, PRIMARY_IND, REQ_JOB_IND, JOB_NUM "
        End If
    Else
        lsSQLStatement = "SELECT EMP_NUM, WG_NUM, PRIMARY_IND, EFF_START_DATE, " & _
                               " EFF_END_DATE, LAST_MOD_EMP_NUM " & _
                          " FROM METHDBA.TTR_EMP_WG " & _
                          " WHERE WG_NUM = " & vlWGNum & _
                          " AND EFF_START_DATE <= '" & Format(vdtEndDate, "mm/dd/yyyy") & "'" & _
                          " AND EFF_END_DATE >= '" & Format(vdtStartDate, "mm/dd/yyyy") & "'" & _
                          " UNION " & _
                          " SELECT EMP_NUM, WG_NUM, 'N' AS PRIMARY_IND, START_DATE AS EFF_START_DATE, " & _
                          " END_DATE AS EFF_END_DATE, LAST_MOD_EMP_NUM " & _
                          " FROM METHDBA.TTR_LINE_ASGNM A " & _
                          " WHERE WG_NUM = " & vlWGNum & _
                          " AND START_DATE <= '" & Format(vdtEndDate, "mm/dd/yyyy") & "'" & _
                          " AND END_DATE >= '" & Format(vdtStartDate, "mm/dd/yyyy") & "'" & _
                          " AND EXISTS(SELECT 1 FROM METHDBA.TTR_EMP_WG B " & _
                                      " WHERE B.EMP_NUM = A.EMP_NUM " & _
                                      " AND B.EFF_START_DATE <= '" & Format(vdtEndDate, "mm/dd/yyyy") & "'" & _
                                      " AND B.EFF_END_DATE >= '" & Format(vdtStartDate, "mm/dd/yyyy") & "'" & _
                                      " AND B.PRIMARY_IND = 'Y' " & _
                                      " AND B.WG_NUM <> A.WG_NUM) " & _
                                      " ORDER BY 1 "
    End If

    'Call ExecuteSQL and retrieve information
    Set ladrs = ExecuteSQL(lsSQLStatement _
                            , adOpenStatic _
                            , adLockReadOnly)

    'Check if recordset is empty
    If (mADOExtensions.RecordsExist(ladrs)) Then
        Set GetEmployeesForWGSchedRpt = ladrs
    Else
        Set GetEmployeesForWGSchedRpt = Nothing
    End If

    'Destroy Recordset
    Set ladrs = Nothing

    Exit Function
GetEmployeesForWGSchedRpt_EH:
    'Destroy Recordset
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForWGSchedRpt"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForWGSchedRpt" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForWGSchedRpt" _
                    , , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
    End Select

End Function
Public Function GetEmployeesForDailyRosterRpt(ByVal vlWGNum As Long _
                    , ByVal vdtRequestDate As Date) As ADODB.Recordset

On Error GoTo GetEmployeesForDailyRosterRpt_EH
'SQL Statement Container
Dim lsSQLStatement As String, ladrs As ADODB.Recordset

    
    lsSQLStatement = " SELECT E.EMP_NUM, E.EMP_NAME, EWG.EMP_NUM, EWG.WG_NUM " & _
                     " , EWG.EFF_START_DATE, EWG.EFF_END_DATE, WG.WG_NUM, DEMAND_IND " & _
                     " FROM METHDBA.TTR_EMP_WG EWG, " & _
                          " METHDBA.TTR_EMPLOYEE E,  " & _
                          " METHDBA.TTR_WORK_GROUP WG " & _
                     " WHERE EWG.WG_NUM = " & vlWGNum & _
                     " AND '" & Format(vdtRequestDate, "mm/dd/yyyy") & "'" & _
                     " BETWEEN EWG.EFF_START_DATE AND EWG.EFF_END_DATE " & _
                     " AND E.EMP_NUM = EWG.EMP_NUM " & _
                     " AND (E.EMP_STATUS = 'A' OR " & _
                     " (E.EMP_STATUS IN ('T','L') AND " & _
                     " E.STATUS_EFF_DATE > '" & Format(vdtRequestDate, "mm/dd/yyyy") & "'))" & _
                     " AND WG.WG_NUM = EWG.WG_NUM " & _
                     " UNION ALL " & _
                     " SELECT E.EMP_NUM, E.EMP_NAME, A.EMP_NUM, A.WG_NUM " & _
                     " , A.START_DATE, A.END_DATE, WG.WG_NUM, DEMAND_IND " & _
                     " FROM METHDBA.TTR_LINE_ASGNM A, " & _
                          " METHDBA.TTR_EMPLOYEE E, " & _
                          " METHDBA.TTR_WORK_GROUP WG " & _
                     " WHERE A.WG_NUM = " & vlWGNum & _
                     " AND '" & Format(vdtRequestDate, "mm/dd/yyyy") & "'" & _
                     " BETWEEN A.START_DATE AND A.END_DATE " & _
                     " AND E.EMP_NUM = A.EMP_NUM " & _
                     " AND (E.EMP_STATUS = 'A' OR "
                                                    
    lsSQLStatement = lsSQLStatement & " (E.EMP_STATUS IN ('T','L') AND " & _
                     " E.STATUS_EFF_DATE > '" & Format(vdtRequestDate, "mm/dd/yyyy") & "'))" & _
                     " AND WG.WG_NUM = A.WG_NUM " & _
                     " ORDER BY 4, 3"
                    
    'Call ExecuteSQL and retrieve information
    Set ladrs = ExecuteSQL(lsSQLStatement _
                            , adOpenStatic _
                            , adLockReadOnly)

    'Check if recordset is empty
    If (mADOExtensions.RecordsExist(ladrs)) Then
        Set GetEmployeesForDailyRosterRpt = ladrs
    Else
        Set GetEmployeesForDailyRosterRpt = Nothing
    End If

    'Destroy Recordset
    Set ladrs = Nothing

    Exit Function
GetEmployeesForDailyRosterRpt_EH:
    'Destroy Recordset
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForDailyRosterRpt"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForDailyRosterRpt" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForDailyRosterRpt" _
                    , , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
    End Select

End Function
Public Function GetEmployeesForDailyRosterRpt2(ByVal vlWGNum As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset

On Error GoTo GetEmployeesForDailyRosterRpt2_EH
'SQL Statement Container
Dim lsSQLStatement As String, ladrs As ADODB.Recordset

    
    lsSQLStatement = " SELECT E.EMP_NUM, E.EMP_NAME, EWG.EMP_NUM, EWG.WG_NUM " & _
                     " , EWG.EFF_START_DATE, EWG.EFF_END_DATE, WG.WG_NUM, DEMAND_IND " & _
                     " FROM METHDBA.TTR_EMP_WG EWG, " & _
                          " METHDBA.TTR_EMPLOYEE E,  " & _
                          " METHDBA.TTR_WORK_GROUP WG " & _
                     " WHERE EWG.WG_NUM = " & vlWGNum & _
                     " AND E.EMP_NUM = EWG.EMP_NUM " & _
                     " AND ('" & Format(vdtStartDate, "mm/dd/yyyy") & "' >= EWG.EFF_START_DATE " & _
                     " OR EWG.EFF_START_DATE BETWEEN '" & Format(vdtStartDate, "mm/dd/yyyy") & _
                     "' AND '" & Format(vdtEndDate, "mm/dd/yyyy") & "') AND ('" & _
                     Format(vdtEndDate, "mm/dd/yyyy") & "' <= EWG.EFF_END_DATE " & _
                     " OR EWG.EFF_END_DATE BETWEEN '" & Format(vdtStartDate, "mm/dd/yyyy") & _
                     "' AND '" & Format(vdtEndDate, "mm/dd/yyyy") & "')" & _
                     " AND (E.EMP_STATUS = 'A' OR " & _
                     " (E.EMP_STATUS IN ('T','L') AND " & _
                     " E.STATUS_EFF_DATE >= '" & Format(vdtEndDate, "mm/dd/yyyy") & "'))" & _
                     " AND WG.WG_NUM = EWG.WG_NUM " & _
                     " UNION ALL " & _
                     " SELECT E.EMP_NUM, E.EMP_NAME, A.EMP_NUM, A.WG_NUM " & _
                     " , A.START_DATE, A.END_DATE, WG.WG_NUM, DEMAND_IND " & _
                     " FROM METHDBA.TTR_LINE_ASGNM A, " & _
                          " METHDBA.TTR_EMPLOYEE E, " & _
                          " METHDBA.TTR_WORK_GROUP WG "

    lsSQLStatement = lsSQLStatement & " WHERE A.WG_NUM = " & vlWGNum & _
                     " AND ('" & Format(vdtStartDate, "mm/dd/yyyy") & "' >= A.START_DATE " & _
                     " OR A.START_DATE BETWEEN '" & Format(vdtStartDate, "mm/dd/yyyy") & "' AND '" & _
                     Format(vdtEndDate, "mm/dd/yyyy") & "') " & _
                     " AND ('" & Format(vdtEndDate, "mm/dd/yyyy") & "' <= A.END_DATE " & _
                     " OR A.END_DATE BETWEEN '" & Format(vdtStartDate, "mm/dd/yyyy") & "' AND '" & _
                     Format(vdtEndDate, "mm/dd/yyyy") & "')" & _
                     " AND E.EMP_NUM = A.EMP_NUM " & _
                     " AND (E.EMP_STATUS = 'A' OR " & _
                     " (E.EMP_STATUS IN ('T','L') AND " & _
                     " E.STATUS_EFF_DATE >= '" & Format(vdtEndDate, "mm/dd/yyyy") & "'))" & _
                     " AND WG.WG_NUM = A.WG_NUM " & _
                     " ORDER BY 4, 3"
                    
    'Call ExecuteSQL and retrieve information
    Set ladrs = ExecuteSQL(lsSQLStatement _
                            , adOpenStatic _
                            , adLockReadOnly)

    'Check if recordset is empty
    If (mADOExtensions.RecordsExist(ladrs)) Then
        Set GetEmployeesForDailyRosterRpt2 = ladrs
    Else
        Set GetEmployeesForDailyRosterRpt2 = Nothing
    End If

    'Destroy Recordset
    Set ladrs = Nothing

    Exit Function
GetEmployeesForDailyRosterRpt2_EH:
    'Destroy Recordset
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForDailyRosterRpt2"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForDailyRosterRpt2" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetEmployeesForDailyRosterRpt2" _
                    , , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
    End Select

End Function


Public Function GetActiveWorkgroups(ByVal vlngEmpNum As Long, ByVal vdtDate As Date) As Long
    
    On Error GoTo GetActiveWorkgroups_eh
    
    Dim lsSQLStatement As String, ladrs As ADODB.Recordset
    
    lsSQLStatement = "SELECT 1, '', A.LAST_MOD_TIMESTAMP, A.WG_NUM " & _
                     " FROM METHDBA.TTR_EMP_WG A, METHDBA.TTR_WORK_GROUP B " & _
                     " WHERE A.EMP_NUM = " & vlngEmpNum & _
                     " AND '" & Format("MM/DD/YYYY", vdtDate) & "'" & _
                     " BETWEEN A.EFF_START_DATE AND A.EFF_END_DATE " & _
                     " AND B.WG_NUM = A.WG_NUM AND B.PROJECT_ID > 0 " & _
                     " UNION " & _
                     "SELECT 2, '', A.LAST_MOD_TIMESTAMP, A.WG_NUM " & _
                     " FROM METHDBA.TTR_LINE_ASGNM A, METHDBA.TTR_WORK_GROUP B " & _
                     " WHERE A.EMP_NUM = " & vlngEmpNum & _
                     " AND '" & Format("MM/DD/YYYY", vdtDate) & "'" & _
                     " BETWEEN A.START_DATE AND A.END_DATE " & _
                     " AND B.WG_NUM = A.WG_NUM AND B.PROJECT_ID > 0 " & _
                     " UNION " & _
                     "SELECT 3, 'N', LAST_MOD_TIMESTAMP, WG_NUM " & _
                     " FROM METHDBA.TTR_LINE_ASGNM " & _
                     " WHERE EMP_NUM = " & vlngEmpNum & _
                     " AND '" & Format("MM/DD/YYYY", vdtDate) & "'" & _
                     " BETWEEN START_DATE AND END_DATE AND RELIEF_COMP = '    '" & _
                     " UNION " & _
                     "SELECT 4, 'N', LAST_MOD_TIMESTAMP, WG_NUM " & _
                     " FROM METHDBA.TTR_LINE_ASGNM " & _
                     " WHERE EMP_NUM = " & vlngEmpNum & _
                     " AND '" & Format("MM/DD/YYYY", vdtDate) & "'" & _
                     " BETWEEN START_DATE AND END_DATE AND RELIEF_COMP = '    '"
                     
    lsSQLStatement = lsSQLStatement & _
                      " UNION " & _
                      "SELECT 5, PRIMARY_IND, LAST_MOD_TIMESTAMP, WG_NUM " & _
                      " FROM METHDBA.TTR_EMP_WG " & _
                      " WHERE EMP_NUM = " & vlngEmpNum & _
                      " AND '" & Format("MM/DD/YYYY", vdtDate) & "'" & _
                      " BETWEEN EFF_START_DATE AND EFF_END_DATE " & _
                      " UNION " & _
                      "SELECT 6, 'N', A.LAST_MOD_TIMESTAMP, A.WG_NUM " & _
                      " FROM METHDBA.TTR_LINE_ASGNM A " & _
                      " WHERE A.EMP_NUM = " & vlngEmpNum & _
                      " AND '" & Format("MM/DD/YYYY", vdtDate) & "'" & _
                      " BETWEEN START_DATE AND END_DATE " & _
                      " AND EXISTS (SELECT 1 FROM METHDBA.TTR_EMP_WG B " & _
                      " WHERE B.EMP_NUM = A.EMP_NUM " & _
                      " AND B.EFF_START_DATE <= A.END_DATE " & _
                      " AND B.EFF_END_DATE >= A.START_DATE " & _
                      " AND B.PRIMARY_IND = 'Y' " & _
                      " AND B.WG_NUM <> A.WG_NUM) " & _
                      " ORDER BY 1,2,3 DESC "
                      
    'Call ExecuteSQL and retrieve information
    Set ladrs = ExecuteSQL(lsSQLStatement _
                            , adOpenStatic _
                            , adLockReadOnly)

    'Check if recordset is empty
    If (mADOExtensions.RecordsExist(ladrs)) Then
        GetActiveWorkgroups = ladrs!WG_NUM
    Else
        GetActiveWorkgroups = 0
    End If

    'Destroy Recordset
    Set ladrs = Nothing

    Exit Function
    
GetActiveWorkgroups_eh:
    'Destroy Recordset
    Set ladrs = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetACTIVEWORKGROUPS"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "Getactiveworkgroups" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "Getactiveworkgroups" _
                    , , , _
                    , DescArgsFrom(lsSQLStatement) _
                    , True
    End Select
End Function
Public Function IsJobTaskWorkgroup(ByVal v_lngWgNum As Long) As Boolean
    
'Does the workgroup use EMPRV JobTasks?
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo IsJobTaskWorkgroup_EH
    
    If v_lngWgNum < 1 Then
        Err.Raise trErrdbInvalidZerolngWgNum
    Else
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT COUNT(*) AS COUNT1 "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_DEPT D, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_WORK_GROUP WG, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_ASSOCIATIONS ASSOC "
    l_strSQL = l_strSQL & "WHERE WG.WG_NUM = " & v_lngWgNum
    l_strSQL = l_strSQL & "  AND WG.WG_NUM = ASSOC.WG_NUM "
    l_strSQL = l_strSQL & "  AND ASSOC.DEPT_NUM = D.DEPT_NUM "
    l_strSQL = l_strSQL & "  AND D.INTERFACE_IND IN ('E', 'B', 'M')  "
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    If Not l_objRS.BOF And Not l_objRS.EOF Then
        If l_objRS!COUNT1 > 0 Then
            IsJobTaskWorkgroup = True
        Else
            IsJobTaskWorkgroup = False
        End If
    Else
        IsJobTaskWorkgroup = False
    End If

    Set l_objRS = Nothing
    Exit Function

IsJobTaskWorkgroup_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "IsJobTaskWorkgroup", trErrdbInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "IsJobTaskWorkgroup", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function

Private Sub Class_Initialize()
    Set m_ResourceMgr = New utilADOHelper.cuResourceMgr
End Sub

Private Sub Class_Terminate()
    Set m_ResourceMgr = Nothing
End Sub
