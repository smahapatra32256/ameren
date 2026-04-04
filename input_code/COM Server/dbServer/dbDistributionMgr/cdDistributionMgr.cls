VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdDistributionMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdDistributionMgr"
Private Const m_DBSCRIPTSROOT As String = "C:\TRIS SOLUTION\dbScripts\TRIS_db2"
Private Const m_DBSCRIPTSPATH As String = "\dbDistributionMgr\cdDistributionMgr"
Private m_ResourceMgr As utilADOHelper.cuResourceMgr

'do not use this public enum - it must be deleted - begin
Public Enum TrsDistributionMgrErrorEnum

    trErrdb_A_GetConnectionString_F = 30008 + vbObjectError
    
    trErrdb_F_GetAccountDescriptionEmptyColWgNum_F = 30060 + vbObjectError
    trErrdb_G_GetDistributionListEmptyColWgNum_F ' = 30061 + vbObjectError
    trErrdb_GetAccountDescription_F          '= 30062 + vbObjectError
    trErrdb_GetDistributionList_F            ' = 30063 + vbObjectError
    trErrdb_GetDOJMWorkRequestsByRegion_F    ' = 30064 + vbObjectError
    trErrdb_IsDOJMDistrict_F                 ' = 30065 + vbObjectError
    trErrdb_ValidateServiceRequest_F         ' = 30066 + vbObjectError
    trErrdb_ValidateWorkRequest_F            ' = 30067 + vbObjectError

End Enum
'do not use this public enum - it must be deleted - end

'copied from util_DBConnectMgr to get DB2 subsystem
'Private Const regSubKey As String = "Software\Ameren\TRIS\DBConnectInfo\"
Private Const defaultProvider As String = "SHADOWDIRECT"
Public Function ValidateWorkRequest( _
            ByVal v_varDistributionId As Variant, _
            ByVal v_strDOJMHeadQuarters As String _
            ) As Boolean
'  Chect the VOMWRKRQ_00R table for the DOJM District and Work Request Number
' RED078
     
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo ValidateWorkRequest_EH
    
     
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT2")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
'    l_strSQL = "SELECT  CDE_WR "
'    l_strSQL = l_strSQL & "FROM METHDBA.VOMWRKRQ_00R "
'    l_strSQL = l_strSQL & "WHERE   ((CDE_STAT IN ('60','6X','70','7X','80','8X') AND CDE_TYP_WR = 'SI') "
'    l_strSQL = l_strSQL & "OR  (CDE_STAT IN ('60','6X','70','7X','80','8X', '90', '9X') AND CDE_TYP_WR = 'SR')) "
'    l_strSQL = l_strSQL & "AND CDE_DIST = '" & v_varDistributionId(0) & "' "
'    l_strSQL = l_strSQL & "AND CDE_WR = '" & v_varDistributionId(2) & "' "
'    l_strSQL = l_strSQL & "AND CDE_HDQTR_CRW like '" & v_strDOJMHeadQuarters & "%'"
        
'SELECT COUNT(*)
'From SYSIBM.SYSDUMMY1
'WHERE EXISTS(SELECT 1
'             From METHDBA.TOMWRKRQ
'             WHERE CDE_DIST = '2DLD'
'             AND CDE_HDQTR_CRW LIKE '08R%'
'             AND ((CDE_STAT IN ('60','6X','70','7X','80','8X') AND CDE_TYP_WR = 'SI')
'             OR  (CDE_STAT IN ('60','6X','70','7X','80','8X', '90', '9X') AND CDE_TYP_WR = 'SR'))
'             AND CDE_WR = '006037')
    
    l_strSQL = "SELECT COUNT(*) AS COUNT1"
    l_strSQL = l_strSQL & " FROM SYSIBM.SYSDUMMY1"
    l_strSQL = l_strSQL & " WHERE EXISTS(SELECT 1"
    l_strSQL = l_strSQL & "  FROM METHDBA.TOMWRKRQ"
    l_strSQL = l_strSQL & "  WHERE CDE_DIST = '" & v_varDistributionId(0) & "'"
    l_strSQL = l_strSQL & "   AND CDE_HDQTR_CRW LIKE '" & v_strDOJMHeadQuarters & "%'"
    l_strSQL = l_strSQL & "   AND ((CDE_STAT IN ('60','6X','70','7X','80','8X') AND CDE_TYP_WR = 'SI')"
    l_strSQL = l_strSQL & "   OR  (CDE_STAT IN ('60','6X','70','7X','80','8X', '90', '9X') AND CDE_TYP_WR = 'SR'))"
    l_strSQL = l_strSQL & "   AND CDE_WR = '" & v_varDistributionId(2) & "')"
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    If Not l_objrs.BOF And Not l_objrs.EOF Then
        If l_objrs!COUNT1 > 0 Then
            ValidateWorkRequest = True
        Else
            ValidateWorkRequest = False
        End If
    Else
        ValidateWorkRequest = False
    End If
        
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
Exit Function

ValidateWorkRequest_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "ValidateWorkRequest", , , , DescArgsFrom(l_strSQL), True
    
End Function

Public Function ValidateServiceRequest( _
                    ByVal v_strServiceReq As String, _
                    ByVal v_strCorpCode As String, _
                    ByRef v_strResponse As String) As Boolean
TrsRaiseError m_strCMODULENAME, "ValidateServiceRequest", trErrObsoleteMethod, , , , True
'    ' RED077
'
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'    Dim l_objCMD As ADODB.Command
'    Dim l_objPCICSRegion As ADODB.Parameter
'    Dim l_objPAccountNum As ADODB.Parameter
'    Dim l_objPCorpCode As ADODB.Parameter
'    Dim l_objPSystemCode As ADODB.Parameter
'    Dim l_objPReturnMessage As ADODB.Parameter
'    Dim l_objPErrorFlags As ADODB.Parameter
'    Dim l_objPSqlCode As ADODB.Parameter
'    Dim l_objPSuccess As ADODB.Parameter
'    Dim l_intSQLCODE As Integer
'    Dim l_strSuccessFailure As String
'    Dim l_strReturnMessage As String
'    Dim l_strErrorFlags As String
'    Dim l_strFirst As String * 16
'    Dim l_strServiceReq As String * 21
'    Dim l_strLast As String * 8
'
'On Error GoTo ValidateServiceRequest_EH
'
'
'    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
'    Set l_objCN = New ADODB.Connection
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'    l_objCN.CursorLocation = adUseClient
'    Set l_objCM = Nothing
'
'    l_objCN.Open
'
'    Set l_objCMD = New ADODB.Command
'
'    Set l_objPCICSRegion = l_objCMD.CreateParameter("aSp_Target_System", adChar, adParamInput, 8)
'    l_objPCICSRegion.Value = "SY7221"
'
'    Set l_objPAccountNum = l_objCMD.CreateParameter("aSp_Distr_Id", adChar, adParamInput, 29)
'    l_objPAccountNum.Value = v_strServiceReq
'
'    Set l_objPCorpCode = l_objCMD.CreateParameter("aSp_Cmpy_Cd", adChar, adParamInput, 3)
'    l_objPCorpCode.Value = v_strCorpCode
'
'    Set l_objPSystemCode = l_objCMD.CreateParameter("aSp_System_Id", adChar, adParamInput, 2)
'    l_objPSystemCode.Value = "TR"
'
'    Set l_objPSqlCode = l_objCMD.CreateParameter("aSp_CICS_Return_CD", adInteger, adParamOutput)
'    l_objPSqlCode.Value = l_intSQLCODE
'
'    Set l_objPReturnMessage = l_objCMD.CreateParameter("aSp_Error_Msg", adChar, adParamOutput, 50)
'    l_objPReturnMessage.Value = l_strReturnMessage
'
'    Set l_objPErrorFlags = l_objCMD.CreateParameter("aSp_Error_Msg", adChar, adParamOutput, 10)
'    l_objPErrorFlags.Value = l_strErrorFlags
'
'    Set l_objPSuccess = l_objCMD.CreateParameter("aSp_Succ_Fail_Flag", adChar, adParamOutput, 1)
'    l_objPSuccess.Value = l_strSuccessFailure
'
'    With l_objCMD
'        .ActiveConnection = l_objCN
'        .CommandType = adCmdStoredProc
'        .CommandText = "SYSPROC.TSTGA030"   ' This is going to change to TSTGA030 (There will also be some other changes)
'        .Parameters.Append l_objPCICSRegion
'        .Parameters.Append l_objPAccountNum
'        .Parameters.Append l_objPCorpCode
'        .Parameters.Append l_objPSystemCode
'        .Parameters.Append l_objPSqlCode
'        .Parameters.Append l_objPReturnMessage
'        .Parameters.Append l_objPErrorFlags
'        .Parameters.Append l_objPSuccess
'    End With
'
'    l_objCMD.Execute
'
' ' check for Null here since problem with Neon and second call
'     If IsNull(l_objPReturnMessage.Value) Then
'        v_strResponse = "Null"
'    Else
'        v_strResponse = l_objPReturnMessage.Value
'    End If
'
'    If l_objPSuccess = "S" Then
'        ValidateServiceRequest = True
'    Else
'        ValidateServiceRequest = False
'    End If
'
'    Set l_objCN = Nothing
'    Set l_objCMD = Nothing
'
'    Set l_objPCICSRegion = Nothing
'    Set l_objPAccountNum = Nothing
'    Set l_objPCorpCode = Nothing
'    Set l_objPSystemCode = Nothing
'    Set l_objPReturnMessage = Nothing
'    Set l_objPErrorFlags = Nothing
'    Set l_objPSqlCode = Nothing
'    Set l_objPSuccess = Nothing
'
'    Exit Function
'
'ValidateServiceRequest_EH:
'
'    Set l_objCN = Nothing
'    Set l_objCMD = Nothing
'
'    Set l_objPCICSRegion = Nothing
'    Set l_objPAccountNum = Nothing
'    Set l_objPCorpCode = Nothing
'    Set l_objPSystemCode = Nothing
'    Set l_objPReturnMessage = Nothing
'    Set l_objPErrorFlags = Nothing
'    Set l_objPSqlCode = Nothing
'    Set l_objPSuccess = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "ValidateServiceRequest", , , , , True
    
End Function

Public Function GetAccountDescription( _
                    ByVal v_colWgNum As Collection) As ADODB.Recordset
    
     
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetAccountDescription_EH
    
    Dim l_strWgNumforSQL As String
     
    If v_colWgNum.Count < 1 Then
        l_strWgNumforSQL = "(0)"
        Err.Raise trErrdbGetAccountDescriptionEmptyColWgNum
    Else
        l_strWgNumforSQL = AsDelimitedString(v_colWgNum, False)
    End If
    
   
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT WG_NUM, DISTRIBUTION_ID, DESCRIPTION From METHDBA.TTR_ACCT_ASSIGNMTS  "
    l_strSQL = l_strSQL & "Where WG_NUM in " & l_strWgNumforSQL
    l_strSQL = l_strSQL & " group by wg_num, distribution_id,description "
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetAccountDescription = l_objrs
       
    Set l_objrs = Nothing
    
    Exit Function

GetAccountDescription_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    Select Case Err.Number
       Case trErrdbGetAccountDescriptionEmptyColWgNum
            TrsRaiseError m_strCMODULENAME, "GetAccountDescription", trErrdbGetAccountDescriptionEmptyColWgNum, , , , True
       Case Else
            TrsRaiseError m_strCMODULENAME, "GetAccountDescription", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function

Public Function IsDOJMDistrict( _
        ByVal v_varDistributionId As Variant) As Boolean
'  Is the Distribution Id a Valid DOJM District

     
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo IsDOJMDistrict_EH
    
     
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT2")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT CDE_DIST  "
    l_strSQL = l_strSQL & "FROM METHDBA.TOMCRWHDQTR "
    l_strSQL = l_strSQL & "WHERE CDE_DIST = '" & v_varDistributionId(0) & "'"
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    If Not l_objrs.BOF And Not l_objrs.EOF Then
        IsDOJMDistrict = True
    Else
        IsDOJMDistrict = False
    End If
    
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    

Exit Function

IsDOJMDistrict_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "IsDOJMDistrict", , , , DescArgsFrom(l_strSQL), True

End Function
        
Public Function GetDOJMWorkRequestsByRegion( _
                    ByVal v_strDistrict As String, _
                    ByVal v_strHeadquarters As String) _
                        As ADODB.Recordset
     
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetDOJMWorkRequestsByRegion_EH
     
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT2")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

'    l_strSQL = "SELECT '0' AS WG_NUM, "
'    l_strSQL = l_strSQL & "CDE_WR AS DISTRIBUTION_ID, "
'    l_strSQL = l_strSQL & "ADDR_HS_JOB||' '|| ADDR_ST_JOB ||' '|| ADDR_TO_JOB AS DESCRIPTION "
'    l_strSQL = l_strSQL & "FROM METHDBA.VOMWRKRQ_00R WHERE CDE_DIST = '" & v_strDistrict & "'"
'    l_strSQL = l_strSQL & "AND CDE_HDQTR_CRW LIKE '" & v_strHeadquarters & "%'"
'    l_strSQL = l_strSQL & " AND ((CDE_STAT IN ('60', '6X', '70', '7X', '80', '8X') AND CDE_TYP_WR = 'SI') "
'    l_strSQL = l_strSQL & " OR   (CDE_STAT IN ('60', '6X', '70', '7X', '80', '8X', '90', '9X') and CDE_TYP_WR = 'SR')) "
'    l_strSQL = l_strSQL & " ORDER BY DISTRIBUTION_ID"

    l_strSQL = "SELECT WG_NUM, DISTRIBUTION_ID, DESCRIPTION, "
    l_strSQL = l_strSQL & " DISTRIBUTION_ID AS JOB_TASK, CAST('' AS VARCHAR(8)) AS EMPRV_DATABASE_ID, "
    l_strSQL = l_strSQL & " CAST('' AS VARCHAR(10)) AS FEEDER_REF, CAST(0 AS INTEGER) AS LAST_MOD_EMP_NUM, CAST('' AS VARCHAR(26)) AS TIME_STAMP, "
    l_strSQL = l_strSQL & " CAST('' AS CHAR(19)) AS TRIS_PROJECT_AND_TASK_IDS, CAST('' AS VARCHAR(10)) AS PROJECT_NUMBER, "
    l_strSQL = l_strSQL & " CAST('' AS VARCHAR(60)) AS PROJECT_NAME, CAST('' AS VARCHAR(21)) AS TASK_NUMBER, CAST('' AS VARCHAR(60)) AS TASK_NAME "
    l_strSQL = l_strSQL & "FROM (SELECT CAST(0 AS INTEGER) AS WG_NUM, "
    l_strSQL = l_strSQL & "CAST(NE1.CDE_WR AS VARCHAR(60)) AS DISTRIBUTION_ID, "
    l_strSQL = l_strSQL & "CAST(LEFT(NE2.ADDR_HS_JOB||' '|| REPLACE(NE2.ADDR_ST_JOB, x'00', SPACE(1)) ||' '|| NE2.ADDR_TO_JOB, 60) AS VARCHAR(60)) AS DESCRIPTION "
    l_strSQL = l_strSQL & "FROM (SELECT CDE_DIST, CDE_WR "
    l_strSQL = l_strSQL & "  FROM METHDBA.TOMWRKRQ "
    l_strSQL = l_strSQL & "  WHERE CDE_DIST = '" & v_strDistrict & "' "
    l_strSQL = l_strSQL & "   AND CDE_HDQTR_CRW LIKE '" & v_strHeadquarters & "%' "
    l_strSQL = l_strSQL & "   AND ((CDE_STAT IN ('60', '6X', '70', '7X', '80', '8X') AND CDE_TYP_WR = 'SI') "
    l_strSQL = l_strSQL & "   OR   (CDE_STAT IN ('60', '6X', '70', '7X', '80', '8X', '90', '9X') and CDE_TYP_WR = 'SR')) "
    l_strSQL = l_strSQL & "  GROUP BY CDE_DIST, CDE_WR) NE1 "
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TOMJOBADDR NE2 "
    l_strSQL = l_strSQL & " ON NE2.CDE_DIST = NE1.CDE_DIST "
    l_strSQL = l_strSQL & " AND NE2.CDE_WR = NE1.CDE_WR ) NE3 "
    l_strSQL = l_strSQL & "GROUP BY DISTRIBUTION_ID, WG_NUM, DESCRIPTION "
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetDOJMWorkRequestsByRegion = l_objrs
        
    Set l_objrs = Nothing
    
    Exit Function

GetDOJMWorkRequestsByRegion_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDOJMWorkRequestsByRegion", , , , DescArgsFrom(l_strSQL), True
   
End Function
        
'Public Function IsPOETActive() As Boolean
'
'    Dim l_objrsSysParms As ADODB.Recordset
'    Dim l_strSQLSysParms As String
'    Dim oADOHelper As utilADOHelper.cuADOHelper
'
'    On Error GoTo eh
'
'    Set oADOHelper = New utilADOHelper.cuADOHelper
'
'    l_strSQLSysParms = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "CheckForOTTC.sql")
'
'    Set l_objrsSysParms = oADOHelper.ExecuteSelect(l_strSQLSysParms, adOpenStatic, adLockBatchOptimistic)
'
'    IsPOETActive = False
'    If RecordsExist(l_objrsSysParms) Then
'        l_objrsSysParms.MoveFirst
'        If l_objrsSysParms.Fields("USE_OTTC").Value = "Y" Then
'            IsPOETActive = True
'        End If
'    End If
'
'    Set l_objrsSysParms = Nothing
'    Set oADOHelper = Nothing
'
'    Exit Function
'
'eh:
'    Set l_objrsSysParms = Nothing
'    Set oADOHelper = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "IsPOETActive", , , , DescArgsFrom(l_strSQLSysParms), True
'End Function
        
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

Public Function GetDistributionList( _
                    ByVal v_colWgNum As Collection) As ADODB.Recordset
    On Error GoTo eh
    
    Set GetDistributionList = GetDistributionList2B(v_colWgNum, 30)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistributionList", , , , v_colWgNum, True
End Function

Public Function GetDistributionList2B( _
                    ByVal v_colWgNum As Collection, _
                    ByVal v_intDescriptionLength As Integer) As ADODB.Recordset
    On Error GoTo eh
    
    Set GetDistributionList2B = GetDistributionList3(v_colWgNum, v_intDescriptionLength, "unknown")
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistributionList2B", , , , v_colWgNum, True
End Function

Public Function GetDistributionList3( _
                    ByVal v_colWgNum As Collection, _
                    ByVal v_intDescriptionLength As Integer, _
                    ByVal v_sCallerMode As String) As ADODB.Recordset
    
'    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
'    Dim l_objrsSysParms As ADODB.Recordset
    Dim l_strSQL As String
'    Dim l_strSQLSysParms As String
    Dim oADOHelper As utilADOHelper.cuADOHelper
    Dim UseOTTC As Boolean
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim sCallerModeWhereClause As String
    
    On Error GoTo eh
    
    Dim l_strWgNumforSQL As String
     
    If v_colWgNum.Count < 1 Then
        l_strWgNumforSQL = "(0)"
        Err.Raise trErrdbGetDistributionListEmptyColWgNum
    Else
        l_strWgNumforSQL = AsDelimitedString(v_colWgNum, False)
    End If
    
    Set oADOHelper = New utilADOHelper.cuADOHelper
    
'    l_strSQLSysParms = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "CheckForOTTC.sql")
'
'    Set l_objrsSysParms = oADOHelper.ExecuteSelect(l_strSQLSysParms, adOpenStatic, adLockBatchOptimistic)
'
'    UseOTTC = False
'    If RecordsExist(l_objrsSysParms) Then
'        l_objrsSysParms.MoveFirst
'        If l_objrsSysParms.Fields("USE_OTTC").Value = "Y" Then
'            UseOTTC = True
'        End If
'    End If
    
    If IsPOETActive Then
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetDistributionListOTTC2.sql")
    Else
        l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetDistributionList.sql")
    End If
    
    l_strSQL = Replace(l_strSQL, "{WgNumInClause}", l_strWgNumforSQL)
    l_strSQL = Replace(l_strSQL, "{DescriptionLength}", v_intDescriptionLength)
    '//allow time collection refresh to merge both GL Code Block Accounting and POET Accounting (only during transition IsOTTCTransitionActive)
    '//won't merge for PickList or LoadPriorDistribution, they will use the IsOTTCEffective sysparm
    '//l_strSQL = Replace(l_strSQL, "{ForceMergeFlag}", IIf(v_sCallerMode = "ActualTimeDesc" Or v_sCallerMode = "LoadPriorPayPeriodDesc", "1=1", "0=1"))
    If v_sCallerMode = "PickList" Then
        sCallerModeWhereClause = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetDistributionListOTTC2_WHERE_PickList.sql")
'        sCallerModeWhereClause = _
'"--picklist, allow 30 days past project finish date " & vbCrLf & _
'"AND CONVERSION_STATUS_CODE = 'P' " & vbCrLf & _
'"AND ACTIVE_STATUS_CODE = 'Y' " & vbCrLf & _
'"--AND (DATE(PM.PROJECT_FINISH_DATE) >= CURRENT DATE - 30 DAYS " & vbCrLf & _
'"AND (DATE(PM.PROJECT_FINISH_DATE) >= CURRENT DATE - (SELECT INT(PARM_VALUE) FROM METHDBA.TTR_SYSPARMS WHERE PARM_NUM = 56 AND PARM_SUFFIX = 1) DAYS " & vbCrLf & _
'"OR PM.PROJECT_FINISH_DATE IS NULL OR PM.PROJECT_FINISH_DATE = '0001-01-01 00:00:00.000000') " & vbCrLf & _
'"AND TM.LABOR_FLAG = 'Y'"
    ElseIf v_sCallerMode = "LoadPriorPayPeriodDesc" Then
        sCallerModeWhereClause = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetDistributionListOTTC2_WHERE_LoadPriorPayPeriodDesc.sql")
'        sCallerModeWhereClause = _
'"--load prior pay period, project finish date is in the future " & vbCrLf & _
'"AND CONVERSION_STATUS_CODE = 'P' " & vbCrLf & _
'"AND ACTIVE_STATUS_CODE = 'Y' " & vbCrLf & _
'"AND (DATE(PM.PROJECT_FINISH_DATE) >= CURRENT DATE - 0 DAYS " & vbCrLf & _
'"OR PM.PROJECT_FINISH_DATE IS NULL OR PM.PROJECT_FINISH_DATE = '0001-01-01 00:00:00.000000') " & vbCrLf & _
'"AND TM.LABOR_FLAG = 'Y'"
    ElseIf v_sCallerMode = "ActualTimeDesc" Then
        sCallerModeWhereClause = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetDistributionListOTTC2_WHERE_ActualTimeDesc.sql")
'        sCallerModeWhereClause = _
'"--ActualTimeDesc" & vbCrLf & _
'"AND 1=1"
    Else
        sCallerModeWhereClause = _
"--UNKNOWN" & vbCrLf & _
"AND 1=1"
    End If
    
    l_strSQL = Replace(l_strSQL, "{CallerModeWhereClause}", sCallerModeWhereClause)
    
    Set l_objrs = oADOHelper.ExecuteSelect(l_strSQL, adOpenStatic, adLockBatchOptimistic)
       
'    Set l_objCN = New ADODB.Connection
'    Set l_objrs = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
'
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open

'    l_strSQL = "SELECT WG_NUM, DISTRIBUTION_ID, DESCRIPTION, LAST_MOD_EMP_NUM, "
'    l_strSQL = l_strSQL & " ' ' AS JOB_TASK, ' ' AS EMPRV_DATABASE_ID, "
'    l_strSQL = l_strSQL & " CHAR(TIME_STAMP) AS TIME_STAMP FROM METHDBA.TTR_ACCT_ASSIGNMTS  "
'    l_strSQL = l_strSQL & "WHERE WG_NUM IN " & l_strWgNumforSQL
'    l_strSQL = l_strSQL & " ORDER BY DESCRIPTION "
'
'    l_objrs.CursorLocation = adUseClient
'    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objrs.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing

    Set GetDistributionList3 = l_objrs
    
    Set l_objrs = Nothing
'    Set l_objrsSysParms = Nothing
    Set oADOHelper = Nothing

'//TEST ONLY
'//    TrsTraceError m_strCMODULENAME, "GetDistributionList3", , False, , DescArgsFrom("sCallerModeWhereClause=""" & sCallerModeWhereClause & """" & vbCrLf, l_strSQL), True

    Exit Function

eh:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
    Set l_objrs = Nothing
'    Set l_objrsSysParms = Nothing
    Set oADOHelper = Nothing
    
    Select Case Err.Number
       Case trErrdbGetDistributionListEmptyColWgNum
            TrsRaiseError m_strCMODULENAME, "GetDistributionList3", trErrdbGetDistributionListEmptyColWgNum, , , , True
       Case Else
            TrsRaiseError m_strCMODULENAME, "GetDistributionList3", , , , DescArgsFrom(l_strSQL), True
    End Select
End Function

Public Function GetPOETMasterFromTRISIds(TRISProjectId As String, TRISTaskId As String) As ADODB.Recordset
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim oADOHelper As utilADOHelper.cuADOHelper
    
    On Error GoTo eh
        
    Set oADOHelper = New utilADOHelper.cuADOHelper
      
    l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetPOETMasterFromTRISIds.sql")
    
    oADOHelper.ADO.CreateCharacterParameter "TRISTaskId", adVarChar, 4, TRISTaskId
    oADOHelper.ADO.CreateCharacterParameter "TRISProjectId", adVarChar, 5, TRISProjectId

    Set l_objrs = oADOHelper.ExecuteSelect(l_strSQL, adOpenStatic, adLockBatchOptimistic)
       
    Set GetPOETMasterFromTRISIds = l_objrs
    
    Set l_objrs = Nothing
    Set oADOHelper = Nothing

    Exit Function

eh:
    Set l_objrs = Nothing
    Set oADOHelper = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetPOETMasterFromTRISIds", , , , DescArgsFrom(l_strSQL), True
End Function

'//called by busDistributionMgr (GetAccountDescriptionJobTask, GetDistributionListEMPRV) from WG maint. and time collection to get the pick lists
Public Function GetDistJobTaskList( _
                    ByVal v_colWgNum As Collection, _
                    ByVal v_strDeptNum As String) As ADODB.Recordset
On Error GoTo eh
    Dim rs As ADODB.Recordset
    
    Set rs = GetDistJobTaskList2(v_colWgNum, v_strDeptNum, False)
    
    Set GetDistJobTaskList = rs
    
    Set rs = Nothing
    Exit Function
eh:
    Set rs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDistJobTaskList"
End Function

'//called by both WG get pick lists and view calendar details, but the latter only needs EMPRV as Maximo will be retrieved through separate calls to get the description
Public Function GetDistJobTaskList2( _
                    ByVal v_colWgNum As Collection, _
                    ByVal v_strDeptNum As String, _
                    ByVal v_blnEMPRVOnly As Boolean) As ADODB.Recordset
         
    Dim l_objCN As New ADODB.Connection
    Dim l_objCMD As New Command
    Dim l_WGNum As New ADODB.Parameter
    Dim l_DeptNum As New ADODB.Parameter
    Dim l_ReturnValue As New ADODB.Parameter
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_rsGetDistJobTaskList As New ADODB.Recordset
    Dim l_objDept As dbDepartment.cdDepartment
    Dim l_strInterfaceInd As String
    Dim l_objrsInterfaceInd As ADODB.Recordset
    Dim l_objrsMAXIMO As ADODB.Recordset
    Dim l_lngWGNum As Long
    
    On Error GoTo GetDistJobTaskList_EH
    
    Dim l_strWgNumforSQL As String
    Dim l_lngWGNumforSQL As Long
    
    Set l_objDept = CreateObject("dbDepartment.cdDepartment")
    Set l_objrsInterfaceInd = l_objDept.GetInterfaceInd(v_strDeptNum)
    If l_objrsInterfaceInd.BOF And l_objrsInterfaceInd.EOF Then
        l_strInterfaceInd = "E"
    Else
        l_objrsInterfaceInd.MoveFirst
        l_strInterfaceInd = l_objrsInterfaceInd!INTERFACE_IND
    End If
    
     
    If v_colWgNum.Count < 1 Then
        l_strWgNumforSQL = "(0)"
        Err.Raise trErrdbGetDistributionListEmptyColWgNum
    Else
        l_lngWGNum = v_colWgNum.Item(1)
        l_strWgNumforSQL = AsDelimitedString(v_colWgNum, False, ",", "")
    End If
    
    l_lngWGNumforSQL = CLng(l_strWgNumforSQL)
    
    If (l_strInterfaceInd = "E" Or l_strInterfaceInd = "B") Then
        Set l_objCN = New ADODB.Connection
        Set l_objCMD = New ADODB.Command
        Set l_ReturnValue = l_objCMD.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue)
        Set l_WGNum = l_objCMD.CreateParameter("@wrkgrpnbr", adInteger, adParamInput, , l_lngWGNumforSQL)
        Set l_DeptNum = l_objCMD.CreateParameter("@payrolldept", adChar, adParamInput, 12, v_strDeptNum)
        Set l_objrs = New ADODB.Recordset
        Set l_objCM = New clsuDBConnectMgr
       
        l_objCN.ConnectionString = l_objCM.GetConnectionString("EMPRV")
        l_objCN.CursorLocation = adUseClient
        
        '//add a string value named CommandTimeout to the registry under HKEY_LOCAL_MACHINE\SOFTWARE\Ameren\TRIS\DBConnectInfo\EMPRV
        l_objCN.CommandTimeout = l_objCM.GetCommandTimeout("EMPRV") '//120
        
        Set l_objCM = Nothing
        
        l_objCN.Open
       
        With l_objCMD
            .Parameters.Append l_ReturnValue
            .Parameters.Append l_WGNum
            .Parameters.Append l_DeptNum
            .ActiveConnection = l_objCN
            .CommandType = adCmdStoredProc
            .CommandText = "dbo.po1382_sp_tris_joblist"
            .CommandTimeout = l_objCN.CommandTimeout
            
            '//debug only for use in case of error
            l_strSQL = .CommandText
            l_strSQL = Replace(l_strSQL, "? = ", "", 1, 1)
            l_strSQL = Replace(l_strSQL, "?", l_WGNum.Value, 1, 1)
            l_strSQL = Replace(l_strSQL, "?", "'" & l_DeptNum.Value & "'", 1, 1)
            '& " " & l_WGNum.Name & " = " & l_WGNum.Value & " " & l_DeptNum.Name & " " & l_DeptNum.Value
            'l_strSQL = .CommandText & " @wrkgrpnbr = " & l_WGNum & " @payrolldept = " & l_DeptNum
        End With
        
        l_objrs.Open l_objCMD
        
        If l_ReturnValue >= 0 Then
            With l_rsGetDistJobTaskList
                .Fields.Append "JOB_NBR", adVarChar, 8, adFldIsNullable
                .Fields.Append "TASK_NBR", adVarChar, 4, adFldIsNullable
                .Fields.Append "FK_EMPRVDBCODE", adChar, 8, adFldIsNullable
                .Fields.Append "DESC60", adVarChar, 60, adFldIsNullable
                .Fields.Append "JOBTASK_AND_DESC", adVarChar, 60, adFldIsNullable
                .Fields.Append "CORP", adChar, 3, adFldIsNullable
                .Fields.Append "UTLTY_CD", adChar, 1, adFldIsNullable
                .Fields.Append "BUS_DV_CD", adChar, 2, adFldIsNullable
                .Fields.Append "MAJ_CD", adChar, 3, adFldIsNullable
                .Fields.Append "MIN_CD", adChar, 3, adFldIsNullable
                .Fields.Append "FMC_CD", adChar, 3, adFldIsNullable
                .Fields.Append "RMC_CD", adChar, 3, adFldIsNullable
                .Fields.Append "TRN_TP_CD", adChar, 1, adFldIsNullable
                .Fields.Append "PROJECT_ID", adChar, 5, adFldIsNullable
                .Fields.Append "PROD_CD", adChar, 2, adFldIsNullable
                .Fields.Append "ACTIV_CD", adChar, 4, adFldIsNullable
                .Fields.Append "RES_TP_CD", adChar, 2, adFldIsNullable
                .Fields.Append "WG_NUM", adInteger, , adFldIsNullable
                .Fields.Append "COSTCNTRCODE", adVarChar, 12, adFldIsNullable
                .Fields.Append "CCDESC", adVarChar, 30, adFldIsNullable
                .Fields.Append "JOBSTCODE", adVarChar, 10, adFldIsNullable
                .Fields.Append "FEEDER_REF", adChar, 10, adFldIsNullable
                .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, adFldIsNullable
                
                .Fields.Append "TRIS_PROJECT_AND_TASK_IDS", adChar, 19, adFldIsNullable
                .Fields.Append "PROJECT_NUMBER", adVarChar, 10, adFldIsNullable
                .Fields.Append "PROJECT_NAME", adVarChar, 60, adFldIsNullable
                .Fields.Append "TASK_NUMBER", adVarChar, 21, adFldIsNullable
                .Fields.Append "TASK_NAME", adVarChar, 60, adFldIsNullable
            
                .CursorLocation = adUseClient
                .LockType = adLockBatchOptimistic
                .CursorType = adOpenStatic
                .Open
                .ActiveConnection = Nothing
                
                If Not l_objrs Is Nothing Then
                    If Not (l_objrs.BOF And l_objrs.EOF) Then
                        While Not l_objrs.EOF
                            .AddNew
                            !JOB_NBR = l_objrs.Fields("jobnbr").Value
                            !TASK_NBR = l_objrs.Fields("tasknbr").Value
                            !FK_EMPRVDBCODE = l_objrs.Fields("fk_emprvdbcode").Value
                            '//SetNullToEmpty
                            If Not IsNull(l_objrs.Fields("desc60").Value) Then
                                !DESC60 = l_objrs.Fields("desc60").Value
                            Else
                                !DESC60 = String(60, " ")
                            End If
                            !JOBTASK_AND_DESC = String(60, " ")
                            !CORP = l_objrs.Fields("corp").Value
                            !UTLTY_CD = l_objrs.Fields("utlty_cd").Value
                            !BUS_DV_CD = l_objrs.Fields("bus_dv_cd").Value
                            !MAJ_CD = l_objrs.Fields("maj_cd").Value
                            !MIN_CD = l_objrs.Fields("min_cd").Value
                            !FMC_CD = l_objrs.Fields("fmc_cd").Value
                            !RMC_CD = l_objrs.Fields("rmc_Cd").Value
                            !TRN_TP_CD = l_objrs.Fields("trn_tp_cd").Value
    
                            If Not IsNull(l_objrs.Fields("project_id").Value) Then
                                If l_objrs.Fields("project_id").ActualSize <= 5 Then
                                    !PROJECT_ID = l_objrs.Fields("project_id").Value
                                Else
                                    !PROJECT_ID = "     "
                                End If
                            Else
                                !PROJECT_ID = "     "
                            End If
                            
                            !PROD_CD = l_objrs.Fields("prod_cd").Value
                            !ACTIV_CD = l_objrs.Fields("activ_cd").Value
                            !RES_TP_CD = l_objrs.Fields("res_tp_cd").Value
                            !WG_NUM = l_objrs.Fields("wrkgrpnbr").Value
                            !COSTCNTRCODE = l_objrs.Fields("costcntrcode").Value
                            !CCDESC = l_objrs.Fields("ccdesc").Value
                            !JOBSTCODE = l_objrs.Fields("jobstcode").Value
                            !FEEDER_REF = Space(10)
                            !DISTRIBUTION_TYPE = "J"
                            
                            !TRIS_PROJECT_AND_TASK_IDS = ""
                            !PROJECT_NUMBER = l_objrs.Fields("project_id").Value
                            !PROJECT_NAME = ""
                            !TASK_NUMBER = ""
                            !TASK_NAME = ""
                            
                            .Update
                            l_objrs.MoveNext
                        Wend
                    End If
                End If
            End With
        End If
    End If
    
    If Not v_blnEMPRVOnly Then
        '//debugging only
        l_strSQL = "GetMAXIMOJobNumbers('" & v_strDeptNum & "', " & l_lngWGNum & ")"
        If (l_strInterfaceInd = "M" Or l_strInterfaceInd = "B") Then
           
            Set l_objrsMAXIMO = GetMAXIMOJobNumbers(v_strDeptNum, l_lngWGNum)
            If Not l_objrsMAXIMO Is Nothing Then
                If Not l_objrsMAXIMO.BOF And Not l_objrsMAXIMO.EOF Then
                    If l_strInterfaceInd = "M" Or l_rsGetDistJobTaskList.Fields.Count < 1 Then '//not build from EMPRV picklist db call
                        With l_rsGetDistJobTaskList
                            .Fields.Append "JOB_NBR", adVarChar, 8, adFldIsNullable
                            .Fields.Append "TASK_NBR", adVarChar, 4, adFldIsNullable
                            .Fields.Append "FK_EMPRVDBCODE", adChar, 8, adFldIsNullable
                            .Fields.Append "DESC60", adVarChar, 60, adFldIsNullable
                            .Fields.Append "JOBTASK_AND_DESC", adVarChar, 60, adFldIsNullable
                            .Fields.Append "CORP", adChar, 3, adFldIsNullable
                            .Fields.Append "UTLTY_CD", adChar, 1, adFldIsNullable
                            .Fields.Append "BUS_DV_CD", adChar, 2, adFldIsNullable
                            .Fields.Append "MAJ_CD", adChar, 3, adFldIsNullable
                            .Fields.Append "MIN_CD", adChar, 3, adFldIsNullable
                            .Fields.Append "FMC_CD", adChar, 3, adFldIsNullable
                            .Fields.Append "RMC_CD", adChar, 3, adFldIsNullable
                            .Fields.Append "TRN_TP_CD", adChar, 1, adFldIsNullable
                            .Fields.Append "PROJECT_ID", adChar, 5, adFldIsNullable
                            .Fields.Append "PROD_CD", adChar, 2, adFldIsNullable
                            .Fields.Append "ACTIV_CD", adChar, 4, adFldIsNullable
                            .Fields.Append "RES_TP_CD", adChar, 2, adFldIsNullable
                            .Fields.Append "WG_NUM", adInteger, , adFldIsNullable
                            .Fields.Append "COSTCNTRCODE", adVarChar, 12, adFldIsNullable
                            .Fields.Append "CCDESC", adVarChar, 30, adFldIsNullable
                            .Fields.Append "JOBSTCODE", adVarChar, 10, adFldIsNullable
                            .Fields.Append "FEEDER_REF", adChar, 10, adFldIsNullable
                            .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, adFldIsNullable
                            
                            .Fields.Append "TRIS_PROJECT_AND_TASK_IDS", adChar, 19, adFldIsNullable
                            .Fields.Append "PROJECT_NUMBER", adVarChar, 10, adFldIsNullable
                            .Fields.Append "PROJECT_NAME", adVarChar, 60, adFldIsNullable
                            .Fields.Append "TASK_NUMBER", adVarChar, 21, adFldIsNullable
                            .Fields.Append "TASK_NAME", adVarChar, 60, adFldIsNullable
                            
                            .CursorLocation = adUseClient
                            .LockType = adLockBatchOptimistic
                            .CursorType = adOpenStatic
                            .Open
                            .ActiveConnection = Nothing
                        End With
                    End If
                    l_objrsMAXIMO.MoveFirst
                    While Not l_objrsMAXIMO.EOF
                        With l_rsGetDistJobTaskList
                            .AddNew
                                !JOB_NBR = l_objrsMAXIMO.Fields("JOB_NBR").Value
                                !TASK_NBR = l_objrsMAXIMO.Fields("TASK_NBR").Value
                                !FK_EMPRVDBCODE = l_objrsMAXIMO.Fields("FK_EMPRVDBCODE").Value
                                '//SetNullToEmpty
                                If Not IsNull(l_objrsMAXIMO.Fields("DESC60").Value) Then
                                    !DESC60 = l_objrsMAXIMO.Fields("DESC60").Value
                                Else
                                    !DESC60 = String(60, " ")
                                End If
                                If Not IsNull(l_objrsMAXIMO.Fields("JOBTASK_AND_DESC").Value) Then
                                    !JOBTASK_AND_DESC = l_objrsMAXIMO.Fields("JOBTASK_AND_DESC").Value
                                Else
                                    !JOBTASK_AND_DESC = String(60, " ")
                                End If
                                !CORP = l_objrsMAXIMO.Fields("CORP").Value
                                !UTLTY_CD = l_objrsMAXIMO.Fields("UTLTY_CD").Value
                                !BUS_DV_CD = l_objrsMAXIMO.Fields("BUS_DV_CD").Value
                                !MAJ_CD = l_objrsMAXIMO.Fields("MAJ_CD").Value
                                !MIN_CD = l_objrsMAXIMO.Fields("MIN_CD").Value
                                !FMC_CD = l_objrsMAXIMO.Fields("FMC_CD").Value
                                !RMC_CD = l_objrsMAXIMO.Fields("RMC_CD").Value
                                !TRN_TP_CD = l_objrsMAXIMO.Fields("TRN_TP_CD").Value
        
                                If Not IsNull(l_objrsMAXIMO.Fields("PROJECT_ID").Value) Then
                                    !PROJECT_ID = l_objrsMAXIMO.Fields("PROJECT_ID").Value
                                Else
                                    !PROJECT_ID = "     "
                                End If
                                
                                !PROD_CD = l_objrsMAXIMO.Fields("PROD_CD").Value
                                !ACTIV_CD = l_objrsMAXIMO.Fields("ACTIV_CD").Value
                                !RES_TP_CD = l_objrsMAXIMO.Fields("RES_TP_CD").Value
                                !FEEDER_REF = l_objrsMAXIMO.Fields("FEEDER_REF").Value
                                !WG_NUM = l_lngWGNum
                                !COSTCNTRCODE = v_strDeptNum
                                !CCDESC = v_strDeptNum
                                !JOBSTCODE = l_objrsMAXIMO.Fields("JOBSTCODE").Value
                                !DISTRIBUTION_TYPE = "M"
                                
                                !TRIS_PROJECT_AND_TASK_IDS = l_objrsMAXIMO.Fields("TRIS_PROJECT_AND_TASK_IDS").Value
                                !PROJECT_NUMBER = l_objrsMAXIMO.Fields("PROJECT_NUMBER").Value
                                !PROJECT_NAME = l_objrsMAXIMO.Fields("PROJECT_NAME").Value
                                !TASK_NUMBER = l_objrsMAXIMO.Fields("TASK_NUMBER").Value
                                !TASK_NAME = l_objrsMAXIMO.Fields("TASK_NAME").Value
                                
                                .Update
                        End With
                        l_objrsMAXIMO.MoveNext
                    Wend
                
                End If
            End If '//If Not l_objrsMAXIMO Is Nothing Then
        End If '//If (l_strInterfaceInd = "M" Or l_strInterfaceInd = "B") Then
    End If '//If Not v_blnEMPRVOnly Then
    
    '//to avoid error down the road, bad data with Maximo job/task, but dist type = "J"
    '//changed to create a default row reguardless of interface type or EMPRV only flag, if nothing has been added
    '//v_blnEMPRVOnly="TRUE"
    '//ErrDescription: Operation is not allowed when the object is closed.
    '//SourceRoutine: busTimecard.cbsTimeCard.GetAccountDescriptionJobTask
    '//GetActualTimeDetailAndSummary->RetrieveActualTimeEntriesForSummaryAndDetail->GetAccountDescriptionJobTask
    If l_rsGetDistJobTaskList.Fields.Count < 1 Then
        With l_rsGetDistJobTaskList
            .Fields.Append "JOB_NBR", adVarChar, 8, adFldIsNullable
            .Fields.Append "TASK_NBR", adVarChar, 4, adFldIsNullable
            .Fields.Append "FK_EMPRVDBCODE", adChar, 8, adFldIsNullable
            .Fields.Append "DESC60", adVarChar, 60, adFldIsNullable
            .Fields.Append "JOBTASK_AND_DESC", adVarChar, 60, adFldIsNullable
            .Fields.Append "CORP", adChar, 3, adFldIsNullable
            .Fields.Append "UTLTY_CD", adChar, 1, adFldIsNullable
            .Fields.Append "BUS_DV_CD", adChar, 2, adFldIsNullable
            .Fields.Append "MAJ_CD", adChar, 3, adFldIsNullable
            .Fields.Append "MIN_CD", adChar, 3, adFldIsNullable
            .Fields.Append "FMC_CD", adChar, 3, adFldIsNullable
            .Fields.Append "RMC_CD", adChar, 3, adFldIsNullable
            .Fields.Append "TRN_TP_CD", adChar, 1, adFldIsNullable
            .Fields.Append "PROJECT_ID", adChar, 5, adFldIsNullable
            .Fields.Append "PROD_CD", adChar, 2, adFldIsNullable
            .Fields.Append "ACTIV_CD", adChar, 4, adFldIsNullable
            .Fields.Append "RES_TP_CD", adChar, 2, adFldIsNullable
            .Fields.Append "WG_NUM", adInteger, , adFldIsNullable
            .Fields.Append "COSTCNTRCODE", adVarChar, 12, adFldIsNullable
            .Fields.Append "CCDESC", adVarChar, 30, adFldIsNullable
            .Fields.Append "JOBSTCODE", adVarChar, 10, adFldIsNullable
            .Fields.Append "FEEDER_REF", adChar, 10, adFldIsNullable
            .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, adFldIsNullable
            
            .Fields.Append "TRIS_PROJECT_AND_TASK_IDS", adChar, 19, adFldIsNullable
            .Fields.Append "PROJECT_NUMBER", adVarChar, 10, adFldIsNullable
            .Fields.Append "PROJECT_NAME", adVarChar, 60, adFldIsNullable
            .Fields.Append "TASK_NUMBER", adVarChar, 21, adFldIsNullable
            .Fields.Append "TASK_NAME", adVarChar, 60, adFldIsNullable
            
            .CursorLocation = adUseClient
            .LockType = adLockBatchOptimistic
            .CursorType = adOpenStatic
            .Open
            .ActiveConnection = Nothing
            
            .AddNew
            !DESC60 = "NO PICK LIST FOR DEPT " & v_strDeptNum
            !JOBTASK_AND_DESC = ""
            !JOB_NBR = "NOTFOUND"
            !TASK_NBR = Space(4)
            !FK_EMPRVDBCODE = Space(8)
            !CORP = Space(3)
            !UTLTY_CD = Space(1)
            !BUS_DV_CD = Space(2)
            !MAJ_CD = Space(3)
            !MIN_CD = Space(3)
            !FMC_CD = Space(3)
            !RMC_CD = Space(3)
            !TRN_TP_CD = Space(1)
            !PROJECT_ID = Space(5)
            !PROD_CD = Space(2)
            !ACTIV_CD = Space(4)
            !RES_TP_CD = Space(2)
            !FEEDER_REF = Space(10)
            !WG_NUM = l_lngWGNum
            !COSTCNTRCODE = v_strDeptNum
            !CCDESC = v_strDeptNum
            !JOBSTCODE = Space(10)
            !DISTRIBUTION_TYPE = "M"
            
            !TRIS_PROJECT_AND_TASK_IDS = ""
            !PROJECT_NUMBER = ""
            !PROJECT_NAME = ""
            !TASK_NUMBER = ""
            !TASK_NAME = ""
            
            .Update
        End With
    End If

    Set GetDistJobTaskList2 = l_rsGetDistJobTaskList
    
    If Not l_objCN Is Nothing Then
        If l_objCN.State = ADODB.adStateOpen Then
            l_objCN.Close
        End If
    End If
    
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    Exit Function

GetDistJobTaskList_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    Set l_objCMD = Nothing
    Set l_rsGetDistJobTaskList = Nothing

    Select Case Err.Number
       Case trErrdbGetDistributionListEmptyColWgNum
            TrsRaiseError m_strCMODULENAME, "GetDistJobTaskList2", trErrdbGetDistributionListEmptyColWgNum, , , , True
       Case Else
            TrsRaiseError m_strCMODULENAME, "GetDistJobTaskList2", , , , DescArgsFrom(l_strSQL), True
    End Select
End Function

'Public Function GetDistandJobTaskListforSelect( _
                    ByVal v_colWgNum As Collection) As ADODB.Recordset
         
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objrs As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
'    On Error GoTo GetDistandJobTaskListforSelect_EH
    
'    Dim l_strWgNumforSQL As String
     
'    If v_colWgNum.Count < 1 Then
'        l_strWgNumforSQL = "(0)"
'        Err.Raise trErrdbGetDistributionListEmptyColWgNum
'    Else
'        l_strWgNumforSQL = AsDelimitedString(v_colWgNum, False)
'    End If
    
   
'    Set l_objCN = New ADODB.Connection
'    Set l_objrs = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
   
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
'    Set l_objCM = Nothing
    
'    l_objCN.Open

    'get job_nbr, jobtask_nbr, and cost_center into description
    'add field for jobtask_desc
    
'    l_strSQL = "SELECT WG_NUM, "
'    l_strSQL = l_strSQL & " accounting AS DISTRIBUTION_ID, "
'    l_strSQL = l_strSQL & " JOB_NBR || '-' || JOBTASK_NBR || ' ' || COST_CENTER AS DESCRIPTION, "
'    l_strSQL = l_strSQL & " LAST_MOD_EMP_NUM, '' AS COST_CENTER, "
'    l_strSQL = l_strSQL & " CHAR(TIME_STAMP) AS TIME_STAMP, "
'    l_strSQL = l_strSQL & " '' AS EMPRV_DATABASE_ID, 'J' AS DISTRIBUTION_TYPE "
'    l_strSQL = l_strSQL & " FROM METHDBA.TTR_ACCT_ASSIGNMTS  "
'    l_strSQL = l_strSQL & "WHERE WG_NUM IN " & l_strWgNumforSQL
'    l_strSQL = l_strSQL & "SELECT WG_NUM, DISTRIBUTION_ID, DESCRIPTION, "
'    l_strSQL = l_strSQL & " '' AS JOB_TASK, "
'    l_strSQL = l_strSQL & " LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP, "
'    l_strSQL = l_strSQL & " '' AS EMPRV_DATABASE_ID, '' AS DISTRIBUTION_TYPE "
'    l_strSQL = l_strSQL & " FROM METHDBA.TTR_ACCT_ASSIGNMTS  "
'    l_strSQL = l_strSQL & "WHERE WG_NUM IN " & l_strWgNumforSQL
'        l_strSQL = l_strSQL & " ORDER BY DESCRIPTION "
    
'    l_objrs.CursorLocation = adUseClient
'    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

'    l_objrs.ActiveConnection = Nothing
'    l_objCN.Close
    
'    Set l_objCN = Nothing

'    Set GetDistandJobTaskListforSelect = l_objrs
    
'    Set l_objrs = Nothing
    
'    Exit Function

'GetDistandJobTaskListforSelect_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objrs = Nothing
    
'    Select Case Err.Number
'       Case trErrdbGetDistributionListEmptyColWgNum
'            TrsRaiseError m_strCMODULENAME, "GetDistandJobTaskListforSelect", trErrdbGetDistributionListEmptyColWgNum, , , , True
'       Case Else
'            TrsRaiseError m_strCMODULENAME, "GetDistandJobTaskListforSelect", , , , DescArgsFrom(l_strSQL), True
'    End Select
   
   
'End Function
'Public Function GetDistributionListforSelect( _
                    ByVal v_colWgNum As Collection) As ADODB.Recordset
         
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objrs As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
'    On Error GoTo GetDistributionListforSelect_EH
    
'    Dim l_strWgNumforSQL As String
     
'    If v_colWgNum.Count < 1 Then
'        l_strWgNumforSQL = "(0)"
'        Err.Raise trErrdbGetDistributionListEmptyColWgNum
'    Else
'        l_strWgNumforSQL = AsDelimitedString(v_colWgNum, False)
'    End If
    
   
'    Set l_objCN = New ADODB.Connection
'    Set l_objrs = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
   
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("EMPRV_DATABASE_A")
    
'    Set l_objCM = Nothing
    
'    l_objCN.Open

    'get job_nbr, jobtask_nbr, and cost_center into description
    'add field for jobtask_desc
    
'    l_strSQL = l_strSQL & "SELECT WG_NUM, DISTRIBUTION_ID, DESCRIPTION, "
'    l_strSQL = l_strSQL & " '' AS JOB_TASK, "
'    l_strSQL = l_strSQL & " LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP, "
'    l_strSQL = l_strSQL & " '' AS EMPRV_DATABASE_ID, '' AS DISTRIBUTION_TYPE "
'    l_strSQL = l_strSQL & " FROM METHDBA.TTR_ACCT_ASSIGNMTS  "
'    l_strSQL = l_strSQL & "WHERE WG_NUM IN " & l_strWgNumforSQL
'        l_strSQL = l_strSQL & " ORDER BY DESCRIPTION "
    
'    l_objrs.CursorLocation = adUseClient
'    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

'    l_objrs.ActiveConnection = Nothing
'    l_objCN.Close
    
'    Set l_objCN = Nothing

'    Set GetDistributionListforSelect = l_objrs
    
'    Set l_objrs = Nothing
    
'    Exit Function

'GetDistributionListforSelect_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objrs = Nothing
    
'    Select Case Err.Number
'       Case trErrdbGetDistributionListEmptyColWgNum
'            TrsRaiseError m_strCMODULENAME, "GetDistributionListforSelect", trErrdbGetDistributionListEmptyColWgNum, , , , True
'       Case Else
'            TrsRaiseError m_strCMODULENAME, "GetDistributionListforSelect", , , , DescArgsFrom(l_strSQL), True
'    End Select
   
   
'End Function

Public Function GetValidECWorkNumbers() As ADODB.Recordset
     TrsRaiseError m_strCMODULENAME, "GetValidECWorkNumbers", trErrObsoleteMethod, , , , True
' RED131
'On Error GoTo GetValidECWorkNumbers_EH
'
'
'    Set GetValidECWorkNumbers = GetWorkNumbers("ALL")
'
'Exit Function
'GetValidECWorkNumbers_EH:
'
'    TrsRaiseError m_strCMODULENAME, "GetValidECWorkNumbers", , , , , True

End Function

Public Function GetValidECWorkNumProdTime() As ADODB.Recordset
TrsRaiseError m_strCMODULENAME, "GetValidECWorkNumProdTime", trErrObsoleteMethod, , , , True
'' RED132
'
'    Dim l_rsValidECWorkNumProdTime As ADODB.Recordset
'
'On Error GoTo GetValidECWorkNumProdTime_EH
'
'
'    Set GetValidECWorkNumProdTime = GetWorkNumbers("PROD")
'
'Exit Function
'GetValidECWorkNumProdTime_EH:
'
'    TrsRaiseError m_strCMODULENAME, "GetValidECWorkNumProdTime", , , , , True
'
End Function

Public Function GetValidECWorkNumNonProdTime() As ADODB.Recordset
TrsRaiseError m_strCMODULENAME, "GetValidECWorkNumNonProdTime", trErrObsoleteMethod, , , , True
'' RED133
'
'    Dim l_rsValidECWorkNumNonProdTime As ADODB.Recordset
'
'On Error GoTo GetValidECWorkNumProdTime_EH
'
'
'    Set GetValidECWorkNumNonProdTime = GetWorkNumbers("NONPROD")
'
'Exit Function
'GetValidECWorkNumProdTime_EH:
'
'    TrsRaiseError m_strCMODULENAME, "GetValidECWorkNumProdTime", , , , , True

End Function


Public Function GetValidECDrawingSizes() As ADODB.Recordset
    TrsRaiseError m_strCMODULENAME, "GetValidECDrawingSizes", trErrObsoleteMethod, , , , True
'' RED139
''  This is a list that is currently Hard Coded into this function
''  At some point this list should be maintained in a database
'
'
'    Dim l_rsValidECDrawingSizes As ADODB.Recordset
'    Dim l_strDrawingSizes As String
'    Dim l_strDrawingSize As String
'    Dim l_lngCount As Long
'    Dim l_intLength As Integer
'
'On Error GoTo GetValidECDrawingSizes_EH
''RED139
'
'    l_strDrawingSizes = "0107101516232425266877788899"
'
'    Set l_rsValidECDrawingSizes = New ADODB.Recordset
'
'    l_rsValidECDrawingSizes.Fields.Append "DRWG_SIZE", adChar, 2, adFldUpdatable
'    l_rsValidECDrawingSizes.Open
'
'    l_intLength = Len(l_strDrawingSizes)
'    l_lngCount = 1
'
'    Do Until l_lngCount > l_intLength
'        l_strDrawingSize = Mid(l_strDrawingSizes, l_lngCount, 2)
'
'        With l_rsValidECDrawingSizes
'            .AddNew
'            .Fields("DRWG_SIZE").Value = l_strDrawingSize
'        End With
'
'        l_lngCount = l_lngCount + 2
'    Loop
'
'    l_rsValidECDrawingSizes.Update
'
'    Set GetValidECDrawingSizes = l_rsValidECDrawingSizes
'
'Exit Function
'GetValidECDrawingSizes_EH:
'
'    TrsRaiseError m_strCMODULENAME, "GetValidECDrawingSizes", , , , , True

End Function

Public Function GetValidECClassNumbers() As ADODB.Recordset
' RED135

     TrsRaiseError m_strCMODULENAME, "GetValidECClassNumbers", trErrObsoleteMethod, , , , True
    
'On Error GoTo GetValidECClassNumbers_EH:
'
'
'    Set GetValidECClassNumbers = GetClassNumbers("ALL")
'
'Exit Function
'GetValidECClassNumbers_EH:
'
'    TrsRaiseError m_strCMODULENAME, "GetValidECClassNumbers", , , , , True

End Function

Public Function GetValidECClassNumbersProdTime() As ADODB.Recordset

    TrsRaiseError m_strCMODULENAME, "GetValidECClassNumbersProdTime", trErrObsoleteMethod, , , , True
'' RED135
'
'
'On Error GoTo GetValidECClassNumbersProdTime_EH:
'
'
'    Set GetValidECClassNumbersProdTime = GetClassNumbers("PROD")
'
'Exit Function
'GetValidECClassNumbersProdTime_EH:
'
'    TrsRaiseError m_strCMODULENAME, "GetValidECWorkNumProdTime", , , , , True

End Function

Public Function GetValidECClassNumbersNonProdTime() As ADODB.Recordset

    TrsRaiseError m_strCMODULENAME, "GetValidECClassNumbersNonProdTime", trErrObsoleteMethod, , , , True
' RED137

'On Error GoTo GetValidECClassNumbersNonProdTime_EH
'
'
'    Set GetValidECClassNumbersNonProdTime = GetClassNumbers("NONPROD")
'
'Exit Function
'GetValidECClassNumbersNonProdTime_EH:
'
'    TrsRaiseError m_strCMODULENAME, "GetValidECWorkNumNonProdTime", , , , , True

End Function

Public Function GetECWorkReason() As ADODB.Recordset
    
    TrsRaiseError m_strCMODULENAME, "GetECWorkReason", trErrObsoleteMethod, , , , True
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objrs As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'
'    On Error GoTo GetECWorkReason_EH
'
'
'    Set l_objCN = New ADODB.Connection
'    Set l_objrs = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
'
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open
'
'    l_strSQL = "SELECT EMP_NUM, WORK_REASON "
'    l_strSQL = l_strSQL & "FROM METHDBA.TTR_EC_WORK_REAS "
'    l_strSQL = l_strSQL & "ORDER BY EMP_NUM"
'
'    l_objrs.CursorLocation = adUseClient
'    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objrs.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'
'    Set GetECWorkReason = l_objrs
'
'    Set l_objrs = Nothing
'
'Exit Function
'GetECWorkReason_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objrs = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetECWorkReason", , , , DescArgsFrom(l_strSQL), True
'
End Function

'Private Function GetClassNumbers(ByVal v_strType As String) As ADODB.Recordset
''  This is a list that is currently Hard Coded into this function
''  At some point this list should be maintained in a database
'
'    Dim l_rsValidECClassNumbers As ADODB.Recordset
'    Dim l_strClassNumbers As String
'    Dim l_strClassNumber As String
'    Dim l_lngCount As Long
'    Dim l_intLength As Integer
'
'On Error GoTo GetClassNumbers_EH
''RED139
'    l_strClassNumbers = "0101002010030100302003030030400305003060030700308003090"
'    l_strClassNumbers = l_strClassNumbers & "04010040200403004040040500406004070"
'    l_strClassNumbers = l_strClassNumbers & "05010050200503005031050320503305034050400505005060"
'    l_strClassNumbers = l_strClassNumbers & "06010060200603007010070200704008010080200803008040"
'    l_strClassNumbers = l_strClassNumbers & "0901009020090300903109032090340904009041090420904409045"
'    l_strClassNumbers = l_strClassNumbers & "100101002010030110101102011030110401105011060"
'    l_strClassNumbers = l_strClassNumbers & "120101201112012120131202012021120301203112032120331203412035"
'    l_strClassNumbers = l_strClassNumbers & "120361203712038120399991099977999809999907030"
'
'    ' Create the Recordset
'    Set l_rsValidECClassNumbers = New ADODB.Recordset
'
'    l_rsValidECClassNumbers.Fields.Append "CLASS_NUM", adChar, 5, adFldUpdatable
'    l_rsValidECClassNumbers.Open
'
'    l_intLength = Len(l_strClassNumbers)
'    l_lngCount = 1
'
'    Do Until l_lngCount > l_intLength
'        l_strClassNumber = Mid(l_strClassNumbers, l_lngCount, 5)
'
'        Select Case UCase(v_strType)
'            Case "ALL"
'                With l_rsValidECClassNumbers
'                    .AddNew
'                    .Fields("CLASS_NUM").Value = l_strClassNumber
'                End With
'
'            Case "NONPROD"
'                Select Case l_strClassNumber
'                    Case "01010", "99910", "99977", "99980", "99999"
'                        With l_rsValidECClassNumbers
'                            .AddNew
'                            .Fields("CLASS_NUM").Value = l_strClassNumber
'                        End With
'                    Case Else
'                        ' Do nothing
'                End Select
'
'            Case "PROD"
'                Select Case l_strClassNumber
'                    Case "99910", "99977", "99980", "99999"
'                        ' Do nothing
'                    Case Else
'                        With l_rsValidECClassNumbers
'                            .AddNew
'                            .Fields("CLASS_NUM").Value = l_strClassNumber
'                        End With
'                End Select
'
'            Case Else
'                ' Bad Request send nothing
'
'        End Select
'        l_lngCount = l_lngCount + 5
'    Loop
'
'    l_rsValidECClassNumbers.Update
'
'    Set GetClassNumbers = l_rsValidECClassNumbers
'    Set l_rsValidECClassNumbers = Nothing
''Why did Juan set this Recordset to Nothing upon successful completion
'
'Exit Function
'GetClassNumbers_EH:
'    Set l_rsValidECClassNumbers = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetClassNumbers", , , , , True
'
'End Function

'Private Function GetWorkNumbers(ByVal v_strType) As ADODB.Recordset
'
''  This is a list that is currently Hard Coded into this function
''  At some point this list should be maintained in a database
'
'    Dim l_rsValidECWorkNumbers As ADODB.Recordset
'    Dim l_strWorkNumbers As String
'    Dim l_strWorkNumber As String
'    Dim l_lngCount As Long
'    Dim l_intLength As Integer
'
'On Error GoTo GetWorkNumbers_EH
'
'    l_strWorkNumbers = "010203040506111213141516212223242526313233343536"
'    l_strWorkNumbers = l_strWorkNumbers & "414243444546515253545556616263646566"
'    l_strWorkNumbers = l_strWorkNumbers & "717273747576818283848586959697"
'
'    Set l_rsValidECWorkNumbers = New ADODB.Recordset
'
'    l_rsValidECWorkNumbers.Fields.Append "WORK_NUM", adChar, 2, adFldUpdatable
'    l_rsValidECWorkNumbers.Open
'
'    l_intLength = Len(l_strWorkNumbers)
'    l_lngCount = 1
'
'    Do Until l_lngCount > l_intLength
'        l_strWorkNumber = Mid(l_strWorkNumbers, l_lngCount, 2)
'
'        Select Case UCase(v_strType)
'            Case "ALL"
'                With l_rsValidECWorkNumbers
'                    .AddNew
'                    .Fields("WORK_NUM").Value = l_strWorkNumber
'                End With
'
'            Case "NONPROD"
'                Select Case l_strWorkNumber
'                    Case "63", "73", "83", "95", "96", "97"
'                        With l_rsValidECWorkNumbers
'                            .AddNew
'                            .Fields("WORK_NUM").Value = l_strWorkNumber
'                        End With
'
'                    Case Else
'                        ' Do noting
'                End Select
'
'            Case "PROD"
'                Select Case l_strWorkNumber
'                    Case "95", "96", "97"
'                        ' Do nothing
'                    Case Else
'                        With l_rsValidECWorkNumbers
'                            .AddNew
'                            .Fields("WORK_NUM").Value = l_strWorkNumber
'                        End With
'                End Select
'
'            Case Else
'                ' Bad request do nothing
'        End Select
'
'        l_lngCount = l_lngCount + 2
'    Loop
'
'    l_rsValidECWorkNumbers.Update
'
'    Set GetWorkNumbers = l_rsValidECWorkNumbers
'    Set l_rsValidECWorkNumbers = Nothing
''Why did Juan set this Recordset to Nothing upon successful completion
'
'Exit Function
'GetWorkNumbers_EH:
'    Set l_rsValidECWorkNumbers = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetWorkNumbers", , , , , True
'
'End Function


Public Function ValidateGLAccounting( _
            ByVal v_strDistributionIDEdits As Variant, _
            ByVal v_intNumberOfEdits As Integer, _
            ByRef r_strDistributionIDsReturned As Variant, _
            ByRef r_strReturnCodes As Variant, _
            ByRef r_strReturnMsgs As Variant, _
            ByRef r_strHighlightFields As Variant _
            ) As Boolean
                
On Error GoTo eh
    
'    ValidateGLAccounting = ValidateGLAccountingTR403(v_strDistributionIDEdits, v_intNumberOfEdits, _
'                                                    r_strDistributionIDsReturned, r_strReturnCodes, r_strReturnMsgs, r_strHighlightFields)
    
'    ValidateGLAccounting = ValidateGLAccountingAAVM(v_strDistributionIDEdits, v_intNumberOfEdits, _
'                                                    r_strDistributionIDsReturned, r_strReturnCodes, r_strReturnMsgs, r_strHighlightFields)
    
    'ValidateGLAccounting = ValidateGLAccountingWS(v_strDistributionIDEdits, v_intNumberOfEdits, _
    '                                                r_strDistributionIDsReturned, r_strReturnCodes, r_strReturnMsgs, r_strHighlightFields)
                                                
    ValidateGLAccounting = ValidateGLAccountingAPI(v_strDistributionIDEdits, v_intNumberOfEdits, _
                                                    r_strDistributionIDsReturned, r_strReturnCodes, r_strReturnMsgs, r_strHighlightFields)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateGLAccounting", , , , , True
End Function

Private Function ValidateGLAccountingTR403( _
            ByVal v_strDistributionIDEdits As Variant, _
            ByVal v_intNumberOfEdits As Integer, _
            ByRef r_strDistributionIDsReturned As Variant, _
            ByRef r_strReturnCodes As Variant, _
            ByRef r_strReturnMsgs As Variant, _
            ByRef r_strHighlightFields As Variant _
            ) As Boolean
            
    Dim l_objCN As ADODB.Connection
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim Acmd As ADODB.Command
    Dim PSp_Sqlcode As ADODB.Parameter
    Dim PSp_Succ_Fail_Flag As ADODB.Parameter
    Dim PSp_Abhi_Abend_Handler_Area As ADODB.Parameter
    Dim PSp_Abhi_Exec_Level_Msg As ADODB.Parameter
    Dim PSp_Applid As ADODB.Parameter
    Dim PSp_System_Id As ADODB.Parameter
    Dim PSp_Set_1_Rows As ADODB.Parameter
    Dim PSp_Set_1 As ADODB.Parameter
    Dim l_strDistributionIDsReturned(221) As String
    Dim l_strReturnCodes(221) As String
    Dim l_strReturnMsgs(221) As String
    Dim l_strHighlightFields(221) As String
    Dim l_strSPReturnValue As String
    Dim l_intI As Integer
    Dim l_intPos As Integer
    Dim l_strServer As String
    
On Error GoTo eh
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT2")
    l_objCN.CursorLocation = adUseClient
    Set l_objCM = Nothing
    l_objCN.Open
    
'RSG Use registry value instead of data provider property since
' DB2Connect and ShadowDirect use it differently.
'    l_strServer = UCase(l_objCN.Properties.Item("Server Name"))
    l_strServer = GetDB2Subsystem("SHADOWDIRECT2")
    
    Set Acmd = New ADODB.Command
    With Acmd
        .ActiveConnection = l_objCN
        .CommandType = adCmdStoredProc

        If l_strServer = "DB2P" Then
'RSG Removed SYSPROC from the high-level qualifier for DB2Connect
' access to stored procedure.  Neon required the SYSPROC qualifier.
'            .CommandText = "SYSPROC.TR1SPROD.TR403"
            .CommandText = "TR1SPROD.TR403"
        Else
            .CommandText = "SYSPROC.TR403"
        End If
        
        Set PSp_Sqlcode = Acmd.CreateParameter("aSp_Sqlcode", adInteger, adParamInputOutput)
        PSp_Sqlcode = -471
        .Parameters.Append PSp_Sqlcode
        
        Set PSp_Succ_Fail_Flag = Acmd.CreateParameter("aSp_Succ_Fail_Flag", adChar, adParamOutput, 1)
        .Parameters.Append PSp_Succ_Fail_Flag
       
        Set PSp_Abhi_Abend_Handler_Area = Acmd.CreateParameter("aSp_ABHI_Abend_Handler_Area", adVarChar, adParamOutput, 2000)
        .Parameters.Append PSp_Abhi_Abend_Handler_Area
     
         Set PSp_Abhi_Exec_Level_Msg = Acmd.CreateParameter("aSp_Abhi_Exec_level_Msg", adChar, adParamOutput, 128)
        .Parameters.Append PSp_Abhi_Exec_Level_Msg
        
        Set PSp_Applid = Acmd.CreateParameter("aSp_Applid", adChar, adParamInput, 8)
        
        PSp_Applid = GetCICSRegion(l_strServer) '//-----convert the db2 subsystem (db2a) to CICS Region (sy7321)-----//
                
        .Parameters.Append PSp_Applid
        
        Set PSp_System_Id = Acmd.CreateParameter("aSp_System_Id", adChar, adParamInput, 2)
        PSp_System_Id = "TR"
        .Parameters.Append PSp_System_Id
    
        Set PSp_Set_1_Rows = Acmd.CreateParameter("aSp_Set_1_Rows", adSmallInt, adParamInput)
        PSp_Set_1_Rows = v_intNumberOfEdits
        .Parameters.Append PSp_Set_1_Rows
    
        Set PSp_Set_1 = Acmd.CreateParameter("aSp_Set_1", adVarChar, adParamInputOutput, 32636)
        PSp_Set_1.Value = ConvertEditArrayToString(v_strDistributionIDEdits)
        .Parameters.Append PSp_Set_1
     
        .Execute
    End With
 
    l_objCN.Close
    Set l_objCN = Nothing
    
    l_strSPReturnValue = PSp_Set_1.Value
    For l_intI = 0 To 221
        l_intPos = l_intI * 147
        l_strReturnCodes(l_intI) = Mid$(l_strSPReturnValue, l_intPos + 45, 1)
        l_strReturnMsgs(l_intI) = Mid$(l_strSPReturnValue, l_intPos + 46, 50)
        l_strHighlightFields(l_intI) = Mid$(l_strSPReturnValue, l_intPos + 96, 13)
        l_strDistributionIDsReturned(l_intI) = Mid$(l_strSPReturnValue, l_intPos + 109, 39)
    Next l_intI
    
    If PSp_Sqlcode.Value < 0 Then
        ValidateGLAccountingTR403 = False
    Else
        ValidateGLAccountingTR403 = True
    End If
    
    r_strDistributionIDsReturned = l_strDistributionIDsReturned
    r_strReturnCodes = l_strReturnCodes
    r_strReturnMsgs = l_strReturnMsgs
    r_strHighlightFields = l_strHighlightFields
    
    Exit Function
    
eh:
    Dim l_colErrorDesc As New Collection
    l_colErrorDesc.Add "Start of the CICS ABHI errors"
    l_colErrorDesc.Add PSp_Abhi_Exec_Level_Msg.Value
    l_colErrorDesc.Add PSp_Abhi_Abend_Handler_Area.Value
    l_colErrorDesc.Add "End of the CICS ABHI errors"
    
    Set l_objCN = Nothing
    Set Acmd = Nothing
    TrsRaiseError m_strCMODULENAME, "ValidateGLAccountingTR403", , , , l_colErrorDesc, True
End Function

Private Function ValidateGLAccountingAAVM( _
            ByVal v_strDistributionIDEdits As Variant, _
            ByVal v_intNumberOfEdits As Integer, _
            ByRef r_strDistributionIDsReturned As Variant, _
            ByRef r_strReturnCodes As Variant, _
            ByRef r_strReturnMsgs As Variant, _
            ByRef r_strHighlightFields As Variant _
            ) As Boolean
            
    Dim l_objCN As ADODB.Connection
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim Acmd As ADODB.Command
    Dim p_sys_id As ADODB.Parameter
    Dim p_edit_reason As ADODB.Parameter
    Dim p_batch_action As ADODB.Parameter
    Dim p_corp As ADODB.Parameter
    Dim p_util As ADODB.Parameter
    Dim p_bus_div As ADODB.Parameter
    Dim p_maj_min As ADODB.Parameter
    Dim p_fmc As ADODB.Parameter
    Dim p_rmc As ADODB.Parameter
    Dim p_tran As ADODB.Parameter
    Dim p_project As ADODB.Parameter
    Dim p_prod As ADODB.Parameter
    Dim p_actv As ADODB.Parameter
    Dim p_res_typ As ADODB.Parameter
    Dim p_feeder_ref As ADODB.Parameter
    Dim p_return_code As ADODB.Parameter
    Dim p_return_msg As ADODB.Parameter
    Dim p_return_corp_flag As ADODB.Parameter
    Dim p_return_util_flag As ADODB.Parameter
    Dim p_return_bus_div_flag As ADODB.Parameter
    Dim p_return_maj_min_flag As ADODB.Parameter
    Dim p_return_fmc_flag As ADODB.Parameter
    Dim p_return_rmc_flag As ADODB.Parameter
    Dim p_return_tran_flag As ADODB.Parameter
    Dim p_return_project_flag As ADODB.Parameter
    Dim p_return_prod_flag As ADODB.Parameter
    Dim p_return_actv_flag As ADODB.Parameter
    Dim p_return_res_typ_flag As ADODB.Parameter
    Dim p_return_feeder_ref_flag As ADODB.Parameter
    Dim p_in_dt As ADODB.Parameter
    Dim strDistributionId As String '//parse this out of the array, so that further substring'ing will match
    Dim strDistributionIdReturned As String
    Dim l_strDistributionIDsReturned(221) As String
    Dim l_strReturnCodes(221) As String
    Dim l_strReturnMsgs(221) As String
    Dim l_strHighlightFields(221) As String
'    Dim l_strSPReturnValue As String
    Dim l_intI As Integer
'    Dim l_intPos As Integer
    'Dim l_strServer As String
    
On Error GoTo eh
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT3")
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
        
    For l_intI = 0 To v_intNumberOfEdits - 1
        strDistributionId = Mid$(v_strDistributionIDEdits(l_intI), 6)
        
        Set Acmd = New ADODB.Command
        With Acmd
            .ActiveConnection = l_objCN
            .CommandType = adCmdStoredProc
            .CommandText = Trim(l_objCM.GetSchema("SHADOWDIRECT3")) & ".AAVM_CONTROL"
            
            Set p_sys_id = Acmd.CreateParameter("p_sys_id", adVarChar, adParamInputOutput, 2)
            p_sys_id = "TR"
            .Parameters.Append p_sys_id
        
            Set p_edit_reason = Acmd.CreateParameter("p_edit_reason", adVarChar, adParamInputOutput, 2)
            p_edit_reason = Left(v_strDistributionIDEdits(l_intI), 2) ' "FE"
            .Parameters.Append p_edit_reason
            
            Set p_batch_action = Acmd.CreateParameter("p_batch_action", adVarChar, adParamInputOutput, 2)
            p_batch_action = ""
            .Parameters.Append p_batch_action
            
            Set p_corp = Acmd.CreateParameter("p_corp", adVarChar, adParamInputOutput, 3)
            p_corp = Mid$(v_strDistributionIDEdits(l_intI), 3, 3)
            .Parameters.Append p_corp
            
            Set p_util = Acmd.CreateParameter("p_util", adVarChar, adParamInputOutput, 1)
            p_util = GetUtility39(strDistributionId) 'Left(strDistributionId, 1)
            .Parameters.Append p_util
            
            Set p_bus_div = Acmd.CreateParameter("p_bus_div", adVarChar, adParamInputOutput, 2)
            p_bus_div = GetBusinessDivision39(strDistributionId) 'Mid$(strDistributionId, 2, 2)
            .Parameters.Append p_bus_div
            
            Set p_maj_min = Acmd.CreateParameter("p_maj_min", adVarChar, adParamInputOutput, 6)
            p_maj_min = GetMajor39(strDistributionId) & GetMinor39(strDistributionId)  'Mid$(strDistributionId, 4, 6)
            .Parameters.Append p_maj_min
            
            Set p_fmc = Acmd.CreateParameter("p_fmc", adVarChar, adParamInputOutput, 3)
            p_fmc = GetFMC39(strDistributionId) 'Mid$(strDistributionId, 10, 3)
            .Parameters.Append p_fmc
            
            Set p_rmc = Acmd.CreateParameter("p_rmc", adVarChar, adParamInputOutput, 3)
            p_rmc = GetRMC39(strDistributionId) ' Mid$(strDistributionId, 13, 3)
            .Parameters.Append p_rmc
    
            Set p_tran = Acmd.CreateParameter("p_tran", adVarChar, adParamInputOutput, 1)
            p_tran = GetTranType39(strDistributionId)
            .Parameters.Append p_tran
            
            Set p_project = Acmd.CreateParameter("p_project", adVarChar, adParamInputOutput, 5)
            p_project = GetProject39(strDistributionId)
            .Parameters.Append p_project
            
            Set p_prod = Acmd.CreateParameter("p_prod", adVarChar, adParamInputOutput, 2)
            p_prod = GetProduct39(strDistributionId)
            .Parameters.Append p_prod
            
            Set p_actv = Acmd.CreateParameter("p_actv", adVarChar, adParamInputOutput, 4)
            p_actv = GetActivity39(strDistributionId)
            .Parameters.Append p_actv
            
            Set p_res_typ = Acmd.CreateParameter("p_res_typ", adVarChar, adParamInputOutput, 2)
            p_res_typ = GetResourceType39(strDistributionId)
            .Parameters.Append p_res_typ
            
            Set p_feeder_ref = Acmd.CreateParameter("p_feeder_ref", adVarChar, adParamInputOutput, 10)
            p_feeder_ref = GetFeederRef39(strDistributionId)
            .Parameters.Append p_feeder_ref
            
            Set p_return_code = Acmd.CreateParameter("p_return_code", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_code
    
            Set p_return_msg = Acmd.CreateParameter("p_return_msg", adVarChar, adParamOutput, 50)
            .Parameters.Append p_return_msg
    
            Set p_return_corp_flag = Acmd.CreateParameter("p_return_corp_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_corp_flag
    
            Set p_return_util_flag = Acmd.CreateParameter("p_return_util_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_util_flag
    
            Set p_return_bus_div_flag = Acmd.CreateParameter("p_return_bus_div_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_bus_div_flag
    
            Set p_return_maj_min_flag = Acmd.CreateParameter("p_return_maj_min_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_maj_min_flag
    
            Set p_return_fmc_flag = Acmd.CreateParameter("p_return_fmc_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_fmc_flag
    
            Set p_return_rmc_flag = Acmd.CreateParameter("p_return_rmc_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_rmc_flag
    
            Set p_return_tran_flag = Acmd.CreateParameter("p_return_tran_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_tran_flag
    
            Set p_return_project_flag = Acmd.CreateParameter("p_return_project_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_project_flag
    
            Set p_return_prod_flag = Acmd.CreateParameter("p_return_prod_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_prod_flag
    
            Set p_return_actv_flag = Acmd.CreateParameter("p_return_actv_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_actv_flag
    
            Set p_return_res_typ_flag = Acmd.CreateParameter("p_return_res_typ_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_res_typ_flag
    
            Set p_return_feeder_ref_flag = Acmd.CreateParameter("p_return_feeder_ref_flag", adVarChar, adParamOutput, 1)
            .Parameters.Append p_return_feeder_ref_flag
    
            Set p_in_dt = Acmd.CreateParameter("p_in_dt", adDate, adParamInputOutput)
            p_in_dt.Value = Date 'if null stored proc will default to current date
            .Parameters.Append p_in_dt
                 
            .Execute
        End With
            
        'l_intI = 0
    '    l_strSPReturnValue = PSp_Set_1.Value
    '    For l_intI = 0 To 221
    '        l_intPos = l_intI * 147
        l_strReturnCodes(l_intI) = p_return_code 'Mid$(l_strSPReturnValue, l_intPos + 45, 1)
        l_strReturnMsgs(l_intI) = IIf(IsNull(p_return_msg), Space(50), p_return_msg) 'Mid$(l_strSPReturnValue, l_intPos + 46, 50)
        '//since major and minor are combined, they are listed twice in order to maintain 13 flags
        l_strHighlightFields(l_intI) = p_return_corp_flag & p_return_util_flag & p_return_bus_div_flag & _
                                        p_return_maj_min_flag & p_return_maj_min_flag & _
                                        p_return_fmc_flag & p_return_rmc_flag & _
                                        p_return_tran_flag & p_return_project_flag & p_return_prod_flag & _
                                        p_return_actv_flag & p_return_res_typ_flag & p_return_feeder_ref_flag
        'Mid$(l_strSPReturnValue, l_intPos + 96, 13)
        strDistributionIdReturned = Space(39)
        SetUtility39 strDistributionIdReturned, p_util
        SetBusinessDivision39 strDistributionIdReturned, p_bus_div
        SetMajor39 strDistributionIdReturned, Left(p_maj_min, 3)
        SetMinor39 strDistributionIdReturned, Mid(p_maj_min, 4, 3)
        SetFMC39 strDistributionIdReturned, p_fmc
        SetRMC39 strDistributionIdReturned, p_rmc
        SetTranType39 strDistributionIdReturned, p_tran
        SetProject39 strDistributionIdReturned, p_project
        SetProduct39 strDistributionIdReturned, p_prod
        SetActivity39 strDistributionIdReturned, p_actv
        SetResourceType39 strDistributionIdReturned, p_res_typ
        SetFeederRef39 strDistributionIdReturned, p_feeder_ref
        l_strDistributionIDsReturned(l_intI) = strDistributionIdReturned
            'Mid$(l_strSPReturnValue, l_intPos + 109, 39)
    '    Next l_intI
    Next
    
    l_objCN.Close
    Set l_objCN = Nothing

'    If p_return_code.Value = "0" Then
        ValidateGLAccountingAAVM = True '//TODO: meaning is tied to sql return code rather than validation result
'    Else
'        ValidateGLAccounting = False
'    End If
    
    r_strDistributionIDsReturned = l_strDistributionIDsReturned
    r_strReturnCodes = l_strReturnCodes
    r_strReturnMsgs = l_strReturnMsgs
    r_strHighlightFields = l_strHighlightFields
    
    Set l_objCM = Nothing
    Set Acmd = Nothing
    Set p_sys_id = Nothing
    Set p_edit_reason = Nothing
    Set p_batch_action = Nothing
    Set p_corp = Nothing
    Set p_util = Nothing
    Set p_bus_div = Nothing
    Set p_maj_min = Nothing
    Set p_fmc = Nothing
    Set p_rmc = Nothing
    Set p_tran = Nothing
    Set p_project = Nothing
    Set p_prod = Nothing
    Set p_actv = Nothing
    Set p_res_typ = Nothing
    Set p_feeder_ref = Nothing
    Set p_return_code = Nothing
    Set p_return_msg = Nothing
    Set p_return_corp_flag = Nothing
    Set p_return_util_flag = Nothing
    Set p_return_bus_div_flag = Nothing
    Set p_return_maj_min_flag = Nothing
    Set p_return_fmc_flag = Nothing
    Set p_return_rmc_flag = Nothing
    Set p_return_tran_flag = Nothing
    Set p_return_project_flag = Nothing
    Set p_return_prod_flag = Nothing
    Set p_return_actv_flag = Nothing
    Set p_return_res_typ_flag = Nothing
    Set p_return_feeder_ref_flag = Nothing
    Set p_in_dt = Nothing
    
    Exit Function
eh:
'    Dim l_colErrorDesc As New Collection
'    l_colErrorDesc.Add "Start of the CICS ABHI errors"
'    l_colErrorDesc.Add PSp_Abhi_Exec_Level_Msg.Value
'    l_colErrorDesc.Add PSp_Abhi_Abend_Handler_Area.Value
'    l_colErrorDesc.Add "End of the CICS ABHI errors"
    
    Set l_objCN = Nothing
    Set l_objCM = Nothing
    Set Acmd = Nothing
    Set p_sys_id = Nothing
    Set p_edit_reason = Nothing
    Set p_batch_action = Nothing
    Set p_corp = Nothing
    Set p_util = Nothing
    Set p_bus_div = Nothing
    Set p_maj_min = Nothing
    Set p_fmc = Nothing
    Set p_rmc = Nothing
    Set p_tran = Nothing
    Set p_project = Nothing
    Set p_prod = Nothing
    Set p_actv = Nothing
    Set p_res_typ = Nothing
    Set p_feeder_ref = Nothing
    Set p_return_code = Nothing
    Set p_return_msg = Nothing
    Set p_return_corp_flag = Nothing
    Set p_return_util_flag = Nothing
    Set p_return_bus_div_flag = Nothing
    Set p_return_maj_min_flag = Nothing
    Set p_return_fmc_flag = Nothing
    Set p_return_rmc_flag = Nothing
    Set p_return_tran_flag = Nothing
    Set p_return_project_flag = Nothing
    Set p_return_prod_flag = Nothing
    Set p_return_actv_flag = Nothing
    Set p_return_res_typ_flag = Nothing
    Set p_return_feeder_ref_flag = Nothing
    Set p_in_dt = Nothing
    
    TrsRaiseError m_strCMODULENAME, "ValidateGLAccountingAAVM", , , , , True
'    TrsRaiseError m_strCMODULENAME, "ValidateGLAccountingAAVM", , , , l_colErrorDesc, True
End Function

Private Function ValidateGLAccountingWS( _
            ByVal v_strDistributionIDEdits As Variant, _
            ByVal v_intNumberOfEdits As Integer, _
            ByRef r_strDistributionIDsReturned As Variant, _
            ByRef r_strReturnCodes As Variant, _
            ByRef r_strReturnMsgs As Variant, _
            ByRef r_strHighlightFields As Variant _
            ) As Boolean
            
    Dim l_obj As Ameren_TRIS_Server_COMIntegration.DistributionNumber
    
    Dim l_bolDBSuccess As Boolean
    
    On Error GoTo eh
    
    Set l_obj = New Ameren_TRIS_Server_COMIntegration.DistributionNumber
    
    '//if the return value of false previously meant that the call to the stored procedure failed
    '//now a value of false means that the validation failed, let any failures result in errors/exceptions
    ReDim Preserve v_strDistributionIDEdits(v_intNumberOfEdits - 1) '//only send what is needed to web service instead of a fixed 222 (0-221) array
    l_bolDBSuccess = l_obj.ValidateDistributionNumbers(v_strDistributionIDEdits, v_intNumberOfEdits, r_strDistributionIDsReturned, r_strReturnCodes, r_strReturnMsgs, r_strHighlightFields)
    
    If l_bolDBSuccess Then
    Else
        r_strReturnMsgs(0) = "Please try again. GL Account Validation Web Service unavailable at this time."
        
'        r_strReturnCodes(0) = 2
'        r_strReturnMsgs(0) = "Call to Account Validation Web Service failed"
'        ValidateGLAccounting = False
'        Err.Raise trErrbusValidateGLAccoutingStoredProcedureFailed
'        Exit Function
    End If
    
'    If Trim(r_strDistributionIDsReturned(0)) = "" Then
'        r_strReturnMsgs(0) = "Please try again. GL Account Validation Web Service unavailable at this time."
'    End If
    
    ValidateGLAccountingWS = True

    Set l_obj = Nothing
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "ValidateGLAccountingWS"
            
End Function

Public Function GetCICSRegion(ByVal v_strSubsystem As String) As String
    Select Case (UCase(v_strSubsystem))
    Case "DB2A"
        GetCICSRegion = "SY7221"
    Case "DB2B"
        GetCICSRegion = "SY7321"
    Case "DB2P"
        GetCICSRegion = "SY7521"
    Case "DB2Q"
        GetCICSRegion = "SY7621"
    Case "DB2T"
        GetCICSRegion = "SY7721"
    Case "DB2D"
        GetCICSRegion = "SY7821"
    Case "DB2E"
        GetCICSRegion = "SY7921"
    Case Else
        'All test DB2 servers will go against
        'CICS in region B
        GetCICSRegion = "SY7321"
    End Select
    Exit Function
End Function

Private Function ConvertEditArrayToString( _
            ByVal v_strDistributionIDEdits As Variant _
            ) As String
            
    Dim l_intI As Integer
    Dim l_strDistIDEdit As String
    Dim l_strDistIDEdits As String
    

    l_strDistIDEdit = ""
    l_strDistIDEdits = ""
    For l_intI = 0 To 221
        l_strDistIDEdit = v_strDistributionIDEdits(l_intI)
        l_strDistIDEdits = l_strDistIDEdits & l_strDistIDEdit & Space(103)
    Next l_intI
    
    ConvertEditArrayToString = l_strDistIDEdits
    
            
End Function
Public Function ValidateWorkRequestsGetWR( _
            ByVal v_aryDOJMDistricts As Variant, _
            ByVal v_aryDOJMWorkRequests As Variant, _
            ByVal v_aryDOJMHeadQuarters As Variant) As ADODB.Recordset

'This function is used for an existance check validation
' on the VOMWRKRQ_00R table
'It retrieves the DOJM District and Work Request Number
' for each one passed in.
' The busDistributionMgr will check to see if those
'  passed in came back.
'  Chect the VOMWRKRQ_00R table for the DOJM District and Work Request Number
' RED078

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo ValidateWorkRequestsGetWR_EH
    
     
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT2")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
'    l_strSQL = "SELECT  CDE_DIST, CDE_HDQTR_CRW, CDE_WR "
'    l_strSQL = l_strSQL & "FROM METHDBA.VOMWRKRQ_00R "
'    l_strSQL = l_strSQL & "WHERE   ((CDE_STAT IN ('60','6X','70','7X','80','8X') AND CDE_TYP_WR = 'SI') "
'    l_strSQL = l_strSQL & "OR  (CDE_STAT IN ('60','6X','70','7X','80','8X', '90', '9X') AND CDE_TYP_WR = 'SR')) "
'    l_strSQL = l_strSQL & "AND " & BuildInClause("CDE_DIST", v_aryDOJMDistricts, True) & " "
'    l_strSQL = l_strSQL & "AND " & BuildInClause("CDE_WR", v_aryDOJMWorkRequests, True) & " "
'    l_strSQL = l_strSQL & "AND " & BuildInClause("SUBSTR(CDE_HDQTR_CRW,1,3)", v_aryDOJMHeadQuarters, True) & " "
'    l_strSQL = l_strSQL & "Order by 1,2,3"

'SELECT  CDE_DIST, CDE_HDQTR_CRW, CDE_WR
'From(
'SELECT  CDE_DIST, CDE_HDQTR_CRW, CDE_WR
'From METHDBA.TOMWRKRQ
'WHERE CDE_DIST IN ('2WWZ')
'AND CDE_WR IN ('073378', '075378', '075299')
'AND   ((CDE_STAT IN ('60','6X','70','7X','80','8X') AND CDE_TYP_WR = 'SI')
'OR  (CDE_STAT IN ('60','6X','70','7X','80','8X', '90', '9X') AND CDE_TYP_WR = 'SR'))
'//'GROUP BY CDE_DIST, CDE_HDQTR_CRW, CDE_WR
') NE1
'WHERE SUBSTR(CDE_HDQTR_CRW,1,3) IN ('06W')
'GROUP BY CDE_DIST, CDE_HDQTR_CRW, CDE_WR;

    l_strSQL = "SELECT CDE_DIST, CDE_HDQTR_CRW, CDE_WR"
    l_strSQL = l_strSQL & " FROM ("
    l_strSQL = l_strSQL & "SELECT CDE_DIST, CDE_HDQTR_CRW, CDE_WR"
    l_strSQL = l_strSQL & " FROM METHDBA.TOMWRKRQ"
    l_strSQL = l_strSQL & " WHERE " & BuildInClause("CDE_DIST", v_aryDOJMDistricts, True)
    l_strSQL = l_strSQL & " AND " & BuildInClause("CDE_WR", v_aryDOJMWorkRequests, True)
    l_strSQL = l_strSQL & " AND   ((CDE_STAT IN ('60','6X','70','7X','80','8X') AND CDE_TYP_WR = 'SI')"
    l_strSQL = l_strSQL & " OR  (CDE_STAT IN ('60','6X','70','7X','80','8X', '90', '9X') AND CDE_TYP_WR = 'SR'))"
    l_strSQL = l_strSQL & " ) NE1"
    l_strSQL = l_strSQL & " WHERE " & BuildInClause("SUBSTR(CDE_HDQTR_CRW,1,3)", v_aryDOJMHeadQuarters, True)
    l_strSQL = l_strSQL & " GROUP BY CDE_DIST, CDE_HDQTR_CRW, CDE_WR"
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objrs.ActiveConnection = Nothing
    Set ValidateWorkRequestsGetWR = l_objrs
    l_objCN.Close
    
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
Exit Function

ValidateWorkRequestsGetWR_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "ValidateWorkRequestsGetWR", , , , DescArgsFrom(l_strSQL), True
    
End Function
'Public Function ValidateJobTasksGetJT( _
'            ByVal v_aryEMPRVJobTasks As Variant) As ADODB.Recordset

'This function is used for an existance check validation
' on the emprv_common table

'    Dim l_objCN As New ADODB.Connection
'    Dim l_objrs As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
'On Error GoTo ValidateJobTasksGetJT_EH
    
'    Set l_objCN = New ADODB.Connection
'    Set l_objrs = New ADODB.Recordset
    'Set l_objCM = New clsuDBConnectMgr
   
    'l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT2")
'    l_objCN.ConnectionString = "DSN=DB2QLUW;UID=DB2ADMIN;PWD=DB2ADMIN"
    
    'Set l_objCM = Nothing
    
'    l_objCN.Open
    
'    l_strSQL = "SELECT JOBTASK_NBR, DESCRIPTION, DISTRIBUTION_ID, WG_NUM " & _
'    l_strSQL = ", EMPRV_DATABASE_ID, COST_CENTER " & _
'    l_strSQL = " FROM DB2ADMIN.TJOBTASK "
'    l_strSQL = " WHERE " & BuildInClause("JOBTASK_NBR", v_aryEMPRVJobTasks, True)
    
'    l_objrs.CursorLocation = adUseClient
'    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

'    l_objrs.ActiveConnection = Nothing
    
    
'    Set ValidateJobTasksGetJT = l_objrs
'    l_objCN.Close
    
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objrs = Nothing
    
'Exit Function

'ValidateJobTasksGetJT_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objrs = Nothing
    
'    TrsRaiseError m_strCMODULENAME, "ValidateJobTasksGetJT", , , , DescArgsFrom(l_strSQL), True
    
'End Function

Public Function GetECAccountNumNonProdTime() As String
    
    TrsRaiseError m_strCMODULENAME, "GetECAccountNumNonProdTime", trErrObsoleteMethod, , , , True
'    Dim lstrAccountNumber As String
'
'On Error GoTo GetECAccountNumNonProdTime_EH
''field changed to 39 characters for AFIP:
''                       "123456789012345678901234567890123456789"
'    lstrAccountNumber = "12118404220009C                        "
'
'    GetECAccountNumNonProdTime = lstrAccountNumber
'
'Exit Function
'GetECAccountNumNonProdTime_EH:
'
'    TrsRaiseError m_strCMODULENAME, "GetECAccountNumNonProdTime", , , , , True

End Function

Public Function GetECDescriptionNonProdTime() As String
    TrsRaiseError m_strCMODULENAME, "GetECAccountNumNonProdTime", trErrObsoleteMethod, , , , True
'    Dim lstrDescriptionNonProd As String
'
'On Error GoTo GetECDescriptionNonProdTime_EH
'
'    lstrDescriptionNonProd = "Non-productive/Productive"
'
'    GetECDescriptionNonProdTime = lstrDescriptionNonProd
'
'Exit Function
'GetECDescriptionNonProdTime_EH:
'
'    TrsRaiseError m_strCMODULENAME, "GetECDescriptionNonProdTime", , , , , True
'
End Function

'//-----wrapper so that COM+ partition can be used-----//
Public Function GetDB2Subsystem(Optional ByVal dataProvider As String) As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'    Dim strValue As String, strProvider As String
'    Dim blnReturn As String
'    Dim strFullSubKey As String
'    Dim strValueName As String
'    Dim ral As REGTool5.Registry
    
    
On Error GoTo eh
   
    Set l_objCM = New clsuDBConnectMgr
    
    GetDB2Subsystem = l_objCM.GetDB2Subsystem(dataProvider)

'    Set ral = New REGTool5.Registry
'
'    If dataProvider = "" Then
'        strProvider = defaultProvider
'    Else
'        strProvider = dataProvider
'    End If
'
'    strFullSubKey = regSubKey & strProvider
'    strValueName = "DB2SubSys"
'
'    blnReturn = ral.GetKeyValue( _
'            HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, strValue)
'    GetDB2Subsystem = strValue
'    Set ral = Nothing
    Set l_objCM = Nothing
    
    Exit Function
eh:
    Set l_objCM = Nothing
    TrsRaiseError _
            m_strCMODULENAME, _
            "GetDB2Subsystem", _
            , _
            True, _
            DescArgsFrom(dataProvider)

'    Set ral = Nothing
'    TrsRaiseError _
'            m_strCMODULENAME, _
'            "GetDB2Subsystem", _
'            , _
'            True, _
'            DescArgsFrom(strFullSubKey, strValueName)

End Function

Private Function GetMAXIMOJobNumbers(ByVal v_strDeptNum As String, ByVal v_lnWgNum As Long) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    Dim job As Ameren_TRIS_Server_COMIntegration4.MAXIMO
On Error GoTo eh

    Set job = New Ameren_TRIS_Server_COMIntegration4.MAXIMO
    
    Set rs = job.GetMaximoJobsVB6_(v_strDeptNum, v_lnWgNum)
    
    Set GetMAXIMOJobNumbers = rs
    
    Set job = Nothing
    Exit Function
eh:
    Set job = Nothing
    TrsRaiseError m_strCMODULENAME, "GetMAXIMOJobNumbers"
End Function

Public Function GetDistJobTaskListbyJob(ByVal v_strJobNum As String, ByVal v_strTaskNum As String) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    Dim job As Ameren_TRIS_Server_COMIntegration4.MAXIMO
On Error GoTo eh

    Set job = New Ameren_TRIS_Server_COMIntegration4.MAXIMO
    
    Set rs = job.GetMaximoJobVB6_(v_strJobNum, v_strTaskNum)
    
    Set GetDistJobTaskListbyJob = rs
    Set job = Nothing
    Exit Function
eh:
    Set job = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDistJobTaskListbyJob"
End Function

Public Function GetDistJobTaskListbyJob2(ByVal v_strJobNum As String, ByVal v_strTaskNum As String, ByVal v_strSiteId As String) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    Dim job As Ameren_TRIS_Server_COMIntegration4.MAXIMO
On Error GoTo eh

    Set job = New Ameren_TRIS_Server_COMIntegration4.MAXIMO
    
    Set rs = job.GetMaximoJob2VB6_(v_strJobNum, v_strTaskNum, v_strSiteId)
    
    Set GetDistJobTaskListbyJob2 = rs
    Set job = Nothing
    Exit Function
eh:
    Set job = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDistJobTaskListbyJob2"
End Function

Public Function GetDistJobTaskListMultiple(ByVal v_rsActualTime As ADODB.Recordset) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    Dim job As Ameren_TRIS_Server_COMIntegration4.MAXIMO
On Error GoTo eh

    Set job = New Ameren_TRIS_Server_COMIntegration4.MAXIMO
    
    Set rs = job.GetMaximoJobMultipleVB6_(v_rsActualTime)
    
    Set GetDistJobTaskListMultiple = rs
    Set job = Nothing
    Exit Function
eh:
    Set job = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDistJobTaskListMultiple"
End Function

Private Function ValidateGLAccountingAPI( _
            ByVal v_strDistributionIDEdits As Variant, _
            ByVal v_intNumberOfEdits As Integer, _
            ByRef r_strDistributionIDsReturned As Variant, _
            ByRef r_strReturnCodes As Variant, _
            ByRef r_strReturnMsgs As Variant, _
            ByRef r_strHighlightFields As Variant _
            ) As Boolean
            
    Dim l_obj As Ameren_TRIS_Server_COMIntegration4.ValidateDistribution
    Dim l_bolDBSuccess As Boolean
    Dim l_strEmptyArray1(221) As String
    Dim l_strEmptyArray2(221) As String
    Dim l_strEmptyArray3(221) As String
    Dim l_strEmptyArray4(221) As String
    
    r_strDistributionIDsReturned = l_strEmptyArray1
    r_strReturnCodes = l_strEmptyArray2
    r_strReturnMsgs = l_strEmptyArray3
    r_strHighlightFields = l_strEmptyArray4
    
    On Error GoTo eh
    
    Set l_obj = New Ameren_TRIS_Server_COMIntegration4.ValidateDistribution
    
    '//if the return value of false previously meant that the call to the stored procedure failed
    '//now a value of false means that the validation failed, let any failures result in errors/exceptions
    ReDim Preserve v_strDistributionIDEdits(v_intNumberOfEdits - 1) '//only send what is needed to web service instead of a fixed 222 (0-221) array
    l_bolDBSuccess = l_obj.ValidateDistributionNumbers_(v_strDistributionIDEdits, v_intNumberOfEdits, r_strDistributionIDsReturned, r_strReturnCodes, r_strReturnMsgs, r_strHighlightFields)
    
    If l_bolDBSuccess Then
    Else
        '//see busDistributionMgr.cbsDistributionMgr ValidateDistributionIDsGenerateBRVs2 for setting the brv being returned to client
        '//actual error message internal error 500 is not that usefull anyway
        r_strReturnMsgs(0) = "Please try again. GL Account Validation API unavailable at this time."
        
'        r_strReturnCodes(0) = 2
'        r_strReturnMsgs(0) = "Call to Account Validation Web Service failed"
'        ValidateGLAccounting = False
'        Err.Raise trErrbusValidateGLAccoutingStoredProcedureFailed
'        Exit Function
    End If
    
    
    ValidateGLAccountingAPI = l_bolDBSuccess '//True

'    l_obj.ReleaseNetObject_
    Set l_obj = Nothing
    Exit Function
eh:
'    If Not l_obj Is Nothing Then
'        l_obj.ReleaseNetObject_
'    End If
    
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "ValidateGLAccountingAPI"
            
End Function

Private Sub Class_Initialize()
    Set m_ResourceMgr = New utilADOHelper.cuResourceMgr
End Sub

Private Sub Class_Terminate()
    Set m_ResourceMgr = Nothing
End Sub

