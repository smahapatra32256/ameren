VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdBuiltDataScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Module Name
Private Const MODULENAME = "dbSecurity.cdBuiltDataScope"
Private Enum DataScopeLevel
    DSEmployee = 60
    DSWorkGroup = 50
    DSDepartment = 30
    DSFunction = 20
    DSDivision = 10
End Enum

Public Function GetAllDataScopes(ByVal intEmpNum As Long _
            , ByVal boolActive As Boolean _
            , ByRef radrsEMP As ADODB.Recordset _
            , ByRef radrsWG As ADODB.Recordset _
            , ByRef radrsDWG As ADODB.Recordset _
            , ByRef radrsNWG As ADODB.Recordset _
            , ByRef radrsDEPT As ADODB.Recordset _
            , ByRef radrsFUN As ADODB.Recordset _
            , ByRef radrsDIV As ADODB.Recordset _
            , ByRef radrsAEEMP As ADODB.Recordset _
            , ByRef radrsAEWG As ADODB.Recordset _
            , ByRef radrsAEDEPT As ADODB.Recordset _
            , ByRef radrsAEFUN As ADODB.Recordset _
            , ByRef radrsAEDIV As ADODB.Recordset _
            , Optional ByVal vbFilterWorkgroups As Boolean = True _
            , Optional ByVal vbFilterApprovalEntry As Boolean = True)
            
    GetAllDataScopes = GetAllDataScopes3(intEmpNum, boolActive, _
            radrsEMP, radrsWG, radrsDWG, radrsNWG, _
            radrsDEPT, radrsFUN, _
            radrsDIV, radrsAEEMP, _
            radrsAEWG, radrsAEDEPT, _
            radrsAEFUN, radrsAEDIV, _
            vbFilterWorkgroups, vbFilterApprovalEntry, False)
End Function

Public Function GetAllDataScopes3(ByVal intEmpNum As Long _
            , ByVal boolActive As Boolean _
            , ByRef radrsEMP As ADODB.Recordset _
            , ByRef radrsWG As ADODB.Recordset _
            , ByRef radrsDWG As ADODB.Recordset _
            , ByRef radrsNWG As ADODB.Recordset _
            , ByRef radrsDEPT As ADODB.Recordset _
            , ByRef radrsFUN As ADODB.Recordset _
            , ByRef radrsDIV As ADODB.Recordset _
            , ByRef radrsAEEMP As ADODB.Recordset _
            , ByRef radrsAEWG As ADODB.Recordset _
            , ByRef radrsAEDEPT As ADODB.Recordset _
            , ByRef radrsAEFUN As ADODB.Recordset _
            , ByRef radrsAEDIV As ADODB.Recordset _
            , Optional ByVal vbFilterWorkgroups As Boolean = True _
            , Optional ByVal vbFilterApprovalEntry As Boolean = True _
            , Optional ByVal vbForceRebuild As Boolean = False)
            
    Dim sSQL As String
    
    Set radrsEMP = GetDataScopeEmp3(intEmpNum, boolActive, vbForceRebuild)
    Set radrsWG = GetDataScopeWG3(intEmpNum, boolActive, vbForceRebuild)
    Set radrsDEPT = GetDataScopeDept3(intEmpNum, vbForceRebuild)
    Set radrsFUN = GetDataScopeFun3(intEmpNum, vbForceRebuild)
    Set radrsDIV = GetDataScopeDiv3(intEmpNum, vbForceRebuild)

    '//-----let the db do the filtering-----//
    If vbFilterWorkgroups Then
        sSQL = m_strDATASCOPE_DWG_select_sql
        Set radrsDWG = GetDataScope(sSQL, intEmpNum)
        
        sSQL = m_strDATASCOPE_NWG_select_sql
        Set radrsNWG = GetDataScope(sSQL, intEmpNum)
    End If
    
    If vbFilterApprovalEntry Then
        sSQL = m_strDATASCOPE_AEEMP_select_sql
        Set radrsAEEMP = GetDataScope(sSQL, intEmpNum)
    
        sSQL = m_strDATASCOPE_AEWG_select_sql
        Set radrsAEWG = GetDataScope(sSQL, intEmpNum)
    
        sSQL = m_strDATASCOPE_AEDEPT_select_sql
        Set radrsAEDEPT = GetDataScope(sSQL, intEmpNum)
    
        sSQL = m_strDATASCOPE_AEFUN_select_sql
        Set radrsAEFUN = GetDataScope(sSQL, intEmpNum)
    
        sSQL = m_strDATASCOPE_AEDIV_select_sql
        Set radrsAEDIV = GetDataScope(sSQL, intEmpNum)
    End If
End Function

'//-----new for force rebuild flag
Public Function GetDataScopeEmp3(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean, _
                                ByVal vbForceRebuild As Boolean) As ADODB.Recordset
    Dim sSQL As String '//= Resource.GetStringResource("DATASCOPE_EMP_select.sql")
    sSQL = m_strDATASCOPE_EMP_select_sql
    CheckAndRebuildDataScopeEmp intEmpNum, boolActive, vbForceRebuild, False '//incremental build

    Set GetDataScopeEmp3 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeWG3(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean, _
                                ByVal vbForceRebuild As Boolean) As ADODB.Recordset
    Dim sSQL As String '//= Resource.GetStringResource("DATASCOPE_WG_select.sql")
    sSQL = m_strDATASCOPE_WG_select_sql
    CheckAndRebuildDataScopeWG intEmpNum, boolActive, vbForceRebuild

    Set GetDataScopeWG3 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeDept3(ByVal intEmpNum As Long, _
                                    ByVal vbForceRebuild As Boolean) As ADODB.Recordset
    Dim sSQL As String '//= Resource.GetStringResource("DATASCOPE_DEPT_select.sql")
    sSQL = m_strDATASCOPE_DEPT_select_sql
    CheckAndRebuildDataScopeDept intEmpNum, vbForceRebuild

    Set GetDataScopeDept3 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeFun3(ByVal intEmpNum As Long, _
                                ByVal vbForceRebuild As Boolean) As ADODB.Recordset
    Dim sSQL As String '//= Resource.GetStringResource("DATASCOPE_FUN_select.sql")
    sSQL = m_strDATASCOPE_FUN_select_sql
    CheckAndRebuildDataScopeFun intEmpNum, vbForceRebuild

    Set GetDataScopeFun3 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeDiv3(ByVal intEmpNum As Long, _
                                ByVal vbForceRebuild As Boolean) As ADODB.Recordset
    Dim sSQL As String '//= Resource.GetStringResource("DATASCOPE_DIV_select.sql")
    sSQL = m_strDATASCOPE_DIV_select_sql
    CheckAndRebuildDataScopeDiv intEmpNum, vbForceRebuild

    Set GetDataScopeDiv3 = GetDataScope(sSQL, intEmpNum)
End Function

'//-----
Public Function GetDataScopeEmp(ByVal intEmpNum As Long, _
                                        ByVal boolActive As Boolean) As ADODB.Recordset
    Dim sSQL As String '//= Resource.GetStringResource("DATASCOPE_EMP_select.sql")
    sSQL = m_strDATASCOPE_EMP_select_sql
    CheckAndRebuildDataScopeEmp intEmpNum, boolActive, False, True '//full rebuild

    Set GetDataScopeEmp = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeWG(ByVal intEmpNum As Long, _
                                        ByVal boolActive As Boolean) As ADODB.Recordset
    Dim sSQL As String '//= Resource.GetStringResource("DATASCOPE_WG_select.sql")
    sSQL = m_strDATASCOPE_WG_select_sql
    CheckAndRebuildDataScopeWG intEmpNum, boolActive, False

    Set GetDataScopeWG = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeDept(ByVal intEmpNum As Long) As ADODB.Recordset
    Dim sSQL As String '//= Resource.GetStringResource("DATASCOPE_DEPT_select.sql")
    sSQL = m_strDATASCOPE_DEPT_select_sql
    CheckAndRebuildDataScopeDept intEmpNum, False

    Set GetDataScopeDept = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeFun(ByVal intEmpNum As Long) As ADODB.Recordset
    Dim sSQL As String '//= Resource.GetStringResource("DATASCOPE_FUN_select.sql")
    sSQL = m_strDATASCOPE_FUN_select_sql
    CheckAndRebuildDataScopeFun intEmpNum, False

    Set GetDataScopeFun = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeDiv(ByVal intEmpNum As Long) As ADODB.Recordset
    Dim sSQL As String '//= Resource.GetStringResource("DATASCOPE_DIV_select.sql")
    sSQL = m_strDATASCOPE_DIV_select_sql
    CheckAndRebuildDataScopeDiv intEmpNum, False

    Set GetDataScopeDiv = GetDataScope(sSQL, intEmpNum)
End Function

Private Function GetDataScope(ByVal sql As String, ByVal intEmpNum As Long) As Recordset
    Dim ladrs As ADODB.Recordset
    
    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = intEmpNum
        .Update
    End With
    
    Set GetDataScope = NewExecuteSQL(sql _
                        , ladrs _
                        , adOpenStatic _
                        , adLockBatchOptimistic)
    '//-----for compatability
    If Not GetDataScope Is Nothing Then
        If Not (GetDataScope.BOF And GetDataScope.EOF) Then
        Else
            Set GetDataScope = Nothing
        End If
    Else
        Set GetDataScope = Nothing
    End If
    
    Set ladrs = Nothing

End Function

'//TODO implement data dictionary for creating parms
'//-----parms are eff_emp_num, emp_num, wg_num
Private Function InvalidateDataScopeInt(ByVal sql As String, ByVal intKey As Long) As Boolean
    Dim ladrs As ADODB.Recordset
    Dim ladcn As ADODB.Connection
        
    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = intKey
        .Update
    End With
    
    Set ladcn = OpenConnection
    NewExecuteOtherSQL sql, ladrs, ladcn
    
    ladcn.Close
    Set ladcn = Nothing
    Set ladrs = Nothing
End Function

'//-----having a problem when using parameters and the update returns no rows, this is ok
Private Function InvalidateDataScopeInt2(ByVal sql As String, ByVal intKey As Long) As Boolean
    Dim ladcn As ADODB.Connection
    Dim sSQL As String
    
    sSQL = Replace(sql, "?", intKey)
    
    Set ladcn = OpenConnection
    ExecuteOTHERSQL sSQL, ladcn
    
    ladcn.Close
    Set ladcn = Nothing
End Function

'//-----having a problem when using parameters and the update returns no rows, this is ok
Private Function InvalidateDataScopeStr2(ByVal sql As String, ByVal strKey As String) As Boolean
    Dim ladcn As ADODB.Connection
    Dim sSQL As String
    
    sSQL = Replace(sql, "?", "'" & strKey & "'")
    
    Set ladcn = OpenConnection
    ExecuteOTHERSQL sSQL, ladcn
    
    ladcn.Close
    Set ladcn = Nothing
End Function

'Private Function InvalidateDataScopeIntTest2(ByVal sql As String, ByVal intKey As Long) As Boolean
'    Dim ladrs As ADODB.Recordset
'    Dim ladcn As ADODB.Connection
'
'    Set ladrs = CreateParameterRS
'    With ladrs
'        .AddNew
'        !PARM_ORDER = 1
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = intKey
'        .Update
'    End With
'
'    Set ladcn = OpenConnection
'    NewExecuteSQLWithConnection sql, ladrs, ladcn, adOpenStatic, adLockBatchOptimistic, adCmdText
'
'    ladcn.Close
'    Set ladcn = Nothing
'    Set ladrs = Nothing
'End Function

'//-----parms are dept_num, function_num, division_num
Private Function InvalidateDataScopeStr(ByVal sql As String, ByVal strKey As String) As Boolean
    Dim ladrs As ADODB.Recordset
    Dim ladcn As ADODB.Connection
    
    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adChar
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 3
        !PARM_VALUE = strKey
        .Update
    End With
    
    Set ladcn = OpenConnection
    NewExecuteOtherSQL sql, ladrs, ladcn
    
    ladcn.Close
    Set ladcn = Nothing
    Set ladrs = Nothing
End Function

Private Function IsDataScopeBuilt(ByVal eDataScopeLevel As DataScopeLevel, _
                                    ByVal sql As String, _
                                    ByVal intEmpNum As Long, _
                                    ByVal blnActiveOnly As Boolean) As Boolean
    Dim ladrs As ADODB.Recordset
    Dim rsDataScope As ADODB.Recordset
    
    IsDataScopeBuilt = False
    
    Set ladrs = CreateParameterRS
    With ladrs
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = intEmpNum
        .Update
        
        If eDataScopeLevel = DSEmployee Or eDataScopeLevel = DSWorkGroup Then
            .AddNew
            !PARM_ORDER = 2
            !PARM_NAME = ""
            !PARM_TYPE = adChar
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 1
            !PARM_VALUE = IIf(blnActiveOnly, "Y", "N")
            .Update
        End If
    End With
    
'    '//-----test
'    sql = Replace(sql, "?", intEmpNum, 1, 1)
'    sql = Replace(sql, "?", IIf(blnActiveOnly, "'Y'", "'N'"), 1, 1)
'    IsDataScopeBuilt = IIf(SingletonQueryWithParms(sql, Nothing, Nothing) > 0, True, False)
    
    Set rsDataScope = NewExecuteSQL(sql _
                        , ladrs _
                        , adOpenStatic _
                        , adLockBatchOptimistic)
    If RecordsExist(rsDataScope) Then
        rsDataScope.MoveFirst
        If rsDataScope.Fields(0) > 0 Then '//SELECT COUNT(*)
            IsDataScopeBuilt = True
        End If
    End If
End Function

Private Sub CheckAndRebuildDataScopeEmp(ByVal intEmpNum As Long, _
                                        ByVal boolActive As Boolean, _
                                        ByVal blnForceRebuild As Boolean, _
                                        ByVal blnFullRebuild As Boolean)
    Dim DataScopeEmpCheckSQL As String '//= Resource.GetStringResource("DATASCOPE_EMP_check.sql")
    DataScopeEmpCheckSQL = m_strDATASCOPE_EMP_check_sql
    Dim DataScopeEmpRebuildSQL As String '//New Text.StringBuilder
    Dim blnIsDataScopeBuilt As Boolean
    Dim strWhereClause As String
    Dim dtLoadFromDate As Date
    
    If blnFullRebuild Then '//do only for the batch rebuild process
        If boolActive Then
            'Begin Prem TRIS 5.9+ April 11, 2007 - PTS 2024
            dtLoadFromDate = CDate(Format(DateAdd("m", -1, Date), "mm/dd/yyyy"))
            'End Prem TRIS 5.9+ April 11, 2007 - PTS 2024
            DataScopeEmpRebuildSQL = m_strDATASCOPE_EMP_ACTIVE_rebuild_sql
            DataScopeEmpRebuildSQL = Replace(DataScopeEmpRebuildSQL, "{0}", intEmpNum)
            DataScopeEmpRebuildSQL = Replace(DataScopeEmpRebuildSQL, "{1}", Format(dtLoadFromDate, "mm/dd/yyyy"))
    '        sbDataScopeEmpRebuildSQL.AppendFormat(Resource.GetStringResource("DATASCOPE_EMP_ACTIVE_rebuild.sql"), _
    '                                            intEmpNum, dtLoadFromDate.ToShortDateString())
        Else '//ALL employees using cutover # of months from sysparms
            dtLoadFromDate = CDate(Format(Date, "mm/dd/yyyy")) '//sql also uses months from sysparms 49, 1
            DataScopeEmpRebuildSQL = m_strDATASCOPE_EMP_ALL_rebuild_sql
            DataScopeEmpRebuildSQL = Replace(DataScopeEmpRebuildSQL, "{0}", intEmpNum)
            DataScopeEmpRebuildSQL = Replace(DataScopeEmpRebuildSQL, "{1}", Format(dtLoadFromDate, "mm/dd/yyyy"))
    '        sbDataScopeEmpRebuildSQL.AppendFormat(Resource.GetStringResource("DATASCOPE_EMP_ALL_rebuild.sql"), _
    '                                            intEmpNum)
        End If
    Else '//do for online only
        DataScopeEmpRebuildSQL = m_strDATASCOPE_EMP_rebuild_changes_sql
        
        If boolActive Then
            dtLoadFromDate = CDate(Format(DateAdd("m", -1, Date), "mm/dd/yyyy"))
            DataScopeEmpRebuildSQL = Replace(DataScopeEmpRebuildSQL, "{0}", intEmpNum)
            strWhereClause = "(B.EMP_STATUS <> 'T'" & vbCrLf & _
                            " OR (B.EMP_STATUS = 'T' AND B.STATUS_EFF_DATE > Cast('{1}' as Date)))"
            strWhereClause = Replace(strWhereClause, "{1}", Format(dtLoadFromDate, "mm/dd/yyyy"))
            DataScopeEmpRebuildSQL = Replace(DataScopeEmpRebuildSQL, "{2}", strWhereClause)
            
        Else '//ALL employees using cutover # of months from sysparms
            dtLoadFromDate = CDate(Format(Date, "mm/dd/yyyy")) '//sql also uses months from sysparms 49, 1
            DataScopeEmpRebuildSQL = Replace(DataScopeEmpRebuildSQL, "{0}", intEmpNum)
            strWhereClause = "(B.EMP_STATUS <> 'T'" & vbCrLf & _
                            " OR (B.EMP_STATUS = 'T' AND B.STATUS_EFF_DATE > Cast('{1}' as Date) - P.ALL_EMP_CUTOFF_MONTHS months))"
            strWhereClause = Replace(strWhereClause, "{1}", Format(dtLoadFromDate, "mm/dd/yyyy"))
            DataScopeEmpRebuildSQL = Replace(DataScopeEmpRebuildSQL, "{2}", strWhereClause)
        End If
    End If
    
    If Not blnForceRebuild Then
        blnIsDataScopeBuilt = IsDataScopeBuilt(DSEmployee, DataScopeEmpCheckSQL, intEmpNum, boolActive)
    End If
    
    If Not blnIsDataScopeBuilt Or blnForceRebuild Then
        ExecuteSQL DataScopeEmpRebuildSQL, adOpenStatic, adLockReadOnly
    End If
End Sub

Private Sub CheckAndRebuildDataScopeWG(ByVal intEmpNum As Long, _
                                        ByVal boolActive As Boolean, _
                                        ByVal blnForceRebuild As Boolean)
    Dim DataScopeWGCheckSQL As String '//= Resource.GetStringResource("DATASCOPE_WG_check.sql")
    DataScopeWGCheckSQL = m_strDATASCOPE_WG_check_sql
    Dim DataScopeWGRebuildSQL As String '//New Text.StringBuilder
    Dim dtLoadFromDate As Date
    Dim blnIsDataScopeBuilt As Boolean
        
'    If boolActive Then
        'Begin Prem TRIS 5.9+ April 11, 2007 - PTS 2024
        dtLoadFromDate = CDate(Format(DateAdd("m", -1, Date), "mm/dd/yyyy"))
        'Begin Prem TRIS 5.9+ April 11, 2007 - PTS 2024
''        dtLoadFromDate = Date.Now.AddMonths(-6)
'    Else
'        'dtLoadFromDate = New Date(1900, 1, 1)
'        dtLoadFromDate = #1/1/1900#
'    End If

    DataScopeWGRebuildSQL = m_strDATASCOPE_WG_rebuild_sql
    DataScopeWGRebuildSQL = Replace(DataScopeWGRebuildSQL, "{0}", intEmpNum)
    DataScopeWGRebuildSQL = Replace(DataScopeWGRebuildSQL, "{1}", dtLoadFromDate)

'    sbDataScopeWGRebuildSQL.AppendFormat(Resource.GetStringResource("DATASCOPE_WG_rebuild.sql"), _
'                        intEmpNum, dtLoadFromDate.ToShortDateString())

    If Not blnForceRebuild Then
        blnIsDataScopeBuilt = IsDataScopeBuilt(DSWorkGroup, DataScopeWGCheckSQL, intEmpNum, True)
    End If
    
    If Not blnIsDataScopeBuilt Or blnForceRebuild Then
        ExecuteSQL DataScopeWGRebuildSQL, adOpenStatic, adLockReadOnly
'        Dim rc As Integer = Database.ExecSQLInTrans(sbDataScopeWGRebuildSQL.ToString())
    End If
End Sub

Private Sub CheckAndRebuildDataScopeDept(ByVal intEmpNum As Long, _
                                        ByVal blnForceRebuild As Boolean)
    Dim DataScopeDeptCheckSQL As String '//= Resource.GetStringResource("DATASCOPE_DEPT_check.sql")
    DataScopeDeptCheckSQL = m_strDATASCOPE_DEPT_check_sql
    Dim DataScopeDeptRebuildSQL As String '//New Text.StringBuilder
    Dim blnIsDataScopeBuilt As Boolean
    
    DataScopeDeptRebuildSQL = m_strDATASCOPE_DEPT_rebuild_sql
    DataScopeDeptRebuildSQL = Replace(DataScopeDeptRebuildSQL, "{0}", intEmpNum)

    '//sbDataScopeDeptRebuildSQL.AppendFormat(Resource.GetStringResource("DATASCOPE_DEPT_rebuild.sql"), intEmpNum)

    If Not blnForceRebuild Then
        blnIsDataScopeBuilt = IsDataScopeBuilt(DSDepartment, DataScopeDeptCheckSQL, intEmpNum, True)
    End If
    
    If Not blnIsDataScopeBuilt Or blnForceRebuild Then
        ExecuteSQL DataScopeDeptRebuildSQL, adOpenStatic, adLockReadOnly
        '//Dim rc As Integer = Database.ExecSQLInTrans(sbDataScopeDeptRebuildSQL.ToString())
    End If
End Sub

Private Sub CheckAndRebuildDataScopeFun(ByVal intEmpNum As Long, _
                                        ByVal blnForceRebuild As Boolean)
    Dim DataScopeFunCheckSQL As String '//= Resource.GetStringResource("DATASCOPE_FUN_check.sql")
    DataScopeFunCheckSQL = m_strDATASCOPE_FUN_check_sql
    Dim DataScopeFunRebuildSQL As String '//New Text.StringBuilder
    Dim blnIsDataScopeBuilt As Boolean
    
    DataScopeFunRebuildSQL = m_strDATASCOPE_FUN_rebuild_sql
    DataScopeFunRebuildSQL = Replace(DataScopeFunRebuildSQL, "{0}", intEmpNum)
    
    '//sbDataScopeFunRebuildSQL.AppendFormat(Resource.GetStringResource("DATASCOPE_FUN_rebuild.sql"), intEmpNum)

    If Not blnForceRebuild Then
        blnIsDataScopeBuilt = IsDataScopeBuilt(DSFunction, DataScopeFunCheckSQL, intEmpNum, True)
    End If
    
    If Not blnIsDataScopeBuilt Or blnForceRebuild Then
        ExecuteSQL DataScopeFunRebuildSQL, adOpenStatic, adLockReadOnly
        '//Dim rc As Integer = Database.ExecSQLInTrans(sbDataScopeFunRebuildSQL.ToString())
    End If
End Sub

Private Sub CheckAndRebuildDataScopeDiv(ByVal intEmpNum As Long, _
                                        ByVal blnForceRebuild As Boolean)
    Dim DataScopeDivCheckSQL As String '//= Resource.GetStringResource("DATASCOPE_DIV_check.sql")
    DataScopeDivCheckSQL = m_strDATASCOPE_DIV_check_sql
    Dim DataScopeDivRebuildSQL As String '//New Text.StringBuilder
    Dim blnIsDataScopeBuilt As Boolean
    
    DataScopeDivRebuildSQL = m_strDATASCOPE_DIV_rebuild_sql
    DataScopeDivRebuildSQL = Replace(DataScopeDivRebuildSQL, "{0}", intEmpNum)
    
    '//sbDataScopeDivRebuildSQL.AppendFormat(Resource.GetStringResource("DATASCOPE_DIV_rebuild.sql"), intEmpNum, intEmpNum, intEmpNum)

    If Not blnForceRebuild Then
        blnIsDataScopeBuilt = IsDataScopeBuilt(DSDivision, DataScopeDivCheckSQL, intEmpNum, True)
    End If
    
    If Not blnIsDataScopeBuilt Or blnForceRebuild Then
        ExecuteSQL DataScopeDivRebuildSQL, adOpenStatic, adLockReadOnly
        '//Dim rc As Integer = Database.ExecSQLInTrans(sbDataScopeDivRebuildSQL.ToString())
    End If
End Sub

'//--------------------------------------------------------------------------------//
'//-----expose the following individually for calling from several threads-----//
'//--------------------------------------------------------------------------------//
Public Sub CheckAndRebuildDataScopeEmp2(ByVal intEmpNum As Long, _
                                        ByVal boolActive As Boolean)
    CheckAndRebuildDataScopeEmp intEmpNum, boolActive, False, True '//full rebuild
                                                
End Sub

Public Sub CheckAndRebuildDataScopeWG2(ByVal intEmpNum As Long, _
                                        ByVal boolActive As Boolean)
    CheckAndRebuildDataScopeWG intEmpNum, boolActive, False
                                                
End Sub

Public Sub CheckAndRebuildDataScopeDept2(ByVal intEmpNum As Long)

    CheckAndRebuildDataScopeDept intEmpNum, False
                                                
End Sub

Public Sub CheckAndRebuildDataScopeFun2(ByVal intEmpNum As Long)

    CheckAndRebuildDataScopeFun intEmpNum, False
                                                
End Sub

Public Sub CheckAndRebuildDataScopeDiv2(ByVal intEmpNum As Long)

    CheckAndRebuildDataScopeDiv intEmpNum, False
                                                
End Sub

Public Function GetDataScopeEmp2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_EMP_select_sql

    Set GetDataScopeEmp2 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeWG2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_WG_select_sql

    Set GetDataScopeWG2 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeDept2(ByVal intEmpNum As Long) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_DEPT_select_sql

    Set GetDataScopeDept2 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeFun2(ByVal intEmpNum As Long) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_FUN_select_sql

    Set GetDataScopeFun2 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeDiv2(ByVal intEmpNum As Long) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_DIV_select_sql

    Set GetDataScopeDiv2 = GetDataScope(sSQL, intEmpNum)
End Function

'//-----for client filtering (approval/entry)-----//
Public Function GetDataScopeAEEmp2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_AEEMP_select_sql

    Set GetDataScopeAEEmp2 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeAEWG2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_AEWG_select_sql

    Set GetDataScopeAEWG2 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeAEDept2(ByVal intEmpNum As Long) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_AEDEPT_select_sql

    Set GetDataScopeAEDept2 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeAEFun2(ByVal intEmpNum As Long) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_AEFUN_select_sql

    Set GetDataScopeAEFun2 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeAEDiv2(ByVal intEmpNum As Long) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_AEDIV_select_sql

    Set GetDataScopeAEDiv2 = GetDataScope(sSQL, intEmpNum)
End Function

'//-----for client filtering (demand/non-demand)-----//
Public Function GetDataScopeDWG2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_DWG_select_sql

    Set GetDataScopeDWG2 = GetDataScope(sSQL, intEmpNum)
End Function

Public Function GetDataScopeNWG2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
    Dim sSQL As String
    sSQL = m_strDATASCOPE_NWG_select_sql

    Set GetDataScopeNWG2 = GetDataScope(sSQL, intEmpNum)
End Function

'//--------------------------------------------------------------------------------//
'//-----expose the following for Invalidation (Batch Process TTR_ASSOCIATIONS)-----//
'//--------------------------------------------------------------------------------//
Public Function InvalidateDataScopeByRefEmp(ByVal intEmpNum As Long, ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    
    InvalidateDataScopeEmpByRefEmp_Private intEmpNum
 
    GetObjectContext().SetComplete
    InvalidateDataScopeByRefEmp = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeByRefEmp" _
        , , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

Public Function InvalidateDataScopeByRefWG(ByVal intWGNum As Long, ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh

    InvalidateDataScopeByRefWG = InvalidateDataScopeByRefWG_Private(intWGNum, True, blnCascade)
'    InvalidateDataScopeWGByRefWG_Private intWGNum
'
'    If blnCascade Then
'        InvalidateDataScopeEmpByRefWG_Private intWGNum
'    End If
 
    GetObjectContext().SetComplete
    InvalidateDataScopeByRefWG = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeByRefWG" _
        , , , _
        , DescArgsFrom("WorkGroup Number = " & CStr(intWGNum)) _
        , True
End Function

Public Function InvalidateDataScopeByRefDept(ByVal strDeptNum As String, ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    Dim sSQL As String
    
    sSQL = m_strINVALIDATE_DS_By_Ref_Dept_sql
    
    InvalidateDataScopeStr2 sSQL, strDeptNum
'    InvalidateDataScopeDeptByRefDept_Private strDeptNum
'
'    If blnCascade Then
'        InvalidateDataScopeWGByRefDept_Private strDeptNum
'        InvalidateDataScopeEmpByRefDept_Private strDeptNum
'    End If
 
    GetObjectContext().SetComplete
    InvalidateDataScopeByRefDept = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeByRefDept" _
        , , , _
        , DescArgsFrom("Department Number = " & strDeptNum) _
        , True
End Function

Public Function InvalidateDataScopeByRefFun(ByVal strFunNum As String, ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    Dim sSQL As String
    
    sSQL = m_strINVALIDATE_DS_By_Ref_Fun_sql
    
    InvalidateDataScopeStr2 sSQL, strFunNum
'    InvalidateDataScopeFunByRefFun_Private strFunNum
'
'    If blnCascade Then
'        InvalidateDataScopeDeptByRefFun_Private strFunNum
'        InvalidateDataScopeWGByRefFun_Private strFunNum
'        InvalidateDataScopeEmpByRefFun_Private strFunNum
'    End If
 
    GetObjectContext().SetComplete
    InvalidateDataScopeByRefFun = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeByRefFun" _
        , , , _
        , DescArgsFrom("Function Number = " & strFunNum) _
        , True
End Function

Public Function InvalidateDataScopeByRefDiv(ByVal strDivNum As String, ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    Dim sSQL As String
        
    sSQL = m_strINVALIDATE_DS_By_Ref_Div_sql
    
    InvalidateDataScopeStr2 sSQL, strDivNum
'    InvalidateDataScopeDivByRefDiv_Private strDivNum
'
'    If blnCascade Then
'        InvalidateDataScopeFunByRefDiv_Private strDivNum
'        InvalidateDataScopeDeptByRefDiv_Private strDivNum
'        InvalidateDataScopeWGByRefDiv_Private strDivNum
'        InvalidateDataScopeEmpByRefDiv_Private strDivNum
'    End If
 
    GetObjectContext().SetComplete
    InvalidateDataScopeByRefDiv = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeByRefDiv" _
        , , , _
        , DescArgsFrom("Division Number = " & strDivNum) _
        , True
End Function

'//-----
'//--------------------------------------------------------------------------------//
'//-----expose the following for Invalidation use the approval and assocations tables-----//
'//--------------------------------------------------------------------------------//

'//--------------------------------------------------------------------------------//
'//-----expose the following for Invalidation (WorkGroup Maint.)-----//
'//--------------------------------------------------------------------------------//

Public Function InvalidateDataScopeEmpByRefWG(ByVal intWGNum As Long, ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    
    InvalidateDataScopeByRefWG_Private intWGNum, False, True
'    InvalidateDataScopeEmpByRefWG_Private intWGNum
  
    GetObjectContext().SetComplete
    InvalidateDataScopeEmpByRefWG = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeEmpByRefWG" _
        , , , _
        , DescArgsFrom("WorkGroup Number = " & CStr(intWGNum)) _
        , True
End Function

'//--------------------------------------------------------------------------------//
'//-----expose the following for Invalidation (Security Profile/Data Access)-----//
'//--------------------------------------------------------------------------------//

Public Function InvalidateDataScopeEmpByEffEmp(ByVal intEmpNum As Long, ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    
    InvalidateDataScopeEmpByEffEmp_Private intEmpNum
  
    GetObjectContext().SetComplete
    InvalidateDataScopeEmpByEffEmp = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeEmpByEffEmp" _
        , , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

Public Function InvalidateDataScopeWGByEffEmp(ByVal intEmpNum As Long, ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    
    InvalidateDataScopeWGByEffEmp_Private intEmpNum
    
    If blnCascade Then
        InvalidateDataScopeEmpByEffEmp_Private intEmpNum
    End If
    
    GetObjectContext().SetComplete
    InvalidateDataScopeWGByEffEmp = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeWGByEffEmp" _
        , , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

Public Function InvalidateDataScopeDeptByEffEmp(ByVal intEmpNum As Long, ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    
    InvalidateDataScopeDeptByEffEmp_Private intEmpNum
    
    If blnCascade Then
        InvalidateDataScopeWGByEffEmp_Private intEmpNum
        InvalidateDataScopeEmpByEffEmp_Private intEmpNum
    End If
    
    GetObjectContext().SetComplete
    InvalidateDataScopeDeptByEffEmp = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeDeptByEffEmp" _
        , , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

Public Function InvalidateDataScopeFunByEffEmp(ByVal intEmpNum As Long, ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    
    InvalidateDataScopeFunByEffEmp_Private intEmpNum
    
    If blnCascade Then
        InvalidateDataScopeDeptByEffEmp_Private intEmpNum
        InvalidateDataScopeWGByEffEmp_Private intEmpNum
        InvalidateDataScopeEmpByEffEmp_Private intEmpNum
    End If
    
    GetObjectContext().SetComplete
    InvalidateDataScopeFunByEffEmp = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeFunByEffEmp" _
        , , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

Public Function InvalidateDataScopeDivByEffEmp(ByVal intEmpNum As Long, ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    
    InvalidateDataScopeDivByEffEmp_Private intEmpNum
    
    If blnCascade Then
        InvalidateDataScopeFunByEffEmp_Private intEmpNum
        InvalidateDataScopeDeptByEffEmp_Private intEmpNum
        InvalidateDataScopeWGByEffEmp_Private intEmpNum
        InvalidateDataScopeEmpByEffEmp_Private intEmpNum
    End If
    
    GetObjectContext().SetComplete
    InvalidateDataScopeDivByEffEmp = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeDivByEffEmp" _
        , , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

'//--------------------------------------------------------------------------------//
'//-----the following is for Batch Invalidation -----//
'//--------------------------------------------------------------------------------//
Public Function InvalidateDataScopeBatch(ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    Dim rs As ADODB.Recordset
    Dim lngEmpNum As Long
    Dim lngWGNum As Long
    Dim strDeptNum As String
    Dim strFunNum As String
    Dim strDivNum As String
    Dim strErrorTrace As String
    
    '//-----Employees
    strErrorTrace = "Set rs = GetBatchToInvalidateEmp"
    Set rs = GetBatchToInvalidateEmp
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngEmpNum = rs.Fields("EMP_NUM").Value
        
            strErrorTrace = "InvalidateDataScopeEmpByEffEmp_Private " & CStr(lngEmpNum)
            InvalidateDataScopeEmpByEffEmp_Private lngEmpNum
            
            strErrorTrace = "InvalidateDataScopeEmpByRefEmp_Private " & CStr(lngEmpNum)
            InvalidateDataScopeEmpByRefEmp_Private lngEmpNum
    
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    lngEmpNum = 0
    
    '//-----WorkGroups
    strErrorTrace = "Set rs = GetBatchToInvalidateWG"
    Set rs = GetBatchToInvalidateWG
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngWGNum = rs.Fields("WG_NUM").Value
        
            strErrorTrace = "InvalidateDataScopeByRefWG_Private " & CStr(lngWGNum) & ", True, " & IIf(blnCascade, "True", "False")
            InvalidateDataScopeByRefWG_Private lngWGNum, True, blnCascade

'            strErrorTrace = "InvalidateDataScopeWGByRefWG_Private " & CStr(lngWGNum)
'            InvalidateDataScopeWGByRefWG_Private lngWGNum
'
'            If blnCascade Then
'                '//WG_NUM on TTR_DS_EMP may be 0
'                '//may need to remove this AND CURRENT DATE BETWEEN EWG.EFF_START_DATE AND EWG.EFF_END_DATE
'                strErrorTrace = "InvalidateDataScopeEmpByRefWG_Private " & CStr(lngWGNum)
'                InvalidateDataScopeEmpByRefWG_Private lngWGNum
'            End If
            
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    lngWGNum = 0
    
    '//-----Departments
    strErrorTrace = "Set rs = GetBatchToInvalidateDept"
    Set rs = GetBatchToInvalidateDept
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            strDeptNum = rs.Fields("DEPT_NUM").Value
        
            strErrorTrace = "InvalidateDataScopeByRefDept " & strDeptNum & ", " & blnCascade
            InvalidateDataScopeByRefDept strDeptNum, blnCascade

'            strErrorTrace = "InvalidateDataScopeDeptByRefDept_Private " & strDeptNum
'            InvalidateDataScopeDeptByRefDept_Private strDeptNum
'
'            If blnCascade Then
'                strErrorTrace = "InvalidateDataScopeWGByRefDept_Private " & strDeptNum
'                InvalidateDataScopeWGByRefDept_Private strDeptNum
'
'                strErrorTrace = "InvalidateDataScopeEmpByRefDept_Private " & strDeptNum
'                InvalidateDataScopeEmpByRefDept_Private strDeptNum
'            End If
    
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    strDeptNum = ""
    
    '//-----Departments from Associations
    strErrorTrace = "Set rs = GetBatchToInvalidateDept2"
    Set rs = GetBatchToInvalidateDept2
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            strDeptNum = rs.Fields("DEPT_NUM").Value
        
            strErrorTrace = "InvalidateDataScopeByRefDept " & strDeptNum & ", " & blnCascade
            InvalidateDataScopeByRefDept strDeptNum, blnCascade
        
'            strErrorTrace = "InvalidateDataScopeDeptByRefDept_Private " & strDeptNum
'            InvalidateDataScopeDeptByRefDept_Private strDeptNum
'
'            If blnCascade Then
'                strErrorTrace = "InvalidateDataScopeWGByRefDept_Private " & strDeptNum
'                InvalidateDataScopeWGByRefDept_Private strDeptNum
'
'                strErrorTrace = "InvalidateDataScopeEmpByRefDept_Private " & strDeptNum
'                InvalidateDataScopeEmpByRefDept_Private strDeptNum
'            End If
    
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    strDeptNum = ""
    
    '//-----Functions from Associations
    strErrorTrace = "Set rs = GetBatchToInvalidateFun2"
    Set rs = GetBatchToInvalidateFun2
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            strFunNum = rs.Fields("FUNCTION_NUM").Value
        
            strErrorTrace = "InvalidateDataScopeByRefFun " & strFunNum & ", " & blnCascade
            InvalidateDataScopeByRefFun strFunNum, blnCascade

'            strErrorTrace = "InvalidateDataScopeFunByRefFun_Private " & strFunNum
'            InvalidateDataScopeFunByRefFun_Private strFunNum
'
'            If blnCascade Then
'                strErrorTrace = "InvalidateDataScopeDeptByRefFun_Private " & strFunNum
'                InvalidateDataScopeDeptByRefFun_Private strFunNum
'
'                strErrorTrace = "InvalidateDataScopeWGByRefFun_Private " & strFunNum
'                InvalidateDataScopeWGByRefFun_Private strFunNum
'
'                strErrorTrace = "InvalidateDataScopeEmpByRefFun_Private " & strFunNum
'                InvalidateDataScopeEmpByRefFun_Private strFunNum
'            End If
    
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    strFunNum = ""
    
    '//-----Divisions from Associations
    strErrorTrace = "Set rs = GetBatchToInvalidateDiv2"
    Set rs = GetBatchToInvalidateDiv2
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            strDivNum = rs.Fields("DIVISION_NUM").Value
        
            strErrorTrace = "InvalidateDataScopeByRefDiv " & strDivNum & ", " & blnCascade
            InvalidateDataScopeByRefDiv strDivNum, blnCascade

'            strErrorTrace = "InvalidateDataScopeDivByRefDiv_Private " & strDivNum
'            InvalidateDataScopeDivByRefDiv_Private strDivNum
'
'            If blnCascade Then
'                strErrorTrace = "InvalidateDataScopeFunByRefDiv_Private " & strDivNum
'                InvalidateDataScopeFunByRefDiv_Private strDivNum
'
'                strErrorTrace = "InvalidateDataScopeDeptByRefDiv_Private " & strDivNum
'                InvalidateDataScopeDeptByRefDiv_Private strDivNum
'
'                strErrorTrace = "InvalidateDataScopeWGByRefDiv_Private " & strDivNum
'                InvalidateDataScopeWGByRefDiv_Private strDivNum
'
'                strErrorTrace = "InvalidateDataScopeEmpByRefDiv_Private " & strDivNum
'                InvalidateDataScopeEmpByRefDiv_Private strDivNum
'            End If
    
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    strDivNum = ""
    
    '//-----Employees (Rebuild if not built within last 2 days)
    strErrorTrace = "Set rs = GetBatchToInvalidateEmp3"
    Set rs = GetBatchToInvalidateEmp3

    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngEmpNum = rs.Fields("EFF_EMP_NUM").Value

            strErrorTrace = "InvalidateDataScopeEmpByEffEmp_Private " & CStr(lngEmpNum)
            InvalidateDataScopeEmpByEffEmp_Private lngEmpNum

            rs.MoveNext
        Loop
    End If

    Set rs = Nothing
    lngEmpNum = 0

    '//-----WorkGroups (Rebuild if not built within last 2 days)
    strErrorTrace = "Set rs = GetBatchToInvalidateWG3"
    Set rs = GetBatchToInvalidateWG3

    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngEmpNum = rs.Fields("EFF_EMP_NUM").Value

            strErrorTrace = "InvalidateDataScopeWGByEffEmp_Private " & CStr(lngEmpNum)
            InvalidateDataScopeWGByEffEmp_Private lngEmpNum

            rs.MoveNext
        Loop
    End If

    Set rs = Nothing
    lngEmpNum = 0

    '//-----Departments (Rebuild if not built within last 2 days)
    strErrorTrace = "Set rs = GetBatchToInvalidateDept3"
    Set rs = GetBatchToInvalidateDept3

    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngEmpNum = rs.Fields("EFF_EMP_NUM").Value

            strErrorTrace = "InvalidateDataScopeDeptByEffEmp_Private " & CStr(lngEmpNum)
            InvalidateDataScopeDeptByEffEmp_Private lngEmpNum

            rs.MoveNext
        Loop
    End If

    Set rs = Nothing
    lngEmpNum = 0

    '//-----Functions (Rebuild if not built within last 2 days)
    strErrorTrace = "Set rs = GetBatchToInvalidateFun3"
    Set rs = GetBatchToInvalidateFun3

    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngEmpNum = rs.Fields("EFF_EMP_NUM").Value

            strErrorTrace = "InvalidateDataScopeFunByEffEmp_Private " & CStr(lngEmpNum)
            InvalidateDataScopeFunByEffEmp_Private lngEmpNum

            rs.MoveNext
        Loop
    End If

    Set rs = Nothing
    lngEmpNum = 0

    '//-----Divisions (Rebuild if not built within last 2 days)
    strErrorTrace = "Set rs = GetBatchToInvalidateDiv3"
    Set rs = GetBatchToInvalidateDiv3

    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngEmpNum = rs.Fields("EFF_EMP_NUM").Value

            strErrorTrace = "InvalidateDataScopeDivByEffEmp_Private " & CStr(lngEmpNum)
            InvalidateDataScopeDivByEffEmp_Private lngEmpNum

            rs.MoveNext
        Loop
    End If

    Set rs = Nothing
    lngEmpNum = 0

    strErrorTrace = "BatchSyncAssociationsToAudit"
    BatchSyncAssociationsToAudit
  
    GetObjectContext().SetComplete
    InvalidateDataScopeBatch = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeBatch" _
        , , , _
        , DescArgsFrom(strErrorTrace) _
        , True
End Function

Private Function GetBatchToInvalidateEmp() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchToInvalidateEmp = ExecuteSQL(m_strBATCH_TO_INVALIDATE_EMP_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchToInvalidateEmp", , , , , True
End Function

Private Function GetBatchToInvalidateWG() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchToInvalidateWG = ExecuteSQL(m_strBATCH_TO_INVALIDATE_WG_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchToInvalidateWG", , , , , True
End Function

Private Function GetBatchToInvalidateDept() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchToInvalidateDept = ExecuteSQL(m_strBATCH_TO_INVALIDATE_DEPT_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchToInvalidateDept", , , , , True
End Function

Private Function GetBatchToInvalidateDept2() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchToInvalidateDept2 = ExecuteSQL(m_strBATCH_TO_INVALIDATE_DEPT2_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchToInvalidateDept2", , , , , True
End Function

Private Function GetBatchToInvalidateFun2() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchToInvalidateFun2 = ExecuteSQL(m_strBATCH_TO_INVALIDATE_FUN2_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchToInvalidateFun2", , , , , True
End Function

Private Function GetBatchToInvalidateDiv2() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchToInvalidateDiv2 = ExecuteSQL(m_strBATCH_TO_INVALIDATE_DIV2_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchToInvalidateDiv2", , , , , True
End Function

Private Function GetBatchToInvalidateEmp3() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchToInvalidateEmp3 = ExecuteSQL(m_strBATCH_TO_INVALIDATE_EMP3_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchToInvalidateEmp3", , , , , True
End Function

Private Function GetBatchToInvalidateWG3() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchToInvalidateWG3 = ExecuteSQL(m_strBATCH_TO_INVALIDATE_WG3_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchToInvalidateWG3", , , , , True
End Function

Private Function GetBatchToInvalidateDept3() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchToInvalidateDept3 = ExecuteSQL(m_strBATCH_TO_INVALIDATE_DEPT3_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchToInvalidateDept3", , , , , True
End Function

Private Function GetBatchToInvalidateFun3() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchToInvalidateFun3 = ExecuteSQL(m_strBATCH_TO_INVALIDATE_FUN3_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchToInvalidateFun3", , , , , True
End Function

Private Function GetBatchToInvalidateDiv3() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchToInvalidateDiv3 = ExecuteSQL(m_strBATCH_TO_INVALIDATE_DIV3_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchToInvalidateDiv3", , , , , True
End Function

Private Function BatchSyncAssociationsToAudit() As Long
    
On Error GoTo eh
    Dim oConn As ADODB.Connection
    
    Set oConn = OpenConnection()
    
    BatchSyncAssociationsToAudit = ExecuteOTHERSQL(m_strBATCH_SYNC_ASSOCIATIONS_TO_AUDIT2_sql, oConn)

    Set oConn = Nothing
    Exit Function
eh:
    Set oConn = Nothing
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "BatchSyncAssociationsToAudit", , , , , True
End Function

'//--------------------------------------------------------------------------------//
'//-----the following is for Batch Rebuild -----//
'//--------------------------------------------------------------------------------//
Public Function RebuildDataScopeBatch(ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh
    Dim rs As ADODB.Recordset
    Dim lngEmpNum As Long
    Dim strActiveOnlyInd As String
    Dim strErrorTrace As String
    
    '//-----Employees
    strErrorTrace = "Set rs = GetBatchNeedingRebuiltEmp"
    Set rs = GetBatchNeedingRebuiltEmp
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngEmpNum = rs.Fields("EFF_EMP_NUM").Value
            strActiveOnlyInd = rs.Fields("ACTIVE_ONLY_IND").Value
        
            strErrorTrace = "CheckAndRebuildDataScopeEmp " & CStr(lngEmpNum) & ", " & CBool(strActiveOnlyInd) & ", True"
            CheckAndRebuildDataScopeEmp lngEmpNum, CBool(strActiveOnlyInd), True, True '//full rebuild from batch process
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    lngEmpNum = 0
    
    '//-----WorkGroups
    strErrorTrace = "Set rs = GetBatchNeedingRebuiltWG"
    Set rs = GetBatchNeedingRebuiltWG
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngEmpNum = rs.Fields("EFF_EMP_NUM").Value
            strActiveOnlyInd = rs.Fields("ACTIVE_ONLY_IND").Value
        
            strErrorTrace = "CheckAndRebuildDataScopeWG " & CStr(lngEmpNum) & ", " & CBool(strActiveOnlyInd) & ", True"
            CheckAndRebuildDataScopeWG lngEmpNum, CBool(strActiveOnlyInd), True
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    lngEmpNum = 0
    
    '//-----Departments
    strErrorTrace = "Set rs = GetBatchNeedingRebuiltDept"
    Set rs = GetBatchNeedingRebuiltDept
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngEmpNum = rs.Fields("EFF_EMP_NUM").Value
        
            strErrorTrace = "CheckAndRebuildDataScopeDept " & CStr(lngEmpNum) & ", True"
            CheckAndRebuildDataScopeDept lngEmpNum, True
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    lngEmpNum = 0
    
    '//-----Functions
    strErrorTrace = "Set rs = GetBatchNeedingRebuiltFun"
    Set rs = GetBatchNeedingRebuiltFun
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngEmpNum = rs.Fields("EFF_EMP_NUM").Value
        
            strErrorTrace = "CheckAndRebuildDataScopeFun " & CStr(lngEmpNum) & ", True"
            CheckAndRebuildDataScopeFun lngEmpNum, True
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    lngEmpNum = 0
    
    '//-----Divisions
    strErrorTrace = "Set rs = GetBatchNeedingRebuiltDiv"
    Set rs = GetBatchNeedingRebuiltDiv
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            lngEmpNum = rs.Fields("EFF_EMP_NUM").Value
        
            strErrorTrace = "CheckAndRebuildDataScopeFun " & CStr(lngEmpNum) & ", True"
            CheckAndRebuildDataScopeDiv lngEmpNum, True
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    lngEmpNum = 0
    
    GetObjectContext().SetComplete
    RebuildDataScopeBatch = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "RebuildDataScopeBatch" _
        , , , _
        , DescArgsFrom(strErrorTrace) _
        , True
End Function

Public Function RemoveDataScopeBatch(ByVal blnCascade As Boolean) As Boolean
On Error GoTo eh

    RemoveDataScopeBatch_Private
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "RemoveDataScopeBatch" _
        , , , _
        , _
        , True
End Function

Public Function GetBatchNeedingRebuiltTrace() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchNeedingRebuiltTrace = ExecuteSQL(m_strBATCH_NEEDING_REBUILT_TRACE_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchNeedingRebuiltTrace", , , , , True
End Function

Public Function GetBatchSyncAssocationsToAuditTrace() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchSyncAssocationsToAuditTrace = ExecuteSQLMulti(m_strBATCH_SYNC_ASSOCIATIONS_TO_AUDIT2_TRACE_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchSyncAssocationsToAuditTrace", , , , , True
End Function

Public Function GetBatchRemoveTrace() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchRemoveTrace = ExecuteSQLMulti(m_strBATCH_REMOVE_TRACE_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchRemoveTrace", , , , , True
End Function

Public Function GetBatchInvalidateTrace() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchInvalidateTrace = ExecuteSQL(m_strBATCH_TO_INVALIDATE_TRACE_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchInvalidateTrace", , , , , True
End Function

Private Function RemoveDataScopeBatch_Private() As Long
    
On Error GoTo eh
    Dim oConn As ADODB.Connection
    
    Set oConn = OpenConnection()
    
    RemoveDataScopeBatch_Private = ExecuteOTHERSQL(m_strBATCH_REMOVE_sql, oConn)

    Set oConn = Nothing
    Exit Function
eh:
    Set oConn = Nothing
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "RemoveDataScopeBatch_Private", , , , , True
End Function

Private Function GetBatchNeedingRebuiltEmp() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchNeedingRebuiltEmp = ExecuteSQL(m_strBATCH_NEEDING_REBUILT_EMP_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchNeedingRebuiltEmp", , , , , True
End Function

Private Function GetBatchNeedingRebuiltWG() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchNeedingRebuiltWG = ExecuteSQL(m_strBATCH_NEEDING_REBUILT_WG_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchNeedingRebuiltWG", , , , , True
End Function

Private Function GetBatchNeedingRebuiltDept() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchNeedingRebuiltDept = ExecuteSQL(m_strBATCH_NEEDING_REBUILT_DEPT_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchNeedingRebuiltDept", , , , , True
End Function

Private Function GetBatchNeedingRebuiltFun() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchNeedingRebuiltFun = ExecuteSQL(m_strBATCH_NEEDING_REBUILT_FUN_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchNeedingRebuiltFun", , , , , True
End Function

Private Function GetBatchNeedingRebuiltDiv() As ADODB.Recordset
    
On Error GoTo eh
    Set GetBatchNeedingRebuiltDiv = ExecuteSQL(m_strBATCH_NEEDING_REBUILT_DIV_sql, adOpenStatic, adLockBatchOptimistic)
    
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME, "GetBatchNeedingRebuiltDiv", , , , , True
End Function

'//--------------------------------------------------------------------------------//
'//-----the following are for Invalidation by Effective EmpNum (key)-----//
'//--------------------------------------------------------------------------------//
Private Function InvalidateDataScopeEmpByEffEmp_Private(ByVal intEmpNum As Long) As Boolean
    Dim sSQL As String

On Error GoTo eh
    sSQL = m_strINVALIDATE_Emp_By_Eff_Emp_sql
    
    InvalidateDataScopeInt2 sSQL, intEmpNum
    GetObjectContext().SetComplete
    InvalidateDataScopeEmpByEffEmp_Private = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeEmpByEffEmp_Private" _
        , _
        , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

Private Function InvalidateDataScopeWGByEffEmp_Private(ByVal intEmpNum As Long) As Boolean
    Dim sSQL As String

On Error GoTo eh
    sSQL = m_strINVALIDATE_WG_By_Eff_Emp_sql
    
    InvalidateDataScopeInt2 sSQL, intEmpNum
    GetObjectContext().SetComplete
    InvalidateDataScopeWGByEffEmp_Private = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeWGByEffEmp_Private" _
        , _
        , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

Private Function InvalidateDataScopeDeptByEffEmp_Private(ByVal intEmpNum As Long) As Boolean
    Dim sSQL As String

On Error GoTo eh
    sSQL = m_strINVALIDATE_Dept_By_Eff_Emp_sql
    
    InvalidateDataScopeInt2 sSQL, intEmpNum
    GetObjectContext().SetComplete
    InvalidateDataScopeDeptByEffEmp_Private = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeDeptByEffEmp_Private" _
        , _
        , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

Private Function InvalidateDataScopeFunByEffEmp_Private(ByVal intEmpNum As Long) As Boolean
    Dim sSQL As String

On Error GoTo eh
    sSQL = m_strINVALIDATE_Fun_By_Eff_Emp_sql
    
    InvalidateDataScopeInt2 sSQL, intEmpNum
    GetObjectContext().SetComplete
    InvalidateDataScopeFunByEffEmp_Private = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeFunByEffEmp_Private" _
        , _
        , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

Private Function InvalidateDataScopeDivByEffEmp_Private(ByVal intEmpNum As Long) As Boolean
    Dim sSQL As String

On Error GoTo eh
    sSQL = m_strINVALIDATE_Div_By_Eff_Emp_sql
    
    InvalidateDataScopeInt2 sSQL, intEmpNum
    GetObjectContext().SetComplete
    InvalidateDataScopeDivByEffEmp_Private = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeDivByEffEmp_Private" _
        , _
        , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

'//--------------------------------------------------------------------------------//
'//-----the following are for Invalidation by Referenced fields (TTR_DS_EMP)-----//
'//--------------------------------------------------------------------------------//
Private Function InvalidateDataScopeEmpByRefEmp_Private(ByVal intEmpNum As Long) As Boolean
    Dim sSQL As String

On Error GoTo eh
    sSQL = m_strINVALIDATE_Emp_By_Ref_Emp_sql
    
    InvalidateDataScopeInt2 sSQL, intEmpNum
    GetObjectContext().SetComplete
    InvalidateDataScopeEmpByRefEmp_Private = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeEmpByRefEmp_Private" _
        , _
        , , _
        , DescArgsFrom("Employee Number = " & CStr(intEmpNum)) _
        , True
End Function

'Private Function InvalidateDataScopeEmpByRefWG_Private(ByVal intWGNum As Long) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_Emp_By_Ref_WG_sql
'
'    InvalidateDataScopeInt2 sSQL, intWGNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeEmpByRefWG_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeEmpByRefWG_Private" _
'        , _
'        , , _
'        , DescArgsFrom("WorkGroup Number = " & CStr(intWGNum)) _
'        , True
'End Function

'Private Function InvalidateDataScopeEmpByRefDept_Private(ByVal strDeptNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_Emp_By_Ref_Dept_sql
'
'    InvalidateDataScopeStr2 sSQL, strDeptNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeEmpByRefDept_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeEmpByRefDept_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Department Number = " & strDeptNum) _
'        , True
'End Function

'Private Function InvalidateDataScopeEmpByRefFun_Private(ByVal strFunNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_Emp_By_Ref_Fun_sql
'
'    InvalidateDataScopeStr2 sSQL, strFunNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeEmpByRefFun_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeEmpByRefFun_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Function Number = " & strFunNum) _
'        , True
'End Function

'Private Function InvalidateDataScopeEmpByRefDiv_Private(ByVal strDivNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_Emp_By_Ref_Div_sql
'
'    InvalidateDataScopeStr2 sSQL, strDivNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeEmpByRefDiv_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeEmpByRefDiv_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Division Number = " & strDivNum) _
'        , True
'End Function

'//--------------------------------------------------------------------------------//
'//-----the following are for Invalidation by Referenced fields (TTR_DS_WG)-----//
'//--------------------------------------------------------------------------------//
Private Function InvalidateDataScopeByRefWG_Private(ByVal lngWGNum As Long, _
                                                    ByVal blnUpdateDSWG As Boolean, _
                                                    ByVal blnUpdateDSEmp As Boolean) As Boolean
    Dim sSQL As String

On Error GoTo eh
    sSQL = m_strINVALIDATE_DS_By_Ref_WG_sql
    
    sSQL = Replace(sSQL, "{0}", lngWGNum)
    sSQL = Replace(sSQL, "{1}", IIf(blnUpdateDSWG, "'Y'", "'N'"))
    sSQL = Replace(sSQL, "{2}", IIf(blnUpdateDSEmp, "'Y'", "'N'"))
    
    InvalidateDataScopeInt2 sSQL, lngWGNum
    GetObjectContext().SetComplete
    InvalidateDataScopeByRefWG_Private = True
    Exit Function
eh:
    GetObjectContext().SetAbort

    TrsRaiseError MODULENAME _
        , "InvalidateDataScopeByRefWG_Private" _
        , _
        , , _
        , DescArgsFrom("WorkGroup Number = " & CStr(lngWGNum)) _
        , True
End Function

'Private Function InvalidateDataScopeWGByRefWG_Private(ByVal intWGNum As Long) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_WG_By_Ref_WG_sql
'
'    InvalidateDataScopeInt2 sSQL, intWGNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeWGByRefWG_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeWGByRefWG_Private" _
'        , _
'        , , _
'        , DescArgsFrom("WorkGroup Number = " & CStr(intWGNum)) _
'        , True
'End Function

'Private Function InvalidateDataScopeWGByRefDept_Private(ByVal strDeptNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_WG_By_Ref_Dept_sql
'
'    InvalidateDataScopeStr2 sSQL, strDeptNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeWGByRefDept_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeWGByRefDept_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Department Number = " & strDeptNum) _
'        , True
'End Function

'Private Function InvalidateDataScopeWGByRefFun_Private(ByVal strFunNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_WG_By_Ref_Fun_sql
'
'    InvalidateDataScopeStr2 sSQL, strFunNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeWGByRefFun_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeWGByRefFun_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Function Number = " & strFunNum) _
'        , True
'End Function

'Private Function InvalidateDataScopeWGByRefDiv_Private(ByVal strDivNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_WG_By_Ref_Div_sql
'
'    InvalidateDataScopeStr2 sSQL, strDivNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeWGByRefDiv_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeWGByRefDiv_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Division Number = " & strDivNum) _
'        , True
'End Function

'//--------------------------------------------------------------------------------//
'//-----the following are for Invalidation by Referenced fields (TTR_DS_DEPT)-----//
'//--------------------------------------------------------------------------------//
'Private Function InvalidateDataScopeDeptByRefDept_Private(ByVal strDeptNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_Dept_By_Ref_Dept_sql
'
'    InvalidateDataScopeStr2 sSQL, strDeptNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeDeptByRefDept_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeDeptByRefDept_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Department Number = " & strDeptNum) _
'        , True
'End Function

'Private Function InvalidateDataScopeDeptByRefFun_Private(ByVal strFunNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_Dept_By_Ref_Fun_sql
'
'    InvalidateDataScopeStr2 sSQL, strFunNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeDeptByRefFun_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeDeptByRefFun_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Function Number = " & strFunNum) _
'        , True
'End Function

'Private Function InvalidateDataScopeDeptByRefDiv_Private(ByVal strDivNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_Dept_By_Ref_Div_sql
'
'    InvalidateDataScopeStr2 sSQL, strDivNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeDeptByRefDiv_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeDeptByRefDiv_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Division Number = " & strDivNum) _
'        , True
'End Function

'//--------------------------------------------------------------------------------//
'//-----the following are for Invalidation by Referenced fields (TTR_DS_FUN)-----//
'//--------------------------------------------------------------------------------//
'Private Function InvalidateDataScopeFunByRefFun_Private(ByVal strFunNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_Fun_By_Ref_Fun_sql
'
'    InvalidateDataScopeStr2 sSQL, strFunNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeFunByRefFun_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeFunByRefFun_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Function Number = " & strFunNum) _
'        , True
'End Function

'Private Function InvalidateDataScopeFunByRefDiv_Private(ByVal strDivNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_Fun_By_Ref_Div_sql
'
'    InvalidateDataScopeStr2 sSQL, strDivNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeFunByRefDiv_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeFunByRefDiv_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Division Number = " & strDivNum) _
'        , True
'End Function

'//--------------------------------------------------------------------------------//
'//-----the following are for Invalidation by Referenced fields (TTR_DS_DIV)-----//
'//--------------------------------------------------------------------------------//
'Private Function InvalidateDataScopeDivByRefDiv_Private(ByVal strDivNum As String) As Boolean
'    Dim sSQL As String
'
'On Error GoTo eh
'    sSQL = m_strINVALIDATE_Div_By_Ref_Div_sql
'
'    InvalidateDataScopeStr2 sSQL, strDivNum
'    GetObjectContext().SetComplete
'    InvalidateDataScopeDivByRefDiv_Private = True
'    Exit Function
'eh:
'    GetObjectContext().SetAbort
'
'    TrsRaiseError MODULENAME _
'        , "InvalidateDataScopeDivByRefDiv_Private" _
'        , _
'        , , _
'        , DescArgsFrom("Division Number = " & strDivNum) _
'        , True
'End Function


'Public Function GetEmployees(ByVal vlEmployee As Long _
'                , ByVal vbActiveInformationOnly As Boolean _
'                ) As ADODB.Recordset
'On Error GoTo GetEmployees_EH
'Dim ladrs As ADODB.Recordset
'Dim lsSQL As String
'
''" WHERE A.EMP_NUM = " & CStr(vlEmployee)
'    Set GetEmployees = Nothing
'
'    Set ladrs = CreateParameterRS
'    With ladrs
'        .AddNew
'        !PARM_ORDER = 1
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlEmployee
'        .Update
'
'        .AddNew
'        !PARM_ORDER = 2
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlEmployee
'        .Update
'    End With
'
'    'Call ExecuteSQL and retrieve information
'    'Set GetEmployees = ExecuteSQL(lsSQL _
'                        , adOpenStatic _
'                        , adLockBatchOptimistic)
'    Set GetEmployees = NewExecuteSQL(lsSQL _
'                        , ladrs _
'                        , adOpenStatic _
'                        , adLockBatchOptimistic)
'
'    Set ladrs = Nothing
'
'    'Complete Transaction
'    GetObjectContext.SetComplete
'
'    Exit Function
'GetEmployees_EH:
'    'Abort Transaction
'    GetObjectContext.SetAbort
'
'    Set ladrs = Nothing
'
'    'Raise the appropriate error message depending
'    'on the error number
'    Select Case (Err.Number)
'        Case trErrutOpenConnection
'            TrsRaiseError MODULENAME _
'                , "GetEmployees"
'        Case trErrdb_ExecuteSQL_F
'            TrsRaiseError MODULENAME _
'                , "GetEmployees" _
'                , trErrdb_ExecuteSQL_F _
'                , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'        Case Else
'            TrsRaiseError MODULENAME _
'                , "GetEmployees" _
'                , trErrdb_Employee_F _
'                , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'    End Select
'
'End Function
'Public Function GetAllWorkgroups(ByVal vlEmployee As Long) As ADODB.Recordset
'On Error GoTo GetAllWorkgroups_EH
'Dim ladrs As ADODB.Recordset
'Dim lsSQL As String
'
''" WHERE D.EMP_NUM = " & CStr(vlEmployee)
''" WHERE B.EMP_NUM = " & CStr(vlEmployee)
'
'    Set GetAllWorkgroups = Nothing
'
'    lsSQL = "SELECT  DISTINCT A.WG_NUM, A.CREW_NUM" & _
'        "       , A.WG_DESC, A.DEMAND_IND" & _
'        "       , CASE (D.APPR_LEVEL_IND)" & _
'        "       WHEN 10 THEN 'DIVISION'" & _
'        "       WHEN 20 THEN 'FUNCTION'" & _
'        "       WHEN 30 THEN 'DEPARTMENT'" & _
'        "       WHEN 50 THEN 'WORKGROUP'" & _
'        "       ELSE '' END AS ACCESS_SOURCE, D.APPR_LEVEL_IND" & _
'        "       , CASE (D.ACCESS_LEVEL)" & _
'        "       WHEN 'A' THEN 'APPROVAL'" & _
'        "       WHEN 'E' THEN 'ENTRY'" & _
'        "       WHEN 'V' THEN 'VIEW'" & _
'        "       ELSE '' END AS APP_OR_ENT, D.ACCESS_LEVEL AS APP_ENT_AUTHORITY" & _
'        "       , COALESCE(A.EFF_START_DATE, DATE('0100-01-01')) AS EFF_START_DATE" & _
'        "       , COALESCE(A.EFF_END_DATE, DATE('0100-01-01')) AS EFF_END_DATE" & _
'        "       , A.WG_SUPERVISOR, B.EMP_NAME AS WG_SUPERVISOR_NAME" & _
'        "       , A.PROJECT_ID, C.DEPT_NUM" & _
'        "       , C.LOC_CD, E.PLANT_CD" & _
'        "       , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
'        "       , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
'        "       , C.HR_LOC, C.DIVISION_NUM AS ASSOC_DIV_NUM" & _
'        "       , C.FUNCTION_NUM AS ASSOC_FUNC_NUM, C.DEPT_NUM AS ASSOC_DEPT_NUM" & _
'        "       , CHAR(D.TIME_STAMP) AS APP_TIME_STAMP, D.LAST_MOD_EMP_NUM AS APP_MOD_EMP_NUM"
'    lsSQL = lsSQL & _
'        " FROM  METHDBA.TTR_WORK_GROUP A" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE B" & _
'        "       ON A.WG_SUPERVISOR = B.EMP_NUM" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS C" & _
'        "       ON A.WG_NUM = C.WG_NUM" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_APPROVAL D" & _
'        "       ON ((INTEGER(C.DIVISION_NUM)= D.GROUP_ID" & _
'        "       AND D.APPR_LEVEL_IND = 10)" & _
'        "       OR (INTEGER(C.FUNCTION_NUM) = D.GROUP_ID" & _
'        "       AND D.APPR_LEVEL_IND = 20)" & _
'        "       OR (C.DEPT_NUM = D.DEPT_NUM" & _
'        "       AND D.APPR_LEVEL_IND = 30)" & _
'        "       OR (C.WG_NUM = D.GROUP_ID" & _
'        "       AND D.APPR_LEVEL_IND = 50))" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_LOCATIONS E" & _
'        "       ON E.LOC_CD = C.LOC_CD" & _
'        " WHERE D.EMP_NUM = ?" & _
'        " AND   A.EFF_END_DATE >= CURRENT DATE - 6 MONTHS"
'
'    lsSQL = lsSQL & " UNION "
'
'    lsSQL = lsSQL & _
'        " SELECT  DISTINCT A.WG_NUM, A.CREW_NUM" & _
'        "       , A.WG_DESC, A.DEMAND_IND" & _
'        "       , 'WORKGROUP' AS ACCESS_SOURCE, INTEGER('50') AS APPR_LEVEL_IND" & _
'        "       , CASE (B.PRIMARY_IND)" & _
'        "       WHEN 'Y' THEN 'PRIMARY'" & _
'        "       ELSE 'SECONDARY' END AS APP_OR_ENT" & _
'        "       , CASE (B.PRIMARY_IND)" & _
'        "       WHEN 'Y' THEN 'P'" & _
'        "       ELSE 'S' END AS APP_ENT_AUTHORITY" & _
'        "       , COALESCE(B.EFF_START_DATE, DATE('0100-01-01')) AS EFF_START_DATE" & _
'        "       , COALESCE(B.EFF_END_DATE, DATE('0100-01-01')) AS EFF_END_DATE" & _
'        "       , A.WG_SUPERVISOR, E.EMP_NAME AS WG_SUPERVISOR_NAME" & _
'        "       , A.PROJECT_ID, C.DEPT_NUM" & _
'        "       , C.LOC_CD, D.PLANT_CD" & _
'        "       , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
'        "       , A.LAST_MOD_EMP_NUM, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP" & _
'        "       , C.HR_LOC, C.DIVISION_NUM AS ASSOC_DIV_NUM" & _
'        "       , C.FUNCTION_NUM AS ASSOC_FUNC_NUM, C.DEPT_NUM AS ASSOC_DEPT_NUM" & _
'        "       , CHAR('0') AS APP_TIME_STAMP, 0 AS APP_MOD_EMP_NUM"
'    lsSQL = lsSQL & _
'        " FROM  METHDBA.TTR_WORK_GROUP A" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_EMP_WG B" & _
'        "       ON A.WG_NUM = B.WG_NUM" & _
'        "       AND CURRENT DATE BETWEEN B.EFF_START_DATE" & _
'        "       AND B.EFF_END_DATE" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS C" & _
'        "       ON A.WG_NUM = C.WG_NUM" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_LOCATIONS D" & _
'        "       ON D.LOC_CD = C.LOC_CD" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_EMPLOYEE E" & _
'        "       ON A.WG_SUPERVISOR = E.EMP_NUM" & _
'        " WHERE B.EMP_NUM = ?" & _
'        " ORDER BY WG_NUM;"
'
'    Set ladrs = CreateParameterRS
'    With ladrs
'        .AddNew
'        !PARM_ORDER = 1
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlEmployee
'        .Update
'
'        .AddNew
'        !PARM_ORDER = 2
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlEmployee
'        .Update
'    End With
'
'    'Call ExecuteSQL and retrieve information
'    'Set GetAllWorkgroups = ExecuteSQL(lsSQL _
'                        , adOpenStatic _
'                        , adLockBatchOptimistic)
'    Set GetAllWorkgroups = NewExecuteSQL(lsSQL _
'                        , ladrs _
'                        , adOpenStatic _
'                        , adLockBatchOptimistic)
'
'
'    Set ladrs = Nothing
'
'    'Complete Transaction
'    GetObjectContext.SetComplete
'
'    Exit Function
'GetAllWorkgroups_EH:
'    'Abort Transaction
'    GetObjectContext.SetAbort
'
'    Set ladrs = Nothing
'
'    'Raise the appropriate error message depending
'    'on the error number
'    Select Case (Err.Number)
'        Case trErrutOpenConnection
'            TrsRaiseError MODULENAME _
'                , "GetAllWorkgroups"
'        Case trErrdb_ExecuteSQL_F
'            TrsRaiseError MODULENAME _
'                , "GetAllWorkgroups" _
'                , trErrdb_ExecuteSQL_F _
'                , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'        Case Else
'            TrsRaiseError MODULENAME _
'                , "GetAllWorkgroups" _
'                , trErrdb_WorkGroup_F _
'                , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'    End Select
'
'End Function
'Public Function GetDepartments(ByVal vlEmployee As Long) As ADODB.Recordset
'On Error GoTo GetDepartments_EH
'Dim ladrs As ADODB.Recordset
'Dim lsSQL As String
'
''" WHERE C.EMP_NUM = " & CStr(vlEmployee)
'
'    Set GetDepartments = Nothing
'
'    lsSQL = "SELECT  DISTINCT A.DEPT_NUM, A.DEPT_NAME" & _
'        "       , CASE (C.APPR_LEVEL_IND)" & _
'        "       WHEN 10 THEN 'DIVISION'" & _
'        "       WHEN 20 THEN 'FUNCTION'" & _
'        "       WHEN 30 THEN 'DEPARTMENT'" & _
'        "       ELSE '' END AS ACCESS_SOURCE, C.APPR_LEVEL_IND" & _
'        "       , CASE (C.ACCESS_LEVEL)" & _
'        "       WHEN 'A' THEN 'APPROVAL'" & _
'        "       WHEN 'E' THEN 'ENTRY'" & _
'        "       WHEN 'V' THEN 'VIEW'" & _
'        "       ELSE '' END AS APP_OR_ENT, C.ACCESS_LEVEL AS APP_ENT_AUTHORITY" & _
'        "       , B.DIVISION_NUM AS ASSOC_DIV_NUM, B.FUNCTION_NUM AS ASSOC_FUNC_NUM" & _
'        "       , CHAR(C.TIME_STAMP) AS APP_TIME_STAMP, C.LAST_MOD_EMP_NUM AS APP_MOD_EMP_NUM"
'    lsSQL = lsSQL & _
'        " FROM  METHDBA.TTR_DEPT A" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS B" & _
'        "       ON A.DEPT_NUM = B.DEPT_NUM" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_APPROVAL C" & _
'        "       ON ((INTEGER(B.DIVISION_NUM) = C.GROUP_ID" & _
'        "       AND C.APPR_LEVEL_IND = 10)" & _
'        "       OR (INTEGER(B.FUNCTION_NUM) = C.GROUP_ID" & _
'        "       AND C.APPR_LEVEL_IND = 20)" & _
'        "       OR (B.DEPT_NUM = C.DEPT_NUM" & _
'        "       AND C.APPR_LEVEL_IND = 30))" & _
'        " WHERE C.EMP_NUM = ?" & _
'        " ORDER BY A.DEPT_NUM;"
'
'    Set ladrs = CreateParameterRS
'    With ladrs
'        .AddNew
'        !PARM_ORDER = 1
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlEmployee
'        .Update
'    End With
'
'    'Call ExecuteSQL and retrieve information
'    'Set GetDepartments = ExecuteSQL(lsSQL _
'                        , adOpenStatic _
'                        , adLockBatchOptimistic)
'
'    Set GetDepartments = NewExecuteSQL(lsSQL _
'                        , ladrs _
'                        , adOpenStatic _
'                        , adLockBatchOptimistic)
'    Set ladrs = Nothing
'
'    'Complete Transaction
'    GetObjectContext.SetComplete
'
'    Exit Function
'GetDepartments_EH:
'    'Abort Transaction
'    GetObjectContext.SetAbort
'
'    Set ladrs = Nothing
'
'    'Raise the appropriate error message depending
'    'on the error number
'    Select Case (Err.Number)
'        Case trErrutOpenConnection
'            TrsRaiseError MODULENAME _
'                , "GetDepartments"
'        Case trErrdb_ExecuteSQL_F
'            TrsRaiseError MODULENAME _
'                , "GetDepartments" _
'                , trErrdb_ExecuteSQL_F _
'                , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'        Case Else
'            TrsRaiseError MODULENAME _
'                , "GetDepartments" _
'                , trErrdb_Department_F _
'                , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'    End Select
'
'End Function
'Public Function GetFunctions(ByVal vlEmployee As Long) As ADODB.Recordset
'On Error GoTo GetFunctions_EH
'Dim ladrs As ADODB.Recordset
'Dim lsSQL As String
'
''" WHERE C.EMP_NUM = " & CStr(vlEmployee)
'    Set GetFunctions = Nothing
'
'    lsSQL = "SELECT  DISTINCT A.FUNCTION_NUM, A.FUNCTION_NAME" & _
'        "       , CASE (C.APPR_LEVEL_IND)" & _
'        "       WHEN 10 THEN 'DIVISION'" & _
'        "       WHEN 20 THEN 'FUNCTION'" & _
'        "       ELSE '' END AS ACCESS_SOURCE, C.APPR_LEVEL_IND" & _
'        "       , CASE (C.ACCESS_LEVEL)" & _
'        "       WHEN 'A' THEN 'APPROVAL'" & _
'        "       WHEN 'E' THEN 'ENTRY'" & _
'        "       WHEN 'V' THEN 'VIEW'" & _
'        "       ELSE '' END AS APP_OR_ENT, C.ACCESS_LEVEL AS APP_ENT_AUTHORITY" & _
'        "       , B.DIVISION_NUM AS ASSOC_DIV_NUM" & _
'        "       , CHAR(C.TIME_STAMP) AS APP_TIME_STAMP, C.LAST_MOD_EMP_NUM AS APP_MOD_EMP_NUM"
'    lsSQL = lsSQL & _
'        " FROM  METHDBA.TTR_FUNCTION A" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS B" & _
'        "       ON A.FUNCTION_NUM = B.FUNCTION_NUM" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_APPROVAL C" & _
'        "       ON ((INTEGER(B.DIVISION_NUM) = C.GROUP_ID" & _
'        "       AND C.APPR_LEVEL_IND = 10)" & _
'        "       OR (INTEGER(B.FUNCTION_NUM) = C.GROUP_ID" & _
'        "       AND C.APPR_LEVEL_IND = 20))" & _
'        " WHERE C.EMP_NUM = ?" & _
'        " ORDER BY A.FUNCTION_NUM;"
'
'    Set ladrs = CreateParameterRS
'    With ladrs
'        .AddNew
'        !PARM_ORDER = 1
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlEmployee
'        .Update
'    End With
'
'    'Call ExecuteSQL and retrieve information
'    'Set GetFunctions = ExecuteSQL(lsSQL _
'                        , adOpenStatic _
'                        , adLockBatchOptimistic)
'    Set GetFunctions = NewExecuteSQL(lsSQL _
'                        , ladrs _
'                        , adOpenStatic _
'                        , adLockBatchOptimistic)
'
'    Set ladrs = Nothing
'
'    'Complete Transaction
'    GetObjectContext.SetComplete
'
'    Exit Function
'GetFunctions_EH:
'    'Abort Transaction
'    GetObjectContext.SetAbort
'
'    Set ladrs = Nothing
'
'    'Raise the appropriate error message depending
'    'on the error number
'    Select Case (Err.Number)
'        Case trErrutOpenConnection
'            TrsRaiseError MODULENAME _
'                , "GetFunctions"
'        Case trErrdb_ExecuteSQL_F
'            TrsRaiseError MODULENAME _
'                , "GetFunctions" _
'                , trErrdb_ExecuteSQL_F _
'                , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'        Case Else
'            TrsRaiseError MODULENAME _
'                , "GetFunctions" _
'                , trErrdb_Function_F _
'                , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'    End Select
'
'End Function
'Public Function GetDivisions(ByVal vlEmployee As Long) As ADODB.Recordset
'On Error GoTo GetDivisions_EH
'Dim ladrs As ADODB.Recordset
'Dim lsSQL As String
'
''" WHERE C.EMP_NUM = " & CStr(vlEmployee)
'
'    Set GetDivisions = Nothing
'
'    lsSQL = "SELECT  DISTINCT A.DIVISION_NUM, A.DIVISION_NAME" & _
'        "       , 'DIVISION' AS ACCESS_SOURCE, C.APPR_LEVEL_IND" & _
'        "       , CASE (C.ACCESS_LEVEL)" & _
'        "       WHEN 'A' THEN 'APPROVAL'" & _
'        "       WHEN 'E' THEN 'ENTRY'" & _
'        "       WHEN 'V' THEN 'VIEW'" & _
'        "       ELSE '' END AS APP_OR_ENT, C.ACCESS_LEVEL AS APP_ENT_AUTHORITY" & _
'        "       , CHAR(C.TIME_STAMP) AS APP_TIME_STAMP, C.LAST_MOD_EMP_NUM AS APP_MOD_EMP_NUM"
'    lsSQL = lsSQL & _
'        " FROM  METHDBA.TTR_DIVISION A" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_ASSOCIATIONS B" & _
'        "       ON A.DIVISION_NUM = B.DIVISION_NUM" & _
'        "       LEFT OUTER JOIN METHDBA.TTR_APPROVAL C" & _
'        "       ON INTEGER(B.DIVISION_NUM) = C.GROUP_ID" & _
'        "       AND C.APPR_LEVEL_IND = 10" & _
'        " WHERE C.EMP_NUM = ?" & _
'        " ORDER BY A.DIVISION_NUM;"
'
'    Set ladrs = CreateParameterRS
'    With ladrs
'        .AddNew
'        !PARM_ORDER = 1
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlEmployee
'        .Update
'    End With
'
'    'Call ExecuteSQL and retrieve information
'    'Set GetDivisions = ExecuteSQL(lsSQL _
'                        , adOpenStatic _
'                        , adLockBatchOptimistic)
'    Set GetDivisions = NewExecuteSQL(lsSQL _
'                        , ladrs _
'                        , adOpenStatic _
'                        , adLockBatchOptimistic)
'
'    Set ladrs = Nothing
'
'    'Complete Transaction
'    GetObjectContext.SetComplete
'
'    Exit Function
'GetDivisions_EH:
'    'Abort Transaction
'    GetObjectContext.SetAbort
'
'    Set ladrs = Nothing
'
'    'Raise the appropriate error message depending
'    'on the error number
'    Select Case (Err.Number)
'        Case trErrutOpenConnection
'            TrsRaiseError MODULENAME _
'                , "GetDivisions"
'        Case trErrdb_ExecuteSQL_F
'            TrsRaiseError MODULENAME _
'                , "GetDivisions" _
'                , trErrdb_ExecuteSQL_F _
'                , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'        Case Else
'            TrsRaiseError MODULENAME _
'                , "GetDivisions" _
'                , trErrdb_Division_F _
'                , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'    End Select
'
'End Function


Private Sub Class_Initialize()
    mdbSecurity.LoadStaticVars
    
'    m_strDATASCOPE_EMP_ACTIVE_rebuild_sql = GetStringResource("DATASCOPE_EMP_ACTIVE_rebuild.sql")
'    m_strDATASCOPE_EMP_ALL_rebuild_sql = GetStringResource("DATASCOPE_EMP_ALL_rebuild.sql")
'    m_strDATASCOPE_EMP_check_sql = GetStringResource("DATASCOPE_EMP_check.sql")
'    m_strDATASCOPE_EMP_select_sql = GetStringResource("DATASCOPE_EMP_select.sql")
'    m_strDATASCOPE_WG_rebuild_sql = GetStringResource("DATASCOPE_WG_rebuild.sql")
'    m_strDATASCOPE_WG_check_sql = GetStringResource("DATASCOPE_WG_check.sql")
'    m_strDATASCOPE_WG_select_sql = GetStringResource("DATASCOPE_WG_select.sql")
'    m_strDATASCOPE_DEPT_rebuild_sql = GetStringResource("DATASCOPE_DEPT_rebuild.sql")
'    m_strDATASCOPE_DEPT_check_sql = GetStringResource("DATASCOPE_DEPT_check.sql")
'    m_strDATASCOPE_DEPT_select_sql = GetStringResource("DATASCOPE_DEPT_select.sql")
'    m_strDATASCOPE_FUN_rebuild_sql = GetStringResource("DATASCOPE_FUN_rebuild.sql")
'    m_strDATASCOPE_FUN_check_sql = GetStringResource("DATASCOPE_FUN_check.sql")
'    m_strDATASCOPE_FUN_select_sql = GetStringResource("DATASCOPE_FUN_select.sql")
'    m_strDATASCOPE_DIV_rebuild_sql = GetStringResource("DATASCOPE_DIV_rebuild.sql")
'    m_strDATASCOPE_DIV_check_sql = GetStringResource("DATASCOPE_DIV_check.sql")
'    m_strDATASCOPE_DIV_select_sql = GetStringResource("DATASCOPE_DIV_select.sql")
    
End Sub

'Private Function GetStringResource(ByVal v_strFilename As String) As String
'    Dim oFS As New FileSystemObject
'    Dim ts As TextStream
'    Dim sFilename As String, sFileContents As String
'
'    sFilename = oFS.BuildPath("C:\TRIS SOLUTION\TRIS DLLs NEW\DBScripts", v_strFilename)
'
'    If oFS.FileExists(sFilename) Then
'        Set ts = oFS.OpenTextFile(sFilename, ForReading, False)
'        sFileContents = ts.ReadAll
'    Else
'        sFileContents = ""
'    End If
'
'    Set oFS = Nothing
'    Set ts = Nothing
'    GetStringResource = sFileContents
'End Function

