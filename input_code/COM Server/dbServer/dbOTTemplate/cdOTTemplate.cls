VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdOTTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cdOTTemplate"

Private Enum CANV_OT_SORT_TYPE
    AVAIL_SORT = 1
    GROUP_SORT = 2
    ALL_SORT_GROUP = 3
End Enum


Public Function GetOvertimeTemplates( _
                    v_intLocationCode As Integer _
                ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strOptionalCriteria As String

    On Error GoTo eh

    If v_intLocationCode > 0 Then
        l_strOptionalCriteria = l_strOptionalCriteria & _
            "  AND A.TEMPL_LOCATION = " & CStr(v_intLocationCode)
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT  A.TEMPLATE_NUM, A.TEMPL_LOCATION, B.LOC_DESC "
    l_strSQL = l_strSQL & ", C.EMP_NAME, A.LAST_MOD_EMP_NUM "
    l_strSQL = l_strSQL & ", cast( left(char(A.TEMPL_LOCATION),3) || '-' || "
    l_strSQL = l_strSQL & " A.CANVASS_DESC as char(34)) as canvass_desc  "
    l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS LAST_TIME_STAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_CANV_TEMPL A, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_LOCATIONS B, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_EMPLOYEE C "
    l_strSQL = l_strSQL & " WHERE A.TEMPL_STATUS <> 'I' "
    l_strSQL = l_strSQL & l_strOptionalCriteria
    l_strSQL = l_strSQL & "  AND A.TEMPL_LOCATION = B.LOC_CD "
    l_strSQL = l_strSQL & "  AND A.LAST_MOD_EMP_NUM = C.EMP_NUM "
    l_strSQL = l_strSQL & "ORDER BY 6 "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetOvertimeTemplates = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeTemplates", , , , DescArgsFrom(l_strSQL)
    
End Function
    

Public Function GetOvertimeTemplateBase( _
                    ByVal v_lngTemplateNum As Long _
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

'    l_strSQL = "select TEMPLATE_NUM, "
'    l_strSQL = l_strSQL & _
'                "case when OT_TEMPLATE_IND = 'Y' then 'Overtime' " & _
'                    " when PROJ_TEMPLATE_IND = 'Y' then 'Project' " & _
'                    " when CIPS_TEMPL_IND = 'Y' then 'GENCO' " & _
'                    " when EMRG_CALL_IND = 'C' then 'COE' " & _
'                    " when EMRG_CALL_IND = 'P' then 'POE' " & _
'                    " when EMRG_CALL_IND = 'U' then 'UOE' " & _
'                " else 'UNKNOWN' " & _
'                "end as CANVASS_TYPE , "
'    l_strSQL = l_strSQL & "CANVASS_DESC, FUNCTION_NUM, TEMPL_LOCATION, "
'    l_strSQL = l_strSQL & "HIGH_MAN_CHRG_IND, REPLC_MAN_CHRG_IND, "
'    l_strSQL = l_strSQL & _
'                "case when HIGH_MAN_CHRG_IND = 'N' " & _
'                    " AND REPLC_MAN_CHRG_IND = 'N' " & _
'                    " then 'Y'" & _
'                " else 'N' " & _
'                "end as OPP_HRS_CHRG_IND , "
'    l_strSQL = l_strSQL & "HRS_OVERRIDE_IND, TEMPL_STATUS, LAST_MOD_EMP_NUM, "
'    l_strSQL = l_strSQL & "char(TIME_STAMP) as LAST_MOD_TIMESTAMP, "
'    l_strSQL = l_strSQL & "CANVASS_DESC, FUNCTION_NUM, TEMPL_LOCATION, "
'    l_strSQL = l_strSQL & "RESPONSE_IND, JOB_CRAFT_IND, WORK_LIMITS, "
'    l_strSQL = l_strSQL & "SORT_ORDER, FORCE_ORDER, GROUP_LIST_BY, "
'    l_strSQL = l_strSQL & "LOC_CD_SELECT, JOB_NUM_SELECT, "
'    l_strSQL = l_strSQL & "ACC_OPP_BANK , REF_OPP_BANK "
'    l_strSQL = l_strSQL & "FROM METHDBA.TTR_CANV_TEMPL  "
'    l_strSQL = l_strSQL & "WHERE TEMPLATE_NUM = " & str(v_lngTemplateNum)

    l_strSQL = "SELECT TEMPLATE_NUM, CANVASS_DESC, TEMPL_LOCATION, "
    l_strSQL = l_strSQL & "HRS_OVERRIDE_IND, TEMPL_STATUS, REFUSE_RULE_NUM, "
    l_strSQL = l_strSQL & " LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "char(TIME_STAMP) as TIME_STAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_CANV_TEMPL  "
    l_strSQL = l_strSQL & "WHERE TEMPLATE_NUM = " & str(v_lngTemplateNum)

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetOvertimeTemplateBase = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeTemplateBase", , , , DescArgsFrom(l_strSQL)
    
End Function


Public Function GetTemplateOTResponses( _
                    ByVal v_lngTemplateNum As Long _
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

    l_strSQL = "select TEMPLATE_NUM, A.OT_RESP_CD, RESPONSE_DESC, "
    l_strSQL = l_strSQL & "A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & "CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_CANV_TMPL_RESP A  "
    l_strSQL = l_strSQL & "   , METHDBA.TTR_OT_RESPONSES B  "
    l_strSQL = l_strSQL & "WHERE TEMPLATE_NUM = " & str(v_lngTemplateNum)
    l_strSQL = l_strSQL & "  AND B.OT_RESP_CD = A.OT_RESP_CD "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTemplateOTResponses = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTemplateOTResponses", , , , DescArgsFrom(l_strSQL)
    
End Function
        
        
Public Function GetTemplateIneligibleRules( _
                    ByVal v_lngTemplateNum As Long _
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

    l_strSQL = "select TEMPLATE_NUM, A.INELG_CD, B.INELG_DESC, "
    l_strSQL = l_strSQL & "A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & "CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_CANV_TMPL_INEL A "
    l_strSQL = l_strSQL & "   , METHDBA.TTR_OT_INELIG B "
    l_strSQL = l_strSQL & "WHERE TEMPLATE_NUM = " & str(v_lngTemplateNum)
    l_strSQL = l_strSQL & "  AND B.INELG_CD = A.INELG_CD "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTemplateIneligibleRules = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTemplateIneligibleRules", , , , DescArgsFrom(l_strSQL)
        
End Function


Public Function GetTemplateShifts( _
                    ByVal v_lngTemplateNum As Long _
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

    l_strSQL = "select TEMPLATE_NUM, SHIFT_CD, "
'    l_strSQL = l_strSQL & "case SHIFT_CD "
'    l_strSQL = l_strSQL & "when 'A' then 'AM' "
'    l_strSQL = l_strSQL & "when 'P' then 'PM' "
'    l_strSQL = l_strSQL & "when 'O' then 'OWL' "
'    l_strSQL = l_strSQL & "when 'D' then 'DAY' "
'    l_strSQL = l_strSQL & "when 'N' then 'NIGHT' "
'    l_strSQL = l_strSQL & "end as SHIFT_CD, "
    l_strSQL = l_strSQL & "START_TIME, END_TIME, "
    l_strSQL = l_strSQL & "LUNCH_START_TIME, LUNCH_END_TIME, "
    l_strSQL = l_strSQL & "TURNOVER_START_TM, TURNOVER_END_TM, "
    l_strSQL = l_strSQL & "COMP_SHIFT_CD, PRIOR_DAY_IND, "
    l_strSQL = l_strSQL & "CREATED_EMP_NUM, LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "char(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & "char(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_CANV_TMPL_SHFT  "
    l_strSQL = l_strSQL & "WHERE TEMPLATE_NUM = " & str(v_lngTemplateNum)
    l_strSQL = l_strSQL & " order by SHIFT_CD "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTemplateShifts = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTemplateShifts", , , , DescArgsFrom(l_strSQL)
    
End Function
        

Public Function GetTemplateSort( _
                    ByVal v_lngTemplateNum As Long, _
                    ByVal v_strSortType As String _
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

    l_strSQL = "select TEMPLATE_NUM, A.SORT_CD, SORT_NM, "
    l_strSQL = l_strSQL & "SORT_ORDER, "
    l_strSQL = l_strSQL & "A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & "CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_CANV_TMPL_SORT A "
    l_strSQL = l_strSQL & "   , METHDBA.TTR_CANV_SORT_CD B "
    l_strSQL = l_strSQL & "WHERE TEMPLATE_NUM = " & str(v_lngTemplateNum)
    l_strSQL = l_strSQL & "  AND SORT_TYPE_CD = '" & v_strSortType & "' "
    l_strSQL = l_strSQL & "  AND B.SORT_CD = A.SORT_CD "
    l_strSQL = l_strSQL & "ORDER BY SORT_ORDER "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTemplateSort = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTemplateSort", , , , DescArgsFrom(l_strSQL)
    
End Function


Public Function GetTemplateOrder( _
                    ByVal v_lngTemplateNum As Long, _
                    ByVal v_strSortType As String _
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

    l_strSQL = "SELECT TEMPLATE_NUM, SORT_CD, "
    l_strSQL = l_strSQL & "SORT_ORDER, SORT_TYPE_CD, "
    l_strSQL = l_strSQL & "CREATED_EMP_NUM, LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & "CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_CANV_TMPL_SORT "
    l_strSQL = l_strSQL & "WHERE TEMPLATE_NUM = " & str(v_lngTemplateNum)
    l_strSQL = l_strSQL & " AND SORT_TYPE_CD = '" & v_strSortType & "' "
    l_strSQL = l_strSQL & " ORDER BY SORT_ORDER "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTemplateOrder = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTemplateOrder", , , , DescArgsFrom(l_strSQL)
    
End Function


Public Function GetAvailableSortColumns() As ADODB.Recordset
                
                
    On Error GoTo eh
    
    Set GetAvailableSortColumns = GetSortColumns(AVAIL_SORT)
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetAvailableSortColumns"
                                
End Function


Public Function GetAvailableGroupColumns() As ADODB.Recordset
                
                
    On Error GoTo eh
    
    Set GetAvailableGroupColumns = GetSortColumns(GROUP_SORT)
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetAvailableGroupColumns"
                                
End Function


Public Function GetAllAvailableSortGroupColumns() As ADODB.Recordset
                
    On Error GoTo eh
    
    Set GetAllAvailableSortGroupColumns = GetSortColumns(ALL_SORT_GROUP)
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetAllAvailableSortGroupColumns"
                                
End Function


Public Function SaveOvertimeTemplate( _
                    ByVal v_lngTemplateNum As Long, _
                    ByVal v_strUserID As String, _
                    ByVal v_rsTemplateBase As ADODB.Recordset, _
                    ByVal v_rsSortOrder As ADODB.Recordset, _
                    ByVal v_rsForceOrder As ADODB.Recordset, _
                    ByVal v_rsGroupOrder As ADODB.Recordset, _
                    ByVal v_rsOTResponses As ADODB.Recordset, _
                    ByVal v_rsIneligibleRules As ADODB.Recordset, _
                    ByVal v_rsShifts As ADODB.Recordset _
                ) As Boolean

    Dim cn As ADODB.Connection
    Dim ret As Boolean
    
    Set cn = OpenConnection

    If Not v_rsTemplateBase Is Nothing Then
        ret = SaveTemplateBase(v_lngTemplateNum, v_rsTemplateBase, cn)
        If TemplateDelete(v_rsTemplateBase) Then
            GoTo xit
        End If
    End If

    If Not v_rsSortOrder Is Nothing Then
        ret = SaveSortOrder(v_rsSortOrder, cn)
    End If
    
    If Not v_rsForceOrder Is Nothing Then
        ret = SaveSortOrder(v_rsForceOrder, cn)
    End If
    
    If Not v_rsGroupOrder Is Nothing Then
        ret = SaveSortOrder(v_rsGroupOrder, cn)
    End If
    
    If Not v_rsOTResponses Is Nothing Then
        ret = SaveOTResponses(v_rsOTResponses, cn)
    End If
    
    If Not v_rsIneligibleRules Is Nothing Then
        ret = SaveIneligibleRules(v_rsIneligibleRules, cn)
    End If

    If Not v_rsShifts Is Nothing Then
        ret = SaveShifts(v_rsShifts, cn)
    End If
    
xit:
    cn.Close
    Set cn = Nothing

    SaveOvertimeTemplate = True
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "SaveOvertimeTemplate"
    
End Function


Private Function TemplateDelete( _
                    ByRef r_rsTemplateBase As ADODB.Recordset _
                ) As Boolean
                
On Error GoTo eh
                
    If r_rsTemplateBase.Status = adRecModified Then
        If r_rsTemplateBase!TEMPL_STATUS.Value = "I" And _
            r_rsTemplateBase!TEMPL_STATUS.OriginalValue = " " Then
            TemplateDelete = True
            Exit Function
        End If
    End If
    
    TemplateDelete = False
    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "TemplateDelete"
                
End Function


Private Function SaveSortOrder( _
                    ByRef r_rsSortOrder As ADODB.Recordset, _
                    ByRef cn As ADODB.Connection _
                ) As Boolean

    Dim rc As Boolean
    Dim bError As Boolean
    
On Error GoTo eh

    Dim l_aryPrimaryKey(2) As String
    Dim l_varIgnore(0) As String
    Dim l_varAdditional(0, 1) As String
    Dim l_strSQL As String
    Dim l_strTimestamp As String
    Dim l_lngRecsAffected As Long
            
    l_strSQL = ""
    l_aryPrimaryKey(0) = "TEMPLATE_NUM"
    l_aryPrimaryKey(1) = "SORT_CD"
    l_aryPrimaryKey(2) = "SORT_TYPE_CD"


On Error GoTo eh

    r_rsSortOrder.Filter = adFilterPendingRecords
    Do Until r_rsSortOrder.EOF
        If r_rsSortOrder.Status = adRecNew Then
            l_strSQL = InsertStatement(r_rsSortOrder, "METHDBA.TTR_CANV_TMPL_SORT")
        ElseIf r_rsSortOrder.Status = adRecDeleted Then
            l_strSQL = DeleteStatement(r_rsSortOrder, "METHDBA.TTR_CANV_TMPL_SORT", l_aryPrimaryKey)
        ElseIf r_rsSortOrder.Status = adRecModified Then
            l_strSQL = UpdateStatement(r_rsSortOrder, "METHDBA.TTR_CANV_TMPL_SORT", l_aryPrimaryKey, l_varIgnore, True)
        End If
    
        If l_strSQL <> "" Then
            cn.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        r_rsSortOrder.MoveNext
    Loop

    r_rsSortOrder.Filter = ""
    SaveSortOrder = True
    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "SaveSortOrder"



'    With r_rsSortOrder
'        bError = False
'        If Not (r_rsSortOrder Is Nothing) Then
'            .Filter = adFilterNone
'            .MoveFirst
'            .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
'             Do While Not .EOF
'                rc = DeleteSortOrder(r_rsSortOrder, cn)
'                If Not rc Then bError = True
'                .MoveNext
'            Loop
'
'            .Filter = adFilterNone
'            .MoveFirst
'            .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
'            Do While Not .EOF
'                rc = InsertSortOrder(r_rsSortOrder, v_lTRISUser, cn)
'                If Not rc Then bError = True
'                .MoveNext
'            Loop
'
'        .Filter = adFilterNone
'        End If
'    End With
'
'    SaveSortOrders = Not bError
    
End Function


Private Function SaveOTResponses( _
                    ByRef r_rsOTResponses As ADODB.Recordset, _
                    ByRef cn As ADODB.Connection _
                ) As Boolean

    Dim rc As Boolean
    Dim bError As Boolean
    
On Error GoTo eh

    Dim l_aryPrimaryKey(1) As String
    Dim l_varIgnore(0) As String
    Dim l_varAdditional(0, 1) As String
    Dim l_strSQL As String
    Dim l_strTimestamp As String
    Dim l_lngRecsAffected As Long
            
    l_strSQL = ""
    l_aryPrimaryKey(0) = "TEMPLATE_NUM"
    l_aryPrimaryKey(1) = "OT_RESP_CD"


On Error GoTo eh

    r_rsOTResponses.Filter = adFilterPendingRecords
    Do Until r_rsOTResponses.EOF
        If r_rsOTResponses.Status = adRecNew Then
            l_strSQL = InsertStatement(r_rsOTResponses, "METHDBA.TTR_CANV_TMPL_RESP")
        ElseIf r_rsOTResponses.Status = adRecDeleted Then
            l_strSQL = DeleteStatement(r_rsOTResponses, "METHDBA.TTR_CANV_TMPL_RESP", l_aryPrimaryKey)
'        ElseIf r_rsOTResponses.Status = adRecModified Then
'            l_strSQL = UpdateStatement(r_rsOTResponses, "METHDBA.TTR_CANV_TMPL_RESP", l_aryPrimaryKey, l_varIgnore, True)
        End If
    
        If l_strSQL <> "" Then
            cn.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        r_rsOTResponses.MoveNext
    Loop

    r_rsOTResponses.Filter = ""
    SaveOTResponses = True
    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "SaveOTResponses"
   
End Function


Private Function SaveIneligibleRules( _
                    ByRef r_rsIneligibleRules As ADODB.Recordset, _
                    ByRef cn As ADODB.Connection _
                ) As Boolean

    Dim rc As Boolean
    Dim bError As Boolean
    
On Error GoTo eh

    Dim l_aryPrimaryKey(2) As String
'    Dim l_varIgnore(0) As String
'    Dim l_varAdditional(0, 1) As String
    Dim l_strSQL As String
    Dim l_strTimestamp As String
    Dim l_lngRecsAffected As Long
            
    l_strSQL = ""
    l_aryPrimaryKey(0) = "TEMPLATE_NUM"
    l_aryPrimaryKey(1) = "INELG_CD"
    l_aryPrimaryKey(2) = "LAST_MOD_TIMESTAMP"


On Error GoTo eh

    r_rsIneligibleRules.Filter = adFilterPendingRecords
    Do Until r_rsIneligibleRules.EOF
        If r_rsIneligibleRules.Status = adRecNew Then
            l_strSQL = InsertStatement(r_rsIneligibleRules, "METHDBA.TTR_CANV_TMPL_INEL")
        ElseIf r_rsIneligibleRules.Status = adRecDeleted Then
            l_strSQL = DeleteStatement(r_rsIneligibleRules, "METHDBA.TTR_CANV_TMPL_INEL", l_aryPrimaryKey)
'        ElseIf r_rsIneligibleRules.Status = adRecModified Then
'            l_strSQL = UpdateStatement(r_rsIneligibleRules, "METHDBA.TTR_CANV_TMPL_RESP", l_aryPrimaryKey, l_varIgnore, True)
        End If
    
        If l_strSQL <> "" Then
            cn.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        r_rsIneligibleRules.MoveNext
    Loop

    r_rsIneligibleRules.Filter = ""
    SaveIneligibleRules = True
    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "SaveIneligibleRules"
   
End Function


Private Function SaveShifts( _
                    ByRef r_rsShifts As ADODB.Recordset, _
                    ByRef cn As ADODB.Connection _
                ) As Boolean

    Dim rc As Boolean
    Dim bError As Boolean
    
On Error GoTo eh

    Dim l_aryPrimaryKey(1) As String
    Dim l_varIgnore(0) As String
    Dim l_varAdditional(0, 1) As String
    Dim l_strSQL As String
    Dim l_strTimestamp As String
    Dim l_lngRecsAffected As Long
            
    l_strSQL = ""
    l_aryPrimaryKey(0) = "TEMPLATE_NUM"
    l_aryPrimaryKey(1) = "SHIFT_CD"


On Error GoTo eh

    r_rsShifts.Filter = adFilterPendingRecords
    Do Until r_rsShifts.EOF
        If r_rsShifts.Status = adRecNew Then
            l_strSQL = InsertStatement(r_rsShifts, "METHDBA.TTR_CANV_TMPL_SHFT")
        ElseIf r_rsShifts.Status = adRecDeleted Then
            l_strSQL = DeleteStatement(r_rsShifts, "METHDBA.TTR_CANV_TMPL_SHFT", l_aryPrimaryKey)
        ElseIf r_rsShifts.Status = adRecModified Then
            l_strSQL = UpdateStatement(r_rsShifts, "METHDBA.TTR_CANV_TMPL_SHFT", l_aryPrimaryKey, l_varIgnore, True)
        End If
    
        If l_strSQL <> "" Then
            cn.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        r_rsShifts.MoveNext
    Loop

    r_rsShifts.Filter = ""
    SaveShifts = True
    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "SaveShifts"
   
End Function



'    With r_rsSortOrder
'        bError = False
'        If Not (r_rsSortOrder Is Nothing) Then
'            .Filter = adFilterNone
'            .MoveFirst
'            .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
'             Do While Not .EOF
'                rc = DeleteSortOrder(r_rsSortOrder, cn)
'                If Not rc Then bError = True
'                .MoveNext
'            Loop
'
'            .Filter = adFilterNone
'            .MoveFirst
'            .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
'            Do While Not .EOF
'                rc = InsertSortOrder(r_rsSortOrder, v_lTRISUser, cn)
'                If Not rc Then bError = True
'                .MoveNext
'            Loop
'
'        .Filter = adFilterNone
'        End If
'    End With
'
'    SaveSortOrders = Not bError

Private Function SaveTemplateBase( _
                    ByVal v_lngTemplateNum As Long, _
                    ByRef r_rsTemplateBase As ADODB.Recordset, _
                    ByRef l_objCN As ADODB.Connection _
                ) As Boolean

    Dim strAction As String


    Dim l_aryPrimaryKey(1) As String
    Dim l_varIgnore(0) As String
    Dim l_varAdditional(0, 1) As String
    Dim l_strSQL As String
    Dim l_strTimestamp As String
    Dim l_lngRecsAffected As Long
            
    l_strSQL = ""
    l_aryPrimaryKey(0) = "TEMPLATE_NUM"
    l_aryPrimaryKey(1) = "TIME_STAMP"
    l_varAdditional(0, 0) = "FUNCTION_NUM"
    l_varAdditional(0, 1) = "'001'"     'default?


On Error GoTo eh

    r_rsTemplateBase.MoveFirst
    If r_rsTemplateBase.Status = adRecNew Then
        l_strSQL = InsertStatement(r_rsTemplateBase, "METHDBA.TTR_CANV_TEMPL", , l_varAdditional)
    ElseIf TemplateDelete(r_rsTemplateBase) Then
        'logical delete
        l_strSQL = "UPDATE METHDBA.TTR_CANV_TEMPL" _
            & " SET TEMPL_STATUS = 'I', TIME_STAMP = CURRENT TIMESTAMP," _
            & " LAST_MOD_EMP_NUM = " & r_rsTemplateBase!LAST_MOD_EMP_NUM _
            & " WHERE TEMPLATE_NUM = " & r_rsTemplateBase!TEMPLATE_NUM _
            & " AND TIME_STAMP = " & "'" & r_rsTemplateBase!TIME_STAMP & "'"
        'SQLStatement = DeleteStatement(v_rsData, v_strTableName, v_varPrimaryKeyNames)
    ElseIf r_rsTemplateBase.Status = adRecModified Then
        l_strSQL = UpdateStatement(r_rsTemplateBase, "METHDBA.TTR_CANV_TEMPL", l_aryPrimaryKey, l_varIgnore, True, "TIME_STAMP")
    End If
    
    If l_strSQL <> "" Then
        l_objCN.Execute l_strSQL, l_lngRecsAffected
        If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
             Err.Raise trErrAppCommitConflict
        End If
    End If

    SaveTemplateBase = True
    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "SaveTemplateBase"
    
End Function


Private Function InsertTemplateBase( _
                    ByVal v_lngTemplateNum As Long, _
                    ByRef r_rsTemplateBase As ADODB.Recordset, _
                    ByRef r_adoConnection As ADODB.Connection _
                ) As Boolean

    Dim l_strSQL As String
    
On Error GoTo eh
    
    With r_rsTemplateBase
        l_strSQL = "INSERT INTO METHDBA.TTR_OT_GROUP "
        l_strSQL = l_strSQL & "(TEMPLATE_NUM, CANVASS_DESC "
        l_strSQL = l_strSQL & ", TEMPL_STATUS, TEMPL_LOCATION "
        l_strSQL = l_strSQL & ", HRS_OVERRIDE_IND, RESPONSE_IND "
        'l_strSQL = l_strSQL & ", HRS_REFUSE_RULE_NUM "
        l_strSQL = l_strSQL & ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP) "
        l_strSQL = l_strSQL & " VALUES "
        l_strSQL = l_strSQL & "(" & CStr(v_lngTemplateNum) & ", '" & Trim(!CANVASS_DESC) & "' "
        l_strSQL = l_strSQL & ", " & "' ', " & CStr(!TEMPL_LOCATION)
        l_strSQL = l_strSQL & ", " & "'Y', ' '"
        l_strSQL = l_strSQL & ", " & CStr(!LAST_MOD_EMP_NUM)
        l_strSQL = l_strSQL & ", 'CURRENT TIMESTAMP'"
    End With
    
    If l_strSQL <> "" Then
        r_adoConnection.Execute l_strSQL
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "InsertTemplateBase", , True, , DescArgsFrom(l_strSQL), True

End Function


Private Function UpdateTemplateBase( _
                    ByRef r_rsTemplateBase As ADODB.Recordset, _
                    ByRef r_adoConnection As ADODB.Connection _
                ) As Boolean

    Dim l_strSQL As String
    
On Error GoTo eh
'
'    With r_rsTemplateBase
'        l_strSQL = "INSERT INTO METHDBA.TTR_OT_GROUP "
'        l_strSQL = l_strSQL & "(TEMPLATE_NUM, CANVASS_DESC "
'        l_strSQL = l_strSQL & ", TEMPL_STATUS, TEMPL_LOCATION "
'        l_strSQL = l_strSQL & ", HRS_OVERRIDE_IND, RESPONSE_IND "
'        'l_strSQL = l_strSQL & ", HRS_REFUSE_RULE_NUM "
'        l_strSQL = l_strSQL & ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP) "
'        l_strSQL = l_strSQL & " VALUES "
'        l_strSQL = l_strSQL & "(" & CStr(v_lngTemplateNum) & ", '" & Trim(!CANVASS_DESC) & "' "
'        l_strSQL = l_strSQL & ", " & "' ', " & CStr(!TEMPL_LOCATION)
'        l_strSQL = l_strSQL & ", " & "'Y', ' '"
'        l_strSQL = l_strSQL & ", " & CStr(!LAST_MOD_EMP_NUM)
'        l_strSQL = l_strSQL & ", 'CURRENT TIMESTAMP'"
'    End With
'
'    If l_strSQL <> "" Then
'        r_adoConnection.Execute l_strSQL
'    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "InsertTemplateBase", , True, , DescArgsFrom(l_strSQL), True

End Function



Private Function OpenConnection() As ADODB.Connection
  
    Dim cm As New util_DBConnectMgr.clsuDBConnectMgr
    Dim cn As ADODB.Connection
    
    On Error GoTo eh
    
    Set cn = New ADODB.Connection
    With cn
        .CursorLocation = adUseClient       'Must use client to support RecordCount
        .ConnectionString = cm.GetConnectionString
        .Open
    End With
    
    Set OpenConnection = cn
    Set cm = Nothing
    Set cn = Nothing
    Exit Function
    
eh:
    Set cm = Nothing
    Set cn = Nothing
    TrsRaiseError m_strCMODULENAME, "OpenConnection"
 
End Function


Private Function GetSortColumns( _
                    v_lngSortType As CANV_OT_SORT_TYPE _
                ) As ADODB.Recordset

'
' v_strSortType values (A, G) Available - Group
'

    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strWhere As String

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If v_lngSortType = GROUP_SORT Then
        l_strWhere = "where SORT_CD >= 1000 "
    ElseIf v_lngSortType = AVAIL_SORT Then
        l_strWhere = "where SORT_CD < 1000 "
    Else
        l_strWhere = ""
    End If

    l_strSQL = "select SORT_CD, SORT_NM, SORT_DIRECTION_IND, "
    l_strSQL = l_strSQL & "SORT_COL_NM, SORT_COL_TAB_NM "
    l_strSQL = l_strSQL & "CREATED_EMP_NUM, LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & "CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_CANV_SORT_CD "
    l_strSQL = l_strSQL & l_strWhere & " "
    l_strSQL = l_strSQL & "ORDER BY SORT_CD "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetSortColumns = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetSortColumns", , , , DescArgsFrom(l_strSQL)
    
End Function
    
    
Public Function GetIneligibleRules() As ADODB.Recordset

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

    l_strSQL = "SELECT INELG_CD, INELG_DESC, "
    l_strSQL = l_strSQL & "CREATED_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_INELIG "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetIneligibleRules = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetIneligibleRules", , , , DescArgsFrom(l_strSQL)
        
End Function
    
    
Public Function GetOTResponses() As ADODB.Recordset

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

    l_strSQL = "SELECT OT_RESP_CD, RESPONSE_DESC, OT_WORKED_IND, OT_CHRG_IND, "
    l_strSQL = l_strSQL & "OT_CONTACTED_IND, LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(TIME_STAMP) AS TIME_STAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_RESPONSES "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetOTResponses = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTResponses", , , , DescArgsFrom(l_strSQL)
    
End Function


Public Function GetRefusalChargeRules() As ADODB.Recordset

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

    l_strSQL = "SELECT REFUSE_RULE_NUM, REFUSE_RULE_DESC, "
    l_strSQL = l_strSQL & "CREATED_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA. TTR_OT_REFUSE_RULE "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetRefusalChargeRules = l_objRS
    
    Set l_objRS = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetRefusalChargeRules", , , , DescArgsFrom(l_strSQL)
        
End Function


    
'
'Public Function GetOTOppIneligibilities( _
'    ByVal v_lngOTOppNum As Long _
'    ) As ADODB.Recordset
'
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
'    l_objCN.ConnectionString = l_objCM.GetConnectionString
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open
'    l_strSQL = "select OT_OPP_NUM, A.EMP_NUM, C.EMP_NAME, A.INELG_CD, B.INELG_DESC, "
'    l_strSQL = l_strSQL & "DATE(START_TIMESTAMP) AS START_DATE,   "
'    l_strSQL = l_strSQL & "REPLACE(LEFT(CHAR(TIME(START_TIMESTAMP)),5),'.',':') AS START_TIME, "
'    l_strSQL = l_strSQL & "DATE(END_TIMESTAMP) AS END_DATE,   "
'     l_strSQL = l_strSQL & "REPLACE(LEFT(CHAR(TIME(END_TIMESTAMP)),5),'.',':') AS END_TIME, "
'    l_strSQL = l_strSQL & "WORK_HRS, EXCEPTION_CD, CHARGE_DATE, "
'    l_strSQL = l_strSQL & "A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP  "
'    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_OPP_INELIG A,  "
'    l_strSQL = l_strSQL & "     METHDBA.TTR_OT_INELIG B ,  "
'    l_strSQL = l_strSQL & "     METHDBA.TTR_EMPLOYEE C  "
'    l_strSQL = l_strSQL & "WHERE OT_OPP_NUM = " & str(v_lngOTOppNum)
'    l_strSQL = l_strSQL & "  AND B.INELG_CD = A.INELG_CD "
'    l_strSQL = l_strSQL & "  AND C.EMP_NUM = A.EMP_NUM "
'    l_strSQL = l_strSQL & "ORDER BY A.EMP_NUM, A.INELG_CD"
'
'
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'
'    Set GetOTOppIneligibilities = l_objRS
'
'    Set l_objRS = Nothing
'
'    GetObjectContext.SetComplete
'
'    Exit Function
'
'eh:
'    GetObjectContext.SetAbort
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetOTOppIneligibilities", , , , DescArgsFrom(l_strSQL)
'
'
'End Function
'
'
'Public Function GetOvertimeCanvasses( _
'    ByVal v_strFunctionNum As String, _
'    ByVal v_datStartDate As Date, _
'    ByVal v_datEndDate As Date, _
'    Optional v_intLocationCode As Integer = 0, _
'    Optional v_blnOpenCanvasses As Boolean = True, _
'    Optional v_blnPostedCanvasses As Boolean = False, _
'    Optional v_lngEmployeeNum As Long = 0 _
'    ) As ADODB.Recordset
'
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'    Dim l_strOptionalCriteria As String
'
'    On Error GoTo eh
'
'    l_strOptionalCriteria = ""
'    If v_intLocationCode > 0 Then
''        l_strOptionalCriteria = l_strOptionalCriteria & _
''            "  AND C.TEMPL_LOCATION = " & str(v_intLocationCode)
'        l_strOptionalCriteria = l_strOptionalCriteria & _
'            "  AND D.LOC_CD = " & str(v_intLocationCode)
'    End If
'    If v_blnOpenCanvasses Then
'        If v_blnPostedCanvasses Then
'            l_strOptionalCriteria = l_strOptionalCriteria & _
'             "  AND A.OPP_STATUS IN ('C','A','R','P') "
'        Else
'            l_strOptionalCriteria = l_strOptionalCriteria & _
'             "  AND A.OPP_STATUS IN ('C','A','R') "
'        End If
'    Else
'        If v_blnPostedCanvasses Then
'            l_strOptionalCriteria = l_strOptionalCriteria & _
'                "  AND A.OPP_STATUS = 'P' "
'        Else
'            'All canvasses Open, Posted, and Voided will be shown
'        End If
'    End If
'    If v_lngEmployeeNum > 0 Then
'        l_strOptionalCriteria = l_strOptionalCriteria & _
'            "  AND A.LAST_MOD_EMP_NUM = " & str(v_lngEmployeeNum)
'    End If
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
'    l_strSQL = "SELECT  A.OT_OPP_NUM, A.OT_OPP_DESC, A.CHARGE_DATE "
'    l_strSQL = l_strSQL & ", A.DATE_OF_OT as START_DATE , A.START_TIME "
'    l_strSQL = l_strSQL & ", A.END_DATE, A.END_TIME "
'    l_strSQL = l_strSQL & ", B.OT_REASON_DESC, A.TEMPLATE_NUM, C.CANVASS_DESC "
'    l_strSQL = l_strSQL & ", D.EMP_NAME, A.LAST_MOD_EMP_NUM, A.OPP_STATUS "
'    l_strSQL = l_strSQL & ", C.TEMPL_LOCATION, CHAR(A.TIME_STAMP) AS LAST_MOD_TIMESTAMP "
'    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_OPPORTUNITY A, "
'    l_strSQL = l_strSQL & "     METHDBA.TTR_OT_REASONS B, "
'    l_strSQL = l_strSQL & "     METHDBA.TTR_CANV_TEMPL C, "
'    l_strSQL = l_strSQL & "     METHDBA.TTR_EMPLOYEE D "
'    l_strSQL = l_strSQL & "WHERE A.DATE_OF_OT <= " & DateAsEnclosedString(v_datEndDate)
'    l_strSQL = l_strSQL & "  AND A.DATE_OF_OT >= " & DateAsEnclosedString(v_datStartDate)
'    l_strSQL = l_strSQL & "  AND A.TEMPLATE_NUM = C.TEMPLATE_NUM "
'    l_strSQL = l_strSQL & "  AND C.FUNCTION_NUM = " & SinqleQuotedString(v_strFunctionNum)
'    l_strSQL = l_strSQL & l_strOptionalCriteria
'    l_strSQL = l_strSQL & "  AND A.TEMPLATE_NUM = C.TEMPLATE_NUM "
'    l_strSQL = l_strSQL & "  AND A.OT_REASON_CD = B.OT_REASON_CD  "
'    l_strSQL = l_strSQL & "  AND A.LAST_MOD_EMP_NUM = D.EMP_NUM "
'    l_strSQL = l_strSQL & "ORDER BY A.OT_OPP_NUM "
''    l_strSQL = l_strSQL & "ORDER BY A.DATE_OF_OT, A.START_TIME, A.END_DATE, A.END_TIME "
'
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'
'    Set GetOvertimeCanvasses = l_objRS
'
'    Set l_objRS = Nothing
'
'    GetObjectContext.SetComplete
'
'    Exit Function
'
'eh:
'    GetObjectContext.SetAbort
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetOvertimeCanvasses", , , , DescArgsFrom(l_strSQL)
'
'
'End Function
'

