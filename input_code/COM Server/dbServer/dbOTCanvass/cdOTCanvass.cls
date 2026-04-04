VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdOTCanvass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'NOTES: PREM CREATED NEW DB METHODS FOR CERTAIN RETRIEVALS
'IN OCTOBER 2005
'ConstructCanvassSheet2, GetOvertimeCanvasses2
', GetOTOpportunity2, GetOTOppCanvassSheet2
', GetOTOppPhoneCalls2, GetOpportunityResponses2
', GetNonCanavassedOvertimeCanvasses2, GetCanvasses2
', GetOTOppCharges2, GetOTOppIneligibilities2
'THESE METHODS USE PARAMETERIZED QUERIES

Private Const m_strCMODULENAME As String = "cdOTCanvass"
Private Const m_DBSCRIPTSROOT As String = "C:\TRIS SOLUTION\dbScripts\TRIS_db2"
Private Const m_DBSCRIPTSPATH As String = "\dbOTCanvass\cdOTCanvass"
Private m_ResourceMgr As utilADOHelper.cuResourceMgr

Private Enum CANV_OT_SORT_TYPE
    AVAIL_SORT = 1
    GROUP_SORT = 2
    ALL_SORT = 3
End Enum
Public Function GetMaintainIneligibilities( _
    ByVal v_aryEmployees As Variant, _
    Optional v_datStartDate As Date = CDate("1980-01-01"), _
    Optional v_datEndDate As Date = CDate("9999-12-31") _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo eh

    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If


    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT  A.EMP_NUM, A.EXCEPTION_CD, B.EXCEPTION_DESC "
    l_strSQL = l_strSQL & ", A.EFF_START_DATE, A.EFF_END_DATE "
    l_strSQL = l_strSQL & ", A.EFF_START_TIME, A.EFF_END_TIME "
    l_strSQL = l_strSQL & ", A.LAST_MOD_EMP_NUM, CHAR(A.TIME_STAMP) AS TIME_STAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_INELIGIBILITY A, METHDBA.TTR_EXCEPTION_CDS B "
    l_strSQL = l_strSQL & "WHERE " & l_strEmployeeSelection
    l_strSQL = l_strSQL & "  AND  A.EXCEPTION_CD = B.EXCEPTION_CD "
    l_strSQL = l_strSQL & "  AND  A.EFF_START_DATE <= " & DateAsEnclosedString(v_datEndDate)
    l_strSQL = l_strSQL & "  AND  A.EFF_END_DATE >= " & DateAsEnclosedString(v_datStartDate)
    l_strSQL = l_strSQL & " "
    l_strSQL = l_strSQL & "ORDER BY A.EMP_NUM,A.EFF_START_DATE,A.EFF_START_TIME"



    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetMaintainIneligibilities = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetMaintainIneligibilities", , , , DescArgsFrom(l_strSQL)

End Function
Public Function GetOvertimeCanvasses( _
    ByVal v_intLocationCode As Integer, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_intOTGroup As Integer = 0, _
    Optional v_blnOpenCanvassesOnly As Boolean = False, _
    Optional v_blnPostedCanvassesOnly As Boolean = False _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strStatusCriteria As String
    Dim l_strOTGroupCriteria As String
    Dim l_strOTLocCriteria As String

    On Error GoTo eh

    If v_blnOpenCanvassesOnly Then
         l_strStatusCriteria = " AND A.OPP_STATUS = 'C'"
    Else
        If v_blnPostedCanvassesOnly Then
            l_strStatusCriteria = " AND A.OPP_STATUS = 'P'"
        Else
            l_strStatusCriteria = " AND A.OPP_STATUS <> 'V'"
        End If
    End If
    
    If v_intOTGroup = 0 Then
        l_strOTGroupCriteria = ""
    Else
        l_strOTGroupCriteria = " AND A.OT_GROUP_NUM = " & str(v_intOTGroup)
    End If

    If v_intLocationCode = 0 Then
        l_strOTLocCriteria = ""
    Else
        l_strOTLocCriteria = "  AND A.LOC_CD = " & str(v_intLocationCode)
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT  A.OT_OPP_NUM, " & _
            "CASE A.OPP_STATUS WHEN 'C' then 'OPEN' " & _
                            "  WHEN 'P' then 'POSTED' " & _
                            "  WHEN 'V' then 'VOID' " & _
                            "  ELSE 'ERROR' end as OPP_STATUS "
    l_strSQL = l_strSQL & _
            ",CASE A.SHIFT_CD WHEN 'A' then 'AM' " & _
                            " WHEN 'P' then 'PM' " & _
                            " WHEN 'O' then 'OWL' " & _
                            " WHEN 'D' then 'DAY' " & _
                            " WHEN 'N' then 'NIGHT' " & _
                            " ELSE 'ERROR' end as SHIFT_DESC "
    l_strSQL = l_strSQL & ", A.OT_OPP_DESC, A.OT_GROUP_NUM, E.OT_GROUP_DESC, A.CHARGE_DATE "
    l_strSQL = l_strSQL & ", A.DATE_OF_OT as START_DATE , A.START_TIME "
    l_strSQL = l_strSQL & ", A.END_DATE, A.END_TIME, A.OT_REASON_CD "
    l_strSQL = l_strSQL & ", B.OT_REASON_DESC, A.TEMPLATE_NUM, C.CANVASS_DESC "
    l_strSQL = l_strSQL & ", D.EMP_NAME, A.LAST_MOD_EMP_NUM "
    l_strSQL = l_strSQL & ", C.TEMPL_LOCATION, F.LOC_DESC "
    l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_OPPORTUNITY A, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_OT_REASONS B, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_CANV_TEMPL C, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_EMPLOYEE D, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_OT_GROUP E, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_LOCATIONS F "
    l_strSQL = l_strSQL & "WHERE OT_OPP_NUM >= 300000 AND OT_OPP_NUM < 32000000 "
    l_strSQL = l_strSQL & "  AND A.DATE_OF_OT <= " & DateAsEnclosedString(v_datEndDate)
    l_strSQL = l_strSQL & "  AND A.DATE_OF_OT >= " & DateAsEnclosedString(v_datStartDate)
    l_strSQL = l_strSQL & l_strOTLocCriteria
'    l_strSQL = l_strSQL & "  AND A.LOC_CD = " & str(v_intLocationCode)
    l_strSQL = l_strSQL & l_strStatusCriteria
    l_strSQL = l_strSQL & l_strOTGroupCriteria
    l_strSQL = l_strSQL & "  AND C.TEMPLATE_NUM = A.TEMPLATE_NUM "
    l_strSQL = l_strSQL & "  AND B.OT_REASON_CD = A.OT_REASON_CD  "
    l_strSQL = l_strSQL & "  AND D.EMP_NUM = A.LAST_MOD_EMP_NUM "
    l_strSQL = l_strSQL & "  AND E.OT_GROUP_NUM = A.OT_GROUP_NUM "
    l_strSQL = l_strSQL & "  AND F.LOC_CD = A.LOC_CD "
    l_strSQL = l_strSQL & "ORDER BY A.OT_OPP_NUM "

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetOvertimeCanvasses = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeCanvasses", , , , DescArgsFrom(l_strSQL)
    
    
End Function
 

Public Function GetTemplateOTResponses( _
    ByVal v_lngTemplateNum As Long _
    ) As ADODB.Recordset

'Will eventually be removed.
'The busOTCanvassTrans will call dbOTTemplate


    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "select TEMPLATE_NUM, A.OT_RESP_CD "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_CANV_TMPL_RESP  "
    l_strSQL = l_strSQL & "WHERE TEMPLATE_NUM = " & str(v_lngTemplateNum)

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTemplateOTResponses = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTemplateOTResponses", , , , DescArgsFrom(l_strSQL)
    
End Function
 
Public Function GetOTOpportunity( _
    ByVal v_lngOTOppNum As Long _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "select OT_OPP_NUM, OT_OPP_DESC, A.LOC_CD, TEMPLATE_NUM, SHIFT_CD, "
    l_strSQL = l_strSQL & "DATE_OF_OT, START_TIME, END_TIME, END_DATE, CHARGE_DATE,  "
    l_strSQL = l_strSQL & "START_LUNCH, END_LUNCH, OT_REASON_CD, OPP_STATUS,  "
    l_strSQL = l_strSQL & "A.LAST_MOD_EMP_NUM, A.CREATED_BY_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(A.TIME_STAMP) AS LAST_MOD_TIMESTAMP, CHAR(A.INDEX_TIME) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & "MOD_REASON_DESC, CALL_OUT_CANV_IND, "
    l_strSQL = l_strSQL & "TURNOVER_START_TM, TURNOVER_END_TM, OT_GROUP_NUM, REFUSE_RULE_NUM "
    l_strSQL = l_strSQL & ",B.EMP_NAME AS LAST_MOD_NAME, C.EMP_NAME AS CREATED_BY_NAME "
    l_strSQL = l_strSQL & ",LOC_DESC, CAST(LEFT(CRAFT_JOB,3) AS CHAR(3)) AS COMP_SHIFT_CD "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_OPPORTUNITY A "
    l_strSQL = l_strSQL & "INNER JOIN METHDBA.TTR_EMPLOYEE B"
    l_strSQL = l_strSQL & " ON A.LAST_MOD_EMP_NUM = B.EMP_NUM "
    l_strSQL = l_strSQL & "INNER JOIN METHDBA.TTR_EMPLOYEE C "
    l_strSQL = l_strSQL & " ON A.LAST_MOD_EMP_NUM = C.EMP_NUM "
    l_strSQL = l_strSQL & "INNER JOIN METHDBA.TTR_LOCATIONS D "
    l_strSQL = l_strSQL & " ON A.LOC_CD = D.LOC_CD "
    l_strSQL = l_strSQL & "WHERE OT_OPP_NUM = " & str(v_lngOTOppNum)

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetOTOpportunity = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTOpportunity", , , , DescArgsFrom(l_strSQL)
            

End Function

Public Function GetOTOpportunityForUpdate( _
    ByVal v_lngOTOppNum As Long _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "select OT_OPP_NUM, OT_OPP_DESC, LOC_CD, TEMPLATE_NUM, SHIFT_CD, "
    l_strSQL = l_strSQL & "DATE_OF_OT, START_TIME, END_TIME, END_DATE, CHARGE_DATE,  "
    l_strSQL = l_strSQL & "START_LUNCH, END_LUNCH, OT_REASON_CD, OPP_STATUS,  "
    l_strSQL = l_strSQL & "LAST_MOD_EMP_NUM, CREATED_BY_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(TIME_STAMP) AS TIME_STAMP, CHAR(INDEX_TIME) AS INDEX_TIME, "
    l_strSQL = l_strSQL & "MOD_REASON_DESC, CALL_OUT_CANV_IND, CRAFT_JOB, "
    l_strSQL = l_strSQL & "TURNOVER_START_TM, TURNOVER_END_TM, OT_GROUP_NUM, REFUSE_RULE_NUM "
    l_strSQL = l_strSQL & ", PROJECT_ID "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_OPPORTUNITY  "
    l_strSQL = l_strSQL & "WHERE OT_OPP_NUM = " & str(v_lngOTOppNum)

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetOTOpportunityForUpdate = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTOpportunityForUpdate", , , , DescArgsFrom(l_strSQL)
            

End Function


Public Function GetOTOppResponsesForUpdate( _
    ByVal v_lngOTOppNum As Long _
    ) As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh


    MsgBox "This method obsolete.  Should now update TTR_CANV_SHEETS"
    '
    ' Remove this method after all references cleaned up.
    '
    

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString

    Set l_objCM = Nothing

    l_objCN.Open
    
    l_strSQL = "SELECT OT_OPP_NUM, EMP_NUM, OT_RESP_CD, "
    l_strSQL = l_strSQL & "CALL_BY_EMP_NUM, LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "CALL_NOTES_TXT, "
    l_strSQL = l_strSQL & "CHAR(CALL_ON_TIME_STAMP) AS CALL_ON_TIME_STAMP, "
    l_strSQL = l_strSQL & "CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP, "
    l_strSQL = l_strSQL & "CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_OPP_RESP "
    l_strSQL = l_strSQL & "WHERE OT_OPP_NUM = " & str(v_lngOTOppNum) & " "
    l_strSQL = l_strSQL & "ORDER BY EMP_NUM "


    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetOTOppResponsesForUpdate = l_objrs

    Set l_objrs = Nothing

    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTOppResponsesForUpdate", , , , DescArgsFrom(l_strSQL)

End Function
    
    
Public Function GetOTOppCharges( _
    ByVal v_lngOTOppNum As Long _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strDistrType As String

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strDistrType = " Case " & _
                    " When RTRIM(LTRIM(DISTRIBUTION_ID)) = '' then SPACE(1) " & _
                    " When RTRIM(LTRIM(SUBSTR(DISTRIBUTION_ID,4,1))) = ''  then 'S' " & _
                    " When SUBSTR(DISTRIBUTION_ID,6,2) = 'WR' " & _
                    "  AND RTRIM(LTRIM(SUBSTR(DISTRIBUTION_ID,5,1))) = '' then 'W' " & _
                    " ELSE 'A' END as DISTRIBUTION_TYPE,  "

    l_strSQL = "select OT_OPP_NUM, A.EMP_NUM, B.EMP_NAME,  "
    l_strSQL = l_strSQL & "DATE_WORKED, CHARGE_DATE, START_TIME, END_TIME, HOURS, "
    l_strSQL = l_strSQL & "A.EXCEPTION_CD, OVER_EXC_CD, A.JOB_NUM, COMP_SHIFT_CD, "
    l_strSQL = l_strSQL & "OT_REASON_CD, WG_NUM, DISTRIBUTION_ID, "
    l_strSQL = l_strSQL & l_strDistrType
    l_strSQL = l_strSQL & "EXCEPTION_DESC AS DISTRIBUTION_DESC, "
    l_strSQL = l_strSQL & "Case TRAVEL_PAY_HRS when 0 then 'N' "
    l_strSQL = l_strSQL & "else 'Y' END AS TRAVEL_PAY_IND, "
    l_strSQL = l_strSQL & "A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & "CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_OPP_TM_CHRG A, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_EMPLOYEE B "
    l_strSQL = l_strSQL & "    ,METHDBA.TTR_EXCEPTION_CDS C "
    l_strSQL = l_strSQL & "WHERE OT_OPP_NUM = " & str(v_lngOTOppNum)
    l_strSQL = l_strSQL & "  AND B.EMP_NUM = A.EMP_NUM "
    l_strSQL = l_strSQL & "  AND C.EXCEPTION_CD = A.EXCEPTION_CD "
    l_strSQL = l_strSQL & "ORDER BY A.EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME "


    

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetOTOppCharges = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTOppCharges", , , , DescArgsFrom(l_strSQL)
            

End Function
Public Function GetOTOppChargesForUpdate( _
    ByVal v_lngOTOppNum As Long _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open


    l_strSQL = "select OT_OPP_NUM, EMP_NUM, "
    l_strSQL = l_strSQL & "DATE_WORKED, CHARGE_DATE, START_TIME, END_TIME, HOURS, "
    l_strSQL = l_strSQL & "EXCEPTION_CD, OVER_EXC_CD, JOB_NUM, COMP_SHIFT_CD, "
    l_strSQL = l_strSQL & "OT_REASON_CD, WG_NUM, DISTRIBUTION_ID, "
    l_strSQL = l_strSQL & "DISTRIBUTION_TYPE, JOB_NBR, JOBTASK_NBR, JOBTASK_STEP_NBR, "
    l_strSQL = l_strSQL & "EMPRV_DATABASE_ID, "
    l_strSQL = l_strSQL & "TRAVEL_PAY_HRS, "
    l_strSQL = l_strSQL & "CREATED_EMP_NUM, LAST_MOD_EMP_NUM, "
    l_strSQL = l_strSQL & "CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & "CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_OPP_TM_CHRG "
    l_strSQL = l_strSQL & "WHERE OT_OPP_NUM = " & str(v_lngOTOppNum)
    l_strSQL = l_strSQL & " ORDER BY EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME "


    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetOTOppChargesForUpdate = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTOppChargesForUpdate", , , , DescArgsFrom(l_strSQL)
            

End Function


Public Function GetScheduledOvertimeForUpdate( _
                    ByVal v_lngOTOppNum As Long _
                ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT EMP_NUM, DATE, SCHED_CD, SCHED_TYPE, START_TIME, END_TIME, "
    l_strSQL = l_strSQL & "EXCEPTION_CD, WG_NUM, JOB_NUM, "
    l_strSQL = l_strSQL & "RELIEF_COMP, OT_OPP_NUM, CASE_NUM, LOC_CD, "
    l_strSQL = l_strSQL & "COMP_CD, SHIFT_COMP, OVER_EXC_CD, FMLA_IND, "
    l_strSQL = l_strSQL & "TRAV_PAY_HRS, CHARGE_DATE, SUBSYS_OWNER, "
    l_strSQL = l_strSQL & "LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP, "
    l_strSQL = l_strSQL & "PENDING_STATUS, LINE_NUM, REQ_IND, PREV_SCHED_CD "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_BASE_HRS_EXC "
    l_strSQL = l_strSQL & "WHERE SCHED_TYPE = 120 "
    l_strSQL = l_strSQL & "AND OT_OPP_NUM = " & CStr(v_lngOTOppNum)
    l_strSQL = l_strSQL & " ORDER BY EMP_NUM, CHARGE_DATE, DATE, START_TIME "

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetScheduledOvertimeForUpdate = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduledOvertimeForUpdate", , , , DescArgsFrom(l_strSQL)
            
End Function


Public Function GetOTOppCanvassSheet( _
    ByVal v_lngOTOppNum As Long _
    ) As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString

    Set l_objCM = Nothing

    l_objCN.Open
    l_strSQL = "select A.OT_OPP_NUM, A.EMP_NUM, B.EMP_NAME, A.LOC_CD, "
    l_strSQL = l_strSQL & " A.QUAL_SKILL_IND, A.INELIG_IND, A.HOME_PHONE_NUM,"
    l_strSQL = l_strSQL & " A.OT_STANDING_HRS, A.JOB_NUM, A.CRAFT_NUM, "
    l_strSQL = l_strSQL & " A.ROW_NUM, A.FORCE_ROW_NUM, "
    l_strSQL = l_strSQL & " LEFT(DIGITS(PLANT_SNRTY_QTY),3) as PLANT_SNRTY_MONTHS, "
    l_strSQL = l_strSQL & " RIGHT(DIGITS(PLANT_SNRTY_QTY),2) as PLANT_SNRTY_DAYS, "
    l_strSQL = l_strSQL & " LEFT(DIGITS(UNION_SNRTY_QTY),3) as UNION_SNRTY_MONTHS, "
    l_strSQL = l_strSQL & " RIGHT(DIGITS(UNION_SNRTY_QTY),2) as UNION_SNRTY_DAYS, "
    l_strSQL = l_strSQL & " A.CANVASS_RESP_CD, FORCE_RESP_CD,"
    l_strSQL = l_strSQL & " A.CALL_NOTES_TXT, "
    l_strSQL = l_strSQL & " CHAR(A.CANVASS_CALL_TIME) AS CANVASS_CALL_TIME, "
    l_strSQL = l_strSQL & " CHAR(A.FORCE_CALL_TIME) AS FORCE_CALL_TIME, "
    l_strSQL = l_strSQL & "  SPACE(20) AS QUAL_CDS, "
    l_strSQL = l_strSQL & "  SPACE(20) AS SPEC_SKILL_CDS "
    l_strSQL = l_strSQL & " ,SPACE(1) AS ON_DUTY_IND "
    l_strSQL = l_strSQL & " ,SPACE(3) AS ON_DUTY_EXC_CD "
    l_strSQL = l_strSQL & " ,A.CREATED_EMP_NUM "
    l_strSQL = l_strSQL & " ,A.LAST_MOD_EMP_NUM "
    l_strSQL = l_strSQL & " ,CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP "
    l_strSQL = l_strSQL & " ,CHAR(A.TIME_STAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_CANV_SHEETS A "
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_EMPLOYEE B "
    l_strSQL = l_strSQL & " ON A.EMP_NUM = B.EMP_NUM "
    l_strSQL = l_strSQL & " WHERE A.OT_OPP_NUM = " & str(v_lngOTOppNum)
    l_strSQL = l_strSQL & " ORDER BY A.ROW_NUM "
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetOTOppCanvassSheet = l_objrs

    Set l_objrs = Nothing

    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTOppCanvassSheet", , , , DescArgsFrom(l_strSQL)


End Function
      
Public Function GetOTOppPhoneCalls( _
    ByVal v_lngOTOppNum As Long _
    ) As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString

    Set l_objCM = Nothing

    l_objCN.Open
    l_strSQL = "SELECT OT_OPP_NUM, A.EMP_NUM, B.EMP_NAME, CALL_NUM, " & _
                    "CHAR(DATE(CALL_ST_TIMESTAMP)) || ' ' || REPLACE(CHAR(TIME(CALL_ST_TIMESTAMP)),'.',':') as CALL_START, " & _
                    "CHAR(DATE(CALL_END_TIMESTAMP)) || ' ' || REPLACE(CHAR(TIME(CALL_END_TIMESTAMP)),'.',':') as CALL_END, " & _
                    "DIALED_NUM, CREATED_EMP_NUM, " & _
                    "C.EMP_NAME AS CALLER_NAME, " & _
                    "CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP "
    l_strSQL = l_strSQL & _
                "FROM METHDBA.TTR_OT_OPP_CALLOUT A " & _
                "INNER JOIN  METHDBA.TTR_EMPLOYEE B " & _
                "ON A.EMP_NUM  = B.EMP_NUM " & _
                "INNER JOIN  METHDBA.TTR_EMPLOYEE C " & _
                "ON A.CREATED_EMP_NUM  = C.EMP_NUM " & _
                "WHERE OT_OPP_NUM = " & str(v_lngOTOppNum) & " " & _
                "ORDER BY A.CALL_ST_TIMESTAMP, A.EMP_NUM"

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetOTOppPhoneCalls = l_objrs

    Set l_objrs = Nothing

    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTOppBankCharges", , , , DescArgsFrom(l_strSQL)


End Function

    
Public Function GetEmployeesInOTGroup( _
    ByVal v_datChargeDate As Date, _
    ByVal v_intOTGroup As Integer _
    ) As ADODB.Recordset

'Might Need to Remove this and use dbOTGroup
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strWhere As String


    On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT DISTINCT B.EMP_NUM "
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_WORK_GROUP A "
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_EMP_WG B "
    l_strSQL = l_strSQL & " ON A.WG_NUM = B.WG_NUM "
    l_strSQL = l_strSQL & " AND B.EFF_START_DATE <= " & SinqleQuotedString(Format(v_datChargeDate, "YYYY-MM-DD"))
    l_strSQL = l_strSQL & " AND B.EFF_END_DATE >= " & SinqleQuotedString(Format(v_datChargeDate, "YYYY-MM-DD"))
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_EMPLOYEE C "
    l_strSQL = l_strSQL & " ON B.EMP_NUM = C.EMP_NUM "
    l_strSQL = l_strSQL & " AND EMP_STATUS <> 'T' "
    l_strSQL = l_strSQL & " LEFT OUTER JOIN METHDBA.TTR_EMP_WG E "
    l_strSQL = l_strSQL & " ON E.EMP_NUM = B.EMP_NUM "
    l_strSQL = l_strSQL & " AND  E.WG_NUM IN (99995,99996,99997,99998,99999) "
    l_strSQL = l_strSQL & " AND E.EFF_START_DATE <= " & SinqleQuotedString(Format(v_datChargeDate, "YYYY-MM-DD"))
    l_strSQL = l_strSQL & " AND E.EFF_END_DATE >= " & SinqleQuotedString(Format(v_datChargeDate, "YYYY-MM-DD"))
    l_strSQL = l_strSQL & " WHERE  A.WG_NUM IN ( "
    l_strSQL = l_strSQL & " SELECT WG_NUM "
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_OT_GROUP_WG "
    l_strSQL = l_strSQL & " WHERE OT_GROUP_NUM = " & str(v_intOTGroup) & " ) "
    l_strSQL = l_strSQL & " AND A.EFF_START_DATE <= " & SinqleQuotedString(Format(v_datChargeDate, "YYYY-MM-DD"))
    l_strSQL = l_strSQL & " AND A.EFF_END_DATE >= " & SinqleQuotedString(Format(v_datChargeDate, "YYYY-MM-DD"))
    l_strSQL = l_strSQL & " AND E.EMP_NUM IS NULL "
    l_strSQL = l_strSQL & " ORDER BY B.EMP_NUM "

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEmployeesInOTGroup = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeesInOTGroup", , , , DescArgsFrom(l_strSQL)
       

End Function

Public Function ConstructCanvassSheet( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datChargeDate As Date, _
    ByVal v_intOTBank As Integer _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String
    Dim l_strBankSel As String

    On Error GoTo eh

    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If
    l_strBankSel = "OT_BANK_CD = " & str(v_intOTBank) & " "


    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT  Y.EMP_NUM, Y.HOME_PHONE_NUM, "
    l_strSQL = l_strSQL & "INTEGER(COALESCE(DEPT_SNRTY_QTY,0))  AS DEPT_SNRTY_QTY, "
    l_strSQL = l_strSQL & "INTEGER(COALESCE(PLANT_SNRTY_QTY,0)) AS PLANT_SNRTY_QTY, "
    l_strSQL = l_strSQL & "INTEGER(COALESCE(UNION_SNRTY_QTY,0)) AS UNION_SNRTY_QTY, "
    l_strSQL = l_strSQL & "INTEGER(COALESCE(CORP_SNRTY_QTY,0))  AS CORP_SNRTY_QTY, "
    l_strSQL = l_strSQL & "INTEGER(COALESCE(LOP_SNRTY_QTY,0))   AS LOP_SNRTY_QTY, "
    l_strSQL = l_strSQL & "INTEGER(COALESCE(JOB_SNRTY_QTY,0))   AS JOB_SNRTY_QTY, "
    l_strSQL = l_strSQL & "INTEGER(COALESCE(CMPY_SNRTY_QTY,0))  AS CMPY_SNRTY_QTY, "
    l_strSQL = l_strSQL & "Y.JOB_NUM, Y.DATE_OF_BIRTH, X.CRAFT_NUM, W.LOC_CD, "
    l_strSQL = l_strSQL & "DECIMAL(COALESCE(TOT_OT_HRS,99999.90),7,2)  AS OT_STANDING_HRS, "
    l_strSQL = l_strSQL & "SMALLINT (0) AS ROW_NUM, SMALLINT (0) AS FORCE_ROW_NUM, "
    l_strSQL = l_strSQL & "SMALLINT (0) AS LOC_SORT, SMALLINT (0) AS WG_SORT, Y.DEPT_NUM, Y.EMP_NAME  "
    l_strSQL = l_strSQL & "FROM  "
    l_strSQL = l_strSQL & "(SELECT EMP_NUM, DEPT_NUM, UNION_CD, JOB_NUM, DATE_OF_BIRTH, HOME_PHONE_NUM, EMP_NAME "
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EMPLOYEE "
    l_strSQL = l_strSQL & " WHERE " & l_strEmployeeSelection & " )Y  "
    l_strSQL = l_strSQL & "INNER JOIN METHDBA.TTR_JOBS X ON Y.JOB_NUM = X.JOB_NUM "
    l_strSQL = l_strSQL & "INNER JOIN "
    l_strSQL = l_strSQL & "(SELECT A.EMP_NUM, A.WG_NUM,B.LOC_CD"
    l_strSQL = l_strSQL & " FROM  METHDBA.TTR_EMP_WG A, METHDBA.TTR_ASSOCIATIONS B "
    l_strSQL = l_strSQL & " WHERE " & l_strEmployeeSelection
    l_strSQL = l_strSQL & "      AND " & SinqleQuotedString(Format(v_datChargeDate, "YYYY-MM-DD"))
    l_strSQL = l_strSQL & "       BETWEEN EFF_START_DATE AND EFF_END_DATE  "
    l_strSQL = l_strSQL & "   AND A.PRIMARY_IND  = 'Y' "
    l_strSQL = l_strSQL & "   AND B.WG_NUM = A.WG_NUM ) W "
    l_strSQL = l_strSQL & "ON W.EMP_NUM = Y.EMP_NUM "
    l_strSQL = l_strSQL & "LEFT OUTER JOIN "
    l_strSQL = l_strSQL & "(SELECT EMP_NUM, TOT_OT_HRS "
    l_strSQL = l_strSQL & " FROM  METHDBA.TTR_OT_BANKS  "
    l_strSQL = l_strSQL & " WHERE " & l_strBankSel & " ) V "
    l_strSQL = l_strSQL & "ON V.EMP_NUM = Y.EMP_NUM "
    l_strSQL = l_strSQL & "LEFT OUTER JOIN "
    l_strSQL = l_strSQL & "(SELECT EMP_NUM, DEPT_NUM, UNION_CD, JOB_NUM, "
    l_strSQL = l_strSQL & "   DEPT_SNRTY_QTY, PLANT_SNRTY_QTY, UNION_SNRTY_QTY, "
    l_strSQL = l_strSQL & "   CORP_SNRTY_QTY, LOP_SNRTY_QTY, JOB_SNRTY_QTY, CMPY_SNRTY_QTY "
    l_strSQL = l_strSQL & " FROM  METHDBA.TTR_SENIORITY  ) Z "
    l_strSQL = l_strSQL & "ON Z.EMP_NUM = Y.EMP_NUM "
    l_strSQL = l_strSQL & " AND Z.DEPT_NUM = Y.DEPT_NUM "
    l_strSQL = l_strSQL & " AND Z.UNION_CD = Y.UNION_CD "
    l_strSQL = l_strSQL & " AND Z.JOB_NUM = Y.JOB_NUM "
    l_strSQL = l_strSQL & "ORDER BY Y.EMP_NUM "




    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set ConstructCanvassSheet = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "ConstructCanvassSheet", , , , DescArgsFrom(l_strSQL)

End Function


Public Function GetWorkGroupSortData( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_intOTGroup As Integer, _
    ByVal v_datChargeDate As Date _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

    On Error GoTo eh

    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "A.EMP_NUM = " & str(CLng(v_aryEmployees)) & " "
    End If


    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT  A.EMP_NUM, B.WG_NUM, B.ORDER_OF_CALL "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_EMP_WG A,"
    l_strSQL = l_strSQL & "     METHDBA.TTR_OT_GROUP_WG B "
    l_strSQL = l_strSQL & "WHERE OT_GROUP_NUM = " & str(v_intOTGroup)
    l_strSQL = l_strSQL & "  AND " & l_strEmployeeSelection
    l_strSQL = l_strSQL & "  AND " & SinqleQuotedString(Format(v_datChargeDate, "YYYY-MM-DD"))
    l_strSQL = l_strSQL & "   BETWEEN EFF_START_DATE AND EFF_END_DATE "
    l_strSQL = l_strSQL & "  AND B.WG_NUM = A.WG_NUM "
    l_strSQL = l_strSQL & "ORDER BY A.EMP_NUM, B.ORDER_OF_CALL"



    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetWorkGroupSortData = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupSortData", , , , DescArgsFrom(l_strSQL)
  
    
End Function

Public Function GetOTOppIneligibilities( _
    ByVal v_lngOTOppNum As Long _
    ) As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString

    Set l_objCM = Nothing

    l_objCN.Open
    l_strSQL = "select OT_OPP_NUM, A.EMP_NUM, C.EMP_NAME, A.INELG_CD, B.INELG_DESC, "
    l_strSQL = l_strSQL & "DATE(START_TIMESTAMP) AS START_DATE,   "
    l_strSQL = l_strSQL & "REPLACE(LEFT(CHAR(TIME(START_TIMESTAMP)),5),'.',':') AS START_TIME, "
    l_strSQL = l_strSQL & "DATE(END_TIMESTAMP) AS END_DATE,   "
     l_strSQL = l_strSQL & "REPLACE(LEFT(CHAR(TIME(END_TIMESTAMP)),5),'.',':') AS END_TIME, "
    l_strSQL = l_strSQL & "WORK_HRS, EXCEPTION_CD, CHARGE_DATE, "
    l_strSQL = l_strSQL & "A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP  "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_OPP_INELIG A,  "
    l_strSQL = l_strSQL & "     METHDBA.TTR_OT_INELIG B ,  "
    l_strSQL = l_strSQL & "     METHDBA.TTR_EMPLOYEE C  "
    l_strSQL = l_strSQL & "WHERE OT_OPP_NUM = " & str(v_lngOTOppNum)
    l_strSQL = l_strSQL & "  AND B.INELG_CD = A.INELG_CD "
    l_strSQL = l_strSQL & "  AND C.EMP_NUM = A.EMP_NUM "
    l_strSQL = l_strSQL & "ORDER BY A.EMP_NUM, A.INELG_CD"


    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetOTOppIneligibilities = l_objrs

    Set l_objrs = Nothing

    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTOppIneligibilities", , , , DescArgsFrom(l_strSQL)


End Function


Public Function SaveOpportunity( _
            ByVal v_rsOpportunity As ADODB.Recordset _
        ) As Boolean
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long
    
On Error GoTo eh

'    v_rsOpportunity.Filter = adFilterPendingRecords
'    If v_rsOpportunity.EOF Then
'        v_rsOpportunity.Filter = adFilterNone
'        SaveOpportunity = True
'        GetObjectContext.SetComplete
'        Exit Function
'    End If
    
'    If recordsexist(v_rsOpportunity) Then
'        v_rsOpportunity.MoveFirst
'    End If
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    Set l_objCM = Nothing
    
    Do Until v_rsOpportunity.EOF

        l_strSQL = SaveOpportunitySQL(v_rsOpportunity)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsOpportunity.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    v_rsOpportunity.Filter = adFilterNone
    SaveOpportunity = True
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort
'    TrsRaiseError m_strCMODULENAME, "SaveOpportunity", , , , DescArgsFrom(l_strSQL), True
    Select Case Err.Number
        Case trErrAppCommitConflict
            TrsTraceError m_strCMODULENAME, "SaveOpportunity", LongAppErr(trErrAppCommitConflict), , , DescArgsFrom(l_strSQL), True
            TrsRaiseError m_strCMODULENAME, "SaveOpportunity", trErrAppCommitConflict
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveOpportunity", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function


Public Function GetEmployeeScheduledHolidays( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date _
                ) As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strEmployeeSelection As String

On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMP_NUM", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMP_NUM = " & CStr(CLng(v_aryEmployees)) & " "
    End If

    l_strSQL = "SELECT EMP_NUM, CHARGE_DATE, HOURS, HOURS_OFFSET_NUM,"
    l_strSQL = l_strSQL & " A.EXCEPTION_CD, WORK_REST_IND"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EMP_SCHED_EXC A, METHDBA.TTR_EXCEPTION_CDS B"
    l_strSQL = l_strSQL & " WHERE " & l_strEmployeeSelection
    l_strSQL = l_strSQL & " (CHARGE_DATE BETWEEN" & DateAsEnclosedString(v_datStartDate)
    l_strSQL = l_strSQL & " AND " & DateAsEnclosedString(v_datEndDate) & ")"
    l_strSQL = l_strSQL & " AND A.EXCEPTION_CD = B.EXCEPTION_CD"
    l_strSQL = l_strSQL & " AND GROUP_IND = 'H'"
    l_strSQL = l_strSQL & " ORDER BY EMP_NUM, CHARGE_DATE "

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetEmployeeScheduledHolidays = l_objrs
    Set l_objrs = Nothing
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeScheduledHolidays", , , , DescArgsFrom(l_strSQL)

End Function


Public Function GetHolidayExceptions() As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr


On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT EXCEPTION_CD, EXCEPTION_DESC, EXCEPTION_TYPE, "
    l_strSQL = l_strSQL & " CONTRACT_EARN_CD, EXEC_EARN_CD"
    l_strSQL = l_strSQL & " SYMBOL_TIME_IND, EXCEPTION_BANK_CD"
    l_strSQL = l_strSQL & " INELIGIBILITY_IND, OTHERPAY_IND"
    l_strSQL = l_strSQL & " SCHED_IND, WORK_REST_IND, SYMBOL_CD, GROUP_IND"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EXCEPTION_CDS"
    l_strSQL = l_strSQL & " WHERE GROUP_IND = 'H'"
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetHolidayExceptions = l_objrs
    Set l_objrs = Nothing
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetHolidayExceptions", , , , DescArgsFrom(l_strSQL)

End Function


Public Function GetAllExceptions() As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr


On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT EXCEPTION_CD, EXCEPTION_DESC, EXCEPTION_TYPE, "
    l_strSQL = l_strSQL & " CONTRACT_EARN_CD, EXEC_EARN_CD"
    l_strSQL = l_strSQL & " SYMBOL_TIME_IND, EXCEPTION_BANK_CD"
    l_strSQL = l_strSQL & " INELIGIBILITY_IND, OTHERPAY_IND"
    l_strSQL = l_strSQL & " SCHED_IND, WORK_REST_IND, SYMBOL_CD, GROUP_IND"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EXCEPTION_CDS"
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    Set GetAllExceptions = l_objrs
    Set l_objrs = Nothing
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllExceptions", , , , DescArgsFrom(l_strSQL)

End Function

'
'Private Function SaveOpportunitySQL( _
'            ByVal v_rsOpportunity As ADODB.Recordset _
'        ) As String
'
'    Dim l_strSQL As String
'    Const CTABLE As String = "METHDBA.TTR_OT_OPPORTUNITY"
'    Dim primKeys As Variant
'    Dim l_strTimeStamp As String
'    Dim l_varIgnore As Variant
'
'    On Error GoTo eh
'    primKeys = Array("OT_OPP_NUM")
'    l_varIgnore = Array("TIME_STAMP", "INDEX_TIME")
'    Select Case v_rsOpportunity.Status
'       Case adRecModified
'            l_strSQL = UpdateStatement(v_rsOpportunity, CTABLE, primKeys, , True, "TIME_STAMP")
'       Case adRecNew
'           l_strSQL = InsertStatement(v_rsOpportunity, CTABLE)
'       Case adRecDeleted
'            l_strSQL = DeleteStatement(v_rsOpportunity, CTABLE, primKeys)
'        Case Else
'           l_strSQL = ""
'    End Select
'
'    SaveOpportunitySQL = l_strSQL
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "SaveOpportunitySQL", , , , DescArgsFrom(l_strSQL), True
'
'End Function


Private Function SaveOpportunitySQL(ByVal v_rsData As ADODB.Recordset) As String
        
    Dim l_strSQL As String
    Dim l_cexRecordset As utilDataServiceLib.CExRecordset
    Dim l_objSQLBuilder As utilDataServiceLib.cuSQLBuilder
    
On Error GoTo eh
    
    Set l_cexRecordset = New CExRecordset
    l_cexRecordset.SetFullTableName "METHDBA.TTR_OT_OPPORTUNITY"
    l_cexRecordset.AddFields "PRIMARYKEY", True, Array("OT_OPP_NUM")
    l_cexRecordset.AddField "INDEX_TIME", Array("SETCURRENTTIMESTAMP"), Array(2)
'    If v_rsData.Status = adRecUnmodified And _
'            Trim(v_rsData!OPP_STATUS) = "C" Then
'        l_cexRecordset.AddField "TIME_STAMP", Array("CONCURRENCYKEY"), Array(True)
'    Else
'        l_cexRecordset.AddField "TIME_STAMP", Array("CONCURRENCYKEY", "TOUCH"), Array(True, True)
'    End If
' RSG - TOUCH opportunity row for each operation to ensure no concurrent operations.
    l_cexRecordset.AddField "TIME_STAMP", Array("CONCURRENCYKEY", "TOUCH"), Array(True, True)
    
    Set l_objSQLBuilder = CreateObject("utilDataServiceLib.cuSQLBuilder")
    
    Select Case v_rsData.Status
       Case adRecModified, adRecUnmodified
            l_strSQL = l_objSQLBuilder.UpdateStatement(v_rsData, l_cexRecordset)
       Case adRecNew
           l_strSQL = l_objSQLBuilder.InsertStatement(v_rsData, l_cexRecordset)
       Case adRecDeleted
            l_strSQL = l_objSQLBuilder.DeleteStatement(v_rsData, l_cexRecordset)
        Case Else
           l_strSQL = ""
    End Select
    
    SaveOpportunitySQL = l_strSQL
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    Exit Function
    
eh:
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveOpportunitySQL", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function SaveCanvassSheet( _
            ByVal v_rsSheet As ADODB.Recordset _
        ) As Boolean
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long
    
On Error GoTo eh

    v_rsSheet.Filter = adFilterPendingRecords
    If v_rsSheet.EOF Then
        SaveCanvassSheet = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    Set l_objCM = Nothing
    
    Do Until v_rsSheet.EOF
        l_strSQL = SaveCanvassSheetSQL(v_rsSheet)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsSheet.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    SaveCanvassSheet = True
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort
'    TrsRaiseError m_strCMODULENAME, "SaveCanvassSheet", , , , DescArgsFrom(l_strSQL), True
    Select Case Err.Number
        Case trErrAppCommitConflict
            TrsTraceError m_strCMODULENAME, "SaveCanvassSheet", LongAppErr(trErrAppCommitConflict), , , DescArgsFrom(l_strSQL), True
            TrsRaiseError m_strCMODULENAME, "SaveCanvassSheet", trErrAppCommitConflict
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveCanvassSheet", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function

'Private Function SaveCanvassSheetSQL( _
'            ByVal v_rsSheet As ADODB.Recordset _
'        ) As String
'
'    Dim l_strSQL As String
'    Const CTABLE As String = "METHDBA.TTR_CANV_SHEETS"
'    Dim primKeys As Variant
'    Dim l_strTimeStamp As String
'    Dim l_varIgnore As Variant
'
'    On Error GoTo eh
'    primKeys = Array("OT_OPP_NUM", "EMP_NUM")
'    Select Case v_rsSheet.Status
'       Case adRecModified
'            l_strSQL = UpdateStatement(v_rsSheet, CTABLE, primKeys, , True, "TIME_STAMP")
'       Case adRecNew
'           l_strSQL = InsertStatement(v_rsSheet, CTABLE)
'       Case adRecDeleted
'            l_strSQL = DeleteStatement(v_rsSheet, CTABLE, primKeys)
'        Case Else
'           l_strSQL = ""
'    End Select
'
'    SaveCanvassSheetSQL = l_strSQL
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "SaveCanvassSheetSQL", , , , DescArgsFrom(l_strSQL), True
'
'End Function


Private Function SaveCanvassSheetSQL(ByVal v_rsData As ADODB.Recordset) As String
        
    Dim l_strSQL As String
    Dim l_cexRecordset As utilDataServiceLib.CExRecordset
    Dim l_objSQLBuilder As utilDataServiceLib.cuSQLBuilder
    
On Error GoTo eh
    
    Set l_cexRecordset = New CExRecordset
    l_cexRecordset.SetFullTableName "METHDBA.TTR_CANV_SHEETS"
    l_cexRecordset.AddFields "PRIMARYKEY", True, Array("OT_OPP_NUM", "EMP_NUM")
    l_cexRecordset.AddField "CREATED_TIMESTAMP", Array("SETCURRENTTIMESTAMP"), Array(2)
    l_cexRecordset.AddField "TIME_STAMP", Array("CONCURRENCYKEY"), Array(True)
    
    Set l_objSQLBuilder = CreateObject("utilDataServiceLib.cuSQLBuilder")
    
    Select Case v_rsData.Status
       Case adRecModified
            l_strSQL = l_objSQLBuilder.UpdateStatement(v_rsData, l_cexRecordset)
       Case adRecNew
           l_strSQL = l_objSQLBuilder.InsertStatement(v_rsData, l_cexRecordset)
       Case adRecDeleted
            l_strSQL = l_objSQLBuilder.DeleteStatement(v_rsData, l_cexRecordset)
        Case Else
           l_strSQL = ""
    End Select
    
    SaveCanvassSheetSQL = l_strSQL
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    Exit Function
    
eh:
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveCanvassSheetSQL", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function GetOTOppCanvassSheetForUpdate( _
    ByVal v_lngOTOppNum As Long _
    ) As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString

    Set l_objCM = Nothing

    l_objCN.Open
    l_strSQL = "select OT_OPP_NUM, EMP_NUM, OT_STANDING_HRS, ROW_NUM, "
    l_strSQL = l_strSQL & "QUAL_SKILL_IND, HOURS_CHARGED, OT_REASON_CD, "
    l_strSQL = l_strSQL & "INELIG_IND, FORCE_ROW_NUM, LOC_CD,  "
    l_strSQL = l_strSQL & "PLANT_SNRTY_QTY, UNION_SNRTY_QTY, "
    l_strSQL = l_strSQL & "OT_BANK_CD, JOB_NUM, CRAFT_NUM, "
    l_strSQL = l_strSQL & "HOME_PHONE_NUM, "
    l_strSQL = l_strSQL & "CANVASS_RESP_CD, CHAR(CANVASS_CALL_TIME) AS CANVASS_CALL_TIME, "
    l_strSQL = l_strSQL & "FORCE_RESP_CD, CHAR(FORCE_CALL_TIME) AS FORCE_CALL_TIME, CALL_NOTES_TXT, "
    l_strSQL = l_strSQL & "CREATED_EMP_NUM, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, "
    l_strSQL = l_strSQL & "LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_CANV_SHEETS "
    l_strSQL = l_strSQL & "WHERE OT_OPP_NUM = " & CStr(v_lngOTOppNum) & " "
    l_strSQL = l_strSQL & "ORDER BY EMP_NUM "


    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetOTOppCanvassSheetForUpdate = l_objrs

    Set l_objrs = Nothing

    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTOppCanvassSheetForUpdate", , , , DescArgsFrom(l_strSQL)


End Function

Public Function GetOTGroupCanvassBank( _
    ByVal v_intOTGroup As Integer _
    ) As ADODB.Recordset
    
'This Method Probably is not needed any more
'It was used to get the PROJ_OR_STAND_IND
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString

    Set l_objCM = Nothing

    l_objCN.Open
    l_strSQL = "SELECT PROJ_OR_STAND_IND, OT_BANK_CD " & _
                "FROM METHDBA.TTR_OT_GROUP, " & _
                    " METHDBA.TTR_OT_BANK_CD " & _
                "WHERE OT_GROUP_NUM = " & str(v_intOTGroup) & _
                "  AND OT_STAND_BANK_CD = OT_BANK_CD"


    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetOTGroupCanvassBank = l_objrs

    Set l_objrs = Nothing

    GetObjectContext.SetComplete

    Exit Function


eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOTGroupCanvassBank", , , , DescArgsFrom(l_strSQL)

    
    
End Function


Public Function GetOpportunityResponses( _
    ByVal v_lngOTOppNum As Long _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString

    Set l_objCM = Nothing

    l_objCN.Open
    l_strSQL = "SELECT B.OT_RESP_CD, RESPONSE_DESC, C.OT_WORKED_IND, C.OT_CHRG_IND, C.OT_CONTACTED_IND " & _
                "FROM METHDBA.TTR_OT_OPPORTUNITY A, " & _
                    " METHDBA.TTR_CANV_TMPL_RESP B, " & _
                    " METHDBA.TTR_OT_RESPONSES C " & _
                "WHERE OT_OPP_NUM = " & CStr(v_lngOTOppNum) & _
                "  AND B.TEMPLATE_NUM = A.TEMPLATE_NUM " & _
                "  AND C.OT_RESP_CD = B.OT_RESP_CD  " & _
                "ORDER BY OT_RESP_CD "


    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetOpportunityResponses = l_objrs

    Set l_objrs = Nothing

    GetObjectContext.SetComplete

    Exit Function
   
    
eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOpportunityResponses", , , , DescArgsFrom(l_strSQL)
    
    
End Function



Public Function SaveOppResponses( _
            ByVal v_rsResponses As ADODB.Recordset _
        ) As Boolean

'Remove when breaking compatibility
SaveOppResponses = False


End Function


Public Function SaveOppTimeCharges( _
            ByVal v_rsTimeCharges As ADODB.Recordset _
        ) As Boolean
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long
    
On Error GoTo eh

    v_rsTimeCharges.Filter = adFilterPendingRecords
    If v_rsTimeCharges.EOF Then
        SaveOppTimeCharges = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    Set l_objCM = Nothing
    
    Do Until v_rsTimeCharges.EOF
        l_strSQL = SaveOppTimeChargesSQL(v_rsTimeCharges)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
'        l_strSQL = SaveOppScheduleSQL(v_rsTimeCharges)
'        If l_strSQL <> "" Then
'            l_objCN.Execute l_strSQL, l_lngRecsAffected
'            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
'                 Err.Raise trErrAppCommitConflict
'            End If
'        End If
        v_rsTimeCharges.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    SaveOppTimeCharges = True
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort
'    TrsRaiseError m_strCMODULENAME, "SaveOppTimeCharges", , , , DescArgsFrom(l_strSQL), True
    Select Case Err.Number
        Case trErrAppCommitConflict
            TrsTraceError m_strCMODULENAME, "SaveOppTimeCharges", LongAppErr(trErrAppCommitConflict), , , DescArgsFrom(l_strSQL), True
            TrsRaiseError m_strCMODULENAME, "SaveOppTimeCharges", trErrAppCommitConflict
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveOppTimeCharges", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function
             
'
'Private Function SaveOppTimeChargesSQL( _
'            ByVal v_rsTimeCharges As ADODB.Recordset _
'        ) As String
'
'    Dim l_strSQL As String
'    Const CTABLE As String = "METHDBA.TTR_OT_OPP_TM_CHRG"
'    Dim primKeys As Variant
'    Dim l_strTimeStamp As String
'    Dim l_varIgnore As Variant
'
'    On Error GoTo eh
'    primKeys = Array("OT_OPP_NUM", "EMP_NUM", "CREATED_TIMESTAMP")
'    Select Case v_rsTimeCharges.Status
'       Case adRecModified
'            l_strSQL = UpdateStatement(v_rsTimeCharges, CTABLE, primKeys, , True, "LAST_MOD_TIMESTAMP")
'       Case adRecNew
'           l_strSQL = InsertStatement(v_rsTimeCharges, CTABLE)
'       Case adRecDeleted
'            l_strSQL = DeleteStatement(v_rsTimeCharges, CTABLE, primKeys)
'        Case Else
'           l_strSQL = ""
'    End Select
'
'    SaveOppTimeChargesSQL = l_strSQL
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "SaveOppTimeChargesSQL", , , , DescArgsFrom(l_strSQL), True
'
'End Function


Private Function SaveOppTimeChargesSQL(ByVal v_rsData As ADODB.Recordset) As String
        
    Dim l_strSQL As String
    Dim l_cexRecordset As utilDataServiceLib.CExRecordset
    Dim l_objSQLBuilder As utilDataServiceLib.cuSQLBuilder
    
On Error GoTo eh
    
    Set l_cexRecordset = New CExRecordset
    l_cexRecordset.SetFullTableName "METHDBA.TTR_OT_OPP_TM_CHRG"
    l_cexRecordset.AddFields "PRIMARYKEY", True, Array("OT_OPP_NUM", "EMP_NUM", "CREATED_TIMESTAMP")
    l_cexRecordset.AddField "CREATED_TIMESTAMP", Array("SETCURRENTTIMESTAMP"), Array(2)
    l_cexRecordset.AddField "LAST_MOD_TIMESTAMP", Array("CONCURRENCYKEY"), Array(True)
    
    Set l_objSQLBuilder = CreateObject("utilDataServiceLib.cuSQLBuilder")
    
    Select Case v_rsData.Status
       Case adRecModified
            l_strSQL = l_objSQLBuilder.UpdateStatement(v_rsData, l_cexRecordset)
       Case adRecNew
           l_strSQL = l_objSQLBuilder.InsertStatement(v_rsData, l_cexRecordset)
       Case adRecDeleted
            l_strSQL = l_objSQLBuilder.DeleteStatement(v_rsData, l_cexRecordset)
        Case Else
           l_strSQL = ""
    End Select
    
    SaveOppTimeChargesSQL = l_strSQL
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    Exit Function
    
eh:
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveOppTimeChargesSQL", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function SaveOppSchedule( _
            ByVal v_rsScheduledOvertime As ADODB.Recordset _
        ) As Boolean
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long
    
On Error GoTo eh

    v_rsScheduledOvertime.Filter = adFilterPendingRecords
    If v_rsScheduledOvertime.EOF Then
        SaveOppSchedule = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    Set l_objCM = Nothing
    
    Do Until v_rsScheduledOvertime.EOF
        l_strSQL = SaveOppScheduleSQL(v_rsScheduledOvertime)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsScheduledOvertime.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    SaveOppSchedule = True
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort
    Select Case Err.Number
        Case trErrAppCommitConflict
            TrsTraceError m_strCMODULENAME, "SaveOppSchedule", LongAppErr(trErrAppCommitConflict), , , DescArgsFrom(l_strSQL), True
            TrsRaiseError m_strCMODULENAME, "SaveOppSchedule", trErrAppCommitConflict
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveOppSchedule", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function


'Private Function SaveOppScheduleSQL( _
'            ByVal v_rsScheduledOvertime As ADODB.Recordset _
'        ) As String
'
'    Dim l_strSQL As String
'    Const CTABLE As String = "METHDBA.TTR_BASE_HRS_EXC"
'    Dim primKeys As Variant
'    Dim l_strTimeStamp As String
'    Dim l_varIgnore As Variant
'    Dim l_aryAdditional(1, 1) As String
'
'On Error GoTo eh
'    primKeys = Array("CHARGE_DATE", "EMP_NUM", "SCHED_TYPE", "DATE", "TIME_STAMP")
'
'    Select Case v_rsScheduledOvertime.Status
'       Case adRecModified
'            l_strSQL = UpdateStatement(v_rsScheduledOvertime, CTABLE, primKeys, , True, "TIME_STAMP")
'       Case adRecNew
'           l_strSQL = InsertStatement(v_rsScheduledOvertime, CTABLE)
'       Case adRecDeleted
'            l_strSQL = DeleteStatement(v_rsScheduledOvertime, CTABLE, primKeys)
'        Case Else
'           l_strSQL = ""
'    End Select
'
'    SaveOppScheduleSQL = l_strSQL
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "SaveOppScheduleSQL", , , , DescArgsFrom(l_strSQL), True
'End Function


Private Function SaveOppScheduleSQL(ByVal v_rsData As ADODB.Recordset) As String
        
    Dim l_strSQL As String
    Dim l_cexRecordset As utilDataServiceLib.CExRecordset
    Dim l_objSQLBuilder As utilDataServiceLib.cuSQLBuilder
    
On Error GoTo eh
    
    Set l_cexRecordset = New CExRecordset
    l_cexRecordset.SetFullTableName "METHDBA.TTR_BASE_HRS_EXC"
    l_cexRecordset.AddFields "PRIMARYKEY", True, Array("CHARGE_DATE", "EMP_NUM", "SCHED_TYPE", "DATE", "TIME_STAMP")
    l_cexRecordset.AddField "TIME_STAMP", Array("CONCURRENCYKEY"), Array(True)
    
    Set l_objSQLBuilder = CreateObject("utilDataServiceLib.cuSQLBuilder")
    
    Select Case v_rsData.Status
       Case adRecModified
            l_strSQL = l_objSQLBuilder.UpdateStatement(v_rsData, l_cexRecordset)
       Case adRecNew
           l_strSQL = l_objSQLBuilder.InsertStatement(v_rsData, l_cexRecordset)
       Case adRecDeleted
            l_strSQL = l_objSQLBuilder.DeleteStatement(v_rsData, l_cexRecordset)
        Case Else
           l_strSQL = ""
    End Select
    
    SaveOppScheduleSQL = l_strSQL
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    Exit Function
    
eh:
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveOppScheduleSQL", , , , DescArgsFrom(l_strSQL), True

End Function


Public Function SaveEMPRVOTSchedule( _
                    ByVal v_rsEMPRVOTSchedule As ADODB.Recordset _
                ) As Boolean
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long
    
On Error GoTo eh
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    Set l_objCM = Nothing
    
    Do Until v_rsEMPRVOTSchedule.EOF
        l_strSQL = SaveEMPRVOTScheduleSQL(v_rsEMPRVOTSchedule)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsEMPRVOTSchedule.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    SaveEMPRVOTSchedule = True
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort
'    TrsRaiseError m_strCMODULENAME, "SaveEMPRVOTSchedule", , , , DescArgsFrom(l_strSQL), True
    Select Case Err.Number
        Case trErrAppCommitConflict
            TrsTraceError m_strCMODULENAME, "SaveEMPRVOTSchedule", LongAppErr(trErrAppCommitConflict), , , DescArgsFrom(l_strSQL), True
            TrsRaiseError m_strCMODULENAME, "SaveEMPRVOTSchedule", trErrAppCommitConflict
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveEMPRVOTSchedule", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function
Public Function SaveJobTaskOTSchedule( _
                    ByVal v_rsJobTaskOTSchedule As ADODB.Recordset _
                ) As Boolean
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long
    
On Error GoTo eh
    
    'Change connection string here for Sybase!!!!!!
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString("EMPRV")
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    Set l_objCM = Nothing
    
    Do Until v_rsJobTaskOTSchedule.EOF
        l_strSQL = SaveJobTaskOTScheduleSQL(v_rsJobTaskOTSchedule)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsJobTaskOTSchedule.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    SaveJobTaskOTSchedule = True
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort
    Select Case Err.Number
        Case trErrAppCommitConflict
            TrsTraceError m_strCMODULENAME, "SaveJobTaskOTSchedule", LongAppErr(trErrAppCommitConflict), , , DescArgsFrom(l_strSQL), True
            TrsRaiseError m_strCMODULENAME, "SaveJobTaskOTSchedule", trErrAppCommitConflict
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveJobTaskOTSchedule", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function


'Private Function SaveEMPRVOTScheduleSQL( _
'                        ByVal v_rsEMPRVOTSchedule As ADODB.Recordset _
'                ) As String
'
'    Dim l_strSQL As String
'    Const CTABLE As String = "METHDBA.TTR_OT_EMPRV"
'
'On Error GoTo eh
'
'    l_strSQL = InsertStatement(v_rsEMPRVOTSchedule, CTABLE)
'
'    SaveEMPRVOTScheduleSQL = l_strSQL
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "SaveEMPRVOTScheduleSQL", , , , DescArgsFrom(l_strSQL), True
'
'End Function


Private Function SaveEMPRVOTScheduleSQL(ByVal v_rsData As ADODB.Recordset) As String
        
    Dim l_strSQL As String
    Dim l_cexRecordset As utilDataServiceLib.CExRecordset
    Dim l_objSQLBuilder As utilDataServiceLib.cuSQLBuilder
    
On Error GoTo eh
    
    Set l_cexRecordset = New CExRecordset
    l_cexRecordset.SetFullTableName "METHDBA.TTR_OT_EMPRV"
    
    Set l_objSQLBuilder = CreateObject("utilDataServiceLib.cuSQLBuilder")
    l_strSQL = l_objSQLBuilder.InsertStatement(v_rsData, l_cexRecordset)
    
    SaveEMPRVOTScheduleSQL = l_strSQL
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    Exit Function
    
eh:
    Set l_objSQLBuilder = Nothing
    Set l_cexRecordset = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveEMPRVOTScheduleSQL", , , , DescArgsFrom(l_strSQL), True

End Function
Private Function SaveJobTaskOTScheduleSQL(ByVal v_rsData As ADODB.Recordset) As String
        
    Dim l_strSQL As String
    'Dim l_cexRecordset As utilDataServiceLib.CExRecordset
    'Dim l_objSQLBuilder As utilDataServiceLib.cuSQLBuilder
    
On Error GoTo eh
    
    'Set l_cexRecordset = New CExRecordset
    'l_cexRecordset.SetFullTableName "METHDBA.TTR_???"
    
    'Set l_objSQLBuilder = CreateObject("utilDataServiceLib.cuSQLBuilder")
    'l_strSQL = l_objSQLBuilder.InsertStatement(v_rsData, l_cexRecordset)
    
    l_strSQL = "insert into dbo.tpotrishours" & _
         "(emplnbr,chrgdt,sttm,workeddt,endtm,wrkgrpnbr,fk_jobtaskjobnbr," & _
         "fk_jobtasktasknbr,fk_jobstepseqnbr,fk_emprvdbcode,excptcd,distrid," & _
         "rptdhrs,otind,chgind,fk_costcatcode)" & _
         "values(" & _
            CStr(v_rsData.Fields("EMP_NUM")) & "," & _
            DateAsEnclosedString(v_rsData.Fields("CHARGE_DATE")) & ", " & _
            SinqleQuotedString(v_rsData.Fields("START_TIME")) & ", " & _
            DateAsEnclosedString(v_rsData.Fields("DATE_WORKED")) & ", " & _
            SinqleQuotedString(v_rsData.Fields("END_TIME")) & ", " & _
            CStr(v_rsData.Fields("WG_NUM")) & ", " & _
            SinqleQuotedString(v_rsData.Fields("JOB_NBR")) & ", " & _
            SinqleQuotedString(v_rsData.Fields("JOBTASK_NBR")) & ", " & _
            CStr(v_rsData.Fields("JOBTASK_STEP_NBR")) & ", " & _
            SinqleQuotedString(v_rsData.Fields("EMPRV_DATABASE_ID")) & ", " & _
            SinqleQuotedString(v_rsData.Fields("EXCEPTION_CD")) & ", " & _
            SinqleQuotedString(v_rsData.Fields("DISTRIBUTION_ID")) & ", " & _
            CStr(v_rsData.Fields("HOURS")) & ", " & _
            SinqleQuotedString(v_rsData.Fields("OT_FLG")) & ", " & _
            SinqleQuotedString(v_rsData.Fields("CHG_IND")) & ", " & _
            SinqleQuotedString(v_rsData.Fields("DEPT_NUM")) & ")"

    SaveJobTaskOTScheduleSQL = l_strSQL
    'Set l_objSQLBuilder = Nothing
    'Set l_cexRecordset = Nothing
    Exit Function
    
eh:
    'Set l_objSQLBuilder = Nothing
    'Set l_cexRecordset = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveJobTaskOTScheduleSQL", , , , DescArgsFrom(l_strSQL), True

End Function



Public Function GetOvertimeExceptions() As ADODB.Recordset
'RED051 and RED062
'This function returns overtime exception codes for
'contract and executive employees
'The exec_contract_ind ('C' or 'E') can be filter
'to get a list specfic to contract or executive employees
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
On Error GoTo GetOvertimeExceptions_EH
    
    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    Set l_objCM = Nothing
    l_objCN.Open
    l_strSQL = "Select exception_cd , EXCEPTION_DESC " & _
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
                "Select exception_cd , EXCEPTION_DESC " & _
                "from methdba.ttr_exception_cds, " & _
                "     methdba.TTR_EARN_CODES " & _
                "Where EARNINGS_CD = CONTRACT_EARN_CD " & _
                "  and EXEC_CONTRACT_IND = 'C' " & _
                "  and otherpay_ind in ('T','B') " & _
                "  and EXCEPTION_BANK_CD = 'O' " & _
                "  and EARNINGS_CD <> '99' " & _
                "  and EXCEPTION_CD not in " & _
                " (select RG_VALUE_CD " & _
                "   from METHDBA.TTR_RG_GROUP " & _
                "   where RG_CD = 6)  "
    l_strSQL = l_strSQL & "UNION " & _
                "Select b.exception_cd , EXCEPTION_DESC " & _
                "from methdba.ttr_exception_cds a, " & _
                    " methdba.ttr_rg_bank_exc_cd b " & _
                "Where b.rg_cd = 107 " & _
                "  and a.exception_cd = b.exception_cd "
    l_strSQL = l_strSQL & " ORDER BY 1 "
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    If l_objrs.BOF And l_objrs.EOF Then
        Set GetOvertimeExceptions = Nothing
    Else
        Set GetOvertimeExceptions = l_objrs
    End If
    GetObjectContext.SetComplete
    Set l_objrs = Nothing
    Exit Function

GetOvertimeExceptions_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetOvertimeExceptions", , , , DescArgsFrom(l_strSQL), True
End Function


Public Function GetOvertimeResponses() As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString

    Set l_objCM = Nothing

    l_objCN.Open
    l_strSQL = "SELECT OT_RESP_CD, RESPONSE_DESC, OT_WORKED_IND, OT_CHRG_IND, OT_CONTACTED_IND " & _
                "FROM METHDBA.TTR_OT_RESPONSES  " & _
                "ORDER BY OT_RESP_CD "

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetOvertimeResponses = l_objrs

    Set l_objrs = Nothing

    GetObjectContext.SetComplete

    Exit Function
   
    
eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeResponses", , , , DescArgsFrom(l_strSQL)
    
    
End Function

Public Function GetNonCanavassedOvertimeCanvasses( _
    ByVal v_intLocationCode As Integer, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_intOTGroup As Integer = 0, _
    Optional v_blnOpenCanvassesOnly As Boolean = False, _
    Optional v_blnPostedCanvassesOnly As Boolean = False _
    ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strStatusCriteria As String
    Dim l_strOTGroupCriteria As String
    Dim l_strOTLocCriteria As String

    On Error GoTo eh

    If v_blnOpenCanvassesOnly Then
         l_strStatusCriteria = " AND A.OPP_STATUS = 'C'"
    Else
        If v_blnPostedCanvassesOnly Then
            l_strStatusCriteria = " AND A.OPP_STATUS = 'P'"
        Else
            l_strStatusCriteria = " AND A.OPP_STATUS <> 'V'"
        End If
    End If
    
    If v_intOTGroup = 0 Then
        l_strOTGroupCriteria = ""
    Else
        l_strOTGroupCriteria = " AND A.OT_GROUP_NUM = " & str(v_intOTGroup)
    End If
    
    If v_intLocationCode = 0 Then
        l_strOTLocCriteria = ""
    Else
        l_strOTLocCriteria = "  AND A.LOC_CD = " & str(v_intLocationCode)
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT  A.OT_OPP_NUM, " & _
            "CASE A.OPP_STATUS WHEN 'C' then 'OPEN' " & _
                            "  WHEN 'P' then 'POSTED' " & _
                            "  WHEN 'V' then 'VOID' " & _
                            "  ELSE 'ERROR' end as OPP_STATUS "
    l_strSQL = l_strSQL & _
            ",CASE A.SHIFT_CD WHEN 'A' then 'AM' " & _
                            " WHEN 'P' then 'PM' " & _
                            " WHEN 'O' then 'OWL' " & _
                            " WHEN 'D' then 'DAY' " & _
                            " WHEN 'N' then 'NIGHT' " & _
                            " ELSE 'ERROR' end as SHIFT_DESC "
    l_strSQL = l_strSQL & ", A.OT_OPP_DESC, A.OT_GROUP_NUM, E.OT_GROUP_DESC, A.CHARGE_DATE "
    l_strSQL = l_strSQL & ", A.DATE_OF_OT as START_DATE , A.START_TIME "
    l_strSQL = l_strSQL & ", A.END_DATE, A.END_TIME, A.OT_REASON_CD "
    l_strSQL = l_strSQL & ", B.OT_REASON_DESC, A.TEMPLATE_NUM, C.CANVASS_DESC "
    l_strSQL = l_strSQL & ", D.EMP_NAME, A.LAST_MOD_EMP_NUM "
    l_strSQL = l_strSQL & ", C.TEMPL_LOCATION, F.LOC_DESC "
    l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_OPPORTUNITY A, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_OT_REASONS B, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_CANV_TEMPL C, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_EMPLOYEE D, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_OT_GROUP E, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_LOCATIONS F "
    l_strSQL = l_strSQL & "WHERE OT_OPP_NUM >= 32000000 "
    l_strSQL = l_strSQL & "  AND A.DATE_OF_OT <= " & DateAsEnclosedString(v_datEndDate)
    l_strSQL = l_strSQL & "  AND A.DATE_OF_OT >= " & DateAsEnclosedString(v_datStartDate)
    l_strSQL = l_strSQL & l_strOTLocCriteria
'    l_strSQL = l_strSQL & "  AND A.LOC_CD = " & str(v_intLocationCode)
    l_strSQL = l_strSQL & l_strStatusCriteria
    l_strSQL = l_strSQL & l_strOTGroupCriteria
    l_strSQL = l_strSQL & "  AND C.TEMPLATE_NUM = A.TEMPLATE_NUM "
    l_strSQL = l_strSQL & "  AND B.OT_REASON_CD = A.OT_REASON_CD  "
    l_strSQL = l_strSQL & "  AND D.EMP_NUM = A.LAST_MOD_EMP_NUM "
    l_strSQL = l_strSQL & "  AND E.OT_GROUP_NUM = A.OT_GROUP_NUM "
    l_strSQL = l_strSQL & "  AND F.LOC_CD = A.LOC_CD "
    l_strSQL = l_strSQL & "ORDER BY A.OT_OPP_NUM "

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetNonCanavassedOvertimeCanvasses = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetNonCanavassedOvertimeCanvasses", , , , DescArgsFrom(l_strSQL)
    
    
End Function

Public Function GetCanvasses( _
    ByVal v_intLocationCode As Integer, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_strCanvassType As String = "A", _
    Optional v_intOTGroup As Integer = 0, _
    Optional v_blnOpenCanvassesOnly As Boolean = False, _
    Optional v_blnPostedCanvassesOnly As Boolean = False _
    ) As ADODB.Recordset
    
'v_strCanvassType can have 1 of 3 values
'       C -- Canvassed Opportunities
'       A -- All Opportunities
'       N -- Non Canvassed Opportunities
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strStatusCriteria As String
    Dim l_strOTGroupCriteria As String
    Dim l_strOTLocCriteria As String
    Dim l_strOppNumCriteria As String

    On Error GoTo eh

    If v_blnOpenCanvassesOnly Then
         l_strStatusCriteria = " AND A.OPP_STATUS = 'C'"
    Else
        If v_blnPostedCanvassesOnly Then
            l_strStatusCriteria = " AND A.OPP_STATUS = 'P'"
        Else
            l_strStatusCriteria = " AND A.OPP_STATUS <> 'V'"
        End If
    End If
    
    If v_intOTGroup = 0 Then
        l_strOTGroupCriteria = ""
    Else
        l_strOTGroupCriteria = " AND A.OT_GROUP_NUM = " & str(v_intOTGroup)
    End If

    If v_intLocationCode = 0 Then
        l_strOTLocCriteria = ""
    Else
        l_strOTLocCriteria = "  AND A.LOC_CD = " & str(v_intLocationCode)
    End If
    
    Select Case v_strCanvassType
        Case "C"
            l_strOppNumCriteria = " AND OT_OPP_NUM >= 300000 " & _
                    " AND OT_OPP_NUM < 32000000 "
        Case "A"
            l_strOppNumCriteria = " AND OT_OPP_NUM >= 300000 "
        Case "N"
            l_strOppNumCriteria = " AND OT_OPP_NUM >= 32000000 "
        Case Else
            l_strOppNumCriteria = " AND OT_OPP_NUM >= 300000 "
    End Select

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT  A.OT_OPP_NUM, " & _
            "CASE A.OPP_STATUS WHEN 'C' then 'OPEN' " & _
                            "  WHEN 'P' then 'POSTED' " & _
                            "  WHEN 'V' then 'VOID' " & _
                            "  ELSE 'ERROR' end as OPP_STATUS "
    l_strSQL = l_strSQL & _
            ",CASE A.SHIFT_CD WHEN 'A' then 'AM' " & _
                            " WHEN 'P' then 'PM' " & _
                            " WHEN 'O' then 'OWL' " & _
                            " WHEN 'D' then 'DAY' " & _
                            " WHEN 'N' then 'NIGHT' " & _
                            " ELSE 'ERROR' end as SHIFT_DESC "
    l_strSQL = l_strSQL & ", A.OT_OPP_DESC, A.OT_GROUP_NUM, E.OT_GROUP_DESC, A.CHARGE_DATE "
    l_strSQL = l_strSQL & ", A.DATE_OF_OT as START_DATE , A.START_TIME "
    l_strSQL = l_strSQL & ", A.END_DATE, A.END_TIME, A.OT_REASON_CD "
    l_strSQL = l_strSQL & ", B.OT_REASON_DESC, A.TEMPLATE_NUM, C.CANVASS_DESC "
    l_strSQL = l_strSQL & ", D.EMP_NAME, A.LAST_MOD_EMP_NUM "
    l_strSQL = l_strSQL & ", C.TEMPL_LOCATION, F.LOC_DESC "
    l_strSQL = l_strSQL & ", CHAR(A.TIME_STAMP) AS LAST_MOD_TIMESTAMP "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_OT_OPPORTUNITY A, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_OT_REASONS B, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_CANV_TEMPL C, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_EMPLOYEE D, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_OT_GROUP E, "
    l_strSQL = l_strSQL & "     METHDBA.TTR_LOCATIONS F "
    l_strSQL = l_strSQL & "WHERE A.CHARGE_DATE <= " & DateAsEnclosedString(v_datEndDate)
    l_strSQL = l_strSQL & "  AND A.CHARGE_DATE >= " & DateAsEnclosedString(v_datStartDate)
    l_strSQL = l_strSQL & l_strOppNumCriteria
    l_strSQL = l_strSQL & l_strOTLocCriteria
'    l_strSQL = l_strSQL & "  AND A.LOC_CD = " & str(v_intLocationCode)
    l_strSQL = l_strSQL & l_strStatusCriteria
    l_strSQL = l_strSQL & l_strOTGroupCriteria
    l_strSQL = l_strSQL & "  AND C.TEMPLATE_NUM = A.TEMPLATE_NUM "
    l_strSQL = l_strSQL & "  AND B.OT_REASON_CD = A.OT_REASON_CD  "
    l_strSQL = l_strSQL & "  AND D.EMP_NUM = A.LAST_MOD_EMP_NUM "
    l_strSQL = l_strSQL & "  AND E.OT_GROUP_NUM = A.OT_GROUP_NUM "
    l_strSQL = l_strSQL & "  AND F.LOC_CD = A.LOC_CD "
    l_strSQL = l_strSQL & "ORDER BY A.OT_OPP_NUM "

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetCanvasses = l_objrs
    
    Set l_objrs = Nothing
    
    GetObjectContext.SetComplete

    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCanvasses", , , , DescArgsFrom(l_strSQL)
    
    
End Function


Public Function GetEmployeePrimaryLocation( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_datEffDate As Date _
                ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open

    l_strSQL = "SELECT A.EMP_NUM, B.WG_NUM, C.LOC_CD, D.LOC_DESC, E.EMP_NAME,"
    l_strSQL = l_strSQL & " F.PROJECT_ID,"
    l_strSQL = l_strSQL & " VALUE(G.PPM_IND, 'N') AS PPM_IND,"
    l_strSQL = l_strSQL & " VALUE(G.EFF_START_DATE, '01/01/0100') AS PROJ_START_DATE,"
    l_strSQL = l_strSQL & " VALUE(G.EFF_END_DATE, '01/01/0100') AS PROJ_END_DATE"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_CANV_SHEETS A"
    l_strSQL = l_strSQL & " INNER JOIN"
    l_strSQL = l_strSQL & " (SELECT EMP_NUM, WG_NUM"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EMP_WG"
    l_strSQL = l_strSQL & " WHERE PRIMARY_IND = 'Y'"
    l_strSQL = l_strSQL & " AND " & DateAsEnclosedString(v_datEffDate)
    l_strSQL = l_strSQL & " BETWEEN  EFF_START_DATE AND EFF_END_DATE) B"
    l_strSQL = l_strSQL & " ON A.EMP_NUM = B.EMP_NUM"
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_ASSOCIATIONS C"
    l_strSQL = l_strSQL & " ON B.WG_NUM = C.WG_NUM"
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_LOCATIONS D"
    l_strSQL = l_strSQL & " ON C.LOC_CD = D.LOC_CD"
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_EMPLOYEE E"
    l_strSQL = l_strSQL & " ON A.EMP_NUM = E.EMP_NUM"
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_WORK_GROUP F"
    l_strSQL = l_strSQL & " ON B.WG_NUM = F.WG_NUM"
    l_strSQL = l_strSQL & " LEFT OUTER JOIN METHDBA.TTR_PROJECTS G"
    l_strSQL = l_strSQL & " ON F.PROJECT_ID = G.PROJECT_ID"
    l_strSQL = l_strSQL & " WHERE A.OT_OPP_NUM = " & v_lngOTOppNum
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing

    Set GetEmployeePrimaryLocation = l_objrs
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function
   
eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeePrimaryLocation", , , , DescArgsFrom(l_strSQL)
    
End Function


Public Function GetEmployeePrimaryLocationFromList( _
                    ByVal v_aryEmpNum As Variant, _
                    ByVal v_datEffDate As Date _
                ) As ADODB.Recordset
    
    Dim l_strSelection As String
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    If IsArray(v_aryEmpNum) Then
        l_strSelection = BuildInClause("A.EMP_NUM", v_aryEmpNum, False)
    Else
        l_strSelection = "A.EMP_NUM = " & CStr(v_aryEmpNum)
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open

    l_strSQL = "SELECT A.EMP_NUM, B.WG_NUM, C.LOC_CD, D.LOC_DESC, A.EMP_NAME"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EMPLOYEE A"
    l_strSQL = l_strSQL & " INNER JOIN"
    l_strSQL = l_strSQL & " (SELECT EMP_NUM, WG_NUM"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EMP_WG"
    l_strSQL = l_strSQL & " WHERE PRIMARY_IND = 'Y'"
    l_strSQL = l_strSQL & " AND " & DateAsEnclosedString(v_datEffDate)
    l_strSQL = l_strSQL & " BETWEEN  EFF_START_DATE AND EFF_END_DATE) B"
    l_strSQL = l_strSQL & " ON A.EMP_NUM = B.EMP_NUM"
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_ASSOCIATIONS C"
    l_strSQL = l_strSQL & " ON B.WG_NUM = C.WG_NUM"
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_LOCATIONS D"
    l_strSQL = l_strSQL & " ON C.LOC_CD = D.LOC_CD"
    l_strSQL = l_strSQL & " WHERE " & l_strSelection
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing

    Set GetEmployeePrimaryLocationFromList = l_objrs
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function
   
eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeePrimaryLocationFromList", , , , DescArgsFrom(l_strSQL)
    
End Function


Public Function GetCanvassLocation( _
                    ByVal v_aryOTOppNum As Variant _
                ) As ADODB.Recordset
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strSelection As String
    

On Error GoTo eh

    If IsArray(v_aryOTOppNum) Then
        l_strSelection = BuildInClause("A.OT_OPP_NUM", v_aryOTOppNum, False)
    Else
        l_strSelection = "A.OT_OPP_NUM = " & CStr(v_aryOTOppNum)
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open

    l_strSQL = "SELECT A.OT_OPP_NUM, A.OT_OPP_DESC, A.LOC_CD, B.LOC_DESC"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_OT_OPPORTUNITY A"
    l_strSQL = l_strSQL & " INNER JOIN METHDBA.TTR_LOCATIONS B"
    l_strSQL = l_strSQL & " ON A.LOC_CD = B.LOC_CD"
    l_strSQL = l_strSQL & " WHERE " & l_strSelection

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing

    Set GetCanvassLocation = l_objrs
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function
   
eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCanvassLocation", , , , DescArgsFrom(l_strSQL)
    
End Function
Public Function GetCallLog( _
                    ByVal v_lngOTOppNum As Long) As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT OT_OPP_NUM, EMP_NUM, CALL_NUM, CHAR(CALL_ST_TIMESTAMP) AS CALL_ST_TIMESTAMP, " & _
                " CHAR(CALL_END_TIMESTAMP) AS CALL_END_TIMESTAMP" & _
                " FROM METHDBA.TTR_OT_OPP_CALLOUT" & _
                " WHERE OT_OPP_NUM = " & v_lngOTOppNum

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetCallLog = l_objrs
    
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCallLog", , , , DescArgsFrom(l_strSQL)

End Function


Public Function GetTTATravelLocations( _
                    ByVal v_lngEffYear As Long) As ADODB.Recordset

    Dim l_objCN As New ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT USER_TYPE, EFF_YEAR, START_LOC_CD, END_LOC_CD " & _
                " FROM METHDBA.TTR_TRAV_ALLOWANCE" & _
                " WHERE EFF_YEAR = " & v_lngEffYear

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetTTATravelLocations = l_objrs
    
    Set l_objrs = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTTATravelLocations", , , , DescArgsFrom(l_strSQL)

End Function
'BEGIN PREM TRIS CONVERSION PHASE
Public Function GetOpportunityBankCharges(ByVal vlOpportunityNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetOpportunityBankCharges_EH
Dim lsSQL As String
Dim ladrsParms As ADODB.Recordset

    Set GetOpportunityBankCharges = Nothing

    lsSQL = "SELECT EMP_NUM, EMP_NAME" & _
        "       , OT_RESP_CD, OT_STANDING_HRS" & _
        "       , SUM(ACC_HRS) AS ACC_HRS, SUM(REF_HRS) AS REF_HRS" & _
        "       , (OT_STANDING_HRS + SUM(ACC_HRS) + SUM(REF_HRS)) AS NEW_STANDING_HRS" & _
        "       , FORCE_RESP_CD, CANVASS_RESP_CD" & _
        "       , ROW_NUM, FORCE_ROW_NUM" & _
        " FROM  (SELECT A.EMP_NUM, C.EMP_NAME" & _
        "               , CASE WHEN (B.FORCE_RESP_CD) <> '' THEN B.FORCE_RESP_CD" & _
        "               ELSE B.CANVASS_RESP_CD END AS OT_RESP_CD" & _
        "               , B.OT_STANDING_HRS" & _
        "               , CASE WHEN (A.ACC_REF_IND) = 'A' THEN A.HOURS_CHARGED" & _
        "               ELSE DECIMAL(0, 7, 2) END AS ACC_HRS" & _
        "               , CASE WHEN (A.ACC_REF_IND) = 'R' THEN A.HOURS_CHARGED" & _
        "               ELSE DECIMAL(0, 7, 2) END AS REF_HRS" & _
        "               , B.FORCE_RESP_CD, B.CANVASS_RESP_CD" & _
        "               , B.ROW_NUM, B.FORCE_ROW_NUM"
    lsSQL = lsSQL & _
        "       FROM    METHDBA.TTR_OT_BANK_TM_CHG A, METHDBA.TTR_CANV_SHEETS B" & _
        "               , METHDBA.TTR_EMPLOYEE C" & _
        "       WHERE   A.OT_OPP_NUM = ?" & _
        "       AND     A.OT_OPP_NUM = B.OT_OPP_NUM" & _
        "       AND     A.EMP_NUM = B.EMP_NUM" & _
        "       AND     A.EMP_NUM = C.EMP_NUM) PREM" & _
        " GROUP BY  EMP_NUM, EMP_NAME" & _
        "           , OT_RESP_CD, OT_STANDING_HRS" & _
        "           , FORCE_RESP_CD, CANVASS_RESP_CD" & _
        "           , ROW_NUM, FORCE_ROW_NUM"

    lsSQL = lsSQL & " UNION ALL "

    lsSQL = lsSQL & _
        " SELECT    A.EMP_NUM, B.EMP_NAME" & _
        "           , CASE WHEN (FORCE_RESP_CD) <> '' THEN FORCE_RESP_CD " & _
        "           ELSE CANVASS_RESP_CD END AS OT_RESP_CD " & _
        "           , A.OT_STANDING_HRS" & _
        "           , DECIMAL(0, 7, 2) AS ACC_HRS, DECIMAL(0, 7, 2) AS REF_HRS" & _
        "           , A.OT_STANDING_HRS AS NEW_STANDING_HRS" & _
        "           , A.FORCE_RESP_CD, A.CANVASS_RESP_CD" & _
        "           , A.ROW_NUM, A.FORCE_ROW_NUM" & _
        " FROM      METHDBA.TTR_CANV_SHEETS A, METHDBA.TTR_EMPLOYEE B" & _
        " WHERE     A.OT_OPP_NUM = ?" & _
        " AND       A.EMP_NUM = B.EMP_NUM" & _
        " AND       A.EMP_NUM NOT IN (SELECT    EMP_NUM" & _
        "                           FROM        METHDBA.TTR_OT_BANK_TM_CHG" & _
        "                           WHERE       OT_OPP_NUM = ?)" & _
        " ORDER BY ROW_NUM" & _
        " FOR FETCH ONLY WITH UR;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOpportunityNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOpportunityNumber
        .Update

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOpportunityNumber
        .Update
    End With

    Set GetOpportunityBankCharges = NewExecuteSQL(lsSQL _
                        , ladrsParms _
                        , adOpenStatic _
                        , adLockReadOnly _
                        , adCmdText)
    Set ladrsParms = Nothing

    GetObjectContext.SetComplete

    Exit Function
GetOpportunityBankCharges_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "GetOpportunityBankCharges"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError m_strCMODULENAME _
                , "GetOpportunityBankCharges" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "GetOpportunityBankCharges" _
                , , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
'END PREM TRIS CONVERSION PHASE

'Begin Prem TRIS 5.6
Public Function ConstructCanvassSheet2(ByVal viOTGroupNumber As Integer _
                , ByVal viOTBankCode As Integer _
                , ByVal vdtChargeDate As Date _
                ) As ADODB.Recordset
On Error GoTo ConstructCanvassSheet2_EH
Dim lsSQL1 As String, lsSQL2 As String
Dim lsSQL3 As String, lsSQL4 As String
Dim ladrsParms As ADODB.Recordset
Dim ladcn As ADODB.Connection

    Set ConstructCanvassSheet2 = Nothing

    'Declare Global Temp SQL Statement
    lsSQL1 = "DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_OT_EMP" & _
        " (EMP_NUM INTEGER NOT NULL DEFAULT 0, HOME_PHONE_NUM CHAR(10) NOT NULL DEFAULT ''" & _
        " , JOB_NUM CHAR(6) NOT NULL DEFAULT '',  DATE_OF_BIRTH DATE NOT NULL" & _
        " , DEPT_NUM CHAR(3) NOT NULL DEFAULT '', UNION_CD CHAR(2) NOT NULL DEFAULT '' " & _
        " , EMP_NAME CHAR(30) NOT NULL DEFAULT '') CCSID EBCDIC" & _
        " ON COMMIT PRESERVE ROWS;"

    'Inser into Global Temp SQL Statement
    lsSQL2 = " INSERT INTO SESSION.GTT_OT_EMP" & _
        " (EMP_NUM, HOME_PHONE_NUM" & _
        " , JOB_NUM, DATE_OF_BIRTH" & _
        " , DEPT_NUM, UNION_CD, EMP_NAME)" & _
        " SELECT    DISTINCT A.EMP_NUM, D.HOME_PHONE_NUM" & _
        "           , D.JOB_NUM, D.DATE_OF_BIRTH" & _
        "           , D.DEPT_NUM, D.UNION_CD, D.EMP_NAME " & _
        " FROM      METHDBA.TTR_EMP_WG A" & _
        "           INNER JOIN METHDBA.TTR_OT_GROUP_WG B" & _
        "           ON B.WG_NUM = A.WG_NUM" & _
        "           AND B.OT_GROUP_NUM = ?" & _
        "           AND ? BETWEEN A.EFF_START_DATE" & _
        "           AND A.EFF_END_DATE" & _
        "           INNER JOIN METHDBA.TTR_WORK_GROUP C" & _
        "           ON C.WG_NUM = A.WG_NUM" & _
        "           AND C.WG_NUM NOT IN (99995, 99996, 99997, 99998, 99999)" & _
        "           INNER JOIN METHDBA.TTR_EMPLOYEE D" & _
        "           ON D.EMP_NUM = A.EMP_NUM" & _
        "           AND D.EMP_STATUS <> 'T' ;"

    'Construct Canvass Sheet SQL Statement
    lsSQL3 = "SELECT    A.EMP_NUM, A.HOME_PHONE_NUM" & _
        "       , INTEGER(COALESCE(DEPT_SNRTY_QTY, 0)) AS DEPT_SNRTY_QTY" & _
        "       , INTEGER(COALESCE(PLANT_SNRTY_QTY, 0)) AS PLANT_SNRTY_QTY" & _
        "       , INTEGER(COALESCE(UNION_SNRTY_QTY, 0)) AS UNION_SNRTY_QTY" & _
        "       , INTEGER(COALESCE(CORP_SNRTY_QTY, 0)) AS CORP_SNRTY_QTY" & _
        "       , INTEGER(COALESCE(LOP_SNRTY_QTY, 0)) AS LOP_SNRTY_QTY" & _
        "       , INTEGER(COALESCE(JOB_SNRTY_QTY, 0)) AS JOB_SNRTY_QTY" & _
        "       , INTEGER(COALESCE(CMPY_SNRTY_QTY, 0)) AS CMPY_SNRTY_QTY" & _
        "       , A.JOB_NUM, A.DATE_OF_BIRTH" & _
        "       , B.CRAFT_NUM, F.LOC_CD" & _
        "       , DECIMAL(COALESCE(TOT_OT_HRS, 99999.90), 7, 2) AS OT_STANDING_HRS" & _
        "       , SMALLINT (0) AS ROW_NUM, SMALLINT (0) AS FORCE_ROW_NUM" & _
        "       , SMALLINT (0) AS LOC_SORT, SMALLINT (0) AS WG_SORT" & _
        "       , A.DEPT_NUM, A.EMP_NAME "

    lsSQL3 = lsSQL3 & _
        " FROM  (SELECT *" & _
        "       FROM    SESSION.GTT_OT_EMP) A" & _
        "       INNER JOIN METHDBA.TTR_JOBS B" & _
        "       ON B.JOB_NUM = A.JOB_NUM" & _
        "       INNER JOIN (SELECT  C.EMP_NUM, D.WG_NUM" & _
        "                           , E.LOC_CD" & _
        "                   FROM    SESSION.GTT_OT_EMP C" & _
        "                           INNER JOIN METHDBA.TTR_EMP_WG D" & _
        "                           ON D.EMP_NUM = C.EMP_NUM" & _
        "                           AND D.PRIMARY_IND = 'Y'" & _
        "                           AND ? BETWEEN D.EFF_START_DATE" & _
        "                           AND D.EFF_END_DATE" & _
        "                           INNER JOIN METHDBA.TTR_ASSOCIATIONS E" & _
        "                           ON E.WG_NUM = D.WG_NUM) F" & _
        "       ON F.EMP_NUM = A.EMP_NUM"

    lsSQL3 = lsSQL3 & _
        "       LEFT OUTER JOIN (SELECT EMP_NUM, TOT_OT_HRS" & _
        "                       FROM    METHDBA.TTR_OT_BANKS" & _
        "                       WHERE   OT_BANK_CD = ?) G" & _
        "       ON G.EMP_NUM = A.EMP_NUM" & _
        "       LEFT OUTER JOIN (SELECT EMP_NUM, DEPT_NUM" & _
        "                               , UNION_CD, JOB_NUM" & _
        "                               , DEPT_SNRTY_QTY, PLANT_SNRTY_QTY" & _
        "                               , UNION_SNRTY_QTY, CORP_SNRTY_QTY" & _
        "                               , LOP_SNRTY_QTY, JOB_SNRTY_QTY" & _
        "                               , CMPY_SNRTY_QTY" & _
        "                       FROM    METHDBA.TTR_SENIORITY) H" & _
        "       ON H.EMP_NUM = A.EMP_NUM" & _
        "       AND H.DEPT_NUM = A.DEPT_NUM" & _
        "       AND H.UNION_CD = A.UNION_CD" & _
        "       AND H.JOB_NUM = A.JOB_NUM " & _
        " ORDER BY A.EMP_NUM;"

    'Drop Global Temp SQL Statement
    lsSQL4 = "DROP TABLE SESSION.GTT_OT_EMP;"

    'Create Parameters
    Set ladrsParms = CreateParameterRS
    'Add records
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = viOTGroupNumber
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtChargeDate
        .Update
    End With

    'Open Connection
    Set ladcn = OpenConnection

    'Declare GTT
    ExecuteOTHERSQL lsSQL1 _
        , ladcn _
        , adCmdText

    'Insert into GTT
    NewExecuteOtherSQL lsSQL2 _
        , ladrsParms _
        , ladcn _
        , adCmdText

    'Delete/Add record
    With ladrsParms
        .MoveFirst
        .Find "PARM_ORDER = 1"
        If Not (.EOF) Then
            .Delete adAffectCurrent
        End If

        .AddNew
        !PARM_ORDER = 3
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = viOTBankCode
        .Update
    End With

    'Get Canvass Sheet
    Set ConstructCanvassSheet2 = NewExecuteSQLWithConnection(lsSQL3 _
                        , ladrsParms _
                        , ladcn _
                        , adOpenStatic _
                        , adLockBatchOptimistic _
                        , adCmdText)
    Set ladrsParms = Nothing

    'Drop GTT
    ExecuteOTHERSQL lsSQL4 _
        , ladcn _
        , adCmdText
    
    If Not (ladcn Is Nothing) Then
        ladcn.Close
    End If
    Set ladcn = Nothing

    GetObjectContext.SetComplete

    Exit Function
ConstructCanvassSheet2_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    If Not (ladcn Is Nothing) Then
        'Drop Global Temp SQL Statement
        lsSQL4 = "DROP TABLE SESSION.GTT_OT_EMP;"
        'Drop GTT
        ExecuteOTHERSQL lsSQL4 _
            , ladcn _
            , adCmdText
    End If

    Set ladcn = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "ConstructCanvassSheet2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError m_strCMODULENAME _
                , "ConstructCanvassSheet2" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL1 & " " & lsSQL2 & " " & lsSQL3 & " " & lsSQL4) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "ConstructCanvassSheet2" _
                , , , _
                , DescArgsFrom(lsSQL1 & " " & lsSQL2 & " " & lsSQL3 & " " & lsSQL4) _
                , True
    End Select

End Function
Public Function GetOvertimeCanvasses2(ByVal viLocation As Integer _
                , ByVal vdtStartDate As Date _
                , ByVal vdtEndDate As Date _
                , Optional viOTGroup As Integer = 0 _
                , Optional vbOpenCanvassOnly As Boolean = False _
                , Optional vbPostedCanvassOnly As Boolean = False _
                ) As ADODB.Recordset
On Error GoTo GetOvertimeCanvasses2_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetOvertimeCanvasses2 = Nothing

    lsSQL = " SELECT    A.OT_OPP_NUM, A.OT_OPP_DESC" & _
        "       , A.CHARGE_DATE" & _
        "       , A.DATE_OF_OT AS START_DATE, A.START_TIME" & _
        "       , A.END_DATE, A.END_TIME" & _
        "       , A.OT_GROUP_NUM, B.OT_GROUP_DESC" & _
        "       , A.TEMPLATE_NUM, C.CANVASS_DESC" & _
        "       , C.TEMPL_LOCATION, D.LOC_DESC" & _
        "       , A.OT_REASON_CD, E.OT_REASON_DESC" & _
        "       , A.LAST_MOD_EMP_NUM, F.EMP_NAME" & _
        "       , CHAR(A.TIME_STAMP) AS LAST_MOD_TIMESTAMP" & _
        "       , CASE (A.OPP_STATUS)" & _
        "       WHEN 'C' THEN 'OPEN'" & _
        "       WHEN 'P' THEN 'POSTED'" & _
        "       WHEN 'V' THEN 'VOID'" & _
        "       ELSE 'ERROR' END AS OPP_STATUS" & _
        "       , CASE (A.SHIFT_CD)" & _
        "       WHEN 'A' THEN 'AM'" & _
        "       WHEN 'P' THEN 'PM'" & _
        "       WHEN 'O' THEN 'OWL'" & _
        "       WHEN 'D' THEN 'DAY'" & _
        "       WHEN 'N' THEN 'NIGHT'" & _
        "       ELSE 'ERROR' END AS SHIFT_DESC"

    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_OT_OPPORTUNITY A" & _
        "       INNER JOIN METHDBA.TTR_OT_GROUP B" & _
        "       ON B.OT_GROUP_NUM = A.OT_GROUP_NUM" & _
        "       INNER JOIN METHDBA.TTR_CANV_TEMPL C" & _
        "       ON C.TEMPLATE_NUM = A.TEMPLATE_NUM" & _
        "       INNER JOIN METHDBA.TTR_LOCATIONS D" & _
        "       ON D.LOC_CD = A.LOC_CD" & _
        "       INNER JOIN METHDBA.TTR_OT_REASONS E" & _
        "       ON E.OT_REASON_CD = A.OT_REASON_CD" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE F" & _
        "       ON F.EMP_NUM = A.LAST_MOD_EMP_NUM" & _
        " WHERE A.OT_OPP_NUM BETWEEN 300000" & _
        " AND   31999999" & _
        " AND   A.DATE_OF_OT BETWEEN ?" & _
        " AND   ?"

    If (vbOpenCanvassOnly) Then
        lsSQL = lsSQL & " AND   A.OPP_STATUS = 'C'"
    Else
        If (vbPostedCanvassOnly) Then
            lsSQL = lsSQL & " AND A.OPP_STATUS = 'P'"
        Else
            lsSQL = lsSQL & " AND A.OPP_STATUS <> 'V'"
        End If
    End If

    Set ladrsParms = CreateParameterRS

    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    If (viLocation > 0) Then
        lsSQL = lsSQL & " AND  A.LOC_CD = ?"

        With ladrsParms
            .AddNew
            !PARM_ORDER = 3
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = viLocation
            .Update
        End With
    End If

    If (viOTGroup > 0) Then
        lsSQL = lsSQL & " AND  A.OT_GROUP_NUM = ?"

        With ladrsParms
            .AddNew
            !PARM_ORDER = 4
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = viOTGroup
            .Update
        End With
    End If

    lsSQL = lsSQL & " ORDER BY A.OT_OPP_NUM" & _
            " FOR FETCH ONLY WITH UR;"

    Set GetOvertimeCanvasses2 = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockBatchOptimistic _
                            , adCmdText)
    Set ladrsParms = Nothing

    GetObjectContext.SetComplete

    Exit Function
GetOvertimeCanvasses2_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "GetOvertimeCanvasses2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError m_strCMODULENAME _
                , "GetOvertimeCanvasses2" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "GetOvertimeCanvasses2" _
                , , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetOTOpportunity2(ByVal vlOpportunity As Long _
                ) As ADODB.Recordset
On Error GoTo GetOTOpportunity2_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetOTOpportunity2 = Nothing

    lsSQL = " SELECT    A.OT_OPP_NUM, A.OT_OPP_DESC" & _
            "           , A.OPP_STATUS, A.LOC_CD" & _
            "           , A.CHARGE_DATE, A.SHIFT_CD" & _
            "           , A.DATE_OF_OT, A.END_DATE" & _
            "           , A.START_TIME, A.END_TIME" & _
            "           , A.START_LUNCH, A.END_LUNCH" & _
            "           , A.OT_REASON_CD, A.TEMPLATE_NUM" & _
            "           , A.OT_GROUP_NUM, A.REFUSE_RULE_NUM" & _
            "           , A.TURNOVER_START_TM, A.TURNOVER_END_TM" & _
            "           , A.MOD_REASON_DESC, A.CALL_OUT_CANV_IND" & _
            "           , C.LOC_DESC, CAST(LEFT(A.CRAFT_JOB, 3) AS CHAR(3)) AS COMP_SHIFT_CD" & _
            "           , A.LAST_MOD_EMP_NUM, A.CREATED_BY_EMP_NUM" & _
            "           , CHAR(A.TIME_STAMP) AS LAST_MOD_TIMESTAMP, CHAR(A.INDEX_TIME) AS CREATED_TIMESTAMP" & _
            "           , B.EMP_NAME AS LAST_MOD_NAME, B.EMP_NAME AS CREATED_BY_NAME" & _
            " FROM      METHDBA.TTR_OT_OPPORTUNITY A" & _
            "           INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
            "           ON B.EMP_NUM = A.LAST_MOD_EMP_NUM" & _
            "           INNER JOIN METHDBA.TTR_LOCATIONS C" & _
            "           ON C.LOC_CD = A.LOC_CD" & _
            " WHERE     A.OT_OPP_NUM = ?;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOpportunity
        .Update
    End With

    Set GetOTOpportunity2 = NewExecuteSQL(lsSQL _
                        , ladrsParms _
                        , adOpenStatic _
                        , adLockBatchOptimistic)
    Set ladrsParms = Nothing

    GetObjectContext.SetComplete

    Exit Function
GetOTOpportunity2_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOpportunity2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOpportunity2" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOpportunity2" _
                , , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetOTOppCanvassSheet2(ByVal vlOpportunity As Long _
                ) As ADODB.Recordset
On Error GoTo GetOTOppCanvassSheet2_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetOTOppCanvassSheet2 = Nothing

    lsSQL = " SELECT    A.OT_OPP_NUM, A.EMP_NUM" & _
            "           , B.EMP_NAME, A.LOC_CD" & _
            "           , A.QUAL_SKILL_IND, A.INELIG_IND" & _
            "           , CASE A.INELIG_IND " & _
            "               WHEN 'N' THEN 'Yes' " & _
            "               WHEN 'Y' THEN 'No' " & _
            "               WHEN 'L' THEN 'Least' " & _
            "               WHEN 'O' THEN 'Owl/AM' " & _
            "               WHEN '1' THEN 'Yes/12H' " & _
            "               ELSE ' ' " & _
            "             END AS INELIG_DESC " & _
            "           , A.HOME_PHONE_NUM, A.OT_STANDING_HRS" & _
            "           , A.JOB_NUM, A.CRAFT_NUM" & _
            "           , A.ROW_NUM, A.FORCE_ROW_NUM" & _
            "           , LEFT(DIGITS(A.PLANT_SNRTY_QTY), 3) AS PLANT_SNRTY_MONTHS" & _
            "           , LEFT(DIGITS(A.PLANT_SNRTY_QTY), 3) AS PLANT_SNRTY_DAYS" & _
            "           , LEFT(DIGITS(A.UNION_SNRTY_QTY), 3) AS UNION_SNRTY_MONTHS" & _
            "           , LEFT(DIGITS(A.UNION_SNRTY_QTY), 3) AS UNION_SNRTY_DAYS" & _
            "           , A.CANVASS_RESP_CD, A.FORCE_RESP_CD" & _
            "           , CHAR(A.CANVASS_CALL_TIME) AS CANVASS_CALL_TIME, CHAR(A.FORCE_CALL_TIME) AS FORCE_CALL_TIME" & _
            "           , A.CALL_NOTES_TXT" & _
            "           , SPACE(20) AS QUAL_CDS, SPACE(20) AS SPEC_SKILL_CDS" & _
            "           , SPACE(1) AS ON_DUTY_IND, SPACE(3) AS ON_DUTY_EXC_CD" & _
            "           , A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM" & _
            "           , CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(A.TIME_STAMP) AS LAST_MOD_TIMESTAMP"
    lsSQL = lsSQL & _
            " FROM      METHDBA.TTR_CANV_SHEETS A" & _
            "           INNER JOIN METHDBA.TTR_EMPLOYEE B" & _
            "           ON B.EMP_NUM = A.EMP_NUM" & _
            "           WHERE A.OT_OPP_NUM = ?" & _
            " ORDER BY A.ROW_NUM;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOpportunity
        .Update
    End With

    Set GetOTOppCanvassSheet2 = NewExecuteSQL(lsSQL _
                                , ladrsParms _
                                , adOpenStatic _
                                , adLockBatchOptimistic _
                                , adCmdText)
    Set ladrsParms = Nothing

    GetObjectContext.SetComplete

    Exit Function
GetOTOppCanvassSheet2_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppCanvassSheet2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppCanvassSheet2" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppCanvassSheet2" _
                , , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetOTOppPhoneCalls2(ByVal vlOpportunity As Long _
                ) As ADODB.Recordset
On Error GoTo GetOTOppPhoneCalls2_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetOTOppPhoneCalls2 = Nothing

    lsSQL = " SELECT    A.OT_OPP_NUM, A.EMP_NUM" & _
            "           , B.EMP_NAME, A.CALL_NUM" & _
            "           , CHAR(DATE(A.CALL_ST_TIMESTAMP)) || ' ' || REPLACE(CHAR(TIME(A.CALL_ST_TIMESTAMP)),'.',':') AS CALL_START" & _
            "           , CHAR(DATE(A.CALL_END_TIMESTAMP)) || ' ' || REPLACE(CHAR(TIME(A.CALL_END_TIMESTAMP)),'.',':') AS CALL_END" & _
            "           , A.DIALED_NUM, A.CREATED_EMP_NUM" & _
            "           , C.EMP_NAME AS CALLER_NAME" & _
            "           , CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
            " FROM      METHDBA.TTR_OT_OPP_CALLOUT A" & _
            "           INNER JOIN  METHDBA.TTR_EMPLOYEE B" & _
            "           ON B.EMP_NUM  = A.EMP_NUM" & _
            "           INNER JOIN METHDBA.TTR_EMPLOYEE C" & _
            "           ON C.EMP_NUM = A.CREATED_EMP_NUM" & _
            " WHERE     A.OT_OPP_NUM = ?" & _
            " ORDER BY  A.CALL_ST_TIMESTAMP, A.EMP_NUM"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOpportunity
        .Update
    End With

    Set GetOTOppPhoneCalls2 = NewExecuteSQL(lsSQL _
                                , ladrsParms _
                                , adOpenStatic _
                                , adLockBatchOptimistic _
                                , adCmdText)
    Set ladrsParms = Nothing

    GetObjectContext.SetComplete

    Exit Function
GetOTOppPhoneCalls2_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppBankCharges2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppBankCharges2" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppBankCharges2" _
                , , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetOpportunityResponses2(ByVal vlOpportunity As Long _
                ) As ADODB.Recordset
On Error GoTo GetOpportunityResponses2_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetOpportunityResponses2 = Nothing

    lsSQL = " SELECT    B.OT_RESP_CD, C.RESPONSE_DESC" & _
            "           , C.OT_WORKED_IND, C.OT_CHRG_IND" & _
            "           , C.OT_CONTACTED_IND" & _
            " FROM      METHDBA.TTR_OT_OPPORTUNITY A" & _
            "           INNER JOIN METHDBA.TTR_CANV_TMPL_RESP B" & _
            "           ON B.TEMPLATE_NUM = A.TEMPLATE_NUM" & _
            "           INNER JOIN METHDBA.TTR_OT_RESPONSES C" & _
            "           ON C.OT_RESP_CD = B.OT_RESP_CD" & _
            " WHERE     A.OT_OPP_NUM = ?" & _
            " ORDER BY OT_RESP_CD;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOpportunity
        .Update
    End With

    Set GetOpportunityResponses2 = NewExecuteSQL(lsSQL _
                                , ladrsParms _
                                , adOpenStatic _
                                , adLockBatchOptimistic _
                                , adCmdText)
    Set ladrsParms = Nothing

    GetObjectContext.SetComplete

    Exit Function
GetOpportunityResponses2_EH:
    GetObjectContext.SetAbort
    
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "GetOpportunityResponses2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError m_strCMODULENAME _
                , "GetOpportunityResponses2" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "GetOpportunityResponses2" _
                , , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetNonCanavassedOvertimeCanvasses2(ByVal viLocation As Integer _
                , ByVal vdtStartDate As Date _
                , ByVal vdtEndDate As Date _
                , Optional viOTGroup As Integer = 0 _
                , Optional vbOpenCanvassOnly As Boolean = False _
                , Optional vbPostedCanvassOnly As Boolean = False _
                ) As ADODB.Recordset
On Error GoTo GetNonCanavassedOvertimeCanvasses2_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetNonCanavassedOvertimeCanvasses2 = Nothing

    lsSQL = " SELECT    A.OT_OPP_NUM, A.OT_OPP_DESC" & _
        "       , A.CHARGE_DATE" & _
        "       , A.DATE_OF_OT AS START_DATE, A.START_TIME" & _
        "       , A.END_DATE, A.END_TIME" & _
        "       , A.OT_GROUP_NUM, B.OT_GROUP_DESC" & _
        "       , A.TEMPLATE_NUM, C.CANVASS_DESC" & _
        "       , C.TEMPL_LOCATION, D.LOC_DESC" & _
        "       , A.OT_REASON_CD, E.OT_REASON_DESC" & _
        "       , A.LAST_MOD_EMP_NUM, F.EMP_NAME" & _
        "       , CHAR(A.TIME_STAMP) AS LAST_MOD_TIMESTAMP" & _
        "       , CASE (A.OPP_STATUS)" & _
        "       WHEN 'C' THEN 'OPEN'" & _
        "       WHEN 'P' THEN 'POSTED'" & _
        "       WHEN 'V' THEN 'VOID'" & _
        "       ELSE 'ERROR' END AS OPP_STATUS" & _
        "       , CASE (A.SHIFT_CD)" & _
        "       WHEN 'A' THEN 'AM'" & _
        "       WHEN 'P' THEN 'PM'" & _
        "       WHEN 'O' THEN 'OWL'" & _
        "       WHEN 'D' THEN 'DAY'" & _
        "       WHEN 'N' THEN 'NIGHT'" & _
        "       ELSE 'ERROR' END AS SHIFT_DESC"

    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_OT_OPPORTUNITY A" & _
        "       INNER JOIN METHDBA.TTR_OT_GROUP B" & _
        "       ON B.OT_GROUP_NUM = A.OT_GROUP_NUM" & _
        "       INNER JOIN METHDBA.TTR_CANV_TEMPL C" & _
        "       ON C.TEMPLATE_NUM = A.TEMPLATE_NUM" & _
        "       INNER JOIN METHDBA.TTR_LOCATIONS D" & _
        "       ON D.LOC_CD = A.LOC_CD" & _
        "       INNER JOIN METHDBA.TTR_OT_REASONS E" & _
        "       ON E.OT_REASON_CD = A.OT_REASON_CD" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE F" & _
        "       ON F.EMP_NUM = A.LAST_MOD_EMP_NUM" & _
        " WHERE A.OT_OPP_NUM >= 32000000" & _
        " AND   A.DATE_OF_OT BETWEEN ?" & _
        " AND   ?"

    If (vbOpenCanvassOnly) Then
        lsSQL = lsSQL & " AND   A.OPP_STATUS = 'C'"
    Else
        If (vbPostedCanvassOnly) Then
            lsSQL = lsSQL & " AND A.OPP_STATUS = 'P'"
        Else
            lsSQL = lsSQL & " AND A.OPP_STATUS <> 'V'"
        End If
    End If

    Set ladrsParms = CreateParameterRS

    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    If (viLocation > 0) Then
        lsSQL = lsSQL & " AND  A.LOC_CD = ?"

        With ladrsParms
            .AddNew
            !PARM_ORDER = 3
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = viLocation
            .Update
        End With
    End If

    If (viOTGroup > 0) Then
        lsSQL = lsSQL & " AND  A.OT_GROUP_NUM = ?"

        With ladrsParms
            .AddNew
            !PARM_ORDER = 4
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = viOTGroup
            .Update
        End With
    End If

    lsSQL = lsSQL & " ORDER BY A.OT_OPP_NUM" & _
            " FOR FETCH ONLY WITH UR;"

    Set GetNonCanavassedOvertimeCanvasses2 = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockBatchOptimistic _
                            , adCmdText)
    Set ladrsParms = Nothing

    GetObjectContext.SetComplete

    Exit Function
GetNonCanavassedOvertimeCanvasses2_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "GetNonCanavassedOvertimeCanvasses2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError m_strCMODULENAME _
                , "GetNonCanavassedOvertimeCanvasses2" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "GetNonCanavassedOvertimeCanvasses2" _
                , , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Public Function GetCanvasses2(ByVal viLocation As Integer _
                , ByVal vdtStartDate As Date _
                , ByVal vdtEndDate As Date _
                , Optional vsCanvassType As String = "A" _
                , Optional viOTGroup As Integer = 0 _
                , Optional vbOpenCanvassOnly As Boolean = False _
                , Optional vbPostedCanvassOnly As Boolean = False _
                ) As ADODB.Recordset
On Error GoTo GetCanvasses2_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetCanvasses2 = Nothing

    lsSQL = " SELECT    A.OT_OPP_NUM, A.OT_OPP_DESC" & _
        "       , A.CHARGE_DATE" & _
        "       , A.DATE_OF_OT AS START_DATE, A.START_TIME" & _
        "       , A.END_DATE, A.END_TIME" & _
        "       , A.OT_GROUP_NUM, B.OT_GROUP_DESC" & _
        "       , A.TEMPLATE_NUM, C.CANVASS_DESC" & _
        "       , C.TEMPL_LOCATION, D.LOC_DESC" & _
        "       , A.OT_REASON_CD, E.OT_REASON_DESC" & _
        "       , A.LAST_MOD_EMP_NUM, F.EMP_NAME" & _
        "       , CHAR(A.TIME_STAMP) AS LAST_MOD_TIMESTAMP" & _
        "       , CASE (A.OPP_STATUS)" & _
        "       WHEN 'C' THEN 'OPEN'" & _
        "       WHEN 'P' THEN 'POSTED'" & _
        "       WHEN 'V' THEN 'VOID'" & _
        "       ELSE 'ERROR' END AS OPP_STATUS" & _
        "       , CASE (A.SHIFT_CD)" & _
        "       WHEN 'A' THEN 'AM'" & _
        "       WHEN 'P' THEN 'PM'" & _
        "       WHEN 'O' THEN 'OWL'" & _
        "       WHEN 'D' THEN 'DAY'" & _
        "       WHEN 'N' THEN 'NIGHT'" & _
        "       ELSE 'ERROR' END AS SHIFT_DESC"

    lsSQL = lsSQL & _
        " FROM  METHDBA.TTR_OT_OPPORTUNITY A" & _
        "       INNER JOIN METHDBA.TTR_OT_GROUP B" & _
        "       ON B.OT_GROUP_NUM = A.OT_GROUP_NUM" & _
        "       INNER JOIN METHDBA.TTR_CANV_TEMPL C" & _
        "       ON C.TEMPLATE_NUM = A.TEMPLATE_NUM" & _
        "       INNER JOIN METHDBA.TTR_LOCATIONS D" & _
        "       ON D.LOC_CD = A.LOC_CD" & _
        "       INNER JOIN METHDBA.TTR_OT_REASONS E" & _
        "       ON E.OT_REASON_CD = A.OT_REASON_CD" & _
        "       INNER JOIN METHDBA.TTR_EMPLOYEE F" & _
        "       ON F.EMP_NUM = A.LAST_MOD_EMP_NUM"

'Canvass Type can have 1 of 3 values
'   C -- Canvassed Opportunities
'   A -- All Opportunities
'   N -- Non Canvassed Opportunities
    Select Case (Trim(vsCanvassType))
        Case "A"
            lsSQL = lsSQL & " WHERE A.OT_OPP_NUM >= 300000"
        Case "C"
            lsSQL = lsSQL & " WHERE A.OT_OPP_NUM BETWEEN 300000" & _
                    " AND   31999999"
        Case "N"
            lsSQL = lsSQL & " WHERE A.OT_OPP_NUM >= 32000000"
    End Select

    lsSQL = lsSQL & " AND   A.DATE_OF_OT BETWEEN ?" & _
            " AND   ?"

    If (vbOpenCanvassOnly) Then
        lsSQL = lsSQL & " AND   A.OPP_STATUS = 'C'"
    Else
        If (vbPostedCanvassOnly) Then
            lsSQL = lsSQL & " AND A.OPP_STATUS = 'P'"
        Else
            lsSQL = lsSQL & " AND A.OPP_STATUS <> 'V'"
        End If
    End If

    Set ladrsParms = CreateParameterRS

    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtStartDate
        .Update

        .AddNew
        !PARM_ORDER = 2
        !PARM_NAME = ""
        !PARM_TYPE = adDate
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vdtEndDate
        .Update
    End With

    If (viLocation > 0) Then
        lsSQL = lsSQL & " AND  A.LOC_CD = ?"

        With ladrsParms
            .AddNew
            !PARM_ORDER = 3
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = viLocation
            .Update
        End With
    End If

    If (viOTGroup > 0) Then
        lsSQL = lsSQL & " AND  A.OT_GROUP_NUM = ?"

        With ladrsParms
            .AddNew
            !PARM_ORDER = 4
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = viOTGroup
            .Update
        End With
    End If

    lsSQL = lsSQL & " ORDER BY A.OT_OPP_NUM" & _
            " FOR FETCH ONLY WITH UR;"

    Set GetCanvasses2 = NewExecuteSQL(lsSQL _
                            , ladrsParms _
                            , adOpenStatic _
                            , adLockBatchOptimistic _
                            , adCmdText)
    Set ladrsParms = Nothing

    GetObjectContext.SetComplete

    Exit Function
GetCanvasses2_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "GetCanvasses2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError m_strCMODULENAME _
                , "GetCanvasses2" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "GetCanvasses2" _
                , , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function

Public Function GetOTOppCharges2(ByVal vlOpportunity As Long _
                ) As ADODB.Recordset
On Error GoTo GetOTOppCharges2_EH

    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim oADOHelper As utilADOHelper.cuADOHelper

    Set oADOHelper = New utilADOHelper.cuADOHelper

    l_strSQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetOTOppCharges2.sql")

    oADOHelper.ADO.CreateNumericParameter "OT_OPP_NUM1", adInteger, 0, 0, vlOpportunity
    oADOHelper.ADO.CreateNumericParameter "OT_OPP_NUM2", adInteger, 0, 0, vlOpportunity
    
    Set l_objrs = oADOHelper.ExecuteSelect(l_strSQL, adOpenStatic, adLockBatchOptimistic)
   
    Set GetOTOppCharges2 = l_objrs
    
    Set l_objrs = Nothing
    Set oADOHelper = Nothing
    
    GetObjectContext.SetComplete
    Exit Function
GetOTOppCharges2_EH:
    GetObjectContext.SetAbort

    Set l_objrs = Nothing
    Set oADOHelper = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppCharges2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppCharges2" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(l_strSQL) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppCharges2" _
                , , , _
                , DescArgsFrom(l_strSQL) _
                , True
    End Select
End Function

'Public Function GetOTOppCharges2(ByVal vlOpportunity As Long _
'                ) As ADODB.Recordset
'On Error GoTo GetOTOppCharges2_EH
'Dim lsSQL As String, ladrsParms As ADODB.Recordset
'
'    Set GetOTOppCharges2 = Nothing
'
'    lsSQL = " SELECT    A.OT_OPP_NUM, A.EMP_NUM" & _
'            "           , B.EMP_NAME, A.DATE_WORKED" & _
'            "           , A.CHARGE_DATE, A.START_TIME" & _
'            "           , A.END_TIME, A.HOURS" & _
'            "           , A.OT_REASON_CD, A.WG_NUM" & _
'            "           , A.EXCEPTION_CD, OVER_EXC_CD" & _
'            "           , A.JOB_NUM, A.COMP_SHIFT_CD" & _
'            "           , A.DISTRIBUTION_ID" & _
'            "           , CASE WHEN RTRIM(LTRIM(DISTRIBUTION_ID)) = '' THEN SPACE(1)" & _
'            "           WHEN RTRIM(LTRIM(SUBSTR(DISTRIBUTION_ID,4,1))) = '' THEN 'S'" & _
'            "           WHEN SUBSTR(DISTRIBUTION_ID,6,2) = 'WR'" & _
'            "           AND RTRIM(LTRIM(SUBSTR(DISTRIBUTION_ID,5,1))) = '' THEN 'W'" & _
'            "           WHEN JOB_NBR <> '' THEN 'J'" & _
'            "           ELSE 'A' END AS DISTRIBUTION_TYPE" & _
'            "           , EXCEPTION_DESC AS DISTRIBUTION_DESC" & _
'            "           , JOB_NBR, JOBTASK_NBR, JOBTASK_STEP_NBR, EMPRV_DATABASE_ID" & _
'            "           , CASE TRAVEL_PAY_HRS" & _
'            "           WHEN 0 THEN 'N'" & _
'            "           ELSE 'Y' END AS TRAVEL_PAY_IND" & _
'            "           , A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM" & _
'            "           , CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP"
'    lsSQL = lsSQL & _
'            " FROM      METHDBA.TTR_OT_OPP_TM_CHRG A, METHDBA.TTR_EMPLOYEE B" & _
'            "           , METHDBA.TTR_EXCEPTION_CDS C" & _
'            " WHERE     A.OT_OPP_NUM = ?" & _
'            " AND       B.EMP_NUM = A.EMP_NUM" & _
'            " AND       C.EXCEPTION_CD = A.EXCEPTION_CD" & _
'            " ORDER BY A.EMP_NUM, A.CHARGE_DATE, A.DATE_WORKED, A.START_TIME"
'
'    Set ladrsParms = CreateParameterRS
'    With ladrsParms
'        .AddNew
'        !PARM_ORDER = 1
'        !PARM_NAME = ""
'        !PARM_TYPE = adInteger
'        !PARM_DIRECTION = adParamInput
'        !PARM_SIZE = 0
'        !PARM_VALUE = vlOpportunity
'        .Update
'    End With
'
'    Set GetOTOppCharges2 = NewExecuteSQL(lsSQL _
'                                , ladrsParms _
'                                , adOpenStatic _
'                                , adLockBatchOptimistic _
'                                , adCmdText)
'    Set ladrsParms = Nothing
'
'    GetObjectContext.SetComplete
'    Exit Function
'GetOTOppCharges2_EH:
'    GetObjectContext.SetAbort
'
'    Set ladrsParms = Nothing
'
'    'Call Standard Module Function
'    'Raise the appropriate error message depending on the error number
'    Select Case (Err.Number)
'        Case trErrutOpenConnection
'            TrsRaiseError m_strCMODULENAME _
'                , "GetOTOppCharges2"
'        Case trErrdb_ExecuteSQL_F
'            TrsRaiseError m_strCMODULENAME _
'                , "GetOTOppCharges2" _
'                , trErrdb_ExecuteSQL_F _
'                , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'        Case Else
'            TrsRaiseError m_strCMODULENAME _
'                , "GetOTOppCharges2" _
'                , , , _
'                , DescArgsFrom(lsSQL) _
'                , True
'    End Select
'
'End Function

Public Function GetOTOppIneligibilities2(ByVal vlOpportunity As Long _
                ) As ADODB.Recordset
On Error GoTo GetOTOppIneligibilities2_EH
Dim lsSQL As String, ladrsParms As ADODB.Recordset

    Set GetOTOppIneligibilities2 = Nothing

    lsSQL = " SELECT    A.OT_OPP_NUM, A.EMP_NUM" & _
            "           , C.EMP_NAME, A.INELG_CD" & _
            "           , B.INELG_DESC, DATE(A.START_TIMESTAMP) AS START_DATE" & _
            "           , REPLACE(LEFT(CHAR(TIME(A.START_TIMESTAMP)),5),'.',':') AS START_TIME" & _
            "           , DATE(A.END_TIMESTAMP) AS END_DATE" & _
            "           , REPLACE(LEFT(CHAR(TIME(A.END_TIMESTAMP)),5),'.',':') AS END_TIME" & _
            "           , A.CHARGE_DATE, A.WORK_HRS, A.EXCEPTION_CD" & _
            "           , A.CREATED_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
            " FROM      METHDBA.TTR_OT_OPP_INELIG A, METHDBA.TTR_OT_INELIG B" & _
            "           , METHDBA.TTR_EMPLOYEE C" & _
            " WHERE     A.OT_OPP_NUM = ?" & _
            " AND       B.INELG_CD = A.INELG_CD" & _
            " AND       C.EMP_NUM = A.EMP_NUM" & _
            " ORDER BY A.EMP_NUM, A.INELG_CD;"

    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlOpportunity
        .Update
    End With

    Set GetOTOppIneligibilities2 = NewExecuteSQL(lsSQL _
                                , ladrsParms _
                                , adOpenStatic _
                                , adLockBatchOptimistic _
                                , adCmdText)
    Set ladrsParms = Nothing

    GetObjectContext.SetComplete
    Exit Function
GetOTOppIneligibilities2_EH:
    GetObjectContext.SetAbort
    Set ladrsParms = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppIneligibilities2"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppIneligibilities2" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "GetOTOppIneligibilities2" _
                , , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
Private Function CreateParameterRS() As ADODB.Recordset
On Error GoTo CreateParameterRS_EH
Dim ladrs As ADODB.Recordset

    Set ladrs = New ADODB.Recordset
    With ladrs
        .Fields.Append "PARM_ORDER", adInteger, , adFldUpdatable
        .Fields.Append "PARM_NAME", adVarChar, 1000, adFldUpdatable + adFldIsNullable
        .Fields.Append "PARM_TYPE", adInteger, , adFldUpdatable
        .Fields.Append "PARM_DIRECTION", adInteger, , adFldUpdatable
        .Fields.Append "PARM_SIZE", adInteger, , adFldUpdatable + adFldIsNullable
        .Fields.Append "PARM_VALUE", adVariant, , adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateParameterRS = ladrs
    Set ladrs = Nothing

    Exit Function
CreateParameterRS_EH:
    Set ladrs = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "CreateParameterRS"
End Function
'End Prem TRIS 5.6

Private Sub Class_Initialize()
    Set m_ResourceMgr = New utilADOHelper.cuResourceMgr
End Sub

Private Sub Class_Terminate()
    Set m_ResourceMgr = Nothing
End Sub

