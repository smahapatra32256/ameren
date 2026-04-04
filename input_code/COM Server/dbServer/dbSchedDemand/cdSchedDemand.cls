VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdSchedDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdSchedDemand"
Private Const m_datCOPTIONALSTARTDATE As Date = #1/1/100#
Private Const m_datCOPTIONALENDDATE As Date = #12/31/9999#
'//-----used the following for testing
Private Const m_strConnectStringReg As String = "SHADOWDIRECT"
'DB2TEST_UDB
'SHADOWDIRECT
' 5/17/02 RSG moved to mTRISConstants.bas
'Public Type tScopeParm
'    lngMinStartDate As Date
'    lngMaxEndDate As Date
'    lngWgNum As Long
'    lngLineNum As Long
'    lngEmpNum As Long
'End Type

'//-----Blue260
'SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE, EMP_NUM
', CREATED_EMP_NUM, LAST_MOD_EMP_NUM, CREATED_TIMESTAMP, LAST_MOD_TIMESTAMP
', WG_NUM AS KEY_WG_NUM, LINE_NUM AS KEY_LINE_NUM, START_DATE AS KEY_START_DATE
', END_DATE AS KEY_END_DATE, EMP_NUM AS KEY_EMP_NUM,' ' AS RECORD_STATUS
'From METHDBA.TTR_LINE_ASGNM
'Where WG_NUM = 4026
'AND START_DATE >= '1/1/2001'
'AND END_DATE <= '12/31/2002'
'AND LINE_NUM BETWEEN 1 AND 10
' ORDER BY WG_NUM, START_DATE, LINE_NUM
Public Function GetAssignments(ByVal v_lngWgNum As Long, ByVal v_datStartDate As Date, Optional ByVal v_datEndDate As Date = m_datCOPTIONALENDDATE, _
                                Optional ByVal v_lngStartLine As Long = -999, Optional v_lngEndLine As Long = -999) As ADODB.Recordset
    Dim l_objCN As ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString(m_strConnectStringReg)
    
    Set l_objCM = Nothing

    l_objCN.Open

    l_strSQL = "SELECT A.WG_NUM, A.LINE_NUM, A.START_DATE, A.END_DATE, A.EMP_NUM, E.EMP_NAME, A.RELIEF_COMP"
    l_strSQL = l_strSQL & ", A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM"
    l_strSQL = l_strSQL & ", CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP"
    l_strSQL = l_strSQL & ", A.WG_NUM AS KEY_WG_NUM, A.LINE_NUM AS KEY_LINE_NUM, A.START_DATE AS KEY_START_DATE"
    l_strSQL = l_strSQL & ", A.END_DATE AS KEY_END_DATE, A.EMP_NUM AS KEY_EMP_NUM, 'L' AS RECORD_STATUS"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_LINE_ASGNM A, METHDBA.TTR_EMPLOYEE E"
    l_strSQL = l_strSQL & " WHERE A.WG_NUM = " & v_lngWgNum
    l_strSQL = l_strSQL & " AND A.EMP_NUM = E.EMP_NUM"
    
'    l_strSQL = l_strSQL & " AND A.START_DATE >= " & dbFormat(v_datStartDate)
'    If v_datEndDate <> #12/31/9999# Then
'        l_strSQL = l_strSQL & " AND A.END_DATE <= " & dbFormat(v_datEndDate)
'    End If
    
    l_strSQL = l_strSQL & " AND A.END_DATE >= " & dbFormat(v_datStartDate)
    If v_datEndDate <> m_datCOPTIONALENDDATE Then
        l_strSQL = l_strSQL & " AND A.START_DATE <= " & dbFormat(v_datEndDate)
    End If
    
    If v_lngStartLine <> -999 And v_lngEndLine <> -999 Then
        l_strSQL = l_strSQL & " AND A.LINE_NUM BETWEEN " & v_lngStartLine & " AND " & v_lngEndLine
    ElseIf v_lngStartLine <> -999 Then
        l_strSQL = l_strSQL & " AND A.LINE_NUM >= " & v_lngStartLine
    End If
    
    l_strSQL = l_strSQL & " ORDER BY A.WG_NUM, A.START_DATE, A.LINE_NUM"

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetAssignments = l_objrs

    Set l_objrs = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing

    TrsRaiseError m_strCMODULENAME, "GetAssignments", , , , DescArgsFrom(l_strSQL), True
End Function

'//-----this will be called from busWorkGroup.cbsWorkGroupTrans.PostValidateAssignmentRequests
Public Function GetAssignmentsInOtherWG(ByVal v_rsWorkGroupAssignments As ADODB.Recordset) As ADODB.Recordset
    Dim l_objCN As ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim bFirst As Boolean
    
    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString(m_strConnectStringReg)
    
    Set l_objCM = Nothing

    l_objCN.Open

    l_strSQL = "SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE, EMP_NUM "
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_LINE_ASGNM WHERE"
    
    If RecordsExist(v_rsWorkGroupAssignments) Then
        With v_rsWorkGroupAssignments
            .MoveFirst
            bFirst = True
            
            Do While Not .EOF
                If bFirst Then
                    bFirst = False
                    l_strSQL = l_strSQL & " (EMP_NUM = " & !EMP_NUM
                Else
                    l_strSQL = l_strSQL & " OR (EMP_NUM = " & !EMP_NUM
                End If
                
                If Not IsNull(!EFF_END_DATE) Then
                    If !EFF_END_DATE <> m_datCOPTIONALENDDATE Then
                        l_strSQL = l_strSQL & " AND START_DATE <= " & dbFormat(!EFF_END_DATE)
                    End If
                End If
                l_strSQL = l_strSQL & " AND END_DATE >= " & dbFormat(!EFF_START_DATE)
                l_strSQL = l_strSQL & " AND WG_NUM <> " & !WG_NUM & ")"
                
                .MoveNext
            Loop
        End With
        l_strSQL = l_strSQL & " ORDER BY EMP_NUM, WG_NUM, LINE_NUM"

        l_objrs.CursorLocation = adUseClient
        l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
        l_objrs.ActiveConnection = Nothing
        l_objCN.Close
    
        Set l_objCN = Nothing
    Else
        Set l_objrs = Nothing
    End If
    
    Set GetAssignmentsInOtherWG = l_objrs

    Set l_objrs = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing

    TrsRaiseError m_strCMODULENAME, "GetAssignmentsInOtherWG", , , , DescArgsFrom(l_strSQL), True
End Function


'//-----retrieve overlaps, need to match layout of assignment request-----//
'//v_varEmpNum is a collection
Public Function GetOverlapInScope(ByVal v_colScopeParm As Collection) As ADODB.Recordset
    Dim l_objCN As ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strScope As String
    Dim l_tScopeParm As Variant '//tScopeParm
    Dim v As Variant
    Dim l_lngCount As Long
    
    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString(m_strConnectStringReg)
    
    Set l_objCM = Nothing

    l_objCN.Open

    '//-----check for overlaps in other WG with EmpNum (1/25/2002)-----//
            
'SELECT WG_NUM, EMP_NUM, LINE_NUM, START_DATE, END_DATE, 'L' AS RECORD_STATUS, WG_NUM AS KEY_WG_NUM, EMP_NUM AS KEY_EMP_NUM, LINE_NUM AS KEY_LINE_NUM, START_DATE AS KEY_START_DATE, END_DATE AS KEY_END_DATE
'From METHDBA.TTR_LINE_ASGNM
'Where START_DATE <= '1/01/2002' AND
' END_DATE >= '12/29/2001' AND
'( (WG_NUM = 3043 AND
' LINE_NUM BETWEEN 1 AND 1) OR
'EMP_NUM IN (16689))
'ORDER BY WG_NUM, START_DATE, LINE_NUM

    l_strSQL = "SELECT WG_NUM, EMP_NUM, LINE_NUM, START_DATE, END_DATE, RELIEF_COMP"
    l_strSQL = l_strSQL & ", WG_NUM AS KEY_WG_NUM, EMP_NUM AS KEY_EMP_NUM, LINE_NUM AS KEY_LINE_NUM"
    l_strSQL = l_strSQL & ", START_DATE AS KEY_START_DATE, END_DATE AS KEY_END_DATE"
    l_strSQL = l_strSQL & ", CREATED_EMP_NUM, LAST_MOD_EMP_NUM"
    l_strSQL = l_strSQL & ", CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP"
    l_strSQL = l_strSQL & ", 'L' AS RECORD_STATUS"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_LINE_ASGNM"
    l_strSQL = l_strSQL & " WHERE "
    
    l_lngCount = 1
    For Each l_tScopeParm In v_colScopeParm
'        l_tScopeParm = v '//-----type cast from variant to tScopeParm
        l_strScope = "START_DATE <= " & dbFormat(l_tScopeParm(lngMaxEndDate))
        l_strScope = l_strScope & " AND END_DATE >= " & dbFormat(l_tScopeParm(lngMinStartDate))
        
        If l_tScopeParm(lngEmpNum) > 0 Then
            l_strScope = l_strScope & " AND EMP_NUM = " & l_tScopeParm(lngEmpNum)
        Else
            l_strScope = l_strScope & " AND WG_NUM = " & l_tScopeParm(lngWgNum)
            l_strScope = l_strScope & " AND LINE_NUM = " & l_tScopeParm(lngLineNum)
        End If

        l_strSQL = l_strSQL & " (" & l_strScope & ")"
        If l_lngCount < v_colScopeParm.Count Then '//-----not the last one-----//
            l_strSQL = l_strSQL & " OR"
        End If
        l_lngCount = l_lngCount + 1
    Next
    
    l_strSQL = l_strSQL & " ORDER BY WG_NUM, LINE_NUM, START_DATE"

    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetOverlapInScope = l_objrs

    Set l_objrs = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing

    TrsRaiseError m_strCMODULENAME, "GetOverlapInScope", , , , DescArgsFrom(l_strSQL), True
End Function


'//-----Make sure that a line exists (ttr_wg_line_sched) for each day in date range-----//
Public Function GetLines(ByVal v_lngWgNum As Long, ByVal v_datStartDate As Date, _
                ByVal v_datEndDate As Date, ByVal v_lngLine As Long) As ADODB.Recordset
        
    Dim l_objCN As ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString(m_strConnectStringReg)
    
    Set l_objCM = Nothing

    l_objCN.Open

'SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE
'From METHDBA.TTR_WG_LINE_SCHED
'WHERE WG_NUM=4026 AND
' LINE_NUM=1 AND
'START_DATE <= '11/30/2001' AND
'END_DATE >= '11/15/2001'

    '//-----retrieve all rows that overlap date range-----//
    l_strSQL = "SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_WG_LINE_SCHED"
    l_strSQL = l_strSQL & " WHERE WG_NUM = " & v_lngWgNum
    l_strSQL = l_strSQL & " AND LINE_NUM = " & v_lngLine
    l_strSQL = l_strSQL & " AND START_DATE <= " & dbFormat(v_datEndDate)
    l_strSQL = l_strSQL & " AND END_DATE >= " & dbFormat(v_datStartDate)
    l_strSQL = l_strSQL & " ORDER BY START_DATE ASC"
        
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing
    
    Set GetLines = l_objrs

    Set l_objrs = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetLines", , , , DescArgsFrom(l_strSQL), True
End Function

'//-----Make sure that a line exists (ttr_wg_line_sched) for each day in date range-----//
Public Function GetLinesBulk(ByVal v_colScopeParm As Collection) As ADODB.Recordset
    Dim l_objCN As ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strParm As String
    Dim l_tScopeParm As Variant '//tScopeParm
    Dim l_lngCount As Long
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString(m_strConnectStringReg)
    
    Set l_objCM = Nothing

    l_objCN.Open

'SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE
'From METHDBA.TTR_WG_LINE_SCHED
'WHERE WG_NUM=4026 AND
' LINE_NUM=1 AND
'START_DATE <= '11/30/2001' AND
'END_DATE >= '11/15/2001'

    '//-----retrieve all rows that overlap date range-----//
    l_strSQL = "SELECT WG_NUM, LINE_NUM, START_DATE, END_DATE"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_WG_LINE_SCHED"
    l_strSQL = l_strSQL & " WHERE"
'    l_strSQL = l_strSQL & " WHERE WG_NUM = " & v_lngWGNum
'    l_strSQL = l_strSQL & " AND LINE_NUM = " & v_lngLine
'    l_strSQL = l_strSQL & " AND START_DATE <= " & dbFormat(v_datEndDate)
'    l_strSQL = l_strSQL & " AND END_DATE >= " & dbFormat(v_datStartDate)
    l_lngCount = 0
    For Each l_tScopeParm In v_colScopeParm
'        l_tScopeParm = v '//-----type cast from variant to tScopeParm

        If l_tScopeParm(lngLineNum) > 0 Then '//-----only need to process the parms with wg/line-----//
            If l_lngCount > 0 Then  '//-----before each line except the first-----//
                l_strSQL = l_strSQL & " OR"
            End If
        
            l_strParm = "WG_NUM=" & l_tScopeParm(lngWgNum)
            l_strParm = l_strParm & " AND LINE_NUM=" & l_tScopeParm(lngLineNum)
            l_strParm = l_strParm & " AND START_DATE<=" & dbFormat(l_tScopeParm(lngMaxEndDate))
            l_strParm = l_strParm & " AND END_DATE>=" & dbFormat(l_tScopeParm(lngMinStartDate))

            l_strSQL = l_strSQL & " (" & l_strParm & ")"
            l_lngCount = l_lngCount + 1
        End If
    Next
    
    l_strSQL = l_strSQL & " ORDER BY WG_NUM, LINE_NUM, START_DATE"
        
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
    
    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing
    
    Set GetLinesBulk = l_objrs

    Set l_objrs = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetLinesBulk", , , , DescArgsFrom(l_strSQL), True
End Function

'//-----White162, employee is effective in WorkGroup for at least 1 day in date range (summary)
Public Function GetEmployeesInWorkGroup(ByVal v_lngWgNum As Long, ByVal v_datStartDate As Date, ByVal v_datEndDate As Date) As ADODB.Recordset
    Dim l_objCN As ADODB.Connection
    Dim l_objrs As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

    On Error GoTo eh

    Set l_objCN = New ADODB.Connection
    Set l_objrs = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr

    l_objCN.ConnectionString = l_objCM.GetConnectionString(m_strConnectStringReg)
    
    Set l_objCM = Nothing

    l_objCN.Open

    l_strSQL = "SELECT DISTINCT LAS.EMP_NUM, E.EMP_NAME"
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_LINE_ASGNM LAS"
    l_strSQL = l_strSQL & " LEFT OUTER JOIN  "
    l_strSQL = l_strSQL & " METHDBA.TTR_EMPLOYEE E "
    l_strSQL = l_strSQL & " ON LAS.EMP_NUM = E.EMP_NUM "
    l_strSQL = l_strSQL & " WHERE WG_NUM = " & v_lngWgNum
    l_strSQL = l_strSQL & " AND START_DATE <= " & dbFormat(v_datEndDate)
    l_strSQL = l_strSQL & " AND END_DATE >= " & dbFormat(v_datStartDate)
    l_strSQL = l_strSQL & " ORDER BY E.EMP_NAME"
            
'    l_strSQL = "SELECT DISTINCT EWG.EMP_NUM, E.EMP_NAME, EWG.WG_NUM, E.EMP_STATUS, E.UNION_CD "
'    l_strSQL = l_strSQL & " FROM (SELECT EMP_NUM, WG_NUM "
'    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EMP_WG "
'    l_strSQL = l_strSQL & " WHERE WG_NUM = " & v_lngWGNum & " AND "
'    l_strSQL = l_strSQL & " EFF_START_DATE <= " & dbFormat(v_datEndDate) & "  AND "
'    l_strSQL = l_strSQL & " EFF_END_DATE >= " & dbFormat(v_datStartDate) & "  ) AS EWG "
'    l_strSQL = l_strSQL & " LEFT OUTER JOIN  (select emp_num ,emp_name ,dept_num ,emp_status, UNION_CD "
'    l_strSQL = l_strSQL & " FROM METHDBA.TTR_EMPLOYEE  ) AS E "
'    l_strSQL = l_strSQL & " ON EWG.emp_num = E.emp_num "
'    l_strSQL = l_strSQL & " ORDER BY E.EMP_NAME"
    
    l_objrs.CursorLocation = adUseClient
    l_objrs.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objrs.ActiveConnection = Nothing
    l_objCN.Close

    Set l_objCN = Nothing

    Set GetEmployeesInWorkGroup = l_objrs

    Set l_objrs = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objrs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroup", , , , DescArgsFrom(l_strSQL), True
End Function


Private Function dbFormat(ByVal v_varData As Variant) As String
    Dim l_strFormated As String, fieldType As Integer
    
    On Error GoTo eh
    
    fieldType = VarType(v_varData)
    Select Case fieldType
        Case vbString
            l_strFormated = "'" & Replace(v_varData, "'", "''") & "'"
        Case vbInteger, vbLong
            l_strFormated = CStr(v_varData)
        Case vbDate
            l_strFormated = "'" & Format(v_varData, "mm/dd/yyyy") & "'"
        Case Else
            l_strFormated = Format(v_varData)
    End Select
    dbFormat = l_strFormated
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "dbFormat", , , , DescArgsFrom(CStr(v_varData)), True
End Function
'BEGIN PREM TRIS CONVERSION PHASE 3
Public Function GetDemandJobPositions(ByVal vlWorkgroup As Long _
        , Optional ByVal vdtStartDate As Date = #1/1/1980# _
        , Optional ByVal vdtEndDate As Date = #12/31/9999# _
        ) As ADODB.Recordset
On Error GoTo GetDemandJobPositions_EH
Dim lobjCNManager As util_DBConnectMgr.clsuDBConnectMgr
Dim lsSQL As String, lsConnection As String
Dim ladcn As ADODB.Connection
Dim ladrs As ADODB.Recordset
Dim ladcmd As ADODB.Command
Dim ladParm As ADODB.Parameter

    lsSQL = "SELECT WG_NUM, LINE_NUM" & _
        "       , CHARGE_DATE, JOB_POSITION" & _
        "       , LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP" & _
        " FROM  METHDBA.TTR_DEMAND_JOB_POS" & _
        " WHERE WG_NUM = ?" & _
        " AND   CHARGE_DATE BETWEEN ?" & _
        " AND   ?" & _
        " ORDER BY LINE_NUM, CHARGE_DATE"
        
    Set ladcmd = New ADODB.Command
    With ladcmd
        .CommandText = lsSQL
        .CommandType = adCmdText

        Set ladParm = .CreateParameter("" _
                        , adInteger _
                        , adParamInput _
                        , 0 _
                        , vlWorkgroup)
        .Parameters.Append ladParm

        Set ladParm = .CreateParameter("" _
                        , adDate _
                        , adParamInput _
                        , 0 _
                        , vdtStartDate)
        .Parameters.Append ladParm

        Set ladParm = .CreateParameter("" _
                        , adDate _
                        , adParamInput _
                        , 0 _
                        , vdtEndDate)
        .Parameters.Append ladParm
    End With

    Set lobjCNManager = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
    lsConnection = lobjCNManager.GetConnectionString
    Set lobjCNManager = Nothing

    Set ladcn = New ADODB.Connection
    With ladcn
        .ConnectionString = lsConnection
        .CursorLocation = adUseClient
        .Open
    End With

    Set ladcmd.ActiveConnection = ladcn

    Set ladrs = New ADODB.Recordset
    With ladrs
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly

        .Open ladcmd
        .ActiveConnection = Nothing
    End With

    Set GetDemandJobPositions = ladrs

    Set ladrs = Nothing
    Set ladParm = Nothing
    Set ladcmd = Nothing
    Set ladcn = Nothing
    Set lobjCNManager = Nothing

    Exit Function
GetDemandJobPositions_EH:
    Set ladrs = Nothing
    Set ladParm = Nothing
    Set ladcmd = Nothing
    Set ladcn = Nothing
    Set lobjCNManager = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetDemandJobPositions" _
        , trErrdbGetDemandJobPositions _
        , , , DescArgsFrom(lsSQL), True

End Function
'END PREM TRIS CONVERSION PHASE 3
