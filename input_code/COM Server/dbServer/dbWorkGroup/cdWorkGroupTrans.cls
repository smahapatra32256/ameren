VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdWorkGroupTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdWorkGroupTrans"

' Static data constants
Private Const CRECORDNOTFOUND = "RECORDNOTFOUND"
Private Const CDEFAULTDATE As String = "01/01/0100"

' Overtime exception code constants
Private Const CWG_GROUPNAME = "WORKGROUPCACHE"
' index for record dynamic array
Private Const CWG_WG_NUM = 0
Private Const CWG_WG_DESC = 1
Private Const CWG_EFF_START_DATE = 2
Private Const CWG_EFF_END_DATE = 3
Private Const CWG_PROJECT_ID = 4
Private Const CWG_DEMAND_IND = 5
Private Const CWG_WG_SUPERVISOR = 6
Private Const CWG_CREW_NUM = 7
Private Const CWG_DEPT_NUM = 8
Private Const CWG_LOC_CD = 9
Private Const CWG_PROJ_START_DATE = 10
Private Const CWG_PROJ_END_DATE = 11
Private Const CWG_DOJM_DIST = 12
Private Const CWG_DOJM_HQ = 13
Private Const CWG_LAST_SLOT = 13

Private maryFieldNames As Variant

'*********************************************************
' Begin Public procedures
'*********************************************************

Public Function SaveWGAssignments( _
        ByVal v_rsWGAssignments As ADODB.Recordset, _
        ByVal v_rsAdjPrimaries As ADODB.Recordset _
                    ) As Boolean
 
       
    Dim l_objCN As New ADODB.Connection
    Dim l_strSQL As String
    Dim l_strTimestamp As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_colErrorDesc As Collection
    Dim l_intI As Integer
    Dim l_lngRecsAffected As Long
    Dim l_lngWgNum As Long
    Dim l_datStartDate As String
    Dim l_lngEmpNum As Long
    
'This Function Saves wg assiggment data in 2 recordsets
' v_rsWGAssignments -- Assigments within a work group
' v_rsAdjPrimaries  -- These are primary assignments in a different
'                       work group that were adjusted because of the
'                       request
' Deletes need to be processed before updates and inserts

    On Error GoTo SaveWGAssignments_EH
    
    
    If RecordsExist(v_rsWGAssignments) Or _
        RecordsExist(v_rsAdjPrimaries) Then
    Else
        GetObjectContext.SetComplete
        SaveWGAssignments = True
        Exit Function
    End If
    
' This code below is used to eliminate the need to get a
' connection if no changes have been made
    v_rsWGAssignments.Filter = adFilterPendingRecords
    v_rsAdjPrimaries.Filter = adFilterPendingRecords
    If RecordsExist(v_rsWGAssignments) Or _
        RecordsExist(v_rsAdjPrimaries) Then
    Else
        GetObjectContext.SetComplete
        SaveWGAssignments = True
        Exit Function
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
'Saving WG Assignment Data
'Process deletes before updates and inserts
    If RecordsExist(v_rsWGAssignments) Then
        For l_intI = 1 To 2
            v_rsWGAssignments.MoveFirst
            While Not v_rsWGAssignments.EOF
                l_lngWgNum = AscertainKeyValue(v_rsWGAssignments, "WG_NUM")
                l_datStartDate = AscertainKeyValue(v_rsWGAssignments, "EFF_START_DATE")
                l_lngEmpNum = AscertainKeyValue(v_rsWGAssignments, "EMP_NUM")
                Select Case l_intI
                    Case 1
                        If v_rsWGAssignments.Status = adRecDeleted Then
                            l_strSQL = SaveWgAssignGetSQL(v_rsWGAssignments)
                        Else
                            l_strSQL = ""
                        End If
                    Case 2
                        If v_rsWGAssignments.Status <> adRecDeleted Then
                            l_strSQL = SaveWgAssignGetSQL(v_rsWGAssignments)
                        Else
                            l_strSQL = ""
                        End If
                End Select
                If l_strSQL <> "" Then
                    l_objCN.Execute l_strSQL, l_lngRecsAffected
                    If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                         Err.Raise trErrAppCommitConflict
                    End If
                End If
                v_rsWGAssignments.MoveNext
            Wend
        Next l_intI
    End If
    
'Saving Adjacent Primary WG Assignment Data
    If RecordsExist(v_rsAdjPrimaries) Then
        For l_intI = 1 To 2
            v_rsAdjPrimaries.MoveFirst
            While Not v_rsAdjPrimaries.EOF
                l_lngWgNum = AscertainKeyValue(v_rsAdjPrimaries, "WG_NUM")
                l_datStartDate = AscertainKeyValue(v_rsAdjPrimaries, "EFF_START_DATE")
                l_lngEmpNum = AscertainKeyValue(v_rsAdjPrimaries, "EMP_NUM")
                Select Case l_intI
                    Case 1
                        If v_rsAdjPrimaries.Status = adRecDeleted Then
                            l_strSQL = SaveWgAssignGetSQL(v_rsAdjPrimaries)
                        Else
                            l_strSQL = ""
                        End If
                    Case 2
                        If v_rsAdjPrimaries.Status <> adRecDeleted Then
                            l_strSQL = SaveWgAssignGetSQL(v_rsAdjPrimaries)
                        Else
                            l_strSQL = ""
                        End If
                End Select
                If l_strSQL <> "" Then
                    l_objCN.Execute l_strSQL, l_lngRecsAffected
                    If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                         Err.Raise trErrAppCommitConflict
                    End If
                End If
                v_rsAdjPrimaries.MoveNext
            Wend
        Next l_intI
    End If
            
    GetObjectContext.SetComplete
    l_objCN.Close
    Set l_objCN = Nothing
    SaveWGAssignments = True
    
    Exit Function

SaveWGAssignments_EH:

    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = Join( _
                            Array( _
                                CStr(TrType_WGM_ASGMT), _
                                CStr(l_lngWgNum), _
                                Format(l_datStartDate, "mm/dd/yyyy"), _
                                "0", _
                                CStr(l_lngEmpNum), _
                                "", _
                                "") _
                            , ";")
            TrsRaiseError m_strCMODULENAME, "SaveWGAssignments", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveWGAssignments", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function


Private Function AscertainKeyValue( _
                    ByVal v_rsDBRequest As ADODB.Recordset, _
                    ByVal v_FieldName As String) As Variant

On Error GoTo eh

    If v_rsDBRequest.Status = adRecDeleted Or v_rsDBRequest.Status = adRecModified Then
        AscertainKeyValue = v_rsDBRequest.Fields(v_FieldName).OriginalValue
    ElseIf v_rsDBRequest.Status = adRecNew Then
        AscertainKeyValue = v_rsDBRequest.Fields(v_FieldName).Value
    Else
        AscertainKeyValue = 0
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "AscertainKeyValue"
End Function


Private Function SaveWgAssignGetSQL( _
            ByVal v_rsWGAssign As ADODB.Recordset _
            ) As String
            
    Dim l_aryPrimaryKey(4) As String
    Dim l_varIgnore(2) As String
    Dim l_strSQL As String
'    Dim l_strTimeStamp As String
            
    l_strSQL = ""
    l_aryPrimaryKey(0) = "EMP_NUM"
    l_aryPrimaryKey(1) = "WG_NUM"
    l_aryPrimaryKey(2) = "PRIMARY_IND"
    l_aryPrimaryKey(3) = "EFF_START_DATE"
    l_aryPrimaryKey(4) = "LAST_MOD_TIMESTAMP"
    l_varIgnore(0) = "CREATED_TIMESTAMP"
    l_varIgnore(1) = "LAST_MOD_TIMESTAMP"

    Select Case v_rsWGAssign.Status
       Case adRecModified
'           l_strTimeStamp = GetTimeStamp()
'           v_rsWGAssign!TIME_STAMP = l_strTimeStamp
           l_strSQL = UpdateStatement(v_rsWGAssign, "METHDBA.TTR_EMP_WG", l_aryPrimaryKey, l_varIgnore, True)
       Case adRecNew
           l_strSQL = InsertStatement(v_rsWGAssign, "METHDBA.TTR_EMP_WG", l_varIgnore)
       Case adRecDeleted
           l_strSQL = DeleteStatement(v_rsWGAssign, "METHDBA.TTR_EMP_WG", l_aryPrimaryKey)
        Case Else
           l_strSQL = ""
    End Select

    SaveWgAssignGetSQL = l_strSQL
            
End Function


Public Function SaveWGLineSchedules( _
            ByRef r_rsWGSchedules As ADODB.Recordset _
        ) As Boolean
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long
    Dim l_lngWgNum As Long
    Dim l_datStartDate As String
    Dim l_lngLineNum As Long
    
On Error GoTo eh

    r_rsWGSchedules.Filter = adFilterPendingRecords
    If r_rsWGSchedules.EOF Then
        SaveWGLineSchedules = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    Set l_objCM = Nothing
    
    Do Until r_rsWGSchedules.EOF
        l_lngWgNum = AscertainKeyValue(r_rsWGSchedules, "WG_NUM")
        l_datStartDate = AscertainKeyValue(r_rsWGSchedules, "START_DATE")
        l_lngLineNum = AscertainKeyValue(r_rsWGSchedules, "LINE_NUM")
        l_strSQL = SaveLineScheduleSQL(r_rsWGSchedules)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        r_rsWGSchedules.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    SaveWGLineSchedules = True
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = Join( _
                            Array( _
                                CStr(TrType_WGM_SCHED), _
                                CStr(l_lngWgNum), _
                                Format(l_datStartDate, "mm/dd/yyyy"), _
                                CStr(l_lngLineNum)) _
                            , ";")
            TrsRaiseError m_strCMODULENAME, "SaveWGLineSchedules", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveWGLineSchedules", , , , DescArgsFrom(l_strSQL), True
    End Select
End Function


Private Function SaveLineScheduleSQL( _
                r_rsWGLineSchedules As ADODB.Recordset) As String
    
    Const cLINESCHEDTABLE As String = "METHDBA.TTR_WG_LINE_SCHED"
    Dim primKeys As Variant
'    Dim l_strTimeStamp As String
    Dim l_strSQL As String
    Dim l_varIgnore As Variant
    
On Error GoTo eh
    l_varIgnore = Array("CREATED_TIMESTAMP", "LAST_MOD_TIMESTAMP")
    primKeys = Array("WG_NUM", "LINE_NUM", "START_DATE", "CREATED_TIMESTAMP", "LAST_MOD_TIMESTAMP")
    Select Case r_rsWGLineSchedules.Status
       Case adRecModified
       '
       ' Replace this with CURRENT TIMESTAMP
       '
'           l_strTimeStamp = GetTimeStamp()
'           r_rsWGLineSchedules!LAST_MOD_TIMESTAMP = l_strTimeStamp
           l_strSQL = UpdateStatement(r_rsWGLineSchedules, cLINESCHEDTABLE, primKeys, , True)
       Case adRecNew
           l_strSQL = InsertStatement(r_rsWGLineSchedules, cLINESCHEDTABLE, l_varIgnore)
       Case adRecDeleted
           l_strSQL = DeleteStatement(r_rsWGLineSchedules, cLINESCHEDTABLE, primKeys)
        Case Else
           l_strSQL = ""
    End Select
    SaveLineScheduleSQL = l_strSQL
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "SaveLineScheduleSQL"
End Function


Public Function SaveWGBaseSchedules( _
            ByRef r_rsWGSchedules As ADODB.Recordset _
        ) As Boolean
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long
    Dim l_lngWgNum As Long
    Dim l_datEffDate As String
    
On Error GoTo eh

    r_rsWGSchedules.Filter = adFilterPendingRecords
    If r_rsWGSchedules.EOF Then
        SaveWGBaseSchedules = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    Set l_objCM = Nothing
    
    Do Until r_rsWGSchedules.EOF
        l_lngWgNum = AscertainKeyValue(r_rsWGSchedules, "WG_NUM")
        l_datEffDate = AscertainKeyValue(r_rsWGSchedules, "EFF_DATE")
        l_strSQL = SaveBaseScheduleSQL(r_rsWGSchedules)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        r_rsWGSchedules.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    SaveWGBaseSchedules = True
    Exit Function
        
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = Join( _
                            Array( _
                                CStr(TrType_WGM_SCHED), _
                                CStr(l_lngWgNum), _
                                Format(l_datEffDate, "mm/dd/yyyy"), _
                                "0") _
                            , ";")
            TrsRaiseError m_strCMODULENAME, "SaveWGBaseSchedules", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveWGBaseSchedules", , , , DescArgsFrom(l_strSQL), True
    End Select
    
End Function


Private Function SaveBaseScheduleSQL( _
                r_rsWGSchedules As ADODB.Recordset) As String
    
    Const CSCHEDTABLE As String = "METHDBA.TTR_WG_BASE_SCHED"
    Dim primKeys As Variant
'    Dim l_strTimeStamp As String
    Dim l_strSQL As String
    Dim l_varIgnore As Variant
    
On Error GoTo eh
    l_varIgnore = Array("CREATED_TIMESTAMP", "LAST_MOD_TIMESTAMP")
    primKeys = Array("WG_NUM", "EFF_DATE", "CREATED_TIMESTAMP", "LAST_MOD_TIMESTAMP")
    Select Case r_rsWGSchedules.Status
       Case adRecModified
       '
       ' Replace this with CURRENT TIMESTAMP
       '
'           l_strTimeStamp = GetTimeStamp()
'           r_rsWGSchedules!LAST_MOD_TIMESTAMP = l_strTimeStamp
           l_strSQL = UpdateStatement(r_rsWGSchedules, CSCHEDTABLE, primKeys, , True)
       Case adRecNew
           l_strSQL = InsertStatement(r_rsWGSchedules, CSCHEDTABLE, l_varIgnore)
       Case adRecDeleted
           l_strSQL = DeleteStatement(r_rsWGSchedules, CSCHEDTABLE, primKeys)
        Case Else
           l_strSQL = ""
    End Select
    SaveBaseScheduleSQL = l_strSQL
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "SaveBaseScheduleSQL"
End Function


Public Function SaveWGAttributes( _
        ByVal v_rsWGAttributes As ADODB.Recordset, _
        ByVal v_strUserId As String _
                    ) As Boolean
 
 
    Dim l_objCN As New ADODB.Connection
    Dim l_strSQL As String
    Dim l_strDeleteAssociationsSQL As String
    Dim l_strAssociationsSQL As String
    Dim l_strDefaultScheduleSQL As String
    Dim l_strApprovalSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_intI As Integer
    Dim l_lngRecsAffected As Long
    Dim l_lngWgNum As Long
    Dim l_strDeleteApprovalSQL As String
    Dim l_strInsertApprovalSQL As String
    Dim l_lngRowsAffected As Long

'This Function Saves wg attributes data in 2 tables
' but only one recordset
' ttr_work_group    -- This is the Work Group Master table
'                       It contains work group number, description
'                        supervisor, project, start date, and end date
' ttr_associations  -- This table associates a work group
'                       to a department and a location

    On Error GoTo SaveWGAttributes_EH
    l_strSQL = ""

    If RecordsExist(v_rsWGAttributes) Then
    Else
        GetObjectContext.SetComplete
        SaveWGAttributes = True
        Exit Function
    End If
    
' This code below is used to eliminate the need to get a
' connection if no changes have been made
    v_rsWGAttributes.Filter = adFilterPendingRecords
    If RecordsExist(v_rsWGAttributes) Then
    Else
        GetObjectContext.SetComplete
        SaveWGAttributes = True
        Exit Function
    End If

    Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    

    v_rsWGAttributes.MoveFirst
    While Not v_rsWGAttributes.EOF
        l_lngWgNum = AscertainKeyValue(v_rsWGAttributes, "WG_NUM")
        l_strSQL = SaveWgAttributesGetSQL(v_rsWGAttributes)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
      
        l_strAssociationsSQL = SaveWGAttGetAssociationsSQL(v_rsWGAttributes, v_strUserId, l_strDeleteAssociationsSQL)
        l_strApprovalSQL = SaveWGAttGetApprovalSQL(v_rsWGAttributes, v_strUserId, l_strDeleteApprovalSQL, l_strInsertApprovalSQL)
        l_strDefaultScheduleSQL = SaveWGAttGetSchedSQL(v_rsWGAttributes, v_strUserId)
        If l_strAssociationsSQL <> "" Then
            If l_strDeleteAssociationsSQL <> "" Then
                l_strSQL = l_strDeleteAssociationsSQL
                l_objCN.Execute l_strSQL
            End If
            l_strSQL = l_strAssociationsSQL
            l_objCN.Execute l_strSQL
        End If
        If l_strApprovalSQL <> "" Then
            If l_strDeleteApprovalSQL <> "" Then
                l_strSQL = l_strDeleteApprovalSQL
                l_objCN.Execute l_strSQL
            End If
            l_strSQL = l_strApprovalSQL
            l_objCN.Execute l_strSQL, l_lngRowsAffected
            If l_strInsertApprovalSQL <> "" Then
                If l_lngRowsAffected > 0 Then
                Else
                    l_strSQL = l_strInsertApprovalSQL
                    l_objCN.Execute l_strSQL
                End If
            End If
        End If
        If l_strDefaultScheduleSQL <> "" Then
            l_strSQL = l_strDefaultScheduleSQL
            l_objCN.Execute l_strSQL
        End If
        v_rsWGAttributes.MoveNext
    Wend
            
    GetObjectContext.SetComplete
    l_objCN.Close
    Set l_objCN = Nothing
    SaveWGAttributes = True
    
    Exit Function

SaveWGAttributes_EH:

    GetObjectContext.SetAbort
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = Join( _
                            Array( _
                                CStr(TrType_WGM), _
                                CStr(l_lngWgNum), _
                                "", _
                                "0", _
                                "0", _
                                "", _
                                "") _
                            , ";")
            TrsRaiseError m_strCMODULENAME, "SaveWGAttributes", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveWGAttributes", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function

Private Function SaveWgAttributesGetSQL( _
            ByVal v_rsWGAttributes As ADODB.Recordset _
            ) As String
            
    Dim l_aryPrimaryKey(1) As String
    Dim l_varIgnore(3) As String
    Dim l_strSQL As String
    Dim l_strTimestamp As String
            
    l_strSQL = ""
    l_aryPrimaryKey(0) = "WG_NUM"
    l_aryPrimaryKey(1) = "LAST_MOD_TIMESTAMP"
'CREATED_EMP_NUM is only set on an insert
'The business process will control this
'    l_varIgnore(0) = "CREATED_EMP_NUM"
    l_varIgnore(0) = "CREATED_TIMESTAMP"
    l_varIgnore(1) = "LAST_MOD_TIMESTAMP"
    l_varIgnore(2) = "DEPT_NUM"
    l_varIgnore(3) = "LOC_CD"

    Select Case v_rsWGAttributes.Status
       Case adRecModified
           l_strSQL = UpdateStatement(v_rsWGAttributes, "METHDBA.TTR_WORK_GROUP", l_aryPrimaryKey, l_varIgnore, True)
       Case adRecNew
           l_strSQL = InsertStatement(v_rsWGAttributes, "METHDBA.TTR_WORK_GROUP", l_varIgnore)
       Case adRecDeleted
           l_strSQL = DeleteStatement(v_rsWGAttributes, "METHDBA.TTR_WORK_GROUP", l_aryPrimaryKey)
        Case Else
           l_strSQL = ""
    End Select

    SaveWgAttributesGetSQL = l_strSQL
            
End Function


Private Function AssociatonChanged( _
    ByVal v_rsWGAttributes As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strDept As String
    Dim l_lngLocCd As Long
    Dim l_strOrigDept As String
    Dim l_lngOrigLocCd As Long
    
    l_strDept = v_rsWGAttributes.Fields("DEPT_NUM").Value
    l_strOrigDept = v_rsWGAttributes.Fields("DEPT_NUM").OriginalValue
    l_lngLocCd = v_rsWGAttributes.Fields("LOC_CD").Value
    l_lngOrigLocCd = v_rsWGAttributes.Fields("LOC_CD").OriginalValue
    
    If l_strDept = l_strOrigDept And _
        l_lngLocCd = l_lngOrigLocCd Then
        AssociatonChanged = False
    Else
        AssociatonChanged = True
    End If
        
    
    
End Function



Private Function SaveWGAttGetAssociationsSQL( _
    ByVal v_rsWGAttributes As ADODB.Recordset, _
    ByVal v_strUserId As String, _
    ByRef r_strDeleteSQL As String _
    ) As String
    
    Dim l_strInsertStmt As String
    Dim l_strDeleteStmt As String
    Dim l_strAssociationsSQL As String
    
    l_strInsertStmt = "INSERT INTO METHDBA.TTR_ASSOCIATIONS " & _
                " Select DISTINCT DIVISION_NUM, FUNCTION_NUM, DEPT_NUM , LOC_CD," & _
                str(v_rsWGAttributes!WG_NUM) & ", ' ', CURRENT TIMESTAMP " & _
                "From METHDBA.TTR_ASSOCIATIONS " & _
                "Where LOC_CD = " & str(v_rsWGAttributes!LOC_CD) & _
                "  and DEPT_NUM = '" & v_rsWGAttributes!DEPT_NUM & "' " & _
                "  and wg_num = 0 "
    
     l_strDeleteStmt = "DELETE FROM METHDBA.TTR_ASSOCIATIONS " & _
                    "WHERE WG_NUM = " & str(v_rsWGAttributes!WG_NUM)
                                                                                
    r_strDeleteSQL = ""
    Select Case v_rsWGAttributes.Status
        Case adRecNew
            l_strAssociationsSQL = l_strInsertStmt
        Case adRecModified
            If AssociatonChanged(v_rsWGAttributes) Then
                l_strAssociationsSQL = l_strInsertStmt
                r_strDeleteSQL = l_strDeleteStmt
            Else
                l_strAssociationsSQL = ""
            End If
            
        Case Else
            l_strAssociationsSQL = ""
    End Select
    
    SaveWGAttGetAssociationsSQL = l_strAssociationsSQL
    
End Function

Private Function SaveWGAttGetApprovalSQL( _
    ByVal v_rsWGAttributes As ADODB.Recordset, _
    ByVal v_strUserId As String, _
    ByRef r_strDeleteSQL As String, _
    ByRef r_strInsertSQL As String _
    ) As String
    
'The r_strDeleteSQL and r_strInsertSQL are only set
' when the work group is updated.
'The r_strDeleteSQL will delete the new supervisor's approval authority
'  if they already have access to the work group
'The r_strInsertSQL is used to set the approval authority of the new supervisor.
' There a quite a few existing work groups where the supervisor is not
' in the ttr_approval table.  In this case the Update would not affect any rows
    
    Dim l_lngWGSupervisor As Long
    Dim l_lngOrigWGSupervisor As Long
    Dim l_lngWorkGroup As Long
    Dim l_strApprovalSQL As String
    Dim l_strInsertSQL As String
    
    r_strDeleteSQL = ""
    r_strInsertSQL = ""
    l_lngWorkGroup = v_rsWGAttributes.Fields("WG_NUM").Value
    l_lngWGSupervisor = v_rsWGAttributes.Fields("WG_SUPERVISOR").Value
    l_strInsertSQL = "INSERT INTO METHDBA.TTR_APPROVAL " & _
                "(EMP_NUM,APPR_LEVEL_IND,GROUP_ID,LAST_MOD_EMP_NUM,ACCESS_LEVEL) " & _
                " VALUES(" & str(l_lngWGSupervisor) & "," & _
                "50," & str(l_lngWorkGroup) & "," & v_strUserId & ",'A') "
    Select Case v_rsWGAttributes.Status
        Case adRecNew
            l_strApprovalSQL = l_strInsertSQL
                
        Case adRecModified
            l_lngOrigWGSupervisor = v_rsWGAttributes.Fields("WG_SUPERVISOR").OriginalValue
            If l_lngWGSupervisor <> l_lngOrigWGSupervisor Then
                r_strInsertSQL = l_strInsertSQL
                r_strDeleteSQL = "DELETE FROM METHDBA.TTR_APPROVAL " & _
                    "WHERE EMP_NUM = " & str(l_lngWGSupervisor) & _
                    " AND APPR_LEVEL_IND = 50 " & _
                    " AND GROUP_ID = " & str(l_lngWorkGroup)
                l_strApprovalSQL = "UPDATE METHDBA.TTR_APPROVAL " & _
                    "SET EMP_NUM = " & str(l_lngWGSupervisor) & _
                    ", LAST_MOD_EMP_NUM = " & v_strUserId & _
                    ", TIME_STAMP = CURRENT TIMESTAMP  " & _
                    "WHERE EMP_NUM = " & str(l_lngOrigWGSupervisor) & _
                    " AND APPR_LEVEL_IND = 50 " & _
                    " AND GROUP_ID = " & str(l_lngWorkGroup)
            Else
                l_strApprovalSQL = ""
            End If
            
        Case Else
            l_strApprovalSQL = ""
            
    End Select
    
    SaveWGAttGetApprovalSQL = l_strApprovalSQL
    
End Function


Private Function SaveWGAttGetSchedSQL( _
    ByVal v_rsWGAttributes As ADODB.Recordset, _
    ByVal v_strUserId As String)
    
        
    Dim l_strInsertStmt As String
    Dim l_strValuesStmt As String

    
    l_strInsertStmt = "INSERT INTO METHDBA.TTR_WG_BASE_SCHED " & _
                "(WG_NUM,EFF_DATE,ROT_CD,ROT_WEEK,SCHED_CD,CREATED_EMP_NUM,LAST_MOD_EMP_NUM,END_DATE) "
                                                           
    Select Case v_rsWGAttributes.Status
        Case adRecNew
            If v_rsWGAttributes!DEMAND_IND = "Y" Then
                l_strValuesStmt = " VALUES(" & str(v_rsWGAttributes!WG_NUM) & "," & _
                        "'" & Format(v_rsWGAttributes!EFF_START_DATE, "YYYY-MM-DD") & "'," & _
                        "0,0,-9," & v_strUserId & "," & v_strUserId & "," & _
                        "'9999-12-31')"
                SaveWGAttGetSchedSQL = l_strInsertStmt & l_strValuesStmt
            Else
                SaveWGAttGetSchedSQL = ""
            End If
        Case adRecModified
            SaveWGAttGetSchedSQL = ""
        Case Else
            SaveWGAttGetSchedSQL = ""
    End Select
    
    
End Function


Public Function SaveWGDistribution( _
            ByVal v_rsWGDistribution As ADODB.Recordset _
        ) As Boolean
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long
    Dim l_lngWgNum As Long
    Dim l_strDistId As String
    
On Error GoTo eh

    v_rsWGDistribution.Filter = adFilterPendingRecords
    If v_rsWGDistribution.EOF Then
        SaveWGDistribution = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    Set l_objCM = Nothing
    
    Do Until v_rsWGDistribution.EOF
        l_lngWgNum = AscertainKeyValue(v_rsWGDistribution, "WG_NUM")
        l_strDistId = AscertainKeyValue(v_rsWGDistribution, "DISTRIBUTION_ID")
        l_strSQL = SaveDistributionSQL(v_rsWGDistribution)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsWGDistribution.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    SaveWGDistribution = True
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort
   
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = Join( _
                            Array( _
                                CStr(TrType_WGM_ACCT), _
                                CStr(l_lngWgNum), _
                                "", _
                                "0", _
                                "0", _
                                l_strDistId) _
                            , ";")
            TrsRaiseError m_strCMODULENAME, "SaveWGDistribution", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveWGDistribution", , , , DescArgsFrom(l_strSQL), True
    End Select
End Function

Private Function SaveDistributionSQL( _
    ByVal v_rsWGDistribution As ADODB.Recordset) As String
    
    Const CSCHEDTABLE As String = "METHDBA.TTR_ACCT_ASSIGNMTS"
    Dim primKeys As Variant
    Dim l_strTimestamp As String
    Dim l_strSQL As String
    Dim l_varIgnore As Variant
'    Dim l_strSetForLogicalDelate As String
'    Dim l_strUpdatePattern As String
'    Dim l_strUpdateReplacePattern As String
'    Dim l_varUpdateRegExGroups As Variant
    Dim ll_WG_NUM As Long
    Dim ls_TIME_STAMP As String
    Dim ls_DISTRIBUTION_ID As String
On Error GoTo eh
    l_varIgnore = Array("TIME_STAMP")
    primKeys = Array("WG_NUM", "DISTRIBUTION_ID", "TIME_STAMP")
    Select Case v_rsWGDistribution.Status
       Case adRecModified
           l_strSQL = UpdateStatement(v_rsWGDistribution, CSCHEDTABLE, primKeys, , True, "TIME_STAMP")
       Case adRecNew
           l_strSQL = InsertStatement(v_rsWGDistribution, CSCHEDTABLE, l_varIgnore)
       Case adRecDeleted
           '//l_strSQL = DeleteStatement(v_rsWGDistribution, CSCHEDTABLE, primKeys)
           '//make it a logical delete
           If Not v_rsWGDistribution.EOF Then
            ll_WG_NUM = v_rsWGDistribution.Fields("WG_NUM").OriginalValue
            ls_TIME_STAMP = v_rsWGDistribution.Fields("TIME_STAMP").OriginalValue
            ls_DISTRIBUTION_ID = v_rsWGDistribution.Fields("DISTRIBUTION_ID").OriginalValue
           
            l_strSQL = "UPDATE METHDBA.TTR_ACCT_ASSIGNMTS " & vbCrLf & _
            " SET ACTIVE_STATUS_CODE = 'N' " & vbCrLf & _
            " WHERE WG_NUM = " & ll_WG_NUM & " AND TIME_STAMP = '" & ls_TIME_STAMP & "' AND DISTRIBUTION_ID = '" & ls_DISTRIBUTION_ID & "'"
           
            '//TEST ONLY
            '//TrsTraceError m_strCMODULENAME, "SaveDistributionSQL(Delete)", , False, , DescArgsFrom(l_strSQL), True
           End If
           '//l_strSQL = UpdateStatement(v_rsWGDistribution, CSCHEDTABLE, primKeys, Array("DISTRIBUTION_ID", "DESCRIPTION"), True, "TIME_STAMP")
           
           
'           l_strUpdatePattern = "((?<update>UPDATE\s*.*)\s*(?<set>SET\s*.*)\s*(?<where>WHERE\s*.*))"
'           l_strSQL = UpdateStatement(v_rsWGDistribution, CSCHEDTABLE, primKeys, , True, "TIME_STAMP")
'           If IsRegExMatch(l_strSQL, l_strUpdatePattern) Then
'                l_varUpdateRegExGroups = GetRegExGroups(l_strSQL, l_strUpdatePattern)
'           '//replace SET
'                l_strSetForLogicalDelate = " SET ACTIVE_STATUS_CODE = 'N' "
'                l_strUpdateReplacePattern = "${update}" & l_strSetForLogicalDelate & "${where}"
'
'                '//TEST ONLY
'                TrsTraceError m_strCMODULENAME, "SaveDistributionSQL(Delete before regex replace)", , False, , DescArgsFrom(l_strSQL), True
'
'                l_strSQL = RegExReplace(l_strSQL, l_strUpdatePattern, l_strUpdateReplacePattern)
'
'                '//TEST ONLY
'                TrsTraceError m_strCMODULENAME, "SaveDistributionSQL(Delete after regex replace)", , False, , DescArgsFrom(l_strSQL), True
'            Else
'                '//TEST ONLY
'                TrsTraceError m_strCMODULENAME, "SaveDistributionSQL(Delete)", , False, , DescArgsFrom("unable to parse update statement for logical delete", l_strSQL), True
'           End If
            
        Case Else
           l_strSQL = ""
    End Select
    SaveDistributionSQL = l_strSQL
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "SaveDistributionSQL"
End Function

Private Function SaveWGActivitiesSQL( _
    ByVal v_rsWGActivities As ADODB.Recordset) As String
    
    Const CSCHEDTABLE As String = "METHDBA.TTR_ACTIVITY_ASGNM"
    Dim primKeys As Variant
    Dim l_strTimestamp As String
    Dim l_strSQL As String
    Dim l_varIgnore As Variant
    
On Error GoTo eh
    l_varIgnore = Array("TIME_STAMP")
    primKeys = Array("ACTIVITY_CD", "WG_NUM", "TIME_STAMP")
    Select Case v_rsWGActivities.Status
       Case adRecModified
           l_strSQL = UpdateStatement(v_rsWGActivities, CSCHEDTABLE, primKeys, , True, "TIME_STAMP")
       Case adRecNew
           l_strSQL = InsertStatement(v_rsWGActivities, CSCHEDTABLE, l_varIgnore)
       Case adRecDeleted
           l_strSQL = DeleteStatement(v_rsWGActivities, CSCHEDTABLE, primKeys)
        Case Else
           l_strSQL = ""
    End Select
    SaveWGActivitiesSQL = l_strSQL
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "SaveWGActivitiesSQL"
End Function


Private Function SaveGLActivitiesSQL( _
    ByVal v_rsGLActivities As ADODB.Recordset) As String
    
    Const CSCHEDTABLE As String = "METHDBA.TTR_ACTIVITY"
    Dim primKeys As Variant
    Dim l_strTimestamp As String
    Dim l_strSQL As String
    Dim l_varIgnore As Variant
    
On Error GoTo eh
    l_varIgnore = Array("TIMESTAMP")
    primKeys = Array("ACTIVITY_CD", "TIMESTAMP")
    Select Case v_rsGLActivities.Status
       Case adRecModified
           l_strSQL = UpdateStatement(v_rsGLActivities, CSCHEDTABLE, primKeys, , True, "TIMESTAMP")
       Case adRecNew
           l_strSQL = InsertStatement(v_rsGLActivities, CSCHEDTABLE, l_varIgnore)
       Case adRecDeleted
           l_strSQL = DeleteStatement(v_rsGLActivities, CSCHEDTABLE, primKeys)
        Case Else
           l_strSQL = ""
    End Select
    SaveGLActivitiesSQL = l_strSQL
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "SaveGLActivitiesSQL"
End Function


Public Function SaveWGActivities( _
            ByVal v_rsWGActivities As ADODB.Recordset _
        ) As Boolean
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long
    Dim l_lngWgNum As Long
    Dim l_strActCd As String
    
On Error GoTo eh

    v_rsWGActivities.Filter = adFilterPendingRecords
    If v_rsWGActivities.EOF Then
        SaveWGActivities = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    Set l_objCM = Nothing
    
    Do Until v_rsWGActivities.EOF
        l_lngWgNum = AscertainKeyValue(v_rsWGActivities, "WG_NUM")
        l_strActCd = AscertainKeyValue(v_rsWGActivities, "ACTIVITY_CD")
        l_strSQL = SaveWGActivitiesSQL(v_rsWGActivities)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsWGActivities.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    SaveWGActivities = True
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = Join( _
                            Array( _
                                CStr(TrType_WGM_ACTIV), _
                                CStr(l_lngWgNum), _
                                "", _
                                "0", _
                                "0", _
                                l_strActCd) _
                            , ";")
            TrsRaiseError m_strCMODULENAME, "SaveWGActivities", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveWGActivities", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function


Public Function SaveGLActivities( _
            ByVal v_rsGLActivities As ADODB.Recordset _
        ) As Boolean
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long
    Dim l_strActCd As String
    
On Error GoTo eh

    v_rsGLActivities.Filter = adFilterPendingRecords
    If v_rsGLActivities.EOF Then
        SaveGLActivities = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open
    
    Set l_objCM = Nothing
    
    Do Until v_rsGLActivities.EOF
        l_strActCd = AscertainKeyValue(v_rsGLActivities, "ACTIVITY_CD")
        l_strSQL = SaveGLActivitiesSQL(v_rsGLActivities)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsGLActivities.MoveNext
    Loop

    l_objCN.Close
    Set l_objCN = Nothing
    GetObjectContext.SetComplete
    SaveGLActivities = True
    Exit Function
    
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort

    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = Join( _
                            Array( _
                                CStr(TrType_WGM_ACTIV), _
                                "0", _
                                "", _
                                "0", _
                                "0", _
                                l_strActCd) _
                            , ";")
            TrsRaiseError m_strCMODULENAME, "SaveGLActivities", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveGLActivities", , , , DescArgsFrom(l_strSQL), True
    End Select

End Function


