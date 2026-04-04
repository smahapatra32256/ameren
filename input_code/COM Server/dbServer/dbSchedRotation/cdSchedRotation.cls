VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdSchedRotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdSchedRotation"
'provides the rotation header information by rotation code or default list of all rotation codes
Public Function GetRotationSchedHeader _
                (Optional ByVal v_intRotationCode As Integer _
                ) As ADODB.Recordset

On Error GoTo GetRotationSchedHeader_EH
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
  
    l_strSQL = RetrieveRotationHeader(v_intRotationCode)
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetRotationSchedHeader = Nothing
    Else
        Set GetRotationSchedHeader = l_objRS
    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing

Exit Function

GetRotationSchedHeader_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    
    TrsRaiseError m_strCMODULENAME, "GetRotationSchedHeader", , , , DescArgsFrom(l_strSQL), True
    
End Function
'provide sql for public method GetRotationSchedHeader (all or specific rotation code)
'or provide executable transactional rotation code to get header information
Private Function RetrieveRotationHeader( _
        Optional ByVal v_intRotationCode As Integer, _
        Optional ByVal cn As ADODB.Connection, _
        Optional ByRef r_rsRotationHeader As ADODB.Recordset) As String

On Error GoTo RetrieveRotationHeader_EH
    
    Dim strSQL As String
    Dim errArgs As New Collection
    Dim l_objRS As ADODB.Recordset

    If v_intRotationCode = 0 Then
        strSQL = "Select ROT_CD , ROT_DESC, LAST_MOD_EMP_NUM, CREATED_EMP_NUM, Char(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, Char(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP, " & _
                   " '' as RECORD_STATUS, '' as RECORD_ID, ROT_CD as KEY_ROT_CD  " & _
                   " FROM       METHDBA.TTR_ROTATION " & _
                   " WHERE ROT_CD > 0 " & _
                   " ORDER BY ROT_CD "
    Else
        strSQL = "Select ROT_CD , ROT_DESC, LAST_MOD_EMP_NUM, CREATED_EMP_NUM, Char(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, Char(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP, " & _
                   " '' as RECORD_STATUS, '' as RECORD_ID, ROT_CD as KEY_ROT_CD  " & _
                   " FROM       METHDBA.TTR_ROTATION " & _
                   " WHERE ROT_CD = " & v_intRotationCode & _
                   " AND ROT_CD > 0 "
    
    End If

    If cn Is Nothing Then
        RetrieveRotationHeader = strSQL
    Else
        If strSQL <> "" Then
            Set r_rsRotationHeader = cn.Execute(strSQL)
            If RecordsExist(r_rsRotationHeader) Then
            Else
               Set r_rsRotationHeader = Nothing
            End If
        End If
    End If
    Set l_objRS = Nothing
    
    
Exit Function

RetrieveRotationHeader_EH:
    Set l_objRS = Nothing
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "RetrieveRotationHeader", , True, , errArgs, True

End Function


'provides the rotation detail information by rotation code and rotation WEEK or default list of all rotation codes

Public Function GetRotationSchedDetail _
                (Optional ByVal v_intRotationCode As Integer, _
                 Optional ByVal v_intRotationWeek As Integer _
                ) As ADODB.Recordset

    
On Error GoTo GetRotationSchedDetail_EH
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = RetrieveRotationDetail(v_intRotationCode, v_intRotationWeek)

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetRotationSchedDetail = Nothing
    Else
        Set GetRotationSchedDetail = l_objRS
    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing
    
Exit Function

GetRotationSchedDetail_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    
    TrsRaiseError m_strCMODULENAME, "GetRotationSchedDetail", , , , DescArgsFrom(l_strSQL), True
    
End Function
'provide sql for public method GetRotationSchedDetail (all or specific rotation code and rotation week)
'or provide executable transactional rotation code/rotation week to get header information

Private Function RetrieveRotationDetail( _
        Optional ByVal v_intRotationCode As Integer, _
        Optional ByVal v_intRotationWeek As Integer, _
        Optional ByVal cn As ADODB.Connection, _
        Optional ByRef r_rsRotationDetail As ADODB.Recordset) As String

On Error GoTo RetrieveRotationDetail_EH
    
    Dim strSQL As String
    Dim errArgs As New Collection
    Dim l_objRS As ADODB.Recordset

    If v_intRotationCode = 0 Then
        strSQL = "Select ROT_CD , ROT_WEEK, SCHED_CD, " & _
            " LAST_MOD_EMP_NUM, CREATED_EMP_NUM, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP," & _
            " '' as RECORD_STATUS, '' as RECORD_ID, ROT_CD as KEY_ROT_CD, ROT_WEEK as KEY_ROT_WEEK " & _
                   " FROM       METHDBA.TTR_ROTATION_SCHED" & _
                   " WHERE ROT_CD > 0 " & _
                   " ORDER BY ROT_CD, ROT_WEEK"
    Else
        If v_intRotationWeek = 0 Then
            strSQL = "Select ROT_CD , ROT_WEEK, SCHED_CD, " & _
            " LAST_MOD_EMP_NUM, CREATED_EMP_NUM, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP," & _
                       " '' as RECORD_STATUS, '' as RECORD_ID, ROT_CD as KEY_ROT_CD, ROT_WEEK as KEY_ROT_WEEK " & _
                       " FROM       METHDBA.TTR_ROTATION_SCHED" & _
                       " WHERE ROT_CD = " & v_intRotationCode & _
                       " AND ROT_CD > 0 " & _
                       " ORDER BY ROT_WEEK"
        Else
            strSQL = "Select ROT_CD , ROT_WEEK, SCHED_CD, " & _
            " LAST_MOD_EMP_NUM, CREATED_EMP_NUM, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP," & _
                       " '' as RECORD_STATUS, '' as RECORD_ID, ROT_CD as KEY_ROT_CD, ROT_WEEK as KEY_ROT_WEEK " & _
                       " FROM       METHDBA.TTR_ROTATION_SCHED" & _
                       " WHERE ROT_CD = " & v_intRotationCode & _
                       " AND ROT_CD > 0 " & _
                       " AND ROT_WEEK = " & v_intRotationWeek
        End If
    End If

    If cn Is Nothing Then
        RetrieveRotationDetail = strSQL
    Else
        If strSQL <> "" Then
            Set r_rsRotationDetail = cn.Execute(strSQL)
            If RecordsExist(r_rsRotationDetail) Then
            Else
               Set r_rsRotationDetail = Nothing
            End If
        End If
    End If
    
    Set l_objRS = Nothing
        
    
Exit Function
RetrieveRotationDetail_EH:
    Set l_objRS = Nothing
    
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "RetrieveRotationDetail", , True, , errArgs, True

End Function




'provides delete, change and update for rotation header
'rotation detail children will be deleted if parent header is deleted
Public Function SaveScheduleRotation(ByVal v_rsScheduleRotationRequest As ADODB.Recordset) As Boolean

On Error GoTo SaveScheduleRotation_EH
    
    Dim cn As ADODB.Connection
    
    If RecordsExist(v_rsScheduleRotationRequest) Then
        v_rsScheduleRotationRequest.Filter = ""
        v_rsScheduleRotationRequest.MoveFirst
    Else
        SaveScheduleRotation = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    
    Set cn = OpenConnection
       
'  Perform Deletes
    ' Filter for Deletes
    v_rsScheduleRotationRequest.Filter = "RECORD_STATUS = 'D'"
    If Not v_rsScheduleRotationRequest.EOF Then
        Do Until v_rsScheduleRotationRequest.EOF
            DeleteRotationDetailChildren v_rsScheduleRotationRequest, cn
            DeleteRotationRequest v_rsScheduleRotationRequest, cn
            v_rsScheduleRotationRequest.MoveNext
        Loop
    End If
'  Perform Changes
    v_rsScheduleRotationRequest.Filter = ""
    v_rsScheduleRotationRequest.MoveFirst
    v_rsScheduleRotationRequest.Filter = "RECORD_STATUS = 'C'"
    Do Until v_rsScheduleRotationRequest.EOF
        UpdateRotationRequest v_rsScheduleRotationRequest, cn
        v_rsScheduleRotationRequest.MoveNext
    Loop
    
'  Perform Inserts
    v_rsScheduleRotationRequest.Filter = ""
    v_rsScheduleRotationRequest.MoveFirst
    v_rsScheduleRotationRequest.Filter = "RECORD_STATUS = 'I'"
    Do Until v_rsScheduleRotationRequest.EOF
        InsertRotationRequest v_rsScheduleRotationRequest, cn
        v_rsScheduleRotationRequest.MoveNext
    Loop

    cn.Close
    
    Set cn = Nothing
    GetObjectContext.SetComplete
    
    SaveScheduleRotation = True
    
Exit Function

SaveScheduleRotation_EH:

    GetObjectContext.SetAbort
    Set cn = Nothing
    
    TrsRaiseError m_strCMODULENAME, "SaveScheduleRotation", , , , , True
    

End Function




'insert a rotation header with the new next available generated rotation code
Private Sub InsertRotationRequest(v_rsScheduleRotationRequest As ADODB.Recordset, cn As ADODB.Connection)

   
On Error GoTo InsertRotationRequest_EH
    
    Dim ignoreFields As Variant
    Dim strSQL As String
    Dim errArgs As New Collection
    Dim addDefaultFields() As Variant

    ignoreFields = Array("RECORD_STATUS", "RECORD_ID", "KEY_ROT_CD", "CREATED_TIMESTAMP", "LAST_MOD_TIMESTAMP")
    strSQL = InsertStatement(v_rsScheduleRotationRequest, "METHDBA.TTR_ROTATION", ignoreFields)

    If strSQL <> "" Then
       cn.Execute strSQL
    End If
    
    GetObjectContext.SetComplete
    
Exit Sub

InsertRotationRequest_EH:

    Dim errLoop As ADODB.Error
    GetObjectContext.SetAbort
    For Each errLoop In cn.Errors
         If errLoop.NativeError = -803 Then
            Err.Raise trErrbusSchedRotDuplicateHeader 'this is a BRV error for duplicates to be handle in the server business component
            Exit Sub
         Else
            errArgs.Add strSQL
            TrsRaiseError m_strCMODULENAME, "InsertRotationRequest", , True, , errArgs, True
            Exit Sub
         End If
    Next

    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "InsertRotationRequest", , True, , errArgs, True

End Sub

'provide the next available rotation code for the new header record
Public Function GetNextRotationCode() As Integer

    
On Error GoTo GetNextRotationCode_EH
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
  
'    l_strSQL = "Select MAX(ROT_CD) + 1 FROM       METHDBA.TTR_ROTATION"
    l_strSQL = "Select Value((MAX(ROT_CD) + 1),1) FROM       METHDBA.TTR_ROTATION where ROT_CD < 20000 "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        GetNextRotationCode = 1  'if empty start at 1 for Rotation Code
    Else
        GetNextRotationCode = l_objRS(0)
    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing
    
Exit Function

GetNextRotationCode_EH:

    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetNextRotationCode, , , DescArgsFrom(l_strSQL), True"

End Function
'update rotation header description
Private Sub UpdateRotationRequest(v_rsScheduleRotationRequest As ADODB.Recordset, cn As ADODB.Connection)
    
On Error GoTo UpdateRotationRequest_EH
    
    Dim primKeys As Variant
    Dim ignoreFields As Variant
    Dim strSQL As String
    Dim errArgs As New Collection


    strSQL = "Update METHDBA.TTR_ROTATION " _
        & " SET ROT_DESC = '" & v_rsScheduleRotationRequest!ROT_DESC & "'" _
        & " ,LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" _
        & " ,LAST_MOD_EMP_NUM = " & v_rsScheduleRotationRequest!LAST_MOD_EMP_NUM _
        & " WHERE ROT_CD = " & v_rsScheduleRotationRequest!KEY_ROT_CD & " " _
        & " AND ROT_CD > 0"
    
    If strSQL <> "" Then
       cn.Execute strSQL
    End If
    GetObjectContext.SetComplete
    
Exit Sub

UpdateRotationRequest_EH:
    
    GetObjectContext.SetAbort
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "UpdateRotationRequest", , True, , errArgs, True

End Sub

'delete a rotation header
Private Sub DeleteRotationRequest(v_rsScheduleRotationRequest As ADODB.Recordset, cn As ADODB.Connection)

    
On Error GoTo DeleteRotationRequest_EH
    
    Dim strSQL As String
    Dim errArgs As New Collection
    
    strSQL = "DELETE FROM METHDBA.TTR_ROTATION" & _
            " WHERE ROT_CD = " & v_rsScheduleRotationRequest!KEY_ROT_CD & _
            " AND ROT_CD > 0 "
    If strSQL <> "" Then
       cn.Execute strSQL
    End If
    
    GetObjectContext.SetComplete
    
Exit Sub

DeleteRotationRequest_EH:
    GetObjectContext.SetAbort
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "DeleteRotationRequest", , True, , errArgs, True

End Sub

'delete rotation detail children for the deleted rotation header

Private Sub DeleteRotationDetailChildren( _
        v_rsScheduleRotationRequest As ADODB.Recordset, _
        cn As ADODB.Connection)

    
On Error GoTo DeleteRotationDetailChildren_EH
    
    Dim strSQL As String
    Dim errArgs As New Collection
    
    strSQL = "DELETE FROM METHDBA.TTR_ROTATION_SCHED " & _
            " WHERE ROT_CD = " & v_rsScheduleRotationRequest!KEY_ROT_CD & _
            " AND ROT_CD > 0"
 
    If strSQL <> "" Then
       cn.Execute strSQL
    End If
    GetObjectContext.SetComplete
    
Exit Sub
DeleteRotationDetailChildren_EH:
    GetObjectContext.SetAbort
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "DeleteRotationDetailChildren", , True, , errArgs, True

End Sub


'Detail*********
'provides delete, change and update for rotation detail
Public Function SaveRotationDetail(ByVal v_rsScheduleRotationDetailRequest As ADODB.Recordset) As Boolean

On Error GoTo SaveRotationDetail_EH
    
    Dim cn As ADODB.Connection
    
    If RecordsExist(v_rsScheduleRotationDetailRequest) Then
        v_rsScheduleRotationDetailRequest.Filter = ""
        v_rsScheduleRotationDetailRequest.MoveFirst
    Else
        SaveRotationDetail = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    
    Set cn = OpenConnection
       
'  Perform Deletes
    ' Filter for Deletes
    v_rsScheduleRotationDetailRequest.Filter = "RECORD_STATUS = 'D'"
    If Not v_rsScheduleRotationDetailRequest.EOF Then
        Do Until v_rsScheduleRotationDetailRequest.EOF
            DeleteRotationDetailRequest v_rsScheduleRotationDetailRequest, cn
            v_rsScheduleRotationDetailRequest.MoveNext
        Loop
    End If
    
'  Perform Changes
    v_rsScheduleRotationDetailRequest.Filter = ""
    v_rsScheduleRotationDetailRequest.MoveFirst
    v_rsScheduleRotationDetailRequest.Filter = "RECORD_STATUS = 'C'"
    Do Until v_rsScheduleRotationDetailRequest.EOF
        UpdateRotationDetailRequest v_rsScheduleRotationDetailRequest, cn
        v_rsScheduleRotationDetailRequest.MoveNext
    Loop
    
'  Perform Inserts
    v_rsScheduleRotationDetailRequest.Filter = ""
    v_rsScheduleRotationDetailRequest.MoveFirst
    v_rsScheduleRotationDetailRequest.Filter = "RECORD_STATUS = 'I'"
    Do Until v_rsScheduleRotationDetailRequest.EOF
       InsertRotationDetailRequest v_rsScheduleRotationDetailRequest, cn
       v_rsScheduleRotationDetailRequest.MoveNext
    Loop

    cn.Close
    
    Set cn = Nothing
    GetObjectContext.SetComplete
    
    
    SaveRotationDetail = True
    
Exit Function

SaveRotationDetail_EH:

    GetObjectContext.SetAbort
    Set cn = Nothing
    
    
    TrsRaiseError m_strCMODULENAME, "SaveRotationDetail", , , , , True
    

End Function



'insert rotation detail
Private Sub InsertRotationDetailRequest(v_rsScheduleRotationRequest As ADODB.Recordset, cn As ADODB.Connection)

On Error GoTo InsertRotationDetailRequest_EH
    
    Dim ignoreFields As Variant
    Dim strSQL As String
    Dim errArgs As New Collection
    Dim addDefaultFields() As Variant

    
    ignoreFields = Array("RECORD_STATUS", "RECORD_ID", "KEY_ROT_CD", "KEY_ROT_WEEK", "CREATED_TIMESTAMP", "LAST_MOD_TIMESTAMP")
    strSQL = InsertStatement(v_rsScheduleRotationRequest, "METHDBA.TTR_ROTATION_SCHED", ignoreFields)
    If strSQL <> "" Then
       cn.Execute strSQL
    End If
    
    GetObjectContext.SetComplete
    
Exit Sub

InsertRotationDetailRequest_EH:

    Dim errLoop As ADODB.Error
    
    GetObjectContext.SetAbort
    
    For Each errLoop In cn.Errors
         If errLoop.NativeError = -803 Then
            Err.Raise trErrbusSchedRotDuplicate
            Exit Sub
         Else
            errArgs.Add strSQL
            TrsRaiseError m_strCMODULENAME, "InsertRotationDetailRequest", , True, , errArgs, True
            Exit Sub
         End If
    Next

End Sub
'update rotation detail record
Private Sub UpdateRotationDetailRequest(v_rsScheduleRotationRequest As ADODB.Recordset, cn As ADODB.Connection)

On Error GoTo UpdateRotationDetailRequest_EH

    Dim primKeys As Variant
    Dim ignoreFields As Variant
    Dim strSQL As String
    Dim errArgs As New Collection
    
    strSQL = "Update METHDBA.TTR_ROTATION_SCHED " _
        & " SET SCHED_CD = " & v_rsScheduleRotationRequest!SCHED_CD & "" _
        & " ,ROT_WEEK = " & v_rsScheduleRotationRequest!ROT_WEEK & "" _
        & " ,LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" _
        & " ,LAST_MOD_EMP_NUM = " & v_rsScheduleRotationRequest!LAST_MOD_EMP_NUM _
        & " WHERE ROT_CD = " & v_rsScheduleRotationRequest!KEY_ROT_CD & " " _
        & " AND ROT_WEEK = " & v_rsScheduleRotationRequest!KEY_ROT_WEEK _
        & " AND ROT_CD > 0"
    
    If strSQL <> "" Then
       cn.Execute strSQL
    End If
    GetObjectContext.SetComplete
    
Exit Sub

UpdateRotationDetailRequest_EH:
    GetObjectContext.SetAbort
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "UpdateRotationDetailRequest", , True, , errArgs, True

End Sub
'Delete rotation detail
Private Sub DeleteRotationDetailRequest(v_rsScheduleRotationRequest As ADODB.Recordset, cn As ADODB.Connection)

On Error GoTo DeleteRotationDetailRequest_EH
    
    Dim primKeys As Variant
    Dim strSQL As String
    Dim errArgs As New Collection
    
    strSQL = "DELETE FROM METHDBA.TTR_ROTATION_SCHED" & _
            " WHERE ROT_CD = " & v_rsScheduleRotationRequest!KEY_ROT_CD & _
            " AND ROT_WEEK = " & v_rsScheduleRotationRequest!KEY_ROT_WEEK & _
            " AND ROT_CD > 0"
 
    If strSQL <> "" Then
       cn.Execute strSQL
    End If
    
    GetObjectContext.SetComplete
    
Exit Sub

DeleteRotationDetailRequest_EH:
    GetObjectContext.SetAbort
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "DeleteRotationDetailRequest", , True, , errArgs, True

End Sub

'Existence check for usage of rotation code - prevent deletion of rotation header when rot code in is use
Public Function IsRotationCodeNotAssigned _
                (ByVal v_intRotationCode As Integer _
                ) As Boolean

On Error GoTo IsRotationCodeNotAssigned_EH
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    If v_intRotationCode < 1 Then
       Err.Raise trErrdbInvalidZerointRotCode
    Else
    End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
  
    l_strSQL = RotationCodeNotAssigned(v_intRotationCode)
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        IsRotationCodeNotAssigned = True
    Else
        IsRotationCodeNotAssigned = False
    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing
    
Exit Function

IsRotationCodeNotAssigned_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerointRotCode
            TrsRaiseError m_strCMODULENAME, "IsRotationCodeNotAssigned", trErrdbInvalidZerointRotCode, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "IsRotationCodeNotAssigned", , , , DescArgsFrom(l_strSQL), True
    End Select
    
End Function

'provide sql for IsRotationCodeNotAssigned
'can be use with transactional items
Private Function RotationCodeNotAssigned( _
        ByVal v_intRotationCode As Integer, _
        Optional ByVal cn As ADODB.Connection, _
        Optional ByRef r_rsRotationHeader As ADODB.Recordset) As String

On Error GoTo RotationCodeNotAssigned_EH
    
    Dim strSQL As String
    Dim errArgs As New Collection
    If cn Is Nothing Then
        strSQL = "SELECT 1 AS EXISTS_SW " & _
                   " From SYSIBM.SYSDUMMY1 " & _
                   " WHERE Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_EMP_SCHED " & _
                   " WHERE ROT_CD = " & v_intRotationCode & ")" & _
                   " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_WG_BASE_SCHED " & _
                   " WHERE ROT_CD = " & v_intRotationCode & ")" & _
                   " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_WG_LINE_SCHED " & _
                   " WHERE ROT_CD = " & v_intRotationCode & ")"
        RotationCodeNotAssigned = strSQL
    Else
        strSQL = "SELECT 1 AS EXISTS_SW " & _
                   " From SYSIBM.SYSDUMMY1 " & _
                   " WHERE Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_EMP_SCHED " & _
                   " WHERE ROT_CD = " & v_intRotationCode & ")" & _
                   " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_WG_BASE_SCHED " & _
                   " WHERE ROT_CD = " & v_intRotationCode & ")" & _
                   " OR Exists " & _
                   " (SELECT * " & _
                   " From METHDBA.TTR_WG_LINE_SCHED " & _
                   " WHERE ROT_CD = " & v_intRotationCode & ")"
        If strSQL <> "" Then
            Set r_rsRotationHeader = cn.Execute(strSQL)
            If RecordsExist(r_rsRotationHeader) Then
            Else
               Set r_rsRotationHeader = Nothing
            End If
        End If
    End If
        
    
Exit Function

RotationCodeNotAssigned_EH:

    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "RotationCodeNotAssigned", , , , DescArgsFrom(strSQL), True

End Function

'Open connection for transactional items

Private Function OpenConnection() As ADODB.Connection
  
On Error GoTo OpenConnection_Eh
    
    Dim cm As New util_DBConnectMgr.clsuDBConnectMgr
    Dim cn As ADODB.Connection
    
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
    
OpenConnection_Eh:
    Set cm = Nothing
    Set cn = Nothing
    
    TrsRaiseError m_strCMODULENAME, "OpenConnection"
 
End Function

'provides if the rotation pattern exists?

Public Function GetRotDetPatterns _
        (ByRef r_rsRotationDetail As ADODB.Recordset _
                ) As ADODB.Recordset

    
On Error GoTo GetRotDetPatterns_EH
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = RetrieveRotDetPattern(r_rsRotationDetail)
    If l_strSQL <> "" Then
    Else
        Set GetRotDetPatterns = Nothing
        Exit Function
    End If

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetRotDetPatterns = Nothing
    Else
        Set GetRotDetPatterns = l_objRS
    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing
    
Exit Function

GetRotDetPatterns_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    
    TrsRaiseError m_strCMODULENAME, "GetRotDetPatterns", , , , DescArgsFrom(l_strSQL), True
    
End Function
'provide sql for public method GetRotDetPatterns (all or specific rotation code and rotation week)
'or provide executable transactional rotation code/rotation week to get header information

Private Function RetrieveRotDetPattern( _
        ByRef r_rsRotationDetail As ADODB.Recordset, _
        Optional ByVal cn As ADODB.Connection _
        ) As String

On Error GoTo RetrieveRotDetPattern_EH
    
    Dim strSQL As String
    Dim errArgs As New Collection
    Dim l_objRS As ADODB.Recordset
    Dim l_lngRotationWeekNoIndex As Long
    Dim l_lngRotationWeekNo As Long
    Dim l_lngRotationWeekNoMaxed As Long
    Dim l_intRotationSchedCd As Integer
    Dim l_lngRotationWeekLO As Long
    Dim l_lngRotationWeekHI As Long
    strSQL = "SELECT a.ROT_CD " & _
        " From (SELECT ROT_CD FROM METHDBA.TTR_ROTATION_SCHED " & _
        " WHERE ROT_CD > 0 " & _
        " group by rot_cd having count(*) = " & r_rsRotationDetail.RecordCount & ") a "
    
    l_intRotationSchedCd = r_rsRotationDetail!SCHED_CD
    l_lngRotationWeekLO = r_rsRotationDetail!ROT_WEEK
    l_lngRotationWeekHI = r_rsRotationDetail!ROT_WEEK
    l_lngRotationWeekNoMaxed = l_lngRotationWeekNoMaxed + 1
    r_rsRotationDetail.MoveNext
    If r_rsRotationDetail.EOF Then  'if there is only one record create this instance
       strSQL = strSQL & AddDetailWeekNoSql(l_lngRotationWeekLO, l_lngRotationWeekHI, l_intRotationSchedCd)
    Else
        For l_lngRotationWeekNoIndex = 2 To r_rsRotationDetail.RecordCount
            If l_intRotationSchedCd = r_rsRotationDetail!SCHED_CD Then 'bypass all duplicates with same sched codes
            Else
                l_lngRotationWeekHI = r_rsRotationDetail!ROT_WEEK - 1 'create this instance using current week -1 for previous HI
                strSQL = strSQL & AddDetailWeekNoSql(l_lngRotationWeekLO, l_lngRotationWeekHI, l_intRotationSchedCd)
                l_intRotationSchedCd = r_rsRotationDetail!SCHED_CD
                l_lngRotationWeekLO = r_rsRotationDetail!ROT_WEEK
                l_lngRotationWeekNoMaxed = l_lngRotationWeekNoMaxed + 1
                If l_lngRotationWeekNoMaxed > 14 Then
                    strSQL = "" 'init and exit
                    Exit Function
                End If
            End If
            r_rsRotationDetail.MoveNext
        Next
        'create last entry with calculated HI
        l_lngRotationWeekHI = r_rsRotationDetail.RecordCount
        strSQL = strSQL & AddDetailWeekNoSql(l_lngRotationWeekLO, l_lngRotationWeekHI, l_intRotationSchedCd)
    End If


    If cn Is Nothing Then
        RetrieveRotDetPattern = strSQL
    Else
        If strSQL <> "" Then
            Set r_rsRotationDetail = cn.Execute(strSQL)
            If RecordsExist(r_rsRotationDetail) Then
            Else
               Set r_rsRotationDetail = Nothing
            End If
        End If
    End If
    
    Set l_objRS = Nothing
        
    
Exit Function
RetrieveRotDetPattern_EH:
    Set l_objRS = Nothing
    
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "RetrieveRotDetPattern", , True, , errArgs, True

End Function

Private Function AddDetailWeekNoSql( _
        ByVal v_lngRotationWeekLO As Long, _
        ByVal v_lngRotationWeekHI As Long, _
        ByVal v_intRotationSched_cd As Integer _
        ) As String

On Error GoTo AddDetailWeekNoSql_EH

    Dim strSQL As String
    Dim errArgs As New Collection

    strSQL = " Inner Join (select distinct rot_cd "
    strSQL = strSQL & " From METHDBA.TTR_ROTATION_SCHED "
    strSQL = strSQL & " where (rot_week between " & v_lngRotationWeekLO & " and " & v_lngRotationWeekHI
    strSQL = strSQL & " and sched_cd = " & v_intRotationSched_cd & ")) "
    strSQL = strSQL & "t" & v_lngRotationWeekLO
    strSQL = strSQL & " on t" & v_lngRotationWeekLO & ".rot_cd = a.rot_cd"
    AddDetailWeekNoSql = strSQL


Exit Function
AddDetailWeekNoSql_EH:
    
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "AddDetailWeekNoSql", , True, , errArgs, True

End Function


''comment out until future use
''provides delete, change and update for rotation assignment
'Public Function SaveRotationAsgnmt(ByVal v_rsRotationAsgnmtRequest As ADODB.Recordset) As Boolean
'
'On Error GoTo SaveRotationAsgnmt_EH
'
'    Dim cn As ADODB.Connection
'
'    If v_rsRotationAsgnmtRequest.BOF And v_rsRotationAsgnmtRequest.EOF Then
'        SaveRotationAsgnmt = True
'        GetObjectContext.SetComplete
'        Exit Function
'    End If
'
'
'    Set cn = OpenConnection
'
''  Perform Deletes
'    v_rsRotationAsgnmtRequest.Filter = ""
'    v_rsRotationAsgnmtRequest.MoveFirst
'    ' Filter for Deletes
'    v_rsRotationAsgnmtRequest.Filter = "RECORD_STATUS = 'D'"
'    If Not v_rsRotationAsgnmtRequest.EOF Then
'        Do Until v_rsRotationAsgnmtRequest.EOF
'            DeleteRotationAsgnmt v_rsRotationAsgnmtRequest, cn
'            v_rsRotationAsgnmtRequest.MoveNext
'        Loop
'    End If
'''  Perform Changes
''    v_rsRotationAsgnmtRequest.Filter = ""
''    v_rsRotationAsgnmtRequest.MoveFirst
''    v_rsRotationAsgnmtRequest.Filter = "RECORD_STATUS = 'C'"
''    Do Until v_rsRotationAsgnmtRequest.EOF
''        UpdateRotationAsgnmt v_rsRotationAsgnmtRequest, cn
''        v_rsRotationAsgnmtRequest.MoveNext
''    Loop
'
''  Perform Inserts
'    v_rsRotationAsgnmtRequest.Filter = ""
'    v_rsRotationAsgnmtRequest.MoveFirst
'    v_rsRotationAsgnmtRequest.Filter = "RECORD_STATUS = 'I'"
'    Do Until v_rsRotationAsgnmtRequest.EOF
'        InsertRotationAsgnmt v_rsRotationAsgnmtRequest, cn
'        v_rsRotationAsgnmtRequest.MoveNext
'    Loop
'
'    cn.Close
'
'    Set cn = Nothing
'    GetObjectContext.SetComplete
'
'    SaveRotationAsgnmt = True
'
'Exit Function
'
'SaveRotationAsgnmt_EH:
'
'    GetObjectContext.SetAbort
'    Set cn = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "SaveRotationAsgnmt", , , , , True
'
'
'End Function
'
''comment out until future use
''insert a rotation header with the new next available generated rotation code
'Private Sub InsertRotationAsgnmt(v_rsScheduleRotationAsgnmt As ADODB.Recordset, cn As ADODB.Connection)
'
'
'On Error GoTo InsertRotationAsgnmt_EH
'
'    Dim ignoreFields As Variant
'    Dim strSQL As String
'    Dim errArgs As New Collection
'    Dim addDefaultFields() As Variant
'
'    ignoreFields = Array("RECORD_STATUS", "RECORD_ID", "KEY_ROT_CD", "KEY_LEVEL_NUM", "KEY_EMP_NUM", "KEY_WG_NUM", "TIME_STAMP")
'    strSQL = InsertStatement(v_rsScheduleRotationAsgnmt, "METHDBA.TTR_ROTATION_ASGNM", ignoreFields)
'
'    If strSQL <> "" Then
'       cn.Execute strSQL
'    End If
'
'Exit Sub
'
'InsertRotationAsgnmt_EH:
'
'    Dim errLoop As ADODB.Error
'    GetObjectContext.SetAbort
'    For Each errLoop In cn.Errors
'         If errLoop.NativeError = -803 Then
'            Err.Raise trErrbusSchedRotAsgnmtDup 'this is a BRV error for duplicates to be handle in the server business component
'            Exit Sub
'         Else
'            errArgs.Add strSQL
'            TrsRaiseError m_strCMODULENAME, "InsertRotationAsgnmt", , True, , errArgs, True
'            Exit Sub
'         End If
'    Next
'
'    errArgs.Add strSQL
'    TrsRaiseError m_strCMODULENAME, "InsertRotationAsgnmt", , True, , errArgs, True
'
'End Sub
'
''update rotation header description
'Private Sub UpdateRotationAsgnmt(v_rsScheduleRotationAsgnmt As ADODB.Recordset, cn As ADODB.Connection)
'
'On Error GoTo UpdateRotationAsgnmt_EH
'
'    Dim primKeys As Variant
'    Dim ignoreFields As Variant
'    Dim strSQL As String
'    Dim errArgs As New Collection
'
'
'    Dim l_utilDBDataServices As utilDBDataServices.cuDBDataServices
'
'    Set l_utilDBDataServices = New utilDBDataServices.cuDBDataServices
'
'    strSQL = "Update METHDBA.TTR_ROTATION_ASGNM " _
'        & " SET ROT_CD = '" & v_rsScheduleRotationAsgnmt!ROT_CD & "'," _
'        & " LEVEL_NUM = '" & v_rsScheduleRotationAsgnmt!LEVEL_NUM & "'," _
'        & " EMP_NUM = '" & v_rsScheduleRotationAsgnmt!EMP_NUM & "'," _
'        & " TIME_STAMP = '" & CVar(l_utilDBDataServices.GetTimeStamp) & "'," _
'        & " LAST_MOD_EMP_NUM = " & v_rsScheduleRotationAsgnmt!LAST_MOD_EMP_NUM _
'        & " WHERE ROT_CD = " & v_rsScheduleRotationAsgnmt!KEY_ROT_CD & " " _
'
'
'    If strSQL <> "" Then
'       cn.Execute strSQL
'    End If
'    GetObjectContext.SetComplete
'
'Exit Sub
'
'UpdateRotationAsgnmt_EH:
'
'    GetObjectContext.SetAbort
'    errArgs.Add strSQL
'    TrsRaiseError m_strCMODULENAME, "UpdateRotationAsgnmt", , True, , errArgs, True
'
'End Sub
''comment out until future use
''delete a rotation header
'Private Sub DeleteRotationAsgnmt(v_rsRotationAsgnmtRequest As ADODB.Recordset, cn As ADODB.Connection)
'
'On Error GoTo DeleteRotationAsgnmt_EH
'
'    Dim strSQL As String
'    Dim errArgs As New Collection
'
'    strSQL = "DELETE FROM METHDBA.TTR_ROTATION_ASGNM" & _
'            " WHERE ROT_CD = " & v_rsRotationAsgnmtRequest!KEY_ROT_CD & _
'            " AND LEVEL_NUM = " & v_rsRotationAsgnmtRequest!KEY_LEVEL_NUM & _
'            " AND EMP_NUM = " & v_rsRotationAsgnmtRequest!KEY_EMP_NUM & _
'            " AND WG_NUM = " & v_rsRotationAsgnmtRequest!KEY_WG_NUM
'
'
'    If strSQL <> "" Then
'       cn.Execute strSQL
'    End If
'
'    GetObjectContext.SetComplete
'
'Exit Sub
'
'DeleteRotationAsgnmt_EH:
'    GetObjectContext.SetAbort
'    errArgs.Add strSQL
'    TrsRaiseError m_strCMODULENAME, "DeleteRotationAsgnmt", , True, , errArgs, True
'
'End Sub
''comment out until future use
''provides the rotation assignment information  or default list of all RotationAsgnmt
'Public Function GetRotationAsgnmt _
'        (Optional ByVal v_intRotationCode As Integer, _
'         Optional ByVal v_intLevelNumber As Integer, _
'         Optional ByVal v_lngEmpNum As Long, _
'         Optional ByVal v_lngWGNum As Long _
'                ) As ADODB.Recordset
'
'On Error GoTo GetRotationAsgnmt_EH
'
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_strSQL As String
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
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
'    l_strSQL = RetrieveRotationAsgnmt(v_intRotationCode, v_intLevelNumber, v_lngEmpNum, v_lngWGNum)
'
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'    If l_objRS.BOF And l_objRS.EOF Then
'        Set GetRotationAsgnmt = Nothing
'    Else
'        Set GetRotationAsgnmt = l_objRS
'    End If
'    GetObjectContext.SetComplete
'    Set l_objRS = Nothing
'
'Exit Function
'
'GetRotationAsgnmt_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'    GetObjectContext.SetAbort
'
'    TrsRaiseError m_strCMODULENAME, "GetRotationAsgnmt", , , , DescArgsFrom(l_strSQL), True
'
'End Function
'comment out until future use
''provide sql for public method RetrieveRotationAsgnmt (all or specific rotation code)
'Private Function RetrieveRotationAsgnmt( _
'        Optional ByVal v_intRotationCode As Integer, _
'        Optional ByVal v_intLevelNumber As Integer, _
'        Optional ByVal v_lngEmpNum As Long, _
'        Optional ByVal v_lngWGNum As Long, _
'        Optional ByVal cn As ADODB.Connection, _
'        Optional ByRef r_rsRotationAsgnmt As ADODB.Recordset) As String
'
'On Error GoTo RetrieveRotationAsgnmt_EH
'
'    Dim strSQL As String
'    Dim errArgs As New Collection
'    Dim l_objRS As ADODB.Recordset
'
'    If v_intRotationCode = 0 Then
'        strSQL = "Select ROT_CD , LEVEL_NUM, EMP_NUM, WG_NUM, LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP, " & _
'                   " '' as RECORD_STATUS, '' as RECORD_ID, '' as KEY_ROT_CD, '' as KEY_LEVEL_NUM,'' as KEY_EMP_NUM,'' as KEY_WG_NUM " & _
'                   " FROM       METHDBA.TTR_ROTATION_ASGNM ORDER BY ROT_CD, LEVEL_NUM, EMP_NUM, WG_NUM"
'    Else
'        strSQL = "Select ROT_CD , LEVEL_NUM, EMP_NUM, WG_NUM, LAST_MOD_EMP_NUM, CHAR(TIME_STAMP) AS TIME_STAMP, " & _
'                   " '' as RECORD_STATUS, '' as RECORD_ID, '' as KEY_ROT_CD, '' as KEY_LEVEL_NUM,'' as KEY_EMP_NUM,'' as KEY_WG_NUM " & _
'                   " FROM       METHDBA.TTR_ROTATION_ASGNM " & _
'                   " WHERE ROT_CD = " & v_intRotationCode
'        If v_intLevelNumber = 0 Then
'        Else
'           strSQL = strSQL & " AND LEVEL_NUM = " & v_intLevelNumber
'        End If
'        If v_lngEmpNum = 0 Then
'        Else
'           strSQL = strSQL & " AND EMP_NUM = " & v_lngEmpNum
'        End If
'        If v_lngWGNum = 0 Then
'        Else
'           strSQL = strSQL & " AND WG_NUM = " & v_lngWGNum
'        End If
'    End If

'
'    If cn Is Nothing Then
'        RetrieveRotationAsgnmt = strSQL
'    Else
'        If strSQL <> "" Then
'            Set r_rsRotationAsgnmt = cn.Execute(strSQL)
'            If RecordsExist(r_rsRotationAsgnmt) Then
'            Else
'               Set r_rsRotationAsgnmt = Nothing
'            End If
'        End If
'    End If
'    Set l_objRS = Nothing
'
'
'Exit Function
'
'RetrieveRotationAsgnmt_EH:
'    Set l_objRS = Nothing
'    errArgs.Add strSQL
'    TrsRaiseError m_strCMODULENAME, "RetrieveRotationAsgnmt", , True, , errArgs, True
'
'End Function
'
