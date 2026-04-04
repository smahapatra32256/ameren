VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdCommentTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strcMODULENAME As String = "cdCommentTrans"
Private Const m_strcConnectStringReg As String = "SHADOWDIRECT"
        
Public Function SaveComment(ByVal v_adrsSaveComment As ADODB.Recordset _
        , ByVal v_lTRISUser As Long) As Boolean
        
    On Error GoTo SaveComment_Err
    
    Dim l_cn As ADODB.Connection
    Dim l_bReturn As Boolean
    Dim l_lHoldEmpNum As Long
    Dim l_lCommentId As Long
    
    Set l_cn = OpenConnection
    
    SaveComment = True
    
    With v_adrsSaveComment
        If RecordsExist(v_adrsSaveComment) Then
            .Sort = "RECORD_STATUS, EMP_NUM"
            'Process deletes
            .Filter = adFilterNone
            .MoveFirst
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
            Do While Not .EOF
                l_bReturn = DeleteComment(v_adrsSaveComment, l_cn)
                If Not l_bReturn Then
                    SaveComment = False
'                    GoTo SaveComment_Err
                End If
                .MoveNext
            Loop
            
            'Process inserts
            .Filter = adFilterNone
            .MoveFirst
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
            
            Do While Not .EOF
                If !EMP_NUM <> l_lHoldEmpNum Then
                    l_lCommentId = GetCommentId(!EMP_NUM)
                Else
                    l_lCommentId = l_lCommentId + 1
                End If
                v_adrsSaveComment!KEY_COMMENT_ID = l_lCommentId
                l_bReturn = InsertComment(v_adrsSaveComment, v_lTRISUser, l_cn)
                If Not l_bReturn Then
                    SaveComment = False
'                    GoTo SaveComment_Err
                End If
                .MoveNext
            Loop
            
            'Process Changes
            .Filter = adFilterNone
            .MoveFirst
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
            Do While Not .EOF
                If !KEY_EMP_NUM <> !EMP_NUM Then
                    If DeleteComment(v_adrsSaveComment, l_cn) Then
                        l_lCommentId = GetCommentId(!EMP_NUM)
                        v_adrsSaveComment!KEY_COMMENT_ID = l_lCommentId
                        l_bReturn = InsertComment(v_adrsSaveComment, v_lTRISUser, l_cn)
                        If Not l_bReturn Then
                            SaveComment = False
                        End If
                    End If
                Else
                    l_bReturn = UpdateComment(v_adrsSaveComment, v_lTRISUser, l_cn)
                    If Not l_bReturn Then
                        SaveComment = False
'                        GoTo SaveComment_Err
                    End If
                End If
                .MoveNext
            Loop
            .Filter = adFilterNone
        End If
    End With

    l_cn.Close
    Set l_cn = Nothing
    
    GetObjectContext.SetComplete

    Exit Function
    
SaveComment_Err:
    Dim sError As String
    
    l_cn.Close
    Set l_cn = Nothing
    
    GetObjectContext.SetAbort
    
    If Not v_adrsSaveComment Is Nothing Then
        sError = v_adrsSaveComment.GetString
    End If
    
    TrsRaiseError m_strcMODULENAME, "SaveComment", , True, , DescArgsFrom(sError), True
    
End Function
    
Private Function DeleteComment(v_adrsSaveComment As ADODB.Recordset, v_cn As ADODB.Connection) As Boolean
    Dim l_strSQL As String

    On Error GoTo DeleteComment_Err

    With v_adrsSaveComment
        l_strSQL = "DELETE FROM methdba.ttr_comments"
        l_strSQL = l_strSQL & " WHERE emp_num = " & !KEY_EMP_NUM
        l_strSQL = l_strSQL & " AND comment_id = " & !KEY_COMMENT_ID
    End With

    v_cn.Execute l_strSQL
    
    DeleteComment = True
    
    Exit Function
    
DeleteComment_Err:
    DeleteComment = False
    GetObjectContext.SetAbort
    TrsRaiseError m_strcMODULENAME, "DeleteComment", , True, , DescArgsFrom(l_strSQL), True
    
End Function

Private Function InsertComment(ByRef r_adrsSaveComment As ADODB.Recordset, ByVal v_lTRISUser As Long, v_cn As ADODB.Connection) As Boolean
    On Error GoTo InsertComment_Err
    
    Dim l_strSQL As String
    
    With r_adrsSaveComment
        l_strSQL = "INSERT INTO methdba.ttr_comments"
        l_strSQL = l_strSQL & " (emp_num, comment_id, start_date, "
        If Trim(!END_DATE) <> "" Then
            l_strSQL = l_strSQL & " end_date, "
        End If
        l_strSQL = l_strSQL & "comment, "
        l_strSQL = l_strSQL & " created_emp_num, last_mod_emp_num, "
        l_strSQL = l_strSQL & " created_timestamp, last_mod_timestamp) "
        l_strSQL = l_strSQL & " VALUES(" & !EMP_NUM
        l_strSQL = l_strSQL & " , " & !KEY_COMMENT_ID
        l_strSQL = l_strSQL & ", " & FormatForDB2(!START_DATE)
        If Trim(!END_DATE) <> "" Then
            l_strSQL = l_strSQL & ", " & FormatForDB2(!END_DATE)
        End If
        l_strSQL = l_strSQL & ", '" & Trim(Replace(!COMMENT, "'", "''")) & "'"
        l_strSQL = l_strSQL & ", " & v_lTRISUser
        l_strSQL = l_strSQL & ", " & v_lTRISUser
        l_strSQL = l_strSQL & ", CURRENT TIMESTAMP "
        l_strSQL = l_strSQL & ", CURRENT TIMESTAMP) "
    End With

    v_cn.Execute l_strSQL
    
    InsertComment = True
    
    Exit Function

InsertComment_Err:
    Dim l_adrsError As ADODB.Error
    
    InsertComment = False
    
    GetObjectContext.SetAbort

    For Each l_adrsError In v_cn.Errors
         If l_adrsError.NativeError = -803 Then '//-----db2 duplicate key-----//
            Err.Raise trErrbusSchedDemandDuplicateLineAsgnm
            Exit Function
         Else
            TrsRaiseError m_strcMODULENAME, "InsertComment", , True, , DescArgsFrom(l_strSQL), True
            Exit Function
         End If
    Next
    
    TrsRaiseError m_strcMODULENAME, "InsertComment", , True, , DescArgsFrom(l_strSQL), True
    
    
End Function

Private Function UpdateComment(v_adrsSaveComment As ADODB.Recordset _
        , ByVal v_lTRISUser As Long _
        , v_cn As ADODB.Connection) As Boolean
    On Error GoTo UpdateComment_Err
    
    Dim l_strSQL As String

    With v_adrsSaveComment
        l_strSQL = "UPDATE methdba.ttr_comments  SET"
        l_strSQL = l_strSQL & " emp_num = " & !EMP_NUM
        l_strSQL = l_strSQL & ", start_date = " & FormatForDB2(!START_DATE)
        If Trim(!END_DATE) <> "" Then
            l_strSQL = l_strSQL & ",  end_date = " & FormatForDB2(!END_DATE)
        Else
            l_strSQL = l_strSQL & ",  end_date = null "
        End If
        l_strSQL = l_strSQL & ", comment = '" & Trim(Replace(!COMMENT, "'", "''")) & "' "
        l_strSQL = l_strSQL & ", last_mod_emp_num = " & v_lTRISUser
        l_strSQL = l_strSQL & ", last_mod_timestamp =  CURRENT TIMESTAMP"
        l_strSQL = l_strSQL & " WHERE emp_num = " & !KEY_EMP_NUM
        l_strSQL = l_strSQL & " AND COMMENT_ID = " & !KEY_COMMENT_ID
    End With

    v_cn.Execute l_strSQL
   
    UpdateComment = True
    
    Exit Function
    
UpdateComment_Err:
    GetObjectContext.SetAbort
    UpdateComment = False
    TrsRaiseError m_strcMODULENAME, "UpdateComment", , True, , DescArgsFrom(l_strSQL), True
    
End Function

Private Function OpenConnection() As ADODB.Connection

    Dim l_objDBConnectMgr As New util_DBConnectMgr.clsuDBConnectMgr
    Dim l_cn As ADODB.Connection

On Error GoTo OpenConnection_Err

    Set l_cn = New ADODB.Connection

    With l_cn
        .CursorLocation = adUseClient       'Must use client to support RecordCount
        .ConnectionString = l_objDBConnectMgr.GetConnectionString(m_strcConnectStringReg)
        .Open
    End With

    Set OpenConnection = l_cn
    Set l_objDBConnectMgr = Nothing
    Set l_cn = Nothing

    Exit Function

OpenConnection_Err:
    Set l_objDBConnectMgr = Nothing
    
    Set l_cn = Nothing

    TrsRaiseError m_strcMODULENAME, "OpenConnection"
End Function

Private Function FormatForDB2(ByVal v_varData As Variant) As String
    Dim l_strFormated As String, fieldType As Integer
    
    On Error GoTo FormatForDB2_Err
    
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
    
    FormatForDB2 = l_strFormated
    
    Exit Function
    
FormatForDB2_Err:
    TrsRaiseError m_strcMODULENAME, "FormatForDB2", , , , DescArgsFrom(CStr(v_varData)), True
    
End Function

Private Function GetCommentId(ByVal v_lEmpNum As Long) As Long

    On Error GoTo GetCommentID_Err
    
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_adrsCommentId As ADODB.Recordset
    Dim l_objDBConnectMgr As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClauseEmp As String
   
    'Get Connection
    Set l_objDBConnectMgr = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    
    l_objCN.ConnectionString = l_objDBConnectMgr.GetConnectionString("SHADOWDIRECT")
    
    Set l_objDBConnectMgr = Nothing

    'Open Connection
    l_objCN.Open
    l_strSQL = "SELECT MAX(comment_id) as max_comment_id " & _
                    " FROM METHDBA.TTR_COMMENTS  " & _
                    " WHERE emp_num = " & v_lEmpNum
    
    Set l_adrsCommentId = New ADODB.Recordset
    With l_adrsCommentId
        .CursorLocation = adUseClient
        .Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
        If RecordsExist(l_adrsCommentId) Then
            If IsNull(!max_comment_id) Then
                GetCommentId = 1
            Else
                GetCommentId = !max_comment_id + 1
            End If
        Else
            GetCommentId = 1
        End If
    End With

    Set l_objCN = Nothing
    Set l_adrsCommentId = Nothing
    GetObjectContext.SetComplete

    Exit Function

GetCommentID_Err:
    GetObjectContext.SetAbort
    Set l_objDBConnectMgr = Nothing
    Set l_objCN = Nothing
    Set l_adrsCommentId = Nothing

    TrsRaiseError m_strcMODULENAME, "GetCommentID", , , , DescArgsFrom(l_strSQL), True
End Function


