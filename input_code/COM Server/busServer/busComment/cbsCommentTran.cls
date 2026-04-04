VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsCommentTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsCommentTran"
Public Function SaveComment(ByVal v_rsSaveComment As ADODB.Recordset, _
                                ByRef r_rsErrors As ADODB.Recordset, _
                                ByVal v_lTRISUser As Long) As Boolean
                                
    On Error GoTo SaveComment_Err
                                
    Dim l_objComment As dbComment.cdCommentTrans
    Dim l_rsErrors As ADODB.Recordset
    Dim l_bError As Boolean, l_bReturn As Boolean

    Set r_rsErrors = CreateNewBusinessErrorSet
    
    If RecordsExist(v_rsSaveComment) Then
        With v_rsSaveComment
            .MoveFirst
'            Do While Not .EOF
'                If Trim(!COMMENT) <> "" Then
'                    !COMMENT = Trim(Replace(!COMMENT, "'", "''"))
'                    .Update
'                End If
'                .MoveNext
'            Loop
'            .MoveFirst
        End With
    End If

    If ValidateComment(v_rsSaveComment, r_rsErrors) Then
        l_bError = False
    Else
        l_bError = True
    End If
    
    If Not l_bError Then
        Set l_objComment = CreateObject("dbComment.cdCommentTrans")
        l_bReturn = l_objComment.SaveComment(v_rsSaveComment, v_lTRISUser)
        If Not l_bReturn Then
            l_bError = True
        End If
        Set l_objComment = Nothing
    End If

    If l_bError Then
        SaveComment = False
        GetObjectContext.SetAbort
    Else
        SaveComment = True
        GetObjectContext.SetComplete
    End If

    Exit Function
    
SaveComment_Err:
    SaveComment = False
    Set l_objComment = Nothing
    GetObjectContext.SetAbort
    
    Dim l_strFieldNames As String
    
    TrsRaiseError m_strCMODULENAME, "SaveComment", , , , , True
    

End Function

Private Function ValidateComment(ByVal v_rsSaveComment As ADODB.Recordset _
        , ByRef r_rsErrors As ADODB.Recordset) As Boolean
        
    Dim l_bCommentError As Boolean
    Dim l_bErrorsFound As Boolean
    Dim l_strBRVMessage As String
    Dim l_objEmployee As dbEmployee.cdEmployee
    Dim l_sStartDate As String, l_sEndDate As String
    Dim l_lApostropheCount As Long, l_iCounter As Integer
    
    On Error GoTo ValidateComment_Err
    
    l_bErrorsFound = False
    'PTS 1005
    
    ValidateComment = False

    If RecordsExist(v_rsSaveComment) Then
        With v_rsSaveComment
            .MoveFirst
            Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
            Do While Not .EOF
                If IsNull(!START_DATE) Then
                    l_sStartDate = ""
                Else
                    l_sStartDate = CStr(!START_DATE)
                End If
                If IsNull(!END_DATE) Then
                    l_sEndDate = ""
                Else
                    l_sEndDate = CStr(!END_DATE)
                End If
                l_bCommentError = False
                If Not l_objEmployee.IsValidEmployee(!KEY_EMP_NUM) Then
                    l_bErrorsFound = True
                    l_bCommentError = True
                    l_strBRVMessage = "Employee (" & !KEY_EMP_NUM & ") is not valid."
                    AppendBusinessError r_rsErrors, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
                        l_strBRVMessage, m_strCMODULENAME, _
                        1, !RECORD_ID, "EMP_NUM", _
                        Array("EMP_NUM", "START_DATE", "END_DATE", "KEY_COMMENT_ID"), _
                        CStr(!EMP_NUM), _
                        l_sStartDate, _
                        l_sEndDate, _
                        !KEY_COMMENT_ID
                Else
                    Select Case !RECORD_STATUS
                        Case "I", "C"
                            If Not l_objEmployee.IsValidEmployee(!EMP_NUM) Then
                                l_bErrorsFound = True
                                l_bCommentError = True
                                l_strBRVMessage = "Employee (" & !EMP_NUM & ") is not valid."
                                AppendBusinessError r_rsErrors, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
                                    l_strBRVMessage, m_strCMODULENAME, _
                                    1, !RECORD_ID, "EMP_NUM", _
                                    Array("EMP_NUM", "START_DATE", "END_DATE", "KEY_COMMENT_ID"), _
                                    CStr(!EMP_NUM), _
                                    l_sStartDate, _
                                    l_sEndDate, _
                                    !KEY_COMMENT_ID
                            End If
                            If Not l_bCommentError Then
                                'Start date is required
                                If Trim(!START_DATE) <> "" Then
                                    If Not IsDate(!START_DATE) Then
                                        l_bErrorsFound = True
                                        l_bCommentError = True
                                        l_strBRVMessage = "Start date is not valid."
                                        AppendBusinessError r_rsErrors, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
                                                l_strBRVMessage, m_strCMODULENAME, _
                                                1, !RECORD_ID, "START_DATE", _
                                                Array("EMP_NUM", "START_DATE", "END_DATE", "KEY_COMMENT_ID"), _
                                                CStr(!EMP_NUM), _
                                                l_sStartDate, _
                                                l_sEndDate, _
                                                !KEY_COMMENT_ID
                                    End If
                                Else
                                    l_bErrorsFound = True
                                    l_bCommentError = True
                                    l_strBRVMessage = "Start date is required."
                                    AppendBusinessError r_rsErrors, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
                                                l_strBRVMessage, m_strCMODULENAME, _
                                                1, !RECORD_ID, "START_DATE", _
                                                Array("EMP_NUM", "START_DATE", "END_DATE", "KEY_COMMENT_ID"), _
                                                CStr(!EMP_NUM), _
                                                l_sStartDate, _
                                                l_sEndDate, _
                                                !KEY_COMMENT_ID
                                End If
                            End If
                            If Not l_bCommentError Then
                                'End date is NOT required, but if provided, must be a valid date
                                If Trim(!END_DATE) <> "" Then
                                    If Not IsDate(!END_DATE) Then
                                        l_bErrorsFound = True
                                        l_bCommentError = True
                                        l_strBRVMessage = "End date is not valid."
                                        AppendBusinessError r_rsErrors, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
                                                l_strBRVMessage, m_strCMODULENAME, _
                                                1, !RECORD_ID, "END_DATE", _
                                                Array("EMP_NUM", "START_DATE", "END_DATE", "KEY_COMMENT_ID"), _
                                                CStr(!EMP_NUM), _
                                                l_sStartDate, _
                                                l_sEndDate, _
                                                !KEY_COMMENT_ID
                                    End If
                                Else
                                    l_bErrorsFound = True
                                    l_bCommentError = True
                                    l_strBRVMessage = "End date is required."
                                    AppendBusinessError r_rsErrors, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
                                                l_strBRVMessage, m_strCMODULENAME, _
                                                1, !RECORD_ID, "END_DATE", _
                                                Array("EMP_NUM", "START_DATE", "END_DATE", "KEY_COMMENT_ID"), _
                                                CStr(!EMP_NUM), _
                                                l_sStartDate, _
                                                l_sEndDate, _
                                                !KEY_COMMENT_ID
                                End If
                            End If
                            If Not l_bCommentError Then
                                'Check the sequence of the start and end dates
                                If Trim(!START_DATE) <> "" And Trim(!END_DATE) <> "" Then
                                    If IsDate(!START_DATE) And IsDate(!END_DATE) Then
                                        If CDate(!START_DATE) > CDate(!END_DATE) Then
                                            l_bErrorsFound = True
                                            l_bCommentError = True
                                            l_strBRVMessage = "End date must be greater or equal to the start date."
                                            AppendBusinessError r_rsErrors, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
                                                    l_strBRVMessage, m_strCMODULENAME, _
                                                    1, !RECORD_ID, "END_DATE", _
                                                    Array("EMP_NUM", "START_DATE", "END_DATE", "KEY_COMMENT_ID"), _
                                                    CStr(!EMP_NUM), _
                                                    l_sStartDate, _
                                                    l_sEndDate, _
                                                    !KEY_COMMENT_ID
                                        End If
                                    End If
                                End If
                            End If
                            If Not l_bCommentError Then
                                l_lApostropheCount = 0
                                For l_iCounter = 1 To Len(Trim(!COMMENT))
                                    If Mid(!COMMENT, l_iCounter, 1) = "'" Then
                                        l_lApostropheCount = l_lApostropheCount + 1
                                    End If
                                Next l_iCounter
                                'Comment length is limited to 255 characters
                                If Len(Trim(!COMMENT)) > (255 - l_lApostropheCount) Then
                                    l_bErrorsFound = True
                                    l_bCommentError = True
                                    l_strBRVMessage = "Comment length is greater than 255 characters.  (Note:  Each apostrophe is counted as two characters.)"
                                    AppendBusinessError r_rsErrors, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
                                        l_strBRVMessage, m_strCMODULENAME, _
                                        1, !RECORD_ID, "COMMENT", _
                                        Array("EMP_NUM", "START_DATE", "END_DATE", "KEY_COMMENT_ID"), _
                                        CStr(!EMP_NUM), _
                                        l_sStartDate, _
                                        l_sEndDate, _
                                        !KEY_COMMENT_ID
                                Else
                                    'Comment length > 0 characters
                                    If Len(Trim(!COMMENT)) < 1 Then
                                        l_bErrorsFound = True
                                        l_bCommentError = True
                                        l_strBRVMessage = "Comment is required."
                                        AppendBusinessError r_rsErrors, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
                                            l_strBRVMessage, m_strCMODULENAME, _
                                            1, !RECORD_ID, "COMMENT", _
                                            Array("EMP_NUM", "START_DATE", "END_DATE", "KEY_COMMENT_ID"), _
                                            CStr(!EMP_NUM), _
                                            l_sStartDate, _
                                            l_sEndDate, _
                                            !KEY_COMMENT_ID
                                    End If
                                End If
                            End If
                        Case "D"
                        '
                    End Select
                End If
                .MoveNext
            Loop
            Set l_objEmployee = Nothing
        End With
    End If
   
    If l_bErrorsFound Then
        ValidateComment = False
        GetObjectContext.SetAbort
    Else
        ValidateComment = True
        GetObjectContext.SetComplete
    End If

    Exit Function
    
ValidateComment_Err:
    Set l_objEmployee = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "ValidateComment", , , , , True
    
End Function



