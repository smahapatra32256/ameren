VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_datCOPTIONALSTARTDATE As Date = #1/1/100#
Private Const m_datCOPTIONALENDDATE As Date = #12/31/9999#

Private Const m_strcMODULENAME As String = "cdComment"

Public Function GetComment(ByVal v_aryEmpNum As Variant, _
                                ByVal v_dtStartDate As Date, _
                                ByVal v_dtEndDate As Date) As ADODB.Recordset

On Error GoTo GetComment_Err

    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_adrsComments As ADODB.Recordset
    Dim l_objDBConnectMgr As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_strInClauseEmp As String
    
    Dim l_aryEmpNum As Variant
    
    'Get Connection
    Set l_objDBConnectMgr = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objDBConnectMgr.GetConnectionString("SHADOWDIRECT")
    Set l_objDBConnectMgr = Nothing

    'Open Connection
    l_objCN.Open

    'Format SQL Statement
    If IsArray(v_aryEmpNum) Then
        l_strInClauseEmp = BuildInClause("A.EMP_NUM", v_aryEmpNum, False)
    Else
        l_strInClauseEmp = "A.EMP_NUM = " & str(CLng(v_aryEmpNum))
    End If
    l_strSQL = "SELECT A.emp_num, B.emp_name, A.start_date , A.end_date , A.comment, "
    l_strSQL = l_strSQL & " A.created_emp_num, A.last_mod_emp_num,"
    l_strSQL = l_strSQL & " CHAR(A.created_timestamp) AS created_timestamp, CHAR(A.last_mod_timestamp) AS last_mod_timestamp, "
    l_strSQL = l_strSQL & "  comment_id "
    l_strSQL = l_strSQL & " FROM METHDBA.TTR_COMMENTS A, METHDBA.TTR_EMPLOYEE B "
    l_strSQL = l_strSQL & " WHERE " & l_strInClauseEmp
    l_strSQL = l_strSQL & " AND A.emp_num = B.emp_num "
    If v_dtStartDate > m_datCOPTIONALSTARTDATE Then
        l_strSQL = l_strSQL & " AND ((START_DATE  >= " & FormatForDB2(v_dtStartDate) & "  AND START_DATE <= " & FormatForDB2(v_dtEndDate) & ") "
        l_strSQL = l_strSQL & "     OR (END_DATE  >= " & FormatForDB2(v_dtStartDate) & " AND END_DATE <= " & FormatForDB2(v_dtEndDate) & ") "
        l_strSQL = l_strSQL & "     OR (START_DATE < " & FormatForDB2(v_dtStartDate) & " AND END_DATE > " & FormatForDB2(v_dtEndDate) & ")) "
    End If
    l_strSQL = l_strSQL & " ORDER BY EMP_NUM, START_DATE, END_DATE"

    Set l_adrsComments = New ADODB.Recordset
    With l_adrsComments
        .CursorLocation = adUseClient
        .Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
        If Not l_adrsComments Is Nothing Then
           Set GetComment = l_adrsComments
        Else
            Set GetComment = Nothing
        End If
    End With

    Set l_objCN = Nothing
    Set l_adrsComments = Nothing

    Exit Function

GetComment_Err:
    Set l_objDBConnectMgr = Nothing
    Set l_objCN = Nothing
    Set l_adrsComments = Nothing

    TrsRaiseError m_strcMODULENAME, "GetComment", , , , DescArgsFrom(l_strSQL), True
End Function

Private Function FormatForDB2(ByVal v_varData As Variant) As String
    Dim l_strFormated As String, l_iFieldType As Integer
    
    On Error GoTo eh
    
    FormatForDB2 = ""
    
    l_iFieldType = VarType(v_varData)
    Select Case l_iFieldType
        Case vbString
            FormatForDB2 = "'" & Replace(v_varData, "'", "''") & "'"
        Case vbInteger, vbLong
            FormatForDB2 = CStr(v_varData)
        Case vbDate
            FormatForDB2 = "'" & Format(v_varData, "mm/dd/yyyy") & "'"
        Case Else
            FormatForDB2 = Format(v_varData)
    End Select
    
    Exit Function
eh:
    TrsRaiseError m_strcMODULENAME, "FormatForDB2", , , , DescArgsFrom(CStr(v_varData)), True
End Function


