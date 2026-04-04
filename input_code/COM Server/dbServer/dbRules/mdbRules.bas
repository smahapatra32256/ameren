Attribute VB_Name = "mdbRules"
Option Explicit

Private Const m_strCMODULENAME As String = "mdbRules"

'*********************************************************
' Begin Private Caching methods
'*********************************************************

Public Function GetCacheRecords(ByVal v_strSQL As String, Optional eLockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long
    
On Error GoTo eh
    
    #If TRDEBUG Then
        stTime = GetTickCount()
    #End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open v_strSQL, l_objCN, adOpenStatic, eLockType
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    
'//-----the cache needs the meta data for formatting the key
'    If l_objRS.BOF And l_objRS.EOF Then
'        Set l_objRS = Nothing
'    End If
    
    Set GetCacheRecords = l_objRS
    
    Set l_objRS = Nothing
    Set l_objCN = Nothing
    Set l_objCM = Nothing
    
    #If TRDEBUG Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        Debug.Print "GetCacheRecords: " & CStr(dfTime / 1000)
    #End If
    
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCacheRecords", , , , DescArgsFrom(v_strSQL), True
End Function

'//-----copied from dbCodeTables
Public Function BuildWhereClause(ByVal v_strSQLKeyWord As String, _
                                    ByVal v_strColumn As String, _
                                    ByVal v_varColumnValues As Variant, _
                                    ByVal v_blnDelimit As Boolean) As String
    Dim l_strWhereClause As String
    Dim l_colErrArgs As Collection
On Error GoTo eh
    
    If IsMissing(v_varColumnValues) Or IsEmpty(v_varColumnValues) Then
        l_strWhereClause = ""
    ElseIf IsArray(v_varColumnValues) Then
        l_strWhereClause = " " & v_strSQLKeyWord & " " & BuildInClause(v_strColumn, v_varColumnValues, v_blnDelimit)
    Else
        l_strWhereClause = " " & v_strSQLKeyWord & " " & v_strColumn & " = " & _
            IIf(v_blnDelimit, StringEnclosedBy(Trim(CStr(v_varColumnValues)), "'"), Trim(CStr(v_varColumnValues)))
    End If

    BuildWhereClause = l_strWhereClause
    Exit Function
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strSQLKeyWord = " & v_strSQLKeyWord
    l_colErrArgs.Add "v_strColumn = " & v_strColumn
    If IsMissing(v_varColumnValues) Or IsEmpty(v_varColumnValues) Then
        l_colErrArgs.Add "v_varColumnValues = empty"
    ElseIf IsArray(v_varColumnValues) Then
        l_colErrArgs.Add "v_varColumnValues() = (" & Join(v_varColumnValues, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varColumnValues = " & v_varColumnValues
    End If
    l_colErrArgs.Add "v_blnDelimit = " & v_blnDelimit
    
    TrsRaiseError m_strCMODULENAME, "BuildWhereClause", , , , l_colErrArgs, True
End Function

Public Function BuildWhereClauseForMultiKey(ByVal v_strSQLKeyWord As Variant, _
                                    ByVal v_strColumn As Variant, _
                                    ByVal v_varColumnValues As Variant, _
                                    ByVal v_blnDelimit As Variant) As Variant
    Dim l_strWhereClause As String
    Dim l_colErrArgs As Collection
    Dim nKey As Long
On Error GoTo eh
    
    If IsMissing(v_varColumnValues) Or IsEmpty(v_varColumnValues) Then
        l_strWhereClause = ""
    Else
        If UBound(v_strSQLKeyWord) = UBound(v_strSQLKeyWord) And _
            UBound(v_strSQLKeyWord) = UBound(v_strColumn) And _
            UBound(v_strSQLKeyWord) = UBound(v_varColumnValues) And _
            UBound(v_strSQLKeyWord) = UBound(v_blnDelimit) Then
            l_strWhereClause = ""
            For nKey = LBound(v_strSQLKeyWord) To UBound(v_strSQLKeyWord)
                l_strWhereClause = l_strWhereClause & BuildWhereClause(v_strSQLKeyWord(nKey), v_strColumn(nKey), v_varColumnValues(nKey), v_blnDelimit(nKey)) & " "
            Next nKey
        Else
        End If
    End If
    
    BuildWhereClauseForMultiKey = l_strWhereClause
    Exit Function
eh:
    Set l_colErrArgs = New Collection
    
    AddToErrorCollection l_colErrArgs, v_strSQLKeyWord, "v_strSQLKeyWord"
    AddToErrorCollection l_colErrArgs, v_strColumn, "v_strColumn"
    AddToErrorCollection l_colErrArgs, v_varColumnValues, "v_varColumnValues"
    AddToErrorCollection l_colErrArgs, v_blnDelimit, "v_blnDelimit"
         
    TrsRaiseError m_strCMODULENAME, "BuildWhereClauseForMultiKey", , , , l_colErrArgs, True
End Function

Private Function AddToErrorCollection(ByRef r_colErrArgs As Collection, _
                                        ByRef v_varArgs As Variant, _
                                        ByVal v_strArgName As String)
    If IsMissing(v_varArgs) Or IsEmpty(v_varArgs) Then
        r_colErrArgs.Add v_strArgName & " = empty"
    ElseIf IsArray(v_varArgs) Then
        r_colErrArgs.Add v_strArgName & "() = (" & Join(v_varArgs, ", ") & ")"
    Else
        r_colErrArgs.Add v_strArgName & " = " & v_varArgs
    End If
End Function


