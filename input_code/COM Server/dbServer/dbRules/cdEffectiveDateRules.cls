VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdEffectiveDateRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdEffectiveDateRules"

'Each Rule will have a separate cache group
'All rules are a simple look up for existence of a value in the rule table
'The key will contain 2 characters for data type and the remaining is the value
'Please see mTRISContstance for the data type definitions


' Rule constants
'The Key is 2 characters for the data type and the remaining is the value
'Private Const CRG_KEY = 0

' Department constants
'Private Const CRG_GROUPNAME = "RULExxxxCACHE_EFFDATE"
' index for record dynamic array
Private Const CRG_RG_CD = 0
Private Const CRG_RG_SUB_CD = 1
Private Const CRG_BUS_TYPE_CD = 2
Private Const CRG_RG_VALUE_CD = 3
Private Const CRG_EFF_START_DT = 4
Private Const CRG_EFF_END_DT = 5
Private Const CRG_SEQ_NUM = 6
Private Const CRG_APPLY_IND = 7

'//TODO: Can each key value contain multiple effective date ranges?
'//TODO: If so, then OneRowForEachKey needs to be False and GetRecord will return a variant array.  (see dbCodeTables.GetComponentAmountOn and FindComponentAmountIn)
'//TODO: Does this need to support an effective date range or more than one date ???
Public Function DoesEffectiveDateRuleApply( _
            ByVal v_lngRule As Long, _
            ByVal v_lngDataType As TRSRulesDataTypes, _
            ByVal v_strValue As String, _
            ByVal v_datDate As Date) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim l_strGroupName As String
    Dim l_strKey As String
    
On Error GoTo eh
    'l_strKey = RuleKey(v_lngDataType, v_strValue)
    l_strGroupName = GetRuleGroupName(v_lngRule)
    l_varRecord = GetRecord(Array(v_lngRule, v_lngDataType, v_strValue), l_strGroupName, lblnExists)
    
    DoesEffectiveDateRuleApply = FindEffectiveDate(v_datDate, l_varRecord)
    '//DoesEffectiveDateRuleApply = Not IsEmpty(l_varRecord)
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "DoesEffectiveDateRuleApply"
End Function
    
Private Function FindEffectiveDate(ByVal v_datEffDate As Date, r_varRecord As Variant) As Boolean
' l_varRowValues is a 2 dim array sorted in EFF_DATE descending order
    Dim i As Integer
    Dim l_varTemp As Variant
    Dim l_varRowValues As Variant
    Dim l_blnFoundDate As Boolean

On Error GoTo eh
    l_blnFoundDate = False
    If IsEmpty(r_varRecord) Then
        FindEffectiveDate = l_blnFoundDate
        Exit Function
    End If
    
    l_varRowValues = r_varRecord(1) '1st record is key, 2nd record is an array of columns and rows
    For i = LBound(l_varRowValues, 2) To UBound(l_varRowValues, 2) '2nd dimension - rows
        If v_datEffDate >= l_varRowValues(CRG_EFF_START_DT, i) And v_datEffDate <= l_varRowValues(CRG_EFF_END_DT, i) Then
            'FindEffectiveDate = l_varRowValues(i) '//cannot easily return just a single row
            l_blnFoundDate = True
            Exit For
        End If
    Next

    FindEffectiveDate = l_blnFoundDate
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "FindEffectiveDate"
End Function

'*********************************************************
' Begin Public Cache management methods
'*********************************************************

Public Function GetGroupRecords( _
                    ByVal v_strGroupName As String, _
                    ByVal v_varKeys As Variant, _
                    ByRef r_varKeyArray As Variant, _
                    ByRef r_varColumnArray As Variant) As ADODB.Recordset
    Dim l_rsRecords As ADODB.Recordset
    Dim l_colErrArgs As Collection
On Error GoTo eh

'    Select Case v_strGroupName
'        Case CRG_GROUPNAME
        Set l_rsRecords = GetRuleCacheRecords(v_strGroupName, v_varKeys)
        r_varKeyArray = Array("RG_CD", "BUS_TYPE_CD", "RG_VALUE_CD")
        r_varColumnArray = Empty
'    End Select
    Set GetGroupRecords = l_rsRecords

    Set l_rsRecords = Nothing
    Exit Function
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    l_colErrArgs.Add "v_varKeys = " & v_varKeys
    l_colErrArgs.Add "l_rsRecords.Source = " & l_rsRecords.Source
    
    Set l_rsRecords = Nothing
    TrsRaiseError m_strCMODULENAME, "GetGroupRecords", , , , l_colErrArgs, True
End Function

Private Function GetRuleCacheRecords(ByVal v_strGroupName As String, _
                                    Optional ByVal v_varKeys As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    Dim l_lngRule As Long
    
On Error GoTo eh
    
    If IsMissing(v_varKeys) Or IsEmpty(v_varKeys) Then '//-----Bulk load for entire cache group-----//
        l_lngRule = GetRuleNumFromGroupName(v_strGroupName)
        l_strWhereClause = " WHERE RG_CD = " & l_lngRule
    ElseIf IsArray(v_varKeys) Then '//-----single row add for invalidate-----//
        If UBound(v_varKeys) - LBound(v_varKeys) = 2 Then
            l_strWhereClause = mdbRules.BuildWhereClauseForMultiKey(Array("WHERE", "AND", "AND"), _
                                                    Array("RG_CD", "BUS_TYPE_CD", "RG_VALUE_CD"), _
                                                    v_varKeys, _
                                                    Array(False, False, True))
        Else
            Err.Raise trErrdbGetCacheSQLKey
        End If
    Else
        l_strWhereClause = " WHERE RG_CD = " & v_varKeys
    End If
    
    l_strSQL = "SELECT RG_CD, RG_SUB_CD, BUS_TYPE_CD, RG_VALUE_CD, EFF_START_DT, EFF_END_DT, SEQ_NUM, APPLY_IND " & _
                " FROM METHDBA.TTR_RG_EFF_DATE " & _
                l_strWhereClause & " " & _
                " ORDER BY RG_CD, BUS_TYPE_CD, RG_VALUE_CD, EFF_START_DT DESC, EFF_END_DT DESC"
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetRuleCacheRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetRuleCacheRecords", , , , DescArgsFrom(l_strSQL), True
End Function

Private Function GetRecord( _
            ByVal v_varKey As Variant, _
            ByVal v_strGroupName As String, _
            ByRef r_blnExists As Boolean) As Variant
' Answer an variant array, representing a record, or empty.
' This method should not invoke InitializeGroup unless the
' group is setup for bulk initialization.  If setup for lazy
' initialization it should only get the single record and add it
' to the cache.  If the record is not found, using lazy, it should
' add a value to the cache to denote invalid so that it does not
' attempt to retrieve it from the database again.  Currently, a string
' constant is added for record not found and the spmCache component
' will answer that it exists, but the property value is a
' RECORDNOTFOUND string.
'
    Dim recCache As utilRecordCache.cuRecordCache
    Dim l_colErrArgs As Collection
           
On Error GoTo eh

'//--------------------------------------------------------------------------------------------------------------//
'//-----http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cossdk/htm/imtxas_2bs6.asp
'//-----Do not use the SafeRef method in COM+. In MTS, SafeRef was used by an object to obtain
'//-----a reference to itself that could be passed outside its context wrapper.
'//-----With contexts and interception fully integrated into COM+, this is no longer necessary.
'//-----
'//-----[C/C++]
'//-----Note   The C++ function SafeRef should be used instead of this method.
'//--------------------------------------------------------------------------------------------------------------//

        Set recCache = CreateObject("utilRecordCache.cuRecordCache")
    
        If IsArray(v_varKey) Then
            GetRecord = recCache.GetRecord2(v_varKey, v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                r_blnExists, Me)
        Else
            GetRecord = recCache.GetRecord(CStr(v_varKey), v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                r_blnExists, Me)
        End If
        
        Set recCache = Nothing
    Exit Function
eh:
    Set recCache = Nothing

    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & CStr(v_varKey)
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    
    TrsRaiseError m_strCMODULENAME, "GetRecord", , , , l_colErrArgs, True
End Function

'*********************************************************
' Begin Private Caching methods
'*********************************************************

Public Sub InvalidateRulesCache(ByVal v_lngRule As Long, _
                                ByVal v_lngBusTypeCd As Long, _
                                ByVal v_strRgValueCd As String)
    Dim l_strGroupName As String

On Error GoTo eh
   
    l_strGroupName = GetRuleGroupName(v_lngRule)
    InvalidateCache Array(v_lngRule, v_lngBusTypeCd, v_strRgValueCd), l_strGroupName
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateRulesCache"
End Sub

'//-----code has been factored into utilRecordCache-----//
Private Sub InvalidateCache( _
                ByVal v_varKey As Variant, _
                ByVal v_strGroupName As String)

    Dim recCache As utilRecordCache.cuRecordCache
    Dim l_colErrArgs As Collection
    Dim l_blnExists As Boolean
On Error GoTo eh

    Set recCache = CreateObject("utilRecordCache.cuRecordCache")
    
    If IsArray(v_varKey) Then '//-----uses metadata to format key, may need to retrieve data to get metadata
        recCache.InvalidateCache2 v_varKey, v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                l_blnExists, Me
    Else
        recCache.InvalidateCache CStr(v_varKey), v_strGroupName
    End If
        
    Set recCache = Nothing

    Exit Sub
eh:
    Set recCache = Nothing
    
    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & CStr(v_varKey)
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    
    TrsRaiseError m_strCMODULENAME, "InvalidateCache", , , , l_colErrArgs, True
End Sub

'//-----used by RecordCache-----//

Private Function IsLazyInitialized(ByVal v_strGroupName As String) As Boolean
' If group is not lazy initialized, then it is bulk initialized.
On Error GoTo eh
'    Select Case Trim(v_strGroupName)
'    Case CRG_GROUPNAME
        IsLazyInitialized = False
'    Case Else
'        IsLazyInitialized = True
'    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsLazyInitialized"
End Function


Private Function GroupSupportsLoadRecord(ByVal v_strGroupName As String) As Boolean
'All group caches that support Invalidate must use parameterized SQL in corresponding GetxxxRecords (GetGroupRecords)
On Error GoTo eh
'    Select Case Trim(v_strGroupName)
'    Case CRG_GROUPNAME
        GroupSupportsLoadRecord = True
'    Case Else
'        GroupSupportsLoadRecord = False
'    End Select
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GroupSupportsLoadRecord"
End Function


Private Function OneRowForEachKey(ByVal v_strGroupName As String) As Boolean
'//-----normally each key/row is cached into a 1-dimensional array that represents each column-----//
'//-----union components will have multiple rows per key and will be cached into a 2-dimensional array-----//
On Error GoTo eh
'    Select Case Trim(v_strGroupName)
'    Case CRG_GROUPNAME
'        OneRowForEachKey = True
'    Case Else
        OneRowForEachKey = False
'    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OneRowForEachKey"
End Function
   
Private Function GetRuleGroupName( _
    ByVal v_lngRule As Long _
    ) As String

    Dim l_strRule As String
    Dim l_str As String
    Dim l_strResult As String

On Error GoTo eh

    l_str = Trim(str(v_lngRule))
    l_strRule = String(4 - Len(l_str), "0") & l_str
    l_strResult = "RULE" & l_strRule & "CACHE_EFFDATE"

    GetRuleGroupName = l_strResult
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetRuleGroupName"

End Function

Private Function GetRuleNumFromGroupName( _
    ByVal v_strGroupName As String _
    ) As Long

    Dim l_str As String

On Error GoTo eh

    l_str = Trim(Mid(v_strGroupName, 5, 4))
    GetRuleNumFromGroupName = CLng(l_str)

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetRuleNumFromGroupName"

End Function
    
Public Function GetEffectiveDateRule( _
    ByVal v_lngRule As Long, _
    Optional ByVal v_strSubRule As String = "" _
    ) As ADODB.Recordset
       
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr


    On Error GoTo eh
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open

    l_strSQL = "select RG_CD, RG_SUB_CD, BUS_TYPE_CD, RG_VALUE_CD, SEQ_NUM " & _
                ", EFF_START_DT, EFF_END_DT  " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN   SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1) " & _
                    "  ELSE '' END AS VALUE_PART1 " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN   SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD)) " & _
                    "  ELSE '' END AS VALUE_PART2 " & _
                "from METHDBA.TTR_RG_EFF_DATE " & _
                "where RG_CD = " & str(v_lngRule)
    If v_strSubRule <> "" Then
         l_strSQL = l_strSQL & " and RG_SUB_CD = '" & v_strSubRule & "'"
    End If
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEffectiveDateRule = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEffectiveDateRule", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function GetEffectiveDateRules(ByVal v_varRules As Variant) As ADODB.Recordset
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo eh
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open

    l_strSQL = "select RG_CD, RG_SUB_CD, BUS_TYPE_CD, RG_VALUE_CD, SEQ_NUM " & _
                ", EFF_START_DT, EFF_END_DT  " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN   SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1) " & _
                    "  ELSE '' END AS VALUE_PART1 " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN   SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD)) " & _
                    "  ELSE '' END AS VALUE_PART2 " & _
                "from METHDBA.TTR_RG_EFF_DATE " & _
                 BuildWhereClause("WHERE", "RG_CD", v_varRules, False)
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetEffectiveDateRules = l_objRS
    
    Set l_objRS = Nothing

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEffectiveDateRules", , , , DescArgsFrom(l_strSQL), True
End Function
