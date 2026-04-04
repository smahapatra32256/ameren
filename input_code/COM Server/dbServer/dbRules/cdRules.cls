VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdRules"

'Each Rule will have a separte cache group
'All rules are a simple look up for existence of a value in the rule table
'The key will contain 2 characters for data type and the remaining is the value
'Please see mTRISContstance for the data type defintitions


' Rule constants
'The Key is 2 characters for the data type and the remaining is the value
'Private Const CRG_KEY = 0

' Department constants
'Private Const CRG_GROUPNAME = "RULExxxxCACHE"
' index for record dynamic array
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
'Private Const CRG_RG_CD = 1
'Private Const CRG_RG_SUB_CD = 2
'Private Const CRG_BUS_TYPE_CD = 3
'Private Const CRG_RG_VALUE_CD = 4

Private Const CRG_RG_CD = 1
Private Const CRG_RG_SUB_CD = 2
Private Const CRG_BUS_TYPE_CD = 3
Private Const CRG_RG_VALUE_CD = 4
Private Const CRG_SEQ_NUM = 5
Private Const CRG_VALUE_PART1 = 6
Private Const CRG_VALUE_PART2 = 7

'also used by cdEffectiveDateRules, when compiled seems to be defined at the dbRules level (typelib)
Public Enum TRSRulesDataTypes
    trExceptionCode = 1
    trEmployeeNumber = 2
    trBenPkgInd = 3
    trDepartmentNumber = 4
    trCompanyCode = 5
    trUnionCode = 6
    trJobNumber = 7
    trLocationCode = 8
    trPayrollGroupCode = 23
    trDeptNumAndUnionCode = 75
End Enum

Public Enum TRSRuleCode
    trIsAmerenCIPS = 11
    trIsAmerenCILCO = 12
    trAICDeptMapping = 203
    trAICEffectiveDates = 206
    trIsSubstationOrTransmissionEmployee = 207
    trIsCIPSForSickLeaveBanks = 209
    trAICDeptCutOverEffDate = 211
End Enum

Private Enum CacheGroupType
    cgtExistence = 0
    cgtLookup
    cgtReverseLookup
End Enum
'Please note : RG_SUB_CD is not being used in the rule cache
' Currently the only use of RG_SUB_CD is for the coloring of View Calendar on the GUI


Public Function GetRule( _
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
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN   SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1) " & _
                    "  ELSE '' END AS VALUE_PART1 " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN   SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD)) " & _
                    "  ELSE '' END AS VALUE_PART2 " & _
                "from METHDBA.TTR_RG_GROUP " & _
                "where RG_CD = " & str(v_lngRule)
    If v_strSubRule <> "" Then
         l_strSQL = l_strSQL & " and RG_SUB_CD = '" & v_strSubRule & "'"
    End If
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetRule = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetRule", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetRules(ByVal v_varRules As Variant) As ADODB.Recordset
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
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN   SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1) " & _
                    "  ELSE '' END AS VALUE_PART1 " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN   SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1) " & _
                    "  ELSE '' END AS VALUE_PART2 " & _
                "from METHDBA.TTR_RG_GROUP " & _
                BuildWhereClause("WHERE", "RG_CD", v_varRules, False) & _
                "ORDER BY RG_CD, RG_SUB_CD, BUS_TYPE_CD, RG_VALUE_CD"
        
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetRules = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetRules", , , , DescArgsFrom(l_strSQL), True
End Function

Public Function DoesRuleApply( _
            ByVal v_lngRule As Long, _
            ByVal v_lngDataType As TRSRulesDataTypes, _
            ByVal v_strValue As String) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim l_strGroupName As String
    
On Error GoTo eh
    l_strGroupName = GetRuleGroupName(v_lngRule, cgtExistence)
    l_varRecord = GetRecord(Array(v_lngRule, v_lngDataType, v_strValue), l_strGroupName, lblnExists)
    DoesRuleApply = Not IsEmpty(l_varRecord)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DoesRuleApply"
End Function
    
Public Function GetRuleLookup( _
            ByVal v_lngRule As Long, _
            ByVal v_lngDataType As TRSRulesDataTypes, _
            ByVal v_strValuePart1 As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim l_strGroupName As String
    
On Error GoTo eh
    l_strGroupName = GetRuleGroupName(v_lngRule, cgtLookup)
    l_varRecord = GetRecord(Array(v_lngRule, v_lngDataType, v_strValuePart1), l_strGroupName, lblnExists)
    If Not IsEmpty(l_varRecord) Then
        GetRuleLookup = l_varRecord(CRG_VALUE_PART2)
    Else
        GetRuleLookup = ""
    End If
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetRuleLookup"
End Function

Public Function GetRuleReverseLookup( _
            ByVal v_lngRule As Long, _
            ByVal v_lngDataType As TRSRulesDataTypes, _
            ByVal v_strValuePart2 As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim l_strGroupName As String
    
On Error GoTo eh
    l_strGroupName = GetRuleGroupName(v_lngRule, cgtReverseLookup)
    l_varRecord = GetRecord(Array(v_lngRule, v_lngDataType, v_strValuePart2), l_strGroupName, lblnExists)
    If Not IsEmpty(l_varRecord) Then
        GetRuleReverseLookup = l_varRecord(CRG_VALUE_PART1)
    Else
        GetRuleReverseLookup = ""
    End If
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetRuleReverseLookup"
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
    Dim l_lngRule As Long
    Dim l_eCacheGroupType As CacheGroupType
On Error GoTo eh

    GetRuleNumFromGroupName v_strGroupName, l_lngRule, l_eCacheGroupType

    Select Case l_eCacheGroupType
    Case cgtExistence
        Set l_rsRecords = GetRuleCacheRecords(v_strGroupName, v_varKeys)
        r_varKeyArray = Array("RG_CD", "BUS_TYPE_CD", "RG_VALUE_CD")
        r_varColumnArray = Empty
    Case cgtLookup
        Set l_rsRecords = GetRuleCacheRecordsLookup(v_strGroupName, v_varKeys)
        r_varKeyArray = Array("RG_CD", "BUS_TYPE_CD", "VALUE_PART1") '//lookup
        r_varColumnArray = Empty
    Case cgtReverseLookup
        Set l_rsRecords = GetRuleCacheRecordsLookup(v_strGroupName, v_varKeys)
        r_varKeyArray = Array("RG_CD", "BUS_TYPE_CD", "VALUE_PART2") '//Reverse lookup
        r_varColumnArray = Empty
    Case Else
        Set l_rsRecords = GetRuleCacheRecords(v_strGroupName, v_varKeys)
        r_varKeyArray = Array("RG_CD", "BUS_TYPE_CD", "RG_VALUE_CD")
        r_varColumnArray = Empty
    End Select

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
    Dim l_eCacheGroupType As CacheGroupType
On Error GoTo eh
    
    If IsMissing(v_varKeys) Or IsEmpty(v_varKeys) Then '//-----Bulk load for entire cache group-----//
        GetRuleNumFromGroupName v_strGroupName, l_lngRule, l_eCacheGroupType
        l_strWhereClause = " WHERE RG_CD = " & l_lngRule
    ElseIf IsArray(v_varKeys) Then '//-----single row add for invalidate-----//
        If UBound(v_varKeys) - LBound(v_varKeys) = 2 Then
            l_strWhereClause = BuildWhereClauseForMultiKey(Array("WHERE", "AND", "AND"), _
                                                    Array("RG_CD", "BUS_TYPE_CD", "RG_VALUE_CD"), _
                                                    v_varKeys, _
                                                    Array(False, False, True))
        Else
            Err.Raise trErrdbGetCacheSQLKey
        End If
    Else
        l_strWhereClause = " WHERE RG_CD = " & v_varKeys
    End If
    
    l_strSQL = "SELECT RG_CD, RG_SUB_CD, BUS_TYPE_CD, RG_VALUE_CD " & _
                " FROM METHDBA.TTR_RG_GROUP " & _
                l_strWhereClause
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetRuleCacheRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetRuleCacheRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'//for lookups needing to convert one value (value_part1) to another (value_part2) (or vice versa), rather than existence check
Private Function GetRuleCacheRecordsLookup(ByVal v_strGroupName As String, _
                                    Optional ByVal v_varKeys As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strWhereClause As String
    Dim l_lngRule As Long
    Dim l_eCacheGroupType As CacheGroupType
On Error GoTo eh
    
    GetRuleNumFromGroupName v_strGroupName, l_lngRule, l_eCacheGroupType
    If IsMissing(v_varKeys) Or IsEmpty(v_varKeys) Then '//-----Bulk load for entire cache group-----//
        l_strWhereClause = " WHERE RG_CD = " & l_lngRule
    ElseIf IsArray(v_varKeys) Then '//-----single row add for invalidate-----//
        If UBound(v_varKeys) - LBound(v_varKeys) = 2 Then
            Dim l_strValueColumn As String
            If l_eCacheGroupType = cgtReverseLookup Then
                l_strValueColumn = "VALUE_PART2" '//reverse lookup
            ElseIf l_eCacheGroupType = cgtLookup Then
                l_strValueColumn = "VALUE_PART1" '// lookup
            Else
                l_strValueColumn = "VALUE_PART1" '//lookup
            End If
            
            l_strWhereClause = BuildWhereClauseForMultiKey(Array("WHERE", "AND", "AND"), _
                                                    Array("RG_CD", "BUS_TYPE_CD", l_strValueColumn), _
                                                    v_varKeys, _
                                                    Array(False, False, True))
        Else
            Err.Raise trErrdbGetCacheSQLKey
        End If
    Else
        l_strWhereClause = " WHERE RG_CD = " & v_varKeys
    End If
    
    l_strSQL = "SELECT * FROM (SELECT RG_CD, RG_SUB_CD, BUS_TYPE_CD, RG_VALUE_CD, SEQ_NUM " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN   SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1) " & _
                    "  ELSE '' END AS VALUE_PART1 " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN   SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD)-LOCATE(':',RG_VALUE_CD,1)) " & _
                    "  ELSE '' END AS VALUE_PART2 " & _
                "FROM METHDBA.TTR_RG_GROUP ) RG " & _
                l_strWhereClause
    
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetRuleCacheRecordsLookup = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetRuleCacheRecordsLookup", , , , DescArgsFrom(l_strSQL), True
End Function

''*********************************************************
'' Begin Private Caching methods
''*********************************************************
'
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

Public Sub InvalidateRulesCache(ByVal v_lngRule As Long, _
                                ByVal v_lngBusTypeCd As Long, _
                                ByVal v_strRgValueCd As String)
    Dim l_strGroupName As String
    Dim recCache As utilRecordCache.cuRecordCache
On Error GoTo eh
   
    Set recCache = CreateObject("utilRecordCache.cuRecordCache")
       
    l_strGroupName = GetRuleGroupName(v_lngRule, cgtExistence)
    If recCache.DoesGroupExist(l_strGroupName) Then
        InvalidateCache Array(v_lngRule, v_lngBusTypeCd, v_strRgValueCd), l_strGroupName
    End If
    
    l_strGroupName = GetRuleGroupName(v_lngRule, cgtLookup)
    If recCache.DoesGroupExist(l_strGroupName) Then
        InvalidateCache Array(v_lngRule, v_lngBusTypeCd, v_strRgValueCd), l_strGroupName
    End If
    
    l_strGroupName = GetRuleGroupName(v_lngRule, cgtReverseLookup)
    If recCache.DoesGroupExist(l_strGroupName) Then
        InvalidateCache Array(v_lngRule, v_lngBusTypeCd, v_strRgValueCd), l_strGroupName
    End If
    
    Set recCache = Nothing
    Exit Sub
eh:
    Set recCache = Nothing
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
        OneRowForEachKey = True
'    Case Else
'        OneRowForEachKey = False
'    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OneRowForEachKey"
End Function

'Private Function GetCacheRecords(ByVal v_strSQL As String, Optional eLockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
'    Dim l_objCN As New ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'On Error GoTo eh
'
'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If
'
'    Set l_objCN = New ADODB.Connection
'    Set l_objRS = New ADODB.Recordset
'    Set l_objCM = New clsuDBConnectMgr
'    l_objCN.ConnectionString = l_objCM.GetConnectionString
'    Set l_objCM = Nothing
'    l_objCN.Open
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open v_strSQL, l_objCN, adOpenStatic, eLockType
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'    Set l_objCN = Nothing
'
''//-----the cache needs the meta data for formatting the key
''    If l_objRS.BOF And l_objRS.EOF Then
''        Set l_objRS = Nothing
''    End If
'
'    Set GetCacheRecords = l_objRS
'
'    Set l_objRS = Nothing
'    Set l_objCN = Nothing
'    Set l_objCM = Nothing
'
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "GetCacheRecords: " & CStr(dfTime / 1000)
'    #End If
'
'    Exit Function
'
'eh:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetCacheRecords", , , , DescArgsFrom(v_strSQL), True
'End Function
    
Private Function GetRuleGroupName(ByVal v_lngRule As Long, ByVal v_CacheGroupType As CacheGroupType) As String
    Dim l_strRule As String
    Dim l_str As String
    Dim l_strResult As String
    Dim l_strGroupNameSuffix As String
On Error GoTo eh

    l_str = Trim(str(v_lngRule))
    l_strRule = String(4 - Len(l_str), "0") & l_str
    
    Select Case v_CacheGroupType
    Case CacheGroupType.cgtExistence
        l_strGroupNameSuffix = ""
    Case CacheGroupType.cgtLookup
        l_strGroupNameSuffix = "LOOKUP"
    Case CacheGroupType.cgtReverseLookup
        l_strGroupNameSuffix = "REVERSELOOKUP"
    Case Else
        l_strGroupNameSuffix = ""
    End Select
    
    l_strResult = "RULE" & l_strRule & "CACHE" & l_strGroupNameSuffix

    GetRuleGroupName = l_strResult
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetRuleGroupName"
End Function
'    |    1   |      2
'123456789012345467890
'RULE0203CACHE
'RULE0203CACHELOOKUP
'RULE0203CACHEREVERSELOOKUP
Private Sub GetRuleNumFromGroupName(ByVal v_strGroupName As String, ByRef r_lngRuleNum As Long, ByRef r_eCacheGroupType As CacheGroupType)
    Dim l_str As String
    Dim l_strCacheGroupType As String
On Error GoTo eh

    l_str = Trim(Mid(v_strGroupName, 5, 4))
    r_lngRuleNum = CLng(l_str)
   
    l_strCacheGroupType = Mid(v_strGroupName, 14)
    Select Case UCase(l_strCacheGroupType)
    Case ""
        r_eCacheGroupType = cgtExistence
    Case "LOOKUP"
       r_eCacheGroupType = cgtLookup
    Case "REVERSELOOKUP"
        r_eCacheGroupType = cgtReverseLookup
    Case Else
        r_eCacheGroupType = cgtExistence
    End Select
    
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "GetRuleNumFromGroupName"
End Sub

''//-----copied from dbCodeTables
'Private Function BuildWhereClause(ByVal v_strSQLKeyWord As String, _
'                                    ByVal v_strColumn As String, _
'                                    ByVal v_varColumnValues As Variant, _
'                                    ByVal v_blnDelimit As Boolean) As String
'    Dim l_strWhereClause As String
'    Dim l_colErrArgs As Collection
'On Error GoTo eh
'
'    If IsMissing(v_varColumnValues) Or IsEmpty(v_varColumnValues) Then
'        l_strWhereClause = ""
'    ElseIf IsArray(v_varColumnValues) Then
'        l_strWhereClause = " " & v_strSQLKeyWord & " " & BuildInClause(v_strColumn, v_varColumnValues, v_blnDelimit)
'    Else
'        l_strWhereClause = " " & v_strSQLKeyWord & " " & v_strColumn & " = " & _
'            IIf(v_blnDelimit, StringEnclosedBy(Trim(CStr(v_varColumnValues)), "'"), Trim(CStr(v_varColumnValues)))
'    End If
'
'    BuildWhereClause = l_strWhereClause
'    Exit Function
'eh:
'    Set l_colErrArgs = New Collection
'    l_colErrArgs.Add "v_strSQLKeyWord = " & v_strSQLKeyWord
'    l_colErrArgs.Add "v_strColumn = " & v_strColumn
'    If IsMissing(v_varColumnValues) Or IsEmpty(v_varColumnValues) Then
'        l_colErrArgs.Add "v_varColumnValues = empty"
'    ElseIf IsArray(v_varColumnValues) Then
'        l_colErrArgs.Add "v_varColumnValues() = (" & Join(v_varColumnValues, ", ") & ")"
'    Else
'        l_colErrArgs.Add "v_varColumnValues = " & v_varColumnValues
'    End If
'    l_colErrArgs.Add "v_blnDelimit = " & v_blnDelimit
'
'    TrsRaiseError m_strCMODULENAME, "BuildWhereClause", , , , l_colErrArgs, True
'End Function
'
'Private Function BuildWhereClauseForMultiKey(ByVal v_strSQLKeyWord As Variant, _
'                                    ByVal v_strColumn As Variant, _
'                                    ByVal v_varColumnValues As Variant, _
'                                    ByVal v_blnDelimit As Variant) As Variant
'    Dim l_strWhereClause As String
'    Dim l_colErrArgs As Collection
'    Dim nKey As Long
'On Error GoTo eh
'
'    If IsMissing(v_varColumnValues) Or IsEmpty(v_varColumnValues) Then
'        l_strWhereClause = ""
'    Else
'        If UBound(v_strSQLKeyWord) = UBound(v_strSQLKeyWord) And _
'            UBound(v_strSQLKeyWord) = UBound(v_strColumn) And _
'            UBound(v_strSQLKeyWord) = UBound(v_varColumnValues) And _
'            UBound(v_strSQLKeyWord) = UBound(v_blnDelimit) Then
'            l_strWhereClause = ""
'            For nKey = LBound(v_strSQLKeyWord) To UBound(v_strSQLKeyWord)
'                l_strWhereClause = l_strWhereClause & BuildWhereClause(v_strSQLKeyWord(nKey), v_strColumn(nKey), v_varColumnValues(nKey), v_blnDelimit(nKey)) & " "
'            Next nKey
'        Else
'        End If
'    End If
'
'    BuildWhereClauseForMultiKey = l_strWhereClause
'    Exit Function
'eh:
'    Set l_colErrArgs = New Collection
'
'    AddToErrorCollection l_colErrArgs, v_strSQLKeyWord, "v_strSQLKeyWord"
'    AddToErrorCollection l_colErrArgs, v_strColumn, "v_strColumn"
'    AddToErrorCollection l_colErrArgs, v_varColumnValues, "v_varColumnValues"
'    AddToErrorCollection l_colErrArgs, v_blnDelimit, "v_blnDelimit"
'
'    TrsRaiseError m_strCMODULENAME, "BuildWhereClauseForMultiKey", , , , l_colErrArgs, True
'End Function
'
'Private Function AddToErrorCollection(ByRef r_colErrArgs As Collection, _
'                                        ByRef v_varArgs As Variant, _
'                                        ByVal v_strArgName As String)
'    If IsMissing(v_varArgs) Or IsEmpty(v_varArgs) Then
'        r_colErrArgs.Add v_strArgName & " = empty"
'    ElseIf IsArray(v_varArgs) Then
'        r_colErrArgs.Add v_strArgName & "() = (" & Join(v_varArgs, ", ") & ")"
'    Else
'        r_colErrArgs.Add v_strArgName & " = " & v_varArgs
'    End If
'End Function

'moved to cdEffectiveDateRules.cls
Public Function GetEffectiveDateRule( _
    ByVal v_lngRule As Long _
    ) As ADODB.Recordset
    
    Dim o As dbRules.cdEffectiveDateRules
    
    On Error GoTo eh
    
    Set o = CreateObject("dbRules.cdEffectiveDateRules")
    
    Set GetEffectiveDateRule = o.GetEffectiveDateRule(v_lngRule)
    
    Set o = Nothing
    Exit Function
eh:
    Set o = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEffectiveDateRule", , , , DescArgsFrom(CStr(v_lngRule)), True
End Function
