VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cuRecordCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Tim,
'   I wished I made more progress, but I ran out of time.  Here a a few notes.
'
'   GetRecord is now public and includes an argument for a safe reference back to the
'   caller, r_objCaller.
'
'
'   I started generalizing InitializeGroup and LoadRecord.
'   The design depends on a callback to the dataservice for record retrieval.
'   Once all is generalized in cuRecordCache, you should be free from the
'   redundant code in cdCodeTables, workgroup, and whatever else has GetRecord.
'
'   As you can see, the client is required to implement GetGroupRecords.  I have
'   implemented this method on cdCodeTables and only made changes to retrieve
'   exception code records.
'   GetGroupRecords should also be flexible enough to handle specific
'   retrieval requests by accepting key information that can be transformed
'   into a WHERE clause.
'
'   I did not have much time to work on the composite keys, but I took a first cut
'   at it.  The caller is responsible for passing an array of field names that must
'   match the column names in the recordset.  For exception codes, I implmemented
'   ExceptionGroupKeys, which only had 1 component.
'
'   I also changed the implementation of GetRecord in cdCodeTables to factor out
'   the instanciation of cuRecordCache.
'
'   I did not really deal with the constants yet.  Ideally, this can be factored
'   out of cuRecordCache.  I did change the exception code constants by making the
'   0 index reserved for the key.  This is unnecessary for single component key, but
'   it seemed like an easy way to generalize it for composite keys.
'
'   Step through the code asking for an exception code attribute and it should
'   be clear enough as to what I was trying to do.
'
'
'
Private Const m_strCMODULENAME As String = "cuRecordCache"

Private Const CCACHEMETADATA As String = "CACHEMETADATA"

'' index for record dynamic array
Private Const CMD_GROUP = 0
Private Const CMD_KEYCOLUMN = 1
Private Const CMD_KEYDATATYPE = 2
Private Const CMD_KEYSIZE = 3
Private Const CMD_LAST_SLOT = 3

Private Declare Function GetTickCount Lib "kernel32" () As Long

' Static data constants
Private Const CRECORDNOTFOUND = "RECORDNOTFOUND"
Private Const CSPINVALID = "SPINVALID"
Private Const CDEFAULTDATE As String = "01/01/0100"

'//-----uses already formated key-----//
Public Function GetRecord( _
            ByVal v_strKey As String, _
            ByVal v_strGroupName As String, _
            ByVal v_blnLazyInitialized As Boolean, _
            ByVal v_blnSupportsInvalidate As Boolean, _
            ByVal v_blnOneRowForEachKey As Boolean, _
            ByRef r_blnExists As Boolean, _
            ByRef r_objCaller As Object) As Variant

    Dim l_colErrArgs As Collection
On Error GoTo eh
    GetRecord = GetRecord3(v_strKey, v_strKey, v_strGroupName, _
                v_blnLazyInitialized, v_blnSupportsInvalidate, v_blnOneRowForEachKey, _
                r_blnExists, r_objCaller)
    
    Exit Function

eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strKey = """ & v_strKey & """"
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    l_colErrArgs.Add "v_blnLazyInitialized = " & v_blnLazyInitialized
    l_colErrArgs.Add "v_blnSupportsInvalidate = " & v_blnSupportsInvalidate
    l_colErrArgs.Add "v_blnOneRowForEachKey = " & v_blnOneRowForEachKey
'    l_colErrArgs.Add "r_blnExists = " & r_blnExists
'    l_colErrArgs.Add "r_objCaller = " & r_objCaller
    If Not r_objCaller Is Nothing Then
        l_colErrArgs.Add "TypeName(r_objCaller) = " & TypeName(r_objCaller)
    Else
        l_colErrArgs.Add "r_objCaller = Nothing"
    End If

    TrsRaiseError m_strCMODULENAME, "GetRecord", , , , l_colErrArgs, True
End Function

'//-----v_strKey is for SPM Cache
'//-----v_varKey will be passed for SQL retrieval (Get<CacheGroup>Records)
Public Function GetRecord3( _
            ByVal v_strKey As String, _
            ByVal v_varKey As Variant, _
            ByVal v_strGroupName As String, _
            ByVal v_blnLazyInitialized As Boolean, _
            ByVal v_blnSupportsInvalidate As Boolean, _
            ByVal v_blnOneRowForEachKey As Boolean, _
            ByRef r_blnExists As Boolean, _
            ByRef r_objCaller As Object) As Variant
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
' Only use r_objCaller for callbacks to retrieve data.
'

    Dim l_varRecord As Variant
    Dim spmWrapper As utilSPMCache.cuSPMCache
    Dim lblnExists As Boolean
    Dim lblnGroupExists As Boolean
    Dim l_colErrArgs As Collection
On Error GoTo eh
    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
    l_varRecord = spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
    If Not lblnGroupExists Then '//-----Group has not been loaded in cache-----//
        If v_blnLazyInitialized Then '//-----Load on demand (one at a time)-----//
            LoadRecord v_strKey, v_varKey, v_strGroupName, v_blnOneRowForEachKey, r_objCaller
            l_varRecord = spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
        Else '//-----Load All (Bulk)-----//
            InitializeGroup v_strGroupName, v_blnOneRowForEachKey, r_objCaller
            l_varRecord = spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
        End If
    Else '//-----Group has been loaded in cache-----//
        If Not r_blnExists Then '//-----Key is not found-----//
            If v_blnLazyInitialized Then '//-----Load on demand (one at a time)-----//
                LoadRecord v_strKey, v_varKey, v_strGroupName, v_blnOneRowForEachKey, r_objCaller
                l_varRecord = spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
            End If
        Else    'it exists and INVALID - reload it
                '***There must be a Load method for every cache that supports Invalidate
            If Not IsArray(l_varRecord) Then
                If l_varRecord = CSPINVALID And v_blnSupportsInvalidate Then '//-----Supports Invalidate (load one)-----//
                    LoadRecord v_strKey, v_varKey, v_strGroupName, v_blnOneRowForEachKey, r_objCaller
                    l_varRecord = spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
                End If
            End If
        End If
    End If
    ' test for array first since 99.9% of time it should be one,
    ' otherwise answer empty.
    If IsArray(l_varRecord) Then
        GetRecord3 = l_varRecord
    Else
        GetRecord3 = Empty
    End If
    Set spmWrapper = Nothing
    Exit Function

eh:
    Set spmWrapper = Nothing
    
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strKey = """ & v_strKey & """"
    If IsMissing(v_varKey) Or IsEmpty(v_varKey) Then
        l_colErrArgs.Add "v_varKey = empty"
    ElseIf IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & v_varKey
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    l_colErrArgs.Add "v_blnLazyInitialized = " & v_blnLazyInitialized
    l_colErrArgs.Add "v_blnSupportsInvalidate = " & v_blnSupportsInvalidate
    l_colErrArgs.Add "v_blnOneRowForEachKey = " & v_blnOneRowForEachKey
'    l_colErrArgs.Add "r_objCaller = " & r_objCaller
    If Not r_objCaller Is Nothing Then
        l_colErrArgs.Add "TypeName(r_objCaller) = " & TypeName(r_objCaller)
    Else
        l_colErrArgs.Add "r_objCaller = Nothing"
    End If
    l_colErrArgs.Add "lblnGroupExists = " & lblnGroupExists
    l_colErrArgs.Add "r_blnExists = " & r_blnExists

    TrsRaiseError m_strCMODULENAME, "GetRecord3", , , , l_colErrArgs, True
End Function

'//-----formats key so that it is consistent with (BuildCacheKey)-----//
Public Function GetRecord2( _
            ByVal v_varKey As Variant, _
            ByVal v_strGroupName As String, _
            ByVal v_blnLazyInitialized As Boolean, _
            ByVal v_blnSupportsInvalidate As Boolean, _
            ByVal v_blnOneRowForEachKey As Boolean, _
            ByRef r_blnExists As Boolean, _
            ByRef r_objCaller As Object) As Variant

    Dim l_strKey As String
    Dim l_varRecord As Variant
    Dim spmWrapper As utilSPMCache.cuSPMCache
    Dim l_blnMetaDataExists As Boolean
    Dim l_blnMetaDataGroupExists As Boolean
    Dim lblnGroupExists As Boolean
    Dim l_colErrArgs As Collection
On Error GoTo eh

'//-----this won't work for the 1st user requesting the cache (chicken/egg syndrome)
'//    l_strKey = FormatKeyFromMetaData(v_strGroupName, v_varKey)
'//    GetRecord2 = GetRecord(l_strKey, v_strGroupName, v_blnLazyInitialized, v_blnSupportsInvalidate, v_blnOneRowForEachKey, r_blnExists, r_objCaller)
               
    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
    l_varRecord = spmWrapper.Item(v_strGroupName, CCACHEMETADATA, l_blnMetaDataExists, l_blnMetaDataGroupExists)
    If Not l_blnMetaDataGroupExists Or Not l_blnMetaDataExists Then '//-----Meta Data has not been loaded in cache-----//
        If v_blnLazyInitialized Then '//-----Load on demand (one at a time)-----//
'//-----currently this would only apply to JobNumbers and WorkGroups, but it should use GetRecord not this one-----//
            Err.Raise trErrRecordCacheNoMetaData
'            LoadRecord v_strKey, v_strGroupName, v_blnOneRowForEachKey, r_objCaller
'
'            l_strKey = FormatKeyFromMetaData(v_strGroupName, v_varKey)
'
'            l_varRecord = spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, lblnGroupExists)
        Else '//-----Load All (Bulk)-----//
            InitializeGroup v_strGroupName, v_blnOneRowForEachKey, r_objCaller
        End If
    End If
    
    l_strKey = FormatKeyFromMetaData(v_strGroupName, v_varKey)
    GetRecord2 = GetRecord3(l_strKey, v_varKey, v_strGroupName, v_blnLazyInitialized, v_blnSupportsInvalidate, v_blnOneRowForEachKey, r_blnExists, r_objCaller)
    
    Set spmWrapper = Nothing
    Exit Function
eh:
    Set spmWrapper = Nothing
    
    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = """ & CStr(v_varKey) & """"
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    l_colErrArgs.Add "v_blnLazyInitialized = " & v_blnLazyInitialized
    l_colErrArgs.Add "v_blnSupportsInvalidate = " & v_blnSupportsInvalidate
    l_colErrArgs.Add "v_blnOneRowForEachKey = " & v_blnOneRowForEachKey
'    l_colErrArgs.Add "r_objCaller = " & r_objCaller
    
    l_colErrArgs.Add "l_blnMetaDataGroupExists = " & l_blnMetaDataGroupExists
    l_colErrArgs.Add "l_blnMetaDataExists = " & l_blnMetaDataExists
    l_colErrArgs.Add "l_strKey = """ & l_strKey & """"
    
    Select Case Err.Number
    Case trErrRecordCacheNoMetaData
        TrsRaiseError m_strCMODULENAME, "GetRecord2", Err.Number, , DescArgsFrom(v_strGroupName), l_colErrArgs, True
    Case Else
        TrsRaiseError m_strCMODULENAME, "GetRecord2", , , , l_colErrArgs, True
    End Select
    
End Function

Public Function DoesGroupExist(ByVal v_strGroupName As String) As Boolean
    Dim spmWrapper As utilSPMCache.cuSPMCache
    Dim l_colErrArgs As Collection
On Error GoTo eh

    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
    
    DoesGroupExist = spmWrapper.GroupExists(v_strGroupName)
    Set spmWrapper = Nothing
    
    Exit Function
eh:
    Set spmWrapper = Nothing
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    TrsRaiseError m_strCMODULENAME, "DoesGroupExist", , , , l_colErrArgs, True
End Function

Public Sub InvalidateCache( _
                ByVal v_strKey As String, _
                ByVal v_strGroupName As String)
' RSG 11/07/02 - Invalidate all servers in cluster.
    Dim spmProxy As utilSPMCacheProxy.cuSPMCacheProxy
    Dim l_strKey As String
    Dim l_strThisComputerName As String
    Dim l_varAllServers As Variant
    Dim i As Integer
    Dim l_colErrArgs As Collection
On Error GoTo eh

    l_strThisComputerName = Trim(UCase(ThisComputerName))
    l_varAllServers = GetAllServerNames(l_strThisComputerName)
    For i = LBound(l_varAllServers) To UBound(l_varAllServers)
        If l_varAllServers(i) = l_strThisComputerName Then
            Set spmProxy = CreateObject("utilSPMCacheProxy.cuSPMCacheProxy")
        Else
            Set spmProxy = CreateObject("utilSPMCacheProxy.cuSPMCacheProxy", l_varAllServers(i))
        End If
        l_strKey = Trim(v_strKey)
        spmProxy.Invalidate l_strKey, v_strGroupName
        Set spmProxy = Nothing
    Next
    Exit Sub
    
eh:
    Set spmProxy = Nothing
    
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strKey = """ & v_strKey & """"
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    l_colErrArgs.Add "l_varAllServers() = (" & Join(l_varAllServers, ", ") & ")"

    TrsRaiseError m_strCMODULENAME, "InvalidateCache", , , , l_colErrArgs, True
End Sub

'//-----for simplicity use same signature as GetRecord2
'//-----v_blnSupportsInvalidate and v_blnSupportsInvalidate aren't used
Public Sub InvalidateCache2( _
            ByVal v_varKey As Variant, _
            ByVal v_strGroupName As String, _
            ByVal v_blnLazyInitialized As Boolean, _
            ByVal v_blnSupportsInvalidate As Boolean, _
            ByVal v_blnOneRowForEachKey As Boolean, _
            ByRef r_blnExists As Boolean, _
            ByRef r_objCaller As Object)
    Dim l_strKey As String
    Dim l_varRecord As Variant
    Dim spmWrapper As utilSPMCache.cuSPMCache
    Dim l_blnMetaDataExists As Boolean, l_blnMetaDataGroupExists As Boolean
    Dim l_blnExists As Boolean
    Dim l_colErrArgs As Collection
On Error GoTo eh

    '//-----there is a chance that a request for an invalidate will ocur before a GetRecord-----//
    '//-----since we don't know the formatting for the key, we could check for the existence of the group
    '//-----if the group does not exist, then it doesn't need invalidated, no error
    '//-----But, what about the other servers in the cluster?
    '//-----basically, do a GetRecord2 to initialize the metadata and continue

    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
    l_varRecord = spmWrapper.Item(v_strGroupName, CCACHEMETADATA, l_blnMetaDataExists, l_blnMetaDataGroupExists)
    If Not l_blnMetaDataGroupExists Or Not l_blnMetaDataExists Then '//-----Meta Data has not been loaded in cache-----//
        If v_blnLazyInitialized Then '//-----Load on demand (one at a time)-----//
'//-----currently this would only apply to JobNumbers and WorkGroups, but it should use GetRecord/InvalidateCache not this one-----//
            Err.Raise trErrRecordCacheNoMetaData
        Else '//-----Load All (Bulk)-----//
            InitializeGroup v_strGroupName, v_blnOneRowForEachKey, r_objCaller
        End If
    End If
    
    l_strKey = FormatKeyFromMetaData(v_strGroupName, v_varKey)
    InvalidateCache l_strKey, v_strGroupName
        
    Set spmWrapper = Nothing
    Exit Sub
eh:
    Set spmWrapper = Nothing
    
    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = """ & CStr(v_varKey) & """"
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName

    l_colErrArgs.Add "l_blnMetaDataGroupExists = " & l_blnMetaDataGroupExists
    l_colErrArgs.Add "l_blnMetaDataExists = " & l_blnMetaDataExists
    l_colErrArgs.Add "l_strKey = " & l_strKey
    
    Select Case Err.Number
    Case trErrRecordCacheNoMetaData
        TrsRaiseError m_strCMODULENAME, "InvalidateCache2", Err.Number, , DescArgsFrom(v_strGroupName), l_colErrArgs, True
    Case Else
        TrsRaiseError m_strCMODULENAME, "InvalidateCache2", , , , l_colErrArgs, True
    End Select

End Sub

'//-----key support must be enabled for this to work
Public Function GetKeys(ByVal v_strGroupName As String) As Variant
    Dim spmWrapper As utilSPMCache.cuSPMCache
    
    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
    
    GetKeys = spmWrapper.Keys(v_strGroupName)
    
    Set spmWrapper = Nothing
End Function

Public Function IsInvalidated(ByVal v_strKey As String, _
                                ByVal v_strGroupName As String) As Boolean
    Dim spmWrapper As utilSPMCache.cuSPMCache
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean, lblnGroupExists As Boolean
    
    IsInvalidated = False
    
    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
    
    l_varRecord = spmWrapper.Item(v_strKey, v_strGroupName, lblnExists, lblnGroupExists)
    If Not lblnGroupExists Then '//-----Group has not been loaded in cache-----//
    
    Else '//-----Group has been loaded in cache-----//
        If Not lblnExists Then '//-----Key is not found-----//
        Else    'it exists and INVALID - reload it
            If Not IsArray(l_varRecord) Then
                If l_varRecord = CSPINVALID Then '//-----Supports Invalidate (load one)-----//
                    IsInvalidated = True
                End If
            End If
        End If
    End If
    Set spmWrapper = Nothing
End Function

Public Function IsInvalidated2(ByVal v_varKey As Variant, _
                                ByVal v_strGroupName As String) As Boolean
                                
    Dim l_strKey As String
    Dim l_varRecord As Variant
    Dim spmWrapper As utilSPMCache.cuSPMCache
    Dim l_blnMetaDataExists As Boolean, l_blnMetaDataGroupExists As Boolean
    Dim l_blnExists As Boolean
'    Dim l_colErrArgs As Collection
    
    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
    l_varRecord = spmWrapper.Item(v_strGroupName, CCACHEMETADATA, l_blnMetaDataExists, l_blnMetaDataGroupExists)
    If Not l_blnMetaDataGroupExists Or Not l_blnMetaDataExists Then '//-----Meta Data has not been loaded in cache-----//
        IsInvalidated2 = False
        Exit Function
    End If
    
    l_strKey = FormatKeyFromMetaData(v_strGroupName, v_varKey)
    IsInvalidated2 = IsInvalidated(l_strKey, v_strGroupName)
                                
End Function

'//-----this could be factored into some other util object for reusablity-----//
Private Function GetAllServerNames(ByVal v_strComputerName As String) As Variant
' Answer an array.
' Get all available servers in the same cluster as v_strComputerName,
' including v_strComputerName.  If GetServersInMyCluster answers empty then
' v_strComputerName is a developer workstation, just add it to the array.

'late bind to prevent near circular reference
'   dbCodeTables calls utilRecordCache
'   utilRecordCache calls utilServerInfo which calls dbCodeTables
    Dim oServerInfo As Object 'utilServerInfo.cuServerInfo
    Dim l_varAllServers As Variant

On Error GoTo eh

    Set oServerInfo = CreateObject("utilServerInfo.cuServerInfo")
    l_varAllServers = oServerInfo.GetServersInMyCluster(v_strComputerName)
    If IsEmpty(l_varAllServers) Then
        ReDim l_varAllServers(0)
        l_varAllServers(0) = v_strComputerName
    End If
    GetAllServerNames = l_varAllServers
    Set oServerInfo = Nothing
    Exit Function

eh:
    Set oServerInfo = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllServerNames"
End Function


Private Function InitializeGroup( _
            ByVal v_strGroupName As String, _
            ByVal v_blnOneRowForEachKey As Boolean, _
            ByRef r_objCaller As Object) As Variant
    Dim l_rsGroupRecords As ADODB.Recordset
    Dim l_varKeyArray As Variant, l_varColumnsArray As Variant
            
On Error GoTo eh

    Set l_rsGroupRecords = r_objCaller.GetGroupRecords(v_strGroupName, Empty, l_varKeyArray, l_varColumnsArray)
    LoadRecordset v_strGroupName, l_varKeyArray, l_varColumnsArray, v_blnOneRowForEachKey, l_rsGroupRecords
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "InitializeGroup"
End Function

Private Function LoadRecordset( _
            ByVal v_strGroupName As String, _
            ByVal v_varKeys As Variant, _
            ByVal v_varColumns As Variant, _
            ByVal v_blnOneRowForEachKey As Boolean, _
            ByRef r_rsRecords As ADODB.Recordset _
        ) As Boolean
    Dim spmWrapper As utilSPMCache.cuSPMCache
    Dim lblnExists As Boolean
    Dim lvarRecordArray() As Variant, lvarRecord() As Variant, lvarRows() As Variant
    Dim sKey As Variant
    Dim llngRetVal As Long
    Dim llngRecordCount As Long
    Dim i As Integer, i2 As Integer
    Dim sKeySave As String
    Dim nRowStart As Long, nKeyGroup As Long
    Dim varBookmark As Variant, varBookmarkStart As Variant
    Dim stTime As Long, ndTime As Long, dfTime As Long
    Dim l_colErrArgs As Collection
On Error GoTo eh
    If r_rsRecords Is Nothing Then
        LoadRecordset = False
        Exit Function
    End If
    
    #If TRDEBUG Then
        stTime = GetTickCount()
    #End If
    
'//    lvarFieldArray = BuildFieldsArray(r_rsRecords)
    AddMetaDataToCache v_strGroupName, v_varKeys, r_rsRecords

    '//-----add the group to the cache for groups not having any data,
    '//-----otherwise additional trips to the db will be made (RULE0015CACHE)
    '//-----(the item being get will be set to CDOESNOTEXIST when GetRecord3 calls utilSPMCache.Item )
    If r_rsRecords.BOF And r_rsRecords.EOF Then
        Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
        spmWrapper.CreateGroup v_strGroupName, lblnExists
        Set spmWrapper = Nothing
        
        LoadRecordset = False
        Exit Function
    End If
    
    '//-----specify columns from recordset to cache-----//
    If IsMissing(v_varColumns) Or IsEmpty(v_varColumns) Then
        v_varColumns = BuildFieldsArray(r_rsRecords) '//-----use all columns in recordset-----//
    End If
    
    '//-----normally each key/row is cached into a 1-dimensional array that represents each column-----//
    If v_blnOneRowForEachKey Then
'        If Not IsEmpty(v_varColumns) Then
            ReDim lvarRecord(UBound(v_varColumns) + 1)   'add 1 for key in index 0
'        End If
        
        llngRecordCount = r_rsRecords.RecordCount
        ReDim lvarRecordArray(1, llngRecordCount - 1)

        r_rsRecords.MoveFirst
        Do Until r_rsRecords.EOF
            lvarRecord(0) = BuildCacheKey(v_varKeys, r_rsRecords) 'key is always index 0
'            '//-----might be able to use the following instead of the loop
'            lvarRows = r_rsRecords.GetRows(1, adBookmarkCurrent, v_varColumns)
'            lvarRecord(0) = BuildCacheKey(v_varKeys, r_rsRecords) 'key is always index 0
'            lvarRecord(1) = lvarRows
                    
            For i2 = 1 To UBound(lvarRecord)
                '//lvarRecord(i2) = GetFieldValue(r_rsRecords.Fields(i2 - 1))
                lvarRecord(i2) = GetFieldValue(r_rsRecords.Fields(v_varColumns(i2 - 1)))
            Next
            lvarRecordArray(0, i) = lvarRecord(0)    'key is always index 0
            lvarRecordArray(1, i) = lvarRecord
            r_rsRecords.MoveNext
            i = i + 1
        Loop
    '//-----union components will have multiple rows per key and will be cached into a 2-dimensional array-----//
    Else
        llngRecordCount = r_rsRecords.RecordCount
'        If Not IsEmpty(v_varColumns) Then
'            ReDim lvarRecord(UBound(v_varColumns))     'add 1 for key in index 0
'        End If
        ReDim lvarRecord(0 To 1)
        ReDim lvarRecordArray(1, llngRecordCount - 1) '//-----allocate maximum

        sKeySave = ""
        r_rsRecords.MoveFirst
        varBookmarkStart = r_rsRecords.Bookmark
        Do While Not r_rsRecords.EOF
            sKey = BuildCacheKey(v_varKeys, r_rsRecords)
            If sKeySave <> sKey Then '//-----check for change in keys-----//
                If sKeySave <> "" Then '//-----flush out, except first-----//
                    varBookmark = r_rsRecords.Bookmark '//-----Save-----//
                    lvarRows = r_rsRecords.GetRows(i - nRowStart, varBookmarkStart, v_varColumns)
                    lvarRecord(0) = sKeySave 'key is always index 0
                    lvarRecord(1) = lvarRows
                    lvarRecordArray(0, nKeyGroup) = sKeySave    'key is always index 0
                    lvarRecordArray(1, nKeyGroup) = lvarRecord
                    nKeyGroup = nKeyGroup + 1
                    r_rsRecords.Bookmark = varBookmark  '//-----Restore-----//
                End If
                nRowStart = i
                varBookmarkStart = r_rsRecords.Bookmark
                sKeySave = sKey
            End If
            
            r_rsRecords.MoveNext
            i = i + 1
        Loop
        
        '//-----flush out last-----//
        lvarRows = r_rsRecords.GetRows(i - nRowStart, varBookmarkStart, v_varColumns)
        lvarRecord(0) = sKeySave 'key is always index 0
        lvarRecord(1) = lvarRows
        lvarRecordArray(0, nKeyGroup) = sKeySave    'key is always index 0
        lvarRecordArray(1, nKeyGroup) = lvarRecord

        ReDim Preserve lvarRecordArray(1, nKeyGroup) '//-----reallocate for groups used-----//
    End If
    
    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
    'group specific
    llngRetVal = spmWrapper.AddAll(lvarRecordArray, v_strGroupName)
    
    #If TRDEBUG Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        Debug.Print "Build " & v_strGroupName & " in seconds: " & CStr(dfTime / 1000)
    #End If
    
    LoadRecordset = (llngRetVal = 0)
    Set spmWrapper = Nothing
    
    Exit Function
    
eh:
    Set spmWrapper = Nothing
    
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    If IsArray(v_varKeys) Then
        l_colErrArgs.Add "v_varKeys() = (" & Join(v_varKeys, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKeys = " & CStr(v_varKeys)
    End If
    If IsArray(v_varColumns) Then
        l_colErrArgs.Add "v_varColumns() = (" & Join(v_varColumns, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varColumns = " & CStr(v_varColumns)
    End If
    l_colErrArgs.Add "v_blnOneRowForEachKey = " & v_blnOneRowForEachKey
    l_colErrArgs.Add "r_rsRecords.Source = " & r_rsRecords.Source
    
    TrsRaiseError m_strCMODULENAME, "LoadRecordset", , , , l_colErrArgs, True
End Function

'//-----could be changed to utilize recordarray returned from AddMetaDataToCache
Private Function BuildCacheKey( _
            ByVal v_varKeyNames As Variant, _
            ByRef r_rsRecords As ADODB.Recordset _
        ) As String
    Dim i As Integer
    Dim lstrInteger As String
    Dim l_rsField As Field
    Dim l_lngKeySize As Integer, l_lngStart As Long, l_lngLength As Long
    Dim l_strKey As String
    
On Error GoTo eh
    l_lngKeySize = 0
    For i = 0 To UBound(v_varKeyNames)
        Set l_rsField = r_rsRecords.Fields(v_varKeyNames(i))
    
        Select Case l_rsField.Type
        Case adChar, adVarChar, adLongVarChar
            l_lngKeySize = l_lngKeySize + l_rsField.DefinedSize
        Case adInteger, adSmallInt '//DefinedSize=size in bytes, Precision=decimal size
            l_lngKeySize = l_lngKeySize + l_rsField.Precision + 1
        Case Else
            l_lngKeySize = l_lngKeySize + l_rsField.DefinedSize
        End Select
    Next i
    
    l_strKey = Space(l_lngKeySize) '//-----allocate memory once-----//
    
    l_lngStart = 1
    For i = 0 To UBound(v_varKeyNames)
        Set l_rsField = r_rsRecords.Fields(v_varKeyNames(i))
        Select Case l_rsField.Type
        Case adChar, adVarChar, adLongVarChar
            l_lngLength = l_rsField.DefinedSize
            Mid(l_strKey, l_lngStart, l_lngLength) = Trim(l_rsField.Value)
        Case adInteger, adSmallInt '//DefinedSize=size in bytes, Precision=decimal size
            l_lngLength = l_rsField.Precision + 1 '//extra for sign
            lstrInteger = String(l_lngLength, "0")  '//-----PadLeft with zeros
            lstrInteger = lstrInteger & Abs(Format(l_rsField.Value))
            Mid(l_strKey, l_lngStart, l_lngLength) = IIf(l_rsField.Value < 0, "-", "0") & Right(lstrInteger, l_lngLength - 1)
        Case Else
            l_lngLength = l_rsField.DefinedSize
            Mid(l_strKey, l_lngStart, l_lngLength) = l_rsField.Value
        End Select
        l_lngStart = l_lngStart + l_lngLength
    Next

    BuildCacheKey = l_strKey
'    BuildCacheKey = BuildCacheKeyString(v_varKeys)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildCacheKey"
End Function

'//-----------------------------------------------------------------------------------------------------
'//similar to BuildCacheKey, used by GetRecord2
Private Function FormatKeyFromMetaData( _
            ByVal v_strGroupName As String, _
            ByVal v_varKeyValue As Variant _
        ) As String
    Dim i As Integer
    Dim l_strInteger As String
    Dim l_lngKeySize As Integer, l_lngStart As Long, l_lngLength As Long
    Dim l_strKey As String
    Dim l_varRecordArray As Variant
    Dim l_varRecord As Variant
    Dim spmWrapper As utilSPMCache.cuSPMCache
    Dim l_blnMetaDataExists As Boolean
    Dim l_blnMetaDataGroupExists As Boolean
    Dim l_colErrArgs As Collection
On Error GoTo eh

    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
    l_varRecordArray = spmWrapper.Item(v_strGroupName, CCACHEMETADATA, l_blnMetaDataExists, l_blnMetaDataGroupExists)
    
    If Not l_blnMetaDataGroupExists Or Not l_blnMetaDataExists Then '//-----Meta Data has not been loaded in cache-----//
        Err.Raise trErrRecordCacheNoMetaData
    End If
    
    If IsArray(l_varRecordArray) Then
        l_lngKeySize = 0
        For i = 0 To UBound(l_varRecordArray, 2)
            If IsArray(l_varRecordArray(1, i)) Then
                l_varRecord = l_varRecordArray(1, i)
                If IsNumeric(l_varRecord(CMD_KEYSIZE)) Then
                    l_lngKeySize = l_lngKeySize + l_varRecord(CMD_KEYSIZE)
                Else
                    Err.Raise trErrRecordCacheInvalidMetaData, , "field CMD_KEYSIZE"
                End If
            Else
                Err.Raise trErrRecordCacheInvalidMetaData, , "l_varRecordArray(1, " & i & ") is not an array"
            End If
        Next i
        
        l_strKey = Space(l_lngKeySize) '//-----allocate memory once-----//
        
        l_lngStart = 1
        For i = 0 To UBound(l_varRecordArray, 2)
            l_varRecord = l_varRecordArray(1, i)
            Select Case l_varRecord(CMD_KEYDATATYPE)
            Case adChar, adVarChar, adLongVarChar
                l_lngLength = l_varRecord(CMD_KEYSIZE)
                Mid(l_strKey, l_lngStart, l_lngLength) = Trim(v_varKeyValue(i))
            Case adInteger, adSmallInt '//DefinedSize=size in bytes, Precision=decimal size
                l_lngLength = l_varRecord(CMD_KEYSIZE)  '//extra for sign
                l_strInteger = String(l_lngLength, "0")  '//-----PadLeft with zeros
                l_strInteger = l_strInteger & Abs(Format(v_varKeyValue(i)))
                Mid(l_strKey, l_lngStart, l_lngLength) = IIf(v_varKeyValue(i) < 0, "-", "0") & Right(l_strInteger, l_lngLength - 1)
            Case Else
                l_lngLength = l_varRecord(CMD_KEYSIZE)
                Mid(l_strKey, l_lngStart, l_lngLength) = v_varKeyValue(i)
            End Select
            l_lngStart = l_lngStart + l_lngLength
        Next
    Else
        Err.Raise trErrRecordCacheInvalidMetaData, , "l_varRecordArray is not an array"
    End If
    
    Set spmWrapper = Nothing
    
    FormatKeyFromMetaData = l_strKey
    
    Exit Function
    
eh:
    Set spmWrapper = Nothing
    
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    If IsArray(v_varKeyValue) Then
        l_colErrArgs.Add "v_varKeyValue() = (" & Join(v_varKeyValue, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKeyValue = " & CStr(v_varKeyValue)
    End If
    
    If IsArray(l_varRecordArray) Then
        l_colErrArgs.Add "CMD_GROUP, CMD_KEYCOLUMN, CMD_KEYDATATYPE, CMD_KEYSIZE"
        For i = 0 To UBound(l_varRecordArray, 2)
            l_varRecord = l_varRecordArray(1, i)

            l_colErrArgs.Add l_varRecord(CMD_GROUP) & ", " & l_varRecord(CMD_KEYCOLUMN) & ", " & l_varRecord(CMD_KEYDATATYPE) & ", " & l_varRecord(CMD_KEYSIZE)
        Next i
    End If
    
    Select Case Err.Number
    Case trErrRecordCacheNoMetaData
        TrsRaiseError m_strCMODULENAME, "FormatKeyFromMetaData", Err.Number, , DescArgsFrom(v_strGroupName), l_colErrArgs, True
    Case trErrRecordCacheInvalidMetaData
        TrsRaiseError m_strCMODULENAME, "FormatKeyFromMetaData", Err.Number, , DescArgsFrom(Err.Description, v_strGroupName), l_colErrArgs, True
    Case Else
        TrsRaiseError m_strCMODULENAME, "FormatKeyFromMetaData", , , , l_colErrArgs, True
    End Select
End Function


Private Function AddMetaDataToCache( _
            ByVal v_strGroupName As String, _
            ByVal v_varKeyNames As Variant, _
            ByVal v_rsRecord As ADODB.Recordset _
        ) As Variant
    
    Dim l_varRecordArray As Variant
    Dim l_varRecordItem As Variant
    Dim spmWrapper As utilSPMCache.cuSPMCache
    Dim l_blnMetaDataExists As Boolean
    Dim l_blnMetaDataGroupExists As Boolean
    Dim rc As Long
    Dim l_colErrArgs As Collection
    Dim i As Integer
    Dim l_varRecord As Variant
On Error GoTo eh
    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")
    l_varRecordArray = spmWrapper.Item(v_strGroupName, CCACHEMETADATA, l_blnMetaDataExists, l_blnMetaDataGroupExists)
    If Not l_blnMetaDataGroupExists Or Not l_blnMetaDataExists Then '//-----Group has not been loaded in cache-----//
        l_varRecordItem = BuildMetaData(v_strGroupName, v_varKeyNames, v_rsRecord)
        rc = spmWrapper.AddAll(l_varRecordItem, CCACHEMETADATA)
        
        '//-----this will be different than the variant array returned from Item
        If IsArray(l_varRecordItem(1, 0)) Then
            l_varRecordArray = l_varRecordItem(1, 0)
        Else
            Err.Raise trErrRecordCacheInvalidMetaData, , "l_varRecordItem(1, 0) is not an array"
        End If
'//        l_varRecordArray = spmWrapper.Item(v_strKey, v_strGroupName, r_blnExists, l_blnMetaDataGroupExists)
    End If
    
    ' test for array first since 99.9% of time it should be one,
    ' otherwise answer empty.
    If IsArray(l_varRecordArray) Then
        AddMetaDataToCache = l_varRecordArray
    Else
        AddMetaDataToCache = Empty
    End If
    Set spmWrapper = Nothing
    
    Exit Function

eh:
    Set spmWrapper = Nothing
                
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    If IsArray(v_varKeyNames) Then
        l_colErrArgs.Add "v_varKeyNames() = (" & Join(v_varKeyNames, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKeyNames = " & CStr(v_varKeyNames)
    End If
    
    l_colErrArgs.Add "v_rsRecord.Source = " & v_rsRecord.Source
    
    If IsArray(l_varRecordArray) Then
        l_colErrArgs.Add "CMD_GROUP, CMD_KEYCOLUMN, CMD_KEYDATATYPE, CMD_KEYSIZE"
        For i = 0 To UBound(l_varRecordArray, 2)
            l_varRecord = l_varRecordArray(1, i)

            l_colErrArgs.Add l_varRecord(CMD_GROUP) & ", " & l_varRecord(CMD_KEYCOLUMN) & ", " & l_varRecord(CMD_KEYDATATYPE) & ", " & l_varRecord(CMD_KEYSIZE)
        Next i
    End If
    
    Select Case Err.Number
    Case trErrRecordCacheInvalidMetaData
        TrsRaiseError m_strCMODULENAME, "AddMetaDataToCache", Err.Number, , DescArgsFrom(Err.Description, v_strGroupName), l_colErrArgs, True
    Case Else
        TrsRaiseError m_strCMODULENAME, "AddMetaDataToCache", , , , l_colErrArgs, True
    End Select
End Function

Private Function BuildMetaData( _
            ByVal v_strGroupName As String, _
            ByVal v_varKeyNames As Variant, _
            ByVal v_rsRecord As ADODB.Recordset _
        ) As Variant
    Dim i As Integer
    Dim l_rsField As Field
    Dim l_strKey As String
    Dim l_varRecordItem(1, 0) As Variant
    Dim l_varRecordArray As Variant
    Dim l_varRecord(CMD_LAST_SLOT) As Variant
    Dim l_colErrArgs As Collection
    Dim l_varRecordErr As Variant
On Error GoTo eh
    
    ReDim l_varRecordArray(1, UBound(v_varKeyNames))
    
    For i = 0 To UBound(v_varKeyNames)
        l_varRecord(CMD_GROUP) = v_strGroupName
        l_varRecord(CMD_KEYCOLUMN) = v_varKeyNames(i)

        Set l_rsField = v_rsRecord.Fields(v_varKeyNames(i))
        l_varRecord(CMD_KEYDATATYPE) = l_rsField.Type
        
        Select Case l_rsField.Type
        Case adChar, adVarChar, adLongVarChar
            l_varRecord(CMD_KEYSIZE) = l_rsField.DefinedSize
        Case adInteger, adSmallInt '//DefinedSize=size in bytes, Precision=decimal size
            l_varRecord(CMD_KEYSIZE) = l_rsField.Precision + 1 '//extra for sign
        Case Else
            l_varRecord(CMD_KEYSIZE) = l_rsField.DefinedSize
        End Select

        l_varRecordArray(0, i) = l_varRecord(0) '//-----key
        l_varRecordArray(1, i) = l_varRecord
    Next

    l_varRecordItem(0, 0) = v_strGroupName '//-----key
    l_varRecordItem(1, 0) = l_varRecordArray

    BuildMetaData = l_varRecordItem
        
    Exit Function
    
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    If IsArray(v_varKeyNames) Then
        l_colErrArgs.Add "v_varKeyNames() = (" & Join(v_varKeyNames, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKeyNames = " & CStr(v_varKeyNames)
    End If
    
    l_colErrArgs.Add "v_rsRecord.Source = " & v_rsRecord.Source
    
    If IsArray(l_varRecordArray) Then
        l_colErrArgs.Add "CMD_GROUP, CMD_KEYCOLUMN, CMD_KEYDATATYPE, CMD_KEYSIZE"
        For i = 0 To UBound(l_varRecordArray, 2)
            l_varRecordErr = l_varRecordArray(1, i)

            l_colErrArgs.Add l_varRecordErr(CMD_GROUP) & ", " & l_varRecordErr(CMD_KEYCOLUMN) & ", " & l_varRecordErr(CMD_KEYDATATYPE) & ", " & l_varRecordErr(CMD_KEYSIZE)
        Next i
    End If
    
    TrsRaiseError m_strCMODULENAME, "BuildMetaData", , , , l_colErrArgs, True
End Function

Private Function BuildFieldsArray(ByRef r_rsRecords As ADODB.Recordset) As Variant
' Answer an array of field names that will be cached.
' Array should be zero-based.
' This will be the default array, used when the recordset is exactly
' what should be cached.  However, design should allow the cache to be
' built based on a recordset that is not exactly same as cache.
    Dim llngFieldCount As Long
    Dim l_Field As ADODB.Field
    Dim lvarFieldArray() As Variant
    Dim i As Integer

On Error GoTo eh

    If r_rsRecords.EOF And r_rsRecords.BOF Then
        BuildFieldsArray = Empty
        Exit Function
    Else
        r_rsRecords.MoveFirst
    End If
    
    llngFieldCount = r_rsRecords.Fields.Count
    ReDim lvarFieldArray(llngFieldCount - 1)
    
    For i = 0 To llngFieldCount - 1
        lvarFieldArray(i) = Trim(r_rsRecords.Fields(i).Name)
    Next
    
    BuildFieldsArray = lvarFieldArray
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildFieldsArray"
End Function


Private Function GetFieldValue(ByRef r_rsField As ADODB.Field) As Variant
' Do as little transformation as possible here to keep it simple.
' SQL statement can do the UCASE.
' Just Trim strings here for now.

On Error GoTo eh

    Select Case r_rsField.Type
    Case adChar, adVarChar, adLongVarChar
        If r_rsField.Name = "ACCT_NUM" Or r_rsField.Name = "UEC_ACCT_NUM" Or r_rsField.Name = "AIC_ACCT_NUM" Or r_rsField.Name = "AMS_ACCT_NUM" Then
            GetFieldValue = r_rsField.Value
        Else
            GetFieldValue = Trim(r_rsField.Value)
        End If
    Case Else
        GetFieldValue = r_rsField.Value
    End Select
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetFieldValue"
End Function


Private Function LoadRecord( _
            ByVal v_strKey As String, _
            ByVal v_varKey As Variant, _
            ByVal v_strGroupName As String, _
            ByVal v_blnOneRowForEachKey As Boolean, _
            ByRef r_objCaller As Object) As Variant
            
    Dim l_rsGroupRecords As ADODB.Recordset
    Dim spmWrapper As utilSPMCache.cuSPMCache
    Dim llngRetVal As Long
    Dim l_varKeyArray As Variant, l_varColumnArray As Variant
    
    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long
    Dim l_colErrArgs As Collection
    Dim sTraceLine As String
On Error GoTo eh

    #If TRDEBUG Then
        stTime = GetTickCount()
    #End If
    '//-----pass the variant key back to dbCodeTables for multi-key SQL
    'Set l_rsGroupRecords = r_objCaller.GetGroupRecords(v_strGroupName, Trim(v_strKey), l_varKeyArray, l_varColumnArray)
sTraceLine = "Set l_rsGroupRecords = r_objCaller.GetGroupRecords(v_strGroupName, v_varKey, l_varKeyArray, l_varColumnArray)"
    Set l_rsGroupRecords = r_objCaller.GetGroupRecords(v_strGroupName, v_varKey, l_varKeyArray, l_varColumnArray)

sTraceLine = "Set spmWrapper = CreateObject(""utilSPMCache.cuSPMCache"")"
    Set spmWrapper = CreateObject("utilSPMCache.cuSPMCache")

    If Not l_rsGroupRecords Is Nothing Then
        If l_rsGroupRecords.BOF And l_rsGroupRecords.EOF Then
sTraceLine = "llngRetVal = spmWrapper.Remove(v_strKey, v_strGroupName)"
            llngRetVal = spmWrapper.Remove(v_strKey, v_strGroupName)
        Else
sTraceLine = "LoadRecordset v_strGroupName, l_varKeyArray, l_varColumnArray, v_blnOneRowForEachKey, l_rsGroupRecords"
            LoadRecordset v_strGroupName, l_varKeyArray, l_varColumnArray, v_blnOneRowForEachKey, l_rsGroupRecords
        End If
    Else
sTraceLine = "llngRetVal = spmWrapper.Remove(v_strKey, v_strGroupName)"
        llngRetVal = spmWrapper.Remove(v_strKey, v_strGroupName)
    End If
    
sTraceLine = "Set spmWrapper = Nothing"
    Set spmWrapper = Nothing
    Set l_rsGroupRecords = Nothing
    
    #If TRDEBUG Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        Debug.Print "LoadRecord-" & v_strGroupName & " in seconds: " & CStr(dfTime / 1000)
    #End If

    Exit Function
eh:
    Set spmWrapper = Nothing
    Set l_rsGroupRecords = Nothing
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strKey = """ & v_strKey & """"
    If IsMissing(v_varKey) Or IsEmpty(v_varKey) Then
        l_colErrArgs.Add "v_varKey = empty"
    ElseIf IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & v_varKey
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    l_colErrArgs.Add "v_blnOneRowForEachKey = " & v_blnOneRowForEachKey
    If Not r_objCaller Is Nothing Then
        l_colErrArgs.Add "TypeName(r_objCaller) = " & TypeName(r_objCaller)
    Else
        l_colErrArgs.Add "r_objCaller = Nothing"
    End If
    l_colErrArgs.Add "TraceLine  = " & sTraceLine
    
    TrsRaiseError m_strCMODULENAME, "LoadRecord", , , , l_colErrArgs, True
End Function



