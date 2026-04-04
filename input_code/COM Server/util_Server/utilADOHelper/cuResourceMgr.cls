VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cuResourceMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'unfortunately using TrsTraceError also writes to the event viewer
' Disable with TRDEBUG = 0.
' Enable with TRDEBUG = -1.
#Const TRDEBUG = 0

'Module Name
Private Const MODULENAME = "utilADOHelper.cuResourceMgr"
Private Const m_lngSerializationChunkSize As Long = 10240

'Public Enum ResourceTypeEnum
'    eString = 0
'    eRecordSet = 1
'End Enum

Private Const STRING_RESOURCE = "STRING_RESOURCE"
Private Const RECORDSET_RESOURCE = "RECORDSET_RESOURCE"

Public Sub InitializeResources(ByVal v_strRoot As String, _
                                ByVal v_strPartialPath As String, _
                                ByVal v_blnRecursive As Boolean)
    Dim l_colErrArgs As Collection
On Error GoTo eh

    InitializeResourcesByType v_strRoot, v_strPartialPath, v_blnRecursive, STRING_RESOURCE
    InitializeResourcesByType v_strRoot, v_strPartialPath, v_blnRecursive, RECORDSET_RESOURCE
        
    Exit Sub
eh:
    Set l_colErrArgs = New Collection

    l_colErrArgs.Add "v_strPartialPath = " & v_strPartialPath
    l_colErrArgs.Add "v_blnRecursive = " & v_blnRecursive

    TrsRaiseError MODULENAME, "InitializeResources", , , , l_colErrArgs, True
End Sub

Public Function GetStringResource(ByVal v_strRoot As String, _
                                    ByVal v_strPartialPath As String, ByVal v_strKey As String) As String
    Dim l_colErrArgs As Collection
On Error GoTo eh
    Dim strPartialPathKey As String
    
'//-----start debug
#If TRDEBUG Then
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strRoot = " & v_strRoot
    l_colErrArgs.Add "v_strPartialPath = " & v_strPartialPath
    l_colErrArgs.Add "v_strKey = " & v_strKey
    TrsTraceError MODULENAME, "GetStringResource", 10000, , , l_colErrArgs, True
#End If
'//-----end debug
    strPartialPathKey = GetPartialPathKey(v_strPartialPath, v_strKey)
    
    GetStringResource = GetResourceFromSPM(v_strRoot, v_strPartialPath, strPartialPathKey, STRING_RESOURCE)
    Exit Function
eh:
    Set l_colErrArgs = New Collection

    l_colErrArgs.Add "v_strRoot = " & v_strRoot
    l_colErrArgs.Add "v_strPartialPath = " & v_strPartialPath
    l_colErrArgs.Add "v_strKey = " & v_strKey

    TrsRaiseError MODULENAME, "GetStringResource", , , , l_colErrArgs, True
End Function

Public Function GetRecordSetResource(ByVal v_strRoot As String, _
                                    ByVal v_strPartialPath As String, ByVal v_strKey As String) As ADODB.Recordset
    Dim l_colErrArgs As Collection
On Error GoTo eh
    Dim strPartialPathKey As String
    Dim rs As ADODB.Recordset
    Dim sRSXml As String, varRS As Variant
    
    strPartialPathKey = GetPartialPathKey(v_strPartialPath, v_strKey)
    
    varRS = GetResourceFromSPM(v_strRoot, v_strPartialPath, strPartialPathKey, RECORDSET_RESOURCE)
    
    If IsArray(varRS) Then
        Set rs = DeSerializeRSFromBinary(varRS)
    Else
        Set rs = DeSerializeRSFromXML(varRS)
    End If
    
    Set GetRecordSetResource = rs
    Exit Function
eh:
    Set l_colErrArgs = New Collection

    l_colErrArgs.Add "v_strRoot = " & v_strRoot
    l_colErrArgs.Add "v_strPartialPath = " & v_strPartialPath
    l_colErrArgs.Add "v_strKey = " & v_strKey

    TrsRaiseError MODULENAME, "GetRecordSetResource", , , , l_colErrArgs, True
End Function

Private Function GetResourceFromSPM(ByVal v_strRoot As String, _
                                    ByVal v_varPartialPath As String, _
                                    ByVal v_varPartialPathKey As String, _
                                    ByVal v_strResourceType As String) As Variant
    Dim l_varRecord As Variant
    Dim oSPMWrapper As utilSPMCache.cuSPMCache
    Dim lblnExists As Boolean
    Dim lblnGroupExists As Boolean
    Dim l_colErrArgs As Collection
    Dim RootAndResourceType As String
On Error GoTo eh
    Set oSPMWrapper = CreateObject("utilSPMCache.cuSPMCache")
    
    RootAndResourceType = v_strRoot & "|" & v_strResourceType
    l_varRecord = oSPMWrapper.Item(LCase(v_varPartialPathKey), LCase(RootAndResourceType), lblnExists, lblnGroupExists)

'//-----start debug
#If TRDEBUG Then
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strRoot = " & v_strRoot
    l_colErrArgs.Add "v_varPartialPath = """ & v_varPartialPath & """"
    l_colErrArgs.Add "v_varPartialPathKey = """ & v_varPartialPathKey & """"
    l_colErrArgs.Add "v_strResourceType = " & v_strResourceType
    l_colErrArgs.Add "l_varRecord = oSPMWrapper.Item(LCase(v_varPartialPathKey), LCase(RootAndResourceType), lblnExists, lblnGroupExists)"
    l_colErrArgs.Add "lblnExists = " & lblnExists
    l_colErrArgs.Add "lblnGroupExists = " & lblnGroupExists
    TrsTraceError MODULENAME, "GetResourceFromSPM", 10000, , , l_colErrArgs, True
#End If
'//-----end debug
    
    If Not lblnGroupExists Then '//-----Group has not been loaded in cache-----//
'        If v_blnLazyInitialized Then '//-----Load on demand (one at a time)-----//
'            LoadRecord v_strKey, v_varKey, v_strResourceType, v_blnOneRowForEachKey, r_objCaller
'            l_varRecord = oSPMWrapper.Item(v_strKey, v_strResourceType, r_blnExists, lblnGroupExists)
'        Else '//-----Load All (Bulk)-----//
        InitializeResourcesByType v_strRoot, "", True, v_strResourceType
        l_varRecord = oSPMWrapper.Item(LCase(v_varPartialPathKey), LCase(RootAndResourceType), lblnExists, lblnGroupExists)
'            InitializeGroup v_strResourceType, v_blnOneRowForEachKey, r_objCaller
'            l_varRecord = oSPMWrapper.Item(v_strKey, v_strResourceType, r_blnExists, lblnGroupExists)
'        End If
    Else '//-----Group has been loaded in cache-----//
        If Not lblnExists Then '//-----Key is not found-----//
        '            If v_blnLazyInitialized Then '//-----Load on demand (one at a time)-----//
            InitializeResourcesByType v_strRoot, v_varPartialPath, False, v_strResourceType
            l_varRecord = oSPMWrapper.Item(LCase(v_varPartialPathKey), LCase(RootAndResourceType), lblnExists, lblnGroupExists)
'                LoadRecord v_strKey, v_varKey, v_strResourceType, v_blnOneRowForEachKey, r_objCaller
'                l_varRecord = oSPMWrapper.Item(v_strKey, v_strResourceType, r_blnExists, lblnGroupExists)
'            End If
        Else    'it exists and INVALID - reload it
                '***There must be a Load method for every cache that supports Invalidate
            If Not IsArray(l_varRecord) Then
'                If l_varRecord = CSPINVALID And v_blnSupportsInvalidate Then '//-----Supports Invalidate (load one)-----//
'                    LoadRecord v_strKey, v_varKey, v_strResourceType, v_blnOneRowForEachKey, r_objCaller
'                    l_varRecord = oSPMWrapper.Item(v_strKey, v_strResourceType, r_blnExists, lblnGroupExists)
'                End If
'//-----start debug
#If TRDEBUG Then
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strRoot = " & v_strRoot
    l_colErrArgs.Add "v_varPartialPath = """ & v_varPartialPath & """"
    l_colErrArgs.Add "v_varPartialPathKey = """ & v_varPartialPathKey & """"
    l_colErrArgs.Add "v_strResourceType = " & v_strResourceType
    l_colErrArgs.Add "Oops not initializing the spm cache!"
    TrsTraceError MODULENAME, "GetResourceFromSPM", 10000, , , l_colErrArgs, True
#End If
'//-----end debug
            End If
            
        End If
    End If
    ' test for array first since 99.9% of time it should be one,
    ' otherwise answer empty.
    If IsArray(l_varRecord) Then
        GetResourceFromSPM = l_varRecord(1) '//-----should be string or variant (binary)
    Else
        GetResourceFromSPM = ""
    End If
    Set oSPMWrapper = Nothing
    Exit Function

eh:
    Set oSPMWrapper = Nothing
    
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strRoot = " & v_strRoot
    l_colErrArgs.Add "v_varPartialPath = """ & v_varPartialPath & """"
    l_colErrArgs.Add "v_varPartialPathKey = """ & v_varPartialPathKey & """"
'    If IsMissing(v_varKey) Or IsEmpty(v_varKey) Then
'        l_colErrArgs.Add "v_varKey = empty"
'    ElseIf IsArray(v_varKey) Then
'        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
'    Else
'        l_colErrArgs.Add "v_varKey = " & v_varKey
'    End If
    l_colErrArgs.Add "v_strResourceType = " & v_strResourceType
    l_colErrArgs.Add "lblnGroupExists = " & lblnGroupExists
    l_colErrArgs.Add "lblnExists = " & lblnExists

    TrsRaiseError MODULENAME, "GetResourceFromSPM", , , , l_colErrArgs, True
End Function

Private Function AddResourcesToSPM(ByVal v_strRoot As String, _
                                    ByVal v_varPartialPathKeyValue As Variant, _
                                    ByVal v_strResourceType As String) As Long
    Dim oSPMWrapper As utilSPMCache.cuSPMCache
    Dim rc As Long
    Dim l_colErrArgs As Collection
    Dim RootAndResourceType As String
On Error GoTo eh

    RootAndResourceType = v_strRoot & "|" & v_strResourceType
    Set oSPMWrapper = CreateObject("utilSPMCache.cuSPMCache")
    
'//-----start debug
#If TRDEBUG Then
    Dim k As String
    Dim kv As Variant

    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strRoot = " & v_strRoot
    If UBound(v_varPartialPathKeyValue, 2) >= 0 Then
        k = v_varPartialPathKeyValue(0, 0)
        l_colErrArgs.Add "v_varPartialPathKeyValue(0, 0) = " & k
        
        If UBound(v_varPartialPathKeyValue, 1) >= 1 Then
            kv = v_varPartialPathKeyValue(1, 0)
    
            If UBound(kv, 1) >= 1 Then
                l_colErrArgs.Add "v_varPartialPathKeyValue(1, 0)(0) = " & kv(0)
                l_colErrArgs.Add "v_varPartialPathKeyValue(1, 0)(1) = " & kv(1)
            End If
        End If
    End If
    l_colErrArgs.Add "v_strResourceType = " & v_strResourceType
    TrsTraceError MODULENAME, "AddResourcesToSPM", 10000, , , l_colErrArgs, True
#End If
'//-----end debug
    
    rc = oSPMWrapper.AddAll(v_varPartialPathKeyValue, LCase(RootAndResourceType))
    
    Set oSPMWrapper = Nothing
    
    AddResourcesToSPM = rc
    Exit Function
eh:
    Set oSPMWrapper = Nothing
    
    Set l_colErrArgs = New Collection
'    l_colErrArgs.Add "v_strKey = """ & v_strKey & """"
'    If IsMissing(v_varKey) Or IsEmpty(v_varKey) Then
'        l_colErrArgs.Add "v_varKey = empty"
'    ElseIf IsArray(v_varKey) Then
'        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
'    Else
'        l_colErrArgs.Add "v_varKey = " & v_varKey
'    End If

    l_colErrArgs.Add "v_strRoot = " & v_strRoot
    l_colErrArgs.Add "v_varPartialPathKeyValue = " & v_varPartialPathKeyValue
    l_colErrArgs.Add "v_strResourceType = " & v_strResourceType
    l_colErrArgs.Add "rc = " & rc

    TrsRaiseError MODULENAME, "AddResourcesToSPM", , , , l_colErrArgs, True
End Function

Private Sub InitializeResourcesByType(ByVal v_strRoot As String, _
                                ByVal v_strPartialPath As String, _
                                ByVal v_blnRecursive As Boolean, _
                                ByVal v_strResourceType As String)
    Dim dctResource As Dictionary, varResource As Variant
    Dim sFullPath As String, sRootPath As String
    Dim rc As Long
    Dim l_colErrArgs As Collection
    
    On Error GoTo eh

    sFullPath = GetFullPath(v_strRoot, v_strPartialPath)
    '//sRootPath = GetRootPath
    
    Set dctResource = New Dictionary
    
'//-----start debug
#If TRDEBUG Then
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strRoot = " & v_strRoot
    l_colErrArgs.Add "v_strPartialPath = " & v_strPartialPath
    l_colErrArgs.Add "v_blnRecursive = " & v_blnRecursive
    l_colErrArgs.Add "v_strResourceType = " & v_strResourceType
    l_colErrArgs.Add "sFullPath = " & sFullPath
    l_colErrArgs.Add "before call to LoadResources"
    TrsTraceError MODULENAME, "InitializeResourcesByType", 10000, , , l_colErrArgs, True
#End If
'//-----end debug

    LoadResources sFullPath, v_blnRecursive, v_strRoot, v_strResourceType, dctResource
    
    varResource = DictionaryToArray(dctResource)
    
    rc = AddResourcesToSPM(v_strRoot, varResource, v_strResourceType)
        
    Exit Sub
eh:
    Set l_colErrArgs = New Collection

    l_colErrArgs.Add "v_strRoot = " & v_strRoot
    l_colErrArgs.Add "v_strPartialPath = " & v_strPartialPath
    l_colErrArgs.Add "v_blnRecursive = " & v_blnRecursive
    l_colErrArgs.Add "v_strResourceType = " & v_strResourceType
    l_colErrArgs.Add "sFullPath = " & sFullPath
    l_colErrArgs.Add "rc = " & rc

    TrsRaiseError MODULENAME, "InitializeResourcesByType", , , , l_colErrArgs, True
End Sub

'Private Function PadPartialPath(ByVal v_strPartialPath As String) As String
'    If Left(v_strPartialPath, 1) = "\" Then
'        PadPartialPath = v_strPartialPath
'    Else
'        PadPartialPath = "\" & v_strPartialPath
'    End If
'
'    If Right(v_strPartialPath, 1) = "\" Then
'    Else
'        PadPartialPath = PadPartialPath & "\"
'    End If
'End Function

Private Function GetPartialPathKey(ByVal v_strPartialPath As String, ByVal v_strKey As String) As String
    Dim oFS As New FileSystemObject

    'GetPartialPathKey = PadPartialPath(v_strPartialPath)
    'GetPartialPathKey = GetPartialPathKey & v_strKey
    
    GetPartialPathKey = oFS.BuildPath(v_strPartialPath, v_strKey)
    
    Set oFS = Nothing
End Function

'Private Function GetRootPath() As String
'    Dim oRegUtil As util_DBConnectMgr.clsuDBConnectMgr
'
'    On Error GoTo eh
'    Set oRegUtil = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
'
'    '//-----get the root path (C:\TRIS SOLUTION\dbScripts\db2)
'    GetRootPath = oRegUtil.GetRegistryStringValue("", "dbScriptsRoot")
'
'    Set oRegUtil = Nothing
'    Exit Function
'eh:
'    Set oRegUtil = Nothing
'    Err.Raise Err.Number
'End Function

Private Function GetFullPath(ByVal v_strRoot, _
                            ByVal v_strPartialPath As String) As String
    Dim oFS As New FileSystemObject
    Dim sRootPath As String, sFullPath As String
    
    On Error GoTo eh

    '//-----get the root path (C:\TRIS SOLUTION\dbScripts\db2)
    '//sRootPath = GetRootPath
    
'    If Left(v_strPartialPath, 1) = "\" Then
'        sFullPath = sRootPath & v_strPartialPath
'    Else
'        sFullPath = sRootPath & "\" & v_strPartialPath
'    End If
'
'    If Right(v_strPartialPath, 1) = "\" Then
'    Else
'        sFullPath = sFullPath & "\"
'    End If

    GetFullPath = oFS.BuildPath(v_strRoot, v_strPartialPath)
'    GetFullPath = sFullPath
    Set oFS = Nothing
    Exit Function
eh:
    Set oFS = Nothing
    Err.Raise Err.Number
End Function

'//-----search for text files and load into dictionary collection
Private Sub LoadResources(ByVal v_strFullPath As String, _
                                ByVal v_blnRecursive As Boolean, _
                                ByVal v_strRootPath As String, _
                                ByVal v_strResourceType As String, _
                                ByRef r_dctResource As Dictionary)
    Dim oFS As New FileSystemObject, oFolder As Folder, oFile As File, oSubFolder As Folder
    Dim ts As TextStream
    Dim strPartialPathKey As String, sExt As String
    Dim l_colErrArgs As Collection
        
    On Error GoTo eh
    
'//-----start debug
#If TRDEBUG Then
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strFullPath = " & v_strFullPath
    l_colErrArgs.Add "v_blnRecursive = " & v_blnRecursive
    l_colErrArgs.Add "v_strRootPath = " & v_strRootPath
    l_colErrArgs.Add "v_strResourceType = " & v_strResourceType
'    TrsTraceError MODULENAME, "LoadResources", 10000, , , l_colErrArgs, True
#End If
'//-----end debug

    If oFS.FolderExists(v_strFullPath) Then
        Set oFolder = oFS.GetFolder(v_strFullPath)
        
        For Each oFile In oFolder.Files
            sExt = LCase(oFS.GetExtensionName(oFile.Path))
            Select Case v_strResourceType
            Case STRING_RESOURCE
                Select Case sExt
                Case "sql", "txt"
                    Set ts = oFile.OpenAsTextStream(ForReading)
                    strPartialPathKey = Replace(oFile.Path, v_strRootPath, "", , , vbTextCompare)
                    r_dctResource.Add LCase(strPartialPathKey), ts.ReadAll
                    ts.Close
'//-----start debug
#If TRDEBUG Then
    l_colErrArgs.Add "oFile.Path = " & oFile.Path
#End If
'//-----end debug
                    
                End Select
            Case RECORDSET_RESOURCE
                Select Case sExt
                Case "xml"
                    Set ts = oFile.OpenAsTextStream(ForReading)
                    strPartialPathKey = Replace(oFile.Path, v_strRootPath, "", , , vbTextCompare)
                    r_dctResource.Add LCase(strPartialPathKey), ts.ReadAll
                    ts.Close
                Case "adtg"
                    Dim oStream As New ADODB.Stream
                    oStream.Open
                    oStream.Type = adTypeBinary
                    oStream.LoadFromFile oFile.Path
                    strPartialPathKey = Replace(oFile.Path, v_strRootPath, "", , , vbTextCompare)
                    r_dctResource.Add LCase(strPartialPathKey), oStream.Read(StreamReadEnum.adReadAll)
                    oStream.Close
                End Select
            End Select
        Next
        
'//-----start debug
#If TRDEBUG Then
    TrsTraceError MODULENAME, "GetResourceFromSPM", 10000, , , l_colErrArgs, True
#End If
'//-----end debug

        If v_blnRecursive Then
            For Each oSubFolder In oFolder.SubFolders
                LoadResources oSubFolder.Path, v_blnRecursive, v_strRootPath, v_strResourceType, r_dctResource
            Next
        End If
    End If
    
    Set oFS = Nothing
    Set oFolder = Nothing
    Set oFile = Nothing
    Set ts = Nothing
    
    Exit Sub
eh:
    Set oFS = Nothing
    Set oFolder = Nothing
    Set oFile = Nothing
    Set ts = Nothing
    
    Set l_colErrArgs = New Collection

    l_colErrArgs.Add "v_strFullPath = " & v_strFullPath
    l_colErrArgs.Add "v_blnRecursive = " & v_blnRecursive
    l_colErrArgs.Add "v_strRootPath = " & v_strRootPath
    l_colErrArgs.Add "v_strResourceType = " & v_strResourceType
    l_colErrArgs.Add "oFolder.Path = " & oFolder.Path
    l_colErrArgs.Add "oFile.Path = " & oFile.Path
    l_colErrArgs.Add "oSubFolder.Path = " & oSubFolder.Path
    
    TrsRaiseError MODULENAME, "LoadResources", , , , l_colErrArgs, True
End Sub

'//-----could be moved into mCollectionLib
Private Function DictionaryToArray(ByVal v_objDictionary As Dictionary, _
                            Optional ByVal v_intBase As Integer = 0) As Variant
' v_intBase should be 0 or 1, if not assume 0
' return empty when collection is empty
    Dim aryReturn() As Variant
    Dim varKey As Variant
    Dim i As Integer
    
On Error GoTo eh

    If v_objDictionary.Count = 0 Then
        DictionaryToArray = Empty
        Exit Function
    End If

    If v_intBase = 1 Then
        i = 1
        ReDim aryReturn(1, i To v_objDictionary.Count)
    Else
        i = 0
        ReDim aryReturn(1, i To v_objDictionary.Count - 1)
    End If

    For Each varKey In v_objDictionary
        aryReturn(0, i) = varKey
        aryReturn(1, i) = Array(varKey, v_objDictionary.Item(varKey))
        i = i + 1
    Next
    DictionaryToArray = aryReturn
    Exit Function
eh:
    Err.Raise Err.Number
End Function

Private Function DictionaryToString(ByVal v_objDictionary As Dictionary, _
                            Optional ByVal singleQuote As Boolean = True, _
                            Optional ByVal delimiter1 As String = "=", _
                            Optional ByVal delimiter2 As String = "," & vbCrLf, _
                            Optional ByVal enclosedBy As String = "()") As String
    Dim varKey As Variant
    Dim str As String, i As Integer
    str = Left(enclosedBy, 1)
    i = 0
    
    For Each varKey In v_objDictionary
        '//-----keys
        If singleQuote Then
            str = str & EncloseWithSingleQuote(varKey)
        Else
            str = str & varKey
        End If
        str = str & delimiter1
        '//-----values
        If singleQuote Then
            str = str & EncloseWithSingleQuote(v_objDictionary.Item(varKey))
        Else
            str = str & v_objDictionary.Item(varKey)
        End If
        
        If Not (i = v_objDictionary.Count) Then
            str = str & delimiter2
        End If
        i = i + 1
    Next
    DictionaryToString = str & Right(enclosedBy, 1)

End Function

'//-----from dbEnvironment
Private Function SerializeRS(ByVal v_rsPreference As ADODB.Recordset) As String
    Dim OutputStream As ADODB.Stream
    Dim iChunkLength As Long
    Dim sSerialized As String
    
    iChunkLength = m_lngSerializationChunkSize
    
    Set OutputStream = CreateObject("ADODB.Stream")
'    If RecordsExist(v_rsPreference) Then
        v_rsPreference.Save OutputStream, adPersistXML
        Do While Not OutputStream.EOS
            sSerialized = sSerialized & OutputStream.ReadText(iChunkLength)
        Loop
'    End If
    
    SerializeRS = sSerialized
    
    Set OutputStream = Nothing

    Exit Function
eh:
    Set OutputStream = Nothing
    
    TrsRaiseError MODULENAME _
        , "SerializeRS" _
        , _
        , , , _
        , True
End Function

Private Function DeSerializeRSFromXML(ByVal sSerializedRecordset As String) As ADODB.Recordset
    Dim InputStream As ADODB.Stream
    Dim rs As ADODB.Recordset
    Dim iChunkLength As Long
    Dim sSerializedRS As String
    
    iChunkLength = m_lngSerializationChunkSize
    
    Set InputStream = CreateObject("ADODB.Stream")
    InputStream.Open
    
    InputStream.WriteText sSerializedRecordset, ADODB.StreamWriteEnum.adWriteChar
    InputStream.Position = 0

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open InputStream

    '//-----Clean Up-----//
    InputStream.Close
    
    Set DeSerializeRSFromXML = rs

    Set InputStream = Nothing
    Set rs = Nothing
    Exit Function
eh:
    Set InputStream = Nothing
    Set rs = Nothing
    
    TrsRaiseError MODULENAME _
        , "DeSerializeRSFromXML" _
        , _
        , , , _
        , True
End Function

Private Function DeSerializeRSFromBinary(ByVal varSerializedRecordset As Variant) As ADODB.Recordset
    Dim InputStream As ADODB.Stream
    Dim rs As ADODB.Recordset
    Dim iChunkLength As Long
    Dim sSerializedRS As String
    
    iChunkLength = m_lngSerializationChunkSize
    
    Set InputStream = CreateObject("ADODB.Stream")
    InputStream.Open
    InputStream.Type = adTypeBinary
    
    InputStream.Write varSerializedRecordset
    InputStream.Position = 0

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open InputStream

    '//-----Clean Up-----//
    InputStream.Close
    
    Set DeSerializeRSFromBinary = rs

    Set InputStream = Nothing
    Set rs = Nothing
    Exit Function
eh:
    Set InputStream = Nothing
    Set rs = Nothing
    
    TrsRaiseError MODULENAME _
        , "DeSerializeRSFromBinary" _
        , _
        , , , _
        , True
End Function


