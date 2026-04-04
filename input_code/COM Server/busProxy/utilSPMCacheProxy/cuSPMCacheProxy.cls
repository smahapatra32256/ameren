VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cuSPMCacheProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function GetIsolationMode ( _
	ByVal v_strGroupName As String,  _
	ByRef r_blnExists As Boolean _
	)  As  Long
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    GetIsolationMode = l_objcuSPMCache.GetIsolationMode(v_strGroupName, r_blnExists)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "GetIsolationMode", , , , , True

End Function


Public Function GetReleaseMode ( _
	ByVal v_strGroupName As String,  _
	ByRef r_blnExists As Boolean _
	)  As  Long
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    GetReleaseMode = l_objcuSPMCache.GetReleaseMode(v_strGroupName, r_blnExists)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "GetReleaseMode", , , , , True

End Function


Public Function GetSupportsKeys ( _
	ByVal v_strGroupName As String,  _
	ByRef r_blnExists As Boolean _
	)  As  Boolean
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    GetSupportsKeys = l_objcuSPMCache.GetSupportsKeys(v_strGroupName, r_blnExists)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "GetSupportsKeys", , , , , True

End Function


Public Function CreateGroup ( _
	ByVal v_strGroupName As String,  _
	ByRef r_blnExists As Boolean,  _
	Optional ByRef r_lngIsolationMode As Long = "1",  _
	Optional ByRef r_lngReleaseMode As Long = "1",  _
	Optional ByRef r_blnKeySupport As Boolean = "False" _
	)  As  Variant
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    CreateGroup = l_objcuSPMCache.CreateGroup(v_strGroupName, r_blnExists, r_lngIsolationMode, r_lngReleaseMode, r_blnKeySupport)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "CreateGroup", , , , , True

End Function


Public Function GroupExists ( _
	ByVal v_strGroupName As String _
	)  As  Boolean
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    GroupExists = l_objcuSPMCache.GroupExists(v_strGroupName)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "GroupExists", , , , , True

End Function


Public Function Item ( _
	ByVal v_strKey As String,  _
	ByVal v_strGroupName As String,  _
	ByRef r_blnExists As Boolean,  _
	ByRef r_blnGroupExists As Boolean _
	)  As  Variant
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    Item = l_objcuSPMCache.Item(v_strKey, v_strGroupName, r_blnExists, r_blnGroupExists)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "Item", , , , , True

End Function


Public Function Add ( _
	ByVal v_strKey As String,  _
	ByVal v_varValue As Variant,  _
	ByVal v_strGroupName As String _
	)  As  Long
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    Add = l_objcuSPMCache.Add(v_strKey, v_varValue, v_strGroupName)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "Add", , , , , True

End Function


Public Function AddAll ( _
	ByVal v_varKeyValueArray As Variant,  _
	ByVal v_strGroupName As String _
	)  As  Long
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    AddAll = l_objcuSPMCache.AddAll(v_varKeyValueArray, v_strGroupName)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "AddAll", , , , , True

End Function


Public Function Count ( _
	ByVal v_strGroupName As String _
	)  As  Long
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    Count = l_objcuSPMCache.Count(v_strGroupName)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "Count", , , , , True

End Function


Public Function Remove ( _
	ByVal v_strKey As String,  _
	ByVal v_strGroupName As String _
	)  As  Long
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    Remove = l_objcuSPMCache.Remove(v_strKey, v_strGroupName)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "Remove", , , , , True

End Function


Public Function RemoveAll ( _
	ByVal v_varKeysArray As Variant,  _
	ByVal v_strGroupName As String _
	)  As  Long
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    RemoveAll = l_objcuSPMCache.RemoveAll(v_varKeysArray, v_strGroupName)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "RemoveAll", , , , , True

End Function


Public Function Keys ( _
	ByVal v_strGroupName As String _
	)  As  Variant
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    Keys = l_objcuSPMCache.Keys(v_strGroupName)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "Keys", , , , , True

End Function


Public Function Invalidate ( _
	ByVal v_strKey As String,  _
	ByVal v_strGroupName As String _
	)  As  Long
  
    Dim l_objcuSPMCache As utilSPMCache.cuSPMCache
    
On Error GoTo eh

    Set l_objcuSPMCache = CreateObject("utilSPMCache.cuSPMCache")
    
    Invalidate = l_objcuSPMCache.Invalidate(v_strKey, v_strGroupName)
    
    Set l_objcuSPMCache = Nothing
    
    Exit Function
    
eh:
    Set l_objcuSPMCache = Nothing

    TrsRaiseError "utilSPMCache.cuSPMCache", "Invalidate", , , , , True

End Function