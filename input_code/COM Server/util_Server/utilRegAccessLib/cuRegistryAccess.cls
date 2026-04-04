VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cuRegistryAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Assuming 32-bit platform

Public Enum RootKeysEnum
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Private lRetVal As Long     'Return value of API
Private phkResult As Long   'handle to an open key

Public Function GetSetting(rootKey As RootKeysEnum, subKey As String, namedValue As String) As Variant
    
    On Error GoTo GetSetting_EH
    phkResult = OpenKey(rootKey, subKey, KEY_QUERY_VALUE)
    GetSetting = QueryValue(phkResult, namedValue)
    RegCloseKey (phkResult)
    Exit Function

GetSetting_EH:
    If phkResult <> 0 Then RegCloseKey (phkResult)
    Err.Raise Err.Number

End Function

Public Function TryGetSetting(ByVal rootKey As RootKeysEnum, ByVal subKey As String, ByVal namedValue As String, ByRef setting As Variant) As Boolean
    Dim rc As Boolean
    On Error GoTo eh
    
    rc = TryOpenKey(rootKey, subKey, KEY_QUERY_VALUE, phkResult)
    If rc Then
        rc = TryQueryValue(phkResult, namedValue, setting)
        TryGetSetting = rc
        rc = TryCloseKey(phkResult)
    End If
    Exit Function
eh:
    If phkResult <> 0 Then TryCloseKey phkResult
    TryGetSetting = False
End Function

Public Function GetAllSettings(rootKey As RootKeysEnum, subKey As String) As Variant
'Answer a 2 dim variant array with all the named values under subKey.
'    Dim phkResult As Long
    Dim varrValues As Variant
    
    On Error GoTo GetAllSettings_EH
    phkResult = OpenKey(rootKey, subKey, KEY_QUERY_VALUE)
    varrValues = QueryAllValues(phkResult)

    RegCloseKey (phkResult)
    GetAllSettings = varrValues
    Exit Function

GetAllSettings_EH:
    If phkResult <> 0 Then RegCloseKey (phkResult)
    Err.Raise Err.Number

End Function

Public Function AllKeys( _
            parentKey As Long, _
            subKey As String, _
            blnRecursive As Boolean) As Variant
' hKeySource and hKeyTarget are open keys.
' Recursive method that will enumerate all keys and call CopyKey.
' Must open/create a target key for each subkey enumerated.
    Dim lvarValues As Variant
    Dim lSubkeyIndex As Long
    Dim lReturn As Long
    Dim lSubkeyBufferSize As Long
    Dim strSubkeyName As String
    Dim lpftLastWriteTime As FILETIME
    Dim varReturnEnum As Variant
    Dim strReturnMe As String
    Dim varReturnAll As Variant
    Dim lphKey As Long, i As Long, j As Long, k As Long
    
On Error GoTo eh
    lphKey = OpenKey(parentKey, subKey, KEY_READ)
'    Debug.Print "opening " & parentKey & " " & subKey & "  " & lphKey
    'strReturnMe = "\" & subKey
    strReturnMe = subKey
    'copy all keys
    
    ReDim varReturnEnum(0)
    varReturnEnum(0) = strReturnMe
    
    If blnRecursive Then
        lSubkeyIndex = 0
        Do Until lReturn = ERROR_NO_MORE_ITEMS
            lSubkeyBufferSize = 255
            strSubkeyName = String(lSubkeyBufferSize, 0)
            lReturn = RegEnumKeyEx(lphKey, lSubkeyIndex, strSubkeyName, lSubkeyBufferSize, 0&, 0&, 0&, lpftLastWriteTime)
            Select Case lReturn
                Case ERROR_SUCCESS, ERROR_MORE_DATA
                    varReturnAll = AllKeys(lphKey, Left(strSubkeyName, lSubkeyBufferSize), blnRecursive)
                    'varReturnEnum = varReturnEnum & varReturnAll
                    i = UBound(varReturnEnum)
                    ReDim Preserve varReturnEnum(i + UBound(varReturnAll) + 1)
                    k = 0
                    For j = i + 1 To i + 1 + UBound(varReturnAll)
                        varReturnEnum(j) = subKey & "\" & varReturnAll(k)
                        k = k + 1
                    Next j
                    
                Case ERROR_NO_MORE_ITEMS
    '                do nothing
                Case Else
                    Err.Raise vbObjectError + 2001, "utilRegAccessLib", _
                        "Error enumerating registry at key " & _
                        CStr(lSubkeyIndex) & " of " & strSubkeyName & "."
            End Select
            lSubkeyIndex = lSubkeyIndex + 1
        Loop
    End If
'    Debug.Print "closing " & lphKey
    CloseKey lphKey
        
    'AllKeys = strReturnMe & varReturnEnum
    AllKeys = varReturnEnum
    
    Exit Function

eh:
    If lphKey Then CloseKey lphKey
    Err.Raise Err.Number
    
End Function

Public Function OpenKey(ByVal rootKey As RootKeysEnum, subKey As String, lSAM As Long) As Long
' Answer a handle to an open key.

    lRetVal = RegOpenKeyEx(rootKey, subKey, 0&, lSAM, phkResult)
    If lRetVal <> ERROR_SUCCESS Then
        Err.Raise vbObjectError + 1001, "cuRegistryAccess", _
            "Error opening key: " & subKey & " under root: " & RootKeyName(rootKey) & ".  API return code: " & CStr(lRetVal)
    End If
    OpenKey = phkResult
    
End Function

Private Function TryOpenKey(ByVal rootKey As RootKeysEnum, ByVal subKey As String, ByVal lSAM As Long, ByRef r_phkResult As Long) As Boolean
' Answer a return code, byref contains a handle to an open key.

    lRetVal = RegOpenKeyEx(rootKey, subKey, 0&, lSAM, r_phkResult)
    TryOpenKey = (lRetVal = ERROR_SUCCESS)
    
End Function

Public Sub CloseKey(ByVal hKey As Long)

    lRetVal = RegCloseKey(hKey)
    If lRetVal <> ERROR_SUCCESS Then
        Err.Raise vbObjectError + 1001, "cuRegistryAccess", "Error closing key.  API return code: " & CStr(lRetVal)
    End If
    
End Sub

Private Function TryCloseKey(ByVal hKey As Long) As Boolean

    lRetVal = RegCloseKey(hKey)
    TryCloseKey = (lRetVal = ERROR_SUCCESS)
    
End Function

Public Function QueryValue(ByVal hKey As Long, sName As String) As Variant
'Must call OpenKey before this method and pass hKey, a handle to an open key.
    Dim lType As Long
    Dim szBuffer As String
    Dim lBufferSize As Long
    Dim lRetVal As Long
        
    lRetVal = RegQueryValueEx(hKey, sName, 0&, lType, 0&, lBufferSize)
    If lRetVal = ERROR_SUCCESS Then
        szBuffer = String(lBufferSize, 0)
        lRetVal = RegQueryValueExStr(hKey, sName, 0&, lType, szBuffer, lBufferSize)
        If lRetVal = ERROR_SUCCESS Then
            QueryValue = Left$(szBuffer, lBufferSize - 1)
        Else
            Err.Raise vbObjectError + 1000, "cuRegistryAccess", "Error querying registry value " & sName & ".  API return code: " & CStr(lRetVal)
        End If
    Else
        Err.Raise vbObjectError + 1000, "cuRegistryAccess", "Error querying registry value " & sName & ".  API return code: " & CStr(lRetVal)
    End If

End Function

Private Function TryQueryValue(ByVal hKey As Long, ByVal sName As String, ByRef value As Variant) As Boolean
'Must call OpenKey before this method and pass hKey, a handle to an open key.
    Dim lType As Long
    Dim szBuffer As String
    Dim lBufferSize As Long
    Dim lRetVal As Long
        
    lRetVal = RegQueryValueEx(hKey, sName, 0&, lType, 0&, lBufferSize)

    If lRetVal = ERROR_SUCCESS Then
        szBuffer = String(lBufferSize, 0)
        lRetVal = RegQueryValueExStr(hKey, sName, 0&, lType, szBuffer, lBufferSize)

        If lRetVal = ERROR_SUCCESS Then
            value = Left$(szBuffer, lBufferSize - 1)
        End If
    End If
    TryQueryValue = (lRetVal = ERROR_SUCCESS)
End Function

Public Function QueryAllValues(ByVal hKey As Long) As Variant
'Answer a variant array.
'Must call OpenKey before this method and pass hKey, a handle to an open key.
    Dim lRetVal As Long
    Dim varrValues() As Variant
    
' RegQueryInfoKey parameters
    Dim szbClass As String
    Dim lClassSize As Long
    Dim lValuesSize As Long
    Dim lMaxValueNameSize As Long
    Dim lMaxValueSize As Long
    
' RegEnumValue parameters
    Dim lIndex As Long
    Dim szbValueName As String
    Dim lValueNameSize As Long
    Dim lType As Long
    Dim szbValueData As String
    Dim lValueDataSize As Long
    
    szbClass = String(255, 0)
    lRetVal = RegQueryInfoKey(hKey, szbClass, lClassSize, 0&, 0&, 0&, 0&, lValuesSize, _
                lMaxValueNameSize, lMaxValueSize, 0&, 0&)
        
    If lRetVal <> 0 Then
        Err.Raise vbObjectError + 1002, "QueryRegistry", _
            "Call RegQueryInfoKey failed.  Return code: " & CStr(lRetVal) & "."
    End If
    
    If lValuesSize > 0 Then
        ReDim varrValues(0 To lValuesSize - 1, 0 To 1)
        
        For lIndex = 0 To lValuesSize - 1
            szbValueName = String(lMaxValueNameSize + 1, 0)
            lValueNameSize = lMaxValueNameSize + 1
            szbValueData = String(lMaxValueSize, 0)
            lValueDataSize = lMaxValueSize
            lRetVal = RegEnumValue(hKey, lIndex, szbValueName, lValueNameSize, 0&, lType, szbValueData, lValueDataSize)
            If lRetVal = ERROR_SUCCESS Then
                varrValues(lIndex, 0) = Left$(szbValueName, lValueNameSize)
                varrValues(lIndex, 1) = Left$(szbValueData, lValueDataSize - 1)
            Else
                Err.Raise vbObjectError + 2001, "ConnectionManager", _
                    "Error enumerating registry at index " & _
                    CStr(lIndex) & " of " & CStr(lValuesSize) & " items."
            End If
        Next
    End If
    
    QueryAllValues = varrValues

End Function

Private Function RootKeyName(rootKey As RootKeysEnum) As String

    Select Case rootKey
        Case HKEY_CLASSES_ROOT
            RootKeyName = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_USER
            RootKeyName = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE
            RootKeyName = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS
            RootKeyName = "HKEY_USERS"
        Case Else
            RootKeyName = "Unknown Root Key"
    End Select
    
End Function

Public Function DoesKeyExist(ByVal rootKey As RootKeysEnum, ByVal subKey As String, Optional ByVal lSAM As Long = &H1) As Boolean
' Answer a handle to an open key.
On Error GoTo eh
    Dim rc As Boolean
    '//Dim phkResult As Long
    
    rc = TryOpenKey(rootKey, subKey, lSAM, phkResult)
    
    DoesKeyExist = rc
    If rc Then
        rc = TryCloseKey(phkResult)
    End If
    Exit Function
eh:
    Err.Raise Err.Number, "DoesKeyExist"
End Function

Public Sub SetStringValue(ByVal rootKey As RootKeysEnum, ByVal subKey As String, ByVal namedValue As String, ByVal value As Variant)
    Dim rc As Boolean
    Dim sValue As String
On Error GoTo eh
    
    phkResult = OpenOrCreateKey(rootKey, subKey, KEY_ALL_ACCESS)

    sValue = value & vbNullChar
    lRetVal = RegSetValueExString(phkResult, namedValue, 0, REG_SZ, sValue, Len(sValue))
    
    If (lRetVal <> ERROR_SUCCESS) Then
        Err.Raise vbObjectError + 1001, "cuRegistryAccess.SetStringValue", "Error setting string value.  API return code: " & CStr(lRetVal)
    End If

    CloseKey phkResult
    
    Exit Sub
eh:
    If phkResult <> 0 Then CloseKey phkResult
    Err.Raise Err.Number ', "SetStringValue"
End Sub

Public Function TrySetStringValue(ByVal rootKey As RootKeysEnum, ByVal subKey As String, ByVal namedValue As String, ByVal value As Variant) As Boolean
    Dim rc As Boolean
    Dim sValue As String
On Error GoTo eh
    
    rc = TryOpenOrCreateKey(rootKey, subKey, KEY_ALL_ACCESS, phkResult)
    If rc Then
        sValue = value & vbNullChar
        lRetVal = RegSetValueExString(phkResult, namedValue, 0, REG_SZ, sValue, Len(sValue))
        
        TrySetStringValue = (lRetVal = ERROR_SUCCESS)
        rc = TryCloseKey(phkResult)
    End If
    Exit Function
eh:
    If phkResult <> 0 Then TryCloseKey phkResult
    TrySetStringValue = False
End Function

Private Function OpenOrCreateKey(ByVal rootKey As RootKeysEnum, ByVal subKey As String, ByVal lSAM As Long) As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim OpenOrAddResult As Long
    Dim r_phkResult As Long
On Error GoTo eh
    sa.nLength = Len(sa)
    sa.lpSecurityDescriptor = 0
    sa.bInheritHandle = CLng(True)
    
    lRetVal = RegCreateKeyEx(rootKey, subKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, lSAM, sa, r_phkResult, OpenOrAddResult)
    If (lRetVal <> ERROR_SUCCESS) Then
        Err.Raise vbObjectError + 1001, "cuRegistryAccess.OpenOrCreateKey", "Error opening or creating key.  API return code: " & CStr(lRetVal)
    End If
    OpenOrCreateKey = r_phkResult
    Exit Function
eh:
    Err.Raise Err.Number ', "OpenOrCreateKey"
End Function

Private Function TryOpenOrCreateKey(ByVal rootKey As RootKeysEnum, ByVal subKey As String, ByVal lSAM As Long, ByRef r_phkResult As Long) As Boolean
    Dim sa As SECURITY_ATTRIBUTES
    Dim OpenOrAddResult As Long
On Error GoTo eh
    sa.nLength = Len(sa)
    sa.lpSecurityDescriptor = 0
    sa.bInheritHandle = CLng(True)
    
    lRetVal = RegCreateKeyEx(rootKey, subKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, lSAM, sa, r_phkResult, OpenOrAddResult)
    TryOpenOrCreateKey = (lRetVal = ERROR_SUCCESS)
'   If lRetVal = ERROR_SUCCESS Then
'    Select Case OpenOrAddResult
'    Case REG_CREATED_NEW_KEY
'    Case REG_OPENED_EXISTING_KEY
'    End Select
'   End If
    Exit Function
eh:
    TryOpenOrCreateKey = False
End Function

'Public Function GetLastReturnCode() As Long
'On Error GoTo eh
'    GetLastReturnCode = lRetVal
'    Exit Function
'eh:
'    GetLastReturnCode = -99999
'End Function
'
'Public Function GetReturnCodeDescription(ByVal v_retVal As Long) As String
'On Error GoTo eh
'
'    Select Case v_retVal
'    Case ERROR_SUCCESS
'        GetReturnCodeDescription = "Success"
'    Case ERROR_BADDB
'        GetReturnCodeDescription = "Corrupted registry"
'    Case ERROR_BADKEY
'        GetReturnCodeDescription = "Invalid registry key"
'    Case ERROR_ACCESS_DENIED
'        GetReturnCodeDescription = "Requested permissions not available"
'    Case ERROR_INVALID_HANDLE
'        GetReturnCodeDescription = "Invalid handle or top-level key"
'    Case ERROR_INVALID_PARAMETER
'        GetReturnCodeDescription = "Bad parameter to a Win32 API function"
'    Case ERROR_INVALID_PARAMETERS
'        GetReturnCodeDescription = "Bad parameters to a Win32 API function"
'    Case ERROR_CANTOPEN
'        GetReturnCodeDescription = "Cannot open registry key"
'    Case ERROR_CANTREAD
'        GetReturnCodeDescription = "Cannot read from registry key"
'    Case ERROR_CANTWRITE
'        GetReturnCodeDescription = "Cannot write to registry key"
'    Case Else
'        GetReturnCodeDescription = "code unknown"
'    End Select
'
'    Exit Function
'eh:
'    GetReturnCodeDescription = "code unknown (resulted in error)"
'End Function
