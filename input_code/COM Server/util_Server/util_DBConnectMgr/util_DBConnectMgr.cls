VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsuDBConnectMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' The purpose is to compose a string from registry settings.
' This will ensure the identical connect string is used for
' all connection requests and best utilize connection pooling.

'To be deleted out begin
Public Enum TrsDBConnectManagerErrorEnum
    trErrGetConnectionString_F = 30008 + vbObjectError
    trErrDecryptPassword_F
    trErrGetConnectionInfo_F
End Enum
'To be deleted out end

Private Const mstrMODULENAME As String = "clsuDBConnectMgr"
Private Const encryptKey As String = "rsg"
Private Const regSubKey As String = "Software\Ameren\TRIS\DBConnectInfo\"
Private Const defaultProvider As String = "SHADOWDIRECT"
Private Const defaultProvider2 As String = "SHADOWDIRECT2"
Private Const defaultPeopleSoft As String = "PeopleSoft"
Private Const defaultEMPRV As String = "EMPRV"
Private Const intMAXRETRY As Integer = 2

Private dsn As String
'Private oledb As String
Private userId As String
Private decryptedPassword As String
Private registryPassword As String
Private isEncrypted As Boolean

Public Function GetConnectionString(Optional dataProvider As String) As String
'Answer a string that should be used as is for ODBC connection.
'Did not use optional parameter to avoid interface change when default changes.
    Dim strProvider As String
    
    On Error GoTo GetConnectString_EH
    
    If dataProvider = "" Then
        strProvider = defaultProvider
    Else
        strProvider = dataProvider
    End If
    
    '//-----Use corresponding registry key for COM+ partitions (TRISPartA, B, Q, T)-----//
    'If strProvider = defaultProvider Then '//-----only override if already "SHADOWDIRECT"-----//
        strProvider = GetCOMPartitionProvider(strProvider)
    'End If
    
    GetConnectionInfo (strProvider)
    
    If isEncrypted Then
        Err.Raise 12345, "", "ConnectionString encryption is not longer supported"
'        decryptedPassword = DecryptPassword(registryPassword)
    Else
        decryptedPassword = registryPassword
    End If

'    GetConnectionString = _
'        "DSN=" & dsn & _
'        ";UID=" & userId & _
'        ";PWD=" & decryptedPassword & _
'        ";OLE DB Services=-4"
'    If Trim(dsn) <> "" Then
        GetConnectionString = _
            "DSN=" & dsn & _
            ";UID=" & userId & _
            ";PWD=" & decryptedPassword
'    Else
'        GetConnectionString = _
'            oledb & _
'            ";UID=" & userId & _
'            ";PWD=" & decryptedPassword
'    End If
        
    Exit Function
            
GetConnectString_EH:

    TrsRaiseError _
            mstrMODULENAME, _
            "GetConnectionString", _
            trErrutGetConnectionString, _
            True, _
            DescArgsFrom(strProvider), , True
        
End Function

Public Function GetCommandTimeout(Optional dataProvider As String) As Long
    Dim strProvider As String
    
On Error GoTo eh
    
    If dataProvider = "" Then
        strProvider = defaultProvider
    Else
        strProvider = dataProvider
    End If
    
    '//-----Use corresponding registry key for COM+ partitions (TRISPartA, B, Q, T)-----//
    strProvider = GetCOMPartitionProvider(strProvider)
    
    GetCommandTimeout = GetCommandTimeoutInfo(strProvider)
    
    Exit Function
eh:
    TrsRaiseError mstrMODULENAME, "GetCommandTimeout", , True, DescArgsFrom(strProvider), , True
End Function

Private Function GetCOMPartitionProvider(ByVal v_sDefaultProvider As String) As String
    Dim o As COMAdmin.COMAdminCatalog
    Dim sProvider As String

    Set o = CreateObject("COMAdmin.COMAdminCatalog")
    
    '//-----only override if already "SHADOWDIRECT"-----//
    If o.CurrentPartitionName = "Base Application Partition" Then
        sProvider = v_sDefaultProvider
    Else
        If v_sDefaultProvider = defaultProvider Then
            sProvider = o.CurrentPartitionName
        ElseIf v_sDefaultProvider = defaultProvider2 Then
            sProvider = o.CurrentPartitionName & "2"
        ElseIf UCase(v_sDefaultProvider) = UCase(defaultPeopleSoft) Or _
                UCase(v_sDefaultProvider) = UCase(defaultEMPRV) Then
            sProvider = GetCOMPartitionProviderName(v_sDefaultProvider, o.CurrentPartitionName)
        Else
            sProvider = v_sDefaultProvider
        End If
    End If
    
'    sProvider = o.CurrentPartitionName
'    If sProvider = "Base Application Partition" Then
'        sProvider = v_sDefaultProvider
'    End If
    
    GetCOMPartitionProvider = sProvider
    
    Set o = Nothing
End Function

'//-----given "PeopleSoft" and "TRISPARTA" return "PeopleSoftA"
Private Function GetCOMPartitionProviderName(ByVal v_sProvider As String, _
                                            ByVal v_PartitionName As String) As String
    If v_PartitionName = "Base Application Partition" Then
        GetCOMPartitionProviderName = v_sProvider
    Else
        GetCOMPartitionProviderName = v_sProvider & Replace(v_PartitionName, "TRISPART", "")
    End If
End Function

'//-----used by dbDistributionMgr, moved here to hook into partition-----//
Public Function GetDB2Subsystem(Optional ByVal dataProvider As String) As String
    Dim strValue As String, strProvider As String
    Dim blnReturn As String
    Dim strFullSubKey As String
    Dim strValueName As String
    Dim ral As REGTool5.Registry
    
    
On Error GoTo eh
    Set ral = New REGTool5.Registry
    
    If dataProvider = "" Then
        strProvider = defaultProvider
    Else
        strProvider = dataProvider
    End If
    
    '//-----Use corresponding registry key for COM+ partitions (TRISPartA, B, Q, T)-----//
    'If strProvider = defaultProvider Then '//-----only override if already "SHADOWDIRECT"-----//
        strProvider = GetCOMPartitionProvider(strProvider)
    'End If
    
    strFullSubKey = regSubKey & strProvider
    strValueName = "DB2SubSys"

    blnReturn = ral.GetKeyValue( _
            HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, strValue)
    GetDB2Subsystem = strValue
    Set ral = Nothing
    Exit Function
    
eh:
    Set ral = Nothing
    TrsRaiseError _
            mstrMODULENAME, _
            "GetDB2Subsystem", _
            , _
            True, _
            DescArgsFrom(strFullSubKey, strValueName)

End Function

''//-----used by utilADOHelper (dbScriptsRoot), moved here to hook into partition-----//
'Public Function IsDBPlatformSQL2005(ByVal dataProvider As String) As Boolean
'    Dim sDBPlatform As String
'
'    sDBPlatform = GetRegistryStringValue(dataProvider, "dbPlatform")
'
'    IsDBPlatformSQL2005 = (sDBPlatform = "SQL2005")
'End Function


'//-----used by utilADOHelper (dbScriptsRoot), moved here to hook into partition-----//
Public Function GetRegistryStringValue(ByVal dataProvider As String, _
                                    ByVal v_strValueName As String) As String
    Dim strValue As String, strProvider As String
    Dim blnReturn As String
    Dim strFullSubKey As String
    Dim strValueName As String
    Dim ral As REGTool5.Registry
    
On Error GoTo eh
    Set ral = New REGTool5.Registry
    
    If dataProvider = "" Then
        strProvider = defaultProvider
    Else
        strProvider = dataProvider
    End If
    
    '//-----Use corresponding registry key for COM+ partitions (TRISPartA, B, Q, T)-----//
    'If strProvider = defaultProvider Then '//-----only override if already "SHADOWDIRECT"-----//
        strProvider = GetCOMPartitionProvider(strProvider)
    'End If
    
    strFullSubKey = regSubKey & strProvider

    blnReturn = ral.GetKeyValue(HKEY_LOCAL_MACHINE, strFullSubKey, v_strValueName, strValue)
    GetRegistryStringValue = strValue
    Set ral = Nothing
    Exit Function
eh:
    Set ral = Nothing
    TrsRaiseError _
            mstrMODULENAME, _
            "GetRegistryStringValue", _
            , _
            True, _
            DescArgsFrom(strFullSubKey, v_strValueName)
End Function

'//-----this has been exposed for use by utilServerInfo-----//
Public Function GetDataSourceName(Optional ByVal dataProvider As String) As String
    Dim strProvider As String, strFullSubKey As String, strValueName As String
    Dim blnReturn As Boolean
    Dim ral As REGTool5.Registry
    Dim l_strDSN As String
    
    On Error GoTo eh
    
    Set ral = New REGTool5.Registry
    
    If dataProvider = "" Then
        strProvider = defaultProvider
    Else
        strProvider = dataProvider
    End If
    
    '//-----Use corresponding registry key for COM+ partitions (TRISPartA, B, Q, T)-----//
    'If strProvider = defaultProvider Then '//-----only override if already "SHADOWDIRECT"-----//
        strProvider = GetCOMPartitionProvider(strProvider)
    'End If
    
    strFullSubKey = regSubKey & strProvider
    strValueName = "DSN"

    blnReturn = ral.GetKeyValue( _
            HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, l_strDSN)
    Set ral = Nothing
    
    GetDataSourceName = l_strDSN
    Exit Function
eh:
    Set ral = Nothing
    TrsRaiseError mstrMODULENAME, "GetDataSourceName", , True, DescArgsFrom(strProvider), , True
End Function

'//-----this has been exposed for use by utilServerInfo-----//
Public Function GetUserId(Optional ByVal dataProvider As String) As String
    Dim strProvider As String, strFullSubKey As String, strValueName As String
    Dim blnReturn As Boolean
    Dim ral As REGTool5.Registry
    Dim l_strUID As String
    
    On Error GoTo eh
    
    Set ral = New REGTool5.Registry
    
    If dataProvider = "" Then
        strProvider = defaultProvider
    Else
        strProvider = dataProvider
    End If
    
    '//-----Use corresponding registry key for COM+ partitions (TRISPartA, B, Q, T)-----//
    'If strProvider = defaultProvider Then '//-----only override if already "SHADOWDIRECT"-----//
        strProvider = GetCOMPartitionProvider(strProvider)
    'End If
    
    strFullSubKey = regSubKey & strProvider
    strValueName = "UID"

    blnReturn = ral.GetKeyValue( _
            HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, l_strUID)
    Set ral = Nothing
    
    GetUserId = l_strUID
    Exit Function
eh:
    Set ral = Nothing
    TrsRaiseError mstrMODULENAME, "GetUserId", , True, DescArgsFrom(strProvider), , True
End Function

Public Function GetSchema(Optional ByVal dataProvider As String) As String
    Dim strProvider As String, strFullSubKey As String, strValueName As String
    Dim blnReturn As Boolean
    Dim ral As REGTool5.Registry
    Dim l_strSCHEMA As String
    
    On Error GoTo eh
    
    Set ral = New REGTool5.Registry
    
    If dataProvider = "" Then
        strProvider = defaultProvider
    Else
        strProvider = dataProvider
    End If
    
    '//-----Use corresponding registry key for COM+ partitions (TRISPartA, B, Q, T)-----//
    'If strProvider = defaultProvider Then '//-----only override if already "SHADOWDIRECT"-----//
        strProvider = GetCOMPartitionProvider(strProvider)
    'End If
    
    strFullSubKey = regSubKey & strProvider
    strValueName = "SCHEMA"

    blnReturn = ral.GetKeyValue(HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, l_strSCHEMA)
    Set ral = Nothing
    
    GetSchema = l_strSCHEMA
    Exit Function
eh:
    Set ral = Nothing
    TrsRaiseError mstrMODULENAME, "GetSchema", , True, DescArgsFrom(strProvider), , True
End Function

'//-----for LUW
Public Function AreDSNsEqual(Optional ByVal dataProvider As String, _
                                Optional ByVal dataProvider2 As String) As Boolean
    Dim strProvider As String, strProvider2 As String
    
    On Error GoTo eh
    
    If dataProvider = "" Then
        strProvider = defaultProvider
    Else
        strProvider = dataProvider
    End If
    If dataProvider2 = "" Then
        strProvider2 = defaultProvider2
    Else
        strProvider2 = dataProvider2
    End If
    
    AreDSNsEqual = Trim(GetDataSourceName(strProvider)) = Trim(GetDataSourceName(strProvider2))
    
    Exit Function
eh:
    TrsRaiseError mstrMODULENAME, "AreDSNsEqual", , True, DescArgsFrom(strProvider, strProvider2), , True
End Function

Public Function IsDSNLUW(Optional ByVal dataProvider As String) As Boolean
    Dim strProvider As String
    
    On Error GoTo eh
    
    If dataProvider = "" Then
        strProvider = defaultProvider
    Else
        strProvider = dataProvider
    End If
    
    IsDSNLUW = InStr(1, GetDataSourceName(strProvider), "LUW") > 0
    
    Exit Function
eh:
    TrsRaiseError mstrMODULENAME, "IsDSNLUW", , True, DescArgsFrom(strProvider), , True
End Function

Public Function GetDB2SubSystemFromDSN(Optional ByVal dataProvider As String) As String
    Dim strProvider As String, lsDSN As String
    
    On Error GoTo eh
    
    If dataProvider = "" Then
        strProvider = defaultProvider
    Else
        strProvider = dataProvider
    End If
    
    lsDSN = GetDataSourceName(strProvider)
    
    If IsDSNLUW(strProvider) Then
        GetDB2SubSystemFromDSN = Trim(Left(lsDSN, 7))
    Else
        GetDB2SubSystemFromDSN = Trim(Left(lsDSN, 4))
    End If
    
    Exit Function
eh:
    TrsRaiseError mstrMODULENAME, "GetDB2SubSystemFromDSN", , True, DescArgsFrom(strProvider), , True
End Function

'Private Function DecryptPassword(pwd As String) As String
'    Dim encr As ARCryptoLib.RC2
'    Dim b64 As ARCryptoLib.Base64
'
'    On Error GoTo DecryptPassword_EH
'
'    Set encr = New ARCryptoLib.RC2
'    Set b64 = New ARCryptoLib.Base64
'    encr.Key = encryptKey
'    DecryptPassword = encr.Decrypt(b64.Decode(pwd))
'    Exit Function
'
'DecryptPassword_EH:
'
'    TrsRaiseError _
'            mstrMODULENAME, _
'            "DecryptPassword", _
'            , _
'            True
'
'End Function

Private Function GetConnectionInfo(strProvider As String)
'Access the registry and store connection values in variables.

    Dim blnReturn As Boolean
    Dim strEncrypt As String
    Dim strFullSubKey As String
    Dim strValueName As String
    Dim ral As REGTool5.Registry
    
    
    On Error GoTo GetConnectionInfo_EH
    Set ral = New REGTool5.Registry
    
    strFullSubKey = regSubKey & strProvider
    
    strValueName = "DSN"
    blnReturn = ral.GetKeyValue( _
            HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, dsn)
    
'    strValueName = "OLEDB"
'    blnReturn = ral.GetKeyValue( _
'            HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, oledb)
    
    strValueName = "UID"
    blnReturn = ral.GetKeyValue( _
            HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, userId)
    
    strValueName = "PWD"
    blnReturn = ral.GetKeyValue( _
            HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, registryPassword)
                
    strValueName = "ENCR"
    blnReturn = ral.GetKeyValue( _
            HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, strEncrypt)
    If blnReturn Then
        If strEncrypt = "0" Then
            isEncrypted = False
        Else
            isEncrypted = True
        End If
    End If
    
    Exit Function
    
GetConnectionInfo_EH:

    TrsRaiseError _
            mstrMODULENAME, _
            "GetConnectionInfo", _
            , _
            True, _
            DescArgsFrom(strFullSubKey, strValueName)

End Function

Private Function GetCommandTimeoutInfo(strProvider As String)
'Access the registry and store connection values in variables.
    Dim blnReturn As Boolean
    Dim strFullSubKey As String
    Dim strValueName As String
    Dim ral As REGTool5.Registry
    Dim CommandTimeout As String
    
    On Error GoTo eh
    Set ral = New REGTool5.Registry
    
    GetCommandTimeoutInfo = 30
    strFullSubKey = regSubKey & strProvider
    
    strValueName = "CommandTimeout"
    blnReturn = ral.GetKeyValue(HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, CommandTimeout)
    
    If blnReturn Then
        If IsNumeric(CommandTimeout) Then
            GetCommandTimeoutInfo = CLng(CommandTimeout)
        End If
    End If
    
    Exit Function
eh:
    TrsRaiseError mstrMODULENAME, "GetCommandTimeoutInfo", , True, DescArgsFrom(strFullSubKey, strValueName)
End Function
'Public Function TweakSQL(ByVal v_strSQL As String, Optional ByVal v_blnIsSelect As Boolean = True) As String
'    Dim blnSQL2005 As Boolean
'
'    blnSQL2005 = IsDBPlatformSQL2005("")
'
'    TweakSQL = TweakSQL2(v_strSQL, v_blnIsSelect, blnSQL2005)
'End Function
'
'Private Function TweakSQL2(ByVal v_strSQL As String, _
'                            ByVal v_blnIsSelect As Boolean, _
'                            ByVal v_blnSQL2005 As Boolean) As String
'
'    Dim oUtil As TRISServerUtilityCOM.Utility
'
'    If v_blnSQL2005 Then
'        Set oUtil = CreateObject("TRISServerUtilityCOM.Utility")
'
'        TweakSQL2 = oUtil.TweakSQL(v_strSQL, v_blnIsSelect)
'        Set oUtil = Nothing
'    Else
'        TweakSQL2 = v_strSQL
'    End If
'End Function

