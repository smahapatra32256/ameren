VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULENAME As String = "dbSecurity.cdFunSecurity"
Private Const m_strCMODULENAME As String = MODULENAME

' Developer Group constants
Private Const DEVELOPER_CACHE = "DEVELOPERGROUP"
Private Const DEVELOPER_CACHE_KEY = 0
Private Const DEVELOPER_CACHE_GROUPNAME = 1
Private Const DEVELOPER_CACHE_EMPNUM = 2
Private Const DEVELOPER_CACHE_LAST_SLOT = 2

' Training Instructor Group constants
Private Const TRAININGINSTRUCTOR_CACHE = "TRAININGINSTRUCTORGROUP"
Private Const TRAININGINSTRUCTOR_CACHE_KEY = 0
Private Const TRAININGINSTRUCTOR_CACHE_GROUPNAME = 1
Private Const TRAININGINSTRUCTOR_CACHE_EMPNUM = 2
Private Const TRAININGINSTRUCTOR_CACHE_LAST_SLOT = 2

' Environment Group constants
Private Const ENVIRONMENT_CACHE = "ENVIRONMENTGROUP"
Private Const ENVIRONMENT_CACHE_KEY = 0
Private Const ENVIRONMENT_CACHE_ENVIRONMENT = 1
Private Const ENVIRONMENT_CACHE_EMPNUM = 2
Private Const ENVIRONMENT_CACHE_LAST_SLOT = 2

' Redirect URL Group constants
Private Const REDIRECTURL_CACHE = "REDIRECTURLGROUP"
Private Const REDIRECTURL_CACHE_KEY = 0
Private Const REDIRECTURL_CACHE_ENVIRONMENT = 1
Private Const REDIRECTURL_CACHE_URL = 2
Private Const REDIRECTURL_CACHE_LAST_SLOT = 2

' Alias Group constants
Private Const ALIAS_CACHE = "ALIASGROUP"
Private Const ALIAS_CACHE_KEY = 0
Private Const ALIAS_CACHE_EMP_NUM = 1
Private Const ALIAS_CACHE_ALIAS_EMP_NUM = 2
Private Const ALIAS_CACHE_LAST_SLOT = 2

' Priority Group constants
Private Const PRIORITY_CACHE = "PRIORITYGROUP"
Private Const PRIORITY_CACHE_KEY = 0
Private Const PRIORITY_CACHE_PRIORITY = 1
Private Const PRIORITY_CACHE_WINGROUP = 2
Private Const PRIORITY_CACHE_EMP_NUM = 3
Private Const PRIORITY_CACHE_LAST_SLOT = 3

' Priority Default Group constants (use the Approval table for the default)
Private Const PRIORITYDEFAULT_CACHE = "PRIORITYDEFAULTGROUP"
Private Const PRIORITYDEFAULT_CACHE_KEY = 0
Private Const PRIORITYDEFAULT_CACHE_EMP_NUM = 1
Private Const PRIORITYDEFAULT_CACHE_PRIORITY = 2
Private Const PRIORITYDEFAULT_CACHE_LAST_SLOT = 2

' Priority Class Group constants
Private Const PRIORITYCLASS_CACHE = "PRIORITYCLASSGROUP"
Private Const PRIORITYCLASS_CACHE_KEY = 0
Private Const PRIORITYCLASS_CACHE_PRIORITY = 1
Private Const PRIORITYCLASS_CACHE_CLASS = 2
Private Const PRIORITYCLASS_CACHE_LAST_SLOT = 2

' SysParms Group constants
Private Const SYSPARMS_CACHE = "SYSPARMSGROUP"
Private Const SYSPARMS_CACHE_KEY = 0
Private Const SYSPARMS_CACHE_PARMNUM = 1
Private Const SYSPARMS_CACHE_PARMSUFFIX = 2
Private Const SYSPARMS_CACHE_PARMDESC = 3
Private Const SYSPARMS_CACHE_PARMVALUE = 4
Private Const SYSPARMS_CACHE_LAST_SLOT = 4

'' Windows Groups/Users Group constants
'Private Const WINGROUP_CACHE = "WINGROUPGROUP"
'Private Const WINGROUP_CACHE_KEY = 0
'Private Const WINGROUP_CACHE_GROUPNAME = 1
'Private Const WINGROUP_CACHE_EMPNUM = 2
'Private Const WINGROUP_CACHE_LAST_SLOT = 2

'//-----Determine Developer by using Rule 102
Public Function IsDeveloper(ByVal v_lEmpNum As Long) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    
    l_varRecord = GetRecord(Array(v_lEmpNum), DEVELOPER_CACHE, lblnExists)
    If IsEmpty(l_varRecord) Then
        IsDeveloper = False
    Else
        IsDeveloper = True
    End If

    Exit Function
eh:
    TrsRaiseError MODULENAME, "IsDeveloper"
End Function

Public Function IsTrainingInstructor(ByVal v_lEmpNum As Long) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    
    l_varRecord = GetRecord(Array(v_lEmpNum), TRAININGINSTRUCTOR_CACHE, lblnExists)
    If IsEmpty(l_varRecord) Then
        IsTrainingInstructor = False
    Else
        IsTrainingInstructor = True
    End If

    Exit Function
eh:
    TrsRaiseError MODULENAME, "IsTrainingInstructor"
End Function

Public Function ValidateEnvironment(ByVal v_sEnvironment As String, ByVal v_lEmpNum As Long) As Boolean
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    
    l_varRecord = GetRecord(Array(v_sEnvironment, v_lEmpNum), ENVIRONMENT_CACHE, lblnExists)
    If IsEmpty(l_varRecord) Then
        ValidateEnvironment = False
    Else
        ValidateEnvironment = True
    End If

    Exit Function
eh:
    TrsRaiseError MODULENAME, "ValidateEnvironment"
End Function

Public Function RedirectURL(ByVal v_sEnvironment As String, ByVal v_sURL As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    
    l_varRecord = GetRecord(Array(v_sEnvironment), REDIRECTURL_CACHE, lblnExists)
    If IsEmpty(l_varRecord) Then
        RedirectURL = v_sURL
    Else
        RedirectURL = l_varRecord(REDIRECTURL_CACHE_URL)
    End If

    Exit Function
eh:
    TrsRaiseError MODULENAME, "RedirectURL"
End Function

Public Function RedirectURL2(ByVal v_lLoginAs As Long, _
                            ByVal v_sEnvironment As String, _
                            ByVal v_sURL As String) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim sUrl As String
    Dim sPriorityClass As String
    
On Error GoTo eh
    
    sUrl = RedirectURL(v_sEnvironment, v_sURL)
    
    'http://TRISPRODNLB/TRISServiceInterfaceP/ServiceInterface.asmx
    If v_sEnvironment = "PROD" Then
        sPriorityClass = GetPriorityClass(v_lLoginAs)
        If sPriorityClass <> "" Then
            sUrl = Replace(sUrl, "TRISServiceInterfaceP", "TRISServiceInterfaceP" & sPriorityClass)
        End If
    End If
    
    RedirectURL2 = sUrl

    Exit Function
eh:
    TrsRaiseError MODULENAME, "RedirectURL"
End Function

Public Function GetAlias(ByVal v_lEmpNum As Long) As Long
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    
    l_varRecord = GetRecord(Array(v_lEmpNum), ALIAS_CACHE, lblnExists)
    If IsEmpty(l_varRecord) Then
        GetAlias = v_lEmpNum
    Else
        GetAlias = l_varRecord(ALIAS_CACHE_ALIAS_EMP_NUM)
    End If

    Exit Function
eh:
    TrsRaiseError MODULENAME, "RedirectURL"
End Function

Private Function GetPriorityClass(ByVal v_lLoginAs As Long) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim sPriority As String

On Error GoTo eh
    
    '//-----defaults: could be E, H, or M
    sPriority = "PRIORITY3"
    GetPriorityClass = "M"
    
    l_varRecord = GetRecord(Array(v_lLoginAs), PRIORITY_CACHE, lblnExists)
    If Not IsEmpty(l_varRecord) Then
        sPriority = l_varRecord(PRIORITY_CACHE_PRIORITY)
    Else
        l_varRecord = GetRecord(Array(v_lLoginAs), PRIORITYDEFAULT_CACHE, lblnExists)
        If Not IsEmpty(l_varRecord) Then
            sPriority = l_varRecord(PRIORITYDEFAULT_CACHE_PRIORITY)
        End If
    End If

    '//-----Convert Priority1, 2, 3 to E, H, M
    l_varRecord = GetRecord(Array(sPriority), PRIORITYCLASS_CACHE, lblnExists)
    If Not IsEmpty(l_varRecord) Then
        GetPriorityClass = l_varRecord(PRIORITYCLASS_CACHE_CLASS)
    End If

    Exit Function
eh:
    TrsRaiseError MODULENAME, "GetPriorityClass"

End Function

Public Function GetSysParms() As ADODB.Recordset
    Dim sSQL As String
On Error GoTo eh
    
    sSQL = sSQL & "SELECT PARM_NUM, PARM_SUFFIX, RTRIM(PARM_DESC) AS PARM_DESC, RTRIM(PARM_VALUE) AS PARM_VALUE "
    sSQL = sSQL & "FROM METHDBA.TTR_SYSPARMS "
    sSQL = sSQL & "ORDER BY PARM_NUM, PARM_SUFFIX"
    
    Set GetSysParms = ExecuteSQL(sSQL, adOpenStatic, adLockBatchOptimistic)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetSysParms"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME, "GetSysParms", trErrdb_ExecuteSQL_F, , , DescArgsFrom(sSQL), True
            Case Else
                TrsRaiseError MODULENAME, "GetSysParms", trErrdb_TRISAvailable_F, , , , True
    End Select
End Function

Public Function GetSysParmsValue(ByVal v_lngParmNum As Long, ByVal v_lngParmSuffix As Long) As String
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
            
On Error GoTo eh
    
    l_varRecord = GetRecord(Array(v_lngParmNum, v_lngParmSuffix), SYSPARMS_CACHE, lblnExists)
    If Not IsEmpty(l_varRecord) Then
        GetSysParmsValue = l_varRecord(SYSPARMS_CACHE_PARMVALUE)
    End If

    Exit Function
eh:
    TrsRaiseError MODULENAME, "GetSysParmsValue"
End Function

'//-----Determine "Login As" (Alias) by using Rule 7
Private Function GetAliasRecords(Optional ByVal v_varKey As Variant) As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String

    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT " & _
            "CASE " & _
                "WHEN SUBSTR(RG_VALUE_CD,1,9) between '000000000' and '999999999'" & _
                "  THEN INTEGER(SUBSTR(RG_VALUE_CD,1,9)) " & _
                "ELSE 0 " & _
                "END AS EMP_NUM, " & _
            "CASE " & _
                "WHEN SUBSTR(RG_VALUE_CD,11,19) between '000000000' and '999999999'" & _
                "  THEN INTEGER(SUBSTR(RG_VALUE_CD,11,19)) " & _
                "ELSE 0 " & _
                "END AS ALIAS_EMP_NUM " & _
        " FROM METHDBA.TTR_RG_GROUP " & _
        " WHERE RG_CD = 7 " & _
        "   AND BUS_TYPE_CD = 2 "

    Set GetAliasRecords = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetAlias"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetAlias" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetAlias" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

'//-----Determine Developer by using Rule 102
Private Function GetDeveloperRecords(Optional ByVal v_varKey As Variant) As ADODB.Recordset
    Dim rsUsers As ADODB.Recordset
    Dim rsGroups As ADODB.Recordset
    
    On Error GoTo eh
        
    Set rsUsers = GetDevelopersByUser
    Set rsGroups = GetDevelopersByGroup

    If RecordsExist(rsGroups) Then
        rsGroups.MoveFirst
        Do While Not rsGroups.EOF
            GetUsersInGroup rsGroups.Fields("GROUP_NAME").Value, "", rsUsers
            
            rsGroups.MoveNext
        Loop
    End If
    
    Set GetDeveloperRecords = rsUsers
    
    Exit Function
eh:
    TrsRaiseError MODULENAME, "GetDeveloperRecords"
End Function

'//-----Determine Training Instructors or those that can see the "Login As" by using Rule 103
Private Function GetTrainingInstructorRecords(Optional ByVal v_varKey As Variant) As ADODB.Recordset
    Dim rsUsers As ADODB.Recordset
    Dim rsGroups As ADODB.Recordset
    
    On Error GoTo eh
    
    Set rsUsers = GetTrainingInstructorsByUser
    Set rsGroups = GetTrainingInstructorsByGroup

    If RecordsExist(rsGroups) Then
        rsGroups.MoveFirst
        Do While Not rsGroups.EOF
            GetUsersInGroup rsGroups.Fields("GROUP_NAME").Value, "", rsUsers
            
            rsGroups.MoveNext
        Loop
    End If
    
    Set GetTrainingInstructorRecords = rsUsers
    
    Exit Function
eh:
    TrsRaiseError MODULENAME, "GetDeveloperRecords"
End Function

'//-----Determine valid users for each environment by using Rule 104
Private Function GetEnvironmentUserRecords(Optional ByVal v_varKey As Variant) As ADODB.Recordset
    Dim rsUsers As ADODB.Recordset
    Dim rsGroups As ADODB.Recordset
    
    On Error GoTo eh
    
    Set rsUsers = GetEnvironmentUsersByUser
    Set rsGroups = GetEnvironmentUsersByGroup

    If RecordsExist(rsGroups) Then
        rsGroups.MoveFirst
        Do While Not rsGroups.EOF
            GetUsersInGroup rsGroups.Fields("GROUP_NAME").Value, rsGroups.Fields("ENVIRONMENT").Value, rsUsers
            
            rsGroups.MoveNext
        Loop
    End If
    
    Set GetEnvironmentUserRecords = rsUsers
    
    Exit Function
eh:
    TrsRaiseError MODULENAME, "GetEnvironmentUserRecords"
End Function

'//-----Determine URL redirect for each environment by using Rule 105
Private Function GetEnvironmentURLRecords(Optional ByVal v_varKey As Variant) As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String

    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT " & _
                "CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) " & _
                    "  ELSE '' END AS ENVIRONMENT " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN RTRIM(SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD))) " & _
                "  ELSE '' END AS URL " & _
        " FROM METHDBA.TTR_RG_GROUP " & _
        " WHERE RG_CD = 105 " & _
        "   AND BUS_TYPE_CD = 21 "
    
    Set GetEnvironmentURLRecords = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetEnvironmentURLRecords"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetEnvironmentURLRecords" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetEnvironmentURLRecords" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

Public Sub InvalidatePriorityCache(ByVal v_lngEmpNum As Long)
    On Error GoTo eh

    InvalidateCache Array(v_lngEmpNum), PRIORITY_CACHE
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InvalidatePriorityCache"
End Sub

Public Sub InvalidatePriorityDefaultCache(ByVal v_lngEmpNum As Long)
    On Error GoTo eh

    InvalidateCache Array(v_lngEmpNum), PRIORITYDEFAULT_CACHE
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InvalidatePriorityDefaultCache"
End Sub

Public Sub InvalidatePriorityClassCache(ByVal v_strPriority As String)
    On Error GoTo eh

    InvalidateCache v_strPriority, PRIORITYCLASS_CACHE
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InvalidatePriorityCache"
End Sub

Public Sub InvalidateSysParmsCache(ByVal v_intParmNum As Integer, ByVal v_intParmSuffix As Integer)
    On Error GoTo eh

    InvalidateCache Array(v_intParmNum, v_intParmSuffix), SYSPARMS_CACHE
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateSysParmsCache"
End Sub

Public Function GetResourceLinks() As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String
    Dim rsTemp As ADODB.Recordset
    Dim nStartPos As Long, sLink As String, nLen As Long
        
    'Set SQL SELECT statment
    l_sSQLStatement = _
        " SELECT" & _
        " CASE WHEN LOCATE(':',RG_VALUE_CD) > 0" & _
        " THEN   SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)" & _
        " Else '' END AS RESOURCE" & _
        " ,CASE WHEN LOCATE(':',RG_VALUE_CD) > 0" & _
        " THEN   SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1)" & _
        " Else '' END AS LINK" & _
        " , SEQ_NUM" & _
        " From METHDBA.TTR_RG_GROUP" & _
        " Where RG_CD = 140" & _
        " ORDER BY RESOURCE, SEQ_NUM"
    
    Set rsTemp = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
    
    '//----combine the link column if needed seq_num > 1
    Set GetResourceLinks = New ADODB.Recordset
    GetResourceLinks.Fields.Append "RESOURCE", adVarChar, 50
    GetResourceLinks.Fields.Append "LINK", adVarChar, 1024
    GetResourceLinks.Open
    
    If RecordsExist(rsTemp) Then
        rsTemp.MoveFirst
        
        Do While Not rsTemp.EOF
            If rsTemp.Fields("SEQ_NUM").Value >= 1 Then
                '//-----SEQ_NUM defines a block of 200 chars
                nStartPos = (CLng(rsTemp.Fields("SEQ_NUM").Value) - 1) * 200 + 1
                GetResourceLinks.Filter = "RESOURCE = '" & rsTemp.Fields("RESOURCE").Value & "'"
                If Not GetResourceLinks.EOF Then
                    nLen = Len(GetResourceLinks.Fields("LINK").Value)
                    If nStartPos > nLen Then '//-----add to the right
                       sLink = GetResourceLinks.Fields("LINK").Value & Space(nStartPos - nLen - 1) & PadR(rsTemp.Fields("LINK").Value, 200)
                    Else '//-----add to the left
                        Mid$(sLink, nStartPos, 200) = PadR(rsTemp.Fields("LINK").Value, 200)
                    End If
                Else
                    GetResourceLinks.AddNew
                    GetResourceLinks.Fields("RESOURCE").Value = rsTemp.Fields("RESOURCE").Value
                    sLink = Space(nStartPos - 1) & PadR(rsTemp.Fields("LINK").Value, 200)
                End If
                
                GetResourceLinks.Fields("LINK").Value = sLink
                GetResourceLinks.Update
            Else
                GetResourceLinks.AddNew
                GetResourceLinks.Fields("RESOURCE").Value = rsTemp.Fields("RESOURCE").Value
                GetResourceLinks.Fields("LINK").Value = rsTemp.Fields("LINK").Value
                GetResourceLinks.Update
            End If

            rsTemp.MoveNext
        Loop
        
        GetResourceLinks.Filter = adFilterNone
        
        '//-----final trim-----//
        If RecordsExist(GetResourceLinks) Then
            GetResourceLinks.MoveFirst
            Do While Not GetResourceLinks.EOF
                GetResourceLinks.Fields("LINK").Value = Trim(GetResourceLinks.Fields("LINK").Value)
                GetResourceLinks.Update
                
                GetResourceLinks.MoveNext
            Loop
        End If
    End If
        
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetResourceLinks"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetResourceLinks" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetResourceLinks" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

Private Function PadR(ByVal sSource As String, ByVal nPad As Long, Optional ByVal sPadChar As String = " ")
    Dim sPad

    If Len(sSource) >= nPad Then
        sPad = Left(sSource, nPad)
    Else
        sPad = sSource & String(nPad - Len(sSource), sPadChar)
    End If
    
    PadR = sPad
    Exit Function
End Function

'//-----Determine Priority for URL Redirection, use differenent db2 userid for COM+ Partition-----//
Private Function GetPriorityRecords(Optional ByVal v_varKey As Variant) As ADODB.Recordset
    On Error GoTo eh
    
    Dim sSQL As String, sWinUser As String, l_strWhere1 As String, l_strWhere2 As String

    sWinUser = "INTEGER(CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sWinUser = sWinUser & " THEN RTRIM(SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+2,LENGTH(RG_VALUE_CD))) "
    sWinUser = sWinUser & "Else '0' END)"
    
    'Set SQL SELECT statment
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "NE1.PRIORITY "
    sSQL = sSQL & ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sSQL = sSQL & "  THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) "
    sSQL = sSQL & "  Else '' END AS WINGROUP "
    sSQL = sSQL & ",INTEGER(CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sSQL = sSQL & "  THEN RTRIM(SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+2,LENGTH(RG_VALUE_CD))) "
    sSQL = sSQL & "  Else '0' END) AS EMP_NUM "
    sSQL = sSQL & "FROM METHDBA.TTR_RG_GROUP, "
    sSQL = sSQL & "   (SELECT "
    sSQL = sSQL & "   CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sSQL = sSQL & "     THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) "
    sSQL = sSQL & "     Else '' END AS PRIORITY "
    sSQL = sSQL & "   ,CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sSQL = sSQL & "     THEN RTRIM(SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD))) "
    sSQL = sSQL & "     Else '' END AS WINGROUP "
    sSQL = sSQL & "   From METHDBA.TTR_RG_GROUP "
    sSQL = sSQL & "   Where RG_CD = 109 "
    sSQL = sSQL & "    AND BUS_TYPE_CD = 20 "
    sSQL = sSQL & "    AND CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sSQL = sSQL & "             THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) "
    sSQL = sSQL & "             Else '' END LIKE 'PRIORITY_') NE1 "
    sSQL = sSQL & "Where RG_CD = 109 "
    sSQL = sSQL & " AND BUS_TYPE_CD = 19 "
    sSQL = sSQL & " AND CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sSQL = sSQL & "          THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) "
    sSQL = sSQL & "          Else '' END = NE1.WINGROUP "
    
    If IsMissing(v_varKey) Or IsEmpty(v_varKey) Then
        l_strWhere1 = ""
    ElseIf IsArray(v_varKey) Then
        l_strWhere1 = " AND " & BuildInClause(sWinUser, v_varKey, False)
    Else
        l_strWhere1 = " AND " & sWinUser & " = " & v_varKey & " "
    End If
    sSQL = sSQL & l_strWhere1
    
    sSQL = sSQL & "Union "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sSQL = sSQL & "  THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) "
    sSQL = sSQL & "  Else '' END AS PRIORITY "
    sSQL = sSQL & ", '' AS WINGROUP "
    sSQL = sSQL & ",INTEGER(CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sSQL = sSQL & "  THEN RTRIM(SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+2,LENGTH(RG_VALUE_CD))) "
    sSQL = sSQL & "  Else '0' END) AS EMP_NUM "
    sSQL = sSQL & "From METHDBA.TTR_RG_GROUP "
    sSQL = sSQL & "Where RG_CD = 109 "
    sSQL = sSQL & " AND BUS_TYPE_CD = 19 "
    sSQL = sSQL & " AND CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sSQL = sSQL & "          THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) "
    sSQL = sSQL & "          Else '' END LIKE 'PRIORITY_' "

    If IsMissing(v_varKey) Or IsEmpty(v_varKey) Then
        l_strWhere2 = ""
    ElseIf IsArray(v_varKey) Then
        l_strWhere2 = " AND " & BuildInClause(sWinUser, v_varKey, False)
    Else
        l_strWhere2 = sWinUser & " = " & v_varKey & " "
    End If
    sSQL = sSQL & l_strWhere2

    Set GetPriorityRecords = ExecuteSQL(sSQL, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetPriorityRecords"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetPriorityRecords" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(sSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetPriorityRecords" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function
          
Private Function GetPriorityDefaultRecords(Optional ByVal v_varKey As Variant) As ADODB.Recordset
    On Error GoTo eh
    
    Dim sSQL As String, l_strWhere As String

    sSQL = sSQL & "select a.emp_num, "
    sSQL = sSQL & "case when Min(appr_level_ind) <=30 "
    sSQL = sSQL & "then 'PRIORITY1' "
    sSQL = sSQL & "Else 'PRIORITY2' "
    sSQL = sSQL & "end As priority "
    sSQL = sSQL & "from methdba.ttr_approval a, methdba.ttr_employee e "
    sSQL = sSQL & "where e.emp_status = 'A' and e.sec_level <> 0 "
    sSQL = sSQL & "and a.emp_num = e.emp_num "
    sSQL = sSQL & "and a.APPR_LEVEL_IND <= 50 "
    sSQL = sSQL & "group by a.emp_num "
    
    If IsMissing(v_varKey) Or IsEmpty(v_varKey) Then
        l_strWhere = ""
    ElseIf IsArray(v_varKey) Then
        l_strWhere = "Having " & BuildInClause("a.emp_num", v_varKey, False)
    Else
        l_strWhere = "Having a.emp_num = " & v_varKey & " "
    End If
    sSQL = sSQL & l_strWhere
    
    sSQL = sSQL & "order by a.emp_num, priority desc "

    Set GetPriorityDefaultRecords = ExecuteSQL(sSQL, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetPriorityDefaultRecords"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetPriorityDefaultRecords" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(sSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetPriorityDefaultRecords" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function
          
Private Function GetPriorityClassRecords(Optional ByVal v_varKey As Variant) As ADODB.Recordset
    On Error GoTo eh
    
    Dim sSQL As String, sPriority As String, l_strWhere As String

    sPriority = " (CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sPriority = sPriority & " THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) "
    sPriority = sPriority & "Else '' END)"
    
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sSQL = sSQL & "  THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) "
    sSQL = sSQL & "  Else '' END AS PRIORITY "
    sSQL = sSQL & ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0 "
    sSQL = sSQL & "  THEN RTRIM(SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD))) "
    sSQL = sSQL & "  Else '' END AS CLASS "
    sSQL = sSQL & "From METHDBA.TTR_RG_GROUP "
    sSQL = sSQL & "Where RG_CD = 109 "
    sSQL = sSQL & " AND BUS_TYPE_CD = 21 "

    If IsMissing(v_varKey) Or IsEmpty(v_varKey) Then
        l_strWhere = ""
    ElseIf IsArray(v_varKey) Then
        l_strWhere = " AND " & BuildInClause(sPriority, v_varKey, True)
    Else
        l_strWhere = " AND " & sPriority & " = '" & v_varKey & "' "
    End If
    sSQL = sSQL & l_strWhere

    Set GetPriorityClassRecords = ExecuteSQL(sSQL, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetPriorityClassRecords"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetPriorityClassRecords" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(sSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetPriorityClassRecords" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

Private Function GetSysParmsRecords(Optional ByVal v_varKeys As Variant) As ADODB.Recordset
    Dim sSQL As String, l_strWhereClause As String
    
On Error GoTo eh
    sSQL = sSQL & "SELECT PARM_NUM, PARM_SUFFIX, RTRIM(PARM_DESC) AS PARM_DESC, RTRIM(PARM_VALUE) AS PARM_VALUE "
    sSQL = sSQL & "FROM METHDBA.TTR_SYSPARMS "
    'sSQL = sSQL & "ORDER BY PARM_NUM, PARM_SUFFIX"
    
    If IsMissing(v_varKeys) Or IsEmpty(v_varKeys) Then
        l_strWhereClause = ""
    'TODO: What is in the array (Parm_Num and Parm_Suffix or multiple Parm_num's) ???
    ElseIf IsArray(v_varKeys) Then 'could use BuildWhereClauseForMultiKey
        'l_strWhere = " AND " & BuildInClause("PARM_NUM", v_varKey, True)
        l_strWhereClause = BuildWhereClauseForMultiKey(Array("WHERE", "AND"), _
                                        Array("PARM_NUM", "PARM_SUFFIX"), _
                                        v_varKeys, _
                                        Array(False, False))
    Else
        l_strWhereClause = " WHERE PARM_NUM = " & v_varKeys
    End If
    sSQL = sSQL & l_strWhereClause

    Set GetSysParmsRecords = ExecuteSQL(sSQL, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetSysParmsRecords"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME, "GetSysParmsRecords", trErrdb_ExecuteSQL_F, , , DescArgsFrom(sSQL), True
            Case Else
                TrsRaiseError MODULENAME, "GetSysParmsRecords", trErrdb_TRISAvailable_F, , , , True
    End Select
End Function

'copied from mdbRules.bas
Private Function BuildWhereClauseForMultiKey(ByVal v_strSQLKeyWord As Variant, _
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

'//-----copied from dbCodeTables and mdbRules.bas
Private Function BuildWhereClause(ByVal v_strSQLKeyWord As String, _
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

'copied from mdbRules.bas
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

'//-----Developers-----//

Private Function GetDevelopersByUser() As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String

    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT '' AS GROUP_NAME, " & _
            "CASE " & _
                "WHEN SUBSTR(RG_VALUE_CD,1,9) between '000000000' and '999999999'" & _
                "  THEN INTEGER(SUBSTR(RG_VALUE_CD,1,9)) " & _
                "ELSE 0 " & _
                "END AS EMP_NUM " & _
        " FROM METHDBA.TTR_RG_GROUP " & _
        " WHERE RG_CD = 102 " & _
        "   AND BUS_TYPE_CD = 17 "

    Set GetDevelopersByUser = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockBatchOptimistic)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetDevelopersByUser"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetDevelopersByUser" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetDevelopersByUser" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

Private Function GetDevelopersByGroup() As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String

    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT RG_VALUE_CD AS GROUP_NAME" & _
        " FROM METHDBA.TTR_RG_GROUP " & _
        " WHERE RG_CD = 102 " & _
        "   AND BUS_TYPE_CD = 18 "

    Set GetDevelopersByGroup = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetDevelopersByGroup"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetDevelopersByGroup" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetDevelopersByGroup" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

''//-----use wmi to find user membership in group and add to user recordset-----//
'Private Function GetUsersInGroup(ByVal v_sGroupName As String, _
'                                    ByVal v_sEnvironment As String, _
'                                    ByRef r_rsUsers As ADODB.Recordset) As ADODB.Recordset
'    Dim objWMIService As Object, colItems As Object, objItem As Object, objItem2
'    Dim strComputer As String
'    Dim l_colErrArgs As Collection
'
'On Error GoTo eh
'
'    strComputer = "."
'    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
'    Set colItems = objWMIService.ExecQuery _
'        ("SELECT * FROM Win32_GroupUser WHERE GroupComponent = 'Win32_Group.Domain=""AMEREN"",Name=""" & v_sGroupName & """'")
'
'    For Each objItem In colItems
'        Set objItem2 = GetObject("winmgmts:" & objItem.PartComponent)
'        If Not objItem2 Is Nothing Then
'            r_rsUsers.AddNew
'            If v_sEnvironment <> "" Then
'                r_rsUsers.Fields("ENVIRONMENT").Value = v_sEnvironment
'            End If
'            r_rsUsers.Fields("GROUP_NAME").Value = v_sGroupName
'            r_rsUsers.Fields("EMP_NUM").Value = ParseUserID(objItem2.Name)
'            r_rsUsers.Update
'        Else
'
'        End If
'    Next
'
'    Exit Function
'eh:
'    Set l_colErrArgs = New Collection
'    l_colErrArgs.Add "v_sGroupName = " & v_sGroupName
'    l_colErrArgs.Add "v_sEnvironment = " & v_sEnvironment
'    If RecordsExist(r_rsUsers) Then
'        r_rsUsers.MoveFirst
'        l_colErrArgs.Add "r_rsUsers = " & r_rsUsers.GetString
'    End If
'
'    TrsRaiseError m_strCMODULENAME, "GetUsersInGroup", , , , l_colErrArgs, True
'End Function

'//-----use RULE 109 to find user membership in group and add to user recordset-----//
Private Function GetUsersInGroup(ByVal v_sGroupName As String, _
                                    ByVal v_sEnvironment As String, _
                                    ByRef r_rsUsers As ADODB.Recordset) As ADODB.Recordset
    Dim rsWinGroup As ADODB.Recordset
    Dim l_colErrArgs As Collection
On Error GoTo eh

        Set rsWinGroup = GetWinGroupRecords(v_sGroupName)

        If RecordsExist(rsWinGroup) Then
            rsWinGroup.MoveFirst
            Do While Not rsWinGroup.EOF
                r_rsUsers.AddNew
                If v_sEnvironment <> "" Then
                    r_rsUsers.Fields("ENVIRONMENT").Value = v_sEnvironment
                End If
                r_rsUsers.Fields("GROUP_NAME").Value = v_sGroupName
                r_rsUsers.Fields("EMP_NUM").Value = ParseUserID(rsWinGroup.Fields("WINUSER"))
                r_rsUsers.Update
                
                rsWinGroup.MoveNext
            Loop
        End If
        Set rsWinGroup = Nothing
    Exit Function
eh:
    Set rsWinGroup = Nothing
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_sGroupName = " & v_sGroupName
    l_colErrArgs.Add "v_sEnvironment = " & v_sEnvironment
    If RecordsExist(r_rsUsers) Then
        r_rsUsers.MoveFirst
        l_colErrArgs.Add "r_rsUsers = " & r_rsUsers.GetString
    End If
    
    TrsRaiseError m_strCMODULENAME, "GetUsersInGroup", , , , l_colErrArgs, True
End Function

'from busFunSecurity.cbsFunSecurity.ParseCustomerID
Private Function ParseUserID(ByVal v_sUserId As String) As Long
    Dim l_lIndex As Long, l_sNumber As String
    
    If (Trim(v_sUserId) <> "") Then
        For l_lIndex = 1 To Len(v_sUserId)
            If (IsNumeric(Mid(v_sUserId, l_lIndex, 1))) Then
                'If numeric then build string
                l_sNumber = l_sNumber & Mid(v_sUserId _
                                        , l_lIndex _
                                        , 1)
            End If
        Next l_lIndex
        If IsNumeric(l_sNumber) Then
            ParseUserID = CLng(l_sNumber)
        Else
        
        End If
    Else
        ParseUserID = 0
    End If

End Function

'//-----Training-----//

Private Function GetTrainingInstructorsByUser() As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String

    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT '' AS GROUP_NAME, " & _
            "CASE " & _
                "WHEN SUBSTR(RG_VALUE_CD,1,9) between '000000000' and '999999999'" & _
                "  THEN INTEGER(SUBSTR(RG_VALUE_CD,1,9)) " & _
                "ELSE 0 " & _
                "END AS EMP_NUM " & _
        " FROM METHDBA.TTR_RG_GROUP " & _
        " WHERE RG_CD = 103 " & _
        "   AND BUS_TYPE_CD = 17 "

    Set GetTrainingInstructorsByUser = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockBatchOptimistic)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetTrainingInstructorsByUser"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetTrainingInstructorsByUser" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetTrainingInstructorsByUser" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

Private Function GetTrainingInstructorsByGroup() As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String

    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT RG_VALUE_CD AS GROUP_NAME" & _
        " FROM METHDBA.TTR_RG_GROUP " & _
        " WHERE RG_CD = 103 " & _
        "   AND BUS_TYPE_CD = 18 "

    Set GetTrainingInstructorsByGroup = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetTrainingInstructorsByGroup"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetTrainingInstructorsByGroup" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetTrainingInstructorsByGroup" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

'//-----Environment/Users-----//

Private Function GetEnvironmentUsersByUser() As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String

    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT '' AS GROUP_NAME, " & _
                "CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) " & _
                    "  ELSE '' END AS ENVIRONMENT " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN   " & _
                    "  CASE WHEN SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD)) between '000000000' and '999999999'" & _
                    "     THEN INTEGER(SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD))) " & _
                    "     ELSE 0 END" & _
                "  ELSE 0 END AS EMP_NUM " & _
        " FROM METHDBA.TTR_RG_GROUP " & _
        " WHERE RG_CD = 104 " & _
        "   AND BUS_TYPE_CD = 19 "
    
    Set GetEnvironmentUsersByUser = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockBatchOptimistic)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetEnvironmentUsersByUser"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetEnvironmentUsersByUser" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetEnvironmentUsersByUser" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

Private Function GetEnvironmentUsersByGroup() As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String

    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT " & _
                "CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) " & _
                    "  ELSE '' END AS ENVIRONMENT " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN RTRIM(SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD))) " & _
                "  ELSE '' END AS GROUP_NAME " & _
        " FROM METHDBA.TTR_RG_GROUP " & _
        " WHERE RG_CD = 104 " & _
        "   AND BUS_TYPE_CD = 20 "
    
    Set GetEnvironmentUsersByGroup = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetEnvironmentUsersByGroup"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetEnvironmentUsersByGroup" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetEnvironmentUsersByGroup" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function


'//-----Determine Windows Groups/Users by using Rule 109
Private Function GetWinGroupRecords(Optional ByVal v_varKey As Variant) As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String

    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT " & _
                "CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) " & _
                    "  ELSE '' END AS WINGROUP " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN RTRIM(SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD))) " & _
                "  ELSE '' END AS WINUSER " & _
        " FROM METHDBA.TTR_RG_GROUP " & _
        " WHERE RG_CD = 109 " & _
        " AND BUS_TYPE_CD = 17 "
        
    If IsArray(v_varKey) Then
    Else
        l_sSQLStatement = l_sSQLStatement & _
                "AND CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                "  THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) " & _
                "  ELSE '' END = '" & CStr(v_varKey) & "'"
    End If
    
    Set GetWinGroupRecords = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetWinGroupRecords"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetWinGroupRecords" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetWinGroupRecords" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
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

    Select Case v_strGroupName
        Case DEVELOPER_CACHE
            Set l_rsRecords = GetDeveloperRecords(v_varKeys)
            r_varKeyArray = Array("EMP_NUM")
            r_varColumnArray = Empty
        Case TRAININGINSTRUCTOR_CACHE
            Set l_rsRecords = GetTrainingInstructorRecords(v_varKeys)
            r_varKeyArray = Array("EMP_NUM")
            r_varColumnArray = Empty
        Case ENVIRONMENT_CACHE
            Set l_rsRecords = GetEnvironmentUserRecords(v_varKeys)
            r_varKeyArray = Array("ENVIRONMENT", "EMP_NUM")
            r_varColumnArray = Empty
        Case REDIRECTURL_CACHE
            Set l_rsRecords = GetEnvironmentURLRecords(v_varKeys)
            r_varKeyArray = Array("ENVIRONMENT")
            r_varColumnArray = Empty
        Case ALIAS_CACHE
            Set l_rsRecords = GetAliasRecords(v_varKeys)
            r_varKeyArray = Array("EMP_NUM")
            r_varColumnArray = Empty
        Case PRIORITY_CACHE
            Set l_rsRecords = GetPriorityRecords(v_varKeys)
            r_varKeyArray = Array("EMP_NUM")
            r_varColumnArray = Array("PRIORITY", "WINGROUP", "EMP_NUM")
        Case PRIORITYDEFAULT_CACHE
            Set l_rsRecords = GetPriorityDefaultRecords(v_varKeys)
            r_varKeyArray = Array("EMP_NUM")
            r_varColumnArray = Array("EMP_NUM", "PRIORITY")
        Case PRIORITYCLASS_CACHE
            Set l_rsRecords = GetPriorityClassRecords(v_varKeys)
            r_varKeyArray = Array("PRIORITY")
            r_varColumnArray = Array("PRIORITY", "CLASS")
        Case SYSPARMS_CACHE
            Set l_rsRecords = GetSysParmsRecords(v_varKeys)
            r_varKeyArray = Array("PARM_NUM", "PARM_SUFFIX")
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

'*********************************************************
' Begin Private Caching methods
'*********************************************************

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
'    Case CJB_GROUPNAME
'        IsLazyInitialized = True
'    Case Else
        IsLazyInitialized = False
'    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsLazyInitialized"
End Function


Private Function GroupSupportsLoadRecord(ByVal v_strGroupName As String) As Boolean
'All group caches that support Invalidate must use parameterized SQL in corresponding GetxxxRecords (GetGroupRecords)
On Error GoTo eh
    Select Case Trim(v_strGroupName)
    Case PRIORITY_CACHE
        GroupSupportsLoadRecord = True
    Case PRIORITYDEFAULT_CACHE
        GroupSupportsLoadRecord = True
    Case PRIORITYCLASS_CACHE
        GroupSupportsLoadRecord = True
    Case SYSPARMS_CACHE
        GroupSupportsLoadRecord = True
    Case Else
        GroupSupportsLoadRecord = False
    End Select
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GroupSupportsLoadRecord"
End Function


Private Function OneRowForEachKey(ByVal v_strGroupName As String) As Boolean
'//-----normally each key/row is cached into a 1-dimensional array that represents each column-----//
'//-----union components will have multiple rows per key and will be cached into a 2-dimensional array-----//
On Error GoTo eh
'    Select Case Trim(v_strGroupName)
'    Case CUC_GROUPNAME
'        OneRowForEachKey = False
'    Case Else
        OneRowForEachKey = True
'    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OneRowForEachKey"
End Function

