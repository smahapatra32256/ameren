VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULENAME As String = "busEnvironment.cbsEnvironment"
Private Const m_strCMODULENAME As String = MODULENAME

Public Function PreLogin(ByVal v_sUserId As String, _
                            ByVal v_sEnvironment As String, _
                            ByVal v_sUrl As String, _
                            ByVal v_rsClientInfo As ADODB.Recordset, _
                            ByRef r_bAllowLoginAs As Boolean, _
                            ByRef r_bAllowUrlRedirect As Boolean, _
                            ByRef r_lLoginAs As Long, _
                            ByRef r_sUrlRedirect As String) As Boolean

    Dim l_obj As dbEnvironment.cdEnvironment
    Dim l_lUserId As Long
    Dim l_colErrArgs As Collection
    
    On Error GoTo eh
    PreLogin = False

    l_lUserId = ParseUserID(v_sUserId)
    
    Set l_obj = CreateObject("dbEnvironment.cdEnvironment")
    
    '//-----Determine Role
    If l_obj.IsDeveloper(l_lUserId) Or _
        (l_obj.IsTrainingInstructor(l_lUserId) And Left(v_sEnvironment, 4) <> "PROD") Then
        r_bAllowLoginAs = True
    Else
        r_bAllowLoginAs = False
    End If
    
    If l_obj.IsDeveloper(l_lUserId) Then
        r_bAllowUrlRedirect = True
    Else
        r_bAllowUrlRedirect = False
    End If
    
    '//-----Determine Alias-----//
    r_lLoginAs = l_obj.GetAlias(l_lUserId)
    
    '//-----Redirect URL if needed-----//
    r_sUrlRedirect = l_obj.RedirectURL(v_sEnvironment, v_sUrl)
    
    '//-----this method will utilize the priority cache groups in order to prioritize users
    '//-----it is also dependent on multiple COM+ partitions and connection strings
    '//r_sUrlRedirect = l_obj.RedirectURL2(r_lLoginAs, v_sEnvironment, v_sUrl)
    
    '//-----Validate Environment
    PreLogin = True '//-----groups not working, to do later-----//
    '//PreLogin = l_obj.ValidateEnvironment(v_sEnvironment, l_lUserId)
    
    Set l_obj = Nothing
    Exit Function
eh:
    Set l_obj = Nothing

    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_sUserId = " & v_sUserId
    l_colErrArgs.Add "v_sEnvironment = " & v_sEnvironment
    l_colErrArgs.Add "v_sUrl = " & v_sUrl
    
    If RecordsExist(v_rsClientInfo) Then
        v_rsClientInfo.MoveFirst
        l_colErrArgs.Add "v_rsClientInfo = " & v_rsClientInfo.GetString
    End If
    
    TrsRaiseError m_strCMODULENAME, "PreLogin", , , , l_colErrArgs, True
End Function

Public Sub InvalidatePriorityCache(ByVal v_lngEmpNum As Long)
    On Error GoTo eh

    Dim l_obj As dbEnvironment.cdEnvironment

    Set l_obj = CreateObject("dbEnvironment.cdEnvironment")
    
    l_obj.InvalidatePriorityCache v_lngEmpNum
    
    Set l_obj = Nothing
    Exit Sub
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "InvalidatePriorityCache"
End Sub

Public Sub InvalidatePriorityDefaultCache(ByVal v_lngEmpNum As Long)
    On Error GoTo eh

    Dim l_obj As dbEnvironment.cdEnvironment

    Set l_obj = CreateObject("dbEnvironment.cdEnvironment")
    
    l_obj.InvalidatePriorityDefaultCache v_lngEmpNum
    
    Set l_obj = Nothing
    Exit Sub
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "InvalidatePriorityDefaultCache"
End Sub

Public Sub InvalidatePriorityClassCache(ByVal v_strPriority As String)
    On Error GoTo eh

    Dim l_obj As dbEnvironment.cdEnvironment

    Set l_obj = CreateObject("dbEnvironment.cdEnvironment")
    
    l_obj.InvalidatePriorityClassCache v_strPriority
    
    Set l_obj = Nothing
    Exit Sub
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "InvalidatePriorityClassCache"
End Sub

Public Function GetResourceLinks() As ADODB.Recordset
    On Error GoTo eh

    Dim l_obj As dbEnvironment.cdEnvironment

    Set l_obj = CreateObject("dbEnvironment.cdEnvironment")
    
    Set GetResourceLinks = l_obj.GetResourceLinks
    
    Set l_obj = Nothing
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetResourceLinks"
End Function

Public Function GetSysParms() As ADODB.Recordset
    On Error GoTo eh

    Dim l_obj As dbEnvironment.cdEnvironment

    Set l_obj = CreateObject("dbEnvironment.cdEnvironment")
    
    Set GetSysParms = l_obj.GetSysParms
    
    Set l_obj = Nothing
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetSysParms"
End Function

'sysparms: parm_num = 44, parm_suffix = 1, parm_desc = Offset days for MaxChargeDate , parm_value = 1 (tomorrow)
Public Function GetMaxChargeDateOffset() As Long
    Dim db As dbEnvironment.cdEnvironment
    Dim strNumberOfDays As String
'    Dim NumberOfDays As Long
    
On Error GoTo eh
    Set db = CreateObject("dbEnvironment.cdEnvironment")
    
    strNumberOfDays = db.GetSysParmsValue(44, 1)
    
    If IsNumeric(strNumberOfDays) Then
        GetMaxChargeDateOffset = CLng(strNumberOfDays)
    Else
        GetMaxChargeDateOffset = 1 'tomorrow
    End If
        
    Set db = Nothing
    Exit Function
eh:
    Set db = Nothing
    TrsRaiseError m_strCMODULENAME, "GetMaxChargeDateOffset"
End Function

Public Function GetMaxChargeDate() As Date
    Dim NumberOfDays As Long
    
On Error GoTo eh
       
    NumberOfDays = GetMaxChargeDateOffset()
    GetMaxChargeDate = DateAdd("d", NumberOfDays, Date)
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetMaxChargeDate"
End Function

'sysparms: parm_num = 44, parm_suffix = 2, parm_desc = Offset days for EffectiveDate, parm_value = 0 (today)
Public Function GetEffectiveDateOffset() As Long
    Dim db As dbEnvironment.cdEnvironment
    Dim strNumberOfDays As String
    
On Error GoTo eh
    Set db = CreateObject("dbEnvironment.cdEnvironment")
    
    strNumberOfDays = db.GetSysParmsValue(44, 2)
    
    If IsNumeric(strNumberOfDays) Then
        GetEffectiveDateOffset = CInt(strNumberOfDays)
    Else
        GetEffectiveDateOffset = 0 'today
    End If
       
    Set db = Nothing
    Exit Function
eh:
    Set db = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEffectiveDateOffset"
End Function

'sysparms: parm_num = 44, parm_suffix = 2, parm_desc = Offset days for EffectiveDate, parm_value = 0 (today)
Public Function GetDefaultEffectiveDate() As Date
    Dim NumberOfDays As Long
    
On Error GoTo eh
   
    NumberOfDays = GetEffectiveDateOffset 'today
    GetDefaultEffectiveDate = DateAdd("d", NumberOfDays, Date)
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDefaultEffectiveDate"
End Function

'Private Function GetAlias( _
'    ByVal v_lngEmpNum As Long _
'    ) As Long
'
'    Dim l_obj As dbEnvironment.cdEnvironment
'    Dim l_rs As ADODB.Recordset
'
'    On Error GoTo eh
'
'
'
'    'Construct the Data Access Object
'    Set l_obj = CreateObject("dbEnvironment.cdEnvironment")
'    'NOTE: It is assumed that a valid employee number will be passed
'    'Call Data Access Object method and assign return
'    Set l_rs = l_obj.GetAlias(v_lngEmpNum)
'    'destroy
'    Set l_obj = Nothing
'
'    If l_rs Is Nothing Then
'        GetAlias = v_lngEmpNum
'        Exit Function
'    End If
'
'    If l_rs.BOF And l_rs.EOF Then
'        GetAlias = v_lngEmpNum
'        Exit Function
'    End If
'
'    'Currently the customer can only have 1 alias
'    'Also the customer can not sign on as themselves, if they have an alias
'
'    l_rs.MoveFirst
'    GetAlias = l_rs!ALIAS_EMP_NUM
'
'    Set l_rs = Nothing
'
'    Exit Function
'
'eh:
'    Set l_obj = Nothing
'    Set l_rs = Nothing
'
'    'Raise Error depending on the Error Number
'    TrsRaiseError MODULENAME _
'            , "GetAlias"
'
'End Function

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

