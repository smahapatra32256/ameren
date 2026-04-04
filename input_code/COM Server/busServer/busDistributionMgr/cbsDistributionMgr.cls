VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cbsDistributionMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsDistributionMgr"
Private Const m_strINVALIDFORCORP As String = "INVLD FOR CORP"
'do not use this public enum - It will be deleted - begin
Public Enum TrsDistributionMgrErrorEnum

    trErrdb_E_Invalid_Zero_v_lngWgNum_F = 30081 + vbObjectError

End Enum
'do not use this public enum - It will be deleted - end
Private l_varWorkGroupDepartmentCache(20, 1) As String
Private l_varDepartmentCache(20, 8) As String
Private l_varWorkGroupDOJMCache(20, 2) As String
Private l_varExceptionCodeCache(200, 8) As String
Private l_lngWorkGroupDepartmentCacheCount As Long
Private l_lngDepartmentCacheCount As Long
Private l_lngWorkGroupDOJMCacheCount As Long
Private l_lngExceptionCodeCacheCount As Long
Private m_rsTravelAllowances As ADODB.Recordset
Private m_datMKSEffDate As Date
Private m_rsMKSEmployeeOverrides As ADODB.Recordset




'For a DOJM Work request the Code block is
'       Positions   Array Element           Description
'       1-4         v_strDistributionID (0) DOJM District
'       5                                   Blank space
'       6-7         v_strDistributionID(1)  "WR" indicates that it is a work request
'       8-13        v_strDistributionID(2)  DOJM Work request
'The DOJM Work request is validated against the
'   DOJM WG, DOJM District, and DOJM Headquarters
'For GL Accounting the Code block is
'       Positions   Array Element           Description
'           1-1     v_strDistributionID(0)  (U)Utility
'           2-3     v_strDistributionID(1)  (BD) Business Division
'           4-6     v_strDistributionID(2)  (MAJ)FERC Major
'           7-9     v_strDistributionID(3)  (MIN)FERC Minor
'           10-12   v_strDistributionID(4)  (FMC) Fund Mgt. Center
'           13-15   v_strDistributionID(5)  (RMC) Resource Mgt. Center
'           16-16   v_strDistributionID(6)  (T) Trans Type
'           17-21   v_strDistributionID(7)  (PROJT) Project
'           22-23   v_strDistributionID(8)  (PR) Product
'           24-27   v_strDistributionID(9)  (ACTV) Activity
'           28-29   v_strDistributionID(10) (RT) Resource Type
'           30-39   v_strDistributionID(11) (FEEDER_REF) Memo

'For Service Request Accounting the Code block is the
' same as above, but only the following fields can be entered
'    FMC, Project, and Activity


Public Function GetDefaultAcct( _
        ByVal v_strExceptionCd As String, _
        ByVal v_strDistDept As String, _
        Optional ByVal v_strBank As String = "", _
        Optional ByVal v_strOverrideExcCd As String = "", _
        Optional ByVal v_strAcctDiv As String = "" _
        ) As String
    
'Note v_strDistDept is actually the payroll department

    Dim l_strInterfaceToCode As String
    Dim l_strNonProdLoad As String
    Dim l_strAcctType As String
    Dim l_strDeptAcctNum As String
    Dim l_strExceptionAcctNum As String
    Dim l_varDistID As Variant
    Dim l_varDefaultDistID As Variant
    Dim l_strPayrollDept As String
    Dim l_strCompany As String
    
On Error GoTo GetDefaultAcct_EH
    
    l_strPayrollDept = v_strDistDept
    l_strCompany = GetCompanyCode(l_strPayrollDept)

    GetDefaultAcct = ""
    
'  If Empty set Account Division via Dept
    If Not (IsPOETActive And IsOTTCEffective) Then
        If Trim(v_strAcctDiv) = "" Then
            v_strAcctDiv = GetAcctDiv(l_strPayrollDept)
        End If
    End If
        
'  Process Request
    l_strInterfaceToCode = GetInterfaceCode(l_strPayrollDept)
    l_strAcctType = GetAcctType(l_strPayrollDept)
    l_strNonProdLoad = GetNonProdLoadInd(l_strPayrollDept)
    l_strDeptAcctNum = GetAcctNum(l_strPayrollDept)
    'add GetInterfaceInd here
    
    If l_strInterfaceToCode = "D" Then
        If Not (AllowDefaultAccountingForDOJM(v_strExceptionCd) Or AllowDefaultAccountingForDOJM(v_strOverrideExcCd)) Then
            GetDefaultAcct = ""
            Exit Function
        End If
    End If
'  If fixed accounting department no default accounting
    If l_strAcctType = "F" Then
        GetDefaultAcct = ""
        Exit Function
    End If
    If l_strAcctType = "E" Then
        If l_strNonProdLoad = "N" Then
            GetDefaultAcct = l_strDeptAcctNum
            Exit Function
        End If
        If Trim(v_strOverrideExcCd) <> "" Then
            l_strExceptionAcctNum = GetExceptionAccountNum(v_strOverrideExcCd, l_strCompany)
            If Trim(l_strExceptionAcctNum) = "" Or Trim(l_strExceptionAcctNum) = "N/A" Then
                GetDefaultAcct = Trim(l_strExceptionAcctNum)
            Else
                l_varDistID = ParseStringDistIDToAnArray(l_strExceptionAcctNum)
                l_varDefaultDistID = FillInMissingPieces(l_varDistID, l_strPayrollDept, 0, "A", "", True)
                GetDefaultAcct = ParseArrayDistIDToAString(l_varDefaultDistID)
            End If
        'If the override exception code had default accounting  then exit
            If Trim(l_strExceptionAcctNum) <> "" Then
                Exit Function
            End If
        End If
        If (Trim(v_strExceptionCd = "ST") Or v_strExceptionCd = "SAS" Or v_strExceptionCd = "SUS" Or v_strExceptionCd = "SSC" Or v_strBank = "O") Then
            GetDefaultAcct = ""
            Exit Function
        End If
        l_strExceptionAcctNum = GetExceptionAccountNum(v_strExceptionCd, l_strCompany)
        If Trim(l_strExceptionAcctNum) = "" Or Trim(l_strExceptionAcctNum) = "N/A" Then
            GetDefaultAcct = Trim(l_strExceptionAcctNum)
        Else
            l_varDistID = ParseStringDistIDToAnArray(l_strExceptionAcctNum)
            l_varDefaultDistID = FillInMissingPieces(l_varDistID, l_strPayrollDept, 0, "A", "", True)
            GetDefaultAcct = ParseArrayDistIDToAString(l_varDefaultDistID)
        End If
        Exit Function
    End If
                  
    If Trim(v_strOverrideExcCd) <> "" Then
        l_strExceptionAcctNum = GetExceptionAccountNum(v_strOverrideExcCd, l_strCompany)
        If l_strNonProdLoad = "N" And Trim(l_strExceptionAcctNum) <> "" Then
            GetDefaultAcct = l_strDeptAcctNum
            Exit Function
        Else
            If Trim(l_strExceptionAcctNum) <> "" Then
                If Trim(l_strExceptionAcctNum) = "N/A" Then
                    GetDefaultAcct = "N/A"
                Else
                    l_varDistID = ParseStringDistIDToAnArray(l_strExceptionAcctNum)
                    l_varDefaultDistID = FillInMissingPieces(l_varDistID, l_strPayrollDept, 0, "A", "", True)
                    GetDefaultAcct = ParseArrayDistIDToAString(l_varDefaultDistID)
                End If
                Exit Function
            End If
        End If
    End If
    l_strExceptionAcctNum = GetExceptionAccountNum(v_strExceptionCd, l_strCompany)
    If l_strNonProdLoad = "N" And Trim(l_strExceptionAcctNum) <> "" Then
        GetDefaultAcct = l_strDeptAcctNum
        Exit Function
    Else
        If Trim(l_strExceptionAcctNum) = "" Then
            GetDefaultAcct = ""
        Else
            If Trim(l_strExceptionAcctNum) = "N/A" Then
                GetDefaultAcct = "N/A"
            Else
                l_varDistID = ParseStringDistIDToAnArray(l_strExceptionAcctNum)
                l_varDefaultDistID = FillInMissingPieces(l_varDistID, l_strPayrollDept, 0, "A", "", True)
                GetDefaultAcct = ParseArrayDistIDToAString(l_varDefaultDistID)
            End If
        End If
    End If
  
    Exit Function

GetDefaultAcct_EH:
    TrsRaiseError m_strCMODULENAME, "GetDefaultAcct", , , , , True


End Function

Private Function ValidateWorkRequest( _
            ByVal v_varDistributionId As Variant, _
            ByVal v_lngChargeWGNum As Long, _
            ByRef r_strReturnMsg As String _
            ) As Boolean
    '  Is the WorkRequest Valid
    Dim l_objValidate As dbDistributionMgr.cdDistributionMgr
    Dim l_blnSuccess As Boolean
    Dim l_strDOJMDistrict As String
    Dim l_strDOJMHeadquarters As String
        
On Error GoTo ValidateWorkRequest_EH

    l_strDOJMDistrict = v_varDistributionId(0)
    l_blnSuccess = ValidateDOJMDistrict(l_strDOJMDistrict, v_lngChargeWGNum, r_strReturnMsg)
    If Not l_blnSuccess Then
        ValidateWorkRequest = False
        Exit Function
    End If
    GetDOJMDistrictAndHeadQuarter v_lngChargeWGNum, l_strDOJMDistrict, l_strDOJMHeadquarters

    Set l_objValidate = CreateObject("dbDistributionMgr.cdDistributionMgr")
    ValidateWorkRequest = l_objValidate.ValidateWorkRequest(v_varDistributionId, l_strDOJMHeadquarters)
    Set l_objValidate = Nothing
    
Exit Function

ValidateWorkRequest_EH:

    Set l_objValidate = Nothing
    TrsRaiseError m_strCMODULENAME, "ValidateWorkRequest"

End Function

Public Function IsDOJMWorkRequest(ByVal v_varDistributionId As Variant) As Boolean
    '  If the first field is a Valid DOJM District then it is a Work Request
    Dim l_objDOJMDistrict As dbDistributionMgr.cdDistributionMgr

On Error GoTo IsDOJMWorkRequest_EH

    Set l_objDOJMDistrict = CreateObject("dbDistributionMgr.cdDistributionMgr")
    
    If l_objDOJMDistrict.IsDOJMDistrict(v_varDistributionId) Then
        IsDOJMWorkRequest = True
    Else
        IsDOJMWorkRequest = False
    End If
    
    Set l_objDOJMDistrict = Nothing
Exit Function

IsDOJMWorkRequest_EH:
    Set l_objDOJMDistrict = Nothing
    TrsRaiseError m_strCMODULENAME, "IsDOJMWorkRequest", , , , , True

End Function
Public Function IsEMPRVJobTaskWG(ByVal v_varWorkGroup As Variant) As Boolean
    '  If the first field is a Valid DOJM District then it is a Work Request
    Dim l_objEMPRVJobTaskWG As dbWorkGroup.cdWorkGroup

On Error GoTo IsEMPRVJobTaskWG_EH

    Set l_objEMPRVJobTaskWG = CreateObject("dbWorkGroup.cdWorkGroup")
    
    If l_objEMPRVJobTaskWG.IsJobTaskWorkgroup(v_varWorkGroup) Then
        IsEMPRVJobTaskWG = True
    Else
        IsEMPRVJobTaskWG = False
    End If
    
    Set l_objEMPRVJobTaskWG = Nothing
Exit Function

IsEMPRVJobTaskWG_EH:
    Set l_objEMPRVJobTaskWG = Nothing
    TrsRaiseError m_strCMODULENAME, "IsEMPRVJobTaskWG", , , , , True

End Function

Public Function IsServiceRequest(ByVal v_varDistributionId As Variant) As Boolean

' A Distribution ID is a service request
' if only the following 3 fields are entered
' FMC, Project, Activity
' The Feeder_ref (Memo) field is ignored
' The Project field must be entered to be a service request

    Dim l_strTempHold As Variant
    Dim l_intNumOfFields As Integer
    
On Error GoTo IsServiceRequest_EH
    
    Dim l_strDistributionId(11) As String
    Dim l_strEmptyFields As String
    
    If Not IsArray(v_varDistributionId) Then
        l_strDistributionId(0) = Left$(v_varDistributionId, 1) '(U)Utility
        l_strDistributionId(1) = Mid$(v_varDistributionId, 2, 2) '(BD) Business Division
        l_strDistributionId(2) = Mid$(v_varDistributionId, 4, 3) '(MAJ)FERC Major
        l_strDistributionId(3) = Mid$(v_varDistributionId, 7, 3) '(MIN)FERC Minor
        l_strDistributionId(4) = Mid$(v_varDistributionId, 10, 3) '(FMC) Fund Mgt. Center
        l_strDistributionId(5) = Mid$(v_varDistributionId, 13, 3) '(RMC) Resource Mgt. Center
        l_strDistributionId(6) = Mid$(v_varDistributionId, 16, 1) '(T) Trans Type
        l_strDistributionId(7) = Mid$(v_varDistributionId, 17, 5) '(PROJT) Project
        l_strDistributionId(8) = Mid$(v_varDistributionId, 22, 2) '(PR) Product
        l_strDistributionId(9) = Mid$(v_varDistributionId, 24, 4) '(ACTV) Activity
        l_strDistributionId(10) = Mid$(v_varDistributionId, 28, 2) '(RT) Resource Type
        l_strDistributionId(11) = Mid$(v_varDistributionId, 30, 10) '(FEEDER_REF) Memo
    Else
        l_strDistributionId(0) = v_varDistributionId(0) '(U)Utility
        l_strDistributionId(1) = v_varDistributionId(1) '(BD)  Business Division
        l_strDistributionId(2) = v_varDistributionId(2) '(MAJ)FERC Major
        l_strDistributionId(3) = v_varDistributionId(3) '(MIN)FERC Minor
        l_strDistributionId(4) = v_varDistributionId(4) '(FMC) Fund Mgt. Center
        l_strDistributionId(5) = v_varDistributionId(5) '(RMC) Resource Mgt. Center
        l_strDistributionId(6) = v_varDistributionId(6) '(T) Trans Type
        l_strDistributionId(7) = v_varDistributionId(7) '(PROJT) Project
        l_strDistributionId(8) = v_varDistributionId(8) '(PR) Product
        l_strDistributionId(9) = v_varDistributionId(9) '(ACTV) Activity
        l_strDistributionId(10) = v_varDistributionId(10) '(RT) Resource Type
        l_strDistributionId(11) = v_varDistributionId(11) '(FEEDER_REF) Memo
    End If
    
    l_strEmptyFields = l_strDistributionId(0) & _
                       l_strDistributionId(1) & _
                       l_strDistributionId(2) & _
                       l_strDistributionId(3) & _
                       l_strDistributionId(5) & _
                       l_strDistributionId(6) & _
                       l_strDistributionId(8) & _
                       l_strDistributionId(10)
    If Trim(l_strEmptyFields) <> "" Then
        IsServiceRequest = False
        Exit Function
    End If
    If Trim(l_strDistributionId(7)) <> "" Then
'Please not that FMC and Activity may be missing
        IsServiceRequest = True
    Else
        IsServiceRequest = False
    End If
    Exit Function
IsServiceRequest_EH:
    TrsRaiseError m_strCMODULENAME, "IsServiceRequest", , , , , True
    
End Function

'//for older TRIS clients, pick list description is 30 chars
Public Function GetAccountDescription(ByVal v_objRsPickListDistribution As ADODB.Recordset, _
                                      ByVal v_lngWgNum As Long, _
                                      ByVal v_strDistributionId As String, _
                                      ByVal v_strDistributionType As String, _
                                      ByVal v_strDistributionDeptNum As String) As String
On Error GoTo eh
                                      
    GetAccountDescription = GetAccountDescriptionPrivate(v_objRsPickListDistribution, v_lngWgNum, v_strDistributionId, v_strDistributionType, v_strDistributionDeptNum)
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetAccountDescription", , , , , True
End Function

'//for newer TRIS clients (OTTC) pick list description increased to 60 chars
'//or will be called by older TRIS Client, busTimecard passing in a length of 30
Public Function GetAccountDescription2(ByVal v_objRsPickListDistribution As ADODB.Recordset, _
                                      ByVal v_lngWgNum As Long, _
                                      ByVal v_strDistributionId As String, _
                                      ByVal v_strDistributionType As String, _
                                      ByVal v_strDistributionDeptNum As String, _
                                      ByVal v_intDescriptionLength) As String
On Error GoTo eh
                                      
    GetAccountDescription2 = GetAccountDescriptionPrivate(v_objRsPickListDistribution, v_lngWgNum, v_strDistributionId, v_strDistributionType, v_strDistributionDeptNum, v_intDescriptionLength)
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetAccountDescription2", , , , , True
End Function

Private Function GetAccountDescriptionPrivate(ByVal v_objRsPickListDistribution As ADODB.Recordset, _
                                      ByVal v_lngWgNum As Long, _
                                      ByVal v_strDistributionId As String, _
                                      ByVal v_strDistributionType As String, _
                                      ByVal v_strDistributionDeptNum As String, _
                                      Optional ByVal v_intDescriptionLength = 39) As String
'    Dim l_objIsWorkReasonIndicatorRequired As busDepartment.cbsDepartment
    Dim l_objGetJobTaskList As dbDistributionMgr.cdDistributionMgr
    'Dim l_objGetJobTaskList As Object
    Dim l_strDeptNum As String

    Dim l_strFilter As String
    Dim l_strAccountDescription As String
    Dim l_strDistributionId As String
    Dim l_blnMatch As Boolean
'    Dim l_blnIsWorkReasonIndicatorRequired As Boolean
    Dim l_rsModifiedList As ADODB.Recordset
    Dim l_rsEMPRVJobTasks As ADODB.Recordset
    
    On Error GoTo GetAccountDescription_EH
    
    If v_lngWgNum < 1 Then
        Err.Raise trErrbusInvalidZerolngWgNum
    End If
    
    If v_strDistributionType = "W" Then
        GetAccountDescriptionPrivate = Trim(v_strDistributionId)
        Exit Function
    End If
    
    'If v_strDistributionType = "J" Then
    '    Dim l_colwg As New Collection
    '    l_colwg.Add v_lngWgNum
    '    l_strDeptNum = GetDepartmentOfWorkGroup(v_lngWgNum)
    '    Set l_objGetJobTaskList = CreateObject("dbDistributionMgr.cdDistributionMgr")    'MTS standard
    '    Set l_rsEMPRVJobTasks = l_objGetJobTaskList.GetDistJobTaskList(l_colwg, l_strDeptNum)
    '    Set l_objGetJobTaskList = Nothing
    '    If RecordsExist(l_rsEMPRVJobTasks) Then
            'l_strFilter = "JOB_NBR <> '' AND DISTRIBUTION_ID = '" & v_strDistributionId & " '"
    '        l_strFilter = "JOB_NBR <> ''" & _
    '        " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
    '        " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
    '        " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
    '        " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
    '        " AND FMC_CD = '" & Mid(v_strDistributionId, 10, 3) & "'" & _
    '        " AND RMC_CD = '" & Mid(v_strDistributionId, 13, 3) & "'" & _
    '        " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
    '        " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
    '        " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'" & _
    '        " AND ACTIV_CD = '" & Mid(v_strDistributionId, 24, 4) & "'" & _
    '        " AND RES_TP_CD = '" & Mid(v_strDistributionId, 28, 2) & "'"

    '        l_rsEMPRVJobTasks.Filter = MDAC25FilterFix(l_strFilter)
    '        If Not (l_rsEMPRVJobTasks.EOF And l_rsEMPRVJobTasks.BOF) Then
    '            l_rsEMPRVJobTasks.MoveFirst
    '            GetAccountDescriptionPrivate = Trim(l_rsEMPRVJobTasks.Fields("DESCRIPTION").Value)
    '        Else
    '            GetAccountDescriptionPrivate = Trim(l_rsEMPRVJobTasks.Fields("JOB_NBR").Value) & _
    '                Trim(l_rsEMPRVJobTasks.Fields("JOBTASK_NBR").Value)
    '        End If
    '    End If
    '    Exit Function
    'End If

'Does it Match Exactly (may not match due to resource type)
    l_strDistributionId = v_strDistributionId
    l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, l_strDistributionId, v_objRsPickListDistribution, v_lngWgNum, v_intDescriptionLength)
    If l_blnMatch Then
        GetAccountDescriptionPrivate = l_strAccountDescription
        GoTo ExitFunction
    End If
    
    '//If Mid(l_strDistributionId, 16, 1) = "P" Then '//POET
    If IsPOET39(l_strDistributionId) Then '//POET
        Dim SearchPOET As String
                
        'Copying the pick list and removing resource type (LU, LM) (match on project, activity/task and feeder ref)
        Set l_rsModifiedList = ModifyDistrList(v_objRsPickListDistribution, v_lngWgNum, True, True, True, False, False)
    
        '// try to match only on TRIS Project Id, TRIS Task Id, and trans type (P)
        If RecordsExist(l_rsModifiedList) Then
            SearchPOET = String(39, " ")
            SetProject39 SearchPOET, GetProject39(l_strDistributionId)
            SetActivity39 SearchPOET, GetActivity39(l_strDistributionId)
            SetTranType39 SearchPOET, GetTranType39(l_strDistributionId)
            SetFeederRef39 SearchPOET, GetFeederRef39(l_strDistributionId)
           
            l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, SearchPOET, l_rsModifiedList, v_lngWgNum, v_intDescriptionLength)
            If l_blnMatch Then
                GetAccountDescriptionPrivate = l_strAccountDescription
                GoTo ExitFunction
            End If
        End If


        'Copying the pick list and removing FeederRef and resource type (LU, LM) (match on project and activity/task)
        Set l_rsModifiedList = ModifyDistrList(v_objRsPickListDistribution, v_lngWgNum, True, True, True, False, True)
    
        '// try to match only on TRIS Project Id, TRIS Task Id, and trans type (P)
        If RecordsExist(l_rsModifiedList) Then
            SearchPOET = String(39, " ")
            SetProject39 SearchPOET, GetProject39(l_strDistributionId)
            SetActivity39 SearchPOET, GetActivity39(l_strDistributionId)
            SetTranType39 SearchPOET, GetTranType39(l_strDistributionId)
           
            '//l_strDistributionId = GetAccountDescRemoveFields(v_strDistributionId, False, False, False, False, True)
           
            l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, SearchPOET, l_rsModifiedList, v_lngWgNum, v_intDescriptionLength)
            If l_blnMatch Then
                GetAccountDescriptionPrivate = l_strAccountDescription
                GoTo ExitFunction
            End If
        End If
    End If
    
'    Set l_objIsWorkReasonIndicatorRequired = CreateObject("busDepartment.cbsDepartment")    'MTS standard
'    l_blnIsWorkReasonIndicatorRequired = l_objIsWorkReasonIndicatorRequired.IsWorkReasonIndicatorRequired(v_strDistributionDeptNum)
'    Set l_objIsWorkReasonIndicatorRequired = Nothing

'    If l_blnIsWorkReasonIndicatorRequired Then
'        l_strDistributionId = GetAccountDescRemoveFields( _
'        v_strDistributionId, False, False, False, False, l_blnIsWorkReasonIndicatorRequired)
'        l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, l_strDistributionId, v_objRsPickListDistribution, v_lngWgNum)
'        If l_blnMatch Then
'            GetAccountDescriptionPrivate = l_strAccountDescription
'            GoTo ExitFunction
'        End If
'    End If
    
 ' Remove RMC & RT
    l_strDistributionId = GetAccountDescRemoveFields(v_strDistributionId, True, True, False, False)
    l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, l_strDistributionId, v_objRsPickListDistribution, v_lngWgNum, v_intDescriptionLength)
    If l_blnMatch Then
        GetAccountDescriptionPrivate = l_strAccountDescription
        GoTo ExitFunction
    End If

 ' Remove RMC & RT & FMC
    l_strDistributionId = GetAccountDescRemoveFields(v_strDistributionId, True, True, True, False)
    l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, l_strDistributionId, v_objRsPickListDistribution, v_lngWgNum, v_intDescriptionLength)
    If l_blnMatch Then
        GetAccountDescriptionPrivate = l_strAccountDescription
        GoTo ExitFunction
    End If
    
    
' Remove RMC & RT & ACTV
    l_strDistributionId = GetAccountDescRemoveFields(v_strDistributionId, True, True, False, True)
    l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, l_strDistributionId, v_objRsPickListDistribution, v_lngWgNum, v_intDescriptionLength)
    If l_blnMatch Then
        GetAccountDescriptionPrivate = l_strAccountDescription
        GoTo ExitFunction
    End If
    
    
' Remove RMC & RT & FMC & ACTV
    l_strDistributionId = GetAccountDescRemoveFields(v_strDistributionId, True, True, True, True)
    l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, l_strDistributionId, v_objRsPickListDistribution, v_lngWgNum, v_intDescriptionLength)
    If l_blnMatch Then
        GetAccountDescriptionPrivate = l_strAccountDescription
        GoTo ExitFunction
    End If
    
' Remove FMC & ACTV
    l_strDistributionId = GetAccountDescRemoveFields(v_strDistributionId, False, False, True, True)
    l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, l_strDistributionId, v_objRsPickListDistribution, v_lngWgNum, v_intDescriptionLength)
    If l_blnMatch Then
        GetAccountDescriptionPrivate = l_strAccountDescription
        GoTo ExitFunction
    End If
    
' Remove ACTV
    l_strDistributionId = GetAccountDescRemoveFields(v_strDistributionId, False, False, False, True)
    l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, l_strDistributionId, v_objRsPickListDistribution, v_lngWgNum, v_intDescriptionLength)
    If l_blnMatch Then
        GetAccountDescriptionPrivate = l_strAccountDescription
        GoTo ExitFunction
    End If

'//for wg 15182, there was a match found for the following, so try to match on FMC first
'//         FMCRMC
'//111920   061061 A028101LMPMLM          MANAG/SUPV  50% UEC TRAIN DRIV
'//111920   EDY061 A028101LMPMLM          PROP MGMT - UEC - ENERGY DELIV
 
 'Copying the pick list and removing RMC, and RT
    Set l_rsModifiedList = ModifyDistrList(v_objRsPickListDistribution, v_lngWgNum, True, True, False, False)
    
 ' Remove RMC & RT
    If RecordsExist(l_rsModifiedList) Then
        l_strDistributionId = GetAccountDescRemoveFields(v_strDistributionId, True, True, False, False)
        l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, l_strDistributionId, l_rsModifiedList, v_lngWgNum, v_intDescriptionLength)
        If l_blnMatch Then
            GetAccountDescriptionPrivate = l_strAccountDescription
            GoTo ExitFunction
        End If
    End If
    
 'Copying the pick list and removing RMC, FMC, and RT
    Set l_rsModifiedList = ModifyDistrList(v_objRsPickListDistribution, v_lngWgNum, True, True, True, False)
 
 ' Remove RMC, FMC & RT
    If RecordsExist(l_rsModifiedList) Then
        l_strDistributionId = GetAccountDescRemoveFields(v_strDistributionId, True, True, True, False)
        l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, l_strDistributionId, l_rsModifiedList, v_lngWgNum, v_intDescriptionLength)
        If l_blnMatch Then
            GetAccountDescriptionPrivate = l_strAccountDescription
            GoTo ExitFunction
        End If
    End If
    
    'Copying the pick list and removing RMC, FMC, RT, and FeederRef
    Set l_rsModifiedList = ModifyDistrList(v_objRsPickListDistribution, v_lngWgNum, True, True, True, False, True)
 
 ' Remove RMC, FMC, RT & FeederRef
    If RecordsExist(l_rsModifiedList) Then
        l_strDistributionId = GetAccountDescRemoveFields(v_strDistributionId, True, True, True, False, True)
        l_blnMatch = GetAccountDescriptionMatch(l_strAccountDescription, l_strDistributionId, l_rsModifiedList, v_lngWgNum, v_intDescriptionLength)
        If l_blnMatch Then
            GetAccountDescriptionPrivate = l_strAccountDescription
            GoTo ExitFunction
        End If
    End If
    
    GetAccountDescriptionPrivate = ""
'     GetAccountDescriptionPrivate = v_strDistributionId

    
ExitFunction:
'    Set l_objIsWorkReasonIndicatorRequired = Nothing
    Exit Function

GetAccountDescription_EH:
'    Set l_objIsWorkReasonIndicatorRequired = Nothing
    Select Case Err.Number
        Case trErrbusInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetAccountDescriptionPrivate", trErrbusInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetAccountDescriptionPrivate", , , , , True
    End Select

'    TrsRaiseError m_strCMODULENAME, "GetAccountDescriptionPrivate", , , , , True

End Function

'//for older TRIS clients, pick list description is 30 chars
Public Function GetAccountDescriptionJobTask(ByVal v_objRsPickListDistribution As ADODB.Recordset, _
                                      ByVal v_lngWgNum As Long, _
                                      ByVal v_strDistributionId As String, _
                                      ByVal v_strDistributionType As String, _
                                      ByVal v_strJobTask As String, _
                                      ByVal v_strEMPRVDBid As String, _
                                      ByVal v_strDistributionDeptNum As String) As String
On Error GoTo eh

GetAccountDescriptionJobTask = GetAccountDescriptionJobTaskPrivate(v_objRsPickListDistribution, v_lngWgNum, v_strDistributionId, v_strDistributionType, _
                                                                    v_strJobTask, v_strEMPRVDBid, v_strDistributionDeptNum)
Exit Function
eh:
    TrsRaiseError "cbsDistributionMgr", "GetAccountDescriptionJobTask", , , , , True
End Function

'//for newer TRIS clients (OTTC) pick list description increased to 60 chars
Public Function GetAccountDescriptionJobTask2(ByVal v_objRsPickListDistribution As ADODB.Recordset, _
                                      ByVal v_lngWgNum As Long, _
                                      ByVal v_strDistributionId As String, _
                                      ByVal v_strDistributionType As String, _
                                      ByVal v_strJobTask As String, _
                                      ByVal v_strEMPRVDBid As String, _
                                      ByVal v_strDistributionDeptNum As String, _
                                      ByVal v_intDescriptionLength As Integer) As String
On Error GoTo eh

GetAccountDescriptionJobTask2 = GetAccountDescriptionJobTaskPrivate(v_objRsPickListDistribution, v_lngWgNum, v_strDistributionId, v_strDistributionType, _
                                                                    v_strJobTask, v_strEMPRVDBid, v_strDistributionDeptNum, v_intDescriptionLength)
Exit Function
eh:
    TrsRaiseError "cbsDistributionMgr", "GetAccountDescriptionJobTask2", , , , , True
End Function

Private Function GetAccountDescriptionJobTaskPrivate(ByVal v_objRsPickListDistribution As ADODB.Recordset, _
                                      ByVal v_lngWgNum As Long, _
                                      ByVal v_strDistributionId As String, _
                                      ByVal v_strDistributionType As String, _
                                      ByVal v_strJobTask As String, _
                                      ByVal v_strEMPRVDBid As String, _
                                      ByVal v_strDistributionDeptNum As String, _
                                      Optional ByVal v_intDescriptionLength As Integer = 39) As String
    Dim l_objGetJobTaskList As dbDistributionMgr.cdDistributionMgr
    'Dim l_objGetJobTaskList As Object

    Dim l_strFilter As String
    Dim l_strAccountDescription As String
    Dim l_strDistributionId As String
    Dim l_blnMatch As Boolean
'    Dim l_blnIsWorkReasonIndicatorRequired As Boolean
    Dim l_rsModifiedList As ADODB.Recordset
    Dim l_rsEMPRVJobTasks As ADODB.Recordset
    
    On Error GoTo GetAccountDescriptionJobTask_EH
    
    If v_lngWgNum < 1 Then
        Err.Raise trErrbusInvalidZerolngWgNum
    End If
    
    Dim l_colwg As New Collection
    Dim l_strDeptNum As String
    l_colwg.Add v_lngWgNum
    l_strDeptNum = GetDepartmentOfWorkGroup(v_lngWgNum)
    Set l_objGetJobTaskList = CreateObject("dbDistributionMgr.cdDistributionMgr")    'MTS standard
    Set l_rsEMPRVJobTasks = l_objGetJobTaskList.GetDistJobTaskList(l_colwg, l_strDeptNum)
    Set l_objGetJobTaskList = Nothing
    If RecordsExist(l_rsEMPRVJobTasks) Then
        'exact match
        l_strFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND FMC_CD = '" & Mid(v_strDistributionId, 10, 3) & "'" & _
            " AND RMC_CD = '" & Mid(v_strDistributionId, 13, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'" & _
            " AND ACTIV_CD = '" & Mid(v_strDistributionId, 24, 4) & "'" & _
            " AND RES_TP_CD = '" & Mid(v_strDistributionId, 28, 2) & "'"
        '//l_rsEMPRVJobTasks.Filter = MDAC25FilterFix(l_strFilter)
        l_rsEMPRVJobTasks.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(l_strFilter, v_strEMPRVDBid)
        If Not (l_rsEMPRVJobTasks.EOF And l_rsEMPRVJobTasks.BOF) Then
            l_rsEMPRVJobTasks.MoveFirst
            GetAccountDescriptionJobTaskPrivate = PadRight(l_rsEMPRVJobTasks.Fields("DESC60").Value, v_intDescriptionLength)
            GoTo ExitFunction
        End If
        
        l_rsEMPRVJobTasks.Filter = adFilterNone
        'remove rmc and rt
        l_strFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND FMC_CD = '" & Mid(v_strDistributionId, 10, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'" & _
            " AND ACTIV_CD = '" & Mid(v_strDistributionId, 24, 4) & "'"
        '//l_rsEMPRVJobTasks.Filter = MDAC25FilterFix(l_strFilter)
        l_rsEMPRVJobTasks.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(l_strFilter, v_strEMPRVDBid)
        If Not (l_rsEMPRVJobTasks.EOF And l_rsEMPRVJobTasks.BOF) Then
            l_rsEMPRVJobTasks.MoveFirst
            GetAccountDescriptionJobTaskPrivate = PadRight(l_rsEMPRVJobTasks.Fields("DESC60").Value, v_intDescriptionLength)
            GoTo ExitFunction
        End If
        
        l_rsEMPRVJobTasks.Filter = adFilterNone
        'remove rmc and fmc and rt
        l_strFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'" & _
            " AND ACTIV_CD = '" & Mid(v_strDistributionId, 24, 4) & "'"
        '//l_rsEMPRVJobTasks.Filter = MDAC25FilterFix(l_strFilter)
        l_rsEMPRVJobTasks.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(l_strFilter, v_strEMPRVDBid)
        If Not (l_rsEMPRVJobTasks.EOF And l_rsEMPRVJobTasks.BOF) Then
            l_rsEMPRVJobTasks.MoveFirst
            GetAccountDescriptionJobTaskPrivate = PadRight(l_rsEMPRVJobTasks.Fields("DESC60").Value, v_intDescriptionLength)
            GoTo ExitFunction
        End If
        
        l_rsEMPRVJobTasks.Filter = adFilterNone
        'remove rmc and rt and actv
        l_strFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND FMC_CD = '" & Mid(v_strDistributionId, 10, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'"
        '//l_rsEMPRVJobTasks.Filter = MDAC25FilterFix(l_strFilter)
        l_rsEMPRVJobTasks.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(l_strFilter, v_strEMPRVDBid)
        If Not (l_rsEMPRVJobTasks.EOF And l_rsEMPRVJobTasks.BOF) Then
            l_rsEMPRVJobTasks.MoveFirst
            GetAccountDescriptionJobTaskPrivate = PadRight(l_rsEMPRVJobTasks.Fields("DESC60").Value, v_intDescriptionLength)
            GoTo ExitFunction
        End If
        
        l_rsEMPRVJobTasks.Filter = adFilterNone
        'remove rmc and rt and fmc and actv
        l_strFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'"
        '//l_rsEMPRVJobTasks.Filter = MDAC25FilterFix(l_strFilter)
        l_rsEMPRVJobTasks.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(l_strFilter, v_strEMPRVDBid)
        If Not (l_rsEMPRVJobTasks.EOF And l_rsEMPRVJobTasks.BOF) Then
            l_rsEMPRVJobTasks.MoveFirst
            GetAccountDescriptionJobTaskPrivate = PadRight(l_rsEMPRVJobTasks.Fields("DESC60").Value, v_intDescriptionLength)
            GoTo ExitFunction
        End If
        
        l_rsEMPRVJobTasks.Filter = adFilterNone
        'remove fmc and actv
        l_strFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND RMC_CD = '" & Mid(v_strDistributionId, 13, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'" & _
            " AND RES_TP_CD = '" & Mid(v_strDistributionId, 28, 2) & "'"
        '//l_rsEMPRVJobTasks.Filter = MDAC25FilterFix(l_strFilter)
        l_rsEMPRVJobTasks.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(l_strFilter, v_strEMPRVDBid)
        If Not (l_rsEMPRVJobTasks.EOF And l_rsEMPRVJobTasks.BOF) Then
            l_rsEMPRVJobTasks.MoveFirst
            GetAccountDescriptionJobTaskPrivate = PadRight(l_rsEMPRVJobTasks.Fields("DESC60").Value, v_intDescriptionLength)
            GoTo ExitFunction
        End If
        
        l_rsEMPRVJobTasks.Filter = adFilterNone
        'remove actv
        l_strFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND FMC_CD = '" & Mid(v_strDistributionId, 10, 3) & "'" & _
            " AND RMC_CD = '" & Mid(v_strDistributionId, 13, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'" & _
            " AND RES_TP_CD = '" & Mid(v_strDistributionId, 28, 2) & "'"
        '//l_rsEMPRVJobTasks.Filter = MDAC25FilterFix(l_strFilter)
        l_rsEMPRVJobTasks.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(l_strFilter, v_strEMPRVDBid)
        If Not (l_rsEMPRVJobTasks.EOF And l_rsEMPRVJobTasks.BOF) Then
            l_rsEMPRVJobTasks.MoveFirst
            GetAccountDescriptionJobTaskPrivate = PadRight(l_rsEMPRVJobTasks.Fields("DESC60").Value, v_intDescriptionLength)
            GoTo ExitFunction
        End If
  
    End If
    
    GetAccountDescriptionJobTaskPrivate = Trim(v_strJobTask)

    Exit Function
    
ExitFunction:
    Exit Function

GetAccountDescriptionJobTask_EH:
    Select Case Err.Number
        Case trErrbusInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "GetAccountDescriptionJobTaskPrivate", trErrbusInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetAccountDescriptionJobTaskPrivate", , , , , True
    End Select

End Function

'//also used in busTimecard.cbsTimecard
Private Function GetAccountDescriptionJobTask_AddFilterForDBId(ByVal v_strBaseFilter As String, ByVal v_strEMPRVDBid As String) As String
On Error GoTo eh
            '//can only use OR at top level of filter
            '//so the following will not work, Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another.
'            vadrsJobTaskDescription.Filter = _
'                v_strBaseFilter & _
'            " AND (FK_EMPRVDBCODE = '" & v_strEMPRVDBid & "')" & _
'            "     OR FK_EMPRVDBCODE = '')" & _
'            "     OR " & IIf(v_strEMPRVDBid = "", "True", "False") & ")"
    '//logical equivalent
    '//match on db id, or allow emprv's db id to be blank, or allow tris's db id to be blank
    GetAccountDescriptionJobTask_AddFilterForDBId = _
            "(" & v_strBaseFilter & _
            "    AND FK_EMPRVDBCODE = '" & v_strEMPRVDBid & "')" & _
            " OR ( " & v_strBaseFilter & _
            "    AND FK_EMPRVDBCODE = '')" & _
            " OR ( " & v_strBaseFilter & _
            "    AND " & IIf(v_strEMPRVDBid = "", "FK_EMPRVDBCODE<>'NotFound(Result=True)'", "FK_EMPRVDBCODE='NotFound(Result=False)'") & ")"
Exit Function
eh:
    TrsRaiseError "cbsTimeCard", "GetAccountDescriptionJobTask_AddFilterForDBId", , , , , True
End Function

Private Function GetAccountDescriptionMatch( _
            ByRef r_strAccountDescription As String, _
            ByVal v_strDistributionId As String, _
            ByVal v_objRsPickListDistribution As ADODB.Recordset, _
            ByVal v_lngWgNum As Long, _
            Optional ByVal v_intDescriptionLength = 39 _
            ) As Boolean
            
    Dim l_strFilter As String
    
    On Error GoTo GetAccountDescriptionMatch_EH
    
    l_strFilter = "WG_NUM= " & v_lngWgNum & " AND DISTRIBUTION_ID= '" & v_strDistributionId & " '"
    v_objRsPickListDistribution.Filter = MDAC25FilterFix(l_strFilter)

    If Not (v_objRsPickListDistribution.EOF And v_objRsPickListDistribution.BOF) Then
        r_strAccountDescription = PadRight(v_objRsPickListDistribution!Description.Value, v_intDescriptionLength)
        GetAccountDescriptionMatch = True
    Else
        r_strAccountDescription = ""
        GetAccountDescriptionMatch = False
    End If

    v_objRsPickListDistribution.Filter = ""
    Exit Function
    
GetAccountDescriptionMatch_EH:
    v_objRsPickListDistribution.Filter = ""
    TrsRaiseError m_strCMODULENAME, "GetAccountDescriptionMatch"
            
End Function

Private Function GetAccountDescRemoveFields( _
            ByVal v_strDistributionId As String, _
            ByVal v_blnRMC As Boolean, _
            ByVal v_blnRT As Boolean, _
            ByVal v_blnFMC As Boolean, _
            ByVal v_blnACTV As Boolean, _
            Optional ByVal v_blnFeederRef As Boolean = False _
            ) As String
            
    Dim l_strDistID As String
    
    On Error GoTo GetAccountDescRemoveFields_EH:
    
    l_strDistID = Left$(v_strDistributionId, 1) '(U)Utility
    l_strDistID = l_strDistID & Mid$(v_strDistributionId, 2, 2) '(BUD) Budget
    l_strDistID = l_strDistID & Mid$(v_strDistributionId, 4, 3) '(MAJ)FERC Major
    l_strDistID = l_strDistID & Mid$(v_strDistributionId, 7, 3) '(MIN)FERC Minor
    If v_blnFMC Then
        l_strDistID = l_strDistID & Space(3)
    Else
        l_strDistID = l_strDistID & Mid$(v_strDistributionId, 10, 3) '(FMC) Fund Mgt. Center
    End If
    If v_blnRMC Then
        l_strDistID = l_strDistID & Space(3)
    Else
        l_strDistID = l_strDistID & Mid$(v_strDistributionId, 13, 3) '(RMC) Resource Mgt. Center
    End If
    l_strDistID = l_strDistID & Mid$(v_strDistributionId, 16, 1) '(T) Trans Type
    l_strDistID = l_strDistID & Mid$(v_strDistributionId, 17, 5) '(PROJT) Project
    l_strDistID = l_strDistID & Mid$(v_strDistributionId, 22, 2) '(PR) Product
    If v_blnACTV Then
        l_strDistID = l_strDistID & Space(4)
    Else
        l_strDistID = l_strDistID & Mid$(v_strDistributionId, 24, 4) '(ACTV) Activity
    End If
    If v_blnRT Then
        l_strDistID = l_strDistID & Space(2)
    Else
        l_strDistID = l_strDistID & Mid$(v_strDistributionId, 28, 2) '(RT) Resource Type
    End If
'    If v_blnIsWorkReasonIndicatorRequired Then
'        l_strDistID = l_strDistID & Mid$(v_strDistributionId, 30, 8) '(FEEDER_REF) Memo
'    Else
    If v_blnFeederRef Then
        l_strDistID = l_strDistID & Space(10)
    Else
        l_strDistID = l_strDistID & Mid$(v_strDistributionId, 30, 10) '(FEEDER_REF) Memo
    End If
'    End If
           
    GetAccountDescRemoveFields = l_strDistID
    Exit Function
    
GetAccountDescRemoveFields_EH:
    TrsRaiseError m_strCMODULENAME, "GetAccountDescRemoveFields"
            
End Function

Public Function GetWGDistributionIDDescriptions(ByVal v_colWgNum As Collection) As ADODB.Recordset
    On Error GoTo eh
    
    Set GetWGDistributionIDDescriptions = GetWGDistributionIDDescriptions2B(v_colWgNum, 30)

    Exit Function
eh:

    TrsRaiseError m_strCMODULENAME, "GetWGDistributionIDDescriptions", , , , , True
End Function

Public Function GetWGDistributionIDDescriptions2B(ByVal v_colWgNum As Collection, ByVal v_intDescriptionLength As Integer) As ADODB.Recordset
    On Error GoTo eh
    
    Set GetWGDistributionIDDescriptions2B = GetWGDistributionIDDescriptions3(v_colWgNum, v_intDescriptionLength, "unknown")

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetWGDistributionIDDescriptions2B", , , , , True
End Function

Public Function GetWGDistributionIDDescriptions3(ByVal v_colWgNum As Collection, ByVal v_intDescriptionLength As Integer, ByVal v_sCallerMode As String) As ADODB.Recordset

    Dim l_objGetDistributionList As dbDistributionMgr.cdDistributionMgr                         'MTS standard

    On Error GoTo GetWGDistributionIDDescriptions_EH
    
    Set l_objGetDistributionList = CreateObject("dbDistributionMgr.cdDistributionMgr")    'MTS standard
    Set GetWGDistributionIDDescriptions3 = l_objGetDistributionList.GetDistributionList3(v_colWgNum, v_intDescriptionLength, v_sCallerMode)

    Set l_objGetDistributionList = Nothing
    Exit Function

GetWGDistributionIDDescriptions_EH:
    Set l_objGetDistributionList = Nothing
    Set GetWGDistributionIDDescriptions3 = Nothing

    TrsRaiseError m_strCMODULENAME, "GetWGDistributionIDDescriptions3", , , , , True
End Function


Public Function GetDOJMWorkRequestsByRegion(ByVal v_strDistrict As String, ByVal v_strHeadquarters As String) As ADODB.Recordset

    Dim l_objGetDOJMWorkRequestsByRegion As dbDistributionMgr.cdDistributionMgr                         'MTS standard

    On Error GoTo GetDOJMWorkRequestsByRegion_EH
    
    Set l_objGetDOJMWorkRequestsByRegion = CreateObject("dbDistributionMgr.cdDistributionMgr")    'MTS standard
    Set GetDOJMWorkRequestsByRegion = l_objGetDOJMWorkRequestsByRegion.GetDOJMWorkRequestsByRegion(v_strDistrict, v_strHeadquarters)
    Set l_objGetDOJMWorkRequestsByRegion = Nothing
    Exit Function

GetDOJMWorkRequestsByRegion_EH:
    Set l_objGetDOJMWorkRequestsByRegion = Nothing
    Set GetDOJMWorkRequestsByRegion = Nothing

    TrsRaiseError m_strCMODULENAME, "GetDOJMWorkRequestsByRegion", , , , , True

End Function


Public Function GetDOJMWorkRequestsByWorkGroup(ByVal v_lngWgNum As Long, r_strDistrict As String, r_strHeadquarters As String) As ADODB.Recordset
    Dim l_GetDOJMDistrictAndHeadquarters As busWorkGroup.cbsWorkGroup
    On Error GoTo GetDOJMWorkRequestsByWorkGroup_EH

    Set l_GetDOJMDistrictAndHeadquarters = CreateObject("busWorkGroup.cbsWorkGroup")    'MTS standard
    Call l_GetDOJMDistrictAndHeadquarters.GetDOJMDistrictAndHeadquarters(v_lngWgNum, r_strDistrict, r_strHeadquarters)
    
    Set GetDOJMWorkRequestsByWorkGroup = GetDOJMWorkRequestsByRegion(r_strDistrict, r_strHeadquarters)
    
    Set l_GetDOJMDistrictAndHeadquarters = Nothing
    Exit Function

GetDOJMWorkRequestsByWorkGroup_EH:
    Set l_GetDOJMDistrictAndHeadquarters = Nothing
    Set GetDOJMWorkRequestsByWorkGroup = Nothing

    TrsRaiseError m_strCMODULENAME, "GetDOJMWorkRequestsByWorkGroup", , , , , True

End Function


Public Function IsDOJMWorkGroup(ByVal v_lngWgNum As Long) As Boolean
    

On Error GoTo IsDOJMWorkGroup_EH
    If v_lngWgNum < 1 Then
        Err.Raise trErrbusInvalidZerolngWgNum
    Else
    End If

    IsDOJMWorkGroup = IsDOJMWorkGroupPrivate(v_lngWgNum)
    
Exit Function

IsDOJMWorkGroup_EH:
    Select Case Err.Number
        Case trErrbusInvalidZerolngWgNum
            TrsRaiseError m_strCMODULENAME, "IsDOJMWorkGroup", trErrbusInvalidZerolngWgNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "IsDOJMWorkGroup", , , , , True
    End Select


End Function

'//what is the definition of Reversed

'//leave for older TRIS clients 6.9.4
Public Function GetDistributionList(ByVal v_lngWgNum As Long, _
                                    Optional r_strDistrict As String, _
                                    Optional r_strHeadquarters As String _
                                    ) As ADODB.Recordset
On Error GoTo eh

    Set GetDistributionList = GetDistributionListPrivate(v_lngWgNum, r_strDistrict, r_strHeadquarters, True)
Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistributionList", , , , , True
End Function

'//for EMPRV the distribution_id contains job task
'//      and the job task contains the distribution id
Public Function GetDistributionList2(ByVal v_lngWgNum As Long, _
                                    Optional r_strDistrict As String, _
                                    Optional r_strHeadquarters As String _
                                    ) As ADODB.Recordset
On Error GoTo eh

    Set GetDistributionList2 = GetDistributionListPrivate2B(v_lngWgNum, r_strDistrict, r_strHeadquarters, True, 30)
Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistributionList2", , , , , True
End Function

'//for EMPRV the distribution_id contains job task
'//      and the job task contains the distribution id
Public Function GetDistributionList2B(ByVal v_lngWgNum As Long, _
                                    ByVal v_intDescriptionLength As Integer, _
                                    Optional r_strDistrict As String, _
                                    Optional r_strHeadquarters As String _
                                    ) As ADODB.Recordset
On Error GoTo eh

    Set GetDistributionList2B = GetDistributionListPrivate2B(v_lngWgNum, r_strDistrict, r_strHeadquarters, True, v_intDescriptionLength)
Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistributionList2B", , , , , True
End Function

'//leave for older TRIS clients 6.9.4
Public Function GetDistributionListReverse(ByVal v_lngWgNum As Long, _
                                    Optional r_strDistrict As String, _
                                    Optional r_strHeadquarters As String _
                                    ) As ADODB.Recordset
On Error GoTo eh

    Set GetDistributionListReverse = GetDistributionListPrivate(v_lngWgNum, r_strDistrict, r_strHeadquarters, False)
Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistributionListReverse", , , , , True
End Function

'//used by the EMPRV account selection window
'//distribution_id and job task are not reversed
Public Function GetDistributionListReverse2(ByVal v_lngWgNum As Long, _
                                    Optional r_strDistrict As String, _
                                    Optional r_strHeadquarters As String _
                                    ) As ADODB.Recordset
On Error GoTo eh

    Set GetDistributionListReverse2 = GetDistributionListPrivate2B(v_lngWgNum, r_strDistrict, r_strHeadquarters, False, 30)
Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistributionListReverse2", , , , , True
End Function

'//used by the EMPRV account selection window
'//distribution_id and job task are not reversed
Public Function GetDistributionListReverse2B(ByVal v_lngWgNum As Long, _
                                            ByVal v_intDescriptionLength As Integer, _
                                    Optional r_strDistrict As String, _
                                    Optional r_strHeadquarters As String _
                                    ) As ADODB.Recordset
On Error GoTo eh

    Set GetDistributionListReverse2B = GetDistributionListPrivate2B(v_lngWgNum, r_strDistrict, r_strHeadquarters, False, v_intDescriptionLength)
Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistributionListReverse2B", , , , , True
End Function

Public Function GetDistributionListDOJM(ByVal v_lngWgNum As Long, _
                                    Optional r_strDistrict As String, _
                                    Optional r_strHeadquarters As String _
                                    ) As ADODB.Recordset
On Error GoTo eh

    If r_strDistrict = "" Or r_strHeadquarters = "" Then
       Set GetDistributionListDOJM = GetDOJMWorkRequestsByWorkGroup(v_lngWgNum, r_strDistrict, r_strHeadquarters)
    Else
       Set GetDistributionListDOJM = GetDOJMWorkRequestsByRegion(r_strDistrict, r_strHeadquarters)
    End If

Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistributionListDOJM", , , , DescArgsFrom(CStr(v_lngWgNum)), True
End Function

Public Function GetDistributionListEMPRV(ByVal v_lngWgNum As Long, _
                                        ByVal v_blnReverseJobTasks As Boolean _
                                    ) As ADODB.Recordset
    Dim colErrArgs As Collection
On Error GoTo eh
    Dim l_getDistributionList As dbDistributionMgr.cdDistributionMgr
    Dim l_colwg As Collection
    'Call GetDistJobTaskList
    Dim l_strDeptNum As String
    Dim l_rsJobTaskList As ADODB.Recordset
    Dim strDistributionId As String
    Dim strJobTask As String
    
    Set l_colwg = New Collection
    l_colwg.Add v_lngWgNum
    l_strDeptNum = GetDepartmentOfWorkGroup(v_lngWgNum)
    
    Set l_getDistributionList = CreateObject("dbDistributionMgr.cdDistributionMgr")
    Set l_rsJobTaskList = l_getDistributionList.GetDistJobTaskList(l_colwg, l_strDeptNum)
    Set l_getDistributionList = Nothing
    
    '//If RecordsExist(l_rsJobTaskList) Then
    Dim l_rsGetDistList As ADODB.Recordset
    'CREATE THIRD RECORDSET HERE AND COMBINE DATA!!!
    Set l_rsGetDistList = New ADODB.Recordset
    With l_rsGetDistList
        'Set l_rsGetDistList = New ADODB.Recordset
        .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
        .Fields.Append "DISTRIBUTION_ID", adVarChar, 39, adFldUpdatable
        .Fields.Append "DESCRIPTION", adVarChar, 30, adFldUpdatable
        .Fields.Append "JOB_TASK", adVarChar, 39, adFldUpdatable
        .Fields.Append "EMPRV_DATABASE_ID", adVarChar, 8, adFldUpdatable
        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "TIME_STAMP", adVarChar, 26, adFldUpdatable
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
        .ActiveConnection = Nothing
        
        If RecordsExist(l_rsJobTaskList) Then
            l_rsJobTaskList.MoveFirst
            While Not l_rsJobTaskList.EOF
                .AddNew
                !WG_NUM = v_lngWgNum
                strJobTask = Trim(l_rsJobTaskList.Fields("JOB_NBR")) & Trim(l_rsJobTaskList.Fields("TASK_NBR"))
                If l_rsJobTaskList!DISTRIBUTION_TYPE = "J" Then
                    strJobTask = strJobTask & " " & Trim(l_rsJobTaskList.Fields("COSTCNTRCODE"))
                    !Description = Mid(SetNullToEmpty(l_rsJobTaskList.Fields("DESC60")), 1, 30)
                '//the project description and task description have been switched in the Maximo db/view
                ElseIf l_rsJobTaskList!DISTRIBUTION_TYPE = "M" Then
                    !Description = Mid(l_rsJobTaskList.Fields("JOBTASK_AND_DESC"), 12, 30) '//Job/Task and task description
                End If

                strDistributionId = l_rsJobTaskList.Fields("UTLTY_CD") & _
                            l_rsJobTaskList.Fields("BUS_DV_CD") & _
                            l_rsJobTaskList.Fields("MAJ_CD") & _
                            l_rsJobTaskList.Fields("MIN_CD") & _
                            l_rsJobTaskList.Fields("FMC_CD") & _
                            l_rsJobTaskList.Fields("RMC_CD") & _
                            l_rsJobTaskList.Fields("TRN_TP_CD") & _
                            l_rsJobTaskList.Fields("PROJECT_ID") & _
                            l_rsJobTaskList.Fields("PROD_CD") & _
                            l_rsJobTaskList.Fields("ACTIV_CD") & _
                            l_rsJobTaskList.Fields("RES_TP_CD") & _
                            l_rsJobTaskList.Fields("FEEDER_REF")
                '//called by GetDistributionList (reversed)
                If v_blnReverseJobTasks Then
                    !DISTRIBUTION_ID = strJobTask
                    !JOB_TASK = strDistributionId
                Else '//called by GetDistributionListReverse (not really reversed)
                    !DISTRIBUTION_ID = strDistributionId
                    !JOB_TASK = strJobTask
                End If
                
                !EMPRV_DATABASE_ID = Trim(l_rsJobTaskList.Fields("FK_EMPRVDBCODE"))
                !LAST_MOD_EMP_NUM = 999999999
                !TIME_STAMP = ""
                .Update
                
                '//check for possible null values
                '//comment out before moving to production
                If Len(!DISTRIBUTION_ID) <> 39 Then
                    If Len(!JOB_TASK) = 0 Then
                        Set colErrArgs = New Collection
                        colErrArgs.Add "WGNum = " & v_lngWgNum
                        colErrArgs.Add "DeptNum = " & l_strDeptNum
                        colErrArgs.Add "DISTRIBUTION_ID = " & !DISTRIBUTION_ID
                        colErrArgs.Add "JOB_TASK = " & !JOB_TASK
                        colErrArgs.Add "EMPRV_DATABASE_ID = " & !EMPRV_DATABASE_ID
                        Err.Source = ""
                        TrsTraceError m_strCMODULENAME, "GetDistributionList", 20004, False, DescArgsFrom("Distribution ID less than 39 chars, possible null coming from EMPRV"), colErrArgs, True
                        Set colErrArgs = Nothing
                    End If
                End If
                
                l_rsJobTaskList.MoveNext
            Wend
        End If
    End With
    
    Set GetDistributionListEMPRV = l_rsGetDistList
    Set l_colwg = Nothing
    Set colErrArgs = Nothing
Exit Function
eh:
    Set l_getDistributionList = Nothing
    Set l_colwg = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDistributionListEMPRV", , , , DescArgsFrom(CStr(v_lngWgNum)), True
End Function

'Public Function GetDistributionListEMPRV2(ByVal v_lngWgNum As Long, _
'                                        ByVal v_blnReverseJobTasks As Boolean _
'                                    ) As ADODB.Recordset
'On Error GoTo eh
'
'    Set GetDistributionListEMPRV2 = GetDistributionListEMPRV3(v_lngWgNum, v_blnReverseJobTasks, 30)
'
'Exit Function
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetDistributionListEMPRV2", , , , DescArgsFrom(CStr(v_lngWgNum)), True
'End Function

'Public Function GetDistributionListEMPRV3(ByVal v_lngWgNum As Long, _
'                                        ByVal v_blnReverseJobTasks As Boolean, _
'                                        ByVal v_intDescriptionLength As Integer _
'                                    ) As ADODB.Recordset

Public Function GetDistributionListEMPRV2(ByVal v_lngWgNum As Long, _
                                        ByVal v_blnReverseJobTasks As Boolean _
                                    ) As ADODB.Recordset
On Error GoTo eh
    Dim l_rsGetDistList As ADODB.Recordset
    
    Set l_rsGetDistList = GetDistributionListEMPRV2B(v_lngWgNum, v_blnReverseJobTasks, 30)

    Set GetDistributionListEMPRV2 = l_rsGetDistList
Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistributionListEMPRV2", , , , DescArgsFrom(CStr(v_lngWgNum)), True
End Function

Public Function GetDistributionListEMPRV2B(ByVal v_lngWgNum As Long, _
                                        ByVal v_blnReverseJobTasks As Boolean, _
                                        ByVal v_intDescriptionLength As Integer _
                                    ) As ADODB.Recordset
    Dim colErrArgs As Collection
On Error GoTo eh
    Dim l_getDistributionList As dbDistributionMgr.cdDistributionMgr
    Dim l_colwg As Collection
    'Call GetDistJobTaskList
    Dim l_strDeptNum As String
    Dim l_rsJobTaskList As ADODB.Recordset
    Dim strDistributionId As String
    Dim strJobTask As String
'    Dim v_intDescriptionLength As Integer
'    v_intDescriptionLength = 60
    
    Set l_colwg = New Collection
    l_colwg.Add v_lngWgNum
    l_strDeptNum = GetDepartmentOfWorkGroup(v_lngWgNum)
    
    Set l_getDistributionList = CreateObject("dbDistributionMgr.cdDistributionMgr")
    Set l_rsJobTaskList = l_getDistributionList.GetDistJobTaskList(l_colwg, l_strDeptNum)
    Set l_getDistributionList = Nothing
    
    '//If RecordsExist(l_rsJobTaskList) Then
    Dim l_rsGetDistList As ADODB.Recordset
    'CREATE THIRD RECORDSET HERE AND COMBINE DATA!!!
    Set l_rsGetDistList = New ADODB.Recordset
    With l_rsGetDistList
        'Set l_rsGetDistList = New ADODB.Recordset
        .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
        '//.Fields.Append "DISTRIBUTION_ID", adVarChar, 39, adFldUpdatable
        .Fields.Append "DISTRIBUTION_ID", adVarChar, v_intDescriptionLength, adFldUpdatable
        .Fields.Append "DESCRIPTION", adVarChar, v_intDescriptionLength, adFldUpdatable
        '//.Fields.Append "JOB_TASK", adVarChar, 39, adFldUpdatable
        .Fields.Append "JOB_TASK", adVarChar, v_intDescriptionLength, adFldUpdatable
        .Fields.Append "EMPRV_DATABASE_ID", adVarChar, 8, adFldUpdatable
        .Fields.Append "FEEDER_REF", adVarChar, 10, adFldUpdatable
        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "TIME_STAMP", adVarChar, 26, adFldUpdatable
        
        .Fields.Append "TRIS_PROJECT_AND_TASK_IDS", adChar, 19, adFldIsNullable
        .Fields.Append "PROJECT_NUMBER", adVarChar, 10, adFldIsNullable
        .Fields.Append "PROJECT_NAME", adVarChar, 60, adFldIsNullable
        .Fields.Append "TASK_NUMBER", adVarChar, 21, adFldIsNullable
        .Fields.Append "TASK_NAME", adVarChar, 60, adFldIsNullable

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
        .ActiveConnection = Nothing
        
        If RecordsExist(l_rsJobTaskList) Then
            l_rsJobTaskList.MoveFirst
            While Not l_rsJobTaskList.EOF
                .AddNew
                !WG_NUM = v_lngWgNum
                strJobTask = Trim(l_rsJobTaskList.Fields("JOB_NBR")) & Trim(l_rsJobTaskList.Fields("TASK_NBR"))
                If l_rsJobTaskList!DISTRIBUTION_TYPE = "J" Then
                    strJobTask = strJobTask & " " & Trim(l_rsJobTaskList.Fields("COSTCNTRCODE"))
                ElseIf l_rsJobTaskList!DISTRIBUTION_TYPE = "M" Then
                    strJobTask = Trim(l_rsJobTaskList.Fields("JOBTASK_AND_DESC")) '//Job/Task and task description
                End If
                !Description = PadRight(SetNullToEmpty(l_rsJobTaskList.Fields("DESC60")), v_intDescriptionLength)
                strDistributionId = l_rsJobTaskList.Fields("UTLTY_CD") & _
                            l_rsJobTaskList.Fields("BUS_DV_CD") & _
                            l_rsJobTaskList.Fields("MAJ_CD") & _
                            l_rsJobTaskList.Fields("MIN_CD") & _
                            l_rsJobTaskList.Fields("FMC_CD") & _
                            l_rsJobTaskList.Fields("RMC_CD") & _
                            l_rsJobTaskList.Fields("TRN_TP_CD") & _
                            l_rsJobTaskList.Fields("PROJECT_ID") & _
                            l_rsJobTaskList.Fields("PROD_CD") & _
                            l_rsJobTaskList.Fields("ACTIV_CD") & _
                            l_rsJobTaskList.Fields("RES_TP_CD") & _
                            l_rsJobTaskList.Fields("FEEDER_REF")
                '//called by GetDistributionList (reversed)
                If v_blnReverseJobTasks Then
                    !DISTRIBUTION_ID = PadRight(strJobTask, v_intDescriptionLength) '//39)
                    !JOB_TASK = strDistributionId
                Else '//called by GetDistributionListReverse (not really reversed)
                    !DISTRIBUTION_ID = strDistributionId
                    !JOB_TASK = PadRight(strJobTask, v_intDescriptionLength)
                End If
                
                !EMPRV_DATABASE_ID = Trim(l_rsJobTaskList.Fields("FK_EMPRVDBCODE"))
                !FEEDER_REF = Trim(l_rsJobTaskList.Fields("FEEDER_REF"))
                !LAST_MOD_EMP_NUM = 999999999
                !TIME_STAMP = ""
                
                !TRIS_PROJECT_AND_TASK_IDS = l_rsJobTaskList.Fields("TRIS_PROJECT_AND_TASK_IDS").Value
                !PROJECT_NUMBER = l_rsJobTaskList.Fields("PROJECT_NUMBER").Value
                !PROJECT_NAME = l_rsJobTaskList.Fields("PROJECT_NAME").Value
                !TASK_NUMBER = l_rsJobTaskList.Fields("TASK_NUMBER").Value
                !TASK_NAME = l_rsJobTaskList.Fields("TASK_NAME").Value
                .Update
                
                '//check for possible null values
                '//comment out before moving to production
                If Len(!DISTRIBUTION_ID) <> 39 Then
                    If Len(!JOB_TASK) = 0 Then
                        Set colErrArgs = New Collection
                        colErrArgs.Add "WGNum = " & v_lngWgNum
                        colErrArgs.Add "DeptNum = " & l_strDeptNum
                        colErrArgs.Add "DISTRIBUTION_ID = " & !DISTRIBUTION_ID
                        colErrArgs.Add "JOB_TASK = " & !JOB_TASK
                        colErrArgs.Add "EMPRV_DATABASE_ID = " & !EMPRV_DATABASE_ID
                        Err.Source = ""
                        TrsTraceError m_strCMODULENAME, "GetDistributionList", 20004, False, DescArgsFrom("Distribution ID less than 39 chars, possible null coming from EMPRV"), colErrArgs, True
                        Set colErrArgs = Nothing
                    End If
                End If
                
                l_rsJobTaskList.MoveNext
            Wend
        End If
    End With
    
    Set GetDistributionListEMPRV2B = l_rsGetDistList
    Set l_colwg = Nothing
    Set colErrArgs = Nothing
Exit Function
eh:
    Set l_getDistributionList = Nothing
    Set l_colwg = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDistributionListEMPRV2B", , , , DescArgsFrom(CStr(v_lngWgNum)), True
End Function

'//leave for older TRIS clients 6.9.4
Private Function GetDistributionListPrivate(ByVal v_lngWgNum As Long, _
                                            ByRef r_strDistrict As String, _
                                            ByRef r_strHeadquarters As String, _
                                            ByVal v_blnReverseJobTasks As Boolean _
                                            ) As ADODB.Recordset
    Dim l_colwg As Collection
    Dim colErrArgs As Collection
    Dim l_rsWGPickList As ADODB.Recordset
    Dim l_rsEMPRVPickList As ADODB.Recordset
On Error GoTo eh
    
    '//If IsDOJMWorkGroup(v_lngWgNum) Then
        '//Set GetDistributionListPrivate = GetDistributionListDOJM(v_lngWgNum, r_strDistrict, r_strHeadquarters)
    '//Else
        Dim l_getDistributionList As dbDistributionMgr.cdDistributionMgr
                    
        Set l_colwg = New Collection
        l_colwg.Add v_lngWgNum
        
        Set l_getDistributionList = CreateObject("dbDistributionMgr.cdDistributionMgr")
        Set l_rsWGPickList = l_getDistributionList.GetDistributionList(l_colwg)
        Set l_getDistributionList = Nothing

        If UseEMPRVJobTasks(v_lngWgNum) Then
            Set l_rsEMPRVPickList = GetDistributionListEMPRV(v_lngWgNum, v_blnReverseJobTasks)
                                
            '//merge in the workgroup pick list
            If RecordsExist(l_rsWGPickList) Then
                l_rsWGPickList.MoveFirst
                While Not l_rsWGPickList.EOF
                    l_rsEMPRVPickList.AddNew
                    l_rsEMPRVPickList!WG_NUM = l_rsWGPickList.Fields("WG_NUM")
                    l_rsEMPRVPickList!DISTRIBUTION_ID = l_rsWGPickList.Fields("DISTRIBUTION_ID")
                    l_rsEMPRVPickList!Description = l_rsWGPickList.Fields("DESCRIPTION")
                    l_rsEMPRVPickList!JOB_TASK = l_rsWGPickList.Fields("JOB_TASK")
                    l_rsEMPRVPickList!EMPRV_DATABASE_ID = l_rsWGPickList.Fields("EMPRV_DATABASE_ID")
                    l_rsEMPRVPickList!LAST_MOD_EMP_NUM = l_rsWGPickList.Fields("LAST_MOD_EMP_NUM")
                    l_rsEMPRVPickList!TIME_STAMP = l_rsWGPickList.Fields("TIME_STAMP")
                    l_rsEMPRVPickList.Update
                    
                    l_rsWGPickList.MoveNext
                Wend
            End If
            
            Set GetDistributionListPrivate = l_rsEMPRVPickList
        Else
            Set GetDistributionListPrivate = l_rsWGPickList
        End If
        
    '//End If
    Exit Function
    
    Set l_colwg = Nothing
eh:
    Set l_getDistributionList = Nothing
    Set l_colwg = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDistributionListPrivate", , , , DescArgsFrom(CStr(v_lngWgNum)), True
End Function

'//called by
'//GetDistributionList2
'//GetDistributionListReverse2
Private Function GetDistributionListPrivate2B(ByVal v_lngWgNum As Long, _
                                            ByRef r_strDistrict As String, _
                                            ByRef r_strHeadquarters As String, _
                                            ByVal v_blnReverseJobTasks As Boolean, _
                                            ByVal v_intDescriptionLength As Integer _
                                            ) As ADODB.Recordset
    Dim l_colwg As Collection
    Dim colErrArgs As Collection
    Dim l_rsWGPickList As ADODB.Recordset
    Dim l_rsEMPRVPickList As ADODB.Recordset
On Error GoTo eh
    
    '//If IsDOJMWorkGroup(v_lngWgNum) Then
        '//Set GetDistributionListPrivate2B = GetDistributionListDOJM(v_lngWgNum, r_strDistrict, r_strHeadquarters)
    '//Else
        Dim l_getDistributionList As dbDistributionMgr.cdDistributionMgr
                    
        Set l_colwg = New Collection
        l_colwg.Add v_lngWgNum
        
        Set l_getDistributionList = CreateObject("dbDistributionMgr.cdDistributionMgr")
        Set l_rsWGPickList = l_getDistributionList.GetDistributionList3(l_colwg, v_intDescriptionLength, "PickList")
        Set l_getDistributionList = Nothing

        If UseEMPRVJobTasks(v_lngWgNum) Then
            Set l_rsEMPRVPickList = GetDistributionListEMPRV2B(v_lngWgNum, v_blnReverseJobTasks, v_intDescriptionLength)
                                
            '//merge in the workgroup pick list
            If RecordsExist(l_rsWGPickList) Then
                l_rsWGPickList.MoveFirst
                While Not l_rsWGPickList.EOF
                    l_rsEMPRVPickList.AddNew
                    l_rsEMPRVPickList!WG_NUM = l_rsWGPickList.Fields("WG_NUM")
                    l_rsEMPRVPickList!DISTRIBUTION_ID = l_rsWGPickList.Fields("DISTRIBUTION_ID")
                    l_rsEMPRVPickList!Description = l_rsWGPickList.Fields("DESCRIPTION")
                    l_rsEMPRVPickList!JOB_TASK = l_rsWGPickList.Fields("JOB_TASK")
                    l_rsEMPRVPickList!EMPRV_DATABASE_ID = l_rsWGPickList.Fields("EMPRV_DATABASE_ID")
                    l_rsEMPRVPickList!FEEDER_REF = l_rsWGPickList.Fields("FEEDER_REF")
                    l_rsEMPRVPickList!LAST_MOD_EMP_NUM = l_rsWGPickList.Fields("LAST_MOD_EMP_NUM")
                    l_rsEMPRVPickList!TIME_STAMP = l_rsWGPickList.Fields("TIME_STAMP")
                    
                    l_rsEMPRVPickList!TRIS_PROJECT_AND_TASK_IDS = l_rsWGPickList.Fields("TRIS_PROJECT_AND_TASK_IDS").Value
                    l_rsEMPRVPickList!PROJECT_NUMBER = l_rsWGPickList.Fields("PROJECT_NUMBER").Value
                    l_rsEMPRVPickList!PROJECT_NAME = l_rsWGPickList.Fields("PROJECT_NAME").Value
                    l_rsEMPRVPickList!TASK_NUMBER = l_rsWGPickList.Fields("TASK_NUMBER").Value
                    l_rsEMPRVPickList!TASK_NAME = l_rsWGPickList.Fields("TASK_NAME").Value
                    l_rsEMPRVPickList.Update
                    
                    l_rsWGPickList.MoveNext
                Wend
            End If
            
            Set GetDistributionListPrivate2B = l_rsEMPRVPickList
        Else
            Set GetDistributionListPrivate2B = l_rsWGPickList
        End If
        
    '//End If
    Exit Function
    
    Set l_colwg = Nothing
eh:
    Set l_getDistributionList = Nothing
    Set l_colwg = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDistributionListPrivate2B", , , , DescArgsFrom(CStr(v_lngWgNum)), True
End Function

'Public Function GetDistributionListReverse(ByVal v_lngWgNum As Long, _
'                                    Optional r_strDistrict As String, _
'                                    Optional r_strHeadquarters As String _
'                                    ) As ADODB.Recordset
'
'    Dim colErrArgs As Collection
'
'On Error GoTo eh
'
'    If IsDOJMWorkGroup(v_lngWgNum) Then
'       If r_strDistrict = "" Or r_strHeadquarters = "" Then
'          Set GetDistributionListReverse = GetDOJMWorkRequestsByWorkGroup(v_lngWgNum, r_strDistrict, r_strHeadquarters)
'       Else
'          Set GetDistributionListReverse = GetDOJMWorkRequestsByRegion(r_strDistrict, r_strHeadquarters)
'       End If
'    Else
'        Dim l_colwg As New Collection
'        Dim l_getDistributionList As dbDistributionMgr.cdDistributionMgr
'        If UseEMPRVJobTasks(v_lngWgNum) Then
'            'Call GetDistJobTaskList
'            Dim l_strDeptNum As String
'            Dim l_rsJobTaskList As ADODB.Recordset
'            Dim l_rsDistList As ADODB.Recordset
'            Set l_getDistributionList = CreateObject("dbDistributionMgr.cdDistributionMgr")
'            l_colwg.Add v_lngWgNum
'            l_strDeptNum = GetDepartmentOfWorkGroup(v_lngWgNum)
'            Set l_rsJobTaskList = l_getDistributionList.GetDistJobTaskList(l_colwg, l_strDeptNum)
'            Set l_rsDistList = l_getDistributionList.GetDistributionList(l_colwg)
'            If RecordsExist(l_rsJobTaskList) Or RecordsExist(l_rsDistList) Then
'                Dim l_rsGetDistList As ADODB.Recordset
'                'CREATE THIRD RECORDSET HERE AND COMBINE DATA!!!
'                Set l_rsGetDistList = New ADODB.Recordset
'                With l_rsGetDistList
'                    'Set l_rsGetDistList = New ADODB.Recordset
'                    .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
'                    .Fields.Append "DISTRIBUTION_ID", adVarChar, 39, adFldUpdatable
'                    .Fields.Append "DESCRIPTION", adVarChar, 30, adFldUpdatable
'                    .Fields.Append "JOB_TASK", adVarChar, 39, adFldUpdatable
'                    .Fields.Append "EMPRV_DATABASE_ID", adVarChar, 8, adFldUpdatable
'                    .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
'                    .Fields.Append "TIME_STAMP", adVarChar, 26, adFldUpdatable
'                    .CursorLocation = adUseClient
'                    .LockType = adLockBatchOptimistic
'                    .CursorType = adOpenStatic
'                    .Open
'                    .ActiveConnection = Nothing
'
'                    If RecordsExist(l_rsJobTaskList) Then
'                        l_rsJobTaskList.MoveFirst
'                        While Not l_rsJobTaskList.EOF
'                            .AddNew
'                            !WG_NUM = v_lngWgNum
'                            !DISTRIBUTION_ID = l_rsJobTaskList.Fields("UTLTY_CD") & _
'                                               l_rsJobTaskList.Fields("BUS_DV_CD") & _
'                                               l_rsJobTaskList.Fields("MAJ_CD") & _
'                                               l_rsJobTaskList.Fields("MIN_CD") & _
'                                               l_rsJobTaskList.Fields("FMC_CD") & _
'                                               l_rsJobTaskList.Fields("RMC_CD") & _
'                                               l_rsJobTaskList.Fields("TRN_TP_CD") & _
'                                               l_rsJobTaskList.Fields("PROJECT_ID") & _
'                                               l_rsJobTaskList.Fields("PROD_CD") & _
'                                               l_rsJobTaskList.Fields("ACTIV_CD") & _
'                                               l_rsJobTaskList.Fields("RES_TP_CD") & _
'                                               "          "
'                            '!DISTRIBUTION_ID = Trim(l_rsJobTaskList.Fields("JOB_NBR")) & Trim(l_rsJobTaskList.Fields("TASK_NBR")) & _
'                                        " " & Trim(l_rsJobTaskList.Fields("COSTCNTRCODE"))
'                            '!Description = Trim(l_rsJobTaskList.Fields("JOB_NBR")) & _
'                                           Trim(l_rsJobTaskList.Fields("TASK_NBR")) & " " & _
'                                           Mid(l_rsJobTaskList.Fields("CCDESC"), 1, 19)
'                            !Description = Mid(SetNullToEmpty(l_rsJobTaskList.Fields("DESC60")), 1, 30)
'                            !JOB_TASK = Trim(l_rsJobTaskList.Fields("JOB_NBR")) & Trim(l_rsJobTaskList.Fields("TASK_NBR")) & _
'                                        " " & Trim(l_rsJobTaskList.Fields("COSTCNTRCODE"))
'                            '!JOB_TASK = l_rsJobTaskList.Fields("UTLTY_CD") & _
'                                        l_rsJobTaskList.Fields("BUS_DV_CD") & _
'                                        l_rsJobTaskList.Fields("MAJ_CD") & _
'                                        l_rsJobTaskList.Fields("MIN_CD") & _
'                                        l_rsJobTaskList.Fields("FMC_CD") & _
'                                        l_rsJobTaskList.Fields("RMC_CD") & _
'                                        l_rsJobTaskList.Fields("TRN_TP_CD") & _
'                                        l_rsJobTaskList.Fields("PROJECT_ID") & _
'                                        l_rsJobTaskList.Fields("PROD_CD") & _
'                                        l_rsJobTaskList.Fields("ACTIV_CD") & _
'                                        l_rsJobTaskList.Fields("RES_TP_CD") & _
'                                        "          "
'                            !EMPRV_DATABASE_ID = Trim(l_rsJobTaskList.Fields("FK_EMPRVDBCODE"))
'                            !LAST_MOD_EMP_NUM = 999999999
'                            !TIME_STAMP = ""
'                            '!DISTRIBUTION_ID = l_rsJobTaskList.Fields("DISTRIBUTION_ID")
'                            '!Description = Mid(l_rsJobTaskList.Fields("DESCRIPTION"), 1, 30)
'                            '!JOB_TASK = l_rsJobTaskList.Fields("JOB_TASK")
'                            '!EMPRV_DATABASE_ID = l_rsJobTaskList.Fields("EMPRV_DATABASE_ID")
'                            '!LAST_MOD_EMP_NUM = l_rsJobTaskList.Fields("LAST_MOD_EMP_NUM")
'                            '!TIME_STAMP = l_rsJobTaskList.Fields("TIME_STAMP")
'                            .Update
'
'                            '//check for possible null values
'                            '//comment out before moving to production
'                            If Len(!DISTRIBUTION_ID) <> 39 Then
'                                If Len(!JOB_TASK) = 0 Then
'                                    Set colErrArgs = New Collection
'                                    colErrArgs.Add "WGNum = " & v_lngWgNum
'                                    colErrArgs.Add "DeptNum = " & l_strDeptNum
'                                    colErrArgs.Add "DISTRIBUTION_ID = " & !DISTRIBUTION_ID
'                                    colErrArgs.Add "JOB_TASK = " & !JOB_TASK
'                                    colErrArgs.Add "EMPRV_DATABASE_ID = " & !EMPRV_DATABASE_ID
'                                    Err.Source = ""
'                                    TrsTraceError m_strCMODULENAME, "GetDistributionList", 20004, False, DescArgsFrom("Distribution ID less than 39 chars, possible null coming from EMPRV"), colErrArgs, True
'                                    Set colErrArgs = Nothing
'                                End If
'                            End If
'
'                            l_rsJobTaskList.MoveNext
'                        Wend
'                    End If
'
'                    If RecordsExist(l_rsDistList) Then
'                        l_rsDistList.MoveFirst
'                        While Not l_rsDistList.EOF
'                            .AddNew
'                            !WG_NUM = l_rsDistList.Fields("WG_NUM")
'                            !DISTRIBUTION_ID = l_rsDistList.Fields("DISTRIBUTION_ID")
'                            !Description = l_rsDistList.Fields("DESCRIPTION")
'                            !JOB_TASK = l_rsDistList.Fields("JOB_TASK")
'                            !EMPRV_DATABASE_ID = l_rsDistList.Fields("EMPRV_DATABASE_ID")
'                            !LAST_MOD_EMP_NUM = l_rsDistList.Fields("LAST_MOD_EMP_NUM")
'                            !TIME_STAMP = l_rsDistList.Fields("TIME_STAMP")
'                            .Update
'                            l_rsDistList.MoveNext
'                        Wend
'                    End If
'                End With
'            End If
'
'            'GetDistributionList = l_getDistributionList.GetDistJobTaskList(l_colwg)
'            Set GetDistributionListReverse = l_rsGetDistList
'
'            Set l_getDistributionList = Nothing
'        Else
'           Set l_getDistributionList = CreateObject("dbDistributionMgr.cdDistributionMgr")
'           l_colwg.Add v_lngWgNum
'           Set GetDistributionListReverse = l_getDistributionList.GetDistributionList(l_colwg)
'
'           Set l_getDistributionList = Nothing
'        End If
'    End If
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetDistributionListReverse", , , , , True
'End Function

'Public Function GetDistributionListforSelect(ByVal v_lngWgNum As Long, _
                                    Optional r_strDistrict As String, _
                                    Optional r_strHeadquarters As String _
                                    ) As ADODB.Recordset
    
'    On Error GoTo GetDistributionListforSelect_EH
    
    'If IsDOJMWorkGroup(v_lngWgNum) Then
    '   If r_strDistrict = "" Or r_strHeadquarters = "" Then
    '      Set GetDistributionListforSelect = GetDOJMWorkReqByWGforSelect(v_lngWgNum, r_strDistrict, r_strHeadquarters)
    '   Else
    '      Set GetDistributionListforSelect = GetDOJMWorkRqByRegionforSelect(r_strDistrict, r_strHeadquarters)
    '   End If
    'Else
'        If UseEMPRVJobTasks(v_lngWgNum) Then
'            'Call GetDistandJobTaskList
'            Dim l_colWG As New Collection
'            Dim l_getDistributionList As dbDistributionMgr.cdDistributionMgr
'            Set l_getDistributionList = CreateObject("dbDistributionMgr.cdDistributionMgr")
'            l_colWG.Add v_lngWgNum
'            Set GetDistributionListforSelect = l_getDistributionList.GetDistandJobTaskListforSelect(l_colWG)
'
'            Set l_getDistributionList = Nothing
'        Else
'           Dim l_colWG As New Collection
'           Dim l_getDistributionList As dbDistributionMgr.cdDistributionMgr
'           Set l_getDistributionList = CreateObject("dbDistributionMgr.cdDistributionMgr")
'           l_colWG.Add v_lngWgNum
'           Set GetDistributionListforSelect = l_getDistributionList.GetDistributionListforSelect(l_colWG)
'
'           Set l_getDistributionList = Nothing
'        End If
    'End If
'    Exit Function

'GetDistributionListforSelect_EH:
'    TrsRaiseError m_strCMODULENAME, "GetDistributionListforSelect", , , , , True

'End Function


Public Sub GetECSupportInformation( _
        ByRef r_rsAllWorkNumbers As ADODB.Recordset, _
        ByRef r_rsAllClassNumbers As ADODB.Recordset, _
        ByRef r_rsAllDrawingSizes As ADODB.Recordset, _
        ByRef r_rsAllWorkReasons As ADODB.Recordset, _
        ByRef r_strECAccountNumNonProdTime As String, _
        ByRef r_strECDescriptionNonProdTime As String)
        
TrsRaiseError m_strCMODULENAME, "GetECSupportInformation", trErrObsoleteMethod, , , , True
'    Dim l_objdbDistMgr As dbDistributionMgr.cdDistributionMgr
'On Error GoTo GetECSupportInformation_EH
'    Set l_objdbDistMgr = CreateObject("dbDistributionMgr.cdDistributionMgr")
'
'    Set r_rsAllWorkNumbers = l_objdbDistMgr.GetValidECWorkNumbers
'    Set r_rsAllClassNumbers = l_objdbDistMgr.GetValidECClassNumbers
'    Set r_rsAllDrawingSizes = l_objdbDistMgr.GetValidECDrawingSizes
'    Set r_rsAllWorkReasons = l_objdbDistMgr.GetECWorkReason
'    r_strECAccountNumNonProdTime = l_objdbDistMgr.GetECAccountNumNonProdTime
'    r_strECDescriptionNonProdTime = l_objdbDistMgr.GetECDescriptionNonProdTime
'
'    Set l_objdbDistMgr = Nothing
'
'    Exit Sub
'GetECSupportInformation_EH:
'
'    Set r_rsAllWorkNumbers = Nothing
'    Set r_rsAllClassNumbers = Nothing
'    Set r_rsAllDrawingSizes = Nothing
'    Set r_rsAllWorkReasons = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetECSupportInformation", , , , , True

End Sub

Public Sub GetAllECInformation( _
        ByRef r_rsECClassNumbers As ADODB.Recordset, _
        ByRef r_rsECWorkNumbers As ADODB.Recordset, _
        ByRef r_rsECDrawingSizes As ADODB.Recordset, _
        ByRef r_rsECWorkReasons As ADODB.Recordset)
        
TrsRaiseError m_strCMODULENAME, "GetAllECInformation", trErrObsoleteMethod, , , , True
'    Dim l_objdbDistMgr As dbDistributionMgr.cdDistributionMgr
'
'
'On Error GoTo GetAllECInformation_EH
'
'
'    Set l_objdbDistMgr = CreateObject("dbDistributionMgr.cdDistributionMgr")
'
'    Set r_rsECClassNumbers = l_objdbDistMgr.GetValidECClassNumbers
'    Set r_rsECWorkNumbers = l_objdbDistMgr.GetValidECWorkNumbers
'    Set r_rsECDrawingSizes = l_objdbDistMgr.GetValidECDrawingSizes
'    Set r_rsECWorkReasons = l_objdbDistMgr.GetECWorkReason
'    Set l_objdbDistMgr = Nothing
'
'Exit Sub
'GetAllECInformation_EH:
'    Set r_rsECClassNumbers = Nothing
'    Set r_rsECWorkNumbers = Nothing
'    Set r_rsECDrawingSizes = Nothing
'    Set r_rsECWorkReasons = Nothing
'    Set l_objdbDistMgr = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetAllECInformation", , , , , True

End Sub

'OBSOLETE, I thought, but is called from busOTTimecard
Public Function ValidateDistributionIdsFromTimeEntries( _
            ByVal v_rsTimeEntries As ADODB.Recordset, _
            ByVal v_rsOtherPayEntries As ADODB.Recordset, _
            ByVal v_rsEmployeeData As ADODB.Recordset, _
            ByVal v_rsEmployeeWGData As ADODB.Recordset, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            ByRef r_rsBRVs As ADODB.Recordset _
            ) As Boolean
    On Error GoTo eh
    
    Dim l_rsJobTaskSteps As ADODB.Recordset
    
    Set l_rsJobTaskSteps = CreateSaveJobTaskStepRecordset
    ValidateDistributionIdsFromTimeEntries = ValidateDistributionIdsFromTimeEntries2(v_rsTimeEntries, v_rsOtherPayEntries, v_rsEmployeeData, v_rsEmployeeWGData, v_datStartDate, v_datEndDate, r_rsBRVs, l_rsJobTaskSteps)
    
    Exit Function
eh:
    Set r_rsBRVs = Nothing
    Set v_rsTimeEntries = Nothing
    Set v_rsOtherPayEntries = Nothing
    Set l_rsJobTaskSteps = Nothing
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIdsFromTimeEntries", Err.Number, , , , True
End Function

'//from TRIS32 client Hack
Private Function CreateSaveJobTaskStepRecordset() As ADODB.Recordset
On Error GoTo CreateSaveJobTaskStepRecordset_EH
Const l_lngCFIELDATTR = adFldUpdatable Or adFldIsNullable
Dim l_adrsJobTaskStep As ADODB.Recordset

    ' Create the empty recordset
    Set l_adrsJobTaskStep = New ADODB.Recordset
    With l_adrsJobTaskStep
        .Fields.Append "JOB_NBR", adChar, 8, l_lngCFIELDATTR
        .Fields.Append "JOBTASK_NBR", adChar, 4, l_lngCFIELDATTR
        .Fields.Append "EMPRV_DATABASE_ID", adChar, 8, l_lngCFIELDATTR
        .Fields.Append "STEP_NBR", adInteger, , l_lngCFIELDATTR
        .Fields.Append "VALIDATION_TYPE", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "PROJECT_NUMBER", adVarChar, 10, l_lngCFIELDATTR
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateSaveJobTaskStepRecordset = l_adrsJobTaskStep

    'Destroy Object
    Set l_adrsJobTaskStep = Nothing
    Exit Function
CreateSaveJobTaskStepRecordset_EH:
    Set l_adrsJobTaskStep = Nothing

    '//Err.Raise Err.Number, MODULENAME & ": CreateSaveJobTaskStepRecordset"
    TrsRaiseError m_strCMODULENAME, "CreateSaveJobTaskStepRecordset", Err.Number, , , , True

End Function

Public Function ValidateDistributionIdsFromTimeEntries2( _
            ByVal v_rsTimeEntries As ADODB.Recordset, _
            ByVal v_rsOtherPayEntries As ADODB.Recordset, _
            ByVal v_rsEmployeeData As ADODB.Recordset, _
            ByVal v_rsEmployeeWGData As ADODB.Recordset, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            ByRef r_rsBRVs As ADODB.Recordset, _
            ByVal v_rsJobTaskSteps As ADODB.Recordset _
            ) As Boolean

' v_rsTimeEntries and v_rsOtherPayEntries are defined as parameters
' ByVal, but they are actually ByRef
' Code in this module requires them to be ByRef
'  because it update distribution_id prior to the
'  time being saved


    Dim l_rsEmployeeDistIDInfo As ADODB.Recordset
    Dim l_bolTimeEntryValidationSuccessful As Boolean
    Dim l_bolOtherPayEntryValidationSuccessful As Boolean
    Dim l_bolWorkOrderValidationSuccessful As Boolean
    Dim l_bolJobTaskValidationSuccessful As Boolean
    Dim l_bolRequiredDistIDsEntered As Boolean
    Dim l_bolRequiredDistIDsEnteredOtherPay As Boolean
    Dim l_bolValidDistribution As Boolean
    
    On Error GoTo eh
    
    ValidateDistributionIdsFromTimeEntries2 = False

    Set l_rsEmployeeDistIDInfo = GetEmployeeDistIdInfo( _
                                   v_rsEmployeeData, _
                                   v_datStartDate, _
                                   v_datEndDate)
                                                           
        
    'Remove Distribution for employees in MKS departments first
    ' The removal process is date driven
    GetMKSEffDate
    MKSDeptsRemoveDistribution v_rsTimeEntries, l_rsEmployeeDistIDInfo, r_rsBRVs
    
    SetChargeWGForDOJMDefaultAccounting v_rsTimeEntries, l_rsEmployeeDistIDInfo
                                                           
    ValidateDistributionIDsDefaultAccounting v_rsTimeEntries, l_rsEmployeeDistIDInfo
    ValidateDistributionIDsDefaultAccounting v_rsOtherPayEntries, l_rsEmployeeDistIDInfo, True

    l_bolTimeEntryValidationSuccessful = ValidateDistributionIDsAccountRequests(l_rsEmployeeDistIDInfo, v_rsEmployeeWGData, v_rsTimeEntries, r_rsBRVs, v_rsJobTaskSteps)
    l_bolOtherPayEntryValidationSuccessful = True
    l_bolOtherPayEntryValidationSuccessful = ValidateDistributionIDsAccountRequests(l_rsEmployeeDistIDInfo, v_rsEmployeeWGData, v_rsOtherPayEntries, r_rsBRVs, v_rsJobTaskSteps, True)
    l_bolWorkOrderValidationSuccessful = ValidateDistributionIDsWorkRequests(l_rsEmployeeDistIDInfo, v_rsEmployeeWGData, v_rsTimeEntries, r_rsBRVs)
    'l_bolJobTaskValidationSuccessful = _
    '        ValidateDistributionIDsJobTasks(l_rsEmployeeDistIDInfo, v_rsEmployeeWGData, v_rsTimeEntries, r_rsBRVs)
    l_bolRequiredDistIDsEntered = ValidateDistributionIDRequiredDistIDsEntered(l_rsEmployeeDistIDInfo, v_rsEmployeeWGData, v_rsTimeEntries, r_rsBRVs)
                                            ' Might need this v_rsEmployeeWGData
                                            
    l_bolRequiredDistIDsEnteredOtherPay = ValidateDistributionIDRequiredDistIDsEntered(l_rsEmployeeDistIDInfo, v_rsEmployeeWGData, v_rsOtherPayEntries, r_rsBRVs, True)
                                            ' Might need this v_rsEmployeeWGData
            
    l_bolValidDistribution = l_bolTimeEntryValidationSuccessful _
                                        And l_bolOtherPayEntryValidationSuccessful _
                                        And l_bolWorkOrderValidationSuccessful _
                                        And l_bolRequiredDistIDsEntered _
                                        And l_bolRequiredDistIDsEnteredOtherPay
                                        
    If l_bolValidDistribution Then
        'Remove Distribution from Other Pay that is not a TTA
        ValidateDistributionIDRemoveOtherPay v_rsOtherPayEntries, r_rsBRVs
    End If
    ValidateDistributionIdsFromTimeEntries2 = l_bolValidDistribution
    v_rsTimeEntries.Filter = ""
                                       
    Exit Function

eh:
    Set r_rsBRVs = Nothing
    Set v_rsTimeEntries = Nothing
    Set v_rsOtherPayEntries = Nothing
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIdsFromTimeEntries2", Err.Number, , , , True

End Function

Private Function GetEmployeeDistIdInfo( _
            ByVal v_varEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As ADODB.Recordset
            
    Dim l_obj As boCommonLib.cboEmployee
    
    Dim l_varEmployees As Variant
    Dim l_rsEmployees As ADODB.Recordset
    Dim l_aryEmployees() As String
    Dim l_intI As Integer
    
'This function assumes that at least one employee was passed
'Allows an array of employees or a recordset of employees
' to be passed
    
    On Error GoTo eh

    If v_varEmployees Is Nothing Then
        Set GetEmployeeDistIdInfo = Nothing
        Exit Function
    End If
    If IsArray(v_varEmployees) Then
        l_varEmployees = v_varEmployees
        Set l_rsEmployees = Nothing
    Else
'Recordset was passed
        Set l_rsEmployees = v_varEmployees
        ReDim l_aryEmployees(l_rsEmployees.RecordCount - 1) As String
        l_rsEmployees.MoveFirst
        l_intI = 0
        While Not l_rsEmployees.EOF
            l_aryEmployees(l_intI) = str(l_rsEmployees!EMP_NUM)
            l_rsEmployees.MoveNext
            l_intI = l_intI + 1
        Wend
        l_varEmployees = l_aryEmployees
    End If
    
    
    Set l_obj = CreateObject("boCommonLib.cboEmployee")
    
    Set GetEmployeeDistIdInfo = l_obj.GetValidateDistIdInfoByChargeDate(l_varEmployees, v_datStartDate, v_datEndDate, l_rsEmployees)
    Set l_obj = Nothing
    
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeDistIdInfo"
            
End Function

Private Sub ValidateDistIDBuildAccountList( _
                ByRef r_strDistributionIDEdits As Variant, _
                ByRef r_intNumofEdits As Integer, _
                ByRef r_lngDistIDRecordPosition As Variant, _
                ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset, _
                ByVal v_rsEmployeeWGData As ADODB.Recordset, _
                ByVal v_rsTimeEntries As ADODB.Recordset, _
                ByRef v_lngStartRecordPosition As Long, _
                ByVal v_rsJobTaskSteps As ADODB.Recordset, _
                Optional v_blnOtherPay As Boolean = False)
                

    Dim l_strDistIDEdit As String
    Dim l_strDistID As String
    Dim l_lngEmployee As Long
    Dim l_datChargeDate As Date
    Dim l_strCompanyCode As String
    Dim l_strPayrollDept As String
    Dim l_lngWGNum As Long
    Dim l_strDistType As String
    Dim l_varDistID As Variant
    Dim l_varDistributionId As Variant
    Dim l_strDistrDept As String
'    Dim l_strTransType As String
    Dim l_blnIsPOET As Boolean
    Dim l_sJobNbr As String, l_sJobTaskNbr As String, l_sDBId As String, l_sFilter As String, l_sProjectNumber As String
    
    On Error GoTo ValidateDistIDBuildAccountList_EH
    

    If RecordsExist(v_rsTimeEntries) Then
    Else
        Exit Sub
    End If
    TimeEntriesFilterAndSort v_rsTimeEntries, v_blnOtherPay
    If v_rsTimeEntries.EOF Then
        Exit Sub
    End If
    v_rsTimeEntries.AbsolutePosition = v_lngStartRecordPosition
    While Not v_rsTimeEntries.EOF And r_intNumofEdits < 221
        l_lngEmployee = v_rsTimeEntries!KEY_EMP_NUM
        l_datChargeDate = v_rsTimeEntries!KEY_CHARGE_DATE
        l_strDistID = GetDistributionID(v_rsTimeEntries)
        l_strDistType = GetDistributionType(v_rsTimeEntries)
        '//l_strTransType = GetTranType39(l_strDistID) '//Mid(l_strDistID, 16, 1) '//check for POET and skip validation
        l_blnIsPOET = IsPOET39(l_strDistID)
        '//If (l_strDistType = "A" And l_strTransType <> "P") Or l_strDistType = "S" Or l_strDistType = "J" Or l_strDistType = "M" Then
        '//If l_strDistType = "A" Or l_strDistType = "S" Or l_strDistType = "J" Or l_strDistType = "M" Then
        '//If ((l_strDistType = "A" Or l_strDistType = "S") And Not (IsOTTCEffective And IsOTTCTransitionActive)) Or _
            '//l_strDistType = "J" Or l_strDistType = "M" Then
        
        '//don't validate here if Old GL Accounting and in the transistion period, where the accounting will be converted else where
        '//Validate if:
        '//Poet
        '//GL (not Poet) and not IsOTTCTransitionActive (not converted else where)
        '//Maximo/Emprv
        If (l_strDistType = "A" And l_blnIsPOET) Or _
            ((l_strDistType = "A" Or l_strDistType = "S") And Not l_blnIsPOET And Not (IsOTTCEffective And IsOTTCTransitionActive)) Or _
            (l_strDistType = "J" Or l_strDistType = "M") Then
            l_strCompanyCode = ValidateDistIDGetCompanyCode(l_lngEmployee, l_datChargeDate, v_rsEmployeeDistIDInfo)
            l_strPayrollDept = ValidateDistIDGetPayrollDept(l_lngEmployee, l_datChargeDate, v_rsEmployeeDistIDInfo)
            l_lngWGNum = GetChargedWG(l_lngEmployee, l_datChargeDate, v_rsEmployeeWGData, v_rsTimeEntries)
            l_varDistID = ParseStringDistIDToAnArray(l_strDistID)
            If l_strPayrollDept = "086" Then
                If IsNull(v_rsTimeEntries!DISTR_DEPT_NUM) Or IsEmpty(v_rsTimeEntries!DISTR_DEPT_NUM) Then
                    l_strDistrDept = ""
                Else
                    l_strDistrDept = v_rsTimeEntries!DISTR_DEPT_NUM
                End If
            Else
                l_strDistrDept = ""
            End If
            l_varDistributionId = FillInMissingPieces(l_varDistID, l_strPayrollDept, l_lngWGNum, l_strDistType, l_strDistrDept)
            l_strDistIDEdit = ParseArrayDistIDToAString(l_varDistributionId)
            '//use PI for ReasonForEdit to validate service requests, this will allow the stored proc/ws to validate against PowerPlant
            '//l_strDistIDEdit = IIf(l_strDistType = "S", "PI", "FE") & l_strCompanyCode & l_strDistIDEdit
            
            If (IsPOETActive And IsOTTCEffective) Then
                Select Case l_strDistType
                Case "M" '//Maximo
                    '//use MX here, will actually use FE for reason for edit and MX for system id
                    '//.NET will lookup the project if empty since there is a limit of 5 chars, new projects will be 8 or 10
                    l_strDistIDEdit = "MX" & l_strCompanyCode & l_strDistIDEdit
                Case "J" '//EMPRV
                    l_strDistIDEdit = "EM" & l_strCompanyCode & l_strDistIDEdit
                Case "S" '//Serivce Request
                    '//use SV here, will actually use FE for reason for edit and TR for system id with default accounting except project and activity
                    l_strDistIDEdit = "SV" & l_strCompanyCode & l_strDistIDEdit
                Case "A" '//Full accounting (GL39) or POET (tran type = "P")
                    l_strDistIDEdit = "FE" & l_strCompanyCode & l_strDistIDEdit
                Case Else
                    l_strDistIDEdit = "FE" & l_strCompanyCode & l_strDistIDEdit
                End Select
            Else
                l_strDistIDEdit = "FE" & l_strCompanyCode & l_strDistIDEdit
            End If
            
            l_strDistIDEdit = l_strDistIDEdit + Format(l_datChargeDate, "mm/dd/yyyy")
            
            l_sProjectNumber = Space(10)
            '//try to get 8 char project id's from client job task steps (from picklists)
            If l_strDistType = "M" Or l_strDistType = "J" Then
                If FieldExists("VALIDATION_TYPE", v_rsJobTaskSteps) Then
                    l_sJobNbr = v_rsTimeEntries!JOB_NBR
                    l_sJobTaskNbr = v_rsTimeEntries!JOBTASK_NBR
                    l_sDBId = v_rsTimeEntries!EMPRV_DATABASE_ID
                    
                    l_sFilter = "JOB_NBR = '" & Trim(l_sJobNbr) & "'" & _
                                " AND JOBTASK_NBR = '" & Trim(l_sJobTaskNbr) & "'" & _
                                " AND EMPRV_DATABASE_ID = '" & Trim(l_sDBId) & "'" & _
                                " AND VALIDATION_TYPE = 'P'"
                    v_rsJobTaskSteps.Filter = l_sFilter
                    '//If Not v_rsJobTaskSteps.BOF And Not v_rsJobTaskSteps.EOF Then
                    If RecordsExist(v_rsJobTaskSteps) Then '//used from mADOExtensions which does not remove filter like timekeeper does
                        v_rsJobTaskSteps.MoveFirst
                        l_sProjectNumber = v_rsJobTaskSteps!PROJECT_NUMBER
                    End If
                End If
            End If
            
            l_strDistIDEdit = l_strDistIDEdit + PadRight(l_sProjectNumber, 10)
            
            r_strDistributionIDEdits(r_intNumofEdits) = l_strDistIDEdit
            r_lngDistIDRecordPosition(r_intNumofEdits) = v_rsTimeEntries.AbsolutePosition
            r_intNumofEdits = r_intNumofEdits + 1
        End If
        v_rsTimeEntries.MoveNext
    Wend
    If v_rsTimeEntries.EOF Then
        v_lngStartRecordPosition = 1
    Else
        v_lngStartRecordPosition = v_rsTimeEntries.AbsolutePosition
    End If
    v_rsTimeEntries.MoveFirst
    
'    If IsOTTCCutoverOutageActive And r_intNumofEdits = 0 Then
'        r_intNumofEdits = 1 '//force failure
'
'        If RecordsExist(v_rsTimeEntries) Then
'            v_rsTimeEntries.MoveFirst
'            l_strDistID = GetDistributionID(v_rsTimeEntries)
'            r_strDistributionIDEdits(r_intNumofEdits - 1) = l_strDistID
'            r_lngDistIDRecordPosition(r_intNumofEdits - 1) = v_rsTimeEntries.AbsolutePosition
'        End If
'    End If
    
    Exit Sub
    
ValidateDistIDBuildAccountList_EH:
        TrsRaiseError m_strCMODULENAME, "ValidateDistIDBuildAccountList"

End Sub

'build/format array list of up to 221 elements composed of the following: ReasonForEdit, CorpCode, DistributionId
Private Sub ValidateDistIDBuildAccountList2( _
                ByVal v_strReasonForEdit As String, _
                ByVal v_strCompanyCode As String, _
                ByVal v_strDistributionIDs As Variant, _
                ByRef r_strDistributionIDEdits As Variant, _
                ByRef r_intNumofEdits As Integer)
    Dim l_strDistID As String
    Dim l_strDistIDEdit As String
    Dim i As Integer
    Dim l_blnFMCDefault As Boolean, l_blnACTVDefault As Boolean
    Dim l_strReasonForEdit As String
    On Error GoTo eh
    
    i = LBound(v_strDistributionIDs)
    Do While i <= UBound(v_strDistributionIDs) And r_intNumofEdits < 221
        l_strDistID = v_strDistributionIDs(i)
        '//l_varDistID = ParseStringDistIDToAnArray(l_strDistID)
        '//l_strDistIDEdit = ParseArrayDistIDToAString(l_varDistributionId)
        '//some distributions will not have the activity, it will be filled in during time collection
        l_strDistID = GetDistIdForValidation(l_strDistID, l_blnFMCDefault, l_blnACTVDefault)
              
        '//use PI for ReasonForEdit to validate service requests, this will allow the stored proc/ws to validate against PowerPlant
        '//HACK: override client that will be sending "FE" for all requests
'        Dim varDistId As Variant
'        varDistId = ParseStringDistIDToAnArray(l_strDistID)
'        If l_strReasonForEdit = "FE" And IsServiceRequest(varDistId) Then
'            l_strReasonForEdit = "PI"
'        Else
            l_strReasonForEdit = v_strReasonForEdit
'        End If
        
        l_strDistIDEdit = PadRight(l_strReasonForEdit, 2) & PadRight(v_strCompanyCode, 3) & PadRight(l_strDistID, 39) & Format(Now, "mm/dd/yyyy")
        r_strDistributionIDEdits(i) = l_strDistIDEdit
        
        r_intNumofEdits = r_intNumofEdits + 1
        i = i + 1
    Loop
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateDistIDBuildAccountList2"
End Sub

Private Sub ValidateDistIDBuildWorkRequestList( _
                ByRef r_strDistributionIDEdits As Variant, _
                ByRef r_intNumofEdits As Integer, _
                ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset, _
                ByVal v_rsEmployeeWGData As ADODB.Recordset, _
                ByVal v_rsTimeEntries As ADODB.Recordset)
                

    Dim l_strDistIDEdit As String
    Dim l_strDistID As String
    Dim l_lngEmployee As Long
    Dim l_datChargeDate As Date
    Dim l_strCompanyCode As String
    Dim l_strDOJMDistrict As String
    Dim l_strDOJMHeadquarters As String
    Dim l_lngChargedWGNum As Long
    Dim l_strDistType As String
    
    On Error GoTo ValidateDistIDBuildWorkRequestList_EH
    
    If RecordsExist(v_rsTimeEntries) Then
    Else
        Exit Sub
    End If
    TimeEntriesFilterAndSort v_rsTimeEntries
    If v_rsTimeEntries.EOF Then
        Exit Sub
    End If
    v_rsTimeEntries.MoveFirst
    While Not v_rsTimeEntries.EOF
        l_lngEmployee = v_rsTimeEntries!KEY_EMP_NUM
        l_datChargeDate = v_rsTimeEntries!KEY_CHARGE_DATE
        l_strDistID = Trim(GetDistributionID(v_rsTimeEntries))
        l_strDistType = GetDistributionType(v_rsTimeEntries)
        If l_strDistType = "W" Then
            l_lngChargedWGNum = GetChargedWG(l_lngEmployee, l_datChargeDate, v_rsEmployeeWGData, v_rsTimeEntries)
            GetDOJMDistrictAndHeadQuarter l_lngChargedWGNum, l_strDOJMDistrict, l_strDOJMHeadquarters
            l_strDistIDEdit = l_strDistID & l_strDOJMHeadquarters
            If IsStringInList(l_strDistIDEdit, r_strDistributionIDEdits) Then
            Else
                r_strDistributionIDEdits(r_intNumofEdits) = l_strDistIDEdit
                r_intNumofEdits = r_intNumofEdits + 1
            End If
        End If
        v_rsTimeEntries.MoveNext
    Wend
    v_rsTimeEntries.MoveFirst
    
    Exit Sub
    
ValidateDistIDBuildWorkRequestList_EH:
            TrsRaiseError m_strCMODULENAME, "ValidateDistIDBuildWorkRequestList"

End Sub
'Private Sub ValidateDistIDBuildJobTaskList( _
'                ByRef r_strDistributionIDEdits As Variant, _
'                ByRef r_intNumofEdits As Integer, _
'                ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset, _
'                ByVal v_rsEmployeeWGData As ADODB.Recordset, _
'                ByVal v_rsTimeEntries As ADODB.Recordset)
                

'    Dim l_strDistIDEdit As String
'    Dim l_strDistID As String
'    Dim l_lngEmployee As Long
'    Dim l_datChargeDate As Date
'    Dim l_strCompanyCode As String
'    Dim l_lngChargedWGNum As Long
'    Dim l_strDistType As String
    
'    On Error GoTo ValidateDistIDBuildJobTaskList_EH
    
'    If RecordsExist(v_rsTimeEntries) Then
'    Else
'        Exit Sub
'    End If
'    TimeEntriesFilterAndSort v_rsTimeEntries
'    If v_rsTimeEntries.EOF Then
'        Exit Sub
'    End If
'    v_rsTimeEntries.MoveFirst
'    While Not v_rsTimeEntries.EOF
'        l_lngEmployee = v_rsTimeEntries!KEY_EMP_NUM
'        l_datChargeDate = v_rsTimeEntries!KEY_CHARGE_DATE
'        l_strDistID = Trim(GetDistributionID(v_rsTimeEntries))
'        l_strDistType = GetDistributionType(v_rsTimeEntries)
'        If l_strDistType = "J" Then
'            l_lngChargedWGNum = GetChargedWG(l_lngEmployee, l_datChargeDate, v_rsEmployeeWGData, v_rsTimeEntries)
'            l_strDistIDEdit = l_strDistID
'            If IsStringInList(l_strDistIDEdit, r_strDistributionIDEdits) Then
'            Else
'                r_strDistributionIDEdits(r_intNumofEdits) = l_strDistIDEdit
'                r_intNumofEdits = r_intNumofEdits + 1
'            End If
'        End If
'        v_rsTimeEntries.MoveNext
'    Wend
'    v_rsTimeEntries.MoveFirst
    
'    Exit Sub
    
'ValidateDistIDBuildJobTaskList_EH:
'            TrsRaiseError m_strCMODULENAME, "ValidateDistIDBuildJobTaskList"

'End Sub



Private Function ValidateDistIDGetCompanyCode( _
            ByVal v_lngEmployee As Long, _
            ByVal v_datChargeDate As Date, _
            ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset _
            ) As String
            
    On Error GoTo ValidateDistIDGetCompanyCode_EH
            
    v_rsEmployeeDistIDInfo.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmployee) _
                & " and CHARGE_DATE = " & str(v_datChargeDate))
                
    If v_rsEmployeeDistIDInfo.EOF Then
        ValidateDistIDGetCompanyCode = "   "
    Else
        ValidateDistIDGetCompanyCode = v_rsEmployeeDistIDInfo!CMPY_CD
    End If
    
    v_rsEmployeeDistIDInfo.Filter = ""
    Exit Function
    
ValidateDistIDGetCompanyCode_EH:
    TrsRaiseError m_strCMODULENAME, "ValidateDistIDGetCompanyCode"
            
End Function
Private Function ValidateDistIDGetPayrollDept( _
            ByVal v_lngEmployee As Long, _
            ByVal v_datChargeDate As Date, _
            ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset _
            ) As String
            
    On Error GoTo ValidateDistIDGetPayrollDept_EH
    
    v_rsEmployeeDistIDInfo.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmployee) _
                & " and CHARGE_DATE = " & str(v_datChargeDate))
                
    If v_rsEmployeeDistIDInfo.EOF Then
        ValidateDistIDGetPayrollDept = "   "
    Else
        ValidateDistIDGetPayrollDept = v_rsEmployeeDistIDInfo!DEPT_NUM
    End If
    
    v_rsEmployeeDistIDInfo.Filter = ""
    Exit Function
    
ValidateDistIDGetPayrollDept_EH:
    TrsRaiseError m_strCMODULENAME, "ValidateDistIDGetPayrollDept"
    
            
End Function

Private Function ValidateDistributionIDsAccountRequests( _
            ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset, _
            ByVal v_rsEmployeeWGData As ADODB.Recordset, _
            ByRef r_rsTimeEntries As ADODB.Recordset, _
            ByRef r_rsBRVs As ADODB.Recordset, _
            ByVal v_rsJobTaskSteps As ADODB.Recordset, _
            Optional v_blnOtherPay As Boolean = False _
            ) As Boolean
            
    Dim l_strDistributionIDEdits(221) As String '* 44
    Dim l_intNumberOfEdits As Integer
    Dim l_lngDistIDRecordPosition(221) As Long
    Dim l_intI As Integer
    Dim l_strReturnCodes As Variant ' (221)String * 1
    Dim l_strReturnMsgs As Variant ' (221)String * 50
    Dim l_strHighlightFields As Variant '(221) String * 13
    Dim l_strDistributionIDsReturned As Variant '(221) String * 39
    Dim l_bolValidationSuccessful As Boolean
    Dim l_lngStartRecordPosition As Long
    Dim l_bolError  As Boolean
    
    On Error GoTo eh
                                         
    For l_intI = 0 To 221
        l_strDistributionIDEdits(l_intI) = Space(64) '//Space(44) added effective or transaction date, added PROJECT_NUMBER for project id's (larger than 5 chars)
        l_lngDistIDRecordPosition(l_intI) = 0
    Next l_intI
    l_intNumberOfEdits = 0
    l_lngStartRecordPosition = 1
    
    If IsOTTCBypassAccountValidationActive() Then '//sysparms 55
        If r_rsBRVs Is Nothing Then
            Set r_rsBRVs = NewErrorset()
        End If
        If RecordsExist(r_rsTimeEntries) Then
            r_rsTimeEntries.MoveFirst
        
            l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 10000, "BI", "Account Validation has been turned off", v_blnOtherPay)
        End If
        
        ValidateDistributionIDsAccountRequests = True
        Exit Function
    End If
    
    If IsOTTCCutoverOutageActive() Then '//sysparms 52
        If r_rsBRVs Is Nothing Then
            Set r_rsBRVs = NewErrorset()
        End If
        If RecordsExist(r_rsTimeEntries) Then
            r_rsTimeEntries.MoveFirst
        
            l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 30000, "BE", "Unable to save accounting because OTTC Cutover/Outage is in progress.", v_blnOtherPay)
        End If
        
        ValidateDistributionIDsAccountRequests = False
        Exit Function
    End If
    
'    If IsOTTCEffective() And Not IsOTTCTransitionActive() Then '//sysparms 53 & 54
'        If RecordsExist(r_rsTimeEntries) Then
'            r_rsTimeEntries.MoveFirst
'
'            r_rsTimeEntries.Filter = ""
'            If r_rsBRVs Is Nothing Then
'                Set r_rsBRVs = NewErrorset()
'            End If
'            If RecordsExist(r_rsTimeEntries) Then
'                r_rsTimeEntries.MoveFirst
'
'                l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 30000, "BE", "Unable to save/convert GL accounting, please remove and select POET instead.", v_blnOtherPay)
'            End If
'
'            ValidateDistributionIDsAccountRequests = False
'            Exit Function
'        End If
'    End If
    
    ValidateDistIDBuildAccountList l_strDistributionIDEdits, l_intNumberOfEdits, l_lngDistIDRecordPosition, v_rsEmployeeDistIDInfo, v_rsEmployeeWGData, r_rsTimeEntries, l_lngStartRecordPosition, v_rsJobTaskSteps, v_blnOtherPay
    
    If l_intNumberOfEdits = 0 Then '//And Not IsOTTCCutoverOutageActive Then '//need to let it fail in ValidateGLAccounting for failing ALL account validation
        ValidateDistributionIDsAccountRequests = True
        Exit Function
    End If
    l_bolValidationSuccessful = ValidateGLAccounting(l_strDistributionIDEdits, l_intNumberOfEdits, l_strDistributionIDsReturned, l_strReturnCodes, l_strReturnMsgs, l_strHighlightFields)
    
    
    If l_bolValidationSuccessful Then
        ValidateDistributionIDsUpdateDistID _
                r_rsTimeEntries, l_intNumberOfEdits, _
                l_strDistributionIDsReturned, _
                l_lngDistIDRecordPosition, r_rsBRVs, v_blnOtherPay
        ValidateDistributionIDsAccountRequests = True
        Exit Function
    End If
    
'Need to generate BRV's for each error
'Accounting that was defaulted but failed validation is still
'valid.
'If this requirement changed, just change the last line of
' this function to return false

    l_bolValidationSuccessful = _
    ValidateDistributionIDsGenerateBRVs( _
        r_rsTimeEntries, l_intNumberOfEdits, _
        l_lngDistIDRecordPosition, _
        l_strDistributionIDsReturned, _
        l_strReturnCodes, l_strReturnMsgs, _
        v_rsEmployeeDistIDInfo, r_rsBRVs, v_blnOtherPay)
        
    ValidateDistributionIDsAccountRequests = l_bolValidationSuccessful


    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIDsAccountRequests"
End Function

'//-----for validating from the client without timeentry dependencies
Public Function ValidateDistributionIDs( _
                ByVal v_strReasonForEdit As String, _
                ByVal v_strCompanyCode As String, _
                ByVal v_strDistributionIDs As Variant, _
                ByRef r_strDistributionIDsReturned As Variant, _
                ByRef r_strReturnCodes As Variant, _
                ByRef r_strReturnMsgs As Variant, _
                ByRef r_strHighlightFields As Variant, _
                ByRef r_rsBRVs As ADODB.Recordset _
            ) As Boolean
            
    Dim l_strDistributionIDEdits(221) As String '* 44
    Dim l_intNumberOfEdits As Integer
    Dim l_intI As Integer
    Dim l_bolValidationSuccessful As Boolean
                       
    On Error GoTo eh
                                         
    For l_intI = 0 To 221
        l_strDistributionIDEdits(l_intI) = Space(54) '// Space(44)
    Next l_intI
    
    l_intNumberOfEdits = 0
    If Not IsOTTCBypassAccountValidationActive() Then '//sysparms 55
        ValidateDistIDBuildAccountList2 v_strReasonForEdit, v_strCompanyCode, v_strDistributionIDs, l_strDistributionIDEdits, l_intNumberOfEdits
    End If
    If l_intNumberOfEdits = 0 Then
        ValidateDistributionIDs = True
        Exit Function
    End If
    l_bolValidationSuccessful = ValidateGLAccounting(l_strDistributionIDEdits, l_intNumberOfEdits, r_strDistributionIDsReturned, r_strReturnCodes, r_strReturnMsgs, r_strHighlightFields)
         
'Need to generate BRV's for each error
'Accounting that was defaulted but failed validation is still
'valid.
'If this requirement changed, just change the last line of
' this function to return false
    
    '//If l_bolValidationSuccessful Then '//if the web api call fails, the return msg and code arrays will not get populated and cause a "Subscript out of range" error
        l_bolValidationSuccessful = ValidateDistributionIDsGenerateBRVs2(v_strReasonForEdit, v_strCompanyCode, _
            v_strDistributionIDs, l_intNumberOfEdits, _
            r_strDistributionIDsReturned, _
            r_strReturnCodes, r_strReturnMsgs, _
            r_rsBRVs) '//, l_bolValidationSuccessful)
    '//End If
    
    '//HACK: Due to a bug on the client which is not associating the return code to an image, resize return codes back to a length of 222
    ReDim Preserve r_strReturnCodes(221)
    ValidateDistributionIDs = l_bolValidationSuccessful

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIDs"
End Function

Private Function GetDistIdForValidation( _
    ByVal v_strDistribution As String, _
    ByRef r_blnFMCDefault As Boolean, _
    ByRef r_blnACTVDefault As Boolean _
    ) As String
    
    Dim l_strFMC As String
    Dim l_strACTV As String
    Dim l_strDept As String
'    Dim l_obj As dbDepartment.cdDepartment

    On Error GoTo eh
    
    r_blnFMCDefault = False
    r_blnACTVDefault = False
    l_strFMC = Mid$(v_strDistribution, 10, 3)
    l_strACTV = Mid$(v_strDistribution, 24, 4)

'    If Trim(l_strACTV) = "" Then
'        l_strACTV = "ADMN"
'        r_blnACTVDefault = True
'    End If

'    If Trim(l_strFMC) = "" Then
'        l_strDept = GetDepartment()
'        Set l_obj = CreateObject("dbDepartment.cdDepartment")
'        l_strFMC = l_obj.GetBudgetNumber(l_strDept)
'        Set l_obj = Nothing
'        r_blnFMCDefault = True
'    End If

    GetDistIdForValidation = Mid$(v_strDistribution, 1, 9) & l_strFMC & _
        Mid$(v_strDistribution, 13, 11) & l_strACTV & Mid$(v_strDistribution, 28, 12)

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistIdForValidation"
End Function

Private Function GetDistIdAfterValidation( _
    ByVal v_strValidatedWGDistribution As String, _
    ByVal v_blnFMCDefault As Boolean, _
    ByVal v_blnACTVDefault As Boolean _
    ) As String
    
    Dim l_strDistIdAfterValidation As String
    
    On Error GoTo eh
    
    l_strDistIdAfterValidation = v_strValidatedWGDistribution
    If v_blnFMCDefault Then
        l_strDistIdAfterValidation = Mid$(l_strDistIdAfterValidation, 1, 9) & _
            Space(3) & Mid$(l_strDistIdAfterValidation, 13, 27)
    End If
    
    If v_blnACTVDefault Then
        l_strDistIdAfterValidation = Mid$(l_strDistIdAfterValidation, 1, 23) & _
            Space(4) & Mid$(l_strDistIdAfterValidation, 28, 12)
    End If
    
    GetDistIdAfterValidation = l_strDistIdAfterValidation
    
                
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistIdAfterValidation"
End Function
'//-----

Private Function ValidateDistributionIDsWorkRequests( _
            ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset, _
            ByVal v_rsEmployeeWGData As ADODB.Recordset, _
            ByVal v_rsTimeEntries As ADODB.Recordset, _
            ByRef r_rsBRVs As ADODB.Recordset _
            ) As Boolean
            
    Dim l_strWorkRequestEdits(200) As String
    Dim l_intNumberOfEdits As Integer
    Dim l_intI As Integer
    Dim l_bol As Boolean
    Dim l_rsDOJMWorkRequests As ADODB.Recordset
    Dim l_strReturnMsg As String
    Dim l_blnValidateWorkRequestsSuccess As Boolean
    Dim l_strDistIDDOJMDistrict As String
    Dim l_strDistIDDOJMWorkRequest As String
    Dim l_lngWGNum As Long
    Dim l_strWGDOJMDistrict As String
    Dim l_strWGDOJMHeadquarters As String
    Dim l_strFilter As String
    Dim l_bolError As Boolean
    Dim l_strDistID As String
    
            
    On Error GoTo eh
                                         
    l_blnValidateWorkRequestsSuccess = True
    For l_intI = 0 To 200
        l_strWorkRequestEdits(l_intI) = ""
    Next l_intI
    l_intNumberOfEdits = 0
    ValidateDistIDBuildWorkRequestList l_strWorkRequestEdits, l_intNumberOfEdits, v_rsEmployeeDistIDInfo, v_rsEmployeeWGData, v_rsTimeEntries
    If l_intNumberOfEdits = 0 Then
        ValidateDistributionIDsWorkRequests = True
        Exit Function
    End If
    
    
    Set l_rsDOJMWorkRequests = ValidateDistributionIDsGetWorkRequests(l_strWorkRequestEdits, l_intNumberOfEdits)
    ValidateDistributionIDsWorkRequests = True
'For each time entry that is a work request
'must check to see if the work request was returned
    v_rsTimeEntries.Filter = MDAC25FilterFix("DISTRIBUTION_TYPE = 'W' AND RECORD_STATUS <> 'D'")
    v_rsTimeEntries.Sort = "EMP_NUM, CHARGE_DATE"
    v_rsTimeEntries.MoveFirst
    While Not v_rsTimeEntries.EOF
        l_rsDOJMWorkRequests.Filter = ""
        l_strDistID = GetDistributionID(v_rsTimeEntries)
        l_strDistIDDOJMDistrict = Left$(l_strDistID, 4)
        l_strDistIDDOJMWorkRequest = Mid$(l_strDistID, 8, 6)
        l_lngWGNum = GetChargedWG(v_rsTimeEntries!KEY_EMP_NUM, v_rsTimeEntries!KEY_CHARGE_DATE, v_rsEmployeeWGData, v_rsTimeEntries)
        If ValidateDOJMDistrict(l_strDistIDDOJMDistrict, l_lngWGNum, l_strReturnMsg) Then
            GetDOJMDistrictAndHeadQuarter l_lngWGNum, l_strWGDOJMDistrict, l_strWGDOJMHeadquarters
            l_strFilter = "CDE_DIST = '" & l_strDistIDDOJMDistrict & "'" _
                       & " AND CDE_HDQTR_CRW LIKE '" & l_strWGDOJMHeadquarters & "%'" _
                       & " AND CDE_WR = '" & l_strDistIDDOJMWorkRequest & "'"
            l_rsDOJMWorkRequests.Filter = MDAC25FilterFix(l_strFilter)
            If l_rsDOJMWorkRequests.EOF Then
            'Write BRV Work Order Invalid for DOJM District and Headquarters
                l_bolError = WriteBRV(r_rsBRVs, v_rsTimeEntries, 30000, "BE", "Distribution is invalid --" & "Work Order Invalid for DOJM District and Headquarters")
                l_blnValidateWorkRequestsSuccess = False
            End If
        Else
        'Write BRV with Message
            l_bolError = WriteBRV(r_rsBRVs, v_rsTimeEntries, 30000, "BE", "Distribution is invalid --" & l_strReturnMsg)
            l_blnValidateWorkRequestsSuccess = False
        End If
        v_rsTimeEntries.MoveNext
    Wend
    
    ValidateDistributionIDsWorkRequests = l_blnValidateWorkRequestsSuccess
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIDsWorkRequests"
               
            
End Function
            
Private Function ValidateDistributionIDRequiredDistIDsEntered( _
            ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset, _
            ByVal v_rsEmployeeWGData As ADODB.Recordset, _
            ByVal v_rsTimeEntries As ADODB.Recordset, _
            ByRef r_rsBRVs As ADODB.Recordset, _
            Optional v_blnOtherPay As Boolean = False _
            ) As Boolean
            
            
    Dim l_lngEmployee As Long
    Dim l_datChargeDate As Date
    Dim l_strDistID As String
    Dim l_strDistIDType As String
    Dim l_strDefaultDistID As String
    Dim l_strAccountType As String
    Dim l_bolRequiredDistIDsEntered As Boolean
    Dim l_bolError As Boolean
    Dim l_lngChargedWG As Long
    Dim l_strDistDepartment As String
    Dim l_strExceptionCode As String
    Dim l_strOverExcCode As String
    Dim l_strBankCode As String
    Dim l_strPayrollDept As String
    Dim l_blnDefaultAcct As Boolean
    
    On Error GoTo ValidateDistributionIDRequiredDistIDsEntered_EH
    
    l_bolRequiredDistIDsEntered = True
    If RecordsExist(v_rsTimeEntries) Then
    Else
        ValidateDistributionIDRequiredDistIDsEntered = True
        Exit Function
    End If
    TimeEntriesFilterAndSort v_rsTimeEntries, v_blnOtherPay
    If v_rsTimeEntries.EOF Then
        ValidateDistributionIDRequiredDistIDsEntered = True
        Exit Function
    End If
    While Not v_rsTimeEntries.EOF
        l_lngEmployee = v_rsTimeEntries!KEY_EMP_NUM
        l_datChargeDate = v_rsTimeEntries!KEY_CHARGE_DATE
        l_strDistID = Trim(GetDistributionID(v_rsTimeEntries))
        If IsNull(v_rsTimeEntries!DISTRIBUTION_TYPE) Then
            l_strDistIDType = DetermineDistributionType(l_strDistID)
        Else
            l_strDistIDType = v_rsTimeEntries!DISTRIBUTION_TYPE
        End If
        l_lngChargedWG = GetChargedWG(l_lngEmployee, l_datChargeDate, v_rsEmployeeWGData, v_rsTimeEntries)
        l_strPayrollDept = ValidateDistIDGetPayrollDept(l_lngEmployee, l_datChargeDate, v_rsEmployeeDistIDInfo)
        If IsNull(v_rsTimeEntries!DISTR_DEPT_NUM) Or IsEmpty(v_rsTimeEntries!DISTR_DEPT_NUM) Then
            l_strDistDepartment = DetermineDistributionDepartment(l_strPayrollDept, l_lngChargedWG)
        Else
            If l_strPayrollDept = "086" Then
                l_strDistDepartment = v_rsTimeEntries!DISTR_DEPT_NUM
            Else
                l_strDistDepartment = DetermineDistributionDepartment(l_strPayrollDept, l_lngChargedWG)
            End If
        End If
        l_strAccountType = GetAcctType(l_strDistDepartment)
    'Can not have an override code without an exception code
    'Assuming straight Time ST
        If IsEmpty(v_rsTimeEntries!EXCEPTION_CD) Or IsNull(v_rsTimeEntries!EXCEPTION_CD) Then
            l_strExceptionCode = "ST"
        Else
            l_strExceptionCode = v_rsTimeEntries!EXCEPTION_CD
        End If
        
        l_strBankCode = GetExceptionBankCode(l_strExceptionCode)
        If v_blnOtherPay Then
            l_strOverExcCode = ""
        Else
            If IsEmpty(v_rsTimeEntries!OVER_EXC_CD) Or IsNull(v_rsTimeEntries!OVER_EXC_CD) Then
                l_strOverExcCode = ""
            Else
                l_strOverExcCode = v_rsTimeEntries!OVER_EXC_CD
            End If
        End If
        Select Case v_blnOtherPay
            Case True
                If Trim(l_strDistID) = "" Then
                    l_bolRequiredDistIDsEntered = False
                    l_bolError = WriteBRV(r_rsBRVs, v_rsTimeEntries, 30000, "BE", "Distribution is required.", v_blnOtherPay)
                Else
                    If l_strDistIDType <> "A" Then
                        l_bolRequiredDistIDsEntered = False
                        l_bolError = WriteBRV(r_rsBRVs, v_rsTimeEntries, 30000, "BE", "Full 39-digit Distribution is required.", v_blnOtherPay)
                    End If
                End If
            Case False
            '
            ' The routine below overrides to account type time for time entries prior to the MKS interface to TRIS
            '
                l_strAccountType = MKSDetermineAcctType(l_strDistDepartment, l_lngEmployee, l_datChargeDate)
                If RequiresAccounting(l_strExceptionCode, l_strDistDepartment, l_strAccountType, l_strBankCode, l_strOverExcCode, l_strPayrollDept, l_blnDefaultAcct) _
                    And Trim(l_strDistIDType) = "" Then
                    l_bolRequiredDistIDsEntered = False
                    l_bolError = WriteBRV(r_rsBRVs, v_rsTimeEntries, 30000, "BE", "Distribution is required.", v_blnOtherPay)
                ' Write BRV
                Else
                    If l_blnDefaultAcct And Not IsDOJMWorkGroupPrivate(l_lngChargedWG) Then
                        v_rsTimeEntries!DISTR_DEPT_NUM = l_strPayrollDept
                    Else
                        v_rsTimeEntries!DISTR_DEPT_NUM = l_strDistDepartment
                    End If
                End If
            End Select
        v_rsTimeEntries.MoveNext
    Wend
    ValidateDistributionIDRequiredDistIDsEntered = l_bolRequiredDistIDsEntered
    Exit Function
    
ValidateDistributionIDRequiredDistIDsEntered_EH:
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIDRequiredDistIDsEntered"
            
End Function

'//moved the call to the Gl account validation web service down to the db level, as it is also called from busWorkGroup
'//could not add reference from busWorkGroup back to busDistributionMgr (cyclic reference)
'//busWorkGroup.cbsWorkGroupTrans is transaction (COM+) and busDistributionMgr (not configured in COM+ because of performance issues)
'//busDistributionMgrProxy (COM+) but mainly there to be used by the client only
Private Function ValidateGLAccounting( _
            ByVal v_strDistributionIDEdits As Variant, _
            ByVal v_intNumberOfEdits As Integer, _
            ByRef r_strDistributionIDsReturned As Variant, _
            ByRef r_strReturnCodes As Variant, _
            ByRef r_strReturnMsgs As Variant, _
            ByRef r_strHighlightFields As Variant _
            ) As Boolean
            
    Dim l_obj As dbDistributionMgr.cdDistributionMgr
    
    Dim l_bolDBSuccess As Boolean
    
    On Error GoTo eh
    
    
    Set l_obj = CreateObject("dbDistributionMgr.cdDistributionMgr")
    
'    If IsOTTCCutoverOutageActive Then '//sysparms 52
'        l_bolDBSuccess = False
'    Else
        l_bolDBSuccess = l_obj.ValidateGLAccounting(v_strDistributionIDEdits, v_intNumberOfEdits, r_strDistributionIDsReturned, r_strReturnCodes, r_strReturnMsgs, r_strHighlightFields)
'    End If
    
    If l_bolDBSuccess Then
    Else
        Dim l_strEmptyArray1(221) As String
        Dim l_strEmptyArray2(221) As String
        Dim l_strEmptyArray3(221) As String
        Dim l_strEmptyArray4(221) As String
        
        r_strDistributionIDsReturned = l_strEmptyArray1
        r_strReturnCodes = l_strEmptyArray2
        r_strReturnMsgs = l_strEmptyArray3
        r_strHighlightFields = l_strEmptyArray4
    
'        If IsOTTCCutoverOutageActive() And v_intNumberOfEdits = 0 Then '//create dummy request in order to fail it
'            v_intNumberOfEdits = 1
'        End If
        
        Dim i As Integer
        For i = 0 To v_intNumberOfEdits - 1
            '//see ValidateDistributionIDsGenerateBRVs2 checking for returned dist id = "", for setting the brv being returned to client
'            If IsOTTCCutoverOutageActive Then '//sysparms 52
'                r_strReturnMsgs(i) = "Unable to save accounting because OTTC Cutover/Outage is in progress."
'            Else
                If IsPOETActive And IsOTTCEffective Then
                    r_strReturnMsgs(i) = "Please try again. OTTC account validation unavailable at this time."
                Else
                    r_strReturnMsgs(i) = "Please try again. GL Distribution validation unavailable at this time."
                End If
'            End If
            
            r_strReturnCodes(i) = "1" '//error
            '//r_strDistributionIDsReturned(i) = Mid(v_strDistributionIDEdits(i), 6) '//ignore (e.g. FEAMS)
            r_strDistributionIDsReturned(i) = "" '//let the brvs be created in ValidateDistributionIDsGenerateBRVs and ValidateDistributionIDsGenerateBRVs2
            r_strHighlightFields(i) = ""
        Next i
'        r_strReturnMsgs(0) = "Call to Stored Procedure Failed"
'        ValidateGLAccounting = False
'        Err.Raise trErrbusValidateGLAccoutingStoredProcedureFailed
'        Exit Function
    End If
    
    If Trim(r_strDistributionIDsReturned(0)) = "" Then
'        r_strReturnMsgs(0) = "Please try again. GL Distribution validation unavailable at this time."
    End If
    
    ValidateGLAccounting = ValidateGLAccountingAllEditsSuccesful(r_strReturnCodes, v_intNumberOfEdits)

    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "ValidateGLAccounting"
            
End Function

'Private Function ValidateGLAccountingWS( _
'            ByVal v_strDistributionIDEdits As Variant, _
'            ByVal v_intNumberOfEdits As Integer, _
'            ByRef r_strDistributionIDsReturned As Variant, _
'            ByRef r_strReturnCodes As Variant, _
'            ByRef r_strReturnMsgs As Variant, _
'            ByRef r_strHighlightFields As Variant _
'            ) As Boolean
'
'    Dim l_obj As Ameren_TRIS_Server_COMIntegration.DistributionNumber
'
'    Dim l_bolDBSuccess As Boolean
'
'    On Error GoTo eh
'
'    Set l_obj = New Ameren_TRIS_Server_COMIntegration.DistributionNumber
'
'    '//if the return value of false previously meant that the call to the stored procedure failed
'    '//now a value of false means that the validation failed, let any failures result in errors/exceptions
'    ReDim Preserve v_strDistributionIDEdits(v_intNumberOfEdits - 1) '//only send what is needed to web service instead of a fixed 222 (0-221) array
'    l_bolDBSuccess = l_obj.ValidateDistributionNumbers(v_strDistributionIDEdits, v_intNumberOfEdits, r_strDistributionIDsReturned, r_strReturnCodes, r_strReturnMsgs, r_strHighlightFields)
'
'    If l_bolDBSuccess Then
'    Else
'        r_strReturnMsgs(0) = "Please try again. GL Account Validation Web Service unavailable at this time."
'
''        r_strReturnCodes(0) = 2
''        r_strReturnMsgs(0) = "Call to Account Validation Web Service failed"
''        ValidateGLAccounting = False
''        Err.Raise trErrbusValidateGLAccoutingStoredProcedureFailed
''        Exit Function
'    End If
'
''    If Trim(r_strDistributionIDsReturned(0)) = "" Then
''        r_strReturnMsgs(0) = "Please try again. GL Account Validation Web Service unavailable at this time."
''    End If
'
'    ValidateGLAccountingWS = ValidateGLAccountingAllEditsSuccesful(r_strReturnCodes, v_intNumberOfEdits)
'
'    Set l_obj = Nothing
'    Exit Function
'eh:
'    Set l_obj = Nothing
'    TrsRaiseError m_strCMODULENAME, "ValidateGLAccountingWS"
'
'End Function

Private Function ValidateDistIDGetAccountType( _
            ByVal v_lngEmployee As Long, _
            ByVal v_datChargeDate As Date, _
            ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset _
            ) As String
            
    On Error GoTo ValidateDistIDGetAccountType_EH
            
    v_rsEmployeeDistIDInfo.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmployee) _
                & " and CHARGE_DATE = " & str(v_datChargeDate))
                
    If v_rsEmployeeDistIDInfo.EOF Then
        ValidateDistIDGetAccountType = ""
    Else
        ValidateDistIDGetAccountType = UCase(v_rsEmployeeDistIDInfo!ACCT_TYPE)
    End If
    
    v_rsEmployeeDistIDInfo.Filter = ""
    Exit Function
    
ValidateDistIDGetAccountType_EH:
    TrsRaiseError m_strCMODULENAME, "ValidateDistIDGetAccountType"
    
End Function

Private Function ValidateGLAccountingAllEditsSuccesful( _
            ByVal v_strReturnCodes As Variant, _
            ByVal v_intNumberOfEdits As Integer _
            ) As Boolean
                      
    Dim l_intI As Integer
    Dim l_bolEditSuccess As Boolean
    
    On Error GoTo ValidateGLAccountingAllEditsSuccesful_EH
    
    l_bolEditSuccess = True
    For l_intI = 0 To v_intNumberOfEdits - 1
        If Trim(v_strReturnCodes(l_intI)) <> "0" Then
            l_bolEditSuccess = False
            Exit For
        End If
    Next l_intI
    
    ValidateGLAccountingAllEditsSuccesful = l_bolEditSuccess
    Exit Function
    
ValidateGLAccountingAllEditsSuccesful_EH:
    TrsRaiseError m_strCMODULENAME, "ValidateDistIDGetAccountType"
               
End Function


Private Sub ValidateDistributionIDsUpdateDistID( _
                ByRef r_rsTimeEntries As ADODB.Recordset, _
                ByVal v_intNumberOfEdits As Integer, _
                ByVal v_strDistributionIDsReturned As Variant, _
                ByVal v_lngDistIDRecordPosition As Variant, _
                ByRef r_rsBRVs As ADODB.Recordset, _
                Optional v_blnOtherPay As Boolean = False)
                
                
    Dim l_intI As Integer
    Dim l_strOrigDistrID As String
    Dim l_strRetDistrID As String
    Dim l_bolError As Boolean
    
    On Error GoTo ValidateDistributionIDsUpdateDistID_EH
    
    TimeEntriesFilterAndSort r_rsTimeEntries, v_blnOtherPay
    If r_rsTimeEntries.EOF Then
        Exit Sub
    End If
    For l_intI = 0 To v_intNumberOfEdits - 1
        r_rsTimeEntries.AbsolutePosition = v_lngDistIDRecordPosition(l_intI)
        l_strOrigDistrID = GetDistributionID(r_rsTimeEntries)
        l_strRetDistrID = v_strDistributionIDsReturned(l_intI)
        If l_strOrigDistrID = l_strRetDistrID Then
'Account Edit Module did not change the distribution id
        Else
            r_rsTimeEntries!DISTRIBUTION_ID = l_strRetDistrID
'            l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 20000, "BW", "System changed the Distribution ID", v_blnOtherPay)
        End If
    Next l_intI
    
    Exit Sub
    
ValidateDistributionIDsUpdateDistID_EH:
    
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIDsUpdateDistID"
                
End Sub


Private Function ValidateDistributionIDsGenerateBRVs( _
            ByRef r_rsTimeEntries As ADODB.Recordset, _
            ByVal v_intNumberOfEdits As Integer, _
            ByVal v_lngDistIDRecordPosition As Variant, _
            ByVal v_strDistributionIDsReturned As Variant, _
            ByVal v_strReturnCodes As Variant, _
            ByVal v_strReturnMsgs As Variant, _
            ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset, _
            ByRef r_rsBRVs As ADODB.Recordset, _
            Optional v_blnOtherPay As Boolean = False _
            ) As Boolean
                
                
    Dim l_intI As Integer
    Dim l_strOrigDistrID As String
    Dim l_bolError As Boolean
    Dim l_strDefaultAccounting As String
    Dim l_bolNoBusinessErrors As Boolean
    Dim l_strDistID As String
    Dim l_strReturnMsg As String
    
    On Error GoTo BRV_EH:
    
    
    l_bolNoBusinessErrors = True
    If r_rsBRVs Is Nothing Then
        Set r_rsBRVs = NewErrorset()
    End If
    TimeEntriesFilterAndSort r_rsTimeEntries, v_blnOtherPay
    If r_rsTimeEntries.EOF Then
        ValidateDistributionIDsGenerateBRVs = True
        Exit Function
    End If
    For l_intI = 0 To v_intNumberOfEdits - 1
        r_rsTimeEntries.AbsolutePosition = v_lngDistIDRecordPosition(l_intI)
        If v_strReturnCodes(l_intI) = "0" Then
'The distribution id passed validation
            l_strDistID = GetDistributionID(r_rsTimeEntries)
            If l_strDistID <> v_strDistributionIDsReturned(l_intI) Then
    'Distribution ID Changed
                r_rsTimeEntries!DISTRIBUTION_ID = v_strDistributionIDsReturned(l_intI)
'                l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 20000, "BW", "System changed the Distribution ID", v_blnOtherPay)
            End If
'Add a BRV stating that the distribution id has changed
        Else
            l_strDefaultAccounting = ValidateDistributionIDsGetDefaultAccount(r_rsTimeEntries, v_rsEmployeeDistIDInfo, v_blnOtherPay)
            If l_strDefaultAccounting <> "" And l_strDefaultAccounting <> "N/A" Then
            'Do not write BRV for default accounting that errored
            'But write the warning if the DISTRIBUTION_ID changed
                l_strDistID = GetDistributionID(r_rsTimeEntries)
                If l_strDistID <> v_strDistributionIDsReturned(l_intI) Then
            'Distribution ID Changed
                    If Trim(v_strDistributionIDsReturned(l_intI)) = "" Then
                         l_bolNoBusinessErrors = False
                        l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 30000, "BE", "Please try again. GL Distribution validation unavailable at this time.", v_blnOtherPay)
                    Else
                        r_rsTimeEntries!DISTRIBUTION_ID = v_strDistributionIDsReturned(l_intI)
'                       l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 20000, "BW", "System changed the Distribution ID", v_blnOtherPay)
                    End If
                End If
            Else
                l_bolNoBusinessErrors = False
                l_strReturnMsg = v_strReturnMsgs(l_intI)
                If l_strReturnMsg = "INVLD FOR CORP -                           ACTV   " Then
                    l_strReturnMsg = "Activity is missing or invalid."
                End If
                If Trim(v_strDistributionIDsReturned(l_intI)) = "" Then
                    If Trim(l_strReturnMsg) = "" Then
                        l_strReturnMsg = "Please try again. GL Distribution validation unavailable at this time."
                    End If
                    l_bolNoBusinessErrors = False
                    l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 30000, "BE", l_strReturnMsg, v_blnOtherPay)
                Else
                    If InStr(1, l_strReturnMsg, m_strINVALIDFORCORP) > 0 Then
                        Dim l_strCompanyCode
                        Dim EmpNum As Long
                        Dim ChargeDate As Date

                        If RecordsExist(r_rsTimeEntries) Then
                            EmpNum = r_rsTimeEntries.Fields("KEY_EMP_NUM").Value
                            ChargeDate = r_rsTimeEntries.Fields("KEY_CHARGE_DATE").Value
                            l_strCompanyCode = ValidateDistIDGetCompanyCode(EmpNum, ChargeDate, v_rsEmployeeDistIDInfo)
                            l_strReturnMsg = Replace(l_strReturnMsg, m_strINVALIDFORCORP, m_strINVALIDFORCORP & " (" & l_strCompanyCode & ")")
                        End If
                    End If
                    
                    '//if a service request and the activity is missing, then add suggestion to check activity
                    Dim varDistId As Variant
                    varDistId = ParseStringDistIDToAnArray(v_strDistributionIDsReturned(l_intI))
                    If Trim(varDistId(9)) = "" Then 'activity is missing
                        If IsServiceRequest(varDistId) Then
                            l_strReturnMsg = Trim(l_strReturnMsg) + ".  Activity is required for Service Requests."
                        End If
                    End If
                    
                    l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 30000, "BE", "Distribution is invalid --" & l_strReturnMsg, v_blnOtherPay)
                End If
            End If
        End If
    Next l_intI
    
'    If IsOTTCBypassAccountValidationActive() Then
'        l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 10000, "BI", "Account Validation has been turned off", v_blnOtherPay)
'    End If
    
    ValidateDistributionIDsGenerateBRVs = l_bolNoBusinessErrors
                           
    Exit Function
    
BRV_EH:
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIDsGenerateBRVs"
End Function

Private Function ValidateDistributionIDsGenerateBRVs2( _
            ByVal v_strReasonForEdit As String, _
            ByVal v_strCompanyCode As String, _
            ByVal v_strDistributionIDs As Variant, _
            ByVal v_intNumberOfEdits As Integer, _
            ByVal v_strDistributionIDsReturned As Variant, _
            ByVal v_strReturnCodes As Variant, _
            ByVal v_strReturnMsgs As Variant, _
            ByRef r_rsBRVs As ADODB.Recordset _
            ) As Boolean
    '//ByVal v_CallToValidationAPISuccess As Boolean
    Dim l_intI As Integer
    Dim l_strOrigDistrID As String
    Dim l_bolError As Boolean
    Dim l_strDefaultAccounting As String
    Dim l_bolNoBusinessErrors As Boolean
    Dim l_strDistID As String
    Dim l_strReturnMsg As String
    Dim l_lngBRVType As Long
    Dim l_strBRVFieldNames As String
    Dim l_strErrorBasedOnType As String
    Dim l_dbDistMgr As dbDistributionMgr.cdDistributionMgr
    
    On Error GoTo eh:
    
    l_bolNoBusinessErrors = True
    If r_rsBRVs Is Nothing Then
        Set r_rsBRVs = CreateNewBusinessErrorSet() 'for workgroup maintenance ' NewErrorset() 'for timecard
    End If

    Select Case v_strReasonForEdit
    Case "FE" '//, "PI"
        l_strErrorBasedOnType = "Distribution is invalid --"
        l_lngBRVType = TrType_WGM_ACCT
        l_strBRVFieldNames = "WG_NUM;DISTRIBUTION_ID"
    Case "AV"
        l_strErrorBasedOnType = "Activity is invalid --"
        l_lngBRVType = TrType_WGM_ACTIV
        l_strBRVFieldNames = "WG_NUM;ACTIVITY_CD"
    Case Else
        l_strErrorBasedOnType = "Distribution is invalid --"
        l_lngBRVType = TrType_WGM_ACCT
        l_strBRVFieldNames = "WG_NUM;DISTRIBUTION_ID"
    End Select

    '//if the web api call fails, the return msg and code arrays will not get populated and cause a "Subscript out of range" error
'    If v_CallToValidationAPISuccess Then
    For l_intI = 0 To v_intNumberOfEdits - 1
        'The distribution id passed validation
        If v_strReturnCodes(l_intI) = "0" Then
            l_strDistID = v_strDistributionIDs(l_intI)
            'Distribution ID Changed
            If l_strDistID <> v_strDistributionIDsReturned(l_intI) Then
'                l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 20000, "BW", "System changed the Distribution ID", v_blnOtherPay)
            End If
        'Add a BRV stating that the distribution id has changed
        Else
'            l_strDefaultAccounting = ValidateDistributionIDsGetDefaultAccount(r_rsTimeEntries, v_rsEmployeeDistIDInfo, v_blnOtherPay)
'            If l_strDefaultAccounting <> "" And l_strDefaultAccounting <> "N/A" Then
'            'Do not write BRV for default accounting that errored
'            'But write the warning if the DISTRIBUTION_ID changed
'                l_strDistID = v_strDistributionIDs(l_intI)
'                If l_strDistID <> v_strDistributionIDsReturned(l_intI) Then
'            'Distribution ID Changed
'                    If Trim(v_strDistributionIDsReturned(l_intI)) = "" Then
'                         l_bolNoBusinessErrors = False
'                        l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 30000, "BE", "Please try again. GL Distribution validation unavailable at this time.", v_blnOtherPay)
'                    Else
'                        r_rsTimeEntries!DISTRIBUTION_ID = v_strDistributionIDsReturned(l_intI)
''                       l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 20000, "BW", "System changed the Distribution ID", v_blnOtherPay)
'                    End If
'                End If
'            Else
                l_bolNoBusinessErrors = False
                l_strReturnMsg = v_strReturnMsgs(l_intI)
                If l_strReturnMsg = "INVLD FOR CORP -                           ACTV   " Then
                    l_strReturnMsg = "Activity is missing or invalid."
                End If
                If Trim(v_strDistributionIDsReturned(l_intI)) = "" Then
                    If Trim(l_strReturnMsg) = "" Then
                        l_strReturnMsg = "Please try again. GL Distribution validation unavailable at this time."
                    End If
                    l_bolNoBusinessErrors = False
                    AppendBusinessError r_rsBRVs, 30000, "BE", l_strReturnMsg, _
                        m_strCMODULENAME, _
                        l_lngBRVType, , , l_strBRVFieldNames, _
                        "", _
                        "", _
                        "0", _
                        "0", _
                        IIf(l_lngBRVType = TrType_WGM_ACCT, v_strDistributionIDsReturned(l_intI), ""), _
                        IIf(l_lngBRVType = TrType_WGM_ACTIV, v_strDistributionIDsReturned(l_intI), "")
                    '//l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 30000, "BE", "Please try again. GL Distribution validation unavailable at this time.", v_blnOtherPay)
                Else
                    If InStr(1, l_strReturnMsg, m_strINVALIDFORCORP) > 0 Then
                        l_strReturnMsg = Replace(l_strReturnMsg, m_strINVALIDFORCORP, m_strINVALIDFORCORP & " (" & v_strCompanyCode & ")")
                    End If
                    
                    '//if a service request and the activity is missing, then add suggestion to check activity
                    Dim varDistId As Variant
                    varDistId = ParseStringDistIDToAnArray(v_strDistributionIDsReturned(l_intI))
                    If Trim(varDistId(9)) = "" Then 'activity is missing
                        If IsServiceRequest(varDistId) Then
                            l_strReturnMsg = Trim(l_strReturnMsg) + ".  Activity is required for Service Requests."
                        End If
                    End If
                    
                    '//If Mid(v_strDistributionIDsReturned(l_intI), 16, 1) = "P" Then  '//POET
                    
                    If IsPOET39(v_strDistributionIDsReturned(l_intI)) Then '//POET
                        If IsPOETActive And IsOTTCEffective Then
                            Dim rsPOET As ADODB.Recordset
                            Set rsPOET = GetPOETMasterFromDistributionId(CStr(v_strDistributionIDsReturned(l_intI)))
                            Dim ProjectNumber As String, ProjectName As String
                            Dim TaskNumber As String, TaskName As String
                            
                            If RecordsExist(rsPOET) Then
                                rsPOET.MoveFirst
                                ProjectNumber = Trim(PadRight(rsPOET.Fields("Project_Number").Value, 39))
                                ProjectName = Trim(PadRight(rsPOET.Fields("Project_Name").Value, 39))
                                TaskNumber = Trim(PadRight(rsPOET.Fields("Task_Number").Value, 39))
                                TaskName = Trim(PadRight(rsPOET.Fields("Task_Name").Value, 39))
                            End If
                            
                            AppendBusinessError r_rsBRVs, 30000, "BE", "POET is invalid --" & l_strReturnMsg, _
                                m_strCMODULENAME, _
                                l_lngBRVType, , , l_strBRVFieldNames, _
                                "", _
                                "", _
                                "0", _
                                "0", _
                                IIf(l_lngBRVType = TrType_WGM_ACCT, v_strDistributionIDsReturned(l_intI), ""), _
                                ProjectNumber, TaskNumber, ProjectName, TaskName
                        Else '//copied below because there is NO AndAlso for the IsPOETActive call used above
                            AppendBusinessError r_rsBRVs, 30000, "BE", l_strErrorBasedOnType & l_strReturnMsg, _
                                m_strCMODULENAME, _
                                l_lngBRVType, , , l_strBRVFieldNames, _
                                "", _
                                "", _
                                "0", _
                                "0", _
                                IIf(l_lngBRVType = TrType_WGM_ACCT, v_strDistributionIDsReturned(l_intI), ""), _
                                IIf(l_lngBRVType = TrType_WGM_ACTIV, v_strDistributionIDsReturned(l_intI), "")
                        End If
                    Else
                        AppendBusinessError r_rsBRVs, 30000, "BE", l_strErrorBasedOnType & l_strReturnMsg, _
                            m_strCMODULENAME, _
                            l_lngBRVType, , , l_strBRVFieldNames, _
                            "", _
                            "", _
                            "0", _
                            "0", _
                            IIf(l_lngBRVType = TrType_WGM_ACCT, v_strDistributionIDsReturned(l_intI), ""), _
                            IIf(l_lngBRVType = TrType_WGM_ACTIV, v_strDistributionIDsReturned(l_intI), "")
                    End If
                    
                    '//l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 30000, "BE", "Distribution is invalid --" & l_strReturnMsg, v_blnOtherPay)
                End If
'            End If
        End If
    Next l_intI
'    Else
'        l_bolNoBusinessErrors = False
'        Dim errmsg As String
'
'        If IsPOETActive Then
'            errmsg = "Please try again. OTTC account validation unavailable at this time."
'        Else
'            errmsg = "Please try again. GL Distribution validation unavailable at this time."
'        End If
'        AppendBusinessError r_rsBRVs, 30000, "BE", errmsg, _
'            m_strCMODULENAME, _
'            l_lngBRVType, , , l_strBRVFieldNames, _
'            "", _
'            "", _
'            "0", _
'            "0", _
'            "", _
'            ""
'    End If
    
    ValidateDistributionIDsGenerateBRVs2 = l_bolNoBusinessErrors
                           
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIDsGenerateBRVs2"
End Function

'//copied
Private Function IsPOETActive() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim UseOTTC As String
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsPOETActive = False
    
    UseOTTC = UCase(dbEnv.GetSysParmsValue(50, 1))
    
    If UseOTTC = "Y" Or UseOTTC = "TRUE" Or UseOTTC = "1" Then
        IsPOETActive = True
    End If
        
    Set dbEnv = Nothing

    Exit Function

eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsPOETActive", , , , DescArgsFrom(UseOTTC), True
End Function

Private Function IsOTTCCutoverOutageActive() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim OTTCCutoverOutage As String
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsOTTCCutoverOutageActive = False
    
    OTTCCutoverOutage = UCase(dbEnv.GetSysParmsValue(52, 1))
    
    If OTTCCutoverOutage = "Y" Or OTTCCutoverOutage = "TRUE" Or OTTCCutoverOutage = "1" Then
        IsOTTCCutoverOutageActive = True
    End If
        
    Set dbEnv = Nothing

    Exit Function

eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsOTTCCutoverOutageActive", , , , DescArgsFrom(OTTCCutoverOutage), True
End Function

'//actual date to start using POET distributions after cutover
Private Function IsOTTCEffective() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim sEffDate As String
    Dim EffDateTime
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsOTTCEffective = False
    
    sEffDate = dbEnv.GetSysParmsValue(53, 1)
    
    EffDateTime = DateTimeFromDB2Timestamp(sEffDate)
    
    IsOTTCEffective = Now > EffDateTime
        
    Set dbEnv = Nothing

    Exit Function
eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsOTTCEffective", , , , DescArgsFrom(sEffDate), True
End Function

Private Function IsOTTCTransitionActive() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim OTTCTransition As String
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsOTTCTransitionActive = False
    
    OTTCTransition = UCase(dbEnv.GetSysParmsValue(54, 1))
    
    If OTTCTransition = "Y" Or OTTCTransition = "TRUE" Or OTTCTransition = "1" Then
        IsOTTCTransitionActive = True
    End If
        
    Set dbEnv = Nothing

    Exit Function
eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsOTTCTransitionActive", , , , DescArgsFrom(OTTCTransition), True
End Function

Private Function IsOTTCBypassAccountValidationActive() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim OTTCBypassAccountValidation As String
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsOTTCBypassAccountValidationActive = False
    
    OTTCBypassAccountValidation = UCase(dbEnv.GetSysParmsValue(55, 1))
    
    If OTTCBypassAccountValidation = "Y" Or OTTCBypassAccountValidation = "TRUE" Or OTTCBypassAccountValidation = "1" Then
        IsOTTCBypassAccountValidationActive = True
    End If
        
    Set dbEnv = Nothing

    Exit Function
eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsOTTCBypassAccountValidation", , , , DescArgsFrom(OTTCBypassAccountValidation), True
End Function

Private Function GetTRISIdsFromDistributionId(DistributionId As String, ByRef TRISProjectId As String, ByRef TRISTaskId As String) As String

    TRISProjectId = ""
    TRISTaskId = ""

    If IsPOET39(DistributionId) Then '//Mid(DistributionId, 16, 1) = "P" Then '//POET
        TRISProjectId = GetProject39(DistributionId) '//Mid(DistributionId, 17, 5)
        TRISTaskId = GetActivity39(DistributionId)  '//Mid(DistributionId, 24, 4)
    End If
   
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetTRISIdsFromDistributionId"
End Function

Private Function GetPOETMasterFromDistributionId(DistributionId As String) As ADODB.Recordset
    Dim TRISProjectId As String, TRISTaskId As String
    Dim rs As ADODB.Recordset
    Dim db As dbDistributionMgr.cdDistributionMgr
    
    GetTRISIdsFromDistributionId DistributionId, TRISProjectId, TRISTaskId
    
    Set db = CreateObject("dbDistributionMgr.cdDistributionMgr")
    
    Set rs = db.GetPOETMasterFromTRISIds(TRISProjectId, TRISTaskId)
    Set GetPOETMasterFromDistributionId = rs
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
eh:
    Set rs = Nothing
    Set db = Nothing

    TrsRaiseError m_strCMODULENAME, "GetPOETMasterFromDistributionId"
End Function


Private Function ValidateDistributionIDsGetWorkRequests( _
            ByVal v_strWorkRequestEdits As Variant, _
            ByVal v_intNumberOfEdits As Integer _
            ) As ADODB.Recordset
    '  Is the WorkRequest Valid
    
    Dim l_obj As dbDistributionMgr.cdDistributionMgr
    Dim l_blnSuccess As Boolean
    Dim l_aryDOJMDistricts() As String
    Dim l_aryDOJMWorkRequests() As String
    Dim l_aryDOJMHeadquarters() As String
    Dim l_intI As Integer
    Dim l_rsDOJMWorkRequests As ADODB.Recordset
        
On Error GoTo eh
    
    ReDim l_aryDOJMDistricts(v_intNumberOfEdits - 1) As String
    ReDim l_aryDOJMWorkRequests(v_intNumberOfEdits - 1) As String
    ReDim l_aryDOJMHeadquarters(v_intNumberOfEdits - 1) As String
    For l_intI = 0 To v_intNumberOfEdits - 1
        l_aryDOJMDistricts(l_intI) = Left$(v_strWorkRequestEdits(l_intI), 4)
        l_aryDOJMWorkRequests(l_intI) = Mid$(v_strWorkRequestEdits(l_intI), 8, 6)
        l_aryDOJMHeadquarters(l_intI) = Mid$(v_strWorkRequestEdits(l_intI), 14, 6)
    Next l_intI
    
    
    Set l_obj = CreateObject("dbDistributionMgr.cdDistributionMgr")

    Set l_rsDOJMWorkRequests = l_obj.ValidateWorkRequestsGetWR(l_aryDOJMDistricts, l_aryDOJMWorkRequests, l_aryDOJMHeadquarters)
    
    Set ValidateDistributionIDsGetWorkRequests = l_rsDOJMWorkRequests
    
    Set l_obj = Nothing
    
Exit Function

eh:

    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIDsGetWorkRequests"

End Function

'obsolete
Public Function ValidateDistributionId( _
            ByVal v_strDistributionId As Variant, _
            ByRef r_strReturnMsg As String, _
            ByVal v_strCorporateCode As String, _
            ByVal v_strDistributionType As String, _
            ByVal v_strPayrollDept As String, _
            ByVal v_lngChargeWGNum As Long, _
            ByRef r_strDistributionIDReturned As Variant, _
            ByRef r_strHighlightFields As String, _
            Optional v_strReasonForEdit As String = "FE" _
            ) As Boolean
            
'This Function is for client side validation
' of one distribution id.
'It will be passed as an array and each part of the
' code block is its on element.
'For a DOJM Work request the Code block is
'       Positions   Array Element           Description
'       1-4         v_strDistributionID (0) DOJM District
'       5                                   Blank space
'       6-7         v_strDistributionID(1)  "WR" indicates that it is a work request
'       8-13        v_strDistributionID(2)  DOJM Work request
'The DOJM Work request is validated against the
'   DOJM WG, DOJM District, and DOJM Headquarters
'For GL Accounting the Code block is
'       Positions   Array Element           Description
'           1-1     v_strDistributionID(0)  (U) Utility
'           2-3     v_strDistributionID(1)  (BD) Business Division
'           4-6     v_strDistributionID(2)  (MAJ)FERC Major
'           7-9     v_strDistributionID(3)  (MIN)FERC Minor
'           10-12   v_strDistributionID(4)  (FMC) Fund Mgt. Center
'           13-15   v_strDistributionID(5)  (RMC) Resource Mgt. Center
'           16-16   v_strDistributionID(6)  (T) Trans Type
'           17-21   v_strDistributionID(7)  (PROJT) Project
'           22-23   v_strDistributionID(8)  (PR) Product
'           24-27   v_strDistributionID(9)  (ACTV) Activity
'           28-29   v_strDistributionID(10) (RT) Resource Type
'           30-39   v_strDistributionID(11) (FEEDER_REF) Memo

'For Service Request Accounting the Code block is the
' same as above, but only the following fields can be entered
'    FMC, Project, and Activity
    
On Error GoTo eh:
    
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionId", trErrObsoleteMethod, , , , True
    ValidateDistributionId = False
    Exit Function
eh:
    r_strDistributionIDReturned = Null
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionId", , , , , True
End Function


Private Function DetermineDistributionType( _
            ByVal v_strDistributionId As Variant _
            ) As String
            
On Error GoTo eh
            
    If Not IsArray(v_strDistributionId) Then
        If Trim(v_strDistributionId) = "" Then
            DetermineDistributionType = " "
            Exit Function
        End If
        If Mid$(v_strDistributionId, 6, 2) = "WR" And Trim(Mid$(v_strDistributionId, 5, 1)) = "" Then
            DetermineDistributionType = "W"
            Exit Function
        End If
        If IsServiceRequest(v_strDistributionId) Then
            DetermineDistributionType = "S"
            Exit Function
        End If
        ' There is no true way to tell if it is full accounting
        ' or the description of the exception code
        ' I am assuming that all exception code descriptions
        ' have at least 4 characters
        If IsNumeric(Mid$(v_strDistributionId, 4, 3)) Or Trim(Mid$(v_strDistributionId, 4, 3)) = "" Then
            DetermineDistributionType = "A"
        Else
            DetermineDistributionType = " "
        End If
    Else
        If Trim(v_strDistributionId(1)) = "WR" Then
            DetermineDistributionType = "W"
            Exit Function
        End If
        If IsServiceRequest(v_strDistributionId) Then
            DetermineDistributionType = "S"
            Exit Function
        End If
        If IsNumeric(v_strDistributionId(2)) Or Trim(v_strDistributionId(2)) = "" Then
            DetermineDistributionType = "A"
        Else
            DetermineDistributionType = " "
        End If
    End If
    
Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineDistributionType"
            
End Function


Private Sub GetDOJMDistrictAndHeadQuarter( _
            ByVal v_lngWgNum As Long, _
            ByRef r_strDOJMDistrict As String, _
            ByRef r_strDOJMHeadQuarters As String)
            
    '  Is the WorkRequest Valid
    
    Dim l_obj As dbWorkGroup.cdWorkGroup
    Dim l_intPosition As Integer
        
On Error GoTo eh

    If IsStringInList2(str(v_lngWgNum), l_varWorkGroupDOJMCache, l_lngWorkGroupDOJMCacheCount, l_intPosition) Then
        r_strDOJMDistrict = l_varWorkGroupDOJMCache(l_intPosition, 1)
        r_strDOJMHeadQuarters = l_varWorkGroupDOJMCache(l_intPosition, 2)
        Exit Sub
    End If
    
    
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    r_strDOJMDistrict = l_obj.DOJMDistrict(v_lngWgNum)
    r_strDOJMHeadQuarters = l_obj.DOJMHeadquarters(v_lngWgNum)
    Set l_obj = Nothing
'''''    If l_lngWorkGroupDOJMCacheCount > 0 Then
'''''        ReDim Preserve l_varWorkGroupDOJMCache(l_lngWorkGroupDOJMCacheCount, 2) As String
'''''    Else
'''''        ReDim l_varWorkGroupDOJMCache(l_lngWorkGroupDOJMCacheCount, 2) As String
'''''    End If
    l_varWorkGroupDOJMCache(l_lngWorkGroupDOJMCacheCount, 0) = str(v_lngWgNum)
    l_varWorkGroupDOJMCache(l_lngWorkGroupDOJMCacheCount, 1) = r_strDOJMDistrict
    l_varWorkGroupDOJMCache(l_lngWorkGroupDOJMCacheCount, 2) = r_strDOJMHeadQuarters
    l_lngWorkGroupDOJMCacheCount = l_lngWorkGroupDOJMCacheCount + 1
    
Exit Sub

eh:

    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDOJMDistrictAndHeadQuarter"

End Sub


Private Function FillInMissingPieces( _
            ByVal v_strDistributionId As Variant, _
            ByVal v_strPayrollDept As String, _
            ByVal v_lngChargeWGNum As Long, _
            ByVal v_strDistributionType As String, _
            ByVal v_strDistrDept As String, _
            Optional v_blnException As Boolean = False _
            ) As Variant
            
'Release 08/10/2001 -- JFB
'  Added the DistrDept parameter
'  It will be blank unless the payroll dept is "086"
'  It will also be blank when defaulting accounting for an exception
            
    Dim l_strUtil As String ' Utility
    Dim l_strBD As String   ' Budject
    Dim l_strFMC As String  ' Funding Management Center
    Dim l_strRMC As String  ' Resource Management Center
    Dim l_strPR As String   ' Product
    Dim l_strRT As String   ' Resource Type
    Dim l_strReturnDistId(11) As String
    Dim l_intI As Integer
    
    On Error GoTo FillInMissingPieces_EH

    If Not v_blnException Then
        If Trim(v_strDistrDept) = "" Then
            v_strDistrDept = DetermineDistributionDepartment(v_strPayrollDept, v_lngChargeWGNum)
        End If
    Else
        v_strDistrDept = v_strPayrollDept
    End If
    
    For l_intI = 0 To 11
        l_strReturnDistId(l_intI) = v_strDistributionId(l_intI)
    Next l_intI
    
'On a service request FMC is the only missing piece
' that can be filled in
    If Not (IsPOETActive And IsOTTCEffective) Then
        If v_strDistributionType = "S" Then
            If Trim(l_strReturnDistId(4)) = "" Then
                l_strFMC = GetBudgetNum(v_strDistrDept)
    '            l_strFMC = GetFMC(v_lngChargeWGNum, v_strPayrollDept)
                l_strReturnDistId(4) = l_strFMC
            End If
            FillInMissingPieces = l_strReturnDistId
            Exit Function
        End If
        
        If v_blnException Then
            l_strReturnDistId(0) = Space(1) 'U Utility
            l_strReturnDistId(1) = Space(2) 'BD Business Division
            l_strReturnDistId(4) = Space(3) 'FMC Funding Management Center
            l_strReturnDistId(5) = Space(3) 'RMC Resource Management Center
            l_strReturnDistId(8) = Space(2) 'PR Product
            l_strReturnDistId(10) = Space(2) 'RT Resource Type
        End If
        
     ' If Business Division is empty fill it in
        l_strBD = l_strReturnDistId(1)
        If l_strBD = Space(2) Then
            l_strBD = GetAcctDiv(v_strPayrollDept)
            l_strReturnDistId(1) = l_strBD
        End If
    'If the utility and product are both empty fill them in
        If Trim(l_strReturnDistId(0)) = "" And Trim(l_strReturnDistId(8)) = "" Then
    '        If IsNumeric(l_strBD) Then
                Select Case Val(l_strBD)
                    Case 80 To 88
                        l_strUtil = "2"
                        l_strPR = "02"
                    Case Else
                        l_strUtil = "1"
                        l_strPR = "01"
                End Select
                l_strReturnDistId(0) = l_strUtil
                l_strReturnDistId(8) = l_strPR
    '        End If
        End If
    'Check FMC
        If Trim(l_strReturnDistId(4)) = "" Then
            If v_blnException Then
                l_strFMC = GetBudgetNum(v_strPayrollDept)
            Else
                l_strFMC = GetFMC(v_lngChargeWGNum, v_strPayrollDept)
                l_strReturnDistId(4) = l_strFMC
            End If
        End If
    'Check RMC
    'Always determine RMC via John S. instructions
    '    If Trim(l_strReturnDistId(5)) = "" Then
            l_strRMC = GetBudgetNum(v_strPayrollDept)
            l_strReturnDistId(5) = l_strRMC
    '    End If
    Else '//only for POET
        If v_blnException Then
            l_strReturnDistId(10) = Space(2) 'RT Resource Type
        End If
    End If

'Check RT
'Always determine RT via John S. instructions
'    If Trim(l_strReturnDistId(10)) = "" Then
        l_strRT = GetResourceType(v_strPayrollDept)
        l_strReturnDistId(10) = l_strRT
'    End If
    
    FillInMissingPieces = l_strReturnDistId
    Exit Function
    
FillInMissingPieces_EH:
    TrsRaiseError m_strCMODULENAME, "FillInMissingPieces"
   
End Function

Private Function GetDepartmentCompanyCode( _
            ByVal v_strDept As String _
            ) As String
            
    Dim l_objDept As dbDepartment.cdDepartment
    
   
On Error GoTo eh:
    
    

    Set l_objDept = CreateObject("dbDepartment.cdDepartment")    'MTS standard
            
    GetDepartmentCompanyCode = l_objDept.GetCompanyCode(v_strDept)
    
    Set l_objDept = Nothing
            
eh:
    Set l_objDept = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDepartmentCompanyCode"

                    
            
End Function


Private Function GetFMC( _
            ByVal v_lngChargeWGNum As Long, _
            ByVal v_strPayrollDept As String _
            ) As String
            
    Dim l_strDept As String
'    Dim l_obj As dbWorkGroup.cdWorkGroup
'
'    Dim l_intPosition As Integer
   
On Error GoTo eh:
    
 'FMC is derived from the Payroll Distribution Department
 'The Payroll Distribution Department (time charge to)
 ' is the same as the Payroll Department (check picked up from)
 ' , except for the following
 '      1. Employee's payroll department is 086
 '      2. Employee's payroll department is a DOJM department
 '      3. Employee's time is charge to a DOJM work group
 ' If an employee meets one of the conditions above,
 '  then the Payroll Distribution Department is derived
 '  by the work group the time is being charged to.
    
'    If IsDOJMDepartmentNumber(v_strPayrollDept) Or _
'        IsDOJMWorkGroupPrivate(v_lngChargeWGNum) Or _
'        v_strPayrollDept = "086" Then
''Determine FMC from Department of Work group
'    Else
'        GetFMC = GetBudgetNum(v_strPayrollDept)
'        Exit Function
'    End If

'    If IsStringInList2(str(v_lngChargeWGNum), l_varWorkGroupDepartmentCache, l_lngWorkGroupDepartmentCacheCount, l_intPosition) Then
'        l_strDept = l_varWorkGroupDepartmentCache(l_intPosition, 1)
'    Else
'
'        Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
'        l_strDept = l_obj.DepartmentNumber(v_lngChargeWGNum)
'        Set l_obj = Nothing
'        l_varWorkGroupDepartmentCache(l_lngWorkGroupDepartmentCacheCount, 0) = str(v_lngChargeWGNum)
'        l_varWorkGroupDepartmentCache(l_lngWorkGroupDepartmentCacheCount, 1) = l_strDept
'        l_lngWorkGroupDepartmentCacheCount = l_lngWorkGroupDepartmentCacheCount + 1
'    End If
    
    l_strDept = DetermineDistributionDepartment(v_strPayrollDept, v_lngChargeWGNum)
    
    GetFMC = GetBudgetNum(l_strDept)
    
    Exit Function
    
eh:
'    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetFMC"

            
End Function

Private Function GetBudgetNum( _
            ByVal v_strDeptNum As String _
            ) As String
            
    Dim l_strBudgetNum As String
    Dim l_intPosition As Integer
   
On Error GoTo eh:
    
    If IsStringInList2(v_strDeptNum, l_varDepartmentCache, l_lngDepartmentCacheCount, l_intPosition) Then
        l_strBudgetNum = l_varDepartmentCache(l_intPosition, 1)
    Else
        l_intPosition = SetDepartmentCache(v_strDeptNum)
        l_strBudgetNum = l_varDepartmentCache(l_intPosition, 1)
    End If
    GetBudgetNum = l_strBudgetNum
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetBudgetNum"
            
End Function

Private Function GetAcctDiv( _
            ByVal v_strDeptNum As String _
            ) As String

    Dim l_strAcctDiv As String
    Dim l_intPosition As Integer
   
On Error GoTo eh:

    If IsStringInList2(v_strDeptNum, l_varDepartmentCache, l_lngDepartmentCacheCount, l_intPosition) Then
        l_strAcctDiv = l_varDepartmentCache(l_intPosition, 2)
    Else
        l_intPosition = SetDepartmentCache(v_strDeptNum)
        l_strAcctDiv = l_varDepartmentCache(l_intPosition, 2)
    End If
    GetAcctDiv = l_strAcctDiv
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetAcctDiv"
            

End Function

Private Function GetResourceType( _
            ByVal v_strDept As String _
            ) As String
            
'Using the department number to determine if it is a
' management or contract department
'The first character of a management department is 5 or greater

' The union_cd field of the department could be used,
' but is not populated for all departments

    Dim l_str1stChar As String
    Dim l_strResourceType As String
    
    On Error GoTo GetResourceType_EH
    
    l_str1stChar = Left(v_strDept, 1)
    If IsNumeric(l_str1stChar) Then
        If Val(l_str1stChar) < 5 Then
            l_strResourceType = "LU"
        Else
            l_strResourceType = "LM"
        End If
    Else
        l_strResourceType = "LU"
    End If
    
    GetResourceType = l_strResourceType
    Exit Function
    
GetResourceType_EH:
    TrsRaiseError m_strCMODULENAME, "GetResourceType"
            
End Function

Private Function ParseStringDistIDToAnArray( _
            ByVal v_strDistID As String _
            ) As Variant
            
    Dim l_aryDistID(11) As String
    Dim l_intI As Integer
    
    On Error GoTo ParseStringDistIDToAnArray_EH
    
    l_aryDistID(0) = Left$(v_strDistID, 1) '(U)Utility
    l_aryDistID(1) = Mid$(v_strDistID, 2, 2) '(BUD) Budget
    l_aryDistID(2) = Mid$(v_strDistID, 4, 3) '(MAJ)FERC Major
    l_aryDistID(3) = Mid$(v_strDistID, 7, 3) '(MIN)FERC Minor
    l_aryDistID(4) = Mid$(v_strDistID, 10, 3) '(FMC) Fund Mgt. Center
    l_aryDistID(5) = Mid$(v_strDistID, 13, 3) '(RMC) Resource Mgt. Center
    l_aryDistID(6) = Mid$(v_strDistID, 16, 1) '(T) Trans Type
    l_aryDistID(7) = Mid$(v_strDistID, 17, 5) '(PROJT) Project
    l_aryDistID(8) = Mid$(v_strDistID, 22, 2) '(PR) Product
    l_aryDistID(9) = Mid$(v_strDistID, 24, 4) '(ACTV) Activity
    l_aryDistID(10) = Mid$(v_strDistID, 28, 2) '(RT) Resource Type
    l_aryDistID(11) = Mid$(v_strDistID, 30, 10) '(FEEDER_REF) Memo
           
    ParseStringDistIDToAnArray = l_aryDistID
    
    Exit Function
    
ParseStringDistIDToAnArray_EH:
    TrsRaiseError m_strCMODULENAME, "ParseStringDistIDToAnArray"
            
            
End Function

Private Function ParseArrayDistIDToAString( _
            ByVal l_varDistributionId As Variant _
            ) As String
            
  Dim l_strDistributionId As String

    On Error GoTo ParseArrayDistIDToAString_EH
    
'Converting Arrary to 39 Character String
    l_strDistributionId = Right(" " & l_varDistributionId(0), 1) & _
                          Right("  " & l_varDistributionId(1), 2) & _
                          Right("   " & l_varDistributionId(2), 3) & _
                          Right("   " & l_varDistributionId(3), 3) & _
                          Right("   " & l_varDistributionId(4), 3) & _
                          Right("   " & l_varDistributionId(5), 3) & _
                          Right(" " & l_varDistributionId(6), 1) & _
                          Right("     " & l_varDistributionId(7), 5) & _
                          Right("  " & l_varDistributionId(8), 2) & _
                          Right("    " & l_varDistributionId(9), 4) & _
                          Right("  " & l_varDistributionId(10), 2) & _
                          Right(Space(10) & l_varDistributionId(11), 10)
    
    ParseArrayDistIDToAString = l_strDistributionId
    Exit Function
    
ParseArrayDistIDToAString_EH:
    TrsRaiseError m_strCMODULENAME, "ParseArrayDistIDToAString"
                   
            
End Function


Private Function GetPrimaryWG( _
            ByVal v_lngEmployee As Long, _
            ByVal v_datChargeDate As Date, _
            ByVal v_rsEmployeeWGData As ADODB.Recordset _
            ) As Long

    Dim l_bolWGfound As Boolean

    On Error GoTo GetPrimaryWG_EH
    
'    l_strSQL = "SELECT B.EMP_NUM, B.WG_NUM, A.WG_DESC, B.EFF_START_DATE, B.EFF_END_DATE, "
'    l_strSQL = l_strSQL & "A.PROJECT_ID, B.PRIMARY_IND, A.WG_SUPERVISOR, C.LOC_CD, C.DEPT_NUM, "
    v_rsEmployeeWGData.Filter = "EMP_NUM = " & str(v_lngEmployee)
    v_rsEmployeeWGData.Sort = "EMP_NUM, EFF_START_DATE"
    If v_rsEmployeeWGData.EOF Then
'should not happen
        v_rsEmployeeWGData.Filter = ""
        GetPrimaryWG = 0
        Exit Function
    End If
    
    l_bolWGfound = False
    v_rsEmployeeWGData.MoveFirst
    While Not v_rsEmployeeWGData.EOF And Not l_bolWGfound
        If v_datChargeDate >= v_rsEmployeeWGData!EFF_START_DATE _
            And v_datChargeDate <= v_rsEmployeeWGData!EFF_END_DATE _
            And v_rsEmployeeWGData!PRIMARY_IND = "Y" Then
            l_bolWGfound = True
        Else
            v_rsEmployeeWGData.MoveNext
        End If
    Wend
    If l_bolWGfound Then
        GetPrimaryWG = v_rsEmployeeWGData!WG_NUM
    Else
        v_rsEmployeeWGData.MoveFirst
        GetPrimaryWG = v_rsEmployeeWGData!WG_NUM
    End If

    v_rsEmployeeWGData.Sort = ""
    v_rsEmployeeWGData.Filter = ""
    Exit Function
    
GetPrimaryWG_EH:
    TrsRaiseError m_strCMODULENAME, "GetPrimaryWG"

End Function

Private Function GetChargedWG( _
            ByVal v_lngEmployee As Long, _
            ByVal v_datChargeDate As Date, _
            ByVal v_rsEmployeeWGData As ADODB.Recordset, _
            ByVal v_rsTimeEntries As ADODB.Recordset _
            ) As Long
            
    Dim l_lngWGNum As Long
            
    On Error GoTo GetChargedWG_EH
            
    If IsEmpty(v_rsTimeEntries!WG_NUM) Or IsNull(v_rsTimeEntries!WG_NUM) Then
        l_lngWGNum = GetPrimaryWG(v_rsTimeEntries!KEY_EMP_NUM, v_rsTimeEntries!KEY_CHARGE_DATE, v_rsEmployeeWGData)
    Else
        If v_rsTimeEntries!WG_NUM = 0 Then
            l_lngWGNum = GetPrimaryWG(v_rsTimeEntries!KEY_EMP_NUM, v_rsTimeEntries!KEY_CHARGE_DATE, v_rsEmployeeWGData)
        Else
            l_lngWGNum = v_rsTimeEntries!WG_NUM
        End If
    End If
    
    GetChargedWG = l_lngWGNum
    Exit Function
    
GetChargedWG_EH:
    TrsRaiseError m_strCMODULENAME, "GetChargedWG"
      
End Function

Private Function GetDistributionType( _
            ByVal v_rsTimeEntries As ADODB.Recordset _
            ) As String
            
    Dim l_strDistID As String
    Dim l_strDistType As String
    
    On Error GoTo GetDistributionType_EH
    
    l_strDistID = GetDistributionID(v_rsTimeEntries)
    
    If IsEmpty(v_rsTimeEntries!DISTRIBUTION_TYPE) Or IsNull(v_rsTimeEntries!DISTRIBUTION_TYPE) Then
        If Trim(v_rsTimeEntries!DISTRIBUTION_TYPE) <> "J" Then
            l_strDistType = DetermineDistributionType(l_strDistID)
            v_rsTimeEntries!DISTRIBUTION_TYPE = l_strDistType
        End If
    Else
        If Trim(v_rsTimeEntries!DISTRIBUTION_TYPE = "") Then
            l_strDistType = DetermineDistributionType(l_strDistID)
            v_rsTimeEntries!DISTRIBUTION_TYPE = l_strDistType
        Else
            l_strDistType = v_rsTimeEntries!DISTRIBUTION_TYPE
        End If
    End If
    
    GetDistributionType = l_strDistType
    
    Exit Function
    
GetDistributionType_EH:
    TrsRaiseError m_strCMODULENAME, "GetDistributionType"
      
        
End Function
Private Function ValidateDOJMDistrict( _
            ByVal v_strDOJMDistrict As String, _
            ByVal v_lngChargeWGNum As Long, _
            ByRef r_strReturnMsg As String _
            ) As Boolean
    
    Dim l_strDOJMDistrict As String
    Dim l_strDOJMHeadquarters As String
    
        
On Error GoTo ValidateDOJMDistrict_EH

    If Not IsDOJMWorkGroup(v_lngChargeWGNum) Then
        r_strReturnMsg = "Can not enter a Work request for a non DOJM work group " & str(v_lngChargeWGNum)
        ValidateDOJMDistrict = False
        Exit Function
    End If
    If Trim(v_strDOJMDistrict) = "" Then
        r_strReturnMsg = "DOJM District is missing "
        ValidateDOJMDistrict = False
        Exit Function
    End If
    GetDOJMDistrictAndHeadQuarter v_lngChargeWGNum, l_strDOJMDistrict, l_strDOJMHeadquarters
    If Trim(l_strDOJMDistrict) = "" Then
        r_strReturnMsg = "Unable to retrieve DOJM District for work group " & str(v_lngChargeWGNum)
        ValidateDOJMDistrict = False
        Exit Function
    End If
    If Trim(l_strDOJMHeadquarters) = "" Then
        r_strReturnMsg = "Unable to retrieve DOJM Headquarters for work group " & str(v_lngChargeWGNum)
        ValidateDOJMDistrict = False
        Exit Function
    End If
    If l_strDOJMDistrict <> v_strDOJMDistrict Then
        r_strReturnMsg = "DOJM District " & v_strDOJMDistrict & " is not valid for work group " & str(v_lngChargeWGNum) & _
                         ".  The DOJM District Should be " & l_strDOJMDistrict
        ValidateDOJMDistrict = False
        Exit Function
    End If
    
    ValidateDOJMDistrict = True

Exit Function

ValidateDOJMDistrict_EH:

    TrsRaiseError m_strCMODULENAME, "ValidateDOJMDistrict"

End Function
Private Function ValidateEMPRVJobTask( _
            ByVal v_lngChargeWGNum As Long, _
            ByRef r_strReturnMsg As String _
            ) As Boolean
    
On Error GoTo ValidateEMPRVJobTask_EH

    If Not IsEMPRVJobTaskWG(v_lngChargeWGNum) Then
        r_strReturnMsg = "Can not enter a Job Task for a non EMPRV work group " & str(v_lngChargeWGNum)
        ValidateEMPRVJobTask = False
        Exit Function
    End If
    
    ValidateEMPRVJobTask = True

Exit Function

ValidateEMPRVJobTask_EH:

    TrsRaiseError m_strCMODULENAME, "ValidateEMPRVJobTask"

End Function

Private Function GetInterfaceCode( _
            ByVal v_strDeptNum As String _
            ) As String

    
    Dim l_strInterfaceCode As String
    Dim l_intPosition As Integer

On Error GoTo eh:

    If IsStringInList2(v_strDeptNum, l_varDepartmentCache, l_lngDepartmentCacheCount, l_intPosition) Then
        l_strInterfaceCode = l_varDepartmentCache(l_intPosition, 3)
    Else
        l_intPosition = SetDepartmentCache(v_strDeptNum)
        l_strInterfaceCode = l_varDepartmentCache(l_intPosition, 3)
    End If
    GetInterfaceCode = l_strInterfaceCode
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetInterfaceCode"
            

End Function

Private Function GetAcctType( _
            ByVal v_strDeptNum As String _
            ) As String

    Dim l_strAcctType As String
    Dim l_intPosition As Integer

   
On Error GoTo eh:

    If IsStringInList2(v_strDeptNum, l_varDepartmentCache, l_lngDepartmentCacheCount, l_intPosition) Then
        l_strAcctType = l_varDepartmentCache(l_intPosition, 4)
    Else
        l_intPosition = SetDepartmentCache(v_strDeptNum)
        l_strAcctType = l_varDepartmentCache(l_intPosition, 4)
    End If
    GetAcctType = l_strAcctType
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetAcctType"
            

End Function

Private Function GetNonProdLoadInd( _
            ByVal v_strDeptNum As String _
            ) As String

    Dim l_strNonProdLoadInd As String
    Dim l_intPosition As Integer

   
On Error GoTo eh:

    If IsStringInList2(v_strDeptNum, l_varDepartmentCache, l_lngDepartmentCacheCount, l_intPosition) Then
        l_strNonProdLoadInd = l_varDepartmentCache(l_intPosition, 5)
    Else
        l_intPosition = SetDepartmentCache(v_strDeptNum)
        l_strNonProdLoadInd = l_varDepartmentCache(l_intPosition, 5)
    End If
    GetNonProdLoadInd = l_strNonProdLoadInd
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetNonProdLoadInd"
            

End Function

Private Function GetAcctNum( _
            ByVal v_strDeptNum As String _
            ) As String

    Dim l_strAcctNum As String
    Dim l_intPosition As Integer

   
On Error GoTo eh:

    If IsStringInList2(v_strDeptNum, l_varDepartmentCache, l_lngDepartmentCacheCount, l_intPosition) Then
        l_strAcctNum = l_varDepartmentCache(l_intPosition, 6)
    Else
        l_intPosition = SetDepartmentCache(v_strDeptNum)
        l_strAcctNum = l_varDepartmentCache(l_intPosition, 6)
    End If
    GetAcctNum = l_strAcctNum
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetAcctNum"
            

End Function


Private Function SetDepartmentCache( _
            ByVal v_strDeptNum As String _
            ) As Integer
            
    Dim l_strAcctType As String
    Dim l_obj As dbDepartment.cdDepartment
    
    Dim l_intPosition As Integer
    Dim l_strBudgetNum As String
    Dim l_strInterfaceCode As String
    Dim l_strAcctDiv As String
    Dim l_strNonLoadProd As String
    Dim l_strAcctNum As String
    Dim l_strCompany As String

    
    On Error GoTo eh:
    
    l_intPosition = l_lngDepartmentCacheCount
    
    Set l_obj = CreateObject("dbDepartment.cdDepartment")
    l_strAcctType = l_obj.GetAccountType(v_strDeptNum)
    l_strAcctDiv = l_obj.GetAccountDivision(v_strDeptNum)
    l_strBudgetNum = l_obj.GetBudgetNumber(v_strDeptNum)
    l_strInterfaceCode = l_obj.GetSystemInterfaceCode(v_strDeptNum)
    l_strNonLoadProd = l_obj.GetNonProdLoadInd(v_strDeptNum)
    l_strAcctNum = l_obj.GetAccountNumber(v_strDeptNum)
    l_strCompany = l_obj.GetCompanyCode(v_strDeptNum)

    Set l_obj = Nothing
    l_varDepartmentCache(l_lngDepartmentCacheCount, 0) = v_strDeptNum
    l_varDepartmentCache(l_lngDepartmentCacheCount, 1) = l_strBudgetNum
    l_varDepartmentCache(l_lngDepartmentCacheCount, 2) = l_strAcctDiv
    l_varDepartmentCache(l_lngDepartmentCacheCount, 3) = l_strInterfaceCode
    l_varDepartmentCache(l_lngDepartmentCacheCount, 4) = l_strAcctType
    l_varDepartmentCache(l_lngDepartmentCacheCount, 5) = l_strNonLoadProd
    l_varDepartmentCache(l_lngDepartmentCacheCount, 6) = l_strAcctNum
    l_varDepartmentCache(l_lngDepartmentCacheCount, 8) = l_strCompany
    If IsMKSDept(v_strDeptNum) Then
        l_varDepartmentCache(l_lngDepartmentCacheCount, 7) = "Y"
    Else
        l_varDepartmentCache(l_lngDepartmentCacheCount, 7) = "N"
    End If
    l_lngDepartmentCacheCount = l_lngDepartmentCacheCount + 1
    
    SetDepartmentCache = l_intPosition
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SetDepartmentCache"
            
End Function


Private Function GetExceptionAccountNum( _
            ByVal v_strExcCd As String, _
            Optional ByVal v_strCompanyCode As String = "" _
            ) As String
            
    Dim l_strExcAcctNum As String
    Dim l_intPosition As Integer
    
    On Error GoTo eh:
                
    If Trim(v_strExcCd) = "" Then
        GetExceptionAccountNum = ""
        Exit Function
    End If

    If IsStringInList2(v_strExcCd, l_varExceptionCodeCache, l_lngExceptionCodeCacheCount, l_intPosition) Then
        l_strExcAcctNum = l_varExceptionCodeCache(l_intPosition, 1)
    Else
        l_intPosition = SetExceptionCache(v_strExcCd)
        l_strExcAcctNum = l_varExceptionCodeCache(l_intPosition, 1)
    End If
        
    Select Case v_strCompanyCode
        Case "AIC"
            l_strExcAcctNum = l_varExceptionCodeCache(l_intPosition, 6)
        Case "AMS"
            l_strExcAcctNum = l_varExceptionCodeCache(l_intPosition, 7)
        Case "UEC"
            l_strExcAcctNum = l_varExceptionCodeCache(l_intPosition, 8)
        Case Else
            l_strExcAcctNum = l_varExceptionCodeCache(l_intPosition, 1)
        
    
    End Select
        
    GetExceptionAccountNum = l_strExcAcctNum
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetExceptionAccountNum"
            
End Function

Private Function GetExceptionBankCode( _
            ByVal v_strExcCd As String _
            ) As String
            
    Dim l_strBankCode As String
    Dim l_intPosition As Integer
    
    On Error GoTo eh:
                
    If Trim(v_strExcCd) = "" Then
        GetExceptionBankCode = ""
        Exit Function
    End If
    
    If IsStringInList2(v_strExcCd, l_varExceptionCodeCache, l_lngExceptionCodeCacheCount, l_intPosition) Then
        l_strBankCode = l_varExceptionCodeCache(l_intPosition, 3)
    Else
        l_intPosition = SetExceptionCache(v_strExcCd)
        l_strBankCode = l_varExceptionCodeCache(l_intPosition, 3)
    End If
        
    GetExceptionBankCode = l_strBankCode
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetExceptionBankCode"
            
End Function

            
Private Sub ValidateDistributionIDsDefaultAccounting( _
            ByRef r_rsTimeEntries As ADODB.Recordset, _
            ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset, _
            Optional v_blnOtherPay As Boolean = False)

    Dim l_strDefaultAccounting As String
    Dim l_strDistID As String
    
On Error GoTo eh:

    If RecordsExist(r_rsTimeEntries) Then
    Else
        Exit Sub
    End If
            
    TimeEntriesFilterAndSort r_rsTimeEntries
    If r_rsTimeEntries.EOF Then
        Exit Sub
    End If
        
    r_rsTimeEntries.MoveFirst
    While Not r_rsTimeEntries.EOF
        l_strDistID = GetDistributionID(r_rsTimeEntries)
        If v_blnOtherPay And l_strDistID <> "" Then
        Else
            l_strDefaultAccounting = ValidateDistributionIDsGetDefaultAccount(r_rsTimeEntries, v_rsEmployeeDistIDInfo, v_blnOtherPay)
            If l_strDefaultAccounting = "" Or l_strDefaultAccounting = "N/A" Then
                If l_strDefaultAccounting = "N/A" Then
                'Actually equals the description of Override
                'or exception code.  Since exception codes can default
                'I am not placing that logic here
                    r_rsTimeEntries!DISTRIBUTION_ID = Space(39)
                    r_rsTimeEntries!DISTRIBUTION_TYPE = Space(1)
                    r_rsTimeEntries!JOB_NBR = Space(8)
                    r_rsTimeEntries!JOBTASK_NBR = Space(4)
                    r_rsTimeEntries!JOBTASK_STEP_NBR = 0
                    r_rsTimeEntries!EMPRV_DATABASE_ID = Space(8)
                End If
            Else
                r_rsTimeEntries!DISTRIBUTION_ID = l_strDefaultAccounting
                If v_blnOtherPay Then
                    r_rsTimeEntries!DISTRIBUTION_TYPE = "A"
                Else
                    If Trim(r_rsTimeEntries!JOB_NBR) <> "" Then
                        '//r_rsTimeEntries!DISTRIBUTION_TYPE = "J"
                        r_rsTimeEntries!DISTRIBUTION_TYPE = DetermineDistributionTypeFromJobTask(r_rsTimeEntries!JOB_NBR, r_rsTimeEntries!JOBTASK_NBR)
                    Else
                        r_rsTimeEntries!DISTRIBUTION_TYPE = "A"
                    End If
                End If
            End If
        End If
        r_rsTimeEntries.MoveNext
    Wend
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIDsDefaultAccounting"
                
            
End Sub

Private Function DetermineDistributionTypeFromJobTask(v_sJobNbr As String, v_sJobTaskNbr As String) As String
On Error GoTo eh
    If IsEMPRVJobTask(v_sJobNbr & v_sJobTaskNbr) Then
       DetermineDistributionTypeFromJobTask = "J"
    Else
       DetermineDistributionTypeFromJobTask = "M"
    End If
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineDistributionTypeFromJobTask"
End Function

Private Function IsEMPRVJobTask(v_sJobTask As String) As Boolean
On Error GoTo eh
    Dim JobType As String
    
    JobType = Left(v_sJobTask, 2)
    
    Select Case JobType
        Case "FU", "JH", "JR", "JS", "JU", "XJ"
            IsEMPRVJobTask = True
        Case Else
            IsEMPRVJobTask = False
    End Select
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsEMPRVJobTask"
End Function

Private Function ValidateDistributionIDsGetDefaultAccount( _
            ByVal v_rsTimeEntries As ADODB.Recordset, _
            ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset, _
            Optional v_blnOtherPay As Boolean = False _
            ) As String


    Dim l_strExceptionCode As String
    Dim l_strOverExcCode As String
    Dim l_strDept As String
    Dim l_strDefaultAccounting As String
    Dim l_strBankCode As String
    Dim l_strExceptionAcctNum As String
    
    On Error GoTo ValidateDistributionIDsGetDefaultAccount_EH

    l_strExceptionCode = ""
    l_strOverExcCode = ""
    If v_blnOtherPay Then
        ValidateDistributionIDsGetDefaultAccount = _
            GetTTADistID(v_rsTimeEntries)
        Exit Function
    End If
'Can not have an override code without an exception code
'Assuming straight Time ST
    If IsEmpty(v_rsTimeEntries!EXCEPTION_CD) Or IsNull(v_rsTimeEntries!EXCEPTION_CD) Then
        l_strExceptionCode = "ST"
    Else
        l_strExceptionCode = v_rsTimeEntries!EXCEPTION_CD
    End If
    
    l_strBankCode = GetExceptionBankCode(l_strExceptionCode)
    If IsEmpty(v_rsTimeEntries!OVER_EXC_CD) Or IsNull(v_rsTimeEntries!OVER_EXC_CD) Then
        l_strOverExcCode = ""
    Else
        l_strOverExcCode = v_rsTimeEntries!OVER_EXC_CD
    End If
    l_strDept = ValidateDistIDGetPayrollDept( _
             v_rsTimeEntries!KEY_EMP_NUM, _
             v_rsTimeEntries!KEY_CHARGE_DATE, _
             v_rsEmployeeDistIDInfo)
             
    l_strDefaultAccounting = GetDefaultAcct(l_strExceptionCode, l_strDept, l_strBankCode, l_strOverExcCode)
    
    'If v_rsTimeEntries!DISTRIBUTION_TYPE = "J" Then
    '    l_strDefaultAccounting = ""
    'End If
        
'JIMB Begin  11/29/2001
'Code to clear default Accounting generated from client
'This happens when an employee in a fixed department
'charges time to a work group in a variable department
    If GetAcctType(l_strDept) = "F" Then
        If Trim(l_strOverExcCode) <> "" Then
            l_strExceptionAcctNum = GetExceptionAccountNum(l_strOverExcCode)
            If Trim(l_strExceptionAcctNum) <> "" Then
                l_strDefaultAccounting = "N/A"
            End If
        Else
            l_strExceptionAcctNum = GetExceptionAccountNum(l_strExceptionCode)
            If Trim(l_strExceptionAcctNum) <> "" Then
                l_strDefaultAccounting = "N/A"
            End If
        End If
    End If
'JIMB End  11/29/2001
    
    ValidateDistributionIDsGetDefaultAccount = l_strDefaultAccounting
    Exit Function
    
ValidateDistributionIDsGetDefaultAccount_EH:
    TrsRaiseError m_strCMODULENAME, "ValidateDistributionIDsGetDefaultAccount"
   
End Function

Private Sub TimeEntriesFilterAndSort( _
            ByRef r_rsTimeEntries As ADODB.Recordset, _
            Optional v_blnOtherPay As Boolean = False)
                       
On Error GoTo Sort_EH:

                       
    If v_blnOtherPay Then
        r_rsTimeEntries.Filter = MDAC25FilterFix("RECORD_STATUS <> 'D' AND EXCEPTION_CD = 'TTA'")
    Else
        r_rsTimeEntries.Filter = "RECORD_STATUS <> 'D'"
    End If
    If r_rsTimeEntries.EOF Then
        Exit Sub
    End If
    r_rsTimeEntries.Sort = "KEY_EMP_NUM, KEY_CHARGE_DATE, KEY_START_TIME, END_TIME"
             
    Exit Sub
    
Sort_EH:
    TrsRaiseError m_strCMODULENAME, "TimeEntriesFilterAndSort"
            
End Sub

Private Function WriteBRV( _
            ByRef r_rsBRVs As ADODB.Recordset, _
            ByVal v_rsTimeEntries As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDesc As String, _
            Optional v_blnOtherPay As Boolean = False _
            ) As Boolean
            
    Dim l_bolError As Boolean
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_decHours As Variant
    Dim l_strExceptionCode As String
    Dim l_strOverExcCode As String
    Dim l_datDateWorked As Date
    
    On Error GoTo WriteBRV_EH
    
    If IsEmpty(v_rsTimeEntries!DATE_WORKED) Or IsNull(v_rsTimeEntries!DATE_WORKED) Then
        l_datDateWorked = v_rsTimeEntries!KEY_CHARGE_DATE
    Else
        l_datDateWorked = v_rsTimeEntries!DATE_WORKED
    End If
    
    If IsEmpty(v_rsTimeEntries!START_TIME) Or IsNull(v_rsTimeEntries!START_TIME) Then
        l_strStartTime = v_rsTimeEntries!KEY_START_TIME
    Else
        l_strStartTime = v_rsTimeEntries!START_TIME
    End If
    
    If IsEmpty(v_rsTimeEntries!END_TIME) Or IsNull(v_rsTimeEntries!END_TIME) Then
        l_strEndTime = ""
    Else
        l_strEndTime = v_rsTimeEntries!END_TIME
    End If
    
    If IsEmpty(v_rsTimeEntries!HOURS) Or IsNull(v_rsTimeEntries!HOURS) Then
        l_decHours = 0
    Else
        l_decHours = v_rsTimeEntries!HOURS
    End If
    
    If IsEmpty(v_rsTimeEntries!EXCEPTION_CD) Or IsNull(v_rsTimeEntries!EXCEPTION_CD) Then
        l_strExceptionCode = ""
    Else
        l_strExceptionCode = v_rsTimeEntries!EXCEPTION_CD
    End If
    
    If v_blnOtherPay Then
    ' Other Pay Does not have an override exception code
        l_strOverExcCode = ""
    Else
        If IsEmpty(v_rsTimeEntries!OVER_EXC_CD) Or IsNull(v_rsTimeEntries!OVER_EXC_CD) Then
            l_strOverExcCode = ""
        Else
            l_strOverExcCode = v_rsTimeEntries!OVER_EXC_CD
        End If
    End If

    
    l_bolError = AppendError(r_rsBRVs, _
            v_rsTimeEntries!KEY_EMP_NUM, _
            v_rsTimeEntries!KEY_CHARGE_DATE, _
            l_datDateWorked, _
            l_strStartTime, _
            l_strEndTime, _
            l_decHours, _
            l_strExceptionCode, _
            l_strOverExcCode, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDesc, _
            "cbsDistributionMgr")
            
    WriteBRV = l_bolError
    Exit Function
WriteBRV_EH:
    TrsRaiseError m_strCMODULENAME, "WriteBRV"

End Function

Private Function GetDistributionID( _
            ByVal v_rsTimeEntries As ADODB.Recordset _
            ) As String
                   
    Dim l_strDistID As String
    
    On Error GoTo GetDistributionID_EH
    
    If IsEmpty(v_rsTimeEntries!DISTRIBUTION_ID) Or IsNull(v_rsTimeEntries!DISTRIBUTION_ID) Then
        l_strDistID = ""
    Else
        l_strDistID = v_rsTimeEntries!DISTRIBUTION_ID
    End If
    
    GetDistributionID = l_strDistID
    Exit Function
    
GetDistributionID_EH:
    TrsRaiseError m_strCMODULENAME, "GetDistributionID"
    
End Function


Private Function GetTTADistID( _
            ByVal v_rsOtherPayEntries As ADODB.Recordset _
            ) As String
            
    Dim l_strDistID
    Dim l_intStartLoc As Integer
    Dim l_intEndLoc As Integer
    Dim l_lngYear As Long
    Dim l_strChemTester As String
    
    On Error GoTo GetTTADistID_EH
    
    l_strDistID = GetDistributionID(v_rsOtherPayEntries)
    
    If l_strDistID <> "" Then
        GetTTADistID = ""
        Exit Function
    End If
    
    l_intStartLoc = GetOtherPayStartLoc(v_rsOtherPayEntries)
    l_intEndLoc = GetOtherPayEndLoc(v_rsOtherPayEntries)
    If l_intStartLoc = 0 _
        Or l_intEndLoc = 0 Then
        GetTTADistID = ""
        Exit Function
    End If
    
    If m_rsTravelAllowances Is Nothing Then
        Set m_rsTravelAllowances = GetTTATravelAllowances()
    End If
    
    l_lngYear = Year(v_rsOtherPayEntries!CHARGE_DATE)
    
    '  Convert Chem Tester
    If GetOtherPayChemTesterInd(v_rsOtherPayEntries) = "Y" Then
        l_strChemTester = "C"
    Else
        l_strChemTester = ""
    End If
    
    '  Filter the recordset with the given arguments
    m_rsTravelAllowances.Filter = MDAC25FilterFix( _
            "START_LOC_CD = " & str(l_intStartLoc) & _
            " AND END_LOC_CD = " & str(l_intEndLoc) & _
            " AND USER_TYPE = '" & l_strChemTester & "'" & _
            " AND EFF_YEAR = " & str(l_lngYear))
                                                
    '  Find DistId and Amount
    If Not m_rsTravelAllowances.EOF And Not m_rsTravelAllowances.BOF Then
            GetTTADistID = m_rsTravelAllowances!DISTRIBUTION_ID
    Else
            GetTTADistID = ""
    End If
    Exit Function
    
GetTTADistID_EH:
    TrsRaiseError m_strCMODULENAME, "GetTTADistID"
               
End Function

Private Function GetOtherPayStartLoc( _
            ByVal v_rsOtherPayEntries As ADODB.Recordset _
            ) As Integer
    
    On Error GoTo GetOtherPayStartLoc_EH

    If IsEmpty(v_rsOtherPayEntries!START_LOC_CD) _
        Or IsNull(v_rsOtherPayEntries!START_LOC_CD) Then
        GetOtherPayStartLoc = 0
    Else
        GetOtherPayStartLoc = v_rsOtherPayEntries!START_LOC_CD
    End If
    Exit Function
    
GetOtherPayStartLoc_EH:
    TrsRaiseError m_strCMODULENAME, "GetOtherPayStartLoc"
               
    
End Function

Private Function GetOtherPayEndLoc( _
            ByVal v_rsOtherPayEntries As ADODB.Recordset _
            ) As Integer

    On Error GoTo GetOtherPayEndLoc_EH
    
    If IsEmpty(v_rsOtherPayEntries!END_LOC_CD) _
        Or IsNull(v_rsOtherPayEntries!END_LOC_CD) Then
        GetOtherPayEndLoc = 0
    Else
        GetOtherPayEndLoc = v_rsOtherPayEntries!END_LOC_CD
    End If
    Exit Function
    
GetOtherPayEndLoc_EH:
    TrsRaiseError m_strCMODULENAME, "GetOtherPayEndLoc"
               
End Function

Private Function GetOtherPayChemTesterInd( _
            ByVal v_rsOtherPayEntries As ADODB.Recordset _
            ) As String

    On Error GoTo GetOtherPayChemTesterInd_EH

    If IsEmpty(v_rsOtherPayEntries!CHEM_MTL_TSTR_IND) _
        Or IsNull(v_rsOtherPayEntries!CHEM_MTL_TSTR_IND) Then
        GetOtherPayChemTesterInd = "N"
    Else
        GetOtherPayChemTesterInd = v_rsOtherPayEntries!CHEM_MTL_TSTR_IND
    End If
    Exit Function
    
GetOtherPayChemTesterInd_EH:
    TrsRaiseError m_strCMODULENAME, "GetOtherPayChemTesterInd"
              
    
End Function

Private Function GetTTATravelAllowances() As ADODB.Recordset


    Dim l_obj As dbCodeTables.cdCodeTables
    
    
    On Error GoTo eh:
    
    
    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    Set GetTTATravelAllowances = l_obj.GetTravelAllowances()
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTTATravelAllowances"
            
End Function

Public Function ValidateECDraftingAcctNum( _
            ByVal v_strDistributionId As String, _
            ByRef r_strReturnMsg As String, _
            ByVal v_strCorporateCode As String, _
            ByVal v_strPayrollDept As String, _
            ByVal v_lngChargeWGNum As Long, _
            ByRef r_strDistributionIDReturned As String, _
            ByRef r_strHighlightFields As String, _
            Optional v_strReasonForEdit As String = "FE" _
            ) As Boolean
TrsRaiseError m_strCMODULENAME, "ValidateECDraftingAcctNum", trErrObsoleteMethod, , , , True
'    Dim l_strDistributionIDEdits(221) As String
'    Dim l_intNumberOfEdits As Integer
'    Dim l_strDistributionIDsReturned As Variant
'    Dim l_strReturnCodes As Variant
'    Dim l_strReturnMsgs As Variant
'    Dim l_strHighlightFields As Variant
'    Dim l_strDistributionType As String
'    Dim l_varDistributionId As Variant
'    Dim l_bolValidationSuccessful As Boolean
'    Dim l_strDistributionId As String
'    Dim l_varDistID As Variant
'
'
'    On Error GoTo eh:
'
'    l_strDistributionType = DetermineDistributionType(v_strDistributionId)
'    Select Case l_strDistributionType
'        Case "A", "S", "J"
'
'        Case "W"
'            r_strReturnMsg = "You can not enter a DOJM work request.  Please contact help desk."
'            ValidateECDraftingAcctNum = False
'            Exit Function
'        Case Else
'            r_strReturnMsg = "No Distribution ID to Validate"
'            ValidateECDraftingAcctNum = False
'            Exit Function
'    End Select
'
'    l_varDistID = ParseStringDistIDToAnArray(v_strDistributionId)
'    l_varDistributionId = FillInMissingPieces(l_varDistID, v_strPayrollDept, v_lngChargeWGNum, l_strDistributionType, "")
'    l_strDistributionId = ParseArrayDistIDToAString(l_varDistributionId)
'    If Trim(v_strCorporateCode) = "" Then
'        v_strCorporateCode = GetDepartmentCompanyCode(v_strPayrollDept)
'    End If
'    l_strDistributionIDEdits(0) = v_strReasonForEdit & v_strCorporateCode & l_strDistributionId
'    l_intNumberOfEdits = 1
'    l_bolValidationSuccessful = ValidateGLAccounting(l_strDistributionIDEdits, l_intNumberOfEdits, l_strDistributionIDsReturned, l_strReturnCodes, l_strReturnMsgs, l_strHighlightFields)
'    r_strReturnMsg = l_strReturnMsgs(0)
'    r_strDistributionIDReturned = l_strDistributionIDsReturned(0)
'    r_strHighlightFields = l_strHighlightFields(0)
'
'    ValidateECDraftingAcctNum = l_bolValidationSuccessful
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "ValidateECDraftingAcctNum", , , , , True
        

End Function
Public Function ValidateECDraftingAcctNum2( _
            ByVal v_strDistributionId As String, _
            ByVal v_strDistributionType As String, _
            ByRef r_strReturnMsg As String, _
            ByVal v_strCorporateCode As String, _
            ByVal v_strPayrollDept As String, _
            ByVal v_lngChargeWGNum As Long, _
            ByRef r_strDistributionIDReturned As String, _
            ByRef r_strHighlightFields As String, _
            Optional v_strReasonForEdit As String = "FE" _
            ) As Boolean
TrsRaiseError m_strCMODULENAME, "ValidateECDraftingAcctNum2", trErrObsoleteMethod, , , , True
'    Dim l_strDistributionIDEdits(221) As String
'    Dim l_intNumberOfEdits As Integer
'    Dim l_strDistributionIDsReturned As Variant
'    Dim l_strReturnCodes As Variant
'    Dim l_strReturnMsgs As Variant
'    Dim l_strHighlightFields As Variant
'    Dim l_strDistributionType As String
'    Dim l_varDistributionId As Variant
'    Dim l_bolValidationSuccessful As Boolean
'    Dim l_strDistributionId As String
'    Dim l_varDistID As Variant
'
'
'    On Error GoTo eh:
'
'    'l_strDistributionType = DetermineDistributionType(v_strDistributionId)
'    l_strDistributionType = v_strDistributionType
'    Select Case l_strDistributionType
'        Case "A", "S", "J"
'
'        Case "W"
'            r_strReturnMsg = "You can not enter a DOJM work request.  Please contact help desk."
'            ValidateECDraftingAcctNum2 = False
'            Exit Function
'        Case Else
'            r_strReturnMsg = "No Distribution ID to Validate"
'            ValidateECDraftingAcctNum2 = False
'            Exit Function
'    End Select
'
'    l_varDistID = ParseStringDistIDToAnArray(v_strDistributionId)
'    l_varDistributionId = FillInMissingPieces(l_varDistID, v_strPayrollDept, v_lngChargeWGNum, l_strDistributionType, "")
'    l_strDistributionId = ParseArrayDistIDToAString(l_varDistributionId)
'    If Trim(v_strCorporateCode) = "" Then
'        v_strCorporateCode = GetDepartmentCompanyCode(v_strPayrollDept)
'    End If
'    l_strDistributionIDEdits(0) = v_strReasonForEdit & v_strCorporateCode & l_strDistributionId
'    l_intNumberOfEdits = 1
'    l_bolValidationSuccessful = ValidateGLAccounting(l_strDistributionIDEdits, l_intNumberOfEdits, l_strDistributionIDsReturned, l_strReturnCodes, l_strReturnMsgs, l_strHighlightFields)
'    r_strReturnMsg = l_strReturnMsgs(0)
'    r_strDistributionIDReturned = l_strDistributionIDsReturned(0)
'    r_strHighlightFields = l_strHighlightFields(0)
'
'    ValidateECDraftingAcctNum2 = l_bolValidationSuccessful
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "ValidateECDraftingAcctNum2", , , , , True
'

End Function



Private Function GetDepartmentOfWorkGroup( _
            ByVal v_lngChargeWGNum As Long _
            ) As String
            
    Dim l_strDept As String
    Dim l_obj As dbWorkGroup.cdWorkGroup
    Dim l_obj2 As dbDepartment.cdDepartment
    
    Dim l_intPosition As Integer
   
On Error GoTo eh:
    
    If IsStringInList2(str(v_lngChargeWGNum), l_varWorkGroupDepartmentCache, l_lngWorkGroupDepartmentCacheCount, l_intPosition) Then
        l_strDept = l_varWorkGroupDepartmentCache(l_intPosition, 1)
    Else
        
        Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
        l_strDept = l_obj.DepartmentNumber(v_lngChargeWGNum)
        Set l_obj = Nothing
        l_varWorkGroupDepartmentCache(l_lngWorkGroupDepartmentCacheCount, 0) = str(v_lngChargeWGNum)
        l_varWorkGroupDepartmentCache(l_lngWorkGroupDepartmentCacheCount, 1) = l_strDept
        l_lngWorkGroupDepartmentCacheCount = l_lngWorkGroupDepartmentCacheCount + 1
    End If
    
    GetDepartmentOfWorkGroup = l_strDept
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    Set l_obj2 = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDepartmentOfWorkGroup"

            
End Function


Private Function RequiresAccounting( _
        ByVal v_strExceptionCd As String, _
        ByVal v_strDistDept As String, _
        ByVal v_strAcctType As String, _
        ByVal v_strBank As String, _
        ByVal v_strOverrideExcCd As String, _
        ByVal v_strPayrollDept As String, _
        ByRef v_blnHasDefaultAccounting As Boolean _
        ) As Boolean
    

    Dim l_strExceptionAcctNum As String
    Dim l_strCompany As String

On Error GoTo eh
    
        
'  Process Request
    
    v_blnHasDefaultAccounting = False
    
    If v_strAcctType <> "V" Then
        RequiresAccounting = False
        Exit Function
    End If
    
    l_strCompany = GetCompanyCode(v_strDistDept)
    
    If IsExecutiveDepartment(v_strDistDept) Then
        If GetExceptionExecutiveEarnCd(v_strExceptionCd) = "32" Then
            RequiresAccounting = False
            Exit Function
        End If
    Else
        If GetExceptionContractEarnCd(v_strExceptionCd) = "99" Then
            RequiresAccounting = False
            Exit Function
        End If
    End If

    
    If Trim(v_strOverrideExcCd) <> "" Then
        l_strExceptionAcctNum = GetExceptionAccountNum(v_strOverrideExcCd, l_strCompany)
        If Trim(l_strExceptionAcctNum) = "" Then
            RequiresAccounting = True
        Else
            RequiresAccounting = False
            v_blnHasDefaultAccounting = True
        End If
    'If the override exception code had default accounting  then exit
        If Trim(l_strExceptionAcctNum) <> "" Then
            Exit Function
        End If
    End If
    
    Select Case UCase(Trim(v_strExceptionCd))
    
        Case "ST", "STP", "SAS", "SUS", "SSC", "RW1", "RW2", "RW3", ""
            RequiresAccounting = True
        Case Else
            If v_strBank = "O" Then
                RequiresAccounting = True
            Else
                l_strExceptionAcctNum = GetExceptionAccountNum(v_strExceptionCd, l_strCompany)
                If Trim(l_strExceptionAcctNum) = "" Then
                    RequiresAccounting = True
                Else
                    RequiresAccounting = False
                    v_blnHasDefaultAccounting = True
                End If
            End If
    End Select
    
    
    Exit Function


eh:
    TrsRaiseError m_strCMODULENAME, "GetDefaultAcct", , , , , True


End Function


Private Function SetExceptionCache( _
            ByVal v_strExcCd As String _
            ) As Integer
            
    Dim l_obj As dbCodeTables.cdCodeTables
    
    Dim l_strExcAcctNum As String
    Dim l_strExcDesc As String
    Dim l_strBankCode As String
    Dim l_strContractEarnCd As String
    Dim l_strExecEarnCd As String
    Dim l_intPosition As Integer
    Dim l_str_AIC_ExcAcctNum As String
    Dim l_str_AMS_ExcAcctNum As String
    Dim l_str_UEC_ExcAcctNum As String

    
    On Error GoTo eh:
    
    l_intPosition = l_lngExceptionCodeCacheCount
    
    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    l_strExcAcctNum = l_obj.ExceptionAccountNumber(v_strExcCd)
    l_str_AIC_ExcAcctNum = l_obj.ExceptionAccountNumberAIC(v_strExcCd)
    l_str_AMS_ExcAcctNum = l_obj.ExceptionAccountNumberAMS(v_strExcCd)
    l_str_UEC_ExcAcctNum = l_obj.ExceptionAccountNumberUEC(v_strExcCd)
    l_strExcDesc = l_obj.ExceptionDescription(v_strExcCd)
    l_strBankCode = l_obj.ExceptionBankCode(v_strExcCd)
    l_strContractEarnCd = l_obj.ExceptionContractEarningsCode(v_strExcCd)
    l_strExecEarnCd = l_obj.ExceptionExecutiveEarningsCode(v_strExcCd)
    Set l_obj = Nothing
    l_varExceptionCodeCache(l_lngExceptionCodeCacheCount, 0) = v_strExcCd
    l_varExceptionCodeCache(l_lngExceptionCodeCacheCount, 1) = l_strExcAcctNum
    l_varExceptionCodeCache(l_lngExceptionCodeCacheCount, 2) = l_strExcDesc
    l_varExceptionCodeCache(l_lngExceptionCodeCacheCount, 3) = l_strBankCode
    l_varExceptionCodeCache(l_lngExceptionCodeCacheCount, 4) = l_strContractEarnCd
    l_varExceptionCodeCache(l_lngExceptionCodeCacheCount, 5) = l_strExecEarnCd
    l_varExceptionCodeCache(l_lngExceptionCodeCacheCount, 6) = l_str_AIC_ExcAcctNum
    l_varExceptionCodeCache(l_lngExceptionCodeCacheCount, 7) = l_str_AMS_ExcAcctNum
    l_varExceptionCodeCache(l_lngExceptionCodeCacheCount, 8) = l_str_UEC_ExcAcctNum
    l_lngExceptionCodeCacheCount = l_lngExceptionCodeCacheCount + 1
    SetExceptionCache = l_intPosition
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SetExceptionCache"
            
End Function


Private Function GetExceptionContractEarnCd( _
            ByVal v_strExcCd As String _
            ) As String
            
    Dim l_strEarningsCd As String
    Dim l_intPosition As Integer
    
    On Error GoTo eh:
                
    If Trim(v_strExcCd) = "" Then
        GetExceptionContractEarnCd = ""
        Exit Function
    End If

    If IsStringInList2(v_strExcCd, l_varExceptionCodeCache, l_lngExceptionCodeCacheCount, l_intPosition) Then
        l_strEarningsCd = l_varExceptionCodeCache(l_intPosition, 4)
    Else
        l_intPosition = SetExceptionCache(v_strExcCd)
        l_strEarningsCd = l_varExceptionCodeCache(l_intPosition, 4)
    End If
        
    GetExceptionContractEarnCd = l_strEarningsCd
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetExceptionContractEarnCd"
            
End Function

Private Function GetExceptionExecutiveEarnCd( _
            ByVal v_strExcCd As String _
            ) As String
            
    Dim l_strEarningsCd As String
    Dim l_intPosition As Integer
    
    On Error GoTo eh:
                
    If Trim(v_strExcCd) = "" Then
        GetExceptionExecutiveEarnCd = ""
        Exit Function
    End If

    If IsStringInList2(v_strExcCd, l_varExceptionCodeCache, l_lngExceptionCodeCacheCount, l_intPosition) Then
        l_strEarningsCd = l_varExceptionCodeCache(l_intPosition, 5)
    Else
        l_intPosition = SetExceptionCache(v_strExcCd)
        l_strEarningsCd = l_varExceptionCodeCache(l_intPosition, 5)
    End If
        
    GetExceptionExecutiveEarnCd = l_strEarningsCd
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetExceptionExecutiveEarnCd"
            
End Function

Private Function IsExecutiveDepartment( _
            ByVal v_strDept As String _
            ) As Boolean
            
    On Error GoTo IsExecutiveDepartment_EH
    
    If Left(v_strDept, 1) >= "5" Then
        IsExecutiveDepartment = True
    Else
        IsExecutiveDepartment = False
    End If
    
    Exit Function
    
IsExecutiveDepartment_EH:
    TrsRaiseError m_strCMODULENAME, "IsExecutiveDepartment"
            
    
End Function

Private Function IsDOJMDepartmentNumber( _
            ByVal v_strDeptNum _
            ) As Boolean
           
    On Error GoTo IsDOJMDepartmentNumber_EH
    
    '//If GetInterfaceCode(v_strDeptNum) = "D" Then
    If DoesRuleApply(1001, 4, v_strDeptNum) Then
        IsDOJMDepartmentNumber = True
    Else
        IsDOJMDepartmentNumber = False
    End If
    Exit Function
    
IsDOJMDepartmentNumber_EH:
    TrsRaiseError m_strCMODULENAME, "IsDOJMDepartmentNumber"
           
End Function

Private Function IsDOJMWorkGroupPrivate( _
            ByVal v_lngWgNum As Long _
            ) As Boolean
    
On Error GoTo IsDOJMWorkGroupPrivate_EH

    Select Case v_lngWgNum
        Case 100 To 999
            'UE DOJM
            IsDOJMWorkGroupPrivate = True
        Case 1000 To 1899
            'CIPS DOJM
            IsDOJMWorkGroupPrivate = True
        Case 1900 To 1999
            'CIPS DOJM Reserved
            IsDOJMWorkGroupPrivate = True
        Case 2000 To 2899
            'CILCO DOJM
            IsDOJMWorkGroupPrivate = True
        Case Else
           IsDOJMWorkGroupPrivate = False
    End Select
            
Exit Function

IsDOJMWorkGroupPrivate_EH:

    TrsRaiseError m_strCMODULENAME, "IsDOJMWorkGroupPrivate"

End Function

Private Function DetermineDistributionDepartment( _
            ByVal v_strPayrollDept As String, _
            ByVal v_lngChargeWGNum As Long _
            ) As String
            
 'The Payroll Distribution Department (time charge to)
 ' is the same as the Payroll Department (check picked up from)
 ' , except for the following
 '      1. Employee's payroll department is 086
 '      2. Employee's payroll department is a DOJM department
 '      3. Employee's time is charge to a DOJM work group
 ' If an employee meets one of the conditions above,
 '  then the Payroll Distribution Department is derived
 '  by the work group the time is being charged to.
 
    On Error GoTo DetermineDistributionDepartment_EH
    
    If IsDOJMDepartmentNumber(v_strPayrollDept) Or _
        IsDOJMWorkGroupPrivate(v_lngChargeWGNum) Or _
        v_strPayrollDept = "086" Then
        DetermineDistributionDepartment = GetDepartmentOfWorkGroup(v_lngChargeWGNum)
    Else
        DetermineDistributionDepartment = v_strPayrollDept
        Exit Function
    End If

    Exit Function
    
DetermineDistributionDepartment_EH:

    TrsRaiseError m_strCMODULENAME, "DetermineDistributionDepartment"
   
End Function


Private Sub ValidateDistributionIDRemoveOtherPay( _
                ByRef r_rsOtherPayEntries As ADODB.Recordset, _
                ByRef r_rsBRVs As ADODB.Recordset)
                
    Dim l_strDistIDType As String
    Dim l_bolError As String
                
    On Error GoTo ValidateDistributionIDRemoveOtherPay_EH
    
    If RecordsExist(r_rsOtherPayEntries) Then
     'Filter non TTA's and sort
        r_rsOtherPayEntries.Filter = MDAC25FilterFix("RECORD_STATUS <> 'D' AND EXCEPTION_CD <> 'TTA'")
        If r_rsOtherPayEntries.EOF Then
            Exit Sub
        End If
        r_rsOtherPayEntries.Sort = "KEY_EMP_NUM, KEY_CHARGE_DATE, KEY_START_TIME, END_TIME"
                        
     'Loop Thru each entry checking for distribution
        r_rsOtherPayEntries.MoveFirst
        While Not r_rsOtherPayEntries.EOF
            'Must check distribution type because a description might be there
            l_strDistIDType = GetDistributionType(r_rsOtherPayEntries)
            If Trim(l_strDistIDType) <> "" Then
                l_bolError = WriteBRV(r_rsBRVs, r_rsOtherPayEntries, 20000, "BW", "Distribution has been removed.  All other pay entries, except for TTAs, can not have distribution", True)
                r_rsOtherPayEntries!DISTRIBUTION_ID = ""
                r_rsOtherPayEntries!DISTRIBUTION_TYPE = ""
            End If
            r_rsOtherPayEntries.MoveNext
        Wend
    End If
    Exit Sub
    
    
ValidateDistributionIDRemoveOtherPay_EH:
    
                
End Sub

Private Function ModifyDistrList( _
    ByVal v_objRsPickListDistribution As ADODB.Recordset, _
    ByVal v_lngWgNum As Long, _
    ByVal v_blnRMC As Boolean, _
    ByVal v_blnRT As Boolean, _
    ByVal v_blnFMC As Boolean, _
    ByVal v_blnACTV As Boolean, _
    Optional ByVal v_blnFeederRef As Boolean = False _
    ) As ADODB.Recordset
    
    Dim l_rsList As ADODB.Recordset
    Dim l_strDistrId As String
    
    On Error GoTo eh
    
    If Not RecordsExist(v_objRsPickListDistribution) Then
        Set ModifyDistrList = Nothing
        Exit Function
    End If
    
    v_objRsPickListDistribution.Filter = "WG_NUM = " & str(v_lngWgNum)
    
    If Not RecordsExist(v_objRsPickListDistribution) Then
        v_objRsPickListDistribution.Filter = ""
        Set ModifyDistrList = Nothing
        Exit Function
    End If
    
    Set l_rsList = CopyRecordset(v_objRsPickListDistribution)
    l_rsList.MoveFirst
    While Not l_rsList.EOF
        l_strDistrId = GetAccountDescRemoveFields(l_rsList("DISTRIBUTION_ID").Value, v_blnRMC, v_blnRT, v_blnFMC, v_blnACTV, v_blnFeederRef)
        
'        If Len(l_strDistrId) > l_rsList("DISTRIBUTION_ID").DefinedSize Then
'            Dim s As Collection
'            Set s = New Collection
'
'            s.Add "v_lngWgNum => " & v_lngWgNum
'            s.Add "l_rsList(""DESCRIPTION"").Value =>" & l_rsList("DESCRIPTION").Value
'            s.Add "l_rsList(""DISTRIBUTION_ID"").Value =>" & l_rsList("DISTRIBUTION_ID").Value
'            s.Add "GetAccountDescRemoveFields => ***" & l_strDistrId & "***"
'            s.Add "Len(l_strDistrId) (source) => " & Len(l_strDistrId)
'            s.Add "l_rsList(""DISTRIBUTION_ID"").DefinedSize (target) =>" & l_rsList("DISTRIBUTION_ID").DefinedSize
'            '//s.Add "l_strDistrId2 => ***" & Left(l_strDistrId, l_rsList("DISTRIBUTION_ID").DefinedSize) & "***"
'
'            TrsTraceError m_strCMODULENAME, "ModifyDistrList", 0, False, , s, True
'        End If
        
        l_rsList("DISTRIBUTION_ID").Value = Left(l_strDistrId, l_rsList("DISTRIBUTION_ID").DefinedSize)
        l_rsList.MoveNext
    Wend
    l_rsList.MoveFirst
    v_objRsPickListDistribution.Filter = ""
    
    Set ModifyDistrList = l_rsList
    
    Exit Function
    
eh:

    TrsRaiseError m_strCMODULENAME, "ModifyDistrList"
    
End Function

Private Function UseEMPRVJobTasks( _
            ByVal v_lngWgNum As Long _
            ) As Boolean
            
    Dim l_obj As dbWorkGroup.cdWorkGroup
    
    Dim l_intPosition As Integer
   
On Error GoTo eh:
    
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    UseEMPRVJobTasks = l_obj.IsJobTaskWorkgroup(v_lngWgNum)
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "UseEMPRVJobTasks"

            
End Function
            
Private Function IsMKSDept(ByVal v_strDept As String) As Boolean

    IsMKSDept = DoesRuleApply(198, 4, v_strDept)
    
End Function

Private Function IsMKSDeptEmployeeException(ByVal v_lngEmp As Long) As Boolean

    IsMKSDeptEmployeeException = DoesRuleApply(199, 2, v_lngEmp)
    
End Function
Private Function DoesRuleApply( _
    ByVal v_lngRuleNum As Long, _
    ByVal v_lngBusType As TRSRulesDataTypes, _
    ByVal v_strValue As String _
    ) As Boolean
       
    Dim l_obj As dbRules.cdRules
       
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    DoesRuleApply = l_obj.DoesRuleApply(v_lngRuleNum, v_lngBusType, v_strValue)
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "DoesRuleApply"
    
End Function

Private Function GetMKSCode( _
            ByVal v_strDeptNum As String _
            ) As String

'The MKSCode is a
' Y -- if it is an MKS dept
' N -- if it is NOT
'
' String was used instead of boolean because the cache array is an array of strings
'
    
    Dim l_strMKSCode As String
    Dim l_intPosition As Integer

On Error GoTo eh:

    If IsStringInList2(v_strDeptNum, l_varDepartmentCache, l_lngDepartmentCacheCount, l_intPosition) Then
        l_strMKSCode = l_varDepartmentCache(l_intPosition, 7)
    Else
        l_intPosition = SetDepartmentCache(v_strDeptNum)
        l_strMKSCode = l_varDepartmentCache(l_intPosition, 7)
    End If
    GetMKSCode = l_strMKSCode
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetMKSCode"
            

End Function


Private Sub MKSDeptsRemoveDistribution( _
            ByRef r_rsTimeEntries As ADODB.Recordset, _
            ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset, _
            ByRef r_rsBRVs As ADODB.Recordset, _
            Optional v_blnOtherPay As Boolean = False)

    Dim l_strDefaultAccounting As String
    Dim l_strDistID As String
    Dim l_bolError As Boolean
    Dim l_strMKSCode As String
    Dim l_strDept As String
    Dim l_lngEmp As Long
    Dim l_ChargeDate As Date
    
On Error GoTo eh:

    If RecordsExist(r_rsTimeEntries) Then
    Else
        Exit Sub
    End If
            
    TimeEntriesFilterAndSort r_rsTimeEntries
    If r_rsTimeEntries.EOF Then
        Exit Sub
    End If
        
    r_rsTimeEntries.MoveFirst
    While Not r_rsTimeEntries.EOF
        l_strDistID = GetDistributionID(r_rsTimeEntries)
        l_strDefaultAccounting = ValidateDistributionIDsGetDefaultAccount(r_rsTimeEntries, v_rsEmployeeDistIDInfo, v_blnOtherPay)
        l_strDept = ValidateDistIDGetPayrollDept( _
             r_rsTimeEntries!KEY_EMP_NUM, _
             r_rsTimeEntries!KEY_CHARGE_DATE, _
             v_rsEmployeeDistIDInfo)
        l_strMKSCode = GetMKSCode(l_strDept)
        l_lngEmp = r_rsTimeEntries!KEY_EMP_NUM
        l_ChargeDate = r_rsTimeEntries!KEY_CHARGE_DATE
        If l_strDefaultAccounting = "" Or l_strDefaultAccounting = "N/A" Then
            If l_ChargeDate >= m_datMKSEffDate _
               And l_strMKSCode = "Y" _
               And Not MKSEmployeeOverride(l_lngEmp, l_ChargeDate, l_strMKSCode) Then
            'Actually equals the description of Override
            'or exception code.  Since exception codes can default
            'I am not placing that logic here
                r_rsTimeEntries!DISTRIBUTION_ID = Space(39)
                r_rsTimeEntries!DISTRIBUTION_TYPE = Space(1)
                r_rsTimeEntries!JOB_NBR = Space(8)
                r_rsTimeEntries!JOBTASK_NBR = Space(4)
                r_rsTimeEntries!JOBTASK_STEP_NBR = 0
                r_rsTimeEntries!EMPRV_DATABASE_ID = Space(8)
                If Trim(DetermineDistributionType(l_strDistID)) <> "" Then
                    l_bolError = WriteBRV(r_rsBRVs, r_rsTimeEntries, 20000, "BW", "System removed the Distribution. Please enter your time to a task in MKS", v_blnOtherPay)
                End If
            End If
        End If
        r_rsTimeEntries.MoveNext
    Wend
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "MKSDeptsRemoveDistribution"
                
            
End Sub

Private Function MKSDetermineAcctType( _
            ByVal v_strDept As String, _
            ByVal v_lngEmp As Long, _
            ByVal v_ChargeDate As Date) As String

'Returns the Account type of the department unless it is a MKS Department
'The department table is not effective date driven
'Also some employees in MKS Departments may still need to enter distribution in TRIS
' because they are not required to enter time in MKS
'

    Dim l_strAccountType As String

    l_strAccountType = GetAcctType(v_strDept)
    If GetMKSCode(v_strDept) <> "Y" Then
        MKSDetermineAcctType = l_strAccountType
        Exit Function
    End If

    If v_ChargeDate < m_datMKSEffDate Then
        MKSDetermineAcctType = "V"
    Else
        If MKSEmployeeOverride(v_lngEmp, v_ChargeDate, "Y") Then
            MKSDetermineAcctType = "V"
        Else
            MKSDetermineAcctType = "E"
        End If
    End If
    

End Function

Private Sub GetMKSEffDate()

    Dim l_obj As dbRules.cdRules
    Dim l_rs As ADODB.Recordset
       
    On Error GoTo eh
    
    If Date > CDate("2010-07-01") Then
        m_datMKSEffDate = CDate("2010-03-01")
        Exit Sub
    End If
    Set l_obj = CreateObject("dbRules.cdRules")
    Set l_rs = l_obj.GetRule(200)
    
    m_datMKSEffDate = CDate(Trim(l_rs!RG_VALUE_CD))
    
    Set l_rs = Nothing
    Set l_obj = Nothing
    Exit Sub
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetMKSEffDate"
    
    
    
End Sub

Private Function MKSEmployeeOverride(l_lngEmp, l_ChargeDate, l_strMKSCode) As Boolean

    Dim l_strFilter As String

    On Error GoTo eh
    
    If l_ChargeDate < m_datMKSEffDate Or l_strMKSCode <> "Y" Then
        MKSEmployeeOverride = False
        Exit Function
    End If
    
    If m_rsMKSEmployeeOverrides Is Nothing Then
        GetMKSEmployeeOverrides
    End If

    If Not RecordsExist(m_rsMKSEmployeeOverrides) Then
        MKSEmployeeOverride = False
        Exit Function
    End If
    
    l_strFilter = "RG_VALUE_CD = " & str(l_lngEmp) & _
                  " AND EFF_START_DT <= " & str(l_ChargeDate) & _
                  " AND EFF_END_DT >= " & str(l_ChargeDate)
    m_rsMKSEmployeeOverrides.Filter = MDAC25FilterFix(l_strFilter)

    If RecordsExist(m_rsMKSEmployeeOverrides) Then
        MKSEmployeeOverride = True
    Else
        MKSEmployeeOverride = False
    End If
    m_rsMKSEmployeeOverrides.Filter = ""
    Exit Function
    
eh:
 
    TrsRaiseError m_strCMODULENAME, "MKSEmployeeOverride"

End Function



Private Sub GetMKSEmployeeOverrides()

    Dim l_obj As dbRules.cdRules
       
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    Set m_rsMKSEmployeeOverrides = l_obj.GetEffectiveDateRule(199)
    
    Set l_obj = Nothing
    Exit Sub
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetMKSEmployeeOverrides"
    
    
    
End Sub

Private Function AllowDefaultAccountingForDOJM(ByVal v_strExceptionCd As String) As Boolean

    Dim l_obj As dbRules.cdRules
    Dim l_blnAllow As Boolean
       
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    
    l_blnAllow = l_obj.DoesRuleApply(274, trExceptionCode, v_strExceptionCd)
    
    Set l_obj = Nothing
    AllowDefaultAccountingForDOJM = l_blnAllow
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "AllowDefaultAccountingForDOJM"
    
    

End Function

Private Sub SetChargeWGForDOJMDefaultAccounting(ByRef r_rsTimeEntries As ADODB.Recordset, ByVal v_rsEmployeeDistIDInfo As ADODB.Recordset)
    
    Dim l_strPayrollDept As String
    Dim l_strInterfaceCode As String
    Dim l_strExceptionCd As String
    Dim l_strOverExceptionCd As String
    Dim l_strWorkGroup As String
    Dim l_obj As dbRules.cdRules
       
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    
    If Not RecordsExist(r_rsTimeEntries) Then
        Exit Sub
    End If
    
    r_rsTimeEntries.MoveFirst
    While Not r_rsTimeEntries.EOF
        If r_rsTimeEntries!RECORD_STATUS <> "D" Then
            l_strPayrollDept = ValidateDistIDGetPayrollDept( _
                 r_rsTimeEntries!KEY_EMP_NUM, _
                 r_rsTimeEntries!KEY_CHARGE_DATE, _
                 v_rsEmployeeDistIDInfo)
            l_strInterfaceCode = GetInterfaceCode(l_strPayrollDept)
        
            If l_strInterfaceCode = "D" Then
                If IsNull(r_rsTimeEntries!EXCEPTION_CD) Or IsEmpty(r_rsTimeEntries!EXCEPTION_CD) Then
                    l_strExceptionCd = ""
                Else
                    l_strExceptionCd = Trim(r_rsTimeEntries!EXCEPTION_CD)
                End If
                If IsNull(r_rsTimeEntries!OVER_EXC_CD) Or IsEmpty(r_rsTimeEntries!OVER_EXC_CD) Then
                    l_strOverExceptionCd = ""
                Else
                    l_strOverExceptionCd = Trim(r_rsTimeEntries!OVER_EXC_CD)
                End If
                If AllowDefaultAccountingForDOJM(l_strOverExceptionCd) Or AllowDefaultAccountingForDOJM(l_strExceptionCd) Then
                    If l_strOverExceptionCd = "" Then
                        l_strWorkGroup = Trim(l_obj.GetRuleLookup(275, trExceptionCode, l_strExceptionCd))
                    Else
                        l_strWorkGroup = Trim(l_obj.GetRuleLookup(275, trExceptionCode, l_strOverExceptionCd))
                    End If
                    If l_strWorkGroup = "" Then
                        r_rsTimeEntries!WG_NUM = 21528
                    Else
                        If IsInWorkGroup(CLng(l_strWorkGroup), r_rsTimeEntries!KEY_EMP_NUM, r_rsTimeEntries!KEY_CHARGE_DATE) Then
                            r_rsTimeEntries!WG_NUM = CLng(l_strWorkGroup)
                        Else
                            If l_strOverExceptionCd = "" Then
                                r_rsTimeEntries!EXCEPTION_CD = "ST"
                            Else
                                r_rsTimeEntries!OVER_EXC_CD = ""
                            End If
                        End If
                    End If
                    r_rsTimeEntries.Update
                End If
            End If
        End If
        r_rsTimeEntries.MoveNext
    Wend
    
    Set l_obj = Nothing
    
    Exit Sub
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SetChargeWGForDOJMDefaultAccounting"
    

End Sub


Private Function IsInWorkGroup( _
    ByVal v_lngWgNum As Long, _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date) As Boolean

    Dim l_obj As dbWorkGroup.cdWorkGroup
    Dim l_rs As ADODB.Recordset
    Dim l_bln As Boolean
       
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    
    l_bln = False
    Set l_rs = l_obj.GetEmployeeWorkGroups(v_lngEmpNum, v_datDate, v_datDate)
    If Not l_rs.BOF And Not l_rs.EOF Then
        l_rs.MoveFirst
        While Not l_rs.EOF And Not l_bln
            If l_rs!WG_NUM = v_lngWgNum Then
                l_bln = True
            End If
            l_rs.MoveNext
        Wend
    End If
    
    IsInWorkGroup = l_bln
    Set l_rs = Nothing
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsInWorkGroup"
    
End Function

Private Function GetCompanyCode( _
            ByVal v_strDeptNum As String _
            ) As String
            
    Dim l_strCompany As String
    Dim l_intPosition As Integer
   
On Error GoTo eh:
    
    If IsStringInList2(v_strDeptNum, l_varDepartmentCache, l_lngDepartmentCacheCount, l_intPosition) Then
        l_strCompany = l_varDepartmentCache(l_intPosition, 8)
    Else
        l_intPosition = SetDepartmentCache(v_strDeptNum)
        l_strCompany = l_varDepartmentCache(l_intPosition, 8)
    End If
    GetCompanyCode = l_strCompany
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetCompanyCode"
            
End Function

