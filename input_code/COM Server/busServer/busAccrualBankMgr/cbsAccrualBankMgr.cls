VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cbsAccrualBankMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Private Const m_strCMODULENAME As String = "busAccrualBankMgr"
Private m_lngLastModEmpNum As Long
Private m_rsAuditTrail As ADODB.Recordset
Private m_rsPriorCalendarYearVacationRule As ADODB.Recordset
Private m_rsUnionsWithMay_April_VacYear As ADODB.Recordset

Private Function GetEmpAccrualBanks( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date _
            ) As ADODB.Recordset

On Error GoTo GetEmpAccrualBanks_EH

    Dim l_dbAccrual As dbAccrualBankMgr.cdAccrualBankMgr
    Set l_dbAccrual = CreateObject("dbAccrualBankMgr.cdAccrualBankMgr")
    Set GetEmpAccrualBanks = l_dbAccrual.GetEmployeeBanks(v_aryEmployees, v_datStartDate)
    Set l_dbAccrual = Nothing
    
    Exit Function

GetEmpAccrualBanks_EH:
    Set l_dbAccrual = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmpAccrualBanks"
                             
End Function

Private Function GetEmpVacationBankAudit( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datStartDate As Date _
                ) As ADODB.Recordset

On Error GoTo eh

    Dim l_dbAccrual As dbAccrualBankMgr.cdAccrualBankMgr
    
    Set l_dbAccrual = CreateObject("dbAccrualBankMgr.cdAccrualBankMgr")
    
    Set GetEmpVacationBankAudit = l_dbAccrual.GetEmployeeVacationBankAudit(v_aryEmployees, v_datStartDate)
         
    Set l_dbAccrual = Nothing
    
    Exit Function

eh:
    Set l_dbAccrual = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmpVacationBankAudit"
End Function


Public Function UpdateBanksForTimeEntry( _
                ByVal v_request As ADODB.Recordset, _
                ByRef r_rsBRVs As ADODB.Recordset, _
                ByVal v_lngLastModEmpNum As Long _
                    ) As Boolean
' r_rsBRVs are not currently used, but may be used for future warnings.
' must set r_rsBRVs to nothing if an error occurs, i.e. discarding warnings.
' Set r_rsBRVs to nothing for now, but if warnings are desired; create
' a recordset for output only.
    Dim l_rsForamattedRequest As ADODB.Recordset
    Dim objCtx As ObjectContext
         
    On Error GoTo UpdateBanksForTimeEntry_EH
    
    Set r_rsBRVs = Nothing
    m_lngLastModEmpNum = v_lngLastModEmpNum
    
    Set objCtx = GetObjectContext
    
    If v_request.EOF And v_request.BOF Then
        objCtx.SetComplete
        Set objCtx = Nothing
        UpdateBanksForTimeEntry = True
        Exit Function
    End If
    
    Set l_rsForamattedRequest = BuildRequestForTimeEntry(v_request)
    
    If l_rsForamattedRequest.EOF And l_rsForamattedRequest.BOF Then
        objCtx.SetComplete
        Set objCtx = Nothing
        UpdateBanksForTimeEntry = True
        Set l_rsForamattedRequest = Nothing
        Exit Function
    End If
    
    ProcessBankUpdateRequests l_rsForamattedRequest, r_rsBRVs
    objCtx.SetComplete
    Set objCtx = Nothing
    UpdateBanksForTimeEntry = True
    Set l_rsForamattedRequest = Nothing
    Exit Function
    
UpdateBanksForTimeEntry_EH:
    objCtx.SetAbort
    Set objCtx = Nothing
    Set r_rsBRVs = Nothing
    Set l_rsForamattedRequest = Nothing
    TrsRaiseError m_strCMODULENAME, "UpdateBanksForTimeEntry", , , , , True
                         
End Function


Public Function UpdateBanksForCasacade( _
                ByVal v_request As ADODB.Recordset, _
                ByRef r_rsBRVs As ADODB.Recordset, _
                ByVal v_lngLastModEmpNum As Long) As Boolean
' Set r_rsBRVs to nothing for now, but if warnings are desired; create
' a recordset for output only.
    Dim l_rsForamattedRequest As ADODB.Recordset
    Dim objCtx As ObjectContext
         
On Error GoTo eh
    
    Set r_rsBRVs = Nothing
    m_lngLastModEmpNum = v_lngLastModEmpNum
    
    Set objCtx = GetObjectContext
    
    If v_request.EOF And v_request.BOF Then
        objCtx.SetComplete
        Set objCtx = Nothing
        UpdateBanksForCasacade = True
        Exit Function
    End If
    
    Set l_rsForamattedRequest = BuildRequestForCascade(v_request)
    If l_rsForamattedRequest.EOF And l_rsForamattedRequest.BOF Then
        objCtx.SetComplete
        Set objCtx = Nothing
        UpdateBanksForCasacade = True
        Set l_rsForamattedRequest = Nothing
        Exit Function
    End If
    
    ProcessBankUpdateRequests l_rsForamattedRequest, r_rsBRVs
    objCtx.SetComplete
    Set objCtx = Nothing
    UpdateBanksForCasacade = True
    Set l_rsForamattedRequest = Nothing
    Exit Function
    
eh:
    objCtx.SetAbort
    Set objCtx = Nothing
    Set r_rsBRVs = Nothing
    Set l_rsForamattedRequest = Nothing
    TrsRaiseError m_strCMODULENAME, "UpdateBanksForCasacade", , , , , True
End Function


Private Function BuildRequestForCascade( _
                    ByVal v_rsCascadeRequest As ADODB.Recordset _
                    ) As ADODB.Recordset
' Only expect deletes so ignore anything else.
    Dim l_rsRequest As ADODB.Recordset
    Dim l_strRecStatus As String
    Dim l_strBankCode As String
    Dim l_varEmpNums As Variant
    Dim l_cdEmployee As dbEmployee.cdEmployee
    Dim l_rsEmpBenPkg As ADODB.Recordset
    Dim l_cbsEmployee As busEmployee.cbsEmployee
    Dim l_strEmpBenPkg As String
    Dim l_strUnionCd As String
    
On Error GoTo eh

    Set l_rsRequest = CreateRequestRecordset()
    
    If Not AnyAccrualBankRelated(v_rsCascadeRequest) Then
        Set BuildRequestForCascade = l_rsRequest
        Exit Function
    End If
    
'    get employee records
    l_varEmpNums = DetermineGetEmployeeScope(v_rsCascadeRequest)
    Set l_cdEmployee = CreateObject("dbEmployee.cdEmployee")
    Set l_rsEmpBenPkg = l_cdEmployee.GetEmployeeBenefitsPackage(l_varEmpNums)
    Set l_cdEmployee = Nothing

    Set l_cbsEmployee = CreateObject("busEmployee.cbsEmployee")
    Do Until v_rsCascadeRequest.EOF
       
        l_rsEmpBenPkg.Filter = "EMP_NUM = " & v_rsCascadeRequest!EMP_NUM
        If l_cbsEmployee.IsExecutive( _
            l_rsEmpBenPkg!JOB_EFF_DATE, _
            l_rsEmpBenPkg!UNION_CD, _
            l_rsEmpBenPkg!PRIOR_UNION_CD, _
            l_rsEmpBenPkg!EXEC_CONTRACT_IND, _
            l_rsEmpBenPkg!PRIOR_EXEC_CON_IND, _
            v_rsCascadeRequest!CHARGE_DATE) Then
            l_strEmpBenPkg = "UE"
        Else
            If l_cbsEmployee.IsAmerenCILCOLocal8Employee( _
                l_rsEmpBenPkg!JOB_EFF_DATE, _
                l_rsEmpBenPkg!UNION_CD, _
                l_rsEmpBenPkg!PRIOR_UNION_CD, _
                v_rsCascadeRequest!CHARGE_DATE) Then
                l_strEmpBenPkg = "UE"
            Else
                l_strEmpBenPkg = l_cbsEmployee.EmployeeBenefitPackage( _
                    l_rsEmpBenPkg!JOB_EFF_DATE, l_rsEmpBenPkg!MGMT_BEN_PKG_IND, _
                    l_rsEmpBenPkg!UNION_CD, l_rsEmpBenPkg!PRIOR_UNION_CD, _
                    l_rsEmpBenPkg!EXEC_CONTRACT_IND, l_rsEmpBenPkg!PRIOR_EXEC_CON_IND, _
                    l_rsEmpBenPkg!UNION_BEN_PKG_IND, l_rsEmpBenPkg!PRIOR_UNION_BN_PKG, _
                    v_rsCascadeRequest!CHARGE_DATE)
            End If
        End If
            
        l_strRecStatus = Trim(v_rsCascadeRequest!REC_STATUS)
        
        l_strUnionCd = l_cbsEmployee.EmployeeUnion( _
            l_rsEmpBenPkg!JOB_EFF_DATE, _
            l_rsEmpBenPkg!UNION_CD, _
            l_rsEmpBenPkg!PRIOR_UNION_CD, _
            v_rsCascadeRequest!CHARGE_DATE)
        
        Select Case l_strRecStatus
            Case "D"
                DeleteRequestForCascade l_rsRequest, v_rsCascadeRequest, l_strEmpBenPkg, l_strUnionCd
            Case "I"
                InsertRequestForCascade l_rsRequest, v_rsCascadeRequest, l_strEmpBenPkg, l_strUnionCd
        End Select
        v_rsCascadeRequest.MoveNext
    Loop
    Set l_cbsEmployee = Nothing
    
    Set BuildRequestForCascade = l_rsRequest
    Set l_rsRequest = Nothing
    
    Exit Function
    
eh:
    Set l_rsRequest = Nothing
    Set l_cdEmployee = Nothing
    Set l_cbsEmployee = Nothing
    TrsRaiseError m_strCMODULENAME, "BuildRequestForCascade"
End Function


Private Function AnyAccrualBankRelated(ByRef r_rsRequest As ADODB.Recordset) As Boolean
    Dim l_strBankCode As String
    Dim l_strOverExcCd As String

On Error GoTo eh

    If r_rsRequest.EOF And r_rsRequest.BOF Then
        AnyAccrualBankRelated = False
        Exit Function
    End If
    
    r_rsRequest.MoveFirst
    Do Until r_rsRequest.EOF
        l_strBankCode = Trim(r_rsRequest!BANK_CODE)
        l_strOverExcCd = Trim(r_rsRequest!OVER_EXC_CD)
        If Trim(r_rsRequest!FMLA_IND) = "Y" Or l_strOverExcCd = "FCL" Or _
            (l_strBankCode <> "" And _
            l_strBankCode <> "N") Then
            AnyAccrualBankRelated = True
            Exit Do
        End If
        r_rsRequest.MoveNext
    Loop
    
    r_rsRequest.MoveFirst
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AnyAccrualBankRelated"
End Function


Private Function DetermineGetEmployeeScope( _
                    ByRef r_rsRequest As ADODB.Recordset) As Variant
    Dim l_colDistinctEmpNums As cuSet
    
On Error GoTo eh
'guard for empty recordset
    If r_rsRequest.BOF And r_rsRequest.EOF Then
        DetermineGetEmployeeScope = Empty
        Exit Function
    Else
        r_rsRequest.MoveFirst
    End If

    Set l_colDistinctEmpNums = New cuSet
    Do Until r_rsRequest.EOF
        l_colDistinctEmpNums.Add CLng(r_rsRequest!EMP_NUM.Value)
        r_rsRequest.MoveNext
    Loop
    
    DetermineGetEmployeeScope = ArrayFromCollection(l_colDistinctEmpNums.Contents)
    r_rsRequest.MoveFirst
    Set l_colDistinctEmpNums = Nothing
    Exit Function

eh:
    Set l_colDistinctEmpNums = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineGetEmployeeScope"
End Function


Private Sub DeleteRequestForCascade( _
                  ByRef r_rsRequest As ADODB.Recordset, _
                  ByRef r_rsCascadeRequest As ADODB.Recordset, _
                  ByVal v_strEmpBenPkg As String, _
                  ByVal v_strUnionCd As String)

    Dim l_strBankCode As String
    Dim l_strBankBucket As String
    Dim l_strOverExcCd As String

On Error GoTo eh

'    Build FMLA request
    If Trim(r_rsCascadeRequest!FMLA_IND) = "Y" Then
        With r_rsRequest
            .AddNew
            !REC_STATUS = "D"
            !EMP_NUM = r_rsCascadeRequest!EMP_NUM
            !CHARGE_DATE = r_rsCascadeRequest!CHARGE_DATE
            !BENEFITS_PACKAGE = v_strEmpBenPkg
            !UNION_CD = v_strUnionCd
            !EXCEPTION_CD = r_rsCascadeRequest!EXCEPTION_CD
            !OVER_EXC_CD = r_rsCascadeRequest!OVER_EXC_CD
            !HOURS = CDec(r_rsCascadeRequest!HOURS) * -1
            !BANK_CODE = "F"
            !BANK_BUCKET = "S"
            .Update
        End With
    End If
   
    l_strBankCode = Trim(r_rsCascadeRequest!BANK_CODE)
    l_strOverExcCd = Trim(r_rsCascadeRequest!OVER_EXC_CD)
    If l_strOverExcCd = "FCL" Then
        l_strBankCode = "S"
    End If
    If l_strBankCode = "" Or l_strBankCode = "N" Then
        Exit Sub
    End If
    
'    Build Accrual bank request
    With r_rsRequest
        .AddNew
        !REC_STATUS = "D"
        !EMP_NUM = r_rsCascadeRequest!EMP_NUM
        !CHARGE_DATE = r_rsCascadeRequest!CHARGE_DATE
        !BENEFITS_PACKAGE = v_strEmpBenPkg
        !UNION_CD = v_strUnionCd
        !EXCEPTION_CD = r_rsCascadeRequest!EXCEPTION_CD
        !OVER_EXC_CD = r_rsCascadeRequest!OVER_EXC_CD
        !HOURS = CDec(r_rsCascadeRequest!HOURS) * -1
        !BANK_CODE = Left(l_strBankCode, 1)
        !BANK_BUCKET = DetermineBankBucket(l_strBankCode)
        .Update
    End With
    
    Exit Sub
           
eh:
    TrsRaiseError m_strCMODULENAME, "DeleteRequestForCascade"
End Sub


Private Function DetermineBankBucket( _
                ByVal v_strBankCode As String, _
                Optional ByVal v_strSchedInd As String = "Y") As String
'The Bank Bucket is converted to character to get a
' clearer undstanding of what us going on
'Delete Requests will be processed before insert requests
'Basically whatever increase the hours available will be done first
'On Inserts the Accrued Bucket (+) is incremented before the Taken bucket (-)
'On Deletes the Taken bucket (-) is decremented before the Accrued Bucket (+)
    Dim l_strBankBucket As String

On Error GoTo eh

    l_strBankBucket = Right(v_strBankCode, 1)
    If Left(v_strBankCode, 1) = "V" Or Left(v_strBankCode, 1) = "P" Or Left(v_strBankCode, 1) = "S" Then
        If v_strSchedInd = "Y" Then
            l_strBankBucket = "S" ' Scheduled
        Else
            l_strBankBucket = "U" ' Used
        End If
    Else
        If l_strBankBucket = "+" Then
            l_strBankBucket = "A" 'Accrued
        Else
            l_strBankBucket = "T" 'Taken
        End If
    End If
    DetermineBankBucket = l_strBankBucket
    Exit Function
           
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineBankBucket"
End Function


Private Sub ProcessBankUpdateRequests(ByVal v_rsRequest As ADODB.Recordset, _
                         ByRef r_rsBRVs As ADODB.Recordset)
    
    Dim l_colEmployees As Collection
    Dim l_intNumOfEmps As Integer
    Dim l_intI As Integer
    Dim l_lngEmployee As Long
    Dim l_datStartDate As Date
    Dim l_rsEmployeeBanks As ADODB.Recordset
    Dim l_strFilter As String
    Dim l_strBankCode As String
    Dim l_strBankBucket As String
    Dim l_datChargeDate As Date
    Dim l_aryEmployees() As String
    Dim l_rsEmpVacationBankAudit As ADODB.Recordset
    Dim l_rsEmpHolidayBankAudit As ADODB.Recordset
    
    On Error GoTo ProcessBankUpdateRequests_EH
    
    
    v_rsRequest.MoveFirst
    Set l_colEmployees = New Collection
    While Not v_rsRequest.EOF
        l_lngEmployee = v_rsRequest!EMP_NUM
        If EmployeeInList(l_lngEmployee, l_colEmployees) Then
        Else
            l_colEmployees.Add l_lngEmployee
        End If
        v_rsRequest.MoveNext
    Wend
    l_intNumOfEmps = l_colEmployees.Count
    ReDim l_aryEmployees(l_intNumOfEmps - 1) As String
    For l_intI = 1 To l_intNumOfEmps
        l_aryEmployees(l_intI - 1) = l_colEmployees.Item(l_intI)
    Next l_intI
    

    v_rsRequest.MoveFirst
    v_rsRequest.Sort = "CHARGE_DATE"
    l_datStartDate = v_rsRequest!CHARGE_DATE
    Set m_rsAuditTrail = CreateAuditTrailRecordeSet()
    Set l_rsEmployeeBanks = GetEmpAccrualBanks(l_aryEmployees, l_datStartDate)
    Set l_rsEmpHolidayBankAudit = GetEmpHolidayBankAudit(l_aryEmployees, l_datStartDate)
'    If there are any delete requests for Vacation in the first week of year,
'    get the audit records to determine the year the bank was charged.
    If AnyDeletedVacationInFirstWeek(v_rsRequest) Then
        Set l_rsEmpVacationBankAudit = GetEmpVacationBankAudit(l_aryEmployees, l_datStartDate)
    End If
    l_intI = 0
    While l_intI < l_intNumOfEmps
        l_intI = l_intI + 1
        l_lngEmployee = l_colEmployees.Item(l_intI)
        l_strFilter = ""
        v_rsRequest.Filter = "EMP_NUM = " & str(l_lngEmployee)
        v_rsRequest.Sort = "CHARGE_DATE"
        v_rsRequest.MoveFirst
        l_datStartDate = v_rsRequest!CHARGE_DATE
        ProcessFMLARequests l_lngEmployee, v_rsRequest, l_rsEmployeeBanks, r_rsBRVs
        ProcessVacationRequests l_lngEmployee, v_rsRequest, l_rsEmployeeBanks, r_rsBRVs, l_rsEmpVacationBankAudit
        ProcessHolidayRequests l_lngEmployee, v_rsRequest, l_rsEmployeeBanks, r_rsBRVs, l_rsEmpHolidayBankAudit
        ProcessCompTimeRequests l_lngEmployee, v_rsRequest, l_rsEmployeeBanks, r_rsBRVs
        ProcessParentalBoundingRequests l_lngEmployee, v_rsRequest, l_rsEmployeeBanks, r_rsBRVs
        ProcessFamilyCareRequests l_lngEmployee, v_rsRequest, l_rsEmployeeBanks, r_rsBRVs
    Wend
    UpdateEmployeeAccrualBanks l_rsEmployeeBanks
    Set l_colEmployees = Nothing
    Set l_rsEmployeeBanks = Nothing
    Exit Sub
    
ProcessBankUpdateRequests_EH:
    
    Set l_colEmployees = Nothing
    Set l_rsEmployeeBanks = Nothing
    TrsRaiseError m_strCMODULENAME, "ProcessBankUpdateRequests"
    
End Sub
 
Private Function AnyDeletedVacationInFirstWeek( _
                        ByRef r_rsRequest As ADODB.Recordset) As Boolean
'    Answer True if there are any delete requests(HOURS < 0) for Vacation in the
'    first week of year that were eligible to be carried over from the prior year.
'    Ensure only scheduled or taken considered, not posting scheduled to used;
'    i.e. l_strBucket=B.
    Dim l_strFilter As String
    Dim l_datChargeDate As Date
    Dim l_strBenefitsPackage As String
    Dim l_strBucket As String
    Dim l_strUnionCd As String
    
On Error GoTo eh

    l_strFilter = "BANK_CODE = 'V'"
    r_rsRequest.Filter = MDAC25FilterFix(l_strFilter)
    r_rsRequest.Sort = ""
    Do Until r_rsRequest.EOF
        l_strBucket = UCase(Trim(r_rsRequest!BANK_BUCKET))
        If r_rsRequest!HOURS < 0 And (l_strBucket = "U" Or l_strBucket = "S") Then
            l_strBenefitsPackage = Trim(r_rsRequest!BENEFITS_PACKAGE)
            l_strUnionCd = UCase(Trim(r_rsRequest!BENEFITS_PACKAGE))
            l_datChargeDate = r_rsRequest!CHARGE_DATE
            If CanEmployeeTakePriorYearVacation(l_datChargeDate, l_strBenefitsPackage, l_strUnionCd) Then
                AnyDeletedVacationInFirstWeek = True
                Exit Do
            End If
        End If
        r_rsRequest.MoveNext
    Loop

    r_rsRequest.Filter = adFilterNone
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AnyDeletedVacationInFirstWeek"
End Function

 
 
Private Function EmployeeInList( _
                    ByVal v_lngEmployee As Long, _
                    ByVal v_colEmployees As Collection _
                ) As Boolean


    Dim l_intIndex As Integer
    Dim l_intCount As Integer
    Dim l_bolInList As Boolean
    
    l_bolInList = False
    l_intIndex = 1
    l_intCount = v_colEmployees.Count
    
    While Not l_bolInList And l_intIndex <= l_intCount
        If v_colEmployees.Item(l_intIndex) = v_lngEmployee Then
            l_bolInList = True
        End If
        l_intIndex = l_intIndex + 1
    Wend
    
    EmployeeInList = l_bolInList
    


End Function

Private Sub ProcessCompTimeRequests(ByVal v_lngEmployee As Long, _
                    ByVal v_rsRequest As ADODB.Recordset, _
                    ByRef r_rsEmpBanks As ADODB.Recordset, _
                    ByRef r_rsBRVs As ADODB.Recordset)
                    
    Dim l_datChargeDate As Date
    Dim l_strExceptionCode As String
    Dim l_decHours As Variant
    Dim l_strBucket As String
    Dim l_strBenefitsPackage As String
    Dim l_bolAccruedBucket As Boolean
    Dim l_bolTakenBucket As Boolean
    Dim l_strFilter As String
    Dim l_lngErrorNumber As Long
    
    On Error GoTo ProcessCompTimeRequests_EH
    
'    m_lngAppErrorInUpdateCompTime = False
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmployee) & " and BANK_CODE = 'C'"
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter)
    If v_rsRequest.EOF Then
        v_rsRequest.Filter = ""
        Exit Sub 'No Comp Time Banks to update
    End If
    
' The order in which the Comp Time Bank is update is important
' All delete tranasactions will occur before insert transactions
' With delete transactions the Taken buckets will be processed
'  before the Accrued Buckets
' With insert transactions the Accrued buckets will be processed
'  before the taken buckets

'Processing Deletes
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'D' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "BANK_BUCKET DESC, CHARGE_DATE DESC"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!EXCEPTION_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        l_strBenefitsPackage = v_rsRequest!BENEFITS_PACKAGE
        If l_strBucket = "A" Then
            l_bolAccruedBucket = True
            l_bolTakenBucket = False
        End If
        If l_strBucket = "T" Then
            l_bolAccruedBucket = False
            l_bolTakenBucket = True
        End If
        UpdateCompTimeBanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolTakenBucket, l_bolAccruedBucket, r_rsEmpBanks, r_rsBRVs, l_strBenefitsPackage
        v_rsRequest.MoveNext
    Wend
    
'Processing Inserts
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'I' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "CHARGE_DATE"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!EXCEPTION_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        l_strBenefitsPackage = v_rsRequest!BENEFITS_PACKAGE
        If l_strBucket = "A" Then
            l_bolAccruedBucket = True
            l_bolTakenBucket = False
        End If
        If l_strBucket = "T" Then
            l_bolAccruedBucket = False
            l_bolTakenBucket = True
        End If
        UpdateCompTimeBanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolTakenBucket, l_bolAccruedBucket, r_rsEmpBanks, r_rsBRVs, l_strBenefitsPackage
        v_rsRequest.MoveNext
    Wend
    
    v_rsRequest.Filter = ""
    Exit Sub
    
ProcessCompTimeRequests_EH:
    TrsRaiseError m_strCMODULENAME, "ProcessCompTimeRequests"
                    
End Sub


Private Sub ProcessHolidayRequests(ByVal v_lngEmployee As Long, _
                    ByVal v_rsRequest As ADODB.Recordset, _
                    ByRef r_rsEmpBanks As ADODB.Recordset, _
                    ByRef r_rsBRVs As ADODB.Recordset, _
                    ByRef l_rsEmpHolidayBankAudit As ADODB.Recordset)
                    
    Dim l_datChargeDate As Date
    Dim l_strExceptionCode As String
    Dim l_decHours As Variant
    Dim l_strBucket As String
    Dim l_strScheduledInd As String
    Dim l_bolAccruedBucket As Boolean
    Dim l_bolTakenBucket As Boolean
    Dim l_strFilter As String
    Dim l_lngErrorNumber As Long
    
    On Error GoTo ProcessHolidayRequests_EH
    
'    m_lngAppErrorInUpdateHoliday = False
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmployee) & " and BANK_CODE = 'H'"
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter)
    If v_rsRequest.EOF Then
        v_rsRequest.Filter = ""
        Exit Sub 'No Holiday Banks to update
    End If
    
    
' The order in which the Holiday Bank is update is important
' All delete tranasactions will occur before insert transactions
' With delete transactions the Taken buckets will be processed
'  before the Accrued Buckets
' With insert transactions the Accrued buckets will be processed
'  before the taken buckets

'Processing Deletes
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'D' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "BANK_BUCKET DESC, CHARGE_DATE DESC"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!EXCEPTION_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        If l_strBucket = "A" Then
            l_bolAccruedBucket = True
            l_bolTakenBucket = False
        End If
        If l_strBucket = "T" Then
            l_bolAccruedBucket = False
            l_bolTakenBucket = True
        End If
        UpdateHolidayBanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolTakenBucket, l_bolAccruedBucket, r_rsEmpBanks, r_rsBRVs, l_rsEmpHolidayBankAudit
        v_rsRequest.MoveNext
    Wend
    
'Processing Inserts
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'I' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "BANK_BUCKET, CHARGE_DATE"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!EXCEPTION_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        If l_strBucket = "A" Then
            l_bolAccruedBucket = True
            l_bolTakenBucket = False
        End If
        If l_strBucket = "T" Then
            l_bolAccruedBucket = False
            l_bolTakenBucket = True
        End If
        UpdateHolidayBanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolTakenBucket, l_bolAccruedBucket, r_rsEmpBanks, r_rsBRVs, l_rsEmpHolidayBankAudit
        v_rsRequest.MoveNext
    Wend
    
    v_rsRequest.Filter = ""
    l_rsEmpHolidayBankAudit.Filter = ""
    Exit Sub
    
ProcessHolidayRequests_EH:
    TrsRaiseError m_strCMODULENAME, "ProcessHolidayRequests"
                    
End Sub
 
Private Sub ProcessFMLARequests(ByVal v_lngEmployee As Long, _
                    ByVal v_rsRequest As ADODB.Recordset, _
                    ByRef r_rsEmpBanks As ADODB.Recordset, _
                    ByRef r_rsBRVs As ADODB.Recordset)
    
    Dim l_datChargeDate As Date
    Dim l_strExceptionCode As String
    Dim l_decHours As Variant
    Dim l_strBucket As String
    Dim l_bolUsedBucket As Boolean
    Dim l_bolSchedBucket As Boolean
    Dim l_lngErrorNumber As Long
    Dim l_strFilter As String
    
    On Error GoTo ProcessFMLARequests_EH
    
'    m_lngAppErrorInUpdateFMLA = False
    l_strFilter = "EMP_NUM = " & str(v_lngEmployee) & " and BANK_CODE = 'F'"
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter)
    If v_rsRequest.EOF Then
        v_rsRequest.Filter = ""
        Exit Sub 'No FMLA Banks to update
    End If
'Processing Deletes
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'D' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "CHARGE_DATE DESC"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!EXCEPTION_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        If l_strBucket = "U" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = False
        End If
        If l_strBucket = "S" Then
            l_bolSchedBucket = True
            l_bolUsedBucket = False
        End If
        If l_strBucket = "B" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = True
        End If
        UpdateFMLABanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolUsedBucket, l_bolSchedBucket, r_rsEmpBanks, r_rsBRVs
        v_rsRequest.MoveNext
    Wend

'Processing Inserts
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'I' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "CHARGE_DATE"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!EXCEPTION_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        If l_strBucket = "U" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = False
        End If
        If l_strBucket = "S" Then
            l_bolSchedBucket = True
            l_bolUsedBucket = False
        End If
        If l_strBucket = "B" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = True
        End If
        UpdateFMLABanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolUsedBucket, l_bolSchedBucket, r_rsEmpBanks, r_rsBRVs
        v_rsRequest.MoveNext
    Wend
    
    v_rsRequest.Filter = ""
    Exit Sub
    
ProcessFMLARequests_EH:
    TrsRaiseError m_strCMODULENAME, "ProcessFMLARequests"
                    
End Sub

Private Sub ProcessVacationRequests(ByVal v_lngEmployee As Long, _
                    ByVal v_rsRequest As ADODB.Recordset, _
                    ByRef r_rsEmpBanks As ADODB.Recordset, _
                    ByRef r_rsBRVs As ADODB.Recordset, _
                    ByRef l_rsEmpVacationBankAudit As ADODB.Recordset)
    
    Dim l_datChargeDate As Date
    Dim l_strExceptionCode As String
    Dim l_decHours As Variant
    Dim l_strBucket As String
    Dim l_strBenefitsPackage As String
    Dim l_strUnionCd As String
    Dim l_bolUsedBucket As Boolean
    Dim l_bolSchedBucket As Boolean
    Dim l_lngErrorNumber As Long
    Dim l_strFilter As String
    
    On Error GoTo ProcessVacationRequests_EH
    
'    m_lngAppErrorInUpdateVacation = False
    l_strFilter = "EMP_NUM = " & str(v_lngEmployee) & " and BANK_CODE = 'V'"
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter)
    If v_rsRequest.EOF Then
        v_rsRequest.Filter = ""
        Exit Sub 'No Vacation Banks to update
    End If
    
    Set m_rsPriorCalendarYearVacationRule = GetCalendarYearVacationRule()
    Set m_rsUnionsWithMay_April_VacYear = GetUnionsWith_May_April_CalYear()
'Processing Deletes
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'D' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "CHARGE_DATE DESC"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!EXCEPTION_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        l_strBenefitsPackage = v_rsRequest!BENEFITS_PACKAGE
        l_strUnionCd = v_rsRequest!UNION_CD
        If l_strBucket = "U" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = False
        End If
        If l_strBucket = "S" Then
            l_bolSchedBucket = True
            l_bolUsedBucket = False
        End If
        If l_strBucket = "B" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = True
        End If
        UpdateVacationBanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolUsedBucket, l_bolSchedBucket, r_rsEmpBanks, r_rsBRVs, l_strBenefitsPackage, l_strUnionCd, l_rsEmpVacationBankAudit
        v_rsRequest.MoveNext
    Wend

'Processing Inserts
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'I' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "CHARGE_DATE"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!EXCEPTION_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        l_strBenefitsPackage = v_rsRequest!BENEFITS_PACKAGE
        l_strUnionCd = v_rsRequest!UNION_CD
        If l_strBucket = "U" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = False
        End If
        If l_strBucket = "S" Then
            l_bolSchedBucket = True
            l_bolUsedBucket = False
        End If
        If l_strBucket = "B" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = True
        End If
        UpdateVacationBanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolUsedBucket, l_bolSchedBucket, r_rsEmpBanks, r_rsBRVs, l_strBenefitsPackage, l_strUnionCd, l_rsEmpVacationBankAudit
        v_rsRequest.MoveNext
    Wend
    
    v_rsRequest.Filter = ""
    Exit Sub
    
ProcessVacationRequests_EH:

    TrsRaiseError m_strCMODULENAME, "ProcessVacationRequests"
                    
End Sub


Private Function BuildRequestForTimeEntry( _
                    ByVal v_TimeEntryRequest As ADODB.Recordset _
                    ) As ADODB.Recordset
                      
    Dim l_rsRequest As ADODB.Recordset
    Dim l_strTimeCardStatus As String
    Dim l_strBankCode As String
    Dim l_varEmpNums As Variant
    
    Set l_rsRequest = CreateRequestRecordset()
    
    If v_TimeEntryRequest.EOF And v_TimeEntryRequest.BOF Then
        Set BuildRequestForTimeEntry = l_rsRequest
        Exit Function
    End If
    
    v_TimeEntryRequest.MoveFirst
    
    While Not v_TimeEntryRequest.EOF
    
        l_strTimeCardStatus = v_TimeEntryRequest!REC_STATUS
        Select Case l_strTimeCardStatus
            Case "D"
                DeleteRequestForTimeEntry l_rsRequest, v_TimeEntryRequest
            Case "I"
                InsertRequestForTimeEntry l_rsRequest, v_TimeEntryRequest
            Case "C"
                InsertRequestForTimeEntry l_rsRequest, v_TimeEntryRequest
                DeleteRequestForTimeEntry l_rsRequest, v_TimeEntryRequest
        End Select
        v_TimeEntryRequest.MoveNext
    Wend
    
    Set BuildRequestForTimeEntry = l_rsRequest
    Set l_rsRequest = Nothing
    
End Function


Private Sub DeleteRequestForTimeEntry( _
                  ByRef r_rsRequest As ADODB.Recordset, _
                  ByVal v_TimeEntryRequest As ADODB.Recordset)
                  
    Dim l_strBankCode As String
    Dim l_strBankBucket As String
    Dim l_strOverExcCd As String
    
        
    If v_TimeEntryRequest!ORIG_FMLA_IND = "Y" Then
' Do not update FMLA bank if the FMLA information did not change
        If v_TimeEntryRequest!FMLA_IND = "Y" And _
             v_TimeEntryRequest!REC_STATUS = "C" And _
             v_TimeEntryRequest!HOURS = v_TimeEntryRequest!ORIG_HOURS Then
        Else
            r_rsRequest.AddNew
            r_rsRequest!REC_STATUS = "D"
            r_rsRequest!EMP_NUM = v_TimeEntryRequest!EMP_NUM
            r_rsRequest!CHARGE_DATE = v_TimeEntryRequest!CHARGE_DATE
            r_rsRequest!BENEFITS_PACKAGE = v_TimeEntryRequest!BENEFITS_PACKAGE
            r_rsRequest!UNION_CD = v_TimeEntryRequest!UNION_CD
            r_rsRequest!EXCEPTION_CD = v_TimeEntryRequest!ORIG_EXCEPTION_CD
            r_rsRequest!OVER_EXC_CD = v_TimeEntryRequest!ORIG_OVER_EXC_CD
            r_rsRequest!HOURS = CDec(v_TimeEntryRequest!ORIG_HOURS) * -1
            r_rsRequest!BANK_CODE = "F"
            If v_TimeEntryRequest!SCHEDULED_IND = "Y" Then
                l_strBankBucket = "S" ' Scheduled
            Else
                l_strBankBucket = "U" ' Used
            End If
            r_rsRequest!BANK_BUCKET = l_strBankBucket
            r_rsRequest.Update
        End If
    End If
           
    l_strBankCode = Trim(v_TimeEntryRequest!ORIG_BANK_CODE)
        l_strOverExcCd = v_TimeEntryRequest!ORIG_OVER_EXC_CD
    If l_strOverExcCd = "FCL" Then
        l_strBankCode = "S"
    End If
    If l_strBankCode = "" Or l_strBankCode = "N" Then
        Exit Sub
    End If
    
    If v_TimeEntryRequest!EXCEPTION_CD = v_TimeEntryRequest!ORIG_EXCEPTION_CD And _
        v_TimeEntryRequest!OVER_EXC_CD = v_TimeEntryRequest!ORIG_OVER_EXC_CD And _
        v_TimeEntryRequest!REC_STATUS = "C" And _
        v_TimeEntryRequest!HOURS = v_TimeEntryRequest!ORIG_HOURS Then
'Delete is from an update and pertinent fields did not change
'Such as the job_number was changed
    Else
        r_rsRequest.AddNew
        r_rsRequest!REC_STATUS = "D"
        r_rsRequest!EMP_NUM = v_TimeEntryRequest!EMP_NUM
        r_rsRequest!CHARGE_DATE = v_TimeEntryRequest!CHARGE_DATE
        r_rsRequest!BENEFITS_PACKAGE = v_TimeEntryRequest!BENEFITS_PACKAGE
        r_rsRequest!EXCEPTION_CD = v_TimeEntryRequest!ORIG_EXCEPTION_CD
        r_rsRequest!OVER_EXC_CD = v_TimeEntryRequest!ORIG_OVER_EXC_CD
        r_rsRequest!HOURS = CDec(v_TimeEntryRequest!ORIG_HOURS) * -1
        r_rsRequest!BANK_CODE = Left(l_strBankCode, 1)
        l_strBankBucket = Right(l_strBankCode, 1)
'The Bank Bucket is converted to character to get a
' clearer undstanding of what us going on
'Delete Requests will be processed before insert requests
'Basically whatever increase the hours available will be done first
'On Inserts the Accrued Bucket (+) is incremented before the Taken bucket (-)
'On Deletes the Taken bucket (-) is decremented before the Accrued Bucket (+)
        If l_strBankBucket = "+" Then
            l_strBankBucket = "A" 'Accrued
        Else
            l_strBankBucket = "T" 'Taken
        End If
        If Left(l_strBankCode, 1) = "V" Or Left(l_strBankCode, 1) = "P" Or Left(l_strBankCode, 1) = "S" Then
            If v_TimeEntryRequest!SCHEDULED_IND = "Y" Then
                l_strBankBucket = "S" ' Scheduled
            Else
                l_strBankBucket = "U" ' Used
            End If
        End If
        r_rsRequest!BANK_BUCKET = l_strBankBucket
        r_rsRequest!UNION_CD = v_TimeEntryRequest!UNION_CD
        r_rsRequest.Update
    End If
           
End Sub

Private Sub InsertRequestForTimeEntry( _
                  ByRef r_rsRequest As ADODB.Recordset, _
                  ByVal v_TimeEntryRequest As ADODB.Recordset)
                  
    Dim l_strBankCode As String
    Dim l_strBankBucket As String
    Dim l_strStrScheduleInd As String
    Dim l_strOverExcCd As String
    
    l_strStrScheduleInd = v_TimeEntryRequest!SCHEDULED_IND
    l_strBankCode = UCase(Trim(v_TimeEntryRequest!BANK_CODE))
    
    If v_TimeEntryRequest!FMLA_IND = "Y" Then
' Do not update FMLA bank if the FMLA information did not change
        If v_TimeEntryRequest!ORIG_FMLA_IND = "Y" And _
            v_TimeEntryRequest!REC_STATUS = "C" And _
             v_TimeEntryRequest!HOURS = v_TimeEntryRequest!ORIG_HOURS Then
        Else
            r_rsRequest.AddNew
            r_rsRequest!REC_STATUS = "I"
            r_rsRequest!EMP_NUM = v_TimeEntryRequest!EMP_NUM
            r_rsRequest!CHARGE_DATE = v_TimeEntryRequest!CHARGE_DATE
            r_rsRequest!BENEFITS_PACKAGE = v_TimeEntryRequest!BENEFITS_PACKAGE
            r_rsRequest!UNION_CD = v_TimeEntryRequest!UNION_CD
            r_rsRequest!EXCEPTION_CD = v_TimeEntryRequest!EXCEPTION_CD
            r_rsRequest!OVER_EXC_CD = v_TimeEntryRequest!OVER_EXC_CD
            r_rsRequest!HOURS = CDec(v_TimeEntryRequest!HOURS)
            r_rsRequest!BANK_CODE = "F"
            If v_TimeEntryRequest!SCHEDULED_IND = "Y" Then
                l_strBankBucket = "B" ' Both used and scheduled
            Else
                l_strBankBucket = "U" ' Used
            End If
            r_rsRequest!BANK_BUCKET = l_strBankBucket
            r_rsRequest.Update
        End If
    End If
    
    l_strOverExcCd = v_TimeEntryRequest!OVER_EXC_CD
    If l_strOverExcCd = "FCL" Then
        l_strBankCode = "S"
    End If
    If l_strBankCode = "" Or l_strBankCode = "N" Then
        Exit Sub
    End If
    
    If l_strStrScheduleInd = "Y" Then
        If Left(l_strBankCode, 1) = "H" Or Left(l_strBankCode, 1) = "C" Then
            Exit Sub
        End If
    End If
    
    If v_TimeEntryRequest!EXCEPTION_CD = v_TimeEntryRequest!ORIG_EXCEPTION_CD And _
        v_TimeEntryRequest!OVER_EXC_CD = v_TimeEntryRequest!ORIG_OVER_EXC_CD And _
        v_TimeEntryRequest!REC_STATUS = "C" And _
        v_TimeEntryRequest!HOURS = v_TimeEntryRequest!ORIG_HOURS Then
'Insert is from an update and pertinent fields did not change
'Such as the job_number was changed
    Else
        r_rsRequest.AddNew
        r_rsRequest!REC_STATUS = "I"
        r_rsRequest!EMP_NUM = v_TimeEntryRequest!EMP_NUM
        r_rsRequest!CHARGE_DATE = v_TimeEntryRequest!CHARGE_DATE
        r_rsRequest!BENEFITS_PACKAGE = v_TimeEntryRequest!BENEFITS_PACKAGE
        r_rsRequest!EXCEPTION_CD = v_TimeEntryRequest!EXCEPTION_CD
        r_rsRequest!OVER_EXC_CD = v_TimeEntryRequest!OVER_EXC_CD
        r_rsRequest!HOURS = CDec(v_TimeEntryRequest!HOURS)
        r_rsRequest!BANK_CODE = Left(l_strBankCode, 1)
        l_strBankBucket = Right(l_strBankCode, 1)
'The Bank Bucket is converted to character to get a
' clearer undstanding of what us going on
'Delete Requests will be processed before insert requests
'Basically whatever increase the hours available will be done first
'On Inserts the Accrued Bucket (+) is incremented before the Taken bucket (-)
'On Deletes the Taken bucket (-) is decremented before the Accrued Bucket (+)
        If l_strBankBucket = "+" Then
            l_strBankBucket = "A" 'Accrued
        Else
            l_strBankBucket = "T" 'Taken
        End If
        If Left(l_strBankCode, 1) = "V" Or Left(l_strBankCode, 1) = "P" Or Left(l_strBankCode, 1) = "S" Then
            If v_TimeEntryRequest!SCHEDULED_IND = "Y" Then
                l_strBankBucket = "B" ' Both used and scheduled
            Else
                l_strBankBucket = "U" ' Used
            End If
        End If
        r_rsRequest!BANK_BUCKET = l_strBankBucket
        r_rsRequest!UNION_CD = v_TimeEntryRequest!UNION_CD
        r_rsRequest.Update
    End If

                      
End Sub

Private Sub AddToAuditTrail( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datDate As Date, _
                ByVal v_ExceptionCode As String, _
                ByVal v_Hours As Variant, _
                ByVal v_BankCode As String, _
                ByVal v_BankBucket As String, _
                ByVal v_BankYear As Integer, _
                ByVal v_Reason As String)
                
    m_rsAuditTrail.AddNew
    
    m_rsAuditTrail!EMP_NUM = v_lngEmpNum
    m_rsAuditTrail!CHARGE_DATE = v_datDate
    m_rsAuditTrail!EXCEPTION_CD = v_ExceptionCode
    m_rsAuditTrail!HOURS = v_Hours
    m_rsAuditTrail!BANK_CD = v_BankCode
    m_rsAuditTrail!BUCKET_IND_CD = v_BankBucket
    m_rsAuditTrail!EFF_YEAR = v_BankYear
    m_rsAuditTrail!REASON = v_Reason
    m_rsAuditTrail!LAST_MOD_EMP_NUM = m_lngLastModEmpNum
    
    m_rsAuditTrail.Update
    
End Sub


Private Function CreateAuditTrailRecordeSet() As ADODB.Recordset
    
    Dim l_rsAudit As ADODB.Recordset
    
    Set l_rsAudit = New ADODB.Recordset
    With l_rsAudit
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("HOURS").Precision = 3
      .Fields.Item("HOURS").NumericScale = 1
      .Fields.Append "BANK_CD", adChar, 1, adFldUpdatable
      .Fields.Append "BUCKET_IND_CD", adChar, 1, adFldUpdatable
      .Fields.Append "EFF_YEAR", adSmallInt, , adFldUpdatable
      .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "REASON", adChar, 30, adFldUpdatable
      
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockBatchOptimistic
      .Open
    End With
    
    Set CreateAuditTrailRecordeSet = l_rsAudit
    Set l_rsAudit = Nothing
    
End Function
Private Function CreateRequestRecordset() As ADODB.Recordset
    Dim l_rsRequest As ADODB.Recordset
    
    Set l_rsRequest = New ADODB.Recordset
    With l_rsRequest
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("HOURS").Precision = 3
      .Fields.Item("HOURS").NumericScale = 1
      .Fields.Append "BANK_CODE", adChar, 1, adFldUpdatable
      .Fields.Append "BANK_BUCKET", adChar, 1, adFldUpdatable
      .Fields.Append "BENEFITS_PACKAGE", adChar, 4, adFldUpdatable
      .Fields.Append "UNION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "REC_STATUS", adChar, 1, adFldUpdatable

      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockBatchOptimistic
      .Open
    End With
    
    Set CreateRequestRecordset = l_rsRequest
    Set l_rsRequest = Nothing

End Function
                      


Private Sub UpdateEmployeeAccrualBanks(ByVal v_rsEmpBanks As ADODB.Recordset)

    Dim l_dbAccrual As dbAccrualBankMgr.cdAccrualBankMgrTrans
    Dim l_bolSuccess As Boolean
    Dim l_lngEmployeeNumber As Long
    Dim l_colDescArgs As Collection
    
On Error GoTo UpdateEmployeeAccrualBanks_EH
    
    Set l_colDescArgs = New Collection
    Set l_dbAccrual = CreateObject("dbAccrualBankMgr.cdAccrualBankMgrTrans")
    
    l_bolSuccess = l_dbAccrual.UpdateEmployeeBanks(v_rsEmpBanks, m_rsAuditTrail, m_lngLastModEmpNum)
         
    If l_bolSuccess = False Then
        Err.Raise trErrbusSystemUnableToUpdateBanks 'app error *
    End If
    
    Set l_dbAccrual = Nothing
    
    Exit Sub

UpdateEmployeeAccrualBanks_EH:
    Set l_dbAccrual = Nothing
    If Not v_rsEmpBanks.EOF Then
        l_lngEmployeeNumber = v_rsEmpBanks!EMP_NUM
        l_colDescArgs.Add Trim(str(l_lngEmployeeNumber))
    Else
        l_colDescArgs.Add "Unknown"
    End If
     
    Select Case Err.Number
        Case trErrbusSystemUnableToUpdateBanks
            TrsRaiseError m_strCMODULENAME, "UpdateEmployeeAccrualBanks", Err.Number, , l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateEmployeeAccrualBanks"
    End Select

                             
End Sub

Private Sub UpdateVacationBanks( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datDate As Date, _
                ByVal v_decHours As Variant, _
                ByVal v_strExceptionCode As String, _
                ByVal v_bolUsedBucket As Boolean, _
                ByVal v_bolSchedBucket As Boolean, _
                ByRef r_rsBanks As ADODB.Recordset, _
                ByRef r_rsBRVs As ADODB.Recordset, _
                ByVal v_strBenefitsPackage As String, _
                ByVal v_strUnionCd As String, _
                ByRef l_rsEmpVacationBankAudit As ADODB.Recordset)
    
    Dim l_intBankYear As Integer
    Dim l_intPriorBankYear As Integer
    Dim l_decAllottedVacationhours As Variant
    Dim l_decUsedVacationhours As Variant
    Dim l_decSchedVacationhours As Variant
    Dim l_decAvailVacationhours As Variant
    Dim l_decPriorYearAllottedVacationhours As Variant
    Dim l_decPriorYearUsedVacationhours As Variant
    Dim l_decPriorYearSchedVacationhours As Variant
    Dim l_decPriorYearAvailVacationhours As Variant
    Dim l_colDescArgs As Collection
    Dim l_lngErrorNumber As Long
    Dim l_bolCanEmployeeTakePriorYearVacation As Boolean
    Dim l_bolCurrentYearError  As Boolean
    Dim l_bolBankBalanceError As Boolean
    Dim l_bolCurrentYearBankBalanceError As Boolean
    Dim l_decUpdateHours As Variant
    Dim l_decNewAvailHours As Variant
    Dim l_bolAuditFound As Boolean
    Dim l_decHoursToProcess As Variant
    Dim l_decHoursCharged As Variant
    Dim l_intExcBankYear As Integer
    Dim l_strErrorDesc As String
    Dim l_colErrArgs As Collection
    Dim l_intBankAuditYear As Integer
    Dim l_decTimeEntryHours As Variant
    Dim l_decAuditHours As Variant
    Dim l_bolMay_April_VacYear As Boolean
    
    On Error GoTo UpdateVacationBanks_EH
    
    If v_bolSchedBucket Or v_bolUsedBucket Then
    Else
        Exit Sub
    End If
    If v_decHours = 0 Then
        Exit Sub
    End If
    l_intBankYear = DetermineBankYear(v_strUnionCd, v_datDate)
    l_intPriorBankYear = l_intBankYear - 1
    
    l_bolCurrentYearError = True

    l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intBankYear, "V", r_rsBanks, l_decAllottedVacationhours, l_decAvailVacationhours, l_decUsedVacationhours, l_decSchedVacationhours)
    
    l_bolCanEmployeeTakePriorYearVacation = CanEmployeeTakePriorYearVacation(v_datDate, v_strBenefitsPackage, v_strUnionCd)

    If l_bolBankBalanceError Then
    'raise error current bank year not found
        If l_bolCanEmployeeTakePriorYearVacation Then
            l_bolCurrentYearBankBalanceError = True
        Else
            Err.Raise trErrbusEmployeeDoesNotHaveBank '1 app error
        End If
    End If

'Getting Prior Year Bank Balances if Needed
    If l_bolCanEmployeeTakePriorYearVacation Then
        l_bolCurrentYearError = False
        l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intPriorBankYear, "V", r_rsBanks, l_decPriorYearAllottedVacationhours, l_decPriorYearAvailVacationhours, l_decPriorYearUsedVacationhours, l_decPriorYearSchedVacationhours)
        If l_bolBankBalanceError And l_bolCurrentYearBankBalanceError Then
            Err.Raise trErrbusEmployeeDoesNotHaveBank
        End If
    Else
        l_decPriorYearUsedVacationhours = 0
        l_decPriorYearSchedVacationhours = 0
        l_decPriorYearAllottedVacationhours = 0
        l_decPriorYearAvailVacationhours = 0
    End If
    
'If No hours are available in the Prior year
'The employee can not TAKE prior year vacation.
'Skip this for DELETED vacation.
    If v_decHours > 0 Then
        If l_bolCanEmployeeTakePriorYearVacation Then
            If l_decPriorYearAvailVacationhours = 0 Then
                If v_bolUsedBucket And v_bolSchedBucket Then
                    If l_decPriorYearSchedVacationhours = 0 Then
                        l_bolCanEmployeeTakePriorYearVacation = False
                    End If
                Else
                    l_bolCanEmployeeTakePriorYearVacation = False
                End If
            End If
        End If
    End If
        
'Inserting a Taken or Scheduled Vacation
' Take remaining hours from prior year if allowed
    If v_decHours > 0 Then
        If l_bolCanEmployeeTakePriorYearVacation Then
            If v_bolUsedBucket And v_bolSchedBucket Then
            Else
                l_decPriorYearAvailVacationhours = l_decPriorYearAvailVacationhours - v_decHours
            End If
            If l_decPriorYearAvailVacationhours >= 0 Then
                'Update Prior Year vacation Bank
                UpdateVacationBanksSetBank v_lngEmpNum, l_intPriorBankYear, v_decHours, r_rsBanks, v_bolUsedBucket, v_bolSchedBucket, v_datDate, v_strExceptionCode
            Else 'Use remaining Prior Year Balance and then take from current year
                l_decUpdateHours = l_decPriorYearAvailVacationhours + v_decHours
                UpdateVacationBanksSetBank v_lngEmpNum, l_intPriorBankYear, l_decUpdateHours, r_rsBanks, v_bolUsedBucket, v_bolSchedBucket, v_datDate, v_strExceptionCode
                'Take Remaining from Current Year
                l_decUpdateHours = v_decHours - l_decUpdateHours
                If v_bolUsedBucket And v_bolSchedBucket Then
                Else
                    l_decAvailVacationhours = l_decAvailVacationhours - (l_decUpdateHours)
                End If
                If l_decAvailVacationhours < 0 Then
                'BRV unavailable Vacation hours *
                    Err.Raise trErrbusNotEnoughVacationHoursAvailable
                End If
                UpdateVacationBanksSetBank v_lngEmpNum, l_intBankYear, l_decUpdateHours, r_rsBanks, v_bolUsedBucket, v_bolSchedBucket, v_datDate, v_strExceptionCode
            End If
        Else ' Take vacation from current year
            If v_bolUsedBucket And v_bolSchedBucket Then
            Else
                l_decAvailVacationhours = l_decAvailVacationhours - v_decHours
            End If
            If l_decAvailVacationhours < 0 Then
                'BRV unavailable Vacation hours *
                Err.Raise trErrbusNotEnoughVacationHoursAvailable
            End If
            UpdateVacationBanksSetBank v_lngEmpNum, l_intBankYear, v_decHours, r_rsBanks, v_bolUsedBucket, v_bolSchedBucket, v_datDate, v_strExceptionCode
        End If
    End If
    
    If v_decHours > 0 Then
        Exit Sub
    End If
    
'Deleting a Taken or Scheduled Vacation
'Delete Hours from current year first,
'  then delete remaining from prior year if allowed
'    If v_decHours < 0 Then
'        l_decAvailVacationhours = l_decAvailVacationhours - v_decHours
'        If l_decAvailVacationhours > (l_decAllottedVacationhours) Then
'            'deduct the remainder hours from prior year if allowed
'            If l_bolCanEmployeeTakePriorYearVacation Then
'            ' Deduct as much as possible from current year
'                l_decUpdateHours = l_decAvailVacationhours - l_decAllottedVacationhours + v_decHours
'                If l_decUpdateHours < 0 Then
'                    UpdateVacationBanksSetBank v_lngEmpNum, l_intBankYear, v_decHours, r_rsBanks, v_bolUsedBucket, v_bolSchedBucket, v_datDate, v_strExceptionCode
'                End If
'            ' Deduct Remainder from prior year
'                l_decUpdateHours = l_decAllottedVacationhours - l_decAvailVacationhours
'                UpdateVacationBanksSetBank v_lngEmpNum, l_intPriorBankYear, l_decUpdateHours, r_rsBanks, v_bolUsedBucket, v_bolSchedBucket, v_datDate, v_strExceptionCode
'            Else
'                'out of balanace condition 'app error 1
'                Err.Raise trErrbusOutofBalanceVacationHours
'            End If
'        Else
'            'Deduct Hours from Current Year
'            UpdateVacationBanksSetBank v_lngEmpNum, l_intBankYear, v_decHours, r_rsBanks, v_bolUsedBucket, v_bolSchedBucket, v_datDate, v_strExceptionCode
'        End If
'    End If
'
'    Determine the year the deleted scheduled/taken hours should be decremented from.
'    If the charge date is in the first week of the year, check the audit table
'    for the year charged.  If an audit record is not found, assume prior year.  This
'    will enable this to work where 2002 is the prior year since 2003 banks were not
'    initialized before TRIS4 went into production.
'    Must handle condition where hours charged are split between years.
'    Raise an error if the sum of hours charged do not match v_decHours.
'    Use the existing logic for any charge date not in first week or not
'    allowed to use vacation in following year.


    If l_bolCanEmployeeTakePriorYearVacation Then
        l_bolAuditFound = FindAuditRecords( _
            v_lngEmpNum, v_datDate, v_bolUsedBucket, _
            v_bolSchedBucket, l_rsEmpVacationBankAudit)
        If l_bolAuditFound Then
    'All hours charged to current year
            l_intBankAuditYear = l_rsEmpVacationBankAudit!EFF_YEAR
            If l_intBankAuditYear = l_intBankYear Then
        'Set to Charged to current year only
                l_bolCanEmployeeTakePriorYearVacation = False
            End If
        Else
        'No audit follow old logic
        '-- Delete Hours from current year first,
        '     then delete remaining from prior year if allowed
            UpdateVacationBanksDeleteProcess _
                 v_lngEmpNum, v_datDate, v_decHours, v_strExceptionCode, _
                 v_bolUsedBucket, v_bolSchedBucket, r_rsBanks, r_rsBRVs, v_strBenefitsPackage, _
                 l_bolCanEmployeeTakePriorYearVacation, _
                 l_decAvailVacationhours, l_decAllottedVacationhours, v_strUnionCd
            Exit Sub
        End If
    End If
    
    If Not l_bolCanEmployeeTakePriorYearVacation Then
    'Charged to current year only
        UpdateVacationBanksDeleteProcess _
            v_lngEmpNum, v_datDate, v_decHours, v_strExceptionCode, _
            v_bolUsedBucket, v_bolSchedBucket, r_rsBanks, r_rsBRVs, v_strBenefitsPackage, _
            l_bolCanEmployeeTakePriorYearVacation, _
            l_decAvailVacationhours, l_decAllottedVacationhours, v_strUnionCd
        Exit Sub
    End If
    
'All or some of the hours were charged to the prior year

    l_rsEmpVacationBankAudit.Sort = "EFF_YEAR DESC"
    l_rsEmpVacationBankAudit.MoveFirst
    l_intBankAuditYear = l_rsEmpVacationBankAudit!EFF_YEAR
    If l_intBankAuditYear = l_intPriorBankYear Then
'All hours charged to prior year
        'Charge to prior year only
        UpdateVacationBanksDeleteProcess _
            v_lngEmpNum, v_datDate, v_decHours, v_strExceptionCode, _
            v_bolUsedBucket, v_bolSchedBucket, r_rsBanks, r_rsBRVs, v_strBenefitsPackage, _
            l_bolCanEmployeeTakePriorYearVacation, _
            l_decPriorYearAvailVacationhours, l_decPriorYearAllottedVacationhours, _
            v_strUnionCd, True
        Exit Sub
    End If
    
'Hours were charge to Current and Prior years on this charge date
'Give back to current year first
'Remember that the audit records were sorted by EFF_YEAR in
' descending sequence. The l_intBankAuditYear has already been set
'this is done just above this code

    l_decTimeEntryHours = v_decHours
    l_decHoursToProcess = Abs(l_decTimeEntryHours)
    Do Until l_rsEmpVacationBankAudit.EOF Or l_decTimeEntryHours >= 0
        l_decHoursCharged = CDec(l_rsEmpVacationBankAudit!HOURS)
        l_decHoursToProcess = l_decHoursToProcess - l_decHoursCharged
    'Charge to prior year
        If l_decHoursToProcess > 0 Then
            If l_intBankAuditYear = l_intPriorBankYear Then
                UpdateVacationBanksDeleteProcess _
                    v_lngEmpNum, v_datDate, l_decHoursCharged * -1, v_strExceptionCode, _
                    v_bolUsedBucket, v_bolSchedBucket, r_rsBanks, r_rsBRVs, v_strBenefitsPackage, _
                    l_bolCanEmployeeTakePriorYearVacation, _
                    l_decPriorYearAvailVacationhours, l_decPriorYearAllottedVacationhours, _
                    v_strUnionCd, True
            Else
                UpdateVacationBanksDeleteProcess _
                    v_lngEmpNum, v_datDate, l_decHoursCharged * -1, v_strExceptionCode, _
                    v_bolUsedBucket, v_bolSchedBucket, r_rsBanks, r_rsBRVs, v_strBenefitsPackage, _
                    l_bolCanEmployeeTakePriorYearVacation, _
                    l_decAvailVacationhours, l_decAllottedVacationhours, v_strUnionCd
            End If
            l_decTimeEntryHours = l_decTimeEntryHours + l_decHoursCharged
            l_rsEmpVacationBankAudit!HOURS = 0
            l_rsEmpVacationBankAudit.MoveNext
            If Not l_rsEmpVacationBankAudit.EOF Then
                l_intBankAuditYear = l_rsEmpVacationBankAudit!EFF_YEAR
            End If
        Else
            If l_intBankAuditYear = l_intPriorBankYear Then
                UpdateVacationBanksDeleteProcess _
                    v_lngEmpNum, v_datDate, l_decTimeEntryHours, v_strExceptionCode, _
                    v_bolUsedBucket, v_bolSchedBucket, r_rsBanks, r_rsBRVs, v_strBenefitsPackage, _
                    l_bolCanEmployeeTakePriorYearVacation, _
                    l_decPriorYearAvailVacationhours, l_decPriorYearAllottedVacationhours, _
                    v_strUnionCd, True
            Else
                UpdateVacationBanksDeleteProcess _
                    v_lngEmpNum, v_datDate, l_decTimeEntryHours, v_strExceptionCode, _
                    v_bolUsedBucket, v_bolSchedBucket, r_rsBanks, r_rsBRVs, v_strBenefitsPackage, _
                    l_bolCanEmployeeTakePriorYearVacation, _
                    l_decAvailVacationhours, l_decAllottedVacationhours, v_strUnionCd
            End If
            l_decAuditHours = l_rsEmpVacationBankAudit!HOURS
            l_decAuditHours = l_decAuditHours + l_decTimeEntryHours
            l_rsEmpVacationBankAudit!HOURS = l_decAuditHours
            l_decTimeEntryHours = 0
        End If
    Loop

    If l_decHoursToProcess > 0 Then
    'The total audit records do not match the entry.
    '-- Delete Hours from current year first,
    '     then delete remaining from prior year if allowed
        UpdateVacationBanksDeleteProcess _
            v_lngEmpNum, v_datDate, l_decHoursToProcess * -1, v_strExceptionCode, _
            v_bolUsedBucket, v_bolSchedBucket, r_rsBanks, r_rsBRVs, v_strBenefitsPackage, _
            l_bolCanEmployeeTakePriorYearVacation, _
            l_decAvailVacationhours, l_decAllottedVacationhours, v_strUnionCd
    End If

    Exit Sub

UpdateVacationBanks_EH:
    Set m_rsAuditTrail = Nothing
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "l_strErrorDesc: " & l_strErrorDesc
    l_colErrArgs.Add "v_lngEmpNum: " & CStr(v_lngEmpNum)
    l_colErrArgs.Add "v_datDate: " & Format(v_datDate, "mm/dd/yyyy")
    l_colErrArgs.Add "v_decHours: " & CStr(v_decHours)
    l_colErrArgs.Add "v_strExceptionCode: " & v_strExceptionCode
    l_colErrArgs.Add "v_bolUsedBucket: " & CStr(v_bolUsedBucket)
    l_colErrArgs.Add "v_bolSchedBucket: " & CStr(v_bolSchedBucket)
    l_colErrArgs.Add "v_strBenefitsPackage: " & v_strBenefitsPackage
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add CStr(v_lngEmpNum)
    If l_bolCurrentYearError Then
        l_colDescArgs.Add CStr(l_intBankYear)
    Else
        l_colDescArgs.Add CStr(l_intPriorBankYear)
    End If
    
    l_lngErrorNumber = Err.Number
    Select Case l_lngErrorNumber
        Case trErrbusNotEnoughVacationHoursAvailable
                Err.Source = FormatBRV(TrType_TC, v_lngEmpNum, v_datDate, v_datDate, "", "", v_decHours, v_strExceptionCode, "")
                TrsRaiseError m_strCMODULENAME, "UpdateVacationBanks", l_lngErrorNumber, , l_colDescArgs, l_colErrArgs
        Case trErrbusOutofBalanceVacationHours, _
             trErrbusEmployeeDoesNotHaveBank
                TrsRaiseError m_strCMODULENAME, "UpdateVacationBanks", l_lngErrorNumber, , l_colDescArgs, l_colErrArgs, True
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateVacationBanks"
    End Select
End Sub



Private Function FindAuditRecords( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datDate As Date, _
                    ByVal v_bolUsedBucket As Boolean, _
                    ByVal v_bolSchedBucket As Boolean, _
                    ByRef r_rsEmpVacationBankAudit As ADODB.Recordset) As Boolean

    Dim l_strBucket As String
    Dim l_blnFound As Boolean
    Dim l_strFilter As String

On Error GoTo eh
        
    If Not RecordsExist(r_rsEmpVacationBankAudit) Then
        Exit Function
    End If
    
    If v_bolUsedBucket And v_bolSchedBucket Then
        l_strBucket = "B"  'This condition will never be met
                           'This method is only called when an
                           ' exception is deleted.
                           'The exceotion can not be both used (reported time)
                           ' and scheduled
    ElseIf v_bolUsedBucket Then
        l_strBucket = "U"
    ElseIf v_bolSchedBucket Then
        l_strBucket = "S"
    End If
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) _
            & " AND CHARGE_DATE = " & v_datDate _
            & " AND BUCKET_IND_CD = '" & l_strBucket & "'" _
            & " AND HOURS > 0 "
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsEmpVacationBankAudit.Filter = l_strFilter
    If Not RecordsExist(r_rsEmpVacationBankAudit) Then
        Exit Function
    End If
    r_rsEmpVacationBankAudit.Sort = "EFF_YEAR"
    r_rsEmpVacationBankAudit.MoveFirst
    FindAuditRecords = True

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "FindAuditRecords"
End Function



Private Sub UpdateVacationBanksSetBank( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_intBankYear As Integer, _
                ByVal v_decHours As Variant, _
                ByRef r_rsBanks As ADODB.Recordset, _
                ByVal v_bolUsedBucket As Boolean, _
                ByVal v_bolSchedBucket As Boolean, _
                ByVal v_datDate As Date, _
                ByVal v_strExceptionCode As String)
    
    Dim l_decUsedHours As Variant
    Dim l_decSchedHours As Variant
    Dim l_decAllottedHours As Variant
    Dim l_colDescArgs As Collection
    Dim l_lngErrorNumber As Long
    
    On Error GoTo UpdateVacationBanksSetBank_EH
    
                
    r_rsBanks.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmpNum) & " and  EFF_YEAR = " & str(v_intBankYear))
    l_decAllottedHours = r_rsBanks!VAC_HRS_ALLOTTED
    l_decUsedHours = r_rsBanks!VAC_HRS_USED
    l_decSchedHours = r_rsBanks!VAC_HRS_SCHED
    If v_bolUsedBucket Then
        l_decUsedHours = l_decUsedHours + v_decHours
        If l_decUsedHours < 0 Then
            Err.Raise trErrbusOutofBalanceVacationHours 'app error 2
        End If
        r_rsBanks!VAC_HRS_USED = l_decUsedHours
        AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "V", "U", v_intBankYear, " "
        If v_bolSchedBucket Then
            l_decSchedHours = l_decSchedHours - v_decHours
            If l_decSchedHours < 0 Then
                Err.Raise trErrbusOutofBalanceVacationHours 'app error 3
            End If
            r_rsBanks!VAC_HRS_SCHED = l_decSchedHours
            AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours * -1, "V", "S", v_intBankYear, " "
        End If
    ' The condition below should not get executed because of the
    ' BRV not enough vacation hours available.  This line makes it bullet proof
        If (l_decUsedHours + l_decSchedHours) > l_decAllottedHours Then
            Err.Raise trErrbusOutofBalanceVacationHours 'app error 4
        End If
    Else
        l_decSchedHours = l_decSchedHours + v_decHours
        If l_decSchedHours < 0 Then
            Err.Raise trErrbusOutofBalanceVacationHours 'app error 5
        End If
    ' The condition below should not get executed because of the
    ' BRV not enough vacation hours available.  This line makes it bullet proof
        If (l_decUsedHours + l_decSchedHours) > l_decAllottedHours Then
            Err.Raise trErrbusOutofBalanceVacationHours 'app error 6
        End If
        r_rsBanks!VAC_HRS_SCHED = l_decSchedHours
        AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "V", "S", v_intBankYear, " "
    End If
    r_rsBanks.Update
    r_rsBanks.Filter = ""
    
    Exit Sub
    
UpdateVacationBanksSetBank_EH:

    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(v_lngEmpNum)
    l_colDescArgs.Add str(v_intBankYear)
    l_lngErrorNumber = Err.Number
    Select Case l_lngErrorNumber
        Case trErrbusOutofBalanceVacationHours
            TrsRaiseError m_strCMODULENAME, "UpdateVacationBanksSetBank", l_lngErrorNumber, , l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateVacationBanksSetBank"
    End Select
        
        
End Sub
                

Private Function GetBankBalances( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_intBankYear As Integer, _
                    ByVal v_strBankType As String, _
                    ByVal v_rsBanks As ADODB.Recordset, _
                    ByRef r_decAllottedHours As Variant, _
                    ByRef r_decAvailHours As Variant, _
                    ByRef r_decUsedHours As Variant, _
                    Optional ByRef r_decSchedHours As Variant, _
                    Optional ByRef r_dateBondingStart As Date _
                    ) As Boolean
                    
    Dim l_colDescArgs As Collection
    
    On Error GoTo GetBankBalances_EH
    
    v_rsBanks.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmpNum) & " and EFF_YEAR = " & str(v_intBankYear))
    If v_rsBanks.EOF And v_rsBanks.BOF Then
        r_decAllottedHours = 0
        r_decAvailHours = 0
        r_decUsedHours = 0
        If v_strBankType = "V" Or v_strBankType = "F" Or v_strBankType = "P" Or v_strBankType = "S" Then
            r_decSchedHours = 0
        End If
        v_rsBanks.Filter = ""
        GetBankBalances = True
        Exit Function
    End If
    Select Case v_strBankType
        Case "V" ' Vacation
            r_decAllottedHours = v_rsBanks!VAC_HRS_ALLOTTED
            r_decUsedHours = v_rsBanks!VAC_HRS_USED
            r_decSchedHours = v_rsBanks!VAC_HRS_SCHED
            r_decAvailHours = r_decAllottedHours - r_decUsedHours - r_decSchedHours
        Case "F" 'FMLA
            r_decAllottedHours = v_rsBanks!FMLA_HRS_ALLOWED
            r_decUsedHours = v_rsBanks!FMLA_HRS_USED
            r_decSchedHours = v_rsBanks!FMLA_HRS_SCHED
            r_decAvailHours = r_decAllottedHours - r_decUsedHours - r_decSchedHours
        Case "H" 'Holiday
            r_decAllottedHours = v_rsBanks!HOL_HRS_ACCRUED
            r_decUsedHours = v_rsBanks!HOL_TAKEN_SCHED
            r_decAvailHours = r_decAllottedHours - r_decUsedHours
        Case "C" 'CompTime
            r_decAllottedHours = v_rsBanks!COMP_STR_TIME_HRS
            r_decUsedHours = v_rsBanks!COMP_TAKEN_SCHED
            r_decAvailHours = r_decAllottedHours - r_decUsedHours
        Case "P" 'Parental Bounding
            r_decAllottedHours = v_rsBanks!PARENTAL_BOUNDING_HRS_ALLOWED
            r_decUsedHours = v_rsBanks!PARENTAL_BOUNDING_HRS_USED
            r_decSchedHours = v_rsBanks!PARENTAL_BOUNDING_HRS_SCHED
            r_decAvailHours = r_decAllottedHours - r_decUsedHours - r_decSchedHours
            r_dateBondingStart = v_rsBanks!PARENTAL_BOUNDING_START_DATE
        Case "S" 'Familiy Care
            r_decAllottedHours = v_rsBanks!FAMILY_CARE_HRS_ALLOWED
            r_decUsedHours = v_rsBanks!FAMILY_CARE_HRS_USED
            r_decSchedHours = v_rsBanks!FAMILY_CARE_HRS_SCHED
            r_decAvailHours = r_decAllottedHours - r_decUsedHours - r_decSchedHours
    End Select
    
'Raise Errors for Negative Balances 'app errors
    If r_decAllottedHours < 0 Then
        Err.Raise GetBadBankErrorNumber(v_strBankType, "A")
    End If
    If r_decAvailHours < 0 Then
        Err.Raise GetBadBankErrorNumber(v_strBankType, "V")
    End If
    If r_decUsedHours < 0 Then
        Err.Raise GetBadBankErrorNumber(v_strBankType, "U")
    End If
    If v_strBankType = "V" Or v_strBankType = "F" Or v_strBankType = "P" Or v_strBankType = "S" Then
        If r_decSchedHours < 0 Then
            Err.Raise GetBadBankErrorNumber(v_strBankType, "S")
        End If
    End If
    
    v_rsBanks.Filter = ""
    GetBankBalances = False
    Exit Function
    
GetBankBalances_EH:
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(v_lngEmpNum)
    l_colDescArgs.Add str(v_intBankYear)
    Select Case Err.Number
        Case trErrbusNegativeVacationAvailableBalance, _
             trErrbusNegativeVacationUsedBalance, _
             trErrbusNegativeVacationSchedBalance, _
             trErrbusNegativeFMLA_AvailableBalance, _
             trErrbusNegativeFMLA_UsedBalance, _
             trErrbusNegativeFMLA_SchedBalance, _
             trErrbusNegativeHolidayAccruedBalance, _
             trErrbusNegativeHolidayAvailableBalance, _
             trErrbusNegativeHolidayTakenBalance, _
             trErrbusNegativeCompTimeAccruedBalance, _
             trErrbusNegativeCompTimeAvailableBalance, _
             trErrbusNegativeCompTimeTakenBalance
            TrsRaiseError m_strCMODULENAME, "GetBankBalances", Err.Number, , l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetBankBalances"
    End Select
    
                
End Function
                    
Private Function GetBadBankErrorNumber( _
                    ByVal v_BankType As String, _
                    ByVal v_BankBucket) As Long

    Select Case v_BankType
        Case "V" 'Vaction
            Select Case v_BankBucket
                Case "V", "A" ' Might need to separted these, but do not feel like creating new error message
                    GetBadBankErrorNumber = trErrbusNegativeVacationAvailableBalance
                Case "U"
                    GetBadBankErrorNumber = trErrbusNegativeVacationUsedBalance
                Case "S"
                    GetBadBankErrorNumber = trErrbusNegativeVacationSchedBalance
            End Select
        Case "F" 'FMLA
            Select Case v_BankBucket
                Case "V", "A"  ' Might need to separted these, but do not feel like creating new error message
                    GetBadBankErrorNumber = trErrbusNegativeFMLA_AvailableBalance
                Case "U"
                    GetBadBankErrorNumber = trErrbusNegativeFMLA_UsedBalance
                Case "S"
                    GetBadBankErrorNumber = trErrbusNegativeFMLA_SchedBalance
            End Select
        Case "H" 'Holiday
            Select Case v_BankBucket
                Case "A"
                    GetBadBankErrorNumber = trErrbusNegativeHolidayAccruedBalance
                Case "V"
                    GetBadBankErrorNumber = trErrbusNegativeHolidayAvailableBalance
                Case "U"
                    GetBadBankErrorNumber = trErrbusNegativeHolidayTakenBalance
            End Select
        Case "C" 'CompTime
            Select Case v_BankBucket
                Case "A"
                    GetBadBankErrorNumber = trErrbusNegativeCompTimeAccruedBalance
                Case "V"
                    GetBadBankErrorNumber = trErrbusNegativeCompTimeAvailableBalance
                Case "U"
                    GetBadBankErrorNumber = trErrbusNegativeCompTimeTakenBalance
                
            End Select
         Case "P" 'Parental Bounding
            Select Case v_BankBucket
                Case "V", "A"  ' Might need to separted these, but do not feel like creating new error message
                    GetBadBankErrorNumber = trErrbusNegativeFMLA_AvailableBalance
                Case "U"
                    GetBadBankErrorNumber = trErrbusNegativeFMLA_UsedBalance
                Case "S"
                    GetBadBankErrorNumber = trErrbusNegativeFMLA_SchedBalance
            End Select
         Case "S" 'Family Care
            Select Case v_BankBucket
                Case "V", "A"  ' Might need to separted these, but do not feel like creating new error message
                    GetBadBankErrorNumber = trErrbusNegativeFMLA_AvailableBalance
                Case "U"
                    GetBadBankErrorNumber = trErrbusNegativeFMLA_UsedBalance
                Case "S"
                    GetBadBankErrorNumber = trErrbusNegativeFMLA_SchedBalance
            End Select
            
    End Select

End Function


Private Function CanEmployeeTakePriorYearVacation( _
                    ByVal v_datDate As Date, _
                    ByVal v_strBenefitsPackage As String, _
                    ByVal v_strUnionCd As String _
                    ) As Boolean

Dim WeekStartDay As Integer
                    
'Management and AmerenUE contract employees may use prior year vactaion hours in the first
'week of the new year.  The first week must have started in the prior year.
'For example these employees can use remaining 2003 vacation hours on
'   01/01/04, 01/02/04, and 01/03/04.  01/04/04 is a Sunday and starts the next week.
'Only 2004 vacation hours can be used from 01/04/04 through 12/31/04.
'All other employees take vacation based on Calendar year only.
'For example 2004 vacation hours can only be used from 01/01/04 through 12/31/04
                    
    
    Select Case UCase(Trim(v_strUnionCd))
        Case "ND"
            If BankYearCILCOLocal8(v_datDate) < 2007 Then
                CanEmployeeTakePriorYearVacation = False
                Exit Function
            End If
        Case Else
    End Select
    
    
'    If Feb - Dec, cannot take prior year
'    If Day = 7, may be first week of year, but only if year started on Sunday,
'    i.e. week started this year, not prior.
'    If Day > 7, must not be first week of year.
    If Month(v_datDate) > 1 Or Day(v_datDate) > 6 Then
        CanEmployeeTakePriorYearVacation = False
        Exit Function
    End If
    
           
     Select Case UCase(Trim(v_strUnionCd))
        Case "NT", "NE", "ND"
            WeekStartDay = vbMonday
        Case Else
            WeekStartDay = vbSunday
    End Select
    
'    The v_datDate is in the first week of year, but
'    did the week start in a prior year?
    If Weekday(v_datDate, WeekStartDay) > Day(v_datDate) Then
        If CanOnlyTakeVacationInCalendarYear(v_strBenefitsPackage) Then
            CanEmployeeTakePriorYearVacation = False
        Else
            CanEmployeeTakePriorYearVacation = True
        End If
    Else
        CanEmployeeTakePriorYearVacation = False
    End If
    
    
End Function

Private Sub UpdateCompTimeBanks( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datDate As Date, _
                ByVal v_decHours As Variant, _
                ByVal v_strExceptionCode As String, _
                ByVal v_bolTakenBucket As Boolean, _
                ByVal v_bolAccruedBucket As Boolean, _
                ByRef r_rsBanks As ADODB.Recordset, _
                ByRef r_rsBRVs As ADODB.Recordset, _
                ByVal v_strBenefitsPackage As String)
                            
'There are 2 buckets for Comp Time
'       1. Accrued -- Gets Incremented when Comp Time is banked
'       2. Taken   -- Get Incremented when Comp Time is taken
'    Note: The Taken bucket contains totals
'           for Actual Time and Scheduled time
    
    Dim l_decEquivalentSTHours 'Contains Equivalent Straight time hours of the hours passed
    Dim l_intBankYear As Integer
    Dim l_intPriorBankYear As Integer
    Dim l_decTakenCompTimeHours As Variant
    Dim l_decAccruedCompTimeHours As Variant
    Dim l_decAvailCompTimeHours As Variant
    Dim l_decPriorYearTakenCompTimeHours As Variant
    Dim l_decPriorYearAccruedCompTimeHours As Variant
    Dim l_decPriorYearAvailCompTimeHours As Variant
    Dim l_decTotalAvailCompTimeHours As Variant
    Dim l_lngErrorNumber As Long
    Dim l_bolCurrentYearError  As Boolean
    Dim l_colDescArgs As Collection
    Dim l_strBankBucket As String
    Dim l_bolBankBalanceError As Boolean
    Dim l_decUpdateHours As Variant

    On Error GoTo UpdateCompTimeBanks_EH
    
    If v_decHours = 0 Then
        Exit Sub
    End If
    If v_bolTakenBucket Or v_bolAccruedBucket Then
    Else
        Exit Sub
    End If
    If v_bolTakenBucket Then
        l_strBankBucket = "T"
    Else
        l_strBankBucket = "A"
    End If
    l_intBankYear = Year(v_datDate)
    l_intPriorBankYear = l_intBankYear - 1
 'Get Current Year Bank Balances
    l_bolCurrentYearError = True
    l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intBankYear, "C", r_rsBanks, l_decAccruedCompTimeHours, l_decAvailCompTimeHours, l_decTakenCompTimeHours)
    
    If l_bolBankBalanceError Then
    'raise error bank year not found '2 app error
        Err.Raise trErrbusEmployeeDoesNotHaveBank
    End If
'Get Prior Year Bank Balances
    l_bolCurrentYearError = False
    l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intPriorBankYear, "C", r_rsBanks, l_decPriorYearAccruedCompTimeHours, l_decPriorYearAvailCompTimeHours, l_decPriorYearTakenCompTimeHours)
    l_decTotalAvailCompTimeHours = l_decAvailCompTimeHours + l_decPriorYearAvailCompTimeHours
    

'Update the Accrued Bucket if the CompTime was banked
'Banked CompTimes hours are always accrued for the year of the CompTime worked
    l_decEquivalentSTHours = GetEquilvalentStraightTimeHours(v_strExceptionCode, v_decHours)
    r_rsBanks.Filter = MDAC25FilterFix(" EMP_NUM = " & str(v_lngEmpNum) & " and EFF_YEAR = " & str(l_intBankYear))
    If v_bolAccruedBucket Then
        If v_strBenefitsPackage = "CIPS" Then
        Else
            l_decTotalAvailCompTimeHours = l_decTotalAvailCompTimeHours + l_decEquivalentSTHours
            If l_decTotalAvailCompTimeHours > 36 Then
                Err.Raise trErrbusEmployeeHasTooMuchCompTime 'BRV *
            End If
        End If
        l_decAvailCompTimeHours = l_decAvailCompTimeHours + l_decEquivalentSTHours
        If l_decAvailCompTimeHours < 0 Then
            Err.Raise trErrbusOutofBalanceCompTimeHours 'app error 1..
        End If
        l_decAccruedCompTimeHours = l_decAccruedCompTimeHours + l_decEquivalentSTHours
        r_rsBanks!COMP_STR_TIME_HRS = l_decAccruedCompTimeHours
        AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, l_decEquivalentSTHours, "C", l_strBankBucket, l_intBankYear, " "
        r_rsBanks.Update
        r_rsBanks.Filter = ""
        Exit Sub
    End If
    
    
'Comp Time taken are deducted from the prior year first, up to the accrued balance
'Any hours exceeding the accrued balance in the prior year are
' deducted from the current year. When deleting or reducing previous entered
' Comp Time taken, the system will decrement the taken bucket for the current
' year first.  This allows the employee to use any remaining hours accrued
' in the prior year before using any accrued hours in the current year.
    

'Deduct Taken hours from prior Year First
    If l_decEquivalentSTHours > 0 Then
        If l_decPriorYearAvailCompTimeHours > 0 Then
            l_decPriorYearAvailCompTimeHours = l_decPriorYearAvailCompTimeHours - l_decEquivalentSTHours
            If l_decPriorYearAvailCompTimeHours < 0 Then
                l_decUpdateHours = l_decPriorYearAvailCompTimeHours + l_decEquivalentSTHours
                UpdateCompTimeBanksSetBank v_lngEmpNum, l_intPriorBankYear, l_decUpdateHours, r_rsBanks, v_datDate, v_strExceptionCode
                l_decAvailCompTimeHours = l_decAvailCompTimeHours + l_decPriorYearAvailCompTimeHours
                If l_decAvailCompTimeHours < 0 Then
                 'BRV not enough hours are available *
                    Err.Raise trErrbusNotEnoughCompTimeHoursAvailable
                End If
                l_decUpdateHours = l_decPriorYearAvailCompTimeHours * -1
                UpdateCompTimeBanksSetBank v_lngEmpNum, l_intBankYear, l_decUpdateHours, r_rsBanks, v_datDate, v_strExceptionCode
            Else
                UpdateCompTimeBanksSetBank v_lngEmpNum, l_intPriorBankYear, l_decEquivalentSTHours, r_rsBanks, v_datDate, v_strExceptionCode
            End If
        Else  'Take for current Year
            l_decAvailCompTimeHours = l_decAvailCompTimeHours - l_decEquivalentSTHours
            If l_decAvailCompTimeHours < 0 Then
                 'BRV not enough hours are available *
                 Err.Raise trErrbusNotEnoughCompTimeHoursAvailable
            End If
            UpdateCompTimeBanksSetBank v_lngEmpNum, l_intBankYear, l_decEquivalentSTHours, r_rsBanks, v_datDate, v_strExceptionCode
        End If
    End If
    
'Deduct Hours from current year first
    If l_decEquivalentSTHours < 0 Then
        If l_decTakenCompTimeHours > 0 Then 'deducting from current year
            l_decTakenCompTimeHours = l_decTakenCompTimeHours + l_decEquivalentSTHours
            If l_decTakenCompTimeHours < 0 Then 'deduct some from prior year
                l_decPriorYearTakenCompTimeHours = l_decPriorYearTakenCompTimeHours + l_decTakenCompTimeHours
                If l_decPriorYearTakenCompTimeHours < 0 Then
                ' Out of Balance
                    Err.Raise trErrbusOutofBalanceCompTimeHours 'app error 2..
                End If
                UpdateCompTimeBanksSetBank v_lngEmpNum, l_intPriorBankYear, l_decTakenCompTimeHours, r_rsBanks, v_datDate, v_strExceptionCode
                l_decUpdateHours = l_decTakenCompTimeHours * -1 + l_decEquivalentSTHours
                UpdateCompTimeBanksSetBank v_lngEmpNum, l_intBankYear, l_decUpdateHours, r_rsBanks, v_datDate, v_strExceptionCode
            Else 'Deducting from current year
                UpdateCompTimeBanksSetBank v_lngEmpNum, l_intBankYear, l_decEquivalentSTHours, r_rsBanks, v_datDate, v_strExceptionCode
            End If
        Else    'Deducting from prior year
            l_decPriorYearTakenCompTimeHours = l_decPriorYearTakenCompTimeHours + l_decEquivalentSTHours
            If l_decPriorYearTakenCompTimeHours < 0 Then
                ' Out of Balance
                Err.Raise trErrbusOutofBalanceCompTimeHours 'app error 3..
            End If
            UpdateCompTimeBanksSetBank v_lngEmpNum, l_intPriorBankYear, l_decEquivalentSTHours, r_rsBanks, v_datDate, v_strExceptionCode
        End If
    End If
    
    r_rsBanks.Filter = ""
    Exit Sub
    
UpdateCompTimeBanks_EH:
    
    Set m_rsAuditTrail = Nothing
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(v_lngEmpNum)
    If l_bolCurrentYearError Then
        l_colDescArgs.Add str(l_intBankYear)
    Else
        l_colDescArgs.Add str(l_intPriorBankYear)
    End If
    l_lngErrorNumber = Err.Number
    Select Case l_lngErrorNumber
        Case trErrbusEmployeeHasTooMuchCompTime, _
             trErrbusNotEnoughCompTimeHoursAvailable
                Err.Source = FormatBRV(TrType_TC, v_lngEmpNum, v_datDate, v_datDate, "", "", v_decHours, v_strExceptionCode, "")
                TrsRaiseError m_strCMODULENAME, "UpdateCompTimeBanks", l_lngErrorNumber, , l_colDescArgs
        Case trErrbusOutofBalanceCompTimeHours, _
             trErrbusEmployeeDoesNotHaveBank
                TrsRaiseError m_strCMODULENAME, "UpdateCompTimeBanks", l_lngErrorNumber, , l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateCompTimeBanks"
    End Select

End Sub
Private Sub UpdateCompTimeBanksSetBank( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_intBankYear As Integer, _
                ByVal v_decHours As Variant, _
                ByRef r_rsBanks As ADODB.Recordset, _
                ByVal v_datDate As Date, _
                ByVal v_strExceptionCode As String)
    
    Dim l_decTakenHours As Variant
    
    On Error GoTo UpdateCompTimeBanksSetBank_EH
                
    r_rsBanks.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmpNum) & " and  EFF_YEAR = " & str(v_intBankYear))
    l_decTakenHours = r_rsBanks!COMP_TAKEN_SCHED
    l_decTakenHours = l_decTakenHours + v_decHours
    r_rsBanks!COMP_TAKEN_SCHED = l_decTakenHours
    AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "C", "T", v_intBankYear, " "
    r_rsBanks.Update
    r_rsBanks.Filter = ""
    
    Exit Sub
    
UpdateCompTimeBanksSetBank_EH:

    TrsRaiseError m_strCMODULENAME, "UpdateCompTimeBanksSetBank"
                
End Sub

Private Function GetEquilvalentStraightTimeHours( _
                    ByVal v_strExceptionCode As String, _
                    ByVal v_decHours As Variant _
                    ) As Variant
                    
    Dim l_decFactor As Variant
    
    Select Case Trim(v_strExceptionCode)
        Case "CB1", "CT"
            l_decFactor = CDec(1#)
        Case "CB", "CTP"
            l_decFactor = CDec(1.5)
        Case "CB2"
            l_decFactor = CDec(2#)
    
    End Select
    
    GetEquilvalentStraightTimeHours = l_decFactor * CDec(v_decHours)
    
    

End Function

Private Sub UpdateHolidayBanks(ByVal v_lngEmpNum As Long, _
                            ByVal v_datDate As Date, _
                            ByVal v_decHours As Variant, _
                            ByVal v_strExceptionCode As String, _
                            ByVal v_bolTakenBucket As Boolean, _
                            ByVal v_bolAccruedBucket As Boolean, _
                            ByRef r_rsBanks As ADODB.Recordset, _
                            ByRef r_rsBRVs As ADODB.Recordset, _
                            ByRef l_rsEmpHolidayBankAudit As ADODB.Recordset)
                            
'There are 2 buckets for Holiday
'       1. Accrued -- Gets Incremented when Holiday's are banked
'       2. Taken   -- Get Incremented when Holiday is taken
'    Note: The Taken bucket contains totals
'           for Actual Time and Scheduled time
    
    Dim l_intBankYear As Integer
    Dim l_intPriorBankYear As Integer
    Dim l_decTakenHolidayHours As Variant
    Dim l_decAccruedHolidayHours As Variant
    Dim l_decAvailHolidayHours As Variant
    Dim l_decPriorYearTakenHolidayHours As Variant
    Dim l_decPriorYearAccruedHolidayHours As Variant
    Dim l_decPriorYearAvailHolidayHours As Variant
    Dim l_decTotalAvailHolidayHours As Variant
    Dim l_lngErrorNumber As Long
    Dim l_bolCurrentYearError  As Boolean
    Dim l_colDescArgs As Collection
    Dim l_strBankBucket As String
    Dim l_bolBankBalanceError As Boolean
    Dim l_decUpdateHours As Variant
    Dim l_decFutureBankedHours As Variant

    On Error GoTo UpdateHolidayBanks_EH
    
    If v_decHours = 0 Then
        Exit Sub
    End If
    If v_bolTakenBucket Or v_bolAccruedBucket Then
    Else
        Exit Sub
    End If
    If v_bolTakenBucket Then
        l_strBankBucket = "T"
    Else
        l_strBankBucket = "A"
    End If
    l_intBankYear = Year(v_datDate)
    l_intPriorBankYear = l_intBankYear - 1
    
    l_bolCurrentYearError = True
    l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intBankYear, "H", r_rsBanks, l_decAccruedHolidayHours, l_decAvailHolidayHours, l_decTakenHolidayHours)
    If l_bolBankBalanceError Then
    'raise error bank year not found '3 app error
        Err.Raise trErrbusEmployeeDoesNotHaveBank
    End If


'Update the Accrued Bucket if the Holiday was banked
'Banked Holidays hours are always accrued for the year of the holiday worked
    r_rsBanks.Filter = MDAC25FilterFix(" EMP_NUM = " & str(v_lngEmpNum) & " and EFF_YEAR = " & str(l_intBankYear))
    If v_bolAccruedBucket Then
        l_decAvailHolidayHours = l_decAvailHolidayHours + v_decHours
        If l_decAvailHolidayHours < 0 Then
            ' Out Of Balance
            Err.Raise trErrbusOutofBalanceHolidayHours 'app error 1.
        End If
        l_decAccruedHolidayHours = l_decAccruedHolidayHours + v_decHours
        r_rsBanks!HOL_HRS_ACCRUED = l_decAccruedHolidayHours
        AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "H", l_strBankBucket, l_intBankYear, " "
        MaintainEmpHolidayBankAudit v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "H", l_strBankBucket, l_intBankYear, " ", l_rsEmpHolidayBankAudit
        r_rsBanks.Update
        r_rsBanks.Filter = ""
        Exit Sub
    End If
    
    
'Banked Holidays taken are deducted from the prior year first, up to the accrued balance
'Any hours exceeding the accrued balance in the prior year are
' deducted from the current year. When deleting or reducing previous entered
' holiday taken, the system will decrement the taken bucket for the current
' year first.  This allows the employee to use any remaining hours accrued
' in the prior year before using any accrued hours in the current year.
    
    l_bolCurrentYearError = False
    l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intPriorBankYear, "H", r_rsBanks, l_decPriorYearAccruedHolidayHours, l_decPriorYearAvailHolidayHours, l_decPriorYearTakenHolidayHours)
        
 'Deduct Hours Banked on Future Dates from the current total
    l_decFutureBankedHours = 0
    l_rsEmpHolidayBankAudit.Filter = MDAC25FilterFix(" EMP_NUM = " & str(v_lngEmpNum) & " and CHARGE_DATE > '" & v_datDate & "' and HOURS > 0 and EFF_YEAR = " & str(l_intBankYear))
    If Not l_rsEmpHolidayBankAudit.EOF Then
        l_rsEmpHolidayBankAudit.MoveFirst
        While Not l_rsEmpHolidayBankAudit.EOF
            l_decFutureBankedHours = l_decFutureBankedHours + l_rsEmpHolidayBankAudit!HOURS
            l_rsEmpHolidayBankAudit.MoveNext
        Wend
    End If
    l_decAvailHolidayHours = l_decAvailHolidayHours - l_decFutureBankedHours
    
    l_decTotalAvailHolidayHours = l_decAvailHolidayHours - l_decPriorYearAvailHolidayHours
    
    
'Deduct Taken hours from prior Year First
    If v_decHours > 0 Then
        If l_decPriorYearAvailHolidayHours > 0 Then
            l_decPriorYearAvailHolidayHours = l_decPriorYearAvailHolidayHours - v_decHours
            If l_decPriorYearAvailHolidayHours < 0 Then
                l_decUpdateHours = l_decPriorYearAvailHolidayHours + v_decHours
                UpdateHolidayBanksSetBank v_lngEmpNum, l_intPriorBankYear, l_decUpdateHours, r_rsBanks, v_datDate, v_strExceptionCode
                MaintainEmpHolidayBankAudit v_lngEmpNum, v_datDate, v_strExceptionCode, l_decUpdateHours, "H", l_strBankBucket, l_intBankYear, " ", l_rsEmpHolidayBankAudit
                l_decAvailHolidayHours = l_decAvailHolidayHours + l_decPriorYearAvailHolidayHours
                If l_decAvailHolidayHours < 0 Then
                 'BRV not enough hours are available *
                    Err.Raise trErrbusNotEnoughHolidayHoursAvailable
                End If
                l_decUpdateHours = l_decPriorYearAvailHolidayHours * -1
                UpdateHolidayBanksSetBank v_lngEmpNum, l_intBankYear, l_decUpdateHours, r_rsBanks, v_datDate, v_strExceptionCode
                MaintainEmpHolidayBankAudit v_lngEmpNum, v_datDate, v_strExceptionCode, l_decUpdateHours, "H", l_strBankBucket, l_intBankYear, " ", l_rsEmpHolidayBankAudit
            Else
                UpdateHolidayBanksSetBank v_lngEmpNum, l_intPriorBankYear, v_decHours, r_rsBanks, v_datDate, v_strExceptionCode
                MaintainEmpHolidayBankAudit v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "H", l_strBankBucket, l_intBankYear, " ", l_rsEmpHolidayBankAudit
            End If
        Else  'Take for current Year
            l_decAvailHolidayHours = l_decAvailHolidayHours - v_decHours
            If l_decAvailHolidayHours < 0 Then
                 'BRV not enough hours are available *
                 Err.Raise trErrbusNotEnoughHolidayHoursAvailable
            End If
            UpdateHolidayBanksSetBank v_lngEmpNum, l_intBankYear, v_decHours, r_rsBanks, v_datDate, v_strExceptionCode
            MaintainEmpHolidayBankAudit v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "H", l_strBankBucket, l_intBankYear, " ", l_rsEmpHolidayBankAudit
        End If
    End If
    
'Deduct Hours from current year first
    If v_decHours < 0 Then
        If l_decTakenHolidayHours > 0 Then 'deducting from current year
            l_decTakenHolidayHours = l_decTakenHolidayHours + v_decHours
            If l_decTakenHolidayHours < 0 Then 'deduct some from prior year
                l_decPriorYearTakenHolidayHours = l_decPriorYearTakenHolidayHours + l_decTakenHolidayHours
                If l_decPriorYearTakenHolidayHours < 0 Then
                ' Out of Balance
                    Err.Raise trErrbusOutofBalanceHolidayHours 'app error 2.
                End If
                UpdateHolidayBanksSetBank v_lngEmpNum, l_intPriorBankYear, l_decTakenHolidayHours, r_rsBanks, v_datDate, v_strExceptionCode
                l_decUpdateHours = l_decTakenHolidayHours * -1 + v_decHours
                UpdateHolidayBanksSetBank v_lngEmpNum, l_intBankYear, l_decUpdateHours, r_rsBanks, v_datDate, v_strExceptionCode
            Else 'Deducting from current year
                UpdateHolidayBanksSetBank v_lngEmpNum, l_intBankYear, v_decHours, r_rsBanks, v_datDate, v_strExceptionCode
            End If
        Else    'Deducting from prior year
            l_decPriorYearTakenHolidayHours = l_decPriorYearTakenHolidayHours + v_decHours
            If l_decPriorYearTakenHolidayHours < 0 Then
                ' Out of Balance
                Err.Raise trErrbusOutofBalanceHolidayHours 'app error 3.
            End If
            UpdateHolidayBanksSetBank v_lngEmpNum, l_intPriorBankYear, v_decHours, r_rsBanks, v_datDate, v_strExceptionCode
        End If
    End If
    
    r_rsBanks.Filter = ""
    l_rsEmpHolidayBankAudit.Filter = ""
    Exit Sub
    
UpdateHolidayBanks_EH:
    
    Set m_rsAuditTrail = Nothing
    
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(v_lngEmpNum)
    If l_bolCurrentYearError Then
        l_colDescArgs.Add str(l_intBankYear)
    Else
        l_colDescArgs.Add str(l_intPriorBankYear)
    End If
    l_lngErrorNumber = Err.Number
    Select Case l_lngErrorNumber
        Case trErrbusNotEnoughHolidayHoursAvailable
                Err.Source = FormatBRV(TrType_TC, v_lngEmpNum, v_datDate, v_datDate, "", "", v_decHours, v_strExceptionCode, "")
                TrsRaiseError m_strCMODULENAME, "UpdateHolidayBanks", l_lngErrorNumber, , l_colDescArgs
        Case trErrbusEmployeeDoesNotHaveBank, _
             trErrbusOutofBalanceHolidayHours
             TrsRaiseError m_strCMODULENAME, "UpdateCompTimeBanks", l_lngErrorNumber, , l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateHolidayBanks"
    End Select


End Sub
Private Sub UpdateHolidayBanksSetBank( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_intBankYear As Integer, _
                ByVal v_decHours As Variant, _
                ByRef r_rsBanks As ADODB.Recordset, _
                ByVal v_datDate As Date, _
                ByVal v_strExceptionCode As String)
    
    Dim l_decTakenHours As Variant
    
    On Error GoTo UpdateHolidayBanksSetBank_EH
                
    r_rsBanks.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmpNum) & " and  EFF_YEAR = " & str(v_intBankYear))
    l_decTakenHours = r_rsBanks!HOL_TAKEN_SCHED
    l_decTakenHours = l_decTakenHours + v_decHours
    r_rsBanks!HOL_TAKEN_SCHED = l_decTakenHours
    AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "H", "T", v_intBankYear, " "
    r_rsBanks.Update
    r_rsBanks.Filter = ""
    
    Exit Sub
    
UpdateHolidayBanksSetBank_EH:

    TrsRaiseError m_strCMODULENAME, "UpdateHolidayBanksSetBank"
                
End Sub


Private Sub UpdateFMLABanks(ByVal v_lngEmpNum As Long, _
                            ByVal v_datDate As Date, _
                            ByVal v_decHours As Variant, _
                            ByVal v_strExceptionCode As String, _
                            ByVal v_bolUsedBucket As Boolean, _
                            ByVal v_bolSchedBucket As Boolean, _
                            ByRef r_rsBanks As ADODB.Recordset, _
                            ByRef r_rsBRVs As ADODB.Recordset)
    
    Dim l_intBankYear As Integer
    Dim l_decAllottedFMLAhours As Variant
    Dim l_decUsedFMLAhours As Variant
    Dim l_decSchedFMLAhours As Variant
    Dim l_decAvailFMLAhours As Variant
    Dim l_colDescArgs As Collection
    Dim l_lngErrorNumber As Long
    Dim l_bolBankBalanceError As Boolean
    
    On Error GoTo UpdateFMLABanks_EH
    
    If v_bolSchedBucket Or v_bolUsedBucket Then
    Else
        Exit Sub
    End If
    l_intBankYear = Year(v_datDate)
    
    l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intBankYear, "F", r_rsBanks, l_decAllottedFMLAhours, l_decAvailFMLAhours, l_decUsedFMLAhours, l_decSchedFMLAhours)
    If l_bolBankBalanceError Then
    'raise error bank year not found '4 app error
        Err.Raise trErrbusEmployeeDoesNotHaveBank
    End If

    If v_bolSchedBucket And v_bolUsedBucket Then
    Else
        l_decAvailFMLAhours = l_decAvailFMLAhours - v_decHours
    End If
    
    If l_decAvailFMLAhours < 0 Then
    'BRV unavailable FMLA hours *
        Err.Raise trErrbusNotEnoughFMLA_HoursAvailable
    End If
    
    r_rsBanks.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmpNum) & " and EFF_YEAR = " & str(l_intBankYear))
    If v_bolUsedBucket Then
        l_decUsedFMLAhours = l_decUsedFMLAhours + v_decHours
        If l_decUsedFMLAhours < 0 Then
            Err.Raise trErrbusOutofBalanceFMLA_Hours 'app error
        End If
        r_rsBanks!FMLA_HRS_USED = l_decUsedFMLAhours
        AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "F", "U", l_intBankYear, " "
        If v_bolSchedBucket Then
            l_decSchedFMLAhours = l_decSchedFMLAhours - v_decHours
            If l_decSchedFMLAhours < 0 Then
                Err.Raise trErrbusOutofBalanceFMLA_Hours 'app error
            End If
            r_rsBanks!FMLA_HRS_SCHED = l_decSchedFMLAhours
            AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours * -1, "F", "S", l_intBankYear, " "
        End If
    Else
        l_decSchedFMLAhours = l_decSchedFMLAhours + v_decHours
        If l_decSchedFMLAhours < 0 Then
            Err.Raise trErrbusOutofBalanceFMLA_Hours 'app error
        End If
        r_rsBanks!FMLA_HRS_SCHED = l_decSchedFMLAhours
        AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "F", "S", l_intBankYear, " "
    End If
    r_rsBanks.Update
    r_rsBanks.Filter = ""
    Exit Sub
    
UpdateFMLABanks_EH:

    Set m_rsAuditTrail = Nothing
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(v_lngEmpNum)
    l_colDescArgs.Add str(l_intBankYear)
    l_lngErrorNumber = Err.Number
    Select Case l_lngErrorNumber
        Case trErrbusNotEnoughFMLA_HoursAvailable
                Err.Source = FormatBRV(TrType_TC, v_lngEmpNum, v_datDate, v_datDate, "", "", v_decHours, v_strExceptionCode, "")
                TrsRaiseError m_strCMODULENAME, "UpdateFMLABanks", l_lngErrorNumber, , l_colDescArgs
        Case trErrbusOutofBalanceFMLA_Hours, _
             trErrbusEmployeeDoesNotHaveBank
                TrsRaiseError m_strCMODULENAME, "UpdateFMLABanks", l_lngErrorNumber, , l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateFMLABanks"
    End Select
    
End Sub


Public Function UpdateBanksForExceptionSched( _
                ByVal v_request As ADODB.Recordset, _
                ByRef r_rsBRVs As ADODB.Recordset, _
                ByVal v_lngLastModEmpNum As Long) As Boolean
' Set r_rsBRVs to nothing for now, but if warnings are desired; create
' a recordset for output only.
    Dim l_rsForamattedRequest As ADODB.Recordset
    Dim objCtx As ObjectContext
         
On Error GoTo eh
    
    Set r_rsBRVs = Nothing
    m_lngLastModEmpNum = v_lngLastModEmpNum
    
    Set objCtx = GetObjectContext
    
    If v_request.EOF And v_request.BOF Then
        objCtx.SetComplete
        Set objCtx = Nothing
        UpdateBanksForExceptionSched = True
        Exit Function
    End If
    
    Set l_rsForamattedRequest = BuildRequestForExceptionSched(v_request)
    If l_rsForamattedRequest.EOF And l_rsForamattedRequest.BOF Then
        objCtx.SetComplete
        Set objCtx = Nothing
        UpdateBanksForExceptionSched = True
        Set l_rsForamattedRequest = Nothing
        Exit Function
    End If
    
    ProcessBankUpdateRequests l_rsForamattedRequest, r_rsBRVs
    objCtx.SetComplete
    Set objCtx = Nothing
    UpdateBanksForExceptionSched = True
    Set l_rsForamattedRequest = Nothing
    Exit Function
    
eh:
    objCtx.SetAbort
    Set objCtx = Nothing
    Set r_rsBRVs = Nothing
    Set l_rsForamattedRequest = Nothing
    TrsRaiseError m_strCMODULENAME, "UpdateBanksForExceptionSched", , , , , True
End Function


Private Function BuildRequestForExceptionSched( _
                    ByVal v_rsExcRequest As ADODB.Recordset _
                    ) As ADODB.Recordset
    Dim l_rsRequest As ADODB.Recordset
    Dim l_strRecStatus As String
    Dim l_strBankCode As String
    Dim l_varEmpNums As Variant
    Dim l_cdEmployee As dbEmployee.cdEmployee
    Dim l_rsEmpBenPkg As ADODB.Recordset
    Dim l_cbsEmployee As busEmployee.cbsEmployee
    Dim l_strEmpBenPkg As String
    Dim l_strUnionCd As String
    
On Error GoTo eh

    Set l_rsRequest = CreateRequestRecordset()
    
    If Not AnyAccrualBankRelated(v_rsExcRequest) Then
        Set BuildRequestForExceptionSched = l_rsRequest
        Exit Function
    End If
    
'    get employee records
    l_varEmpNums = DetermineGetEmployeeScope(v_rsExcRequest)
    Set l_cdEmployee = CreateObject("dbEmployee.cdEmployee")
    Set l_rsEmpBenPkg = l_cdEmployee.GetEmployeeBenefitsPackage(l_varEmpNums)
    Set l_cdEmployee = Nothing

    Set l_cbsEmployee = CreateObject("busEmployee.cbsEmployee")
    Do Until v_rsExcRequest.EOF
       
        l_rsEmpBenPkg.Filter = "EMP_NUM = " & v_rsExcRequest!EMP_NUM
        If l_cbsEmployee.IsExecutive( _
            l_rsEmpBenPkg!JOB_EFF_DATE, _
            l_rsEmpBenPkg!UNION_CD, _
            l_rsEmpBenPkg!PRIOR_UNION_CD, _
            l_rsEmpBenPkg!EXEC_CONTRACT_IND, _
            l_rsEmpBenPkg!PRIOR_EXEC_CON_IND, _
            v_rsExcRequest!CHARGE_DATE) Then
            l_strEmpBenPkg = "UE"
        Else
            If l_cbsEmployee.IsAmerenCILCOLocal8Employee( _
                l_rsEmpBenPkg!JOB_EFF_DATE, _
                l_rsEmpBenPkg!UNION_CD, _
                l_rsEmpBenPkg!PRIOR_UNION_CD, _
                v_rsExcRequest!CHARGE_DATE) Then
                l_strEmpBenPkg = "UE"
            Else
                l_strEmpBenPkg = l_cbsEmployee.EmployeeBenefitPackage( _
                    l_rsEmpBenPkg!JOB_EFF_DATE, l_rsEmpBenPkg!MGMT_BEN_PKG_IND, _
                    l_rsEmpBenPkg!UNION_CD, l_rsEmpBenPkg!PRIOR_UNION_CD, _
                    l_rsEmpBenPkg!EXEC_CONTRACT_IND, l_rsEmpBenPkg!PRIOR_EXEC_CON_IND, _
                    l_rsEmpBenPkg!UNION_BEN_PKG_IND, l_rsEmpBenPkg!PRIOR_UNION_BN_PKG, _
                    v_rsExcRequest!CHARGE_DATE)
            End If
        End If
        l_strRecStatus = Trim(v_rsExcRequest!REC_STATUS)
        l_strUnionCd = l_cbsEmployee.EmployeeUnion( _
                    l_rsEmpBenPkg!JOB_EFF_DATE, _
                    l_rsEmpBenPkg!UNION_CD, _
                    l_rsEmpBenPkg!PRIOR_UNION_CD, _
                    v_rsExcRequest!CHARGE_DATE)
        RequestForExceptionSched l_strRecStatus, l_rsRequest, v_rsExcRequest, l_strEmpBenPkg, l_strUnionCd
        v_rsExcRequest.MoveNext
    Loop
    Set l_cbsEmployee = Nothing
    
    Set BuildRequestForExceptionSched = l_rsRequest
    Set l_rsRequest = Nothing
    
    Exit Function
    
eh:
    Set l_rsRequest = Nothing
    Set l_cdEmployee = Nothing
    Set l_cbsEmployee = Nothing
    TrsRaiseError m_strCMODULENAME, "BuildRequestForExceptionSched"
End Function

Private Sub RequestForExceptionSched( _
    ByVal v_strReqType As String, _
    ByRef r_rsRequest As ADODB.Recordset, _
    ByRef r_rsExcRequest As ADODB.Recordset, _
    ByVal v_strEmpBenPkg As String, _
    ByVal v_strUnionCd As String)

    Dim l_strBankCode As String
    Dim l_strBankBucket As String
    Dim l_strSchedInd As String
    Dim l_strOverExcCd As String

On Error GoTo eh

'    Build FMLA request
    If Trim(r_rsExcRequest!FMLA_IND) = "Y" Then
        With r_rsRequest
            .AddNew
            !REC_STATUS = v_strReqType
            !EMP_NUM = r_rsExcRequest!EMP_NUM
            !CHARGE_DATE = r_rsExcRequest!CHARGE_DATE
            !BENEFITS_PACKAGE = v_strEmpBenPkg
            !EXCEPTION_CD = r_rsExcRequest!EXCEPTION_CD
            !OVER_EXC_CD = r_rsExcRequest!OVER_EXC_CD
            !UNION_CD = v_strUnionCd
            If v_strReqType = "D" Then
                !HOURS = CDec(r_rsExcRequest!HOURS) * -1
            Else
                !HOURS = CDec(r_rsExcRequest!HOURS)
            End If
            !BANK_CODE = "F"
            If r_rsExcRequest!ACTUAL_TIME_IND <> "Y" Then
                !BANK_BUCKET = "S"
            Else
                !BANK_BUCKET = "U"
            End If
            .Update
        End With
    End If
   
    l_strBankCode = Trim(r_rsExcRequest!BANK_CODE)
    l_strOverExcCd = Trim(r_rsExcRequest!OVER_EXC_CD)
    If l_strOverExcCd = "FCL" Then
        l_strBankCode = "S"
    End If
    
    If l_strBankCode = "" Or l_strBankCode = "N" Then
        Exit Sub
    End If
    
'    Build Accrual bank request
    With r_rsRequest
        .AddNew
        !REC_STATUS = v_strReqType
        !EMP_NUM = r_rsExcRequest!EMP_NUM
        !CHARGE_DATE = r_rsExcRequest!CHARGE_DATE
        !BENEFITS_PACKAGE = v_strEmpBenPkg
        !UNION_CD = v_strUnionCd
        !EXCEPTION_CD = r_rsExcRequest!EXCEPTION_CD
        !OVER_EXC_CD = r_rsExcRequest!OVER_EXC_CD
        If v_strReqType = "D" Then
            !HOURS = CDec(r_rsExcRequest!HOURS) * -1
        Else
            !HOURS = CDec(r_rsExcRequest!HOURS)
        End If
        !BANK_CODE = Left(l_strBankCode, 1)
        If r_rsExcRequest!ACTUAL_TIME_IND <> "Y" Then
            l_strSchedInd = "Y"
        Else
            l_strSchedInd = "N"
        End If
        !BANK_BUCKET = DetermineBankBucket(l_strBankCode, l_strSchedInd)
        .Update
    End With
    
    Exit Sub
           
eh:
    TrsRaiseError m_strCMODULENAME, "RequestForExceptionSched"
End Sub

Private Sub UpdateVacationBanksDeleteProcess( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date, _
    ByVal v_decHours As Variant, _
    ByVal v_strExceptionCode As String, _
    ByVal v_bolUsedBucket As Boolean, _
    ByVal v_bolSchedBucket As Boolean, _
    ByRef r_rsBanks As ADODB.Recordset, _
    ByRef r_rsBRVs As ADODB.Recordset, _
    ByVal v_strBenefitsPackage As String, _
    ByVal l_bolCanEmployeeTakePriorYearVacation As Boolean, _
    ByVal l_decAvailVacationhours As Variant, _
    ByVal l_decAllottedVacationhours As Variant, _
    ByVal v_strUnionCd As String, _
    Optional v_blnChargePriorYearOnly As Boolean = False)

    Dim l_intBankYear As Integer
    Dim l_intPriorBankYear As Integer
    Dim l_decUpdateHours As Variant
    Dim l_strErrorDesc As String
    Dim l_bolBankBalanceError As Boolean
    Dim l_bolCurrentYearError As Boolean
    
    On Error GoTo eh
    
    
    l_intBankYear = DetermineBankYear(v_strUnionCd, v_datDate)
    l_intPriorBankYear = l_intBankYear - 1
    
    If v_blnChargePriorYearOnly Then
        l_decAvailVacationhours = l_decAvailVacationhours - v_decHours
        If l_decAvailVacationhours > (l_decAllottedVacationhours) Then
            'out of balanace condition 'app error 1
            l_bolCurrentYearError = False
            l_strErrorDesc = "An error occured while giving hours back.  " _
                    & "Prior year only logic. "
            Err.Raise trErrbusOutofBalanceVacationHours
        Else
            UpdateVacationBanksSetBank v_lngEmpNum, l_intPriorBankYear, _
                v_decHours, r_rsBanks, v_bolUsedBucket, _
                v_bolSchedBucket, v_datDate, v_strExceptionCode
        End If
        Exit Sub
    End If
    
    l_decAvailVacationhours = l_decAvailVacationhours - v_decHours
    If l_decAvailVacationhours > (l_decAllottedVacationhours) Then
        'deduct the remainder hours from prior year if allowed
        If l_bolCanEmployeeTakePriorYearVacation Then
        ' Deduct as much as possible from current year
            l_decUpdateHours = l_decAvailVacationhours - l_decAllottedVacationhours + v_decHours
            If l_decUpdateHours < 0 Then
                UpdateVacationBanksSetBank v_lngEmpNum, l_intBankYear, v_decHours, r_rsBanks, v_bolUsedBucket, v_bolSchedBucket, v_datDate, v_strExceptionCode
            End If
        ' Deduct Remainder from prior year
            l_decUpdateHours = l_decAllottedVacationhours - l_decAvailVacationhours
            UpdateVacationBanksSetBank v_lngEmpNum, l_intPriorBankYear, l_decUpdateHours, r_rsBanks, v_bolUsedBucket, v_bolSchedBucket, v_datDate, v_strExceptionCode
        Else
            'out of balanace condition 'app error 1
            l_bolCurrentYearError = True
            l_strErrorDesc = "An error occured while giving hours back.  " _
                    & "Current year only logic. "
            Err.Raise trErrbusOutofBalanceVacationHours
        End If
    Else
        'Deduct Hours from Current Year
        UpdateVacationBanksSetBank v_lngEmpNum, l_intBankYear, v_decHours, r_rsBanks, v_bolUsedBucket, v_bolSchedBucket, v_datDate, v_strExceptionCode
    End If
    
    Exit Sub
           
eh:
   Dim l_colErrArgs As Collection
   Dim l_colDescArgs As Collection
   Dim l_lngErrorNumber As Long
   
   Set l_colErrArgs = New Collection
    l_colErrArgs.Add "l_strErrorDesc: " & l_strErrorDesc
    l_colErrArgs.Add "v_lngEmpNum: " & CStr(v_lngEmpNum)
    l_colErrArgs.Add "v_datDate: " & Format(v_datDate, "mm/dd/yyyy")
    l_colErrArgs.Add "v_decHours: " & CStr(v_decHours)
    l_colErrArgs.Add "v_strExceptionCode: " & v_strExceptionCode
    l_colErrArgs.Add "v_bolUsedBucket: " & CStr(v_bolUsedBucket)
    l_colErrArgs.Add "v_bolSchedBucket: " & CStr(v_bolSchedBucket)
    l_colErrArgs.Add "v_strBenefitsPackage: " & v_strBenefitsPackage
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add CStr(v_lngEmpNum)
    If l_bolCurrentYearError Then
        l_colDescArgs.Add CStr(l_intBankYear)
    Else
        l_colDescArgs.Add CStr(l_intPriorBankYear)
    End If
    
    l_lngErrorNumber = Err.Number
    Select Case l_lngErrorNumber
        Case trErrbusOutofBalanceVacationHours
                TrsRaiseError m_strCMODULENAME, "UpdateVacationBanksDeleteProcess", l_lngErrorNumber, , l_colDescArgs, l_colErrArgs, True
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateVacationBanksDeleteProcess"
    End Select
                
End Sub

Private Sub InsertRequestForCascade( _
                  ByRef r_rsRequest As ADODB.Recordset, _
                  ByRef r_rsCascadeRequest As ADODB.Recordset, _
                  ByVal v_strEmpBenPkg As String, _
                  ByVal v_strUnionCd As String)

    Dim l_strBankCode As String
    Dim l_strBankBucket As String
    Dim l_strOverExcCd As String

On Error GoTo eh

'    Build FMLA request
    If Trim(r_rsCascadeRequest!FMLA_IND) = "Y" Then
        With r_rsRequest
            .AddNew
            !REC_STATUS = "I"
            !EMP_NUM = r_rsCascadeRequest!EMP_NUM
            !CHARGE_DATE = r_rsCascadeRequest!CHARGE_DATE
            !BENEFITS_PACKAGE = v_strEmpBenPkg
            !UNION_CD = v_strUnionCd
            !EXCEPTION_CD = r_rsCascadeRequest!EXCEPTION_CD
            !OVER_EXC_CD = r_rsCascadeRequest!OVER_EXC_CD
            !HOURS = CDec(r_rsCascadeRequest!HOURS)
            !BANK_CODE = "F"
            !BANK_BUCKET = "S"
            .Update
        End With
    End If
   
    l_strBankCode = Trim(r_rsCascadeRequest!BANK_CODE)
    l_strOverExcCd = Trim(r_rsCascadeRequest!OVER_EXC_CD)
    If l_strOverExcCd = "FCL" Then
        l_strBankCode = "S"
    End If
    If l_strBankCode = "" Or l_strBankCode = "N" Then
        Exit Sub
    End If
    
'    Build Accrual bank request
    With r_rsRequest
        .AddNew
        !REC_STATUS = "I"
        !EMP_NUM = r_rsCascadeRequest!EMP_NUM
        !CHARGE_DATE = r_rsCascadeRequest!CHARGE_DATE
        !BENEFITS_PACKAGE = v_strEmpBenPkg
        !EXCEPTION_CD = r_rsCascadeRequest!EXCEPTION_CD
        !OVER_EXC_CD = r_rsCascadeRequest!OVER_EXC_CD
        !UNION_CD = v_strUnionCd
        !HOURS = CDec(r_rsCascadeRequest!HOURS)
        !BANK_CODE = Left(l_strBankCode, 1)
        !BANK_BUCKET = DetermineBankBucket(l_strBankCode)
        .Update
    End With
    
    Exit Sub
           
eh:
    TrsRaiseError m_strCMODULENAME, "InsertRequestForCascade"
End Sub

Private Function CanOnlyTakeVacationInCalendarYear( _
    ByVal v_strBenefitsPackage As String _
    ) As Boolean
    
'RULE0008
    
    On Error GoTo eh
    
    If Not RecordsExist(m_rsPriorCalendarYearVacationRule) Then
        CanOnlyTakeVacationInCalendarYear = False
        Exit Function
    End If
    
    m_rsPriorCalendarYearVacationRule.Filter = "RG_VALUE_CD = '" & Trim(v_strBenefitsPackage) & "'"
    If Not RecordsExist(m_rsPriorCalendarYearVacationRule) Then
        CanOnlyTakeVacationInCalendarYear = False
    Else
        CanOnlyTakeVacationInCalendarYear = True
    End If
    m_rsPriorCalendarYearVacationRule.Filter = ""
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "CanOnlyTakeVacationInCalendarYear"
    
End Function

Private Function GetCalendarYearVacationRule( _
    ) As ADODB.Recordset
    
'RULE0008
'This will get a list of Benefits packages that can only use vacation in calendar year allotted
'Management and AmerenUE contract employees may use prior year vactaion hours in the first
'week of the new year.  The first week must have started in the prior year.
'For example these employees can use remaining 2003 vacation hours on
'   01/01/04, 01/02/04, and 01/03/04.  01/04/04 is a Sunday and starts the next week.
'Only 2004 vacation hours can be used for 01/04/04 through 12/31/04.
'All other employees take vacation based on Calendar year only.
'For example 2004 vacation hours can only be used from 01/01/04 through 12/31/04


    Dim l_obj As dbRules.cdRules

    On Error GoTo eh

    Set l_obj = CreateObject("dbRules.cdRules")
    Set GetCalendarYearVacationRule = l_obj.GetRule(8)
         
    Set l_obj = Nothing
 
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCalendarYearVacationRule"
    
End Function


Private Function DetermineBankYear( _
    ByVal v_strUnionCd As String, _
    ByVal v_datDate As Date _
    ) As Integer

On Error GoTo eh

    Select Case UCase(Trim(v_strUnionCd))
        Case "ND"
            DetermineBankYear = BankYearCILCOLocal8(v_datDate)
            Exit Function
        Case Else
    End Select
        
    
    If Month(v_datDate) < 5 Then
        If HasMay_AprilCalYear(v_strUnionCd) Then
            DetermineBankYear = Year(v_datDate) - 1
        Else
            DetermineBankYear = Year(v_datDate)
        End If
    Else
        DetermineBankYear = Year(v_datDate)
    End If
    
    Exit Function
eh:

    TrsRaiseError m_strCMODULENAME, "DetermineBankYear"
      
    
End Function

Private Function GetUnionsWith_May_April_CalYear( _
    ) As ADODB.Recordset
    
'RULE0120
'The employess in this union get

    Dim l_obj As dbRules.cdRules

    On Error GoTo eh

    Set l_obj = CreateObject("dbRules.cdRules")
    Set GetUnionsWith_May_April_CalYear = l_obj.GetRule(120)
         
    Set l_obj = Nothing
 
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetUnionsWith_May_April_CalYear"
    
End Function

Private Function HasMay_AprilCalYear( _
    ByVal v_strUnion As String _
    ) As Boolean
    
    Dim l_bln As Boolean
    
    On Error GoTo eh

    m_rsUnionsWithMay_April_VacYear.Filter = "RG_VALUE_CD = '" & Trim(v_strUnion) & "'"
    If RecordsExist(m_rsUnionsWithMay_April_VacYear) Then
        HasMay_AprilCalYear = True
    Else
        HasMay_AprilCalYear = False
    End If
    m_rsUnionsWithMay_April_VacYear.Filter = ""

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "HasMay_AprilCalYear"
    
End Function


Private Function BankYearCILCOLocal8( _
    ByVal v_datDate As Date) As Integer
    
    Select Case v_datDate
        Case CDate("2005-01-10") To CDate("2006-01-08")
            BankYearCILCOLocal8 = 2005
        Case CDate("2006-01-09") To CDate("2007-01-07")
            BankYearCILCOLocal8 = 2006
        Case Else
            BankYearCILCOLocal8 = Year(v_datDate)
    End Select
        

End Function


Private Function GetEmpHolidayBankAudit( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datStartDate As Date _
                ) As ADODB.Recordset

On Error GoTo eh

    Dim l_dbAccrual As dbAccrualBankMgr.cdAccrualBankMgr
    
    Set l_dbAccrual = CreateObject("dbAccrualBankMgr.cdAccrualBankMgr")
    
    Set GetEmpHolidayBankAudit = l_dbAccrual.GetEmployeeHolidayBankAudit(v_aryEmployees, v_datStartDate)
         
    Set l_dbAccrual = Nothing
    
    Exit Function

eh:
    Set l_dbAccrual = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmpHolidayBankAudit"
End Function


Private Sub MaintainEmpHolidayBankAudit( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datDate As Date, _
                ByVal v_ExceptionCode As String, _
                ByVal v_Hours As Variant, _
                ByVal v_BankCode As String, _
                ByVal v_BankBucket As String, _
                ByVal v_BankYear As Integer, _
                ByVal v_Reason As String, _
                ByRef l_rsEmpHolidayBankAudit As ADODB.Recordset)
                
                
    Dim l_PriorHours As Variant
    Dim l_NewHours As Variant
    l_rsEmpHolidayBankAudit.Filter = MDAC25FilterFix(" EMP_NUM = " & str(v_lngEmpNum) & " and CHARGE_DATE = '" & v_datDate & "'")
    If l_rsEmpHolidayBankAudit.BOF And l_rsEmpHolidayBankAudit.EOF Then
        l_NewHours = v_Hours
    Else
        l_PriorHours = 0
        l_rsEmpHolidayBankAudit.MoveFirst
        While Not l_rsEmpHolidayBankAudit.EOF
            l_PriorHours = l_PriorHours + l_rsEmpHolidayBankAudit!HOURS
            l_rsEmpHolidayBankAudit.MoveNext
        Wend
        l_NewHours = l_PriorHours + v_Hours
        l_rsEmpHolidayBankAudit.MoveFirst
        While Not l_rsEmpHolidayBankAudit.EOF
            l_rsEmpHolidayBankAudit.Delete
            If Not l_rsEmpHolidayBankAudit.EOF Then
                l_rsEmpHolidayBankAudit.MoveNext
            End If
        Wend
    End If


    If l_NewHours > 0 Then
        l_rsEmpHolidayBankAudit.AddNew
        l_rsEmpHolidayBankAudit!EMP_NUM = v_lngEmpNum
        l_rsEmpHolidayBankAudit!CHARGE_DATE = v_datDate
        l_rsEmpHolidayBankAudit!EFF_YEAR = v_BankYear
        l_rsEmpHolidayBankAudit!HOURS = v_Hours
        l_rsEmpHolidayBankAudit.Update
    End If

    l_rsEmpHolidayBankAudit.Filter = ""
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "MaintainEmpHolidayBankAudit"

End Sub

Private Sub ProcessParentalBoundingRequests(ByVal v_lngEmployee As Long, _
                    ByVal v_rsRequest As ADODB.Recordset, _
                    ByRef r_rsEmpBanks As ADODB.Recordset, _
                    ByRef r_rsBRVs As ADODB.Recordset)
    
    Dim l_datChargeDate As Date
    Dim l_strExceptionCode As String
    Dim l_decHours As Variant
    Dim l_strBucket As String
    Dim l_bolUsedBucket As Boolean
    Dim l_bolSchedBucket As Boolean
    Dim l_lngErrorNumber As Long
    Dim l_strFilter As String
    
    On Error GoTo ProcessParentalBoundingRequests_EH
    
    '//l_strFilter = "EMP_NUM = " & str(v_lngEmployee) & " and EXCEPTION_CD = 'PBL'"
    '//l_strFilter = "EMP_NUM = " & str(v_lngEmployee) & " and ( EXCEPTION_CD = 'PBL' or EXCEPTION_CD = 'PBH')  "
    l_strFilter = "EMP_NUM = " & str(v_lngEmployee) & " and BANK_CODE = 'P' "

    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter)
    If v_rsRequest.EOF Then
        v_rsRequest.Filter = ""
        Exit Sub 'No PBL Banks to update
    End If
'Processing Deletes
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'D' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "CHARGE_DATE DESC"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!EXCEPTION_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        If l_strBucket = "U" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = False
        End If
        If l_strBucket = "S" Then
            l_bolSchedBucket = True
            l_bolUsedBucket = False
        End If
        If l_strBucket = "B" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = True
        End If
        UpdateParentalBoundingBanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolUsedBucket, l_bolSchedBucket, r_rsEmpBanks, r_rsBRVs
        v_rsRequest.MoveNext
    Wend

'Processing Inserts
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'I' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "CHARGE_DATE"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!EXCEPTION_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        If l_strBucket = "U" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = False
        End If
        If l_strBucket = "S" Then
            l_bolSchedBucket = True
            l_bolUsedBucket = False
        End If
        If l_strBucket = "B" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = True
        End If
        UpdateParentalBoundingBanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolUsedBucket, l_bolSchedBucket, r_rsEmpBanks, r_rsBRVs
        v_rsRequest.MoveNext
    Wend
    
    v_rsRequest.Filter = ""
    Exit Sub
    
ProcessParentalBoundingRequests_EH:
    TrsRaiseError m_strCMODULENAME, "ProcessParentalBoundingRequests"
                    
End Sub



Private Sub UpdateParentalBoundingBanks(ByVal v_lngEmpNum As Long, _
                            ByVal v_datDate As Date, _
                            ByVal v_decHours As Variant, _
                            ByVal v_strExceptionCode As String, _
                            ByVal v_bolUsedBucket As Boolean, _
                            ByVal v_bolSchedBucket As Boolean, _
                            ByRef r_rsBanks As ADODB.Recordset, _
                            ByRef r_rsBRVs As ADODB.Recordset)
    
    Dim l_intBankYear As Integer
    Dim l_decAllottedPBLhours As Variant
    Dim l_decUsedPBLhours As Variant
    Dim l_decSchedPBLhours As Variant
    Dim l_decAvailPBLhours As Variant
    Dim l_colDescArgs As Collection
    Dim l_lngErrorNumber As Long
    Dim l_bolBankBalanceError As Boolean
    Dim ldateParentalBoundingStart As Date
    Dim ldateCompare As Date
    Dim l_lngError As Long
    
    On Error GoTo UpdateParentalBoundingBanks_EH
    
    l_lngError = 0
    
    If v_bolSchedBucket Or v_bolUsedBucket Then
    Else
        Exit Sub
    End If
    l_intBankYear = Year(v_datDate)
    ldateCompare = DateAdd("yyyy", -1, v_datDate)
    
    l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intBankYear, "P", r_rsBanks, l_decAllottedPBLhours, l_decAvailPBLhours, l_decUsedPBLhours, l_decSchedPBLhours, ldateParentalBoundingStart)
    If l_bolBankBalanceError Then
    'raise error bank year not found '4 app error
        Err.Raise trErrbusEmployeeDoesNotHaveBank
    End If
    
    If ldateParentalBoundingStart = CDate("0100-01-01") Then
        'get prior year date
        l_intBankYear = l_intBankYear - 1
        l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intBankYear, "P", r_rsBanks, l_decAllottedPBLhours, l_decAvailPBLhours, l_decUsedPBLhours, l_decSchedPBLhours, ldateParentalBoundingStart)
        If l_bolBankBalanceError Then
            l_intBankYear = l_intBankYear + 1
            l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intBankYear, "P", r_rsBanks, l_decAllottedPBLhours, l_decAvailPBLhours, l_decUsedPBLhours, l_decSchedPBLhours, ldateParentalBoundingStart)
        Else
            If ldateParentalBoundingStart = CDate("0100-01-01") Then 'not set in prior year
               l_intBankYear = l_intBankYear + 1
                l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intBankYear, "P", r_rsBanks, l_decAllottedPBLhours, l_decAvailPBLhours, l_decUsedPBLhours, l_decSchedPBLhours, ldateParentalBoundingStart)
            End If
        End If
    Else
        If ldateParentalBoundingStart > v_datDate Then
            'get prior year date
            l_intBankYear = l_intBankYear - 1
            l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intBankYear, "P", r_rsBanks, l_decAllottedPBLhours, l_decAvailPBLhours, l_decUsedPBLhours, l_decSchedPBLhours, ldateParentalBoundingStart)
        End If
    End If
    
    
    If ldateParentalBoundingStart < ldateCompare Then
        If ldateParentalBoundingStart = CDate("0100-01-01") Then
'''''            l_lngError = 4
'''''            Err.Raise 40180 + vbObjectError, , "Parental Bonding Birth Date not set "
'''''                    'Err.Raise trErrbusNotEnoughFMLA_HoursAvailable
            ldateParentalBoundingStart = v_datDate
            l_intBankYear = Year(v_datDate)
            r_rsBanks.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmpNum) & " and EFF_YEAR = " & str(l_intBankYear))
            r_rsBanks!PARENTAL_BOUNDING_START_DATE = v_datDate
            r_rsBanks.Update
        Else
            l_lngError = 4
            Err.Raise 40178 + vbObjectError, , "Parental Bonding Birth has been more than a year " & CStr(ldateParentalBoundingStart)
        End If
    End If
    

    If v_bolSchedBucket And v_bolUsedBucket Then
    Else
        l_decAvailPBLhours = l_decAvailPBLhours - v_decHours
    End If
    
    If l_decAvailPBLhours < 0 Then
    'BRV unavailable PBL hours *
        l_lngError = 2
        Err.Raise 40179 + vbObjectError, , "Not enough Parental Bonding hours available "
    End If
    
    r_rsBanks.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmpNum) & " and EFF_YEAR = " & str(l_intBankYear))
    If v_bolUsedBucket Then
        l_decUsedPBLhours = l_decUsedPBLhours + v_decHours
        If l_decUsedPBLhours < 0 Then
            l_lngError = 3
            Err.Raise 30359 + vbObjectError, , "Parental Bonding hours out of balance "
        End If
        r_rsBanks!PARENTAL_BOUNDING_HRS_USED = l_decUsedPBLhours
        AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "P", "U", l_intBankYear, " "
        If v_bolSchedBucket Then
            l_decSchedPBLhours = l_decSchedPBLhours - v_decHours
            If l_decSchedPBLhours < 0 Then
                l_lngError = 3
                Err.Raise 30359 + vbObjectError, , "Parental Bonding hours out of balance "
            End If
            r_rsBanks!PARENTAL_BOUNDING_HRS_SCHED = l_decSchedPBLhours
            AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours * -1, "P", "S", l_intBankYear, " "
        End If
    Else
        l_decSchedPBLhours = l_decSchedPBLhours + v_decHours
        If l_decSchedPBLhours < 0 Then
            l_lngError = 3
            Err.Raise 30359 + vbObjectError, , "Parental Bonding hours out of balance "
        End If
        r_rsBanks!PARENTAL_BOUNDING_HRS_SCHED = l_decSchedPBLhours
        AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "P", "S", l_intBankYear, " "
    End If
    r_rsBanks.Update
    r_rsBanks.Filter = ""
    Exit Sub
    
UpdateParentalBoundingBanks_EH:

    Set m_rsAuditTrail = Nothing
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(v_lngEmpNum)
    l_colDescArgs.Add str(l_intBankYear)
    l_lngErrorNumber = Err.Number
    Select Case l_lngErrorNumber
        Case trErrbusNotEnoughFMLA_HoursAvailable
                Err.Source = FormatBRV(TrType_TC, v_lngEmpNum, v_datDate, v_datDate, "", "", v_decHours, v_strExceptionCode, "")
                TrsRaiseError m_strCMODULENAME, "UpdateParentalBoundingBanks", l_lngErrorNumber, , l_colDescArgs
        Case trErrbusOutofBalanceFMLA_Hours, _
             trErrbusEmployeeDoesNotHaveBank
                TrsRaiseError m_strCMODULENAME, "UpdateParentalBoundingBanks", l_lngErrorNumber, , l_colDescArgs, , True
        Case Else
            Select Case l_lngError
                Case 1
                    TrsRaiseError m_strCMODULENAME, "UpdateParentalBoundingBanks", l_lngErrorNumber, , l_colDescArgs, , True
                Case 2
                    Err.Source = FormatBRV(TrType_TC, v_lngEmpNum, v_datDate, v_datDate, "", "", v_decHours, v_strExceptionCode, "")
                    TrsRaiseError m_strCMODULENAME, "UpdateParentalBoundingBanks", l_lngErrorNumber, , l_colDescArgs
                Case 3
                    TrsRaiseError m_strCMODULENAME, "UpdateParentalBoundingBanks", l_lngErrorNumber, , l_colDescArgs
                Case 4
                    Err.Source = FormatBRV(TrType_TC, v_lngEmpNum, v_datDate, v_datDate, "", "", v_decHours, v_strExceptionCode, "")
                    TrsRaiseError m_strCMODULENAME, "UpdateParentalBoundingBanks", l_lngErrorNumber, , l_colDescArgs
                Case Else
                    TrsRaiseError m_strCMODULENAME, "UpdateParentalBoundingBanks"
        End Select
    End Select
    
End Sub


Private Sub ProcessFamilyCareRequests(ByVal v_lngEmployee As Long, _
                    ByVal v_rsRequest As ADODB.Recordset, _
                    ByRef r_rsEmpBanks As ADODB.Recordset, _
                    ByRef r_rsBRVs As ADODB.Recordset)
    
    Dim l_datChargeDate As Date
    Dim l_strExceptionCode As String
    Dim l_decHours As Variant
    Dim l_strBucket As String
    Dim l_bolUsedBucket As Boolean
    Dim l_bolSchedBucket As Boolean
    Dim l_lngErrorNumber As Long
    Dim l_strFilter As String
    
    On Error GoTo ProcessFamilyCareRequests_EH
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmployee) & " and OVER_EXC_CD = 'FCL'"
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter)
    If v_rsRequest.EOF Then
        v_rsRequest.Filter = ""
        Exit Sub 'No FCL Banks to update
    End If
'Processing Deletes
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'D' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "CHARGE_DATE DESC"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!OVER_EXC_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        If l_strBucket = "U" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = False
        End If
        If l_strBucket = "S" Then
            l_bolSchedBucket = True
            l_bolUsedBucket = False
        End If
        If l_strBucket = "B" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = True
        End If
        UpdateFamilyCareBanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolUsedBucket, l_bolSchedBucket, r_rsEmpBanks, r_rsBRVs
        v_rsRequest.MoveNext
    Wend

'Processing Inserts
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter & " and REC_STATUS = 'I' ")
    If Not v_rsRequest.EOF And Not v_rsRequest.BOF Then
        v_rsRequest.Sort = "CHARGE_DATE"
        v_rsRequest.MoveFirst
    End If
    While Not v_rsRequest.EOF
        l_datChargeDate = v_rsRequest!CHARGE_DATE
        l_strExceptionCode = v_rsRequest!OVER_EXC_CD
        l_strBucket = UCase(v_rsRequest!BANK_BUCKET)
        l_decHours = CDec(v_rsRequest!HOURS)
        If l_strBucket = "U" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = False
        End If
        If l_strBucket = "S" Then
            l_bolSchedBucket = True
            l_bolUsedBucket = False
        End If
        If l_strBucket = "B" Then
            l_bolUsedBucket = True
            l_bolSchedBucket = True
        End If
        UpdateFamilyCareBanks v_lngEmployee, l_datChargeDate, l_decHours, l_strExceptionCode, l_bolUsedBucket, l_bolSchedBucket, r_rsEmpBanks, r_rsBRVs
        v_rsRequest.MoveNext
    Wend
    
    v_rsRequest.Filter = ""
    Exit Sub
    
ProcessFamilyCareRequests_EH:
    TrsRaiseError m_strCMODULENAME, "ProcessFamilyCareRequests"
                    
End Sub



Private Sub UpdateFamilyCareBanks(ByVal v_lngEmpNum As Long, _
                            ByVal v_datDate As Date, _
                            ByVal v_decHours As Variant, _
                            ByVal v_strExceptionCode As String, _
                            ByVal v_bolUsedBucket As Boolean, _
                            ByVal v_bolSchedBucket As Boolean, _
                            ByRef r_rsBanks As ADODB.Recordset, _
                            ByRef r_rsBRVs As ADODB.Recordset)
    
    Dim l_intBankYear As Integer
    Dim l_decAllottedFCLhours As Variant
    Dim l_decUsedFCLhours As Variant
    Dim l_decSchedFCLhours As Variant
    Dim l_decAvailFCLhours As Variant
    Dim l_colDescArgs As Collection
    Dim l_lngErrorNumber As Long
    Dim l_bolBankBalanceError As Boolean
    Dim l_bolRaiseError As Boolean
    
    On Error GoTo UpdateFamilyCareBanks_EH
    
    If v_bolSchedBucket Or v_bolUsedBucket Then
    Else
        Exit Sub
    End If
    l_bolRaiseError = False
    l_intBankYear = Year(v_datDate)
    
    l_bolBankBalanceError = GetBankBalances(v_lngEmpNum, l_intBankYear, "S", r_rsBanks, l_decAllottedFCLhours, l_decAvailFCLhours, l_decUsedFCLhours, l_decSchedFCLhours)
    If l_bolBankBalanceError Then
    'raise error bank year not found '4 app error
        Err.Raise trErrbusEmployeeDoesNotHaveBank
    End If

    If v_bolSchedBucket And v_bolUsedBucket Then
    Else
        l_decAvailFCLhours = l_decAvailFCLhours - v_decHours
    End If
    
    If l_decAvailFCLhours < 0 Then
    'BRV unavailable FMLA hours *
        l_bolRaiseError = True
        Err.Raise 40181 + vbObjectError, , "Not enough Family Care hours available "
    End If
    
    r_rsBanks.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmpNum) & " and EFF_YEAR = " & str(l_intBankYear))
    If v_bolUsedBucket Then
        l_decUsedFCLhours = l_decUsedFCLhours + v_decHours
        If l_decUsedFCLhours < 0 Then
            l_bolRaiseError = True
            Err.Raise 30360 + vbObjectError, , "Family Care hours out of balance "
        End If
        r_rsBanks!FAMILY_CARE_HRS_USED = l_decUsedFCLhours
        AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "S", "U", l_intBankYear, " "
        If v_bolSchedBucket Then
            l_decSchedFCLhours = l_decSchedFCLhours - v_decHours
            If l_decSchedFCLhours < 0 Then
                l_bolRaiseError = True
                Err.Raise 30360 + vbObjectError, , "Family Care hours out of balance "
            End If
            r_rsBanks!FAMILY_CARE_HRS_SCHED = l_decSchedFCLhours
            AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours * -1, "S", "S", l_intBankYear, " "
        End If
    Else
        l_decSchedFCLhours = l_decSchedFCLhours + v_decHours
        If l_decSchedFCLhours < 0 Then
            l_bolRaiseError = True
            Err.Raise 30360 + vbObjectError, , "Family Care hours out of balance "
        End If
        r_rsBanks!FAMILY_CARE_HRS_SCHED = l_decSchedFCLhours
        AddToAuditTrail v_lngEmpNum, v_datDate, v_strExceptionCode, v_decHours, "S", "S", l_intBankYear, " "
    End If
    r_rsBanks.Update
    r_rsBanks.Filter = ""
    Exit Sub
    
UpdateFamilyCareBanks_EH:

    Set m_rsAuditTrail = Nothing
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(v_lngEmpNum)
    l_colDescArgs.Add str(l_intBankYear)
    l_lngErrorNumber = Err.Number
    Select Case l_lngErrorNumber
        Case trErrbusNotEnoughFMLA_HoursAvailable
                Err.Source = FormatBRV(TrType_TC, v_lngEmpNum, v_datDate, v_datDate, "", "", v_decHours, v_strExceptionCode, "")
                TrsRaiseError m_strCMODULENAME, "UpdateFamilyCareBanks", l_lngErrorNumber, , l_colDescArgs
        Case trErrbusOutofBalanceFMLA_Hours, _
             trErrbusEmployeeDoesNotHaveBank
                TrsRaiseError m_strCMODULENAME, "UpdateFamilyCareBanks", l_lngErrorNumber, , l_colDescArgs, , True
        Case Else
            If l_bolRaiseError = True Then
                Err.Source = FormatBRV(TrType_TC, v_lngEmpNum, v_datDate, v_datDate, "", "", v_decHours, v_strExceptionCode, "")
                TrsRaiseError m_strCMODULENAME, "UpdateFamilyCareBanks", l_lngErrorNumber, , l_colDescArgs
            Else
                TrsRaiseError m_strCMODULENAME, "UpdateFamilyCareBanks"
            End If
    End Select
    
End Sub


