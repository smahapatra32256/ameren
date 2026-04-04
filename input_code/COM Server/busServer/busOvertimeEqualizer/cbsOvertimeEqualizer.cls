VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cbsOTEqualizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Const m_lngMINOFFALLOWCONTINUOUS As Long = 1440
Private Const m_strMODULENAME As String = "cbsOvertimeEqualizer"
Private Const m_intNOREFUSERULE As Integer = 0
Private Const m_intREPLACINGMAN As Integer = 1
Private Const m_intHIGHMAN As Integer = 2
Private Const m_lngSystemEmpNumber As Long = 999999999
Private Const m_intStandardBankCode As Integer = 1
Private Const m_intVolunteerBankCode As Integer = 2
'Private m_rsOTOpportunity As ADODB.Recordset
Private m_intOTOppLoc As Integer
Private m_strOTOppStartTime As String
Private m_strOTOppEndTime As String
Private m_datOTOppChargeDate As Date
Private m_datOTOppStartDate As Date
Private m_lngOTOppNum As Long
Private m_strOTShiftCode As String
Private m_intOTGroup As Integer
Private m_intRefuseRuleNum As Integer
Private m_blnFirstPosting As Boolean
Private m_lngLastModEmpNum As Long
Private m_bolQuarterlyExec As Boolean
Private m_strModDesc As String
Private m_intOTGroupStandingBank As Integer


Private m_rsErrors As ADODB.Recordset
Private m_rsOppTempBankCharges As ADODB.Recordset
Private m_rsOppBankCharges As ADODB.Recordset
Private m_rsOTGroupBanks As ADODB.Recordset
Private m_rsSchedule As ADODB.Recordset
Private m_rsAuditTrail As ADODB.Recordset
Private m_rsOTResponses As ADODB.Recordset
Private m_rsOTBankTMCharges As ADODB.Recordset
Private m_rsOTExceptionCodes As ADODB.Recordset
Private m_rsHolidayRestExcCds As ADODB.Recordset
Private m_rsEmployeeOvertimeBanks As ADODB.Recordset
Private m_rsEmployeeOvertimeTempBanks As ADODB.Recordset
Private m_rsCollectionPeriods As ADODB.Recordset
Private m_rsEmployeeOTData As ADODB.Recordset
Private m_rsEmployeeData As ADODB.Recordset
            'This recordset is needed to determine if an
            'employee is an executive or not on a
            'specific date
            ' It contains the following
                    'Charge Date
                    'Union Code
                    'OT Type  - Indicates if employee is a quarterly executive
                            ' This field is driven by the job num
                    'Exec Contract Ind
                    'Shift Ind - Indicates if an employee shift worker
                            ' and gets paid for working OT on Holidays
                            ' without meeting an exclusion.
                            
Public Function UnChargeOvertimeBanksForOpportunity( _
    ByVal v_rsOTOpportunity As ADODB.Recordset, _
    ByVal v_lngLastModEmpNum As Long, _
    ByRef r_rsBrvs As ADODB.Recordset _
    ) As Boolean

    Dim l_rsForamattedRequest As ADODB.Recordset
    
 On Error GoTo eh
        
    m_lngOTOppNum = v_rsOTOpportunity!OT_OPP_NUM
    m_lngLastModEmpNum = v_lngLastModEmpNum
    Set m_rsErrors = CreateNewBusinessErrorSet()
    m_strModDesc = v_rsOTOpportunity!MOD_REASON_DESC
    m_intOTGroup = v_rsOTOpportunity!OT_GROUP_NUM
    m_datOTOppChargeDate = v_rsOTOpportunity!CHARGE_DATE
    m_strOTShiftCode = Trim(v_rsOTOpportunity!SHIFT_CD)
    
    Set m_rsEmployeeData = Nothing
    Set m_rsOppBankCharges = GetOppBankCharges()
    If Not RecordsExist(m_rsOppBankCharges) Then
        GetObjectContext.SetComplete
        UnChargeOvertimeBanksForOpportunity = True
        GoTo XIT
    End If
    GetEmployeeStateForUnPost
    Set m_rsOTGroupBanks = GetOTGroupBanks(m_intOTGroup)
    Set m_rsOTResponses = GetOvertimeResponses()
    Set m_rsOTBankTMCharges = GetOvertimeBankTimeCharges()
    
    BuildRequestForUnPosting l_rsForamattedRequest
    RemoveBankTimeCharges
    
    If Not RecordsExist(l_rsForamattedRequest) Then
        GetObjectContext.SetComplete
        UnChargeOvertimeBanksForOpportunity = True
        GoTo XIT
    End If
    
    
    ProcessOvertimeBankUpdateRequests l_rsForamattedRequest
    GetObjectContext.SetComplete
        
    UnChargeOvertimeBanksForOpportunity = True
XIT:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsBrvs = Nothing
        Else
            Set r_rsBrvs = m_rsErrors
        End If
    Else
        Set r_rsBrvs = Nothing
    End If
'    Set l_rsForamattedRequest = Nothing
    ClearGlobalRecordsets

    Exit Function
    
eh:
' The client is expecting a BRV recordset when errors occur
' Calling TrsRaiseError in this function will cause the
' BRVs to no be passed back, because the statement
' end function was not executed

    Set r_rsBrvs = Nothing
    GetObjectContext.SetAbort
    
    Set l_rsForamattedRequest = Nothing
    ClearGlobalRecordsets
    TrsRaiseError m_strMODULENAME, "UnChargeOvertimeBanksForOpportunity", , , , , True

    
End Function

Public Function UpdateOvertimeBanksForTimeEntry( _
                ByVal v_request As ADODB.Recordset, _
                ByVal v_rsEmployeeData As ADODB.Recordset, _
                ByRef r_rsBrvs As ADODB.Recordset, _
                ByVal v_lngLastModEmpNum As Long _
                    ) As Boolean
                         
    Dim l_rsForamattedRequest As ADODB.Recordset
    
    On Error GoTo UpdateOvertimeBanksForTimeEntry_EH
    
    m_lngLastModEmpNum = v_lngLastModEmpNum
    Set m_rsErrors = CreateNewBusinessErrorSet()
    m_lngOTOppNum = 90000000
    
    If v_request.EOF And v_request.BOF Then
        GetObjectContext.SetComplete
        UpdateOvertimeBanksForTimeEntry = True
        GoTo XIT
    End If
    Set l_rsForamattedRequest = BuildRequestForTimeEntry(v_request)
    If l_rsForamattedRequest.EOF And l_rsForamattedRequest.BOF Then
        GetObjectContext.SetComplete
        UpdateOvertimeBanksForTimeEntry = True
        GoTo XIT
    End If
    
    If v_rsEmployeeData Is Nothing Then
        Set m_rsEmployeeData = Nothing
    Else
        Set m_rsEmployeeData = v_rsEmployeeData
    End If
    
    ProcessOvertimeBankUpdateRequests l_rsForamattedRequest
    GetObjectContext.SetComplete
    UpdateOvertimeBanksForTimeEntry = True
    
XIT:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsBrvs = Nothing
        Else
            Set r_rsBrvs = m_rsErrors
        End If
    Else
        Set r_rsBrvs = Nothing
    End If
    Set l_rsForamattedRequest = Nothing
    ClearGlobalRecordsets

    Exit Function
    
UpdateOvertimeBanksForTimeEntry_EH:
' The client is expecting a BRV recordset when errors occur
' Calling TrsRaiseError in this function will cause the
' BRVs to no be passed back, because the statement
' end function was not executed

    Set r_rsBrvs = Nothing
    GetObjectContext.SetAbort
    
    Set l_rsForamattedRequest = Nothing
    ClearGlobalRecordsets
    TrsRaiseError m_strMODULENAME, "UpdateBanksForTimeEntry", , , , , True

End Function

Private Function BuildRequestForTimeEntry( _
                    ByVal v_TimeEntryRequest As ADODB.Recordset _
                    ) As ADODB.Recordset
                      
    Dim l_rsRequest As ADODB.Recordset
    Dim l_strTimeCardStatus As String
    Dim l_strBankCode As String
    
    On Error GoTo eh
    
    Set l_rsRequest = CreateRequestRecordset()
    
    If v_TimeEntryRequest.EOF And v_TimeEntryRequest.BOF Then
        Set BuildRequestForTimeEntry = l_rsRequest
        Set l_rsRequest = Nothing
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
    
    Exit Function

eh:
    Set l_rsRequest = Nothing
    TrsRaiseError m_strMODULENAME, "BuildRequestForTimeEntry"
                  
End Function


Private Function CreateRequestRecordset() As ADODB.Recordset

    Dim l_rsRequest As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rsRequest = New ADODB.Recordset
    With l_rsRequest
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("HOURS").Precision = 3
      .Fields.Item("HOURS").NumericScale = 1
      .Fields.Append "REC_STATUS", adChar, 1, adFldUpdatable
      .Fields.Append "BANK_CHARGE_DATE", adDate, , adFldMayBeNull
      .Fields.Append "OT_MOD_DESC", adChar, 30, adFldUpdatable
      .Fields.Append "OT_BANK_CODE", adSmallInt, , adFldMayBeNull
      .Fields.Append "OT_RESPONSE", adChar, 5, adFldUpdatable
      .Fields.Append "OT_OPP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "TIME_ENTRY_IND", adChar, 1, adFldUpdatable
      .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
     'New for TRIS 5
      .Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
      .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "SHIFT_CD", adChar, 1, adFldUpdatable
      .Fields.Append "ACC_REF_IND", adChar, 1, adFldUpdatable
      .Fields.Append "CALL_TIME", adChar, 26, adFldUpdatable
      .Fields.Append "HOURS_CHARGED", adDecimal, , adFldUpdatable
      .Fields.Item("HOURS_CHARGED").Precision = 3
      .Fields.Item("HOURS_CHARGED").NumericScale = 1
      .Fields.Append "TRAVEL_PAY_HRS", adDecimal, , adFldUpdatable
      .Fields.Item("TRAVEL_PAY_HRS").Precision = 3
      .Fields.Item("TRAVEL_PAY_HRS").NumericScale = 1
      
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockBatchOptimistic
      .Open
    End With
    
    Set CreateRequestRecordset = l_rsRequest
    Set l_rsRequest = Nothing
    
    Exit Function

eh:
    Set l_rsRequest = Nothing
    TrsRaiseError m_strMODULENAME, "CreateRequestRecordset"

End Function


Private Function CreateAuditTrailRecordset() As ADODB.Recordset

    Dim l_rsRequest As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rsRequest = New ADODB.Recordset
    With l_rsRequest
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "REASON", adChar, 30, adFldUpdatable
      .Fields.Append "HRS_ADJ", adDecimal, , adFldUpdatable
      .Fields.Item("HRS_ADJ").Precision = 7
      .Fields.Item("HRS_ADJ").NumericScale = 2
      .Fields.Append "OT_BANK_CD", adSmallInt, , adFldUpdatable
      .Fields.Append "OT_OPP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "ACC_REF_IND", adChar, 1, adFldUpdatable
      .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
      .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "SHIFT_CD", adChar, 1, adFldUpdatable
      .Fields.Append "HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("HOURS").Precision = 3
      .Fields.Item("HOURS").NumericScale = 1
      .Fields.Append "RATE", adDecimal, , adFldUpdatable
      .Fields.Item("RATE").Precision = 3
      .Fields.Item("RATE").NumericScale = 2
      
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockBatchOptimistic
      .Open
    End With
    
    Set CreateAuditTrailRecordset = l_rsRequest
    Set l_rsRequest = Nothing
    
    
    Exit Function

eh:
    Set l_rsRequest = Nothing
    TrsRaiseError m_strMODULENAME, "CreateAuditTrailRecordset"
    
    

End Function


Private Sub ProcessOvertimeBankUpdateRequests(ByVal v_rsRequest As ADODB.Recordset)
    
    Dim l_colEmployees As Collection
    Dim l_intNumOfEmps As Integer
    Dim l_intI As Integer
    Dim l_lngEmployee As Long
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
'    Dim l_rsEmployeeOvertimeBanks As ADODB.Recordset
    Dim l_strFilter As String
    Dim l_datChargeDate As Date
    Dim l_aryEmployees() As String
    
    On Error GoTo ProcessOvertimeBankUpdateRequests_EH
    
    
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
    
    v_rsRequest.Sort = "CHARGE_DATE"
    v_rsRequest.MoveFirst
    l_datStartDate = v_rsRequest!CHARGE_DATE
    v_rsRequest.MoveLast
    l_datEndDate = v_rsRequest!CHARGE_DATE
    
    GetState l_aryEmployees, l_datStartDate, l_datEndDate
    l_intI = 0
    While l_intI < l_intNumOfEmps
'        Set m_rsAuditTrail = CreateAuditTrailRecordset()
        l_intI = l_intI + 1
        l_lngEmployee = l_colEmployees.Item(l_intI)
        l_strFilter = ""
        v_rsRequest.Filter = "EMP_NUM = " & str(l_lngEmployee)
        v_rsRequest.Sort = "CHARGE_DATE"

        ProcessOvertimeBankUpdates l_lngEmployee, v_rsRequest
    Wend
    SaveUpdateToEmployeeOTBanks
    Set l_colEmployees = Nothing
    Exit Sub
    
ProcessOvertimeBankUpdateRequests_EH:
    
    Set l_colEmployees = Nothing
    TrsRaiseError m_strMODULENAME, "ProcessOvertimeBankUpdateRequests"
                            
End Sub

Private Function GetChargeAsPayInd( _
            ByVal v_strUnionCode As String _
            ) As String
            
    On Error GoTo eh
    
    Select Case Trim(v_strUnionCode)
        Case "OA", "GA", "LA", "PA", "VA" _
            , "IP", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BJ", "BK", "BL", "BM"
            GetChargeAsPayInd = "Y"
        Case Else
            GetChargeAsPayInd = "N"
    End Select
    
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "GetChargeAsPayInd"
            
End Function


Private Sub InsertRequestForTimeEntry( _
                  ByRef r_rsRequest As ADODB.Recordset, _
                  ByVal v_TimeEntryRequest As ADODB.Recordset)
                  
    Dim l_strBankCode As String
    Dim l_strStrScheduleInd As String
    Dim l_decHours As Variant
    Dim l_decOrigHours As Variant
    Dim l_decTravelHours As Variant
    Dim l_decOrigTravelHours As Variant
    Dim l_decTotalHours As Variant
    Dim l_decOrigTotalHours As Variant
    
    On Error GoTo eh
    
    l_strStrScheduleInd = v_TimeEntryRequest!SCHEDULED_IND
    l_strBankCode = UCase(Trim(v_TimeEntryRequest!BANK_CODE))
    l_strBankCode = Left(l_strBankCode, 1)
    
    If l_strBankCode = "O" Then
    'Overtime Bank Exception
    Else
    ' Does not affect OT bank
        Exit Sub
    End If
    
    If l_strStrScheduleInd = "Y" Then
    'overtime bank already updated
            Exit Sub
    End If
    
    
    l_decHours = CDec(v_TimeEntryRequest!HOURS)
    l_decTravelHours = CDec(v_TimeEntryRequest!TRAVEL_PAY_HRS)
    l_decTotalHours = l_decHours + l_decTravelHours
    l_decOrigHours = CDec(v_TimeEntryRequest!ORIG_HOURS)
    l_decOrigTravelHours = CDec(v_TimeEntryRequest!ORIG_TRAVEL_PAY_HRS)
    l_decOrigTotalHours = l_decOrigHours + l_decOrigTravelHours
    
    If v_TimeEntryRequest!REC_STATUS = "C" Then
        If v_TimeEntryRequest!EXCEPTION_CD = v_TimeEntryRequest!ORIG_EXCEPTION_CD And _
            l_decTotalHours = l_decOrigTotalHours Then
    'Insert is from an update and pertinent fields did not change
    'Such as the job_number was changed
            Exit Sub
        End If
    End If
    
' Add Request
    r_rsRequest.AddNew
    r_rsRequest!REC_STATUS = "I"
    r_rsRequest!EMP_NUM = v_TimeEntryRequest!EMP_NUM
    r_rsRequest!CHARGE_DATE = v_TimeEntryRequest!CHARGE_DATE
    r_rsRequest!EXCEPTION_CD = v_TimeEntryRequest!EXCEPTION_CD
    r_rsRequest!HOURS = l_decHours
    r_rsRequest!TRAVEL_PAY_HRS = l_decTravelHours
    r_rsRequest!OT_BANK_CODE = 1
    r_rsRequest!OT_RESPONSE = "Y"
    r_rsRequest!OT_OPP_NUM = v_TimeEntryRequest!OT_OPP_NUM
    r_rsRequest!TIME_ENTRY_IND = "Y"
    r_rsRequest!START_TIME = v_TimeEntryRequest!START_TIME
    r_rsRequest!OVER_EXC_CD = v_TimeEntryRequest!OVER_EXC_CD
'Date worked is not on the Time Entry Request may need to add it
'for now assuming date worked is the same date as charge date
'Also probably need end time as well
'Ask Rick for this stuff
    r_rsRequest!DATE_WORKED = v_TimeEntryRequest!CHARGE_DATE
    r_rsRequest!END_TIME = v_TimeEntryRequest!START_TIME
    
    If v_TimeEntryRequest!REC_STATUS = "C" Then
        If v_TimeEntryRequest!OT_BANK_CHRG_DT = CDate("01/01/0100") Then
            r_rsRequest!BANK_CHARGE_DATE = Null
        Else
            r_rsRequest!BANK_CHARGE_DATE = v_TimeEntryRequest!OT_BANK_CHRG_DT
        End If
        r_rsRequest!OT_MOD_DESC = v_TimeEntryRequest!OT_MOD_DESC
    Else
        r_rsRequest!BANK_CHARGE_DATE = Null
        r_rsRequest!OT_MOD_DESC = "Posted New Time Entry"
    End If
    r_rsRequest.Update
    
    Exit Sub

eh:
    TrsRaiseError m_strMODULENAME, "InsertRequestForTimeEntry"

                      
End Sub

Private Sub DeleteRequestForTimeEntry( _
                  ByRef r_rsRequest As ADODB.Recordset, _
                  ByVal v_TimeEntryRequest As ADODB.Recordset)
                  
    Dim l_strBankCode As String
    Dim l_decHours As Variant
    Dim l_decOrigHours As Variant
    Dim l_decTravelHours As Variant
    Dim l_decOrigTravelHours As Variant
    Dim l_decTotalHours As Variant
    Dim l_decOrigTotalHours As Variant
    
    On Error GoTo eh
    
    l_strBankCode = UCase(Trim(v_TimeEntryRequest!ORIG_BANK_CODE))
    l_strBankCode = Left(l_strBankCode, 1)
    
    If l_strBankCode = "O" Then
    'Overtime Bank Exception
    Else
    ' Does not affect OT bank
        Exit Sub
    End If
        
    l_decHours = CDec(v_TimeEntryRequest!HOURS)
    l_decTravelHours = CDec(v_TimeEntryRequest!TRAVEL_PAY_HRS)
    l_decTotalHours = l_decHours + l_decTravelHours
    l_decOrigHours = CDec(v_TimeEntryRequest!ORIG_HOURS)
    l_decOrigTravelHours = CDec(v_TimeEntryRequest!ORIG_TRAVEL_PAY_HRS)
    l_decOrigTotalHours = l_decOrigHours + l_decOrigTravelHours
    
    If v_TimeEntryRequest!REC_STATUS = "C" Then
        If v_TimeEntryRequest!EXCEPTION_CD = v_TimeEntryRequest!ORIG_EXCEPTION_CD And _
            l_decTotalHours = l_decOrigTotalHours Then
        'Delete is from an update and pertinent fields did not change
        'Such as the job_number was changed
            Exit Sub
        End If
    End If
 ' Add delete request
    r_rsRequest.AddNew
    r_rsRequest!REC_STATUS = "D"
    r_rsRequest!EMP_NUM = v_TimeEntryRequest!EMP_NUM
    r_rsRequest!CHARGE_DATE = v_TimeEntryRequest!CHARGE_DATE
    r_rsRequest!EXCEPTION_CD = v_TimeEntryRequest!ORIG_EXCEPTION_CD
    r_rsRequest!HOURS = l_decOrigHours * -1
    r_rsRequest!TRAVEL_PAY_HRS = l_decOrigTravelHours * -1
    r_rsRequest!OT_BANK_CODE = 1
    If v_TimeEntryRequest!OT_BANK_CHRG_DT = CDate("01/01/0100") Then
        r_rsRequest!BANK_CHARGE_DATE = Null
    Else
        r_rsRequest!BANK_CHARGE_DATE = v_TimeEntryRequest!OT_BANK_CHRG_DT
    End If
    r_rsRequest!OT_MOD_DESC = v_TimeEntryRequest!OT_MOD_DESC
    r_rsRequest!OT_RESPONSE = "Y"
    r_rsRequest!OT_OPP_NUM = v_TimeEntryRequest!ORIG_OT_OPP_NUM
    r_rsRequest!TIME_ENTRY_IND = "Y"
    r_rsRequest!START_TIME = v_TimeEntryRequest!START_TIME
    r_rsRequest!OVER_EXC_CD = v_TimeEntryRequest!ORIG_OVER_EXC_CD
'Date worked is not on the Time Entry Request may need to add it
'for now assuming date worked is the same date as charge date
'Also probably need end time as well
'Ask Rick for this stuff
    r_rsRequest!DATE_WORKED = v_TimeEntryRequest!CHARGE_DATE
    r_rsRequest!END_TIME = v_TimeEntryRequest!START_TIME

    r_rsRequest.Update
        
    Exit Sub

eh:
    TrsRaiseError m_strMODULENAME, "DeleteRequestForTimeEntry"
    
    
End Sub

Private Function EmployeeInList( _
                    ByVal v_lngEmployee As Long, _
                    ByVal v_colEmployees As Collection _
                ) As Boolean


    Dim l_intIndex As Integer
    Dim l_intCount As Integer
    Dim l_bolInList As Boolean
    
    On Error GoTo eh
    
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
    
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "EmployeeInList"
    

End Function


Private Function GetEmpOTData( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As ADODB.Recordset
            

On Error GoTo GetEmpOTData_EH


    Dim l_boCommonLib As boCommonLib.cboEmployee

    
    Set l_boCommonLib = CreateObject("boCommonLib.cboEmployee")
    
    Set GetEmpOTData = l_boCommonLib.GetOTBankChargeInfoByChargeDate(v_aryEmployees, v_datStartDate, v_datEndDate, m_rsEmployeeData)
         
    Set l_boCommonLib = Nothing
    
    
    Exit Function

GetEmpOTData_EH:
    Set l_boCommonLib = Nothing
    GetObjectContext.SetAbort
    
    TrsRaiseError m_strMODULENAME, "GetEmpOTData"
                               
End Function

'' All New Code is Below

Private Function GetEmpOvertimeBanks( _
            ByVal v_aryEmployees As Variant _
            ) As ADODB.Recordset
   
 
On Error GoTo eh


    
    Dim l_dbOvertimeEqualizer As dbOvertimeEqualizer.cdOvertimeEqualizer

    
    Set l_dbOvertimeEqualizer = CreateObject("dbOvertimeEqualizer.cdOvertimeEqualizer")
    
    Set GetEmpOvertimeBanks = l_dbOvertimeEqualizer.GetEmployeeOvertimeBanks(v_aryEmployees)
         
    Set l_dbOvertimeEqualizer = Nothing
    
    
    Exit Function

eh:
    Set l_dbOvertimeEqualizer = Nothing
    TrsRaiseError m_strMODULENAME, "GetEmpOvertimeBanks"
                    
End Function

Private Function GetEmpOvertimeTempBanks( _
            ByVal v_aryEmployees As Variant _
            ) As ADODB.Recordset
   
 
On Error GoTo GetEmpOvertimeTempBanks_EH


    Dim l_dbOvertimeEqualizer As dbOvertimeEqualizer.cdOvertimeEqualizer

    
    Set l_dbOvertimeEqualizer = CreateObject("dbOvertimeEqualizer.cdOvertimeEqualizer")
    
    Set GetEmpOvertimeTempBanks = l_dbOvertimeEqualizer.GetEmployeeOvertimeTempBanks(v_aryEmployees)
         
    Set l_dbOvertimeEqualizer = Nothing
    
    
    Exit Function

GetEmpOvertimeTempBanks_EH:
    Set l_dbOvertimeEqualizer = Nothing
    GetObjectContext.SetAbort
    
    TrsRaiseError m_strMODULENAME, "GetEmpOvertimeTempBanks"
                    
   
End Function



Private Function GetOvertimeResponses( _
            ) As ADODB.Recordset

 On Error GoTo GetOvertimeResponses_EH

    Dim l_dbOvertimeEqualizer As dbOvertimeEqualizer.cdOvertimeEqualizer

    
    Set l_dbOvertimeEqualizer = CreateObject("dbOvertimeEqualizer.cdOvertimeEqualizer")
    
    Set GetOvertimeResponses = l_dbOvertimeEqualizer.GetOvertimeResponses()
    
    Set l_dbOvertimeEqualizer = Nothing
    
    
    Exit Function

GetOvertimeResponses_EH:
    Set l_dbOvertimeEqualizer = Nothing
    GetObjectContext.SetAbort
    
    TrsRaiseError m_strMODULENAME, "GetOvertimeResponses"

End Function

Private Function GetOvertimeExceptions( _
            ) As ADODB.Recordset

 On Error GoTo GetOvertimeExceptions_EH

    Dim l_dbCodeTables As dbCodeTables.cdCodeTables

    
    Set l_dbCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    Set GetOvertimeExceptions = l_dbCodeTables.GetOvertimeExceptions()
         
    Set l_dbCodeTables = Nothing
    
    
    Exit Function

GetOvertimeExceptions_EH:
    Set l_dbCodeTables = Nothing
    GetObjectContext.SetAbort
    
    TrsRaiseError m_strMODULENAME, "GetOvertimeExceptions"
                    
End Function



Private Sub ProcessOvertimeBankUpdates( _
                ByVal v_lngEmployee As Long, _
                ByVal v_rsRequest As ADODB.Recordset)
                
    Dim l_datBankChargeDate As Date
    Dim l_datTimeChargeDate As Date
    Dim l_datTodaysDate As Date
    Dim l_decHours As Variant
    Dim l_strExceptionCode As String
    
    On Error GoTo eh
             
    l_datTodaysDate = Date
'Postion on First Record

    v_rsRequest.Filter = "EMP_NUM = " & str(v_lngEmployee)
    v_rsRequest.Sort = "CHARGE_DATE"
    v_rsRequest.MoveFirst
    
    While Not v_rsRequest.EOF
'Might want to consider always determining bank charge date
'Intially decided not to because I need to get pay period
'Information for union IA employees
        l_datTimeChargeDate = v_rsRequest!CHARGE_DATE
        If IsNull(v_rsRequest!EXCEPTION_CD) Or IsEmpty(v_rsRequest!EXCEPTION_CD) Then
            l_strExceptionCode = ""
        Else
            l_strExceptionCode = Trim(v_rsRequest!EXCEPTION_CD)
        End If
        If IsNull(v_rsRequest!BANK_CHARGE_DATE) Or IsEmpty(v_rsRequest!BANK_CHARGE_DATE) Then
            l_datBankChargeDate = DetermineBankChargeDate(v_lngEmployee, l_datTimeChargeDate, l_strExceptionCode)
            v_rsRequest!BANK_CHARGE_DATE = l_datBankChargeDate
        Else
            l_datBankChargeDate = v_rsRequest!BANK_CHARGE_DATE
        End If

        If l_datBankChargeDate > l_datTimeChargeDate And _
            l_datBankChargeDate > l_datTodaysDate Then
            UpdateTemporaryOTBanks v_lngEmployee, v_rsRequest, l_datBankChargeDate
        Else
            UpdatePermanentOTBanks v_lngEmployee, v_rsRequest
        End If
        v_rsRequest.MoveNext
        
    Wend
                
    Exit Sub

eh:
    TrsRaiseError m_strMODULENAME, "ProcessOvertimeBankUpdates"
    
End Sub

Private Function DetermineBankChargeDate( _
                    ByVal v_lngEmployee As Long, _
                    ByVal v_datChargeDate As Date, _
                    Optional v_strExceptionCode As String = "" _
                    ) As Date
                    
    Dim l_datBankChargeDate As Date
    Dim l_datPayPeriodEndDate As Date
    Dim l_datTodaysDate As Date
    Dim l_dat30DaysPriorToToday As Date
    
    On Error GoTo eh
             
    l_datTodaysDate = Date
    
        
    Select Case GetBankChargeDateRule(v_lngEmployee, v_datChargeDate, v_strExceptionCode)
        Case 0 'Charge immediately using system date
            l_datBankChargeDate = l_datTodaysDate
            
        Case 1 'Charge Date of the Overtime -- Ameren GOB Management
            l_datBankChargeDate = v_datChargeDate
        
        Case 2 '2 Calendar Sundays from Pay Period End Date -- Union IA
            l_datPayPeriodEndDate = DeterminePayPeriodEndDate(v_lngEmployee, v_datChargeDate)
            l_datBankChargeDate = GetNextSunday(l_datPayPeriodEndDate)
            l_datBankChargeDate = DateAdd("d", 7, l_datBankChargeDate)
        
        Case 3 'Sunday Following End of Pay Period -- WA loc 63 depts
            l_datPayPeriodEndDate = DeterminePayPeriodEndDate(v_lngEmployee, v_datChargeDate)
            l_datBankChargeDate = GetNextSunday(l_datPayPeriodEndDate)
        
        Case 4 'Monday Following End of Pay Period
            l_datPayPeriodEndDate = DeterminePayPeriodEndDate(v_lngEmployee, v_datChargeDate)
            l_datBankChargeDate = GetNextSunday(l_datPayPeriodEndDate)
            l_datBankChargeDate = DateAdd("d", 1, l_datBankChargeDate)
        
        Case 5 'Tuesday Following End of Pay Period
            l_datPayPeriodEndDate = DeterminePayPeriodEndDate(v_lngEmployee, v_datChargeDate)
            l_datBankChargeDate = GetNextSunday(l_datPayPeriodEndDate)
            l_datBankChargeDate = DateAdd("d", 2, l_datBankChargeDate)
        
        Case 6 'Wednesday Following End of Pay Period -- WA loc 72 depts
            l_datPayPeriodEndDate = DeterminePayPeriodEndDate(v_lngEmployee, v_datChargeDate)
            l_datBankChargeDate = GetNextSunday(l_datPayPeriodEndDate)
            l_datBankChargeDate = DateAdd("d", 3, l_datBankChargeDate)
        
        Case 7 '1st of next Month following charge date
                    ' Go to first of month
            l_datBankChargeDate = DateAdd("d", 1 - Day(v_datChargeDate), v_datChargeDate)
                    ' Add one month
            l_datBankChargeDate = DateAdd("m", 1, l_datBankChargeDate)
            
        Case 8 'Monday Following charge date
            l_datBankChargeDate = GetNextMonday(v_datChargeDate)
            
        Case 9 'Wednesday Following charge date
            l_datBankChargeDate = DateAdd("d", 2, GetNextMonday(v_datChargeDate))
            
        Case 10 '2nd Wednesday Following End of Pay Period -- Missouri Call Centers
            l_datPayPeriodEndDate = DeterminePayPeriodEndDate(v_lngEmployee, v_datChargeDate)
            l_datBankChargeDate = GetNextSunday(l_datPayPeriodEndDate)
            l_datBankChargeDate = DateAdd("d", 10, l_datBankChargeDate)
        
            
    End Select
        
    
    DetermineBankChargeDate = l_datBankChargeDate
                    
     Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "DetermineBankChargeDate"

End Function

Private Sub UpdatePermanentOTBank( _
                ByVal v_lngEmployee As Long, _
                ByVal v_decAccHours As Variant, _
                ByVal v_decRefHours As Variant, _
                ByVal v_intOTBank As Integer, _
                ByVal v_rsRequest As ADODB.Recordset, _
                ByVal v_decRate As Variant)
                
    Dim l_datTimeChargeDate As Date
    Dim l_lngOTOppNum As Long
    Dim l_strModDesc As String
    Dim l_strExceptionCode As String
    Dim l_intOTBankCode As Integer
    Dim l_strOTResponse As String
    Dim l_decHours As Variant
    Dim l_decAcceptedHours As Variant
    Dim l_decRefusedHours As Variant
    Dim l_decOTBankTotalHours As Variant
    Dim l_decOTBankRefusedHours As Variant
    Dim l_strFilter As String
    
    On Error GoTo eh
    
    l_decAcceptedHours = v_decAccHours
    l_decRefusedHours = v_decRefHours
    l_datTimeChargeDate = v_rsRequest!CHARGE_DATE
    l_strModDesc = v_rsRequest!OT_MOD_DESC
    l_lngOTOppNum = v_rsRequest!OT_OPP_NUM
    l_strExceptionCode = v_rsRequest!EXCEPTION_CD
    l_decHours = v_rsRequest!HOURS
    l_intOTBankCode = v_intOTBank
    l_strOTResponse = v_rsRequest!OT_RESPONSE
    
    m_rsEmployeeOvertimeBanks.Filter = ""
    If m_rsEmployeeOvertimeBanks.EOF Then
        If BankShouldBeCreated(v_lngEmployee, l_datTimeChargeDate, l_intOTBankCode) Then
            CreateNewOvertimeBank v_lngEmployee, l_intOTBankCode
        Else
            m_rsEmployeeOvertimeBanks.Filter = ""
            Exit Sub
        End If
    Else
        m_rsEmployeeOvertimeBanks.MoveFirst
        l_strFilter = "EMP_NUM = " & str(v_lngEmployee) & _
                    " and OT_BANK_CD = " & str(l_intOTBankCode)
        m_rsEmployeeOvertimeBanks.Filter = MDAC25FilterFix(l_strFilter)
        If m_rsEmployeeOvertimeBanks.EOF Then
            If BankShouldBeCreated(v_lngEmployee, l_datTimeChargeDate, l_intOTBankCode) Then
                CreateNewOvertimeBank v_lngEmployee, l_intOTBankCode
            Else
                m_rsEmployeeOvertimeBanks.Filter = ""
                Exit Sub
            End If
        End If
    End If
    l_decOTBankTotalHours = m_rsEmployeeOvertimeBanks!TOT_OT_HRS
    If l_intOTBankCode = m_intVolunteerBankCode _
        And l_decOTBankTotalHours <= 0 Then
        m_rsEmployeeOvertimeBanks.Filter = ""
        Exit Sub
    End If
    l_decOTBankRefusedHours = m_rsEmployeeOvertimeBanks!TOT_OT_HRS_REF
    l_decOTBankRefusedHours = l_decOTBankRefusedHours + l_decRefusedHours
    l_decOTBankTotalHours = l_decOTBankTotalHours + l_decAcceptedHours + l_decRefusedHours
    m_rsEmployeeOvertimeBanks!TOT_OT_HRS = l_decOTBankTotalHours
    m_rsEmployeeOvertimeBanks!TOT_OT_HRS_REF = l_decOTBankRefusedHours
    m_rsEmployeeOvertimeBanks!LAST_MOD_EMP_NUM = m_lngLastModEmpNum
    m_rsEmployeeOvertimeBanks.Update
    If l_decAcceptedHours <> 0 Then
        AddToAuditTrail v_lngEmployee, l_decAcceptedHours, l_intOTBankCode, "A", l_lngOTOppNum, l_strModDesc, v_rsRequest, v_decRate
    End If
    If l_decRefusedHours <> 0 Then
        AddToAuditTrail v_lngEmployee, l_decRefusedHours, l_intOTBankCode, "R", l_lngOTOppNum, l_strModDesc, v_rsRequest, v_decRate
    End If
    m_rsEmployeeOvertimeBanks.Filter = ""
    
    Exit Sub

eh:
    TrsRaiseError m_strMODULENAME, "UpdatePermanentOTBank"
                    
                
End Sub

Private Sub UpdateTemporaryOTBank( _
                ByVal v_lngEmployee As Long, _
                ByVal v_decAccHours As Variant, _
                ByVal v_decRefHours As Variant, _
                ByVal v_intOTBank As Integer, _
                ByVal v_rsRequest As ADODB.Recordset, _
                ByVal v_datBankChargeDate As Date)

    Dim l_lngOTOppNum As Long
    Dim l_intOTBankCode As Integer
    Dim l_strFilter As String
    Dim l_decTotalHours As Variant
    Dim l_decRefusedHours As Variant
    Dim l_decHoursCharged As Variant
    Dim l_decHours As Variant
    Dim l_strOTResponse As String
    Dim l_strExceptionCode As String
    Dim l_datTimeChargeDate As Date
        
    On Error GoTo eh
            
    l_datTimeChargeDate = v_rsRequest!CHARGE_DATE
    l_lngOTOppNum = v_rsRequest!OT_OPP_NUM
    l_strExceptionCode = v_rsRequest!EXCEPTION_CD
    l_intOTBankCode = v_intOTBank
    l_strOTResponse = v_rsRequest!OT_RESPONSE
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmployee) & _
                " and OT_OPP_NUM = " & str(l_lngOTOppNum) & _
                " and OT_BANK_CD = " & str(l_intOTBankCode) & _
                " and EFF_DATE = " & str(v_datBankChargeDate)
                               
    m_rsEmployeeOvertimeTempBanks.Filter = MDAC25FilterFix(l_strFilter)
    If m_rsEmployeeOvertimeTempBanks.EOF Then
        CreateNewOvertimeTempBank v_lngEmployee, l_intOTBankCode, v_datBankChargeDate, l_lngOTOppNum
    End If
    
    l_decTotalHours = m_rsEmployeeOvertimeTempBanks!TOT_OT_HRS
    l_decRefusedHours = m_rsEmployeeOvertimeTempBanks!TOT_OT_HRS_REF
    l_decTotalHours = l_decTotalHours + v_decAccHours + v_decRefHours
    l_decRefusedHours = l_decRefusedHours + v_decRefHours

    
    m_rsEmployeeOvertimeTempBanks!TOT_OT_HRS = l_decTotalHours
    m_rsEmployeeOvertimeTempBanks!TOT_OT_HRS_REF = l_decRefusedHours
    m_rsEmployeeOvertimeTempBanks!LAST_MOD_EMP_NUM = m_lngLastModEmpNum
    m_rsEmployeeOvertimeTempBanks!CHARGE_DATE = v_rsRequest!CHARGE_DATE
    m_rsEmployeeOvertimeTempBanks!DATE_WORKED = v_rsRequest!DATE_WORKED
    m_rsEmployeeOvertimeTempBanks!START_TIME = v_rsRequest!START_TIME
    m_rsEmployeeOvertimeTempBanks!END_TIME = v_rsRequest!END_TIME
    m_rsEmployeeOvertimeTempBanks.Update
    
    m_rsEmployeeOvertimeTempBanks.Filter = ""
    
    Exit Sub

eh:
    TrsRaiseError m_strMODULENAME, "UpdateTemporaryOTBank"
                        

End Sub


Private Function IsAcceptedResponse( _
                    ByVal v_strOTResponse As String _
                    ) As Boolean
                    
On Error GoTo eh
                                            
     m_rsOTResponses.MoveFirst
     m_rsOTResponses.Find "OT_RESP_CD = '" & v_strOTResponse & "'"

     If m_rsOTResponses.EOF Then
     ' Raise Error Should have found it
     ' This code will be put in on phase 3
        IsAcceptedResponse = False
     Else
        If m_rsOTResponses!OT_WORKED_IND = "Y" Then
            IsAcceptedResponse = True
        Else
            IsAcceptedResponse = False
        End If
     End If
     
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "IsAcceptedResponse"
                    
End Function

Private Function CalculateChargedHours( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_decHours As Variant, _
                    ByVal v_strOTResponse As String, _
                    ByVal v_strExceptionCode As String, _
                    ByVal v_datTimeChargeDate As Date, _
                    Optional v_blnRefusal As Boolean = False _
                    ) As Variant

    Dim l_decChargedHours As Variant
    Dim l_decFactor As Variant
    
On Error GoTo eh

    If IsResponseChargeable(v_strOTResponse) Then
        If IsHoursChargedAsPay(v_lngEmpNum, v_datTimeChargeDate) _
                Or v_blnRefusal Then
            l_decFactor = GetFactor(v_strExceptionCode)
            l_decChargedHours = v_decHours * l_decFactor
        Else
            l_decChargedHours = v_decHours
        End If
    Else
        l_decChargedHours = CDec(0)
    End If
    
    CalculateChargedHours = l_decChargedHours
    
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "CalculateChargedHours"
                       
End Function


Private Function GetFactor( _
                    ByVal v_strExceptionCode As String _
                    ) As Variant

    Dim l_obj As dbCodeTables.cdCodeTables
    
    Dim l_varFactor As Variant
    
 On Error GoTo GetFactor_EH

    m_rsOTExceptionCodes.Filter = ""
    If Not RecordsExist(m_rsOTExceptionCodes) Then
    'Should not happen. Always will have overtime exception codes
        GetFactor = CDec(1)
        m_rsOTExceptionCodes.Filter = ""
        Exit Function
    End If
    m_rsOTExceptionCodes.Filter = "EXCEPTION_CD = '" & v_strExceptionCode & "'"
    If Not RecordsExist(m_rsOTExceptionCodes) Then
    'Should not happen. Always will have overtime exception codes
        GetFactor = CDec(1)
        m_rsOTExceptionCodes.Filter = ""
        Exit Function
    End If
    GetFactor = CDec(m_rsOTExceptionCodes!FACTOR)
    m_rsOTExceptionCodes.Filter = ""
    
    Exit Function
    
GetFactor_EH:

    GetObjectContext.SetAbort
    Set l_obj = Nothing
    TrsRaiseError m_strMODULENAME, "GetFactor"
            
End Function


Private Function IsResponseChargeable( _
                    ByVal v_strOTResponse As String _
                    ) As Boolean
                    
    Dim l_strChargeInd As String
    
    On Error GoTo eh
    
     If Not RecordsExist(m_rsOTResponses) Then
     ' This should not happen
        IsResponseChargeable = False
        Exit Function
     End If

     m_rsOTResponses.MoveFirst
     m_rsOTResponses.Find "OT_RESP_CD = '" & v_strOTResponse & "'"

     If m_rsOTResponses.EOF Then
     ' Raise Error Should have found it
     ' This code will be put in on phase 3
        IsResponseChargeable = False
     Else
'This field is "C" (Charge) or "N" (do not charge)
        If m_rsOTResponses!OT_CHRG_IND = "C" Then
            IsResponseChargeable = True
        Else
            IsResponseChargeable = False
        End If
     End If
                    
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "IsResponseChargeable"

End Function

Private Sub AddToAuditTrail( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_Hours As Variant, _
                ByVal v_OTBankCode As Integer, _
                ByVal v_AccRefInd As String, _
                ByVal v_OTOppNum As Long, _
                ByVal v_Reason As String, _
                ByVal v_rsRequest As ADODB.Recordset, _
                ByVal v_decRate As Variant)
                    
    On Error GoTo eh
                
    m_rsAuditTrail.AddNew
    
    m_rsAuditTrail!EMP_NUM = v_lngEmpNum
    m_rsAuditTrail!HRS_ADJ = v_Hours
    m_rsAuditTrail!OT_BANK_CD = v_OTBankCode
    m_rsAuditTrail!ACC_REF_IND = v_AccRefInd
    m_rsAuditTrail!OT_OPP_NUM = v_OTOppNum
    m_rsAuditTrail!REASON = v_Reason
    m_rsAuditTrail!LAST_MOD_EMP_NUM = m_lngLastModEmpNum
    m_rsAuditTrail!CHARGE_DATE = v_rsRequest!CHARGE_DATE
    m_rsAuditTrail!DATE_WORKED = v_rsRequest!DATE_WORKED
    m_rsAuditTrail!START_TIME = v_rsRequest!START_TIME
    m_rsAuditTrail!END_TIME = v_rsRequest!END_TIME
    m_rsAuditTrail!SHIFT_CD = v_rsRequest!SHIFT_CD
    m_rsAuditTrail!HOURS = v_rsRequest!HOURS
    m_rsAuditTrail!Rate = v_decRate
    
    m_rsAuditTrail.Update
    
    Exit Sub

eh:
    TrsRaiseError m_strMODULENAME, "AddToAuditTrail"
    
End Sub

Private Function IsHoursChargedAsPay( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datTimeChargeDate As Date _
                    ) As Boolean
                    
    Dim l_strUnion As String
    
    On Error GoTo eh
    
    m_rsEmployeeOTData.Filter = ""
    If m_rsEmployeeOTData.EOF And m_rsEmployeeOTData.BOF Then
    ' Did not get employee OT Data
    ' This Condition should not happen.
        IsHoursChargedAsPay = False
        Exit Function
    End If
    
    m_rsEmployeeOTData.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmpNum) & _
                                " and CHARGE_DATE = " & str(v_datTimeChargeDate))
    If m_rsEmployeeOTData.EOF And m_rsEmployeeOTData.BOF Then
    ' Did not get employee OT Data for Charge Date
    ' This Condition should not happen.
        IsHoursChargedAsPay = False
        m_rsEmployeeOTData.Filter = ""
        Exit Function
    End If
    
    l_strUnion = m_rsEmployeeOTData!UNION_CD
    If GetChargeAsPayInd(l_strUnion) = "Y" Then
        IsHoursChargedAsPay = True
    Else
        IsHoursChargedAsPay = False
    End If
    
    m_rsEmployeeOTData.Filter = ""
    
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "IsHoursChargedAsPay"
    
                    
                    
End Function


Private Function IsUnionIAOvertime( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datDate As Date _
                    ) As Boolean
                    
    Dim l_strUnionCode As String
    
    On Error GoTo eh
    
    l_strUnionCode = GetEmployeeUnionCode(v_lngEmpNum, v_datDate)
    
    If l_strUnionCode = "IA" Then
        IsUnionIAOvertime = True
    Else
        IsUnionIAOvertime = False
    End If
    
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "IsUnionIAOvertime"
                    
End Function

Private Sub CreateNewOvertimeBank( _
                ByVal v_lngEmployee As Long, _
                ByVal v_intOTBankCode As Integer, _
                Optional ByVal v_decTotalHours As Variant = 0, _
                Optional ByVal v_decRefusedHours As Variant = 0)
    
    On Error GoTo eh
                
    m_rsEmployeeOvertimeBanks.AddNew
    m_rsEmployeeOvertimeBanks!EMP_NUM = v_lngEmployee
    m_rsEmployeeOvertimeBanks!OT_BANK_CD = v_intOTBankCode
    m_rsEmployeeOvertimeBanks!TOT_OT_HRS = CDec(v_decTotalHours)
    m_rsEmployeeOvertimeBanks!TOT_OT_HRS_REF = CDec(v_decRefusedHours)
    m_rsEmployeeOvertimeBanks!LAST_MOD_EMP_NUM = m_lngLastModEmpNum
    m_rsEmployeeOvertimeBanks.Update
    
    Exit Sub

eh:
    TrsRaiseError m_strMODULENAME, "CreateNewOvertimeBank"
                

End Sub


Private Sub CreateNewOvertimeTempBank( _
                ByVal v_lngEmployee As Long, _
                ByVal v_intOTBankCode As Integer, _
                ByVal v_datDate As Date, _
                Optional ByVal v_lngOTOppNum As Long = 0, _
                Optional ByVal v_decTotalHours As Variant = 0, _
                Optional ByVal v_decRefusedHours As Variant = 0 _
                )
                
    On Error GoTo eh
                
    m_rsEmployeeOvertimeTempBanks.AddNew
    m_rsEmployeeOvertimeTempBanks!EMP_NUM = v_lngEmployee
    m_rsEmployeeOvertimeTempBanks!OT_OPP_NUM = v_lngOTOppNum
    m_rsEmployeeOvertimeTempBanks!OT_BANK_CD = v_intOTBankCode
    m_rsEmployeeOvertimeTempBanks!EFF_DATE = v_datDate
    m_rsEmployeeOvertimeTempBanks!TOT_OT_HRS = CDec(v_decTotalHours)
    m_rsEmployeeOvertimeTempBanks!TOT_OT_HRS_REF = CDec(v_decRefusedHours)
    m_rsEmployeeOvertimeTempBanks!LAST_MOD_EMP_NUM = m_lngLastModEmpNum
    m_rsEmployeeOvertimeTempBanks.Update
    
    Exit Sub

eh:
    TrsRaiseError m_strMODULENAME, "CreateNewOvertimeTempBank"
                

End Sub




Private Sub SaveUpdateToEmployeeOTBanks()

    Dim l_dbOTEqualizer As dbOvertimeEqualizer.cdOvertimeEqualizer
    Dim l_bolUpdateSuccesful As Boolean
    
        
On Error GoTo SaveUpdateToEmployeeOTBanks_EH

    Set l_dbOTEqualizer = CreateObject("dbOvertimeEqualizer.cdOvertimeEqualizer")
    l_bolUpdateSuccesful = l_dbOTEqualizer.UpdateEmployeeOvertimeBanks(m_rsEmployeeOvertimeBanks, m_rsAuditTrail)
    l_bolUpdateSuccesful = l_dbOTEqualizer.UpdateEmployeeOvertimeTempBanks(m_rsEmployeeOvertimeTempBanks)
    If m_rsOTBankTMCharges Is Nothing Then
    Else
        m_rsOTBankTMCharges.Filter = ""
        l_bolUpdateSuccesful = l_dbOTEqualizer.SaveOvertimeBankTimeCharges(m_rsOTBankTMCharges)
    End If
    
    Set l_dbOTEqualizer = Nothing

    Exit Sub

SaveUpdateToEmployeeOTBanks_EH:
    Set l_dbOTEqualizer = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strMODULENAME, "SaveUpdateToEmployeeOTBanks"
                                 
End Sub


Private Function GetCollectionPeriods() As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    Dim l_col As Collection
    Dim l_obj As dbCollectionPeriod.cdCollectionPeriod
    
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    
 On Error GoTo GetCollectionPeriods_EH

    Set l_col = New Collection
    l_col.Add "A"
    l_col.Add "B"
    l_datStartDate = DateAdd("d", -30, Date)
    l_datEndDate = DateAdd("d", 30, Date)
    Set l_obj = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    Set GetCollectionPeriods = l_obj.GetCollectionPeriods(l_datStartDate, l_datEndDate, l_col)
    Set l_obj = Nothing
    
    Set l_rs = Nothing
    Set l_col = Nothing
    Exit Function
    
GetCollectionPeriods_EH:

    GetObjectContext.SetAbort
    Set l_obj = Nothing
    
    Set l_rs = Nothing
    Set l_col = Nothing

    TrsRaiseError m_strMODULENAME, "GetCollectionPeriods"

End Function


Private Function DeterminePayPeriodEndDate( _
            ByVal v_lngEmployee As Long, _
            ByVal v_datDate As Date _
            ) As Date
            
    Dim l_strPayrollGroupCode As String
    Dim l_strFilter As String
    
    On Error GoTo DeterminePayPeriodEndDate_EH
    
    m_rsEmployeeOTData.Filter = ""
    If m_rsEmployeeOTData.EOF And m_rsEmployeeOTData.BOF Then
        'should not happen
        DeterminePayPeriodEndDate = Date
        Exit Function
    End If
    

'Getting the Payroll Group Code for Given Date
    m_rsEmployeeOTData.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmployee) & _
                                " and CHARGE_DATE = " & str(v_datDate))
    If m_rsEmployeeOTData.EOF And m_rsEmployeeOTData.BOF Then
        'should not happen
        DeterminePayPeriodEndDate = Date
        m_rsEmployeeOTData.Filter = ""
        Exit Function
    End If
    m_rsEmployeeOTData.Filter = ""
    
    l_strPayrollGroupCode = UCase(Trim(m_rsEmployeeOTData!PAYROLL_GROUP_CD))

    l_strFilter = "PAYROLL_GROUP_CD = '" & l_strPayrollGroupCode & "'"
    l_strFilter = l_strFilter & " and PAY_PER_START_DATE <= " & str(v_datDate)
    l_strFilter = l_strFilter & " and PAY_PER_END_DATE >= " & str(v_datDate)
    If Not RecordsExist(m_rsCollectionPeriods) Then
'use the Spam Cache
        DeterminePayPeriodEndDate = GetPayPeriodEndDate(l_strPayrollGroupCode, v_datDate)
        Exit Function
    End If
    
    m_rsCollectionPeriods.Filter = l_strFilter
    If Not RecordsExist(m_rsCollectionPeriods) Then
        'should not happen
        DeterminePayPeriodEndDate = Date
        m_rsCollectionPeriods.Filter = ""
        Exit Function
    End If

    DeterminePayPeriodEndDate = m_rsCollectionPeriods!PAY_PER_END_DATE
    m_rsCollectionPeriods.Filter = ""
    
    Exit Function


DeterminePayPeriodEndDate_EH:

    TrsRaiseError m_strMODULENAME, "DeterminePayPeriodEndDate"

End Function


Private Function GetPayPeriodEndDate( _
            ByVal v_strPayrollGroup As String, _
            ByVal v_datDate As Date _
            ) As Date

    Dim l_obj As dbCollectionPeriod.cdCollectionPeriod
    
    Dim l_varPayPerEndDate As Variant
    
 On Error GoTo GetPayPeriodEndDate_EH

    
    Set l_obj = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    l_varPayPerEndDate = l_obj.PayPeriodEndDate(v_strPayrollGroup, v_datDate)
    If IsDate(l_varPayPerEndDate) Then
        GetPayPeriodEndDate = l_varPayPerEndDate
    Else
        GetPayPeriodEndDate = Date
    End If
    
    Set l_obj = Nothing
    
    Exit Function
    
GetPayPeriodEndDate_EH:

    GetObjectContext.SetAbort
    Set l_obj = Nothing
    TrsRaiseError m_strMODULENAME, "GetPayPeriodEndDate"
    
End Function


Private Function GetChargeInd( _
                    ByVal v_strResponse As String _
                    ) As String

    Dim l_obj As dbCodeTables.cdCodeTables
    
    Dim l_varFactor As Variant
    
 On Error GoTo GetChargeInd_EH

    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    GetChargeInd = l_obj.OvertimeResponseChargeIndicator(v_strResponse)
    Set l_obj = Nothing

    Exit Function
    
GetChargeInd_EH:

    GetObjectContext.SetAbort
    Set l_obj = Nothing
    TrsRaiseError m_strMODULENAME, "GetChargeInd"
            
End Function

Private Function GetWorkInd( _
                    ByVal v_strResponse As String _
                    ) As String

    Dim l_obj As dbCodeTables.cdCodeTables
    
    Dim l_varFactor As Variant
    
 On Error GoTo GetWorkInd_EH

    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    GetWorkInd = l_obj.OvertimeResponseWorkIndicator(v_strResponse)
    
    Set l_obj = Nothing
    
    Exit Function
    
GetWorkInd_EH:

    GetObjectContext.SetAbort
    Set l_obj = Nothing
    TrsRaiseError m_strMODULENAME, "GetWorkInd"
            
End Function


Public Function GetBankChargeDate( _
                    ByVal v_rsEmployeeOTData As ADODB.Recordset _
                    ) As Date
                    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_bankChargeDate As Date
                    
    On Error GoTo GetBankChargeDate_EH


    If v_rsEmployeeOTData Is Nothing Then
        GetBankChargeDate = CDate("0100-01-01")
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    
    If v_rsEmployeeOTData.EOF And v_rsEmployeeOTData.BOF Then
        GetBankChargeDate = CDate("0100-01-01")
        GetObjectContext.SetComplete
        Exit Function
    End If

    v_rsEmployeeOTData.MoveFirst
    l_lngEmpNum = v_rsEmployeeOTData!EMP_NUM
    l_datChargeDate = v_rsEmployeeOTData!CHARGE_DATE
    
    Set m_rsEmployeeOTData = v_rsEmployeeOTData
    
    l_bankChargeDate = DetermineBankChargeDate(l_lngEmpNum, l_datChargeDate, 1)
    
    GetBankChargeDate = l_bankChargeDate
    Set m_rsEmployeeOTData = Nothing
    GetObjectContext.SetComplete
    Exit Function
    

GetBankChargeDate_EH:
    Set m_rsEmployeeOTData = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strMODULENAME, "GetBankChargeDate", , , , , True
                    
End Function

Private Function GetBankChargeDateRule( _
                    ByVal v_lngEmployee As Long, _
                    ByVal v_datChargeDate As Date, _
                    Optional v_strExceptionCode As String = "" _
                    ) As Integer
                    
    'Based on union contracts and AmerenUE Management Employee Handbook
    'The bank charge date is always the current date unless the employee is a
    '    1. UnionIA employee
    '    2. AmerenUE Management Employee
    'There are other areas of the company doing there on thing
    'Below are the rule numbers
    '  0. Current System Date
    '  1. Charge Date of Time Entry  (AmerenUE Management)
    '  2. 2 Calendar Sundays from pay period end date (Union IA)
    '  3. Sunday Following End of Pay Period  (WA loc 63)
    '  4. Monday Following End of Pay Period
    '  5. Tuesday Following End of Pay Period
    '  6. Wednesday Following End of Pay Period (WA loc 72) and Missouri Call Centers
    '  7. 1st Day of Month follwing Charge Date
    '  8. Charge on Monday following charge date
    '  9. Charge on the Wednesday after the Monday following charge date
    '  10. 2nd Wednesday Following End of Pay Period - Missouri Call Center
    
                    
    Dim l_BankChargeDateRule As Integer
    Dim l_datTodaysDate As Date
    Dim l_dat30DaysPriorToToday As Date
    Dim l_strUnionCode As String
    Dim l_strDeptNum As String
    
    On Error GoTo eh
             
    l_datTodaysDate = Date
    l_dat30DaysPriorToToday = DateAdd("d", -30, l_datTodaysDate)
    
'    If IsExecutive(v_lngEmployee, v_datChargeDate) Then
'        GetBankChargeDateRule = 1
'        Exit Function
'    End If


'If charging for a day 30 days ago, use system date
' Should not happen, because pay period should be closed
    If v_datChargeDate < l_dat30DaysPriorToToday Then
        GetBankChargeDateRule = 0
        Exit Function
    End If
    
    l_strUnionCode = GetEmployeeUnionCode(v_lngEmployee, v_datChargeDate)
    Select Case l_strUnionCode
        Case "IP", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BJ", "BK", "BL", "BM"
            'AmerenIP Unions and management (IP)
            '  9. Charge on the Wednesday after the Monday following charge date
            GetBankChargeDateRule = 9
            
        Case "IA"
            Select Case GetEmployeeDeptNum(v_lngEmployee, v_datChargeDate)
                Case "156", "16S"
                    GetBankChargeDateRule = 3  'Sunday after PP End
                Case Else
                    GetBankChargeDateRule = 2   '2 Calendar Sundays from pay period end date
            End Select
        Case "WA"
            Select Case GetEmployeeDeptNum(v_lngEmployee, v_datChargeDate)
                Case "187", "189"
                    GetBankChargeDateRule = 6  'Wednesday after PP End
                Case Else
                    GetBankChargeDateRule = 3  'Sunday after PP End
            End Select
        
        Case "VA"
            Select Case GetEmployeeDeptNum(v_lngEmployee, v_datChargeDate)
                Case "330", "332", "336"
                    GetBankChargeDateRule = 7  '1st of month after date of OT
                Case "351"
                    GetBankChargeDateRule = 3  'Sunday after PP End
                Case Else
                    GetBankChargeDateRule = 0  'immediately (system date)
            End Select
        Case "JA", "JB"
             Select Case GetEmployeeDeptNum(v_lngEmployee, v_datChargeDate)
                Case "156", "16S" 'Callaway
                    GetBankChargeDateRule = 3  'Sunday after PP End
                Case "174" 'Callaway
                    GetBankChargeDateRule = 5  'Tuesday after PP End
                Case "154", "226", "326" 'Callaway
                    GetBankChargeDateRule = 8 'Charge on Monday following charge date
                Case "331", "335", "431", "432" 'Loc 22 Depts (Alton
                    'Had mixed results between first of month rule
                    ' and charge immediately
                    GetBankChargeDateRule = 0  'immediately (system date)
                Case Else
                    If IsMissouriCallCenter(GetEmployeeDeptNum(v_lngEmployee, v_datChargeDate)) Then
                        GetBankChargeDateRule = 6 'Wednesday after PP End
                    Else
                        GetBankChargeDateRule = 0  'immediately (system date)
                    End If
            End Select
        Case "TA", "TB"
            If IsMissouriCallCenter(GetEmployeeDeptNum(v_lngEmployee, v_datChargeDate)) Then
                GetBankChargeDateRule = 6 'Wednesday after PP End
            Else
                GetBankChargeDateRule = 0  'immediately (system date)
            End If
        Case Else
            GetBankChargeDateRule = 0
    End Select


    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "GetBankChargeDateRule"
    
End Function



Private Function GetEmployeeDeptNum( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDate As Date _
            ) As String
            
  On Error GoTo eh
   
    m_rsEmployeeOTData.Filter = ""
    If m_rsEmployeeOTData.EOF And m_rsEmployeeOTData.BOF Then
        'should not happen
        GetEmployeeDeptNum = ""
        Exit Function
    End If
    
    m_rsEmployeeOTData.Filter = "EMP_NUM = " & str(v_lngEmpNum) & _
                                " and CHARGE_DATE = " & str(v_datDate)
        
    If m_rsEmployeeOTData.EOF And m_rsEmployeeOTData.BOF Then
        'should not happen
        GetEmployeeDeptNum = ""
        m_rsEmployeeOTData.Filter = ""
        Exit Function
    End If
    
    GetEmployeeDeptNum = Trim(m_rsEmployeeOTData!DEPT_NUM)
    
    m_rsEmployeeOTData.Filter = ""
    
    Exit Function
    
eh:
    Set m_rsEmployeeOTData = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strMODULENAME, "GetEmployeeUnionCode"
    
            
End Function
            
    
Private Function GetEmployeeUnionCode( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datDate As Date _
                    ) As String
                    
On Error GoTo eh
                   
    m_rsEmployeeOTData.Filter = ""
    If m_rsEmployeeOTData.EOF And m_rsEmployeeOTData.BOF Then
        'should not happen
        GetEmployeeUnionCode = ""
        Exit Function
    End If
    
    m_rsEmployeeOTData.Filter = "EMP_NUM = " & str(v_lngEmpNum) & _
                                " and CHARGE_DATE = " & str(v_datDate)
        
    If m_rsEmployeeOTData.EOF And m_rsEmployeeOTData.BOF Then
        'should not happen
        GetEmployeeUnionCode = ""
        m_rsEmployeeOTData.Filter = ""
        Exit Function
    End If
    
    GetEmployeeUnionCode = Trim(m_rsEmployeeOTData!UNION_CD)
        
    m_rsEmployeeOTData.Filter = ""
    
    Exit Function
    
eh:
    Set m_rsEmployeeOTData = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strMODULENAME, "GetEmployeeUnionCode"
    
                    
End Function

Public Function HighmanHoursWorkGroupAssignments( _
                ByVal v_rsRequest As ADODB.Recordset, _
                ByVal v_lngLastModEmpNum As Long _
                    ) As Boolean
                    
    Dim l_aryEmployees As Variant
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_rsLineAssignments As ADODB.Recordset
    Dim l_rsWGAssignments As ADODB.Recordset
    Dim l_rsHighHoursRequest As ADODB.Recordset
    
    On Error GoTo eh
    
    MsgBox "No longer used", , "See Jim or Greg"
    HighmanHoursWorkGroupAssignments = True
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strMODULENAME, "HighmanHoursWorkGroupAssignments"
                    
End Function

Public Function HighmanHoursDemandLineAssignments( _
                ByVal v_rsRequest As ADODB.Recordset, _
                ByVal v_lngLastModEmpNum As Long _
                    ) As Boolean
                    
    Dim l_aryEmployees As Variant
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_rsLineAssignments As ADODB.Recordset
    Dim l_rsWGAssignments As ADODB.Recordset
    Dim l_rsHighHoursRequest As ADODB.Recordset
    

    On Error GoTo eh
    
    MsgBox "HighmanHoursDemandLineAssignments is no longer used", , "See Jim or Greg"
    HighmanHoursDemandLineAssignments = True
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strMODULENAME, "HighmanHoursDemandLineAssignments"
                    
End Function

Public Function ChargeOvertimeBanksForOpportunity( _
                ByVal v_rsOTOpportunity As ADODB.Recordset, _
                ByVal v_rsCanvassSheet As ADODB.Recordset, _
                ByVal v_rsEmployeeData As ADODB.Recordset, _
                ByVal v_rsSchedule As ADODB.Recordset, _
                ByVal v_lngLastModEmpNum As Long, _
                ByRef r_OppTimeCharges As ADODB.Recordset, _
                ByRef r_rsBrvs As ADODB.Recordset _
                    ) As Boolean
                         
    Dim l_rsForamattedRequest As ADODB.Recordset
    
    On Error GoTo eh
    
    If v_rsOTOpportunity!OPP_STATUS <> "P" Then
        m_blnFirstPosting = True
    Else
        If v_rsOTOpportunity("OPP_STATUS").OriginalValue <> "P" Then
            m_blnFirstPosting = True
        Else
            m_blnFirstPosting = False
        End If
    End If
    
    If m_blnFirstPosting Then
        v_rsCanvassSheet.Filter = "CANVASS_RESP_CD <> '' or FORCE_RESP_CD <> ''"
        If Not RecordsExist(v_rsCanvassSheet) Then
        'Write BRV stating that there are no responses to process
            ChargeOvertimeBanksForOpportunity = True
            GoTo XIT
        End If
    End If
    v_rsCanvassSheet.Sort = "ROW_NUM"
    
    m_lngLastModEmpNum = v_lngLastModEmpNum
    Set m_rsErrors = CreateNewBusinessErrorSet()
    
'Opportunity State was passed by the caller
    v_rsOTOpportunity.Filter = ""
    m_intOTGroup = v_rsOTOpportunity!OT_GROUP_NUM
    m_lngOTOppNum = v_rsOTOpportunity!OT_OPP_NUM
    m_intRefuseRuleNum = v_rsOTOpportunity!REFUSE_RULE_NUM
    m_strOTShiftCode = Trim(v_rsOTOpportunity!SHIFT_CD)
    m_datOTOppChargeDate = v_rsOTOpportunity!CHARGE_DATE
    m_datOTOppStartDate = v_rsOTOpportunity!DATE_OF_OT
    m_strOTOppStartTime = v_rsOTOpportunity!START_TIME
    m_strOTOppEndTime = v_rsOTOpportunity!END_TIME
    m_intOTOppLoc = v_rsOTOpportunity!LOC_CD
    Set m_rsSchedule = Nothing
    
        
    r_OppTimeCharges.Filter = ""
    If Not RecordsExist(r_OppTimeCharges) Then
        If m_blnFirstPosting Then
            GetObjectContext.SetComplete
            ChargeOvertimeBanksForOpportunity = True
            GoTo XIT
        Else
            r_OppTimeCharges.Filter = adFilterPendingRecords
            If Not RecordsExist(r_OppTimeCharges) Then
                GetObjectContext.SetComplete
                ChargeOvertimeBanksForOpportunity = True
                GoTo XIT
            End If
        End If
    End If
   
    If Not RecordsExist(v_rsEmployeeData) Then
        Set m_rsEmployeeData = Nothing
    Else
        Set m_rsEmployeeData = v_rsEmployeeData
    End If
    GetEmployeeState r_OppTimeCharges, v_rsCanvassSheet
    
    GetSupportInfo
    
    If Not DetermineRefusalCharges(v_rsCanvassSheet, r_OppTimeCharges) Then
        ChargeOvertimeBanksForOpportunity = False
        GoTo XIT
    End If
    If Not DetermineAcceptorCharges(v_rsCanvassSheet, r_OppTimeCharges, l_rsForamattedRequest) Then
        ChargeOvertimeBanksForOpportunity = False
        GoTo XIT
    End If

    If m_blnFirstPosting Then
        BuildRequestForPosting r_OppTimeCharges, v_rsCanvassSheet, l_rsForamattedRequest
    Else
        Set m_rsOppBankCharges = GetOppBankCharges()
        BuildRequestForUnPosting l_rsForamattedRequest
        BuildRequestForPosting r_OppTimeCharges, v_rsCanvassSheet, l_rsForamattedRequest
        RemoveBankTimeCharges
    End If

    If Not RecordsExist(l_rsForamattedRequest) Then
        GetObjectContext.SetComplete
        ChargeOvertimeBanksForOpportunity = True
        GoTo XIT
    End If
    
    ProcessOvertimeBankUpdateRequests l_rsForamattedRequest
    GetObjectContext.SetComplete
    ChargeOvertimeBanksForOpportunity = True
XIT:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsBrvs = Nothing
        Else
            Set r_rsBrvs = m_rsErrors
        End If
    Else
        Set r_rsBrvs = Nothing
    End If
    Set l_rsForamattedRequest = Nothing
    ClearGlobalRecordsets
    Exit Function
    
eh:
' The client is expecting a BRV recordset when errors occur
' Calling TrsRaiseError in this function will cause the
' BRVs to not be passed back, because the statement
' end function was not executed

    Set r_rsBrvs = Nothing
    GetObjectContext.SetAbort
    
    Set l_rsForamattedRequest = Nothing
    ClearGlobalRecordsets
    TrsRaiseError m_strMODULENAME, "ChargeOvertimeBanksForOpportunity", , , , , True

End Function

Private Sub ClearGlobalRecordsets()
    
    Set m_rsOTGroupBanks = Nothing
    Set m_rsSchedule = Nothing
    Set m_rsAuditTrail = Nothing
    Set m_rsOTResponses = Nothing
    Set m_rsOTBankTMCharges = Nothing
    Set m_rsOTExceptionCodes = Nothing
    Set m_rsEmployeeOvertimeBanks = Nothing
    Set m_rsEmployeeOvertimeTempBanks = Nothing
    Set m_rsCollectionPeriods = Nothing
    Set m_rsEmployeeOTData = Nothing
    Set m_rsEmployeeData = Nothing
    Set m_rsOppBankCharges = Nothing
    Set m_rsHolidayRestExcCds = Nothing
    Set m_rsErrors = Nothing

End Sub

Private Function IsExecutive( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datDate As Date _
                    ) As Boolean
                    
    Dim l_strExecConInd As String
                   
    On Error GoTo eh
    
    m_rsEmployeeOTData.Filter = ""
    If m_rsEmployeeOTData.EOF And m_rsEmployeeOTData.BOF Then
        'should not happen
        IsExecutive = False
        Exit Function
    End If
    
    m_rsEmployeeOTData.Filter = MDAC25FilterFix("EMP_NUM = " & str(v_lngEmpNum) & _
                                " and CHARGE_DATE = " & str(v_datDate))
        
    If m_rsEmployeeOTData.EOF And m_rsEmployeeOTData.BOF Then
        'should not happen
        IsExecutive = False
        m_rsEmployeeOTData.Filter = ""
        Exit Function
    End If
    
    l_strExecConInd = Trim(m_rsEmployeeOTData!EXEC_CONTRACT_IND)
    
    If l_strExecConInd = "E" Then
        IsExecutive = True
    Else
        IsExecutive = False
    End If
    
    m_rsEmployeeOTData.Filter = ""
    
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "IsExecutive"
                    
End Function

Private Function IsUnionOAOvertime( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datDate As Date _
                    ) As Boolean
                    
    Dim l_strUnionCode As String
    
    On Error GoTo eh
    
    l_strUnionCode = GetEmployeeUnionCode(v_lngEmpNum, v_datDate)
    
    If l_strUnionCode = "OA" Then
        IsUnionOAOvertime = True
    Else
        IsUnionOAOvertime = False
    End If
    
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "IsUnionOAOvertime"
                    
End Function


Private Sub AppendRedTimeNotChargedWarning( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_strErrorDescription As String)
    
    Dim l_bolError As Boolean
    
    On Error GoTo eh

    l_bolError = AppendBusinessError(m_rsErrors, _
        20000, _
        mstrCBUSINESSWAR, _
        v_strErrorDescription, _
        m_strMODULENAME, _
        0, , , "EMP_NUM", _
        str(v_lngEmpNum))
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "AppendRedTimeNotChargedWarning"

End Sub

Private Sub AppendRedTimeNotChargedError( _
            ByVal v_lngEmpNum As Long)
    
    Dim l_strErrDesc As String
    Dim l_bolError As Boolean
    
    On Error GoTo eh
    
    If m_intRefuseRuleNum = 0 Then
        l_strErrDesc = "The canvass had no refusal rule.  Please enter a refusal time charge " & _
                        "or take away the refusal."
    Else
        l_strErrDesc = "No employees accepted the opportunity. System could not automate " & _
                        "the refusal charge.  Please enter a refusal time charge or take away the refusal."
    End If

    l_bolError = AppendBusinessError(m_rsErrors, _
        40000, _
        mstrCBUSINESSERR, _
        l_strErrDesc, _
        m_strMODULENAME, _
        0, , , "EMP_NUM", _
        str(v_lngEmpNum))
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "AppendRedTimeNotChargedError"

End Sub

Private Function GetOTGroupBanks( _
    ByVal v_intOTGroup As Integer _
    ) As ADODB.Recordset
    
    Dim l_rs As ADODB.Recordset
 
On Error GoTo eh

    
    Dim l_dbOvertimeEqualizer As dbOvertimeEqualizer.cdOvertimeEqualizer
   
    Set l_dbOvertimeEqualizer = CreateObject("dbOvertimeEqualizer.cdOvertimeEqualizer")
    
    Set l_rs = l_dbOvertimeEqualizer.GetOTGroupCanvassBank(v_intOTGroup)
    
    If RecordsExist(l_rs) Then
        m_intOTGroupStandingBank = l_rs!OT_STAND_BANK_CD
    Else
        m_intOTGroupStandingBank = 0
    End If
    Set l_rs = Nothing
    
    Set GetOTGroupBanks = l_dbOvertimeEqualizer.GetOTGroupBanks(v_intOTGroup)
         
    Set l_dbOvertimeEqualizer = Nothing
    
    
    Exit Function

eh:
    Set l_rs = Nothing
    Set l_dbOvertimeEqualizer = Nothing
    TrsRaiseError m_strMODULENAME, "GetOTGroupBanks"
End Function

Private Sub GetState( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date)

Dim l_dbOvertimeEqualizer As dbOvertimeEqualizer.cdOvertimeEqualizer

On Error GoTo eh
   
    Set l_dbOvertimeEqualizer = CreateObject("dbOvertimeEqualizer.cdOvertimeEqualizer")

    Set m_rsEmployeeOvertimeBanks = l_dbOvertimeEqualizer.GetEmployeeOvertimeBanksForUpdate(v_aryEmployees)
    Set m_rsEmployeeOvertimeTempBanks = l_dbOvertimeEqualizer.GetEmployeeOvertimeTempBanks(v_aryEmployees)
    
    Set l_dbOvertimeEqualizer = Nothing
    Set m_rsAuditTrail = CreateAuditTrailRecordset()
    If m_rsEmployeeOTData Is Nothing Then
        Set m_rsEmployeeOTData = GetEmpOTData(v_aryEmployees, v_datStartDate, v_datEndDate)
    End If
    If m_rsOTResponses Is Nothing Then
        Set m_rsOTResponses = GetOvertimeResponses()
    End If
    If m_rsOTExceptionCodes Is Nothing Then
        Set m_rsOTExceptionCodes = GetOvertimeExceptions()
    End If
    Set m_rsCollectionPeriods = GetCollectionPeriods()

    Exit Sub
    
eh:
    Set l_dbOvertimeEqualizer = Nothing
    TrsRaiseError m_strMODULENAME, "GetState"
End Sub
Private Sub BuildRequestForPosting( _
                    ByRef r_rsTimeCharges As ADODB.Recordset, _
                    ByVal v_rsCanvassSheet As ADODB.Recordset, _
                    ByRef r_rsRequest As ADODB.Recordset)
                      
    Dim l_lngEmpNum As Long
    Dim l_strResponse As String
    Dim l_strCallTime As String
    Dim l_decAccHours As Variant
    Dim l_decRefHours As Variant
    Dim l_decTotAccHours As Variant
    Dim l_decTotRefHours As Variant
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_StartDateTime As Date
    Dim l_EndDateTime As Date
    Dim l_strExcCd As String
    Dim l_decTotHoursWorked As Variant
    Dim l_decTotHoursRefused As Variant
    

    On Error GoTo eh
    
    If r_rsRequest Is Nothing Then
        Set r_rsRequest = CreateRequestRecordset()
    End If
    v_rsCanvassSheet.Filter = ""
    r_rsTimeCharges.Filter = ""
    
    If Not RecordsExist(v_rsCanvassSheet) Or Not RecordsExist(r_rsTimeCharges) Then
        Exit Sub
    End If
    
    v_rsCanvassSheet.MoveFirst
    While Not v_rsCanvassSheet.EOF
        l_decTotHoursWorked = 0
        l_decTotHoursRefused = 0
        l_decTotAccHours = 0
        l_decTotRefHours = 0
        l_decAccHours = 0
        l_decRefHours = 0
        l_lngEmpNum = v_rsCanvassSheet!EMP_NUM
        l_strResponse = DetermineOppResponse(v_rsCanvassSheet)
        If l_strResponse <> "" Then
            l_strCallTime = DetermineOppResponseCallTime(v_rsCanvassSheet)
            r_rsTimeCharges.Filter = "EMP_NUM = " & str(l_lngEmpNum)
            If RecordsExist(r_rsTimeCharges) Then
                r_rsTimeCharges.MoveFirst
                While Not r_rsTimeCharges.EOF
                    DetermineHoursToCharge l_lngEmpNum, l_strResponse, r_rsTimeCharges, l_decAccHours, l_decRefHours
                    l_decTotAccHours = l_decTotAccHours + l_decAccHours
                    l_decTotRefHours = l_decTotRefHours + l_decRefHours
                    If l_decAccHours > 0 Then
                        l_decTotHoursWorked = l_decTotHoursWorked + r_rsTimeCharges!HOURS + r_rsTimeCharges!TRAVEL_PAY_HRS
                    Else
                        l_decTotHoursRefused = l_decTotHoursRefused + r_rsTimeCharges!HOURS + r_rsTimeCharges!TRAVEL_PAY_HRS
                    End If
                    r_rsTimeCharges.MoveNext
                Wend
                r_rsTimeCharges.MoveFirst
                If l_decTotAccHours > 0 Then
                    l_decTotAccHours = Round(l_decTotAccHours, 1)
                    DetermineContinuousStartEndTimes r_rsTimeCharges, l_StartDateTime, l_EndDateTime
                    l_strExcCd = DeterimineExceptionCode(r_rsTimeCharges)
                    InsertRequestForOpportunity r_rsRequest, l_lngEmpNum, l_decTotHoursWorked, l_decTotAccHours, l_strResponse, l_StartDateTime, l_EndDateTime, "A", l_strCallTime, l_strExcCd
                End If
                If l_decTotRefHours > 0 Then
                    l_decTotRefHours = Round(l_decTotRefHours, 1)
                    DetermineContinuousStartEndTimes r_rsTimeCharges, l_StartDateTime, l_EndDateTime, True
                    l_strExcCd = DeterimineExceptionCode(r_rsTimeCharges)
                    InsertRequestForOpportunity r_rsRequest, l_lngEmpNum, l_decTotHoursRefused, l_decTotRefHours, l_strResponse, l_StartDateTime, l_EndDateTime, "R", l_strCallTime, l_strExcCd
                End If
                If l_decTotAccHours = 0 And l_decTotRefHours = 0 Then
                    If IsAcceptedResponse(l_strResponse) Then
                        DetermineContinuousStartEndTimes r_rsTimeCharges, l_StartDateTime, l_EndDateTime
                        l_strExcCd = DeterimineExceptionCode(r_rsTimeCharges)
                        InsertRequestForOpportunity r_rsRequest, l_lngEmpNum, l_decTotHoursWorked, 0, l_strResponse, l_StartDateTime, l_EndDateTime, "A", l_strCallTime, l_strExcCd
                    Else
                        If IsRefusedResponse(l_strResponse) Then
                            DetermineContinuousStartEndTimes r_rsTimeCharges, l_StartDateTime, l_EndDateTime, True
                            l_strExcCd = DeterimineExceptionCode(r_rsTimeCharges)
                            InsertRequestForOpportunity r_rsRequest, l_lngEmpNum, l_decTotHoursRefused, 0, l_strResponse, l_StartDateTime, l_EndDateTime, "R", l_strCallTime, l_strExcCd
                        End If
                    End If
                End If
            End If
        End If
        v_rsCanvassSheet.MoveNext
    Wend
    r_rsTimeCharges.Filter = ""
    
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "BuildRequestForPosting"
                  
End Sub


Private Sub InsertRequestForOpportunity( _
                  ByRef r_rsRequest As ADODB.Recordset, _
                  ByVal v_lngEmpNum As Long, _
                  ByVal v_decHours As Variant, _
                  ByVal v_decHoursCharged As Variant, _
                  ByVal v_strResponse As String, _
                  ByVal v_StartDateTime As Date, _
                  ByVal v_EndDateTime As Date, _
                  ByVal v_AccRefInd As String, _
                  ByVal l_strCallTime As String, _
                  ByVal v_strExcCd As String)
                      
    Dim l_datStartDate As Date
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_decHours As Variant
    
                      
    On Error GoTo eh
    
    l_datStartDate = CDate(Format(v_StartDateTime, "YYYY-MM-DD"))
    l_strStartTime = Format(v_StartDateTime, "HH:mm")
    l_strEndTime = Format(v_EndDateTime, "HH:mm")
    l_decHours = Round(v_decHours, 1)
    
' Add Request
    r_rsRequest.AddNew
    r_rsRequest!REC_STATUS = "I"
    r_rsRequest!EMP_NUM = v_lngEmpNum
    r_rsRequest!CHARGE_DATE = m_datOTOppChargeDate
    r_rsRequest!EXCEPTION_CD = v_strExcCd
    r_rsRequest!HOURS = l_decHours
    r_rsRequest!HOURS_CHARGED = v_decHoursCharged
    r_rsRequest!OT_BANK_CODE = Null
    r_rsRequest!OT_RESPONSE = v_strResponse
    r_rsRequest!OT_OPP_NUM = m_lngOTOppNum
    r_rsRequest!TIME_ENTRY_IND = "N"
    r_rsRequest!START_TIME = l_strStartTime
    r_rsRequest!OVER_EXC_CD = ""
    r_rsRequest!DATE_WORKED = l_datStartDate
    r_rsRequest!END_TIME = l_strEndTime
    r_rsRequest!BANK_CHARGE_DATE = Null
    r_rsRequest!OT_MOD_DESC = "Posted Overtime Entry"
    r_rsRequest!CALL_TIME = l_strCallTime
    r_rsRequest!SHIFT_CD = m_strOTShiftCode
    r_rsRequest!ACC_REF_IND = v_AccRefInd

    r_rsRequest.Update
    
        
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "InsertRequestForOpportunity"

                      
End Sub



Private Sub UnPostRequestForOpportunity( _
                  ByRef r_rsRequest As ADODB.Recordset, _
                  ByVal v_rsOppBankCharge As ADODB.Recordset)
                  
    Dim l_decTotalHours As Variant
    
    On Error GoTo eh
    
    l_decTotalHours = CDec(v_rsOppBankCharge!HOURS_CHARGED)

' Add Request
    r_rsRequest.AddNew
    r_rsRequest!REC_STATUS = "D"
    r_rsRequest!EMP_NUM = v_rsOppBankCharge!EMP_NUM
    r_rsRequest!CHARGE_DATE = v_rsOppBankCharge!CHARGE_DATE
    r_rsRequest!EXCEPTION_CD = ""
    r_rsRequest!HOURS = v_rsOppBankCharge!HOURS
    r_rsRequest!HOURS_CHARGED = l_decTotalHours * -1
    r_rsRequest!OT_BANK_CODE = Null
    r_rsRequest!OT_RESPONSE = ""
    r_rsRequest!OT_OPP_NUM = v_rsOppBankCharge!OT_OPP_NUM
    r_rsRequest!TIME_ENTRY_IND = "N"
    r_rsRequest!START_TIME = v_rsOppBankCharge!START_TIME
    r_rsRequest!OVER_EXC_CD = ""
    r_rsRequest!DATE_WORKED = v_rsOppBankCharge!DATE_WORKED
    r_rsRequest!END_TIME = v_rsOppBankCharge!END_TIME
    r_rsRequest!BANK_CHARGE_DATE = Null
    r_rsRequest!OT_MOD_DESC = "Removed Bank Charge"
    r_rsRequest!CALL_TIME = ""
    r_rsRequest!ACC_REF_IND = v_rsOppBankCharge!ACC_REF_IND

    r_rsRequest.Update
    
        
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "UnPostRequestForOpportunity"

                      
End Sub

Private Sub UpdatePermanentOTBanks( _
                ByVal v_lngEmployee As Long, _
                ByVal v_rsRequest As ADODB.Recordset)
                
    Dim l_intOTBank As Integer
    Dim l_decAccHours As Variant
    Dim l_decRefHours As Variant
    Dim l_strAccRefInd As String
    Dim l_decRate As Variant
    Dim l_decHoursCharged As Variant
    Dim l_decHours As Variant
    Dim l_strExcCd As String
    Dim l_strResponse As String
    
    On Error GoTo eh
    
    If v_rsRequest!TIME_ENTRY_IND = "N" Then
        l_decHours = CDec(v_rsRequest!HOURS)
        l_decHoursCharged = CDec(v_rsRequest!HOURS_CHARGED)
        l_strAccRefInd = v_rsRequest!ACC_REF_IND
        l_strExcCd = ""
        If l_decHours = 0 Then
            l_decRate = CDec(0)
        Else
            l_decRate = l_decHoursCharged / l_decHours
            l_decRate = Round(l_decRate, 2)
        End If
        If l_strAccRefInd = "A" Then
            l_decAccHours = l_decHoursCharged
            l_decRefHours = 0
        Else
            l_decAccHours = 0
            l_decRefHours = l_decHoursCharged
        End If
    Else
        l_strResponse = v_rsRequest!OT_RESPONSE
        DetermineHoursToCharge v_lngEmployee, l_strResponse, v_rsRequest, l_decAccHours, l_decRefHours
        l_decHoursCharged = l_decAccHours + l_decRefHours
        If l_decAccHours <> 0 Then
            l_strAccRefInd = "A"
        Else
            l_strAccRefInd = "R"
        End If
        l_strExcCd = Trim(v_rsRequest!EXCEPTION_CD)
        l_decRate = GetFactor(l_strExcCd)
    End If
    
'JIMB not sure if this is in the correct place yet
'It is ok as long as the hours are not negative
    If v_rsRequest!REC_STATUS = "I" And _
        v_rsRequest!TIME_ENTRY_IND = "N" Then
        AddOTBankTimeCharge v_rsRequest, l_decRate, l_decHoursCharged, l_strAccRefInd
    End If
            
    If RecordsExist(m_rsOTGroupBanks) Then
        m_rsOTGroupBanks.MoveFirst
        While Not m_rsOTGroupBanks.EOF
            l_intOTBank = m_rsOTGroupBanks!OT_BANK_CD
            UpdatePermanentOTBank v_lngEmployee, l_decAccHours, l_decRefHours, l_intOTBank, v_rsRequest, l_decRate
            m_rsOTGroupBanks.MoveNext
        Wend
    Else
'use standard permanent overtime bank
        UpdatePermanentOTBank v_lngEmployee, l_decAccHours, l_decRefHours, 1, v_rsRequest, l_decRate
    End If
        
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "UpdatePermanentOTBanks"
End Sub

Private Sub UpdateTemporaryOTBanks( _
                ByVal v_lngEmployee As Long, _
                ByVal v_rsRequest As ADODB.Recordset, _
                ByVal v_datBankChargeDate As Date)
                
    Dim l_intOTBank As Integer
    Dim l_decAccHours As Variant
    Dim l_decRefHours As Variant
    Dim l_strResponse As String
    Dim l_decRate As Variant
    Dim l_decHoursCharged As Variant
    Dim l_decHours As Variant
    Dim l_strExcCd As String
    Dim l_strAccRefInd As String
    
    On Error GoTo eh
    
    If v_rsRequest!TIME_ENTRY_IND = "N" Then
        l_decHours = CDec(v_rsRequest!HOURS)
        l_decHoursCharged = CDec(v_rsRequest!HOURS_CHARGED)
        l_strAccRefInd = v_rsRequest!ACC_REF_IND
        l_strExcCd = ""
        If l_decHours = 0 Then
            l_decRate = CDec(0)
        Else
            l_decRate = l_decHoursCharged / l_decHours
            l_decRate = Round(l_decRate, 2)
        End If
        If l_strAccRefInd = "A" Then
            l_decAccHours = l_decHoursCharged
            l_decRefHours = 0
        Else
            l_decAccHours = 0
            l_decRefHours = l_decHoursCharged
        End If
    Else
        l_strResponse = v_rsRequest!OT_RESPONSE
        DetermineHoursToCharge v_lngEmployee, l_strResponse, v_rsRequest, l_decAccHours, l_decRefHours
        l_decHoursCharged = l_decAccHours + l_decRefHours
        If l_decAccHours <> 0 Then
            l_strAccRefInd = "A"
        Else
            l_strAccRefInd = "R"
        End If
        l_strExcCd = Trim(v_rsRequest!EXCEPTION_CD)
        l_decRate = GetFactor(l_strExcCd)
    End If
    
    
    If v_rsRequest!REC_STATUS = "I" And _
        v_rsRequest!TIME_ENTRY_IND = "N" Then
        AddOTBankTimeCharge v_rsRequest, l_decRate, l_decHoursCharged, l_strAccRefInd
    End If
    
    If RecordsExist(m_rsOTGroupBanks) Then
        m_rsOTGroupBanks.MoveFirst
        While Not m_rsOTGroupBanks.EOF
            l_intOTBank = m_rsOTGroupBanks!OT_BANK_CD
            UpdateTemporaryOTBank v_lngEmployee, l_decAccHours, l_decRefHours, _
                l_intOTBank, v_rsRequest, v_datBankChargeDate
            m_rsOTGroupBanks.MoveNext
        Wend
    Else
'use standard permanent overtime bank
        UpdateTemporaryOTBank v_lngEmployee, l_decAccHours, l_decRefHours, 1, v_rsRequest, v_datBankChargeDate
    End If
    
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "UpdateTemporaryOTBanks"
                
End Sub




Private Function DetermineOppResponse( _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    Optional v_blnOriginalResp As Boolean = False _
    ) As String
    
Dim l_strCanvassResp As String
Dim l_strForceResp As String

    On Error GoTo eh

    If v_blnOriginalResp Then
        l_strCanvassResp = Trim(v_rsCanvassSheet.Fields("CANVASS_RESP_CD").OriginalValue)
        l_strForceResp = Trim(v_rsCanvassSheet.Fields("FORCE_RESP_CD").OriginalValue)
    Else
        l_strCanvassResp = Trim(v_rsCanvassSheet!CANVASS_RESP_CD)
        l_strForceResp = Trim(v_rsCanvassSheet!FORCE_RESP_CD)
    End If
    If Trim(l_strForceResp) = "" Then
        l_strForceResp = l_strCanvassResp
    End If

    If l_strForceResp <> "" Then
        DetermineOppResponse = l_strForceResp
    Else
        DetermineOppResponse = l_strCanvassResp
    End If
        
    Exit Function
    
eh:
    TrsRaiseError m_strMODULENAME, "DetermineOppResponse"
End Function

Private Function DetermineOppResponseCallTime( _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    Optional v_blnOriginalResp As Boolean = False _
    ) As String
    
Dim l_strCanvassCallTime As String
Dim l_strForceCallTime  As String
Dim l_strForceResp As String
Dim l_datCallDate As Date
Dim l_strCallDateTime As String
Dim l_datOppDate As Date
Dim l_lngDiff As Long

    On Error GoTo eh

    If v_blnOriginalResp Then
        l_strCanvassCallTime = Trim(v_rsCanvassSheet.Fields("CANVASS_CALL_TIME").OriginalValue)
        l_strForceCallTime = Trim(v_rsCanvassSheet.Fields("FORCE_CALL_TIME").OriginalValue)
        l_strForceResp = Trim(v_rsCanvassSheet.Fields("FORCE_RESP_CD").OriginalValue)
    Else
        l_strCanvassCallTime = Trim(v_rsCanvassSheet!CANVASS_CALL_TIME)
        l_strForceCallTime = Trim(v_rsCanvassSheet!FORCE_CALL_TIME)
        l_strForceResp = Trim(v_rsCanvassSheet!FORCE_RESP_CD)
    End If
    If l_strForceResp = "" Then
        l_strForceCallTime = l_strCanvassCallTime
    End If

    If l_strForceResp <> "" Then
        l_strCallDateTime = l_strForceCallTime
    Else
        l_strCallDateTime = l_strCanvassCallTime
    End If
    
    l_datCallDate = CDate(Left(l_strCallDateTime, 10))
    l_datOppDate = DateAdd("d", 1, m_datOTOppChargeDate)
    If l_datCallDate > l_datOppDate Then
        l_strCallDateTime = ""
    Else
        l_lngDiff = DateDiff("d", l_datCallDate, l_datOppDate)
        If l_lngDiff > 15 Then
            l_strCallDateTime = ""
        End If
    End If
    
    DetermineOppResponseCallTime = l_strCallDateTime
        
    
    Exit Function
    
eh:
    TrsRaiseError m_strMODULENAME, "DetermineOppResponseCallTime"
End Function
Private Function BuildMissingRefusalCharges( _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    ByRef r_rsTimeCharges As ADODB.Recordset _
    ) As Boolean
    
'creates a time charge for all refusals that do not have a time charge entry

    Dim l_lngEmpNum As Long
    Dim l_strResponse As String
    Dim l_rsTimeClone As ADODB.Recordset
    Dim l_lngAcceptor As Long
    Dim l_blnOK As Boolean
    
    On Error GoTo eh
    
    l_blnOK = True
    v_rsCanvassSheet.MoveFirst
    Set l_rsTimeClone = r_rsTimeCharges.Clone
    While Not v_rsCanvassSheet.EOF
        l_lngEmpNum = v_rsCanvassSheet!EMP_NUM
        l_strResponse = DetermineOppResponse(v_rsCanvassSheet)
        If l_strResponse <> "" Then
            If IsRefusedResponse(l_strResponse) Then
                r_rsTimeCharges.Filter = "EMP_NUM = " & str(l_lngEmpNum)
                If Not RecordsExist(r_rsTimeCharges) Then
                    l_lngAcceptor = FindMatchingAcceptor(v_rsCanvassSheet, l_rsTimeClone)
                    If l_lngAcceptor = 0 Then
                    'Write BRV stating that replacing man and/or high man was not found
                        AppendRedTimeNotChargedError l_lngEmpNum
                        l_blnOK = False
                    Else
                        AddRefusalTimeCharge l_lngEmpNum, r_rsTimeCharges, l_lngAcceptor, l_rsTimeClone
                    End If
                End If
            End If
        End If
        v_rsCanvassSheet.MoveNext
    Wend
    
    BuildMissingRefusalCharges = l_blnOK
    Exit Function
    
eh:
    TrsRaiseError m_strMODULENAME, "BuildMissingRefusalCharges"
End Function

Private Function FindMatchingAcceptor( _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    ByVal v_rsTimeClone As ADODB.Recordset _
    ) As Long

'Returns false if a BRV is written

    Dim l_intCanvassRow As Integer
    Dim l_intForceRow As Integer
    Dim l_rsClone As ADODB.Recordset
    Dim l_strResponse As String
    Dim l_lngAcceptor As Long
    Dim l_decHours As Variant
    Dim l_decMaxHours As Variant
    Dim l_strExcCd As String
    Dim l_decChargeHours As Variant
    Dim l_lngEmpNum As Long

    On Error GoTo eh

    If m_intRefuseRuleNum = m_intNOREFUSERULE Then
        'return 0 because there was no matching acceptor
        'the calling method will write a BRV ??????
        FindMatchingAcceptor = 0
        Exit Function
    End If
    
    l_intCanvassRow = v_rsCanvassSheet!ROW_NUM
    l_intForceRow = v_rsCanvassSheet!FORCE_ROW_NUM
        
    Set l_rsClone = v_rsCanvassSheet.Clone
    l_rsClone.Filter = ""
    l_rsClone.Sort = "ROW_NUM"
    l_lngAcceptor = 0
    
    If m_intRefuseRuleNum = m_intREPLACINGMAN Then
        l_rsClone.AbsolutePosition = l_intCanvassRow
        While Not l_rsClone.EOF And l_lngAcceptor = 0
            l_strResponse = DetermineOppResponse(l_rsClone)
            If Not Trim(l_strResponse) = "" Then
                If IsAcceptedResponse(l_strResponse) Then
                    l_lngAcceptor = l_rsClone!EMP_NUM
                End If
            End If
            l_rsClone.MoveNext
        Wend
    End If
    
    If l_lngAcceptor > 0 Then
        FindMatchingAcceptor = l_lngAcceptor
        Exit Function
    End If
    
'HighMan rule
    l_decHours = CDec(0)
    l_decMaxHours = CDec(0)
    l_rsClone.MoveFirst
    While Not l_rsClone.EOF
        l_strResponse = DetermineOppResponse(l_rsClone)
        If Not Trim(l_strResponse) = "" Then
            If IsAcceptedResponse(l_strResponse) Then
                l_lngEmpNum = l_rsClone!EMP_NUM
                l_decHours = TotalChargeHours(l_strResponse, l_lngEmpNum, v_rsTimeClone)
                If l_decHours > l_decMaxHours Then
                    l_lngAcceptor = l_lngEmpNum
                    l_decMaxHours = l_decHours
                End If
            End If
        End If
        l_rsClone.MoveNext
    Wend
    
    FindMatchingAcceptor = l_lngAcceptor
    Exit Function

    
eh:
    TrsRaiseError m_strMODULENAME, "FindMatchingAcceptor"

End Function

Private Function TotalChargeHours( _
    ByVal v_strResponse As String, _
    ByVal v_lngEmpNum As Long, _
    ByVal v_rsTime As ADODB.Recordset _
    ) As Variant
    
    Dim l_decHours As Variant
    Dim l_decHoursCharged As Variant
    Dim l_decTotHoursCharged As Variant
    Dim l_strExceptionCode As String
    Dim l_datTimeChargeDate As Date
    
    
    On Error GoTo eh
    
    l_decHours = CDec(0)
    l_decHoursCharged = CDec(0)
    l_decTotHoursCharged = CDec(0)
    v_rsTime.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    v_rsTime.MoveFirst
    While Not v_rsTime.EOF
        l_strExceptionCode = v_rsTime!EXCEPTION_CD
        l_datTimeChargeDate = v_rsTime!CHARGE_DATE
        l_decHours = CDec(v_rsTime!HOURS) + CDec(v_rsTime!TRAVEL_PAY_HRS)
        l_decHoursCharged = CalculateChargedHours(v_lngEmpNum, l_decHours, v_strResponse, l_strExceptionCode, l_datTimeChargeDate)
        l_decTotHoursCharged = l_decTotHoursCharged + l_decHoursCharged
        v_rsTime.MoveNext
    Wend
    v_rsTime.Filter = ""
    
    TotalChargeHours = l_decTotHoursCharged
    
    Exit Function

    
eh:
    TrsRaiseError m_strMODULENAME, "TotalChargeHours"
    
End Function

Private Sub AddRefusalTimeCharge( _
    ByVal v_lngRefuser As Long, _
    ByRef r_rsTimeCharges As ADODB.Recordset, _
    ByVal v_lngAcceptor As Long, _
    ByVal v_rsTimeClone As ADODB.Recordset)
    
    Dim l_strExcCd As String
    Dim l_strRefExcCd As String
    
    On Error GoTo eh
    
    v_rsTimeClone.Filter = "EMP_NUM = " & str(v_lngAcceptor)
    If Not RecordsExist(v_rsTimeClone) Then
'should not happen
        Exit Sub
    End If
    
    v_rsTimeClone.MoveFirst
    While Not v_rsTimeClone.EOF
        r_rsTimeCharges.AddNew
        r_rsTimeCharges!OT_OPP_NUM = m_lngOTOppNum
        r_rsTimeCharges!EMP_NUM = v_lngRefuser
        r_rsTimeCharges!CHARGE_DATE = v_rsTimeClone!CHARGE_DATE
        r_rsTimeCharges!DATE_WORKED = v_rsTimeClone!DATE_WORKED
        r_rsTimeCharges!START_TIME = v_rsTimeClone!START_TIME
        r_rsTimeCharges!END_TIME = v_rsTimeClone!END_TIME
        r_rsTimeCharges!HOURS = CDec(v_rsTimeClone!HOURS)
        r_rsTimeCharges!TRAVEL_PAY_HRS = CDec(v_rsTimeClone!TRAVEL_PAY_HRS)
        l_strExcCd = Trim(v_rsTimeClone!EXCEPTION_CD)
        Select Case l_strExcCd
            Case "ON1"
                l_strRefExcCd = "RF1"
            Case "ON2"
                l_strRefExcCd = "RF5"
            Case "ON7"
                l_strRefExcCd = "RF3"
            Case Else
                Select Case Right(l_strExcCd, 1)
                    Case 1, 2, 3, 4, 5
                        l_strRefExcCd = "RF" & Right(l_strExcCd, 1)
                    Case Else
                        l_strRefExcCd = "REF"
                End Select
        End Select
        r_rsTimeCharges!EXCEPTION_CD = l_strRefExcCd
        r_rsTimeCharges!CREATED_EMP_NUM = m_lngSystemEmpNumber
        r_rsTimeCharges!LAST_MOD_EMP_NUM = m_lngSystemEmpNumber
        r_rsTimeCharges!WG_NUM = 0

        v_rsTimeClone.MoveNext
    Wend
    r_rsTimeCharges.Update
    
    Exit Sub

    
eh:
    TrsRaiseError m_strMODULENAME, "AddRefusalTimeCharge"
End Sub

Private Function GetEmployeeSchedules( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As String _
        ) As ADODB.Recordset
        
    Dim l_obj As busSchedule.cbsSchedule
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("busSchedule.cbsSchedule")
'    Set GetEmployeeSchedules = l_obj.GetAdjustedBaseSchedule( _
'        v_aryEmployees, v_datStartDate, v_datEndDate, Nothing, , _
'        trAdjBaseSchedWithExceptionsAndOvertime, trAllScheduleLevels, _
'        , trAllExceptions)

    Set GetEmployeeSchedules = l_obj.GetScheduleForOTCanvassing( _
        v_aryEmployees, v_datStartDate, v_datEndDate)
    Set l_obj = Nothing
    
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strMODULENAME, "GetEmployeeSchedules"

End Function

Private Function ApplyRedTimeRules( _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    ByRef r_rsTimeCharges As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strResp As String
    Dim l_lngEmpNum As Long
    Dim l_colRefusedEmps As Collection
    Dim l_colNotContactedEmps As Collection
    Dim l_intCount As Integer
    Dim l_blnOK As Boolean
    Dim l_strFilter As String
    
    On Error GoTo eh
    
'Builds a list of employees for which is refusal is to be charged
'Build another list of employees that were not physically contacted to
'determine if a refusal should be charged

    Set l_colRefusedEmps = New Collection
    Set l_colNotContactedEmps = New Collection
    v_rsCanvassSheet.MoveFirst
    While Not v_rsCanvassSheet.EOF
        l_strResp = DetermineOppResponse(v_rsCanvassSheet)
        If l_strResp <> "" Then
            If IsRefusedResponse(l_strResp) Then
                l_lngEmpNum = v_rsCanvassSheet!EMP_NUM
                l_colRefusedEmps.Add l_lngEmpNum
                If Not IsContactedResponse(l_strResp) Then
                    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
                        " AND LAST_MOD_EMP_NUM = " & str(m_lngSystemEmpNumber)
                    l_strFilter = MDAC25FilterFix(l_strFilter)
                    r_rsTimeCharges.Filter = l_strFilter
                    If RecordsExist(r_rsTimeCharges) Then
                        l_colNotContactedEmps.Add l_lngEmpNum
                    End If
                End If
            End If
        End If
        v_rsCanvassSheet.MoveNext
    Wend
    
    l_intCount = l_colNotContactedEmps.Count
    If l_intCount > 0 Then
        ApplyRedTimeNotContactedRules _
            l_colNotContactedEmps, v_rsCanvassSheet, r_rsTimeCharges
    End If
    
    ApplyRedTimeAlreadyChargedRule _
        l_colRefusedEmps, v_rsCanvassSheet, r_rsTimeCharges
        
    ApplyRedTimeNotAssignedRule _
        l_colRefusedEmps, v_rsCanvassSheet, r_rsTimeCharges
    
    ApplyRedTimeRules = True
    
    Exit Function
eh:
    
    TrsRaiseError m_strMODULENAME, "ApplyRedTimeRules"

End Function

Private Function GetEmployeeData( _
            ByVal v_aryEmployees As Variant _
            ) As ADODB.Recordset
            

On Error GoTo eh


    Dim l_obj As dbEmployee.cdEmployee

    
    Set l_obj = CreateObject("dbEmployee.cdEmployee")
    
    Set GetEmployeeData = l_obj.GetEmployee(v_aryEmployees)
         
    Set l_obj = Nothing
    
    
    Exit Function

eh:
    Set l_obj = Nothing
    GetObjectContext.SetAbort
    
    TrsRaiseError m_strMODULENAME, "GetEmployeeData"
                               
End Function

Private Function IsRefusedResponse( _
                    ByVal v_strOTResponse As String _
                    ) As Boolean

'A Refused response is a reponse for time that is not worked and the overtime bank is charged
'aka Redtime

    On Error GoTo eh

     m_rsOTResponses.MoveFirst
     m_rsOTResponses.Find "OT_RESP_CD = '" & v_strOTResponse & "'"

     If m_rsOTResponses.EOF Then
     ' Raise Error Should have found it
     ' This code will be put in on phase 3
        IsRefusedResponse = False
     Else
        If m_rsOTResponses!OT_WORKED_IND = "N" Then
            If m_rsOTResponses!OT_CHRG_IND = "C" Then
                IsRefusedResponse = True
            Else
                IsRefusedResponse = False
            End If
        Else
            IsRefusedResponse = False
        End If
     End If
     
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "IsRefusedResponse"
     
                    
End Function

Private Function IsContactedResponse( _
                    ByVal v_strOTResponse As String _
                    ) As Boolean

'A Refused response is a reponse for time that is not worked and the overtime bank is charged
'aka Redtime

    On Error GoTo eh

     m_rsOTResponses.MoveFirst
     m_rsOTResponses.Find "OT_RESP_CD = '" & v_strOTResponse & "'"

     If m_rsOTResponses.EOF Then
     ' Raise Error Should have found it
     ' This code will be put in on phase 3
        IsContactedResponse = False
     Else
        If m_rsOTResponses!OT_CONTACTED_IND <> "N" Then
            IsContactedResponse = True
        Else
            IsContactedResponse = False
        End If
     End If
     
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "IsContactedResponse"
     
                    
End Function

Private Sub ApplyRedTimeNotContactedRules( _
    ByVal v_colEmps As Collection, _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    ByRef r_rsTimeCharges As ADODB.Recordset)


    Dim l_lngEmpNum As Long
    Dim l_aryEmps As Variant
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_intCount As Integer
    Dim l_intI As Integer
    Dim l_strCallDateTime As String
    Dim l_datCallDate As Date
    Dim l_strCallTime As String
    Dim l_datCallDateTime As Date
    Dim l_datDateWorked As Date
    Dim l_datSchedStartDateTime As Date
    Dim l_datSchedEndDateTime As Date
    Dim l_strTime As String
    Dim l_strFilter As String
    Dim l_blnCalledDuringWork As Boolean
    Dim l_lngStartDiff As Long
    Dim l_lngEndDiff As Long
    Dim l_blnCalledWithin2HoursBefore As Boolean
    Dim l_blnCalledWithin2HoursAfter As Boolean
    Dim l_decHoursSched As Variant
    Dim l_datOppStartDateTime As Date
    Dim l_blnSchedToWorkPrior As Boolean
    Dim l_strErrorDesc As String
    Dim l_strExcCd As String
    Dim l_strOverExcCd As String
    
    On Error GoTo eh

    l_aryEmps = ArrayFromCollection(v_colEmps)
    
    l_datOppStartDateTime = DateTimeFrom(m_datOTOppStartDate, m_strOTOppStartTime)
    If m_rsSchedule Is Nothing Then
        v_rsCanvassSheet.Filter = ""
        v_rsCanvassSheet.MoveFirst
        l_datStartDate = DateAdd("d", -1, m_datOTOppChargeDate)
        While Not v_rsCanvassSheet.EOF
            l_lngEmpNum = v_rsCanvassSheet!EMP_NUM
            If IncludesString(Trim(str(l_lngEmpNum)), v_colEmps) Then
                l_strCallDateTime = DetermineOppResponseCallTime(v_rsCanvassSheet)
                If l_strCallDateTime <> "" Then
                    l_datCallDate = CDate(Left(l_strCallDateTime, 10))
                    If l_datCallDate < l_datStartDate Then
                        l_datStartDate = l_datCallDate
                    End If
                End If
            End If
            v_rsCanvassSheet.MoveNext
        Wend
        l_datEndDate = DateAdd("d", 1, m_datOTOppChargeDate)
        Set m_rsSchedule = GetEmployeeSchedules(l_aryEmps, l_datStartDate, l_datEndDate)
    End If
'Applying Conacted within 2hours coming/going to work rule
    l_intCount = v_colEmps.Count
    For l_intI = 1 To l_intCount
        l_lngEmpNum = v_colEmps(l_intI)
        v_rsCanvassSheet.Filter = "EMP_NUM = " & str(l_lngEmpNum)
        l_strCallDateTime = DetermineOppResponseCallTime(v_rsCanvassSheet)
        If l_strCallDateTime <> "" Then
            l_datCallDate = CDate(Left(l_strCallDateTime, 10))
            l_strCallTime = Mid(l_strCallDateTime, 12, 2) & ":" & Mid(l_strCallDateTime, 15, 2)
            l_datCallDateTime = DateTimeFrom(l_datCallDate, l_strCallTime)
            l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
                " AND SYSTEM_EXC_CD <> 'SLT' " & _
                " AND SYSTEM_EXC_CD <> 'SSO' "
            m_rsSchedule.Filter = MDAC25FilterFix(l_strFilter)
            m_rsSchedule.Sort = "CHARGE_DATE, DATE_WORKED, START_TIME"
            l_blnCalledDuringWork = False
            l_blnCalledWithin2HoursBefore = False
            l_blnCalledWithin2HoursAfter = False
            l_blnSchedToWorkPrior = False
          If RecordsExist(m_rsSchedule) Then
            m_rsSchedule.MoveFirst
            While Not m_rsSchedule.EOF And Not l_blnCalledDuringWork
                l_strExcCd = m_rsSchedule!EXCEPTION_CD
                l_strOverExcCd = Trim(m_rsSchedule!OVER_EXC_CD)
                If m_rsSchedule!HOURS > 0 And Not (IsHolidayRestException(l_strExcCd) Or l_strOverExcCd = "OA") Then
                    l_datDateWorked = m_rsSchedule!DATE_WORKED
                    l_strTime = m_rsSchedule!START_TIME
                    l_datSchedStartDateTime = DateTimeFrom(l_datDateWorked, l_strTime)
                    l_strTime = m_rsSchedule!END_TIME
                    l_datSchedEndDateTime = DateTimeFrom(l_datDateWorked, l_strTime)
                    If TimePeriodsIntersect(l_datCallDateTime, l_datCallDateTime, _
                        l_datSchedStartDateTime, l_datSchedEndDateTime) Then
                        l_blnCalledDuringWork = True
                        l_blnCalledWithin2HoursBefore = False
                        l_blnCalledWithin2HoursAfter = False
                    Else
                        l_lngStartDiff = Abs(DateDiff("n", l_datCallDateTime, l_datSchedStartDateTime))
                        l_lngEndDiff = Abs(DateDiff("n", l_datCallDateTime, l_datSchedEndDateTime))
                        If l_lngStartDiff < 120 Then
                            l_blnCalledWithin2HoursBefore = True
                        End If
                        If l_lngEndDiff < 120 Then
                            l_blnCalledWithin2HoursAfter = True
                        End If
                        If l_datCallDateTime < l_datSchedStartDateTime Then
                            If l_datOppStartDateTime >= l_datSchedStartDateTime Then
                                If m_rsSchedule!SHIFT_NUM > 0 Then
                                    l_blnSchedToWorkPrior = True
                                End If
                            End If
                        End If
                    End If
                End If
                m_rsSchedule.MoveNext
            Wend
          End If
            If l_blnSchedToWorkPrior Then
                'write BRV stating that the employee is scheduled to work prior to the opportunity
                'set the time charge to 0 hours
                    l_strErrorDesc = "No RedTime Charged.  The employee has a shift between the contact time and the opportunity"
                    AppendRedTimeNotChargedWarning l_lngEmpNum, l_strErrorDesc
                    SetTimeChargesToZero l_lngEmpNum, r_rsTimeCharges
            Else
                If l_blnCalledWithin2HoursBefore Then
                    'write BRV stating that the employee is scheduled to work within 2 hours
                    'set the time charge to 0 hours
                    l_strErrorDesc = "No RedTime Charged.  The employee is scheduled to work within 2 hours of the call(" & l_strCallDateTime & ")."
                    AppendRedTimeNotChargedWarning l_lngEmpNum, l_strErrorDesc
                    SetTimeChargesToZero l_lngEmpNum, r_rsTimeCharges
                End If
                If l_blnCalledWithin2HoursAfter Then
                    'write BRV stating that the employee has left work within 2 hours
                    'set the time charge to 0 hours
                    l_strErrorDesc = "No RedTime Charged.  The call time (" & l_strCallDateTime & ") is within 2 hours of the employee leaving work."
                    AppendRedTimeNotChargedWarning l_lngEmpNum, l_strErrorDesc
                    SetTimeChargesToZero l_lngEmpNum, r_rsTimeCharges
                End If
            End If
        End If
    Next l_intI
    
    v_rsCanvassSheet.Filter = ""
    
Exit Sub
eh:
   
    TrsRaiseError m_strMODULENAME, "ApplyRedTimeNotContactedRules"
                               
End Sub

Private Sub SetTimeChargesToZero( _
    ByVal v_lngEmpNum As Long, _
    ByRef r_rsTimeCharges As ADODB.Recordset)
    
    On Error GoTo eh
    
    r_rsTimeCharges.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    r_rsTimeCharges.MoveFirst
    While Not r_rsTimeCharges.EOF
        r_rsTimeCharges!EXCEPTION_CD = "RF0"
        r_rsTimeCharges.MoveNext
    Wend
    
    r_rsTimeCharges.Filter = ""
    
    
Exit Sub
eh:
   
    TrsRaiseError m_strMODULENAME, "SetTimeChargesToZero"
                                 
    
End Sub

Private Sub ApplyRedTimeAlreadyChargedRule( _
    ByVal v_colEmps As Collection, _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    ByRef r_rsTimeCharges As ADODB.Recordset)


    Dim l_lngEmpNum As Long
    Dim l_intCount As Integer
    Dim l_intI As Integer
    Dim l_strFilter As String
    Dim l_strErrorDesc As String
    Dim l_strTimeFilter As String
    Dim l_blnSystemCharge As Boolean
    Dim l_TC_DateWorked As Date
    Dim l_TC_StartTime As String
    Dim l_TC_EndTime As String
    Dim l_TC_StartDateTime As Date
    Dim l_TC_EndDateTime As Date
    Dim l_BC_DateWorked As Date
    Dim l_BC_StartTime As String
    Dim l_BC_EndTime As String
    Dim l_BC_StartDateTime As Date
    Dim l_BC_EndDateTime As Date
    Dim l_blnPrevCharge As Boolean
    
    On Error GoTo eh

    l_intCount = v_colEmps.Count
    For l_intI = 1 To l_intCount
        l_lngEmpNum = v_colEmps(l_intI)
        l_strTimeFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
            "  AND LAST_MOD_EMP_NUM <> " & str(m_lngSystemEmpNumber)
        l_strTimeFilter = MDAC25FilterFix(l_strTimeFilter)
        r_rsTimeCharges.Filter = l_strTimeFilter
        If RecordsExist(r_rsTimeCharges) Then
            l_blnSystemCharge = False
        Else
            l_blnSystemCharge = True
        End If
        l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
            "  AND CHARGE_DATE = " & m_datOTOppChargeDate & _
            "  AND HOURS_CHARGED <> 0 " & _
            "  AND SHIFT_CD = '" & m_strOTShiftCode & "'" & _
            "  AND OT_OPP_NUM <> " & str(m_lngOTOppNum)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        m_rsOTBankTMCharges.Filter = l_strFilter
        If RecordsExist(m_rsOTBankTMCharges) Then
            l_blnPrevCharge = True
            If m_intOTOppLoc = 36 Then
            'Probably need to do something different for Lakeside
            'Lakeside probably needs more refined to start and end time
                l_blnPrevCharge = False
                m_rsOTBankTMCharges.Sort = "DATE_WORKED, START_TIME"
                r_rsTimeCharges.Filter = "EMP_NUM = " & str(l_lngEmpNum)
                r_rsTimeCharges.MoveFirst
                m_rsOTBankTMCharges.MoveFirst
                While Not m_rsOTBankTMCharges.EOF
                    l_BC_DateWorked = m_rsOTBankTMCharges!DATE_WORKED
                    l_BC_StartTime = m_rsOTBankTMCharges!START_TIME
                    l_BC_EndTime = m_rsOTBankTMCharges!END_TIME
                    l_BC_StartDateTime = DateTimeFrom(l_BC_DateWorked, l_BC_StartTime)
                    l_BC_EndDateTime = DateTimeFrom(l_BC_DateWorked, l_BC_EndTime)
'                    r_rsTimeCharges.MoveFirst
                    While Not r_rsTimeCharges.EOF And Not l_blnPrevCharge
                        l_TC_DateWorked = r_rsTimeCharges!DATE_WORKED
                        l_TC_StartTime = r_rsTimeCharges!START_TIME
                        l_TC_EndTime = r_rsTimeCharges!END_TIME
                        l_TC_StartDateTime = DateTimeFrom(l_TC_DateWorked, l_TC_StartTime)
                        l_TC_EndDateTime = DateTimeFrom(l_TC_DateWorked, l_TC_EndTime)
                        If TimePeriodsIntersect(l_TC_StartDateTime, l_TC_EndDateTime, _
                                l_BC_StartDateTime, l_BC_EndDateTime) Then
                            l_blnPrevCharge = True
                        End If
                        r_rsTimeCharges.MoveNext
                    Wend
                    m_rsOTBankTMCharges.MoveNext
                Wend
            End If
            If l_blnPrevCharge Then
                If l_blnSystemCharge Then
                    'write BRV stating that the employee has already been charged for this shift
                    'set the time charge to 0 hours
                    l_strErrorDesc = "No RedTime Charged.  The employee has already been charged for this shift."
                    AppendRedTimeNotChargedWarning l_lngEmpNum, l_strErrorDesc
                    SetTimeChargesToZero l_lngEmpNum, r_rsTimeCharges
                Else
                    l_strErrorDesc = "User overrode system and the employee was charged Redtime.  The system determined that the employee has already been charged for this shift."
                    AppendRedTimeNotChargedWarning l_lngEmpNum, l_strErrorDesc
                End If
            End If
        End If
            

    Next l_intI
    
Exit Sub
eh:
   
    TrsRaiseError m_strMODULENAME, "ApplyRedTimeAlreadyChargedRule"
                               
End Sub


Private Function GetOvertimeWorkedInd( _
    ByVal v_strResp As String) As String
    
    Dim l_strOTWorkedInd As String
    
    On Error GoTo eh
    
                                                
     m_rsOTResponses.MoveFirst
     m_rsOTResponses.Find "OT_RESP_CD = '" & v_strResp & "'"

     If m_rsOTResponses.EOF Then
        l_strOTWorkedInd = "N"
     Else
        l_strOTWorkedInd = Trim(m_rsOTResponses!OT_WORKED_IND)
     End If
    
    GetOvertimeWorkedInd = l_strOTWorkedInd
    Exit Function
    
eh:
    TrsRaiseError m_strMODULENAME, "GetOvertimeWorkedInd"
                
End Function

Private Function GetOvertimeChargedInd( _
    ByVal v_strResp As String) As String
    
    Dim l_strOTChargedInd As String
    
    On Error GoTo eh
    
                                                
     m_rsOTResponses.MoveFirst
     m_rsOTResponses.Find "OT_RESP_CD = '" & v_strResp & "'"

     If m_rsOTResponses.EOF Then
        l_strOTChargedInd = "N"
     Else
        l_strOTChargedInd = Trim(m_rsOTResponses!OT_CHRG_IND)
     End If
    
    GetOvertimeChargedInd = l_strOTChargedInd
    Exit Function
    
eh:
    TrsRaiseError m_strMODULENAME, "GetOvertimeChargedInd"
                
End Function


Private Function DetermineRefusalCharges( _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    ByRef r_OppTimeCharges As ADODB.Recordset _
    ) As Boolean

    Dim l_lngEmp As Long
    Dim l_colEmp As Collection
    
    On Error GoTo eh
                                        
        If Not m_blnFirstPosting Then
            If AnyAcceptorChanges(r_OppTimeCharges) Then
                RemoveSystemRefusalTimeCharges r_OppTimeCharges, l_colEmp
            Else
                If RefusalCodeChanges(r_OppTimeCharges, v_rsCanvassSheet, l_colEmp) Then
                    RemoveSystemRefusalTimeCharges r_OppTimeCharges, l_colEmp
                End If
            End If
        End If

        If Not BuildMissingRefusalCharges(v_rsCanvassSheet, r_OppTimeCharges) Then
            DetermineRefusalCharges = False
            Exit Function
        End If
        
'Making an assumption that all employees on the canvass are
'from the same union
        v_rsCanvassSheet.MoveFirst
        m_rsEmployeeOTData.Filter = ""
        If RecordsExist(m_rsEmployeeOTData) Then
            m_rsEmployeeOTData.MoveFirst
            l_lngEmp = m_rsEmployeeOTData!EMP_NUM
        Else
            l_lngEmp = v_rsCanvassSheet!EMP_NUM
        End If
        If IsUnionOAOvertime(l_lngEmp, m_datOTOppChargeDate) Then
            If Not ApplyRedTimeRules(v_rsCanvassSheet, r_OppTimeCharges) Then
                DetermineRefusalCharges = False
                Exit Function
            End If
        End If

    DetermineRefusalCharges = True
    
    Exit Function
    
eh:
    TrsRaiseError m_strMODULENAME, "DetermineRefusalCharges"

End Function


Private Function GetOvertimeBankTimeCharges( _
    Optional v_lngOTOppNum As Long = 0 _
    ) As ADODB.Recordset

On Error GoTo eh


    Dim l_dbOvertimeEqualizer As dbOvertimeEqualizer.cdOvertimeEqualizer
    
    Set l_dbOvertimeEqualizer = CreateObject("dbOvertimeEqualizer.cdOvertimeEqualizer")
    
    Set GetOvertimeBankTimeCharges = l_dbOvertimeEqualizer.GetOvertimeBankTimeCharges(m_datOTOppChargeDate, m_strOTShiftCode, v_lngOTOppNum)
    
    Set l_dbOvertimeEqualizer = Nothing
    
    

    Exit Function
    
eh:
    TrsRaiseError m_strMODULENAME, "GetOvertimeBankTimeCharges"

End Function

Private Sub GetEmployeeState( _
    ByRef r_OppTimeCharges As ADODB.Recordset, _
    ByVal v_rsCanvassSheet As ADODB.Recordset)
    
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_aryEmps As Variant
    Dim l_colEmps As Collection
    Dim l_strResp As String
    Dim l_strOrigResp As String
    Dim l_strFResp As String
    Dim l_strFOrigResp As String
    Dim l_lngEmp As Long
    
    On Error GoTo eh
    
    l_datStartDate = m_datOTOppChargeDate
    l_datEndDate = m_datOTOppChargeDate
    Set l_colEmps = New Collection
    
    v_rsCanvassSheet.MoveFirst
    While Not v_rsCanvassSheet.EOF
        l_strResp = Trim(v_rsCanvassSheet.Fields("CANVASS_RESP_CD").Value)
        l_strOrigResp = Trim(v_rsCanvassSheet.Fields("CANVASS_RESP_CD").OriginalValue)
        l_strFResp = Trim(v_rsCanvassSheet.Fields("FORCE_RESP_CD").Value)
        l_strFOrigResp = Trim(v_rsCanvassSheet.Fields("FORCE_RESP_CD").OriginalValue)
        If l_strResp <> "" Or l_strOrigResp <> "" _
            Or l_strFResp <> "" Or l_strFOrigResp <> "" Then
            l_lngEmp = v_rsCanvassSheet!EMP_NUM
            l_colEmps.Add l_lngEmp
        End If
        v_rsCanvassSheet.MoveNext
    Wend

    
    If l_colEmps.Count = 0 Then
        Set m_rsEmployeeOTData = Nothing
        Exit Sub
    End If
    
    l_aryEmps = ArrayFromCollection(l_colEmps)
    
    
'This probably needs some work if we allow more than 1 charge date in the opportunity
    r_OppTimeCharges.Filter = ""
    If RecordsExist(r_OppTimeCharges) Then
        r_OppTimeCharges.Sort = "CHARGE_DATE"
        r_OppTimeCharges.MoveFirst
        While Not r_OppTimeCharges.EOF
            If r_OppTimeCharges!CHARGE_DATE < l_datStartDate Then
                l_datStartDate = r_OppTimeCharges!CHARGE_DATE
            End If
            If r_OppTimeCharges!CHARGE_DATE > l_datEndDate Then
                l_datEndDate = r_OppTimeCharges!CHARGE_DATE
            End If
            If IsEmpty(r_OppTimeCharges!CHARGE_DATE) Or IsNull(r_OppTimeCharges!CHARGE_DATE) Then
            Else
                If r_OppTimeCharges.Fields("CHARGE_DATE").OriginalValue < l_datStartDate Then
                    l_datStartDate = r_OppTimeCharges!CHARGE_DATE
                End If
                If r_OppTimeCharges.Fields("CHARGE_DATE").OriginalValue > l_datEndDate Then
                    l_datEndDate = r_OppTimeCharges!CHARGE_DATE
                End If
            End If
            r_OppTimeCharges.MoveNext
        Wend
    End If
    
    Set m_rsEmployeeOTData = GetEmpOTData(l_aryEmps, l_datStartDate, l_datEndDate)
    
    
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "GetEmployeeState"
End Sub

Private Sub AddOTBankTimeCharge( _
        ByVal v_rsRequest As ADODB.Recordset, _
        ByVal v_decRate As Variant, _
        ByVal v_decHoursCharged As Variant, _
        ByVal v_strAccRefInd As String)

'JIMB???  Where do I call this from
    On Error GoTo eh

    m_rsOTBankTMCharges.AddNew
    m_rsOTBankTMCharges!OT_OPP_NUM = v_rsRequest!OT_OPP_NUM
    m_rsOTBankTMCharges!EMP_NUM = v_rsRequest!EMP_NUM
    m_rsOTBankTMCharges!CHARGE_DATE = v_rsRequest!CHARGE_DATE
    m_rsOTBankTMCharges!DATE_WORKED = v_rsRequest!DATE_WORKED
    m_rsOTBankTMCharges!START_TIME = v_rsRequest!START_TIME
    m_rsOTBankTMCharges!END_TIME = v_rsRequest!END_TIME
    m_rsOTBankTMCharges!HOURS = v_rsRequest!HOURS
    m_rsOTBankTMCharges!Rate = v_decRate
    m_rsOTBankTMCharges!HOURS_CHARGED = v_decHoursCharged
    m_rsOTBankTMCharges!ACC_REF_IND = v_strAccRefInd
    m_rsOTBankTMCharges!SHIFT_CD = m_strOTShiftCode
    m_rsOTBankTMCharges!CREATED_EMP_NUM = m_lngLastModEmpNum
    m_rsOTBankTMCharges!LAST_MOD_EMP_NUM = m_lngLastModEmpNum
    
    
    m_rsOTBankTMCharges.Update
        
        
    
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "AddOTBankTimeCharge"

End Sub


Private Sub DetermineHoursToCharge( _
    ByVal v_lngEmployee As Long, _
    ByVal v_strResponse As String, _
    ByVal v_rsRequest As ADODB.Recordset, _
    ByRef r_decAccHours As Variant, _
    ByRef r_decRefHours As Variant)
    
    Dim l_strExceptionCode As String
    Dim l_strOTResponse As String
    Dim l_decHours As Variant
    Dim l_datTimeChargeDate As Date

    
    On Error GoTo eh
    
    l_datTimeChargeDate = v_rsRequest!CHARGE_DATE
    l_strExceptionCode = Trim(v_rsRequest!EXCEPTION_CD)
    l_decHours = CDec(v_rsRequest!HOURS) + CDec(v_rsRequest!TRAVEL_PAY_HRS)
    l_strOTResponse = v_strResponse

    If IsAcceptedResponse(l_strOTResponse) Then
        If Not IsRefusalExceptionCode(l_strExceptionCode) Then
            If l_strExceptionCode = "ONP" Then
                r_decAccHours = CDec(0)
                r_decRefHours = CDec(0)
            Else
                r_decAccHours = CalculateChargedHours(v_lngEmployee, l_decHours, l_strOTResponse, l_strExceptionCode, l_datTimeChargeDate)
                r_decRefHours = CDec(0)
            End If
        Else
                r_decRefHours = CalculateChargedHours(v_lngEmployee, l_decHours, l_strOTResponse, l_strExceptionCode, l_datTimeChargeDate, True)
                r_decAccHours = CDec(0)
        End If
    Else
        If IsRefusedResponse(l_strOTResponse) Then
            r_decRefHours = CalculateChargedHours(v_lngEmployee, l_decHours, l_strOTResponse, l_strExceptionCode, l_datTimeChargeDate, True)
            r_decAccHours = CDec(0)
        Else
            r_decAccHours = CDec(0)
            r_decRefHours = CDec(0)
        End If
    End If
    
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "DetermineHoursToCharge"
   
    
End Sub


Private Function IsRefusalExceptionCode( _
    ByVal v_strExcCd As String _
    ) As Boolean
    
   On Error GoTo eh
   
    m_rsOTExceptionCodes.Filter = "EXCEPTION_CD = '" & v_strExcCd & "'"
    If RecordsExist(m_rsOTExceptionCodes) Then
        If m_rsOTExceptionCodes!REFUSED_IND = "Y" Then
            IsRefusalExceptionCode = True
        Else
            IsRefusalExceptionCode = False
        End If
    Else
        IsRefusalExceptionCode = False
    End If
    
       
    Exit Function
    
eh:
    TrsRaiseError m_strMODULENAME, "IsRefusalExceptionCode"
   
    
End Function

Private Sub BuildRequestForUnPosting( _
                    ByRef r_rsRequest As ADODB.Recordset)
                      
    Dim l_lngEmpNum As Long
    Dim l_strResponse As String
    Dim l_strOrigResponse As String
    Dim l_strCallTime As String

    On Error GoTo eh
    
    If r_rsRequest Is Nothing Then
        Set r_rsRequest = CreateRequestRecordset()
    End If
    
    If Not RecordsExist(m_rsOppBankCharges) Then
        Exit Sub
    End If
    
    m_rsOppBankCharges.MoveFirst
    While Not m_rsOppBankCharges.EOF
        UnPostRequestForOpportunity r_rsRequest, m_rsOppBankCharges
        m_rsOppBankCharges.MoveNext
    Wend
    
    
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "BuildRequestForUnPosting"
                  
End Sub

Private Sub RemoveBankTimeCharges()

    On Error GoTo eh
        
    m_rsOTBankTMCharges.Filter = "OT_OPP_NUM = " & str(m_lngOTOppNum)
    If Not RecordsExist(m_rsOTBankTMCharges) Then
        Exit Sub
    End If
    
    m_rsOTBankTMCharges.MoveFirst
    While Not m_rsOTBankTMCharges.EOF
        m_rsOTBankTMCharges.Delete
        m_rsOTBankTMCharges.MoveNext
    Wend
    
    m_rsOTBankTMCharges.Filter = ""
    
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "RemoveBankTimeCharges"
End Sub

Private Sub RemoveSystemRefusalTimeCharges( _
    ByRef rsTimeCharges As ADODB.Recordset, _
    ByVal r_colEmps As Collection)

    Dim intI As Integer
    Dim l_lngEmp As Long

    On Error GoTo eh
    
    rsTimeCharges.Filter = "LAST_MOD_EMP_NUM = " & str(m_lngSystemEmpNumber)
    
    
    If RecordsExist(rsTimeCharges) Then
        If r_colEmps Is Nothing Then
            rsTimeCharges.MoveFirst
            While Not rsTimeCharges.EOF
                rsTimeCharges.Delete
                rsTimeCharges.MoveNext
            Wend
        Else
            For intI = 1 To r_colEmps.Count
                l_lngEmp = r_colEmps.Item(intI)
                rsTimeCharges.Filter = "EMP_NUM = " & str(l_lngEmp) & _
                        " AND " & "LAST_MOD_EMP_NUM = " & str(m_lngSystemEmpNumber)
                If RecordsExist(rsTimeCharges) Then
                            rsTimeCharges.MoveFirst
                    While Not rsTimeCharges.EOF
                        rsTimeCharges.Delete
                        rsTimeCharges.MoveNext
                    Wend
                
                End If
            Next
        End If
    End If
    rsTimeCharges.Filter = ""
    
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "RemoveSystemRefusalTimeCharges"
End Sub

Private Function GetOppBankCharges() As ADODB.Recordset

 
On Error GoTo eh

    
    Dim l_dbOvertimeEqualizer As dbOvertimeEqualizer.cdOvertimeEqualizer
   
    Set l_dbOvertimeEqualizer = CreateObject("dbOvertimeEqualizer.cdOvertimeEqualizer")
    
    Set GetOppBankCharges = l_dbOvertimeEqualizer.GetOppOTBankCharges(m_lngOTOppNum)
         
    Set l_dbOvertimeEqualizer = Nothing
    
    
    Exit Function

eh:
    Set l_dbOvertimeEqualizer = Nothing
    TrsRaiseError m_strMODULENAME, "GetOppBankCharges"

End Function


Private Sub GetEmployeeStateForUnPost()
    
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_aryEmps As Variant
    Dim l_colEmps As Collection
    Dim l_strEmp As String
    
    On Error GoTo eh
    
    l_datStartDate = m_datOTOppChargeDate
    l_datEndDate = m_datOTOppChargeDate
    Set l_colEmps = New Collection
    
    m_rsOppBankCharges.MoveFirst
    While Not m_rsOppBankCharges.EOF
        l_strEmp = Trim(str(m_rsOppBankCharges!EMP_NUM))
        If Not IncludesString(l_strEmp, l_colEmps) Then
            l_colEmps.Add l_strEmp
        End If
        m_rsOppBankCharges.MoveNext
    Wend

    
    If l_colEmps.Count = 0 Then
        Set m_rsEmployeeOTData = Nothing
        Exit Sub
    End If
    
    l_aryEmps = ArrayFromCollection(l_colEmps)
    
    
    m_rsOppBankCharges.Sort = "CHARGE_DATE"
    m_rsOppBankCharges.MoveFirst
    While Not m_rsOppBankCharges.EOF
        If m_rsOppBankCharges!CHARGE_DATE < l_datStartDate Then
            l_datStartDate = m_rsOppBankCharges!CHARGE_DATE
        End If
        If m_rsOppBankCharges!CHARGE_DATE > l_datEndDate Then
            l_datEndDate = m_rsOppBankCharges!CHARGE_DATE
        End If
        m_rsOppBankCharges.MoveNext
    Wend
    
    Set m_rsEmployeeOTData = GetEmpOTData(l_aryEmps, l_datStartDate, l_datEndDate)
    
    
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "GetEmployeeState"
End Sub

Private Function BankShouldBeCreated( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datDate As Date, _
    ByVal v_intOTBankCode As Integer _
    ) As Boolean
    
'The Newton Power plant uses a Volunteer System to canvass overtime
'The Newton employees that wish to be canvassed for overtime have a Volunteer bank
'Currently TRIS identifies an employee as a Volunteer if the employee has a volunteer bank
'So the Volunteer bank will not created in this process

'The current TRIS 16-bit application only charged the POE, UOE, and COE banks
'if the employee already had that bank

'This method is invoke because the employee does not have the bank

'Please note that the OT group has a standing bank
'   that is used to rank employees on a canvass
'The OT Group can also charge other banks, which allows the cross charging between groups
'The v_intOTBankCode is in the list of banks to charge for the OT Group
'The criteria below determines if the bank should be created

'To make it more flexible for change the following was done
'   1.  Do not create a Volunteer Bank if the employee does not have one
'   2.  Always create the standard bank
'   3.  Always create the standing bank for the OT Group
'   4.  Do not create a bank if it is not 3 or 4 above

'JimB 07/20/05
' Modified the logic to only create the Standard overtime bank
' This forces a new overtime bank's standing to be set before it is used

    
    On Error GoTo eh
        
'    If v_intOTBankCode = m_intVolunteerBankCode Then
'        BankShouldBeCreated = False
'        Exit Function
'    End If
    
    If v_intOTBankCode = m_intStandardBankCode Then
        BankShouldBeCreated = True
    Else
        BankShouldBeCreated = False
    End If
    
'    If v_intOTBankCode = m_intOTGroupStandingBank Then
'        BankShouldBeCreated = True
'    Else
'        BankShouldBeCreated = False
'    End If
       
    Exit Function
    
eh:
    TrsRaiseError m_strMODULENAME, "BankShouldBeCreated"
       

    
End Function

Private Function DetermineAcceptorCharges( _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    ByRef r_OppTimeCharges As ADODB.Recordset, _
    ByRef r_rsOTBankRequest As ADODB.Recordset _
    ) As Boolean

    Dim l_lngEmp As Long
    Dim l_blnOK As Boolean
    Dim l_strResp As String
    
    On Error GoTo eh
                                        
        
'Making an assumption that all employees on the canvass are
'from the same union
    v_rsCanvassSheet.Filter = ""
    v_rsCanvassSheet.MoveFirst
    m_rsEmployeeOTData.Filter = ""
    If RecordsExist(m_rsEmployeeOTData) Then
        m_rsEmployeeOTData.MoveFirst
        l_lngEmp = m_rsEmployeeOTData!EMP_NUM
    Else
        l_lngEmp = v_rsCanvassSheet!EMP_NUM
    End If
    If IsUnionOAOvertime(l_lngEmp, m_datOTOppChargeDate) Then
        'continue
    Else
        DetermineAcceptorCharges = True
        Exit Function
    End If
        
    l_blnOK = True
    While Not v_rsCanvassSheet.EOF
        l_strResp = DetermineOppResponse(v_rsCanvassSheet)
        If l_strResp <> "" Then
            If IsAcceptedResponse(l_strResp) And IsResponseChargeable(l_strResp) Then
                If m_intOTOppLoc = 36 Then
                    If Not ApplyAcceptorTimeRule(v_rsCanvassSheet, r_OppTimeCharges, r_rsOTBankRequest) Then
                        l_blnOK = False
                    End If
                Else
                    If Not ApplyAcceptorShiftRule(v_rsCanvassSheet, r_OppTimeCharges, r_rsOTBankRequest) Then
                        l_blnOK = False
                    End If
                End If
            End If
        End If
        v_rsCanvassSheet.MoveNext
    Wend
        

    DetermineAcceptorCharges = l_blnOK
    
    Exit Function
    
eh:
    TrsRaiseError m_strMODULENAME, "DetermineAcceptorCharges"

End Function


Private Function ApplyAcceptorShiftRule( _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    ByRef r_rsTimeCharges As ADODB.Recordset, _
    ByRef r_rsOTBankRequest As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strResp As String
    Dim l_lngEmpNum As Long
    Dim l_blnOK As Boolean
    Dim l_strFilter As String
    
    On Error GoTo eh
    
    l_lngEmpNum = v_rsCanvassSheet!EMP_NUM
    
    r_rsTimeCharges.Filter = "EMP_NUM = " & str(l_lngEmpNum)
    If Not RecordsExist(r_rsTimeCharges) Then
        ApplyAcceptorShiftRule = False
        AppendAcceptorTimeChargeError l_lngEmpNum, 1
        GoTo XIT
    End If
                
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
            "  AND CHARGE_DATE = " & m_datOTOppChargeDate & _
            "  AND HOURS_CHARGED <> 0 " & _
            "  AND SHIFT_CD = '" & m_strOTShiftCode & "'" & _
            "  AND OT_OPP_NUM <> " & str(m_lngOTOppNum)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsOTBankTMCharges.Filter = l_strFilter
    If Not RecordsExist(m_rsOTBankTMCharges) Then
        ApplyAcceptorShiftRule = True
        GoTo XIT
    End If
    
    m_rsOTBankTMCharges.MoveFirst
    While Not m_rsOTBankTMCharges.EOF
        If m_rsOTBankTMCharges!ACC_REF_IND = "A" Then
            If IsCallOutStandBy(r_rsTimeCharges) Then
                ApplyAcceptorShiftRule = True
                AppendAcceptorTimeChargeInformational l_lngEmpNum, 6
                GoTo XIT
            Else
                ApplyAcceptorShiftRule = False
                AppendAcceptorTimeChargeError l_lngEmpNum, 2
                GoTo XIT
            End If
        End If
        m_rsOTBankTMCharges.MoveNext
    Wend
    
'Only Refusal Charges exist for shift
    m_rsOTBankTMCharges.MoveFirst
    While Not m_rsOTBankTMCharges.EOF
        RemovePrevRefusal r_rsOTBankRequest
        AppendAcceptorTimeChargeInformational l_lngEmpNum, 1
        m_rsOTBankTMCharges!HOURS_CHARGED = CDec(0)
        m_rsOTBankTMCharges!Rate = CDec(0)
        m_rsOTBankTMCharges!LAST_MOD_EMP_NUM = m_lngLastModEmpNum
        m_rsOTBankTMCharges.MoveNext
    Wend
    
    ApplyAcceptorShiftRule = True
    

XIT:
    r_rsTimeCharges.Filter = ""
    m_rsOTBankTMCharges.Filter = ""
    
    Exit Function
eh:
    
    TrsRaiseError m_strMODULENAME, "ApplyAcceptorShiftRule"

End Function

Private Function ApplyAcceptorTimeRule( _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    ByRef r_rsTimeCharges As ADODB.Recordset, _
    ByRef r_rsOTBankRequest As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strResp As String
    Dim l_lngEmpNum As Long
    Dim l_blnOK As Boolean
    Dim l_strFilter As String
    Dim l_datBKStartDateTime As Date
    Dim l_datBKEndDateTime As Date
    Dim l_datTCStartDateTime As Date
    Dim l_datTCEndDateTime As Date
    Dim l_intDatesIntersect As Integer
    Dim l_decChargeHoursRemoved As Variant
    Dim l_decChargeHoursAdded As Variant
    
    On Error GoTo eh
    
    l_lngEmpNum = v_rsCanvassSheet!EMP_NUM
    
    r_rsTimeCharges.Filter = "EMP_NUM = " & str(l_lngEmpNum)
    If Not RecordsExist(r_rsTimeCharges) Then
        ApplyAcceptorTimeRule = False
        AppendAcceptorTimeChargeError l_lngEmpNum, 1
        GoTo XIT
    End If
                
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
            "  AND CHARGE_DATE = " & m_datOTOppChargeDate & _
            "  AND HOURS_CHARGED <> 0 " & _
            "  AND OT_OPP_NUM <> " & str(m_lngOTOppNum)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsOTBankTMCharges.Filter = l_strFilter
    If Not RecordsExist(m_rsOTBankTMCharges) Then
        ApplyAcceptorTimeRule = True
        GoTo XIT
    End If
    
    m_rsOTBankTMCharges.Sort = "DATE_WORKED, START_TIME"
    m_rsOTBankTMCharges.MoveFirst
    While Not m_rsOTBankTMCharges.EOF
        GetStartEndDateTimes m_rsOTBankTMCharges, l_datBKStartDateTime, l_datBKEndDateTime
        r_rsTimeCharges.MoveFirst
        DetermineContinuousStartEndTimes r_rsTimeCharges, l_datTCStartDateTime, l_datTCEndDateTime
        If TimePeriodsIntersect(l_datBKStartDateTime, l_datBKEndDateTime, _
                    l_datTCStartDateTime, l_datTCEndDateTime) Then
            If m_rsOTBankTMCharges!ACC_REF_IND = "A" Then
                AppendAcceptorTimeChargeInformational l_lngEmpNum, 2
            Else
                l_intDatesIntersect = DateIntersection( _
                    l_datBKStartDateTime, l_datBKEndDateTime, _
                    l_datTCStartDateTime, l_datTCEndDateTime)
                Select Case l_intDatesIntersect
                    Case 1
                        'Split Refusal Charge
                        If l_datBKStartDateTime <> l_datTCStartDateTime Then
                            AlterPrevRefusal r_rsOTBankRequest, l_decChargeHoursRemoved, l_datBKStartDateTime, l_datTCStartDateTime, True
                            AddPrevRefusal r_rsOTBankRequest, l_decChargeHoursAdded, l_datTCEndDateTime, l_datBKEndDateTime
                            l_decChargeHoursRemoved = l_decChargeHoursRemoved + l_decChargeHoursAdded
                            AppendAcceptorTimeChargeInformational l_lngEmpNum, 5, l_decChargeHoursRemoved
                        Else
                            If l_datBKEndDateTime = l_datTCEndDateTime Then
                                RemovePrevRefusal r_rsOTBankRequest
                                l_decChargeHoursRemoved = CDec(m_rsOTBankTMCharges!HOURS_CHARGED) * -1
                                AppendAcceptorTimeChargeInformational l_lngEmpNum, 3, l_decChargeHoursRemoved
                                m_rsOTBankTMCharges!HOURS_CHARGED = CDec(0)
                                m_rsOTBankTMCharges!Rate = CDec(0)
                                m_rsOTBankTMCharges!LAST_MOD_EMP_NUM = m_lngLastModEmpNum
                            Else
                                AlterPrevRefusal r_rsOTBankRequest, l_decChargeHoursRemoved, l_datTCEndDateTime, l_datBKEndDateTime, False
                                AppendAcceptorTimeChargeInformational l_lngEmpNum, 4, l_decChargeHoursRemoved
                            End If
                        End If
                    Case 2
                        'Remove Refusal Charge
                        RemovePrevRefusal r_rsOTBankRequest
                        l_decChargeHoursRemoved = CDec(m_rsOTBankTMCharges!HOURS_CHARGED) * -1
                        AppendAcceptorTimeChargeInformational l_lngEmpNum, 3, l_decChargeHoursRemoved
                        m_rsOTBankTMCharges!HOURS_CHARGED = CDec(0)
                        m_rsOTBankTMCharges!Rate = CDec(0)
                        m_rsOTBankTMCharges!LAST_MOD_EMP_NUM = m_lngLastModEmpNum
                    Case 3
                        'End Refusal charge at time charge start
                        AlterPrevRefusal r_rsOTBankRequest, l_decChargeHoursRemoved, l_datBKStartDateTime, l_datTCStartDateTime, True
                        AppendAcceptorTimeChargeInformational l_lngEmpNum, 4, l_decChargeHoursRemoved
                    Case 4
                        'Start Refusal charge at time charge end
                        AlterPrevRefusal r_rsOTBankRequest, l_decChargeHoursRemoved, l_datTCEndDateTime, l_datBKEndDateTime, False
                        AppendAcceptorTimeChargeInformational l_lngEmpNum, 4, l_decChargeHoursRemoved
                    Case 5
                        'Do nothing
                End Select
            End If
        End If
        m_rsOTBankTMCharges.MoveNext
    Wend
    
    ApplyAcceptorTimeRule = True
    

XIT:
    r_rsTimeCharges.Filter = ""
    m_rsOTBankTMCharges.Filter = ""
    
    Exit Function
eh:
    
    TrsRaiseError m_strMODULENAME, "ApplyAcceptorTimeRule"

End Function
Private Sub RemovePrevRefusal(ByRef r_rsRequest As ADODB.Recordset)

 
    Dim l_decHours As Variant
    Dim l_decTotalHours As Variant
    
    On Error GoTo eh
    
    If r_rsRequest Is Nothing Then
        Set r_rsRequest = CreateRequestRecordset()
    End If
    
    l_decHours = CDec(m_rsOTBankTMCharges!HOURS_CHARGED)
    l_decTotalHours = l_decHours

' Add Request
    r_rsRequest.AddNew
    r_rsRequest!REC_STATUS = "D"
    r_rsRequest!EMP_NUM = m_rsOTBankTMCharges!EMP_NUM
    r_rsRequest!CHARGE_DATE = m_rsOTBankTMCharges!CHARGE_DATE
    r_rsRequest!EXCEPTION_CD = ""
    r_rsRequest!HOURS = m_rsOTBankTMCharges!HOURS
    r_rsRequest!HOURS_CHARGED = l_decTotalHours * -1
    r_rsRequest!OT_BANK_CODE = Null
    r_rsRequest!OT_RESPONSE = ""
    r_rsRequest!OT_OPP_NUM = m_rsOTBankTMCharges!OT_OPP_NUM
    r_rsRequest!TIME_ENTRY_IND = "N"
    r_rsRequest!START_TIME = m_rsOTBankTMCharges!START_TIME
    r_rsRequest!OVER_EXC_CD = ""
    r_rsRequest!DATE_WORKED = m_rsOTBankTMCharges!DATE_WORKED
    r_rsRequest!END_TIME = m_rsOTBankTMCharges!END_TIME
    r_rsRequest!BANK_CHARGE_DATE = Null
    r_rsRequest!OT_MOD_DESC = "Worked Shift -Refusal Removed"
    r_rsRequest!CALL_TIME = ""
    r_rsRequest!END_TIME = m_rsOTBankTMCharges!END_TIME
    r_rsRequest!ACC_REF_IND = m_rsOTBankTMCharges!ACC_REF_IND

    r_rsRequest.Update
    
    Exit Sub
eh:
    
    TrsRaiseError m_strMODULENAME, "RemovePrevRefusal"

End Sub

Private Sub AppendAcceptorTimeChargeError( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_intErrorType As Integer)
    
    Dim l_strErrDesc As String
    Dim l_bolError As Boolean
    
    On Error GoTo eh
    
    Select Case v_intErrorType
        Case 1
            l_strErrDesc = "Employee has no time charges "
        Case 2
            l_strErrDesc = "Employee accepted overtime for this shift.  " & _
                "It was opportunity (" & str(m_rsOTBankTMCharges!OT_OPP_NUM) & ")."
    End Select

    l_bolError = AppendBusinessError(m_rsErrors, _
        40000, _
        mstrCBUSINESSERR, _
        l_strErrDesc, _
        m_strMODULENAME, _
        0, , , "EMP_NUM", _
        str(v_lngEmpNum))
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "AppendAcceptorTimeChargeError"

End Sub

Private Sub AppendAcceptorTimeChargeInformational( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_intErrorType As Integer, _
            Optional v_decChargeHours As Variant = 0#)
            
    
    Dim l_strErrDesc As String
    Dim l_bolError As Boolean
    
    On Error GoTo eh
    
    Select Case v_intErrorType
        Case 1
            l_strErrDesc = "Employee's refusal charge for opportunity " & _
                "(" & str(m_rsOTBankTMCharges!OT_OPP_NUM) & ") " & _
                "has been removed because the employee worked " & _
                "the same shift on this opportunity. "
        Case 2
            l_strErrDesc = "Employee has worked time charged to this same time period.  " & _
                "It was for opportunity " & _
                "(" & str(m_rsOTBankTMCharges!OT_OPP_NUM) & ").  " & _
                "The employee's overtime bank was charged for both opportunities. "
        Case 3
            l_strErrDesc = "Employee's refusal charge for opportunity " & _
                "(" & str(m_rsOTBankTMCharges!OT_OPP_NUM) & ") " & _
                "has been removed because the employee worked " & _
                "the same time period on this opportunity. " & _
                " The refusal charge has been reduced by " & Format(v_decChargeHours, "##.0") & _
                " charged hours."

        Case 4
            l_strErrDesc = "Employee's refusal charge for opportunity " & _
                "(" & str(m_rsOTBankTMCharges!OT_OPP_NUM) & ") " & _
                "has been reduced by " & Format(v_decChargeHours, "##.0") & _
                " charged hours.  " & _
                "This is because the employee's time worked " & _
                "on this opportunity intersected the previous opportunity. "
                
        Case 5
            l_strErrDesc = "Employee's refusal charge for opportunity " & _
                "(" & str(m_rsOTBankTMCharges!OT_OPP_NUM) & ") " & _
                "has been reduced by " & Format(v_decChargeHours, "##.0") & _
                " charged hours.  " & _
                "This is because the employee's time worked " & _
                "on this opportunity intersected the previous opportunity. "
                
        Case 6
            l_strErrDesc = "Employee has worked time charged to this same shift.  " & _
                "It was for opportunity " & _
                "(" & str(m_rsOTBankTMCharges!OT_OPP_NUM) & ").  " & _
                "The employee's overtime bank was charged for both opportunities. "
    End Select

    l_bolError = AppendBusinessError(m_rsErrors, _
        10000, _
        mstrCBUSINESSINF, _
        l_strErrDesc, _
        m_strMODULENAME, _
        0, , , "EMP_NUM", _
        str(v_lngEmpNum))
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "AppendAcceptorTimeChargeInformational"

End Sub

Private Sub DetermineContinuousStartEndTimes( _
                ByVal v_rsTimeCharge As ADODB.Recordset, _
                ByRef r_datStartShift As Date, _
                ByRef r_datEndShift As Date, _
                Optional l_blnRefusedTime As Boolean = False)
' Split charges, ie. refusals should only be entered at the beginning
' or ending of a shift.  If entered in the middle, it will break the
' continuity of the shift.

    Dim l_varBookmark As Variant
    Dim l_strOrigSort As String
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_datDateWorkedStart As Date
    Dim l_datDateWorkedEnd As Date
    Dim l_strExcCd As String

On Error GoTo eh

    l_varBookmark = v_rsTimeCharge.Bookmark
    l_strOrigSort = v_rsTimeCharge.Sort

    v_rsTimeCharge.Sort = "DATE_WORKED ASC, START_TIME ASC"
    v_rsTimeCharge.MoveFirst
    If l_blnRefusedTime Then
        l_strExcCd = v_rsTimeCharge!EXCEPTION_CD
        While Not IsRefusalExceptionCode(l_strExcCd)
            v_rsTimeCharge.MoveNext
            l_strExcCd = v_rsTimeCharge!EXCEPTION_CD
        Wend
    Else
        l_strExcCd = v_rsTimeCharge!EXCEPTION_CD
        While IsRefusalExceptionCode(l_strExcCd)
            v_rsTimeCharge.MoveNext
            l_strExcCd = v_rsTimeCharge!EXCEPTION_CD
        Wend
    End If
    l_strStartTime = v_rsTimeCharge!START_TIME
    l_strEndTime = v_rsTimeCharge!END_TIME
    l_datDateWorkedStart = v_rsTimeCharge!DATE_WORKED
    l_datDateWorkedEnd = v_rsTimeCharge!DATE_WORKED

    Do Until v_rsTimeCharge.EOF
'        If TimespanContinous(DateTimeFrom(l_datDateWorkedStart, l_strStartTime), _
'                            DateTimeFrom(l_datDateWorkedEnd, l_strEndTime), _
'                            DateTimeFrom(v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!START_TIME), _
'                            DateTimeFrom(v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!END_TIME, True), _
'                            m_lngMINOFFALLOWCONTINUOUS) Then
            If l_blnRefusedTime Then
                If IsRefusalExceptionCode(v_rsTimeCharge!EXCEPTION_CD) Then
                    l_strEndTime = v_rsTimeCharge!END_TIME
                    l_datDateWorkedEnd = v_rsTimeCharge!DATE_WORKED
                End If
            Else
                If Not IsRefusalExceptionCode(v_rsTimeCharge!EXCEPTION_CD) Then
                    l_strEndTime = v_rsTimeCharge!END_TIME
                    l_datDateWorkedEnd = v_rsTimeCharge!DATE_WORKED
                End If
            End If
'        Else
'            Exit Do
'        End If
        v_rsTimeCharge.MoveNext
    Loop
    
    r_datStartShift = DateTimeFrom(l_datDateWorkedStart, l_strStartTime)
    r_datEndShift = DateTimeFrom(l_datDateWorkedEnd, l_strEndTime)
    v_rsTimeCharge.Sort = l_strOrigSort
    v_rsTimeCharge.Bookmark = l_varBookmark

    Exit Sub

eh:
    TrsRaiseError m_strMODULENAME, "DetermineContinuousStartEndTimes"
    
End Sub

Private Function TimespanContinous( _
                    ByVal v_datStart1 As Date, _
                    ByVal v_datEnd1 As Date, _
                    ByVal v_datStart2 As Date, _
                    ByVal v_datEnd2 As Date, _
                    Optional ByVal v_lngMinBreak As Long = 0) _
                As Boolean
                    
On Error GoTo eh

    If v_datStart1 >= v_datStart2 Then
        If DateAdd("n", v_lngMinBreak, v_datEnd2) >= v_datStart1 Then
            TimespanContinous = True
        Else
            TimespanContinous = False
        End If
    Else
        If DateAdd("n", v_lngMinBreak, v_datEnd1) >= v_datStart2 Then
            TimespanContinous = True
        Else
            TimespanContinous = False
        End If
    End If

    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "TimespanContinous"

End Function

Private Sub GetStartEndDateTimes( _
    ByVal v_rsCharges As ADODB.Recordset, _
    ByRef r_datStartDateTime As Date, _
    ByRef r_datEndDateTime As Date)
    
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    
    On Error GoTo eh
    
        l_strStartTime = v_rsCharges!START_TIME
        l_strEndTime = v_rsCharges!END_TIME
        l_datStartDate = v_rsCharges!DATE_WORKED
        If l_strStartTime > l_strEndTime Then
            l_datEndDate = DateAdd("d", 1, l_datStartDate)
        Else
            l_datEndDate = l_datStartDate
        End If
        r_datStartDateTime = DateTimeFrom(l_datStartDate, l_strStartTime)
        r_datEndDateTime = DateTimeFrom(l_datEndDate, l_strEndTime)
    
    
    Exit Sub
    
eh:
    TrsRaiseError m_strMODULENAME, "GetStartEndDateTimes"
End Sub



Private Sub AlterPrevRefusal( _
    ByRef r_rsRequest As ADODB.Recordset, _
    ByRef r_decHoursCharged As Variant, _
    ByVal v_datStartTime As Date, _
    ByVal v_datEndTime As Date, _
    ByVal v_blnAlterEnd As Boolean)

 
    Dim l_decHours As Variant
    Dim l_decChargedHours As Variant
    Dim l_decOrigHoursCharged As Variant
    Dim l_strExcCd As String
    Dim l_decFactor As Variant
    Dim l_decOrigHours As Variant
    Dim l_decOrigFactor As Variant
    
    On Error GoTo eh
    
    If r_rsRequest Is Nothing Then
        Set r_rsRequest = CreateRequestRecordset()
    End If
    
    l_decOrigHoursCharged = m_rsOTBankTMCharges!HOURS_CHARGED
    l_decOrigHours = m_rsOTBankTMCharges!HOURS
    l_decOrigFactor = m_rsOTBankTMCharges!Rate
    l_decHours = CDec(DateDiff("n", v_datStartTime, v_datEndTime) / 60)
    l_decHours = Round(l_decHours, 1)
    l_decChargedHours = l_decHours * l_decOrigFactor
    l_decChargedHours = Round(l_decChargedHours, 1)
    
    m_rsOTBankTMCharges!LAST_MOD_EMP_NUM = m_lngLastModEmpNum
    
    If v_blnAlterEnd Then
        m_rsOTBankTMCharges!END_TIME = Format(v_datEndTime, "HH:MM")
    Else
        m_rsOTBankTMCharges!START_TIME = Format(v_datStartTime, "HH:MM")
        m_rsOTBankTMCharges!DATE_WORKED = CDate(Format(v_datStartTime, "YYYY-MM-DD"))
    End If
    m_rsOTBankTMCharges!HOURS_CHARGED = CDec(l_decChargedHours)
    m_rsOTBankTMCharges!HOURS = CDec(l_decHours)
    
    r_decHoursCharged = l_decChargedHours - l_decOrigHoursCharged


' Add Request
    r_rsRequest.AddNew
    r_rsRequest!REC_STATUS = "D"
    r_rsRequest!EMP_NUM = m_rsOTBankTMCharges!EMP_NUM
    r_rsRequest!CHARGE_DATE = m_rsOTBankTMCharges!CHARGE_DATE
    r_rsRequest!EXCEPTION_CD = ""
    r_rsRequest!HOURS = m_rsOTBankTMCharges!HOURS
    r_rsRequest!HOURS_CHARGED = l_decChargedHours - l_decOrigHoursCharged
    r_rsRequest!OT_BANK_CODE = Null
    r_rsRequest!OT_RESPONSE = ""
    r_rsRequest!OT_OPP_NUM = m_rsOTBankTMCharges!OT_OPP_NUM
    r_rsRequest!TIME_ENTRY_IND = "N"
    r_rsRequest!START_TIME = m_rsOTBankTMCharges!START_TIME
    r_rsRequest!OVER_EXC_CD = ""
    r_rsRequest!DATE_WORKED = m_rsOTBankTMCharges!DATE_WORKED
    r_rsRequest!END_TIME = m_rsOTBankTMCharges!END_TIME
    r_rsRequest!BANK_CHARGE_DATE = Null
    r_rsRequest!OT_MOD_DESC = "Worked Shift -Refusal Altered"
    r_rsRequest!CALL_TIME = ""
    r_rsRequest!END_TIME = m_rsOTBankTMCharges!END_TIME
    r_rsRequest!ACC_REF_IND = m_rsOTBankTMCharges!ACC_REF_IND

    r_rsRequest.Update
    
    Exit Sub
eh:
    
    TrsRaiseError m_strMODULENAME, "AlterPrevRefusal"

End Sub

Private Function DeterimineExceptionCode( _
    ByVal v_rsTimeCharges As ADODB.Recordset _
    ) As String
    
    Dim l_strExcCd As String
        
    On Error GoTo eh
    
    v_rsTimeCharges.MoveFirst
    l_strExcCd = v_rsTimeCharges!EXCEPTION_CD
    While Not v_rsTimeCharges.EOF And l_strExcCd = ""
        If l_strExcCd = v_rsTimeCharges!EXCEPTION_CD Then
        Else
            l_strExcCd = ""
        End If
        v_rsTimeCharges.MoveNext
    Wend
    
    DeterimineExceptionCode = Trim(l_strExcCd)

    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "DeterimineExceptionCode"

End Function

Private Sub AddPrevRefusal( _
    ByRef r_rsRequest As ADODB.Recordset, _
    ByRef r_decChargeHoursAdded As Variant, _
    ByVal v_datStartTime As Date, _
    ByVal v_datEndTime As Date)

 
    Dim l_decHours As Variant
    Dim l_decChargedHours As Variant
    Dim l_decFactor As Variant
    Dim l_lngEmpNum As Long
    Dim l_lngOTOppNum As Long
    Dim l_datChargeDate As Date
    Dim l_strAccRefInd As String
    
    
    On Error GoTo eh
    
    If r_rsRequest Is Nothing Then
        Set r_rsRequest = CreateRequestRecordset()
    End If
    
    l_decHours = CDec(DateDiff("n", v_datStartTime, v_datEndTime) / 60)
    l_decHours = Round(l_decHours, 1)
    l_decFactor = CDec(m_rsOTBankTMCharges("Rate").OriginalValue)
    l_decChargedHours = l_decHours * l_decFactor
    l_decChargedHours = Round(l_decChargedHours, 1)
    r_decChargeHoursAdded = l_decChargedHours

' Add Request
    r_rsRequest.AddNew
    r_rsRequest!REC_STATUS = "I"
    r_rsRequest!EMP_NUM = m_rsOTBankTMCharges!EMP_NUM
    r_rsRequest!CHARGE_DATE = m_rsOTBankTMCharges!CHARGE_DATE
    r_rsRequest!EXCEPTION_CD = ""
    r_rsRequest!HOURS = l_decHours
    r_rsRequest!HOURS_CHARGED = l_decChargedHours
    r_rsRequest!OT_BANK_CODE = Null
    r_rsRequest!OT_RESPONSE = ""
    r_rsRequest!OT_OPP_NUM = m_rsOTBankTMCharges!OT_OPP_NUM
    r_rsRequest!TIME_ENTRY_IND = "N"
    r_rsRequest!START_TIME = Format(v_datStartTime, "HH:MM")
    r_rsRequest!OVER_EXC_CD = ""
    r_rsRequest!DATE_WORKED = CDate(Format(v_datStartTime, "YYYY-MM-DD"))
    r_rsRequest!END_TIME = Format(v_datEndTime, "HH:MM")
    r_rsRequest!BANK_CHARGE_DATE = Null
    r_rsRequest!OT_MOD_DESC = "Worked Shift -Refusal Altered"
    r_rsRequest!CALL_TIME = ""
    r_rsRequest!ACC_REF_IND = "R"

    r_rsRequest.Update
    
    Exit Sub
eh:
    
    TrsRaiseError m_strMODULENAME, "AddPrevRefusal"

End Sub

Public Function IsHolidayRestException( _
    ByVal v_strExcCd As String) As Boolean
    
    Dim l_strExcCd As String
    
    On Error GoTo eh
    
    l_strExcCd = Trim(v_strExcCd)
    
    If l_strExcCd = "" Then
        IsHolidayRestException = False
        Exit Function
    End If
    
    
    m_rsHolidayRestExcCds.Filter = "EXCEPTION_CD = '" & l_strExcCd & "'"
    If RecordsExist(m_rsHolidayRestExcCds) Then
        IsHolidayRestException = True
    Else
        IsHolidayRestException = False
    End If

    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "IsHolidayRestException"
    
End Function

Private Function AnyAcceptorChanges( _
    ByVal v_OppTimeCharges As ADODB.Recordset _
    ) As Boolean
    
    Dim l_blnAcceptorChanged As Boolean
    Dim l_strExcCd As String
    
    On Error GoTo eh
    
    v_OppTimeCharges.Filter = adFilterPendingRecords
    If Not RecordsExist(v_OppTimeCharges) Then
        l_blnAcceptorChanged = False
        GoTo XIT
    End If
    
    l_blnAcceptorChanged = False
    v_OppTimeCharges.MoveFirst
    While Not v_OppTimeCharges.EOF And Not l_blnAcceptorChanged
        Select Case v_OppTimeCharges.Status
            Case adRecNew
                l_strExcCd = v_OppTimeCharges!EXCEPTION_CD
            Case adRecModified
                l_strExcCd = v_OppTimeCharges("EXCEPTION_CD").OriginalValue
            Case adRecDeleted
                l_strExcCd = v_OppTimeCharges("EXCEPTION_CD").OriginalValue
            Case Else
                l_strExcCd = ""
        End Select

        If Not IsRefusalExceptionCode(l_strExcCd) Then
            l_blnAcceptorChanged = True
        End If
        v_OppTimeCharges.MoveNext
    Wend
    
XIT:
    v_OppTimeCharges.Filter = ""
    AnyAcceptorChanges = l_blnAcceptorChanged

    
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "AnyAcceptorChanges"
End Function


Private Function IsCallOutStandBy( _
    ByVal v_rsTimeCharges As ADODB.Recordset) As Boolean

    Dim l_blnAllCStime As Boolean
    Dim l_strOverExcCd As String
    
    On Error GoTo eh
    
    v_rsTimeCharges.MoveFirst
    l_blnAllCStime = True
    While Not v_rsTimeCharges.EOF
        l_strOverExcCd = Trim(v_rsTimeCharges!OVER_EXC_CD)
        If l_strOverExcCd <> "CS" Then
            l_blnAllCStime = False
        End If
        v_rsTimeCharges.MoveNext
    Wend
    
    IsCallOutStandBy = l_blnAllCStime
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "IsCallOutStandBy"

End Function
    
    
Private Sub GetSupportInfo()

    Dim l_dbOvertimeEqualizer As dbOvertimeEqualizer.cdOvertimeEqualizer
    Dim l_rs As ADODB.Recordset

    On Error GoTo eh
    
    Set l_dbOvertimeEqualizer = CreateObject("dbOvertimeEqualizer.cdOvertimeEqualizer")
    
    Set l_rs = l_dbOvertimeEqualizer.GetOTGroupCanvassBank(m_intOTGroup)
    
    If RecordsExist(l_rs) Then
        m_intOTGroupStandingBank = l_rs!OT_STAND_BANK_CD
    Else
        m_intOTGroupStandingBank = 0
    End If
    Set l_rs = Nothing
    
    Set m_rsOTGroupBanks = l_dbOvertimeEqualizer.GetOTGroupBanks(m_intOTGroup)
    Set m_rsOTResponses = l_dbOvertimeEqualizer.GetOvertimeResponses()
    Set m_rsOTBankTMCharges = l_dbOvertimeEqualizer.GetOvertimeBankTimeCharges(m_datOTOppChargeDate, m_strOTShiftCode, 0)
    Set m_rsHolidayRestExcCds = l_dbOvertimeEqualizer.GetHolidayRestExceptions
    Set l_dbOvertimeEqualizer = Nothing
    Set m_rsOTExceptionCodes = GetOvertimeExceptions()
    
    Exit Sub

eh:
    Set l_dbOvertimeEqualizer = Nothing
    TrsRaiseError m_strMODULENAME, "GetSupportInfo"
End Sub


Private Sub ApplyRedTimeNotAssignedRule( _
    ByVal v_colEmps As Collection, _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    ByRef r_rsTimeCharges As ADODB.Recordset)


    Dim l_lngEmpNum As Long
    Dim l_intEmpLoc As Integer
    Dim l_intCount As Integer
    Dim l_intI As Integer
    Dim l_strFilter As String
    Dim l_strErrorDesc As String
    Dim l_strTimeFilter As String
    Dim l_blnTimeCharge As Boolean
    Dim l_strEmpDept As String

'Need to find out if this is only for GCMS employees only (dept 180)
'If so, uncomment out the commented code and remove
'        If m_intOTOppLoc <> l_intEmpLoc Then
    
    On Error GoTo eh

    l_intCount = v_colEmps.Count
    For l_intI = 1 To l_intCount
        l_lngEmpNum = v_colEmps(l_intI)
        v_rsCanvassSheet.Filter = "EMP_NUM = " & str(l_lngEmpNum)
        l_intEmpLoc = v_rsCanvassSheet!LOC_CD
        l_strEmpDept = GetEmployeeDeptNum(l_lngEmpNum, m_datOTOppChargeDate)
        If m_intOTOppLoc <> l_intEmpLoc And l_strEmpDept = "180" Then
            l_strTimeFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
                " AND EXCEPTION_CD <> 'RF0' "
            l_strTimeFilter = MDAC25FilterFix(l_strTimeFilter)
            r_rsTimeCharges.Filter = l_strTimeFilter
            If RecordsExist(r_rsTimeCharges) Then
                    'write BRV stating that the employee has already been charged for this shift
                    'set the time charge to 0 hours
                    l_strErrorDesc = "No RedTime Charged.  The employee is assigned to location " _
                        & str(l_intEmpLoc) & " and can not be charged for refused overtime at the" _
                        & " canvass location " & str(m_intOTOppLoc) & "."
                    AppendRedTimeNotChargedWarning l_lngEmpNum, l_strErrorDesc
                    SetTimeChargesToZero l_lngEmpNum, r_rsTimeCharges
            End If
        End If
            

    Next l_intI
    v_rsCanvassSheet.Filter = ""
    
Exit Sub
eh:
   
    TrsRaiseError m_strMODULENAME, "ApplyRedTimeNotAssignedRule"
                               
End Sub


Private Function OverlapsPreviousCharge( _
    ByVal v_rsTimeCharges As ADODB.Recordset _
    ) As Boolean

    Dim l_blnOverlaps As Boolean
    Dim l_strOverExcCd As String
    Dim l_strPrevChrgStartTime As String
    Dim l_strPrevChrgEndTime As String
    Dim l_datPrevChrgStartDate As Date
    Dim l_datPrevChrgEndDate As Date
    Dim l_datPrevChrgStartDateTime As Date
    Dim l_datPrevChrgEndDateTime As Date
    Dim l_strCurrChrgStartTime As String
    Dim l_strCurrChrgEndTime As String
    Dim l_datCurrChrgStartDate As Date
    Dim l_datCurrChrgEndDate As Date
    Dim l_datCurrChrgStartDateTime As Date
    Dim l_datCurrChrgEndDateTime As Date
    Dim l_intDateIntersection As Integer
    
    On Error GoTo eh
    
    l_blnOverlaps = False
    l_strPrevChrgStartTime = m_rsOTBankTMCharges!START_TIME
    l_strPrevChrgEndTime = m_rsOTBankTMCharges!END_TIME
    l_datPrevChrgStartDate = m_rsOTBankTMCharges!DATE_WORKED
    If l_strPrevChrgStartTime > l_strPrevChrgEndTime Then
        l_datPrevChrgEndDate = DateAdd("d", 1, l_datPrevChrgStartDate)
    Else
        l_datPrevChrgEndDate = l_datPrevChrgStartDate
    End If
    l_datPrevChrgStartDateTime = DateTimeFrom(l_datPrevChrgStartDate, l_strPrevChrgStartTime)
    l_datPrevChrgEndDateTime = DateTimeFrom(l_datPrevChrgEndDate, l_strPrevChrgEndTime)
    
    v_rsTimeCharges.MoveFirst
    While Not v_rsTimeCharges.EOF
        l_strCurrChrgStartTime = v_rsTimeCharges!START_TIME
        l_strCurrChrgEndTime = v_rsTimeCharges!END_TIME
        l_datCurrChrgStartDate = v_rsTimeCharges!DATE_WORKED
        If l_strCurrChrgStartTime > l_strCurrChrgEndTime Then
            l_datCurrChrgEndDate = DateAdd("d", 1, l_datCurrChrgStartDate)
        Else
            l_datCurrChrgEndDate = l_datCurrChrgStartDate
        End If
        l_datCurrChrgStartDateTime = DateTimeFrom(l_datCurrChrgStartDate, l_strCurrChrgStartTime)
        l_datCurrChrgEndDateTime = DateTimeFrom(l_datCurrChrgEndDate, l_strCurrChrgEndTime)
        l_intDateIntersection = DateIntersection( _
                l_datPrevChrgStartDateTime, l_datPrevChrgEndDateTime, _
                l_datCurrChrgStartDateTime, l_datCurrChrgEndDateTime)
        If l_intDateIntersection = 0 Or l_intDateIntersection = 5 Then
            'No Intersection
        Else
            l_blnOverlaps = True
        End If
        v_rsTimeCharges.MoveNext
    Wend
    
    OverlapsPreviousCharge = l_blnOverlaps
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "OverlapsPreviousCharge"

End Function

Private Function RefusalCodeChanges( _
    ByVal v_OppTimeCharges As ADODB.Recordset, _
    ByVal v_rsCanvassSheet As ADODB.Recordset, _
    ByRef r_colEmps As Collection _
    ) As Boolean
    
    Dim l_blnRefusalChanged As Boolean
    Dim l_strExcCd As String
    Dim l_lngEmp As Long
    Dim l_strResponse As String
    Dim l_strOrigResponse As String
    
    
    On Error GoTo eh
    
    l_blnRefusalChanged = False
    Set r_colEmps = New Collection
    v_rsCanvassSheet.MoveFirst
    While Not v_rsCanvassSheet.EOF
        l_lngEmp = v_rsCanvassSheet!EMP_NUM
        l_strResponse = DetermineOppResponse(v_rsCanvassSheet)
        l_strOrigResponse = DetermineOppResponse(v_rsCanvassSheet, True)
        If l_strOrigResponse <> l_strResponse Then
            If IsRefusedResponse(l_strOrigResponse) And IsRefusedResponse(l_strResponse) Then
                l_blnRefusalChanged = True
                If EmployeeInList(l_lngEmp, r_colEmps) Then
                Else
                    r_colEmps.Add l_lngEmp
                End If
            End If
        End If
        v_rsCanvassSheet.MoveNext
    Wend
    
    
    v_OppTimeCharges.Filter = adFilterPendingRecords
    If Not RecordsExist(v_OppTimeCharges) Then
        GoTo XIT
    End If
    
    v_OppTimeCharges.MoveFirst
    While Not v_OppTimeCharges.EOF
        Select Case v_OppTimeCharges.Status
            Case adRecNew
                l_strExcCd = v_OppTimeCharges!EXCEPTION_CD
            Case adRecModified
                l_strExcCd = v_OppTimeCharges("EXCEPTION_CD").OriginalValue
            Case adRecDeleted
'                l_strExcCd = v_OppTimeCharges("EXCEPTION_CD").OriginalValue
                l_strExcCd = ""
            Case Else
                l_strExcCd = ""
        End Select

        If IsRefusalExceptionCode(l_strExcCd) Then
            l_blnRefusalChanged = True
            l_lngEmp = v_OppTimeCharges!EMP_NUM
            If EmployeeInList(l_lngEmp, r_colEmps) Then
            Else
                r_colEmps.Add l_lngEmp
            End If
        End If
        v_OppTimeCharges.MoveNext
    Wend
    

    
XIT:
    v_OppTimeCharges.Filter = ""
    v_rsCanvassSheet.Filter = ""
    RefusalCodeChanges = l_blnRefusalChanged

    
    Exit Function

eh:
    TrsRaiseError m_strMODULENAME, "RefusalCodeChanges"
End Function
    
    
Private Function IsMissouriCallCenter(ByVal vstrDept As String) As Boolean


    On Error GoTo eh
    
    Dim l_obj As dbRules.cdRules
    
    Set l_obj = CreateObject("dbRules.cdRules")
    IsMissouriCallCenter = l_obj.DoesRuleApply(258, 4, vstrDept)
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    TrsRaiseError m_strMODULENAME, "IsMissouriCallCenter"
End Function
        


