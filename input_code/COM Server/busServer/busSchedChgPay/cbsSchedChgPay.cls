VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsSchedChgPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cdSchedChgPay"
Private m_datMinChargeDate As Date
Private m_datMaxChargeDate As Date
Private m_aryEmpNum As Variant
Private m_colEmployees As Collection
Private m_aryEmployees As Variant
Private m_rsEmployee As ADODB.Recordset

Public Function GetSchedChgPay(ByVal v_aryEmpNum As Variant, _
                                ByVal v_dtStartDate As Date, _
                                ByVal v_dtEndDate As Date) As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_objSchedChgPay As dbSchedChgPay.cdSchedChgPay
    Dim l_rsSchedChgPay As ADODB.Recordset
    Dim l_rsScopeParms  As ADODB.Recordset
    Dim l_rsSchedule As ADODB.Recordset
    Dim l_strFilter As String
    
    Set l_objSchedChgPay = CreateObject("dbSchedChgPay.cdSchedChgPay")
    Set l_rsSchedChgPay = l_objSchedChgPay.GetSchedChgPay(v_aryEmpNum, v_dtStartDate, v_dtEndDate)
    Set l_objSchedChgPay = Nothing
    
    If RecordsExist(l_rsSchedChgPay) Then
        '//-----build summary recordset with EmpNum, StartDate, and EndDate-----//
        Set l_rsScopeParms = GetScopeParms(l_rsSchedChgPay)
        '//-----Check SchedChgPay against schedule for hours > 0 and flag off days (Orphan_Ind=N for hours > 0, otherwise Y)
        Set l_rsSchedule = GetSchedule(l_rsScopeParms)
        
        If RecordsExist(l_rsSchedule) Then
            With l_rsSchedChgPay
                .MoveFirst
                Do While Not .EOF
                    l_strFilter = MDAC25FilterFix("EMP_NUM = " & !EMP_NUM & " AND CHARGE_DATE = #" & !CHARGE_DATE & "# AND SHIFT_HOURS > 0")
                    l_rsSchedule.Filter = l_strFilter
                    If RecordsExist(l_rsSchedule) Then
                        !ORPHAN_IND = "N"
                        .Update
                    End If
                
                    .MoveNext
                Loop
            End With
        End If
    End If
    
    Set GetSchedChgPay = l_rsSchedChgPay
    
    Set l_objSchedChgPay = Nothing
    Set l_rsSchedChgPay = Nothing
    Set l_rsScopeParms = Nothing
    Set l_rsSchedule = Nothing

    Exit Function
eh:
    Set GetSchedChgPay = Nothing
    Set l_objSchedChgPay = Nothing
    Set l_rsSchedChgPay = Nothing
    Set l_rsScopeParms = Nothing
    Set l_rsSchedule = Nothing

    TrsRaiseError m_strCMODULENAME, "GetSchedChgPay", , , , , True
End Function

'.Fields.Append "EMP_NUM", adInteger, , l_lngCFIELDATTR
'.Fields.Append "CHARGE_DATE", adDate, , l_lngCFIELDATTR
'.Fields.Append "BASELINE_SCHED_CD", adInteger, , l_lngCFIELDATTR
'.Fields.Append "EXCEPTION_CD", adChar, 3, l_lngCFIELDATTR
'.Fields.Append "PAYMENT_TYPE_CD", adChar, 1, l_lngCFIELDATTR
'
'.Fields.Append "LAST_MOD_EMP_NUM", adInteger, , l_lngCFIELDATTR
'.Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, l_lngCFIELDATTR
'.Fields.Append "CREATED_EMP_NUM", adInteger, , l_lngCFIELDATTR
'.Fields.Append "CREATED_TIMESTAMP", adChar, 26, l_lngCFIELDATTR
'
'.Fields.Append "KEY_EMP_NUM", adInteger, , l_lngCFIELDATTR
'.Fields.Append "KEY_CHARGE_DATE", adDate, , l_lngCFIELDATTR
'.Fields.Append "RECORD_STATUS", adChar, 1, l_lngCFIELDATTR

Public Function SaveSchedChgPay(ByVal v_rsSchedChgPayRequest As ADODB.Recordset, _
                                ByRef r_rsErrors As ADODB.Recordset, _
                                ByVal v_lTRISUser As Long) As Boolean
    Dim l_objSchedChgPay As dbSchedChgPay.cdSchedChgPayTrans
    Dim l_rsErrors As ADODB.Recordset
    Dim bError As Boolean

    On Error GoTo eh
    bError = False

    Set l_rsErrors = CreateNewBusinessErrorSet
    If PreValidateSchedChgPay(v_rsSchedChgPayRequest, l_rsErrors) Then
        '//Modify RECORD_STATUS = "C" requests, to "D" and "I" for requests that are changing the key columns (10/23/2001)
        If Not ProcessKeyChange(v_rsSchedChgPayRequest) Then
            bError = True
        End If
            
        If ValidateSchedChgPay(v_rsSchedChgPayRequest, l_rsErrors) Then
            If PostValidateSchedChgPay(v_rsSchedChgPayRequest, l_rsErrors) Then
            Else
                bError = True
            End If
        Else
            bError = True
        End If
    Else
        bError = True
    End If
    
    If Not bError Then
        Set l_objSchedChgPay = CreateObject("dbSchedChgPay.cdSchedChgPayTrans")
        bError = Not l_objSchedChgPay.SaveSchedChgPay(v_rsSchedChgPayRequest, v_lTRISUser)
        Set l_objSchedChgPay = Nothing
    End If

    If bError Then
        SaveSchedChgPay = False
        GetObjectContext.SetAbort
    Else
        SaveSchedChgPay = True
        GetObjectContext.SetComplete
    End If

    Set r_rsErrors = l_rsErrors '//-----pass back to client

    Set l_objSchedChgPay = Nothing
    Set l_rsErrors = Nothing
    Exit Function
eh:
    Set l_objSchedChgPay = Nothing
    Set l_rsErrors = Nothing
    
    SaveSchedChgPay = False
    GetObjectContext.SetAbort
'    Select Case Err.Number
'        Case trErrbusSchedDemandDuplicateLineAsgnm
'            Err.Clear
'            With v_rsSchedChgPayRequest
'                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
'                    "Cannot perform " & IIf(!RECORD_STATUS = "I", "insert", "update") & _
'                    " because an SchedChgPay already exists.", m_strCMODULENAME, _
'                    1, !RECORD_ID, "LINE_NUM", _
'                    Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
'                    CStr(!WG_NUM), _
'                    CStr(!LINE_NUM), _
'                    Format(!START_DATE, "mm/dd/yyyy"), _
'                    Format(!END_DATE, "mm/dd/yyyy"), _
'                    CStr(!EMP_NUM)
'            End With
'
'            Exit Function
'        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveSchedChgPay", , , , , True
'    End Select

End Function


Public Function GetSchedChgPayPending( _
                    ByVal r_rsEmpSchedDates As ADODB.Recordset, _
                    Optional ByVal v_rsWGBaseSchedChanges As ADODB.Recordset, _
                    Optional ByVal v_rsWGLineSchedChanges As ADODB.Recordset, _
                    Optional ByVal v_rsWGAssignChanges As ADODB.Recordset, _
                    Optional ByVal v_rsWGAssignOtherChanges As ADODB.Recordset, _
                    Optional ByVal v_rsEmpLineAssignChanges As ADODB.Recordset, _
                    Optional ByVal v_rsPersonalChanges As ADODB.Recordset) As ADODB.Recordset
' Check for any schedule change payments during the scope of the schedule change.
' If any found, determine whether the payment is scheduled on an off-day.  Must call
' GAB with the pending changes to determine the schedule based on the unsaved
' schedule changes.
' Answer Nothing if invalid request.
' Answer an empty recordset when schedule change payments do not exist.
' 8/8/02 RSG - Currently busSchedDemand and busWorkgroup use this method.  If there
' are changes to the process, may need to make same change in busSchedPersonal and busSchedTrade.
' busSchedPersonal can be easily converted to use this method, instead of copying change.
' busSchedTrade may never want to use this method since it already has the net schedule
' available to it without calling GAB.  Also, the GABPending is not ready to accept the
' trade updates since it was coded first and differently for cascade.
    Dim l_varEmployees As Variant
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_dbSchedChgPay As dbSchedChgPay.cdSchedChgPay
    Dim l_rsSchedChgPay As ADODB.Recordset
    Dim l_rsSchedChgPayDates As ADODB.Recordset
    Dim l_cbsSchedule As busSchedule.cbsSchedule
    Dim l_rsSchedule As ADODB.Recordset
                    
On Error GoTo eh

    DetermineGetSchedChgPayScope r_rsEmpSchedDates, l_varEmployees, l_datStart, l_datEnd
    
    If IsEmpty(l_varEmployees) Then
        Set GetSchedChgPayPending = Nothing
        Exit Function
    End If
                          
    Set l_dbSchedChgPay = CreateObject("dbSchedChgPay.cdSchedChgPay")
    Set l_rsSchedChgPay = l_dbSchedChgPay.GetSchedChgPay(l_varEmployees, l_datStart, l_datEnd)
    Set l_dbSchedChgPay = Nothing
    
    If Not RecordsExist(l_rsSchedChgPay) Then
        Set GetSchedChgPayPending = l_rsSchedChgPay
        Exit Function
    End If

' Reduce scope of GAB by only requesting the dates schedule change payments exist.
    Set l_rsSchedChgPayDates = GetSchedChgPayDates(l_rsSchedChgPay)

    Set l_cbsSchedule = CreateObject("busSchedule.cbsSchedule")
    Set l_rsSchedule = l_cbsSchedule.GetPendingAdjustedBaseSchedule( _
            l_rsSchedChgPayDates, Nothing, False, _
            trBaseSchedNoDetails, trAllScheduleLevels, Nothing, _
            v_rsWGBaseSchedChanges, v_rsWGLineSchedChanges, _
            v_rsWGAssignChanges, v_rsWGAssignOtherChanges, _
            v_rsEmpLineAssignChanges, v_rsPersonalChanges)
    Set l_cbsSchedule = Nothing

    Do Until l_rsSchedChgPay.EOF
        If IsPaymentScheduledOnOffDay(l_rsSchedChgPay, l_rsSchedule) Then
            l_rsSchedChgPay!ORPHAN_IND = "Y"
        Else
            l_rsSchedChgPay!ORPHAN_IND = "N"
        End If
        l_rsSchedChgPay.Update
        l_rsSchedChgPay.MoveNext
    Loop

    Set GetSchedChgPayPending = l_rsSchedChgPay
    Set l_rsSchedChgPay = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsSchedChgPayDates = Nothing
    GetObjectContext.SetComplete
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    Set l_rsSchedChgPayDates = Nothing
    Set l_dbSchedChgPay = Nothing
    Set l_rsSchedChgPay = Nothing
    Set l_rsSchedule = Nothing
    Set l_cbsSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetSchedChgPayPending"
End Function


Private Sub DetermineGetSchedChgPayScope( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_varEmployees As Variant, _
                    ByRef r_datStartDate As Date, _
                    ByRef r_datEndDate As Date)
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_colDistinctEmpNums As cuSet
    Const l_datLowDate As Date = #1/1/1900#
    Const l_datHighDate As Date = #12/31/2079#

On Error GoTo eh

'guard for empty recordset
    If Not RecordsExist(r_rsEmpSchedDates) Then
        r_varEmployees = Empty
        r_datStartDate = 0
        r_datEndDate = 0
        Exit Sub
    Else
        r_rsEmpSchedDates.MoveFirst
    End If

' seed the dates with high and low dates
' there should be at least 1 valid record
    l_datStartDate = l_datHighDate
    l_datEndDate = l_datLowDate
    Set l_colDistinctEmpNums = New cuSet
    Do Until r_rsEmpSchedDates.EOF
        l_colDistinctEmpNums.Add CLng(r_rsEmpSchedDates!EMP_NUM.Value)
        l_datStartDate = MinDate(r_rsEmpSchedDates!START_DATE.Value, l_datStartDate)
        l_datEndDate = MaxDate(r_rsEmpSchedDates!END_DATE.Value, l_datEndDate)
        r_rsEmpSchedDates.MoveNext
    Loop

    r_datStartDate = l_datStartDate
    r_datEndDate = l_datEndDate
    r_varEmployees = ArrayFromCollection(l_colDistinctEmpNums.Contents)
    r_rsEmpSchedDates.MoveFirst
    Set l_colDistinctEmpNums = Nothing
    Exit Sub

eh:
    Set l_colDistinctEmpNums = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineGetSchedChgPayScope"
End Sub


Private Function GetSchedChgPayDates( _
                    ByRef r_rsSchedChgPay As ADODB.Recordset) _
                As ADODB.Recordset
'    Build a recordset with intervals for each day an schedule change payment is scheduled.
'    The purpose is to optimize the output of gab, refining a large scope to just
'    the dates of interest.
    Dim l_rsEmpDates As ADODB.Recordset

On Error GoTo eh

    Set l_rsEmpDates = New ADODB.Recordset
    With l_rsEmpDates
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "START_DATE", adDate, , adFldUpdatable
        .Fields.Append "END_DATE", adDate, , adFldUpdatable
        .Open
    End With

    r_rsSchedChgPay.MoveFirst
    Do Until r_rsSchedChgPay.EOF
        l_rsEmpDates.AddNew
        l_rsEmpDates!EMP_NUM = r_rsSchedChgPay!EMP_NUM
        l_rsEmpDates!START_DATE = r_rsSchedChgPay!CHARGE_DATE
        l_rsEmpDates!END_DATE = r_rsSchedChgPay!CHARGE_DATE
        l_rsEmpDates.Update
        r_rsSchedChgPay.MoveNext
    Loop
    r_rsSchedChgPay.MoveFirst
    
    Set GetSchedChgPayDates = l_rsEmpDates
    Set l_rsEmpDates = Nothing

    Exit Function

eh:
    Set l_rsEmpDates = Nothing
    TrsRaiseError m_strCMODULENAME, "GetSchedChgPayDates"
End Function


Private Function IsPaymentScheduledOnOffDay( _
                    ByVal v_rsSchedChgPay As ADODB.Recordset, _
                    ByVal v_rsSchedule As ADODB.Recordset) As Boolean
    Dim l_lngEmpNum As Long
    Dim l_datChgDate As Date
    Dim l_strFilter As String

On Error GoTo eh
    l_lngEmpNum = v_rsSchedChgPay!EMP_NUM
    l_datChgDate = v_rsSchedChgPay!CHARGE_DATE
    l_strFilter = "EMP_NUM = " & CStr(l_lngEmpNum) & _
        " AND CHARGE_DATE = #" & Format(l_datChgDate, "mm/dd/yyyy") & "#" & _
        " AND SHIFT_HOURS > 0"
    v_rsSchedule.Filter = MDAC25FilterFix(l_strFilter)
    IsPaymentScheduledOnOffDay = Not RecordsExist(v_rsSchedule)
    v_rsSchedule.Filter = 0
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsPaymentScheduledOnOffDay"
End Function


Private Function PreValidateSchedChgPay(ByVal v_rsSchedChgPayRequest As ADODB.Recordset, ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim bError As Boolean, bErrorWithNull As Boolean
    Dim l_strBRVMessage As String, l_strSchedChgInd As String
    Dim l_objCodeTables As dbCodeTables.cdCodeTables
    Dim l_objEmployee As dbEmployee.cdEmployee
    Dim l_strEarningsCode As String
    Dim l_strEmpType As String
    Dim bRecordsExist As Boolean

    On Error GoTo eh
    bError = False
    PreValidateSchedChgPay = False
    
    If RecordsExist(v_rsSchedChgPayRequest) Then
        bRecordsExist = True
    Else
        bRecordsExist = False
    End If
    
    If bRecordsExist Then
        Set l_objCodeTables = CreateObject("dbCodeTables.cdCodeTables")
        Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
        Set m_colEmployees = GetEmployeesInRequest(v_rsSchedChgPayRequest)
        m_aryEmployees = ArrayFromCollection(m_colEmployees)
        Set m_rsEmployee = l_objEmployee.Getemployee(m_aryEmployees)
    End If

    If bRecordsExist Then
        With v_rsSchedChgPayRequest
            .MoveFirst
                                                                          
            Do While Not .EOF
                bErrorWithNull = False
                If IsNull(!RECORD_ID) Then
                    bErrorWithNull = True
                    bError = True
                    l_strBRVMessage = "Record ID cannot be null."
                    AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                        l_strBRVMessage, m_strCMODULENAME, _
                        1, 1, "RECORD_ID"
                Else
                    If (!RECORD_STATUS = "I" Or !RECORD_STATUS = "C" Or !RECORD_STATUS = "D") Then
                        '//-----key fields are required-----//
                        If (!RECORD_STATUS = "D" Or !RECORD_STATUS = "C") Then
                            If IsNull(!KEY_CHARGE_DATE) And IsNull(!KEY_EMP_NUM) Then
                                bErrorWithNull = True
                                l_strBRVMessage = "Key charge date and key employee number cannot be null."
                                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                    l_strBRVMessage, m_strCMODULENAME, _
                                    1, !RECORD_ID, "KEY_EMP_NUM"
                            ElseIf IsNull(!KEY_CHARGE_DATE) Then
                                bErrorWithNull = True
                                l_strBRVMessage = "Key charge date cannot be null."
                                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                    l_strBRVMessage, m_strCMODULENAME, _
                                    1, !RECORD_ID, "KEY_CHARGE_DATE", _
                                    Array("KEY_EMP_NUM"), _
                                    CStr(!KEY_EMP_NUM)
                            ElseIf IsNull(!KEY_EMP_NUM) Then
                                bErrorWithNull = True
                                l_strBRVMessage = "Key employee number cannot be null."
                                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                    l_strBRVMessage, m_strCMODULENAME, _
                                    1, !RECORD_ID, "KEY_EMP_NUM", _
                                    Array("KEY_CHARGE_DATE"), _
                                    CStr(!KEY_CHARGE_DATE)
                            End If
                        End If
                        
                        '//-----non key fields are required-----//
                        If (!RECORD_STATUS = "I" Or !RECORD_STATUS = "C") Then
                            If IsNull(!CHARGE_DATE) And IsNull(!EMP_NUM) Then
                                bErrorWithNull = True
                                l_strBRVMessage = "Charge date and employee number cannot be null."
                                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                    l_strBRVMessage, m_strCMODULENAME, _
                                    1, !RECORD_ID, "EMP_NUM"
                            ElseIf IsNull(!CHARGE_DATE) Then
                                bErrorWithNull = True
                                l_strBRVMessage = "Charge date cannot be null."
                                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                    l_strBRVMessage, m_strCMODULENAME, _
                                    1, !RECORD_ID, "CHARGE_DATE", _
                                    Array("EMP_NUM"), _
                                    CStr(!EMP_NUM)
                            ElseIf IsNull(!EMP_NUM) Then
                                bErrorWithNull = True
                                l_strBRVMessage = "Employee number cannot be null."
                                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                    l_strBRVMessage, m_strCMODULENAME, _
                                    1, !RECORD_ID, "EMP_NUM", _
                                    Array("CHARGE_DATE"), _
                                    CStr(!CHARGE_DATE)
                            End If
                        End If
                                            
                        If Not bErrorWithNull Then
                            '//-----to make it easier for setting filters and BRV's (use key fields for all-----//
                            '//-----otherwise must use key fields for "C", "D" and non-key fields for "I")-----//
                            If !RECORD_STATUS = "I" Then
                                '//-----optional fields-----//
                                If IsNull(!KEY_EMP_NUM) Then
                                    !KEY_EMP_NUM = !EMP_NUM
                                End If
                                If IsNull(!KEY_CHARGE_DATE) Then
                                    !KEY_CHARGE_DATE = !CHARGE_DATE
                                End If
                                .Update
                            End If
                            If !RECORD_STATUS = "D" Then
                                '//-----optional fields-----//
                                If IsNull(!EMP_NUM) Then
                                    !EMP_NUM = !KEY_EMP_NUM
                                End If
                                If IsNull(!CHARGE_DATE) Then
                                    !CHARGE_DATE = !KEY_CHARGE_DATE
                                End If
                                .Update
                            End If
                            
                            m_rsEmployee.Filter = "EMP_NUM = " & !EMP_NUM
                                                    
                            
                            If m_rsEmployee.BOF And m_rsEmployee.EOF Then
                                bError = True
                                l_strBRVMessage = "Employee (" & !EMP_NUM & ") is not valid."
                                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                    l_strBRVMessage, m_strCMODULENAME, _
                                    1, !RECORD_ID, "EMP_NUM", _
                                    Array("EMP_NUM", "CHARGE_DATE"), _
                                    CStr(!EMP_NUM), _
                                    CStr(!CHARGE_DATE)
                            End If
            
                            If IsDate(!CHARGE_DATE) Then
                                'Schedule change payment must be within the last 30 days or into the future.
                                If CDate(Format(!CHARGE_DATE, "mm/dd/yyyy")) < CDate(Format(DateAdd("d", -30, Now()), "mm/dd/yyyy")) Then
                                    bError = True
                                    l_strBRVMessage = "Charge Date (" & !CHARGE_DATE & ") cannot be more than 30 days in the past."
                                    AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                        l_strBRVMessage, m_strCMODULENAME, _
                                        1, !RECORD_ID, "CHARGE_DATE", _
                                        Array("EMP_NUM", "CHARGE_DATE"), _
                                        CStr(!EMP_NUM), _
                                        CStr(!CHARGE_DATE)
                                End If
                            Else
                                bError = True
                                l_strBRVMessage = "Charge Date (" & !CHARGE_DATE & ") is not a valid date."
                                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                    l_strBRVMessage, m_strCMODULENAME, _
                                    1, !RECORD_ID, "CHARGE_DATE", _
                                    Array("EMP_NUM", "CHARGE_DATE"), _
                                    CStr(!EMP_NUM), _
                                    CStr(!CHARGE_DATE)
                            End If
                            
                            Select Case !RECORD_STATUS
                            Case "I", "C"
                                '//-----other required fields-----//
                                If IsNull(!EXCEPTION_CD) Then
                                    bErrorWithNull = True
                                    l_strBRVMessage = "Exception code cannot be null."
                                    AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                        l_strBRVMessage, m_strCMODULENAME, _
                                        1, !RECORD_ID, "EXCEPTION_CD", _
                                        Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                                        CStr(!KEY_EMP_NUM), _
                                        CStr(!KEY_CHARGE_DATE)
                                End If
                                
                                If IsNull(!PAYMENT_TYPE) Then
                                    bErrorWithNull = True
                                    l_strBRVMessage = "Payment type cannot be null."
                                    AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                        l_strBRVMessage, m_strCMODULENAME, _
                                        1, !RECORD_ID, "PAYMENT_TYPE", _
                                        Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                                        CStr(!KEY_EMP_NUM), _
                                        CStr(!KEY_CHARGE_DATE)
                                End If
                            
                                If Not bErrorWithNull Then
                                    l_strSchedChgInd = l_objCodeTables.ExceptionScheduleChgInd(!EXCEPTION_CD)
                                    If l_strSchedChgInd <> "S" Then
                                        bError = True
                                        l_strBRVMessage = "Invalid Exception code.  Schedule change indicator (" & l_strSchedChgInd & _
                                                        ") for Exception code (" & !EXCEPTION_CD & ") must be 'S'."
                                        AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                            l_strBRVMessage, m_strCMODULENAME, _
                                            1, !RECORD_ID, "EXCEPTION_CD", _
                                            Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                                            CStr(!KEY_EMP_NUM), _
                                            CStr(!KEY_CHARGE_DATE)
                                    Else
                                        If m_rsEmployee!EXEC_CONTRACT_IND! = "E" Then
                                            l_strEarningsCode = l_objCodeTables.ExceptionExecutiveEarningsCode(!EXCEPTION_CD)
                                            l_strEmpType = "Management"
                                        Else
                                            l_strEarningsCode = l_objCodeTables.ExceptionContractEarningsCode(!EXCEPTION_CD)
                                            l_strEmpType = "Contract"
                                        End If
                                        If Trim(l_strEarningsCode) = "" Then
                                            bError = True
                                            l_strBRVMessage = "Exception code. (" & !EXCEPTION_CD & _
                                                        ") is not valid for " & l_strEmpType & " employee."
                                            AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                                l_strBRVMessage, m_strCMODULENAME, _
                                                1, !RECORD_ID, "EXCEPTION_CD", _
                                                Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                                                CStr(!KEY_EMP_NUM), _
                                                CStr(!KEY_CHARGE_DATE)
                                        End If
                                    End If
                                    
                                    '//-----validate only when payment type='D'
                                    If !PAYMENT_TYPE = "D" Then
                                        If IsNull(!BASELINE_SCHED_CD) Or !BASELINE_SCHED_CD = 0 Then '//-----default to current schedule-----//
                                            '//-----set default is set in ValidateSchedChgPay
                                        Else
                                            If Not l_objCodeTables.ValidateScheduleCode(!BASELINE_SCHED_CD) Then
                                                bError = True
                                                l_strBRVMessage = "Schedule code " & !BASELINE_SCHED_CD & " is not valid."
                                                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                                    l_strBRVMessage, m_strCMODULENAME, _
                                                    1, !RECORD_ID, "BASELINE_SCHED_CD", _
                                                    Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                                                    CStr(!KEY_EMP_NUM), _
                                                    CStr(!KEY_CHARGE_DATE)
                                            End If
                                        End If
                                    Else    'assume PAYMENT_TYPE = "F"
                                        If Not IsNull(!BASELINE_SCHED_CD) And Not IsEmpty(!BASELINE_SCHED_CD) Then
                                            If !BASELINE_SCHED_CD <> 0 Then
                                                l_strBRVMessage = "Schedule code is ignored for full payments."
                                                AppendBusinessError r_rsErrors, trWarBaseSchedCdIgnored, mstrCBUSINESSWAR, _
                                                    l_strBRVMessage, m_strCMODULENAME, _
                                                    1, !RECORD_ID, "BASELINE_SCHED_CD", _
                                                    Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                                                    CStr(!KEY_EMP_NUM), _
                                                    CStr(!KEY_CHARGE_DATE)
                                            End If
                                        End If
                                        !BASELINE_SCHED_CD = Null
                                    End If
                                    
'                                    If Not IsNull(!BASELINE_SCHED_CD) Then
'                                        If Not l_objCodeTables.ValidateScheduleCode(!BASELINE_SCHED_CD) Then
'                                            bError = True
'                                            l_strBRVMessage = "Schedule code " & !BASELINE_SCHED_CD & " is not valid."
'                                            AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
'                                                l_strBRVMessage, m_strCMODULENAME, _
'                                                1, !RECORD_ID, "BASELINE_SCHED_CD", _
'                                                Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
'                                                CStr(!KEY_EMP_NUM), _
'                                                CStr(!KEY_CHARGE_DATE)
'                                        End If
'                                    Else '//-----may need to provide a default when null and payment type='D'
'                                    End If
                                    
                                    If Not IsPaymentTypeValid(!PAYMENT_TYPE, l_strBRVMessage) Then
                                        bError = True
                                        AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                            l_strBRVMessage, m_strCMODULENAME, _
                                            1, !RECORD_ID, "PAYMENT_TYPE", _
                                            Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                                            CStr(!KEY_EMP_NUM), _
                                            CStr(!KEY_CHARGE_DATE)
                                    End If
                                Else '//-----skip the rest of the pre-val or may get "invalid use of null"-----//
                                    bError = True
                                End If
                            Case "D"
                            End Select
                        Else '//-----skip the rest of the pre-val or may get "invalid use of null"-----//
                            bError = True
                        End If
                    Else
                        bError = True
                        l_strBRVMessage = "RECORD_STATUS must be I-insert, C-change or D-delete"
                        AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                            l_strBRVMessage, m_strCMODULENAME, _
                            1, !RECORD_ID, "RECORD_STATUS"
                    End If
                End If
                
                .MoveNext
            Loop
            Set l_objCodeTables = Nothing
            Set l_objEmployee = Nothing
        End With
        
        If Not bError Then
            '//-----check for duplicate rows in request-----//
            If DoesDuplicateExist(v_rsSchedChgPayRequest, r_rsErrors) Then
                bError = True
            Else
                bError = False
            End If
        End If
    
    End If
   
    If bError Then
        PreValidateSchedChgPay = False
        GetObjectContext.SetAbort
    Else
        PreValidateSchedChgPay = True
        GetObjectContext.SetComplete
    End If

    Exit Function
eh:
    Set l_objCodeTables = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "PreValidateSchedChgPay", , , , , True
End Function


Private Function ValidateSchedChgPay(ByVal v_rsSchedChgPayRequest As ADODB.Recordset, ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim bError As Boolean, bErrorHours As Boolean, bErrorExceptions As Boolean
    Dim l_rsScopeParms As ADODB.Recordset
    Dim l_rsScope As ADODB.Recordset
    Dim l_rsSchedule As ADODB.Recordset
    Dim l_strFilter As String, sBRVMsg As String
    Dim l_rsExceptions As ADODB.Recordset
    Dim l_strExcCd As String
    
    bError = False
    ValidateSchedChgPay = False
    On Error GoTo eh

    If RecordsExist(v_rsSchedChgPayRequest) Then
        With v_rsSchedChgPayRequest
            '//-----build summary recordset with EmpNum, StartDate, and EndDate-----//
            Set l_rsScopeParms = GetScopeParms(v_rsSchedChgPayRequest)

            '//-----get Scope-----//
            Set l_rsScope = GetScope(l_rsScopeParms)
            
            .MoveFirst
'            Do While Not .EOF
                '//-----client request swapping rows
                '//-----7/24 change to 7/25
                '//-----7/25 change to 7/24
                '//-----ProcessKeyChange produces
                '//-----7/24 D (1)
                '//-----7/25 I (2)
                '//-----7/25 D (3)
                '//-----7/24 I (4)
                '//-----ExistenceCheck complains that (2) and (4) already exist,
                '//-----even though they will be deleted by (1) and (3)
                
                '//-----Check existence of change, delete, and insert requests applied in the net effect order-----//
                If Not ApplyPendingChanges(v_rsSchedChgPayRequest, l_rsScope, r_rsErrors) Then
                    bError = True
                End If
                
'                '//-----Check existence of change and delete requests-----//
''//                If !RECORD_STATUS = "C" Or !RECORD_STATUS = "D" Then
'                    If Not ExistenceCheck(v_rsSchedChgPayRequest, l_rsScope, r_rsErrors) Then
'                        bError = True
'                    End If
''//                End If
'                .MoveNext
'            Loop

            If Not bError Then
                '//-----Check SchedChgPay against schedule for hours > 0 and flag off days (Orphan_Ind=N for hours > 0, otherwise Y)
                Set l_rsSchedule = GetSchedule(l_rsScopeParms)
                Set l_rsExceptions = GetScheduledExceptions(l_rsScopeParms)

                .MoveFirst
                Do While Not .EOF
                    If !RECORD_STATUS = "I" Or !RECORD_STATUS = "C" Then
                        '//-----Check SchedChgPay against schedule for hours > 0 and create a BRV for off days
                        bErrorHours = True
                        l_rsSchedule.Filter = adFilterNone '//-----PTS 901
                        If RecordsExist(l_rsSchedule) Then
                            l_strFilter = MDAC25FilterFix("EMP_NUM = " & !EMP_NUM & " AND CHARGE_DATE = #" & !CHARGE_DATE & "# AND SHIFT_HOURS > 0")
                            l_rsSchedule.Filter = l_strFilter
                            If RecordsExist(l_rsSchedule) Then
                                bErrorHours = False
                            End If
                        End If

                        If bErrorHours Then
                            bError = True
                            sBRVMsg = "Scheduled hours do not exist for " & !EMP_NUM & " on " & !CHARGE_DATE & "."

                            AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                sBRVMsg, m_strCMODULENAME, _
                                1, !RECORD_ID, "CHARGE_DATE", _
                                Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                                CStr(!KEY_EMP_NUM), _
                                CStr(!KEY_CHARGE_DATE)
                        End If
                        
                        '//-----Check SchedChgPay against scheduled exceptions and create a BRV for exceptions not payable
                        bErrorExceptions = False
                        l_rsExceptions.Filter = adFilterNone
                        If RecordsExist(l_rsExceptions) Then
                            l_strFilter = MDAC25FilterFix("EMP_NUM = " & !EMP_NUM & " AND CHARGE_DATE = #" & !CHARGE_DATE & "#")
                            l_rsExceptions.Filter = l_strFilter
                            If RecordsExist(l_rsExceptions) Then
                                Do Until l_rsExceptions.EOF Or bErrorExceptions
                                    bErrorExceptions = _
                                        Not IsScheduleChangePaymentPayable(l_rsExceptions!EXCEPTION_CD)
                                    If bErrorExceptions Then
                                        l_strExcCd = Trim(l_rsExceptions!EXCEPTION_CD)
                                    End If
                                    l_rsExceptions.MoveNext
                                Loop
                            End If
                        End If
                        If bErrorExceptions Then
                            bError = True
                            sBRVMsg = "Payment for " & !EMP_NUM & " is not valid with exception " _
                                        & l_strExcCd & " on " _
                                        & Format(v_rsSchedChgPayRequest!CHARGE_DATE, "mm/dd/yyyy") & "."
                            AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                sBRVMsg, m_strCMODULENAME, _
                                1, !RECORD_ID, "CHARGE_DATE", _
                                Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                                CStr(!KEY_EMP_NUM), _
                                CStr(!KEY_CHARGE_DATE)
                        End If
                        
                        '//-----validate only when payment type='D'
                        If !PAYMENT_TYPE = "D" Then
                            If IsNull(!BASELINE_SCHED_CD) Or !BASELINE_SCHED_CD = 0 Then '//-----default to current schedule-----//
                                If RecordsExist(l_rsSchedule) Then
                                    !BASELINE_SCHED_CD = l_rsSchedule!SCHED_CD
                                End If
                            End If
                        End If

                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
        
    If bError Then
        ValidateSchedChgPay = False
        GetObjectContext.SetAbort
    Else
        ValidateSchedChgPay = True
        GetObjectContext.SetComplete
    End If
        
    Set l_rsScopeParms = Nothing
    Set l_rsScope = Nothing
    Set l_rsSchedule = Nothing
    
    Exit Function
eh:
    Set l_rsScopeParms = Nothing
    Set l_rsScope = Nothing
    Set l_rsSchedule = Nothing
    ValidateSchedChgPay = False
    
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "ValidateSchedChgPay", , , , , True
End Function


Private Function PostValidateSchedChgPay(ByVal v_rsSchedChgPayRequest As ADODB.Recordset, ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim bError As Boolean

    On Error GoTo eh
'    If DoesDuplicateExist(v_rsSchedChgPayRequest, r_rsErrors) Then
'        bError = True
'    Else
        bError = False
'    End If
    
    If bError Then
        PostValidateSchedChgPay = False
        GetObjectContext.SetAbort
    Else
        PostValidateSchedChgPay = True
        GetObjectContext.SetComplete
    End If

    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "PostValidateSchedChgPay", , , , , True
End Function


'//-----borrowed from busSchedDemand-----//
'//-----check for any overlaps within the client request (duplicate requests)-----//
Private Function DoesDuplicateExist(ByVal v_rsRequest As ADODB.Recordset, ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim bError As Boolean
    Dim xx As Long
    Dim l_strFilter As String '//, l_strOverlapMsg As String
    Dim l_rsOverlap As ADODB.Recordset
    Dim l_varRequestBookmark As Variant
    
    On Error GoTo eh
    bError = False
    DoesDuplicateExist = True

    '//-----exclude deletes
    l_strFilter = "RECORD_STATUS <> 'D'"
    v_rsRequest.Filter = MDAC25FilterFix(l_strFilter)

    If RecordsExist(v_rsRequest) Then
        With v_rsRequest
            Set l_rsOverlap = .Clone
            .MoveFirst
            Do While Not .EOF
                '//-----don't allow duplicates in net result-----//
                l_varRequestBookmark = .Bookmark
                l_strFilter = "EMP_NUM = " & !EMP_NUM & _
                            " AND CHARGE_DATE = #" & !CHARGE_DATE & "#" & _
                            " AND RECORD_STATUS <> 'D'"
                If DuplicateExistFilter(l_rsOverlap, l_strFilter, l_varRequestBookmark) Then
                    bError = True
                    '//l_strOverlapMsg = "Overlap was found creating ambiguity with line " & l_rsOverlap!KEY_LINE_NUM
                    DuplicateExistBRV v_rsRequest, l_rsOverlap, "EMP_NUM", r_rsErrors
                End If

                '//-----don't allow one row to be updated twice (7/19 changed to 7/20 and 7/19 changed to 7/21)
                If !RECORD_STATUS = "C" Then
                    l_strFilter = "KEY_EMP_NUM = " & !KEY_EMP_NUM & _
                                " AND KEY_CHARGE_DATE = #" & !KEY_CHARGE_DATE & "#" & _
                                " AND RECORD_STATUS = 'C'"
                    If DuplicateExistFilter(l_rsOverlap, l_strFilter, l_varRequestBookmark) Then
                        bError = True
                        '//l_strOverlapMsg = "Overlap was found creating ambiguity with line " & l_rsOverlap!KEY_LINE_NUM
                        DuplicateExistBRV v_rsRequest, l_rsOverlap, "KEY_EMP_NUM", r_rsErrors
                    End If
                End If
                
                .MoveNext
            Loop
            
            Set l_rsOverlap = Nothing
        End With
    End If
    v_rsRequest.Filter = adFilterNone
        
    If bError Then
        DoesDuplicateExist = True
        GetObjectContext.SetAbort
    Else
        DoesDuplicateExist = False
        GetObjectContext.SetComplete
    End If

    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "DoesDuplicateExist", , , , , True
End Function

'//-----apply filter and check for any row other than the one matching the bookmark-----//
'//-----hey MS what about rs.Filter = " <column expression> And BOOKMARK <> " & bmkAssignment
Private Function DuplicateExistFilter(ByVal v_rsOverlap As ADODB.Recordset, _
                                        ByVal v_strFilter As String, _
                                        ByVal v_varRequestBookmark As Variant) As Boolean
    Dim l_bolOverlapFound As Boolean
    On Error GoTo eh
    
    l_bolOverlapFound = False
    
    With v_rsOverlap
        .Filter = MDAC25FilterFix(v_strFilter)
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF And Not l_bolOverlapFound
                If .Bookmark <> v_varRequestBookmark Then
                    l_bolOverlapFound = True
                    Exit Do
                End If
                .MoveNext
            Loop
        End If
    End With
    
    DuplicateExistFilter = l_bolOverlapFound
    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "DuplicateExistFilter", , , , , True
End Function

'//-----suppress duplicate error messages for overlap
Private Function DuplicateExistBRV(ByVal v_rsRequest As ADODB.Recordset, _
                                ByVal v_rsOverlap As ADODB.Recordset, _
                                ByVal v_strField As String, _
                                ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_strFilter As String, l_strBRVMessage As String
    
    On Error GoTo eh

    With v_rsRequest
        Do While Not v_rsOverlap.EOF
            If .Bookmark <> v_rsOverlap.Bookmark Then '//-----ignore same row
            '//LongVBErr(trErrbusSchedChgPayDuplicate) = 40102
                If v_strField = "KEY_EMP_NUM" Then
                    l_strFilter = "(ERR_NUMBER = " & LongAppErr(trErrbusSchedChgPayDuplicate) & _
                                " AND KEY_FIELD1 = " & !KEY_EMP_NUM & _
                                " AND KEY_FIELD2 = #" & !KEY_CHARGE_DATE & "#)"
                Else
                    l_strFilter = "(ERR_NUMBER = " & LongAppErr(trErrbusSchedChgPayDuplicate) & _
                                " AND KEY_FIELD1 = " & !EMP_NUM & _
                                " AND KEY_FIELD2 = #" & !CHARGE_DATE & "#)"
                End If
                
                r_rsErrors.Filter = MDAC25FilterFix(l_strFilter)
                If Not RecordsExist(r_rsErrors) Then
                    '//-----add extra fields to BRV so that duplicate BRVs can be suppressed-----//
                    If v_strField = "KEY_EMP_NUM" Then
                        l_strBRVMessage = "Duplicate was found creating ambiguity with " & !KEY_EMP_NUM & " on " & !KEY_CHARGE_DATE
                        AppendBusinessError r_rsErrors, trErrbusSchedChgPayDuplicate, mstrCBUSINESSERR, _
                            l_strBRVMessage, m_strCMODULENAME, _
                            1, !RECORD_ID, "KEY_EMP_NUM", _
                            Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                            CStr(!KEY_EMP_NUM), _
                            CStr(!KEY_CHARGE_DATE)
                    Else
                        l_strBRVMessage = "Duplicate was found creating ambiguity with " & !EMP_NUM & " on " & !CHARGE_DATE
                        AppendBusinessError r_rsErrors, trErrbusSchedChgPayDuplicate, mstrCBUSINESSERR, _
                            l_strBRVMessage, m_strCMODULENAME, _
                            1, !RECORD_ID, "EMP_NUM", _
                            Array("EMP_NUM", "CHARGE_DATE"), _
                            CStr(!EMP_NUM), _
                            CStr(!CHARGE_DATE)
                    End If
                End If
            End If
            v_rsOverlap.MoveNext
        Loop
    End With
    
    r_rsErrors.Filter = adFilterNone
    DuplicateExistBRV = True
    
    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "DuplicateExistBRV", , , , , True
End Function

'//-----the following modifies the data retrieved for validation to achieve the net effect-----//
Private Function ApplyPendingChanges(ByVal v_rsPendingChanges As ADODB.Recordset, _
                                    ByRef r_rsScopeData As ADODB.Recordset, _
                                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim f As Field
    Dim l_strFilter As String
    Dim l_blnError As Boolean
    
    On Error GoTo eh

    l_blnError = False
    
    '//-----apply changes from overlap request to data retrieved from db for net effect-----//
    If RecordsExist(v_rsPendingChanges) Then
        With v_rsPendingChanges
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'") '//-----so this is applied once-----//
            If RecordsExist(v_rsPendingChanges) Then
                .MoveFirst
                Do While Not .EOF
                    If Not ApplyPendingChange(v_rsPendingChanges, r_rsScopeData, r_rsErrors) Then
                        l_blnError = True
                    End If
                    .MoveNext
                Loop
            End If
            
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'") '//-----so this is applied once-----//
            If RecordsExist(v_rsPendingChanges) Then
                .MoveFirst
                Do While Not .EOF
                    If Not ApplyPendingChange(v_rsPendingChanges, r_rsScopeData, r_rsErrors) Then
                        l_blnError = True
                    End If
                    .MoveNext
                Loop
            End If
            
            '//-----don't expect this code to execute, because changes made to key fields produce a delete/insert
            '//-----will get here if only the RELIEF_COMP changes but the keys remain the same
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'") '//-----so this is applied once-----//
            If RecordsExist(v_rsPendingChanges) Then
                .MoveFirst
                Do While Not .EOF
                    If Not ApplyPendingChange(v_rsPendingChanges, r_rsScopeData, r_rsErrors) Then
                        l_blnError = True
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
    
    r_rsScopeData.Filter = adFilterNone
    v_rsPendingChanges.Filter = adFilterNone
    ApplyPendingChanges = Not l_blnError
    
    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "ApplyPendingChanges", , , , , True
End Function

Private Function ApplyPendingChange(ByRef r_rsPendingChange As ADODB.Recordset, _
                                    ByRef r_rsScopeData As ADODB.Recordset, _
                                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim f As Field
    Dim l_strFilter As String
    Dim l_blnError As Boolean
    Dim sBRVMsg As String
    
    On Error GoTo eh

    l_blnError = False
    
    '//-----apply changes from current row of request (overlap or assignment) to data retrieved from db for net effect-----//
    If RecordsExist(r_rsPendingChange) Then
        With r_rsPendingChange
            l_strFilter = "EMP_NUM = " & !KEY_EMP_NUM & _
                        " AND CHARGE_DATE = #" & !KEY_CHARGE_DATE & "#"
            r_rsScopeData.Filter = MDAC25FilterFix(l_strFilter)
            
            Select Case !RECORD_STATUS
            Case "D"
                If RecordsExist(r_rsScopeData) Then
'                    r_rsScopeData.MoveFirst
'                    Do While Not r_rsScopeData.EOF
                        r_rsScopeData.Delete
                        r_rsScopeData.Update
'                        r_rsScopeData.MoveNext
'                    Loop
                Else
                    sBRVMsg = "Invalid delete request." & _
                                "  Payment not found for " & !KEY_EMP_NUM & " on " & !KEY_CHARGE_DATE & "."
                                
                    AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                        sBRVMsg, m_strCMODULENAME, _
                        1, !RECORD_ID, "CHARGE_DATE", _
                        Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                        CStr(!KEY_EMP_NUM), _
                        CStr(!KEY_CHARGE_DATE)
                    l_blnError = True
                End If
'                !RECORD_STATUS = "D" '//-----so this is applied once-----//
'                .Update
            Case "I"
                '//-----fields in both recordsets need to match-----//
                If Not RecordsExist(r_rsScopeData) Then
                    r_rsScopeData.AddNew
                    CopyRequestFields r_rsPendingChange, r_rsScopeData, "I"
                    r_rsScopeData.Update
                Else
                    sBRVMsg = "Invalid insert request." & _
                                "  Payment already exists for " & !KEY_EMP_NUM & " on " & !KEY_CHARGE_DATE & "."
                                
                    AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                        sBRVMsg, m_strCMODULENAME, _
                        1, !RECORD_ID, "CHARGE_DATE", _
                        Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                        CStr(!KEY_EMP_NUM), _
                        CStr(!KEY_CHARGE_DATE)
                    l_blnError = True
                End If
'                !RECORD_STATUS = "I" '//-----so this is applied once-----//
'                .Update
            Case "C"
                If RecordsExist(r_rsScopeData) Then
                    CopyRequestFields r_rsPendingChange, r_rsScopeData, "C"
                    r_rsScopeData.Update
                Else
                    sBRVMsg = "Invalid change request." & _
                                "  Payment not found for " & !KEY_EMP_NUM & " on " & !KEY_CHARGE_DATE & "."
                                
                    AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                        sBRVMsg, m_strCMODULENAME, _
                        1, !RECORD_ID, "CHARGE_DATE", _
                        Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                        CStr(!KEY_EMP_NUM), _
                        CStr(!KEY_CHARGE_DATE)
                    l_blnError = True
                End If
'                !RECORD_STATUS = "C" '//-----so this is applied once-----//
'                .Update
            End Select
        End With
    End If
    
    ApplyPendingChange = Not l_blnError
    
    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "ApplyPendingChange", , , , , True
End Function

Private Function CopyRequestFields(ByVal rsSource As ADODB.Recordset, ByRef rsDestination As ADODB.Recordset, _
                    ByVal v_strRecordStatus As String) As Boolean
    
    On Error GoTo eh
        
    With rsDestination
        '//-----because an Update ("C") request changes to a delete/insert,
        '//----- need to carry the created columns over-----//
        '//-----also needed for conversion of 21st Watch which will preserve the last_mod columns-----//
        If Not IsNull(rsSource!CREATED_EMP_NUM) Then
            !CREATED_EMP_NUM = rsSource!CREATED_EMP_NUM
        End If
        If Not IsNull(rsSource!CREATED_TIMESTAMP) Then
            !CREATED_TIMESTAMP = rsSource!CREATED_TIMESTAMP
        End If
        If Not IsNull(rsSource!LAST_MOD_EMP_NUM) Then
            !LAST_MOD_EMP_NUM = rsSource!LAST_MOD_EMP_NUM
        End If
        If Not IsNull(rsSource!LAST_MOD_TIMESTAMP) Then
            !LAST_MOD_TIMESTAMP = rsSource!LAST_MOD_TIMESTAMP
        End If
        Select Case v_strRecordStatus
        Case "I" '//, m_strCPENDINGINSERT
            !EMP_NUM = rsSource!EMP_NUM
            !CHARGE_DATE = CDate(rsSource!CHARGE_DATE)
            If Not IsNull(rsSource!BASELINE_SCHED_CD) Then
                !BASELINE_SCHED_CD = rsSource!BASELINE_SCHED_CD
            End If
            !EXCEPTION_CD = rsSource!EXCEPTION_CD
            !PAYMENT_TYPE = rsSource!PAYMENT_TYPE
        Case "C" '//, m_strCPENDINGCHANGE
            !EMP_NUM = rsSource!EMP_NUM
            !CHARGE_DATE = CDate(rsSource!CHARGE_DATE)
            If Not IsNull(rsSource!BASELINE_SCHED_CD) Then
                !BASELINE_SCHED_CD = rsSource!BASELINE_SCHED_CD
            End If
            !EXCEPTION_CD = rsSource!EXCEPTION_CD
            !PAYMENT_TYPE = rsSource!PAYMENT_TYPE
'        Case "D" '//, m_strCPENDINGDELETE
'            !WG_NUM = rsSource!KEY_WG_NUM
'            !EMP_NUM = rsSource!KEY_EMP_NUM
'            !LINE_NUM = rsSource!KEY_LINE_NUM
'            !START_DATE = CDate(rsSource!KEY_START_DATE)
'            !END_DATE = CDate(rsSource!KEY_END_DATE)
        End Select
        
'        !RECORD_STATUS = v_strRecordStatus
    
    End With
    CopyRequestFields = True
    Exit Function
eh:
    CopyRequestFields = False
    TrsRaiseError m_strCMODULENAME, "CopyRequestFields", , , , , True
End Function

'//-----end borrowing from busSchedDemand-----//

'blue 144 sched chg ind='S'

'blue 117
Private Function IsPaymentTypeValid(ByVal v_strPaymentType As String, ByRef r_strBRVMessage As String) As Boolean
    On Error GoTo eh
    If v_strPaymentType = "F" Or v_strPaymentType = "D" Then
        r_strBRVMessage = ""
        IsPaymentTypeValid = True
    Else
        r_strBRVMessage = "Payment type can only be F-Full Shift or D-Difference."
        IsPaymentTypeValid = False
    End If
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsPaymentTypeValid", , , , , True
End Function


'//-----summarize date to get employee num, start date, and end date (for GetScope & GetSchedule)-----//
Private Function GetScopeParms(ByVal v_rsSchedChgPayData As ADODB.Recordset) As ADODB.Recordset
    Dim l_rsClone As ADODB.Recordset
    Dim l_rsGetScopeParms As ADODB.Recordset
    Dim l_lngEmpNumSave As Long
    Dim l_datStartDate As Date, l_datEndDate As Date
    Const l_lngCFIELDATTR = adFldUpdatable Or adFldIsNullable

    On Error GoTo eh
    
    If RecordsExist(v_rsSchedChgPayData) Then
        Set l_rsClone = v_rsSchedChgPayData.Clone
        With l_rsClone
            '//-----paramters for GetAdjustedBase and GetScope-----//
            Set l_rsGetScopeParms = New ADODB.Recordset
            
            With l_rsGetScopeParms
                .Fields.Append "EMP_NUM", adInteger, , l_lngCFIELDATTR
                .Fields.Append "START_DATE", adDate, , l_lngCFIELDATTR
                .Fields.Append "END_DATE", adDate, , l_lngCFIELDATTR
                .CursorLocation = adUseClient
                .CursorType = adOpenStatic
                .LockType = adLockBatchOptimistic
                .Open
            End With
            
            '//-----build list of necessary lines-----//
            .Sort = "KEY_EMP_NUM ASC, KEY_CHARGE_DATE ASC"
        
            .MoveFirst
    
            l_lngEmpNumSave = !KEY_EMP_NUM
            l_datStartDate = !KEY_CHARGE_DATE
            l_datEndDate = !KEY_CHARGE_DATE
                
            Do While Not .EOF
                '//-----check for change in wg/line
                If !KEY_EMP_NUM <> l_lngEmpNumSave Then
                    With l_rsGetScopeParms
                        .AddNew
                        !EMP_NUM = l_lngEmpNumSave
                        !START_DATE = l_datStartDate
                        !END_DATE = l_datEndDate
                        .Update
                    End With
                    
                    l_lngEmpNumSave = !KEY_EMP_NUM
                    l_datStartDate = !KEY_CHARGE_DATE
                    l_datEndDate = !KEY_CHARGE_DATE
                End If
                
                '//-----Determine scope for GetAdjustedBase-----
    '//-----first row will be the start
    '            If !KEY_CHARGE_DATE < l_datStartDate Then
    '                l_datStartDate = !KEY_CHARGE_DATE
    '            End If
    '//-----last row will be the end
    '            If !KEY_CHARGE_DATE > l_datEndDate Then
                    l_datEndDate = !KEY_CHARGE_DATE
    '            End If
                
                .MoveNext
            Loop
            
            '//-----flush out last employee-----//
            With l_rsGetScopeParms
                .AddNew
                !EMP_NUM = l_lngEmpNumSave
                !START_DATE = l_datStartDate
                !END_DATE = l_datEndDate
                .Update
            End With
        End With
        Set GetScopeParms = l_rsGetScopeParms
    Else
        Set GetScopeParms = Nothing
    End If
    
    Set l_rsClone = Nothing
    Set l_rsGetScopeParms = Nothing
    
    Exit Function
eh:
    Set l_rsClone = Nothing
    Set l_rsGetScopeParms = Nothing

    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetScopeParms", , , , , True
End Function


'//-----find minimum and maximum values in order to retrieve all scope for validation-----//
Private Function GetScope(ByVal v_rsScopeParms As ADODB.Recordset) As ADODB.Recordset
    Dim l_objSchedChgPay As dbSchedChgPay.cdSchedChgPay
    Dim l_rsScope As ADODB.Recordset

    On Error GoTo eh
    
    If RecordsExist(v_rsScopeParms) Then
        Set l_objSchedChgPay = CreateObject("dbSchedChgPay.cdSchedChgPay")
        Set l_rsScope = l_objSchedChgPay.GetScope(v_rsScopeParms)
        Set l_objSchedChgPay = Nothing
        
        Set GetScope = l_rsScope
    Else
        Set GetScope = Nothing
    End If
    
    Set l_objSchedChgPay = Nothing
    Set l_rsScope = Nothing
    
    Exit Function
eh:
    Set l_objSchedChgPay = Nothing
    Set l_rsScope = Nothing

    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetScope", , , , , True
End Function


'//-----need employee schedule for determining "ORPHAN_IND" (Get) or validation (BRV)
Private Function GetSchedule(ByVal v_rsScopeParms As ADODB.Recordset) As ADODB.Recordset
    Dim l_objSchedule As busSchedule.cbsSchedule
    Dim l_rsSchedule As ADODB.Recordset
    Dim l_rsDummy As ADODB.Recordset

    On Error GoTo eh
    
    If RecordsExist(v_rsScopeParms) Then
        Set l_objSchedule = CreateObject("busSchedule.cbsSchedule")
        Set l_rsSchedule = l_objSchedule.GABEmployeeSchedOnSpecificDates(v_rsScopeParms, l_rsDummy, _
                                            False, trBaseSchedNoDetails, trAllScheduleLevels)
        
        Set GetSchedule = l_rsSchedule
    Else
        Set GetSchedule = Nothing
    End If
    
    Set l_objSchedule = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsDummy = Nothing
    
    Exit Function
eh:
    Set l_objSchedule = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsDummy = Nothing

    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetSchedule", , , , , True
End Function


'//-----make sure that changes/deletes exist
Private Function ExistenceCheck(ByVal v_rsRequest As ADODB.Recordset, _
                                    ByVal v_rsScope As ADODB.Recordset, _
                                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_strFilter As String
    Dim sBRVMsg As String
    
    On Error GoTo eh
    ExistenceCheck = True
    
    With v_rsRequest
        l_strFilter = "EMP_NUM = " & !KEY_EMP_NUM & _
                    " AND CHARGE_DATE = #" & !KEY_CHARGE_DATE & "#"
    
        v_rsScope.Filter = MDAC25FilterFix(l_strFilter)
    
        If !RECORD_STATUS = "C" Or !RECORD_STATUS = "D" Then
            If Not RecordsExist(v_rsScope) Then
                sBRVMsg = "Invalid " & IIf(!RECORD_STATUS = "C", "change", "delete") & " request." & _
                            "  Payment not found for " & !KEY_EMP_NUM & " on " & !KEY_CHARGE_DATE & "."
                            
                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                    sBRVMsg, m_strCMODULENAME, _
                    1, !RECORD_ID, "CHARGE_DATE", _
                    Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                    CStr(!KEY_EMP_NUM), _
                    CStr(!KEY_CHARGE_DATE)
                ExistenceCheck = False
            End If
        ElseIf !RECORD_STATUS = "I" Then
            If RecordsExist(v_rsScope) Then
                sBRVMsg = "Invalid insert request." & _
                            "  Payment already exists for " & !KEY_EMP_NUM & " on " & !KEY_CHARGE_DATE & "."
                            
                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                    sBRVMsg, m_strCMODULENAME, _
                    1, !RECORD_ID, "CHARGE_DATE", _
                    Array("KEY_EMP_NUM", "KEY_CHARGE_DATE"), _
                    CStr(!KEY_EMP_NUM), _
                    CStr(!KEY_CHARGE_DATE)
                ExistenceCheck = False
            End If
        End If
        v_rsScope.Filter = adFilterNone
    End With
    
    Exit Function
eh:
    ExistenceCheck = False
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "ExistenceCheck", , , , , True
End Function


'//-----copied from busSchedDemand using AppendRecordset instead of CopyRequestFields
Private Function ProcessKeyChange(ByRef r_rsRequest As ADODB.Recordset, _
                    Optional ByVal v_varKeyColumns As Variant, Optional ByVal v_varOriginalKeyColumns As Variant) As Boolean
    Dim lngKey As Long, xx As Long
    Dim rsAppend As ADODB.Recordset

    On Error GoTo eh

    If Not (r_rsRequest Is Nothing) Then
        If IsMissing(v_varKeyColumns) Then
            v_varKeyColumns = FindKeyColumns(r_rsRequest)
            v_varOriginalKeyColumns = FindOriginalKeyColumns(v_varKeyColumns)
        ElseIf IsMissing(v_varOriginalKeyColumns) Then
            v_varOriginalKeyColumns = FindOriginalKeyColumns(v_varKeyColumns)
        End If

        If Not ValidateKeyColumns(r_rsRequest, v_varKeyColumns, v_varOriginalKeyColumns) Then
            GetObjectContext.SetAbort
            ProcessKeyChange = False
            Exit Function
        End If

        '//-----Validate Key Columns-----//
        With r_rsRequest
            If Not (.BOF And .EOF) Then
                Set rsAppend = .Clone '//-----keep separate cursors but share data-----//
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    Do While Not .EOF
                        '//-----Compare KEY_ fields-----//
                        For lngKey = LBound(v_varKeyColumns) To UBound(v_varKeyColumns)
                            If .Fields(v_varKeyColumns(lngKey)).Value <> .Fields(v_varOriginalKeyColumns(lngKey)).Value Then
                                '//-----Add "I" row and Change "C" to "D"-----//
                                AppendRecordset r_rsRequest, rsAppend, 1, r_rsRequest.AbsolutePosition
                                rsAppend!RECORD_STATUS = "I"
                                '//-----move non-key columns to key columns-----//
                                For xx = LBound(v_varKeyColumns) To UBound(v_varKeyColumns)
                                    rsAppend.Fields(v_varKeyColumns(xx)).Value = rsAppend.Fields(v_varOriginalKeyColumns(xx)).Value
                                Next
'//                                rsAppend!KEY_EMP_NUM = rsAppend!EMP_NUM
'//                                rsAppend!KEY_CHARGE_DATE = rsAppend!CHARGE_DATE
                                rsAppend.Update
                                '//-----
                                !RECORD_STATUS = "D"
                                '//-----move key columns to non-key columns-----//
                                For xx = LBound(v_varKeyColumns) To UBound(v_varKeyColumns)
                                    .Fields(v_varOriginalKeyColumns(xx)).Value = .Fields(v_varKeyColumns(xx)).Value
                                Next
'//                                !EMP_NUM = !KEY_EMP_NUM
'//                                !CHARGE_DATE = !KEY_CHARGE_DATE
                                .Update
                                
                                Exit For
                            End If
                        Next lngKey
                        .MoveNext
                    Loop
                End If
                .Filter = adFilterNone
            End If
        End With
    End If

    GetObjectContext.SetComplete
    ProcessKeyChange = True
    Set rsAppend = Nothing
    Exit Function
eh:
    Set rsAppend = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "ProcessKeyChange", , , , , True
End Function

Private Function FindKeyColumns(ByVal r_rsRequest As ADODB.Recordset) As Variant
    Dim lngKeyCount As Long, f As Field
    Dim l_varKeyColumns() As Variant

    On Error GoTo eh

    With r_rsRequest
        lngKeyCount = 0
        For Each f In .Fields
            If Left(f.Name, 4) = "KEY_" Then
                ReDim Preserve l_varKeyColumns(lngKeyCount)
                l_varKeyColumns(lngKeyCount) = f.Name
                lngKeyCount = lngKeyCount + 1
            End If
        Next f
    End With

    GetObjectContext.SetComplete
    FindKeyColumns = l_varKeyColumns
    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "FindKeyColumns", , , , , True
End Function

Private Function FindOriginalKeyColumns(ByVal v_varKeyColumns As Variant) As Variant
    Dim l_varOriginalKeyColumns() As Variant
    Dim lngKey As Long

    On Error GoTo eh

    ReDim l_varOriginalKeyColumns(LBound(v_varKeyColumns) To UBound(v_varKeyColumns))

    For lngKey = LBound(v_varKeyColumns) To UBound(v_varKeyColumns)
        l_varOriginalKeyColumns(lngKey) = Mid(v_varKeyColumns(lngKey), 5)
    Next

    GetObjectContext.SetComplete
    FindOriginalKeyColumns = l_varOriginalKeyColumns
    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "FindOriginalKeyColumns", , , , , True
End Function

Private Function ValidateKeyColumns(ByVal r_rsRequest As ADODB.Recordset, ByVal v_varKeyColumns As Variant, ByVal v_varOriginalKeyColumns As Variant) As Boolean
    Dim lngKey As Long, lngKeyCount As Long
    Dim sTest As String

    On Error GoTo eh

    If UBound(v_varKeyColumns) <> UBound(v_varOriginalKeyColumns) Then
        ValidateKeyColumns = False
        Exit Function
    End If

    sTest = r_rsRequest.Fields("RECORD_STATUS").Name

    For lngKey = LBound(v_varKeyColumns) To UBound(v_varKeyColumns)
        sTest = r_rsRequest.Fields(v_varKeyColumns(lngKey)).Name
        sTest = r_rsRequest.Fields(v_varOriginalKeyColumns(lngKey)).Name
    Next lngKey

    GetObjectContext.SetComplete
    ValidateKeyColumns = True
    Exit Function
eh:
    If Err.Number = 3265 Then '//-----field does not exist (item cannot be found)-----//
        GetObjectContext.SetAbort
        ValidateKeyColumns = False
        Exit Function
    End If

    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "ValidateKeyColumns", , , , , True
End Function


Private Function GetScheduledExceptions( _
                    ByVal v_rsScopeParms As ADODB.Recordset _
                ) As ADODB.Recordset
                
    Dim l_obj As dbSchedException.cdSchedException
    Dim l_aryEmployees As Variant
    Dim l_datStart As Date
    Dim l_datEnd As Date

On Error GoTo eh

    ExtractEmployeeDates v_rsScopeParms, l_aryEmployees, l_datStart, l_datEnd
    If ArraySize(l_aryEmployees) < 1 Then
        Set GetScheduledExceptions = Nothing
        Exit Function
    End If

    Set l_obj = CreateObject("dbSchedException.cdSchedException")
    Set GetScheduledExceptions = _
        l_obj.GetScheduledExceptions(l_aryEmployees, l_datStart, l_datEnd)
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduledExceptions"
End Function


Private Sub ExtractEmployeeDates( _
                    ByVal v_rsScopeParms As ADODB.Recordset, _
                    ByRef r_aryEmpNums As Variant, _
                    ByRef r_datStartDate As Date, _
                    ByRef r_datEndDate As Date)
' Answer an array of employee numbers; i.e. longs, in v_aryEmpNums
' Answer min StartDate and max EndDate
    Dim l_cuSet As utilTRISCommonLib.cuSet
    Dim l_datMinDate As Date
    Dim l_datMaxDate As Date

On Error GoTo eh
    l_datMinDate = #12/31/2079#
    l_datMaxDate = #1/1/1900#
    If RecordsExist(v_rsScopeParms) Then
        v_rsScopeParms.MoveFirst
        Set l_cuSet = New cuSet
        Do Until v_rsScopeParms.EOF
            l_cuSet.Add v_rsScopeParms.Fields("EMP_NUM").Value
            l_datMinDate = MinDate(l_datMinDate, v_rsScopeParms.Fields("START_DATE").Value)
            l_datMaxDate = MaxDate(l_datMaxDate, v_rsScopeParms.Fields("END_DATE").Value)
            v_rsScopeParms.MoveNext
        Loop
        r_aryEmpNums = ArrayFromCollection(l_cuSet.Contents)
        r_datStartDate = l_datMinDate
        r_datEndDate = l_datMaxDate
        v_rsScopeParms.MoveFirst
        Set l_cuSet = Nothing
    Else
        r_aryEmpNums = Array()
    End If

    Exit Sub

eh:
    Set l_cuSet = Nothing
    TrsRaiseError m_strCMODULENAME, "ExtractEmployeeDates"
End Sub

Private Function IsScheduleChangePaymentPayable( _
                    ByVal v_strExcCd As String) As Boolean

    Dim l_obj As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    IsScheduleChangePaymentPayable = l_obj.ExceptionScheduleChgInd(v_strExcCd) = "Y"
    
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsScheduleChangePaymentPayable"
End Function


'Private Function CopyRequestFields(ByVal rsSource As ADODB.Recordset, ByRef rsDestination As ADODB.Recordset, _
'                    ByVal v_strRecordStatus As String) As Boolean
'
'    On Error GoTo eh
'
'    With rsDestination
'        '//-----because an Update ("C") request changes to a delete/insert,
'        '//----- need to carry the created columns over-----//
'        '//-----also needed for conversion of 21st Watch which will preserve the last_mod columns-----//
'        If Not IsNull(rsSource!CREATED_EMP_NUM) Then
'            !CREATED_EMP_NUM = rsSource!CREATED_EMP_NUM
'        End If
'        If Not IsNull(rsSource!CREATED_TIMESTAMP) Then
'            !CREATED_TIMESTAMP = rsSource!CREATED_TIMESTAMP
'        End If
'        If Not IsNull(rsSource!LAST_MOD_EMP_NUM) Then
'            !LAST_MOD_EMP_NUM = rsSource!LAST_MOD_EMP_NUM
'        End If
'        If Not IsNull(rsSource!LAST_MOD_TIMESTAMP) Then
'            !LAST_MOD_TIMESTAMP = rsSource!LAST_MOD_TIMESTAMP
'        End If
'        If Not IsNull(rsSource!RELIEF_COMP) Then
'            !RELIEF_COMP = rsSource!RELIEF_COMP
'        End If
'        Select Case v_strRecordStatus
'        Case "I", m_strCPENDINGINSERT
'            !WG_NUM = rsSource!WG_NUM
'            !EMP_NUM = rsSource!EMP_NUM
'            !LINE_NUM = rsSource!LINE_NUM
'            !START_DATE = CDate(rsSource!START_DATE)
'            !END_DATE = CDate(rsSource!END_DATE)
'
'            '//-----even though the non-key fields are only defined, the key fields will be needed later-----//
'            !KEY_WG_NUM = rsSource!WG_NUM
'            !KEY_EMP_NUM = rsSource!EMP_NUM
'            !KEY_LINE_NUM = rsSource!LINE_NUM
'            !KEY_START_DATE = CDate(rsSource!START_DATE)
'            !KEY_END_DATE = CDate(rsSource!END_DATE)
'        Case "C", m_strCPENDINGCHANGE
'            !WG_NUM = rsSource!WG_NUM
'            !EMP_NUM = rsSource!EMP_NUM
'            !LINE_NUM = rsSource!LINE_NUM
'            !START_DATE = CDate(rsSource!START_DATE)
'            !END_DATE = CDate(rsSource!END_DATE)
'
'            !KEY_WG_NUM = rsSource!KEY_WG_NUM
'            !KEY_EMP_NUM = rsSource!KEY_EMP_NUM
'            !KEY_LINE_NUM = rsSource!KEY_LINE_NUM
'            !KEY_START_DATE = CDate(rsSource!KEY_START_DATE)
'            !KEY_END_DATE = CDate(rsSource!KEY_END_DATE)
'        Case "D", m_strCPENDINGDELETE
'            !WG_NUM = rsSource!KEY_WG_NUM
'            !EMP_NUM = rsSource!KEY_EMP_NUM
'            !LINE_NUM = rsSource!KEY_LINE_NUM
'            !START_DATE = CDate(rsSource!KEY_START_DATE)
'            !END_DATE = CDate(rsSource!KEY_END_DATE)
'
'            !KEY_WG_NUM = rsSource!KEY_WG_NUM
'            !KEY_EMP_NUM = rsSource!KEY_EMP_NUM
'            !KEY_LINE_NUM = rsSource!KEY_LINE_NUM
'            !KEY_START_DATE = CDate(rsSource!KEY_START_DATE)
'            !KEY_END_DATE = CDate(rsSource!KEY_END_DATE)
'        End Select
'
'        !RECORD_STATUS = v_strRecordStatus
'
'    End With
'    CopyRequestFields = True
'    Exit Function
'eh:
'    CopyRequestFields = False
'    TrsRaiseError m_strCMODULENAME, "CopyRequestFields", , , , , True
'End Function


'Public Function ExistenceCheckSchedChgPay _
'    (ByVal v_aryEmployees As Variant, _
'    ByVal v_dtStartDate As Date, _
'    ByVal v_dtEndDate As Date) As Boolean
'
'    On Error GoTo ExistenceCheckSchedChgPay_EH
'
'    Dim l_objSchedChgPay As dbSchedChgPay.cdSchedChgPay
'
'    ExistenceCheckSchedChgPay = False
'
'    Set l_objSchedChgPay = CreateObject("dbSchedChgPay.cdSchedChgPay")
'
'    ExistenceCheckSchedChgPay = l_objSchedChgPay.ExistenceCheckSchedChgPay _
'            (v_aryEmployees, v_dtStartDate, v_dtEndDate)
'
'    Set l_objSchedChgPay = Nothing
'
'Exit Function
'
'ExistenceCheckSchedChgPay_EH:
'    ExistenceCheckSchedChgPay = False
'    Set l_objSchedChgPay = Nothing
'    TrsRaiseError m_strCMODULENAME, "ExistenceCheckSchedChgPay ", , , , , True
'
'End Function
'
'Public Function InsertSchedChgPmts _
'    (ByVal v_adrsSchedChgPmts As ADODB.Recordset) As Boolean
'
''*****************************************************************************************
''*  Schedule Change Payments ADO
''*****************************************************************************************
''* EMP_NUM
''* SCHEDULE_DATE
''* START_TIME
''* EXCEPTION_CD
''* CHARGE_DATE
''* PAYMENT_TYPE
''* ORIG_SCHED_CODE
''* LAST_MOD_EMP
''*****************************************************************************************
'    On Error GoTo InsertSchedChgPmts_EH
'
'    Dim l_objSchedChgPay As dbSchedChgPay.cdSchedChgPayTrans
'
'    Set l_objSchedChgPay = CreateObject("dbSchedChgPay.cdSchedChgPayTrans ")
'
'    InsertSchedChgPmts = l_objSchedChgPay.InsertSchedChgPmts _
'            (v_adrsSchedChgPmts)
'
'    Set l_objSchedChgPay = Nothing
'
'    Exit Function
'
'InsertSchedChgPmts_EH:
'    InsertSchedChgPmts = False
'    Set l_objSchedChgPay = Nothing
'    TrsRaiseError m_strCMODULENAME, "InsertSchedChgPmts", , , , , True
'
'End Function
'
'Public Function DeleteSchedChgPmts _
'    (ByVal v_adrsSchedChgPmts As ADODB.Recordset) As Boolean
'
''*****************************************************************************************
''*  Delete Schedule Change Payments ADO
''*****************************************************************************************
''* EMP_NUM
''* SCHED_DATE
''* START_TIME
''*****************************************************************************************
'    On Error GoTo DeleteSchedChgPmts_EH
'
'    Dim l_objSchedChgPay As dbSchedChgPay.cdSchedChgPayTrans
'
'    Set l_objSchedChgPay = CreateObject("dbSchedChgPay.cdSchedChgPayTrans ")
'
'    DeleteSchedChgPmts = l_objSchedChgPay.DeleteSchedChgPmts(v_adrsSchedChgPmts)
'
'    Set l_objSchedChgPay = Nothing
'
'    Exit Function
'
'DeleteSchedChgPmts_EH:
'    DeleteSchedChgPmts = False
'    Set l_objSchedChgPay = Nothing
'    TrsRaiseError m_strCMODULENAME, "DeleteSchedChgPmts", , , , , True
'
'End Function
'
'Public Function DeleteSchedChgPmtsForRange _
'    (ByVal v_lEmpNum As Long, _
'    ByVal v_dtStartDate As Date, _
'    ByVal v_dtEndDate As Date) As Boolean
'
''*****************************************************************************************
''*  Delete Schedule Change Payments ADO
''*****************************************************************************************
''* EMP_NUM
''* SCHED_DATE
''* START_TIME
''*****************************************************************************************
'    On Error GoTo DeleteSchedChgPmtsForRange_EH
'
'    Dim l_objSchedChgPay As dbSchedChgPay.cdSchedChgPayTrans
'
'    Set l_objSchedChgPay = CreateObject("dbSchedChgPay.cdSchedChgPayTrans ")
'
'    DeleteSchedChgPmtsForRange = l_objSchedChgPay.DeleteSchedChgPmtsForRange _
'            (v_lEmpNum, v_dtStartDate, v_dtEndDate)
'
'    Set l_objSchedChgPay = Nothing
'
'    Exit Function
'
'DeleteSchedChgPmtsForRange_EH:
'    DeleteSchedChgPmtsForRange = False
'    Set l_objSchedChgPay = Nothing
'    TrsRaiseError m_strCMODULENAME, "DeleteSchedChgPmtsForRange", , , , , True
'
'End Function

Private Function GetEmployeesInRequest( _
    ByVal v_objRS As ADODB.Recordset _
    ) As Collection

    Dim l_colEmployees As Collection
    Dim l_strEmployee As String
    Dim l_aryEmployees As Variant

    On Error GoTo eh

    Set l_colEmployees = New Collection

    v_objRS.Sort = "EMP_NUM"
    v_objRS.MoveFirst
    l_strEmployee = ""
    While Not v_objRS.EOF
        If Trim(str(v_objRS!EMP_NUM)) <> l_strEmployee Then
            l_strEmployee = Trim(str(v_objRS!EMP_NUM))
            If Not IncludesString(l_strEmployee, l_colEmployees) Then
                l_colEmployees.Add l_strEmployee
            End If
        End If
        v_objRS.MoveNext
    Wend
    v_objRS.Sort = ""
    v_objRS.MoveFirst

    Set GetEmployeesInRequest = l_colEmployees

    Exit Function

eh:

  TrsRaiseError m_strCMODULENAME, "GetEmployeesInRequest"

End Function

