VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cboTimecard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrCMODULENAME As String = "cboTimecard"

Private mobjEmployee As boCommonLib.cboEmployee
Private mobjSchedule As cboSchedule
Private mobjTimecardEntries As ADODB.Recordset
Private mobjBSTimecard As cbsTimecard
Private mobjScheduledEntries As ADODB.Recordset
Private maryOrigOTOpportunities As Variant
Private mstrCaller As String
'for older clients (6.0.0) that are not sending this column to the server
Private m_blnPreArrangedOTIndMissing As Boolean

Private mcolTimecardEntries As Collection
Private Enum EMPRVTranEnum
    etEMPNUM = 0
    etDEPTNUM
    etEXCCD
    etCHGDATE
    etHOURS
    etTRANCD
    etLOCCD
End Enum

'*************************************************************
'    Public procedures
'*************************************************************

Public Function TimecardEntryRecordset() As ADODB.Recordset
    Set TimecardEntryRecordset = mobjTimecardEntries
End Function

Public Function ResetRecordStatus()
'Ensure all record status has been reset to restored.

    If RecordsExist(mobjTimecardEntries) Then
        mobjTimecardEntries.MoveFirst
        Do Until mobjTimecardEntries.EOF
            mobjTimecardEntries!RECORD_STATUS = " "
            mobjTimecardEntries.MoveNext
        Loop
        mobjTimecardEntries.MoveFirst
    End If

End Function


Public Property Get TimecardEntriesOnChargeDate( _
                        ByVal vChgDate As Date) As Collection
    Dim colCopy As Collection
    Dim tce As cboTimecardEntry
    
On Error GoTo eh

    Set colCopy = New Collection
    For Each tce In TimecardEntries
        If tce.ChargeDate = vChgDate Then
            colCopy.Add tce
        End If
    Next
    Set TimecardEntriesOnChargeDate = colCopy
    Exit Property
    
eh:
    TrsRaiseError mstrCMODULENAME, "TimecardEntriesOnChargeDate"

End Property

Public Property Get TimecardEntriesWithOpportunity( _
                        ByVal vlngOTOppNum As Long, _
                        Optional ByVal vblnIncludeDeleted As Boolean = False) As Collection
' Answer a collection of time entries.
    Dim colCopy As Collection
    Dim tce As cboTimecardEntry
    
On Error GoTo eh

    Set colCopy = New Collection
    For Each tce In TimecardEntriesPrivate
        If tce.OvertimeOpportunityNumber = vlngOTOppNum Then
            If vblnIncludeDeleted Then
                colCopy.Add tce
            ElseIf Not tce.IsDeleted Then
                colCopy.Add tce
            End If
        End If
    Next
    Set TimecardEntriesWithOpportunity = colCopy
    Exit Property
    
eh:
    TrsRaiseError mstrCMODULENAME, "TimecardEntriesWithOpportunity"

End Property

Public Property Get TimecardEntries() As Collection
' filter deleted entries from collection
    Dim colCopy As Collection
    Dim tce As cboTimecardEntry
    
On Error GoTo eh

    Set colCopy = New Collection
    For Each tce In TimecardEntriesPrivate
        If Not tce.IsDeleted Then
            colCopy.Add tce
        End If
    Next
    Set TimecardEntries = colCopy
    Exit Property
    
eh:
    TrsRaiseError mstrCMODULENAME, "TimecardEntries"

End Property

Public Property Get PreArrangedOTIndMissing() As Boolean
    PreArrangedOTIndMissing = m_blnPreArrangedOTIndMissing
End Property

Public Property Let PreArrangedOTIndMissing(blnPreArrangedOTIndMissing As Boolean)
    m_blnPreArrangedOTIndMissing = blnPreArrangedOTIndMissing
End Property

Public Function TimecardEntriesPrivate() As Collection
' reference to private collection, only allow friends access.

On Error GoTo eh
    
    Set TimecardEntriesPrivate = mcolTimecardEntries
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "TimecardEntriesPrivate"

End Function

Public Function Schedule() As cboSchedule
    Set Schedule = mobjSchedule
End Function

Public Function SetSchedule(sched As cboSchedule)
    Set mobjSchedule = sched
    mobjSchedule.SetTimecard SafeRef(Me)
End Function

Public Function Employee() As cboEmployee
    Set Employee = mobjEmployee
End Function

Public Function SetEmployee(emp As cboEmployee)
    Set mobjEmployee = emp
End Function

Public Sub SetBSTimecard(r_objBSTimecard As cbsTimecard)
    Set mobjBSTimecard = r_objBSTimecard
End Sub

Public Function GetBSTimecard() As cbsTimecard
    Set GetBSTimecard = mobjBSTimecard
End Function

Public Function UserEmployeeNumber() As Long
    UserEmployeeNumber = GetBSTimecard.UserEmployeeNumber
End Function

Public Function GetDepartmentManager() As dbDepartment.cdDepartment
    Set GetDepartmentManager = Me.GetBSTimecard.GetDepartmentManager
End Function

Public Function GetJobManager() As dbCodeTables.cdCodeTables
    Set GetJobManager = Me.GetBSTimecard.GetJobManager
End Function

Public Function GetCaller() As String
    GetCaller = mstrCaller
End Function

Public Function SetCaller(ByVal v_strData As String)
    mstrCaller = v_strData
End Function

Public Function SetCallerOvertime()
    mstrCaller = "O"
End Function

Public Function IsCallerOvertime() As Boolean
    IsCallerOvertime = mstrCaller = "O"
End Function

'Error set is for validation and warnings only.
Public Function ErrorSet() As ADODB.Recordset
    Set ErrorSet = GetBSTimecard.ErrorSet
End Function


Public Function RecordTimeEntry( _
            timeEntry As ADODB.Recordset) As Boolean
    Dim l_objTimeEntry As cboTimecardEntry
    Dim lcolErrDescArgs As Collection
    Dim lcolErrArgs As Collection
    Dim l_intPos As Integer
    
On Error GoTo eh
    
    Select Case timeEntry!RECORD_STATUS
    Case "I"
        Set l_objTimeEntry = AddTimeEntry(timeEntry)
        ApplyRules l_objTimeEntry
    Case "C"
        Set l_objTimeEntry = FindTimecardEntry(timeEntry)
        If Not l_objTimeEntry Is Nothing Then
            UpdateTimecardEntry l_objTimeEntry, timeEntry
            ApplyRules l_objTimeEntry, False
        Else
            Set lcolErrArgs = New Collection
            lcolErrArgs.Add "timeEntry.Filter=" & timeEntry.Filter
            'Getting all Requests up to the current Request
            l_intPos = timeEntry.AbsolutePosition
            lcolErrArgs.Add "Current Request Row = " & str(l_intPos)
            timeEntry.MoveFirst
            lcolErrArgs.Add timeEntry.GetString()
            timeEntry.AbsolutePosition = l_intPos
            Err.Raise trErrbusInvalidChangeEntryNotFound
        End If
    Case "D"
    ' delete time entry requests with sched_time_ind = Y are
    ' scheduled time only delete request, not actual time entry
    ' delete requests.
        If timeEntry!SCHED_TIME_IND = "Y" Then
            AddScheduledEntry timeEntry
        Else
            Set l_objTimeEntry = FindTimecardEntry(timeEntry)
            If Not l_objTimeEntry Is Nothing Then
                DeleteTimecardEntry l_objTimeEntry, timeEntry
            Else
                Set lcolErrArgs = New Collection
                lcolErrArgs.Add "timeEntry.Filter=" & timeEntry.Filter
    'Getting all Requests up to the current Request
                l_intPos = timeEntry.AbsolutePosition
                lcolErrArgs.Add "Current Request Row = " & str(l_intPos)
                timeEntry.MoveFirst
                lcolErrArgs.Add timeEntry.GetString()
                timeEntry.AbsolutePosition = l_intPos
                Err.Raise trErrbusInvalidDeleteEntryNotFound
            End If
        End If
    Case Else
        Set lcolErrDescArgs = New Collection
        lcolErrDescArgs.Add CStr(timeEntry!RECORD_STATUS)
        Err.Raise trErrbusInvalidRecordStatus
    End Select
    
    RecordTimeEntry = True
    Exit Function
eh:

' dont attempt to write any more errors if a run-time error occurred
' during write.
'    If Err.Number <> trErrAppendEntryError Then
'        AppendEntryError l_objErrors, v_rsTimeEntriesRequest, _
'            Err.Number, mstrCAPPLICATIONERR, Err.Description, Err.Source
'    End If
'
'    Set r_rsErrors = l_objErrors

    Select Case Err.Number
        Case trErrbusInvalidRecordStatus
            TrsRaiseError mstrCMODULENAME, "RecordTimeEntry", Err.Number, , lcolErrDescArgs, , True
        Case trErrbusInvalidChangeEntryNotFound, trErrbusInvalidDeleteEntryNotFound
            TrsRaiseError mstrCMODULENAME, "RecordTimeEntry", Err.Number, , , lcolErrArgs, True
        Case Else
            TrsRaiseError mstrCMODULENAME, "RecordTimeEntry"
    End Select

End Function

Public Function AddTimecardEntry(entry As cboTimecardEntry)
    
    mcolTimecardEntries.Add entry
    Set entry.Timecard = SafeRef(Me)
    
End Function


Public Function GetChargeDatesWithTime() As Collection
    Dim l_colChgDatesWithTime As cuSet
    Dim l_objTimecardEntry As cboTimecardEntry

On Error GoTo eh
    Set l_colChgDatesWithTime = New cuSet
    For Each l_objTimecardEntry In TimecardEntriesPrivate
        l_colChgDatesWithTime.Add l_objTimecardEntry.ChargeDate
    Next
' copy for return
    Set GetChargeDatesWithTime = l_colChgDatesWithTime.Contents
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetChargeDatesWithTime"

End Function


Public Function GetChargeDatesWithModifiedTime() As Collection
' Answer a collection of dates that represents the charge dates
' that have changed.
    Dim l_colChgDatesWithModTime As cuSet
    Dim l_objTimecardEntry As cboTimecardEntry

On Error GoTo eh
    Set l_colChgDatesWithModTime = New cuSet
    For Each l_objTimecardEntry In TimecardEntriesPrivate
        If l_objTimecardEntry.IsNew Or l_objTimecardEntry.IsModified _
            Or l_objTimecardEntry.IsDeleted Then
            l_colChgDatesWithModTime.Add l_objTimecardEntry.ChargeDate
        End If
    Next

    Set GetChargeDatesWithModifiedTime = l_colChgDatesWithModTime.Contents
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetChargeDatesWithModifiedTime"

End Function


Public Function SumHoursInSchedule(ByVal v_datChargeDate As Date) As Variant
' Only sum time within schedule, unless it is callout standby or lunch.
' May also add an optional parameter for sum option to designate exceptions
' to summing all.
' May use IsAnyPeriodWithinSchedule since all time is split on schedule.
    Dim eEntry As cboTimecardEntry
    Dim l_decHours As Variant

On Error GoTo SumHoursInSchedule_EH
    
    l_decHours = CDec(0)
    For Each eEntry In Me.TimecardEntries
        If eEntry.ChargeDate = v_datChargeDate Then
            If Schedule.IsAnyPeriodWithinSchedule(v_datChargeDate, eEntry.StartTimeAsDate, eEntry.EndTimeAsDate) Then
                If eEntry.ExceptionCode <> "LU" And _
                    eEntry.ExceptionCode <> "LP" And _
                    eEntry.ExceptionCode <> "ONP" Then
                    l_decHours = l_decHours + eEntry.Hours
                End If
            End If
        End If
    Next
    
    SumHoursInSchedule = l_decHours
    Exit Function
    
SumHoursInSchedule_EH:
    TrsRaiseError mstrCMODULENAME, "SumHoursInSchedule"

End Function

Public Sub RestoreWithState(ByRef r_rsTimecardEntries As ADODB.Recordset)

    On Error GoTo Restore_eh
    
    SetState r_rsTimecardEntries
    
    Exit Sub
    
Restore_eh:
    TrsRaiseError mstrCMODULENAME, "RestoreWithState"
    
End Sub


Public Sub Restore(ByVal v_lngEmpNum As Long, _
        Optional startDt As Variant, Optional endDt As Variant)
' startDt and endDt should be dates when provided
    Dim dbTimeCard As dbTimeCard.cdTimeCard
    Dim rsTimecard As ADODB.Recordset
    Dim stDt As Date
    Dim ndDt As Date
    Dim lcolErrDescArgs As New Collection

    On Error GoTo Restore_eh
    
    If IsMissing(startDt) Then
        DetermineStartEndDates v_lngEmpNum, stDt, ndDt
    Else
        stDt = startDt
        ndDt = endDt
    End If
    
    Set dbTimeCard = CreateObject("dbTimeCard.cdTimeCard")
    Set rsTimecard = _
        dbTimeCard.Get_TimeCard(v_lngEmpNum, stDt, ndDt)
    If rsTimecard Is Nothing Then
        lcolErrDescArgs.Add CStr(v_lngEmpNum)
        lcolErrDescArgs.Add CStr(stDt)
        lcolErrDescArgs.Add CStr(ndDt)
        Err.Raise trErrbusTimecardRestore
    Else
        SetState rsTimecard
    End If
    Exit Sub
    
Restore_eh:
    Select Case Err.Number
        Case trErrbusTimecardRestore
            TrsRaiseError mstrCMODULENAME, "Restore", Err.Number, , lcolErrDescArgs
        Case Else
            TrsRaiseError mstrCMODULENAME, "Restore"
    End Select
    
End Sub

'*************************************************************
'    Private procedures
'*************************************************************

'Private Function IsValid() As Boolean
'    IsValid = True
'End Function

Private Function AddScheduledEntry(ByRef r_rsTimeEntry As ADODB.Recordset)
' Copy the scheduled time entry to private recordset for later
' use with accrual bank manager.
    Dim l_rsScheduleEntries As ADODB.Recordset
    Dim l_fld As ADODB.Field
    Dim lstrBRVDesc As String
    Dim lstrExcCD As String
    Dim l_datChargeDate As Date
    Dim l_strOverExcCd As String
    
On Error GoTo eh

    Set l_rsScheduleEntries = ScheduledEntries
    With l_rsScheduleEntries
        .AddNew
        !EMP_NUM = r_rsTimeEntry!EMP_NUM
        !CHARGE_DATE = r_rsTimeEntry!CHARGE_DATE
        !BENEFITS_PACKAGE = Employee.EmployeeBenefitPackage
        !REC_STATUS = r_rsTimeEntry!RECORD_STATUS
        !SCHEDULED_IND = r_rsTimeEntry!SCHED_TIME_IND
        !ORIG_EXCEPTION_CD = r_rsTimeEntry!EXCEPTION_CD
        !ORIG_HOURS = r_rsTimeEntry!Hours
        !ORIG_BANK_CODE = ExceptionBankCodeFor(r_rsTimeEntry!EXCEPTION_CD)
'        !ORIG_FMLA_IND = r_rsTimeEntry!FAMILY_LEAVE_IND
        !ORIG_FMLA_IND = "N"
        If IsNull(r_rsTimeEntry!WG_NUM) Or IsEmpty(r_rsTimeEntry!WG_NUM) Then
            !WG_NUM = 0
        Else
            !WG_NUM = r_rsTimeEntry!WG_NUM
        End If
        If IsNull(r_rsTimeEntry!OVER_EXC_CD) Or IsEmpty(r_rsTimeEntry!OVER_EXC_CD) Then
            !ORIG_OVER_EXC_CD = ""
        Else
            !ORIG_OVER_EXC_CD = r_rsTimeEntry!OVER_EXC_CD
        End If
        .Update
    End With
    
    lstrExcCD = Trim(r_rsTimeEntry!EXCEPTION_CD)
    l_datChargeDate = r_rsTimeEntry!CHARGE_DATE
    If IsNull(r_rsTimeEntry!OVER_EXC_CD) Or IsEmpty(r_rsTimeEntry!OVER_EXC_CD) Then
        l_strOverExcCd = ""
    Else
        l_strOverExcCd = r_rsTimeEntry!OVER_EXC_CD
    End If
    Select Case lstrExcCD
        Case "VCO"
'            lstrBRVDesc = "Scheduled VCO (Vacation Carry Over) was deleted or changed.  Please verify that all VCO exceptions are reported and/or scheduled"
            lstrBRVDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                    & " is part of a partial week of vacation that has been carried over from the previous year. " _
                    & " This partial week of vacation cannot be further split and must be scheduled in a block."
        Case "VEF"
'            lstrBRVDesc = "Scheduled VEF (Vacation for Illness Frozen) was deleted or changed.  Please verify that all VEF exceptions are reported and/or scheduled"
            lstrBRVDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                & " is part of a third week of vacation that has been split for vacation in lieu of sick leave. " _
                & " The remaining hours of vacation in this week cannot be moved except to be taken as vacation in lieu of sick leave up to this date."
        Case Else
            lstrBRVDesc = ""
    End Select
    If lstrBRVDesc <> "" Then
        AppendError ErrorSet, _
            r_rsTimeEntry!EMP_NUM, _
            r_rsTimeEntry!CHARGE_DATE, r_rsTimeEntry!CHARGE_DATE, _
            "", "", CStr(r_rsTimeEntry!Hours), _
            lstrExcCD, l_strOverExcCd, _
            vbObjectError + 20000, mstrCBUSINESSWAR, _
                lstrBRVDesc, mstrCMODULENAME
    End If
    
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "AddScheduledEntry"
    
End Function

Public Function ScheduledEntries() As ADODB.Recordset
    If mobjScheduledEntries Is Nothing Then
        Set mobjScheduledEntries = CreateScheduledEntryContainer
    End If
    Set ScheduledEntries = mobjScheduledEntries
End Function

Private Function CreateScheduledEntryContainer() As ADODB.Recordset
' This recordset will contain scheduled only entries, i.e. do
' not exist on actual time, that are marked for deletion.
' These must be passed to accrual bank manager.
    Dim l_rsScheduleEntries As ADODB.Recordset
    
On Error GoTo eh
    ' Create the empty recordset
    Set l_rsScheduleEntries = New ADODB.Recordset
    With l_rsScheduleEntries
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "BENEFITS_PACKAGE", adChar, 4, adFldUpdatable
      .Fields.Append "REC_STATUS", adChar, 1, adFldUpdatable
      .Fields.Append "SCHEDULED_IND", adChar, 1, adFldUpdatable
      .Fields.Append "ORIG_EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "ORIG_HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("ORIG_HOURS").Precision = 3
      .Fields.Item("ORIG_HOURS").NumericScale = 1
      .Fields.Append "ORIG_BANK_CODE", adChar, 2, adFldUpdatable
      .Fields.Append "ORIG_FMLA_IND", adChar, 1, adFldUpdatable
      .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
      .Fields.Append "ORIG_OVER_EXC_CD", adChar, 3, adFldUpdatable
      .Open
    End With

    Set CreateScheduledEntryContainer = l_rsScheduleEntries
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "CreateScheduledEntryContainer"
   
End Function


Private Function ApplyRules(ByRef r_objTimeEntry As cboTimecardEntry, _
      Optional ByVal v_blnNewEntry As Boolean = True) As Boolean
'Split the time entry for schedule change payments and
' build details for the time entry passed and each time entry
' split from it.
' BLUE049
    
    Dim l_colTimeEntries As Collection
    Dim l_objTimeEntry As cboTimecardEntry
    Dim l_objTimeEntrySCP As cboTimecardEntry
    Dim l_varSchedChgPmt As Variant
    Dim splitAllEntries As Boolean
    Dim l_colSchedChgPmts As Collection
    Dim l_intCount As Integer
    Dim l_intI As Integer
    Dim l_colTimeEntries2 As Collection
    Dim l_blnSchedSplits As Boolean
    
On Error GoTo eh
    Set l_colTimeEntries = New Collection
    l_colTimeEntries.Add r_objTimeEntry
    
 'Spilt Time Entry for schedule change payments
 If v_blnNewEntry Then
 'now calling applyrules on a changed time entries to split for shift component
 'NOT doing the schedule change payment logic on a changed time entry
    If r_objTimeEntry.ScheduleChangeExceptionCode = "" Then
        Set l_colSchedChgPmts = _
            GetScheduleChangePayments( _
                r_objTimeEntry.ChargeDate, _
                r_objTimeEntry.DateWorked, _
                r_objTimeEntry.startTime, _
                r_objTimeEntry.EndTime)
        l_intCount = l_colSchedChgPmts.Count
        Set l_objTimeEntry = r_objTimeEntry
        Dim l_objNextTimeEntry As cboTimecardEntry
        For l_intI = 1 To l_intCount
            l_varSchedChgPmt = l_colSchedChgPmts(l_intI)
            Set l_objTimeEntrySCP = l_objTimeEntry.SplitComplex( _
            l_varSchedChgPmt(4), _
            l_varSchedChgPmt(5), _
            l_colTimeEntries, l_objNextTimeEntry)
            If l_objTimeEntrySCP Is Nothing Then
' 8/5/02 RSG
' SplitComplex will answer nothing when no split.  For now, assume that
' there is no split since the payment interval is the same as, or overlaps
' the time entry.  Seems we need to clean this up.
' Also assuming only 1 payment when this occurs, otherwise l_objTimeEntry will
' be set to Nothing from l_objNextTimeEntry.
' 8/12/02 RSG - Jim verified these assumptions.
                l_objTimeEntry.ScheduleChangeExceptionCode = l_varSchedChgPmt(6)
            Else
                l_objTimeEntrySCP.ScheduleChangeExceptionCode = l_varSchedChgPmt(6)
                Set l_objTimeEntry = l_objNextTimeEntry
            End If
        Next l_intI
    End If
 End If
    
    Set l_colTimeEntries2 = New Collection
    For Each l_objTimeEntry In l_colTimeEntries
        ApplyShiftComponentSplits l_objTimeEntry, l_colTimeEntries2, l_blnSchedSplits
    Next
    
    If v_blnNewEntry Or l_blnSchedSplits Then
        For Each l_objTimeEntry In l_colTimeEntries2
            BuildMissingComponents l_objTimeEntry
        Next
    End If
    
    Exit Function
    
eh:
    Select Case Err.Number
        Case trErrbusEntrySplitError
            TrsRaiseError mstrCMODULENAME, "ApplyRules", Err.Number
        Case Else
            TrsRaiseError mstrCMODULENAME, "ApplyRules"
    End Select

End Function

Private Sub BuildMissingComponents(ByRef r_objTimeEntry As cboTimecardEntry)
    
On Error GoTo eh
    r_objTimeEntry.DeterminePayrollDepartmentNumber
        
    r_objTimeEntry.DetermineUnionCode

    If r_objTimeEntry.WorkgroupNumber = 0 Then
        r_objTimeEntry.SetDefaultChargedWorkgroupNumber
    Else
        r_objTimeEntry.DetermineLocationCode
    End If
    
    If r_objTimeEntry.DistributionDepartmentNumber = "" Then
        r_objTimeEntry.DetermineDistributionDepartmentNumber
    End If
    
    If r_objTimeEntry.ExceptionCode = "" Then
        r_objTimeEntry.DetermineExceptionCode
    Else
        r_objTimeEntry.DetermineOvertimeBankChargeDate (r_objTimeEntry.ExceptionCode)
        r_objTimeEntry.DetermineEarningsCode
    End If
    
' 9/9/00 rsg this should be combined with exception code
    If r_objTimeEntry.OvertimeReasonCode = "" Then
        r_objTimeEntry.DetermineOvertimeReasonCode
    End If

' JIMB The Distribution ID is being determined in
' busDistributionMgr prior to calling busTimCard
' Commenting out line below
'    r_objTimeEntry.DetermineDistributionId


'JIMB need to place description of exception code
' in the distribution id field for fixed departments
' and for exceptions that have defualt accouting = "N/A"
    r_objTimeEntry.DetermineDistributionDescription
    
    If r_objTimeEntry.ScheduleChangeExceptionCode = "" Then
        r_objTimeEntry.DetermineScheduleChangeExceptionCode
    Else
        r_objTimeEntry.DetermineScheduleChangeEarningsCode
    End If
    
    If r_objTimeEntry.JobNumber = "" Then
        r_objTimeEntry.SetDefaultJobNumber
    Else
        r_objTimeEntry.DetermineJobUpgrade
    End If
    
    If r_objTimeEntry.TradeExceptionCode = "" Then
        r_objTimeEntry.DetermineTradeExceptionCode
    End If
    
    If r_objTimeEntry.ReliefCode = "" Then
        r_objTimeEntry.DetermineVacationReliefCode
    End If
    
    If r_objTimeEntry.ShiftComponentCode = "" Then
        r_objTimeEntry.DetermineShiftComponentCode
    Else
        r_objTimeEntry.DetermineShiftComponentAmount
    End If
    
    If r_objTimeEntry.ComponentCode = "" Then
        r_objTimeEntry.DetermineComponentCode
    Else
        r_objTimeEntry.DetermineComponentAmount
    End If
    
    r_objTimeEntry.DetermineScheduleCode

    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "BuildMissingComponents"
    
End Sub

Private Function FindTimecardEntry(timeEntry As ADODB.Recordset) As cboTimecardEntry
' timeentry should have all properties that uniquely identify a timecard entry
' employeeNumber, chargeDate, startTime, indexTime
' Answer nothing if a matching timecard entry is not found.
    Dim l_objTimecardEntry As cboTimecardEntry
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strStartTime As String
    Dim l_strIndexTime As String

On Error GoTo eh
    l_lngEmpNum = timeEntry!KEY_EMP_NUM
    l_datChargeDate = timeEntry!KEY_CHARGE_DATE
    l_strStartTime = Trim(timeEntry!KEY_START_TIME)
    l_strIndexTime = Trim(timeEntry!KEY_INDEX_TIME)

    For Each l_objTimecardEntry In TimecardEntries
        If l_objTimecardEntry.EmployeeNumber = l_lngEmpNum And _
            l_objTimecardEntry.ChargeDate = l_datChargeDate And _
            l_objTimecardEntry.startTime = l_strStartTime And _
            l_objTimecardEntry.IndexTime = l_strIndexTime Then
            Set FindTimecardEntry = l_objTimecardEntry
            Exit Function
        End If
    Next
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "FindTimecardEntry"

End Function

Private Sub UpdateTimecardEntry( _
        r_objTimecardEntry As cboTimecardEntry, _
        r_objTimeEntryRequest As ADODB.Recordset)
' check each updatable field for not Null
    Dim l_blnIgnoreREL_CD As Boolean
    Dim lstrJobNumRequest As String
    Dim l_colErrDescArgs As New Collection
    Dim lblnExceptionCodeChanged As Boolean
    Dim lstrOrigExcCd As String
    Dim lstrNewExcCd As String
    Dim lstrBRVDesc As String
    Dim l_datChargeDate As Date

On Error GoTo eh

    If Not IsNull(r_objTimeEntryRequest.Fields("DATE_WORKED")) Then
        r_objTimecardEntry.DateWorked = _
            r_objTimeEntryRequest.Fields("DATE_WORKED")
    End If

' capture these values for brv before set to null
    Dim l_varKeyFields As Variant
    l_varKeyFields = CaptureKeyFieldValues(r_objTimeEntryRequest)
    l_datChargeDate = l_varKeyFields(1)

    SetUnchangedFieldsToNull r_objTimecardEntry, r_objTimeEntryRequest
'rsg 2/27/01 hack
    If Not IsNull(r_objTimeEntryRequest.Fields("EXCEPTION_CD")) Then
        If Trim(r_objTimeEntryRequest.Fields("EXCEPTION_CD")) = _
                r_objTimecardEntry.ExceptionCode Then
            lblnExceptionCodeChanged = False
         Else
            lblnExceptionCodeChanged = True
        End If
    Else
        lblnExceptionCodeChanged = False
    End If

    r_objTimecardEntry.ScheduleTimeIndicator = _
            r_objTimeEntryRequest!SCHED_TIME_IND
            
    If Not IsNull(r_objTimeEntryRequest!OT_ADJUST_DESC) Then
        r_objTimecardEntry.OTAdjustDesc = _
                r_objTimeEntryRequest!OT_ADJUST_DESC
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("HOURS")) Then
        r_objTimecardEntry.Hours = _
            r_objTimeEntryRequest.Fields("HOURS")
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("START_TIME")) Then
        r_objTimecardEntry.startTime = _
            r_objTimeEntryRequest.Fields("START_TIME")
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("END_TIME")) Then
        r_objTimecardEntry.EndTime = _
            r_objTimeEntryRequest.Fields("END_TIME")
    End If

' Set the case number before UpdateExceptionCode because
' it will reset the case number to zero when the exception
' code is updated from a case related exception to a
' non-case related exception.  Reseting the case number to
' zero overrides any value that is sent from the client for
' case number, effectively ignoring the request.  This is
' required since the update request will have the
' original case number instead of only providing
' changes the user requested.
'
' setting to null first should remove this dependency
' 9/8/00 rsg I moved this down to ensure set to null corrected problem.
'    If Not IsNull(r_objTimeEntryRequest.Fields("CASE_NUM")) Then
'        r_objTimecardEntry.CaseNumber = _
'            r_objTimeEntryRequest.Fields("CASE_NUM")
'    End If

    If lblnExceptionCodeChanged Then
        lstrOrigExcCd = r_objTimecardEntry.ExceptionCode
        lstrNewExcCd = Trim(r_objTimeEntryRequest.Fields("EXCEPTION_CD"))
        r_objTimecardEntry.UpdateExceptionCode ( _
            r_objTimeEntryRequest.Fields("EXCEPTION_CD"))
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("OVER_EXC_CD")) Then
        r_objTimecardEntry.UpdateOverrideExceptionCode ( _
            r_objTimeEntryRequest.Fields("OVER_EXC_CD"))
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("WG_NUM")) Then
        r_objTimecardEntry.UpdateChargedWorkgroupNumber (r_objTimeEntryRequest.Fields("WG_NUM"))
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("OT_OPP_DESC")) Then
        r_objTimecardEntry.OvertimeOpportunityDescription = r_objTimeEntryRequest.Fields("OT_OPP_DESC")
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("OT_REASON_CD")) Then
        r_objTimecardEntry.OvertimeReasonCode = r_objTimeEntryRequest.Fields("OT_REASON_CD")
    End If
    
    If Not PreArrangedOTIndMissing Then
        If Not IsNull(r_objTimeEntryRequest.Fields("PREARRANGED_OT_IND")) Then
            r_objTimecardEntry.PreArrangedOTIndicator = r_objTimeEntryRequest.Fields("PREARRANGED_OT_IND")
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("OT_TRAVEL_PAY_IND")) Then
        If Trim(r_objTimeEntryRequest.Fields("OT_TRAVEL_PAY_IND")) = "Y" Then
            r_objTimecardEntry.TravelPayHours = CDec(trTravelPayHours)
        Else
            r_objTimecardEntry.TravelPayHours = CDec(0)
        End If
    End If
    
    ' reminder: when client sends value same as existing, set to null
    ' no longer eligible based on new exception code
    If Not r_objTimecardEntry.IsEligibleForJobUpgrade(Employee.UnionCode(r_objTimecardEntry.ChargeDate)) Then
        ' existing upgrade
        If r_objTimecardEntry.JobNumber <> Employee.JobNumber(r_objTimecardEntry.ChargeDate) Then
            If Not IsNull(r_objTimeEntryRequest.Fields("JOB_NUM")) Then
                lstrJobNumRequest = Trim(r_objTimeEntryRequest.Fields("JOB_NUM"))
                If lstrJobNumRequest = "" Or _
                    lstrJobNumRequest = Employee.JobNumber(r_objTimecardEntry.ChargeDate) Then
                    ' and user request is to clear it,
                    ' update with base job number, ensuring upgrade not pulled from schedule.
                    ' donot send warning
                    r_objTimecardEntry.UpdateJobNumber ( _
                        Employee.JobNumber(r_objTimecardEntry.ChargeDate))
                Else
                    ' client sent a different job upgrade, let it error
                    r_objTimecardEntry.UpdateJobNumber (lstrJobNumRequest)
                End If
            Else
                ' client send no value, update with the base job
                l_colErrDescArgs.Add r_objTimecardEntry.JobNumber
                l_colErrDescArgs.Add r_objTimecardEntry.ExceptionCode
                r_objTimecardEntry.AppendWarning trWarDiscardJobUpgrade, , l_colErrDescArgs
                r_objTimecardEntry.UpdateJobNumber ( _
                    Employee.JobNumber(r_objTimecardEntry.ChargeDate))
            End If
        Else
            If Not IsNull(r_objTimeEntryRequest.Fields("JOB_NUM")) Then
                r_objTimecardEntry.UpdateJobNumber ( _
                    r_objTimeEntryRequest.Fields("JOB_NUM"))
            End If
        End If
    Else
        If Not IsNull(r_objTimeEntryRequest.Fields("JOB_NUM")) Then
            r_objTimecardEntry.UpdateJobNumber ( _
                r_objTimeEntryRequest.Fields("JOB_NUM"))
        End If
    End If
    
    
    
    If Not IsNull(r_objTimeEntryRequest.Fields("COMP_SHIFT_CD")) Then
'Jimb 09/27/2001
'Always determine shift component for Coffeen
        If Employee.IsCoffeen(r_objTimecardEntry.ChargeDate) Then
            If Not IsDayLightSavingsDay(r_objTimecardEntry.ChargeDate) Then
                r_objTimecardEntry.DetermineShiftComponentCode
            Else
                r_objTimecardEntry.ShiftComponentCode = _
                    r_objTimeEntryRequest.Fields("COMP_SHIFT_CD")
                r_objTimecardEntry.DetermineShiftComponentAmount
            End If
        Else
            r_objTimecardEntry.ShiftComponentCode = _
                r_objTimeEntryRequest.Fields("COMP_SHIFT_CD")
            r_objTimecardEntry.DetermineShiftComponentAmount
        End If
    Else
        If lblnExceptionCodeChanged Then
            If IsWorkingException(lstrNewExcCd) And Not IsWorkingException(lstrOrigExcCd) Then
                r_objTimecardEntry.DetermineShiftComponentCode
            End If
        Else
            If Employee.IsCoffeen(r_objTimecardEntry.ChargeDate) Then
                If Not IsDayLightSavingsDay(r_objTimecardEntry.ChargeDate) Then
                    r_objTimecardEntry.DetermineShiftComponentCode
                    '//Received error - Invalid use of Null
                    '//Login as 32972 (Trina Britton)
                    '//added time for 50551 on 11/2
                    '//Change Job_num from LA1310 to LA1315 (comp shift code = C12)
'                Else
'                    r_objTimecardEntry.ShiftComponentCode = _
'                        r_objTimeEntryRequest.Fields("COMP_SHIFT_CD")
'                    r_objTimecardEntry.DetermineShiftComponentAmount
                End If
            End If
        End If
    End If
    
' Set distDeptNum after update charged workgroup since it may have
' defaulted the dist dept, but user may have changed it.
'    If Not IsNull(r_objTimeEntryRequest.Fields("DISTR_DEPT_NUM")) Then
'        r_objTimecardEntry.DistributionDepartmentNumber = _
'            r_objTimeEntryRequest.Fields("DISTR_DEPT_NUM")
'    End If
    If Not IsNull(r_objTimeEntryRequest.Fields("DISTR_DEPT_NUM")) Then
        r_objTimecardEntry.UpdateDistributionDepartementNumber ( _
            r_objTimeEntryRequest.Fields("DISTR_DEPT_NUM"))
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("FAMILY_LEAVE_IND")) Then
        r_objTimecardEntry.FamilyLeaveIndicator = _
            r_objTimeEntryRequest.Fields("FAMILY_LEAVE_IND")
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("CASE_NUM")) Then
        r_objTimecardEntry.CaseNumber = _
            r_objTimeEntryRequest.Fields("CASE_NUM")
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("SCHED_CHG_EXC_CD")) Then
        r_objTimecardEntry.UpdateScheduleChangeExceptionCode ( _
            r_objTimeEntryRequest.Fields("SCHED_CHG_EXC_CD"))
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("COMP_CD")) Then
        r_objTimecardEntry.UpdateComponentCode ( _
            r_objTimeEntryRequest.Fields("COMP_CD"))
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("REL_CD")) Then
        r_objTimecardEntry.ReliefCode = _
            r_objTimeEntryRequest.Fields("REL_CD")
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("DISTRIBUTION_TYPE")) Then
        r_objTimecardEntry.DistributionType = _
            r_objTimeEntryRequest.Fields("DISTRIBUTION_TYPE")
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("DISTRIBUTION_ID")) Then
        r_objTimecardEntry.DistributionId = _
            r_objTimeEntryRequest.Fields("DISTRIBUTION_ID")
    End If
        
    If Not IsNull(r_objTimeEntryRequest.Fields("JOB_NBR")) Then
        r_objTimecardEntry.JobNbr = _
            r_objTimeEntryRequest.Fields("JOB_NBR")
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("JOBTASK_NBR")) Then
        r_objTimecardEntry.JobTaskNbr = _
            r_objTimeEntryRequest.Fields("JOBTASK_NBR")
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("JOBTASK_STEP_NBR")) Then
        r_objTimecardEntry.JobTaskStepNbr = _
            r_objTimeEntryRequest.Fields("JOBTASK_STEP_NBR")
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("EMPRV_DATABASE_ID")) Then
        r_objTimecardEntry.EMPRVDatabaseID = _
            r_objTimeEntryRequest.Fields("EMPRV_DATABASE_ID")
            
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("APPR_IND")) Then
        r_objTimecardEntry.SetApprovalIndicator _
            r_objTimeEntryRequest.Fields("APPR_IND"), _
            UserEmployeeNumber
    End If
    
' **Remove DetermineDistributionId from above.**
' This may need to be changed in the future, questions about OO issues
' There are several fields that affect the determination of the Distribution Id
' The Distribution Id will have to be recalculated if any of the following change
' Workgroup Number, Distribution Department Number
' Exception Code, Override Exception Code

'    If Not IsNull("WG_NUM") Or _
'        Not IsNull("DISTR_DEPT_NUM") Or _
'        Not IsNull("EXCEPTION_CD") Or _
'        Not IsNull("OVER_EXC_CD") Or _
'        Not IsNull("DISTRIBUTION_ID") Then
'
'            r_objTimecardEntry.DetermineDistributionId
'
'    End If
    If Not IsNull(r_objTimeEntryRequest.Fields("EXCEPTION_CD")) Or _
        Not IsNull(r_objTimeEntryRequest.Fields("OVER_EXC_CD")) Or _
        Not IsNull(r_objTimeEntryRequest.Fields("JOB_NUM")) Then
        r_objTimecardEntry.DetermineDistributionDescription
    End If

'***DIST***
    
' 3/15/02 RSG Ensure the SCHED_CD and SCHED_SRC_CD are still
' correct after any change.
    r_objTimecardEntry.DetermineScheduleCode
    
    
    If lblnExceptionCodeChanged Then
        Select Case lstrOrigExcCd
            Case "VCO"
'                lstrBRVDesc = "VCO (Vacation Carry Over) was changed to " & lstrNewExcCd & "." _
                        & "  Please verify that all VCO exceptions are reported and/or scheduled in the same week."
                lstrBRVDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                    & " is part of a partial week of vacation that has been carried over from the previous year. " _
                    & " This partial week of vacation cannot be further split and must be scheduled in a block."
                        
            Case "VEF"
                If lstrNewExcCd = "VE" Then
                    lstrBRVDesc = ""
                Else
'                    lstrBRVDesc = "VEF (Vacation for Illness Frozen) was changed to " & lstrNewExcCd & "." _
                        & "  Please verify that all VEF exceptions are reported and/or scheduled in the same week."
                    lstrBRVDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                        & " is part of a third week of vacation that has been split for vacation in lieu of sick leave. " _
                        & " The remaining hours of vacation in this week cannot be moved except to be taken as vacation in lieu of sick leave up to this date."
                End If
            Case Else
                lstrBRVDesc = ""
        End Select
        If lstrBRVDesc <> "" Then
            AppendError ErrorSet, _
                r_objTimecardEntry.EmployeeNumber, _
                r_objTimecardEntry.ChargeDate, r_objTimecardEntry.DateWorked, _
                r_objTimecardEntry.startTime, r_objTimecardEntry.EndTime, CStr(r_objTimecardEntry.Hours), _
                r_objTimecardEntry.ExceptionCode, r_objTimecardEntry.OverrideExceptionCode, _
                vbObjectError + 20000, mstrCBUSINESSWAR, _
                    lstrBRVDesc, mstrCMODULENAME
        End If
    End If
    
    
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "UpdateTimecardEntry"
    
End Sub

Private Function CaptureKeyFieldValues( _
        r_objTimeEntryRequest As ADODB.Recordset) As Variant

On Error GoTo eh

    Dim l_varKeyFields(7) As Variant

    If Not IsNull(r_objTimeEntryRequest.Fields("EMP_NUM")) Then
        l_varKeyFields(0) = r_objTimeEntryRequest.Fields("EMP_NUM")
    Else
        l_varKeyFields(0) = 0
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("CHARGE_DATE")) Then
        l_varKeyFields(1) = r_objTimeEntryRequest.Fields("CHARGE_DATE")
    Else
        l_varKeyFields(1) = Date
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("DATE_WORKED")) Then
        l_varKeyFields(2) = r_objTimeEntryRequest.Fields("DATE_WORKED")
    Else
        l_varKeyFields(2) = Date
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("START_TIME")) Then
        l_varKeyFields(3) = r_objTimeEntryRequest.Fields("START_TIME")
    Else
        l_varKeyFields(3) = "00:00"
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("END_TIME")) Then
        l_varKeyFields(4) = r_objTimeEntryRequest.Fields("END_TIME")
    Else
        l_varKeyFields(4) = "00:00"
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("HOURS")) Then
        l_varKeyFields(5) = r_objTimeEntryRequest.Fields("HOURS")
    Else
        l_varKeyFields(5) = CDec("0.0")
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("EXCEPTION_CD")) Then
        l_varKeyFields(6) = r_objTimeEntryRequest.Fields("EXCEPTION_CD")
    Else
        l_varKeyFields(6) = ""
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("OVER_EXC_CD")) Then
        l_varKeyFields(7) = r_objTimeEntryRequest.Fields("OVER_EXC_CD")
    Else
        l_varKeyFields(7) = ""
    End If
    CaptureKeyFieldValues = l_varKeyFields

    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "CaptureKeyFieldValues"

End Function


Private Sub SetUnchangedFieldsToNull( _
        r_objTimecardEntry As cboTimecardEntry, _
        r_objTimeEntryRequest As ADODB.Recordset)

On Error GoTo eh

' ignore unchanged values

    If Not IsNull(r_objTimeEntryRequest.Fields("OT_ADJUST_DESC")) Then
        If r_objTimeEntryRequest.Fields("OT_ADJUST_DESC") = _
            r_objTimecardEntry.OTAdjustDesc Then
            r_objTimeEntryRequest.Fields("OT_ADJUST_DESC") = Null
        End If
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("HOURS")) Then
        If r_objTimeEntryRequest.Fields("HOURS") = _
            r_objTimecardEntry.Hours Then
            r_objTimeEntryRequest.Fields("HOURS") = Null
        End If
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("START_TIME")) Then
        If Trim(r_objTimeEntryRequest.Fields("START_TIME")) = _
                r_objTimecardEntry.startTime Then
            r_objTimeEntryRequest.Fields("START_TIME") = Null
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("END_TIME")) Then
        If Trim(r_objTimeEntryRequest.Fields("END_TIME")) = _
                r_objTimecardEntry.EndTime Then
            r_objTimeEntryRequest.Fields("END_TIME") = Null
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("CASE_NUM")) Then
        If r_objTimeEntryRequest.Fields("CASE_NUM") = _
                r_objTimecardEntry.CaseNumber Then
            r_objTimeEntryRequest.Fields("CASE_NUM") = Null
        End If
    End If

' distribution id is used for both distribution id and exception code desc.
' need to ignore distId when distType is spaces on both request and existing
' timecard entry
'    If Not IsNull(r_objTimeEntryRequest.Fields("DISTRIBUTION_TYPE")) Then
'        If Trim(r_objTimeEntryRequest.Fields("DISTRIBUTION_TYPE")) = "" _
'                And r_objTimecardEntry.DistributionType = "" Then
'            r_objTimeEntryRequest.Fields("DISTRIBUTION_ID") = Null
'        Else
'
'        End If
'    End If


' 2/26/01 RSG Hack
' Moved this back to UpdateTimecardEntry, but do not set the exception
' code to null on the recordset.  Changing the exception_cd here causes
' the position of the recordset to change due to an existing sort on
' exception_cd.  This is really bad since we are in the middle of
' processing the recordset sequentially.  It results in attempting to
' process the same request twice, which should result in an overlapping
' duplicate error.
'    If Not IsNull(r_objTimeEntryRequest.Fields("EXCEPTION_CD")) Then
'        If Trim(r_objTimeEntryRequest.Fields("EXCEPTION_CD")) = _
'                r_objTimecardEntry.ExceptionCode Then
'            r_objTimeEntryRequest.Fields("EXCEPTION_CD") = Null
'        End If
'    End If


    If Not IsNull(r_objTimeEntryRequest.Fields("OVER_EXC_CD")) Then
        If Trim(r_objTimeEntryRequest.Fields("OVER_EXC_CD")) = _
                r_objTimecardEntry.OverrideExceptionCode Then
            r_objTimeEntryRequest.Fields("OVER_EXC_CD") = Null
        End If
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("WG_NUM")) Then
        If r_objTimeEntryRequest.Fields("WG_NUM") = _
                r_objTimecardEntry.WorkgroupNumber Then
            r_objTimeEntryRequest.Fields("WG_NUM") = Null
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("OT_OPP_DESC")) Then
        If Trim(r_objTimeEntryRequest.Fields("OT_OPP_DESC")) = _
                r_objTimecardEntry.OvertimeOpportunityDescription Then
            r_objTimeEntryRequest.Fields("OT_OPP_DESC") = Null
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("OT_REASON_CD")) Then
        If Trim(r_objTimeEntryRequest.Fields("OT_REASON_CD")) = _
                r_objTimecardEntry.OvertimeReasonCode Then
            r_objTimeEntryRequest.Fields("OT_REASON_CD") = Null
        End If
    End If
    
    If Not PreArrangedOTIndMissing Then
        If Not IsNull(r_objTimeEntryRequest.Fields("PREARRANGED_OT_IND")) Then
            If Trim(r_objTimeEntryRequest.Fields("PREARRANGED_OT_IND")) = _
                    r_objTimecardEntry.PreArrangedOTIndicator Then
                r_objTimeEntryRequest.Fields("PREARRANGED_OT_IND") = Null
            End If
        End If
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("OT_TRAVEL_PAY_IND")) Then
        If Trim(r_objTimeEntryRequest.Fields("OT_TRAVEL_PAY_IND")) = "Y" Then
            If r_objTimecardEntry.TravelPayHours = CDec(trTravelPayHours) Then
                r_objTimeEntryRequest.Fields("OT_TRAVEL_PAY_IND") = Null
            End If
        Else
            If r_objTimecardEntry.TravelPayHours = CDec(0) Then
                r_objTimeEntryRequest.Fields("OT_TRAVEL_PAY_IND") = Null
            End If
        End If
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("JOB_NUM")) Then
        If Trim(r_objTimeEntryRequest.Fields("JOB_NUM")) = _
                r_objTimecardEntry.JobNumber Then
            r_objTimeEntryRequest.Fields("JOB_NUM") = Null
        End If
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("COMP_SHIFT_CD")) Then
        If Trim(r_objTimeEntryRequest.Fields("COMP_SHIFT_CD")) = _
                r_objTimecardEntry.ShiftComponentCode Then
            r_objTimeEntryRequest.Fields("COMP_SHIFT_CD") = Null
        End If
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("DISTR_DEPT_NUM")) Then
        If Trim(r_objTimeEntryRequest.Fields("DISTR_DEPT_NUM")) = _
                r_objTimecardEntry.DistributionDepartmentNumber Then
            r_objTimeEntryRequest.Fields("DISTR_DEPT_NUM") = Null
        End If
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("FAMILY_LEAVE_IND")) Then
        If Trim(r_objTimeEntryRequest.Fields("FAMILY_LEAVE_IND")) = _
                r_objTimecardEntry.FamilyLeaveIndicator Then
            r_objTimeEntryRequest.Fields("FAMILY_LEAVE_IND") = Null
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("SCHED_CHG_EXC_CD")) Then
        If Trim(r_objTimeEntryRequest.Fields("SCHED_CHG_EXC_CD")) = _
                r_objTimecardEntry.ScheduleChangeExceptionCode Then
            r_objTimeEntryRequest.Fields("SCHED_CHG_EXC_CD") = Null
        End If
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("COMP_CD")) Then
        If Trim(r_objTimeEntryRequest.Fields("COMP_CD")) = _
                r_objTimecardEntry.ComponentCode Then
            r_objTimeEntryRequest.Fields("COMP_CD") = Null
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("REL_CD")) Then
        If Trim(r_objTimeEntryRequest.Fields("REL_CD")) = _
                r_objTimecardEntry.ReliefCode Then
            r_objTimeEntryRequest.Fields("REL_CD") = Null
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("DISTRIBUTION_TYPE")) Then
        If Trim(r_objTimeEntryRequest.Fields("DISTRIBUTION_TYPE")) = _
                r_objTimecardEntry.DistributionType Then
            r_objTimeEntryRequest.Fields("DISTRIBUTION_TYPE") = Null
        End If
    End If
    
'RSG 8/1/01 Use RTrim instead of Trim on DISTRIBUTION_ID since it may
'have leading spaces that could be lost.
'Currently DataManager applies RTrim to string fields so this should use
'the same function and/or right pad with spaces to equal length before compare.
'When DataManager is removed, do not trim on restoring the timecardEntry
'nor here.
    If Not IsNull(r_objTimeEntryRequest.Fields("DISTRIBUTION_ID")) Then
        If RTrim(r_objTimeEntryRequest.Fields("DISTRIBUTION_ID")) = _
                r_objTimecardEntry.DistributionId Then
            r_objTimeEntryRequest.Fields("DISTRIBUTION_ID") = Null
        End If
    End If

    If Not IsNull(r_objTimeEntryRequest.Fields("APPR_IND")) Then
        If Trim(r_objTimeEntryRequest.Fields("APPR_IND")) = _
                r_objTimecardEntry.ApprovalIndicator Then
            r_objTimeEntryRequest.Fields("APPR_IND") = Null
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("JOB_NBR")) Then
        If Trim(r_objTimeEntryRequest.Fields("JOB_NBR")) = _
                r_objTimecardEntry.JobNbr Then
            r_objTimeEntryRequest.Fields("JOB_NBR") = Null
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("JOBTASK_NBR")) Then
        If Trim(r_objTimeEntryRequest.Fields("JOBTASK_NBR")) = _
                r_objTimecardEntry.JobTaskNbr Then
            r_objTimeEntryRequest.Fields("JOBTASK_NBR") = Null
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("JOBTASK_STEP_NBR")) Then
        If r_objTimeEntryRequest.Fields("JOBTASK_STEP_NBR") = _
                r_objTimecardEntry.JobTaskNbr Then
            r_objTimeEntryRequest.Fields("JOBTASK_STEP_NBR") = Null
        End If
    End If
    
    If Not IsNull(r_objTimeEntryRequest.Fields("EMPRV_DATABASE_ID")) Then
        If Trim(r_objTimeEntryRequest.Fields("EMPRV_DATABASE_ID")) = _
                r_objTimecardEntry.EMPRVDatabaseID Then
            r_objTimeEntryRequest.Fields("EMPRV_DATABASE_ID") = Null
        End If
    End If
    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "SetUnchangedFieldsToNull"
    
End Sub

'Private Function ApprovedByEmployeeNumber(v_varData As Variant) As Long
'
'    If IsNull(v_varData) Then
'        ApprovedByEmployeeNumber = 0
'    Else
'        ApprovedByEmployeeNumber = CLng(v_varData)
'    End If
'
'End Function

Private Sub DeleteTimecardEntry( _
                r_objTimecardEntry As cboTimecardEntry, _
                r_rsTimeEntry As ADODB.Recordset)

Dim lstrBRVDesc As String
Dim lstrExcCD As String
Dim l_datChargeDate As Date

On Error GoTo eh
    r_objTimecardEntry.ScheduleTimeIndicator = _
            r_rsTimeEntry!SCHED_TIME_IND
            
    If Not IsNull(r_rsTimeEntry!OT_ADJUST_DESC) Then
        r_objTimecardEntry.OTAdjustDesc = _
                r_rsTimeEntry!OT_ADJUST_DESC
    End If
    
    lstrExcCD = r_objTimecardEntry.ExceptionCode
    l_datChargeDate = r_objTimecardEntry.ChargeDate
    Select Case lstrExcCD
        Case "VCO"
'            lstrBRVDesc = "VCO (Vacation Carry Over) was deleted.  Please verify that all VCO exceptions are reported and/or scheduled"
            lstrBRVDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                & " is part of a partial week of vacation that has been carried over from the previous year. " _
                & " This partial week of vacation cannot be further split and must be scheduled in a block."
        Case "VEF"
'            lstrBRVDesc = "VEF (Vacation for Illness Frozen) was deleted.  Please verify that all VEF exceptions are reported and/or scheduled"
            lstrBRVDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                & " is part of a third week of vacation that has been split for vacation in lieu of sick leave. " _
                & " The remaining hours of vacation in this week cannot be moved except to be taken as vacation in lieu of sick leave up to this date."
        Case Else
            lstrBRVDesc = ""
    End Select
    If lstrBRVDesc <> "" Then
        AppendError ErrorSet, _
            r_objTimecardEntry.EmployeeNumber, _
            r_objTimecardEntry.ChargeDate, r_objTimecardEntry.DateWorked, _
            r_objTimecardEntry.startTime, r_objTimecardEntry.EndTime, CStr(r_objTimecardEntry.Hours), _
            r_objTimecardEntry.ExceptionCode, r_objTimecardEntry.OverrideExceptionCode, _
            vbObjectError + 20000, mstrCBUSINESSWAR, _
                lstrBRVDesc, mstrCMODULENAME
    End If
            
    r_objTimecardEntry.Delete
    Exit Sub

eh:
    TrsRaiseError mstrCMODULENAME, "DeleteTimecardEntry"

End Sub

Private Function AddTimeEntry(newRequest As ADODB.Recordset) As cboTimecardEntry
    Dim te As cboTimecardEntry
    Dim l_strExcCd As String
    
On Error GoTo eh
    Set te = CreateObject("busTimecard.cboTimecardEntry")
    te.SetNew
    
'    These are required for new entry
    te.EmployeeNumber = newRequest!EMP_NUM
    te.ChargeDate = newRequest!CHARGE_DATE
    te.DateWorked = newRequest!DATE_WORKED
    te.startTime = newRequest!START_TIME
    te.EndTime = newRequest!END_TIME
    te.Hours = newRequest!Hours
    te.ScheduleTimeIndicator = newRequest!SCHED_TIME_IND
    
    Me.AddTimecardEntry te
    
'    Optional fields will be be null when not provided
    te.OvertimeBankChargeDate = CDate(CDEFAULTDATE)

'RSG - added for equalized overtime entry
    If FieldExists("OT_OPP_NUM", newRequest) Then
        If Not IsNull(newRequest!OT_OPP_NUM) Then
            te.OvertimeOpportunityNumber = newRequest!OT_OPP_NUM
        Else
            te.OvertimeOpportunityNumber = 0
        End If
    End If

'The DetermineStraghtTimeExceptionCode is dependent on DistributionDepartmentNumber
'Currently the client is always passing it in the request
'The client must continue to do this or
' clear exception code upon new entry, if the exception code is ST

    If Not IsNull(newRequest!DISTR_DEPT_NUM) Then
        te.DistributionDepartmentNumber = newRequest!DISTR_DEPT_NUM
    End If

    If Not IsNull(newRequest!EXCEPTION_CD) Then
        l_strExcCd = Trim(newRequest!EXCEPTION_CD)
        Select Case l_strExcCd
            Case "ST"
                If te.IsRegularScheduledHours Then
                    l_strExcCd = te.DetermineStraightTimeExceptionCode
                    te.UpdateExceptionCode (l_strExcCd)
                Else
                End If
            Case "ONP"
                If te.IsRegularScheduledHours Then
                Else
                    te.UpdateExceptionCode (newRequest!EXCEPTION_CD)
                End If
            
            Case Else
                te.UpdateExceptionCode (newRequest!EXCEPTION_CD)
        End Select
    End If
    
    If Not IsNull(newRequest!OVER_EXC_CD) Then
        te.OverrideExceptionCode = newRequest!OVER_EXC_CD
    End If
    
    If Not IsNull(newRequest!WG_NUM) Then
        te.WorkgroupNumber = newRequest!WG_NUM
    End If
    
' this needs to get travel pay hours from sysparms,
' just use constant for now
    If Not IsNull(newRequest!OT_TRAVEL_PAY_IND) Then
        If Trim(newRequest!OT_TRAVEL_PAY_IND) = "Y" Then
            te.TravelPayHours = CDec(trTravelPayHours)
        End If
    End If
    
    If Not IsNull(newRequest!OT_OPP_DESC) Then
        te.OvertimeOpportunityDescription = newRequest!OT_OPP_DESC
    End If
    
    If Not IsNull(newRequest!OT_REASON_CD) Then
        te.OvertimeReasonCode = newRequest!OT_REASON_CD
    End If
    
    If Not PreArrangedOTIndMissing Then
        If Not IsNull(newRequest!PREARRANGED_OT_IND) Then
            te.PreArrangedOTIndicator = newRequest!PREARRANGED_OT_IND
        Else 'the timecard entry was getting initialized with empty string rather than null
            te.PreArrangedOTIndicator = "N" 'would rather let the db assign
        End If
    Else
        te.PreArrangedOTIndicator = "N" 'would rather let the db assign
    End If
    
'
' not a time entry attribute, just pass to busOvertimeEqualizer
    If Not IsNull(newRequest!OT_ADJUST_DESC) Then
        te.OvertimeOpportunityDescription = newRequest!OT_ADJUST_DESC
    End If
    
    If Not IsNull(newRequest!JOB_NUM) Then
        te.JobNumber = newRequest!JOB_NUM
    End If

    If Not IsNull(newRequest!COMP_SHIFT_CD) Then
    'Jimb begin 09/27/2001
'Always determine shift component for Coffeen
    If Employee.IsCoffeen(te.ChargeDate) Then
        te.ShiftComponentCode = ""
    Else
        If Employee.IsAmerenCILCOLocal8Employee(te.ChargeDate) Then
            If Trim(newRequest!COMP_SHIFT_CD) = "R" Then
                te.ShiftComponentCode = newRequest!COMP_SHIFT_CD
            Else
'Always determine shift component for Local 8 non rotating
                te.ShiftComponentCode = ""
            End If
        Else
            te.ShiftComponentCode = newRequest!COMP_SHIFT_CD
        End If
    End If
'Jimb end 09/27/2001
    End If

    If IsNull(newRequest!FAMILY_LEAVE_IND) Then
        te.FamilyLeaveIndicator = "N"
    Else
        te.FamilyLeaveIndicator = newRequest!FAMILY_LEAVE_IND
    End If
    
    If Not IsNull(newRequest!CASE_NUM) Then
        te.CaseNumber = newRequest!CASE_NUM
    End If
    
    If Not IsNull(newRequest!SCHED_CHG_EXC_CD) Then
        te.ScheduleChangeExceptionCode = newRequest!SCHED_CHG_EXC_CD
    End If
    
    If Not IsNull(newRequest!DISTRIBUTION_ID) Then
        te.DistributionId = newRequest!DISTRIBUTION_ID
    End If
    
    If Not IsNull(newRequest!DISTRIBUTION_TYPE) Then
        te.DistributionType = newRequest!DISTRIBUTION_TYPE
    End If
    
    If Not IsNull(newRequest!JOB_NBR) Then
        te.JobNbr = newRequest!JOB_NBR
    End If
    
    If Not IsNull(newRequest!JOBTASK_NBR) Then
        te.JobTaskNbr = newRequest!JOBTASK_NBR
    End If
    
    If Not IsNull(newRequest!JOBTASK_STEP_NBR) Then
        te.JobTaskStepNbr = newRequest!JOBTASK_STEP_NBR
    End If
    
    If Not IsNull(newRequest!EMPRV_DATABASE_ID) Then
        te.EMPRVDatabaseID = newRequest!EMPRV_DATABASE_ID
    End If

    If Not IsNull(newRequest!COMP_CD) Then
        te.ComponentCode = newRequest!COMP_CD
    End If

    If Not IsNull(newRequest!REL_CD) Then
        te.ReliefCode = newRequest!REL_CD
    End If

    If Not IsNull(newRequest!APPR_IND) Then
        te.ApprovalIndicator = newRequest!APPR_IND
    Else
        te.ApprovalIndicator = "N"
    End If
    
    If te.ApprovalIndicator = "Y" Then
        te.ApprovedDate = Date
        te.ApprovedBy = UserEmployeeNumber
    Else
        te.ApprovedDate = CDate(CDEFAULTDATE)
        te.ApprovedBy = 0
    End If
    
    If Not IsNull(newRequest!TIME_STAMP) Then
        te.GroupTimestamp = newRequest!TIME_STAMP
    Else
        te.GroupTimestamp = ""
    End If
    
    Set AddTimeEntry = te
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "AddTimeEntry"

End Function

Private Function DetermineStartEndDates(ByVal v_lngEmpNum As Long, _
        ByRef startDt As Date, ByRef endDt As Date)
' get the collection period start and end date based on the
' current date.

    GetCollectionPeriod _
        v_lngEmpNum, _
        Date, _
        startDt, _
        endDt

End Function

Private Sub GetCollectionPeriod( _
                ByVal v_lngEmployeeNumber As Long, _
                ByVal v_datChargeDate As Date, _
                ByRef r_datCollPerStart As Date, _
                ByRef r_datCollPerEnd As Date)
    Dim l_objCollPeriod As busCollectionPeriod.cbsCollectionPeriod
    Dim l_rsCollPeriods As ADODB.Recordset
    Dim lcolErrDescArgs As New Collection
    
On Error GoTo eh
    Set l_objCollPeriod = _
        CreateObject("busCollectionPeriod.cbsCollectionPeriod")
    Set l_rsCollPeriods = _
        l_objCollPeriod.GetCollectionPeriodStatusByEmployee( _
                            v_lngEmployeeNumber, v_datChargeDate, 1)
    
'    this error handling needs work - write a brv
    If l_rsCollPeriods.BOF And l_rsCollPeriods.EOF Then
        lcolErrDescArgs.Add CStr(v_lngEmployeeNumber)
        lcolErrDescArgs.Add CStr(v_datChargeDate)
        Err.Raise trErrbusCollectionPeriodNotFound
    Else
        r_datCollPerStart = l_rsCollPeriods!PAY_PER_START_DATE
        r_datCollPerEnd = l_rsCollPeriods!PAY_PER_END_DATE
    End If
    Exit Sub

eh:
    Select Case Err.Number
        Case trErrbusCollectionPeriodNotFound
            TrsRaiseError mstrCMODULENAME, "GetCollectionPeriod", Err.Number, , lcolErrDescArgs
        Case Else
            TrsRaiseError mstrCMODULENAME, "GetCollectionPeriod"
    End Select
    
End Sub

Public Sub Release()

'Donot enable an error handler since this is called from an error handler.
    
    Dim tce As cboTimecardEntry

    Set mobjBSTimecard = Nothing
    If Not mobjEmployee Is Nothing Then
        mobjEmployee.Release
        Set mobjEmployee = Nothing
    End If
    
    If Not mobjSchedule Is Nothing Then
        mobjSchedule.Release
        Set mobjSchedule = Nothing
    End If
    
    For Each tce In mcolTimecardEntries
        tce.Release
    Next
    Set tce = Nothing
    Set mcolTimecardEntries = Nothing
    
    Set mobjTimecardEntries = Nothing
    Set mobjScheduledEntries = Nothing
    GetObjectContext.SetComplete
    
End Sub

Private Function DefaultStartDate() As Date
    DefaultStartDate = WorkWeekBeginDateFor(Date)
End Function

Private Function DefaultEndDate(stDate As Date) As Date
    DefaultEndDate = WorkWeekEndDateFor(stDate)
End Function

Public Function SetState(rs As ADODB.Recordset)
    Dim restoredEntry As cboTimecardEntry
    
On Error GoTo eh
    Set mobjTimecardEntries = rs
    
    'create collection of cboTimecardEntries
    Set mcolTimecardEntries = New Collection
    
' REMOVE DATAMANAGER
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set restoredEntry = CreateObject("busTimecard.cboTimecardEntry")
            
            'for older clients (6.0.0) that are not sending this column to the server
            restoredEntry.PreArrangedOTIndMissing = PreArrangedOTIndMissing

            restoredEntry.Restore rs
            AddTimecardEntry restoredEntry
            rs.MoveNext
        Loop
    End If
       
    SetOriginalOvertimeOpportunities
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "SetState"
    
End Function

Private Function SetOriginalOvertimeOpportunities()
' set canvass and non canvass opportunity details.
' maryOrigOTOpportunities is a variant array that contains 1 to n
' arrays of variants.  each array of variants contains:
' otOppNum, startTime, endTime, totalHours
' This is used later for validation of attempts to update
' canvass/noncanvass opportunity overtime.

On Error GoTo eh
    maryOrigOTOpportunities = GetAllOvertimeOpportunityDetails
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "SetOriginalOvertimeOpportunities"
    
End Function

Public Function GetOriginalOvertimeOpportunities() As Variant
' maryOrigOTOpportunities is a variant array that contains 1 to n
' arrays of variants.  each array of variants contains:
' otOppNum, startTime, endTime, totalHours

On Error GoTo eh
    GetOriginalOvertimeOpportunities = maryOrigOTOpportunities
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "GetOriginalOvertimeOpportunities"
    
End Function

Public Function GetAllOvertimeOpportunityDetails() As Variant
' canvass and non canvass
' answer array for each canvass/noncanvass opportunity with following:
' otOppNum, startTime, endTime, totalHours
    Dim lcolOTOppNums As Collection
    Dim eOTOppNum As Variant
    Dim lvarOTOppNumDetails As Variant
    Dim laryOTOppNumDetails() As Variant
    Dim lintOTOppIndex As Integer

On Error GoTo eh
    Set lcolOTOppNums = GetOvertimeOpportunityNumbers
    If lcolOTOppNums.Count = 0 Then
        GetAllOvertimeOpportunityDetails = Empty
        Exit Function
    End If
    
    For Each eOTOppNum In lcolOTOppNums
        lvarOTOppNumDetails = GetOvertimeOpportunityDetails(CLng(eOTOppNum))
        If Not IsEmpty(lvarOTOppNumDetails) Then
            lintOTOppIndex = lintOTOppIndex + 1
            ReDim Preserve laryOTOppNumDetails(1 To lintOTOppIndex)
            laryOTOppNumDetails(lintOTOppIndex) = lvarOTOppNumDetails
        End If
    Next
    
    If lintOTOppIndex > 0 Then
        GetAllOvertimeOpportunityDetails = laryOTOppNumDetails
    Else
        GetAllOvertimeOpportunityDetails = Empty
    End If
    Set lcolOTOppNums = Nothing
    Exit Function
    
eh:
    Set lcolOTOppNums = Nothing
    TrsRaiseError mstrCMODULENAME, "GetAllOvertimeOpportunityDetails"
    
End Function

Public Function GetOvertimeOpportunityNumbers() As Collection
' Answer a collection of longs that represents the distinct
' canvass and non canvass opportunity numbers referenced in my
' set of timecard entries.
    Dim l_colOTOppNums As cuSet
    Dim l_objTimecardEntry As cboTimecardEntry

On Error GoTo eh
    Set l_colOTOppNums = New cuSet
    For Each l_objTimecardEntry In TimecardEntriesPrivate
        If l_objTimecardEntry.IsCanvassNonCanvass Then
            l_colOTOppNums.Add l_objTimecardEntry.OvertimeOpportunityNumber
        End If
    Next
    Set GetOvertimeOpportunityNumbers = l_colOTOppNums.Contents
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetOvertimeOpportunityNumbers"

End Function

Public Function GetOvertimeOpportunityDetails( _
            ByVal vlngOTOppNum As Long) As Variant
' Answer a variant array that has the following details:
' otOppNum, earliest startTime, latest endTime, total hours
' If not canvass/noncanvass or no details are available for
' vlngOTOppNum, answer Empty.
    Dim eEntry As cboTimecardEntry
    Dim laryOTOppDetails(1 To 4) As Variant
    Dim ldatStartTime As Date
    Dim ldatEndTime As Date
    Dim ldecHours As Variant
    Dim ldatEntryTime As Date

On Error GoTo eh
    
    If vlngOTOppNum = 0 Or vlngOTOppNum >= trLegacyMinDefaultOTOppNum Then
        GetOvertimeOpportunityDetails = Empty
        Exit Function
    End If
    
    ldecHours = CDec(0)
    ldatStartTime = CDate("12/31/2100")
    ldatEndTime = CDate("01/01/0100")
    
    For Each eEntry In Me.TimecardEntries
        If eEntry.OvertimeOpportunityNumber = vlngOTOppNum Then
            ldatEntryTime = eEntry.StartTimeAsDate
            If ldatEntryTime < ldatStartTime Then
                ldatStartTime = ldatEntryTime
            End If
            ldatEntryTime = eEntry.EndTimeAsDate
            If ldatEntryTime > ldatEndTime Then
                ldatEndTime = ldatEntryTime
            End If
            ldecHours = ldecHours + eEntry.Hours
        End If
    Next
    
    laryOTOppDetails(1) = vlngOTOppNum
    laryOTOppDetails(2) = ldatStartTime
    laryOTOppDetails(3) = ldatEndTime
    laryOTOppDetails(4) = ldecHours
    If ldecHours > 0 Then
        GetOvertimeOpportunityDetails = laryOTOppDetails
    Else
        GetOvertimeOpportunityDetails = Empty
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "GetOvertimeOpportunityDetails"

End Function


Public Function BuildAccrualBankUpdates() As ADODB.Recordset
'BLUE023
    Dim l_strBANKCODE As String
    Dim l_strOriginalEXCEPTION As String
    Dim l_varOriginalHours As Variant
    Dim l_strOriginalBANKCODE As String
    Dim l_strOriginalFMLAIND As String
    Dim l_strEmployeeBenefitPackage As String
    Dim l_rsRequest As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim l_strOriginalOVEREXCCD As String

On Error GoTo eh
    Set l_rsRequest = CreateInterfaceRecordset()
    Set rs = TimecardEntryRecordset
    l_strEmployeeBenefitPackage = Employee.EmployeeBenefitPackage
    rs.Filter = MDAC25FilterFix("RECORD_STATUS = 'D' OR RECORD_STATUS = 'I' OR RECORD_STATUS = 'C'")
    Do Until rs.EOF
       l_strBANKCODE = ExceptionBankCodeFor(Trim(rs!EXCEPTION_CD.Value))
       
        If IsEmpty(rs!OVER_EXC_CD.OriginalValue) Then
          l_strOriginalOVEREXCCD = ""
       Else
          l_strOriginalOVEREXCCD = rs!OVER_EXC_CD.OriginalValue
       End If

       If IsEmpty(rs!EXCEPTION_CD.OriginalValue) Then
          l_strOriginalEXCEPTION = ""
          l_strOriginalBANKCODE = ""
       Else
          l_strOriginalEXCEPTION = rs!EXCEPTION_CD.OriginalValue
          l_strOriginalBANKCODE = ExceptionBankCodeFor(Trim(l_strOriginalEXCEPTION))
       End If

       If IsEmpty(rs!FAMILY_LEAVE_IND.OriginalValue) Then
          l_strOriginalFMLAIND = ""
       Else
          l_strOriginalFMLAIND = rs!FAMILY_LEAVE_IND.OriginalValue
       End If
       If IsEmpty(rs!Hours.OriginalValue) Then
          l_varOriginalHours = CDec(0)
       Else
          l_varOriginalHours = CDec(rs!Hours.OriginalValue)
       End If
       With l_rsRequest
           .AddNew
           .Fields("EMP_NUM") = rs!EMP_NUM.Value
           .Fields("CHARGE_DATE") = rs!CHARGE_DATE.Value
           .Fields("EXCEPTION_CD") = rs!EXCEPTION_CD.Value
           .Fields("OVER_EXC_CD") = rs!OVER_EXC_CD.Value
           .Fields("HOURS") = rs!Hours.Value
           .Fields("BANK_CODE") = l_strBANKCODE
           .Fields("FMLA_IND") = rs!FAMILY_LEAVE_IND.Value
           .Fields("SCHEDULED_IND") = rs!SCHED_TIME_IND
           .Fields("BENEFITS_PACKAGE") = Trim(l_strEmployeeBenefitPackage)
           .Fields("ORIG_EXCEPTION_CD") = l_strOriginalEXCEPTION
           .Fields("ORIG_OVER_EXC_CD") = l_strOriginalOVEREXCCD
           .Fields("ORIG_HOURS") = l_varOriginalHours
           .Fields("ORIG_BANK_CODE") = l_strOriginalBANKCODE
           .Fields("ORIG_FMLA_IND") = l_strOriginalFMLAIND
           .Fields("REC_STATUS") = rs!RECORD_STATUS.Value
           .Fields("UNION_CD") = Trim(Employee.UnionCode(rs!CHARGE_DATE.Value))
           .Update
       End With
       rs.MoveNext
    Loop
    rs.Filter = ""

' deleted scheduled time only - add to recordset here
    If Not mobjScheduledEntries Is Nothing Then
        Set rs = ScheduledEntries
        If Not rs.RecordCount = 0 Then
            rs.MoveFirst
        End If
        Do Until rs.EOF
            With l_rsRequest
                .AddNew
                .Fields("EMP_NUM") = rs!EMP_NUM
                .Fields("CHARGE_DATE") = rs!CHARGE_DATE
                .Fields("SCHEDULED_IND") = rs!SCHEDULED_IND
                .Fields("BENEFITS_PACKAGE") = rs!BENEFITS_PACKAGE
                .Fields("ORIG_EXCEPTION_CD") = rs!ORIG_EXCEPTION_CD
                .Fields("ORIG_HOURS") = rs!ORIG_HOURS
                .Fields("ORIG_BANK_CODE") = rs!ORIG_BANK_CODE
                .Fields("ORIG_FMLA_IND") = rs!ORIG_FMLA_IND
                .Fields("REC_STATUS") = rs!REC_STATUS
                .Fields("UNION_CD") = Trim(Employee.UnionCode(rs!CHARGE_DATE.Value))
                .Update
            End With
            rs.MoveNext
        Loop
    End If
    Set BuildAccrualBankUpdates = l_rsRequest
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "BuildAccrualBankUpdates"

End Function

Private Function CreateInterfaceRecordset() As ADODB.Recordset

    Dim l_rsRequest As ADODB.Recordset

On Error GoTo CreateInterfaceRecordset_EH

    Set l_rsRequest = New ADODB.Recordset
    With l_rsRequest
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "OT_BANK_CHRG_DT", adDate, , adFldMayBeNull
      .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("HOURS").Precision = 3
      .Fields.Item("HOURS").NumericScale = 1
      .Fields.Append "BANK_CODE", adChar, 2, adFldUpdatable
      .Fields.Append "FMLA_IND", adChar, 1, adFldUpdatable
      .Fields.Append "OT_OPP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "TRAVEL_PAY_HRS", adDecimal, , adFldUpdatable
      .Fields.Item("TRAVEL_PAY_HRS").Precision = 3
      .Fields.Item("TRAVEL_PAY_HRS").NumericScale = 1
      .Fields.Append "BENEFITS_PACKAGE", adChar, 4, adFldUpdatable
      .Fields.Append "UNION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "REC_STATUS", adChar, 1, adFldUpdatable
      .Fields.Append "SCHEDULED_IND", adChar, 1, adFldUpdatable
      .Fields.Append "OT_MOD_DESC", adChar, 30, adFldUpdatable
      .Fields.Append "ORIG_EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "ORIG_OVER_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "ORIG_HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("ORIG_HOURS").Precision = 3
      .Fields.Item("ORIG_HOURS").NumericScale = 1
      .Fields.Append "ORIG_BANK_CODE", adChar, 2, adFldUpdatable
      .Fields.Append "ORIG_FMLA_IND", adChar, 1, adFldUpdatable
      .Fields.Append "ORIG_OT_OPP_NUM", adInteger, , adFldUpdatable
       .Fields.Append "ORIG_TRAVEL_PAY_HRS", adDecimal, , adFldUpdatable
      .Fields.Item("ORIG_TRAVEL_PAY_HRS").Precision = 3
      .Fields.Item("ORIG_TRAVEL_PAY_HRS").NumericScale = 1

      .CursorLocation = adUseClientBatch
      .CursorType = adOpenStatic
      .LockType = adLockBatchOptimistic
      .Open
    End With
    Set CreateInterfaceRecordset = l_rsRequest
    Exit Function
    
CreateInterfaceRecordset_EH:
    TrsRaiseError mstrCMODULENAME, "CreateInterfaceRecordset"

End Function

Public Function BuildOvertimeBankUpdates() As ADODB.Recordset
    Dim l_strBANKCODE As String
    Dim l_strRecStatus As String
    Dim l_strOriginalEXCEPTION As String
    Dim l_decOriginalHours As Variant
    Dim l_strOriginalBANKCODE As String
    Dim l_strOriginalFMLAIND As String
    Dim l_strEmployeeBenefitPackage As String
    Dim l_strOriginalOverrideExcCD As String
    Dim l_decOriginalTravelPayHours As Variant
    Dim l_lngOriginalOTOppNum As Long
    Dim l_strOvertimeModificationDesc As String
    Dim l_rsRequest As ADODB.Recordset
    Dim l_strUnionCd As String
    Dim l_datChargeDate As Date
    Dim rs As ADODB.Recordset
    
On Error GoTo eh

    Set l_rsRequest = CreateInterfaceRecordset()
    Set rs = TimecardEntryRecordset
    l_strEmployeeBenefitPackage = Employee.EmployeeBenefitPackage
    rs.Filter = MDAC25FilterFix("RECORD_STATUS = 'D' OR RECORD_STATUS = 'I' OR RECORD_STATUS = 'C'")
    Do Until rs.EOF
        l_strBANKCODE = ExceptionBankCodeFor(Trim(rs!EXCEPTION_CD.Value))
       
        If IsEmpty(rs!EXCEPTION_CD.OriginalValue) Then
            l_strOriginalEXCEPTION = ""
            l_strOriginalBANKCODE = ""
        Else
            l_strOriginalEXCEPTION = rs!EXCEPTION_CD.OriginalValue
            l_strOriginalBANKCODE = ExceptionBankCodeFor(Trim(l_strOriginalEXCEPTION))
        End If
       
        If IsEmpty(rs!FAMILY_LEAVE_IND.OriginalValue) Then
            l_strOriginalFMLAIND = ""
        Else
            l_strOriginalFMLAIND = rs!FAMILY_LEAVE_IND.OriginalValue
        End If
        If IsEmpty(rs!Hours.OriginalValue) Then
            l_decOriginalHours = CDec(0)
        Else
            l_decOriginalHours = CDec(rs!Hours.OriginalValue)
        End If
        If IsEmpty(rs!OVER_EXC_CD.OriginalValue) Then
            l_strOriginalOverrideExcCD = ""
        Else
            l_strOriginalOverrideExcCD = rs!OVER_EXC_CD.OriginalValue
        End If
        If IsEmpty(rs!TRAVEL_PAY_HRS.OriginalValue) Then
            l_decOriginalTravelPayHours = CDec(0)
        Else
            l_decOriginalTravelPayHours = CDec(rs!TRAVEL_PAY_HRS.OriginalValue)
        End If
        If IsEmpty(rs!OT_OPP_NUM.OriginalValue) Then
            l_lngOriginalOTOppNum = CLng(0)
        Else
            l_lngOriginalOTOppNum = CLng(rs!OT_OPP_NUM.OriginalValue)
        End If
        If IsEmpty(rs!OT_ADJUST_DESC) Or IsNull(rs!OT_ADJUST_DESC) Then
            l_strOvertimeModificationDesc = ""
        Else
            l_strOvertimeModificationDesc = rs!OT_ADJUST_DESC
        End If
        With l_rsRequest
            .AddNew
            .Fields("EMP_NUM") = rs!EMP_NUM
            .Fields("CHARGE_DATE") = rs!CHARGE_DATE
            .Fields("START_TIME") = rs!START_TIME
            .Fields("EXCEPTION_CD") = rs!EXCEPTION_CD
            .Fields("OVER_EXC_CD") = rs!OVER_EXC_CD
            .Fields("HOURS") = CDec(rs!Hours)
            .Fields("BANK_CODE") = l_strBANKCODE
            .Fields("FMLA_IND") = rs!FAMILY_LEAVE_IND
            .Fields("OT_OPP_NUM") = rs!OT_OPP_NUM
            .Fields("TRAVEL_PAY_HRS") = rs!TRAVEL_PAY_HRS
            .Fields("SCHEDULED_IND") = rs!SCHED_TIME_IND
            .Fields("BENEFITS_PACKAGE") = Trim(l_strEmployeeBenefitPackage)
            l_datChargeDate = rs!CHARGE_DATE
            l_strUnionCd = Employee.UnionCode(l_datChargeDate)
            .Fields("UNION_CD") = Trim(l_strUnionCd)
            .Fields("ORIG_EXCEPTION_CD") = l_strOriginalEXCEPTION
            .Fields("ORIG_HOURS") = l_decOriginalHours
            .Fields("ORIG_BANK_CODE") = l_strOriginalBANKCODE
            .Fields("ORIG_FMLA_IND") = l_strOriginalFMLAIND
            .Fields("ORIG_OVER_EXC_CD") = l_strOriginalOverrideExcCD
            .Fields("ORIG_OT_OPP_NUM") = l_lngOriginalOTOppNum
            .Fields("ORIG_TRAVEL_PAY_HRS") = l_decOriginalTravelPayHours
            l_strRecStatus = UCase(Trim(rs!RECORD_STATUS))
            .Fields("REC_STATUS") = l_strRecStatus
            If l_strRecStatus = "I" Then
                .Fields("OT_BANK_CHRG_DT") = Null
                .Fields("OT_MOD_DESC") = ""
            Else
                .Fields("OT_BANK_CHRG_DT") = rs!OT_BANK_CHRG_DT
                .Fields("OT_MOD_DESC") = l_strOvertimeModificationDesc
            End If
            .Update
        End With
        rs.MoveNext
    Loop
    rs.Filter = ""

' deleted scheduled time only - add to recordset here
    If Not mobjScheduledEntries Is Nothing Then
        Set rs = ScheduledEntries
        If Not rs.RecordCount = 0 Then
            rs.MoveFirst
        End If
        Do Until rs.EOF
            With l_rsRequest
                .AddNew
                .Fields("EMP_NUM") = rs!EMP_NUM
                .Fields("CHARGE_DATE") = rs!CHARGE_DATE
                ' 9/13/00 jim does not need START_TIME,EXCEPTION_CD, OVER_EXC_CD
                ' for delete scheduled time.
                .Fields("START_TIME") = ""
                .Fields("EXCEPTION_CD") = ""
                .Fields("OVER_EXC_CD") = ""
                .Fields("SCHEDULED_IND") = rs!SCHEDULED_IND
                .Fields("BENEFITS_PACKAGE") = rs!BENEFITS_PACKAGE
                .Fields("ORIG_EXCEPTION_CD") = rs!ORIG_EXCEPTION_CD
                .Fields("ORIG_HOURS") = rs!ORIG_HOURS
                .Fields("ORIG_BANK_CODE") = rs!ORIG_BANK_CODE
                .Fields("ORIG_FMLA_IND") = rs!ORIG_FMLA_IND
                .Fields("ORIG_OVER_EXC_CD") = l_strOriginalOverrideExcCD
                .Fields("ORIG_OT_OPP_NUM") = l_lngOriginalOTOppNum
                .Fields("ORIG_TRAVEL_PAY_HRS") = l_decOriginalTravelPayHours
                .Fields("OT_BANK_CHRG_DT") = Null
                .Fields("OT_MOD_DESC") = l_strOvertimeModificationDesc
                .Fields("REC_STATUS") = rs!REC_STATUS
                .Update
            End With
            rs.MoveNext
        Loop
    End If

    Set BuildOvertimeBankUpdates = l_rsRequest
    Exit Function
    
eh:
     TrsRaiseError mstrCMODULENAME, "BuildOvertimeBankUpdates"

End Function


Public Function IsDateAHoliday( _
            ByVal v_datDate As Date _
            ) As Boolean
'Release 07/25/2001 -- JFB removed the hard coding of H and HR

    Dim tce As cboTimecardEntry
    Dim l_blnIsAHoliday As Boolean
  
On Error GoTo eh
    l_blnIsAHoliday = False
    
' Schedule provided for time collection does not have exceptions, so it will always
' check time entries for answer.  Schedule provided from overtime will have scheduled
' exceptions.
    If Me.Schedule.IsScheduledHoliday(v_datDate) Then
        l_blnIsAHoliday = True
        GoTo xit
    End If
    
    For Each tce In TimecardEntries
        If IsAHolidayException(tce.ExceptionCode) And tce.ChargeDate = v_datDate Then
            l_blnIsAHoliday = True
            Exit For
        End If
    Next
    
xit:
    IsDateAHoliday = l_blnIsAHoliday
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsDateAHoliday"
            
End Function


Public Function IsDateAVacation( _
            ByVal v_datDate As Date _
            ) As Boolean
'Release 07/25/2001 -- JFB  Added this function to get to see if
'                             vacation was reported on a day
            
    Dim tce As cboTimecardEntry
    Dim l_blnIsAVacation As Boolean
  
On Error GoTo eh
    l_blnIsAVacation = False
    
' Schedule provided for time collection does not have exceptions, so it will always
' check time entries for answer.  Schedule provided from overtime will have scheduled
' exceptions.
    If Me.Schedule.IsScheduledVacation(v_datDate) Then
        l_blnIsAVacation = True
        GoTo xit
    End If
    
    For Each tce In TimecardEntries
        If IsAVacationException(tce.ExceptionCode) And tce.ChargeDate = v_datDate Then
            l_blnIsAVacation = True
            Exit For
        End If
    Next
    
xit:
    IsDateAVacation = l_blnIsAVacation
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsDateAVacation"
            
End Function


Public Function DoesOvertimeOverlapWithVacation( _
            ByRef v_Entry As cboTimecardEntry) As Boolean

    Dim colCurrentOverlap As Collection
    Dim l_OverlapEntry As cboTimecardEntry
    
    On Error GoTo eh

    Set colCurrentOverlap = GetBSTimecard.FindOverlappingEntries(v_Entry, mcolTimecardEntries)
    If colCurrentOverlap.Count = 0 Then
        DoesOvertimeOverlapWithVacation = False
        Exit Function
    End If
    Set l_OverlapEntry = colCurrentOverlap.Item(1)
    If IsAVacationException(l_OverlapEntry.ExceptionCode) Then
        DoesOvertimeOverlapWithVacation = True
    Else
        DoesOvertimeOverlapWithVacation = False
    End If
    
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "DoesOvertimeOverlapWithVacation"

End Function


Public Function DoesTimeEntryOverlapWithHoliday( _
            ByRef v_Entry As cboTimecardEntry) As Boolean

    Dim colCurrentOverlap As Collection
    Dim l_OverlapEntry As cboTimecardEntry
    
    On Error GoTo eh

    Set colCurrentOverlap = GetBSTimecard.FindOverlappingEntries(v_Entry, mcolTimecardEntries)
    If colCurrentOverlap.Count = 0 Then
        DoesTimeEntryOverlapWithHoliday = False
        Exit Function
    End If
    Set l_OverlapEntry = colCurrentOverlap.Item(1)
    If IsAHolidayException(l_OverlapEntry.ExceptionCode) Then
        DoesTimeEntryOverlapWithHoliday = True
    Else
        DoesTimeEntryOverlapWithHoliday = False
    End If
    
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "DoesTimeEntryOverlapWithHoliday"

End Function


Public Function DoesTimeEntryOverlapWithHolidayOrVacation( _
            ByRef v_Entry As cboTimecardEntry) As Boolean

    Dim colCurrentOverlap As Collection
    Dim l_OverlapEntry As cboTimecardEntry
    
    On Error GoTo eh

    Set colCurrentOverlap = GetBSTimecard.FindOverlappingEntries(v_Entry, TimecardEntries)
    If colCurrentOverlap.Count = 0 Then
        DoesTimeEntryOverlapWithHolidayOrVacation = False
        Exit Function
    End If
    Set l_OverlapEntry = colCurrentOverlap.Item(1)
    If IsAHolidayException(l_OverlapEntry.ExceptionCode) Or _
       IsAVacationException(l_OverlapEntry.ExceptionCode) Or _
       IsSpecialOverrideCode(l_OverlapEntry.OverrideExceptionCode) Then
        DoesTimeEntryOverlapWithHolidayOrVacation = True
    Else
        DoesTimeEntryOverlapWithHolidayOrVacation = False
    End If
    
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "DoesTimeEntryOverlapWithHolidayOrVacation"

End Function


Private Function GetScheduleChangePayments( _
    ByVal v_datChargeDate As Date, _
    ByVal v_datDateWorked As Date, _
    ByVal v_strStartTime As String, _
    ByVal v_strEndTime As String _
    ) As Collection
    
    Dim l_colSchedChgPmts As Collection
    Dim l_colSCPTemp As Collection
    Dim l_varSchedChgPmt As Variant
    Dim l_intCount As Integer
    Dim l_intI As Integer
    Dim l_datStartDateIn As Date
    Dim l_datEndDateIn As Date
    Dim l_datStartDateSCP As Date
    Dim l_datEndDateSCP As Date
    Dim l_datEndTimeSCP As String
    
    On Error GoTo eh
    
    l_datStartDateIn = DateTimeFrom(v_datDateWorked, v_strStartTime)
    l_datEndDateIn = DateTimeFrom(v_datDateWorked, v_strEndTime)
    Set l_colSCPTemp = Me.Schedule.ScheduleChangePaymentsOn(v_datChargeDate)
    l_intCount = l_colSCPTemp.Count
    Set l_colSchedChgPmts = New Collection
    For l_intI = 1 To l_intCount
        l_varSchedChgPmt = l_colSCPTemp(l_intI)
        If Not IsArray(l_varSchedChgPmt) Then
            Set GetScheduleChangePayments = New Collection
            Exit Function
        End If
        l_datStartDateSCP = DateTimeFrom(l_varSchedChgPmt(2), l_varSchedChgPmt(4))
        If l_varSchedChgPmt(4) > l_varSchedChgPmt(5) Then
            l_datEndDateSCP = DateTimeFrom(DateAdd("d", 1, l_varSchedChgPmt(2)), l_varSchedChgPmt(5))
        Else
            l_datEndDateSCP = DateTimeFrom(l_varSchedChgPmt(2), l_varSchedChgPmt(5))
        End If

        If TimePeriodsIntersect( _
            l_datStartDateSCP, l_datEndDateSCP, _
            l_datStartDateIn, l_datEndDateIn) Then
            If l_varSchedChgPmt(4) > l_varSchedChgPmt(5) Then
            'Split Payment to before and After Midnight
                l_datEndTimeSCP = l_varSchedChgPmt(5)
                l_varSchedChgPmt(5) = "24:00"
                l_colSchedChgPmts.Add l_varSchedChgPmt
                l_varSchedChgPmt(4) = "00:00"
                l_varSchedChgPmt(5) = l_datEndTimeSCP
                l_varSchedChgPmt(2) = DateAdd("d", 1, l_varSchedChgPmt(2))
                l_colSchedChgPmts.Add l_varSchedChgPmt
            Else
                l_colSchedChgPmts.Add l_varSchedChgPmt
            End If
        End If
    Next l_intI
        
    Set GetScheduleChangePayments = l_colSchedChgPmts
    
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "GetScheduleChangePayments"
    
    
End Function

Public Function BuildEMPRVInterface() As Collection


    Dim l_colEMPRVInt As Collection
    Dim rs As ADODB.Recordset
    Dim l_strDept As String
    Dim l_datChargeDate As Date
    Dim l_strRecStatus As String
    Dim l_strException As String
    Dim l_strOverExc As String
    Dim l_strOrigException As String
    Dim l_strOrigOverExc As String
    
On Error GoTo eh



    Set l_colEMPRVInt = New Collection
    Set rs = TimecardEntryRecordset
    rs.Filter = MDAC25FilterFix("RECORD_STATUS = 'D' OR RECORD_STATUS = 'I' OR RECORD_STATUS = 'C'")
    While Not rs.EOF
        l_datChargeDate = rs!CHARGE_DATE
        l_strDept = mobjEmployee.DepartmentNumber(l_datChargeDate)
        If IsEMPRVEmployee(l_strDept) Then
            l_strRecStatus = rs!RECORD_STATUS
            Select Case l_strRecStatus
                Case "I"
                    If rs!SCHED_TIME_IND <> "Y" Then
                        AddEMRVInsert rs, l_colEMPRVInt
                    End If
                Case "D"
                    AddEMRVDelete rs, l_colEMPRVInt
                Case "C"
                    l_strException = UCase(Trim(rs!EXCEPTION_CD))
                    l_strOverExc = UCase(Trim(rs!OVER_EXC_CD))
                    l_strOrigException = UCase(Trim(rs("EXCEPTION_CD").OriginalValue))
                    l_strOrigOverExc = UCase(Trim(rs("OVER_EXC_CD").OriginalValue))
                    If l_strException = l_strOrigException And _
                        l_strOverExc = l_strOrigOverExc Then
                    'Skip do not send request
                    Else
                        AddEMRVDelete rs, l_colEMPRVInt
                        AddEMRVInsert rs, l_colEMPRVInt
                    End If
            End Select
        End If
        rs.MoveNext
    Wend
    rs.Filter = ""

' deleted scheduled time only - add to interface here
    If RecordsExist(mobjScheduledEntries) Then
        Set rs = ScheduledEntries
        rs.MoveFirst
        While Not rs.EOF
            l_datChargeDate = rs!CHARGE_DATE
            l_strDept = mobjEmployee.DepartmentNumber(l_datChargeDate)
            If IsEMPRVEmployee(l_strDept) Then
                AddEMRVScheduleDelete rs, l_colEMPRVInt
            End If
            rs.MoveNext
        Wend
    End If

    Set BuildEMPRVInterface = l_colEMPRVInt
    Exit Function
    
eh:
     TrsRaiseError mstrCMODULENAME, "BuildEMPRVInterface"

End Function
Public Function BuildEMPRVJobTasks() As ADODB.Recordset

    Dim l_rsEMPRVJobTasks As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim l_strDept As String
    Dim l_datChargeDate As Date
    Dim l_strRecStatus As String
    Dim l_strException As String
    Dim l_strOverExc As String
    Dim l_strOrigException As String
    Dim l_strOrigOverExc As String
    
On Error GoTo eh

    Set rs = TimecardEntryRecordset
    rs.Filter = "(RECORD_STATUS = 'D' AND DISTRIBUTION_TYPE = 'J') OR (RECORD_STATUS = 'I' AND DISTRIBUTION_TYPE = 'J') OR (RECORD_STATUS = 'C')"
    Set l_rsEMPRVJobTasks = New ADODB.Recordset
    With l_rsEMPRVJobTasks
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
        .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
        .Fields.Append "JOB_NBR", adVarChar, 8, adFldUpdatable
        .Fields.Append "JOBTASK_NBR", adVarChar, 4, adFldUpdatable
        .Fields.Append "JOBTASK_STEP_NBR", adInteger, , adFldUpdatable
        .Fields.Append "EMPRV_DATABASE_ID", adVarChar, 8, adFldUpdatable
        .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, adFldUpdatable
        .Fields.Append "HOURS", adDecimal, , adFldUpdatable
        .Fields.Item("HOURS").Precision = 3
        .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "OT_IND", adChar, 1, adFldUpdatable
        .Fields.Append "CHG_IND", adChar, 1, adFldUpdatable
        .Fields.Append "DEPT_NUM", adChar, 3, adFldUpdatable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open

        While Not rs.EOF
            l_datChargeDate = rs!CHARGE_DATE
            l_strDept = mobjEmployee.DepartmentNumber(l_datChargeDate)
            If IsJobTaskEmployee(l_strDept) Then
                l_strRecStatus = rs!RECORD_STATUS
                Select Case l_strRecStatus
                    Case "I"
                        .AddNew
                        .Fields("CHG_IND") = "I"
                        .Fields("EMP_NUM") = rs.Fields("EMP_NUM")
                        .Fields("DATE_WORKED") = rs.Fields("DATE_WORKED")
                        .Fields("CHARGE_DATE") = rs.Fields("CHARGE_DATE")
                        .Fields("START_TIME") = rs.Fields("START_TIME")
                        .Fields("END_TIME") = rs.Fields("END_TIME")
                        .Fields("WG_NUM") = rs.Fields("WG_NUM")
                        .Fields("JOB_NBR") = rs.Fields("JOB_NBR")
                        .Fields("JOBTASK_NBR") = rs.Fields("JOBTASK_NBR")
                        .Fields("JOBTASK_STEP_NBR") = rs.Fields("JOBTASK_STEP_NBR")
                        .Fields("EXCEPTION_CD") = rs.Fields("EXCEPTION_CD")
                        .Fields("DISTRIBUTION_ID") = rs.Fields("DISTRIBUTION_ID")
                        .Fields("HOURS") = rs.Fields("HOURS")
                        .Fields("EMPRV_DATABASE_ID") = rs.Fields("EMPRV_DATABASE_ID")
                        If IsJobTaskOTExcep(.Fields("EXCEPTION_CD")) Then
                            .Fields("OT_IND") = "Y"
                        Else
                            .Fields("OT_IND") = "N"
                        End If
                        .Fields("DEPT_NUM") = rs.Fields("DEPT_NUM")
                        .Update
                    Case "D"
                        .AddNew
                        .Fields("CHG_IND") = "D"
                        .Fields("EMP_NUM") = rs.Fields("EMP_NUM")
                        .Fields("DATE_WORKED") = rs.Fields("DATE_WORKED")
                        .Fields("CHARGE_DATE") = rs.Fields("CHARGE_DATE")
                        .Fields("START_TIME") = rs.Fields("START_TIME")
                        .Fields("END_TIME") = rs.Fields("END_TIME")
                        .Fields("WG_NUM") = rs.Fields("WG_NUM")
                        .Fields("JOB_NBR") = rs.Fields("JOB_NBR")
                        .Fields("JOBTASK_NBR") = rs.Fields("JOBTASK_NBR")
                        .Fields("JOBTASK_STEP_NBR") = rs.Fields("JOBTASK_STEP_NBR")
                        .Fields("EXCEPTION_CD") = rs.Fields("EXCEPTION_CD")
                        .Fields("DISTRIBUTION_ID") = rs.Fields("DISTRIBUTION_ID")
                        .Fields("HOURS") = rs.Fields("HOURS")
                        .Fields("EMPRV_DATABASE_ID") = rs.Fields("EMPRV_DATABASE_ID")
                        If IsJobTaskOTExcep(.Fields("EXCEPTION_CD")) Then
                            .Fields("OT_IND") = "Y"
                        Else
                            .Fields("OT_IND") = "N"
                        End If
                        .Fields("DEPT_NUM") = rs.Fields("DEPT_NUM")
                        .Update
                    Case "C"
                        If rs.Fields("DISTRIBUTION_TYPE").Value = "J" Then
                            .AddNew
                            .Fields("EMP_NUM") = rs.Fields("EMP_NUM").Value
                            .Fields("DATE_WORKED") = rs.Fields("DATE_WORKED").Value
                            .Fields("CHARGE_DATE") = rs.Fields("CHARGE_DATE").Value
                            .Fields("START_TIME") = rs.Fields("START_TIME").Value
                            .Fields("END_TIME") = rs.Fields("END_TIME").Value
                            .Fields("WG_NUM") = rs.Fields("WG_NUM").Value
                            .Fields("JOB_NBR") = rs.Fields("JOB_NBR").Value
                            .Fields("JOBTASK_NBR") = rs.Fields("JOBTASK_NBR").Value
                            .Fields("JOBTASK_STEP_NBR") = rs.Fields("JOBTASK_STEP_NBR").Value
                            .Fields("EXCEPTION_CD") = rs.Fields("EXCEPTION_CD").Value
                            .Fields("DISTRIBUTION_ID") = rs.Fields("DISTRIBUTION_ID").Value
                            .Fields("HOURS") = rs.Fields("HOURS").Value
                            .Fields("EMPRV_DATABASE_ID") = rs.Fields("EMPRV_DATABASE_ID").Value
                            If IsJobTaskOTExcep(.Fields("EXCEPTION_CD").Value) Then
                                .Fields("OT_IND") = "Y"
                            Else
                                .Fields("OT_IND") = "N"
                            End If
                            .Fields("CHG_IND") = "I"
                            .Fields("DEPT_NUM") = rs.Fields("DEPT_NUM").Value
                            .Update
                        End If
                        If rs.Fields("DISTRIBUTION_TYPE").OriginalValue = "J" Then
                            .AddNew
                            .Fields("EMP_NUM") = rs.Fields("EMP_NUM").OriginalValue
                            .Fields("DATE_WORKED") = rs.Fields("DATE_WORKED").OriginalValue
                            .Fields("CHARGE_DATE") = rs.Fields("CHARGE_DATE").OriginalValue
                            .Fields("START_TIME") = rs.Fields("START_TIME").OriginalValue
                            .Fields("END_TIME") = rs.Fields("END_TIME").OriginalValue
                            .Fields("WG_NUM") = rs.Fields("WG_NUM").OriginalValue
                            .Fields("JOB_NBR") = rs.Fields("JOB_NBR").OriginalValue
                            .Fields("JOBTASK_NBR") = rs.Fields("JOBTASK_NBR").OriginalValue
                            .Fields("JOBTASK_STEP_NBR") = rs.Fields("JOBTASK_STEP_NBR").OriginalValue
                            .Fields("EXCEPTION_CD") = rs.Fields("EXCEPTION_CD").OriginalValue
                            .Fields("DISTRIBUTION_ID") = rs.Fields("DISTRIBUTION_ID").OriginalValue
                            .Fields("HOURS") = rs.Fields("HOURS").OriginalValue
                            .Fields("EMPRV_DATABASE_ID") = rs.Fields("EMPRV_DATABASE_ID").OriginalValue
                            If IsJobTaskOTExcep(.Fields("EXCEPTION_CD").OriginalValue) Then
                                .Fields("OT_IND") = "Y"
                            Else
                                .Fields("OT_IND") = "N"
                            End If
                            .Fields("CHG_IND") = "D"
                            .Fields("DEPT_NUM") = rs.Fields("DEPT_NUM").Value
                            .Update
                        End If
                End Select
            End If
            rs.MoveNext
        Wend
    End With
    rs.Filter = ""

    Set BuildEMPRVJobTasks = l_rsEMPRVJobTasks
    Exit Function
    
eh:
     TrsRaiseError mstrCMODULENAME, "BuildEMPRVJobTasks"

End Function

Private Function IsEMPRVEmployee( _
    ByVal v_strDept As String _
    ) As Boolean
    
    Dim l_obj As busDepartment.cbsDepartment
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("busDepartment.cbsDepartment")
    IsEMPRVEmployee = l_obj.IsEMPRVDepartmentNumber(v_strDept)
    Exit Function
    
eh:
     TrsRaiseError mstrCMODULENAME, "IsEMPRVEmployee"
    
End Function
Private Function IsJobTaskEmployee( _
    ByVal v_strDept As String _
    ) As Boolean
    
    Dim l_obj As busDepartment.cbsDepartment
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("busDepartment.cbsDepartment")
    IsJobTaskEmployee = l_obj.IsJobTaskDepartmentNumber(v_strDept)
    Exit Function
    
eh:
     TrsRaiseError mstrCMODULENAME, "IsEMPRVEmployee"
    
End Function
Private Function IsJobTaskOTExcep( _
    ByVal v_strException As String _
    ) As Boolean
    
    Dim l_obj As dbRules.cdRules
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    IsJobTaskOTExcep = l_obj.DoesRuleApply(143, trExceptionCode, v_strException)
    Exit Function
    
eh:
     TrsRaiseError mstrCMODULENAME, "IsJobTaskOTExcep"
    
End Function

Private Sub AddEMRVInsert( _
    ByVal v_rsTimeEntry As ADODB.Recordset, _
    ByRef r_colEMPRVInt As Collection)
    
    Dim l_strException As String
    Dim l_strOverExc As String
    Dim l_lngEmp As Long
    Dim l_datChargeDate As Date
    Dim l_decHours As Variant
    Dim l_lngWorkGroup As Long
    
    
    On Error GoTo eh
    
    l_strException = UCase(Trim(v_rsTimeEntry!EXCEPTION_CD))
    l_strOverExc = UCase(Trim(v_rsTimeEntry!OVER_EXC_CD))
    
    If IsEMPRVException(l_strException, l_strOverExc) Then
        l_lngEmp = v_rsTimeEntry!EMP_NUM
        l_datChargeDate = v_rsTimeEntry!CHARGE_DATE
        l_decHours = CDec(v_rsTimeEntry!Hours)
        l_lngWorkGroup = v_rsTimeEntry!WG_NUM
    
        r_colEMPRVInt.Add BuildEMPRVTransaction( _
            l_lngEmp, l_datChargeDate, l_strException, _
            l_strOverExc, l_decHours, l_lngWorkGroup, "I")
    End If
    
    Exit Sub
    
eh:
     TrsRaiseError mstrCMODULENAME, "AddEMRVInsert"
    
End Sub

Private Sub AddEMRVDelete( _
    ByVal v_rsTimeEntry As ADODB.Recordset, _
    ByRef r_colEMPRVInt As Collection)
    
    Dim l_strException As String
    Dim l_strOverExc As String
    Dim l_lngEmp As Long
    Dim l_datChargeDate As Date
    Dim l_decHours As Variant
    Dim l_lngWorkGroup As Long
    
    
    On Error GoTo eh
    
    If IsEmpty(v_rsTimeEntry!EXCEPTION_CD.OriginalValue) Then
       l_strException = ""
    Else
       l_strException = UCase(Trim(v_rsTimeEntry!EXCEPTION_CD.OriginalValue))
    End If
    
    If IsEmpty(v_rsTimeEntry!OVER_EXC_CD.OriginalValue) Then
       l_strOverExc = ""
    Else
       l_strOverExc = UCase(Trim(v_rsTimeEntry!OVER_EXC_CD.OriginalValue))
    End If

    If IsEMPRVException(l_strException, l_strOverExc) Then
        l_lngEmp = v_rsTimeEntry!EMP_NUM
        l_datChargeDate = v_rsTimeEntry!CHARGE_DATE
        If IsEmpty(v_rsTimeEntry!Hours.OriginalValue) Then
            l_decHours = CDec(0)
        Else
            l_decHours = CDec(v_rsTimeEntry!Hours.OriginalValue)
        End If
        If IsEmpty(v_rsTimeEntry!WG_NUM.OriginalValue) Then
            l_lngWorkGroup = 0
        Else
            l_lngWorkGroup = CDec(v_rsTimeEntry!WG_NUM.OriginalValue)
        End If
    
        r_colEMPRVInt.Add BuildEMPRVTransaction( _
            l_lngEmp, l_datChargeDate, l_strException, _
            l_strOverExc, l_decHours, l_lngWorkGroup, "D")
    End If
    
    Exit Sub
    
eh:
     TrsRaiseError mstrCMODULENAME, "AddEMRVDelete"
    
End Sub

Private Sub AddEMRVScheduleDelete( _
    ByVal v_rsTimeEntry As ADODB.Recordset, _
    ByRef r_colEMPRVInt As Collection)
    
    Dim l_strException As String
    Dim l_strOverExc As String
    Dim l_lngEmp As Long
    Dim l_datChargeDate As Date
    Dim l_decHours As Variant
    Dim l_lngWorkGroup As Long
    
    
    On Error GoTo eh
    
    l_strException = UCase(Trim(v_rsTimeEntry!ORIG_EXCEPTION_CD))
    l_strOverExc = UCase(Trim(v_rsTimeEntry!ORIG_OVER_EXC_CD))
    
    If IsEMPRVException(l_strException, l_strOverExc) Then
        l_lngEmp = v_rsTimeEntry!EMP_NUM
        l_datChargeDate = v_rsTimeEntry!CHARGE_DATE
        l_decHours = CDec(v_rsTimeEntry!ORIG_HOURS)
        l_lngWorkGroup = v_rsTimeEntry!WG_NUM
    
        r_colEMPRVInt.Add BuildEMPRVTransaction( _
            l_lngEmp, l_datChargeDate, l_strException, _
            l_strOverExc, l_decHours, l_lngWorkGroup, "D")
    End If
    
    Exit Sub
    
eh:
     TrsRaiseError mstrCMODULENAME, "AddEMRVScheduleDelete"
    
End Sub



Private Function BuildEMPRVTransaction( _
    ByVal v_lngEmployee As Long, _
    ByVal v_datChargeDate As Date, _
    ByVal v_strException As String, _
    ByVal v_strOverExcCd As String, _
    ByVal v_decHours As Variant, _
    ByVal v_lngWorkGroup As Long, _
    ByVal v_strTransType As String _
    ) As Variant
    
    
    Dim l_varEMPRVTran(0 To 6) As Variant

    Dim l_strDept As String
    Dim l_intLocCd As Integer
    

On Error GoTo eh

    l_intLocCd = EmployeeWorkgroupLocation( _
             v_lngWorkGroup, v_datChargeDate)
    l_strDept = mobjEmployee.DepartmentNumber(v_datChargeDate)
    l_varEMPRVTran(etEMPNUM) = v_lngEmployee
    l_varEMPRVTran(etDEPTNUM) = l_strDept
    If v_strOverExcCd <> "" Then
        l_varEMPRVTran(etEXCCD) = v_strOverExcCd
    Else
        l_varEMPRVTran(etEXCCD) = v_strException
    End If
    l_varEMPRVTran(etCHGDATE) = v_datChargeDate
    l_varEMPRVTran(etHOURS) = v_decHours
    l_varEMPRVTran(etLOCCD) = l_intLocCd
    l_varEMPRVTran(etTRANCD) = v_strTransType

    BuildEMPRVTransaction = l_varEMPRVTran
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "BuildEMPRVTransaction"
End Function


Private Function IsEMPRVException( _
    ByVal v_strException As String, _
    ByVal v_strOverExc As String _
    ) As Boolean
    
    
    On Error GoTo eh
    
    If v_strException = "" And v_strOverExc = "" Then
        IsEMPRVException = False
        Exit Function
    End If
    
    If v_strException = "ST" And v_strOverExc = "" Then
        IsEMPRVException = False
        Exit Function
    End If
        
    IsEMPRVException = True
    
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsEMPRVException"

End Function


Private Function EmployeeWorkgroupLocation( _
    ByVal v_lngWorkGroup As Long, _
    ByVal v_datChargeDate As Date _
    ) As Integer
    
    Dim l_lngWG As Long
    
    On Error GoTo eh
    
    If v_lngWorkGroup > 0 Then
        l_lngWG = v_lngWorkGroup
    Else
        l_lngWG = mobjEmployee.GetPrimaryWorkgroupNumber(v_datChargeDate)
    End If
    EmployeeWorkgroupLocation = mobjEmployee.GetWorkgroupLocationCode(l_lngWG)

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "EmployeeWorkgroupLocation"

End Function


Private Sub ApplyShiftComponentSplits( _
    ByRef r_objTimeEntry As cboTimecardEntry, _
    ByRef r_colTimeEntries As Collection, _
    ByRef r_blnSchedSplits As Boolean)
    
    Dim l_blnIsCoffeen As Boolean
    Dim l_blnIsCILCOLocal8 As Boolean
    Dim l_strJobNum As String
    Dim l_aryShifts As Variant
    Dim l_strShiftComp As String
    Dim l_blnCoffeen12HourAgreement As Boolean
    Dim l_blnHutsonville148 As Boolean
    Dim l_blnHutsonville12HourAgreement As Boolean
    Dim l_blnLocal8MaterialHandler As Boolean
    Dim l_datStartTimeAsDate As Date
    Dim l_datEndTimeAsDate As Date
    Dim l_blnHutsonvilleMaint As Boolean
    Dim l_blnCallawaySecurity As Boolean
    Dim l_strOverExcCd As String
    
    
    On Error GoTo eh
    
    l_blnIsCoffeen = mobjEmployee.IsCoffeen(r_objTimeEntry.ChargeDate)
    l_blnIsCILCOLocal8 = mobjEmployee.IsAmerenCILCOLocal8Employee(r_objTimeEntry.ChargeDate)
    l_blnCoffeen12HourAgreement = Me.Employee.IsCoffeen12HourAgreementJob(r_objTimeEntry.ChargeDate, r_objTimeEntry.JobNumber)
    l_blnHutsonville148 = Me.Employee.IsHutsonvilleLocal148(r_objTimeEntry.ChargeDate)
    l_blnHutsonville12HourAgreement = Me.Employee.IsHutsonville12HourAgreementJob(r_objTimeEntry.ChargeDate, r_objTimeEntry.JobNumber)
    l_blnLocal8MaterialHandler = Me.Employee.IsCILCOLocal8MaterialHandler(r_objTimeEntry.ChargeDate)
    l_blnHutsonvilleMaint = Me.Employee.IsHutsonvilleMaint(r_objTimeEntry.ChargeDate) _
                                Or Me.Employee.IsHutsonvilleMaintJob(r_objTimeEntry.ChargeDate, r_objTimeEntry.JobNumber)
    l_blnCallawaySecurity = Me.Employee.IsCallawaySecurity(r_objTimeEntry.ChargeDate)
    
    
    If l_blnIsCoffeen Or l_blnIsCILCOLocal8 Or l_blnHutsonville148 Or l_blnCallawaySecurity Then
    'continue
        r_blnSchedSplits = True
    Else
        r_colTimeEntries.Add r_objTimeEntry
        r_blnSchedSplits = False
        Exit Sub
    End If
            
    If l_blnIsCILCOLocal8 Then
        If Trim(r_objTimeEntry.ShiftComponentCode) = "R" Then
             r_colTimeEntries.Add r_objTimeEntry
             r_blnSchedSplits = False
             Exit Sub
        Else
            l_strShiftComp = mobjSchedule.ShiftComponentOnJobNum(r_objTimeEntry.ChargeDate, r_objTimeEntry.StartTimeAsDate, r_objTimeEntry.EndTimeAsDate, r_objTimeEntry.JobNumber)
            If l_strShiftComp = "R" Then
                r_objTimeEntry.ShiftComponentCode = l_strShiftComp
                r_colTimeEntries.Add r_objTimeEntry
                r_blnSchedSplits = False
                Exit Sub
            End If
            If l_blnLocal8MaterialHandler Then
                If Left(r_objTimeEntry.JobNumber, 2) = "ND" Then 'Duck Creek
                    l_aryShifts = GetCILCOLocal8MaterialHandlerShifts()
                Else
'Temporary fix for Edwards Station
'Edwards sold to Dynegy on 10/01/2013
                    l_aryShifts = GetCILCOLocal8MaterialHandlerShifts2()
                End If
            Else
                l_aryShifts = GetCILCOLocal8Shifts()
            End If
        
        End If
    End If
        
    If l_blnIsCoffeen Then
        If l_blnCoffeen12HourAgreement Then
            l_aryShifts = GetCoffeen12HourShifts()
        Else
            l_aryShifts = GetCoffeenShifts()
        End If
    End If
    
    If l_blnHutsonville148 Then
        If l_blnHutsonville12HourAgreement Then
            l_aryShifts = GetStandardShifts()
        Else
'only split if it is overtime
            l_datStartTimeAsDate = DateTimeFrom(r_objTimeEntry.DateWorked, r_objTimeEntry.startTime)
            l_datEndTimeAsDate = DateTimeFrom(r_objTimeEntry.DateWorked, r_objTimeEntry.EndTime)
            If mobjSchedule.IsAnyPeriodWithinSchedule(r_objTimeEntry.ChargeDate, l_datStartTimeAsDate, l_datEndTimeAsDate) Then
                r_colTimeEntries.Add r_objTimeEntry
                r_blnSchedSplits = False
                Exit Sub
            Else
            'Overtime
                If l_blnHutsonvilleMaint Then
                    l_aryShifts = GetMaintenanceShifts()
                Else
                    l_aryShifts = GetStandardShifts()
                End If
            End If
        End If
    End If
            
    
    If l_blnCallawaySecurity Then
    'only split if it is overtime
            l_datStartTimeAsDate = DateTimeFrom(r_objTimeEntry.DateWorked, r_objTimeEntry.startTime)
            l_datEndTimeAsDate = DateTimeFrom(r_objTimeEntry.DateWorked, r_objTimeEntry.EndTime)
            If mobjSchedule.IsAnyPeriodWithinSchedule(r_objTimeEntry.ChargeDate, l_datStartTimeAsDate, l_datEndTimeAsDate) Then
                r_colTimeEntries.Add r_objTimeEntry
                r_blnSchedSplits = False
                Exit Sub
            Else
            'Overtime
            'Allow user to enter NC (No Component), which will remove the shift component
                If IsNull(r_objTimeEntry.ShiftComponentCode) Or IsEmpty(r_objTimeEntry.ShiftComponentCode) Then
                    l_strShiftComp = ""
                Else
                    l_strShiftComp = r_objTimeEntry.ShiftComponentCode
                End If
                If l_strShiftComp = "NC" Then
                    r_colTimeEntries.Add r_objTimeEntry
                    r_blnSchedSplits = False
                    Exit Sub
                End If
            'do not apply shift component splits to turnover time and combat drills
                If IsNull(r_objTimeEntry.OverrideExceptionCode) Or IsEmpty(r_objTimeEntry.OverrideExceptionCode) Then
                    l_strOverExcCd = ""
                Else
                    l_strOverExcCd = r_objTimeEntry.OverrideExceptionCode
                End If
                If CallawaySecurityShiftComponentEligible(l_strOverExcCd) Then
                    l_aryShifts = GetCallawaySecurity12HourShifts()
                Else
                    r_colTimeEntries.Add r_objTimeEntry
                    r_blnSchedSplits = False
                    Exit Sub
                End If
            End If
    End If
    
    ShiftComponentSplitsDetermine l_aryShifts, r_objTimeEntry, r_colTimeEntries
    
    
    Exit Sub

eh:
    TrsRaiseError mstrCMODULENAME, "ApplyShiftComponentSplits"
    
End Sub

Private Function GetCoffeenShifts() As Variant

    Dim l_aryShifts(3, 2) As String

    
    On Error GoTo eh
    
    l_aryShifts(0, 0) = "O"
    l_aryShifts(0, 1) = "00:00"
    l_aryShifts(0, 2) = "07:00"
    l_aryShifts(1, 0) = ""
    l_aryShifts(1, 1) = "07:00"
    l_aryShifts(1, 2) = "15:00"
    l_aryShifts(2, 0) = "P"
    l_aryShifts(2, 1) = "15:00"
    l_aryShifts(2, 2) = "23:00"
    l_aryShifts(3, 0) = "O"
    l_aryShifts(3, 1) = "23:00"
    l_aryShifts(3, 2) = "24:00"
    
    GetCoffeenShifts = l_aryShifts
    

  Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetCoffeenShifts"
    
End Function

Private Function GetCILCOLocal8Shifts() As Variant

    Dim l_aryShifts(3, 2) As String

    
    On Error GoTo eh
    
    l_aryShifts(0, 0) = "P"
    l_aryShifts(0, 1) = "00:00"
    l_aryShifts(0, 2) = "06:30"
    l_aryShifts(1, 0) = ""
    l_aryShifts(1, 1) = "06:30"
    l_aryShifts(1, 2) = "14:30"
    l_aryShifts(2, 0) = "P"
    l_aryShifts(2, 1) = "14:30"
    l_aryShifts(2, 2) = "24:00"
    l_aryShifts(3, 0) = "P"
    l_aryShifts(3, 1) = "24:00"
    l_aryShifts(3, 2) = "24:00"
    
    GetCILCOLocal8Shifts = l_aryShifts
    

  Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetCILCOLocal8Shifts"
    
End Function
    
Private Function GetCILCOLocal8MaterialHandlerShifts() As Variant

    Dim l_aryShifts(3, 2) As String

    
    On Error GoTo eh
    
    l_aryShifts(0, 0) = "P"
    l_aryShifts(0, 1) = "00:00"
    l_aryShifts(0, 2) = "07:00"
    l_aryShifts(1, 0) = ""
    l_aryShifts(1, 1) = "07:00"
    l_aryShifts(1, 2) = "15:00"
    l_aryShifts(2, 0) = "P"
    l_aryShifts(2, 1) = "15:00"
    l_aryShifts(2, 2) = "24:00"
    l_aryShifts(3, 0) = "P"
    l_aryShifts(3, 1) = "24:00"
    l_aryShifts(3, 2) = "24:00"
    
    GetCILCOLocal8MaterialHandlerShifts = l_aryShifts
    

  Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetCILCOLocal8MaterialHandlerShifts"
    
End Function

Private Function GetCILCOLocal8MaterialHandlerShifts2() As Variant

'Temporary fix for Edwards Station
'Edwards sold to Dynegy on 10/01/2013

    Dim l_aryShifts(3, 2) As String

    
    On Error GoTo eh
    
    l_aryShifts(0, 0) = "P"
    l_aryShifts(0, 1) = "00:00"
    l_aryShifts(0, 2) = "06:00"
    l_aryShifts(1, 0) = ""
    l_aryShifts(1, 1) = "06:00"
    l_aryShifts(1, 2) = "14:00"
    l_aryShifts(2, 0) = "P"
    l_aryShifts(2, 1) = "14:00"
    l_aryShifts(2, 2) = "24:00"
    l_aryShifts(3, 0) = "P"
    l_aryShifts(3, 1) = "24:00"
    l_aryShifts(3, 2) = "24:00"
    
    GetCILCOLocal8MaterialHandlerShifts2 = l_aryShifts
    

  Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetCILCOLocal8MaterialHandlerShifts2"
    
End Function

Private Sub ShiftComponentSplitsDetermine( _
    ByRef r_aryShifts As Variant, _
    ByRef r_objTimeEntry As cboTimecardEntry, _
    ByRef r_colTimeEntries As Collection)
    
    
'Please note that time that crosses midnight will be split prior to this method
'Start time will always be less than end time

    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_aryShifts(3, 2) As String
    Dim l_strCompStartTime As String
    Dim l_strCompEndTime As String
    Dim l_strShiftComp As String
    Dim l_blnContinue As Boolean
    Dim l_decHours As Variant
    Dim l_strSystemExceptionCode As String
    Dim l_curTimeEntry As cboTimecardEntry
    Dim l_NewTimeEntry As cboTimecardEntry

    
    On Error GoTo eh
    
    Set l_curTimeEntry = r_objTimeEntry
    l_strStartTime = l_curTimeEntry.startTime
    l_strEndTime = l_curTimeEntry.EndTime
    If l_strEndTime = "00:00" Then
        l_strEndTime = "24:00"
    End If
    
    l_blnContinue = True
    While l_blnContinue
        GetShiftCompAndTimes l_strStartTime, l_strShiftComp, l_strCompStartTime, l_strCompEndTime, r_aryShifts
        l_curTimeEntry.ShiftComponentCode = l_strShiftComp
        
        If l_strEndTime > l_strCompEndTime Then
'JIMB Experiment
'I need to create a new timeentry object
'            l_curTimeEntry = l_curTimeEntry.Split(l_strCompEndTime, r_colTimeEntries)
            Set l_NewTimeEntry = CreateNewTimeEntry(l_curTimeEntry)
            l_NewTimeEntry.startTime = l_strStartTime
            l_NewTimeEntry.EndTime = l_strCompEndTime
            l_decHours = DetermineHours(l_strStartTime, l_strCompEndTime)
            l_NewTimeEntry.Hours = l_decHours
'            AddTimecardEntry l_curTimeEntry
            r_colTimeEntries.Add l_NewTimeEntry
            l_strStartTime = l_strCompEndTime
        Else
            l_curTimeEntry.startTime = l_strStartTime
            l_curTimeEntry.EndTime = l_strEndTime
            l_decHours = DetermineHours(l_strStartTime, l_strEndTime)
            l_curTimeEntry.Hours = l_decHours
            r_colTimeEntries.Add l_curTimeEntry
            l_blnContinue = False
        End If
    Wend
        
    Exit Sub

eh:
    TrsRaiseError mstrCMODULENAME, "ShiftComponentSplitsDetermine"
    
    
End Sub


Private Function DetermineHours( _
    ByVal v_strStartTime As String, _
    ByVal v_strEndTime As String _
    ) As Variant
    
    Dim l_decHours As Variant
    
    On Error GoTo eh
    
    If v_strEndTime <> "24:00" Then
        l_decHours = (Hour(v_strEndTime) + Minute(v_strEndTime) / 60) - _
            (Hour(v_strStartTime) + Minute(v_strStartTime) / 60)
    Else
        l_decHours = 24 - (Hour(v_strStartTime) + Minute(v_strStartTime) / 60)
    End If
    
    DetermineHours = Round(l_decHours, 1)
    
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "DetermineHours"


End Function

Private Sub GetShiftCompAndTimes( _
    ByVal v_strStartTime As String, _
    ByRef r_strShiftComp As String, _
    ByRef r_strCompStartTime As String, _
    ByRef r_strCompEndTime As String, _
    ByRef r_aryShifts As Variant)
    
'r_aryShifts is a 2 demensional array of defined shifts for shift component
'Should by defined as Dim r_aryShifts(3,2) as String
'The 1st element is the shift component, 2nd element is start time, the 3rd is end time

    
    Dim l_lngLow As Long
    Dim l_lngHigh As Long
    Dim l_lngI As Long
    Dim l_strCompStartTime As String
    Dim l_strCompEndTime As String
    Dim l_strShiftComp As String
    
   On Error GoTo eh
    
    l_lngLow = LBound(r_aryShifts)
    l_lngHigh = UBound(r_aryShifts)
    
    For l_lngI = l_lngLow To l_lngHigh
        r_strShiftComp = r_aryShifts(l_lngI, 0)
        r_strCompStartTime = r_aryShifts(l_lngI, 1)
        r_strCompEndTime = r_aryShifts(l_lngI, 2)
        If v_strStartTime < r_strCompEndTime Then
            Exit For
        End If
    Next l_lngI
    
    
    Exit Sub

eh:
    TrsRaiseError mstrCMODULENAME, "GetShiftCompAndTimes"
    
End Sub

Private Function CreateNewTimeEntry( _
    ByVal v_objTimeEntry As cboTimecardEntry _
    ) As cboTimecardEntry
    
    On Error GoTo eh
    
    Dim newEntry As cboTimecardEntry
    Dim lcolErrArgs As New Collection
    
    Set newEntry = CreateObject("busTimecard.cboTimecardEntry")
    newEntry.SetNew
    
    AddTimecardEntry newEntry

    With newEntry
        .EmployeeNumber = v_objTimeEntry.EmployeeNumber
        .ChargeDate = v_objTimeEntry.ChargeDate
        .DateWorked = v_objTimeEntry.DateWorked
        .startTime = v_objTimeEntry.startTime
        .EndTime = v_objTimeEntry.EndTime
        .DepartmentNumber = v_objTimeEntry.DepartmentNumber
        .ExceptionCode = v_objTimeEntry.ExceptionCode
        .DistributionId = v_objTimeEntry.DistributionId
        .DistributionType = v_objTimeEntry.DistributionType
        .OverrideExceptionCode = v_objTimeEntry.OverrideExceptionCode
        .DepartmentNumber = v_objTimeEntry.DepartmentNumber
        .DistributionDepartmentNumber = v_objTimeEntry.DistributionDepartmentNumber
        .WorkgroupNumber = v_objTimeEntry.WorkgroupNumber
        .JobNumber = v_objTimeEntry.JobNumber
        .ShiftComponentCode = v_objTimeEntry.ShiftComponentCode
        .FamilyLeaveIndicator = v_objTimeEntry.FamilyLeaveIndicator
        .ScheduleChangeExceptionCode = v_objTimeEntry.ScheduleChangeExceptionCode
        .ComponentCode = v_objTimeEntry.ComponentCode
        .ComponentAmount = v_objTimeEntry.ComponentAmount
        .OvertimeReasonCode = v_objTimeEntry.OvertimeReasonCode
        .PreArrangedOTIndicator = v_objTimeEntry.PreArrangedOTIndicator
        .CaseNumber = v_objTimeEntry.CaseNumber
        .ReliefCode = v_objTimeEntry.ReliefCode
        .OvertimeOpportunityNumber = v_objTimeEntry.OvertimeOpportunityNumber
        .OvertimeOpportunityDescription = v_objTimeEntry.OvertimeOpportunityDescription
        .ApprovedBy = v_objTimeEntry.ApprovedBy
        .ApprovedDate = v_objTimeEntry.ApprovedDate
        .ApprovalIndicator = v_objTimeEntry.ApprovalIndicator
        .ScheduleTimeIndicator = v_objTimeEntry.ScheduleTimeIndicator
        .GroupTimestamp = v_objTimeEntry.GroupTimestamp
        .ScheduleCode = v_objTimeEntry.ScheduleCode
        .ScheduleSourceCode = v_objTimeEntry.ScheduleSourceCode
        .JobNbr = v_objTimeEntry.JobNbr
        .JobTaskNbr = v_objTimeEntry.JobTaskNbr
        .JobTaskStepNbr = v_objTimeEntry.JobTaskStepNbr
        .EMPRVDatabaseID = v_objTimeEntry.EMPRVDatabaseID
    End With
    
    Set CreateNewTimeEntry = newEntry
    
    Exit Function
    
    
eh:
    TrsRaiseError mstrCMODULENAME, "CreateNewTimeEntry"
End Function

Private Sub Class_Initialize()
    mstrCaller = "T"
End Sub

Private Function GetStandardShifts() As Variant

    Dim l_aryShifts(3, 2) As String

    
    On Error GoTo eh
    
    l_aryShifts(0, 0) = "O"
    l_aryShifts(0, 1) = "00:00"
    l_aryShifts(0, 2) = "07:00"
    l_aryShifts(1, 0) = ""
    l_aryShifts(1, 1) = "07:00"
    l_aryShifts(1, 2) = "15:00"
    l_aryShifts(2, 0) = "P"
    l_aryShifts(2, 1) = "15:00"
    l_aryShifts(2, 2) = "23:00"
    l_aryShifts(3, 0) = "O"
    l_aryShifts(3, 1) = "23:00"
    l_aryShifts(3, 2) = "24:00"
    
    GetStandardShifts = l_aryShifts
    

  Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetStandardShifts"
    
End Function

Private Function GetCoffeen12HourShifts() As Variant

    Dim l_aryShifts(2, 2) As String

    
    On Error GoTo eh
    
    l_aryShifts(0, 0) = "C12"
    l_aryShifts(0, 1) = "00:00"
    l_aryShifts(0, 2) = "07:00"
    l_aryShifts(1, 0) = ""
    l_aryShifts(1, 1) = "07:00"
    l_aryShifts(1, 2) = "19:00"
    l_aryShifts(2, 0) = "C12"
    l_aryShifts(2, 1) = "19:00"
    l_aryShifts(2, 2) = "24:00"
    
    GetCoffeen12HourShifts = l_aryShifts
    

  Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetCoffeen12HourShifts"
    
End Function


Private Function GetMaintenanceShifts() As Variant

    Dim l_aryShifts(3, 2) As String

    
    On Error GoTo eh
    
    l_aryShifts(0, 0) = "O"
    l_aryShifts(0, 1) = "00:00"
    l_aryShifts(0, 2) = "07:00"
    l_aryShifts(1, 0) = ""
    l_aryShifts(1, 1) = "07:00"
    l_aryShifts(1, 2) = "15:30"
    l_aryShifts(2, 0) = "P"
    l_aryShifts(2, 1) = "15:30"
    l_aryShifts(2, 2) = "23:00"
    l_aryShifts(3, 0) = "O"
    l_aryShifts(3, 1) = "23:00"
    l_aryShifts(3, 2) = "24:00"
    
    GetMaintenanceShifts = l_aryShifts
    

  Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetMaintenanceShifts"
    
End Function

Private Function GetCallawaySecurity12HourShifts() As Variant

    Dim l_aryShifts(2, 2) As String
    
    On Error GoTo eh
    
    l_aryShifts(0, 0) = ""
    l_aryShifts(0, 1) = "00:00"
    l_aryShifts(0, 2) = "06:00"
    l_aryShifts(1, 0) = ""
    l_aryShifts(1, 1) = "06:00"
    l_aryShifts(1, 2) = "18:00"
    l_aryShifts(2, 0) = ""
    l_aryShifts(2, 1) = "18:00"
    l_aryShifts(2, 2) = "24:00"
    
    GetCallawaySecurity12HourShifts = l_aryShifts
    

  Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "GetCallawaySecurity12HourShifts"
    
End Function

Public Function IsDateASpecialOverride( _
            ByVal v_datDate As Date _
            ) As Boolean
'Release 03/27/2020 -- JFB  Added this function to get to see if
'                             vacation was reported on a day
            
    Dim tce As cboTimecardEntry
    Dim l_blnIsASpecialOverride As Boolean
  
On Error GoTo eh
    l_blnIsASpecialOverride = False
    
' Schedule provided for time collection does not have exceptions, so it will always
' check time entries for answer.  Schedule provided from overtime will have scheduled
' exceptions.
    If Me.Schedule.IsScheduledSpecialOverride(v_datDate) Then
        l_blnIsASpecialOverride = True
        GoTo xit
    End If
    
    For Each tce In TimecardEntries
        If IsSpecialOverrideCode(tce.OverrideExceptionCode) And tce.ChargeDate = v_datDate Then
            l_blnIsASpecialOverride = True
            Exit For
        End If
    Next
    
xit:
    IsDateASpecialOverride = l_blnIsASpecialOverride
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsDateASpecialOverride"
            
End Function




Public Function IsDateAHolidayBanked( _
            ByVal v_datDate As Date _
            ) As Boolean

            
    Dim tce As cboTimecardEntry
    Dim l_blnIsAHolidayBanked  As Boolean
  
On Error GoTo eh
    IsDateAHolidayBanked = False

    
    For Each tce In TimecardEntries
        If Trim(tce.ExceptionCode) = "HB" And tce.ChargeDate = v_datDate Then
            l_blnIsAHolidayBanked = True
            Exit For
        End If
    Next
    
xit:
    IsDateAHolidayBanked = l_blnIsAHolidayBanked
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsDateAHolidayBanked"
            
End Function


Public Function DoesOvertimeOverlapWithHolidayBanked( _
            ByRef v_Entry As cboTimecardEntry) As Boolean

    Dim colCurrentOverlap As Collection
    Dim l_OverlapEntry As cboTimecardEntry
    
    On Error GoTo eh

    Set colCurrentOverlap = GetBSTimecard.FindOverlappingEntries(v_Entry, mcolTimecardEntries)
    If colCurrentOverlap.Count = 0 Then
        DoesOvertimeOverlapWithHolidayBanked = False
        Exit Function
    End If
    Set l_OverlapEntry = colCurrentOverlap.Item(1)
    If Trim(l_OverlapEntry.ExceptionCode) = "HB" Then
        DoesOvertimeOverlapWithHolidayBanked = True
    Else
        DoesOvertimeOverlapWithHolidayBanked = False
    End If
    
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "DoesOvertimeOverlapWithHolidayBanked"

End Function

Public Function BuildMAXIMOJobTasks() As ADODB.Recordset

    Dim l_rsMAXIMOJobTasks As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim l_strDept As String
    Dim l_datChargeDate As Date
    Dim l_strRecStatus As String
    Dim l_strException As String
    Dim l_strOverExc As String
    Dim l_strOrigException As String
    Dim l_strOrigOverExc As String
    Dim l_bolChanged As Boolean
    
On Error GoTo eh

    Set rs = TimecardEntryRecordset
    rs.Filter = "(RECORD_STATUS = 'D' AND DISTRIBUTION_TYPE = 'M') OR (RECORD_STATUS = 'I' AND DISTRIBUTION_TYPE = 'M') OR (RECORD_STATUS = 'C')"
    Set l_rsMAXIMOJobTasks = New ADODB.Recordset
    With l_rsMAXIMOJobTasks
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
        .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "JOB_NBR", adVarChar, 8, adFldUpdatable
        .Fields.Append "JOBTASK_NBR", adVarChar, 4, adFldUpdatable
        .Fields.Append "JOBTASK_STEP_NBR", adInteger, , adFldUpdatable
        .Fields.Append "EMPRV_DATABASE_ID", adVarChar, 8, adFldUpdatable
        .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
        .Fields.Append "HOURS", adDecimal, , adFldUpdatable
        .Fields.Item("HOURS").Precision = 3
        .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "OT_IND", adChar, 1, adFldUpdatable
        .Fields.Append "CHG_IND", adChar, 1, adFldUpdatable
        .Fields.Append "DEPT_NUM", adChar, 3, adFldUpdatable
        .Fields.Append "SRC_ROW_ID", adDecimal, , adFldUpdatable
        .Fields.Item("SRC_ROW_ID").Precision = 20
        .Fields.Item("SRC_ROW_ID").NumericScale = 0
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open

        While Not rs.EOF
            l_datChargeDate = rs!CHARGE_DATE
            l_strDept = mobjEmployee.DepartmentNumber(l_datChargeDate)
'            If IsJobTaskEmployee(l_strDept) Then
                l_strRecStatus = rs!RECORD_STATUS
                Select Case l_strRecStatus
                    Case "I"
                        .AddNew
                        .Fields("CHG_IND") = "N"
                        .Fields("EMP_NUM") = rs.Fields("EMP_NUM")
                        .Fields("DATE_WORKED") = rs.Fields("DATE_WORKED")
                        .Fields("START_TIME") = rs.Fields("START_TIME")
                        .Fields("END_TIME") = rs.Fields("END_TIME")
                        .Fields("JOB_NBR") = rs.Fields("JOB_NBR")
                        .Fields("JOBTASK_NBR") = rs.Fields("JOBTASK_NBR")
                        .Fields("JOBTASK_STEP_NBR") = rs.Fields("JOBTASK_STEP_NBR")
                        .Fields("EXCEPTION_CD") = rs.Fields("EXCEPTION_CD")
                        .Fields("HOURS") = rs.Fields("HOURS")
                        .Fields("SRC_ROW_ID") = rs.Fields("ROW_ID")
                        .Fields("EMPRV_DATABASE_ID") = rs.Fields("EMPRV_DATABASE_ID")
                        If IsJobTaskOTExcep(.Fields("EXCEPTION_CD")) Then
                            .Fields("OT_IND") = "Y"
                        Else
                            .Fields("OT_IND") = "N"
                        End If
                        .Fields("DEPT_NUM") = rs.Fields("DEPT_NUM")
                        .Update
                    Case "D"
                        .AddNew
                        .Fields("CHG_IND") = "D"
                        .Fields("EMP_NUM") = rs.Fields("EMP_NUM")
                        .Fields("DATE_WORKED") = rs.Fields("DATE_WORKED")
                        .Fields("START_TIME") = rs.Fields("START_TIME")
                        .Fields("END_TIME") = rs.Fields("END_TIME")
                        .Fields("JOB_NBR") = rs.Fields("JOB_NBR")
                        .Fields("JOBTASK_NBR") = rs.Fields("JOBTASK_NBR")
                        .Fields("EXCEPTION_CD") = rs.Fields("EXCEPTION_CD")
                        .Fields("HOURS") = rs.Fields("HOURS")
                        .Fields("SRC_ROW_ID") = rs.Fields("ROW_ID")
                        .Fields("EMPRV_DATABASE_ID") = rs.Fields("EMPRV_DATABASE_ID")
                        If IsJobTaskOTExcep(.Fields("EXCEPTION_CD")) Then
                            .Fields("OT_IND") = "Y"
                        Else
                            .Fields("OT_IND") = "N"
                        End If
                        .Fields("DEPT_NUM") = rs.Fields("DEPT_NUM")
                        .Update
                    Case "C"
                        If rs.Fields("DISTRIBUTION_TYPE").Value = "M" And rs.Fields("DISTRIBUTION_TYPE").OriginalValue <> "M" Then
                        '' wasn't MAXIMO before
                            .AddNew
                            .Fields("EMP_NUM") = rs.Fields("EMP_NUM").Value
                            .Fields("DATE_WORKED") = rs.Fields("DATE_WORKED").Value
                            .Fields("START_TIME") = rs.Fields("START_TIME").Value
                            .Fields("END_TIME") = rs.Fields("END_TIME").Value
                            .Fields("JOB_NBR") = rs.Fields("JOB_NBR").Value
                            .Fields("JOBTASK_NBR") = rs.Fields("JOBTASK_NBR").Value
                            .Fields("JOBTASK_STEP_NBR") = rs.Fields("JOBTASK_STEP_NBR").Value
                            .Fields("EXCEPTION_CD") = rs.Fields("EXCEPTION_CD").Value
                            .Fields("HOURS") = rs.Fields("HOURS").Value
                            .Fields("SRC_ROW_ID") = rs.Fields("ROW_ID")
                            .Fields("EMPRV_DATABASE_ID") = rs.Fields("EMPRV_DATABASE_ID").Value
                            If IsJobTaskOTExcep(.Fields("EXCEPTION_CD").Value) Then
                                .Fields("OT_IND") = "Y"
                            Else
                                .Fields("OT_IND") = "N"
                            End If
                            .Fields("CHG_IND") = "I"
                            .Fields("DEPT_NUM") = rs.Fields("DEPT_NUM").Value
                            .Update
                        End If
                        If rs.Fields("DISTRIBUTION_TYPE").OriginalValue = "M" And rs.Fields("DISTRIBUTION_TYPE").Value <> "M" Then
                        '' was MAXIMO before and is not now
                            .AddNew
                            .Fields("EMP_NUM") = rs.Fields("EMP_NUM").OriginalValue
                            .Fields("DATE_WORKED") = rs.Fields("DATE_WORKED").OriginalValue
                            .Fields("START_TIME") = rs.Fields("START_TIME").OriginalValue
                            .Fields("END_TIME") = rs.Fields("END_TIME").OriginalValue
                            .Fields("JOB_NBR") = rs.Fields("JOB_NBR").OriginalValue
                            .Fields("JOBTASK_NBR") = rs.Fields("JOBTASK_NBR").OriginalValue
                            .Fields("JOBTASK_STEP_NBR") = rs.Fields("JOBTASK_STEP_NBR").OriginalValue
                            .Fields("EXCEPTION_CD") = rs.Fields("EXCEPTION_CD").OriginalValue
                            .Fields("HOURS") = rs.Fields("HOURS").OriginalValue
                            .Fields("SRC_ROW_ID") = rs.Fields("ROW_ID")
                            .Fields("EMPRV_DATABASE_ID") = rs.Fields("EMPRV_DATABASE_ID").OriginalValue
                            If IsJobTaskOTExcep(.Fields("EXCEPTION_CD").OriginalValue) Then
                                .Fields("OT_IND") = "Y"
                            Else
                                .Fields("OT_IND") = "N"
                            End If
                            .Fields("CHG_IND") = "D"
                            .Fields("DEPT_NUM") = rs.Fields("DEPT_NUM").Value
                            .Update
                        End If
                        If rs.Fields("DISTRIBUTION_TYPE").Value = "M" And rs.Fields("DISTRIBUTION_TYPE").OriginalValue = "M" Then
                            '' determine if there was a change in job number, task, hours, or start time
                            l_bolChanged = False
                            If rs.Fields("JOB_NBR").Value <> rs.Fields("JOB_NBR").OriginalValue _
                              Or rs.Fields("JOBTASK_NBR").Value <> rs.Fields("JOBTASK_NBR").OriginalValue _
                              Or rs.Fields("HOURS").Value <> rs.Fields("HOURS").OriginalValue _
                              Or rs.Fields("START_TIME").Value <> rs.Fields("START_TIME").OriginalValue Then
                                .AddNew
                                .Fields("EMP_NUM") = rs.Fields("EMP_NUM").Value
                                .Fields("DATE_WORKED") = rs.Fields("DATE_WORKED").Value
                                .Fields("START_TIME") = rs.Fields("START_TIME").Value
                                .Fields("END_TIME") = rs.Fields("END_TIME").Value
                                .Fields("JOB_NBR") = rs.Fields("JOB_NBR").Value
                                .Fields("JOBTASK_NBR") = rs.Fields("JOBTASK_NBR").Value
                                .Fields("JOBTASK_STEP_NBR") = rs.Fields("JOBTASK_STEP_NBR").Value
                                .Fields("EXCEPTION_CD") = rs.Fields("EXCEPTION_CD").Value
                                .Fields("HOURS") = rs.Fields("HOURS").Value
                                .Fields("SRC_ROW_ID") = rs.Fields("ROW_ID")
                                .Fields("EMPRV_DATABASE_ID") = rs.Fields("EMPRV_DATABASE_ID").Value
                                If IsJobTaskOTExcep(.Fields("EXCEPTION_CD").Value) Then
                                    .Fields("OT_IND") = "Y"
                                Else
                                    .Fields("OT_IND") = "N"
                                End If
                                .Fields("CHG_IND") = "C"
                                .Fields("DEPT_NUM") = rs.Fields("DEPT_NUM").Value
                                .Update
                            End If
                        End If
                End Select
'            End If
            rs.MoveNext
        Wend
    End With
    rs.Filter = ""

    Set BuildMAXIMOJobTasks = l_rsMAXIMOJobTasks
    Exit Function
    
eh:
     TrsRaiseError mstrCMODULENAME, "BuildMAXIMOJobTasks"

End Function
