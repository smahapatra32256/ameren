VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cboSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Currently, creating local shift objects as needed, but we
'may want to cache them if it will improve performance.
'Balance with increase resources required.

Private Const mstrCMODULENAME As String = "cboSchedule"
Private mobjScheduleState As ADODB.Recordset
Private mobjScheduleChangePayments As ADODB.Recordset
Private mobjTimecard As busTimecard.cboTimecard
Private mstrEmpNum As Long
Private mobjAvailSched As ADODB.Recordset
Private mstrScheduleTimeStamp As String

Public Function SetTimecard(ByRef robjTimecard As busTimecard.cboTimecard)
    Set mobjTimecard = robjTimecard
End Function

Public Function Timecard() As busTimecard.cboTimecard
    Set Timecard = mobjTimecard
End Function

Public Function Employee() As boCommonLib.cboEmployee
    Set Employee = Timecard.Employee
End Function


Public Function GetScheduledHours( _
            ByVal v_ChgDate As Date, _
            Optional ByVal v_datCutoff As Variant) As Variant
'  Get the Scheduled Hours for this Charge Date.
'  v_datCutoff should be a datetime, where date represents the
'  actual work date.
    Dim l_objShifts As Collection
    Dim l_objShift As cboScheduledShift
    Dim l_decShiftHours As Variant
    Dim l_datCutoff As Variant

On Error GoTo GetScheduledHours_EH
    
    If IsMissing(v_datCutoff) Then
        l_datCutoff = Empty
    Else
        l_datCutoff = v_datCutoff
    End If

    '  Call ShiftsOn to get a collection of shifts for the specified Charge Date
    Set l_objShifts = ShiftsOn(v_ChgDate)
    
    '  Get the total Shift Hours for the Shifts for the Charge Date
    l_decShiftHours = CDec(0)
    For Each l_objShift In l_objShifts
        If l_objShift.IsScheduledWorkDay Then
            If IsEmpty(l_datCutoff) Then
                l_decShiftHours = l_decShiftHours + l_objShift.ShiftHours
            Else
                If l_datCutoff >= l_objShift.ShiftStartDateTime Then
                    l_decShiftHours = l_decShiftHours + l_objShift.ShiftHours
                End If
            End If
        End If
    Next
    
    GetScheduledHours = l_decShiftHours
    Exit Function
    
GetScheduledHours_EH:
    Set l_objShifts = Nothing
    Set l_objShift = Nothing
    TrsRaiseError mstrCMODULENAME, "GetScheduledHours"

End Function



' 7/26/00
' ShiftOn should be FirstShiftOn since it only answers the first
' shift regardless of if there is a second.
' The following methods currently use ShiftOn, but they
' need to account for multiple shifts on a day:
'  IsScheduledFirstDayOff
'  IsScheduledSecondDayOff
'  IsScheduledWorkDay
'  ShiftStart
'  ShiftEnd
' 7/27/00  Maybe its taken care of now?
'
'Public Function ShiftOn(chgDate As Date) As cboScheduledShift
'    Dim shift As cboScheduledShift
'    Dim l_colShifts As Collection
'
'    Set l_colShifts = ShiftsOn(chgDate)
'    If l_colShifts.Count = 0 Then
'        Err.Raise vbObjectError + 1500, "No shift for charge date: " & Format(chgDate, "MM/DD/YYYY") & " found."
'    End If
'    Set ShiftOn = l_colShifts(1)
'
'End Function
'
'
'Private Sub PositionRecordsetOn(chgDate As Date)
'
'    mobjScheduleState.MoveFirst
'    mobjScheduleState.Find "CHARGE_DATE = " & Format(chgDate, "mm/dd/yyyy")
'    If mobjScheduleState.EOF Then
'        Err.Raise vbObjectError + 1500, "No shift for charge date: " & Format(chgDate, "MM/DD/YYYY") & " found."
'    End If
'
'End Sub

'
'Public Function LunchPeriod(chgDate As Date) As Collection
''Answer a collection with lunch start first, and lunch end second.
'
''This must be handle multiple shifts, so provide method
''for timecard entry to ask about a lunch break for a specific
''shift.
'
'    Dim eShift As cboScheduledShift
'
'    Set eShift = ShiftOn(chgDate)
'    Set LunchPeriod = eShift.LunchPeriod
'
'End Function

Private Function CurrentShift() As cboScheduledShift
    Dim shift As cboScheduledShift
    Dim l_blnLunchFound As Boolean
    Dim l_blnNextScheduleFound As Boolean

    Set shift = CreateNewShift
    shift.ShiftChargeDate = mobjScheduleState("CHARGE_DATE")
    shift.ShiftWorkDate = mobjScheduleState("DATE_WORKED")
    shift.ShiftDay = Weekday(mobjScheduleState("CHARGE_DATE"), vbSunday)
    shift.ScheduleCode = mobjScheduleState("SCHED_CD")
    shift.ShiftStart = mobjScheduleState("SHIFT_START_TIME")
    shift.ShiftEnd = mobjScheduleState("SHIFT_END_TIME")
    shift.ShiftHours = mobjScheduleState("SHIFT_HOURS")
    shift.ShiftComponentCode = mobjScheduleState("COMP_SHIFT_CD")
    shift.ShiftCode = mobjScheduleState("SHIFT_CODE")
    shift.SystemExceptionCode = mobjScheduleState("SYSTEM_EXC_CD")
    shift.ReliefCode = mobjScheduleState("REL_CD")
    shift.ScheduleSource = mobjScheduleState("SCHED_SRC_CD")
    shift.ScheduleForShiftComp = mobjScheduleState("SCHED_CD_FOR_SHIFT_COMP")
    shift.LineNumber = mobjScheduleState("LINE_NUM")
    shift.RequiredLineInd = mobjScheduleState("REQ_LINE_IND")
    shift.ShiftNumber = mobjScheduleState("SHIFT_NUM")
    shift.ExceptionCode = mobjScheduleState("EXCEPTION_CD")
    shift.OverrideExceptionCode = mobjScheduleState("OVER_EXC_CD")

'    Check for lunch records.
'    Assuming first record is schedule work time.
'    Scan to next schedule or eof

    shift.ShiftLunchHours = 0   'Intializing Lunch time to zero hours
                                ' incase the shift does not have a lunch
    
    Do Until mobjScheduleState.EOF
        If mobjScheduleState!SCHED_CD <> shift.ScheduleCode Then
            Exit Do
        ElseIf mobjScheduleState!SYSTEM_EXC_CD = "SLT" Then
            shift.ShiftLunchStart = mobjScheduleState("START_TIME")
            shift.ShiftLunchEnd = mobjScheduleState("END_TIME")
            shift.ShiftLunchHours = mobjScheduleState("HOURS")
        End If
        mobjScheduleState.MoveNext
    Loop

    Set CurrentShift = shift
    
End Function


Public Function ScheduledShiftsOn(v_datChgDate As Date) As Collection
    Dim lcolSchedShifts As Collection
    Dim lcolShifts As Collection
    Dim eShift As cboScheduledShift

On Error GoTo eh
    Set lcolShifts = ShiftsOn(v_datChgDate)
    Set lcolSchedShifts = New Collection
    For Each eShift In lcolShifts
        If eShift.IsScheduledWorkDay Then
            lcolSchedShifts.Add eShift
        End If
    Next
    Set ScheduledShiftsOn = lcolSchedShifts
    Set lcolSchedShifts = Nothing
    Set lcolShifts = Nothing
    Set eShift = Nothing
    Exit Function

eh:
    Set lcolSchedShifts = Nothing
    Set lcolShifts = Nothing
    Set eShift = Nothing
    TrsRaiseError mstrCMODULENAME, "ScheduledShiftsOn"

End Function


Private Function ShiftsOn(v_datChgDate As Date) As Collection
' Employee may be scheduled for upto 2 shifts per charge date.
' Mask all schedule records from view for dates other than v_datChgDate.
'
    Dim l_objShift As cboScheduledShift
    Dim l_colShifts As Collection

    Set l_colShifts = New Collection
    SetScheduleFilterOn v_datChgDate
    If mobjScheduleState.EOF Then
        Set ShiftsOn = l_colShifts
        Exit Function
    End If

    mobjScheduleState.Sort = "SHIFT_NUM"
    mobjScheduleState.MoveFirst
    Do Until mobjScheduleState.EOF
        l_colShifts.Add CurrentShift
    Loop
    
    Set ShiftsOn = l_colShifts
    
End Function

Private Sub SetScheduleFilterOn(v_datChgDate)
    Dim l_strFilter As String
    
    l_strFilter = "EMP_NUM = " & mstrEmpNum & " AND " & _
        "CHARGE_DATE = " & _
        "#" & Format(v_datChgDate, "MM/DD/YYYY") & "#"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    mobjScheduleState.Filter = l_strFilter

End Sub

Private Sub ClearScheduleFilter()

    mobjScheduleState.Filter = "EMP_NUM = " & mstrEmpNum

End Sub

Private Sub SetScheduleChangePmtFilterOn(v_datChgDate)
    Dim l_strFilter As String
    
    l_strFilter = "EMP_NUM = " & mstrEmpNum & " AND " & _
        "CHARGE_DATE = " & _
        "#" & Format(v_datChgDate, "MM/DD/YYYY") & "#"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    mobjScheduleChangePayments.Filter = l_strFilter

End Sub

Private Sub ClearScheduleChangePmtFilter()

    mobjScheduleChangePayments.Filter = "EMP_NUM = " & mstrEmpNum

End Sub

Public Function IsScheduledFirstDayOff(ByVal v_datChargeDate As Date) As Boolean
' Is any shift on chgDate a work day and sunday?
    Dim shifts As Collection
    Dim eShift As cboScheduledShift
    
    Set shifts = ShiftsOn(v_datChargeDate)
    For Each eShift In shifts
        If eShift.IsScheduledFirstDayOff Then
            IsScheduledFirstDayOff = True
            Exit Function
        End If
    Next
    
    Set eShift = Nothing
    Set shifts = Nothing
        
End Function

Public Function IsScheduledSecondDayOff(ByVal v_datChargeDate As Date) As Boolean
' Is any shift on chgDate a work day and saturday?
    Dim shifts As Collection
    Dim eShift As cboScheduledShift
    
    Set shifts = ShiftsOn(v_datChargeDate)
    For Each eShift In shifts
        If eShift.IsScheduledSecondDayOff Then
            IsScheduledSecondDayOff = True
            Exit Function
        End If
    Next
    
End Function

Public Function IsScheduledWorkDay(ByVal v_datChargeDate As Date) As Boolean
' Is any shift on chgDate a work day.
    Dim shifts As Collection
    Dim eShift As cboScheduledShift
    
    Set shifts = ShiftsOn(v_datChargeDate)
    For Each eShift In shifts
        If eShift.IsScheduledWorkDay Then
            IsScheduledWorkDay = True
            Exit Function
        End If
    Next
    
End Function


Public Function IsScheduledHolidayOrVacation( _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_datStartTime As Date, _
                    ByVal v_datEndTime As Date) As Boolean

    Dim l_objShift As cboScheduledShift
    
    Set l_objShift = ShiftScheduledDuring( _
            v_datChargeDate, v_datStartTime, v_datEndTime)
    If l_objShift Is Nothing Then
        IsScheduledHolidayOrVacation = False
    Else
        IsScheduledHolidayOrVacation = _
            l_objShift.IsScheduledHoliday Or l_objShift.IsScheduledVacation Or l_objShift.IsScheduledSpecialOverride
    End If
    Set l_objShift = Nothing

End Function


Public Function IsScheduledHoliday(ByVal v_datChargeDate As Date) As Boolean
' Is any shift on chgDate a holiday.
    Dim shifts As Collection
    Dim eShift As cboScheduledShift
    
    Set shifts = ShiftsOn(v_datChargeDate)
    For Each eShift In shifts
        If eShift.IsScheduledHoliday Then
            IsScheduledHoliday = True
            Exit Function
        End If
    Next
    
End Function


Public Function IsScheduledVacation(ByVal v_datChargeDate As Date) As Boolean
' Is any shift on chgDate a vacation.
    Dim shifts As Collection
    Dim eShift As cboScheduledShift
    
    Set shifts = ShiftsOn(v_datChargeDate)
    For Each eShift In shifts
        If eShift.IsScheduledVacation Then
            IsScheduledVacation = True
            Exit Function
        End If
    Next
    
End Function


Public Function IsAnyPeriodWithinSchedule( _
            ByVal v_datChargeDate As Date, _
            ByVal v_datStartTime As Date, _
            ByVal v_datEndTime As Date) As Boolean
'Lunch time is outside of schedule

    Dim eShift As cboScheduledShift
    Dim l_blnWithinSchedule As Boolean
    Dim l_colLunchTimes As Collection
    
On Error GoTo eh
    Set eShift = ShiftScheduledDuring(v_datChargeDate, v_datStartTime, v_datEndTime)
    l_blnWithinSchedule = Not eShift Is Nothing
    If l_blnWithinSchedule Then
        If eShift.IsScheduledLunch Then
            If Not IsFloatingLunchEmployee(v_datChargeDate) Then
                Set l_colLunchTimes = eShift.LunchPeriodAsDates
                If IsTimeDuringLunch(v_datStartTime, v_datEndTime, l_colLunchTimes(1), l_colLunchTimes(2)) Then
                    l_blnWithinSchedule = False
                End If
            End If
        End If
    End If
    IsAnyPeriodWithinSchedule = l_blnWithinSchedule
    Exit Function
    
eh:
    Set l_colLunchTimes = Nothing
    Set eShift = Nothing
    TrsRaiseError mstrCMODULENAME, "IsAnyPeriodWithinSchedule"
    
End Function


Private Function IsFloatingLunchEmployee(ByVal v_datChargeDate As Date) As Boolean
    
On Error GoTo eh
    
    IsFloatingLunchEmployee = _
        Timecard.Employee.IsLunchFloater(v_datChargeDate)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsFloatingLunchEmployee"
End Function


Public Function IsAnyPeriodOutsideSchedule( _
        ByVal v_datChargeDate As Date, _
        ByVal v_datStartTime As Date, _
        ByVal v_datEndTime As Date) As Boolean
    Dim shifts As Collection
    Dim eShift As cboScheduledShift
    
On Error GoTo eh
    Set shifts = ShiftsOn(v_datChargeDate)
    If shifts.Count = 0 Then
        IsAnyPeriodOutsideSchedule = True
        Exit Function
    End If
    
    For Each eShift In shifts
        If eShift.IsScheduledWorkDay Then
            If (v_datStartTime < eShift.ShiftStartDateTime) Or _
                (v_datEndTime > eShift.ShiftEndDateTime) Then
                IsAnyPeriodOutsideSchedule = True
                Exit For
            End If
        End If
    Next
    Exit Function

eh:
    Set shifts = Nothing
    Set eShift = Nothing
    TrsRaiseError mstrCMODULENAME, "IsAnyPeriodOutsideSchedule", Err.Number

End Function

Public Function IsSelfTradedShift( _
        ByVal v_datChargeDate As Date, _
        ByVal v_datStartTime As Date, _
        ByVal v_datEndTime As Date) As Boolean
        
    Dim l_objShift As cboScheduledShift
    
    Set l_objShift = ShiftScheduledDuring( _
            v_datChargeDate, v_datStartTime, v_datEndTime)
    If l_objShift Is Nothing Then
        IsSelfTradedShift = False
    Else
        IsSelfTradedShift = l_objShift.IsSelfTrade
    End If
    Set l_objShift = Nothing

End Function

Public Function IsTradedWorkShift( _
        ByVal v_datChargeDate As Date, _
        ByVal v_datStartTime As Date, _
        ByVal v_datEndTime As Date) As Boolean
        
    Dim l_objShift As cboScheduledShift
    
    Set l_objShift = ShiftScheduledDuring( _
            v_datChargeDate, v_datStartTime, v_datEndTime)
    If l_objShift Is Nothing Then
        IsTradedWorkShift = False
    Else
        IsTradedWorkShift = l_objShift.IsTradedWorkShift
    End If
    Set l_objShift = Nothing

End Function

Public Function ShiftComponentOn( _
        ByVal v_datChargeDate As Date, _
        ByVal v_datStartDateTime As Date, _
        ByVal v_datEndDateTime As Date) As String

'No longer used
' ShiftComponentOnJobNum is now the one called


    Dim l_objShift As cboScheduledShift
    Dim l_blnCoffeen As Boolean
    Dim l_colShifts As Collection
    Dim l_blnLocal8 As Boolean
    Dim l_blnLocal8MaterialHan As Boolean
    Dim l_strJobNum As String
    Dim l_strShiftComponentOn As String
    Dim l_blnOvertime As Boolean
    Dim l_blnHutsonville148 As Boolean
    Dim l_blnHutsonville12HourAgreement As Boolean
    Dim l_blnCoffeen12HourAgreement As Boolean
    Dim l_blnLocal8MaterialHandler As Boolean
    
On Error GoTo eh
    l_blnCoffeen = Me.Employee.IsCoffeen(v_datChargeDate)
    l_blnLocal8 = Me.Employee.IsAmerenCILCOLocal8Employee(v_datChargeDate)
    l_blnCoffeen12HourAgreement = Me.Employee.IsCoffeen12HourAgreement(v_datChargeDate)
    l_blnHutsonville148 = Me.Employee.IsHutsonvilleLocal148(v_datChargeDate)
    l_blnHutsonville12HourAgreement = Me.Employee.IsHutsonville12HourAgreement(v_datChargeDate)
    l_blnLocal8MaterialHandler = Me.Employee.IsCILCOLocal8MaterialHandler(v_datChargeDate)
    
    If l_blnHutsonville12HourAgreement Then
        ShiftComponentOn = _
        DetermineStandardShiftComponent( _
            Format(v_datStartDateTime, "hh:nn"))
    Exit Function
    
    End If
    If l_blnCoffeen12HourAgreement Then
            ShiftComponentOn = _
            DetermineCoffeen12ShiftComponent( _
                Format(v_datStartDateTime, "hh:nn"))
        Exit Function
    End If
    
    If l_blnCoffeen Then
        ShiftComponentOn = _
            DetermineCoffeenShiftComponent( _
                Format(v_datStartDateTime, "hh:nn"))
        Exit Function
    End If
    
    Set l_objShift = ShiftScheduledDuring( _
            v_datChargeDate, v_datStartDateTime, v_datEndDateTime)
    If l_objShift Is Nothing Then
        l_blnOvertime = True
'Determing Shift Component for Overtime
        If Me.Employee.IsAmerenCIPS(v_datChargeDate) Then
'Currently we are not determining Shift Component for
'AmerenCIPS Overtime
'09-08-08 now assigning shift component for Huttsonville
            If l_blnHutsonville148 Then
                l_strShiftComponentOn = DetermineStandardShiftComponent( _
                    Format(v_datStartDateTime, "hh:nn"))
            Else
                l_strShiftComponentOn = ""
            End If
        Else
            Set l_colShifts = ShiftsOn(v_datChargeDate)
            If l_colShifts.Count = 0 Then
                l_strShiftComponentOn = ""
            Else
               Set l_objShift = l_colShifts.Item(1)
               l_strShiftComponentOn = l_objShift.ShiftComponentCode
            End If
        End If
    Else
        l_blnOvertime = False
'Straight Time Shift Component
        l_strShiftComponentOn = l_objShift.ShiftComponentCode
    End If
    
    
    
    If l_blnLocal8 Then
        If l_strShiftComponentOn = "R" Or _
            (l_strShiftComponentOn = "NC" And Not l_blnOvertime) Then
        Else
            If l_blnLocal8MaterialHandler Then
                    l_strShiftComponentOn = _
                        DetermineCILCOLocal8MaterialHandlersShiftComponent( _
                            Format(v_datStartDateTime, "hh:nn"))
                
            Else
                    l_strShiftComponentOn = _
                        DetermineCILCOLocal8ShiftComponent( _
                            Format(v_datStartDateTime, "hh:nn"))
            End If
        End If
    End If
    
    ShiftComponentOn = l_strShiftComponentOn
    Exit Function
    
eh:
    Set l_objShift = Nothing
    Set l_colShifts = Nothing
    TrsRaiseError mstrCMODULENAME, "ShiftComponentOn"

End Function

Private Function DetermineCoffeenShiftComponent( _
            ByVal v_strStartTime As String) As String
' BLUE014
On Error GoTo eh
    
    If IsTimeBetween(v_strStartTime, "15:00", "22:54") Then
        DetermineCoffeenShiftComponent = "P"
    ElseIf IsTimeBetween(v_strStartTime, "23:00", "23:54") Or _
        IsTimeBetween(v_strStartTime, "00:00", "06:54") Then
        DetermineCoffeenShiftComponent = "O"
    Else
        DetermineCoffeenShiftComponent = ""
    End If
    
    Exit Function
        
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineCoffeenShiftComponent"
    
End Function

Public Function TradeExceptionCode( _
        ByVal v_datChargeDate As Date, _
        ByVal v_datStartDateTime As Date, _
        ByVal v_datEndDateTime As Date) As String
' Answer a formatted trade exception code if the period
' is during a traded work shift, otherwise answer an empty string.
' BLUE029
    Dim l_objShift As cboScheduledShift
    
    Set l_objShift = ShiftScheduledDuring(v_datChargeDate, _
            v_datStartDateTime, v_datEndDateTime)
    If l_objShift Is Nothing Then
        TradeExceptionCode = ""
    Else
        TradeExceptionCode = _
            l_objShift.FormatTradeExceptionCode
    End If
    Set l_objShift = Nothing
    
End Function

Public Function ReliefCode( _
        ByVal v_datChargeDate As Date, _
        ByVal v_datStartDateTime As Date, _
        ByVal v_datEndDateTime As Date) As String
' Answer the vacation relief code for a scheduled work shift
' BLUE064
    Dim l_objShift As cboScheduledShift
    
    Set l_objShift = ShiftScheduledDuring(v_datChargeDate, _
            v_datStartDateTime, v_datEndDateTime)
    If l_objShift Is Nothing Then
        ReliefCode = ""
    Else
        ReliefCode = l_objShift.ReliefCode
    End If
    Set l_objShift = Nothing
    
End Function


Public Function ScheduledJobUpgrade( _
            ByVal v_datChargeDate As Date, _
            ByVal v_datStartTime As Date, _
            ByVal v_datEndTime As Date) As String
' Find the schedule row that this time period is a subset of
' and answer the sched_job_num.
' filter by chargedate
' check each row for subset or equal with period passed
'
' 7/20/00
' Ask jim what kind of record this is on?
' should the job number be on the scheduled shift object?
'
    Dim l_datStartDateIn As Date
    Dim l_datEndDateIn As Date
    Dim l_datStartDateSched As Date
    Dim l_datEndDateSched As Date
    Dim l_strSchedJobNum As String
    Dim l_blnSchedJobUpgrade As Boolean
    
    SetScheduleFilterOn v_datChargeDate
    If mobjScheduleState.BOF And _
        mobjScheduleState.EOF Then
        ClearScheduleFilter
        Exit Function
    End If
    
    l_datStartDateIn = v_datStartTime
    l_datEndDateIn = v_datEndTime
    
    Do Until mobjScheduleState.EOF
        If Not Trim(mobjScheduleState!START_TIME) = "" Then
            l_datStartDateSched = _
                DateTimeFrom(mobjScheduleState!DATE_WORKED, _
                mobjScheduleState!START_TIME)
            l_datEndDateSched = EndDateTime(mobjScheduleState)
            If TimePeriodsSubset( _
                    l_datStartDateIn, l_datEndDateIn, _
                    l_datStartDateSched, l_datEndDateSched) Then
                l_strSchedJobNum = _
                    Trim(mobjScheduleState!SCHED_JOB_NUM)       'SCHED_JOB_NUM
                ClearScheduleFilter
                l_blnSchedJobUpgrade = True
                Exit Do
            End If
        End If
        mobjScheduleState.MoveNext
    Loop
    If Me.Employee.IsUnion148(v_datChargeDate) Then
        If Not IsAnyPeriodWithinSchedule(v_datChargeDate, l_datStartDateIn, l_datEndDateIn) Then
            SetScheduleFilterOn v_datChargeDate
            mobjScheduleState.MoveFirst
            l_blnSchedJobUpgrade = False
            While Not mobjScheduleState.EOF And Not l_blnSchedJobUpgrade
                If mobjScheduleState!SYSTEM_EXC_CD <> "SLT" Then
                    l_strSchedJobNum = Trim(mobjScheduleState!SCHED_JOB_NUM)
                    If l_strSchedJobNum <> "" Then
                        l_blnSchedJobUpgrade = True
                    End If
                End If
                mobjScheduleState.MoveNext
            Wend
        End If
    End If
        
    If l_blnSchedJobUpgrade Then
        ScheduledJobUpgrade = l_strSchedJobNum
    Else
        ScheduledJobUpgrade = ""
    End If

    ClearScheduleFilter

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "ScheduledJobUpgrade"

End Function

Private Function EndDateTime(ByRef mobjScheduleState As ADODB.Recordset) As Date
    
On Error GoTo eh
' off day, answer default datetime
    If Trim(mobjScheduleState!END_TIME) = "" Then
        EndDateTime = CDate(CDEFAULTDATE)
        Exit Function
    End If

    If mobjScheduleState!END_TIME < mobjScheduleState!START_TIME Then
        EndDateTime = _
            DateTimeFrom(DateAdd("d", 1, mobjScheduleState!DATE_WORKED), _
                mobjScheduleState!END_TIME)
    Else
        EndDateTime = _
            DateTimeFrom(mobjScheduleState!DATE_WORKED, _
                mobjScheduleState!END_TIME)
    End If
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "EndDateTime"
            
End Function

Public Function ShiftScheduledDuring( _
            ByVal v_datChargeDate As Date, _
            ByVal v_datStartDateTime As Date, _
            ByVal v_datEndDateTime As Date) As cboScheduledShift
'Answer the cboScheduleShift that the passed time period
' falls between.  If none found, answer nothing.
'
    Dim shifts As Collection
    Dim eShift As cboScheduledShift
    
On Error GoTo eh
    Set shifts = ShiftsOn(v_datChargeDate)
    For Each eShift In shifts
        If eShift.IsScheduledWorkDay Then
            If TimePeriodsIntersect( _
                    v_datStartDateTime, _
                    v_datEndDateTime, _
                    eShift.ShiftStartDateTime, _
                    eShift.ShiftEndDateTime) Then
                Set ShiftScheduledDuring = eShift
                Exit For
            End If
        End If
    Next
    Exit Function
    
eh:
    Set shifts = Nothing
    Set eShift = Nothing
    TrsRaiseError mstrCMODULENAME, "ShiftScheduledDuring"
    
End Function

Public Function ScheduleChangePaymentScheduled( _
    ByVal v_datChargeDate As Date, ByVal v_datDateWorked As Date, _
    ByVal v_strStartTime As String, ByVal v_strEndTime As String) As String
'filter by chargedate
'check each row for subset or equal with period passed

    Dim l_varSchedChgPmt As Variant
    Dim l_strFilter As String
    Dim l_datStartDateIn As Date
    Dim l_datEndDateIn As Date
    Dim l_datStartDateSCP As Date
    Dim l_datEndDateSCP As Date
    Dim l_strPaymentType As String
    Dim l_colShifts As Collection
    Dim l_objShift As cboScheduledShift
    Dim l_strShiftStartTime As String
    Dim l_strShiftEndTime As String
    Dim l_colSchedChgPmts As Collection
    
On Error GoTo eh

    Set l_colSchedChgPmts = ScheduleChangePaymentsOn(v_datChargeDate)
    If l_colSchedChgPmts.Count = 0 Then
        Exit Function
    End If

    SetScheduleChangePmtFilterOn v_datChargeDate
    mobjScheduleChangePayments.MoveFirst
    
    l_varSchedChgPmt = l_colSchedChgPmts(1)
    If Not IsArray(l_varSchedChgPmt) Then
        If l_varSchedChgPmt = "OK" Then
            'New Schedule is completly within Base Schedule
            'No Payment Paid
            'May need BRV in Future
        Else
            'Can not find hours for schedule change payment
            'May need to write a BRV warning
        End If
        Exit Function
    End If
    
    l_datStartDateIn = DateTimeFrom(v_datDateWorked, v_strStartTime)
    l_datEndDateIn = DateTimeFrom(v_datDateWorked, v_strEndTime)
    l_datStartDateSCP = DateTimeFrom(l_varSchedChgPmt(2), l_varSchedChgPmt(4))
    If l_varSchedChgPmt(4) > l_varSchedChgPmt(5) Then
        l_datEndDateSCP = DateTimeFrom(DateAdd("d", 1, l_varSchedChgPmt(2)), l_varSchedChgPmt(5))
    Else
        l_datEndDateSCP = DateTimeFrom(l_varSchedChgPmt(2), l_varSchedChgPmt(5))
    End If
    
    If TimePeriodsSubset( _
                l_datStartDateIn, l_datEndDateIn, _
                l_datStartDateSCP, l_datEndDateSCP) Then
        ScheduleChangePaymentScheduled = _
            mobjScheduleChangePayments("EXCEPTION_CD")
        Exit Function
    End If
    
    If l_colSchedChgPmts.Count = 1 Then
        Exit Function
    End If
    
    l_varSchedChgPmt = l_colSchedChgPmts(2)
    l_datStartDateSCP = DateTimeFrom(l_varSchedChgPmt(2), l_varSchedChgPmt(4))
    If l_varSchedChgPmt(4) > l_varSchedChgPmt(5) Then
        l_datEndDateSCP = DateTimeFrom(DateAdd("d", 1, l_varSchedChgPmt(2)), l_varSchedChgPmt(5))
    Else
        l_datStartDateSCP = DateTimeFrom(l_varSchedChgPmt(2), l_varSchedChgPmt(5))
    End If
    
    If TimePeriodsSubset( _
                l_datStartDateIn, l_datEndDateIn, _
                l_datStartDateSCP, l_datEndDateSCP) Then
        ScheduleChangePaymentScheduled = _
            mobjScheduleChangePayments("EXCEPTION_CD")
    End If
    
    
    
'JIMB 06/12/2002 commented out old code below
'''''    Do While Not mobjScheduleChangePayments.EOF
'''''        l_datStartDateSCP = _
'''''            DateTimeFrom(mobjScheduleChangePayments("DATE"), _
'''''            mobjScheduleChangePayments("START_TIME"))
'''''
''''''24:00/00:00 fix
'''''        If Trim(mobjScheduleChangePayments("END_TIME")) = "00:00" Then
'''''            l_datEndDateSCP = _
'''''                DateTimeFrom(mobjScheduleChangePayments("DATE"), "24:00")
'''''        Else
'''''            l_datEndDateSCP = _
'''''                DateTimeFrom(mobjScheduleChangePayments("DATE"), _
'''''                mobjScheduleChangePayments("END_TIME"))
'''''        End If
'''''
'''''        If TimePeriodsSubset( _
'''''                l_datStartDateIn, l_datEndDateIn, _
'''''                l_datStartDateSCP, l_datEndDateSCP) Then
'''''            ScheduleChangePaymentScheduled = mobjScheduleChangePayments("EXCEPTION_CD")
'''''            Exit Function
'''''        End If
'''''        mobjScheduleChangePayments.MoveNext
'''''    Loop
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ScheduleChangePaymentScheduled"

End Function

Public Function HasScheduleChangePayment( _
    ByVal v_datChargeDate, ByVal v_datDateWorked, _
    ByVal v_strStartTime, ByVal v_strEndTime) As Boolean
    
'07/03/2002 -- JIMB
'This methdod is dead code
 
    HasScheduleChangePayment = Not IsEmpty(FirstScheduleChangePayment( _
        v_datChargeDate, v_datDateWorked, _
        v_strStartTime, v_strEndTime))

End Function

Public Function FirstScheduleChangePayment( _
    ByVal v_datChargeDate, ByVal v_datDateWorked, _
    ByVal v_strStartTime, ByVal v_strEndTime) As Variant
    
'07/03/2002 -- JIMB
'This methdod is dead code
    
'filter by chargedate
'check each row for intersection with period passed
'answer an empty variant when sched chg pmt not found

    Dim l_varSchedChgPmt As Variant
    Dim l_datStartDateIn As Date
    Dim l_datEndDateIn As Date
    Dim l_datStartDateSCP As Date
    Dim l_datEndDateSCP As Date

On Error GoTo eh
    SetScheduleChangePmtFilterOn v_datChargeDate
    If mobjScheduleChangePayments.BOF And _
        mobjScheduleChangePayments.EOF Then
        Exit Function
    Else
        mobjScheduleChangePayments.MoveFirst
    End If

    l_datStartDateIn = DateTimeFrom(v_datDateWorked, v_strStartTime)
    l_datEndDateIn = DateTimeFrom(v_datDateWorked, v_strEndTime)
    Do While Not mobjScheduleChangePayments.EOF
        l_datStartDateSCP = _
            DateTimeFrom(mobjScheduleChangePayments("DATE"), _
            mobjScheduleChangePayments("START_TIME"))
        l_datEndDateSCP = _
            DateTimeFrom(mobjScheduleChangePayments("DATE"), _
            mobjScheduleChangePayments("END_TIME"))
        If TimePeriodsIntersect( _
                l_datStartDateIn, l_datEndDateIn, _
                l_datStartDateSCP, l_datEndDateSCP) Then
            ReDim l_varSchedChgPmt(1 To 6)
            l_varSchedChgPmt(1) = mobjScheduleChangePayments("EMP_NUM")
            l_varSchedChgPmt(2) = mobjScheduleChangePayments("DATE")
            l_varSchedChgPmt(3) = mobjScheduleChangePayments("CHARGE_DATE")
            l_varSchedChgPmt(4) = mobjScheduleChangePayments("START_TIME")
            l_varSchedChgPmt(5) = mobjScheduleChangePayments("END_TIME")
            l_varSchedChgPmt(6) = mobjScheduleChangePayments("EXCEPTION_CD")
            FirstScheduleChangePayment = l_varSchedChgPmt
            Exit Function
        End If
        mobjScheduleChangePayments.MoveNext
    Loop
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "FirstScheduleChangePayment"

End Function

Public Function GetScheduleChangePayments() As ADODB.Recordset
    Set GetScheduleChangePayments = mobjScheduleChangePayments
End Function

Public Function ShiftDetailsOn( _
        ByVal v_datChargeDate As Date, _
        ByVal v_datStartDateTime As Date, _
        ByVal v_datEndDateTime As Date) As Variant
' Answer a zero-based array with SCHED_CD and SCHED_SRC_CD
' in positions 1 and 2, respectively.
    Dim l_objShift As cboScheduledShift
    
On Error GoTo eh
    
    Set l_objShift = ShiftScheduledDuring( _
            v_datChargeDate, v_datStartDateTime, v_datEndDateTime)
    If l_objShift Is Nothing Then
        ShiftDetailsOn = Array(0, 0)
    Else
        ShiftDetailsOn = Array(l_objShift.ScheduleCode, l_objShift.ScheduleSource)
    End If
    Exit Function
    
eh:
    Set l_objShift = Nothing
    TrsRaiseError mstrCMODULENAME, "ShiftDetailsOn"
End Function


Public Sub Restore(empId As Long, _
        schedStartDate As Date, schedEndDate As Date)

    Dim l_dbSchedule As dbSchedule.cdSchedule
    Dim l_bsSchedule As busSchedule.cbsSchedule
    Dim l_rsSchedule As ADODB.Recordset
    Dim l_rsScheduleWOExc As ADODB.Recordset
    
On Error GoTo Restore_eh
    Set l_bsSchedule = CreateObject("busSchedule.cbsSchedule")
    Set l_rsSchedule = l_bsSchedule.GetAdjustedBaseSchedule(empId, schedStartDate, schedEndDate, l_rsScheduleWOExc)
    SetState l_rsSchedule
    
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    Set mobjScheduleChangePayments = _
        l_dbSchedule.GetScheduleChangePayments(empId, schedStartDate, schedEndDate)
    If RecordsExist(mobjScheduleChangePayments) Then
        mobjScheduleChangePayments.Filter = "PAYMENT_TYPE <> 'F'"
        If RecordsExist(mobjScheduleChangePayments) Then
            Dim l_colSched As Collection
            Set l_colSched = New Collection
            mobjScheduleChangePayments.MoveFirst
            While Not mobjScheduleChangePayments.EOF
                l_colSched.Add mobjScheduleChangePayments!BASELINE_SCHED_CD.Value
                mobjScheduleChangePayments.MoveNext
            Wend
            Set mobjAvailSched = l_dbSchedule.GetAvailSchedInfo(l_colSched)
        Else
            Set mobjAvailSched = Nothing
        End If
        mobjScheduleChangePayments.Filter = ""
    Else
        Set mobjAvailSched = Nothing
    End If
    Set l_dbSchedule = Nothing
    Set l_bsSchedule = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsScheduleWOExc = Nothing
    Exit Sub
    
Restore_eh:
    Set l_dbSchedule = Nothing
    Set l_bsSchedule = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsScheduleWOExc = Nothing
    TrsRaiseError mstrCMODULENAME, "Restore"
End Sub


Public Sub RestoreWithState( _
            ByRef r_rsSchedule As ADODB.Recordset, _
            ByRef r_rsScheduleChangePayments As ADODB.Recordset)

    Dim l_dbSchedule As dbSchedule.cdSchedule
    
On Error GoTo eh
    
    SetState r_rsSchedule
    If r_rsScheduleChangePayments Is Nothing Then
        Set mobjScheduleChangePayments = BuildEmptySchedChgPmtRecordset
    Else
        Set mobjScheduleChangePayments = r_rsScheduleChangePayments
    End If
    
    If RecordsExist(mobjScheduleChangePayments) Then
        mobjScheduleChangePayments.Filter = "PAYMENT_TYPE <> 'F'"
        If RecordsExist(mobjScheduleChangePayments) Then
            Dim l_colSched As Collection
            Set l_colSched = New Collection
            mobjScheduleChangePayments.MoveFirst
            While Not mobjScheduleChangePayments.EOF
                l_colSched.Add mobjScheduleChangePayments!BASELINE_SCHED_CD.Value
                mobjScheduleChangePayments.MoveNext
            Wend
            Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
            Set mobjAvailSched = l_dbSchedule.GetAvailSchedInfo(l_colSched)
            Set l_dbSchedule = Nothing
        Else
            Set mobjAvailSched = Nothing
        End If
        mobjScheduleChangePayments.Filter = ""
    Else
        Set mobjAvailSched = Nothing
    End If
    
    
    Exit Sub
    
eh:
    Set l_dbSchedule = Nothing
    TrsRaiseError mstrCMODULENAME, "RestoreWithState"
    
End Sub

Private Function BuildEmptySchedChgPmtRecordset() As ADODB.Recordset
' Minimum definition to enable the filter to work
    Dim lrsSchedChgPmt As ADODB.Recordset
    
    Set lrsSchedChgPmt = New ADODB.Recordset
    lrsSchedChgPmt.Fields.Append "EMP_NUM", adInteger
    lrsSchedChgPmt.Fields.Append "CHARGE_DATE", adDate
    lrsSchedChgPmt.Open
    Set BuildEmptySchedChgPmtRecordset = lrsSchedChgPmt

End Function

Private Sub SetState(rs As ADODB.Recordset)
    Set mobjScheduleState = rs
    mstrEmpNum = CStr(rs!EMP_NUM)
    mstrScheduleTimeStamp = Trim(rs!LAST_MOD_SCHED_TM)
End Sub

Public Sub Release()

'Donot enable an error handler since this is called from an error handler.

    Set mobjScheduleState = Nothing
    Set mobjScheduleChangePayments = Nothing
    Set mobjTimecard = Nothing
    GetObjectContext.SetComplete
    
End Sub

Private Function CreateNewShift() As cboScheduledShift
    Set CreateNewShift = CreateObject("busTimecard.cboScheduledShift")
End Function

Private Function IsTimeDuringLunch( _
            ByVal v_datStartTime As Date, _
            ByVal v_datEndTime As Date, _
            ByVal v_datLunchStartTime As Date, _
            ByVal v_datLunchEndTime As Date _
            ) As Boolean
            
On Error GoTo eh
    If TimePeriodsSubset( _
        v_datStartTime, v_datEndTime, v_datLunchStartTime, v_datLunchEndTime) Then
        IsTimeDuringLunch = True
    Else
        IsTimeDuringLunch = False
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsTimeDuringLunch"
            
End Function

Public Function ScheduleChangePaymentsOn( _
    ByVal v_datChargeDate As Date _
    ) As Collection


'filter by chargedate
'check each row for subset or equal with period passed

    Dim l_strFilter As String
    Dim l_datStartDateSCP As Date
    Dim l_datEndDateSCP As Date
    Dim l_strPaymentType As String
    Dim l_colShifts As Collection
    Dim l_objShift As cboScheduledShift
    Dim l_varSchedChgPmt(1 To 6) As Variant
    Dim l_colSchedChangePMTs As Collection
    Dim l_intBaseLineSchedCode As Integer
    Dim l_strBaseStartTime As String
    Dim l_strBaseEndTime As String
    Dim l_datBaseStartDate As Date
    Dim l_datBaseEndDate As Date
    Dim l_datBaseStartDateTime As Date
    Dim l_datBaseEndDateTime As Date
    Dim l_datShiftStartDateTime As Date
    Dim l_datShiftEndDateTime As Date
    Dim l_decShiftHours As Variant
    Dim l_intHowDatesIntersect As Integer
    
On Error GoTo eh

    Set l_colSchedChangePMTs = New Collection
    
    SetScheduleChangePmtFilterOn v_datChargeDate
    If mobjScheduleChangePayments.BOF And _
        mobjScheduleChangePayments.EOF Then
        Set ScheduleChangePaymentsOn = l_colSchedChangePMTs
        Exit Function
    Else
        mobjScheduleChangePayments.MoveFirst
    End If
    
'There can only be one schedule change payment on a charge date for an employee
' START and END time are no longer on the table
'A double shift does not get the payment on both shifts

    Set l_colShifts = ShiftsOn(v_datChargeDate)
    Set l_objShift = l_colShifts(1)
    If l_objShift.ShiftHours > 0 Then
    Else
        If l_colShifts.Count > 1 Then
            Set l_objShift = l_colShifts(2)
            l_decShiftHours = CDec(l_objShift.ShiftHours)
        Else
            l_decShiftHours = CDec(0)
        End If
        If l_decShiftHours > 0 Then
        Else
        'Can not find hours for schedule change payment
        'May need to write a BRV warning
            l_colSchedChangePMTs.Add "BAD"
            Set ScheduleChangePaymentsOn = l_colSchedChangePMTs
            Exit Function
        End If
    End If

    l_strPaymentType = mobjScheduleChangePayments!PAYMENT_TYPE
    If l_strPaymentType = "F" Then
        l_varSchedChgPmt(1) = mobjScheduleChangePayments("EMP_NUM")
        l_varSchedChgPmt(2) = l_objShift.ShiftWorkDate
        l_varSchedChgPmt(3) = l_objShift.ShiftChargeDate
        l_varSchedChgPmt(4) = l_objShift.ShiftStart
        l_varSchedChgPmt(5) = l_objShift.ShiftEnd
        l_varSchedChgPmt(6) = mobjScheduleChangePayments("EXCEPTION_CD")
        l_colSchedChangePMTs.Add l_varSchedChgPmt
        Set ScheduleChangePaymentsOn = l_colSchedChangePMTs
        Exit Function
    End If
    
    
'TRY and find Schedule Change PMTs for Differences
    l_intBaseLineSchedCode = mobjScheduleChangePayments!BASELINE_SCHED_CD
    GetBaseLineScheduleData l_intBaseLineSchedCode, v_datChargeDate, _
        l_strBaseStartTime, l_strBaseEndTime, _
        l_datBaseStartDate, l_datBaseEndDate, _
        l_datBaseStartDateTime, l_datBaseEndDateTime
    l_datStartDateSCP = l_objShift.ShiftStartDateTime
    l_datEndDateSCP = l_objShift.ShiftEndDateTime


    l_intHowDatesIntersect = DateIntersection( _
        l_datStartDateSCP, l_datEndDateSCP, _
        l_datBaseStartDateTime, l_datBaseEndDateTime)
        
    Select Case l_intHowDatesIntersect
        Case 0, 5 'No intersection
            'Can not find hours for schedule change payment
            'May need to write a BRV warning
'            l_colSchedChangePMTs.Add "BAD"
'JimB 09/26/2002
'Per Judi Gosik's request
'The system will give a schedule change payment for the entire shift
' when the new shifts falls outside of the prior (Base) shift.
        l_varSchedChgPmt(1) = mobjScheduleChangePayments("EMP_NUM")
        l_varSchedChgPmt(2) = l_objShift.ShiftWorkDate
        l_varSchedChgPmt(3) = l_objShift.ShiftChargeDate
        l_varSchedChgPmt(4) = l_objShift.ShiftStart
        l_varSchedChgPmt(5) = l_objShift.ShiftEnd
        l_varSchedChgPmt(6) = mobjScheduleChangePayments("EXCEPTION_CD")
        l_colSchedChangePMTs.Add l_varSchedChgPmt
        Case 1
            'New Schedule extends Base Schedule
            '1 or 2 payments need paid
            If DateDiff("n", l_datStartDateSCP, l_datBaseStartDateTime) <> 0 Then
                l_varSchedChgPmt(1) = mobjScheduleChangePayments("EMP_NUM")
                l_varSchedChgPmt(2) = l_datStartDateSCP
                l_varSchedChgPmt(3) = l_objShift.ShiftChargeDate
                l_varSchedChgPmt(4) = Format(l_datStartDateSCP, "HH:MM")
                l_varSchedChgPmt(5) = Format(l_datBaseStartDateTime, "HH:MM")
                l_varSchedChgPmt(6) = mobjScheduleChangePayments("EXCEPTION_CD")
                l_colSchedChangePMTs.Add l_varSchedChgPmt
            End If
            If DateDiff("n", l_datEndDateSCP, l_datBaseEndDateTime) <> 0 Then
                l_varSchedChgPmt(1) = mobjScheduleChangePayments("EMP_NUM")
                l_varSchedChgPmt(2) = l_datBaseEndDateTime
                l_varSchedChgPmt(3) = l_objShift.ShiftChargeDate
                l_varSchedChgPmt(4) = Format(l_datBaseEndDateTime, "HH:MM")
                l_varSchedChgPmt(5) = Format(l_datEndDateSCP, "HH:MM")
                If l_varSchedChgPmt(5) = "00:00" Then
                    l_varSchedChgPmt(5) = "24:00"
                End If
                l_varSchedChgPmt(6) = mobjScheduleChangePayments("EXCEPTION_CD")
                l_colSchedChangePMTs.Add l_varSchedChgPmt
            End If
        
        Case 2
            'New Schedule is completly within Base Schedule
            'No Payment Paid
            'May need to write a BRV warning
            l_colSchedChangePMTs.Add "OK"
        
        Case 3
            'New Schedule starts on or prior to Base Schedule start
            'New Schedule ends before Base Schedule ends
            '1 or 0 payments need paid
            If DateDiff("n", l_datStartDateSCP, l_datBaseStartDateTime) <> 0 Then
                l_varSchedChgPmt(1) = mobjScheduleChangePayments("EMP_NUM")
                l_varSchedChgPmt(2) = l_datStartDateSCP
                l_varSchedChgPmt(3) = l_objShift.ShiftChargeDate
                l_varSchedChgPmt(4) = Format(l_datStartDateSCP, "HH:MM")
                l_varSchedChgPmt(5) = Format(l_datBaseStartDateTime, "HH:MM")
                l_varSchedChgPmt(6) = mobjScheduleChangePayments("EXCEPTION_CD")
                l_colSchedChangePMTs.Add l_varSchedChgPmt
            Else
                'May need to write a BRV warning
                l_colSchedChangePMTs.Add "OK"
            End If

        Case 4
            'New Schedule starts after Base Schedule starts
            'New Schedule on or after Base Schedule ends
            '1 or 0 payments need paid
            If DateDiff("n", l_datEndDateSCP, l_datBaseEndDateTime) <> 0 Then
                l_varSchedChgPmt(1) = mobjScheduleChangePayments("EMP_NUM")
                l_varSchedChgPmt(2) = l_datBaseEndDateTime
                l_varSchedChgPmt(3) = l_objShift.ShiftChargeDate
                l_varSchedChgPmt(4) = Format(l_datBaseEndDateTime, "HH:MM")
                l_varSchedChgPmt(5) = Format(l_datEndDateSCP, "HH:MM")
                l_varSchedChgPmt(6) = mobjScheduleChangePayments("EXCEPTION_CD")
                l_colSchedChangePMTs.Add l_varSchedChgPmt
            Else
                'May need to write a BRV warning
                l_colSchedChangePMTs.Add "OK"
            End If
    End Select
    
    Set ScheduleChangePaymentsOn = l_colSchedChangePMTs
    
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ScheduleChangePaymentsOn"

    
    
End Function



Private Sub GetBaseLineScheduleData( _
    ByVal v_intSchedCode As Integer, _
    ByVal v_datChargeDate As Date, _
    ByRef r_strStartTime As String, _
    ByRef r_strEndTime As String, _
    ByRef r_datStartDate As Date, _
    ByRef r_datEndDate As Date, _
    ByRef r_datStartDateTime As Date, _
    ByRef r_datEndDateTime As Date)
    
    
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
'''''    Dim l_datStartDateTime As Date
'''''    Dim l_datEndDateTime As Date
    Dim l_strDay As String
    Dim l_decShiftHours As Variant
    Dim l_strShiftStartDay As String
    
    mobjAvailSched.Filter = "SCHED_CD = " & str(v_intSchedCode)
    l_strDay = UCase(WeekdayName(Weekday(v_datChargeDate), True))
    l_strStartTime = mobjAvailSched(l_strDay & "_START_TIME")
    l_strEndTime = mobjAvailSched(l_strDay & "_END_TIME")
    l_decShiftHours = mobjAvailSched(l_strDay & "_STR_HRS")
    l_strShiftStartDay = mobjAvailSched(l_strDay & "_ACT_DAY")
    DetermineShiftDates l_strShiftStartDay, l_strDay, v_datChargeDate, _
        l_strStartTime, l_strEndTime, _
        l_datStartDate, l_datEndDate
        
    If Trim(l_strStartTime) = "" Then
        l_strStartTime = "00:00"
    End If
    If Trim(l_strEndTime) = "" Then
        l_strEndTime = "00:00"
    End If
    r_datStartDateTime = DateTimeFrom(l_datStartDate, l_strStartTime)
    r_datEndDateTime = DateTimeFrom(l_datEndDate, l_strEndTime)
    r_strStartTime = l_strStartTime
    r_strEndTime = l_strEndTime
    r_datStartDate = l_datStartDate
    r_datEndDate = l_datEndDate
        
        
End Sub


Private Sub DetermineShiftDates( _
                ByVal v_strShiftStartDay As String, _
                ByVal v_strChargeDateDayOfWeek As String, _
                ByVal v_datChargeDate As String, _
                ByVal v_strShiftStartTime As String, _
                ByVal v_strSiftEndTime As String, _
                ByRef r_datShiftStartDate As Date, _
                ByRef r_datShiftEndDate As Date)
                
'All shifts start on the charge date or 1 day prior of the charge date
                              
'No Shift
    If Trim(v_strShiftStartDay) = "" Then
        r_datShiftStartDate = v_datChargeDate
        r_datShiftEndDate = v_datChargeDate
        Exit Sub
    End If
    
' Shift starts on charge date
    If v_strShiftStartDay = v_strChargeDateDayOfWeek Then
        r_datShiftStartDate = v_datChargeDate
        If v_strShiftStartTime <= v_strSiftEndTime Then
            r_datShiftEndDate = v_datChargeDate
        Else
            r_datShiftEndDate = DateAdd("d", 1, v_datChargeDate)
        End If
        Exit Sub
    End If
    
' Shift starts prior to charge date
    r_datShiftStartDate = DateAdd("d", -1, v_datChargeDate)
    r_datShiftEndDate = v_datChargeDate

         
End Sub

Public Property Let ScheduleTimeStamp(ByVal vData As String)
    mstrScheduleTimeStamp = vData
End Property

Public Property Get ScheduleTimeStamp() As String
    ScheduleTimeStamp = mstrScheduleTimeStamp
End Property

Private Function DetermineCILCOLocal8ShiftComponent( _
            ByVal v_strStartTime As String) As String
' BLUE014
On Error GoTo eh
    
    If IsTimeBetween(v_strStartTime, "06:30", "14:24") Then
        DetermineCILCOLocal8ShiftComponent = ""
    Else
        DetermineCILCOLocal8ShiftComponent = "P"
    End If
    
    Exit Function
        
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineCILCOLocal8ShiftComponent"
    
End Function

Private Function DetermineCILCOLocal8MaterialHandlersShiftComponent( _
            ByVal v_strStartTime As String) As String
' BLUE014
On Error GoTo eh
    
    If IsTimeBetween(v_strStartTime, "06:00", "13:54") Then
        DetermineCILCOLocal8MaterialHandlersShiftComponent = ""
    Else
        DetermineCILCOLocal8MaterialHandlersShiftComponent = "P"
    End If
    
    Exit Function
        
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineCILCOLocal8MaterialHandlersShiftComponent"
    
End Function

Private Function DetermineCoffeen12ShiftComponent( _
            ByVal v_strStartTime As String) As String
' The employees under the Coffeen overtime agreement get the shift component C12
' for all hours worked between 19:00 and 07:00

On Error GoTo eh
    
    If IsTimeBetween(v_strStartTime, "19:00", "23:54") Then
        DetermineCoffeen12ShiftComponent = "C12"
    ElseIf IsTimeBetween(v_strStartTime, "00:00", "06:54") Then
        DetermineCoffeen12ShiftComponent = "C12"
    Else
        DetermineCoffeen12ShiftComponent = ""
    End If
    
    Exit Function
        
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineCoffeen12ShiftComponent"
    
End Function

Private Function DetermineStandardShiftComponent( _
            ByVal v_strStartTime As String) As String
' Assign the P shift component for time worked between 15:00 and 23:00
' Assign the O shift component for time worked between 23:00 and 07:00
' Assign no component for time worked between 07:00 and 15:00
On Error GoTo eh
    
    If IsTimeBetween(v_strStartTime, "15:00", "22:54") Then
        DetermineStandardShiftComponent = "P"
    ElseIf IsTimeBetween(v_strStartTime, "23:00", "23:54") Or _
        IsTimeBetween(v_strStartTime, "00:00", "06:54") Then
        DetermineStandardShiftComponent = "O"
    Else
        DetermineStandardShiftComponent = ""
    End If
    
    Exit Function
        
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineStandardShiftComponent"
    
End Function


Public Function ShiftComponentOnJobNum( _
        ByVal v_datChargeDate As Date, _
        ByVal v_datStartDateTime As Date, _
        ByVal v_datEndDateTime As Date, _
        ByVal v_JobNum As String) As String
        
    Dim l_objShift As cboScheduledShift
    Dim l_blnCoffeen As Boolean
    Dim l_colShifts As Collection
    Dim l_blnLocal8 As Boolean
    Dim l_blnLocal8MaterialHan As Boolean
    Dim l_strJobNum As String
    Dim l_strShiftComponentOn As String
    Dim l_blnOvertime As Boolean
    Dim l_blnHutsonville148 As Boolean
    Dim l_blnHutsonville12HourAgreement As Boolean
    Dim l_blnCoffeen12HourAgreement As Boolean
    Dim l_blnLocal8MaterialHandler As Boolean
    Dim l_blnHutsonvilleMaint As Boolean
    Dim l_blnCallawaySecurity As Boolean
    
On Error GoTo eh
    l_blnCoffeen = Me.Employee.IsCoffeen(v_datChargeDate)
    l_blnLocal8 = Me.Employee.IsAmerenCILCOLocal8Employee(v_datChargeDate)
    l_blnCoffeen12HourAgreement = Me.Employee.IsCoffeen12HourAgreementJob(v_datChargeDate, v_JobNum)
    l_blnHutsonville148 = Me.Employee.IsHutsonvilleLocal148(v_datChargeDate)
    l_blnHutsonville12HourAgreement = Me.Employee.IsHutsonville12HourAgreementJob(v_datChargeDate, v_JobNum)
    l_blnLocal8MaterialHandler = Me.Employee.IsCILCOLocal8MaterialHandler(v_datChargeDate)
    l_blnHutsonvilleMaint = Me.Employee.IsHutsonvilleMaint(v_datChargeDate) _
                                Or Me.Employee.IsHutsonvilleMaintJob(v_datChargeDate, v_JobNum)
    l_blnCallawaySecurity = Me.Employee.IsCallawaySecurity(v_datChargeDate)
    
    If l_blnHutsonville12HourAgreement Then
        ShiftComponentOnJobNum = _
        DetermineStandardShiftComponent( _
            Format(v_datStartDateTime, "hh:nn"))
    Exit Function
    
    End If
    If l_blnCoffeen12HourAgreement Then
            ShiftComponentOnJobNum = _
            DetermineCoffeen12ShiftComponent( _
                Format(v_datStartDateTime, "hh:nn"))
        Exit Function
    End If
    
    If l_blnCoffeen Then
        ShiftComponentOnJobNum = _
            DetermineCoffeenShiftComponent( _
                Format(v_datStartDateTime, "hh:nn"))
        Exit Function
    End If
    
    Set l_objShift = ShiftScheduledDuring( _
            v_datChargeDate, v_datStartDateTime, v_datEndDateTime)
    If l_objShift Is Nothing Then
        l_blnOvertime = True
'Determing Shift Component for Overtime
        If Me.Employee.IsAmerenCIPS(v_datChargeDate) Then
'Currently we are not determining Shift Component for
'AmerenCIPS Overtime
'09-08-08 now assigning shift component for Huttsonville
            If l_blnHutsonville148 Then
                If l_blnHutsonvilleMaint Then
                    l_strShiftComponentOn = DetermineMaintenanceShiftComponent( _
                        Format(v_datStartDateTime, "hh:nn"))
                Else
                    l_strShiftComponentOn = DetermineStandardShiftComponent( _
                        Format(v_datStartDateTime, "hh:nn"))
                End If
            Else
                l_strShiftComponentOn = ""
            End If
        Else
            Set l_colShifts = ShiftsOn(v_datChargeDate)
            If l_colShifts.Count = 0 Then
                l_strShiftComponentOn = ""
            Else
               Set l_objShift = l_colShifts.Item(1)
               l_strShiftComponentOn = l_objShift.ShiftComponentCode
            End If
        End If
    Else
        l_blnOvertime = False
'Straight Time Shift Component
        l_strShiftComponentOn = l_objShift.ShiftComponentCode
    End If
    
    
    
    If l_blnLocal8 Then
        If l_strShiftComponentOn = "R" Or _
            (l_strShiftComponentOn = "NC" And Not l_blnOvertime) Then
        Else
            If l_blnLocal8MaterialHandler Then
                    l_strShiftComponentOn = _
                        DetermineCILCOLocal8MaterialHandlersShiftComponent( _
                            Format(v_datStartDateTime, "hh:nn"))
                
            Else
                    l_strShiftComponentOn = _
                        DetermineCILCOLocal8ShiftComponent( _
                            Format(v_datStartDateTime, "hh:nn"))
            End If
        End If
    End If
    
'Determine Shift Component on overtime for Callaway Security based on time worked
    If l_blnCallawaySecurity And l_blnOvertime Then
            l_strShiftComponentOn = DetermineCallawaySecurityShiftComponent(Format(v_datStartDateTime, "hh:nn"), Not Me.Employee.IsNonExemptExecutive(v_datChargeDate))
    End If
    
    ShiftComponentOnJobNum = l_strShiftComponentOn
    Exit Function
    
eh:
    Set l_objShift = Nothing
    Set l_colShifts = Nothing
    TrsRaiseError mstrCMODULENAME, "ShiftComponentOnJobNum"

End Function


Private Function DetermineMaintenanceShiftComponent( _
            ByVal v_strStartTime As String) As String
' Assign the P shift component for time worked between 15:30 and 23:00
' Assign the O shift component for time worked between 23:00 and 07:00
' Assign no component for time worked between 07:00 and 15:30
On Error GoTo eh
    
    If IsTimeBetween(v_strStartTime, "15:30", "22:54") Then
        DetermineMaintenanceShiftComponent = "P"
    ElseIf IsTimeBetween(v_strStartTime, "23:00", "23:54") Or _
        IsTimeBetween(v_strStartTime, "00:00", "06:54") Then
        DetermineMaintenanceShiftComponent = "O"
    Else
        DetermineMaintenanceShiftComponent = ""
    End If
    
    Exit Function
        
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineMaintenanceShiftComponent"
    
End Function

Private Function DetermineCallawaySecurityShiftComponent( _
    ByVal v_strStartTime As String, _
    ByVal v_blnExemptEmployee As Boolean _
    ) As String
       
    Dim l_strShiftComp As String
    
    On Error GoTo eh
    
    If IsTimeBetween(v_strStartTime, "06:00", "17:54") Then
        l_strShiftComp = ""
    Else
        l_strShiftComp = GetCallawayNightShiftComponent(v_blnExemptEmployee)
    End If
    
    DetermineCallawaySecurityShiftComponent = l_strShiftComp
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "DetermineCallawaySecurityShiftComponent"
    
End Function

Public Function IsScheduledSpecialOverride(ByVal v_datChargeDate As Date) As Boolean
' Is any shift on chgDate a vacation.
    Dim shifts As Collection
    Dim eShift As cboScheduledShift
    
    Set shifts = ShiftsOn(v_datChargeDate)
    For Each eShift In shifts
        If eShift.IsScheduledSpecialOverride Then
            IsScheduledSpecialOverride = True
            Exit Function
        End If
    Next
    
End Function
