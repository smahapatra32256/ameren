Attribute VB_Name = "mbusCascade"
Option Explicit

Private Const m_strCMODULENAME As String = "mbusCascade"

Public Const CFULLTYPE As String = "F"
Public Const CPARTTYPE As String = "P"
Public Const CMODTYPE As String = "M"
Public Const CRESCHEDEXC As String = "R"
Public Const CDELETEEXC As String = "D"
Public Const CHOLIDAY As String = "H"
Public Const CHOLIDAYBANKED As String = "HB"
Public Const CHOLIDAYONREST As String = "HR"
Public Const CSYSWORK As String = "SWT"
Public Const CSYSLUNCH As String = "SLT"
Public Const CSYSOVERTIME As String = "SOT"
Public Const CHRSTARTTIME As String = "08:00"
Public Const CHRENDTIME As String = "16:00"
Public Const CHRHOURS As Integer = 8
Public Const CFULLSHIFTEXC As Long = 115
Public Const CFULLPARTEXC As Long = 110
Public Const CMIDNIGHT As Long = 1
Public Const CMINUTESINDAY As Long = 1440
Public Const CLOWDATE As Date = #1/1/1900#
Public Const CHIGHDATE As Date = #12/31/2079#
Public Const CEMPRVTRANCD As String = "D"

Public Enum ScheduleSourceEnum
    ssScheduleChange = 40
End Enum

Type ScheduledShift
    ScheduleCode As Long
    ShiftCode As String
    ShiftComponent As String
    StartTime As String
    EndTime As String
    Hours As Variant            'decimal 4,2
    LunchStart As String
    LunchEnd As String
    OvertimeStart As String
    OvertimeEnd As String
    OvertimeHours As Variant    'decimal 4,2
    ActualStartDay As String
    ActualStartDate As Date
    ChargeDay As String
End Type

Type ScheduledException
    EmployeeNumber As Long
    ChargeDate As Date
    ShiftNumber As Integer
    ScheduleCode As Long
    ScheduleSourceCode As Long
    ExceptionCode As String
    OverrideExceptionCode As String
    Hours As Variant        'decimal 3,2
    HoursOffset As Variant  'decimal 3,2
' determine start/end
    StartTime As String
    EndTime As String
    JobNumber As String
    ReliefComponent As String
    ComponentCode As String
    ShiftComponent As String
    CaseNumber As Long
    WorkgroupNumber As Long
    FamilyLeaveIndicator As String
'    LastModEmpNum As Long
    LastModTimestamp As String
End Type

Public Enum EMPRVTranEnum
    etEMPNUM = 0
    etDEPTNUM
    etEXCCD
    etCHGDATE
    etHOURS
    etTRANCD
    etLOCCD
End Enum

'Test Only
Public Sub DebugPrint(ByVal v_varStringArray As Variant)

    Dim i As Integer
    
    For i = LBound(v_varStringArray) To UBound(v_varStringArray)
        Debug.Print v_varStringArray(i)
    Next

End Sub



'******************************************************************************
'  Public procedures
'******************************************************************************

Public Function AllocateScheduledShift( _
                    ByRef r_rsAvailSched As ADODB.Recordset, _
                    ByVal v_lngSchedCode As Long, _
                    ByVal v_datChargeDate As Date) _
                As ScheduledShift
'Get shift data from available schedule for a given date
    Dim l_strDayofWeek As String
    Dim l_SchedShift As ScheduledShift
    
On Error GoTo eh
    r_rsAvailSched.Filter = "SCHED_CD = " & v_lngSchedCode
    If Not RecordsExist(r_rsAvailSched) Then
        l_SchedShift.ScheduleCode = -1
        Exit Function
    End If
    l_strDayofWeek = UCase(WeekdayName(Weekday(v_datChargeDate), True))   'returns day name SUN,MON,TUE, ..., SAT
    l_SchedShift = _
        AllocateScheduledShiftByDay(r_rsAvailSched, l_strDayofWeek)
    'Start day must be the day before when ActualDay does not equal ChargeDay.
    'ActualDay is an empty string for off days.
    If Trim(l_SchedShift.ActualStartDay) = "" Or _
        l_SchedShift.ActualStartDay = l_SchedShift.ChargeDay Then
        l_SchedShift.ActualStartDate = v_datChargeDate
    Else
        l_SchedShift.ActualStartDate = DateAdd("d", -1, v_datChargeDate)
    End If
    AllocateScheduledShift = l_SchedShift
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "AllocateScheduledShiftByDate"
End Function

Public Function CopyScheduledException( _
                    ByRef r_rsExceptions As ADODB.Recordset) _
                As ScheduledException
'Copy exception record from current position of r_rsExceptions
    Dim l_SchedExc As ScheduledException
    
On Error GoTo eh
    l_SchedExc.EmployeeNumber = r_rsExceptions("EMP_NUM")
    l_SchedExc.ChargeDate = r_rsExceptions("CHARGE_DATE")
    l_SchedExc.ShiftNumber = r_rsExceptions("SHIFT_NUM")
    l_SchedExc.ScheduleCode = r_rsExceptions("SCHED_CD")
    l_SchedExc.ScheduleSourceCode = r_rsExceptions("SCHED_SRC_CD")
    
    l_SchedExc.ExceptionCode = r_rsExceptions("EXCEPTION_CD")
    l_SchedExc.OverrideExceptionCode = r_rsExceptions("OVER_EXC_CD")
    
    l_SchedExc.Hours = r_rsExceptions("HOURS")
    l_SchedExc.HoursOffset = r_rsExceptions("HOURS_OFFSET_NUM")

    l_SchedExc.JobNumber = r_rsExceptions("JOB_NUM")
    l_SchedExc.ReliefComponent = r_rsExceptions("RELIEF_COMP")
    l_SchedExc.ComponentCode = r_rsExceptions("COMP_CD")
    l_SchedExc.ComponentCode = r_rsExceptions("SHIFT_COMP")
    l_SchedExc.CaseNumber = r_rsExceptions("CASE_NUM")
    l_SchedExc.WorkgroupNumber = r_rsExceptions("WG_NUM")
    l_SchedExc.FamilyLeaveIndicator = r_rsExceptions("FMLA_IND")

'    l_SchedExc.LastModEmpNum = r_rsExceptions("LAST_MOD_EMP_NUM")
    l_SchedExc.LastModTimestamp = r_rsExceptions("LAST_MOD_TIMESTAMP")
    
    CopyScheduledException = l_SchedExc
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CopyScheduledException"
End Function


Public Function IsScheduledWorkDay( _
                    ByRef r_udtScheduleShift As ScheduledShift) _
                As Boolean
    
On Error GoTo eh
    
    IsScheduledWorkDay = IsTimeValid(r_udtScheduleShift.StartTime)
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsScheduledWorkDay"
End Function

Public Function IsScheduledOffDay( _
                    ByRef r_udtScheduleShift As ScheduledShift) _
                As Boolean
    
On Error GoTo eh
    
    IsScheduledOffDay = Not IsScheduledWorkDay(r_udtScheduleShift)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsScheduledOffDay"
End Function

Public Function HasScheduledLunch( _
                    ByRef r_udtScheduleShift As ScheduledShift) _
                As Boolean
    
On Error GoTo eh
    
    HasScheduledLunch = _
        r_udtScheduleShift.LunchStart <> "" And _
        r_udtScheduleShift.LunchEnd <> ""
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "HasScheduledLunch"
End Function


'Public Function ShiftTime( _
'                    ByVal v_strTime As String, _
'                    ByVal v_varHoursOffset As Variant) As String
'    Dim l_lngMinutes As Long
'
'On Error GoTo eh
'    l_lngMinutes = CDec(v_varHoursOffset) * 60
'    ShiftTime = AddMinutes(v_strTime, l_lngMinutes)
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "ShiftTime"
'End Function


Public Function ShiftTime( _
                    ByVal v_strTime As String, _
                    ByVal v_varHoursOffset As Variant, _
                    Optional ByVal v_strLunchStart As String, _
                    Optional ByVal v_strLunchEnd As String, _
                    Optional ByVal v_bln2400 As Boolean) As String
' Answer time, as string hh:mm, by adding v_varHoursOffset to v_strTime.
' If the time between v_strTime and v_varHoursOffset intersects a lunch interval,
' add the lunch interval.
' If requested, convert 00:00 to 24:00.
    Dim l_lngMinutes As Long
    Dim l_lngLunchMinutes As Long
    Dim l_datStart As Date
    Dim l_datLunchStart As Date
    Dim l_datLunchEnd As Date
    Dim l_datStartPlusOffset As Date
    Dim l_datBaseDate As Date
    Dim l_datShiftTime As Date
    
On Error GoTo eh
    If Trim(v_strLunchStart) <> "" Then
        If Trim(v_strLunchEnd) = "" Then
            Err.Raise 30001 + vbObjectError, "Lunch end is required when lunch start is provided"
        End If
    End If

    l_datStart = CDate(v_strTime)
    l_lngMinutes = CDec(v_varHoursOffset) * 60

    If Trim(v_strLunchStart) <> "" Then
        l_datBaseDate = IIf(v_strLunchStart < v_strTime, 1, 0)
        l_datLunchStart = DateTimeFrom(l_datBaseDate, v_strLunchStart)
        l_datBaseDate = IIf(v_strLunchEnd < v_strLunchStart, 1, 0)
        l_datLunchEnd = DateTimeFrom(l_datBaseDate, v_strLunchEnd)
        l_lngLunchMinutes = DateDiff("n", l_datLunchStart, l_datLunchEnd)
    End If

    l_datStartPlusOffset = DateAdd("n", l_lngMinutes, l_datStart)
    If l_lngLunchMinutes Then
        If l_lngMinutes > DateDiff("n", l_datStart, l_datLunchStart) Then
            l_datShiftTime = DateAdd("n", l_lngLunchMinutes, l_datStartPlusOffset)
        Else
            l_datShiftTime = l_datStartPlusOffset
        End If
    Else
        l_datShiftTime = l_datStartPlusOffset
    End If

    ShiftTime = FormatTime(l_datShiftTime, v_bln2400)

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ShiftTime"
End Function


'******************************************************************************
'  Private procedures
'******************************************************************************

Private Function AllocateScheduledShiftByDay( _
                    ByRef r_rsAvailSched As ADODB.Recordset, _
                    ByVal v_strDayofWeek As String) _
                As ScheduledShift
'Get shift data from available schedule for a given day
    Dim l_SchedShift As ScheduledShift
    
On Error GoTo eh
    l_SchedShift.ScheduleCode = r_rsAvailSched("SCHED_CD")
    l_SchedShift.ShiftCode = Trim(r_rsAvailSched(v_strDayofWeek & "_SHIFT_CD"))
    l_SchedShift.ShiftComponent = Trim(r_rsAvailSched(v_strDayofWeek & "_COMP_SHIFT_CD"))
    l_SchedShift.StartTime = Trim(r_rsAvailSched(v_strDayofWeek & "_START_TIME"))
    l_SchedShift.EndTime = Trim(r_rsAvailSched(v_strDayofWeek & "_END_TIME"))
    l_SchedShift.Hours = r_rsAvailSched(v_strDayofWeek & "_STR_HRS")
    l_SchedShift.LunchStart = Trim(r_rsAvailSched(v_strDayofWeek & "_LUNCH_START"))
    l_SchedShift.LunchEnd = Trim(r_rsAvailSched(v_strDayofWeek & "_LUNCH_END"))
    l_SchedShift.OvertimeStart = Trim(r_rsAvailSched(v_strDayofWeek & "_OT_START_TIME"))
    l_SchedShift.OvertimeEnd = Trim(r_rsAvailSched(v_strDayofWeek & "_OT_END_TIME"))
    l_SchedShift.OvertimeHours = r_rsAvailSched(v_strDayofWeek & "_OT_HOURS")
    l_SchedShift.ActualStartDay = Trim(r_rsAvailSched(v_strDayofWeek & "_ACT_DAY"))
    l_SchedShift.ChargeDay = v_strDayofWeek
    AllocateScheduledShiftByDay = l_SchedShift
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "AllocateScheduledShiftByDay"
End Function



Private Sub DetermineShiftDates( _
                ByVal v_strShiftStartDay As String, _
                ByVal v_strChargeDateDayOfWeek As String, _
                ByVal v_datChargeDate As String, _
                ByVal v_strShiftStartTime As String, _
                ByVal v_strShiftEndTime As String, _
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
        If v_strShiftStartTime <= v_strShiftEndTime Then
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

